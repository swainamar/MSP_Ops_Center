#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
msp_runbook.py - ITIL v4 MSP Infrastructure Operations Center  (version 4)
==========================================================================

Single-file web application for an MSP operations team:

  Execution plans   Year > Month > Date & Day > Customer > Task (Ticket ID) > numbered phases
                    (1, 2 ...) > steps (1.1, 1.2.1 ...), chained planned times, go/no-go
                    checkpoints, rollback, run mode, templates, print view.
  Runbook           Customer > Environment > Asset > dated activity records.
  Asset inventory   Full CMDB record per device; device names link to the runbook; one-click
                    SSH (in-browser, recorded) and RDP (.rdp file or Apache Guacamole).
  Incidents         ART standard (Action / Reason / Technical tracking), P1-P4.
  Shift log         Year > Month > Date > Shift > Section > rows; clone whole shifts.
  Bulletin board    Scrolling ticker with read receipts.
  Security          App logins and/or Microsoft 365 (Entra ID) sign-in; roles Viewer, Engineer,
                    Lead, Admin enforced by the server; CSRF protection; login lockout; idle
                    timeout; optional HTTPS; admin-only terminal session recordings.

Central storage (one server, whole team in the browser)
-------------------------------------------------------
  msp_runbook_data.xlsx       assets, runbook, incidents, handovers, bulletins, audit trail
  msp_execution_plans.json    execution plans and plan templates
  msp_shift_hierarchy.json    shift log tree
  msp_metadata.json           version history, import batches, read receipts, terminal session index
  msp_users.json              users, roles, password hashes, settings, trusted SSH host keys
  .recordings/                terminal session transcripts (admin only)
  backups/                    automatic daily JSON backups

Requirements: Python 3.10+, openpyxl. Terminals also need paramiko (+ cryptography, installed
with it). The Windows launcher Start_MSP_Ops_Center.bat installs everything automatically.
Run:  python msp_runbook.py --help      or double-click Start_MSP_Ops_Center.bat
"""
from __future__ import annotations

import argparse
import base64
import calendar
import codecs
import configparser
import copy
import csv
import dataclasses
import errno
import datetime as dt
import hashlib
import hmac
import io
import ipaddress
import json
import os
import re
import secrets
import socket
import ssl
import struct
import sys
import threading
import time
import traceback
import urllib.error
import urllib.parse
import urllib.request
import uuid
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

try:
    import openpyxl
    from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
except ImportError:  # pragma: no cover
    sys.stderr.write("openpyxl is required. Install it with:  pip install openpyxl\n")
    raise

try:  # optional: the shift roster lives in msp_roster.py next to this file
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
    import msp_roster
except ImportError:  # pragma: no cover - the app runs without a roster
    msp_roster = None
try:  # optional: team access rules (isolation) live in msp_access.py next to this file
    import msp_access
except ImportError:  # pragma: no cover
    msp_access = None
try:  # optional: teams, customers and projects live in msp_org.py next to this file
    import msp_org
except ImportError:  # pragma: no cover
    msp_org = None
try:  # optional: the built-in assistant lives in msp_ai.py next to this file
    import msp_ai
except ImportError:  # pragma: no cover - the app runs without the assistant
    msp_ai = None

STATIC_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "static")
STATIC_TYPES = {".woff2": "font/woff2", ".js": "text/javascript; charset=utf-8", ".css": "text/css; charset=utf-8",
                ".svg": "image/svg+xml", ".png": "image/png", ".txt": "text/plain; charset=utf-8"}
STATIC_NAME = re.compile(r"(?:[a-z0-9_-]+/){0,3}[A-Za-z0-9][A-Za-z0-9._-]*")

APP_NAME = "MSP Ops Center"
APP_VERSION = "4.0.0"

XLSX_NAME = "msp_runbook_data.xlsx"
PLANS_NAME = "msp_execution_plans.json"
JSON_NAME = "msp_shift_hierarchy.json"
META_NAME = "msp_metadata.json"
USERS_NAME = "msp_users.json"
RECORDINGS_DIR = ".recordings"
BACKUPS_DIR = "backups"

ID_COL = "Record_ID"
SYSTEM_COLS = ["Version", "Created_By", "Created_At", "Updated_By", "Updated_At"]
AUDIT_SHEET = "Audit_Trail"
COLUMNS_SHEET = "_Columns"
AUDIT_HEADERS = ["Timestamp", "User", "Area", "Action", "Target", "Details"]
AUDIT_CAP = 5000
HISTORY_PER_RECORD = 30
IMPORT_BATCH_CAP = 200
MAX_BODY = 30 * 1024 * 1024
MAX_COLUMNS = 80

PRIORITIES = ["P1", "P2", "P3", "P4"]
SHIFTS = ["Morning", "Afternoon", "Night"]
SECTION_TEMPLATES: list[tuple[str, list[str]]] = [
    ("Shift Handover Overview", ["Check", "Client", "Status", "Details", "Owner"]),
    ("Scheduled", ["Client", "Event", "Assets", "Window_Start", "Window_End", "Change_Ref", "Status"]),
    ("On-Hold OR Customer Hold", ["Ticket", "Client", "Priority", "Blocked_On", "Next_Follow_Up", "Notes"]),
    ("Patch and firmware details", ["Client", "Asset", "Current_Level", "Target_Level", "Compliance", "Notes"]),
    ("In Work - All Clients", ["Ticket", "Client", "Priority", "Task", "Owner", "Status", "Next_Step"]),
    ("Weekly Meetings", ["Meeting", "Client", "Schedule", "Attendees", "MOM", "Updates"]),
    ("Weekly Tasks", ["Task", "Client", "Frequency", "Playbook", "Owner", "Status"]),
]
CUSTOM_SECTION_COLUMNS = ["Item", "Details", "Owner", "Status"]

# ---- flat tables (stored in the workbook) --------------------------------------------
TABLES: dict[str, dict] = {
    "Asset_Master": {
        "module": "assets", "title": "CMDB: configuration items", "noun": "CI", "label_key": "Device",
        "core": [ID_COL, "Customer", "Serial_No", "Device", "CI_Type", "Hostname_FQDN", "IP_Address", "Environment",
                 "Platform_Type", "Virtualization_Cloud", "OS", "OS_Version", "Application", "Role", "CPU",
                 "Memory", "Root_Disk", "Data_Disk", "Data_Centre_Rack", "Criticality", "Patch_Group",
                 "Maintenance_Window", "Backup_Policy", "Monitoring", "Warranty_End", "OS_EOL",
                 "Upstream_Dependencies", "Downstream_Dependencies", "Special_Care", "Owner", "Contacts",
                 "Status", "Owner_Team", "Sharing", "Shared_With", "Shared_Fields"],
        "required": ["Device"], "dup_keys": [["Serial_No"], ["Customer", "Device"]],
        "choices": {"Platform_Type": ["Physical", "Virtualization", "Cloud"],
                    "Environment": ["Production", "DR", "UAT", "Test", "Development"],
                    "Criticality": ["Critical", "High", "Medium", "Low"],
                    "Status": ["Planned", "In build", "Active", "Maintenance", "Decommissioned"],
                    "Sharing": ["Shared core", "Fully shared", "Team only"],
                    "CI_Type": ["Server", "Virtual machine", "Database", "Application", "Middleware", "Network device",
                                "Storage", "Cloud service", "Other"]},
        "dates": ["Warranty_End", "OS_EOL"],
    },
    "Runbook": {
        "module": "runbook", "title": "Runbook", "noun": "runbook entry", "label_key": "Title",
        "core": [ID_COL, "Customer", "Environment", "Asset", "Date", "Category", "Title", "Note", "Solution",
                 "Commands", "Ticket_Ref", "Engineer", "Tags", "Known_Issue", "Owner_Team", "Shared_With"],
        "required": ["Asset"], "dup_keys": [["Customer", "Asset", "Date", "Title"]],
        "choices": {"Category": ["Incident", "Change", "Maintenance", "Known issue", "Workaround", "Health check",
                                 "Information"],
                    "Known_Issue": ["No", "Yes"]},
        "dates": ["Date"],
    },
    "Incidents": {
        "module": "incidents", "title": "Incidents", "noun": "incident", "label_key": "Ticket_ID",
        "core": [ID_COL, "Ticket_ID", "Customer", "Asset", "Priority", "Status", "Summary", "Action", "Reason",
                 "Technical_Tracking", "Engineer", "Opened_At", "Resolved_At"],
        "required": ["Action", "Reason", "Technical_Tracking"], "dup_keys": [["Ticket_ID"]],
        "choices": {"Priority": PRIORITIES, "Status": ["New", "In Progress", "On Hold", "Resolved", "Closed"]},
        "dates": [],
    },
    "Shift_Transitions": {
        "module": "handovers", "title": "Shift handovers", "noun": "handover", "label_key": "Date",
        "core": [ID_COL, "Date", "Shift", "Region", "Outgoing_Engineer", "Incoming_Engineer", "P1_P2_In_Flight",
                 "Planned_Changes", "High_Risk_Assets", "Notes", "Acknowledged", "Team"],
        "required": ["Date", "Shift"], "dup_keys": [["Date", "Shift", "Region"]],
        "choices": {"Shift": SHIFTS, "Region": ["US Day", "APAC Night", "EMEA Morning"],
                    "Acknowledged": ["No", "Yes"]},
        "dates": ["Date"],
    },
    "Bulletins": {
        "module": "bulletins", "title": "Bulletin board", "noun": "bulletin", "label_key": "Title",
        "core": [ID_COL, "Title", "Severity", "Customers", "Body", "Link", "Start", "Expiry", "Pinned",
                 "Posted_By"],
        "required": ["Title"], "dup_keys": [["Title", "Start"]],
        "choices": {"Severity": ["Info", "Advisory", "Critical"], "Pinned": ["No", "Yes"]},
        "dates": [],
    },
}
SHEET_ORDER = list(TABLES)
MODULES = {t["module"]: s for s, t in TABLES.items()}
LABELS = {
    "Record_ID": "Record ID", "Serial_No": "Serial No", "Hostname_FQDN": "Hostname / FQDN", "IP_Address": "IP address",
    "Platform_Type": "Platform type", "Virtualization_Cloud": "Hypervisor / cloud account", "OS": "OS",
    "OS_Version": "OS version", "CPU": "CPU", "Root_Disk": "Root disk", "Data_Disk": "Data disk",
    "Data_Centre_Rack": "Data centre / rack", "Patch_Group": "Patch group", "Maintenance_Window": "Maintenance window",
    "Backup_Policy": "Backup policy", "Warranty_End": "Warranty end", "OS_EOL": "OS end of life",
    "Upstream_Dependencies": "Upstream dependencies", "Downstream_Dependencies": "Downstream dependencies",
    "Special_Care": "Special care", "Ticket_Ref": "Ticket ref", "Known_Issue": "Known issue",
    "Ticket_ID": "Ticket ID", "Technical_Tracking": "Technical tracking", "Opened_At": "Opened at",
    "Resolved_At": "Resolved at", "Outgoing_Engineer": "Outgoing engineer", "Incoming_Engineer": "Incoming engineer",
    "P1_P2_In_Flight": "P1/P2 in flight", "Planned_Changes": "Planned changes", "High_Risk_Assets": "High-risk CIs",
    "CI_Type": "CI type",
    "Owner_Team": "Owner team", "Sharing": "Sharing", "Shared_With": "Shared with teams", "Shared_Fields": "Shared fields",
    "Posted_By": "Posted by", "Created_By": "Created by", "Created_At": "Created at", "Updated_By": "Updated by",
    "Updated_At": "Updated at",
}
WRAP_KEYS = {"Special_Care", "Action", "Reason", "Technical_Tracking", "Commands", "Command", "Notes", "Note",
             "Solution", "Summary", "Body", "Upstream_Dependencies", "Downstream_Dependencies", "P1_P2_In_Flight",
             "Planned_Changes", "High_Risk_Assets", "Step", "Rollback", "Expected", "Details"}

HEADER_ALIASES = {
    "id": ID_COL, "recordid": ID_COL, "rowid": ID_COL,
    "client": "Customer", "customername": "Customer", "account": "Customer",
    "serial": "Serial_No", "serialnumber": "Serial_No", "sn": "Serial_No", "assettag": "Serial_No",
    "device": "Device", "devicename": "Device", "hostname": "Device", "host": "Device", "servername": "Device",
    "server": "Device", "fqdn": "Hostname_FQDN", "ip": "IP_Address", "ipaddress": "IP_Address", "ipaddr": "IP_Address",
    "env": "Environment", "platform": "Platform_Type", "hypervisor": "Virtualization_Cloud",
    "virtualizationplatform": "Virtualization_Cloud", "cloud": "Virtualization_Cloud", "os": "OS",
    "operatingsystem": "OS", "version": "OS_Version", "osversion": "OS_Version", "app": "Application",
    "applications": "Application", "cpus": "CPU", "vcpu": "CPU", "cores": "CPU", "ram": "Memory", "mem": "Memory",
    "rootdisk": "Root_Disk", "osdisk": "Root_Disk", "datadisk": "Data_Disk", "datadisks": "Data_Disk",
    "location": "Data_Centre_Rack", "datacenter": "Data_Centre_Rack", "datacentre": "Data_Centre_Rack",
    "rack": "Data_Centre_Rack", "warranty": "Warranty_End", "eol": "OS_EOL", "upstream": "Upstream_Dependencies",
    "downstream": "Downstream_Dependencies", "specialcare": "Special_Care", "specialcarenotes": "Special_Care",
    "caveats": "Special_Care", "asset": "Asset", "assetname": "Asset", "notes": ["Notes", "Note"], "note": ["Note", "Notes"],
    "problem": "Note", "issue": "Note", "solution": "Solution", "resolution": "Solution", "fix": "Solution",
    "commands": ["Commands", "Technical_Tracking"], "command": ["Commands"], "ticket": ["Ticket_ID", "Ticket_Ref"], "ticketno": ["Ticket_ID", "Ticket_Ref"],
    "ticketnumber": ["Ticket_ID", "Ticket_Ref"], "incidentid": "Ticket_ID", "ticketref": ["Ticket_Ref", "Ticket_ID"], "severity": ["Severity", "Priority"],
    "actiontaken": "Action", "rootcause": "Reason", "technical": "Technical_Tracking",
    "commandsexecuted": "Technical_Tracking", "knownissue": "Known_Issue", "category": "Category",
    "type": "Category", "message": "Body", "details": "Body", "url": "Link", "expires": "Expiry",
    "expirydate": "Expiry", "startdate": "Start", "pinned": "Pinned",
}
# --------------------------------------------------------------------------------------
# Errors and small helpers
# --------------------------------------------------------------------------------------
class ApiError(Exception):
    """An error that maps directly onto an HTTP status and a user-facing message."""

    def __init__(self, status: int, message: str):
        super().__init__(message)
        self.status = status
        self.message = message


class StorageError(ApiError):
    """Central storage (disk or SharePoint) could not be read or written."""


class ConflictError(Exception):
    """The stored file changed since we last read it (ETag / version mismatch)."""


def now_iso() -> str:
    return dt.datetime.now().astimezone().isoformat(timespec="seconds")


def new_id(prefix: str = "") -> str:
    return prefix + uuid.uuid4().hex[:12]


def norm_key(value) -> str:
    return re.sub(r"[^a-z0-9]", "", str(value or "").lower())


def cell_str(value) -> str:
    """Convert any spreadsheet cell value into the canonical string form used in memory."""
    if value is None:
        return ""
    if isinstance(value, bool):
        return "Yes" if value else "No"
    if isinstance(value, dt.datetime):
        if value.hour == 0 and value.minute == 0 and value.second == 0:
            return value.date().isoformat()
        return value.isoformat(sep=" ", timespec="minutes")
    if isinstance(value, dt.date):
        return value.isoformat()
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def clean_text(value, limit: int = 32000) -> str:
    """Sanitise free text for storage (removes characters Excel rejects, caps length)."""
    if value is None:
        return ""
    if isinstance(value, bool):
        value = "Yes" if value else "No"
    text = ILLEGAL_CHARACTERS_RE.sub("", str(value))
    return text[:limit]


def truthy(value) -> bool:
    return str(value).strip().lower() in ("1", "y", "yes", "true", "done", "x", "✓", "complete", "completed")


def http_open(req, timeout: float):
    """urlopen restricted to http(s) URLs (never file:// or custom schemes)."""
    url = req.full_url if isinstance(req, urllib.request.Request) else str(req)
    if urllib.parse.urlsplit(url).scheme not in ("https", "http"):
        raise StorageError(400, "Only http(s) URLs are allowed.")
    return urllib.request.urlopen(req, timeout=timeout)  # nosec B310 - scheme checked above


def hide_path(path: str):
    """Mark a folder hidden on Windows (best effort; no effect elsewhere)."""
    if os.name == "nt":
        try:
            import ctypes
            os.makedirs(path, exist_ok=True)
            ctypes.windll.kernel32.SetFileAttributesW(str(path), 0x02)
        except Exception:  # nosec B110 - cosmetic, best effort
            pass


def audit_entry(user: str, area: str, action: str, target: str, details: str = "") -> dict:
    return {
        "Timestamp": dt.datetime.now().astimezone().isoformat(timespec="milliseconds"),
        "User": clean_text(user, 120) or "anonymous",
        "Area": area,
        "Action": action,
        "Target": clean_text(target, 300),
        "Details": clean_text(details, 2000),
    }


# --------------------------------------------------------------------------------------
# Storage back ends
# --------------------------------------------------------------------------------------
class LocalStorage:
    """Files in a directory. Can be a OneDrive-synced SharePoint folder on the server host."""

    kind = "local"

    def __init__(self, data_dir: str):
        self.data_dir = os.path.abspath(data_dir)
        os.makedirs(self.data_dir, exist_ok=True)
        self._lock = threading.Lock()

    def _path(self, name: str) -> str:
        raw = str(name).replace("\\", "/")
        parts = [p for p in raw.split("/") if p]
        if raw.startswith("/") or not parts or any(p in (".", "..") or ":" in p for p in parts):
            raise StorageError(400, f"Invalid storage path '{name}'.")
        return os.path.join(self.data_dir, *parts)

    def version(self, name: str):
        try:
            st = os.stat(self._path(name))
        except FileNotFoundError:
            return None
        return f"{st.st_mtime_ns}:{st.st_size}"

    def read(self, name: str):
        path = self._path(name)
        try:
            with open(path, "rb") as fh:
                data = fh.read()
        except FileNotFoundError:
            return None, None
        except PermissionError as exc:
            raise StorageError(423, f"{name} is locked by another program (is it open in Excel?): {exc}")
        return data, self.version(name)

    def write(self, name: str, data: bytes, expected):
        path = self._path(name)
        with self._lock:
            current = self.version(name)
            if current != expected:
                raise ConflictError(name)
            tmp = f"{path}.{uuid.uuid4().hex[:8]}.tmp"
            os.makedirs(os.path.dirname(path), exist_ok=True)
            if name.startswith(RECORDINGS_DIR):
                hide_path(os.path.join(self.data_dir, RECORDINGS_DIR))
            try:
                with open(tmp, "wb") as fh:
                    fh.write(data)
                    fh.flush()
                    os.fsync(fh.fileno())
                os.replace(tmp, path)
            except PermissionError as exc:
                try:
                    os.remove(tmp)
                except OSError:
                    pass
                raise StorageError(423, f"{name} is locked by another program (close it in Excel and retry): {exc}")
            return self.version(name)

    def delete(self, name: str):
        try:
            os.remove(self._path(name))
        except FileNotFoundError:
            pass

    def describe(self) -> dict:
        synced = any(k in self.data_dir.lower() for k in ("onedrive", "sharepoint"))
        return {
            "kind": self.kind,
            "label": "OneDrive-synced folder" if synced else "Server folder",
            "location": self.data_dir,
        }


class GraphStorage:
    """Files in a SharePoint document library or OneDrive for Business drive via Microsoft Graph.

    Authentication uses the OAuth2 client-credentials flow of an Entra ID app registration.
    Recommended application permission: Sites.Selected (granted on the one team site only).
    """

    kind = "sharepoint"

    def __init__(self, *, tenant_id: str, client_id: str, client_secret: str,
                 site_host: str = "", site_path: str = "", library: str = "",
                 drive_id: str = "", onedrive_user: str = "", folder: str = "MSP-Runbook",
                 graph_base: str = "https://graph.microsoft.com/v1.0",
                 auth_base: str = "https://login.microsoftonline.com", timeout: float = 30.0):
        missing = [n for n, v in (("tenant id", tenant_id), ("client id", client_id),
                                  ("client secret", client_secret)) if not v]
        if missing:
            raise ValueError("SharePoint storage needs: " + ", ".join(missing))
        if not (drive_id or site_host or onedrive_user):
            raise ValueError("SharePoint storage needs a site host (+ path), a drive id or a OneDrive user")
        self.tenant_id, self.client_id, self.client_secret = tenant_id, client_id, client_secret
        self.site_host = site_host.strip().replace("https://", "").strip("/")
        self.site_path = site_path.strip()
        self.library = library.strip()
        self.onedrive_user = onedrive_user.strip()
        self.folder = folder.strip().strip("/")
        self.graph_base = graph_base.rstrip("/")
        self.auth_base = auth_base.rstrip("/")
        self.timeout = timeout
        self._drive = drive_id.strip() or None
        self._token_value = None
        self._token_expiry = 0.0
        self._ensured: set = set()
        self._lock = threading.RLock()

    # -- HTTP plumbing ---------------------------------------------------------------
    def _token(self, force: bool = False) -> str:
        with self._lock:
            if not force and self._token_value and time.time() < self._token_expiry - 60:
                return self._token_value
            body = urllib.parse.urlencode({
                "client_id": self.client_id,
                "client_secret": self.client_secret,
                "grant_type": "client_credentials",
                "scope": "https://graph.microsoft.com/.default",
            }).encode()
            req = urllib.request.Request(
                f"{self.auth_base}/{urllib.parse.quote(self.tenant_id)}/oauth2/v2.0/token",
                data=body, method="POST",
                headers={"Content-Type": "application/x-www-form-urlencoded"})
            try:
                with http_open(req, self.timeout) as resp:
                    payload = json.loads(resp.read().decode("utf-8"))
            except urllib.error.HTTPError as exc:
                detail = exc.read()[:400].decode("utf-8", "replace")
                raise StorageError(502, f"SharePoint sign-in failed (HTTP {exc.code}): {detail}")
            except urllib.error.URLError as exc:
                raise StorageError(503, f"Cannot reach Microsoft sign-in: {exc.reason}")
            self._token_value = payload["access_token"]
            self._token_expiry = time.time() + int(payload.get("expires_in", 3600))
            return self._token_value

    def _call(self, method: str, url: str, data: bytes | None = None, headers: dict | None = None):
        attempt, refreshed = 0, False
        while True:
            hdrs = {"Authorization": f"Bearer {self._token()}", "Accept": "application/json"}
            hdrs.update(headers or {})
            req = urllib.request.Request(url, data=data, method=method, headers=hdrs)
            try:
                with http_open(req, self.timeout) as resp:
                    return resp.status, dict(resp.headers), resp.read()
            except urllib.error.HTTPError as exc:
                status, body = exc.code, exc.read()
                rh = dict(exc.headers or {})
                if status == 401 and not refreshed:
                    refreshed = True
                    self._token(force=True)
                    continue
                if status in (429, 503, 504) and attempt < 4:
                    attempt += 1
                    try:
                        wait = float(rh.get("Retry-After", 0)) or 2 ** attempt
                    except ValueError:
                        wait = 2 ** attempt
                    time.sleep(min(wait, 30))
                    continue
                return status, rh, body
            except urllib.error.URLError as exc:
                if attempt < 2:
                    attempt += 1
                    time.sleep(1)
                    continue
                raise StorageError(503, f"Cannot reach Microsoft Graph: {exc.reason}")

    @staticmethod
    def _snippet(body: bytes) -> str:
        try:
            err = json.loads(body.decode("utf-8")).get("error", {})
            return f"{err.get('code', '')}: {err.get('message', '')}".strip(": ")
        except Exception:
            return body[:300].decode("utf-8", "replace")

    # -- location resolution ---------------------------------------------------------
    def _drive_id(self) -> str:
        with self._lock:
            if self._drive:
                return self._drive
            g = self.graph_base
            if self.site_host:
                rel = self.site_path.strip("/")
                url = f"{g}/sites/{self.site_host}" + (f":/{urllib.parse.quote(rel)}" if rel else "")
                st, _, body = self._call("GET", url)
                if st != 200:
                    raise StorageError(502, f"SharePoint site '{self.site_host}/{rel}' not reachable "
                                            f"(HTTP {st}). Check the URL and that the app has access. "
                                            f"{self._snippet(body)}")
                site_id = json.loads(body)["id"]
                sid = urllib.parse.quote(site_id, safe=",")
                st, _, body = self._call("GET", f"{g}/sites/{sid}/drives")
                drives = json.loads(body).get("value", []) if st == 200 else []
                want = norm_key(self.library or "Documents")
                aliases = {"documents", "shareddocuments"}
                chosen = None
                for d in drives:
                    name = norm_key(d.get("name", ""))
                    if name == want or (want in aliases and name in aliases):
                        chosen = d["id"]
                        break
                if not chosen:
                    if self.library:
                        names = ", ".join(d.get("name", "?") for d in drives) or "none visible"
                        raise StorageError(502, f"Library '{self.library}' not found on the site. "
                                                f"Available: {names}")
                    st, _, body = self._call("GET", f"{g}/sites/{sid}/drive")
                    if st != 200:
                        raise StorageError(502, f"Default library not reachable: {self._snippet(body)}")
                    chosen = json.loads(body)["id"]
                self._drive = chosen
            elif self.onedrive_user:
                st, _, body = self._call("GET", f"{g}/users/{urllib.parse.quote(self.onedrive_user)}/drive")
                if st != 200:
                    raise StorageError(502, f"OneDrive of '{self.onedrive_user}' not reachable: "
                                            f"{self._snippet(body)}")
                self._drive = json.loads(body)["id"]
            return self._drive

    def _rel(self, name: str | None = None) -> str:
        parts = [p for p in self.folder.split("/") if p]
        if name:
            parts.append(name)
        return "/".join(urllib.parse.quote(p) for p in parts)

    def _item_url(self, name: str) -> str:
        return f"{self.graph_base}/drives/{self._drive_id()}/root:/{self._rel(name)}"

    def _ensure_folder(self, extra: list[str] | None = None):
        segments = [p for p in self.folder.split("/") if p] + list(extra or [])
        key = "/".join(segments)
        if key in self._ensured or not segments:
            return
        drive = self._drive_id()
        done: list[str] = []
        for seg in segments:
            parent = "/".join(urllib.parse.quote(p) for p in done)
            done.append(seg)
            path = "/".join(urllib.parse.quote(p) for p in done)
            st, _, body = self._call("GET", f"{self.graph_base}/drives/{drive}/root:/{path}")
            if st == 200:
                continue
            if st != 404:
                raise StorageError(502, f"Cannot inspect folder '{seg}': {self._snippet(body)}")
            url = (f"{self.graph_base}/drives/{drive}/root"
                   + (f":/{parent}:" if parent else "") + "/children")
            payload = json.dumps({"name": seg, "folder": {},
                                  "@microsoft.graph.conflictBehavior": "fail"}).encode()
            st, _, body = self._call("POST", url, payload, {"Content-Type": "application/json"})
            if st not in (200, 201, 409):
                raise StorageError(502, f"Cannot create folder '{seg}': {self._snippet(body)}")
        self._ensured.add(key)

    # -- storage interface ------------------------------------------------------------
    def version(self, name: str):
        st, _, body = self._call("GET", self._item_url(name))
        if st == 404:
            return None
        if st != 200:
            raise StorageError(502, f"Cannot read {name} metadata from SharePoint: {self._snippet(body)}")
        return json.loads(body).get("eTag")

    def read(self, name: str):
        st, _, body = self._call("GET", self._item_url(name))
        if st == 404:
            return None, None
        if st != 200:
            raise StorageError(502, f"Cannot read {name} from SharePoint: {self._snippet(body)}")
        meta = json.loads(body)
        etag = meta.get("eTag")
        download = meta.get("@microsoft.graph.downloadUrl")
        if download:
            # Pre-authenticated URL: must be fetched WITHOUT the bearer token.
            try:
                with http_open(urllib.request.Request(download), self.timeout) as r:
                    data = r.read()
            except (urllib.error.HTTPError, urllib.error.URLError) as exc:
                raise StorageError(502, f"Cannot download {name} from SharePoint: {exc}")
        else:
            st, _, data = self._call("GET", self._item_url(name) + ":/content")
            if st != 200:
                raise StorageError(502, f"Cannot download {name}: {self._snippet(data)}")
        return data, etag

    def write(self, name: str, data: bytes, expected):
        self._ensure_folder([p for p in name.split("/") if p][:-1])
        url = self._item_url(name) + ":/content"
        headers = {"Content-Type": "application/octet-stream"}
        if expected:
            headers["If-Match"] = expected
        else:
            url += "?@microsoft.graph.conflictBehavior=fail"
        st, _, body = self._call("PUT", url, data, headers)
        if st in (409, 412):
            raise ConflictError(name)
        if st == 423:
            raise StorageError(423, f"{name} is locked in SharePoint (checked out or open for exclusive "
                                    f"editing). Close it and retry.")
        if st not in (200, 201):
            raise StorageError(502, f"Cannot save {name} to SharePoint (HTTP {st}): {self._snippet(body)}")
        return json.loads(body).get("eTag")

    def delete(self, name: str):
        st, _, body = self._call("DELETE", self._item_url(name))
        if st not in (200, 204, 404):
            raise StorageError(502, f"Cannot delete {name} in SharePoint: {self._snippet(body)}")

    def describe(self) -> dict:
        if self.site_host:
            where = f"{self.site_host}/{self.site_path.strip('/')}".rstrip("/")
            where += f" > {self.library or 'Documents'}"
        elif self.onedrive_user:
            where = f"OneDrive of {self.onedrive_user}"
        else:
            where = f"drive {self._drive}"
        if self.folder:
            where += f" > {self.folder}"
        return {"kind": self.kind, "label": "SharePoint", "location": where}




# --------------------------------------------------------------------------------------
# Flat tables: configurable columns, header mapping, validation
# --------------------------------------------------------------------------------------
DEFAULT_HIDDEN = {
    "Shift_Transitions": ["Team"],  # team-owned: each team only sees its own, so the column is not needed on screen
    "Asset_Master": {"Hostname_FQDN", "Role", "Data_Centre_Rack", "Patch_Group", "Maintenance_Window",
                     "Backup_Policy", "Monitoring", "Warranty_End", "OS_EOL", "Upstream_Dependencies",
                     "Downstream_Dependencies", "Owner", "Contacts"},
    "Runbook": set(),
    "Incidents": {"Action", "Reason", "Technical_Tracking", "Resolved_At"},
    "Shift_Transitions": {"Notes"},
    "Bulletins": {"Body", "Link"},
}
SHEET_ALIASES: dict[str, str] = {}
for _s, _t in TABLES.items():
    for _n in (_s, _t["module"], _t["title"]):
        SHEET_ALIASES[norm_key(_n)] = _s
SHEET_ALIASES.update({
    "assets": "Asset_Master", "asset": "Asset_Master", "assetinventory": "Asset_Master", "cmdb": "Asset_Master",
    "inventory": "Asset_Master", "masterassetinventory": "Asset_Master", "runbooks": "Runbook",
    "activities": "Runbook", "assetactivities": "Runbook", "activitylog": "Runbook", "incident": "Incidents",
    "incidentlog": "Incidents", "tickets": "Incidents", "handover": "Shift_Transitions",
    "shifthandover": "Shift_Transitions", "shifttransition": "Shift_Transitions", "bulletin": "Bulletins",
    "notices": "Bulletins", "announcements": "Bulletins",
    "executionplans": "plans", "executionplan": "plans", "plans": "plans", "plan": "plans", "wbs": "plans",
    "executionwbs": "plans", "changeplan": "plans", "changeplans": "plans",
    norm_key(AUDIT_SHEET): AUDIT_SHEET, "audit": AUDIT_SHEET, "auditlog": AUDIT_SHEET,
    norm_key(COLUMNS_SHEET): COLUMNS_SHEET, "shifthierarchy": "shiftlog", "shiftlog": "shiftlog", "shifts": "shiftlog", "shifthandoverlog": "shiftlog",
    "instructions": "instructions", "readme": "instructions",
})


def canonical_sheet(name: str):
    return SHEET_ALIASES.get(norm_key(name))


# ITIL service configuration management wording for the CMDB (labels only; keys and stored data are unchanged)
SHEET_LABELS = {"Asset_Master": {"Device": "CI name", "Status": "CI status"},
                "Runbook": {"Asset": "CI"}, "Incidents": {"Asset": "CI"}}
OLD_DEFAULT_LABELS = {"Device": "Device", "Status": "Status", "Asset": "Asset", "High_Risk_Assets": "High-risk assets"}


# Tables whose rows belong to one team once teams exist: sheet -> column holding the team id (hidden by default).
TEAM_TABLES = {"Shift_Transitions": "Team"}


def default_label(key: str, sheet: str = "") -> str:
    return SHEET_LABELS.get(sheet, {}).get(key) or LABELS.get(key, key.replace("_", " "))


def schema_keys(sheet: str) -> list[str]:
    return TABLES[sheet]["core"] + SYSTEM_COLS


def default_columns(sheet: str) -> list[dict]:
    cols = []
    for k in schema_keys(sheet):
        system = k == ID_COL or k in SYSTEM_COLS
        cols.append({"key": k, "label": default_label(k, sheet), "core": True, "system": system,
                     "hidden": system or k in DEFAULT_HIDDEN[sheet]})
    return cols


def ensure_columns(sheet: str, cols: list[dict] | None) -> list[dict]:
    """Merge a stored/edited column config with the schema: core columns can never disappear."""
    schema = schema_keys(sheet)
    defaults = {c["key"]: c for c in default_columns(sheet)}
    out, seen = [], set()
    for c in cols or []:
        key = str(c.get("key") or "").strip()
        if not key or key in seen:
            continue
        seen.add(key)
        base = defaults.get(key)
        label = clean_text(c.get("label") or (base["label"] if base else key), 80).strip() or key
        if base and label == OLD_DEFAULT_LABELS.get(key):
            label = base["label"]  # an untouched old default label moves to the CMDB wording; custom names stay
        out.append({"key": key, "label": label, "core": key in schema, "system": bool(base and base["system"]),
                    "hidden": True if base and base["system"] else bool(c.get("hidden", False))})
    for k in schema:
        if k not in seen:
            out.append(dict(defaults[k]))
    for c in out:  # a team-ownership column holds an internal id and is never shown: each team only sees its own rows
        if c["key"] == TEAM_TABLES.get(sheet):
            c["hidden"] = True
    return out


def match_header(header: str, columns: list[dict], sheet: str):
    k = norm_key(header)
    if not k:
        return None
    for c in columns:
        if norm_key(c["label"]) == k or norm_key(c["key"]) == k:
            return c["key"]
    schema = schema_keys(sheet)
    for key in schema:
        if norm_key(key) == k or norm_key(default_label(key)) == k:
            return key
    alias = HEADER_ALIASES.get(k)
    for a in (alias if isinstance(alias, list) else [alias]):
        if a in schema:
            return a
    return None


def custom_key(label: str, existing: set) -> str:
    base = "x_" + (re.sub(r"[^A-Za-z0-9]+", "_", label).strip("_")[:30] or "col")
    key, n = base, 2
    while key in existing:
        key, n = f"{base}_{n}", n + 1
    return key


def map_headers(headers: list[str], columns: list[dict], sheet: str):
    """Return (keys per header position, new custom columns). Unknown headers become custom columns."""
    keys, new_cols, used = [], [], {c["key"] for c in columns}
    for h in headers:
        key = match_header(h, columns + new_cols, sheet)
        if key is None and str(h).strip() and not re.fullmatch(r"Column_\d+", str(h)):
            key = custom_key(str(h), used)
            used.add(key)
            new_cols.append({"key": key, "label": clean_text(h, 80).strip(), "core": False,
                             "system": False, "hidden": False})
        keys.append(key)
    return keys, new_cols


def empty_table(sheet: str) -> dict:
    return {"columns": default_columns(sheet), "rows": []}


def default_tables() -> dict:
    return {s: empty_table(s) for s in SHEET_ORDER}


def blank_row(table: dict) -> dict:
    return {c["key"]: "" for c in table["columns"]}


def set_columns(sheet: str, table: dict, columns) -> dict:
    """Apply a column edit from the UI: rename, reorder, hide/show, add or remove custom columns."""
    if not isinstance(columns, list) or not columns:
        raise ApiError(400, "Provide the full list of columns.")
    if len(columns) > MAX_COLUMNS:
        raise ApiError(400, f"A table can have at most {MAX_COLUMNS} columns.")
    current = {c["key"]: c for c in table["columns"]}
    used, labels, new_cfg = set(current), set(), []
    for c in columns:
        if not isinstance(c, dict):
            raise ApiError(400, "Each column must be an object with a label.")
        label = clean_text(c.get("label", ""), 80).strip()
        if not label:
            raise ApiError(400, "Column names cannot be empty.")
        if label.lower() in labels:
            raise ApiError(400, f"Duplicate column name '{label}'.")
        labels.add(label.lower())
        key = c.get("key")
        if key not in current:
            key = custom_key(label, used)
            used.add(key)
        new_cfg.append({"key": key, "label": label, "hidden": bool(c.get("hidden", False))})
    # system columns are not edited in the UI; keep their labels, placed where they were
    for k, c in current.items():
        if c.get("system") and k not in {x["key"] for x in new_cfg}:
            new_cfg.append({"key": k, "label": c["label"], "hidden": True})
    merged = ensure_columns(sheet, new_cfg)
    kept = {c["key"] for c in merged}
    removed = [c["label"] for k, c in current.items() if k not in kept]
    for r in table["rows"]:
        for k in list(r):
            if k not in kept:
                del r[k]
        for k in kept:
            r.setdefault(k, "")
    table["columns"] = merged
    return {"removed": removed, "columns": merged}


# --------------------------------------------------------------------------------------
# Date / time normalisation
# --------------------------------------------------------------------------------------
_DATE_FORMATS = ["%Y-%m-%d", "%Y/%m/%d", "%d/%m/%Y", "%d-%m-%Y", "%d.%m.%Y", "%d-%b-%Y", "%d %b %Y",
                 "%d-%B-%Y", "%d %B %Y", "%b %d %Y", "%Y%m%d"]


def norm_date(value, field: str = "Date") -> str:
    s = str(value or "").strip()
    if not s:
        return ""
    s = re.split(r"[ T]", s, maxsplit=1)[0] if re.match(r"\d{4}-\d{2}-\d{2}[ T]", s) else s
    for fmt in _DATE_FORMATS:
        try:
            return dt.datetime.strptime(s, fmt).date().isoformat()
        except ValueError:
            pass
    raise ApiError(400, f"{default_label(field)} '{value}' is not a date. Use YYYY-MM-DD (or DD/MM/YYYY).")


def norm_datetime(value, field: str = "Date & time") -> str:
    s = str(value or "").strip().replace("T", " ")
    if not s:
        return ""
    s = re.sub(r"(\d{2}:\d{2}):\d{2}(\.\d+)?$", r"\1", s)
    for fmt in ("%Y-%m-%d %H:%M", "%d/%m/%Y %H:%M", "%d-%m-%Y %H:%M", "%Y/%m/%d %H:%M", "%d-%b-%Y %H:%M"):
        try:
            return dt.datetime.strptime(s, fmt).strftime("%Y-%m-%d %H:%M")
        except ValueError:
            pass
    try:
        return norm_date(s, field) + " 00:00"
    except ApiError:
        raise ApiError(400, f"{field} '{value}' is not a date/time. Use YYYY-MM-DD HH:MM.")


def now_local() -> str:
    return dt.datetime.now().strftime("%Y-%m-%d %H:%M")


def parse_dt(s: str):
    try:
        return dt.datetime.strptime(s, "%Y-%m-%d %H:%M") if s else None
    except ValueError:
        return None


def _choice(value: str, options: list[str]):
    for o in options:
        if norm_key(o) == norm_key(value):
            return o
    return None


STRICT_CHOICES = {"Priority", "Severity", "Shift"}
YES_NO = {"Known_Issue", "Pinned", "Acknowledged"}


def prepare_row(sheet: str, row: dict, user: str, creating: bool) -> dict:
    """Validate and default a flat-table row in place (raises ApiError with a user-facing message)."""
    t = TABLES[sheet]
    for k, opts in t["choices"].items():
        v = str(row.get(k, "")).strip()
        if k in YES_NO:
            row[k] = "Yes" if truthy(v) else "No"
            continue
        if not v:
            continue
        match = _choice(v, opts)
        if match:
            row[k] = match
        elif k in STRICT_CHOICES:
            raise ApiError(400, f"{default_label(k)} must be one of {', '.join(opts)}.")
    for k in t["dates"]:
        if row.get(k):
            row[k] = norm_date(row[k], k)
    if sheet == "Asset_Master":
        v = norm_key(row.get("Platform_Type", ""))
        if v and row["Platform_Type"] not in t["choices"]["Platform_Type"]:
            if any(w in v for w in ("virt", "vmware", "hyperv", "kvm", "nutanix", "vm", "esx", "ovm", "ldom")):
                row["Platform_Type"] = "Virtualization"
            elif any(w in v for w in ("cloud", "aws", "azure", "gcp", "oci", "ec2")):
                row["Platform_Type"] = "Cloud"
            elif any(w in v for w in ("physical", "baremetal", "metal")):
                row["Platform_Type"] = "Physical"
    elif sheet == "Incidents":
        missing = [k for k in ("Action", "Reason", "Technical_Tracking") if not str(row.get(k, "")).strip()]
        if missing:
            raise ApiError(400, "ART standard: every incident needs Action (A), Reason (R) and Technical tracking "
                                "(T). Missing: " + ", ".join(default_label(m) for m in missing) + ".")
        row["Priority"] = row.get("Priority") or "P3"
        row["Status"] = row.get("Status") or "New"
        row["Opened_At"] = norm_datetime(row.get("Opened_At")) if row.get("Opened_At") else now_local()
        if row.get("Resolved_At"):
            row["Resolved_At"] = norm_datetime(row["Resolved_At"])
    elif sheet == "Runbook":
        row["Date"] = row.get("Date") or dt.date.today().isoformat()
        row["Category"] = row.get("Category") or "Information"
        if not (str(row.get("Note", "")).strip() or str(row.get("Solution", "")).strip()
                or str(row.get("Title", "")).strip()):
            raise ApiError(400, "A runbook entry needs a title, a note or a solution.")
    elif sheet == "Shift_Transitions":
        row["Shift"] = parse_shift(row["Shift"])
    elif sheet == "Bulletins":
        row["Severity"] = row.get("Severity") or "Info"
        row["Start"] = norm_datetime(row.get("Start")) if row.get("Start") else now_local()
        if row.get("Expiry"):
            row["Expiry"] = norm_datetime(row["Expiry"], "Expiry")
            if row["Expiry"] < row["Start"]:
                raise ApiError(400, "Expiry must be after the start.")
        if creating and not row.get("Posted_By"):
            row["Posted_By"] = user
    for k in t["required"]:
        if not str(row.get(k, "")).strip():
            raise ApiError(400, f"{default_label(k)} is required.")
    return row


def stamp(row: dict, user: str, previous: dict | None):
    ts = now_iso()
    if previous is None:
        row["Version"], row["Created_By"], row["Created_At"] = "1", user, ts
    else:
        try:
            row["Version"] = str(int(previous.get("Version") or 0) + 1)
        except ValueError:
            row["Version"] = "1"
        row["Created_By"] = previous.get("Created_By", "")
        row["Created_At"] = previous.get("Created_At", "")
    row["Updated_By"], row["Updated_At"] = user, ts


def dup_signatures(sheet: str, row: dict) -> list[str]:
    sigs = []
    for keyset in TABLES[sheet]["dup_keys"]:
        vals = [norm_key(row.get(k, "")) for k in keyset]
        if all(vals):
            sigs.append("|".join(keyset) + "=" + "|".join(vals))
    return sigs


def row_label(sheet: str, row: dict) -> str:
    t = TABLES[sheet]
    main = row.get(t["label_key"], "") or row.get(ID_COL, "")
    extra = row.get("Customer", "")
    return f"{main} ({extra})" if extra else str(main)


# --------------------------------------------------------------------------------------
# Workbook IO
# --------------------------------------------------------------------------------------
_BANNER_FILL = PatternFill("solid", fgColor="0C1A2B")
_HEADER_FILL = PatternFill("solid", fgColor="24405F")


def detect_header_row(rows: list[list[str]], known_headers=None, scan: int = 25):
    """Locate the real header row, skipping decorative banner/title rows.

    1. The first row containing two or more known column names wins.
    2. Otherwise the first row with >= 2 filled cells and at least half the filled
       width of the widest row in the scanned window.
    """
    window = rows[:scan]
    counts = [sum(1 for v in r if str(v).strip()) for r in window]
    if not any(counts):
        return None
    known = {norm_key(h) for h in (known_headers or [])} | set(HEADER_ALIASES)
    for i, row in enumerate(window):
        if sum(1 for v in row if str(v).strip() and norm_key(v) in known) >= 2:
            return i
    widest = max(counts)
    for i, c in enumerate(counts):
        if c >= 2 and c >= widest * 0.5:
            return i
    return next(i for i, c in enumerate(counts) if c)


def all_known_headers() -> list[str]:
    out = list(AUDIT_HEADERS) + PLAN_IMPORT_COLS + ["Sheet", "Key", "Label"]
    for s in SHEET_ORDER:
        out += schema_keys(s) + [default_label(k) for k in schema_keys(s)]
    return out


def grid_to_section(values: list[list[str]], known=None) -> dict:
    idx = detect_header_row(values, known or all_known_headers())
    if idx is None:
        return {"headers": [], "rows": [], "header_row": None}
    raw_headers = list(values[idx])
    while raw_headers and not str(raw_headers[-1]).strip():
        raw_headers.pop()
    headers, seen = [], set()
    for i, h in enumerate(raw_headers):
        h = str(h).strip() or f"Column_{i + 1}"
        base, n = h, 2
        while h.lower() in seen:
            h, n = f"{base}_{n}", n + 1
        seen.add(h.lower())
        headers.append(h)
    rows = []
    for raw in values[idx + 1:]:
        cells = ([cell_str(c) for c in raw] + [""] * len(headers))[:len(headers)]
        if not any(c.strip() for c in cells):
            continue
        rows.append(dict(zip(headers, cells)))
    return {"headers": headers, "rows": rows, "header_row": idx + 1}


def read_workbook(data: bytes) -> dict:
    """Parse every sheet of an .xlsx into {sheet name: {headers, rows, header_row}}."""
    if not data or data[:2] != b"PK":
        raise ApiError(400, "The file is not an .xlsx workbook.")
    try:
        wb = openpyxl.load_workbook(io.BytesIO(data), read_only=True, data_only=True)
    except Exception as exc:
        raise ApiError(400, f"The workbook could not be read: {exc}")
    out = {}
    try:
        for ws in wb.worksheets:
            values = []
            for i, row in enumerate(ws.iter_rows(values_only=True)):
                if i > 200000:
                    break
                values.append([cell_str(v) for v in row])
            out[ws.title] = grid_to_section(values)
    finally:
        wb.close()
    return out


def _put(cell, value):
    text = clean_text(value)
    if text.isdigit() and len(text) < 10 and not (len(text) > 1 and text.startswith("0")):
        cell.value = int(text)
        return
    cell.value = text
    if text.startswith("="):
        cell.data_type = "s"  # never let user text become an Excel formula


def write_sheet(wb, title: str, keys: list[str], labels: list[str], rows: list[dict], banner: str | None):
    ws = wb.create_sheet(title)
    ncols = max(len(keys), 1)
    hdr_row = 1
    if banner:
        ws.cell(row=1, column=1, value=banner).font = Font(name="Arial", size=12, bold=True, color="FFFFFF")
        for c in range(1, ncols + 1):
            ws.cell(row=1, column=c).fill = _BANNER_FILL
        if ncols > 1:
            ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncols)
        ws.row_dimensions[1].height = 22
        hdr_row = 2
    for c, label in enumerate(labels, start=1):
        cell = ws.cell(row=hdr_row, column=c, value=label)
        cell.font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
        cell.fill = _HEADER_FILL
    body_font = Font(name="Arial", size=10)
    widths = [len(str(lb)) + 2 for lb in labels]
    for r_i, row in enumerate(rows, start=hdr_row + 1):
        for c, key in enumerate(keys, start=1):
            cell = ws.cell(row=r_i, column=c)
            _put(cell, row.get(key, ""))
            cell.font = body_font
            if key in WRAP_KEYS:
                cell.alignment = Alignment(wrap_text=True, vertical="top")
            longest = max((len(x) for x in str(row.get(key, "")).splitlines()), default=0)
            widths[c - 1] = max(widths[c - 1], min(longest + 2, 60))
    for c, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(c)].width = max(9, min(w, 60))
    ws.freeze_panes = ws.cell(row=hdr_row + 1, column=1)
    if keys:
        ws.auto_filter.ref = f"A{hdr_row}:{get_column_letter(len(keys))}{max(hdr_row, len(rows) + hdr_row)}"
    return ws


def write_tables_workbook(tables: dict, audit: list[dict], extra_sheets=None) -> bytes:
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    stamp_txt = dt.datetime.now().strftime("%Y-%m-%d %H:%M")
    for sheet in SHEET_ORDER:
        t = tables.get(sheet) or empty_table(sheet)
        keys = [c["key"] for c in t["columns"]]
        labels = [c["label"] for c in t["columns"]]
        write_sheet(wb, sheet, keys, labels, t["rows"], f"{APP_NAME} - {TABLES[sheet]['title']} - saved {stamp_txt}")
    for title, keys, labels, rows, banner in extra_sheets or []:
        write_sheet(wb, title, keys, labels, rows, banner)
    write_sheet(wb, AUDIT_SHEET, AUDIT_HEADERS, AUDIT_HEADERS, audit, f"{APP_NAME} - Audit trail - {stamp_txt}")
    cfg_rows = []
    for sheet in SHEET_ORDER:
        for i, c in enumerate((tables.get(sheet) or empty_table(sheet))["columns"]):
            cfg_rows.append({"Sheet": sheet, "Key": c["key"], "Label": c["label"], "Core": "Yes" if c["core"] else "No",
                             "Hidden": "Yes" if c["hidden"] else "No", "Order": str(i + 1)})
    ws = write_sheet(wb, COLUMNS_SHEET, ["Sheet", "Key", "Label", "Core", "Hidden", "Order"],
                     ["Sheet", "Key", "Label", "Core", "Hidden", "Order"], cfg_rows, None)
    ws.sheet_state = "hidden"
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def section_to_table(sheet: str, section: dict, columns: list[dict]):
    """Map a parsed sheet section onto a table with the given column config. Returns (table, mapped rows)."""
    cols = [dict(c) for c in columns]
    keys, new_cols = map_headers(section["headers"], cols, sheet)
    cols += new_cols
    rows = []
    for raw in section["rows"]:
        row = {c["key"]: "" for c in cols}
        for h, key in zip(section["headers"], keys):
            if key and raw.get(h, "") != "" and (raw.get(h) or not row[key]):
                row[key] = clean_text(raw[h])
        rows.append(row)
    return {"columns": cols, "rows": rows}, new_cols


def tables_from_workbook(data: bytes):
    """Load the central workbook: honours the hidden _Columns config so renamed headers round-trip."""
    raw = read_workbook(data)
    cfg: dict[str, list] = {}
    for name, sec in raw.items():
        if canonical_sheet(name) == COLUMNS_SHEET:
            for r in sec["rows"]:
                s = r.get("Sheet", "")
                if s in TABLES and r.get("Key"):
                    cfg.setdefault(s, []).append({"key": r["Key"], "label": r.get("Label") or r["Key"],
                                                  "hidden": truthy(r.get("Hidden", "No")),
                                                  "_order": int(r.get("Order") or 0) if str(r.get("Order", "")).isdigit() else 0})
    tables, audit = default_tables(), []
    for name, sec in raw.items():
        canon = canonical_sheet(name)
        if canon == AUDIT_SHEET:
            for r in sec["rows"]:
                audit.append({h: r.get(h, "") for h in AUDIT_HEADERS})
            continue
        if canon not in TABLES:
            continue
        stored = sorted(cfg.get(canon, []), key=lambda c: c["_order"])
        table, _ = section_to_table(canon, sec, ensure_columns(canon, stored or None))
        table["columns"] = ensure_columns(canon, table["columns"])
        seen = set()
        for r in table["rows"]:
            for c in table["columns"]:
                r.setdefault(c["key"], "")
            if not r.get(ID_COL) or r[ID_COL] in seen:
                r[ID_COL] = new_id("R")
            seen.add(r[ID_COL])
        tables[canon] = table
    return tables, audit


# --------------------------------------------------------------------------------------
# Sample data for a first start
# --------------------------------------------------------------------------------------
def seed_tables(user: str = "system") -> dict:
    t = default_tables()
    today = dt.date.today()

    def add(sheet, **vals):
        row = blank_row(t[sheet])
        row.update(vals)
        prepare_row(sheet, row, user, True)
        row[ID_COL] = new_id("R")
        stamp(row, user, None)
        t[sheet]["rows"].append(row)

    add("Asset_Master", Customer="Acme Corp", Serial_No="CZJ8340XQ1", Device="acme-orarac01",
        Hostname_FQDN="acme-orarac01.acme.local", IP_Address="10.20.1.11", Environment="Production",
        Platform_Type="Physical", Virtualization_Cloud="", OS="Oracle Linux", OS_Version="8.9",
        Application="Oracle RAC 19c / ASM", Role="RAC node 1", CPU="2 x 24 cores", Memory="512 GB",
        Root_Disk="2 x 480 GB RAID1", Data_Disk="ASM DATA 8 TB, FRA 4 TB", Data_Centre_Rack="DC1 / R12-U07",
        Criticality="Critical", Patch_Group="RAC-Monthly", Maintenance_Window="Sun 22:00-02:00",
        Backup_Policy="RMAN daily incremental", Monitoring="Enabled", Warranty_End="2027-03-31",
        OS_EOL="2029-07-01", Upstream_Dependencies="SAN VNX-02, acme-orarac02",
        Downstream_Dependencies="ERP app tier", Special_Care="Patch node 2 first. Never stop CRS on both nodes "
                                                              "in the same window.", Owner="DB Platform",
        Status="Active")
    add("Asset_Master", Customer="Acme Corp", Serial_No="VMW-42-1a7f", Device="acme-boomi01",
        IP_Address="10.20.3.40", Environment="Production", Platform_Type="Virtualization",
        Virtualization_Cloud="VMware vSphere 8 / cluster PROD-01", OS="RHEL", OS_Version="9.3",
        Application="Boomi Molecule", CPU="8 vCPU", Memory="32 GB", Root_Disk="80 GB", Data_Disk="200 GB /boomi",
        Criticality="High", Special_Care="Drain the molecule (stop listeners) before reboot.", Status="Active")
    add("Asset_Master", Customer="Globex", Serial_No="i-0a91c3f2e7", Device="gbx-sql01", IP_Address="172.16.5.21",
        Environment="Production", Platform_Type="Cloud", Virtualization_Cloud="Azure / sub-prod / westeurope",
        OS="Windows Server", OS_Version="2022", Application="MSSQL 2022 AlwaysOn", CPU="16 vCPU", Memory="128 GB",
        Root_Disk="127 GB", Data_Disk="2 x 1 TB Premium SSD", Criticality="Critical", Status="Active")
    add("Runbook", Customer="Acme Corp", Environment="Production", Asset="acme-orarac01",
        Date=(today - dt.timedelta(days=3)).isoformat(), Category="Known issue",
        Title="ASM DATA diskgroup fills when RMAN credential expires",
        Note="Archive log backups fail silently after the backup host credential rotation.",
        Solution="Re-register the credential on the backup host, run a manual archivelog backup, then purge.",
        Commands="asmcmd lsdg\nrman target / cmdfile=arch_backup.rman", Ticket_Ref="INC-104233",
        Engineer="On-call DBA", Tags="asm, rman", Known_Issue="Yes")
    add("Runbook", Customer="Acme Corp", Environment="Production", Asset="acme-boomi01",
        Date=(today - dt.timedelta(days=10)).isoformat(), Category="Maintenance", Title="Molecule node drained for patching",
        Note="Listeners stopped via Atom Management before OS patching.", Solution="Rejoined cluster after reboot.",
        Commands="systemctl stop atom\nsystemctl start atom && systemctl is-active atom", Engineer="Integration L2")
    add("Incidents", Ticket_ID="INC-104233", Customer="Acme Corp", Asset="acme-orarac01", Priority="P1",
        Status="In Progress", Summary="ASM diskgroup DATA at 96%; archiver at risk",
        Action="Purged archived logs older than 2 days after confirming backup; requested 200 GB LUN.",
        Reason="RMAN archive-log backup failed for 3 nights (expired credential on backup host).",
        Technical_Tracking="asmcmd lsdg\nrman: delete noprompt archivelog until time 'sysdate-2' backed up 1 times\n"
                           "asmcmd lsdg  # DATA now 71%", Engineer="On-call DBA")
    add("Shift_Transitions", Date=today.isoformat(), Shift="Morning", Region="EMEA Morning",
        Outgoing_Engineer="APAC Night team", Incoming_Engineer="EMEA Morning team",
        P1_P2_In_Flight="INC-104233 (Acme ASM space)", Planned_Changes="CHG-2041 Acme RAC patching on Sunday",
        High_Risk_Assets="acme-orarac01", Acknowledged="No")
    add("Bulletins", Title="Acme RAC patching this Sunday 22:00", Severity="Advisory", Customers="Acme Corp",
        Body="CHG-2041: node 2 first, then node 1. Expect ERP read-only for 20 minutes during CRS restart.",
        Start=now_local(), Expiry=(dt.datetime.now() + dt.timedelta(days=5)).strftime("%Y-%m-%d %H:%M"),
        Pinned="Yes")
    add("Bulletins", Title="OpenSSH security advisory: patch Linux bastions", Severity="Critical",
        Customers="All", Body="Apply vendor OpenSSH updates to all bastion hosts within 72 hours.",
        Start=now_local(), Expiry=(dt.datetime.now() + dt.timedelta(days=3)).strftime("%Y-%m-%d %H:%M"))
    return t


# --------------------------------------------------------------------------------------
# Execution plans: Year > Month > Date > Customer > Task > Phase (1, 2 ...) > Step (1.1, 1.2.1 ...)
# --------------------------------------------------------------------------------------
PHASE_PRESETS = ["Pre-execution", "During outage", "Post execution", "Pre-live", "During cut-over", "Post-live"]
BUILTIN_PLAN_TEMPLATES = {
    "blank": ("Blank (no phases)", []),
    "change": ("Standard change", ["Pre-execution", "During outage", "Post execution"]),
    "golive": ("Go-live / cut-over", ["Pre-live", "During cut-over", "Post-live"]),
    "full": ("All six phases", PHASE_PRESETS),
}
TASK_STATUSES = ["Draft", "Scheduled", "In Progress", "Completed", "Rolled Back", "Cancelled"]
STEP_STATUSES = ["Pending", "In Progress", "Done", "Skipped", "Failed"]
STEP_CORE = [("title", "Task / step"), ("owner", "Owner"), ("planned_start", "Date & time"),
             ("approx_min", "Approx (min)"), ("actual_start", "Actual start"), ("actual_end", "Actual end"),
             ("status", "Status"), ("notes", "Status notes"), ("command", "Command"), ("rollback", "Rollback"),
             ("expected", "Expected output")]
STEP_CORE_KEYS = [k for k, _ in STEP_CORE]
TASK_FIELDS = ["date", "customer", "name", "ticket_id", "change_ref", "window_start", "window_end", "approver",
               "status", "notes", "assets"]
MAX_LEVEL = 3  # 1.1 / 1.1.1 / 1.1.1.1 / 1.1.1.1.1
PLAN_IMPORT_COLS = ["Date", "Customer", "Task_Name", "Ticket_ID", "Change_Ref", "Window_Start", "Window_End",
                    "Approver", "Task_Status", "Phase", "Checkpoint", "Step_No", "Step", "Owner", "Planned_Start",
                    "Approx_Min", "Actual_Start", "Actual_End", "Status", "Notes", "Command", "Rollback", "Expected"]
PLAN_EXPORT_COLS = ["Year", "Month", "Day"] + PLAN_IMPORT_COLS[:9] + ["Phase_No"] + PLAN_IMPORT_COLS[9:15] + \
                   ["Planned_Calculated"] + PLAN_IMPORT_COLS[15:18] + ["Actual_Min"] + PLAN_IMPORT_COLS[18:]
PLAN_HEADER_ALIASES = {
    "client": "Customer", "customername": "Customer", "taskname": "Task_Name", "task": "Task_Name",
    "changetask": "Task_Name", "ticket": "Ticket_ID", "ticketno": "Ticket_ID", "ticketnumber": "Ticket_ID",
    "change": "Change_Ref", "changeid": "Change_Ref", "crq": "Change_Ref", "maintenancestart": "Window_Start",
    "maintenanceend": "Window_End", "taskstatus": "Task_Status", "header": "Phase", "phaseheader": "Phase",
    "stage": "Phase", "section": "Phase", "gonogo": "Checkpoint", "no": "Step_No", "number": "Step_No",
    "seq": "Step_No", "sno": "Step_No", "stepnumber": "Step_No", "wbs": "Step_No", "stepdescription": "Step",
    "description": "Step", "activity": "Step", "taskorstep": "Step", "steps": "Step", "approx": "Approx_Min",
    "approxtime": "Approx_Min", "approxmins": "Approx_Min", "estimate": "Approx_Min", "duration": "Approx_Min",
    "durationmin": "Approx_Min", "datetime": "Planned_Start", "dateandtime": "Planned_Start",
    "plannedtime": "Planned_Start", "starttime": "Planned_Start", "commands": "Command", "cmd": "Command",
    "commandtoexecute": "Command", "statusnotes": "Notes", "remarks": "Notes", "comments": "Notes",
    "rollbackcommand": "Rollback", "backout": "Rollback", "expectedoutput": "Expected", "verification": "Expected",
    "validation": "Expected", "plandate": "Date", "executiondate": "Date", "stepstatus": "Status",
}


def parse_minutes(value) -> int | str:
    s = str(value if value is not None else "").strip().lower()
    if not s:
        return ""
    if re.fullmatch(r"\d+", s):
        return int(s)
    if re.fullmatch(r"\d+(\.\d+)?", s):
        return int(round(float(s)))
    m = re.fullmatch(r"(\d+):(\d{1,2})", s)
    if m:
        return int(m.group(1)) * 60 + int(m.group(2))
    m = re.fullmatch(r"(?:(\d+(?:\.\d+)?)\s*h(?:rs?|ours?)?)?\s*(?:(\d+)\s*m(?:in(?:ute)?s?)?)?", s)
    if m and (m.group(1) or m.group(2)):
        return int(round(float(m.group(1) or 0) * 60)) + int(m.group(2) or 0)
    raise ApiError(400, f"Approx time '{value}' is not a duration. Use minutes (45), 1h30m or 1:30.")


def default_step_columns() -> list[dict]:
    return [{"key": k, "label": lb, "core": True, "hidden": False} for k, lb in STEP_CORE]


def ensure_step_columns(cols) -> list[dict]:
    out, seen = [], set()
    labels = dict(STEP_CORE)
    for c in cols or []:
        if not isinstance(c, dict):
            continue
        key = str(c.get("key") or "")
        if not key or key in seen:
            continue
        if key not in labels and not key.startswith("x_"):
            continue
        seen.add(key)
        out.append({"key": key, "label": clean_text(c.get("label") or labels.get(key, key), 80).strip() or key,
                    "core": key in labels, "hidden": bool(c.get("hidden"))})
    for k, lb in STEP_CORE:
        if k not in seen:
            out.append({"key": k, "label": lb, "core": True, "hidden": False})
    return out


def new_step(level: int = 0) -> dict:
    return {"id": new_id("st_"), "level": level, "title": "", "owner": "", "planned_start": "", "approx_min": "",
            "actual_start": "", "actual_end": "", "status": "Pending", "notes": "", "command": "", "rollback": "",
            "expected": "", "custom": {}}


def new_phase(name: str, checkpoint: bool = False) -> dict:
    return {"id": new_id("ph_"), "name": clean_text(name, 120).strip() or "Phase", "checkpoint": bool(checkpoint),
            "signed_off_by": "", "signed_off_at": "", "steps": []}


def normalize_levels(steps: list[dict]):
    prev = -1
    for st in steps:
        try:
            lvl = int(st.get("level") or 0)
        except (TypeError, ValueError):
            lvl = 0
        lvl = max(0, min(lvl, MAX_LEVEL, prev + 1))
        st["level"] = lvl
        prev = lvl


def normalize_task(obj: dict, keep_ids: bool = True) -> dict:
    """Validate/repair a task loaded from disk or imported from JSON."""
    if not isinstance(obj, dict):
        raise ApiError(400, "Each plan task must be an object.")
    t = {"id": str(obj.get("id")) if keep_ids and obj.get("id") else new_id("t_")}
    for f in TASK_FIELDS:
        t[f] = clean_text(obj.get(f, ""), 4000)
    t["date"] = norm_date(t["date"] or dt.date.today().isoformat())
    for f in ("window_start", "window_end"):
        t[f] = norm_datetime(t[f]) if t[f] else ""
    t["status"] = _choice(t["status"], TASK_STATUSES) or "Draft"
    for f in ("created_by", "created_at", "updated_by", "updated_at"):
        t[f] = clean_text(obj.get(f, ""), 120)
    t["step_columns"] = ensure_step_columns(obj.get("step_columns"))
    t["owner_team"] = clean_text(obj.get("owner_team", ""), 80)  # isolation: the team that owns the plan
    t["shared_with"] = [clean_text(x, 80) for x in (obj.get("shared_with") or []) if str(x).strip()][:10]
    custom_keys = {c["key"] for c in t["step_columns"] if not c["core"]}
    t["phases"] = []
    for ph in obj.get("phases") or []:
        if not isinstance(ph, dict):
            continue
        p = new_phase(ph.get("name", "Phase"), ph.get("checkpoint", False))
        if keep_ids and ph.get("id"):
            p["id"] = str(ph["id"])
        p["signed_off_by"] = clean_text(ph.get("signed_off_by", ""), 120)
        p["signed_off_at"] = clean_text(ph.get("signed_off_at", ""), 40)
        p["team"] = clean_text(ph.get("team", ""), 80)  # a phase given to another team (partial sharing)
        for s in ph.get("steps") or []:
            if not isinstance(s, dict):
                continue
            st = new_step()
            if keep_ids and s.get("id"):
                st["id"] = str(s["id"])
            st["level"] = s.get("level", 0)
            for k in ("title", "owner", "notes", "command", "rollback", "expected"):
                st[k] = clean_text(s.get(k, ""))
            for k in ("planned_start", "actual_start", "actual_end"):
                try:
                    st[k] = norm_datetime(s.get(k, "")) if s.get(k) else ""
                except ApiError:
                    st[k] = ""
            try:
                st["approx_min"] = parse_minutes(s.get("approx_min", ""))
            except ApiError:
                st["approx_min"] = ""
            st["status"] = _choice(str(s.get("status", "")), STEP_STATUSES) or "Pending"
            st["custom"] = {k: clean_text(v) for k, v in (s.get("custom") or {}).items() if k in custom_keys}
            st["team"] = clean_text(s.get("team", ""), 80)  # a single step given to another team
            p["steps"].append(st)
        normalize_levels(p["steps"])
        t["phases"].append(p)
    return t


def empty_plans() -> dict:
    return {"schema_version": 1, "tasks": [], "templates": [], "audit": []}


def normalize_plans(doc) -> dict:
    if not isinstance(doc, dict) or not isinstance(doc.get("tasks", []), list):
        raise ApiError(400, "Execution plans JSON must be an object with a 'tasks' list.")
    out = empty_plans()
    out["tasks"] = [normalize_task(t) for t in doc.get("tasks", [])]
    for tpl in doc.get("templates", []) or []:
        if isinstance(tpl, dict) and tpl.get("name"):
            body = normalize_task({**tpl, "date": dt.date.today().isoformat()})
            out["templates"].append({"id": str(tpl.get("id") or new_id("tpl_")), "name": clean_text(tpl["name"], 120),
                                     "created_by": clean_text(tpl.get("created_by", ""), 120),
                                     "created_at": clean_text(tpl.get("created_at", ""), 40),
                                     "phases": body["phases"], "step_columns": body["step_columns"]})
    audit = doc.get("audit", [])
    out["audit"] = [a for a in audit if isinstance(a, dict)][-AUDIT_CAP:] if isinstance(audit, list) else []
    return out


def _fmt(d: dt.datetime | None) -> str:
    return d.strftime("%Y-%m-%d %H:%M") if d else ""


def compute_task(task: dict) -> dict:
    """Return a copy with numbering (1, 1.1, 1.2.1), chained planned times, actual durations and totals."""
    t = copy.deepcopy(task)
    start = parse_dt(t.get("window_start")) or parse_dt(f"{t['date']} 00:00")
    running, total, done, count = start, 0, 0, 0
    for pi, ph in enumerate(t["phases"]):
        ph["num"] = str(pi + 1)
        counters = [0] * (MAX_LEVEL + 1)
        ph_total = 0
        for st in ph["steps"]:
            lvl = st["level"]
            counters[lvl] += 1
            for j in range(lvl + 1, MAX_LEVEL + 1):
                counters[j] = 0
            st["num"] = ph["num"] + "." + ".".join(str(c) for c in counters[:lvl + 1])
            override = parse_dt(st.get("planned_start"))
            planned = override or running
            mins = st["approx_min"] if isinstance(st.get("approx_min"), int) else 0
            st["planned"] = _fmt(planned)
            st["planned_inherited"] = override is None
            end = planned + dt.timedelta(minutes=mins) if planned else None
            st["planned_end"] = _fmt(end)
            running = end
            a1, a2 = parse_dt(st.get("actual_start")), parse_dt(st.get("actual_end"))
            st["actual_min"] = int((a2 - a1).total_seconds() // 60) if a1 and a2 and a2 >= a1 else ""
            st["variance_min"] = (st["actual_min"] - mins) if st["actual_min"] != "" and mins else ""
            total += mins
            ph_total += mins
            count += 1
            done += st["status"] in ("Done", "Skipped")
        ph["approx_total"] = ph_total
    wend = parse_dt(t.get("window_end"))
    overrun = int((running - wend).total_seconds() // 60) if running and wend and running > wend else 0
    t["computed"] = {"start": _fmt(start), "projected_end": _fmt(running), "approx_total": total, "steps": count,
                     "done": done, "progress": round(100 * done / count) if count else 0, "overrun_min": overrun}
    return t


def task_label(t: dict) -> str:
    ref = f" [{t['ticket_id']}]" if t.get("ticket_id") else ""
    return f"{t.get('date', '')} {t.get('customer', '')} - {t.get('name', '')}{ref}"


def task_dup_sig(t: dict) -> str:
    third = t.get("ticket_id") or t.get("name") or ""
    return "|".join(norm_key(x) for x in (t.get("date", ""), t.get("customer", ""), third))


def reset_actuals(phases: list[dict], new_ids: bool = True):
    for ph in phases:
        if new_ids:
            ph["id"] = new_id("ph_")
        ph["signed_off_by"] = ph["signed_off_at"] = ""
        for st in ph["steps"]:
            if new_ids:
                st["id"] = new_id("st_")
            st["actual_start"] = st["actual_end"] = ""
            st["status"] = "Pending"


def _find_task(doc: dict, tid) -> dict:
    for t in doc["tasks"]:
        if t["id"] == tid:
            return t
    raise ApiError(404, "That task no longer exists.")


def _find_phase(task: dict, pid) -> dict:
    for ph in task["phases"]:
        if ph["id"] == pid:
            return ph
    raise ApiError(404, "That phase no longer exists.")


def _step_index(phase: dict, sid) -> int:
    for i, st in enumerate(phase["steps"]):
        if st["id"] == sid:
            return i
    raise ApiError(404, "That step no longer exists.")


def _block_end(steps: list[dict], i: int) -> int:
    lvl, j = steps[i]["level"], i + 1
    while j < len(steps) and steps[j]["level"] > lvl:
        j += 1
    return j


def apply_task_fields(task: dict, fields: dict):
    if not isinstance(fields, dict):
        raise ApiError(400, "Task fields must be an object.")
    for f in TASK_FIELDS:
        if f not in fields:
            continue
        v = clean_text(fields[f], 4000).strip()
        if f == "date":
            v = norm_date(v, "Date")
            if not v:
                raise ApiError(400, "Date is required.")
        elif f in ("window_start", "window_end"):
            v = norm_datetime(v, "Maintenance window") if v else ""
        elif f == "status":
            v = _choice(v, TASK_STATUSES) or ""
            if not v:
                raise ApiError(400, f"Task status must be one of {', '.join(TASK_STATUSES)}.")
        task[f] = v
    for f in ("customer", "name"):
        if not str(task.get(f, "")).strip():
            raise ApiError(400, ("Customer" if f == "customer" else "Task name") + " is required.")
    if task.get("window_start") and task.get("window_end") and task["window_end"] < task["window_start"]:
        raise ApiError(400, "The maintenance window must end after it starts.")


def apply_step_fields(task: dict, step: dict, fields: dict):
    if not isinstance(fields, dict):
        raise ApiError(400, "Step fields must be an object.")
    by_label = {norm_key(c["label"]): c["key"] for c in task["step_columns"]}
    custom = {c["key"] for c in task["step_columns"] if not c["core"]}
    for k, v in fields.items():
        key = k if (k in STEP_CORE_KEYS or k in custom) else by_label.get(norm_key(k))
        if not key:
            continue
        v = clean_text(v)
        if key in custom:
            step["custom"][key] = v
        elif key == "approx_min":
            step[key] = parse_minutes(v)
        elif key in ("planned_start", "actual_start", "actual_end"):
            step[key] = norm_datetime(v, dict(STEP_CORE)[key]) if v.strip() else ""
        elif key == "status":
            s = _choice(v, STEP_STATUSES)
            if not s:
                raise ApiError(400, f"Step status must be one of {', '.join(STEP_STATUSES)}.")
            step[key] = s
        else:
            step[key] = v
    if step.get("actual_start") and step.get("actual_end") and step["actual_end"] < step["actual_start"]:
        raise ApiError(400, "Actual end must be after actual start.")


def _template_phases(doc: dict, template) -> tuple[list[dict], list[dict] | None]:
    template = template or "change"
    if template in BUILTIN_PLAN_TEMPLATES:
        names = BUILTIN_PLAN_TEMPLATES[template][1]
        return [new_phase(n, checkpoint=(i == 0 and len(names) > 1)) for i, n in enumerate(names)], None
    for tpl in doc["templates"]:
        if tpl["id"] == template:
            phases = copy.deepcopy(tpl["phases"])
            reset_actuals(phases)
            return phases, copy.deepcopy(tpl["step_columns"])
    raise ApiError(404, "That plan template no longer exists.")


def plan_op(doc: dict, op: str, p: dict, user: str, ctx: dict):
    """Apply one execution-plan operation to the plans document (in place)."""
    hist = ctx.setdefault("history", [])
    ts = now_iso()

    def touched(task, before, action):
        task["updated_by"], task["updated_at"] = user, ts
        hist.append((f"plan/{task['id']}", action, before, copy.deepcopy(task)))
        ctx["target"] = task_label(task)

    # ---- tasks -------------------------------------------------------------------------
    if op == "task.create":
        fields = p.get("task") or {}
        task = {"id": new_id("t_"), **{f: "" for f in TASK_FIELDS}, "status": "Draft", "created_by": user,
                "created_at": ts, "updated_by": user, "updated_at": ts, "step_columns": default_step_columns(),
                "phases": []}
        apply_task_fields(task, {"date": dt.date.today().isoformat(), **fields})
        phases, cols = _template_phases(doc, p.get("template"))
        task["phases"] = phases
        if cols:
            task["step_columns"] = cols
        if not p.get("allow_duplicate"):
            sig = task_dup_sig(task)
            dup = next((t for t in doc["tasks"] if task_dup_sig(t) == sig), None)
            if dup:
                raise ApiError(409, f"A task already exists for this date, customer and ticket/name: "
                                    f"{task_label(dup)}. Open it, or create anyway.")
        doc["tasks"].append(task)
        hist.append((f"plan/{task['id']}", "Create", None, copy.deepcopy(task)))
        ctx["target"] = task_label(task)
        return {"task_id": task["id"]}

    if op == "template.delete":
        before = len(doc["templates"])
        doc["templates"] = [t for t in doc["templates"] if t["id"] != p.get("template_id")]
        if len(doc["templates"]) == before:
            raise ApiError(404, "That template no longer exists.")
        ctx["target"] = f"template {p.get('template_id')}"
        return {"deleted": p.get("template_id")}

    task = _find_task(doc, p.get("task_id"))
    before = copy.deepcopy(task)

    if op == "task.share":  # owner team and teams that see the whole plan
        if "owner_team" in p:
            task["owner_team"] = clean_text(p.get("owner_team"), 80)
        if "shared_with" in p:
            task["shared_with"] = [clean_text(x, 80) for x in (p.get("shared_with") or []) if str(x).strip()][:10]
        touched(task, before, "Change sharing")
        return {"task_id": task["id"]}
    if op == "task.update":
        apply_task_fields(task, p.get("task") or {})
        touched(task, before, "Update task")
        return {"task_id": task["id"]}
    if op == "task.delete":
        doc["tasks"].remove(task)
        hist.append((f"plan/{task['id']}", "Delete", before, None))
        ctx["target"] = task_label(before)
        return {"deleted": task["id"]}
    if op == "task.clone":
        clone = copy.deepcopy(task)
        clone.update({"id": new_id("t_"), "status": "Draft", "created_by": user, "created_at": ts,
                      "updated_by": user, "updated_at": ts})
        overrides = {k: v for k, v in (p.get("task") or {}).items() if k in TASK_FIELDS}
        if overrides.get("date") and "window_start" not in overrides and task.get("window_start"):
            # move the maintenance window to the new date, keeping times and length
            shift = (dt.date.fromisoformat(norm_date(overrides["date"])) - dt.date.fromisoformat(task["date"]))
            for f in ("window_start", "window_end"):
                d0 = parse_dt(task.get(f))
                clone[f] = _fmt(d0 + shift) if d0 else ""
        apply_task_fields(clone, overrides)
        for ph in clone["phases"]:
            for st in ph["steps"]:
                if st.get("planned_start") and overrides.get("date"):
                    d0 = parse_dt(st["planned_start"])
                    delta = dt.date.fromisoformat(clone["date"]) - dt.date.fromisoformat(task["date"])
                    st["planned_start"] = _fmt(d0 + delta) if d0 else ""
        reset_actuals(clone["phases"])
        if task_dup_sig(clone) == task_dup_sig(task):
            clone["name"] = f"{task['name']} (copy)"
        doc["tasks"].append(clone)
        hist.append((f"plan/{clone['id']}", f"Clone of {task['id']}", None, copy.deepcopy(clone)))
        ctx["target"] = task_label(clone)
        return {"task_id": clone["id"]}
    if op == "task.columns":
        cols = p.get("columns")
        if not isinstance(cols, list) or not cols:
            raise ApiError(400, "Provide the full list of step columns.")
        current = {c["key"]: c for c in task["step_columns"]}
        new_cfg, labels, used = [], set(), set(current)
        for c in cols:
            label = clean_text((c or {}).get("label", ""), 80).strip()
            if not label:
                raise ApiError(400, "Column names cannot be empty.")
            if label.lower() in labels:
                raise ApiError(400, f"Duplicate column name '{label}'.")
            labels.add(label.lower())
            key = c.get("key") if c.get("key") in current else None
            if not key:
                key = custom_key(label, used)
                used.add(key)
            new_cfg.append({"key": key, "label": label, "hidden": bool(c.get("hidden"))})
        task["step_columns"] = ensure_step_columns(new_cfg)
        kept = {c["key"] for c in task["step_columns"] if not c["core"]}
        removed = [c["label"] for k, c in current.items() if not c["core"] and k not in kept]
        for ph in task["phases"]:
            for st in ph["steps"]:
                st["custom"] = {k: v for k, v in st["custom"].items() if k in kept}
        touched(task, before, "Configure step columns")
        return {"removed": removed}
    if op == "template.save":
        name = clean_text(p.get("name", ""), 120).strip() or task["name"]
        phases = copy.deepcopy(task["phases"])
        reset_actuals(phases)
        tpl = {"id": new_id("tpl_"), "name": name, "created_by": user, "created_at": ts, "phases": phases,
               "step_columns": copy.deepcopy(task["step_columns"])}
        doc["templates"].append(tpl)
        ctx["target"] = f"template '{name}'"
        return {"template_id": tpl["id"]}

    # ---- phases ------------------------------------------------------------------------
    if op == "phase.add":
        ph = new_phase(p.get("name") or "New phase", p.get("checkpoint", False))
        after = p.get("after_phase_id")
        if after:
            idx = task["phases"].index(_find_phase(task, after)) + 1
            task["phases"].insert(idx, ph)
        else:
            task["phases"].append(ph)
        touched(task, before, f"Add phase '{ph['name']}'")
        return {"phase_id": ph["id"]}
    phase = _find_phase(task, p.get("phase_id")) if op.startswith(("phase.", "step.")) else None
    if op == "phase.team":  # partial sharing: this phase is run by another team
        phase["team"] = clean_text(p.get("team", ""), 80)
        touched(task, before, f"Give phase '{phase['name']}' to {phase['team'] or 'the owning team'}")
        return {"phase_id": phase["id"]}
    if op == "phase.update":
        if "name" in p:
            name = clean_text(p["name"], 120).strip()
            if not name:
                raise ApiError(400, "Phase name cannot be empty.")
            phase["name"] = name
        if "checkpoint" in p:
            phase["checkpoint"] = bool(p["checkpoint"])
            if not phase["checkpoint"]:
                phase["signed_off_by"] = phase["signed_off_at"] = ""
        touched(task, before, f"Update phase '{phase['name']}'")
        return {"phase_id": phase["id"]}
    if op == "phase.move":
        i = task["phases"].index(phase)
        j = i + (-1 if p.get("dir") in ("up", -1, "-1") else 1)
        if not 0 <= j < len(task["phases"]):
            raise ApiError(400, "The phase cannot move further.")
        task["phases"][i], task["phases"][j] = task["phases"][j], task["phases"][i]
        touched(task, before, f"Move phase '{phase['name']}'")
        return {"phase_id": phase["id"]}
    if op == "phase.clone":
        ph = copy.deepcopy(phase)
        ph["name"] = f"{phase['name']} (copy)"
        reset_actuals([ph])
        task["phases"].insert(task["phases"].index(phase) + 1, ph)
        touched(task, before, f"Clone phase '{phase['name']}'")
        return {"phase_id": ph["id"]}
    if op == "phase.delete":
        task["phases"].remove(phase)
        touched(task, before, f"Delete phase '{phase['name']}'")
        return {"deleted": phase["id"]}
    if op == "phase.signoff":
        if not phase["checkpoint"]:
            raise ApiError(400, "This phase is not a go/no-go checkpoint.")
        if p.get("undo"):
            phase["signed_off_by"] = phase["signed_off_at"] = ""
        else:
            phase["signed_off_by"], phase["signed_off_at"] = user, now_local()
        touched(task, before, ("Withdraw" if p.get("undo") else "Sign off") + f" checkpoint '{phase['name']}'")
        return {"phase_id": phase["id"]}

    # ---- steps -------------------------------------------------------------------------
    steps = phase["steps"]
    if op == "step.add":
        st = new_step()
        after = p.get("after_step_id")
        if after:
            i = _step_index(phase, after)
            if p.get("as_child"):
                st["level"] = min(steps[i]["level"] + 1, MAX_LEVEL)
                pos = i + 1
            else:
                st["level"] = steps[i]["level"]
                pos = _block_end(steps, i)
        else:
            pos = len(steps)
        apply_step_fields(task, st, p.get("step") or {})
        steps.insert(pos, st)
        normalize_levels(steps)
        touched(task, before, "Add step")
        return {"step_id": st["id"]}
    i = _step_index(phase, p.get("step_id"))
    if op == "step.team":
        steps[i]["team"] = clean_text(p.get("team", ""), 80)
        touched(task, before, f"Give step to {steps[i]['team'] or 'the phase team'}")
        return {"step_id": steps[i]["id"]}
    if op == "step.update":
        apply_step_fields(task, steps[i], p.get("step") or {})
        touched(task, before, "Edit step")
        return {"step_id": steps[i]["id"]}
    if op == "step.delete":
        end = _block_end(steps, i)
        del steps[i:end]
        normalize_levels(steps)
        touched(task, before, f"Delete step ({end - i} line(s))")
        return {"deleted": end - i}
    if op == "step.clone":
        end = _block_end(steps, i)
        block = copy.deepcopy(steps[i:end])
        reset_actuals([{"id": "", "steps": block}])
        steps[end:end] = block
        touched(task, before, "Clone step")
        return {"step_id": block[0]["id"]}
    if op == "step.move":
        d = p.get("dir")
        lvl, end = steps[i]["level"], _block_end(steps, i)
        if d == "indent":
            if i == 0 or steps[i - 1]["level"] < lvl or lvl >= MAX_LEVEL:
                raise ApiError(400, "This step cannot be indented further.")
            for st in steps[i:end]:
                st["level"] = min(st["level"] + 1, MAX_LEVEL)
        elif d == "outdent":
            if lvl == 0:
                raise ApiError(400, "This step is already at the top level.")
            for st in steps[i:end]:
                st["level"] -= 1
        elif d in ("up", "down"):
            if d == "up":
                j = i - 1
                while j >= 0 and steps[j]["level"] > lvl:
                    j -= 1
                if j < 0 or steps[j]["level"] < lvl:
                    raise ApiError(400, "The step cannot move further up.")
                steps[j:end] = steps[i:end] + steps[j:i]
            else:
                if end >= len(steps) or steps[end]["level"] < lvl:
                    raise ApiError(400, "The step cannot move further down.")
                nxt_end = _block_end(steps, end)
                steps[i:nxt_end] = steps[end:nxt_end] + steps[i:end]
        else:
            raise ApiError(400, "dir must be up, down, indent or outdent.")
        normalize_levels(steps)
        touched(task, before, f"Move step {d}")
        return {"step_id": p.get("step_id")}
    if op == "step.action":
        action = p.get("action")
        st = steps[i]
        if action in ("start", "done"):
            for prev in task["phases"][:task["phases"].index(phase)]:
                if prev["checkpoint"] and not prev["signed_off_by"]:
                    raise ApiError(409, f"Go/no-go checkpoint '{prev['name']}' has not been signed off yet.")
        now = now_local()
        if action == "start":
            st["actual_start"], st["actual_end"], st["status"] = now, "", "In Progress"
        elif action == "done":
            st["actual_start"] = st["actual_start"] or now
            st["actual_end"], st["status"] = now, "Done"
        elif action == "skip":
            st["status"] = "Skipped"
        elif action == "fail":
            st["actual_end"], st["status"] = now, "Failed"
        elif action == "reset":
            st["actual_start"] = st["actual_end"] = ""
            st["status"] = "Pending"
        else:
            raise ApiError(400, "action must be start, done, skip, fail or reset.")
        all_steps = [s for ph in task["phases"] for s in ph["steps"]]
        if action in ("start", "done") and task["status"] in ("Draft", "Scheduled"):
            task["status"] = "In Progress"
        if all_steps and all(s["status"] in ("Done", "Skipped") for s in all_steps):
            task["status"] = "Completed"
        elif task["status"] == "Completed" and action in ("reset", "start", "fail"):
            task["status"] = "In Progress"
        touched(task, before, f"Step {st.get('title') or st['id']}: {action}")
        return {"step_id": st["id"], "task_status": task["status"]}
    raise ApiError(400, f"Unknown plan operation '{op}'.")


# ---- Excel rows <-> tasks -------------------------------------------------------------
def plans_to_rows(tasks: list[dict]) -> list[dict]:
    rows = []
    for raw in sorted(tasks, key=lambda t: (t["date"], t["customer"], t["name"])):
        t = compute_task(raw)
        d = dt.date.fromisoformat(t["date"])
        base = {"Year": str(d.year), "Month": f"{d.month:02d}", "Date": t["date"], "Day": d.strftime("%A"),
                "Customer": t["customer"], "Task_Name": t["name"], "Ticket_ID": t["ticket_id"],
                "Change_Ref": t["change_ref"], "Window_Start": t["window_start"], "Window_End": t["window_end"],
                "Approver": t["approver"], "Task_Status": t["status"]}
        if not any(ph["steps"] for ph in t["phases"]):
            rows.append(base)
        for ph in t["phases"]:
            for st in ph["steps"] or [None]:
                row = dict(base, Phase_No=ph["num"], Phase=ph["name"], Checkpoint="Yes" if ph["checkpoint"] else "No")
                if st:
                    row.update({"Step_No": st["num"], "Step": st["title"], "Owner": st["owner"],
                                "Planned_Start": st["planned_start"], "Planned_Calculated": st["planned"],
                                "Approx_Min": st["approx_min"], "Actual_Start": st["actual_start"],
                                "Actual_End": st["actual_end"], "Actual_Min": st["actual_min"],
                                "Status": st["status"], "Notes": st["notes"], "Command": st["command"],
                                "Rollback": st["rollback"], "Expected": st["expected"]})
                rows.append(row)
    return rows


def plan_header_key(h: str):
    k = norm_key(h)
    for c in PLAN_IMPORT_COLS + PLAN_EXPORT_COLS:
        if norm_key(c) == k:
            return c
    return PLAN_HEADER_ALIASES.get(k)


def rows_to_tasks(section: dict) -> list[dict]:
    """Group flat plan rows (one per step) into tasks. Returns [{task, errors, rows}]."""
    keys = [plan_header_key(h) for h in section["headers"]]
    groups, order = {}, []
    for n, raw in enumerate(section["rows"], start=1):
        r = {k: str(raw.get(h, "")).strip() for h, k in zip(section["headers"], keys) if k}
        gkey = (r.get("Date", ""), norm_key(r.get("Customer", "")), norm_key(r.get("Ticket_ID") or r.get("Task_Name", "")))
        if gkey not in groups:
            groups[gkey] = {"rows": [], "first": r}
            order.append(gkey)
        groups[gkey]["rows"].append((n, r))
    items = []
    for gkey in order:
        g, errors = groups[gkey], []
        first = g["first"]
        task = {"id": new_id("t_"), **{f: "" for f in TASK_FIELDS}, "status": "Draft",
                "step_columns": default_step_columns(), "phases": []}
        try:
            apply_task_fields(task, {"date": first.get("Date", ""), "customer": first.get("Customer", ""),
                                     "name": first.get("Task_Name") or first.get("Ticket_ID", ""),
                                     "ticket_id": first.get("Ticket_ID", ""), "change_ref": first.get("Change_Ref", ""),
                                     "window_start": first.get("Window_Start", ""),
                                     "window_end": first.get("Window_End", ""), "approver": first.get("Approver", ""),
                                     **({"status": first["Task_Status"]} if first.get("Task_Status") else {})})
            if not first.get("Date"):
                raise ApiError(400, "Date is required.")
        except ApiError as exc:
            errors.append(f"row {g['rows'][0][0]}: {exc.message}")
        phases = {}
        for n, r in g["rows"]:
            pname = r.get("Phase") or "Steps"
            ph = phases.get(norm_key(pname))
            if ph is None:
                ph = new_phase(pname, truthy(r.get("Checkpoint", "")))
                phases[norm_key(pname)] = ph
                task["phases"].append(ph)
            if not (r.get("Step") or r.get("Command")):
                continue
            st = new_step()
            num = r.get("Step_No", "")
            if re.fullmatch(r"\d+(\.\d+)+", num):
                st["level"] = min(num.count(".") - 1, MAX_LEVEL)
            try:
                apply_step_fields(task, st, {"title": r.get("Step", ""), "owner": r.get("Owner", ""),
                                             "planned_start": r.get("Planned_Start", ""),
                                             "approx_min": r.get("Approx_Min", ""),
                                             "actual_start": r.get("Actual_Start", ""),
                                             "actual_end": r.get("Actual_End", ""),
                                             "status": r.get("Status") or "Pending", "notes": r.get("Notes", ""),
                                             "command": r.get("Command", ""), "rollback": r.get("Rollback", ""),
                                             "expected": r.get("Expected", "")})
            except ApiError as exc:
                errors.append(f"row {n}: {exc.message}")
            ph["steps"].append(st)
        for ph in task["phases"]:
            normalize_levels(ph["steps"])
        items.append({"task": task, "errors": errors, "rows": len(g["rows"])})
    return items


def seed_plans(user: str = "system") -> dict:
    doc = empty_plans()
    sunday = dt.date.today() + dt.timedelta(days=(6 - dt.date.today().weekday()) % 7 or 7)
    ctx = {}
    plan_op(doc, "task.create", {"task": {"date": sunday.isoformat(), "customer": "Acme Corp",
                                          "name": "Oracle RAC quarterly OS patching", "ticket_id": "CHG-2041",
                                          "change_ref": "CHG-2041", "window_start": f"{sunday} 22:00",
                                          "window_end": f"{sunday + dt.timedelta(days=1)} 02:00",
                                          "approver": "Acme CAB", "status": "Scheduled",
                                          "assets": "acme-orarac01, acme-orarac02"},
                                 "template": "change"}, user, ctx)
    t = doc["tasks"][0]
    pre, out, post = t["phases"]
    data = [
        (pre, 0, "Confirm last good RMAN backup", "DBA", 10, "rman target / <<< 'list backup summary;'", "", "Backup completed in last 24 h"),
        (pre, 0, "Capture baseline", "Linux L2", 15, "df -hT; systemctl --failed; uptime", "", ""),
        (pre, 1, "Verify RAC resources online", "DBA", 5, "crsctl stat res -t", "", "All resources ONLINE"),
        (out, 0, "Patch node 2", "Linux L2", 0, "", "", ""),
        (out, 1, "Stop CRS on node 2", "DBA", 10, "crsctl stop crs", "crsctl start crs", ""),
        (out, 1, "Apply OS security updates and reboot", "Linux L2", 40, "dnf -y update --security && systemctl reboot",
         "dnf history undo last", ""),
        (out, 0, "Patch node 1 (repeat 2.1 sub-steps)", "Linux L2", 50, "", "", ""),
        (post, 0, "Validate cluster and services", "DBA", 15, "crsctl stat res -t; srvctl status database -d ACMEP",
         "", "Instances open on both nodes"),
        (post, 0, "Customer sign-off recorded in change ticket", "Shift lead", 10, "", "", ""),
    ]
    for ph, lvl, title, owner, mins, cmd, rb, exp in data:
        st = new_step(lvl)
        st.update({"title": title, "owner": owner, "approx_min": mins, "command": cmd, "rollback": rb,
                   "expected": exp})
        ph["steps"].append(st)
    for ph in t["phases"]:
        normalize_levels(ph["steps"])
    plan_op(doc, "template.save", {"task_id": t["id"], "name": "Oracle RAC patching (2 nodes)"}, user, {})
    doc["audit"].append(audit_entry(user, "Execution plans", "Initialise", PLANS_NAME, "Sample plan created"))
    return doc
# --------------------------------------------------------------------------------------
# Hierarchy (Year > Month > Date & Day > Shift > Section > Rows)
# --------------------------------------------------------------------------------------
def empty_hierarchy() -> dict:
    return {"schema_version": 1, "years": {}, "audit": []}


def make_section(name: str, columns: list[str], template: bool) -> dict:
    return {"id": new_id("s_"), "name": name, "template": template,
            "columns": [{"id": new_id("c_"), "name": c} for c in columns], "rows": []}


def parse_year(value) -> str:
    s = str(value if value is not None else "").strip()
    if not re.fullmatch(r"\d{4}", s) or not 1970 <= int(s) <= 2200:
        raise ApiError(400, f"Year must be a four-digit year, got '{s}'.")
    return s


def parse_month(value) -> str:
    try:
        m = int(str(value).strip())
    except (TypeError, ValueError):
        raise ApiError(400, f"Month must be 1-12, got '{value}'.")
    if not 1 <= m <= 12:
        raise ApiError(400, f"Month must be 1-12, got '{value}'.")
    return f"{m:02d}"


def parse_date(value) -> dt.date:
    try:
        return dt.date.fromisoformat(str(value).strip())
    except (TypeError, ValueError):
        raise ApiError(400, f"Date must be YYYY-MM-DD, got '{value}'.")


def parse_shift(value) -> str:
    for s in SHIFTS:
        if str(value or "").strip().lower() == s.lower():
            return s
    raise ApiError(400, f"Shift must be one of {', '.join(SHIFTS)}.")


def resolve_date(p: dict) -> tuple[str, str, str]:
    """Return (year, month, date) from a payload; year/month are derived from date and cross-checked."""
    d = parse_date(p.get("date"))
    y, m = f"{d.year:04d}", f"{d.month:02d}"
    if p.get("year") not in (None, "") and parse_year(p["year"]) != y:
        raise ApiError(400, f"Date {d} does not belong to year {p['year']}.")
    if p.get("month") not in (None, "") and parse_month(p["month"]) != m:
        raise ApiError(400, f"Date {d} does not belong to month {p['month']}.")
    return y, m, d.isoformat()


def normalise_hierarchy(h) -> dict:
    """Validate/repair a hierarchy document (used on load and on JSON import)."""
    if not isinstance(h, dict) or not isinstance(h.get("years", {}), dict):
        raise ApiError(400, "Shift hierarchy JSON must be an object with a 'years' object.")
    out = empty_hierarchy()
    audit = h.get("audit", [])
    out["audit"] = [a for a in audit if isinstance(a, dict)][-AUDIT_CAP:] if isinstance(audit, list) else []
    for y, ynode in h.get("years", {}).items():
        y = parse_year(y)
        ny = out["years"].setdefault(y, {"months": {}})
        for m, mnode in (ynode or {}).get("months", {}).items():
            m = parse_month(m)
            nm = ny["months"].setdefault(m, {"name": calendar.month_name[int(m)], "dates": {}})
            for d, dnode in (mnode or {}).get("dates", {}).items():
                dd = parse_date(d)
                if f"{dd.year:04d}" != y or f"{dd.month:02d}" != m:
                    raise ApiError(400, f"Date {d} is filed under the wrong year/month ({y}-{m}).")
                nd = nm["dates"].setdefault(dd.isoformat(), {"day": dd.strftime("%A"), "shifts": {}})
                for s, snode in (dnode or {}).get("shifts", {}).items():
                    s = parse_shift(s)
                    sections = []
                    for sec in (snode or {}).get("sections", []):
                        cols, col_ids = [], set()
                        for c in sec.get("columns", []):
                            if isinstance(c, str):
                                c = {"name": c}
                            cid = str(c.get("id") or new_id("c_"))
                            if cid in col_ids:
                                cid = new_id("c_")
                            col_ids.add(cid)
                            cols.append({"id": cid, "name": clean_text(c.get("name") or "Column", 80)})
                        rows = []
                        for r in sec.get("rows", []):
                            if not isinstance(r, dict):
                                continue
                            cells = r.get("cells", {}) if isinstance(r.get("cells"), dict) else {}
                            rows.append({"id": str(r.get("id") or new_id("r_")),
                                         "cells": {k: clean_text(v) for k, v in cells.items() if k in col_ids},
                                         "updated_by": clean_text(r.get("updated_by", ""), 120),
                                         "updated_at": clean_text(r.get("updated_at", ""), 40),
                                         **({"carried_from": r["carried_from"]} if r.get("carried_from") else {})})
                        sections.append({"id": str(sec.get("id") or new_id("s_")),
                                         "name": clean_text(sec.get("name") or "Section", 120),
                                         "template": bool(sec.get("template", False)),
                                         "columns": cols, "rows": rows})
                    nd["shifts"][s] = {"created_at": (snode or {}).get("created_at", now_iso()),
                                       "created_by": (snode or {}).get("created_by", ""),
                                       "sections": sections}
                    if (snode or {}).get("cloned_from"):
                        nd["shifts"][s]["cloned_from"] = clean_text(snode["cloned_from"], 60)
    teams = h.get("teams")  # each team's own shift log: same shape and the same checks as the main tree
    if isinstance(teams, dict):
        out["teams"] = {}
        for tid, sub in teams.items():
            if re.fullmatch(r"[A-Za-z0-9_]{1,40}", str(tid)) and isinstance(sub, dict):
                out["teams"][str(tid)] = {"years": normalise_hierarchy({"years": sub.get("years", {}) or {}})["years"]}
    return out


def _shift_node(h: dict, p: dict) -> tuple[dict, str, str]:
    y, m, d = resolve_date(p)
    s = parse_shift(p.get("shift"))
    try:
        return h["years"][y]["months"][m]["dates"][d]["shifts"][s], d, s
    except KeyError:
        raise ApiError(404, f"Shift {d} {s} does not exist yet. Create it first.")


def _section(shift: dict, section_id) -> dict:
    for sec in shift["sections"]:
        if sec["id"] == section_id:
            return sec
    raise ApiError(404, f"Section '{section_id}' not found in this shift.")


def _cells_from_input(section: dict, cells) -> dict:
    if cells is None:
        return {}
    if not isinstance(cells, dict):
        raise ApiError(400, "'cells' must be an object of column id (or column name) to value.")
    by_id = {c["id"] for c in section["columns"]}
    by_name = {norm_key(c["name"]): c["id"] for c in section["columns"]}
    out = {}
    for k, v in cells.items():
        cid = k if k in by_id else by_name.get(norm_key(k))
        if cid:
            out[cid] = clean_text(v)
    return out


def _previous_shift(h: dict, date: str, shift: str):
    target = (date, SHIFTS.index(shift))
    best = None
    for ynode in h["years"].values():
        for mnode in ynode["months"].values():
            for d, dnode in mnode["dates"].items():
                for s, snode in dnode["shifts"].items():
                    key = (d, SHIFTS.index(s))
                    if key < target and (best is None or key > best[0]):
                        best = (key, snode)
    return best


def h_create_node(h: dict, p: dict, user: str) -> dict:
    level = str(p.get("level", "")).lower()
    if level == "year":
        y = parse_year(p.get("year"))
        created = y not in h["years"]
        h["years"].setdefault(y, {"months": {}})
        return {"level": level, "key": y, "created": created}
    if level == "month":
        y, m = parse_year(p.get("year")), parse_month(p.get("month"))
        months = h["years"].setdefault(y, {"months": {}})["months"]
        created = m not in months
        months.setdefault(m, {"name": calendar.month_name[int(m)], "dates": {}})
        return {"level": level, "key": f"{y}-{m}", "created": created}
    if level in ("date", "shift"):
        y, m, d = resolve_date(p)
        months = h["years"].setdefault(y, {"months": {}})["months"]
        dates = months.setdefault(m, {"name": calendar.month_name[int(m)], "dates": {}})["dates"]
        created = d not in dates
        dnode = dates.setdefault(d, {"day": parse_date(d).strftime("%A"), "shifts": {}})
        if level == "date":
            return {"level": level, "key": d, "created": created}
        s = parse_shift(p.get("shift"))
        if s in dnode["shifts"]:
            return {"level": level, "key": f"{d} {s}", "created": False, "carried_from": None}
        carried_from = None
        prev = _previous_shift(h, d, s) if p.get("carry_forward") else None
        if prev:
            (pd, ps_idx), pnode = prev
            carried_from = f"{pd} {SHIFTS[ps_idx]}"
            sections = []
            for sec in pnode["sections"]:
                ns = copy.deepcopy(sec)
                ns["id"] = new_id("s_")
                for r in ns["rows"]:
                    r["id"] = new_id("r_")
                    r["carried_from"] = carried_from
                sections.append(ns)
        else:
            sections = [make_section(n, cols, True) for n, cols in SECTION_TEMPLATES]
        dnode["shifts"][s] = {"created_at": now_iso(), "created_by": user, "sections": sections}
        return {"level": level, "key": f"{d} {s}", "created": True, "carried_from": carried_from}
    raise ApiError(400, "level must be one of: year, month, date, shift.")


def h_delete_node(h: dict, p: dict) -> dict:
    level = str(p.get("level", "")).lower()
    try:
        if level == "year":
            y = parse_year(p.get("year"))
            del h["years"][y]
            return {"deleted": y}
        if level == "month":
            y, m = parse_year(p.get("year")), parse_month(p.get("month"))
            del h["years"][y]["months"][m]
            return {"deleted": f"{y}-{m}"}
        if level == "date":
            y, m, d = resolve_date(p)
            del h["years"][y]["months"][m]["dates"][d]
            return {"deleted": d}
        if level == "shift":
            y, m, d = resolve_date(p)
            s = parse_shift(p.get("shift"))
            del h["years"][y]["months"][m]["dates"][d]["shifts"][s]
            return {"deleted": f"{d} {s}"}
    except KeyError:
        raise ApiError(404, f"Nothing to delete at that {level}.")
    raise ApiError(400, "level must be one of: year, month, date, shift.")


def h_set_columns(section: dict, columns) -> list[str]:
    if not isinstance(columns, list) or not columns:
        raise ApiError(400, "Provide at least one column.")
    if len(columns) > 40:
        raise ApiError(400, "A section can have at most 40 columns.")
    existing_ids = {c["id"] for c in section["columns"]}
    existing_by_name = {c["name"].strip().lower(): c["id"] for c in section["columns"]}
    new_cols, seen, used = [], set(), set()
    for c in columns:
        if isinstance(c, str):
            c = {"name": c}
        if not isinstance(c, dict):
            raise ApiError(400, "Each column must be a name or an object with 'name'.")
        name = clean_text(c.get("name", ""), 80).strip()
        if not name:
            raise ApiError(400, "Column names cannot be empty.")
        if name.lower() in seen:
            raise ApiError(400, f"Duplicate column name '{name}'.")
        seen.add(name.lower())
        cid = c.get("id")
        if cid not in existing_ids or cid in used:
            cid = existing_by_name.get(name.lower())
            if cid in used:
                cid = None
        cid = cid or new_id("c_")
        used.add(cid)
        new_cols.append({"id": cid, "name": name})
    removed = [c["name"] for c in section["columns"] if c["id"] not in used]
    for r in section["rows"]:
        r["cells"] = {k: v for k, v in r["cells"].items() if k in used}
    section["columns"] = new_cols
    return removed




def h_clone_shift(h: dict, p: dict, user: str) -> dict:
    """Copy a whole shift (all sections, columns and optionally rows) to another date and/or shift."""
    src, d, s = _shift_node(h, p)
    target_date = parse_date(p.get("to_date") or d)
    ty, tm, td = f"{target_date.year:04d}", f"{target_date.month:02d}", target_date.isoformat()
    ts = parse_shift(p.get("to_shift") or s)
    if (td, ts) == (d, s):
        raise ApiError(400, "Choose a different date or shift to clone into.")
    mode = "structure" if p.get("mode") == "structure" else "all"
    months = h["years"].setdefault(ty, {"months": {}})["months"]
    dates = months.setdefault(tm, {"name": calendar.month_name[int(tm)], "dates": {}})["dates"]
    dnode = dates.setdefault(td, {"day": target_date.strftime("%A"), "shifts": {}})
    if ts in dnode["shifts"] and not p.get("overwrite"):
        raise ApiError(409, f"The {ts} shift on {td} already exists. Choose another slot or overwrite it.")
    sections = copy.deepcopy(src["sections"])
    for sec in sections:
        sec["id"] = new_id("s_")
        if mode == "structure":
            sec["rows"] = []
        for r in sec["rows"]:
            r["id"] = new_id("r_")
            r["carried_from"] = f"{d} {s}"
            r["updated_by"], r["updated_at"] = user, now_iso()
    dnode["shifts"][ts] = {"created_at": now_iso(), "created_by": user, "cloned_from": f"{d} {s}",
                           "sections": sections}
    return {"key": f"{td} {ts}", "mode": mode, "rows": sum(len(x["rows"]) for x in sections)}


# --------------------------------------------------------------------------------------
# Users, roles and sign-in
# --------------------------------------------------------------------------------------
ROLES = ["Viewer", "Engineer", "Lead", "Admin"]
ROLE_RANK = {r: i + 1 for i, r in enumerate(ROLES)}
ROLE_HELP = {
    "Viewer": "Read everything. Mark bulletins read. Change own password.",
    "Engineer": "Viewer rights plus add and edit records, run execution plans, open SSH/RDP sessions, export.",
    "Lead": "Engineer rights plus delete, sign off go/no-go checkpoints, post bulletins, import and undo, "
            "change columns, restore versions, manage plan templates.",
    "Admin": "Everything, plus the Admin area: users and roles, sessions, terminal recordings, trusted host keys, "
             "settings and backups.",
}
AUTH_METHODS = ["local", "entra", "both"]
PBKDF2_ITERATIONS = 390_000
USERNAME_RE = re.compile(r"^[A-Za-z0-9][A-Za-z0-9._@\-]{1,119}$")
COMMON_PASSWORDS = {"password123", "password1234", "welcome123", "qwerty12345", "letmein123", "admin12345",
                    "changeme123", "p@ssw0rd123", "1234567890"}
DEFAULT_SETTINGS = {
    "session_idle_minutes": 30, "session_max_hours": 12, "lockout_attempts": 5, "lockout_minutes": 15,
    "entra_default_role": "", "terminal_restrict_to_inventory": True, "terminal_allowed_networks": "",
    "terminal_max_per_user": 6, "terminal_idle_minutes": 30, "terminal_record": True,
    "auto_backup": True, "backup_keep": 14,
    "ai_mode": "search", "ai_endpoint": "http://127.0.0.1:11434", "ai_api": "ollama", "ai_model": "", "ai_timeout_seconds": 60,
}
SETTING_TYPES = {k: type(v) for k, v in DEFAULT_SETTINGS.items()}


def hash_password(password: str, iterations: int | None = None) -> str:
    it = iterations or PBKDF2_ITERATIONS
    salt = secrets.token_bytes(16)
    dk = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt, it)
    return f"pbkdf2_sha256${it}${base64.b64encode(salt).decode()}${base64.b64encode(dk).decode()}"


def verify_password(password: str, stored: str) -> bool:
    try:
        algo, it, salt, dk = stored.split("$")
        if algo != "pbkdf2_sha256":
            return False
        calc = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), base64.b64decode(salt), int(it))
        return hmac.compare_digest(calc, base64.b64decode(dk))
    except (ValueError, TypeError, AttributeError):
        return False


_DUMMY_HASH = None


def burn_password_time(password: str):
    """Spend the same time as a real check so unknown usernames cannot be told apart by timing."""
    global _DUMMY_HASH
    if _DUMMY_HASH is None:
        _DUMMY_HASH = hash_password(secrets.token_hex(8))
    verify_password(password, _DUMMY_HASH)


def check_password_policy(password: str, username: str = ""):
    if not isinstance(password, str) or len(password) < 10:
        raise ApiError(400, "Passwords need at least 10 characters.")
    if len(password) > 256:
        raise ApiError(400, "Passwords can be at most 256 characters.")
    if not (re.search(r"[A-Za-z]", password) and re.search(r"[0-9]", password)):
        raise ApiError(400, "Passwords need at least one letter and one number.")
    if password.lower() in COMMON_PASSWORDS or (username and username.lower() in password.lower()):
        raise ApiError(400, "That password is too easy to guess. Choose another.")


def temp_password() -> str:
    alphabet = "ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnopqrstuvwxyz23456789"
    while True:
        pw = "".join(secrets.choice(alphabet) for _ in range(14))
        if re.search(r"[0-9]", pw) and re.search(r"[A-Za-z]", pw):
            return pw


def empty_users() -> dict:
    return {"schema_version": 1, "users": {}, "settings": dict(DEFAULT_SETTINGS), "known_hosts": {}, "audit": []}


def normalize_settings(raw) -> dict:
    out = dict(DEFAULT_SETTINGS)
    for k, v in (raw or {}).items():
        if k not in DEFAULT_SETTINGS:
            continue
        t = SETTING_TYPES[k]
        try:
            if t is bool:
                out[k] = v if isinstance(v, bool) else truthy(v)
            elif t is int:
                out[k] = max(0, min(int(v), 100000))
            else:
                out[k] = clean_text(v, 2000)
        except (TypeError, ValueError):
            continue
    if out["entra_default_role"] not in ("", *ROLES):
        out["entra_default_role"] = ""
    out["session_idle_minutes"] = max(out["session_idle_minutes"], 5)
    out["session_max_hours"] = max(out["session_max_hours"], 1)
    out["lockout_attempts"] = max(out["lockout_attempts"], 3)
    out["terminal_max_per_user"] = max(out["terminal_max_per_user"], 1)
    out["terminal_idle_minutes"] = max(out["terminal_idle_minutes"], 1)
    out["backup_keep"] = max(out["backup_keep"], 1)
    if out["ai_mode"] not in ("off", "search", "model"):
        out["ai_mode"] = "search"
    if out["ai_api"] not in ("ollama", "openai"):
        out["ai_api"] = "ollama"
    out["ai_timeout_seconds"] = max(5, min(out["ai_timeout_seconds"], 600))
    return out


def normalize_users(doc) -> dict:
    out = empty_users()
    if not isinstance(doc, dict):
        return out
    for key, u in (doc.get("users") or {}).items():
        if not isinstance(u, dict):
            continue
        username = str(u.get("username") or key).strip()
        out["users"][username.lower()] = {
            "username": username, "name": clean_text(u.get("name", ""), 120) or username,
            "email": clean_text(u.get("email", ""), 200).strip().lower(),
            "role": u.get("role") if u.get("role") in ROLES else "Viewer",
            "auth": u.get("auth") if u.get("auth") in AUTH_METHODS else "local",
            "password": str(u.get("password", "")), "must_change": bool(u.get("must_change", False)),
            "disabled": bool(u.get("disabled", False)), "created_by": clean_text(u.get("created_by", ""), 120),
            "created_at": clean_text(u.get("created_at", ""), 40), "last_login": clean_text(u.get("last_login", ""), 40),
            "password_changed": clean_text(u.get("password_changed", ""), 40),
        }
    out["settings"] = normalize_settings(doc.get("settings"))
    if isinstance(doc.get("known_hosts"), dict):
        out["known_hosts"] = {str(k): v for k, v in doc["known_hosts"].items() if isinstance(v, dict)}
    audit = doc.get("audit", [])
    out["audit"] = [a for a in audit if isinstance(a, dict)][-AUDIT_CAP:] if isinstance(audit, list) else []
    return out


def public_user(u: dict) -> dict:
    return {k: v for k, v in u.items() if k != "password"} | {"has_password": bool(u.get("password")),
                                                                "rank": ROLE_RANK.get(u.get("role"), 0)}


def find_user_by_identity(users: dict, identity: str):
    ident = str(identity or "").strip().lower()
    if not ident:
        return None
    if ident in users:
        return users[ident]
    for u in users.values():
        if u.get("email") and u["email"] == ident:
            return u
    return None


class SessionManager:
    """In-memory web sessions (signing in again after a server restart is expected)."""

    def __init__(self):
        self.lock = threading.Lock()
        self.items: dict = {}

    def create(self, username: str, ip: str, method: str) -> dict:
        sess = {"id": secrets.token_urlsafe(32), "csrf": secrets.token_urlsafe(24), "user": username.lower(),
                "ip": ip, "method": method, "created": time.time(), "last": time.time(), "created_iso": now_iso()}
        with self.lock:
            self.items[sess["id"]] = sess
        return sess

    def get(self, sid: str, settings: dict, touch: bool = True):
        if not sid:
            return None
        with self.lock:
            sess = self.items.get(sid)
            if not sess:
                return None
            now = time.time()
            if (now - sess["last"] > settings["session_idle_minutes"] * 60
                    or now - sess["created"] > settings["session_max_hours"] * 3600):
                del self.items[sid]
                return None
            if touch:
                sess["last"] = now
            return sess

    def drop(self, sid: str):
        with self.lock:
            self.items.pop(sid, None)

    def drop_handle(self, handle: str) -> bool:
        with self.lock:
            for sid in list(self.items):
                if hmac.compare_digest(session_handle(sid), str(handle)):
                    del self.items[sid]
                    return True
        return False

    def drop_user(self, username: str, except_sid: str = ""):
        with self.lock:
            for sid in [s for s, v in self.items.items() if v["user"] == username.lower() and s != except_sid]:
                del self.items[sid]

    def list(self) -> list[dict]:
        with self.lock:
            return [{"handle": session_handle(s), "user": v["user"], "ip": v["ip"], "method": v["method"],
                     "created": v["created_iso"], "idle_seconds": int(time.time() - v["last"])}
                    for s, v in self.items.items()]


def session_handle(sid: str) -> str:
    """Opaque, non-reversible reference to a session (admins never see real session ids)."""
    return hashlib.sha256(("msp-session:" + sid).encode()).hexdigest()[:20]


class LoginThrottle:
    """Lock an account for a while after repeated failures (per username, any source)."""

    IP_LIMIT = 30          # failed sign-ins from one address ...
    IP_WINDOW = 15 * 60    # ... within this many seconds

    def __init__(self):
        self.lock = threading.Lock()
        self.fails: dict = {}
        self.ip_fails: dict = {}

    def check_ip(self, ip: str):
        with self.lock:
            now = time.time()
            recent = [t for t in self.ip_fails.get(ip, []) if now - t < self.IP_WINDOW]
            self.ip_fails[ip] = recent
            if len(recent) >= self.IP_LIMIT:
                raise ApiError(429, "Too many failed sign-ins from this computer. Wait 15 minutes.")

    def failed_ip(self, ip: str):
        with self.lock:
            self.ip_fails.setdefault(ip, []).append(time.time())

    def check(self, username: str):
        with self.lock:
            rec = self.fails.get(username.lower())
            if rec and rec["until"] > time.time():
                mins = int((rec["until"] - time.time()) // 60) + 1
                raise ApiError(429, f"Too many failed sign-ins. Try again in {mins} minute(s) or ask an admin.")

    def failed(self, username: str, settings: dict):
        with self.lock:
            rec = self.fails.setdefault(username.lower(), {"n": 0, "until": 0})
            rec["n"] += 1
            if rec["n"] >= settings["lockout_attempts"]:
                rec["until"] = time.time() + settings["lockout_minutes"] * 60
                rec["n"] = 0

    def reset(self, username: str):
        with self.lock:
            self.fails.pop(username.lower(), None)


class EntraAuth:
    """Microsoft 365 sign-in: OpenID Connect authorization-code flow with PKCE (no extra packages).

    The ID token is received directly from the Microsoft token endpoint over TLS using the client
    secret, so per OpenID Connect Core 3.1.3.7 TLS validation stands in for signature checking;
    audience, issuer tenant, expiry and nonce are still verified.
    """

    def __init__(self, tenant: str, client_id: str, client_secret: str, redirect_uri: str,
                 auth_base: str = "https://login.microsoftonline.com", timeout: float = 20.0):
        self.tenant, self.client_id, self.client_secret = tenant, client_id, client_secret
        self.redirect_uri, self.auth_base, self.timeout = redirect_uri, auth_base.rstrip("/"), timeout
        self.pending: dict = {}
        self.lock = threading.Lock()

    def login_url(self, next_path: str) -> str:
        verifier = secrets.token_urlsafe(48)
        challenge = base64.urlsafe_b64encode(hashlib.sha256(verifier.encode()).digest()).decode().rstrip("=")
        state, nonce = secrets.token_urlsafe(24), secrets.token_urlsafe(24)
        with self.lock:
            now = time.time()
            self.pending = {k: v for k, v in self.pending.items() if now - v["at"] < 600}
            self.pending[state] = {"verifier": verifier, "nonce": nonce, "next": safe_next(next_path), "at": now}
        q = urllib.parse.urlencode({"client_id": self.client_id, "response_type": "code", "redirect_uri": self.redirect_uri,
                                    "response_mode": "query", "scope": "openid profile email", "state": state,
                                    "nonce": nonce, "code_challenge": challenge, "code_challenge_method": "S256",
                                    "prompt": "select_account"})
        return f"{self.auth_base}/{urllib.parse.quote(self.tenant)}/oauth2/v2.0/authorize?{q}"

    def complete(self, code: str, state: str) -> tuple[dict, str]:
        with self.lock:
            pend = self.pending.pop(state or "", None)
        if not pend or time.time() - pend["at"] > 600:
            raise ApiError(400, "This sign-in link has expired. Start again.")
        body = urllib.parse.urlencode({"client_id": self.client_id, "client_secret": self.client_secret,
                                       "grant_type": "authorization_code", "code": code,
                                       "redirect_uri": self.redirect_uri, "code_verifier": pend["verifier"],
                                       "scope": "openid profile email"}).encode()
        req = urllib.request.Request(f"{self.auth_base}/{urllib.parse.quote(self.tenant)}/oauth2/v2.0/token",
                                     data=body, method="POST",
                                     headers={"Content-Type": "application/x-www-form-urlencoded"})
        try:
            with http_open(req, self.timeout) as resp:
                tokens = json.loads(resp.read().decode("utf-8"))
        except urllib.error.HTTPError as exc:
            raise ApiError(502, f"Microsoft sign-in failed (HTTP {exc.code}).")
        except urllib.error.URLError as exc:
            raise ApiError(503, f"Cannot reach Microsoft sign-in: {exc.reason}")
        claims = decode_jwt_payload(tokens.get("id_token", ""))
        now = time.time()
        aud = claims.get("aud")
        if (aud if isinstance(aud, str) else (aud or [None])[0]) != self.client_id:
            raise ApiError(401, "Sign-in token was not issued for this application.")
        if claims.get("nonce") != pend["nonce"]:
            raise ApiError(401, "Sign-in token did not match this sign-in attempt.")
        if float(claims.get("exp", 0)) < now - 300:
            raise ApiError(401, "Sign-in token has expired.")
        if self.tenant.lower() not in ("common", "organizations") and claims.get("tid", "").lower() != self.tenant.lower() \
                and self.tenant.lower() not in str(claims.get("iss", "")).lower():
            raise ApiError(401, "Sign-in came from a different Microsoft 365 tenant.")
        return claims, pend["next"]


def decode_jwt_payload(token: str) -> dict:
    try:
        payload = token.split(".")[1]
        payload += "=" * (-len(payload) % 4)
        return json.loads(base64.urlsafe_b64decode(payload.encode()).decode("utf-8"))
    except (IndexError, ValueError):
        raise ApiError(401, "Microsoft returned an unreadable sign-in token.")


def safe_next(path) -> str:
    """Only allow same-site relative redirect targets (prevents open redirects)."""
    p = str(path or "/")
    if not p.startswith("/") or p.startswith("//") or "\\" in p or "\n" in p or "\r" in p:
        return "/"
    return p[:500]


def guac_encrypted_token(secret_hex: str, payload: dict) -> str:
    """Build data for Guacamole's guacamole-auth-json extension:
    HMAC-SHA256 signature + JSON, encrypted with AES-128-CBC (zero IV, PKCS#7), base64."""
    try:
        from cryptography.hazmat.primitives import padding
        from cryptography.hazmat.primitives.ciphers import Cipher, algorithms, modes
    except ImportError:
        raise ApiError(501, "In-browser RDP needs the 'cryptography' package on the server (pip install cryptography).")
    try:
        key = bytes.fromhex(secret_hex.strip())
    except ValueError:
        raise ApiError(500, "The Guacamole secret key must be 32 hexadecimal characters.")
    if len(key) != 16:
        raise ApiError(500, "The Guacamole secret key must be 32 hexadecimal characters (128 bits).")
    data = json.dumps(payload, separators=(",", ":")).encode("utf-8")
    signed = hmac.new(key, data, hashlib.sha256).digest() + data
    padder = padding.PKCS7(128).padder()
    padded = padder.update(signed) + padder.finalize()
    enc = Cipher(algorithms.AES(key), modes.CBC(b"\0" * 16)).encryptor()  # nosec B305 - format mandated by Guacamole
    return base64.b64encode(enc.update(padded) + enc.finalize()).decode()


# ---- what each request needs ---------------------------------------------------------------
LEAD_PLAN_OPS = {"task.delete", "phase.delete", "phase.signoff", "template.save", "template.delete", "task.columns"}


def required_rank(method: str, path: str, body_fn) -> int:
    """Minimum role rank for an API call. body_fn() lazily returns the parsed JSON body."""
    if path.startswith("/api/admin/"):
        return 4
    if path == "/api/me/password" or re.fullmatch(r"/api/bulletins/[A-Za-z0-9_\-]+/read", path):
        return 1
    if path == "/api/ai/ask":
        return 1  # read-only: every signed-in person may ask
    if path.startswith("/api/org"):
        if method == "GET" and not path.endswith("/history"):
            return 1  # everyone sees their own teams; msp_org decides what else
        return 3  # changes: team leads and administrators; msp_org checks which team
    if path.startswith("/api/roster"):
        if method == "GET":
            if path == "/api/roster/deleted":
                return 3
            return 2 if path.endswith("/export") or path == "/api/roster/template" else 1
        if re.fullmatch(r"/api/roster/[A-Za-z0-9_\-]+/(leave|swap|cancel|consent)", path):
            return 2  # the roster module also limits Engineers to their own requests
        return 3
    if method == "GET":
        if path.startswith(("/api/export/", "/api/template/", "/api/rdp/", "/api/terminal")):
            return 2
        return 1
    if path.startswith(("/api/import", "/api/imports")) or path == "/api/history/restore":
        return 3
    m = re.fullmatch(r"/api/table/([A-Za-z_]+)(/.*)?", path)
    if m:
        try:
            sheet = resolve_module(m.group(1))
        except ApiError:
            sheet = ""
        if sheet == "Bulletins" or method == "DELETE" or (m.group(2) or "").endswith("/columns"):
            return 3
        return 2
    if path == "/api/plans/op":
        return 3 if str(body_fn().get("op", "")) in LEAD_PLAN_OPS else 2
    if path.startswith("/api/hierarchy"):
        if (path == "/api/hierarchy/node" and method == "DELETE") or (path == "/api/hierarchy/section" and method == "DELETE") \
                or path == "/api/hierarchy/section/columns":
            return 3
        return 2
    return 2


# --------------------------------------------------------------------------------------
# Metadata document (version history, import batches, bulletin read receipts)
# --------------------------------------------------------------------------------------
def empty_meta() -> dict:
    return {"schema_version": 1, "history": {}, "imports": [], "bulletin_reads": {}, "terminal_sessions": [],
            "backups": [], "audit": []}


def normalize_meta(doc) -> dict:
    out = empty_meta()
    if isinstance(doc, dict):
        if isinstance(doc.get("history"), dict):
            out["history"] = {k: v for k, v in doc["history"].items() if isinstance(v, list)}
        if isinstance(doc.get("imports"), list):
            out["imports"] = [b for b in doc["imports"] if isinstance(b, dict)][-IMPORT_BATCH_CAP:]
        if isinstance(doc.get("bulletin_reads"), dict):
            out["bulletin_reads"] = {k: list(v) for k, v in doc["bulletin_reads"].items() if isinstance(v, list)}
        for key, cap in (("terminal_sessions", 5000), ("backups", 500), ("audit", AUDIT_CAP)):
            if isinstance(doc.get(key), list):
                out[key] = [x for x in doc[key] if isinstance(x, dict)][-cap:]
    return out


def _short(v) -> str:
    s = str(v if v is not None else "")
    return s if len(s) <= 300 else s[:297] + "..."


def diff_records(before, after) -> list[dict]:
    """Field-level differences between two versions of a row or a plan task (max 60 entries)."""
    if before is None or after is None:
        return []
    changes = []
    skip = set(SYSTEM_COLS) | {"updated_by", "updated_at"}
    if "phases" in before or "phases" in after:
        for f in TASK_FIELDS:
            if before.get(f, "") != after.get(f, ""):
                changes.append({"field": f, "old": _short(before.get(f)), "new": _short(after.get(f))})
        bp = {ph["id"]: ph for ph in before.get("phases", [])}
        ap = {ph["id"]: ph for ph in after.get("phases", [])}
        for pid, ph in ap.items():
            if pid not in bp:
                changes.append({"field": "phase added", "old": "", "new": ph["name"]})
            elif bp[pid]["name"] != ph["name"]:
                changes.append({"field": "phase renamed", "old": bp[pid]["name"], "new": ph["name"]})
        for pid, ph in bp.items():
            if pid not in ap:
                changes.append({"field": "phase removed", "old": ph["name"], "new": ""})
        bs = {s["id"]: s for ph in before.get("phases", []) for s in ph["steps"]}
        as_ = {s["id"]: s for ph in after.get("phases", []) for s in ph["steps"]}
        for sid, st in as_.items():
            if sid not in bs:
                changes.append({"field": "step added", "old": "", "new": st.get("title", "")})
                continue
            for k in STEP_CORE_KEYS + ["level"]:
                if bs[sid].get(k, "") != st.get(k, ""):
                    changes.append({"field": f"step '{_short(st.get('title'))[:40]}' {k}",
                                    "old": _short(bs[sid].get(k)), "new": _short(st.get(k))})
        for sid, st in bs.items():
            if sid not in as_:
                changes.append({"field": "step removed", "old": st.get("title", ""), "new": ""})
        bo = [s["id"] for ph in before.get("phases", []) for s in ph["steps"]]
        ao = [s["id"] for ph in after.get("phases", []) for s in ph["steps"]]
        if [x for x in bo if x in as_] != [x for x in ao if x in bs]:
            changes.append({"field": "step order", "old": "", "new": "changed"})
    else:
        for k in sorted(set(before) | set(after)):
            if k in skip:
                continue
            if before.get(k, "") != after.get(k, ""):
                changes.append({"field": k, "old": _short(before.get(k)), "new": _short(after.get(k))})
    return changes[:60]


def flatten_hierarchy(hier: dict) -> list[dict]:
    rows = []
    for y in sorted(hier.get("years", {})):
        for m in sorted(hier["years"][y].get("months", {})):
            for d in sorted(hier["years"][y]["months"][m].get("dates", {})):
                dnode = hier["years"][y]["months"][m]["dates"][d]
                shifts = dnode.get("shifts", {})
                for s in sorted(shifts, key=lambda k: SHIFTS.index(k) if k in SHIFTS else 9):
                    for sec in shifts[s].get("sections", []):
                        names = {c["id"]: c["name"] for c in sec.get("columns", [])}
                        for row in sec.get("rows", []):
                            details = " | ".join(f"{names[cid]}: {row['cells'].get(cid, '')}"
                                                 for cid in names if row["cells"].get(cid, ""))
                            rows.append({"Year": y, "Month": m, "Date": d, "Day": dnode.get("day", ""), "Shift": s,
                                         "Section": sec.get("name", ""), "Row_ID": row["id"], "Details": details,
                                         "Updated_By": row.get("updated_by", ""),
                                         "Updated_At": row.get("updated_at", "")})
    return rows


def resolve_module(name: str) -> str:
    """Accept a module name (assets), sheet name (Asset_Master) or title; return sheet name, 'plans' or 'shiftlog'."""
    n = str(name or "").strip()
    if n in TABLES:
        return n
    if n in MODULES:
        return MODULES[n]
    canon = canonical_sheet(n)
    if canon in TABLES or canon in ("plans", "shiftlog"):
        return canon
    raise ApiError(404, f"Unknown module '{name}'. Use one of: {', '.join(list(MODULES) + ['plans', 'shiftlog'])}.")


def detect_module(headers: list[str]):
    scores = {}
    for sheet in SHEET_ORDER:
        cols = default_columns(sheet)
        scores[sheet] = sum(1 for h in headers if match_header(h, cols, sheet) not in (None, ID_COL))
    distinctive = {"Task_Name", "Step", "Phase", "Step_No", "Approx_Min", "Window_Start", "Planned_Start", "Rollback"}
    scores["plans"] = sum(2 for h in headers if plan_header_key(h) in distinctive)
    keys = {norm_key(h) for h in headers}
    if {"shift", "section"} <= keys:
        scores["shiftlog"] = 100
    best = max(scores, key=scores.get)
    return best if scores[best] >= 2 else None


def decode_text(data: bytes) -> str:
    for enc in ("utf-8-sig", "utf-16", "cp1252"):
        try:
            text = data.decode(enc)
            if enc == "utf-16" and not data[:2] in (b"\xff\xfe", b"\xfe\xff"):
                continue
            return text
        except UnicodeDecodeError:
            continue
    return data.decode("latin-1")


def parse_delimited(text: str) -> list[list[str]]:
    sample = "\n".join(text.splitlines()[:30])
    try:
        dialect = csv.Sniffer().sniff(sample, delimiters=",\t;|")
        delim = dialect.delimiter
    except csv.Error:
        counts = {d: sample.count(d) for d in ",\t;|"}
        delim = max(counts, key=counts.get) if any(counts.values()) else ","
    return [list(r) for r in csv.reader(io.StringIO(text), delimiter=delim)]




# --------------------------------------------------------------------------------------
# Shift log as flat rows: one line per shift-log row (template, export and import share it)
#   Date | Shift | Section | <columns of every section side by side> | Row_ID
# --------------------------------------------------------------------------------------
SHIFTLOG_KEYS = {"date": "Date", "shift": "Shift", "section": "Section", "rowid": "Row_ID", "id": "Row_ID",
                 "year": None, "month": None, "day": None, "updatedby": None, "updatedat": None}
SHIFTLOG_EXAMPLES = {
    "Shift Handover Overview": {"Check": "Backup monitoring", "Client": "Acme Corp", "Status": "Green",
                                "Details": "All nightly jobs completed", "Owner": "Night shift"},
    "Scheduled": {"Client": "Acme Corp", "Event": "RAC OS patching", "Assets": "acme-orarac01, acme-orarac02",
                  "Window_Start": "22:00", "Window_End": "02:00", "Change_Ref": "CHG-2041", "Status": "Approved"},
    "On-Hold OR Customer Hold": {"Ticket": "INC-104250", "Client": "Globex", "Priority": "P3",
                                 "Blocked_On": "Customer approval", "Next_Follow_Up": "Tomorrow 10:00"},
    "Patch and firmware details": {"Client": "Initech", "Asset": "ini-sol01", "Current_Level": "11.4.62",
                                   "Target_Level": "11.4.66", "Compliance": "Due"},
    "In Work - All Clients": {"Ticket": "INC-104233", "Client": "Acme Corp", "Priority": "P1",
                              "Task": "ASM DATA diskgroup space", "Owner": "DBA on call", "Status": "In progress",
                              "Next_Step": "Add 200 GB LUN"},
    "Weekly Meetings": {"Meeting": "Acme service review", "Client": "Acme Corp", "Schedule": "Thu 15:00",
                        "Attendees": "SDM, DBA lead", "MOM": "Patch plan agreed"},
    "Weekly Tasks": {"Task": "Weekend health checks", "Client": "All", "Frequency": "Weekly",
                     "Playbook": "RB-HC-01", "Owner": "Weekend shift", "Status": "Planned"},
}


def shiftlog_value_columns(hier: dict | None = None) -> list[str]:
    """Union of column names: standard sections first, then any custom ones in the shift log."""
    names, seen = [], set()

    def add(n):
        if norm_key(n) and norm_key(n) not in seen:
            seen.add(norm_key(n))
            names.append(n)
    for _, cols in SECTION_TEMPLATES:
        for c in cols:
            add(c)
    for y in sorted((hier or {}).get("years", {})):
        for m in sorted(hier["years"][y]["months"]):
            for d in sorted(hier["years"][y]["months"][m]["dates"]):
                for snode in hier["years"][y]["months"][m]["dates"][d]["shifts"].values():
                    for sec in snode["sections"]:
                        for c in sec["columns"]:
                            add(c["name"])
    return names


def shiftlog_export_rows(hier: dict) -> tuple[list[str], list[dict]]:
    cols = ["Year", "Month", "Date", "Day", "Shift", "Section"] + shiftlog_value_columns(hier) + \
           ["Row_ID", "Updated_By", "Updated_At"]
    canon = {norm_key(c): c for c in cols}
    rows = []
    for y in sorted(hier.get("years", {})):
        for m in sorted(hier["years"][y]["months"]):
            for d in sorted(hier["years"][y]["months"][m]["dates"]):
                dnode = hier["years"][y]["months"][m]["dates"][d]
                for s in sorted(dnode["shifts"], key=lambda k: SHIFTS.index(k) if k in SHIFTS else 9):
                    for sec in dnode["shifts"][s]["sections"]:
                        for r in sec["rows"]:
                            row = {"Year": y, "Month": m, "Date": d, "Day": dnode.get("day", ""), "Shift": s,
                                   "Section": sec["name"], "Row_ID": r["id"], "Updated_By": r.get("updated_by", ""),
                                   "Updated_At": r.get("updated_at", "")}
                            for c in sec["columns"]:
                                if r["cells"].get(c["id"]):
                                    row[canon[norm_key(c["name"])]] = r["cells"][c["id"]]
                            rows.append(row)
    return cols, rows


def parse_shiftlog_rows(section: dict) -> list[dict]:
    """Read flat shift-log lines. Returns [{date, shift, section, row_id, cells{column name: value}, errors}]."""
    out = []
    for n, raw in enumerate(section["rows"], start=1):
        r = {"date": "", "shift": "", "section": "", "row_id": "", "cells": {}, "errors": [], "line": n}
        for h, v in raw.items():
            v = clean_text(v).strip()
            key = SHIFTLOG_KEYS.get(norm_key(h), "cell")
            if key == "Date":
                r["date"] = v
            elif key == "Shift":
                r["shift"] = v
            elif key == "Section":
                r["section"] = clean_text(v, 120)
            elif key == "Row_ID":
                r["row_id"] = v
            elif key == "cell" and v and str(h).strip() and not re.fullmatch(r"Column_\d+", str(h)):
                r["cells"][clean_text(h, 80).strip()] = v
        try:
            r["date"] = parse_date(norm_date(r["date"], "Date") if r["date"] else "").isoformat()
        except ApiError as exc:
            r["errors"].append(exc.message if r["date"] else "Date is required.")
        try:
            r["shift"] = parse_shift(r["shift"])
        except ApiError as exc:
            r["errors"].append(exc.message)
        if not r["section"]:
            r["errors"].append("Section is required (for example In Work - All Clients).")
        if not r["cells"]:
            r["errors"].append("The line has no values.")
        out.append(r)
    return out


def _find_section_by_name(shift_node: dict | None, name: str):
    if not shift_node:
        return None
    return next((s for s in shift_node["sections"] if norm_key(s["name"]) == norm_key(name)), None)


def _template_section_name(name: str) -> str:
    """Use the standard spelling when an imported section matches a template section."""
    return next((n for n, _ in SECTION_TEMPLATES if norm_key(n) == norm_key(name)), name)


def _hier_shift(h: dict, date: str, shift: str):
    y, m = date[:4], date[5:7]
    return h["years"].get(y, {}).get("months", {}).get(m, {}).get("dates", {}).get(date, {}).get("shifts", {}).get(shift)


def row_cells_by_name(sec: dict, row: dict) -> dict:
    names = {c["id"]: c["name"] for c in sec["columns"]}
    return {norm_key(names[k]): v for k, v in row["cells"].items() if k in names and v}

# --------------------------------------------------------------------------------------
# Data store
# --------------------------------------------------------------------------------------
class DataStore:
    DOCS = {"xlsx": XLSX_NAME, "plans": PLANS_NAME, "shift": JSON_NAME, "meta": META_NAME, "users": USERS_NAME}

    def __init__(self, storage, refresh_interval: float = 5.0, seed: bool = True):
        self.storage = storage
        self.refresh_interval = refresh_interval
        self.seed = seed
        self.lock = threading.RLock()
        self.state: dict = {}
        self.versions: dict = {}
        self.previews: dict = {}
        self.last_sync = None
        self.last_error = None
        self._last_check = 0.0
        self.roster_store = None  # set by AppServer when msp_roster.py is present
        self._req = threading.local()  # per-request team scope (each request runs in its own thread)
        self.org_store = None  # set by AppServer when msp_org.py is present
        with self.lock:
            for k in self.DOCS:
                self._load(k)
            self._last_check = time.monotonic()

    # -- codecs --------------------------------------------------------------------------
    def _init_state(self, k):
        if k == "xlsx":
            return {"tables": seed_tables() if self.seed else default_tables(),
                    "audit": [audit_entry("system", "Storage", "Initialise", XLSX_NAME, "Workbook created")]}
        if k == "plans":
            return seed_plans() if self.seed else empty_plans()
        if k == "shift":
            h = empty_hierarchy()
            h["audit"].append(audit_entry("system", "Shift log", "Initialise", JSON_NAME, "Shift log created"))
            return h
        if k == "users":
            return empty_users()
        return empty_meta()

    @staticmethod
    def _parse(k, data: bytes):
        if k == "xlsx":
            tables, audit = tables_from_workbook(data)
            return {"tables": tables, "audit": audit}
        doc = json.loads(data.decode("utf-8-sig"))
        return {"plans": normalize_plans, "shift": normalise_hierarchy, "meta": normalize_meta,
                "users": normalize_users}[k](doc)

    @staticmethod
    def _dump(k, state) -> bytes:
        if k == "xlsx":
            return write_tables_workbook(state["tables"], state["audit"])
        return json.dumps(state, indent=1, ensure_ascii=False).encode("utf-8")

    def _load(self, k, attempt: int = 0):
        name = self.DOCS[k]
        data, ver = self.storage.read(name)
        if data is None:
            state = self._init_state(k)
            try:
                ver = self.storage.write(name, self._dump(k, state), None)
            except ConflictError:
                if attempt < 3:
                    return self._load(k, attempt + 1)
                raise
        else:
            try:
                state = self._parse(k, data)
            except (ApiError, ValueError) as exc:
                msg = exc.message if isinstance(exc, ApiError) else str(exc)
                raise StorageError(500, f"Central file {name} is unreadable ({msg}). "
                                        f"Restore a previous version from SharePoint version history.")
        self.state[k], self.versions[k] = state, ver
        self.last_sync = now_iso()

    def refresh(self, force: bool = False) -> bool:
        with self.lock:
            now = time.monotonic()
            if not force and now - self._last_check < self.refresh_interval:
                return False
            self._last_check = now
            changed = False
            try:
                for k, name in self.DOCS.items():
                    if self.storage.version(name) != self.versions[k]:
                        self._load(k)
                        changed = True
                self.last_error = None
            except StorageError as exc:
                self.last_error = exc.message
            return changed

    # -- generic read-modify-write --------------------------------------------------------
    def mutate(self, k, fn, user, area, action, target="", details=""):
        with self.lock:
            self.refresh()
            for _ in range(4):
                state = copy.deepcopy(self.state[k])
                ctx = {"target": target, "details": details, "history": []}
                result = fn(state, ctx)
                if (k != "meta" or ctx.get("audit_meta")) and not ctx.get("no_audit"):
                    state["audit"] = (state["audit"] + [audit_entry(user, area, ctx.get("action") or action,
                                                                    ctx["target"], ctx["details"])])[-AUDIT_CAP:]
                try:
                    ver = self.storage.write(self.DOCS[k], self._dump(k, state), self.versions[k])
                except ConflictError:
                    self._load(k)  # someone else saved first: reload and re-apply our change
                    continue
                self.state[k], self.versions[k] = state, ver
                self.last_sync = now_iso()
                break
            else:
                raise ApiError(409, f"{self.DOCS[k]} kept changing while saving. Please retry.")
            if ctx["history"]:
                self._record_history(ctx["history"], user, ctx.get("batch", ""))
            return result

    def _record_history(self, entries, user, batch_id=""):
        def fn(meta, ctx):
            for key, action, before, after in entries:
                lst = meta["history"].setdefault(key, [])
                lst.append({"v": (lst[-1]["v"] + 1) if lst else 1, "at": now_iso(), "by": user, "action": action,
                            "changes": diff_records(before, after),
                            "snapshot": copy.deepcopy(after if after is not None else before),
                            "deleted": after is None, "batch": batch_id})
                del lst[:-HISTORY_PER_RECORD]
        try:
            self.mutate("meta", fn, user, "", "")
        except (StorageError, ApiError) as exc:  # history must never block the user's actual change
            self.last_error = f"Version history not saved: {exc.message}"

    # -- read API ------------------------------------------------------------------------
    @property
    def revision(self) -> str:
        return "|".join(str(self.versions.get(k)) for k in self.DOCS)

    def audit(self, limit: int = 500) -> list[dict]:
        with self.lock:
            keyed = []
            for k in ("xlsx", "plans", "shift", "users", "meta"):
                keyed += [(str(a.get("Timestamp", "")), i, a) for i, a in enumerate(self.state[k].get("audit", []))]
        keyed.sort(key=lambda x: (x[0], x[1]), reverse=True)
        return [x[2] for x in keyed[:max(1, min(int(limit), 20000))]]

    def plans_view(self) -> dict:
        doc = self.state["plans"]
        acc = self.scope()["access"] if msp_access else None
        tasks = doc["tasks"] if acc is None else [v for v in (acc.plan_view(t) for t in doc["tasks"]) if v is not None]
        return {"tasks": [compute_task(t) for t in tasks],
                "templates": [{"id": t["id"], "name": t["name"], "created_by": t["created_by"],
                               "phases": len(t["phases"]), "steps": sum(len(p["steps"]) for p in t["phases"])}
                              for t in doc["templates"]],
                "builtin": [{"id": k, "name": v[0], "phases": v[1]} for k, v in BUILTIN_PLAN_TEMPLATES.items()],
                "presets": PHASE_PRESETS, "step_statuses": STEP_STATUSES, "task_statuses": TASK_STATUSES}

    def snapshot(self, user: str = "", audit_limit: int = 100) -> dict:
        with self.lock:
            self.refresh()
            hier = self._hier_view()
            reads = self.state["meta"]["bulletin_reads"]
            sc = self.scope()
            acc = sc.get("access")
            audit = self.audit(audit_limit * 5 if acc else audit_limit)
            if acc is not None:
                audit = [a for a in audit if acc.audit_visible(a)][:audit_limit]
            return {"tables": self._scoped_tables(copy.deepcopy(self.state["xlsx"]["tables"])), "plans": self.plans_view(),
                    "hierarchy": copy.deepcopy(hier), "audit": audit, "revision": self.revision,
                    "my_reads": [b for b, users in reads.items() if user in users],
                    "imports": self.imports(30), "schema": self.schema(),
                    "team_scope": sc["active"], "team": sc["team"], "my_teams": sc["teams"],
                    "legacy": self.legacy_counts() if sc["active"] and sc["team"] else None}

    @staticmethod
    def schema() -> dict:
        return {s: {"module": t["module"], "title": t["title"], "noun": t["noun"], "required": t["required"],
                    "choices": t["choices"], "dates": t["dates"], "label_key": t["label_key"]}
                for s, t in TABLES.items()}

    def imports(self, limit: int = 100) -> list[dict]:
        out = []
        for b in reversed(self.state["meta"]["imports"][-limit:]):
            out.append({k: v for k, v in b.items() if k not in ("created", "updated", "created_shifts")} |
                       {"created_count": len(b.get("created", [])), "updated_count": len(b.get("updated", []))})
        return out

    def status(self) -> dict:
        with self.lock:
            self.refresh()
            counts = {s: len(t["rows"]) for s, t in self.state["xlsx"]["tables"].items()}
            counts["plans"] = len(self.state["plans"]["tasks"])
            counts["users"] = len(self.state["users"]["users"])
            return {"app": APP_NAME, "version": APP_VERSION, "storage": self.storage.describe(),
                    "files": dict(self.DOCS), "revision": self.revision, "last_sync": self.last_sync,
                    "last_error": self.last_error, "counts": counts}

    # -- flat tables -----------------------------------------------------------------------
    @staticmethod
    def _sheet(name: str) -> str:
        s = resolve_module(name)
        if s not in TABLES:
            raise ApiError(404, f"'{name}' is not a table.")
        return s

    @staticmethod
    def _assign(table: dict, sheet: str, row: dict, values: dict) -> list[str]:
        if not isinstance(values, dict):
            raise ApiError(400, "Row must be a JSON object.")
        changed = []
        for k, v in values.items():
            key = k if k in row else match_header(k, table["columns"], sheet)
            if not key or key == ID_COL or key in SYSTEM_COLS or key not in row:
                continue
            v = clean_text(v)
            if row[key] != v:
                row[key] = v
                changed.append(key)
        return changed

    @staticmethod
    def _autofill(tables: dict, sheet: str, row: dict):
        """Runbook entries inherit Customer / Environment from the asset inventory when left blank."""
        if sheet in ("Runbook", "Incidents") and row.get("Asset"):
            asset = next((a for a in tables["Asset_Master"]["rows"]
                          if norm_key(a.get("Device")) == norm_key(row["Asset"])), None)
            if asset:
                row["Customer"] = row.get("Customer") or asset.get("Customer", "")
                if "Environment" in row:
                    row["Environment"] = row.get("Environment") or asset.get("Environment", "")

    def _find_dup(self, table: dict, sheet: str, row: dict, exclude_id: str = ""):
        sigs = set(dup_signatures(sheet, row))
        if not sigs:
            return None
        for r in table["rows"]:  # only rows this person may see: another team's records are not duplicates of yours
            if r[ID_COL] != exclude_id and self._row_visible(sheet, r) and sigs & set(dup_signatures(sheet, r)):
                return r
        return None

    def add_row(self, sheet: str, values: dict, user: str, allow_duplicate: bool = False) -> dict:
        sheet = self._sheet(sheet)

        def fn(state, ctx):
            tables = state["tables"]
            t = tables[sheet]
            row = blank_row(t)
            self._assign(t, sheet, row, values)
            self._stamp_team(sheet, row)
            self._autofill(tables, sheet, row)
            prepare_row(sheet, row, user, True)
            dup = None if allow_duplicate else self._find_dup(t, sheet, row)
            if dup:
                raise ApiError(409, f"Possible duplicate of an existing {TABLES[sheet]['noun']}: "
                                    f"{row_label(sheet, dup)}. Edit that record or save anyway.")
            row[ID_COL] = new_id("R")
            stamp(row, user, None)
            t["rows"].append(row)
            ctx["target"], ctx["details"] = f"{sheet}/{row[ID_COL]}", row_label(sheet, row)
            ctx["history"].append((f"table/{sheet}/{row[ID_COL]}", "Create", None, copy.deepcopy(row)))
            return row

        return self.mutate("xlsx", fn, user, sheet, "Create")

    def update_row(self, sheet: str, rid: str, values: dict, user: str, allow_duplicate: bool = False) -> dict:
        sheet = self._sheet(sheet)

        def fn(state, ctx):
            tables = state["tables"]
            t = tables[sheet]
            row = next((r for r in t["rows"] if r[ID_COL] == rid and self._row_visible(sheet, r)), None)
            if row is None:
                raise ApiError(404, f"That {TABLES[sheet]['noun']} no longer exists.")
            self._check_change(sheet, row, values)
            before = copy.deepcopy(row)
            changed = self._assign(t, sheet, row, values)
            if TEAM_TABLES.get(sheet) and self.scope()["active"]:
                row[TEAM_TABLES[sheet]] = before.get(TEAM_TABLES[sheet], "")
            self._autofill(tables, sheet, row)
            prepare_row(sheet, row, user, False)
            key_fields = {k for ks in TABLES[sheet]["dup_keys"] for k in ks}
            dup = None if allow_duplicate or not key_fields & set(changed) else \
                self._find_dup(t, sheet, row, exclude_id=rid)
            if dup:
                raise ApiError(409, f"This change makes it a duplicate of {row_label(sheet, dup)}. Save anyway?")
            stamp(row, user, before)
            ctx["details"] = f"{row_label(sheet, row)}; changed: {', '.join(changed) or 'nothing'}"
            ctx["history"].append((f"table/{sheet}/{rid}", "Update", before, copy.deepcopy(row)))
            return row

        return self.mutate("xlsx", fn, user, sheet, "Update", f"{sheet}/{rid}")

    def delete_row(self, sheet: str, rid: str, user: str) -> dict:
        sheet = self._sheet(sheet)

        def fn(state, ctx):
            t = state["tables"][sheet]
            row = next((r for r in t["rows"] if r[ID_COL] == rid and self._row_visible(sheet, r)), None)
            if row is None:
                raise ApiError(404, f"That {TABLES[sheet]['noun']} no longer exists.")
            self._check_change(sheet, row, None)
            t["rows"].remove(row)
            ctx["details"] = row_label(sheet, row)
            ctx["history"].append((f"table/{sheet}/{rid}", "Delete", copy.deepcopy(row), None))
            return {"deleted": rid}

        return self.mutate("xlsx", fn, user, sheet, "Delete", f"{sheet}/{rid}")

    def clone_row(self, sheet: str, rid: str, user: str, overrides: dict | None = None) -> dict:
        sheet = self._sheet(sheet)

        def fn(state, ctx):
            t = state["tables"][sheet]
            src = next((r for r in t["rows"] if r[ID_COL] == rid and self._row_visible(sheet, r)), None)
            if src is None:
                raise ApiError(404, f"That {TABLES[sheet]['noun']} no longer exists.")
            self._check_change(sheet, src, None)
            row = copy.deepcopy(src)
            self._stamp_team(sheet, row)
            if sheet == "Asset_Master":
                row["Serial_No"] = ""
            self._assign(t, sheet, row, overrides or {})
            if self._find_dup(t, sheet, row):
                lk = TABLES[sheet]["label_key"]
                row[lk] = f"{row.get(lk, '')} (copy)".strip()
            prepare_row(sheet, row, user, True)
            row[ID_COL] = new_id("R")
            stamp(row, user, None)
            t["rows"].insert(t["rows"].index(src) + 1, row)
            ctx["target"], ctx["details"] = f"{sheet}/{row[ID_COL]}", f"clone of {row_label(sheet, src)}"
            ctx["history"].append((f"table/{sheet}/{row[ID_COL]}", f"Clone of {rid}", None, copy.deepcopy(row)))
            return row

        return self.mutate("xlsx", fn, user, sheet, "Clone")

    def set_table_columns(self, sheet: str, columns, user: str) -> dict:
        sheet = self._sheet(sheet)

        def fn(state, ctx):
            res = set_columns(sheet, state["tables"][sheet], columns)
            ctx["details"] = ", ".join(c["label"] + (" (hidden)" if c["hidden"] else "")
                                       for c in res["columns"] if not c["system"])
            if res["removed"]:
                ctx["details"] += f"; removed: {', '.join(res['removed'])}"
            return res

        return self.mutate("xlsx", fn, user, sheet, "Configure columns", sheet)

    # -- plans / shift log -----------------------------------------------------------------
    def plan_op(self, op: str, p: dict, user: str) -> dict:
        if not isinstance(p, dict):
            raise ApiError(400, "Body must be a JSON object.")
        acc = self.scope()["access"] if msp_access else None
        if acc is None:
            return self.mutate("plans", lambda d, ctx: plan_op(d, op, p, user, ctx), user, "Execution plans", op)

        def guarded(d, ctx):  # isolation: the rules are in msp_access.py
            task = next((t for t in d["tasks"] if t["id"] == p.get("task_id")), None) if p.get("task_id") else None
            try:
                acc.check_plan_op(op, task, p)
            except LookupError:
                raise ApiError(404, "That execution plan no longer exists.")
            except PermissionError as exc:
                raise ApiError(403, str(exc))
            result = plan_op(d, op, p, user, ctx)
            if op in ("task.create", "task.clone"):
                new = next((t for t in d["tasks"] if t["id"] == (result or {}).get("task_id")), None)
                if new is not None:
                    try:
                        new["owner_team"] = acc.plan_default_owner(new)
                    except PermissionError as exc:
                        raise ApiError(403, str(exc))
                    if op == "task.create":
                        new["shared_with"] = []
            return result
        return self.mutate("plans", guarded, user, "Execution plans", op)

    # -- team scope: once teams exist, each team has its own shift log and handovers ------------
    def set_scope(self, active: bool, team=None, teams=None, admin: bool = False, access=None) -> None:
        self._req.scope = {"active": bool(active), "team": team, "teams": list(teams or []), "admin": admin,
                           "access": access if active else None}

    def scope(self) -> dict:
        return getattr(self._req, "scope", None) or {"active": False, "team": None, "teams": [], "admin": False, "access": None}

    # -- isolation bundle A hooks: the rules themselves are in msp_access.py ------------------------
    def _access(self, sheet: str):
        acc = self.scope()["access"]
        return acc if acc is not None and msp_access and sheet in msp_access.SCOPED_SHEETS else None

    def _cols(self, sheet: str) -> list:
        return self.state["xlsx"]["tables"][sheet]["columns"]

    def _check_change(self, sheet: str, row: dict, values) -> None:
        acc = self._access(sheet)
        if acc is None:
            return
        changed = None if values is None else [k for k, v in values.items() if str(v) != str(row.get(k, ""))]
        try:
            acc.check_change(sheet, row, self._cols(sheet), changed)
            owner_key = next((c["key"] for c in self._cols(sheet) if c["key"] == "Owner_Team"), None)
            if values and owner_key and owner_key in (changed or []):
                acc.default_owner(sheet, {owner_key: values[owner_key], "Customer": row.get("Customer")}, self._cols(sheet))
        except LookupError:
            raise ApiError(404, f"That {TABLES[sheet]['noun']} no longer exists.")
        except PermissionError as exc:
            raise ApiError(403, str(exc))

    def _hier_view(self, doc=None) -> dict:
        doc = self.state["shift"] if doc is None else doc
        sc = self.scope()
        if not sc["active"]:
            return {k: v for k, v in doc.items() if k not in ("audit", "teams")}
        return {"years": (doc.get("teams", {}).get(sc["team"]) or {}).get("years", {})} if sc["team"] else {"years": {}}

    def _row_visible(self, sheet: str, row: dict) -> bool:
        if sheet == "Incidents" and self._access("Asset_Master") is not None and self._incident_hidden(row):
            return False
        acc = self._access(sheet)
        if acc is not None:
            return acc.view_row(sheet, row, self._cols(sheet)) is not None
        col, sc = TEAM_TABLES.get(sheet), self.scope()
        return not col or not sc["active"] or (bool(sc["team"]) and row.get(col) == sc["team"])

    def _stamp_team(self, sheet: str, row: dict) -> None:
        acc = self._access(sheet)
        if acc is not None:
            try:
                row["Owner_Team"] = acc.default_owner(sheet, row, self._cols(sheet))
            except PermissionError as exc:
                raise ApiError(403, str(exc))
            return
        col, sc = TEAM_TABLES.get(sheet), self.scope()
        if col and sc["active"]:
            if not sc["team"]:
                raise ApiError(403, "This belongs to a team, and you are not in one. A team lead can add you to their team.")
            row[col] = sc["team"]

    def _hidden_cis(self) -> set:
        """(customer, CI name) of every CI this person may not see (isolation; empty without teams)."""
        acc = self._access("Asset_Master")
        if acc is None:
            return set()
        cols = self._cols("Asset_Master")
        return {(str(r.get("Customer", "")).strip().lower(), str(r.get("Device", "")).strip().lower())
                for r in self.state["xlsx"]["tables"]["Asset_Master"]["rows"] if acc.view_row("Asset_Master", r, cols) is None}

    def _incident_hidden(self, row: dict, hidden=None) -> bool:
        """An incident about a CI you may not see is hidden too: otherwise it would reveal that CI."""
        hidden = self._hidden_cis() if hidden is None else hidden
        return (str(row.get("Customer", "")).strip().lower(), str(row.get("Asset", "")).strip().lower()) in hidden

    def _scoped_tables(self, tables: dict) -> dict:
        if "Incidents" in tables and self._access("Asset_Master") is not None:
            hidden = self._hidden_cis()
            tables["Incidents"]["rows"] = [r for r in tables["Incidents"]["rows"] if not self._incident_hidden(r, hidden)]
        for sheet in (msp_access.SCOPED_SHEETS if msp_access else ()):
            acc = self._access(sheet)
            if acc is not None and sheet in tables:
                cols = tables[sheet]["columns"]
                tables[sheet]["rows"] = [v for v in (acc.view_row(sheet, r, cols) for r in tables[sheet]["rows"]) if v is not None]
        for sheet in TEAM_TABLES:
            if sheet in tables:
                tables[sheet]["rows"] = [r for r in tables[sheet]["rows"] if self._row_visible(sheet, r)]
        return tables

    def history_allowed(self, key: str) -> bool:
        """Version history of a team-owned record is only for that team (the record, or its last version, says which)."""
        pm = re.fullmatch(r"plan/(.+)", str(key or ""))
        acc = self.scope()["access"] if msp_access else None
        if pm and acc is not None:
            task = next((t for t in self.state["plans"]["tasks"] if t["id"] == pm.group(1)), None)
            if task is None:
                hist = self.state["meta"]["history"].get(key) or []
                task = (hist[-1].get("snapshot") if hist else None) or {}
            return acc.plan_mode(task) in ("owner", "shared")
        m = re.fullmatch(r"table/([A-Za-z_]+)/(.+)", str(key or ""))
        if m and m.group(1) == "Incidents" and self._access("Asset_Master") is not None:
            row = next((r for r in self.state["xlsx"]["tables"]["Incidents"]["rows"] if r[ID_COL] == m.group(2)), None)
            if row is None:
                hist = self.state["meta"]["history"].get(key) or []
                row = (hist[-1].get("snapshot") if hist else None) or {}
            return not self._incident_hidden(row)
        if m and self._access(m.group(1)) is not None:
            row = next((r for r in self.state["xlsx"]["tables"][m.group(1)]["rows"] if r[ID_COL] == m.group(2)), None)
            if row is None:
                hist = self.state["meta"]["history"].get(key) or []
                row = (hist[-1].get("snapshot") if hist else None) or {}
            view = self._access(m.group(1)).view_row(m.group(1), row, self._cols(m.group(1)))
            return view is not None and not view.get("_view_only")
        if not m or m.group(1) not in TEAM_TABLES or not self.scope()["active"]:
            return True
        sheet, rid = m.group(1), m.group(2)
        row = next((r for r in self.state["xlsx"]["tables"][sheet]["rows"] if r[ID_COL] == rid), None)
        if row is None:
            hist = self.state["meta"]["history"].get(key) or []
            row = (hist[-1].get("snapshot") if hist else None) or {}
        return self._row_visible(sheet, row)

    def legacy_counts(self) -> dict:
        """Shift-log days and handovers from before teams existed (not visible until a team lead moves them)."""
        days = sum(len(m.get("dates", {})) for y in self.state["shift"].get("years", {}).values() for m in y.get("months", {}).values())
        col = TEAM_TABLES["Shift_Transitions"]
        hand = sum(1 for r in self.state["xlsx"]["tables"]["Shift_Transitions"]["rows"] if not r.get(col))
        return {"shift_days": days, "handovers": hand}

    def claim_legacy(self, team: str, user: str) -> dict:
        """Move the shift log and handovers from before teams into a team. Days the team already has are kept."""
        moved = {"shift_days": 0, "handovers": 0, "kept": 0}

        def sfn(doc, ctx):
            dest = doc.setdefault("teams", {}).setdefault(team, {"years": {}})["years"]
            for y, ynode in list(doc.get("years", {}).items()):
                for m, mnode in list(ynode.get("months", {}).items()):
                    for d, dnode in list(mnode.get("dates", {}).items()):
                        tgt = dest.setdefault(y, {"months": {}})["months"].setdefault(m, {"dates": {}})["dates"]
                        if d in tgt:
                            moved["kept"] += 1
                        else:
                            tgt[d] = dnode
                            moved["shift_days"] += 1
            doc["years"] = {}
            ctx["details"] = f"moved {moved['shift_days']} day(s) into team {team}"
        self.mutate("shift", sfn, user, "Shift log", "Move shift log into team", team)
        col = TEAM_TABLES["Shift_Transitions"]

        def tfn(state, ctx):
            for r in state["tables"]["Shift_Transitions"]["rows"]:
                if not r.get(col):
                    r[col] = team
                    moved["handovers"] += 1
            ctx["target"], ctx["details"] = team, f"{moved['handovers']} handover(s)"
        self.mutate("xlsx", tfn, user, "Shift handovers", "Move handovers into team")
        return moved

    def mutate_hier(self, fn, user, action, target="", details="", team=None):
        """Change the shift log. With teams, fn receives the current team's own tree (same shape as before)."""
        sc = self.scope()
        tid = team or sc["team"]
        if not sc["active"] and not team:
            return self.mutate("shift", fn, user, "Shift log", action, target, details)
        if not tid:
            raise ApiError(403, "Shift logs belong to teams, and you are not in one. A team lead can add you to their team.")

        def scoped(doc, ctx):
            teams = doc.setdefault("teams", {})
            sub = {"years": (teams.get(tid) or {}).get("years", {}), "audit": doc.get("audit", [])}
            result = fn(sub, ctx)
            teams[tid] = {"years": sub.get("years", {}) or {}}  # only the tree itself, never anything nested
            return result
        return self.mutate("shift", scoped, user, "Shift log", action, target, details)

    # -- bulletins -------------------------------------------------------------------------
    def mark_read(self, bid: str, user: str, read: bool = True) -> dict:
        if not any(r[ID_COL] == bid for r in self.state["xlsx"]["tables"]["Bulletins"]["rows"]):
            raise ApiError(404, "That bulletin no longer exists.")

        def fn(meta, ctx):
            users = set(meta["bulletin_reads"].get(bid, []))
            (users.add if read else users.discard)(user)
            meta["bulletin_reads"][bid] = sorted(users)
            return {"read": read}

        return self.mutate("meta", fn, user, "", "")

    # -- version history -------------------------------------------------------------------
    def history(self, key: str) -> list[dict]:
        with self.lock:
            self.refresh()
            return [{k: v for k, v in h.items() if k != "snapshot"}
                    for h in reversed(self.state["meta"]["history"].get(key, []))]

    def restore(self, key: str, v: int, user: str) -> dict:
        with self.lock:
            self.refresh()
            ver = next((h for h in self.state["meta"]["history"].get(key, []) if h["v"] == int(v)), None)
        if ver is None:
            raise ApiError(404, "That version is no longer in the history.")
        snap = copy.deepcopy(ver["snapshot"])
        parts = key.split("/")
        if parts[0] == "table" and len(parts) == 3:
            sheet, rid = parts[1], parts[2]

            def fn(state, ctx):
                t = state["tables"][sheet]
                row = next((r for r in t["rows"] if r[ID_COL] == rid), None)
                before = copy.deepcopy(row) if row else None
                new = blank_row(t)
                new.update({k: v for k, v in snap.items() if k in new})
                new[ID_COL] = rid
                stamp(new, user, before or snap)
                if row:
                    t["rows"][t["rows"].index(row)] = new
                else:
                    t["rows"].append(new)
                ctx["details"] = f"restored version {v}"
                ctx["history"].append((key, f"Restore v{v}", before, copy.deepcopy(new)))
                return new

            return {"row": self.mutate("xlsx", fn, user, sheet, "Restore version", key)}
        if parts[0] == "plan" and len(parts) == 2:
            def pfn(doc, ctx):
                task = normalize_task(snap)
                task["updated_by"], task["updated_at"] = user, now_iso()
                cur = next((t for t in doc["tasks"] if t["id"] == task["id"]), None)
                before = copy.deepcopy(cur) if cur else None
                if cur:
                    doc["tasks"][doc["tasks"].index(cur)] = task
                else:
                    doc["tasks"].append(task)
                ctx["target"], ctx["details"] = task_label(task), f"restored version {v}"
                ctx["history"].append((key, f"Restore v{v}", before, copy.deepcopy(task)))
                return {"task_id": task["id"]}

            return self.mutate("plans", pfn, user, "Execution plans", "Restore version")
        raise ApiError(400, "Unknown history key.")

    # -- import: parse -> preview -> commit -> undo ----------------------------------------
    def _parse_upload(self, filename: str, data: bytes, hint: str | None):
        ext = os.path.splitext(filename.lower())[1]
        sections, ignored = [], []

        def add(source, sec, module):
            module = module or detect_module(sec["headers"]) or hint
            if module is None:
                ignored.append(f"{source} (could not tell what kind of data this is; choose the module)")
                return
            sections.append({"source": source, "module": module, **sec})

        if ext in (".xlsx", ".xlsm"):
            for name, sec in read_workbook(data).items():
                canon = canonical_sheet(name)
                if canon in (COLUMNS_SHEET, AUDIT_SHEET, "instructions") or not sec["rows"]:
                    if canon not in (COLUMNS_SHEET,) and sec["rows"]:
                        ignored.append(name)
                    continue
                add(name, sec, canon if canon in TABLES or canon in ("plans", "shiftlog") else None)
        elif ext in (".csv", ".txt", ".tsv"):
            grid = parse_delimited(decode_text(data))
            add(filename, grid_to_section([[c.strip() for c in r] for r in grid]), hint)
        elif ext == ".json":
            try:
                doc = json.loads(decode_text(data))
            except ValueError as exc:
                raise ApiError(400, f"Invalid JSON: {exc}")

            def rows_section(rows):
                rows = [r for r in rows if isinstance(r, dict)]
                headers = []
                for r in rows:
                    headers += [k for k in r if k not in headers]
                return {"headers": headers, "rows": [{k: cell_str(v) for k, v in r.items()} for r in rows],
                        "header_row": None}

            if isinstance(doc, list):
                add(filename, rows_section(doc), hint)
            elif isinstance(doc, dict) and ("years" in doc and "tasks" not in doc and "tables" not in doc):
                sections.append({"source": filename, "module": "shiftlog", "json": doc})
            elif isinstance(doc, dict):
                tables = doc.get("tables") if isinstance(doc.get("tables"), dict) else {}
                for name, t in tables.items():
                    rows = t.get("rows") if isinstance(t, dict) else t
                    if isinstance(rows, list):
                        try:
                            add(name, rows_section(rows), resolve_module(name))
                        except ApiError:
                            ignored.append(name)
                plans = doc.get("plans", doc) if "tasks" in doc.get("plans", doc) else None
                if isinstance(plans, dict) and isinstance(plans.get("tasks"), list):
                    sections.append({"source": "tasks", "module": "plans", "json_tasks": plans["tasks"]})
                if isinstance(doc.get("hierarchy"), dict):
                    sections.append({"source": "hierarchy", "module": "shiftlog", "json": doc["hierarchy"]})
                for name, rows in doc.items():
                    if name in ("tables", "plans", "hierarchy", "tasks", "audit", "templates") or not isinstance(rows, list):
                        continue
                    try:
                        add(name, rows_section(rows), resolve_module(name))
                    except ApiError:
                        ignored.append(name)
        else:
            raise ApiError(400, "Upload an .xlsx, .csv, .txt or .json file.")
        return sections, ignored

    def import_preview(self, filename: str, data: bytes, module_hint: str, user: str) -> dict:
        hint = resolve_module(module_hint) if module_hint and module_hint != "auto" else None
        sections, ignored = self._parse_upload(filename, data, hint)
        if not sections:
            raise ApiError(400, "Nothing to import was found in the file. " +
                           (f"Skipped: {'; '.join(ignored)}." if ignored else "Download a template to see the format."))
        out = []
        with self.lock:
            self.refresh()
            for sec in sections:
                out.append(self._preview_section(sec, user))
        token = new_id("pv_")
        now = time.monotonic()
        self.previews = {k: v for k, v in self.previews.items() if now - v["at"] < 3600}
        self.previews[token] = {"at": now, "user": user, "filename": filename, "sections": out}
        public = [{k: v for k, v in s.items() if k != "items"} |
                  {"items": [{k: v for k, v in it.items() if not k.startswith("_")} for it in s["items"]]}
                  for s in out]
        return {"token": token, "filename": filename, "sections": public, "ignored": ignored}

    def _preview_section(self, sec: dict, user: str) -> dict:
        module = sec["module"]
        items = []
        if module in TABLES:
            cur = self.state["xlsx"]["tables"][module]
            table, new_cols = section_to_table(module, sec, cur["columns"])
            by_id = {r[ID_COL]: r for r in cur["rows"] if self._row_visible(module, r)}  # never match another team's record
            seen = {}
            for i, raw in enumerate(table["rows"]):
                row = dict(raw)
                item = {"i": i, "label": row_label(module, row) or f"row {i + 1}", "errors": [], "status": "new"}
                try:
                    self._autofill(self.state["xlsx"]["tables"], module, row)
                    prepare_row(module, row, user, True)
                except ApiError as exc:
                    item["status"], item["errors"] = "invalid", [exc.message]
                match = by_id.get(row.get(ID_COL)) or self._find_dup(cur, module, row)
                if item["status"] != "invalid":
                    if match:
                        item.update(status="duplicate", match_id=match[ID_COL], match_label=row_label(module, match),
                                    changes=[default_label(k) if k in schema_keys(module) else k
                                             for k, v in row.items() if v and k not in (ID_COL, *SYSTEM_COLS)
                                             and k in match and match[k] != v])
                    else:
                        clash = next((seen[s] for s in dup_signatures(module, row) if s in seen), None)
                        if clash is not None:
                            item.update(status="duplicate_in_file", errors=[f"Same as row {clash + 1} of this file"])
                    for s in dup_signatures(module, row):
                        seen.setdefault(s, i)
                item["_row"] = row
                items.append(item)
            extra = {"new_columns": [c["label"] for c in new_cols], "_new_cols": new_cols,
                     "header_row": sec.get("header_row")}
        elif module == "plans":
            if "json_tasks" in sec:
                parsed = []
                for t in sec["json_tasks"]:
                    try:
                        task = normalize_task(t, keep_ids=False)
                        apply_task_fields(task, {})
                        parsed.append({"task": task, "errors": [], "rows": sum(len(p["steps"]) for p in task["phases"])})
                    except ApiError as exc:
                        parsed.append({"task": None, "errors": [exc.message], "rows": 0})
            else:
                parsed = rows_to_tasks(sec)
            acc = self.scope()["access"] if msp_access else None
            existing = {task_dup_sig(t): t for t in self.state["plans"]["tasks"] if acc is None or acc.plan_mode(t) == "owner"}
            seen = {}
            for i, p in enumerate(parsed):
                task = p["task"]
                steps = sum(len(ph["steps"]) for ph in task["phases"]) if task else 0
                item = {"i": i, "label": (task_label(task) if task else f"task {i + 1}") + f" ({steps} steps)",
                        "errors": p["errors"], "status": "invalid" if p["errors"] or not task else "new"}
                if item["status"] == "new":
                    sig = task_dup_sig(task)
                    if sig in existing:
                        item.update(status="duplicate", match_id=existing[sig]["id"],
                                    match_label=task_label(existing[sig]), changes=["phases and steps"])
                    elif sig in seen:
                        item.update(status="duplicate_in_file", errors=[f"Same task as item {seen[sig] + 1}"])
                    seen.setdefault(sig, i)
                item["_task"] = task
                items.append(item)
            extra = {"new_columns": [], "header_row": sec.get("header_row")}
        elif module == "shiftlog" and "json" not in sec:
            hier = self._hier_view()
            seen, new_cols = {}, []
            for i, r in enumerate(parse_shiftlog_rows(sec)):
                first = next(iter(r["cells"].values()), "")
                item = {"i": i, "label": f"{r['date'] or '?'} {r['shift'] or '?'} / {r['section'] or '?'}: {first}"[:160],
                        "errors": list(r["errors"]), "status": "invalid" if r["errors"] else "new"}
                if item["status"] == "new":
                    node = _hier_shift(hier, r["date"], r["shift"])
                    tsec = _find_section_by_name(node, r["section"])
                    if node is None:
                        item["label"] += " (creates the shift)"
                    want = {norm_key(k): v for k, v in r["cells"].items()}
                    known = {norm_key(c["name"]) for c in tsec["columns"]} if tsec else \
                        {norm_key(c) for n, cols in SECTION_TEMPLATES if norm_key(n) == norm_key(r["section"]) for c in cols}
                    for k in r["cells"]:
                        tag = f"{r['section']}: {k}"
                        if norm_key(k) not in known and tag not in new_cols:
                            new_cols.append(tag)
                    match = None
                    if tsec:
                        match = next((x for x in tsec["rows"] if r["row_id"] and x["id"] == r["row_id"]), None) or \
                            next((x for x in tsec["rows"] if row_cells_by_name(tsec, x) == want), None)
                    if match:
                        have = row_cells_by_name(tsec, match)
                        item.update(status="duplicate", match_id=match["id"],
                                    match_label=f"existing row in {r['date']} {r['shift']} / {tsec['name']}",
                                    changes=[k for k in r["cells"] if have.get(norm_key(k), "") != r["cells"][k]])
                    else:
                        sig = (r["date"], r["shift"], norm_key(r["section"]), tuple(sorted(want.items())))
                        if sig in seen:
                            item.update(status="duplicate_in_file", errors=[f"Same as line {seen[sig] + 1} of this file"])
                        seen.setdefault(sig, i)
                item["_line"] = r
                items.append(item)
            extra = {"new_columns": new_cols, "header_row": sec.get("header_row")}
        elif module == "shiftlog":
            hier = normalise_hierarchy(sec["json"])
            items.append({"i": 0, "label": f"Replace the shift log ({len(hier['years'])} year(s))", "errors": [],
                          "status": "replace", "_hier": hier})
            extra = {"new_columns": [], "header_row": None}
        else:
            raise ApiError(400, f"Cannot import module '{module}'.")
        counts = {}
        for it in items:
            counts[it["status"]] = counts.get(it["status"], 0) + 1
        return {"module": module, "title": TABLES[module]["title"] if module in TABLES else
                {"plans": "Execution plans", "shiftlog": "Shift log"}[module],
                "source": sec["source"], "counts": counts, "items": items, **extra}

    def import_commit(self, token: str, decisions: dict, default_duplicate: str, user: str) -> dict:
        pv = self.previews.pop(token, None)
        if pv is None:
            raise ApiError(404, "This import preview has expired. Upload the file again.")
        decisions = decisions if isinstance(decisions, dict) else {}
        default_duplicate = default_duplicate if default_duplicate in ("skip", "update", "keep") else "skip"
        batch = {"id": new_id("imp_"), "at": now_iso(), "by": user, "file": pv["filename"], "sections": [], "team": self.scope()["team"],
                 "created": [], "updated": [], "undone": False}

        def choice(si, it):
            d = (decisions.get(str(si)) or {}).get(str(it["i"]))
            if it["status"] == "invalid":
                return "skip"
            if it["status"] == "new":
                return d if d in ("add", "skip") else "add"
            if it["status"] == "replace":
                return d if d in ("replace", "skip") else "replace"
            if it["status"] == "duplicate_in_file":
                return d if d in ("keep", "skip") else "skip"
            return d if d in ("skip", "update", "keep") else default_duplicate

        for si, sec in enumerate(pv["sections"]):
            module, counts = sec["module"], {"added": 0, "updated": 0, "skipped": 0}
            if module in TABLES:
                def fn(state, ctx, sec=sec, si=si, module=module, counts=counts):
                    t = state["tables"][module]
                    plan = [(it, choice(si, it)) for it in sec["items"]]
                    if any(c != "skip" for _, c in plan) and sec["_new_cols"]:
                        known = {c["key"] for c in t["columns"]}
                        for c in sec["_new_cols"]:
                            if c["key"] not in known:
                                t["columns"].append(dict(c))
                                for r in t["rows"]:
                                    r[c["key"]] = ""
                    for it, c in plan:
                        src = it["_row"]
                        if c in ("add", "keep"):
                            row = blank_row(t)
                            row.update({k: v for k, v in src.items() if k in row})
                            self._stamp_team(module, row)
                            row[ID_COL] = new_id("R")
                            stamp(row, user, None)
                            t["rows"].append(row)
                            key = f"table/{module}/{row[ID_COL]}"
                            batch["created"].append(key)
                            ctx["history"].append((key, "Import (new)", None, copy.deepcopy(row)))
                            counts["added"] += 1
                        elif c == "update":
                            target = next((r for r in t["rows"] if r[ID_COL] == it.get("match_id") and self._row_visible(module, r)), None)
                            if target is not None and self._access(module) is not None:
                                try:
                                    self._check_change(module, target, {k: v for k, v in src.items() if v})
                                except ApiError:
                                    target = None  # shared with you, but owned by another team: not yours to update
                            if target is None:
                                counts["skipped"] += 1
                                continue
                            before = copy.deepcopy(target)
                            for k, v in src.items():  # an import never moves a record to another team
                                if v and k in target and k not in (ID_COL, *SYSTEM_COLS) and k != TEAM_TABLES.get(module):
                                    target[k] = v
                            stamp(target, user, before)
                            key = f"table/{module}/{target[ID_COL]}"
                            batch["updated"].append({"key": key, "before": before})
                            ctx["history"].append((key, "Import (update)", before, copy.deepcopy(target)))
                            counts["updated"] += 1
                        else:
                            counts["skipped"] += 1
                    ctx["batch"] = batch["id"]
                    ctx["details"] = f"{pv['filename']} / {sec['source']}: {counts}"
                self.mutate("xlsx", fn, user, module, "Import", pv["filename"])
            elif module == "plans":
                def pfn(doc, ctx, sec=sec, si=si, counts=counts):
                    for it in sec["items"]:
                        c, task = choice(si, it), it["_task"]
                        if c in ("add", "keep"):
                            new = copy.deepcopy(task)
                            new.update({"id": new_id("t_"), "created_by": user, "created_at": now_iso(),
                                        "updated_by": user, "updated_at": now_iso()})
                            acc = self.scope()["access"] if msp_access else None
                            if acc is not None:
                                try:
                                    new["owner_team"] = acc.plan_default_owner(new)
                                except PermissionError as exc:
                                    raise ApiError(403, str(exc))
                            doc["tasks"].append(new)
                            batch["created"].append(f"plan/{new['id']}")
                            ctx["history"].append((f"plan/{new['id']}", "Import (new)", None, copy.deepcopy(new)))
                            counts["added"] += 1
                        elif c == "update":
                            cur = next((t for t in doc["tasks"] if t["id"] == it.get("match_id")), None)
                            acc = self.scope()["access"] if msp_access else None
                            if cur is not None and acc is not None and acc.plan_mode(cur) != "owner":
                                cur = None  # only the owning team updates a plan
                            if cur is None:
                                counts["skipped"] += 1
                                continue
                            before = copy.deepcopy(cur)
                            for f in TASK_FIELDS:
                                if task.get(f):
                                    cur[f] = task[f]
                            cur["phases"], cur["step_columns"] = copy.deepcopy(task["phases"]), task["step_columns"]
                            cur["updated_by"], cur["updated_at"] = user, now_iso()
                            batch["updated"].append({"key": f"plan/{cur['id']}", "before": before})
                            ctx["history"].append((f"plan/{cur['id']}", "Import (update)", before, copy.deepcopy(cur)))
                            counts["updated"] += 1
                        else:
                            counts["skipped"] += 1
                    ctx["batch"] = batch["id"]
                    ctx["details"] = f"{pv['filename']}: {counts}"
                self.mutate("plans", pfn, user, "Execution plans", "Import", pv["filename"])
            elif module == "shiftlog" and sec["items"] and "_line" in sec["items"][0]:
                def lfn(h, ctx, sec=sec, si=si, counts=counts):
                    for it in sec["items"]:
                        c, r = choice(si, it), it["_line"]
                        if c not in ("add", "keep", "update"):
                            counts["skipped"] += 1
                            continue
                        if _hier_shift(h, r["date"], r["shift"]) is None:
                            h_create_node(h, {"level": "shift", "date": r["date"], "shift": r["shift"]}, user)
                            batch.setdefault("created_shifts", []).append(f"{r['date']}|{r['shift']}")
                        node = _hier_shift(h, r["date"], r["shift"])
                        tsec = _find_section_by_name(node, r["section"])
                        if tsec is None:
                            tsec = make_section(_template_section_name(r["section"]), list(r["cells"]), False)
                            node["sections"].append(tsec)
                        by_name = {norm_key(col["name"]): col["id"] for col in tsec["columns"]}
                        for k in r["cells"]:
                            if norm_key(k) not in by_name:
                                col = {"id": new_id("c_"), "name": k}
                                tsec["columns"].append(col)
                                by_name[norm_key(k)] = col["id"]
                        cells = {by_name[norm_key(k)]: v for k, v in r["cells"].items()}
                        key_base = f"shiftrow/{r['date']}/{r['shift']}/{tsec['id']}"
                        target = next((x for x in tsec["rows"] if x["id"] == it.get("match_id")), None)
                        if c == "update" and target is not None:
                            batch["updated"].append({"key": f"{key_base}/{target['id']}", "before": copy.deepcopy(target)})
                            target["cells"].update(cells)
                            target["updated_by"], target["updated_at"] = user, now_iso()
                            counts["updated"] += 1
                        elif c == "update":
                            counts["skipped"] += 1
                        else:
                            row = {"id": new_id("r_"), "cells": cells, "updated_by": user, "updated_at": now_iso()}
                            tsec["rows"].append(row)
                            batch["created"].append(f"{key_base}/{row['id']}")
                            counts["added"] += 1
                    ctx["details"] = f"{pv['filename']}: {counts}"
                self.mutate_hier(lfn, user, "Import", pv["filename"], team=batch.get("team"))
            elif module == "shiftlog":
                it = sec["items"][0]
                if choice(si, it) == "replace":
                    def hfn(h, ctx, it=it):
                        batch["updated"].append({"key": "shiftlog", "before": {k: v for k, v in h.items() if k != "audit"}})
                        audit = h["audit"]
                        h.clear()
                        h.update(copy.deepcopy(it["_hier"]))
                        h["audit"] = audit
                        ctx["details"] = f"{pv['filename']}: shift log replaced"
                    self.mutate_hier(hfn, user, "Import", pv["filename"], team=batch.get("team"))
                    counts["updated"] = 1
                else:
                    counts["skipped"] = 1
            batch["sections"].append({"module": module, "source": sec["source"], **counts})

        def mfn(meta, ctx):
            meta["imports"] = (meta["imports"] + [batch])[-IMPORT_BATCH_CAP:]
        self.mutate("meta", mfn, user, "", "")
        return {k: v for k, v in batch.items() if k not in ("created", "updated", "created_shifts")}

    def import_undo(self, batch_id: str, user: str) -> dict:
        with self.lock:
            self.refresh()
            batch = next((b for b in self.state["meta"]["imports"] if b["id"] == batch_id), None)
        if batch is None:
            raise ApiError(404, "That import is no longer in the list.")
        if batch.get("undone"):
            raise ApiError(409, "That import has already been undone.")
        table_created = [k for k in batch["created"] if k.startswith("table/")]
        table_updated = [u for u in batch["updated"] if u["key"].startswith("table/")]
        plan_created = [k for k in batch["created"] if k.startswith("plan/")]
        plan_updated = [u for u in batch["updated"] if u["key"].startswith("plan/")]
        result = {"removed": 0, "restored": 0}
        if table_created or table_updated:
            def fn(state, ctx):
                for key in table_created:
                    _, sheet, rid = key.split("/")
                    t = state["tables"][sheet]
                    row = next((r for r in t["rows"] if r[ID_COL] == rid), None)
                    if row:
                        t["rows"].remove(row)
                        ctx["history"].append((key, "Undo import", copy.deepcopy(row), None))
                        result["removed"] += 1
                for u in table_updated:
                    _, sheet, rid = u["key"].split("/")
                    t = state["tables"][sheet]
                    row = next((r for r in t["rows"] if r[ID_COL] == rid), None)
                    restored = blank_row(t)
                    restored.update({k: v for k, v in u["before"].items() if k in restored})
                    if row:
                        t["rows"][t["rows"].index(row)] = restored
                    else:
                        t["rows"].append(restored)
                    ctx["history"].append((u["key"], "Undo import", copy.deepcopy(row), copy.deepcopy(restored)))
                    result["restored"] += 1
            self.mutate("xlsx", fn, user, "Import", "Undo import", batch["file"])
        if plan_created or plan_updated:
            def pfn(doc, ctx):
                for key in plan_created:
                    tid = key.split("/", 1)[1]
                    cur = next((t for t in doc["tasks"] if t["id"] == tid), None)
                    if cur:
                        doc["tasks"].remove(cur)
                        ctx["history"].append((key, "Undo import", copy.deepcopy(cur), None))
                        result["removed"] += 1
                for u in plan_updated:
                    tid = u["key"].split("/", 1)[1]
                    cur = next((t for t in doc["tasks"] if t["id"] == tid), None)
                    restored = copy.deepcopy(u["before"])
                    if cur:
                        doc["tasks"][doc["tasks"].index(cur)] = restored
                    else:
                        doc["tasks"].append(restored)
                    ctx["history"].append((u["key"], "Undo import", copy.deepcopy(cur), copy.deepcopy(restored)))
                    result["restored"] += 1
            self.mutate("plans", pfn, user, "Execution plans", "Undo import", batch["file"])
        row_created = [k for k in batch["created"] if k.startswith("shiftrow/")]
        row_updated = [u for u in batch["updated"] if u["key"].startswith("shiftrow/")]
        if row_created or row_updated or batch.get("created_shifts"):
            def rfn(h, ctx):
                def locate(key):
                    _, date, shift, sid, rid = key.split("/")
                    node = _hier_shift(h, date, shift)
                    sec = next((s for s in node["sections"] if s["id"] == sid), None) if node else None
                    return sec, rid
                for key in row_created:
                    sec, rid = locate(key)
                    row = next((x for x in sec["rows"] if x["id"] == rid), None) if sec else None
                    if row:
                        sec["rows"].remove(row)
                        result["removed"] += 1
                for u in row_updated:
                    sec, rid = locate(u["key"])
                    row = next((x for x in sec["rows"] if x["id"] == rid), None) if sec else None
                    if row:
                        row.update(copy.deepcopy(u["before"]))
                        result["restored"] += 1
                for item in batch.get("created_shifts", []):
                    date, shift = item.split("|")
                    node = _hier_shift(h, date, shift)
                    if node is not None and not any(s["rows"] for s in node["sections"]):
                        del h["years"][date[:4]]["months"][date[5:7]]["dates"][date]["shifts"][shift]
                ctx["details"] = f"{batch['file']}: {result}"
            self.mutate_hier(rfn, user, "Undo import", batch["file"], team=batch.get("team"))
        for u in batch["updated"]:
            if u["key"] == "shiftlog":
                def hfn(h, ctx, before=u["before"]):
                    audit = h["audit"]
                    h.clear()
                    h.update(copy.deepcopy(before))
                    h["audit"] = audit
                self.mutate_hier(hfn, user, "Undo import", batch["file"], team=batch.get("team"))
                result["restored"] += 1

        def mfn(meta, ctx):
            for b in meta["imports"]:
                if b["id"] == batch_id:
                    b["undone"], b["undone_by"], b["undone_at"] = True, user, now_iso()
        self.mutate("meta", mfn, user, "", "")
        return result

    # -- templates and exports -------------------------------------------------------------
    def template_file(self, module: str, fmt: str):
        target = resolve_module(module)
        fmt = (fmt or "xlsx").lower()
        if fmt not in ("xlsx", "csv", "txt", "json"):
            raise ApiError(400, "format must be xlsx, csv, txt or json.")
        if target in TABLES:
            cols = [c for c in self.state["xlsx"]["tables"][target]["columns"] if not c["system"]]
            keys, labels = [c["key"] for c in cols], [c["label"] for c in cols]
            sample = seed_tables()[target]["rows"][0]
            examples = [{k: sample.get(k, "") for k in keys}]
            title, name = TABLES[target]["title"], target
            notes = []
            for c in cols:
                req = "Yes" if c["key"] in TABLES[target]["required"] else ""
                allowed = ", ".join(TABLES[target]["choices"].get(c["key"], []))
                if c["key"] in TABLES[target]["dates"]:
                    allowed = "Date: YYYY-MM-DD (DD/MM/YYYY also accepted)"
                notes.append({"Column": c["label"], "Required": req, "Allowed values / format": allowed,
                              "Notes": "Used to detect duplicates" if any(c["key"] in ks for ks in TABLES[target]["dup_keys"]) else ""})
            notes.append({"Column": "(any extra column)", "Notes": "Extra columns are added to the table automatically."})
        elif target == "plans":
            keys = labels = PLAN_IMPORT_COLS
            examples = [{k: r.get(k, "") for k in keys} for r in plans_to_rows(seed_plans()["tasks"])]
            title, name = "Execution plans", "Execution_Plans"
            notes = [{"Column": "Date / Customer / Task_Name / Ticket_ID",
                      "Notes": "Rows with the same Date, Customer and Ticket ID (or Task name) form one task."},
                     {"Column": "Phase", "Notes": "Phase header, e.g. " + ", ".join(PHASE_PRESETS) + ". Numbered 1, 2, 3 in order of appearance."},
                     {"Column": "Step_No", "Notes": "Optional. Dots set the level: 1.1 = step, 1.1.1 = sub-step. Numbers are recalculated."},
                     {"Column": "Planned_Start", "Notes": "Leave blank to chain from the previous step (start + approx time). Fill to override."},
                     {"Column": "Approx_Min", "Notes": "Minutes, or 1h30m, or 1:30."},
                     {"Column": "Checkpoint", "Notes": "Yes = go/no-go sign-off needed before the next phase."}]
        elif target == "shiftlog":
            values = shiftlog_value_columns(self._hier_view())
            keys = labels = ["Date", "Shift", "Section"] + values
            day = dt.date.today().isoformat()
            examples = [{"Date": day, "Shift": "Morning", "Section": name, **SHIFTLOG_EXAMPLES.get(name, {})}
                        for name, _ in SECTION_TEMPLATES]
            title, name = "Shift log", "Shift_Log"
            notes = [{"Column": "Date", "Required": "Yes", "Allowed values / format": "YYYY-MM-DD (DD/MM/YYYY also accepted)"},
                     {"Column": "Shift", "Required": "Yes", "Allowed values / format": ", ".join(SHIFTS)},
                     {"Column": "Section", "Required": "Yes", "Allowed values / format": "One of the sections below, or a "
                      "new name to create a custom section"}]
            notes += [{"Column": f"Section: {n}", "Notes": "Fill these columns: " + ", ".join(cols)} for n, cols in SECTION_TEMPLATES]
            notes.append({"Column": "(any extra column)", "Notes": "Added to that section automatically. Missing shifts "
                          "are created with the seven standard sections."})
        else:
            raise ApiError(400, "Templates are available for tables, execution plans and the shift log.")
        stamp_txt = dt.date.today().isoformat()
        if fmt == "json":
            body = json.dumps([{lb: ex.get(k, "") for k, lb in zip(keys, labels)} for ex in examples],
                              indent=1, ensure_ascii=False).encode("utf-8")
            return body, "application/json; charset=utf-8", f"{name}_template.json"
        if fmt in ("csv", "txt"):
            buf = io.StringIO()
            w = csv.writer(buf, delimiter="," if fmt == "csv" else "\t", lineterminator="\r\n")
            w.writerow(labels)
            for ex in examples:
                w.writerow([ex.get(k, "") for k in keys])
            return (buf.getvalue().encode("utf-8-sig"), "text/csv; charset=utf-8" if fmt == "csv" else
                    "text/plain; charset=utf-8", f"{name}_template.{fmt}")
        wb = openpyxl.Workbook()
        wb.remove(wb.active)
        write_sheet(wb, name, keys, labels, examples, f"{title} import template ({stamp_txt}) - replace the example row(s)")
        write_sheet(wb, "Instructions", ["Column", "Required", "Allowed values / format", "Notes"],
                    ["Column", "Required", "Allowed values / format", "Notes"], notes, f"How to fill in the {title} template")
        buf = io.BytesIO()
        wb.save(buf)
        return (buf.getvalue(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                f"{name}_template.xlsx")

    def export_xlsx(self) -> bytes:
        with self.lock:
            self.refresh()
            tables = self._scoped_tables(copy.deepcopy(self.state["xlsx"]["tables"]))
            acc = self.scope()["access"] if msp_access else None
            rows = plans_to_rows(self.state["plans"]["tasks"] if acc is None
                                 else [v for v in (acc.plan_view(t) for t in self.state["plans"]["tasks"]) if v is not None])
            hier = copy.deepcopy(self._hier_view())
        stamp_txt = dt.datetime.now().strftime("%Y-%m-%d %H:%M")
        hcols, hrows = shiftlog_export_rows(hier)
        extra = [("Execution_Plans", PLAN_EXPORT_COLS, PLAN_EXPORT_COLS, rows, f"{APP_NAME} - Execution plans - {stamp_txt}"),
                 ("Shift_Log", hcols, hcols, hrows, f"{APP_NAME} - Shift log - {stamp_txt}")]
        return write_tables_workbook(tables, list(reversed(self.audit(20000))), extra_sheets=extra)

    def export_json(self) -> bytes:
        with self.lock:
            self.refresh()
            doc = {"exported_at": now_iso(), "app": APP_NAME, "version": APP_VERSION,
                   "tables": copy.deepcopy(self.state["xlsx"]["tables"]),
                   "plans": {k: v for k, v in self.state["plans"].items() if k != "audit"},
                   "hierarchy": {k: v for k, v in self.state["shift"].items() if k != "audit"},
                   "audit": self.audit(20000), "imports": self.imports(IMPORT_BATCH_CAP)}
        if self.org_store is not None:
            try:
                org = self.org_store.load()[0]
                doc["org"] = {"teams": org["teams"], "customers": org["customers"]}
            except (StorageError, ValueError, OSError) as exc:
                doc["org_error"] = f"Teams not included: {exc}"
        if self.roster_store is not None:
            try:
                doc["rosters"] = self.roster_store.load()[0]["rosters"]
            except (StorageError, ValueError, OSError) as exc:
                doc["rosters_error"] = f"Rosters not included: {exc}"
        return json.dumps(doc, indent=1, ensure_ascii=False).encode("utf-8")

    # -- settings, users and trusted hosts (Admin) -------------------------------------------
    def settings(self) -> dict:
        return dict(self.state["users"]["settings"])

    def get_user(self, username: str):
        return self.state["users"]["users"].get(str(username or "").lower())

    def users_list(self) -> list[dict]:
        with self.lock:
            self.refresh()
            return sorted((public_user(u) for u in self.state["users"]["users"].values()), key=lambda u: u["username"].lower())

    @staticmethod
    def _active_admins(users: dict, excluding: str = "") -> int:
        return sum(1 for k, u in users.items() if u["role"] == "Admin" and not u["disabled"] and k != excluding)

    def create_user(self, data: dict, actor: str, setup: bool = False) -> tuple[dict, str]:
        if not isinstance(data, dict):
            raise ApiError(400, "User details must be an object.")
        username = str(data.get("username", "")).strip()
        if not USERNAME_RE.fullmatch(username):
            raise ApiError(400, "User names are 2-120 characters: letters, numbers, dot, dash, underscore or @.")
        role = data.get("role", "Viewer")
        if role not in ROLES:
            raise ApiError(400, f"Role must be one of {', '.join(ROLES)}.")
        auth = data.get("auth", "local")
        if auth not in AUTH_METHODS:
            raise ApiError(400, "Sign-in must be local, entra or both.")
        email = clean_text(data.get("email", ""), 200).strip().lower()
        if auth in ("entra", "both") and not (email or "@" in username):
            raise ApiError(400, "Microsoft 365 users need an email / UPN.")
        password, temp = data.get("password") or "", ""
        if auth in ("local", "both"):
            if password:
                check_password_policy(password, username)
            else:
                password = temp = temp_password()

        def fn(doc, ctx):
            users = doc["users"]
            if username.lower() in users:
                raise ApiError(409, f"User '{username}' already exists.")
            if email and any(u["email"] == email for u in users.values()):
                raise ApiError(409, f"Another user already uses {email}.")
            if setup and users:
                raise ApiError(409, "Setup has already been completed.")
            users[username.lower()] = {
                "username": username, "name": clean_text(data.get("name", ""), 120).strip() or username,
                "email": email, "role": role, "auth": auth, "password": hash_password(password) if password else "",
                "must_change": bool(temp) or bool(data.get("must_change", False)), "disabled": False,
                "created_by": actor, "created_at": now_iso(), "last_login": "",
                "password_changed": now_iso() if password else ""}
            ctx["target"], ctx["details"] = username, f"role {role}, sign-in {auth}"
            return users[username.lower()]

        u = self.mutate("users", fn, actor, "Admin", "Create user")
        return public_user(u), temp

    def update_user(self, username: str, data: dict, actor: str) -> dict:
        def fn(doc, ctx):
            users = doc["users"]
            u = users.get(username.lower())
            if not u:
                raise ApiError(404, "That user no longer exists.")
            changes = []
            if "role" in data and data["role"] != u["role"]:
                if data["role"] not in ROLES:
                    raise ApiError(400, f"Role must be one of {', '.join(ROLES)}.")
                if u["role"] == "Admin" and self._active_admins(users, username.lower()) == 0:
                    raise ApiError(400, "There must always be at least one active Admin.")
                changes.append(f"role {u['role']} -> {data['role']}")
                u["role"] = data["role"]
            if "disabled" in data and bool(data["disabled"]) != u["disabled"]:
                if data["disabled"] and u["role"] == "Admin" and self._active_admins(users, username.lower()) == 0:
                    raise ApiError(400, "You cannot disable the last active Admin.")
                u["disabled"] = bool(data["disabled"])
                changes.append("disabled" if u["disabled"] else "enabled")
            if "name" in data:
                u["name"] = clean_text(data["name"], 120).strip() or u["username"]
            if "email" in data:
                email = clean_text(data["email"], 200).strip().lower()
                if email and any(x["email"] == email and k != username.lower() for k, x in users.items()):
                    raise ApiError(409, f"Another user already uses {email}.")
                u["email"] = email
            if "auth" in data:
                if data["auth"] not in AUTH_METHODS:
                    raise ApiError(400, "Sign-in must be local, entra or both.")
                u["auth"] = data["auth"]
            ctx["target"], ctx["details"] = u["username"], "; ".join(changes) or "details updated"
            return public_user(u)

        res = self.mutate("users", fn, actor, "Admin", "Update user")
        return res

    def set_password(self, username: str, new_password: str, actor: str, must_change: bool = False,
                     current: str | None = None) -> None:
        u = self.get_user(username)
        if not u:
            raise ApiError(404, "That user no longer exists.")
        if current is not None and not verify_password(current, u["password"]):
            raise ApiError(400, "Your current password is not correct.")
        check_password_policy(new_password, u["username"])
        if current is not None and verify_password(new_password, u["password"]):
            raise ApiError(400, "Choose a password different from the current one.")
        hashed = hash_password(new_password)

        def fn(doc, ctx):
            x = doc["users"].get(username.lower())
            if not x:
                raise ApiError(404, "That user no longer exists.")
            x["password"], x["must_change"], x["password_changed"] = hashed, must_change, now_iso()
            if x["auth"] == "entra":
                x["auth"] = "both"
            ctx["target"] = x["username"]
        self.mutate("users", fn, actor, "Admin" if actor.lower() != username.lower() else "Account",
                    "Reset password" if must_change else "Change password")

    def delete_user(self, username: str, actor: str):
        if username.lower() == actor.lower():
            raise ApiError(400, "You cannot delete your own account.")

        def fn(doc, ctx):
            users = doc["users"]
            u = users.get(username.lower())
            if not u:
                raise ApiError(404, "That user no longer exists.")
            if u["role"] == "Admin" and self._active_admins(users, username.lower()) == 0:
                raise ApiError(400, "You cannot delete the last active Admin.")
            del users[username.lower()]
            ctx["target"] = u["username"]
            return {"deleted": u["username"]}
        return self.mutate("users", fn, actor, "Admin", "Delete user")

    def record_login(self, username: str, method: str, ip: str):
        def fn(doc, ctx):
            u = doc["users"].get(username.lower())
            if u:
                u["last_login"] = now_iso()
            ctx["target"], ctx["details"] = username, f"{method} sign-in from {ip}"
        try:
            self.mutate("users", fn, username, "Account", "Sign in")
        except (StorageError, ApiError):
            pass

    def update_settings(self, data: dict, actor: str) -> dict:
        if not isinstance(data, dict):
            raise ApiError(400, "Settings must be an object.")

        def fn(doc, ctx):
            before = doc["settings"]
            doc["settings"] = normalize_settings({**before, **data})
            ctx["details"] = ", ".join(f"{k}={doc['settings'][k]}" for k in data if k in doc["settings"]
                                       and doc["settings"][k] != before.get(k)) or "no change"
            return doc["settings"]
        return self.mutate("users", fn, actor, "Admin", "Change settings", "settings")

    def trust_host(self, host_id: str, info: dict, user: str):
        def fn(doc, ctx):
            doc["known_hosts"][host_id] = {**info, "added_by": user, "added_at": now_iso()}
            ctx["target"], ctx["details"] = host_id, f"{info['type']} {info['fingerprint']}"
        self.mutate("users", fn, user, "Terminal", "Trust SSH host key")

    def untrust_host(self, host_id: str, actor: str):
        def fn(doc, ctx):
            if host_id not in doc["known_hosts"]:
                raise ApiError(404, "That host key is not in the trusted list.")
            del doc["known_hosts"][host_id]
            ctx["target"] = host_id
            return {"removed": host_id}
        return self.mutate("users", fn, actor, "Admin", "Remove trusted host key")

    # -- terminal session index and recordings ---------------------------------------------------
    def log_terminal_event(self, t, event: str, note: str = "", path: str = "", size: int = 0, sha256: str = "",
                           kind: str = "ssh"):
        def fn(meta, ctx):
            lst = meta["terminal_sessions"]
            rec = next((x for x in lst if x["id"] == t.sid), None)
            if rec is None:
                rec = {"id": t.sid, "type": kind, "user": t.user, "role": t.role, "ip": t.ip, "host": t.host,
                       "port": t.port, "username": t.username, "asset": t.asset, "started": t.started_iso,
                       "ended": "", "status": "", "file": "", "bytes": 0, "sha256": "", "note": ""}
                lst.append(rec)
            rec["status"] = event
            if event in ("ended", "refused", "launched"):
                rec["ended"] = now_iso()
            if path:
                rec["file"] = path
            if size:
                rec["bytes"], rec["sha256"] = size, sha256
            if note:
                rec["note"] = note
            del lst[:-5000]
            ctx["audit_meta"] = True
            ctx["target"] = f"{t.username}@{t.host}:{t.port}"
            ctx["details"] = f"{kind} {event}" + (f" ({note})" if note else "")
        try:
            self.mutate("meta", fn, t.user, "Terminal", f"{kind.upper()} {event}")
        except (StorageError, ApiError) as exc:
            self.last_error = f"Terminal log not saved: {exc.message}"

    def terminal_sessions(self, limit: int = 1000) -> list[dict]:
        with self.lock:
            self.refresh()
            return list(reversed(self.state["meta"]["terminal_sessions"][-limit:]))

    def read_recording(self, sid: str) -> tuple[bytes, dict]:
        rec = next((x for x in self.state["meta"]["terminal_sessions"] if x["id"] == sid), None)
        if not rec or not rec.get("file"):
            raise ApiError(404, "There is no recording for that session.")
        if not rec["file"].startswith(RECORDINGS_DIR + "/") or ".." in rec["file"]:
            raise ApiError(400, "Invalid recording path.")
        data, _ = self.storage.read(rec["file"])
        if data is None:
            raise ApiError(404, "The recording file is missing from storage.")
        return data, rec

    # -- backups ----------------------------------------------------------------------------------
    def backup_now(self, user: str, auto: bool = False) -> dict:
        data = self.export_json()
        stamp_txt = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
        name = f"{BACKUPS_DIR}/msp_backup_{stamp_txt}_{secrets.token_hex(3)}{'_auto' if auto else ''}.json"
        self.storage.write(name, data, None)
        keep = self.settings()["backup_keep"]
        removed = []

        def fn(meta, ctx):
            meta["backups"].append({"file": name, "at": now_iso(), "by": user, "bytes": len(data), "auto": auto})
            while len(meta["backups"]) > keep:
                removed.append(meta["backups"].pop(0)["file"])
            ctx["audit_meta"], ctx["target"] = True, name
            ctx["details"] = f"{len(data)} bytes" + (f"; pruned {len(removed)} old backup(s)" if removed else "")
        self.mutate("meta", fn, user, "Admin", "Automatic backup" if auto else "Backup")
        for f in removed:
            try:
                self.storage.delete(f)
            except (StorageError, OSError):
                pass
        return {"file": name, "bytes": len(data)}

    def auto_backup_if_due(self):
        if not self.settings().get("auto_backup"):
            return None
        today = dt.date.today().strftime("%Y%m%d")
        if any(today in b["file"] for b in self.state["meta"]["backups"]):
            return None
        return self.backup_now("system", auto=True)


# --------------------------------------------------------------------------------------
# Terminals: WebSocket transport, SSH sessions (paramiko), recordings, RDP helpers
# --------------------------------------------------------------------------------------
try:  # optional: terminals are disabled with a clear message when paramiko is missing
    import paramiko
except Exception:  # pragma: no cover
    paramiko = None

WS_GUID = "258EAFA5-E914-47DA-95CA-C5AB0DC85B11"
WS_MAX_FRAME = 1 << 20
ANSI_RE = re.compile(r"\x1b\[[0-?]*[ -/]*[@-~]|\x1b\][^\x07\x1b]*(?:\x07|\x1b\\)|\x1b[@-Z\\-_]|[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]")
SECRET_PROMPT_RE = re.compile(r"(password|passphrase|passcode|pin|secret|token|otp|verification code)[^\n]*[:>]\s*$", re.I)
KEY_NAMES = {"\x1b[A": "<Up>", "\x1b[B": "<Down>", "\x1b[C": "<Right>", "\x1b[D": "<Left>", "\t": "<Tab>",
             "\x1b": "<Esc>", "\x1b[3~": "<Del>", "\x1b[H": "<Home>", "\x1b[F": "<End>"}


def ws_accept_key(key: str) -> str:
    return base64.b64encode(hashlib.sha1((key + WS_GUID).encode()).digest()).decode()  # nosec B324 - RFC 6455


class WebSocket:
    """Minimal RFC 6455 server endpoint over an existing HTTP connection."""

    def __init__(self, rfile, wfile, sock):
        self.rfile, self.wfile, self.sock = rfile, wfile, sock
        self.lock = threading.Lock()
        self.closed = False

    def _read(self, n: int) -> bytes:
        data = self.rfile.read(n)
        if data is None or len(data) < n:
            raise ConnectionError("websocket closed")
        return data

    def recv(self):
        """Return (opcode, bytes) of the next data message; handles ping/pong/close and fragments."""
        chunks, first_op = [], None
        while True:
            b1, b2 = self._read(2)
            fin, op, masked, ln = b1 & 0x80, b1 & 0x0F, b2 & 0x80, b2 & 0x7F
            if not masked:
                raise ConnectionError("client frames must be masked")
            if ln == 126:
                ln = struct.unpack(">H", self._read(2))[0]
            elif ln == 127:
                ln = struct.unpack(">Q", self._read(8))[0]
            if ln > WS_MAX_FRAME:
                raise ConnectionError("frame too large")
            mask = self._read(4)
            payload = bytes(b ^ mask[i % 4] for i, b in enumerate(self._read(ln)))
            if op == 0x8:
                self.close()
                raise ConnectionError("websocket closed by client")
            if op == 0x9:
                self._send(0xA, payload)
                continue
            if op == 0xA:
                continue
            if op in (0x1, 0x2):
                first_op, chunks = op, [payload]
            elif op == 0x0 and first_op is not None:
                chunks.append(payload)
                if sum(len(c) for c in chunks) > 4 * WS_MAX_FRAME:
                    raise ConnectionError("message too large")
            else:
                raise ConnectionError("protocol error")
            if fin:
                return first_op, b"".join(chunks)

    def _send(self, op: int, payload: bytes):
        header = bytes([0x80 | op])
        n = len(payload)
        if n < 126:
            header += bytes([n])
        elif n < 65536:
            header += bytes([126]) + struct.pack(">H", n)
        else:
            header += bytes([127]) + struct.pack(">Q", n)
        with self.lock:
            if self.closed and op != 0x8:
                return
            self.wfile.write(header + payload)
            self.wfile.flush()

    def send_json(self, obj):
        try:
            self._send(0x1, json.dumps(obj, ensure_ascii=False).encode("utf-8"))
        except (OSError, ValueError):
            self.closed = True

    def close(self, code: int = 1000):
        if self.closed:
            return
        try:
            self._send(0x8, struct.pack(">H", code))
        except (OSError, ValueError):
            pass
        self.closed = True


class UnknownHostKey(Exception):
    def __init__(self, info):
        super().__init__("unknown host key")
        self.info = info


class HostKeyMismatch(Exception):
    def __init__(self, info, expected):
        super().__init__("host key mismatch")
        self.info, self.expected = info, expected


def host_key_info(key) -> dict:
    fp = "SHA256:" + base64.b64encode(hashlib.sha256(key.asbytes()).digest()).decode().rstrip("=")
    return {"type": key.get_name(), "fingerprint": fp}


def _make_policy(known: dict, host_id: str, trust_fp: str, on_trust):
    class Policy(paramiko.MissingHostKeyPolicy):
        def missing_host_key(self, client, hostname, key):
            info = host_key_info(key)
            have = known.get(host_id)
            if have:
                if have.get("fingerprint") == info["fingerprint"]:
                    return
                raise HostKeyMismatch(info, have)
            if trust_fp and trust_fp == info["fingerprint"]:
                on_trust(info)
                return
            raise UnknownHostKey(info)
    return Policy()


def load_private_key(text: str, passphrase: str | None):
    errors = []
    for cls_name in ("Ed25519Key", "ECDSAKey", "RSAKey", "DSSKey"):
        cls = getattr(paramiko, cls_name, None)
        if cls is None:
            continue
        try:
            return cls.from_private_key(io.StringIO(text), password=passphrase or None)
        except paramiko.PasswordRequiredException:
            raise ApiError(400, "This private key needs its passphrase.")
        except Exception as exc:  # try the next key type
            errors.append(str(exc))
    raise ApiError(400, "The private key could not be read (supported: OpenSSH / PEM RSA, ECDSA, Ed25519).")


class Recorder:
    """Readable transcript of a terminal session: commands (with secrets masked) and screen output."""

    MAX_BYTES = 50 * 1024 * 1024

    def __init__(self, store, path: str, header: str, enabled: bool):
        self.store, self.path, self.enabled = store, path, enabled
        self.parts: list[str] = [header]
        self.size = len(header)
        self.version = None
        self.dirty = enabled
        self.line = ""
        self.tail = ""
        self.hidden = False
        self.truncated = False
        self.lock = threading.Lock()
        self.last_flush = 0.0

    def _add(self, text: str):
        if not self.enabled or self.truncated:
            return
        if self.size + len(text) > self.MAX_BYTES:
            text = "\n[recording size limit reached; later activity not recorded]\n"
            self.truncated = True
        self.parts.append(text)
        self.size += len(text)
        self.dirty = True

    def output(self, text: str):
        clean = ANSI_RE.sub("", text.replace("\r\n", "\n").replace("\r", ""))
        with self.lock:
            self.tail = (self.tail + clean)[-300:]
            self._add(clean)

    def input(self, data: str):
        with self.lock:
            last_line = self.tail.splitlines()[-1] if self.tail.strip() else ""
            if not self.line and not self.hidden and SECRET_PROMPT_RE.search(last_line):
                self.hidden = True
            i = 0
            while i < len(data):
                ch = data[i]
                if ch in "\r\n" or ch == "\x03":
                    if ch == "\x03":
                        self.line += "^C"
                    shown = "[hidden input]" if self.hidden else self.line
                    self._add(f"\n[{dt.datetime.now().strftime('%H:%M:%S')}] >>> {shown}\n")
                    self.line, self.hidden = "", False
                    i += 1
                    continue
                if ch == "\x1b":
                    m = re.match(r"\x1b[\[O][0-9;]*[A-Za-z~]", data[i:])
                    seq = m.group(0) if m else "\x1b"
                    if not self.hidden:
                        self.line += KEY_NAMES.get(seq.replace("\x1bO", "\x1b["), "<Esc>" if seq == "\x1b" else "")
                    i += len(seq)
                    continue
                if ch in ("\x7f", "\x08"):
                    self.line = self.line[:-1]
                elif not self.hidden:
                    self.line += KEY_NAMES.get(ch) or (f"^{chr(ord(ch) + 64)}" if ord(ch) < 32 else ch)
                i += 1

    def note(self, text: str):
        with self.lock:
            self._add(f"\n[{dt.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] # {text}\n")

    def content(self) -> bytes:
        with self.lock:
            return "".join(self.parts).encode("utf-8")

    def flush(self, force: bool = False):
        if not self.enabled or not self.dirty:
            return
        if not force and time.time() - self.last_flush < 5:
            return
        data = self.content()
        try:
            self.version = self.store.storage.write(self.path, data, self.version)
            self.dirty = False
        except ConflictError:
            self.version = self.store.storage.version(self.path)
        except (StorageError, OSError) as exc:
            self.store.last_error = f"Recording not saved: {exc}"
        self.last_flush = time.time()


class TerminalSession:
    def __init__(self, sid, user, role, ip, host, port, username, asset):
        self.sid, self.user, self.role, self.ip = sid, user, role, ip
        self.host, self.port, self.username, self.asset = host, port, username, asset
        self.token = secrets.token_urlsafe(24)
        self.created = time.time()
        self.started_iso = now_iso()
        self.state = "pending"
        self.ws = None
        self.client = None
        self.chan = None
        self.recorder = None
        self.closed = threading.Event()


class TerminalManager:
    def __init__(self, store):
        self.store = store
        self.lock = threading.Lock()
        self.sessions: dict[str, TerminalSession] = {}

    @staticmethod
    def available() -> bool:
        return paramiko is not None

    def _allowed_target(self, host: str) -> bool:
        settings = self.store.settings()
        h = host.strip().lower()
        extra = [x.strip() for x in str(settings.get("terminal_allowed_networks", "")).split(",") if x.strip()]
        for item in extra:
            try:
                if ipaddress.ip_address(h) in ipaddress.ip_network(item, strict=False):
                    return True
            except ValueError:
                if item.lower() == h:
                    return True
        if not settings.get("terminal_restrict_to_inventory", True):
            return True
        for a in self.store.state["xlsx"]["tables"]["Asset_Master"]["rows"]:
            if h in {str(a.get(k, "")).strip().lower() for k in ("Device", "Hostname_FQDN", "IP_Address")} - {""}:
                if self.store._row_visible("Asset_Master", a):  # with teams: owned by, or shared with, your team
                    return True
        return False

    def create(self, ident: dict, ip: str, host: str, port, username: str, asset: str) -> TerminalSession:
        if paramiko is None:
            raise ApiError(501, "SSH terminals need the 'paramiko' package on the server. Run Start_MSP_Ops_Center.bat "
                                "or: pip install paramiko")
        host = str(host or "").strip()
        if not host or len(host) > 255 or not re.fullmatch(r"[A-Za-z0-9._:\-\[\]%]+", host):
            raise ApiError(400, "Enter a valid host name or IP address.")
        try:
            port = int(port or 22)
        except ValueError:
            raise ApiError(400, "Port must be a number.")
        if not 1 <= port <= 65535:
            raise ApiError(400, "Port must be between 1 and 65535.")
        username = str(username or "").strip()
        if not username or len(username) > 100 or re.search(r"[\s\x00-\x1f]", username):
            raise ApiError(400, "Enter the user name to sign in with.")
        if not self._allowed_target(host):
            raise ApiError(403, f"'{host}' is not a configuration item in the CMDB. Add it as a CI first, or ask an admin to "
                                f"allow it in Admin > Settings.")
        settings = self.store.settings()
        with self.lock:
            now = time.time()
            for sid in [s for s, t in self.sessions.items() if t.state == "pending" and now - t.created > 120]:
                del self.sessions[sid]
            mine = [t for t in self.sessions.values() if t.user == ident["username"]]
            if len(mine) >= settings["terminal_max_per_user"]:
                raise ApiError(429, f"You already have {len(mine)} terminal sessions open. Close one first.")
            sess = TerminalSession(new_id("ts_"), ident["username"], ident["role"], ip, host, port, username,
                                   clean_text(asset, 200))
            self.sessions[sess.sid] = sess
        return sess

    def live(self, username: str | None = None) -> list[dict]:
        with self.lock:
            return [{"sid": t.sid, "user": t.user, "host": t.host, "port": t.port, "username": t.username,
                     "state": t.state, "started": t.started_iso, "ip": t.ip}
                    for t in self.sessions.values() if username is None or t.user == username]

    def kill(self, sid: str) -> bool:
        with self.lock:
            t = self.sessions.get(sid)
        if not t:
            return False
        if t.recorder:
            t.recorder.note("Session terminated by an administrator")
        if t.ws:
            t.ws.send_json({"t": "error", "m": "An administrator ended this session."})
        self._shutdown(t)
        return True

    def _shutdown(self, t: TerminalSession):
        t.closed.set()
        for obj in (t.chan, t.client):
            try:
                if obj:
                    obj.close()
            except Exception:  # nosec B110 - already closing; nothing useful to report
                pass
        if t.ws:
            t.ws.close()
            try:
                t.ws.sock.shutdown(socket.SHUT_RDWR)
            except OSError:
                pass

    def attach(self, sid: str, token: str, ident: dict, ws: WebSocket):
        with self.lock:
            t = self.sessions.get(sid)
        if not t or not hmac.compare_digest(t.token, token or "") or t.user != ident["username"] or t.state != "pending":
            ws.send_json({"t": "error", "m": "This terminal link is not valid. Open a new session."})
            ws.close(1008)
            return
        t.state, t.ws = "authenticating", ws
        settings = self.store.settings()
        try:
            ws.sock.settimeout(180)
            self._authenticate(t, ws, settings)
            if t.state != "connected":
                return
            ws.sock.settimeout(settings["terminal_idle_minutes"] * 60)
            self._relay(t, ws)
        except socket.timeout:
            if t.recorder:
                t.recorder.note("Session closed after inactivity")
            ws.send_json({"t": "error", "m": "Session closed after inactivity."})
        except (ConnectionError, OSError):
            pass
        finally:
            self._finish(t)

    def _authenticate(self, t: TerminalSession, ws: WebSocket, settings: dict):
        host_id = f"{t.host.lower()}:{t.port}"
        while True:
            op, raw = ws.recv()
            try:
                msg = json.loads(raw.decode("utf-8"))
            except ValueError:
                continue
            if msg.get("t") != "auth":
                continue
            cols, rows = int(msg.get("cols") or 120), int(msg.get("rows") or 32)
            pkey = None
            if msg.get("key"):
                try:
                    pkey = load_private_key(str(msg["key"]), msg.get("passphrase"))
                except ApiError as exc:
                    ws.send_json({"t": "error", "m": exc.message, "retry": True})
                    continue
            known = self.store.state["users"]["known_hosts"]

            def trust(info):
                self.store.trust_host(host_id, info, t.user)

            client = paramiko.SSHClient()
            client.set_missing_host_key_policy(_make_policy(known, host_id, str(msg.get("trust") or ""), trust))
            ws.send_json({"t": "status", "m": f"Connecting to {t.host}:{t.port} ..."})
            try:
                client.connect(t.host, port=t.port, username=t.username, password=msg.get("password") or None,
                               pkey=pkey, look_for_keys=False, allow_agent=False, timeout=15, banner_timeout=20,
                               auth_timeout=25)
            except UnknownHostKey as exc:
                client.close()
                ws.send_json({"t": "hostkey", "type": exc.info["type"], "fingerprint": exc.info["fingerprint"]})
                continue
            except HostKeyMismatch as exc:
                client.close()
                ws.send_json({"t": "error", "m": f"WARNING: the host key of {t.host} has CHANGED (now "
                                                 f"{exc.info['fingerprint']}, expected {exc.expected.get('fingerprint')}). "
                                                 f"This can mean a rebuilt server or an interception attempt. The "
                                                 f"connection was refused. An admin can remove the old key in Admin > "
                                                 f"Trusted hosts after checking."})
                self.store.log_terminal_event(t, "refused", "host key mismatch")
                return
            except paramiko.AuthenticationException:
                client.close()
                ws.send_json({"t": "error", "m": "Sign-in to the server was rejected. Check the user name and "
                                                 "password or key.", "retry": True})
                continue
            except (paramiko.SSHException, OSError, EOFError) as exc:
                client.close()
                ws.send_json({"t": "error", "m": f"Could not connect to {t.host}:{t.port}: {exc}", "retry": True})
                continue
            t.client = client
            t.chan = client.invoke_shell(term="xterm-256color", width=max(20, min(cols, 500)),
                                         height=max(5, min(rows, 200)))
            ym = dt.date.today().strftime("%Y/%m")
            path = f"{RECORDINGS_DIR}/{ym}/{dt.datetime.now().strftime('%Y%m%d_%H%M%S')}_{re.sub(r'[^A-Za-z0-9_.-]', '_', t.user)}_" \
                   f"{re.sub(r'[^A-Za-z0-9_.-]', '_', t.host)}_{t.sid}.log"
            header = (f"# {APP_NAME} terminal recording\n# session : {t.sid}\n# user    : {t.user} ({t.role}) from {t.ip}\n"
                      f"# target  : {t.username}@{t.host}:{t.port}  asset: {t.asset or '-'}\n# started : {now_iso()}\n"
                      f"# note    : output is shown as displayed; typed input after password prompts is masked\n"
                      f"{'-' * 72}\n")
            t.recorder = Recorder(self.store, path, header, settings.get("terminal_record", True))
            t.state = "connected"
            self.store.log_terminal_event(t, "connected", "", path if t.recorder.enabled else "")
            ws.send_json({"t": "connected", "recorded": t.recorder.enabled})
            return

    def _relay(self, t: TerminalSession, ws: WebSocket):
        def pump():
            decoder = codecs.getincrementaldecoder("utf-8")("replace")
            try:
                while not t.closed.is_set():
                    data = t.chan.recv(32768)
                    if not data:
                        break
                    text = decoder.decode(data)
                    if text:
                        t.recorder.output(text)
                        ws.send_json({"t": "out", "d": text})
                    t.recorder.flush()
            except (OSError, EOFError, paramiko.SSHException):
                pass
            finally:
                if not t.closed.is_set():
                    ws.send_json({"t": "closed", "m": "The server closed the session."})
                    self._shutdown(t)

        th = threading.Thread(target=pump, daemon=True)
        th.start()
        while not t.closed.is_set():
            op, raw = ws.recv()
            try:
                msg = json.loads(raw.decode("utf-8"))
            except ValueError:
                continue
            kind = msg.get("t")
            if kind == "in":
                data = str(msg.get("d", ""))[:65536]
                t.recorder.input(data)
                t.chan.send(data.encode("utf-8"))
            elif kind == "resize":
                try:
                    t.chan.resize_pty(width=max(20, min(int(msg.get("cols", 80)), 500)),
                                      height=max(5, min(int(msg.get("rows", 24)), 200)))
                except (ValueError, paramiko.SSHException):
                    pass
            elif kind == "close":
                break
            t.recorder.flush()

    def _finish(self, t: TerminalSession):
        was_connected = t.recorder is not None
        self._shutdown(t)
        with self.lock:
            self.sessions.pop(t.sid, None)
        if was_connected:
            t.recorder.note("Session ended")
            t.recorder.flush(force=True)
            data = t.recorder.content() if t.recorder.enabled else b""
            self.store.log_terminal_event(t, "ended", "", t.recorder.path if t.recorder.enabled else "",
                                          size=len(data), sha256=hashlib.sha256(data).hexdigest() if data else "")


def rdp_file(host: str, username: str = "", port: int = 3389) -> bytes:
    address = host if port == 3389 else f"{host}:{port}"
    lines = [f"full address:s:{address}", "prompt for credentials:i:1", "authentication level:i:2",
             "screen mode id:i:2", "use multimon:i:0", "redirectclipboard:i:1", "redirectdrives:i:0",
             "redirectprinters:i:0", "audiomode:i:0", "autoreconnection enabled:i:1"]
    if username:
        lines.append(f"username:s:{username}")
    return ("\r\n".join(lines) + "\r\n").encode("utf-16")


# --------------------------------------------------------------------------------------
# Multipart parsing (no dependency on the removed `cgi` module)
# --------------------------------------------------------------------------------------
def parse_multipart(content_type: str, body: bytes):
    m = re.search(r'boundary="?([^";]+)"?', content_type or "", re.I)
    if not m:
        raise ApiError(400, "Upload must be multipart/form-data with a boundary.")
    boundary = b"--" + m.group(1).encode("latin-1")
    fields, files = {}, []
    for part in body.split(boundary)[1:]:
        if part.startswith(b"--"):
            break
        if part.startswith(b"\r\n"):
            part = part[2:]
        head, sep, content = part.partition(b"\r\n\r\n")
        if not sep:
            continue
        if content.endswith(b"\r\n"):
            content = content[:-2]
        head_text = head.decode("utf-8", "replace")
        disp = re.search(r"content-disposition:([^\r\n]*)", head_text, re.I)
        if not disp:
            continue
        name = re.search(r'\bname="([^"]*)"', disp.group(1))
        filename = re.search(r'\bfilename="([^"]*)"', disp.group(1))
        key = name.group(1) if name else ""
        if filename:
            files.append({"field": key, "filename": os.path.basename(filename.group(1)), "data": content})
        else:
            fields[key] = content.decode("utf-8", "replace")
    return fields, files




# --------------------------------------------------------------------------------------
# Configuration
# --------------------------------------------------------------------------------------
@dataclasses.dataclass
class Config:
    host: str = "127.0.0.1"
    port: int = 8080
    storage: str = "local"               # "local" or "sharepoint"
    data_dir: str = "./msp_data"
    refresh_interval: float = 5.0        # seconds between checks for changes made elsewhere
    seed: bool = True                    # create sample records when the workbook does not exist yet
    trust_proxy_user: bool = False       # trust identity headers from Azure App Service Easy Auth / a reverse proxy
    frame_ancestors: str = "'self'"     # e.g. "'self' https://contoso.sharepoint.com" to embed in your site
    quiet: bool = False
    public_url: str = ""                 # e.g. https://msp-ops.contoso.com (needed for Microsoft sign-in)
    tls_cert: str = ""
    tls_key: str = ""
    open_browser: bool = False
    backup_thread: bool = True
    setup_code_file: str = ""
    # Microsoft 365 / Entra ID sign-in (blank values fall back to the SharePoint app registration)
    entra_enabled: bool = False
    entra_tenant_id: str = ""
    entra_client_id: str = ""
    entra_client_secret: str = ""
    # Apache Guacamole (in-browser RDP)
    guac_url: str = ""
    guac_secret: str = ""
    guac_recording_path: str = ""
    # SharePoint / OneDrive (Microsoft Graph)
    sp_tenant_id: str = ""
    sp_client_id: str = ""
    sp_client_secret: str = ""
    sp_site_host: str = ""
    sp_site_path: str = ""
    sp_library: str = ""
    sp_drive_id: str = ""
    sp_onedrive_user: str = ""
    sp_folder: str = "MSP-Runbook"
    graph_base: str = "https://graph.microsoft.com/v1.0"
    auth_base: str = "https://login.microsoftonline.com"


def build_storage(cfg: Config):
    if cfg.storage == "sharepoint":
        return GraphStorage(tenant_id=cfg.sp_tenant_id, client_id=cfg.sp_client_id,
                            client_secret=cfg.sp_client_secret, site_host=cfg.sp_site_host,
                            site_path=cfg.sp_site_path, library=cfg.sp_library, drive_id=cfg.sp_drive_id,
                            onedrive_user=cfg.sp_onedrive_user, folder=cfg.sp_folder,
                            graph_base=cfg.graph_base, auth_base=cfg.auth_base)
    if cfg.storage == "local":
        return LocalStorage(cfg.data_dir)
    raise ValueError("storage must be 'local' or 'sharepoint'")


# --------------------------------------------------------------------------------------
# HTTP layer
# --------------------------------------------------------------------------------------
NO_CACHE = {
    "Cache-Control": "no-cache, no-store, must-revalidate, max-age=0",
    "Pragma": "no-cache",
    "Expires": "0",
}
SESSION_COOKIE = "msp_session"
HOST_RE = re.compile(r"^[A-Za-z0-9.\-]+(:\d{1,5})?$|^\[[0-9A-Fa-f:.]+\](:\d{1,5})?$")


class Handler(BaseHTTPRequestHandler):
    server_version = "MSPOpsCenter"
    sys_version = ""
    timeout = 60
    _body_cache = None
    _cookies_out = None

    # -- plumbing ------------------------------------------------------------------------
    def log_message(self, fmt, *args):
        if not self.server.config.quiet:
            sys.stderr.write(f"[{self.log_date_time_string()}] {self._client_ip()} {fmt % args}\n")

    def _client_ip(self) -> str:
        if self.server.config.trust_proxy_user and self.headers.get("X-Forwarded-For"):
            return self.headers["X-Forwarded-For"].split(",")[0].strip()[:64]
        return self.client_address[0]

    def _is_https(self) -> bool:
        if isinstance(self.connection, ssl.SSLSocket):
            return True
        return self.server.config.trust_proxy_user and self.headers.get("X-Forwarded-Proto", "").lower() == "https"

    def _safe_host(self) -> str:
        h = self.headers.get("Host", "")
        return h if HOST_RE.match(h or "") else ""

    def _headers_common(self):
        for k, v in NO_CACHE.items():
            self.send_header(k, v)
        host = self._safe_host()
        ws = f" ws://{host} wss://{host}" if host else ""
        self.send_header("X-Content-Type-Options", "nosniff")
        self.send_header("Referrer-Policy", "same-origin")
        self.send_header("Permissions-Policy", "camera=(), microphone=(), geolocation=(), payment=()")
        self.send_header("Cross-Origin-Opener-Policy", "same-origin")
        self.send_header("Content-Security-Policy",
                         f"default-src 'self'; script-src 'self' {self._script_src()}; "
                         "style-src 'self' 'unsafe-inline'; font-src 'self' data:; img-src 'self' data:; "
                         f"connect-src 'self'{ws}; object-src 'none'; base-uri 'self'; form-action 'self'; "
                         f"frame-ancestors {self.server.config.frame_ancestors}")
        fa = self.server.config.frame_ancestors.strip()
        if fa in ("'self'", "'none'"):  # the same rule for older browsers that ignore frame-ancestors
            self.send_header("X-Frame-Options", "SAMEORIGIN" if fa == "'self'" else "DENY")
        if self._is_https():
            self.send_header("Strict-Transport-Security", "max-age=31536000; includeSubDomains")
        for c in self._cookies_out or []:
            self.send_header("Set-Cookie", c)

    def _script_src(self) -> str:
        nonce = getattr(self, "_nonce", "")
        return f"'nonce-{nonce}'" if nonce else "'none'"

    def _send(self, status: int, body: bytes, ctype: str, extra: dict | None = None):
        self.send_response(status)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self._headers_common()
        for k, v in (extra or {}).items():
            self.send_header(k, v)
        self.end_headers()
        if self.command != "HEAD":
            self.wfile.write(body)

    def _json(self, payload, status: int = 200):
        self._send(status, json.dumps(payload, ensure_ascii=False).encode("utf-8"), "application/json; charset=utf-8")

    def _redirect(self, location: str):
        self._send(302, b"", "text/plain", {"Location": location})

    def _raw_body(self) -> bytes:
        if self._body_cache is None:
            try:
                length = int(self.headers.get("Content-Length") or 0)
            except ValueError:
                raise ApiError(400, "Invalid Content-Length.")
            if length > MAX_BODY or length < 0:
                raise ApiError(413, f"Request too large (max {MAX_BODY // (1024 * 1024)} MB).")
            self._body_cache = self.rfile.read(length) if length > 0 else b""
        return self._body_cache

    def _body(self) -> dict:
        raw = self._raw_body()
        if not raw:
            return {}
        try:
            data = json.loads(raw.decode("utf-8"))
        except (ValueError, UnicodeDecodeError):
            raise ApiError(400, "Request body must be valid JSON.")
        if not isinstance(data, dict):
            raise ApiError(400, "Request body must be a JSON object.")
        return data

    def _cookie(self, name: str) -> str:
        for part in (self.headers.get("Cookie") or "").split(";"):
            k, _, v = part.strip().partition("=")
            if k == name:
                return v
        return ""

    def _set_session_cookie(self, sid: str | None):
        secure = "; Secure" if self._is_https() else ""
        if sid:
            c = f"{SESSION_COOKIE}={sid}; Path=/; HttpOnly; SameSite=Lax{secure}"
        else:
            c = f"{SESSION_COOKIE}=; Path=/; HttpOnly; SameSite=Lax; Max-Age=0{secure}"
        self._cookies_out = (self._cookies_out or []) + [c]

    # -- identity ------------------------------------------------------------------------
    def _start_session(self, user: dict, method: str) -> dict:
        srv = self.server
        srv.sessions.drop(self._cookie(SESSION_COOKIE))  # never reuse a pre-login session id
        sess = srv.sessions.create(user["username"], self._client_ip(), method)
        self._set_session_cookie(sess["id"])
        srv.store.record_login(user["username"], method, self._client_ip())
        return sess

    def _entra_user(self, identity: str, display: str = "") -> dict:
        store = self.server.store
        u = find_user_by_identity(store.state["users"]["users"], identity)
        if u is None:
            role = store.settings().get("entra_default_role")
            if not role or not identity or "@" not in identity:
                raise ApiError(403, f"The Microsoft 365 account {identity or '(unknown)'} has no access to MSP Ops "
                                    f"Center yet. Ask an administrator to add you.")
            uname = re.sub(r"[^A-Za-z0-9._@\-]", "", identity)[:120]
            store.create_user({"username": uname, "name": display or uname, "email": identity.lower(),
                               "role": role, "auth": "entra"}, "Microsoft sign-in")
            u = store.get_user(uname)
        if u["disabled"]:
            raise ApiError(403, "This account is disabled. Ask an administrator.")
        if u["auth"] == "local":
            raise ApiError(403, "This account signs in with a user name and password, not Microsoft 365.")
        return u

    def _current(self, touch: bool = True):
        srv = self.server
        settings = srv.store.settings()
        sid = self._cookie(SESSION_COOKIE)
        sess = srv.sessions.get(sid, settings, touch)
        if sess:
            u = srv.store.get_user(sess["user"])
            if u and not u["disabled"]:
                return sess, u
            srv.sessions.drop(sid)
        if srv.config.trust_proxy_user:
            principal = (self.headers.get("X-MS-CLIENT-PRINCIPAL-NAME") or self.headers.get("X-Forwarded-Email")
                         or self.headers.get("X-Forwarded-User") or "").strip()
            if principal:
                u = self._entra_user(clean_text(principal, 200), "")
                return self._start_session(u, "proxy"), u
        return None, None

    @staticmethod
    def _ident(sess, u) -> dict:
        return {"name": u["username"], "username": u["username"], "display": u["name"], "role": u["role"],
                "rank": ROLE_RANK[u["role"]], "verified": True, "session": sess}

    # -- verbs ---------------------------------------------------------------------------
    def do_GET(self):
        self._dispatch("GET")

    def do_HEAD(self):
        self._dispatch("GET")

    def do_POST(self):
        self._dispatch("POST")

    def do_PUT(self):
        self._dispatch("PUT")

    def do_DELETE(self):
        self._dispatch("DELETE")

    def _dispatch(self, method: str):
        parts = urllib.parse.urlsplit(self.path)
        path = parts.path.rstrip("/") or "/"
        query = urllib.parse.parse_qs(parts.query)
        try:
            if method == "GET" and path in ("/", "/index.html"):
                self._nonce = secrets.token_urlsafe(18)
                page = HTML_PAGE.replace("<script>", f'<script nonce="{self._nonce}">')
                return self._send(200, page.encode("utf-8"), "text/html; charset=utf-8")
            if path == "/favicon.ico":
                return self._send(204, b"", "image/x-icon")
            if path == "/healthz":
                return self._json({"ok": True, "status": "healthy"})
            if method == "GET" and path.startswith("/static/"):
                return self._static(path[len("/static/"):])
            if method == "GET" and path == "/app":
                index = os.path.join(STATIC_DIR, "app", "index.html")
                if not os.path.isfile(index):
                    raise ApiError(404, "The new interface is not built yet. Copy the static/app folder next to msp_runbook.py.")
                self._nonce = secrets.token_urlsafe(18)  # keeps the script policy well-formed; the page loads only its own files
                with open(index, "rb") as fh:
                    return self._send(200, fh.read(), "text/html; charset=utf-8")
            if method == "GET" and path.startswith("/app/"):
                return self._static("app/" + path[len("/app/"):])
            if path == "/ws/terminal" and method == "GET":
                return self._ws_terminal(query)
            if path.startswith("/auth/entra/") and method == "GET":
                return self._entra_route(path, query)
            if not path.startswith("/api/"):
                raise ApiError(404, "Not found.")
            if path.startswith("/api/auth/"):
                return self._json({"ok": True, **self._auth_api(method, path)})
            sess, u = self._current(touch=self.headers.get("X-Background") != "1")
            if not sess:
                raise ApiError(401, "Please sign in.")
            if method not in ("GET", "HEAD") and not hmac.compare_digest(
                    self.headers.get("X-CSRF-Token", "").encode(), sess["csrf"].encode()):
                raise ApiError(403, "Security check failed (missing CSRF token). Reload the page.")
            if u["must_change"] and path != "/api/me/password":
                raise ApiError(403, "Change your password to continue.")
            need = required_rank(method, path, self._body)
            if ROLE_RANK[u["role"]] < need:
                raise ApiError(403, f"Your role ({u['role']}) cannot do this. It needs {ROLES[need - 1]} or higher.")
            result = self._api(method, path, query, self._ident(sess, u))
            if isinstance(result, tuple) and result[0] == "file":
                _, data, ctype, filename = result
                safe = re.sub(r'[^A-Za-z0-9._\-]', "_", filename)[:150]
                return self._send(200, data, ctype, {"Content-Disposition": f'attachment; filename="{safe}"'})
            return self._json({"ok": True, **result})
        except ApiError as exc:
            self._json({"ok": False, "error": exc.message}, exc.status)
        except (BrokenPipeError, ConnectionResetError, ConnectionAbortedError):
            self.close_connection = True  # the browser went away mid-reply (tab closed, page reloaded): nothing to report
        except Exception:  # pragma: no cover - defensive: never leak internals
            traceback.print_exc()
            try:
                self._json({"ok": False, "error": "Unexpected server error. Details are in the server console."}, 500)
            except OSError:
                self.close_connection = True

    def _static(self, rel: str):
        """Fonts and the terminal component shipped in the static folder; public, read-only, no directory access."""
        ext = os.path.splitext(rel)[1].lower()
        if not STATIC_NAME.fullmatch(rel) or ".." in rel or ext not in STATIC_TYPES:
            raise ApiError(404, "Not found.")
        root = os.path.realpath(STATIC_DIR)
        full = os.path.realpath(os.path.join(root, rel))
        if not full.startswith(root + os.sep) or not os.path.isfile(full):
            raise ApiError(404, "Not found.")
        with open(full, "rb") as fh:
            data = fh.read()
        self.send_response(200)
        self.send_header("Content-Type", STATIC_TYPES[ext])
        self.send_header("Content-Length", str(len(data)))
        self.send_header("Cache-Control", "public, max-age=86400")
        self.send_header("X-Content-Type-Options", "nosniff")
        self.end_headers()
        if self.command != "HEAD":
            self.wfile.write(data)

    # -- sign-in API ---------------------------------------------------------------------
    def _auth_api(self, method: str, path: str) -> dict:
        srv, store = self.server, self.server.store
        settings = store.settings()
        if path == "/api/auth/status" and method == "GET":
            try:
                sess, u = self._current(touch=self.headers.get("X-Background") != "1")
            except ApiError as exc:
                sess, u = None, None
                self._auth_note = exc.message
            out = {"app": APP_NAME, "setup_required": not store.state["users"]["users"], "authenticated": bool(sess),
                   "methods": {"local": True, "entra": srv.entra is not None, "proxy": srv.config.trust_proxy_user}}
            if getattr(self, "_auth_note", ""):
                out["note"] = self._auth_note
            if sess:
                out.update(user=public_user(u), csrf=sess["csrf"], idle_minutes=settings["session_idle_minutes"],
                           roles=ROLES, role_help=ROLE_HELP, terminals=TerminalManager.available(),
                           guacamole=bool(srv.config.guac_url and srv.config.guac_secret))
            return out
        if method != "POST":
            raise ApiError(404, "Not found.")
        if not (self.headers.get("Content-Type") or "").lower().startswith("application/json"):
            raise ApiError(415, "Send JSON.")
        body = self._body()
        if path == "/api/auth/setup":
            if store.state["users"]["users"]:
                raise ApiError(409, "Setup is already complete. Sign in instead.")
            srv.throttle.check("__setup__")
            if not srv.setup_code or not hmac.compare_digest(str(body.get("code", "")).strip().upper().encode(),
                                                             srv.setup_code.encode()):
                srv.throttle.failed("__setup__", settings)
                raise ApiError(403, "The setup code is not correct. It is in FIRST_RUN_SETUP_CODE.txt next to the "
                                    "program, and in the server window.")
            user, _ = store.create_user({"username": body.get("username"), "name": body.get("name"),
                                         "email": body.get("email", ""), "role": "Admin", "auth": "local",
                                         "password": body.get("password")}, "first-run setup", setup=True)
            srv.clear_setup_code()
            sess = self._start_session(store.get_user(user["username"]), "local")
            return {"user": public_user(store.get_user(user["username"])), "csrf": sess["csrf"]}
        if path == "/api/auth/login":
            username, password = str(body.get("username", "")).strip()[:200], str(body.get("password", ""))[:300]
            if not username or not password:
                raise ApiError(400, "Enter your user name and password.")
            ip = self._client_ip()
            srv.throttle.check_ip(ip)
            srv.throttle.check(username)
            u = find_user_by_identity(store.state["users"]["users"], username)
            if not u or not u["password"] or u["auth"] == "entra" or u["disabled"]:
                burn_password_time(password)
                srv.throttle.failed(username, settings)
                srv.throttle.failed_ip(ip)
                raise ApiError(401, "User name or password is not correct.")
            if not verify_password(password, u["password"]):
                srv.throttle.failed(username, settings)
                srv.throttle.failed_ip(ip)
                raise ApiError(401, "User name or password is not correct.")
            srv.throttle.reset(username)
            sess = self._start_session(u, "local")
            return {"user": public_user(u), "csrf": sess["csrf"]}
        if path == "/api/auth/logout":
            srv.sessions.drop(self._cookie(SESSION_COOKIE))
            self._set_session_cookie(None)
            return {}
        raise ApiError(404, "Not found.")

    def _entra_route(self, path: str, query: dict):
        srv = self.server
        if srv.entra is None:
            raise ApiError(404, "Microsoft 365 sign-in is not configured.")
        q = lambda k: (query.get(k) or [""])[0]  # noqa: E731
        if path == "/auth/entra/login":
            return self._redirect(srv.entra.login_url(q("next") or "/"))
        if path == "/auth/entra/callback":
            try:
                if q("error"):
                    raise ApiError(401, "Microsoft sign-in was cancelled or refused: " + clean_text(q("error"), 80))
                claims, nxt = srv.entra.complete(q("code"), q("state"))
                identity = str(claims.get("preferred_username") or claims.get("email") or claims.get("upn") or "")
                u = self._entra_user(clean_text(identity, 200).lower(), clean_text(claims.get("name", ""), 120))
                self._start_session(u, "entra")
                return self._redirect(nxt)
            except ApiError as exc:
                return self._redirect("/?auth_error=" + urllib.parse.quote(exc.message))
        raise ApiError(404, "Not found.")

    # -- terminal websocket -----------------------------------------------------------------
    def _ws_terminal(self, query: dict):
        srv = self.server
        try:
            sess, u = self._current()
        except ApiError:
            sess, u = None, None
        origin = self.headers.get("Origin", "")
        same_origin = bool(origin) and urllib.parse.urlsplit(origin).netloc == self.headers.get("Host", "")
        if (not sess or ROLE_RANK[u["role"]] < 2 or not same_origin
                or self.headers.get("Upgrade", "").lower() != "websocket"
                or self.headers.get("Sec-WebSocket-Version") != "13" or not self.headers.get("Sec-WebSocket-Key")):
            return self._json({"ok": False, "error": "Terminal connection refused."}, 403)
        accept = ws_accept_key(self.headers["Sec-WebSocket-Key"].strip())
        self.wfile.write(("HTTP/1.1 101 Switching Protocols\r\nUpgrade: websocket\r\nConnection: Upgrade\r\n"
                          f"Sec-WebSocket-Accept: {accept}\r\n\r\n").encode())
        self.wfile.flush()
        self.log_message('"GET /ws/terminal" 101 -')
        ws = WebSocket(self.rfile, self.wfile, self.connection)
        try:
            srv.terminals.attach((query.get("sid") or [""])[0], (query.get("token") or [""])[0], self._ident(sess, u), ws)
        finally:
            self.close_connection = True

    # -- API -----------------------------------------------------------------------------
    def _api(self, method: str, path: str, query: dict, ident: dict):
        store: DataStore = self.server.store
        user = ident["name"]
        q = lambda k, d="": (query.get(k) or [d])[0]  # noqa: E731

        self._team_scope(ident)
        pub_ident = {k: v for k, v in ident.items() if k != "session"}
        roster = self._api_roster(method, path, ident, q)
        if roster is not None:
            return roster
        answer = self._api_ai(method, path, ident)
        if answer is not None:
            return answer
        org = self._api_org(method, path, ident)
        if org is not None:
            return org
        extra = self._api_v4(method, path, query, ident, q)
        if extra is not None:
            return extra
        if method == "GET":
            if path == "/api/status":
                return {**store.status(), "user": pub_ident}
            if path == "/api/whoami":
                return {"user": pub_ident}
            if path == "/api/state":
                return {**store.snapshot(user), "user": pub_ident, "status": store.status()}
            if path == "/api/audit":
                acc = store.scope().get("access")
                entries = store.audit(int(q("limit", "500") or 500) * (5 if acc else 1))
                if acc is not None:  # isolation: admins see all, team leads their teams, everyone else themselves
                    entries = [a for a in entries if acc.audit_visible(a)][:int(q("limit", "500") or 500)]
                return {"audit": entries}
            if path == "/api/plans":
                with store.lock:
                    store.refresh()
                    return {"plans": store.plans_view()}
            if path == "/api/hierarchy":
                return {"hierarchy": store.snapshot(user)["hierarchy"]}
            if path == "/api/history":
                if not store.history_allowed(q("key")):
                    raise ApiError(404, "No history for that record.")
                return {"key": q("key"), "versions": store.history(q("key"))}
            if path == "/api/imports":
                return {"imports": store.imports(int(q("limit", "100") or 100))}
            m = re.fullmatch(r"/api/template/([A-Za-z_]+)", path)
            if m:
                data, ctype, fname = store.template_file(m.group(1), q("format", "xlsx"))
                return ("file", data, ctype, fname)
            if path == "/api/export/xlsx":
                stamp_txt = dt.datetime.now().strftime("%Y%m%d_%H%M")
                return ("file", store.export_xlsx(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        f"msp_runbook_export_{stamp_txt}.xlsx")
            if path == "/api/export/json":
                stamp_txt = dt.datetime.now().strftime("%Y%m%d_%H%M")
                return ("file", store.export_json(), "application/json; charset=utf-8",
                        f"msp_runbook_backup_{stamp_txt}.json")
            m = re.fullmatch(r"/api/table/([A-Za-z_]+)", path)
            if m:
                sheet = store._sheet(m.group(1))
                return {"table": store.snapshot(user)["tables"][sheet], "sheet": sheet}

        # ---- flat tables ----------------------------------------------------------------
        m = re.fullmatch(r"/api/table/([A-Za-z_]+)/columns", path)
        if m and method == "PUT":
            return store.set_table_columns(m.group(1), self._body().get("columns"), user)
        m = re.fullmatch(r"/api/table/([A-Za-z_]+)", path)
        if m and method == "POST":
            body = self._body()
            return {"row": store.add_row(m.group(1), body.get("row", body), user, bool(body.get("allow_duplicate")))}
        m = re.fullmatch(r"/api/table/([A-Za-z_]+)/([A-Za-z0-9_\-]+)/clone", path)
        if m and method == "POST":
            return {"row": store.clone_row(m.group(1), m.group(2), user, self._body().get("row"))}
        m = re.fullmatch(r"/api/table/([A-Za-z_]+)/([A-Za-z0-9_\-]+)", path)
        if m and method == "PUT":
            body = self._body()
            return {"row": store.update_row(m.group(1), m.group(2), body.get("row", body), user,
                                            bool(body.get("allow_duplicate")))}
        if m and method == "DELETE":
            return store.delete_row(m.group(1), m.group(2), user)

        # ---- execution plans --------------------------------------------------------------
        if path == "/api/plans/op" and method == "POST":
            body = self._body()
            return store.plan_op(str(body.get("op", "")), body, user)

        # ---- bulletins, history, imports ---------------------------------------------------
        m = re.fullmatch(r"/api/bulletins/([A-Za-z0-9_\-]+)/read", path)
        if m and method == "POST":
            return store.mark_read(m.group(1), user, self._body().get("read", True) is not False)
        if path == "/api/history/restore" and method == "POST":
            body = self._body()
            if not store.history_allowed(str(body.get("key", ""))):
                raise ApiError(404, "No history for that record.")
            return store.restore(str(body.get("key", "")), body.get("v", 0), user)
        if path == "/api/import/preview" and method == "POST":
            fields, files = parse_multipart(self.headers.get("Content-Type", ""), self._raw_body())
            if not files:
                raise ApiError(400, "Attach a file in the 'file' field.")
            return store.import_preview(files[0]["filename"], files[0]["data"], fields.get("module", "auto"), user)
        if path == "/api/import/commit" and method == "POST":
            body = self._body()
            return {"batch": store.import_commit(str(body.get("token", "")), body.get("decisions") or {},
                                                 body.get("default_duplicate", "skip"), user)}
        m = re.fullmatch(r"/api/imports/([A-Za-z0-9_\-]+)/undo", path)
        if m and method == "POST":
            return store.import_undo(m.group(1), user)

        # ---- shift log ----------------------------------------------------------------------
        if path == "/api/hierarchy/node" and method in ("POST", "DELETE"):
            body = self._body()
            label = " ".join(str(body.get(k, "")) for k in ("year", "month", "date", "shift") if body.get(k))
            if method == "POST":
                return store.mutate_hier(lambda h, ctx: h_create_node(h, body, user), user,
                                         f"Create {body.get('level', '')}", label)
            return store.mutate_hier(lambda h, ctx: h_delete_node(h, body), user, f"Delete {body.get('level', '')}", label)
        if path == "/api/hierarchy/section":
            body = self._body()
            if method == "POST":
                def add_section(h, ctx):
                    shift, d, s = _shift_node(h, body)
                    name = clean_text(body.get("name", ""), 120).strip()
                    if not name:
                        raise ApiError(400, "Section name is required.")
                    if any(x["name"].lower() == name.lower() for x in shift["sections"]):
                        raise ApiError(409, f"Section '{name}' already exists in this shift.")
                    cols = body.get("columns") or CUSTOM_SECTION_COLUMNS
                    sec = make_section(name, [clean_text(c, 80) for c in cols if str(c).strip()], False)
                    shift["sections"].append(sec)
                    ctx["target"] = f"{d} {s} / {name}"
                    return {"section": sec}
                return store.mutate_hier(add_section, user, "Add section")
            if method == "DELETE":
                def del_section(h, ctx):
                    shift, d, s = _shift_node(h, body)
                    sec = _section(shift, body.get("section_id"))
                    if sec.get("template"):
                        raise ApiError(400, "Standard template sections cannot be deleted.")
                    shift["sections"].remove(sec)
                    ctx["target"] = f"{d} {s} / {sec['name']}"
                    return {"deleted": sec["id"]}
                return store.mutate_hier(del_section, user, "Delete section")
        if path == "/api/hierarchy/section/columns" and method == "PUT":
            body = self._body()

            def set_cols(h, ctx):
                shift, d, s = _shift_node(h, body)
                sec = _section(shift, body.get("section_id"))
                removed = h_set_columns(sec, body.get("columns"))
                ctx["target"] = f"{d} {s} / {sec['name']}"
                ctx["details"] = "columns: " + ", ".join(c["name"] for c in sec["columns"]) + (
                    f"; removed: {', '.join(removed)}" if removed else "")
                return {"section": sec, "removed": removed}
            return store.mutate_hier(set_cols, user, "Configure columns")
        if path in ("/api/hierarchy/row", "/api/hierarchy/row/clone"):
            body = self._body()

            def locate(h):
                shift, d, s = _shift_node(h, body)
                return _section(shift, body.get("section_id")), d, s

            def find_row(sec):
                row = next((r for r in sec["rows"] if r["id"] == body.get("row_id")), None)
                if row is None:
                    raise ApiError(404, "That row no longer exists.")
                return row

            if path.endswith("/clone") and method == "POST":
                def clone(h, ctx):
                    sec, d, s = locate(h)
                    src = find_row(sec)
                    row = copy.deepcopy(src)
                    row.update({"id": new_id("r_"), "updated_by": user, "updated_at": now_iso()})
                    row.pop("carried_from", None)
                    row["cells"].update(_cells_from_input(sec, body.get("cells")))
                    sec["rows"].insert(sec["rows"].index(src) + 1, row)
                    ctx["target"] = f"{d} {s} / {sec['name']} / {row['id']}"
                    return {"row": row}
                return store.mutate_hier(clone, user, "Clone row")
            if method == "POST":
                def add(h, ctx):
                    sec, d, s = locate(h)
                    row = {"id": new_id("r_"), "cells": _cells_from_input(sec, body.get("cells")),
                           "updated_by": user, "updated_at": now_iso()}
                    sec["rows"].append(row)
                    ctx["target"] = f"{d} {s} / {sec['name']} / {row['id']}"
                    return {"row": row}
                return store.mutate_hier(add, user, "Add row")
            if method == "PUT":
                def edit(h, ctx):
                    sec, d, s = locate(h)
                    row = find_row(sec)
                    row["cells"].update(_cells_from_input(sec, body.get("cells")))
                    row["updated_by"], row["updated_at"] = user, now_iso()
                    ctx["target"] = f"{d} {s} / {sec['name']} / {row['id']}"
                    return {"row": row}
                return store.mutate_hier(edit, user, "Edit row")
            if method == "DELETE":
                def delete(h, ctx):
                    sec, d, s = locate(h)
                    row = find_row(sec)
                    sec["rows"].remove(row)
                    ctx["target"] = f"{d} {s} / {sec['name']} / {row['id']}"
                    return {"deleted": row["id"]}
                return store.mutate_hier(delete, user, "Delete row")

        raise ApiError(404, f"No API route for {method} {path}.")



    # -- v4 routes: account, terminals, RDP, shift cloning, administration -----------------------
    # ---- team scope (isolation) -------------------------------------------------------------------
    def _org_state(self) -> dict | None:
        srv = self.server
        if srv.org is None:
            return None
        with srv.ai_lock:
            at, state = getattr(srv, "_org_cache", (0.0, None))
            if state is None or time.monotonic() - at > 3:
                state = srv.org.load()[0]
                srv._org_cache = (time.monotonic(), state)
            return state

    def _team_scope(self, ident: dict) -> None:
        """Once teams exist, each request sees one team's shift log and handovers: the team chosen with the
        X-Team header if the person belongs to it, otherwise their first team. Administrators belong to none."""
        store, state = self.server.store, self._org_state()
        if not state or not state["teams"]:
            store.set_scope(False)
            return
        admin = ident.get("role") == "Admin"
        mine = [] if admin else msp_org.teams_of(state, ident["username"])
        mine.sort(key=lambda t: state["teams"][t]["name"].lower())
        want = self.headers.get("X-Team", "")
        team = want if want in mine else (mine[0] if mine else None)
        access = msp_access.Access(state, ident["username"], ident.get("role", "")) if msp_access else None
        store.set_scope(True, team, [{"id": t, "name": state["teams"][t]["name"]} for t in mine], admin, access)

    # ---- teams, customers and projects (msp_org.py) ----------------------------------------------
    def _org_users(self) -> dict:
        return {k: public_user(u) for k, u in self.server.store.state["users"]["users"].items()}

    def _api_org(self, method: str, path: str, ident: dict):
        if not path.startswith("/api/org"):
            return None
        srv, O = self.server, msp_org
        if srv.org is None:
            raise ApiError(404, "Teams are not installed: put msp_org.py next to msp_runbook.py and restart.")
        actor = {"user": ident["username"], "role": ident["role"]}
        with srv.store.lock:
            srv.store.refresh()
            users = self._org_users()

        def change(action, target, fn):
            res = srv.org.mutate(actor, action, target, fn)
            srv._org_cache = (0.0, None)

            def audit(meta, ctx):
                ctx["audit_meta"] = True
            try:
                srv.store.mutate("meta", audit, actor["user"], "Teams", action, target)
            except (StorageError, ApiError) as exc:
                srv.store.last_error = f"Team audit entry not saved: {exc.message}"
            return res
        try:
            if path == "/api/org" and method == "GET":
                return O.view(srv.org.load()[0], actor, users)
            if path == "/api/org/claim-legacy" and method == "POST":
                sc = srv.store.scope()
                state = srv.org.load()[0]
                if not sc["team"] or not O.leads_team(state, actor["user"], sc["team"]):
                    raise ApiError(403, "Only a lead of the team you are viewing can move the older shift log into it.")
                return {"moved": srv.store.claim_legacy(sc["team"], actor["user"])}
            if path == "/api/org/people" and method == "GET":
                state = srv.org.load()[0]
                if not (O.is_admin(actor) or any(O.leads_team(state, actor["user"], t) for t in state["teams"])):
                    raise ApiError(403, "Only team leads and administrators can list people.")
                return {"people": [{"username": u["username"], "name": u["name"], "role": u["role"], "disabled": u["disabled"]}
                                   for u in sorted(users.values(), key=lambda x: x["name"].lower()) if u["role"] != "Admin"]}
            if path == "/api/org/history" and method == "GET":
                if not O.is_admin(actor):
                    raise ApiError(403, "Only an administrator can see the team setup history.")
                return {"versions": srv.org.history()[-30:]}
            if path == "/api/org/restore" and method == "POST":
                srv.org.restore(actor, int(self._body().get("index", -1)))
                return {"ok": True}
            body = self._body() if method in ("POST", "PUT") else {}
            if path == "/api/org/teams" and method == "POST":
                t = change("Create team", str(body.get("name", "")), lambda st: O.create_team(st, actor, body, users))
                return {"team": {"id": t["id"], "name": t["name"]}}
            m = re.fullmatch(r"/api/org/teams/([A-Za-z0-9_]+)(?:/members(?:/([A-Za-z0-9._@\-]+))?)?", path)
            if m:
                tid, who = m.group(1), m.group(2)
                is_members = "/members" in path
                if not is_members and method == "PUT":
                    t = change("Update team", tid, lambda st: O.update_team(st, actor, tid, body, users))
                    return {"team": {"id": tid, "name": t["name"]}}
                if not is_members and method == "DELETE":
                    change("Delete team", tid, lambda st: O.delete_team(st, actor, tid))
                    return {"deleted": tid}
                if is_members and not who and method == "POST":
                    state = srv.org.load()[0]
                    if tid not in state["teams"]:
                        raise ApiError(404, "Team not found.")
                    if not (O.is_admin(actor) or O.leads_team(state, actor["user"], tid)):
                        raise ApiError(403, "Only the team's leads or an administrator can add members.")
                    temp = ""
                    if isinstance(body.get("create"), dict):  # a team lead creates a new Engineer or Viewer account
                        new = dict(body["create"])
                        if new.get("role", "Engineer") not in ("Engineer", "Viewer"):
                            raise ApiError(400, "Team leads create Engineers or Viewers. An administrator grants Lead or Admin.")
                        new.setdefault("role", "Engineer")
                        created, temp = srv.store.create_user(new, actor["user"])
                        username = created["username"]
                        users = self._org_users()
                    else:
                        username = str(body.get("username") or "")
                    res = change("Add team member", f"{username} to {tid}", lambda st: O.add_member(st, actor, tid, username, users))
                    return {"member": res, "temp_password": temp}
                if is_members and who and method == "PUT":
                    projects = change("Assign projects", f"{who} in {tid}",
                                      lambda st: O.set_member_projects(st, actor, tid, who, body.get("projects") or []))
                    return {"projects": projects}
                if is_members and who and method == "DELETE":
                    change("Remove team member", f"{who} from {tid}", lambda st: O.remove_member(st, actor, tid, who))
                    return {"removed": who}
            if path == "/api/org/customers" and method == "POST":
                c = change("Create customer", str(body.get("name", "")), lambda st: O.create_customer(st, actor, body))
                return {"customer": {"id": c["id"], "name": c["name"]}}
            m = re.fullmatch(r"/api/org/customers/([A-Za-z0-9_]+)(?:/projects(?:/([A-Za-z0-9_]+))?)?", path)
            if m:
                cid, pid = m.group(1), m.group(2)
                if "/projects" not in path:
                    if method == "PUT":
                        c = change("Update customer", cid, lambda st: O.update_customer(st, actor, cid, body))
                        return {"customer": {"id": cid, "name": c["name"]}}
                    if method == "DELETE":
                        change("Delete customer", cid, lambda st: O.delete_customer(st, actor, cid))
                        return {"deleted": cid}
                elif method in ("POST", "PUT"):
                    p = change("Save project", f"{cid}/{pid or 'new'}", lambda st: O.save_project(st, actor, cid, body, pid))
                    return {"project": p}
                elif method == "DELETE" and pid:
                    change("Delete project", f"{cid}/{pid}", lambda st: O.delete_project(st, actor, cid, pid))
                    return {"deleted": pid}
        except O.OrgError as exc:
            raise ApiError(exc.status, exc.message)
        raise ApiError(404, f"No API route for {method} {path}.")

    # ---- built-in assistant (msp_ai.py) ---------------------------------------------------------
    def _ai_data(self) -> dict:
        """What the assistant may read: operational records only (never users, sessions, settings or recordings)."""
        snap = self.server.store.snapshot()
        rosters, rver = [], None
        if self.server.rosters is not None:
            state, rver = self.server.rosters.load()
            rosters = list(state["rosters"].values())
        sc = self.server.store.scope()
        org = None
        if self.server.org is not None and msp_org:
            ident_user = getattr(self, "_ai_user", None) or {}
            org = msp_org.view(self.server.org.load()[0], ident_user, self._org_users()) if ident_user else None
        # Since isolation each person sees different records, so the search index is cached per person and team.
        who = (getattr(self, "_ai_user", None) or {}).get("user", "")
        return {"tables": snap["tables"], "plans": snap["plans"], "hierarchy": snap["hierarchy"], "rosters": rosters, "org": org,
                "signature": f"{snap['revision']}:{rver}:{who}:{sc.get('team')}:{bool(sc.get('access'))}"}

    def _ai_limit(self, user: str, per_minute: int = 20) -> None:
        now = time.time()
        with self.server.ai_lock:
            hits = [t for t in self.server.ai_hits.get(user, []) if now - t < 60]
            if len(hits) >= per_minute:
                raise ApiError(429, "That is a lot of questions in one minute. Wait a moment and ask again.")
            hits.append(now)
            self.server.ai_hits[user] = hits

    def _api_ai(self, method: str, path: str, ident: dict):
        if not (path.startswith("/api/ai/") or path == "/api/admin/ai/test"):
            return None
        srv = self.server
        if srv.assistant is None:
            raise ApiError(404, "The assistant is not installed: put msp_ai.py next to msp_runbook.py and restart.")
        settings = srv.store.settings()
        srv.assistant.settings = settings
        try:
            if path == "/api/ai/status" and method == "GET":
                mode = settings["ai_mode"]
                return {"mode": mode, "model": settings["ai_model"] if mode == "model" else "", "roster": srv.rosters is not None}
            if path == "/api/admin/ai/test" and method == "POST":
                given = self._body().get("settings")
                trial = normalize_settings({**settings, **(given if isinstance(given, dict) else {}), "ai_mode": "model"})
                return {"status": msp_ai.Assistant(trial).status()}
            if path == "/api/ai/ask" and method == "POST":
                self._ai_limit(ident["username"])
                self._ai_user = {"user": ident["username"], "role": ident["role"]}
                out = srv.assistant.ask(self._body().get("question"), self._ai_user, self._ai_data())

                def fn(meta, ctx):
                    ctx["audit_meta"] = True
                try:
                    srv.store.mutate("meta", fn, ident["username"], "Assistant", "Ask", clean_text(out["question"], 160),
                                     f"{out['mode']} answer, {len(out['sources'])} source(s), {'found' if out['found'] else 'not found'}")
                except (StorageError, ApiError) as exc:
                    srv.store.last_error = f"Assistant audit entry not saved: {exc.message}"
                return {"answer": out}
        except msp_ai.AssistantError as exc:
            raise ApiError(exc.status, exc.message)
        raise ApiError(404, f"No API route for {method} {path}.")

    # ---- shift roster (msp_roster.py) ---------------------------------------------------------
    def _api_roster(self, method: str, path: str, ident: dict, q):
        if not path.startswith("/api/roster"):
            return None
        rs = self.server.rosters
        if rs is None:
            raise ApiError(404, "The shift roster is not installed: put msp_roster.py next to msp_runbook.py and restart.")
        actor = {"user": ident["username"], "name": ident.get("display", ""), "role": ident["role"]}
        try:
            return self._roster_route(rs, method, path, actor, q)
        except msp_roster.RosterError as exc:
            raise ApiError(exc.status, exc.message)

    def _roster_guard(self, team: str, changing: bool) -> None:
        """Isolation: everyone sees every roster, but only a lead of the roster's team changes it (msp_access)."""
        acc = self.server.store.scope().get("access")
        if changing and acc is not None and not acc.can_manage_roster(team):
            raise ApiError(403, f"Only a lead of the {team or 'roster'} team can change this roster.")

    def _roster_audit(self, actor: dict, action: str, target: str, details: str = "") -> None:
        """Mirror roster changes into the app's audit trail (Activity page)."""
        def fn(meta, ctx):
            ctx["audit_meta"] = True
        try:
            self.server.store.mutate("meta", fn, actor["user"], "Roster", action, target, details)
        except (StorageError, ApiError) as exc:  # the roster change itself is already saved
            self.server.store.last_error = f"Roster audit entry not saved: {exc.message}"

    @staticmethod
    def _roster_summary(r: dict) -> dict:
        return {"id": r["id"], "name": r["name"], "team": r["team"], "layout": r.get("layout", "by_person"),
                "template": r.get("template", "standard"), "members": len(r.get("members", [])),
                "pending": sum(1 for x in r.get("leaves", []) + r.get("swaps", []) if x.get("status") == "pending")}

    def _roster_view(self, rs, r: dict, start: str, mode: str, actor: dict) -> dict:
        R = msp_roster
        cal = R.calendar(r, start or dt.date.today().isoformat(), mode or "week")
        a, b = cal["start"], cal["end"]
        sched = R.generate(r, a, b)
        days = sorted(sched)
        hols = R.holidays_for(r, a, b)
        shifts = r.get("shifts", [])
        others = [o for o in rs.rosters() if o["id"] != r["id"]]
        names = {m["id"]: m["name"] for m in r.get("members", [])}
        need = {d: {s["code"]: (R.required_staff(s, d, hols) if (sched[d] and R.parse_date(d).weekday() in s["days"]) else None)
                    for s in shifts} for d in days}
        grid = [{"id": m["id"], "name": m["name"], "role": m.get("role", ""), "user": m.get("user", ""),
                 "cells": [sched[d].get(m["id"]) or {"code": "", "source": ""} for d in days]} for m in r.get("members", [])]
        leaves = [dict(x, kind="leave", who=names.get(x["member_id"], "?")) for x in r.get("leaves", [])]
        swaps = [dict(x, who=f"{names.get(x['member_a'], '?')} and {names.get(x['member_b'], '?')}") for x in r.get("swaps", [])]
        order = {"awaiting_partner": 0, "pending": 1, "approved": 2, "rejected": 3, "declined": 4, "cancelled": 5}
        requests = sorted(leaves + swaps, key=lambda x: (order.get(x.get("status"), 9), x.get("requested_at", "")), reverse=False)
        me = next((m["id"] for m in r.get("members", []) if m.get("user") and R.norm(m["user"]) == R.norm(actor["user"])), None)
        config = {k: v for k, v in r.items() if k not in ("leaves", "swaps", "overrides")}
        return {"rosters": [self._roster_summary(x) for x in rs.rosters()], "roster": self._roster_summary(r), "config": config,
                "start": a, "end": b, "mode": cal["mode"], "days": cal["days"], "need": need, "grid": grid,
                "checks": (checks := R.check(r, sched, others)), "checks_grouped": R.group_issues(checks),
                "holidays": hols, "requests": requests[-300:],
                "history": rs.history(r["id"])[-30:], "me": me, "templates": R.TEMPLATES,
                "regions": R.HOLIDAY_REGIONS, "holidays_available": R.holidays_available(),
                "can": {"edit": ROLE_RANK[actor["role"]] >= 3 and (self.server.store.scope().get("access") is None
                                                                   or self.server.store.scope()["access"].can_manage_roster(r.get("team", ""))),
                        "request": ROLE_RANK[actor["role"]] >= 2,
                        "admin": ROLE_RANK[actor["role"]] >= 4}, "user": actor["user"]}

    def _roster_check_shared(self, rs, roster_id: str, changes: dict) -> None:
        """Members stay dedicated to one roster unless that roster allows sharing."""
        cur = rs.get(roster_id)
        merged = dict(cur, **changes)
        if (merged.get("rules") or {}).get("allow_shared_members"):
            return
        mine = {msp_roster.norm(m.get("user") or m["name"]): m["name"] for m in merged.get("members", [])}
        for o in rs.rosters():
            if o["id"] == roster_id:
                continue
            for m in o.get("members", []):
                key = msp_roster.norm(m.get("user") or m["name"])
                if key in mine:
                    raise ApiError(409, f"{mine[key]} already belongs to the roster '{o['name']}'. "
                                        "Allow shared members in this roster's rules to add them here too.")

    def _roster_route(self, rs, method: str, path: str, actor: dict, q):
        R = msp_roster
        if path == "/api/roster" and method == "GET":
            return {"rosters": [self._roster_summary(x) for x in rs.rosters()], "templates": R.TEMPLATES,
                    "regions": R.HOLIDAY_REGIONS, "holidays_available": R.holidays_available()}
        if path == "/api/roster/subdivisions" and method == "GET":
            return {"subdivisions": R.holiday_subdivisions(q("region"))}
        if path == "/api/roster/now" and method == "GET":
            everyone = rs.rosters()
            mine = next((x for x in everyone if any(m.get("user") and R.norm(m["user"]) == R.norm(actor["user"])
                                                    for m in x.get("members", []))), None)
            r = mine or (everyone[0] if everyone else None)
            if r is None:
                return {"roster": None}
            day = R.calendar(r, dt.date.today(), "day")["days"][0]
            today_by_log = {}
            for s_ in r.get("shifts", []):
                if day["has_data"] and dt.date.today().weekday() in s_["days"]:
                    today_by_log.setdefault(s_.get("log_shift") or s_["name"], []).extend(day["shifts"].get(s_["code"]) or [])
            return {"roster": self._roster_summary(r), "now": R.on_now(r), "today": today_by_log,
                    "shifts": [{"code": s["code"], "name": s["name"], "start": s["start"], "end": s["end"], "color": s.get("color", "")}
                               for s in r.get("shifts", [])]}
        if path == "/api/roster/mine" and method == "GET":
            days = max(1, min(int(q("days", "7") or 7), 31))
            start = dt.date.today()
            end = start + dt.timedelta(days=days - 1)
            out = []
            for r in rs.rosters():
                for m in r.get("members", []):
                    if m.get("user") and R.norm(m["user"]) == R.norm(actor["user"]):
                        out += [dict(x, roster=r["name"], roster_id=r["id"]) for x in R.person_view(r, m["id"], start, end)]
            todo = []
            for r in rs.rosters():
                names = {m["id"]: m["name"] for m in r.get("members", [])}
                for x in R.waiting_for(r, actor):
                    what = (f"Leave: {names.get(x.get('member_id'), '?')}, {x['start']}" + (f" to {x['end']}" if x["end"] != x["start"] else "")
                            if x["id"].startswith("lv_") else
                            f"{x['kind'].title()}: {names.get(x['member_a'], '?')} with {names.get(x['member_b'], '?')}, {x['date_a']}"
                            + (f" / {x['date_b']}" if x.get("date_b") and x["date_b"] != x["date_a"] else ""))
                    todo.append({"roster": r["name"], "roster_id": r["id"], "id": x["id"], "action": x["action"],
                                 "kind": "leave" if x["id"].startswith("lv_") else "swap", "what": what,
                                 "by": x.get("requested_by", ""), "warnings": x.get("warnings") or []})
            return {"shifts": sorted(out, key=lambda x: (x["date"], x.get("start", ""))), "todo": todo}
        if path == "/api/roster/template" and method == "GET":
            data, ctype, fname = R.template_file(q("layout", "by_shift"), q("format", "xlsx"), month=q("month") or None)
            return ("file", data, ctype, fname)
        if path == "/api/roster/import" and method == "POST":
            body = self._body()
            self._roster_guard(body.get("team", ""), True)
            fname = str(body.get("filename") or "roster.xlsx")
            ext = os.path.splitext(fname.lower())[1]
            if ext not in (".xlsx", ".xlsm", ".csv", ".txt", ".tsv"):
                raise ApiError(400, "Choose an Excel (.xlsx), CSV or tab-separated text file.")
            try:
                raw = base64.b64decode(str(body.get("data") or ""), validate=True)
            except (ValueError, TypeError):
                raise ApiError(400, "The file could not be read.")
            if ext in (".xlsx", ".xlsm"):
                if not raw.startswith(b"PK\x03\x04"):
                    raise ApiError(400, "That file is not an Excel workbook. Save it as .xlsx and try again.")
                res = rs.import_new(actor, raw, body.get("team", ""), body.get("name", ""))  # read from memory: no temporary file
            else:
                res = rs.import_new(actor, raw.decode("utf-8-sig", errors="replace"), body.get("team", ""), body.get("name", ""))
            self._roster_audit(actor, "Import roster", res["roster"]["name"], f"{fname}: {res.get('layout', '')}")
            return {"roster": self._roster_summary(res["roster"]), "warnings": res["warnings"], "layout": res.get("layout")}
        if path == "/api/roster" and method == "POST":
            body = self._body()
            self._roster_guard(body.get("team", ""), True)
            r = R.default_roster(body.get("team", ""), body.get("name", ""), body.get("start_date") or dt.date.today().isoformat(),
                                 body.get("template") or "standard")
            if isinstance(body.get("holidays"), dict):
                r["holidays"] = {**r["holidays"], **body["holidays"]}
            r = rs.create(actor, r)
            for m in body.get("members") or []:
                rs.add_member(actor, r["id"], m.get("name", ""), m.get("role") or "Engineer", m.get("user", ""), int(m.get("offset") or 0))
            self._roster_audit(actor, "Create roster", r["name"], f"template {r.get('template')}")
            return {"roster": self._roster_summary(rs.get(r["id"]))}
        if path == "/api/roster/deleted" and method == "GET":
            return {"deleted": rs.deleted()}
        m = re.fullmatch(r"/api/roster/([A-Za-z0-9_\-]+)(?:/([a-z]+))?", path)
        if not m:
            raise ApiError(404, f"No API route for {method} {path}.")
        rid, sub = m.group(1), m.group(2) or ""
        if method == "POST" and sub == "undelete":
            self._roster_guard(next((x.get("team", "") for x in rs.deleted() if x["id"] == rid), ""), True)
            r = rs.undelete(actor, rid)
            self._roster_audit(actor, "Restore deleted roster", r["name"])
            return {"roster": self._roster_summary(r)}
        r = rs.get(rid)
        if method == "GET" and sub == "view":
            return self._roster_view(rs, r, q("start"), q("mode", "week"), actor)
        if method == "GET" and sub == "export":
            start = q("start") or dt.date.today().replace(day=1).isoformat()
            end = q("end") or (R.parse_date(start).replace(day=28) + dt.timedelta(days=4)).replace(day=1).isoformat()
            end = (R.parse_date(end) - dt.timedelta(days=1)).isoformat() if not q("end") else end
            base = "roster_" + (R.norm(r["name"])[:30] or "export")
            if q("format", "xlsx") == "html":
                page = R.export_html(r, start, end, q("layout") or None, history=rs.history(rid))
                return ("file", page.encode("utf-8"), "text/html; charset=utf-8", f"{base}_{start}.html")
            return ("file", R.export_excel(r, start, end), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    f"{base}_{start}.xlsx")
        body = self._body() if method in ("POST", "PUT") else {}
        self._roster_guard(r.get("team", ""), (method == "PUT" and not sub) or (method == "DELETE" and not sub)
                           or sub in ("override", "overrides", "restore", "clone"))
        if method == "PUT" and not sub:
            changes = body.get("changes") if isinstance(body.get("changes"), dict) else {}
            if "team" in changes:
                self._roster_guard(changes.get("team", ""), True)
            if "members" in changes:
                self._roster_check_shared(rs, rid, changes)
            out = rs.update(actor, rid, changes)
            self._roster_audit(actor, "Update roster", out["name"], ", ".join(sorted(changes)))
            return {"roster": self._roster_summary(out)}
        if method == "DELETE" and not sub:
            rs.delete(actor, rid)
            self._roster_audit(actor, "Delete roster", r["name"])
            return {"deleted": rid}
        if method != "POST":
            raise ApiError(405, "Method not allowed.")
        if sub == "leave":
            x = rs.request_leave(actor, rid, body.get("member_id"), body.get("start"), body.get("end"),
                                 body.get("type") or "Annual", body.get("note", ""), body.get("approver") or None)
            self._roster_audit(actor, "Request leave", r["name"], f"{x['start']} to {x['end']}: {x['status']}")
            return {"request": x}
        if sub == "swap":
            x = rs.request_swap(actor, rid, body.get("member_a"), body.get("date_a"), body.get("member_b"),
                                body.get("date_b") or None, body.get("kind") or "swap", body.get("note", ""),
                                body.get("approver") or None)
            self._roster_audit(actor, "Request " + x["kind"], r["name"], f"{x['date_a']}: {x['status'].replace('_', ' ')}")
            return {"request": x}
        if sub == "consent":
            accept = bool(body.get("accept"))
            x = rs.consent(actor, rid, body.get("id"), accept, body.get("note", ""))
            self._roster_audit(actor, ("Accept " if accept else "Decline ") + x["kind"], r["name"], f"{x['date_a']}: {x['status']}")
            return {"request": x}
        if sub == "decide":
            approve = bool(body.get("approve"))
            x = rs.decide(actor, rid, body.get("kind"), body.get("id"), approve, body.get("note", ""))
            self._roster_audit(actor, ("Approve " if approve else "Reject ") + str(body.get("kind")), r["name"], str(body.get("id")))
            return {"request": x}
        if sub == "cancel":
            x = rs.cancel(actor, rid, body.get("kind"), body.get("id"))
            self._roster_audit(actor, "Cancel " + str(body.get("kind")), r["name"], str(body.get("id")))
            return {"request": x}
        if sub == "override":
            rs.set_override(actor, rid, body.get("member_id"), body.get("date"), body.get("code") or None)
            self._roster_audit(actor, "Change shift", r["name"], f"{body.get('date')}: {body.get('code') or 'rotation'}")
            return {"ok": True}
        if sub == "overrides":
            n = rs.set_overrides(actor, rid, body.get("changes"))
            self._roster_audit(actor, "Change shifts", r["name"], f"{n} change(s)")
            return {"changed": n}
        if sub == "clone":
            new = rs.clone(actor, rid, body.get("name", ""), body.get("team", ""), bool(body.get("members")))
            self._roster_audit(actor, "Clone roster", new["name"], f"from {r['name']}")
            return {"roster": self._roster_summary(new)}
        if sub == "restore":
            rs.restore(actor, rid, int(body.get("index", -1)))
            self._roster_audit(actor, "Restore roster version", r["name"], str(body.get("index")))
            return {"roster": self._roster_summary(rs.get(rid))}
        raise ApiError(404, f"No API route for {method} {path}.")

    def _api_v4(self, method: str, path: str, query: dict, ident: dict, q):
        srv, store, cfg = self.server, self.server.store, self.server.config
        user = ident["username"]

        if path == "/api/me/password" and method == "POST":
            body = self._body()
            u = store.get_user(user)
            if u["auth"] == "entra" and not u["password"]:
                raise ApiError(400, "Your account signs in with Microsoft 365; its password is managed there.")
            store.set_password(user, str(body.get("new", "")), user, False, current=str(body.get("current", "")))
            srv.sessions.drop_user(user, except_sid=ident["session"]["id"])
            return {"changed": True}
        if path == "/api/hierarchy/shift/clone" and method == "POST":
            body = self._body()
            label = f"{body.get('date', '')} {body.get('shift', '')} -> {body.get('to_date', '')} {body.get('to_shift', '')}"
            return store.mutate_hier(lambda h, ctx: h_clone_shift(h, body, user), user, "Clone shift", label)

        # ---- terminals and remote desktop ------------------------------------------------------
        def asset_by_id(aid):
            if not aid:
                return None
            a = next((r for r in store.state["xlsx"]["tables"]["Asset_Master"]["rows"] if r[ID_COL] == aid
                      and store._row_visible("Asset_Master", r)), None)
            if a is None:
                raise ApiError(404, "That asset no longer exists.")
            return a

        def asset_target(a, host):
            host = str(host or "").strip() or (a and (a.get("IP_Address") or a.get("Hostname_FQDN") or a.get("Device"))) or ""
            label = (a.get("Device") if a else "") or host
            return host, label

        if path == "/api/terminal/live" and method == "GET":
            return {"sessions": srv.terminals.live(user), "available": TerminalManager.available()}
        if path == "/api/terminal/sessions" and method == "POST":
            body = self._body()
            a = asset_by_id(body.get("asset_id"))
            host, label = asset_target(a, body.get("host"))
            t = srv.terminals.create(ident, self._client_ip(), host, body.get("port") or 22, body.get("username"), label)
            return {"sid": t.sid, "token": t.token, "host": t.host, "port": t.port, "username": t.username,
                    "label": label}
        if path == "/api/rdp/file" and method == "GET":
            a = asset_by_id(q("asset_id"))
            host, label = asset_target(a, q("host"))
            if not host or not re.fullmatch(r"[A-Za-z0-9._:\-\[\]]+", host):
                raise ApiError(400, "This asset has no IP address or host name to connect to.")
            port = int(q("port", "3389") or 3389)
            uname = re.sub(r"[\r\n]", "", q("username"))[:100]
            t = TerminalSession(new_id("rdp_"), user, ident["role"], self._client_ip(), host, port, uname, label)
            store.log_terminal_event(t, "launched", "Remote Desktop file downloaded (session not recorded)",
                                     kind="rdp-file")
            return ("file", rdp_file(host, uname, port), "application/x-rdp",
                    f"{re.sub(r'[^A-Za-z0-9_.-]', '_', label)}.rdp")
        if path == "/api/rdp/guac" and method == "POST":
            if not (cfg.guac_url and cfg.guac_secret):
                raise ApiError(501, "In-browser RDP is not set up. Add the Guacamole URL and secret key in "
                                    "msp_settings.ini, or use the Remote Desktop file option.")
            body = self._body()
            a = asset_by_id(body.get("asset_id"))
            host, label = asset_target(a, body.get("host"))
            if not host or not srv.terminals._allowed_target(host):
                raise ApiError(403, f"'{host}' is not a configuration item in the CMDB.")
            port = int(body.get("port") or 3389)
            sid = new_id("rdp_")
            params = {"hostname": host, "port": str(port), "security": "any", "ignore-cert": "true"}
            uname, domain = clean_text(body.get("username", ""), 100).strip(), clean_text(body.get("domain", ""), 100).strip()
            if uname:
                params["username"] = uname
            if domain:
                params["domain"] = domain
            rec_name = ""
            if cfg.guac_recording_path:
                rec_name = f"{dt.datetime.now().strftime('%Y%m%d_%H%M%S')}_{re.sub(r'[^A-Za-z0-9_.-]', '_', user)}_" \
                           f"{re.sub(r'[^A-Za-z0-9_.-]', '_', host)}_{sid}"
                params.update({"recording-path": cfg.guac_recording_path, "recording-name": rec_name,
                               "create-recording-path": "true"})
            payload = {"username": user, "expires": int((time.time() + 300) * 1000),
                       "connections": {f"{label} ({host})": {"protocol": "rdp", "parameters": params}}}
            token = guac_encrypted_token(cfg.guac_secret, payload)
            t = TerminalSession(sid, user, ident["role"], self._client_ip(), host, port, uname, label)
            store.log_terminal_event(t, "launched", f"Guacamole recording: {rec_name}" if rec_name
                                     else "Guacamole session (recording not configured)", kind="rdp-guac")
            return {"url": cfg.guac_url.rstrip("/") + "/?data=" + urllib.parse.quote(token, safe="")}

        # ---- administration ------------------------------------------------------------------------
        if not path.startswith("/api/admin/"):
            return None
        if path == "/api/admin/users":
            if method == "GET":
                return {"users": store.users_list(), "roles": ROLES, "role_help": ROLE_HELP}
            if method == "POST":
                u, temp = store.create_user(self._body(), user)
                return {"user": u, "temp_password": temp}
        m = re.fullmatch(r"/api/admin/users/([A-Za-z0-9._@\-]+)(/reset)?", path)
        if m:
            target = m.group(1)
            if m.group(2) and method == "POST":
                temp = temp_password()
                store.set_password(target, temp, user, must_change=True)
                srv.sessions.drop_user(target)
                return {"temp_password": temp}
            if method == "PUT":
                res = store.update_user(target, self._body(), user)
                if res.get("disabled"):
                    srv.sessions.drop_user(target)
                return {"user": res}
            if method == "DELETE":
                res = store.delete_user(target, user)
                srv.sessions.drop_user(target)
                if srv.org is not None and any(msp_org.norm(target) in t["members"] or msp_org.norm(target) in
                                               [msp_org.norm(x) for x in t.get("leads", [])] for t in srv.org.load()[0]["teams"].values()):
                    srv.org.mutate({"user": user, "role": "Admin"}, "Remove deleted user from teams", target,
                                   lambda st: msp_org.forget_user(st, target))
                return res
        if path == "/api/admin/sessions" and method == "GET":
            return {"web": srv.sessions.list(), "terminals": srv.terminals.live()}
        m = re.fullmatch(r"/api/admin/sessions/([a-f0-9]{20})", path)
        if m and method == "DELETE":
            if hmac.compare_digest(m.group(1), session_handle(ident["session"]["id"])):
                raise ApiError(400, "Use Sign out to end your own session.")
            if not srv.sessions.drop_handle(m.group(1)):
                raise ApiError(404, "That session has already ended.")
            return {"ended": True}
        m = re.fullmatch(r"/api/admin/terminals/([A-Za-z0-9_\-]+)/kill", path)
        if m and method == "POST":
            if not srv.terminals.kill(m.group(1)):
                raise ApiError(404, "That terminal session is no longer running.")
            return {"killed": m.group(1)}
        if path == "/api/admin/recordings" and method == "GET":
            return {"sessions": store.terminal_sessions(int(q("limit", "1000") or 1000))}
        m = re.fullmatch(r"/api/admin/recordings/([A-Za-z0-9_\-]+)(/download)?", path)
        if m and method == "GET":
            data, rec = store.read_recording(m.group(1))
            if m.group(2):
                return ("file", data, "text/plain; charset=utf-8", os.path.basename(rec["file"]))
            text = data.decode("utf-8", "replace")
            ok = (hashlib.sha256(data).hexdigest() == rec.get("sha256")) if rec.get("sha256") else None
            return {"session": rec, "content": text[-2_000_000:], "integrity_ok": ok}
        if path == "/api/admin/known-hosts":
            if method == "GET":
                return {"hosts": [{"host": k, **v} for k, v in sorted(store.state["users"]["known_hosts"].items())]}
            if method == "DELETE":
                return store.untrust_host(str(self._body().get("host", "")), user)
        if path == "/api/admin/settings":
            if method == "GET":
                return {"settings": store.settings()}
            if method == "PUT":
                new = self._body().get("settings") or {}
                trial = normalize_settings({**store.settings(), **(new if isinstance(new, dict) else {})})
                if msp_ai and trial["ai_mode"] == "model":
                    try:
                        msp_ai.check_local_endpoint(trial["ai_endpoint"])
                    except msp_ai.AssistantError as exc:
                        raise ApiError(400, exc.message)
                return {"settings": store.update_settings(new, user)}
        if path == "/api/admin/backup" and method == "POST":
            return {"backup": store.backup_now(user)}
        if path == "/api/admin/health" and method == "GET":
            with store.lock:
                backups = list(reversed(store.state["meta"]["backups"][-50:]))
            return {"status": store.status(), "backups": backups, "checks": {
                "https": isinstance(self.connection, ssl.SSLSocket) or self._is_https(),
                "ssh_terminals": TerminalManager.available(),
                "guacamole": bool(cfg.guac_url and cfg.guac_secret),
                "guacamole_recording": bool(cfg.guac_recording_path),
                "microsoft_signin": srv.entra is not None, "proxy_identity": cfg.trust_proxy_user,
                "web_sessions": len(srv.sessions.list()), "live_terminals": len(srv.terminals.live()),
                "python": sys.version.split()[0], "storage": store.storage.describe(),
                "roster": srv.rosters is not None, "teams": srv.org is not None,
                "assistant": srv.assistant is not None, "assistant_mode": store.settings().get("ai_mode", "search"),
                "local_assets": os.path.isfile(os.path.join(STATIC_DIR, "vendor", "xterm.js"))
                and os.path.isfile(os.path.join(STATIC_DIR, "fonts", "ibm-plex-sans-latin-400-normal.woff2")),
                "roster_holidays": bool(msp_roster and msp_roster.holidays_available())}}
        raise ApiError(404, f"No API route for {method} {path}.")
class AppServer(ThreadingHTTPServer):
    daemon_threads = True
    # On Windows, SO_REUSEADDR lets a second program share a port that is already in use, so the browser may
    # reach the wrong program. There the port is taken exclusively instead (see server_bind).
    allow_reuse_address = os.name != "nt"

    def server_bind(self):
        if os.name == "nt" and hasattr(socket, "SO_EXCLUSIVEADDRUSE"):
            self.socket.setsockopt(socket.SOL_SOCKET, socket.SO_EXCLUSIVEADDRUSE, 1)
        super().server_bind()

    def __init__(self, config: Config, store: DataStore):
        self.config = config
        self.store = store
        self.sessions = SessionManager()
        self.throttle = LoginThrottle()
        self.terminals = TerminalManager(store)
        self.rosters = msp_roster.RosterStore(store.storage) if msp_roster else None
        store.roster_store = self.rosters
        self.assistant = msp_ai.Assistant() if msp_ai else None
        self.org = msp_org.OrgStore(store.storage) if msp_org else None
        store.org_store = self.org
        self.ai_hits = {}
        self.ai_lock = threading.Lock()
        self.entra = None
        if config.entra_enabled:
            tenant = config.entra_tenant_id or config.sp_tenant_id
            client = config.entra_client_id or config.sp_client_id
            secret = config.entra_client_secret or config.sp_client_secret
            if not (tenant and client and secret and config.public_url):
                raise ValueError("Microsoft 365 sign-in needs tenant id, client id, client secret and public_url "
                                 "(the address people use to open the portal).")
            self.entra = EntraAuth(tenant, client, secret, config.public_url.rstrip("/") + "/auth/entra/callback",
                                   config.auth_base)
        self.setup_code = ""
        if not store.state["users"]["users"]:
            self.setup_code = "-".join(secrets.token_hex(2).upper() for _ in range(3))
            path = self.setup_code_path()
            try:
                with open(path, "w", encoding="utf-8") as fh:
                    fh.write(f"{APP_NAME} first-run setup code: {self.setup_code}\n"
                             "Open the portal in your browser and enter this code to create the first administrator.\n"
                             "This file is deleted automatically once setup is complete.\n")
            except OSError:
                pass
        super().__init__((config.host, config.port), Handler)
        if config.tls_cert:
            ctx = ssl.SSLContext(ssl.PROTOCOL_TLS_SERVER)
            ctx.minimum_version = ssl.TLSVersion.TLSv1_2
            ctx.load_cert_chain(config.tls_cert, config.tls_key or None)
            self.socket = ctx.wrap_socket(self.socket, server_side=True)
        self._stop = threading.Event()
        if config.backup_thread:
            threading.Thread(target=self._housekeeping, daemon=True).start()

    def setup_code_path(self) -> str:
        return self.config.setup_code_file or os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                                           "FIRST_RUN_SETUP_CODE.txt")

    def clear_setup_code(self):
        self.setup_code = ""
        try:
            os.remove(self.setup_code_path())
        except OSError:
            pass

    def _housekeeping(self):
        while not self._stop.wait(60):
            try:
                with self.store.lock:
                    self.store.refresh()
                self.store.auto_backup_if_due()
            except Exception as exc:  # keep the thread alive
                self.store.last_error = f"Automatic backup failed: {exc}"
            self._stop.wait(3540)

    def server_close(self):
        self._stop.set()
        super().server_close()


PORT_TRIES = 20
PORT_BUSY = {errno.EADDRINUSE, errno.EACCES, 10048, 10013}  # in use; reserved or blocked (Windows 10048 / 10013)


def port_free(host: str, port: int) -> bool:
    """Whether this server could listen on host:port, tested with the same socket options the server uses."""
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as probe:
        if os.name == "nt" and hasattr(socket, "SO_EXCLUSIVEADDRUSE"):
            probe.setsockopt(socket.SOL_SOCKET, socket.SO_EXCLUSIVEADDRUSE, 1)
        else:
            probe.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
        try:
            probe.bind((host or "0.0.0.0", port))  # nosec B104 - a short test bind on the configured address
            return True
        except OSError as exc:
            if exc.errno in PORT_BUSY or getattr(exc, "winerror", None) in PORT_BUSY:
                return False
            raise


def create_server(config: Config) -> AppServer:
    """Build the server. If the configured port is taken, use the next free one and say so (server.port_note)."""
    note = ""
    if config.port:
        first = config.port
        free = next((p for p in range(first, min(first + PORT_TRIES, 65536)) if port_free(config.host, p)), None)
        if free is None:
            raise ValueError(f"Ports {first} to {first + PORT_TRIES - 1} are all in use by other programs. "
                             "Choose a free port in msp_settings.ini ([server] port = ...) and start again.")
        if free != first:
            note = (f"Port {first} is already used by another program, so MSP Ops Center is using port {free}. "
                    f"To keep a fixed port, set [server] port in msp_settings.ini.")
            config = dataclasses.replace(config, port=free)
    storage = build_storage(config)
    store = DataStore(storage, refresh_interval=config.refresh_interval, seed=config.seed)
    server = AppServer(config, store)
    server.port_note = note
    return server


# ---- settings file (msp_settings.ini) ----------------------------------------------------------
SETTINGS_TEMPLATE = """; MSP Ops Center settings. Edit in Notepad, save, then restart Start_MSP_Ops_Center.bat.
; Lines starting with ; are comments. Leave a value empty to use the default.
; Keep this file private: it can contain secrets.

[server]
; 127.0.0.1 = only this computer.  0.0.0.0 = the whole team can connect to this computer.
host = 127.0.0.1
port = 8080
open_browser = yes
; Address people type to reach the portal (needed for Microsoft 365 sign-in), e.g. https://msp-ops.contoso.com
public_url =
; HTTPS certificate and key (PEM). Strongly recommended when host = 0.0.0.0
tls_cert =
tls_key =
; yes ONLY behind Azure App Service sign-in or a reverse proxy that sets (and strips) identity headers.
; Never enable it when people can reach this server directly: anyone could then claim any identity.
trust_proxy_headers = no
; Sites allowed to embed the portal in a frame, e.g. 'self' https://contoso.sharepoint.com
frame_ancestors = 'self'

[storage]
; local = files in data_dir (can be a OneDrive-synced SharePoint folder)
; sharepoint = read/write directly in a SharePoint library through Microsoft Graph
mode = local
data_dir = ./msp_data
site_host =
site_path =
library =
folder = MSP-Runbook
tenant_id =
client_id =
client_secret =

[signin]
; Microsoft 365 sign-in in addition to user name + password. Blank ids reuse the [storage] app registration.
microsoft = no
tenant_id =
client_id =
client_secret =

[guacamole]
; In-browser RDP through Apache Guacamole with the guacamole-auth-json extension.
url =
; 32 hexadecimal characters, the same value as json-secret-key in guacamole.properties
secret_key =
; Folder on the Guacamole server for session recordings, e.g. ${HISTORY_PATH}/${HISTORY_UUID}
recording_path =
"""

INI_MAP = {
    ("server", "host"): "MSP_HOST", ("server", "port"): "MSP_PORT", ("server", "open_browser"): "MSP_OPEN_BROWSER",
    ("server", "public_url"): "MSP_PUBLIC_URL", ("server", "tls_cert"): "MSP_TLS_CERT",
    ("server", "tls_key"): "MSP_TLS_KEY", ("server", "trust_proxy_headers"): "MSP_TRUST_PROXY_USER",
    ("server", "frame_ancestors"): "MSP_FRAME_ANCESTORS",
    ("storage", "mode"): "MSP_STORAGE", ("storage", "data_dir"): "MSP_DATA_DIR",
    ("storage", "site_host"): "MSP_SP_SITE_HOST", ("storage", "site_path"): "MSP_SP_SITE_PATH",
    ("storage", "library"): "MSP_SP_LIBRARY", ("storage", "folder"): "MSP_SP_FOLDER",
    ("storage", "tenant_id"): "MSP_SP_TENANT_ID", ("storage", "client_id"): "MSP_SP_CLIENT_ID",
    ("storage", "client_secret"): "MSP_SP_CLIENT_SECRET", ("storage", "drive_id"): "MSP_SP_DRIVE_ID",
    ("storage", "onedrive_user"): "MSP_SP_ONEDRIVE_USER",
    ("signin", "microsoft"): "MSP_ENTRA_ENABLED", ("signin", "tenant_id"): "MSP_ENTRA_TENANT_ID",
    ("signin", "client_id"): "MSP_ENTRA_CLIENT_ID", ("signin", "client_secret"): "MSP_ENTRA_CLIENT_SECRET",
    ("guacamole", "url"): "MSP_GUAC_URL", ("guacamole", "secret_key"): "MSP_GUAC_SECRET",
    ("guacamole", "recording_path"): "MSP_GUAC_RECORDING_PATH",
}


def load_settings_file(path: str) -> dict:
    """Read msp_settings.ini into {ENV_NAME: value}; empty values are ignored."""
    out = {}
    if not path or not os.path.exists(path):
        return out
    cp = configparser.ConfigParser(interpolation=None, inline_comment_prefixes=(";",))
    cp.read(path, encoding="utf-8")
    for (sec, key), env in INI_MAP.items():
        if cp.has_option(sec, key):
            v = cp.get(sec, key).strip()
            if v:
                out[env] = v
    return out


def config_from_args(argv=None) -> Config:
    pre = argparse.ArgumentParser(add_help=False)
    pre.add_argument("--config", default=os.environ.get("MSP_CONFIG", ""))
    known, _ = pre.parse_known_args(argv)
    cfg_path = known.config
    if not cfg_path:
        default = os.path.join(os.path.dirname(os.path.abspath(__file__)), "msp_settings.ini")
        cfg_path = default if os.path.exists(default) else ""
    file_vals = load_settings_file(cfg_path)
    env = lambda k, d="": os.environ.get(k, file_vals.get(k, d))  # noqa: E731 - env beats file
    yes = lambda k: truthy(env(k, ""))  # noqa: E731

    p = argparse.ArgumentParser(
        description=f"{APP_NAME} - ITIL v4 MSP operations portal",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="Settings come from (highest first): command line, environment variables, msp_settings.ini.")
    p.add_argument("--config", default=cfg_path, help="settings file (default: msp_settings.ini next to the program)")
    p.add_argument("--write-settings-template", metavar="PATH", help="write a commented settings file and exit")
    p.add_argument("--host", default=env("MSP_HOST", "127.0.0.1"), help="bind address; 0.0.0.0 lets the team connect")
    p.add_argument("--port", type=int, default=int(env("MSP_PORT", env("PORT", "8080"))))
    p.add_argument("--storage", choices=["local", "sharepoint"], default=env("MSP_STORAGE", "local"))
    p.add_argument("--data-dir", default=env("MSP_DATA_DIR", "./msp_data"))
    p.add_argument("--refresh-interval", type=float, default=float(env("MSP_REFRESH_SECONDS", "5")))
    p.add_argument("--no-seed", action="store_true", default=yes("MSP_NO_SEED"))
    p.add_argument("--open-browser", action="store_true", default=yes("MSP_OPEN_BROWSER"))
    p.add_argument("--public-url", default=env("MSP_PUBLIC_URL"))
    p.add_argument("--tls-cert", default=env("MSP_TLS_CERT"))
    p.add_argument("--tls-key", default=env("MSP_TLS_KEY"))
    p.add_argument("--trust-proxy-user", action="store_true", default=yes("MSP_TRUST_PROXY_USER"))
    p.add_argument("--frame-ancestors", default=env("MSP_FRAME_ANCESTORS", "'self'"),
                   help="who may embed the portal, e.g. \"'self' https://contoso.sharepoint.com\"")
    g = p.add_argument_group("SharePoint / OneDrive storage")
    for opt, envname in (("--sp-tenant-id", "MSP_SP_TENANT_ID"), ("--sp-client-id", "MSP_SP_CLIENT_ID"),
                         ("--sp-client-secret", "MSP_SP_CLIENT_SECRET"), ("--sp-site-host", "MSP_SP_SITE_HOST"),
                         ("--sp-site-path", "MSP_SP_SITE_PATH"), ("--sp-library", "MSP_SP_LIBRARY"),
                         ("--sp-drive-id", "MSP_SP_DRIVE_ID"), ("--sp-onedrive-user", "MSP_SP_ONEDRIVE_USER")):
        g.add_argument(opt, default=env(envname))
    g.add_argument("--sp-folder", default=env("MSP_SP_FOLDER", "MSP-Runbook"))
    e = p.add_argument_group("Microsoft 365 sign-in")
    e.add_argument("--entra", action="store_true", default=yes("MSP_ENTRA_ENABLED"))
    e.add_argument("--entra-tenant-id", default=env("MSP_ENTRA_TENANT_ID"))
    e.add_argument("--entra-client-id", default=env("MSP_ENTRA_CLIENT_ID"))
    e.add_argument("--entra-client-secret", default=env("MSP_ENTRA_CLIENT_SECRET"))
    gu = p.add_argument_group("Apache Guacamole (in-browser RDP)")
    gu.add_argument("--guac-url", default=env("MSP_GUAC_URL"))
    gu.add_argument("--guac-secret", default=env("MSP_GUAC_SECRET"))
    gu.add_argument("--guac-recording-path", default=env("MSP_GUAC_RECORDING_PATH"))
    a = p.parse_args(argv)
    if a.write_settings_template:
        with open(a.write_settings_template, "w", encoding="utf-8") as fh:
            fh.write(SETTINGS_TEMPLATE)
        print(f"Wrote {a.write_settings_template}")
        sys.exit(0)
    data_dir = a.data_dir
    if cfg_path and not os.path.isabs(data_dir):
        data_dir = os.path.join(os.path.dirname(os.path.abspath(cfg_path)), data_dir)
    return Config(host=a.host, port=a.port, storage=a.storage, data_dir=data_dir, refresh_interval=a.refresh_interval,
                  seed=not a.no_seed, trust_proxy_user=a.trust_proxy_user, frame_ancestors=a.frame_ancestors,
                  public_url=a.public_url, tls_cert=a.tls_cert, tls_key=a.tls_key, open_browser=a.open_browser,
                  entra_enabled=a.entra, entra_tenant_id=a.entra_tenant_id, entra_client_id=a.entra_client_id,
                  entra_client_secret=a.entra_client_secret, guac_url=a.guac_url, guac_secret=a.guac_secret,
                  guac_recording_path=a.guac_recording_path,
                  sp_tenant_id=a.sp_tenant_id, sp_client_id=a.sp_client_id, sp_client_secret=a.sp_client_secret,
                  sp_site_host=a.sp_site_host, sp_site_path=a.sp_site_path, sp_library=a.sp_library,
                  sp_drive_id=a.sp_drive_id, sp_onedrive_user=a.sp_onedrive_user, sp_folder=a.sp_folder,
                  graph_base=env("MSP_GRAPH_BASE", "https://graph.microsoft.com/v1.0"),
                  auth_base=env("MSP_AUTH_BASE", "https://login.microsoftonline.com"))


def main(argv=None):
    cfg = config_from_args(argv)
    try:
        server = create_server(cfg)
    except (ValueError, ApiError, OSError, ssl.SSLError) as exc:
        msg = exc.message if isinstance(exc, ApiError) else str(exc)
        sys.stderr.write(f"Startup failed: {msg}\n")
        return 2
    info = server.store.storage.describe()
    scheme = "https" if cfg.tls_cert else "http"
    shown = "127.0.0.1" if cfg.host in ("127.0.0.1", "0.0.0.0", "", "localhost") else cfg.host  # nosec B104 - display only
    url = f"{scheme}://{shown}:{server.server_address[1]}/"
    print(f"{APP_NAME} {APP_VERSION}")
    print(f"  Storage   : {info['label']} -> {info['location']}")
    print(f"  Open      : {url}")
    if getattr(server, "port_note", ""):
        print(f"  Note      : {server.port_note}")
    if cfg.host == "0.0.0.0":  # nosec B104 - informational message only
        print(f"  Team      : {scheme}://{socket.gethostname()}:{server.server_address[1]}/")
    print(f"  Sign-in   : user name + password{' and Microsoft 365' if server.entra else ''}")
    print(f"  Terminals : {'SSH ready' if TerminalManager.available() else 'install paramiko to enable SSH'}"
          f"{', Guacamole RDP ready' if cfg.guac_url and cfg.guac_secret else ''}")
    if server.setup_code:
        print("\n  FIRST RUN: open the portal and enter this setup code to create the administrator:")
        print(f"      {server.setup_code}")
        print(f"  (also saved in {server.setup_code_path()})\n")
    if cfg.trust_proxy_user:
        print("  Security  : trust_proxy_headers is ON. Only safe behind App Service sign-in / a reverse proxy.")
    if cfg.host not in ("127.0.0.1", "localhost") and not cfg.tls_cert and not cfg.trust_proxy_user:
        print("  Warning   : the team can connect but traffic is not encrypted. Set tls_cert/tls_key, or put the")
        print("              portal behind an HTTPS reverse proxy, before using passwords and terminals.")
    if cfg.open_browser:
        threading.Timer(1.0, lambda: webbrowser.open(url)).start()
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nStopping.")
    finally:
        server.server_close()
    return 0


# --------------------------------------------------------------------------------------
# Web UI (single page, served from memory)
# --------------------------------------------------------------------------------------
HTML_PAGE = r"""<!DOCTYPE html>
<html lang="en" data-theme="dark">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>MSP Ops Center</title>
<style>
:root{
  --bg:#0A1120;--bg2:#0D1628;--side:#0B1324;--panel:#111C31;--panel2:#16233D;--hover:#1B2A47;--line:#1F2E4B;--line2:#2B3D60;
  --text:#E6EDF7;--muted:#97A6C0;--dim:#6C7C99;--accent:#2DD4BF;--accent-soft:rgba(45,212,191,.13);--accent2:#7AA7FF;--ink:#04201C;
  --care:#F2B44B;--p1:#F0605A;--p2:#F59345;--p3:#E0C153;--p4:#6FA8F0;--ok:#43C08F;
  --morning:#F2C76E;--afternoon:#7AA7FF;--night:#A391F5;--shadow:0 18px 50px rgba(0,0,0,.45);
  --sans:"IBM Plex Sans","Segoe UI",system-ui,-apple-system,sans-serif;
  --mono:"IBM Plex Mono",Consolas,"Cascadia Mono",Menlo,monospace;
  color-scheme:dark;
}
html[data-theme="light"]{
  --bg:#F3F5F9;--bg2:#FFFFFF;--side:#0F1A2E;--panel:#FFFFFF;--panel2:#F4F6FA;--hover:#EEF2F8;--line:#E2E7F0;--line2:#CDD5E3;
  --text:#13203A;--muted:#55647F;--dim:#8592AA;--accent:#0E9F8E;--accent-soft:rgba(14,159,142,.12);--accent2:#2F6FDB;--ink:#FFFFFF;
  --care:#B7791F;--ok:#1E8E62;--shadow:0 18px 50px rgba(20,32,58,.18);color-scheme:light;
}
*{box-sizing:border-box}
html,body{margin:0;background:var(--bg);color:var(--text);font-family:var(--sans);font-size:14px;line-height:1.45}
a{color:var(--accent2)}
:focus-visible{outline:2px solid var(--accent);outline-offset:2px}
.hidden{display:none!important}
/* ---------- app frame with sidebar blades ---------- */
.app{display:grid;grid-template-columns:248px 1fr;min-height:100vh;transition:grid-template-columns .18s}
.app.collapsed{grid-template-columns:68px 1fr}
.side{position:sticky;top:0;height:100vh;background:var(--side);color:#DCE6F5;display:flex;flex-direction:column;border-right:1px solid rgba(255,255,255,.06);overflow:hidden}
.brand{display:flex;gap:11px;align-items:center;padding:16px 16px 14px}
.brand .logo{width:34px;height:34px;flex:none;border-radius:9px;display:grid;place-items:center;background:linear-gradient(135deg,#2DD4BF,#3B82F6);color:#04121A;font-weight:700;font-size:15px}
.brand b{display:block;font-size:15px;font-weight:600;color:#F2F6FC;white-space:nowrap}
.brand span{font-size:11.5px;color:#8FA0BD;white-space:nowrap}
.nav{flex:1;overflow-y:auto;padding:4px 10px 10px}
.nav .grp{font-size:11.5px;color:#7486A6;margin:14px 10px 6px;white-space:nowrap}
.blade{display:flex;align-items:center;gap:11px;width:100%;padding:7px 9px;margin:2px 0;border:0;border-radius:10px;background:none;color:#C5D1E5;font:inherit;font-size:13.5px;cursor:pointer;text-align:left;position:relative;white-space:nowrap}
.blade .ic{width:30px;height:30px;flex:none;border-radius:8px;display:grid;place-items:center;background:rgba(255,255,255,.05);color:#A9B8D2;transition:background .15s,color .15s}
.blade .ic svg{width:17px;height:17px}
.blade:hover{background:rgba(255,255,255,.05);color:#fff}
.blade:hover .ic{color:#fff}
.blade.on{background:rgba(45,212,191,.12);color:#fff}
.blade.on .ic{background:#2DD4BF;color:#04201C}
.blade.on::before{content:"";position:absolute;left:-10px;top:8px;bottom:8px;width:3px;border-radius:0 3px 3px 0;background:#2DD4BF}
.blade .lbl{flex:1;overflow:hidden;text-overflow:ellipsis}
.blade .n{font-size:11px;min-width:20px;text-align:center;padding:1px 6px;border-radius:9px;background:rgba(255,255,255,.08);color:#C5D1E5}
.blade .n.hot{background:rgba(240,96,90,.25);color:#FFB3AE}
.collapsed .brand div,.collapsed .blade .lbl,.collapsed .blade .n,.collapsed .nav .grp,.collapsed .me .who,.collapsed .me .acts{display:none}
.side-foot{border-top:1px solid rgba(255,255,255,.07);padding:10px}
.me{display:flex;align-items:center;gap:10px;padding:6px}
.avatar{width:34px;height:34px;border-radius:50%;flex:none;display:grid;place-items:center;font-weight:600;font-size:13px;background:#24365A;color:#E7EEF9;border:0;cursor:pointer;font-family:inherit}
.me .who{flex:1;min-width:0}
.me .who b{display:block;font-size:13px;font-weight:600;color:#EEF3FB;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.rolepill{display:inline-block;font-size:11px;padding:0 7px;border-radius:9px;background:rgba(45,212,191,.16);color:#7FE9DA}
.me .acts{display:flex;gap:4px}
.iconbtn{width:30px;height:30px;border-radius:8px;border:0;background:rgba(255,255,255,.05);color:#B8C6DD;cursor:pointer;display:grid;place-items:center}
.iconbtn:hover{background:rgba(255,255,255,.1);color:#fff}
.iconbtn svg{width:16px;height:16px}
.main{min-width:0;display:flex;flex-direction:column}
.topbar{position:sticky;top:0;z-index:40;display:flex;align-items:center;gap:12px;padding:10px 22px;background:color-mix(in srgb,var(--bg) 88%,transparent);backdrop-filter:blur(8px);border-bottom:1px solid var(--line)}
.tb-toggle{background:none;border:1px solid var(--line2);color:var(--muted);border-radius:8px;width:34px;height:34px;cursor:pointer;display:grid;place-items:center;flex:none}
.tb-toggle svg{width:17px;height:17px}
.crumbs{font-size:13px;color:var(--muted);white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.crumbs b{color:var(--text);font-weight:600}
.searchbtn{margin-left:auto;display:flex;align-items:center;gap:10px;min-width:250px;padding:7px 10px;border:1px solid var(--line2);border-radius:9px;background:var(--panel);color:var(--dim);font:inherit;font-size:13px;cursor:pointer}
.searchbtn kbd{margin-left:auto;font-family:var(--mono);font-size:11px;padding:1px 6px;border:1px solid var(--line2);border-radius:5px;color:var(--muted)}
.store{display:flex;align-items:center;gap:7px;font-size:12px;color:var(--muted);max-width:260px}
.store .loc{overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.dot{width:8px;height:8px;border-radius:50%;background:var(--ok);flex:none;box-shadow:0 0 0 3px color-mix(in srgb,var(--ok) 25%,transparent)}
.dot.err{background:var(--p1)}
main#view{padding:22px 26px 40px;max-width:1600px;width:100%}
/* ---------- bulletin ticker ---------- */
.ticker{display:flex;align-items:stretch;border-bottom:1px solid var(--line);background:var(--bg2);overflow:hidden}
.ticker .lbl{flex:none;padding:7px 14px;font-size:12.5px;font-weight:600;color:var(--muted);border-right:1px solid var(--line);display:flex;align-items:center;gap:6px}
.ticker .lbl button{background:none;border:0;color:var(--accent2);font:inherit;font-weight:500;cursor:pointer;padding:0}
.ticker .rail{flex:1;overflow:hidden;mask-image:linear-gradient(90deg,transparent,#000 30px,#000 calc(100% - 30px),transparent)}
.ticker .track{display:inline-flex;gap:34px;white-space:nowrap;padding:7px 0;animation:tick var(--dur,40s) linear infinite}
.ticker:hover .track,.ticker:focus-within .track{animation-play-state:paused}
@keyframes tick{from{transform:translateX(0)}to{transform:translateX(-50%)}}
.tk{background:none;border:0;color:var(--text);font:inherit;font-size:13px;cursor:pointer;display:inline-flex;gap:8px;align-items:center;padding:0}
.tk:hover .tt{text-decoration:underline}
.tk .unread{width:7px;height:7px;border-radius:50%;background:var(--accent)}
.sev{display:inline-block;padding:0 6px;border-radius:4px;font-size:11.5px;font-weight:600}
.sev-Critical{background:rgba(240,96,90,.18);color:var(--p1)}
.sev-Advisory{background:rgba(245,147,69,.17);color:var(--p2)}
.sev-Info{background:rgba(111,168,240,.17);color:var(--p4)}
/* ---------- shared components ---------- */
h2{font-size:21px;font-weight:600;margin:0 0 4px;letter-spacing:-.01em}
h3{font-size:15px;font-weight:600;margin:0}
.sub{color:var(--muted);margin:0 0 16px;max-width:860px}
.btn{background:var(--panel);border:1px solid var(--line2);color:var(--text);border-radius:8px;padding:7px 12px;font:inherit;font-size:13px;cursor:pointer;white-space:nowrap;transition:border-color .12s,background .12s}
.btn:hover{border-color:var(--accent);background:var(--hover)}
.btn.primary{background:var(--accent);border-color:var(--accent);color:var(--ink);font-weight:600}
.btn.primary:hover{filter:brightness(1.07);background:var(--accent)}
.btn.danger{background:transparent;border-color:color-mix(in srgb,var(--p1) 55%,transparent);color:var(--p1)}
.btn.sm{padding:4px 9px;font-size:12px}
.btn.xs{padding:2px 7px;font-size:11.5px;border-radius:6px}
.btn.ghost{background:transparent;border-color:transparent;color:var(--muted)}
.btn.ghost:hover{color:var(--text);border-color:var(--line2)}
.btn[disabled]{opacity:.45;cursor:not-allowed}
input,select,textarea{background:var(--bg);border:1px solid var(--line2);color:var(--text);border-radius:8px;padding:7px 9px;font:inherit;font-size:13px;max-width:100%}
input:focus,select:focus,textarea:focus{border-color:var(--accent);outline:none;box-shadow:0 0 0 3px var(--accent-soft)}
input[type=checkbox],input[type=radio]{accent-color:var(--accent)}
input[disabled],select[disabled],textarea[disabled]{opacity:.75}
textarea{width:100%;min-height:70px;resize:vertical}
textarea.mono,.mono,code{font-family:var(--mono);font-size:12.5px}
.toolbar{display:flex;flex-wrap:wrap;gap:8px;align-items:center;margin:0 0 12px}
.toolbar .grow{flex:1 1 200px}
.toolbar input[type=search]{width:100%}
.tablewrap{overflow:auto;border:1px solid var(--line);border-radius:10px;background:var(--panel);max-height:72vh}
table.grid{width:100%;border-collapse:collapse;font-size:13px}
.grid th{position:sticky;top:0;background:var(--panel2);text-align:left;font-weight:600;color:var(--muted);padding:9px 10px;border-bottom:1px solid var(--line2);white-space:nowrap;z-index:1;font-size:12.5px}
.grid td{padding:8px 10px;border-bottom:1px solid var(--line);vertical-align:top;max-width:320px}
.grid tbody tr:hover td{background:var(--hover)}
.grid td.actions{white-space:nowrap;text-align:right;width:1%}
.r2 .grid td[data-ie],.r3 .grid td[data-ie],.r4 .grid td[data-ie]{cursor:text}
.r2 .grid td[data-ie]:hover,.r3 .grid td[data-ie]:hover,.r4 .grid td[data-ie]:hover{box-shadow:inset 0 0 0 1px var(--line2)}
.ie{width:100%;min-width:120px}
.clip{display:-webkit-box;-webkit-line-clamp:3;-webkit-box-orient:vertical;overflow:hidden;white-space:pre-wrap}
tr.care td:first-child{box-shadow:inset 3px 0 0 var(--care)}
.empty{padding:30px;text-align:center;color:var(--muted)}
.badge{display:inline-block;padding:1px 7px;border-radius:5px;font-size:12px;font-weight:600;border:1px solid transparent;white-space:nowrap}
.pri-P1{background:color-mix(in srgb,var(--p1) 15%,transparent);color:var(--p1);border-color:color-mix(in srgb,var(--p1) 45%,transparent)}
.pri-P2{background:color-mix(in srgb,var(--p2) 15%,transparent);color:var(--p2);border-color:color-mix(in srgb,var(--p2) 45%,transparent)}
.pri-P3{background:color-mix(in srgb,var(--p3) 15%,transparent);color:var(--p3);border-color:color-mix(in srgb,var(--p3) 45%,transparent)}
.pri-P4{background:color-mix(in srgb,var(--p4) 15%,transparent);color:var(--p4);border-color:color-mix(in srgb,var(--p4) 45%,transparent)}
.stb{display:inline-block;padding:1px 8px;border-radius:10px;font-size:11.5px;font-weight:500;background:var(--panel2);color:var(--muted);white-space:nowrap;border:1px solid var(--line)}
.stb.Done,.stb.Completed,.stb.ended{background:color-mix(in srgb,var(--ok) 15%,transparent);color:var(--ok);border-color:transparent}
.stb.InProgress,.stb.connected,.stb.launched{background:color-mix(in srgb,var(--accent2) 16%,transparent);color:var(--accent2);border-color:transparent}
.stb.Failed,.stb.RolledBack,.stb.Cancelled,.stb.refused{background:color-mix(in srgb,var(--p1) 15%,transparent);color:var(--p1);border-color:transparent}
.stb.Scheduled{background:color-mix(in srgb,var(--night) 17%,transparent);color:var(--night);border-color:transparent}
.st{color:var(--muted);font-size:12.5px}
.st.done{color:var(--ok)}
.warn{color:var(--p2)}
.bad{color:var(--p1)}
.care-tag{color:var(--care);font-size:12.5px}
.care-tag::before{content:"\26A0  "}
.devlink{font-weight:600;text-decoration:none;color:var(--accent2);border-bottom:1px dotted var(--accent2)}
.devlink:hover{border-bottom-style:solid}
.cnt{font-size:11px;color:var(--dim);margin-left:4px}
.ki{color:var(--care);font-size:11px;margin-left:4px}
.art{display:inline-flex;gap:3px}
.art:not(.login .art) span{width:18px;height:18px;border-radius:4px;display:inline-grid;place-items:center;font-size:11px;font-weight:600;background:color-mix(in srgb,var(--ok) 16%,transparent);color:var(--ok)}
.art:not(.login .art) span.miss{background:color-mix(in srgb,var(--p1) 16%,transparent);color:var(--p1)}
.chip{font-size:12px;padding:1px 8px;border-radius:10px;background:var(--panel2);color:var(--muted);border:1px solid var(--line)}
.chip.new{color:var(--ok)}.chip.duplicate,.chip.duplicate_in_file{color:var(--p2)}.chip.invalid{color:var(--p1)}
/* ---------- home / overview ---------- */
.hello{display:flex;align-items:baseline;gap:14px;flex-wrap:wrap;margin:2px 0 18px}
.hello h2{font-size:27px;font-weight:500;letter-spacing:-.02em}
.sun{display:grid;grid-template-columns:repeat(3,1fr);border:1px solid var(--line);border-radius:12px;overflow:hidden;margin-bottom:20px}
.tile{position:relative;padding:15px 18px 17px;background:var(--panel);border:0;border-right:1px solid var(--line);text-align:left;color:var(--text);font:inherit;cursor:pointer}
.tile:last-child{border-right:0}
.tile::before{content:"";position:absolute;left:0;right:0;top:0;height:3px;background:var(--c)}
.tile:hover{background:var(--hover)}
.tile h3{display:flex;justify-content:space-between;align-items:baseline}
.tile h3 small{color:var(--muted);font-weight:400;font-size:12px}
.tile .nums{display:flex;gap:18px;margin-top:10px}
.tile .nums div{display:flex;flex-direction:column}
.tile .nums b{font-size:22px;font-weight:500}
.tile .nums span{color:var(--muted);font-size:12px}
.tile .start{margin-top:12px;color:var(--muted)}
.cols{display:grid;grid-template-columns:repeat(auto-fit,minmax(330px,1fr));gap:16px}
.panel{background:var(--panel);border:1px solid var(--line);border-radius:12px;padding:15px 17px}
.panel h3{margin-bottom:8px;display:flex;align-items:center;gap:8px}
.panel h3 .n{font-size:12px;font-weight:500;color:var(--muted)}
.list{list-style:none;margin:0;padding:0}
.list li{padding:8px 0;border-top:1px solid var(--line);display:flex;gap:10px;align-items:flex-start}
.list li:first-child{border-top:0}
.list .t{flex:1;min-width:0}
.list .t small{display:block;color:var(--muted)}
.list button.linkish{background:none;border:0;color:var(--text);font:inherit;padding:0;cursor:pointer;text-align:left}
.list button.linkish:hover{color:var(--accent2);text-decoration:underline}
.bar{height:6px;background:var(--panel2);border-radius:3px;overflow:hidden;margin-top:5px;min-width:60px;border:1px solid var(--line)}
.bar i{display:block;height:100%;background:var(--ok)}
/* ---------- trees ---------- */
details.node{border-left:1px solid var(--line);margin:3px 0 3px 7px;padding-left:12px}
details.node.lvl0{border-left:0;margin-left:0;padding-left:0}
details.node>summary{list-style:none;cursor:pointer;display:flex;align-items:center;gap:10px;flex-wrap:wrap;padding:6px 8px;border-radius:8px}
details.node>summary::-webkit-details-marker{display:none}
details.node>summary::before{content:"\25B8";color:var(--dim);transition:transform .15s;width:10px}
details.node[open]>summary::before{transform:rotate(90deg)}
details.node>summary:hover{background:var(--hover)}
details.node>summary .label{font-weight:600}
details.node>summary .meta{color:var(--muted);font-size:12.5px}
details.node>summary .sp{flex:1}
.shiftsum{border-left:3px solid var(--c);background:var(--panel)}
.sec{background:var(--panel);border:1px solid var(--line);border-radius:10px;margin:10px 0}
.sec-head{display:flex;align-items:center;gap:10px;flex-wrap:wrap;padding:9px 12px;border-bottom:1px solid var(--line)}
.sec-head h4{margin:0;font-size:13.5px;font-weight:600}
.sec-head .meta{color:var(--muted);font-size:12px}
.sec-head .sp{flex:1}
.carried{color:var(--dim);font-size:11.5px}
/* ---------- execution plans ---------- */
.task{background:var(--panel);border:1px solid var(--line);border-radius:12px;margin:8px 0 14px;overflow:hidden}
.task-info{display:flex;flex-wrap:wrap;gap:6px 22px;padding:12px 16px;border-bottom:1px solid var(--line);font-size:13px}
.task-info div{display:flex;flex-direction:column}
.task-info span{color:var(--muted);font-size:11.5px}
.task-tools{display:flex;flex-wrap:wrap;gap:6px;padding:10px 16px;border-bottom:1px solid var(--line);background:var(--panel2)}
.phase{border-bottom:1px solid var(--line)}
.phase:last-of-type{border-bottom:0}
.phase-head{display:flex;align-items:center;gap:12px;flex-wrap:wrap;padding:10px 16px;background:color-mix(in srgb,var(--accent) 6%,var(--panel))}
.phase-no{font-size:22px;font-weight:500;color:var(--accent);min-width:22px;font-variant-numeric:tabular-nums}
.phase-head h4{margin:0;font-size:14.5px;font-weight:600}
.phase-head .meta{color:var(--muted);font-size:12px}
.phase-head .sp{flex:1}
.gate{font-size:12px;padding:2px 8px;border-radius:5px;border:1px dashed var(--care);color:var(--care)}
.gate.ok{border-style:solid;border-color:color-mix(in srgb,var(--ok) 60%,transparent);color:var(--ok)}
.steps td.num{font-family:var(--mono);font-size:12px;color:var(--muted);white-space:nowrap}
.steps tr.l1 td.ttl{padding-left:28px}.steps tr.l2 td.ttl{padding-left:46px}.steps tr.l3 td.ttl{padding-left:64px}
.steps tr.parent td.ttl{font-weight:600}
.steps td .inh{color:var(--dim)}
.steps td .ovr{color:var(--text);font-weight:600}
.steps tr.cmdrow td{border-bottom:1px solid var(--line);padding-top:0}
.steps tr.stepr td{border-bottom:0}
.steps tr.stepr:has(+ tr:not(.cmdrow)) td,.steps tr.stepr:last-child td{border-bottom:1px solid var(--line)}
pre.cmd{position:relative;margin:4px 0 0;padding:7px 62px 7px 10px;background:var(--bg);border:1px solid var(--line);border-radius:7px;white-space:pre-wrap;word-break:break-word;font-family:var(--mono);font-size:12.5px}
pre.cmd .btn{position:absolute;top:4px;right:4px}
pre.cmd.rb{border-color:color-mix(in srgb,var(--p1) 40%,transparent)}
.cmdlbl{font-size:11.5px;color:var(--muted);margin-top:4px}
.overrun{color:var(--p1);font-weight:600}
/* ---------- runbook ---------- */
.rb-asset{background:var(--panel)}
.rb-info{font-size:12.5px;color:var(--muted);padding:4px 8px 6px 22px}
.entry{background:var(--panel);border:1px solid var(--line);border-radius:10px;padding:10px 14px;margin:8px 0}
.entry.ki-entry{border-left:3px solid var(--care)}
.entry header{display:flex;gap:10px;align-items:baseline;flex-wrap:wrap}
.entry header .t{font-weight:600}
.entry header .sp{flex:1}
.entry .f{margin-top:6px}
.entry .f b{display:block;font-size:11.5px;color:var(--muted);font-weight:500}
.entry .f div{white-space:pre-wrap}
.focus{box-shadow:0 0 0 2px var(--accent)}
/* ---------- terminals ---------- */
#termHost{padding:18px 22px 22px;display:flex;flex-direction:column;gap:10px;height:calc(100vh - 58px)}
.term-tabs{display:flex;gap:4px;flex-wrap:wrap;align-items:flex-end;border-bottom:1px solid var(--line2)}
.term-tab{display:flex;align-items:center;gap:8px;padding:6px 8px 6px 12px;border:1px solid var(--line2);border-bottom:0;border-radius:9px 9px 0 0;background:var(--panel2);color:var(--muted);font:inherit;font-size:12.5px;cursor:pointer}
.term-tab.on{background:#05080F;color:#E6EDF7;border-color:#05080F}
.term-tab .sdot{width:8px;height:8px;border-radius:50%;background:var(--dim)}
.term-tab .sdot.connected{background:var(--ok)}.term-tab .sdot.error{background:var(--p1)}.term-tab .sdot.connecting{background:var(--p3)}
.term-tab .x{border:0;background:none;color:inherit;cursor:pointer;font-size:15px;line-height:1;padding:0 3px}
.term-panes{flex:1;position:relative;min-height:300px}
.term-pane{position:absolute;inset:0;background:#05080F;border-radius:0 0 10px 10px;padding:8px;display:none}
.term-pane.on{display:block}
.term-pane .xt{position:absolute;inset:8px}
.term-pane .fallback{position:absolute;inset:8px;display:flex;flex-direction:column;gap:6px}
.term-pane .fallback pre{flex:1;margin:0;overflow:auto;color:#D7E2F2;font-family:var(--mono);font-size:13px;white-space:pre-wrap}
.term-pane .fallback .row{display:flex;gap:6px}
.term-pane .fallback input{flex:1;background:#0E1626;border-color:#22314F;color:#E6EDF7;font-family:var(--mono)}
.term-banner{position:absolute;right:14px;top:12px;z-index:5;font-size:11.5px;color:#9FB0CC;background:rgba(10,17,32,.85);border:1px solid #22314F;padding:3px 8px;border-radius:6px}
.term-empty{flex:1;display:grid;place-items:center;border:1px dashed var(--line2);border-radius:12px;color:var(--muted);text-align:center;padding:30px}
/* ---------- admin ---------- */
.subtabs{display:flex;gap:4px;border-bottom:1px solid var(--line);margin:4px 0 16px;flex-wrap:wrap}
.subtabs button{background:none;border:0;border-bottom:2px solid transparent;color:var(--muted);font:inherit;padding:8px 12px;cursor:pointer}
.subtabs button.on{color:var(--text);border-bottom-color:var(--accent)}
.checks{display:grid;grid-template-columns:repeat(auto-fit,minmax(240px,1fr));gap:10px;margin-bottom:16px}
.check{display:flex;gap:10px;align-items:center;padding:10px 12px;border:1px solid var(--line);border-radius:10px;background:var(--panel)}
.check i{width:10px;height:10px;border-radius:50%;background:var(--dim);flex:none}
.check i.ok{background:var(--ok)}.check i.no{background:var(--p2)}
pre.rec{max-height:62vh;overflow:auto;background:#05080F;color:#D7E2F2;padding:12px;border-radius:8px;font-family:var(--mono);font-size:12.5px;white-space:pre-wrap;margin:0}
pre.rec mark{background:#F2C76E;color:#111}
/* ---------- modal, palette, toast, drop ---------- */
.overlay{position:fixed;inset:0;z-index:10000;background:rgba(3,8,16,.66);backdrop-filter:blur(2px);display:flex;align-items:flex-start;justify-content:center;padding:6vh 16px 24px;overflow:auto}
.dialog{width:min(640px,100%);background:var(--panel);border:1px solid var(--line2);border-radius:14px;box-shadow:var(--shadow);animation:pop .14s ease-out}
@keyframes pop{from{transform:translateY(6px);opacity:.6}to{transform:none;opacity:1}}
.dialog.wide{width:min(980px,100%)}
.dialog.full{width:min(1320px,100%)}
.dialog>header{display:flex;align-items:center;justify-content:space-between;gap:10px;padding:14px 18px;border-bottom:1px solid var(--line)}
.dialog>header h3{font-size:16px}
.dialog .body{padding:16px 18px;display:grid;gap:12px;grid-template-columns:1fr 1fr}
.dialog .body .full{grid-column:1/-1}
.dialog .body p{margin:0}
.dialog footer{display:flex;justify-content:flex-end;gap:8px;padding:12px 18px;border-top:1px solid var(--line)}
.field{display:flex;flex-direction:column;gap:5px}
.field label{font-size:12.5px;color:var(--muted)}
.field label em{color:var(--p1);font-style:normal}
.field .hint{font-size:12px;color:var(--dim)}
.field input,.field select{width:100%}
.group{grid-column:1/-1;margin:6px 0 -4px;padding-top:10px;border-top:1px solid var(--line);font-weight:600;font-size:13px}
.metaline{grid-column:1/-1;font-size:12px;color:var(--dim)}
.colrow{display:flex;gap:6px;align-items:center;margin-bottom:6px}
.colrow input[type=text]{flex:1}
.colrow .tag{font-size:11px;color:var(--dim);width:48px}
.colrow label{font-size:12px;color:var(--muted);display:flex;gap:4px;align-items:center;white-space:nowrap}
.pv-sec{border:1px solid var(--line);border-radius:10px;margin-bottom:14px;overflow:hidden}
.pv-sec>header{display:flex;flex-wrap:wrap;gap:10px;align-items:center;padding:10px 12px;background:var(--panel2)}
.hist{border-top:1px solid var(--line);padding:10px 0}
.hist:first-child{border-top:0}
.hist table{font-size:12.5px;border-collapse:collapse;margin-top:6px;width:100%}
.hist td{padding:3px 6px;border-bottom:1px solid var(--line);vertical-align:top}
.hist del{color:var(--p1);text-decoration:line-through}
.hist ins{color:var(--ok);text-decoration:none}
.run-step{display:grid;grid-template-columns:70px 1fr auto;gap:12px;padding:10px 0;border-top:1px solid var(--line)}
.run-step.cur{background:var(--accent-soft);margin:0 -10px;padding:10px;border-radius:8px}
.run-step .n{font-family:var(--mono);color:var(--muted)}
.run-step .btns{display:flex;gap:6px;flex-wrap:wrap;align-items:flex-start;justify-content:flex-end}
.run-head{display:flex;gap:18px;flex-wrap:wrap;align-items:center;font-size:13px}
.palette{width:min(660px,100%);background:var(--panel);border:1px solid var(--line2);border-radius:14px;box-shadow:var(--shadow);overflow:hidden}
.palette input{width:100%;border:0;border-bottom:1px solid var(--line);border-radius:0;padding:15px 16px;font-size:15px;background:var(--panel)}
.palette input:focus{box-shadow:none}
.palette ul{list-style:none;margin:0;padding:6px;max-height:56vh;overflow:auto}
.palette li{display:flex;gap:10px;align-items:center;padding:8px 10px;border-radius:8px;cursor:pointer}
.palette li.on,.palette li:hover{background:var(--hover)}
.palette li .k{font-size:11px;color:var(--dim);width:84px;flex:none}
.palette li small{color:var(--muted);margin-left:auto;font-size:12px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;max-width:45%}
#toasts{position:fixed;right:16px;bottom:16px;z-index:20000;display:flex;flex-direction:column;gap:8px;max-width:calc(100vw - 32px)}
.toast{background:var(--panel);border:1px solid var(--line2);border-left:3px solid var(--accent2);border-radius:9px;padding:10px 14px;max-width:420px;box-shadow:var(--shadow);animation:pop .18s ease-out;display:flex;gap:12px;align-items:center}
.toast.success{border-left-color:var(--ok)}
.toast.error{border-left-color:var(--p1)}
#drop{position:fixed;inset:0;z-index:15000;display:none;align-items:center;justify-content:center;background:rgba(5,12,22,.86);border:2px dashed var(--accent);font-size:18px;color:#fff}
#drop.on{display:flex}
/* ---------- sign-in ---------- */
.login{position:fixed;inset:0;z-index:9000;display:grid;grid-template-columns:minmax(0,1.1fr) minmax(360px,1fr);background:var(--bg)}
.login .art{background:#0A1120;color:#DCE6F5;padding:48px;display:flex;flex-direction:column;justify-content:space-between;position:relative;overflow:hidden}
.login .art::after{content:"";position:absolute;inset:0;background-image:linear-gradient(rgba(45,212,191,.08) 1px,transparent 1px),linear-gradient(90deg,rgba(45,212,191,.08) 1px,transparent 1px);background-size:44px 44px;mask-image:radial-gradient(circle at 30% 70%,#000,transparent 70%);pointer-events:none}
.login .art h1{font-size:40px;line-height:1.1;font-weight:500;letter-spacing:-.03em;margin:0;max-width:520px;position:relative;z-index:1}
.login .art p{color:#9FB0CC;max-width:460px;position:relative;z-index:1}
.login .art .brand{padding:0;position:relative;z-index:1}
.login .form{display:flex;align-items:center;justify-content:center;padding:32px}
.login .card{width:min(400px,100%);display:flex;flex-direction:column;gap:12px}
.login .card h2{font-size:24px}
.login .card input{width:100%;padding:10px 12px;font-size:14px}
.login .card .btn{padding:10px 14px;font-size:14px}
.login .msg{font-size:13px;padding:9px 11px;border-radius:8px;background:color-mix(in srgb,var(--p1) 12%,transparent);color:var(--p1)}
.login .msg.info{background:var(--accent-soft);color:var(--accent)}
.login .or{display:flex;align-items:center;gap:10px;color:var(--dim);font-size:12px}
.login .or::before,.login .or::after{content:"";flex:1;height:1px;background:var(--line)}
.ms{display:flex;align-items:center;justify-content:center;gap:10px;text-decoration:none}
.ms i{display:grid;grid-template-columns:8px 8px;gap:2px}
.ms i b{width:8px;height:8px;display:block}
@media (max-width:900px){
  .app,.app.collapsed{grid-template-columns:1fr}.side{position:fixed;z-index:60;width:248px;transform:translateX(-100%);transition:transform .18s}
  .app.mobile-open .side{transform:none}
  .login{grid-template-columns:1fr}.login .art{display:none}
  .sun{grid-template-columns:1fr}.tile{border-right:0;border-bottom:1px solid var(--line)}
  .dialog .body{grid-template-columns:1fr} main#view{padding:16px}.store,.searchbtn kbd{display:none}.searchbtn{min-width:0}
}
@media (prefers-reduced-motion:reduce){*{transition:none!important;animation:none!important}.ticker .rail{overflow-x:auto}}
@media print{.side,.topbar,#ticker,.toolbar,.task-tools,.actions,#toasts{display:none!important}.app{display:block}}
.ro-legend{display:flex;flex-wrap:wrap;gap:8px;margin:0 0 10px}
.ro-chip{display:inline-block;padding:2px 9px;border-radius:12px;font-size:12px;font-weight:700;color:#0b1324;cursor:pointer;border:1px solid rgba(0,0,0,.15);white-space:nowrap}
.ro-chip.off{opacity:.3;text-decoration:line-through}
#roTable td,#roTable th{vertical-align:top}
.ro-cell{text-align:center;font-weight:700;min-width:40px;color:#0b1324}
.ro-edit{cursor:pointer}.ro-edit:hover{outline:2px solid var(--acc);outline-offset:-2px}
.ro-short{box-shadow:inset 0 0 0 2px var(--p1)}.ro-short .ro-n{color:var(--p1);font-weight:700;opacity:1}
tr.ro-hol>td,tr.ro-hol>th{background:color-mix(in srgb,#F59345 14%,transparent)}th.ro-holh{background:color-mix(in srgb,#F59345 30%,transparent)}
tr.ro-we>th{opacity:.85}.ro-hb{display:block;font-size:11px;color:var(--care);font-weight:700}
.ro-n{font-size:11px;opacity:.7;margin-left:4px}tr.ro-nodata>td{font-style:italic;opacity:.65}tr.ro-nodata>th{opacity:.65}.ro-picks{display:grid;grid-template-columns:repeat(auto-fill,minmax(220px,1fr));gap:6px 14px}.ro-dim{opacity:.15}.ro-log{font-size:11px}
/* ===================== MSP Ops Center design system, v5 ===================== */
@font-face{font-family:"IBM Plex Sans";font-weight:400;font-display:swap;src:url(/static/fonts/ibm-plex-sans-latin-400-normal.woff2) format("woff2")}
@font-face{font-family:"IBM Plex Sans";font-weight:500;font-display:swap;src:url(/static/fonts/ibm-plex-sans-latin-500-normal.woff2) format("woff2")}
@font-face{font-family:"IBM Plex Sans";font-weight:600;font-display:swap;src:url(/static/fonts/ibm-plex-sans-latin-600-normal.woff2) format("woff2")}
@font-face{font-family:"IBM Plex Sans Condensed";font-weight:500;font-display:swap;src:url(/static/fonts/ibm-plex-sans-condensed-latin-500-normal.woff2) format("woff2")}
@font-face{font-family:"IBM Plex Sans Condensed";font-weight:600;font-display:swap;src:url(/static/fonts/ibm-plex-sans-condensed-latin-600-normal.woff2) format("woff2")}
@font-face{font-family:"IBM Plex Mono";font-weight:400;font-display:swap;src:url(/static/fonts/ibm-plex-mono-latin-400-normal.woff2) format("woff2")}
@font-face{font-family:"IBM Plex Mono";font-weight:500;font-display:swap;src:url(/static/fonts/ibm-plex-mono-latin-500-normal.woff2) format("woff2")}
:root{
  --bg:#070D19;--bg2:#0A1221;--side:#050A14;--panel:#0D1627;--panel2:#121E34;--hover:#16243D;--line:#1A2841;--line2:#26385A;
  --text:#E8EEF8;--muted:#A3B1C9;--dim:#8291AB;--accent:#2DD4BF;--accent-soft:rgba(45,212,191,.12);--accent2:#8AB4FF;--ink:#032220;
  --display:"IBM Plex Sans Condensed","IBM Plex Sans","Segoe UI",system-ui,sans-serif;
  --r-sm:7px;--r:10px;--r-lg:14px;
  --e1:inset 0 1px 0 rgba(255,255,255,.035),0 1px 2px rgba(0,0,0,.35);
  --e2:inset 0 1px 0 rgba(255,255,255,.04),0 10px 30px -12px rgba(0,0,0,.6);
  --glow:radial-gradient(1200px 520px at 12% -8%,rgba(45,212,191,.075),transparent 60%),radial-gradient(900px 480px at 100% -10%,rgba(122,167,255,.06),transparent 55%);
}
html[data-theme="light"]{
  --bg:#F2F5F9;--bg2:#FFFFFF;--side:#0A1323;--panel:#FFFFFF;--panel2:#F5F7FB;--hover:#EDF1F7;--line:#E1E6EF;--line2:#CBD3E1;
  --text:#101B30;--muted:#4B5A75;--dim:#5E6C86;--accent:#08776B;--accent-soft:rgba(8,119,107,.10);--accent2:#2459C2;--ink:#FFFFFF;
  --p1:#B42318;--p2:#A04C0A;--p3:#7A5C00;--p4:#1D55B0;--night:#5B3EC2;--care:#8F5A00;
  --e1:0 1px 2px rgba(16,27,48,.06);--e2:0 12px 30px -14px rgba(16,27,48,.22);
  --glow:radial-gradient(1200px 520px at 12% -8%,rgba(11,140,126,.07),transparent 60%);
}
body{background:var(--bg) var(--glow) no-repeat fixed;font-feature-settings:"ss01","tnum" 0;-webkit-font-smoothing:antialiased;text-rendering:optimizeLegibility}
.grid td,.grid th,.n,.cnt,.tile .nums b,.ro-n,.badge,kbd,.metaline{font-variant-numeric:tabular-nums}
::selection{background:color-mix(in srgb,var(--accent) 35%,transparent)}
/* side navigation */
.side{background:linear-gradient(180deg,#071021 0%,var(--side) 60%);border-right:1px solid rgba(255,255,255,.05)}
.brand{padding:18px 16px 16px}
.brand .logo{border-radius:10px;background:conic-gradient(from 210deg,#2DD4BF,#3B82F6,#2DD4BF);box-shadow:0 0 0 1px rgba(255,255,255,.12) inset,0 6px 18px -6px rgba(45,212,191,.55)}
.brand b{font-family:var(--display);font-size:16px;letter-spacing:.01em}
.nav .grp{font-size:10.5px;font-weight:600;letter-spacing:.09em;text-transform:uppercase;color:#7F8FAC;margin:16px 10px 6px}
.blade{border-radius:var(--r);padding:6px 9px}
.blade .ic{border-radius:8px;background:transparent}
.blade:hover{background:rgba(255,255,255,.045)}
.blade.on{background:linear-gradient(90deg,rgba(45,212,191,.16),rgba(45,212,191,.04));box-shadow:inset 0 0 0 1px rgba(45,212,191,.18)}
.blade.on .ic{background:#2DD4BF;color:#032220;box-shadow:0 4px 14px -4px rgba(45,212,191,.7)}
/* top bar */
.topbar{padding:10px 26px;gap:14px;background:color-mix(in srgb,var(--bg) 82%,transparent);backdrop-filter:saturate(1.4) blur(10px)}
.crumbs{font-size:12.5px}
.crumbs b{font-family:var(--display);font-size:15px;letter-spacing:.01em}
.searchbtn{border-radius:var(--r);min-width:270px;background:color-mix(in srgb,var(--panel) 80%,transparent);box-shadow:var(--e1)}
.searchbtn:hover{border-color:color-mix(in srgb,var(--accent) 45%,var(--line2));color:var(--muted)}
.onshift{display:none;align-items:center;gap:10px;padding:5px 12px 5px 8px;border-radius:999px;border:1px solid var(--line2);background:color-mix(in srgb,var(--panel) 85%,transparent);color:var(--text);font:inherit;font-size:12.5px;cursor:pointer;max-width:420px;white-space:nowrap;box-shadow:var(--e1)}
.onshift.show{display:inline-flex}
.onshift:hover{border-color:var(--sc,var(--accent))}
.onshift .pulse{width:9px;height:9px;border-radius:50%;background:var(--sc,var(--accent));box-shadow:0 0 0 0 var(--sc,var(--accent));animation:pulse 2.4s ease-out infinite;flex:none}
@keyframes pulse{0%{box-shadow:0 0 0 0 color-mix(in srgb,var(--sc,var(--accent)) 60%,transparent)}70%{box-shadow:0 0 0 8px transparent}100%{box-shadow:0 0 0 0 transparent}}
.onshift b{font-family:var(--display);font-weight:600;letter-spacing:.02em}
.onshift .who{color:var(--muted);overflow:hidden;text-overflow:ellipsis}
.onshift .left{font-size:11.5px;color:var(--dim);font-variant-numeric:tabular-nums}
.askbtn{display:inline-flex;align-items:center;gap:8px;height:34px;padding:0 13px 0 11px;border-radius:var(--r);border:1px solid color-mix(in srgb,var(--accent) 55%,transparent);background:linear-gradient(180deg,color-mix(in srgb,var(--accent) 20%,transparent),color-mix(in srgb,var(--accent) 8%,transparent));color:var(--text);font:inherit;font-size:13px;font-weight:600;cursor:pointer;box-shadow:var(--e1)}
.askbtn svg{width:16px;height:16px;color:var(--accent)}
.askbtn:hover{border-color:var(--accent);background:color-mix(in srgb,var(--accent) 22%,transparent)}
.askbtn kbd{font-family:var(--mono);font-size:10.5px;font-weight:500;padding:0 5px;border:1px solid var(--line2);border-radius:5px;color:var(--muted)}
/* page type */
main#view{padding:26px 30px 48px}
h2{font-family:var(--display);font-size:26px;font-weight:600;letter-spacing:-.005em;line-height:1.15}
.hello h2{font-size:30px}
.sub{font-size:14px;line-height:1.55}
/* surfaces */
.panel,.tablewrap,.task,.sec,.entry,.check{border-radius:var(--r-lg);border-color:var(--line);box-shadow:var(--e1);background:var(--panel)}
.panel{padding:16px 18px}
.panel h3{font-size:11.5px;font-weight:600;text-transform:uppercase;letter-spacing:.08em;color:var(--muted);margin-bottom:6px;display:flex;align-items:center;gap:8px}
.panel h3 .n{font-size:11px;font-weight:600;letter-spacing:0;padding:1px 7px;border-radius:999px;background:var(--accent-soft);color:var(--accent)}
.panel h3 .n:empty{display:none}
/* a plan shared with your team: its structure and go/no-go stay with the owning team */
.task.ro [data-act=phase-clone],.task.ro [data-act=phase-del],.task.ro [data-act=phase-edit],.task.ro [data-act=phase-move],
.task.ro [data-act=phase-sign],.task.ro [data-act=step-add],.task.ro [data-act=step-clone],.task.ro [data-act=step-del],
.task.ro [data-act=step-move],.task.ro [data-act=step-cols]{display:none}
.field label:has(>input[type=checkbox]){display:inline-flex;align-items:center;gap:8px;width:auto;margin-top:6px;font-size:13px;color:var(--text);cursor:pointer}
.field input[type=checkbox]{width:16px;height:16px;flex:none;margin:0;accent-color:var(--accent)}
#roTable .ro-log{display:inline-flex;align-items:center;margin-left:4px;padding:0 7px;height:18px;border-radius:999px;border:1px solid var(--line2);background:transparent;color:var(--accent2);font-size:10.5px;font-weight:600;letter-spacing:.03em;text-transform:uppercase;vertical-align:1px;cursor:pointer}
#roTable .ro-log:hover{border-color:var(--accent2);background:color-mix(in srgb,var(--accent2) 12%,transparent)}
#roHost .toolbar .btn:not(.primary){padding:6px 11px;font-size:12.5px}
#roHost .toolbar{row-gap:8px}
.list li{padding:10px 0}
.list .t{font-size:13.5px}
.cols{gap:16px}
.tile{border-radius:var(--r-lg);box-shadow:var(--e1)}
.tile h3{font-family:var(--display);font-size:17px}
.tile .nums b{font-family:var(--display);font-size:24px}
.sun{border-radius:var(--r-lg);overflow:hidden;box-shadow:var(--e2)}
.sun .tile{border-radius:0;box-shadow:none}
.tile .rost{margin-top:10px;padding-top:8px;border-top:1px dashed var(--line2);font-size:12.5px;color:var(--muted)}
.tile .rost b{color:var(--text);font-weight:500}
/* tables */
.grid th{font-size:11px;font-weight:600;text-transform:uppercase;letter-spacing:.07em;color:var(--muted);background:var(--panel2);padding-top:10px;padding-bottom:10px}
.grid td{border-color:var(--line)}
.grid tbody tr:hover td{background:color-mix(in srgb,var(--accent) 5%,var(--panel))}
/* controls */
.btn{border-radius:var(--r-sm);padding:7px 13px;font-weight:500;box-shadow:var(--e1)}
.btn.primary{background:linear-gradient(180deg,#3BE0CB,#22C3AE);border-color:#22C3AE;color:var(--ink);box-shadow:inset 0 1px 0 rgba(255,255,255,.35),0 6px 16px -8px rgba(45,212,191,.8)}
html[data-theme="light"] .btn.primary{background:linear-gradient(180deg,#0A8276,#08776B);border-color:#08776B}
.btn.primary:hover{filter:brightness(1.06)}
.btn.ghost,.btn.xs{box-shadow:none}
input,select,textarea{border-radius:var(--r-sm)}
.badge,.stb,.chip{border-radius:999px}
.dialog{border-radius:18px;box-shadow:0 30px 80px -20px rgba(0,0,0,.7),inset 0 1px 0 rgba(255,255,255,.05)}
.dialog>header h3{font-family:var(--display);font-size:18px}
.toast{border-radius:var(--r);box-shadow:var(--e2)}
/* sign-in */
.login .art h1{font-family:var(--display);font-weight:600;letter-spacing:-.01em}
.login .card{border-radius:18px}
/* ===================== assistant ===================== */
.ask-scrim{position:fixed;inset:0;background:rgba(3,7,15,.45);backdrop-filter:blur(2px);z-index:85;opacity:0;pointer-events:none;transition:opacity .18s}
.ask-scrim.on{opacity:1;pointer-events:auto}
.ask{position:fixed;top:0;right:0;bottom:0;width:min(560px,100vw);z-index:90;display:flex;flex-direction:column;background:var(--bg2);border-left:1px solid var(--line2);box-shadow:-30px 0 80px -30px rgba(0,0,0,.7);transform:translateX(102%);transition:transform .22s cubic-bezier(.2,.8,.2,1)}
.ask.on{transform:none}
.ask header{display:flex;align-items:center;gap:12px;padding:16px 18px;border-bottom:1px solid var(--line)}
.ask header .mark{width:34px;height:34px;border-radius:10px;display:grid;place-items:center;background:conic-gradient(from 210deg,#2DD4BF,#3B82F6,#2DD4BF);color:#032220;flex:none}
.ask header .mark svg{width:18px;height:18px}
.ask header h3{font-family:var(--display);font-size:17px;font-weight:600;margin:0}
.ask header small{display:block;color:var(--muted);font-size:12px}
.ask header .sp{flex:1}
.ask .mode{font-size:11px;font-weight:600;padding:2px 8px;border-radius:999px;border:1px solid var(--line2);color:var(--muted);white-space:nowrap}
.ask .mode.model{border-color:color-mix(in srgb,var(--accent) 55%,transparent);color:var(--accent)}
.ask .log{flex:1;overflow-y:auto;padding:18px;display:flex;flex-direction:column;gap:16px}
.ask .welcome h4{font-family:var(--display);font-size:22px;font-weight:600;margin:4px 0 6px}
.ask .welcome p{color:var(--muted);margin:0 0 14px}
.ask .sugg{display:flex;flex-wrap:wrap;gap:8px}
.ask .sugg button{border:1px solid var(--line2);background:var(--panel);color:var(--text);border-radius:999px;padding:6px 12px;font:inherit;font-size:12.5px;cursor:pointer}
.ask .sugg button:hover{border-color:var(--accent)}
.ask .q{align-self:flex-end;max-width:85%;background:color-mix(in srgb,var(--accent) 16%,var(--panel));border:1px solid color-mix(in srgb,var(--accent) 30%,transparent);border-radius:14px 14px 4px 14px;padding:9px 13px;font-size:13.5px}
.ask .a{background:var(--panel);border:1px solid var(--line);border-radius:14px 14px 14px 4px;padding:14px 15px;box-shadow:var(--e1);font-size:13.5px;line-height:1.55}
.ask .a p{margin:0 0 8px}.ask .a p:last-child{margin-bottom:0}
.ask .a ul{margin:0 0 8px;padding-left:18px}.ask .a li{margin:2px 0}
.ask .a pre{position:relative;margin:4px 0 10px;padding:10px 12px;background:var(--bg);border:1px solid var(--line);border-radius:var(--r-sm);font-family:var(--mono);font-size:12.5px;white-space:pre-wrap;word-break:break-word}
.ask .a pre button{position:absolute;top:6px;right:6px}
.ask .a.nf{border-style:dashed}
.ask .cite{display:inline-grid;place-items:center;min-width:18px;height:18px;padding:0 4px;margin:0 1px;border-radius:5px;border:0;background:var(--accent-soft);color:var(--accent);font:600 11px/1 var(--sans);cursor:pointer;vertical-align:1px}
.ask .cite:hover{background:var(--accent);color:var(--ink)}
.ask .srcs{margin-top:12px;padding-top:10px;border-top:1px solid var(--line);display:grid;gap:6px}
.ask .srcs .h{font-size:10.5px;font-weight:600;text-transform:uppercase;letter-spacing:.08em;color:var(--dim)}
.ask .src{display:flex;gap:10px;align-items:flex-start;text-align:left;width:100%;border:1px solid transparent;background:none;color:inherit;font:inherit;padding:6px 8px;border-radius:var(--r-sm);cursor:pointer}
.ask .src:hover,.ask .src.hl{background:var(--panel2);border-color:var(--line2)}
.ask .src .num{flex:none;min-width:20px;height:20px;border-radius:6px;display:grid;place-items:center;background:var(--accent-soft);color:var(--accent);font-size:11px;font-weight:600}
.ask .src .kind{font-size:10.5px;font-weight:600;text-transform:uppercase;letter-spacing:.06em;color:var(--dim);display:block}
.ask .src .ttl{font-size:13px}
.ask .src .whr{display:block;font-size:12px;color:var(--muted)}
.ask .meta{margin-top:10px;font-size:11.5px;color:var(--dim)}
.ask .note{margin-top:8px;font-size:12px;color:var(--care)}
.ask .think{display:flex;gap:10px;align-items:center;color:var(--muted);font-size:13px}
.ask .think i{width:7px;height:7px;border-radius:50%;background:var(--accent);animation:blink 1.2s infinite both}
.ask .think i:nth-child(2){animation-delay:.2s}.ask .think i:nth-child(3){animation-delay:.4s}
@keyframes blink{0%,80%,100%{opacity:.25}40%{opacity:1}}
.ask form{padding:14px 18px 16px;border-top:1px solid var(--line);background:var(--bg2)}
.ask .box{display:flex;gap:8px;align-items:flex-end;border:1px solid var(--line2);border-radius:14px;background:var(--panel);padding:8px 8px 8px 12px;box-shadow:var(--e1)}
.ask .box:focus-within{border-color:var(--accent);box-shadow:0 0 0 3px var(--accent-soft)}
.ask textarea{flex:1;border:0;background:transparent;resize:none;min-height:22px;max-height:140px;padding:4px 0;box-shadow:none;font-size:14px}
.ask textarea:focus{box-shadow:none}
.ask .send{width:34px;height:34px;flex:none;border-radius:10px;border:0;background:var(--accent);color:var(--ink);cursor:pointer;display:grid;place-items:center}
.ask .send[disabled]{opacity:.4;cursor:not-allowed}
.ask .send svg{width:16px;height:16px}
.ask .fine{margin:8px 2px 0;font-size:11.5px;color:var(--dim)}
.topbar>*{flex-shrink:0}
.crumbs{flex:0 1 auto;min-width:120px}
.searchbtn{margin-left:auto;flex:0 1 300px;min-width:180px;white-space:nowrap;overflow:hidden}
.searchbtn kbd,.askbtn kbd{white-space:nowrap}
.onshift{flex:0 1 auto;min-width:0}
.store{flex:0 1 auto;min-width:0}
@media (max-width:1560px){.store .loc,.store b{display:none}.store{gap:0}}
@media (max-width:1320px){.onshift .who{display:none}.searchbtn kbd{display:none}}
@media (max-width:1120px){.searchbtn{flex-basis:44px;min-width:44px;font-size:0;gap:0}.searchbtn svg{width:16px;height:16px}}
@media (max-width:900px){.onshift{display:none!important}.askbtn kbd{display:none}}
@media (prefers-reduced-motion:reduce){.ask,.ask-scrim{transition:none}.onshift .pulse,.ask .think i{animation:none}}
@media print{.ask,.ask-scrim,.askbtn,.onshift{display:none!important}}

</style>
</head>
<body class="r0">
<div id="login" class="login hidden"></div>
<div class="app hidden" id="app">
  <aside class="side" id="side" aria-label="Main navigation">
    <div class="brand"><div class="logo" aria-hidden="true">M</div><div><b>MSP Ops Center</b><span>ITIL v4 service operations</span></div></div>
    <nav class="nav" id="nav"></nav>
    <div class="side-foot" id="sideFoot"></div>
  </aside>
  <div class="main">
    <header class="topbar">
      <button class="tb-toggle" data-act="side-toggle" aria-label="Collapse or expand the menu"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M4 6h16M4 12h16M4 18h16"/></svg></button>
      <div class="crumbs" id="crumbs"></div>
      <button class="onshift" id="onShift" data-act="onshift" aria-label="Shift running now"></button>
      <button class="searchbtn" data-act="palette" aria-label="Search everything"><svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="11" cy="11" r="7"/><path d="m20 20-3.5-3.5"/></svg>Search devices, tickets, tasks<kbd>Ctrl K</kbd></button>
      <button class="askbtn" data-act="ask-open" aria-label="Ask MSP Ops Center (Ctrl J)"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M12 3l1.8 4.6L18.5 9l-4.7 1.6L12 15l-1.8-4.4L5.5 9l4.7-1.4z"/><path d="M19 15l.9 2.1L22 18l-2.1.9L19 21l-.9-2.1L16 18l2.1-.9z"/></svg>Ask<kbd>Ctrl J</kbd></button>
      <div class="store" id="storeChip"><span class="dot" id="syncDot"></span><span class="loc">Connecting…</span></div>
      <button class="btn sm" data-act="import" data-module="auto">Import</button>
      <button class="btn sm" data-act="export-menu">Export</button>
    </header>
    <div id="ticker"></div>
    <main id="view"></main>
    <section id="termHost" class="hidden" aria-label="Terminals"></section>
  </div>
</div>
<div id="modalRoot"></div>
<div id="toasts" role="status" aria-live="polite"></div>
<div id="drop">Drop an .xlsx, .csv, .txt or .json file to import it</div>
<script>
"use strict";
const $ = s => document.querySelector(s);
function today(){return new Date().toLocaleDateString('en-CA');}
function currentShift(){const h=new Date().getHours();return h<6?'Night':h<14?'Morning':h<22?'Afternoon':'Night';}
function nowLocal(){const d=new Date(),p=n=>String(n).padStart(2,'0');return `${d.getFullYear()}-${p(d.getMonth()+1)}-${p(d.getDate())} ${p(d.getHours())}:${p(d.getMinutes())}`;}
const S = {
  user:'', me:null, rank:0, csrf:'', caps:{}, roleHelp:{}, adminTab:'users', terms:[], activeTerm:null, started:false,
  tables:{}, plans:{tasks:[],templates:[],builtin:[],presets:[],step_statuses:[],task_statuses:[]}, hier:{years:{}},
  audit:[], imports:[], schema:{}, reads:new Set(), revision:null, status:null, identity:null,
  view:(location.hash||'#mywork').slice(1).split('=')[0], filters:{}, open:new Set(JSON.parse(localStorage.getItem('msp.open')||'[]')),
  modals:[], hDate:today(), hShift:currentShift(), carry:true, rbFocus:null, runTask:null
};
const SHIFTS=['Morning','Afternoon','Night'];
const SHIFT_COLOR={Morning:'var(--morning)',Afternoon:'var(--afternoon)',Night:'var(--night)'};
const CLOSED=['resolved','closed','cancelled'];
const SEV_RANK={Critical:3,Advisory:2,Info:1};
const ICONS={
  home:'<path d="M3 11 12 4l9 7"/><path d="M5 10v10h14V10"/><path d="M10 20v-6h4v6"/>',
  grid:'<rect x="3" y="3" width="7" height="7" rx="1.5"/><rect x="14" y="3" width="7" height="7" rx="1.5"/><rect x="3" y="14" width="7" height="7" rx="1.5"/><rect x="14" y="14" width="7" height="7" rx="1.5"/>',
  plan:'<path d="M9 5h11M9 12h11M9 19h11"/><path d="m3 5 1.5 1.5L7 4M3 12l1.5 1.5L7 11M3 19l1.5 1.5L7 18"/>',
  cal:'<rect x="3" y="5" width="18" height="16" rx="2"/><path d="M3 10h18M8 3v4M16 3v4"/>',
  swap:'<path d="M4 8h13l-3-3M20 16H7l3 3"/>',
  term:'<rect x="3" y="4" width="18" height="16" rx="2"/><path d="m7 9 3 3-3 3M12 15h5"/>',
  book:'<path d="M4 5a2 2 0 0 1 2-2h13v16H6a2 2 0 0 0-2 2z"/><path d="M4 19V5M9 7h6"/>',
  server:'<rect x="3" y="4" width="18" height="7" rx="1.5"/><rect x="3" y="13" width="18" height="7" rx="1.5"/><path d="M7 7.5h.01M7 16.5h.01"/>',
  alert:'<path d="M12 3 2 20h20z"/><path d="M12 10v4M12 17h.01"/>',
  mega:'<path d="M3 11v2a1 1 0 0 0 1 1h3l6 4V6L7 10H4a1 1 0 0 0-1 1z"/><path d="M17 8a5 5 0 0 1 0 8"/>',
  pulse:'<path d="M3 12h4l3-8 4 16 3-8h4"/>',
  shield:'<path d="M12 3 4 6v6c0 5 3.5 8 8 9 4.5-1 8-4 8-9V6z"/><path d="m9 12 2 2 4-4"/>',
  sun:'<circle cx="12" cy="12" r="4"/><path d="M12 2v2M12 20v2M2 12h2M20 12h2M4.9 4.9l1.4 1.4M17.7 17.7l1.4 1.4M4.9 19.1l1.4-1.4M17.7 6.3l1.4-1.4"/>',
  key:'<circle cx="8" cy="15" r="4"/><path d="m11 12 9-9M16 7l3 3"/>',
  out:'<path d="M15 4h4v16h-4M10 17l5-5-5-5M15 12H3"/>',
  roster:'<rect x="3" y="4" width="18" height="17" rx="2"/><path d="M3 9h18M8 2v4M16 2v4M7 13h3M7 17h3M14 13h3M14 17h3"/>',
  briefcase:'<rect x="3" y="7" width="18" height="13" rx="2"/><path d="M9 7V5a2 2 0 0 1 2-2h2a2 2 0 0 1 2 2v2M3 13h18"/>',
  users:'<circle cx="9" cy="8" r="3.2"/><path d="M3.5 19c.6-3.2 2.8-5 5.5-5s4.9 1.8 5.5 5"/><circle cx="17" cy="9" r="2.6"/><path d="M15.5 14.2c2.4.2 4.2 1.8 4.9 4.8"/>',
  spark:'<path d="M12 3l1.8 4.6L18.5 9l-4.7 1.6L12 15l-1.8-4.4L5.5 9l4.7-1.4z"/><path d="M19 15l.9 2.1L22 18l-2.1.9L19 21l-.9-2.1L16 18l2.1-.9z"/>'};
const svg=n=>`<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${ICONS[n]}</svg>`;
const NAV=[['Home',[['mywork','My work','home'],['overview','Overview','grid']]],
  ['Operations',[['plans','Execution plans','plan'],['shiftlog','Shift log','cal'],['Shift_Transitions','Shift handovers','swap'],['roster','Shift roster','roster'],['terminals','Terminals','term',2]]],
  ['Knowledge',[['runbook','Runbook','book'],['Asset_Master','CMDB','server'],['Incidents','Incidents','alert']]],
  ['Communication',[['Bulletins','Bulletins','mega']]],
  ['Service relationships',[['customers','Customers and services','briefcase']]],
  ['Administration',[['teams','Teams','users'],['activity','Activity','pulse'],['admin','Admin','shield',4]]]];
const VIEWS=NAV.flatMap(g=>g[1]);
const SUBS={Asset_Master:'Configuration management database: every configuration item (CI) with its type, status, platform, sizing, dependencies and special-care caveats. Select a CI name to open its runbook.',
  Incidents:'Incident and problem management. Every record follows ART: Action, Reason, Technical tracking.',
  Shift_Transitions:'Follow-the-sun handover summaries between regional shifts.',
  Bulletins:'Notices shown in the scrolling ticker. They disappear after their expiry date.'};
const LONG=new Set(['Special_Care','Action','Reason','Note','Solution','Summary','Body','Upstream_Dependencies','Downstream_Dependencies','P1_P2_In_Flight','Planned_Changes','High_Risk_Assets','Notes','Contacts']);
const MONO=new Set(['Commands','Technical_Tracking']);
const DT_KEYS=new Set(['Start','Expiry','Opened_At','Resolved_At']);
const HINTS={Action:'Direct remediation and stabilisation steps taken.',Reason:'Root cause, trigger event, dependency validation.',
  Technical_Tracking:'Exact commands run and verification proof.',Device:'Device / server name. Links the asset to its runbook.',
  Serial_No:'Serial number, asset tag or cloud instance ID. Used to detect duplicates.'};

/* ---------- utilities ---------- */
function esc(v){return String(v ?? '').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));}
function label(k){return String(k).replace(/^x_/,'').replace(/_/g,' ');}
function nk(v){return String(v??'').toLowerCase().replace(/[^a-z0-9]/g,'');}
function uniq(a){return [...new Set(a.filter(x=>x!==undefined&&x!==null&&String(x).trim()!==''))].sort((x,y)=>String(x).localeCompare(String(y)));}
function truthy(v){return ['yes','y','true','1','done','x'].includes(String(v).trim().toLowerCase());}
function isOpen(r){return !CLOSED.includes(String(r.Status||'').toLowerCase());}
function fmtTime(iso){if(!iso)return '';const d=new Date(iso);return isNaN(d)?iso:d.toLocaleString([], {dateStyle:'medium',timeStyle:'short'});}
function fmtMin(m){m=+m||0;if(!m)return '0m';const h=Math.floor(Math.abs(m)/60),r=Math.abs(m)%60;return (m<0?'-':'')+((h?h+'h ':'')+(r?r+'m':'')).trim();}
function hm(v,base){if(!v)return '';const [d,t]=String(v).split(' ');return d===base?t:`${d.slice(5)} ${t||''}`;}
const toInput=v=>v?String(v).replace(' ','T').slice(0,16):'';
const fromInput=v=>v?String(v).replace('T',' '):'';
function addDays(d,n){const x=new Date(d+'T00:00');x.setDate(x.getDate()+n);return x.toLocaleDateString('en-CA');}
function persistOpen(){localStorage.setItem('msp.open',JSON.stringify([...S.open].slice(-400)));}
function toast(msg,type='info',action){
  const el=document.createElement('div');el.className='toast '+type;
  const sp=document.createElement('span');sp.textContent=msg;el.appendChild(sp);
  if(action){const b=document.createElement('button');b.className='btn sm';b.textContent=action.label;b.onclick=()=>{el.remove();action.fn();};el.appendChild(b);}
  $('#toasts').appendChild(el);setTimeout(()=>el.remove(),action?12000:(type==='error'?6500:3500));
}
async function copyText(text){
  try{if(navigator.clipboard&&window.isSecureContext){await navigator.clipboard.writeText(text);}
      else{const t=document.createElement('textarea');t.value=text;t.style.position='fixed';t.style.opacity='0';document.body.appendChild(t);t.select();document.execCommand('copy');t.remove();}
      toast('Copied to clipboard','success');}catch(e){toast('Copy failed: select the text and copy it manually','error');}
}

/* ---------- API ---------- */
if(!S.team)S.team=localStorage.getItem('msp.team')||'';
async function api(method,path,body,opts={}){
  const url=path+(path.includes('?')?'&':'?')+'_ts='+Date.now();
  const headers={};
  if(method!=='GET'&&S.csrf)headers['X-CSRF-Token']=S.csrf;
  if(opts.bg)headers['X-Background']='1';
  if(S.team)headers['X-Team']=S.team;
  let payload=body;
  if(body!==undefined&&!(body instanceof FormData)){headers['Content-Type']='application/json';payload=JSON.stringify(body);}
  const r=await fetch(url,{method,headers,body:payload,cache:'no-store',credentials:'same-origin'});
  if(r.status===401&&!opts.noAuth){sessionEnded();const e=new Error('Your session has ended. Please sign in again.');e.status=401;e.shown=true;throw e;}
  if(opts.raw){if(!r.ok){let m='Request failed';try{m=(await r.json()).error||m;}catch(_){} const e=new Error(m);e.status=r.status;throw e;}return r;}
  let data;try{data=await r.json();}catch(_){data={ok:false,error:'Unexpected response from server ('+r.status+')'};}
  if(!r.ok||data.ok===false){const e=new Error(data.error||r.statusText);e.status=r.status;throw e;}
  return data;
}
const planOp=(op,body)=>api('POST','/api/plans/op',{op,...body});
async function run(fn,okMsg){
  try{const res=await fn();if(okMsg)toast(okMsg,'success');await load();return res;}
  catch(e){if(!e.shown){toast(e.message,'error');e.shown=true;}throw e;}
}
async function dupSafe(call){
  try{return await call(false);}
  catch(e){
    if(e.status===409&&/duplicate|already exists/i.test(e.message)){
      if(await confirmBox(e.message,'Save anyway',false))return await call(true);
      e.shown=true;throw e;
    }
    throw e;
  }
}
async function load(){
  const d=await api('GET','/api/state');
  Object.assign(S,{tables:d.tables,plans:d.plans,hier:d.hierarchy,audit:d.audit,imports:d.imports,schema:d.schema,
    reads:new Set(d.my_reads),revision:d.revision,status:d.status,
    teamScope:!!d.team_scope,team:d.team||'',myTeams:d.my_teams||[],legacy:d.legacy||null});
  if(S.team)localStorage.setItem('msp.team',S.team);
  renderChrome();render();if(S.runTask)renderRun();
}
async function poll(){
  if(document.hidden||S.modals.length||document.activeElement?.classList.contains('ie'))return;
  try{const st=await api('GET','/api/status',undefined,{bg:true});S.status=st;renderStore();
    if(S.revision&&st.revision!==S.revision){await load();toast('Updated with the latest changes from the team');}}
  catch(e){$('#syncDot').classList.add('err');}
}

/* ---------- modals (stackable) ---------- */
function openModal(o){
  const ov=document.createElement('div');ov.className='overlay';
  ov.innerHTML=`<div class="dialog ${o.size||''}" role="dialog" aria-modal="true" aria-label="${esc(o.title)}">
    <header><h3>${esc(o.title)}</h3><button class="btn ghost sm" data-mclose>Close</button></header>
    <div class="body">${o.body}</div>
    <footer>${o.footExtra||''}<button class="btn" data-mclose>${o.onSave?'Cancel':'Close'}</button>${o.onSave?`<button class="btn ${o.danger?'danger':'primary'}" data-msave>${esc(o.save||'Save')}</button>`:''}</footer></div>`;
  $('#modalRoot').appendChild(ov);
  const m={el:ov,o};S.modals.push(m);
  ov.addEventListener('mousedown',e=>{m.down=e.target===ov;});
  ov.addEventListener('click',async e=>{
    if((e.target===ov&&m.down)||e.target.closest('[data-mclose]'))return closeModal(m,'cancel');
    const sb=e.target.closest('[data-msave]');
    if(sb){sb.disabled=true;try{const keep=await o.onSave(ov);if(keep!==false)closeModal(m,'saved');}
      catch(err){if(err&&!err.shown&&err.message)toast(err.message,'error');}finally{sb.disabled=false;}}
  });
  if(o.onMount)o.onMount(ov,m);
  setTimeout(()=>{const f=ov.querySelector('[autofocus]')||ov.querySelector('.body input:not([type=checkbox]):not([type=file]),.body select,.body textarea');if(f)f.focus();},0);
  return m;
}
function closeModal(m,reason){m=m||S.modals[S.modals.length-1];const i=S.modals.indexOf(m);if(i<0)return;S.modals.splice(i,1);m.el.remove();if(m.o.onClose)m.o.onClose(reason);}
function confirmBox(message,action='Delete',danger=true){
  return new Promise(res=>{openModal({title:'Please confirm',body:`<p class="full" style="white-space:pre-wrap">${esc(message)}</p>`,save:action,danger,
    onSave:()=>{res(true);},onClose:r=>{if(r!=='saved')res(false);}});});
}
document.addEventListener('keydown',e=>{if(e.key==='Escape'&&S.modals.length&&!e.target.classList?.contains('ie'))closeModal();
  else if(e.key==='Escape'&&$('#app')?.classList.contains('mobile-open'))$('#app').classList.remove('mobile-open');});
document.addEventListener('click',e=>{const app=$('#app');if(app?.classList.contains('mobile-open')&&!e.target.closest('.side')&&!e.target.closest('[data-act=side-toggle]'))app.classList.remove('mobile-open');});

function fieldHtml(f){
  const id='f_'+Math.random().toString(36).slice(2,9),req=f.required?' <em>*</em>':'',v=f.value??'';
  let input;
  if(f.type==='select'){
    const opts=[...new Set(['',...(f.options||[]),v])];
    input=`<select id="${id}" data-field="${esc(f.key)}">${opts.map(o=>`<option ${o===v?'selected':''} value="${esc(o)}">${esc(o||'—')}</option>`).join('')}</select>`;
  }else if(f.type==='textarea'||f.type==='mono'){
    input=`<textarea id="${id}" data-field="${esc(f.key)}" class="${f.type==='mono'?'mono':''}" spellcheck="${f.type!=='mono'}">${esc(v)}</textarea>`;
  }else if(f.type==='checkbox'){
    input=`<label class="st"><input type="checkbox" id="${id}" data-field="${esc(f.key)}" data-bool="1" ${v?'checked':''}> ${esc(f.text||'')}</label>`;
  }else{
    const t=f.type==='date'?'date':f.type==='month'?'month':f.type==='datetime'?'datetime-local':f.type==='password'?'password':'text';
    input=`<input id="${id}" type="${t}" data-field="${esc(f.key)}" value="${esc(t==='datetime-local'?toInput(v):v)}" ${f.list?`list="${f.list}"`:''} ${f.placeholder?`placeholder="${esc(f.placeholder)}"`:''} ${t==='password'?'autocomplete="new-password"':''}>`;
  }
  const wide=(f.type==='textarea'||f.type==='mono'||f.full)?' full':'';
  return `<div class="field${wide}"><label for="${id}">${esc(f.label||label(f.key))}${req}</label>${input}${f.hint?`<span class="hint">${esc(f.hint)}</span>`:''}</div>`;
}
function readFields(root){const o={};root.querySelectorAll('[data-field]').forEach(el=>{
  o[el.dataset.field]=el.dataset.bool?el.checked:(el.type==='datetime-local'?fromInput(el.value):el.value);});return o;}
function datalists(){
  const cust=uniq([...Object.values(S.tables).flatMap(t=>t.rows.map(r=>r.Customer)),...S.plans.tasks.map(t=>t.customer)]);
  const dev=uniq((S.tables.Asset_Master?.rows||[]).map(r=>r.Device));
  const env=uniq([...(S.tables.Asset_Master?.rows||[]).map(r=>r.Environment),'Production','DR','UAT','Test','Development']);
  return `<datalist id="dl-customers">${cust.map(c=>`<option value="${esc(c)}">`).join('')}</datalist>
    <datalist id="dl-assets">${dev.map(c=>`<option value="${esc(c)}">`).join('')}</datalist>
    <datalist id="dl-env">${env.map(c=>`<option value="${esc(c)}">`).join('')}</datalist>
    <datalist id="dl-phases">${S.plans.presets.map(c=>`<option value="${esc(c)}">`).join('')}</datalist>`;
}

/* ---------- chrome ---------- */
function renderStore(){
  const st=S.status;if(!st)return;
  const chip=$('#storeChip');chip.querySelector('.loc').innerHTML=`<b>${esc(st.storage.label)}</b> ${esc(st.storage.location)}`;
  chip.title=st.storage.label+': '+st.storage.location+'\n'+(st.last_error?'Sync problem: '+st.last_error+'\n':'')+'Last synced '+fmtTime(st.last_sync);
  $('#syncDot').classList.toggle('err',!!st.last_error);
}
function renderChrome(){
  renderStore();
  const counts={plans:S.plans.tasks.filter(t=>!['Completed','Cancelled'].includes(t.status)).length,runbook:S.tables.Runbook?.rows.length,
    Asset_Master:S.tables.Asset_Master?.rows.length,Incidents:(S.tables.Incidents?.rows||[]).filter(isOpen).length,
    Bulletins:activeBulletins().filter(b=>!S.reads.has(b.Record_ID)).length,terminals:S.terms.length||null,mywork:myWorkCount()||null};
  const hot={Incidents:(S.tables.Incidents?.rows||[]).some(r=>isOpen(r)&&['P1','P2'].includes(r.Priority)),Bulletins:counts.Bulletins>0};
  $('#nav').innerHTML=NAV.map(([g,items])=>{const vis=items.filter(i=>(i[3]||1)<=S.rank);if(!vis.length)return '';
    return `<div class="grp">${esc(g)}</div>`+vis.map(([k,l,ic])=>`<button class="blade ${S.view===k?'on':''}" data-act="nav" data-view="${k}" title="${esc(l)}" ${S.view===k?'aria-current="page"':''}><span class="ic">${svg(ic)}</span><span class="lbl">${esc(l)}</span>${counts[k]?`<span class="n ${hot[k]?'hot':''}">${counts[k]}</span>`:''}</button>`).join('');}).join('');
  const me=S.me||{name:'',role:''},ini=(me.name||me.username||'?').split(/\s+/).map(x=>x[0]).join('').slice(0,2).toUpperCase();
  $('#sideFoot').innerHTML=`<div class="me"><button class="avatar" data-act="account" title="Your account">${esc(ini)}</button><div class="who"><b>${esc(me.name)}</b><span class="rolepill">${esc(me.role)}</span></div>
    <div class="acts"><a class="iconbtn" href="/app/" title="Try the new interface (preview)" aria-label="Try the new interface (preview)">${svg('spark')}</a><button class="iconbtn" data-act="theme" title="Switch light / dark" aria-label="Switch light or dark theme">${svg('sun')}</button><button class="iconbtn" data-act="logout" title="Sign out" aria-label="Sign out">${svg('out')}</button></div></div>`;
  const cur=VIEWS.find(v=>v[0]===S.view),grp=NAV.find(g=>g[1].some(i=>i[0]===S.view));
  $('#crumbs').innerHTML=cur?`${esc(grp[0])} / <b>${esc(cur[1])}</b>`:'';
  renderTicker();
}
function activeBulletins(){
  const now=nowLocal();
  return (S.tables.Bulletins?.rows||[]).filter(b=>(!b.Start||b.Start<=now)&&(!b.Expiry||b.Expiry>=now))
    .sort((a,b)=>(truthy(b.Pinned)-truthy(a.Pinned))||((SEV_RANK[b.Severity]||0)-(SEV_RANK[a.Severity]||0))||String(b.Start).localeCompare(String(a.Start)));
}
function renderTicker(){
  const list=activeBulletins(),el=$('#ticker');
  if(!list.length){el.innerHTML='';return;}
  const unread=list.filter(b=>!S.reads.has(b.Record_ID)).length;
  const items=dup=>list.map(b=>`<button class="tk" data-act="bulletin" data-id="${esc(b.Record_ID)}" ${dup?'tabindex="-1" aria-hidden="true"':''}>${S.reads.has(b.Record_ID)?'':'<i class="unread"></i>'}<span class="sev sev-${esc(b.Severity)}">${esc(b.Severity)}</span><span class="tt">${esc(b.Title)}</span>${b.Customers?`<span class="st">${esc(b.Customers)}</span>`:''}</button>`).join('');
  el.innerHTML=`<div class="ticker"><div class="lbl">Bulletins${unread?` <span class="chip">${unread} unread</span>`:''} <button data-act="nav" data-view="Bulletins">See all</button></div>
    <div class="rail"><div class="track" style="--dur:${Math.max(28,list.length*14)}s">${items(false)}${items(true)}</div></div></div>`;
}
function go(view){
  let param='';view=String(view||'');
  if(view.includes('=')){param=decodeURIComponent(view.slice(view.indexOf('=')+1));view=view.slice(0,view.indexOf('='));}
  if(view==='shiftlog'&&param){const [d,s]=param.split('/');
    if(/^[0-9]{4}-[0-9]{2}-[0-9]{2}$/.test(d||'')){S.hDate=d;if(SHIFTS.includes(s))S.hShift=s;S.slFocus={d,s:SHIFTS.includes(s)?s:S.hShift};}}
  const item=VIEWS.find(v=>v[0]===view);
  if(!item||(item[3]||1)>S.rank)view='mywork';
  S.view=view;history.replaceState(null,'','#'+view);renderChrome();render();window.scrollTo(0,0);
  $('#app').classList.remove('mobile-open');
  if(view==='shiftlog'&&S.slFocus){const {d,s}=S.slFocus;S.slFocus=null;const [y,m]=d.split('-');
    if(S.hier.years?.[y]?.months?.[m]?.dates?.[d]?.shifts?.[s]){['y:'+y,`m:${y}-${m}`,'d:'+d,`s:${d}:${s}`].forEach(k=>S.open.add(k));persistOpen();render();
      requestAnimationFrame(()=>document.querySelector(`[data-key="s:${d}:${s}"]`)?.scrollIntoView({block:'start',behavior:'smooth'}));}
    else toast(`There is no ${s} shift on ${d} yet. The date and shift are filled in above: select Open shift to start it.`);}
}
function render(){
  const v=S.view,el=$('#view'),term=v==='terminals';
  el.classList.toggle('hidden',term);$('#termHost').classList.toggle('hidden',!term);
  if(term){renderTerminals();return;}
  if(v==='mywork'){el.innerHTML=renderMyWork();fillMyShifts();}
  else if(v==='roster')renderRoster(el);
  else if(v==='teams')renderTeams(el);
  else if(v==='overview'){el.innerHTML=renderOverview();fillOverviewRoster();}
  else if(v==='plans'){el.innerHTML=plansShell();renderPlans();}
  else if(v==='runbook'){el.innerHTML=runbookShell();renderRunbook();}
  else if(S.schema[v]&&v==='Shift_Transitions'&&S.teamScope&&!S.team)el.innerHTML=teamBar();
  else if(S.schema[v]){el.innerHTML=(v==='Shift_Transitions'?teamBar():'')+tableShell(v);renderRows(v);}
  else if(v==='shiftlog')el.innerHTML=S.teamScope&&!S.team?teamBar():teamBar()+renderShiftLog();
  else if(v==='customers')renderCustomers(el);
  else if(v==='activity')renderActivity(el);
  else if(v==='admin')renderAdmin(el);
  else return go('mywork');
}
function renderHost(scope){if(scope==='plans')renderPlans();else if(scope==='runbook')renderRunbook();else if(S.schema[scope])renderRows(scope);}

/* ---------- team scope bar (shift log and handovers belong to a team) ---------- */
function teamBar(){
  if(!S.teamScope)return '';
  if(!S.team)return `<div class="panel" style="margin-bottom:16px"><h3>Team only</h3><p class="st" style="margin:0">${S.me.role==='Admin'
    ?'Shift logs and handovers belong to teams. Administrators are not team members, so they do not see them (separation of duties). Set up teams and leads on the Teams page.'
    :'Shift logs and handovers belong to teams, and you are not in one yet. A team lead can add you to their team.'}</p></div>`;
  const name=(S.myTeams.find(t=>t.id===S.team)||{}).name||'';
  const pick=S.myTeams.length>1?`<select data-team-pick aria-label="Team">${S.myTeams.map(t=>`<option value="${esc(t.id)}" ${t.id===S.team?'selected':''}>${esc(t.name)}</option>`).join('')}</select>`
    :`<b>${esc(name)}</b>`;
  const lg=S.legacy&&(S.legacy.shift_days||S.legacy.handovers)&&S.me.rank>=3?`<div class="panel" style="margin:10px 0 0;border-style:dashed"><p class="st" style="margin:0 0 8px">There is a shift log from before teams (${S.legacy.shift_days} day(s)) and ${S.legacy.handovers} handover(s) that no team owns yet, so nobody sees them. As a lead of ${esc(name)} you can move them into this team.</p>
    <button class="btn" data-act="team-claim">Move them into ${esc(name)}</button></div>`:'';
  return `<div class="toolbar" style="margin:0 0 12px;align-items:center"><span class="st">Team:</span>${pick}<span class="st">Only ${esc(name)} sees this team's shift log and handovers.</span></div>${lg}`;
}
document.addEventListener('change',e=>{if(!e.target.matches('[data-team-pick]'))return;S.team=e.target.value;localStorage.setItem('msp.team',S.team);load();});

/* ---------- generic tables ---------- */
function findRow(sheet,id){return (S.tables[sheet]?.rows||[]).find(r=>r.Record_ID===id);}
function fieldType(sheet,key){
  const sc=S.schema[sheet];if(!sc)return 'text';
  if(sc.choices[key])return 'select';if(sc.dates.includes(key))return 'date';if(DT_KEYS.has(key))return 'datetime';
  if(MONO.has(key))return 'mono';if(LONG.has(key))return 'textarea';return 'text';
}
function tf(sheet){return S.filters[sheet]||(S.filters[sheet]={customer:'',q:'',pri:'',open:sheet==='Incidents'});}
function moduleButtons(module,extra=''){return `${extra}<button class="btn" data-act="import" data-module="${module}">Import</button><button class="btn" data-act="template" data-module="${module}">Download template</button>`;}
function tableShell(sheet){
  const sc=S.schema[sheet],f=tf(sheet),rows=S.tables[sheet].rows;
  const customers=uniq(rows.map(r=>r.Customer));
  return `<h2>${esc(sc.title)}</h2><p class="sub">${esc(SUBS[sheet]||'')}</p>
  <div class="toolbar">
    ${customers.length?`<select data-filter="customer" data-scope="${sheet}" aria-label="Customer"><option value="">All customers</option>${customers.map(c=>`<option ${c===f.customer?'selected':''}>${esc(c)}</option>`).join('')}</select>`:''}
    ${sheet==='Incidents'?`<select data-filter="pri" data-scope="${sheet}" aria-label="Priority"><option value="">All priorities</option>${['P1','P2','P3','P4'].map(p=>`<option ${p===f.pri?'selected':''}>${p}</option>`).join('')}</select>
      <label class="st"><input type="checkbox" data-filter="open" data-scope="${sheet}" ${f.open?'checked':''}> Open only</label>`:''}
    <div class="grow"><input type="search" data-filter="q" data-scope="${sheet}" placeholder="Search ${esc(sheet==='Asset_Master'?'configuration items':sc.title.toLowerCase())}" value="${esc(f.q)}" aria-label="Search"></div>
    <button class="btn primary" data-act="add-row" data-sheet="${sheet}">Add ${esc(sc.noun)}</button>
    <button class="btn" data-act="columns" data-sheet="${sheet}">Columns</button>${moduleButtons(sc.module)}
  </div><div id="tableHost"></div><p class="st" style="margin-top:8px">Double-click any cell to edit it in place. Enter saves, Esc cancels.</p>`;
}
function filtered(sheet){
  const f=tf(sheet),q=(f.q||'').toLowerCase();
  return S.tables[sheet].rows.filter(r=>(!f.customer||r.Customer===f.customer)&&(!f.pri||r.Priority===f.pri)&&(!f.open||isOpen(r))
    &&(!q||Object.values(r).some(v=>String(v).toLowerCase().includes(q))));
}
function runbookIndex(){
  const m={};(S.tables.Runbook?.rows||[]).forEach(r=>{const k=nk(r.Asset);if(!k)return;(m[k]=m[k]||{n:0,ki:0});m[k].n++;if(truthy(r.Known_Issue))m[k].ki++;});return m;
}
function cellView(sheet,k,r,rbi){
  const v=r[k]??'';
  if(k==='Priority')return v?`<span class="badge pri-${esc(v)}">${esc(v)}</span>`:'';
  if(k==='Severity')return v?`<span class="sev sev-${esc(v)}">${esc(v)}</span>`:'';
  if(k==='Special_Care')return v?`<span class="care-tag clip">${esc(v)}</span>`:'';
  if(['Acknowledged','Pinned','Known_Issue'].includes(k))return `<span class="st ${truthy(v)?'done':''}">${truthy(v)?'Yes':'No'}</span>`;
  if(sheet==='Asset_Master'&&k==='Device'){const i=rbi[nk(v)];return `<a href="#runbook" class="devlink" data-act="goto-runbook" data-id="${esc(r.Record_ID)}" title="Open this device's runbook">${esc(v)}</a>${i?`<span class="cnt">${i.n} entries</span>`:''}${i&&i.ki?`<span class="ki">known issue</span>`:''}`;}
  if((k==='Warranty_End'||k==='OS_EOL')&&v){const days=(new Date(v)-new Date(today()))/864e5;return `<span class="${days<0?'bad':days<90?'warn':''}">${esc(v)}</span>`;}
  if(sheet==='Bulletins'&&k==='Expiry'&&v&&v<nowLocal())return `<span class="st">${esc(v)} (expired)</span>`;
  if(MONO.has(k))return `<div class="clip mono">${esc(v)}</div>`;
  return `<div class="clip">${esc(v)}</div>`;
}
function rowActions(sheet,id){
  const remote=sheet==='Asset_Master'?`<button class="btn xs primary" data-act="ssh" data-id="${esc(id)}" title="Open an SSH terminal">SSH</button>
    <button class="btn xs" data-act="rdp" data-id="${esc(id)}" title="Open Remote Desktop">RDP</button>`:'';
  return remote+`<button class="btn xs" data-act="edit-row" data-sheet="${sheet}" data-id="${esc(id)}">${S.rank<2?'View':'Edit'}</button>
    <button class="btn xs ghost" data-act="clone-row" data-sheet="${sheet}" data-id="${esc(id)}">Clone</button>
    <button class="btn xs ghost" data-act="history" data-key="table/${sheet}/${esc(id)}">History</button>
    <button class="btn xs ghost" data-act="del-row" data-sheet="${sheet}" data-id="${esc(id)}">Delete</button>`;
}
function renderRows(sheet){
  const host=$('#tableHost');if(!host)return;
  const t=S.tables[sheet],cols=t.columns.filter(c=>!c.hidden&&!c.system),rows=filtered(sheet),rbi=runbookIndex();
  if(!rows.length){host.innerHTML=`<div class="tablewrap"><div class="empty">${t.rows.length?'No records match these filters.':'No records yet. Add the first one, or import a file.'}</div></div>`;return;}
  const art=sheet==='Incidents';
  host.innerHTML=`<div class="tablewrap"><table class="grid"><thead><tr>${cols.map(c=>`<th>${esc(c.label)}</th>`).join('')}${art?'<th>ART</th>':''}<th></th></tr></thead><tbody>
  ${rows.map(r=>`<tr class="${String(r.Special_Care||'').trim()?'care':''}">${cols.map(c=>`<td data-ie="t" data-sheet="${sheet}" data-id="${esc(r.Record_ID)}" data-key="${esc(c.key)}">${cellView(sheet,c.key,r,rbi)}</td>`).join('')}
    ${art?`<td><span class="art">${['Action','Reason','Technical_Tracking'].map((k,i)=>`<span class="${String(r[k]||'').trim()?'':'miss'}" title="${label(k)}">${'ART'[i]}</span>`).join('')}</span></td>`:''}
    <td class="actions">${rowActions(sheet,r.Record_ID)}</td></tr>`).join('')}</tbody></table></div>`;
}
function editRow(sheet,id,prefill){
  const t=S.tables[sheet],sc=S.schema[sheet],row=id?findRow(sheet,id):null;
  if(id&&!row)return toast('That record no longer exists. The list has been refreshed.','error');
  const defaults={...(prefill||{})};
  if(!row){
    if(sheet==='Shift_Transitions'){defaults.Date=defaults.Date||today();defaults.Shift=defaults.Shift||currentShift();}
    if(sheet==='Runbook'){defaults.Date=defaults.Date||today();defaults.Engineer=defaults.Engineer||S.user;}
    if(sheet==='Incidents'){defaults.Priority='P3';defaults.Status='New';defaults.Engineer=S.user;}
    if(sheet==='Bulletins'){defaults.Severity='Info';defaults.Start=nowLocal();}
  }
  const body=t.columns.filter(c=>!c.system).map(c=>{
    const grp=sheet==='Incidents'&&c.key==='Action'?'<div class="group">ART record</div>':'';
    return grp+fieldHtml({key:c.key,label:c.label,type:fieldType(sheet,c.key),options:sc.choices[c.key],value:row?row[c.key]:(defaults[c.key]||''),
      required:sc.required.includes(c.key),hint:HINTS[c.key],list:c.key==='Customer'?'dl-customers':(c.key==='Asset'?'dl-assets':(c.key==='Environment'?'dl-env':''))});
  }).join('')+datalists()+(row?`<div class="metaline">Version ${esc(row.Version||'1')}. Created by ${esc(row.Created_By||'?')} ${esc(fmtTime(row.Created_At))}. Last updated by ${esc(row.Updated_By||'?')} ${esc(fmtTime(row.Updated_At))}.</div>`:'');
  const ro=S.rank<2||(sheet==='Bulletins'&&S.rank<3);
  openModal({title:ro?label(sc.noun.replace(/^./,c=>c.toUpperCase()))+' details':(row?'Edit ':'Add ')+sc.noun,body,size:'wide',save:row?'Save changes':'Add '+sc.noun,
    onMount:ov=>{if(ro)ov.querySelectorAll('[data-field]').forEach(f=>f.disabled=true);},onSave:ro?null:async ov=>{
    const vals=readFields(ov);
    const miss=sc.required.filter(k=>!String(vals[k]||'').trim());
    if(miss.length){toast('Required: '+miss.map(k=>t.columns.find(c=>c.key===k)?.label||label(k)).join(', '),'error');return false;}
    if(row)await run(()=>dupSafe(a=>api('PUT',`/api/table/${sheet}/${encodeURIComponent(id)}`,{row:vals,allow_duplicate:a})),'Saved');
    else await run(()=>dupSafe(a=>api('POST',`/api/table/${sheet}`,{row:vals,allow_duplicate:a})),'Added '+sc.noun);
  }});
}
async function cloneRow(sheet,id){
  const r=await run(()=>api('POST',`/api/table/${sheet}/${encodeURIComponent(id)}/clone`,{}),'Cloned. Adjust the copy and save.');
  editRow(sheet,r.row.Record_ID);
}
async function showHistory(key,title){
  let d;try{d=await api('GET','/api/history?key='+encodeURIComponent(key));}catch(e){return toast(e.message,'error');}
  const vs=d.versions;
  const body=`<div class="full">${vs.length?vs.map((v,i)=>`<div class="hist"><div><b>Version ${v.v}</b> <span class="st">${esc(fmtTime(v.at))}, ${esc(v.by)}: ${esc(v.action)}${v.batch?' (import)':''}${v.deleted?' (record deleted)':''}</span>
    ${i>0||v.deleted?`<button class="btn xs" data-restore="${v.v}">Restore this version</button>`:'<span class="chip">current</span>'}</div>
    ${v.changes.length?`<table>${v.changes.map(c=>`<tr><td>${esc(label(c.field))}</td><td><del>${esc(c.old)}</del></td><td><ins>${esc(c.new)}</ins></td></tr>`).join('')}</table>`:''}</div>`).join('')
    :'<p class="st">No history recorded yet for this record.</p>'}</div>`;
  openModal({title:'History: '+(title||key),body,size:'wide',onMount:ov=>ov.addEventListener('click',async e=>{
    const b=e.target.closest('[data-restore]');if(!b)return;
    if(!await confirmBox(`Restore version ${b.dataset.restore}? The current state stays in the history.`,'Restore',false))return;
    await run(()=>api('POST','/api/history/restore',{key,v:+b.dataset.restore}),'Version restored');closeModal();
  })});
}
function columnEditor({title,cols,onSave,note}){
  const list=cols.map(c=>({...c}));
  const draw=ov=>{ov.querySelector('#colList').innerHTML=list.map((c,i)=>`<div class="colrow">
    <span class="tag">${c.core?'built-in':'custom'}</span><input type="text" data-idx="${i}" value="${esc(c.label)}" aria-label="Column ${i+1} name">
    <label><input type="checkbox" data-show="${i}" ${c.hidden?'':'checked'}> show</label>
    <button class="btn xs ghost" data-cmove="${i}" data-dir="-1" ${i===0?'disabled':''} aria-label="Move up">Up</button>
    <button class="btn xs ghost" data-cmove="${i}" data-dir="1" ${i===list.length-1?'disabled':''} aria-label="Move down">Down</button>
    <button class="btn xs danger" data-cdel="${i}" ${c.core?'disabled title="Built-in columns can be renamed or hidden, not removed"':''}>Remove</button></div>`).join('');};
  openModal({title,save:'Save columns',size:'wide',
    body:`<div class="full"><p class="st" style="margin-bottom:10px">${esc(note||'Rename, reorder, show or hide columns. Renaming keeps all data. Removing a custom column deletes its values.')}</p><div id="colList"></div><button class="btn sm" data-cadd>Add column</button></div>`,
    onMount:ov=>{draw(ov);
      ov.addEventListener('input',e=>{if(e.target.dataset.idx!==undefined)list[+e.target.dataset.idx].label=e.target.value;});
      ov.addEventListener('change',e=>{if(e.target.dataset.show!==undefined)list[+e.target.dataset.show].hidden=!e.target.checked;});
      ov.addEventListener('click',e=>{
        const mv=e.target.closest('[data-cmove]'),del=e.target.closest('[data-cdel]'),add=e.target.closest('[data-cadd]');
        if(mv){const i=+mv.dataset.cmove,j=i+(+mv.dataset.dir);[list[i],list[j]]=[list[j],list[i]];draw(ov);}
        if(del){list.splice(+del.dataset.cdel,1);draw(ov);}
        if(add){list.push({label:'New column '+(list.length+1),hidden:false,core:false});draw(ov);ov.querySelector(`[data-idx="${list.length-1}"]`).select();}
      });},
    onSave:async()=>{if(list.some(c=>!c.label.trim())){toast('Column names cannot be empty','error');return false;}
      await onSave(list.map(c=>({key:c.key,label:c.label.trim(),hidden:!!c.hidden})));}});
}
function tableColumns(sheet){
  columnEditor({title:'Columns: '+S.schema[sheet].title,cols:S.tables[sheet].columns.filter(c=>!c.system),
    onSave:cols=>run(()=>api('PUT',`/api/table/${sheet}/columns`,{columns:cols}),'Columns saved')});
}

/* ---------- inline editing ---------- */
function stepOf(tid,pid,sid){const t=S.plans.tasks.find(x=>x.id===tid);const ph=t?.phases.find(p=>p.id===pid);return {t,ph,st:ph?.steps.find(s=>s.id===sid)};}
const STEP_TYPES={planned_start:'datetime',actual_start:'datetime',actual_end:'datetime',status:'select',notes:'textarea',command:'mono',rollback:'mono',expected:'textarea'};
function inlineMeta(ds){
  if(ds.ie==='t'){const r=findRow(ds.sheet,ds.id);if(!r)return null;return {value:r[ds.key]??'',type:fieldType(ds.sheet,ds.key),options:S.schema[ds.sheet].choices[ds.key]};}
  if(ds.ie==='s'){const {st}=stepOf(ds.task,ds.phase,ds.step);if(!st)return null;
    const v=ds.key.startsWith('x_')?(st.custom[ds.key]||''):(st[ds.key]??'');
    return {value:v,type:STEP_TYPES[ds.key]||'text',options:ds.key==='status'?S.plans.step_statuses:null};}
  return null;
}
function saveInline(ds,v){
  if(ds.ie==='t')return dupSafe(a=>api('PUT',`/api/table/${ds.sheet}/${encodeURIComponent(ds.id)}`,{row:{[ds.key]:v},allow_duplicate:a}));
  return planOp('step.update',{task_id:ds.task,phase_id:ds.phase,step_id:ds.step,step:{[ds.key]:v}});
}
function startInline(td){
  const ds={...td.dataset},m=inlineMeta(ds);if(!m)return;
  let ed;
  if(m.type==='select'){ed=document.createElement('select');[...new Set(['',...(m.options||[]),m.value])].forEach(o=>{const op=document.createElement('option');op.value=o;op.textContent=o||'—';if(o===m.value)op.selected=true;ed.appendChild(op);});}
  else if(m.type==='textarea'||m.type==='mono'){ed=document.createElement('textarea');if(m.type==='mono')ed.className='mono';}
  else{ed=document.createElement('input');ed.type=m.type==='date'?'date':m.type==='datetime'?'datetime-local':'text';}
  ed.classList.add('ie');if(ed.tagName!=='SELECT')ed.value=m.type==='datetime'?toInput(m.value):String(m.value);
  td.innerHTML='';td.appendChild(ed);ed.focus();if(ed.type==='text')ed.select();
  let done=false;
  const finish=async save=>{if(done)return;done=true;let v=ed.value;if(m.type==='datetime')v=fromInput(v);
    if(!save||String(v)===String(m.value??'')){render();return;}
    try{await run(()=>saveInline(ds,v),'Saved');}catch(e){render();}};
  ed.addEventListener('keydown',e=>{if(e.key==='Escape'){e.preventDefault();e.stopPropagation();finish(false);}
    else if(e.key==='Enter'&&(ed.tagName!=='TEXTAREA'||e.ctrlKey||e.metaKey)){e.preventDefault();finish(true);}});
  ed.addEventListener('blur',()=>finish(true));
  if(ed.tagName==='SELECT')ed.addEventListener('change',()=>finish(true));
}
document.addEventListener('dblclick',e=>{if(S.rank<2||e.target.closest('a,button'))return;const td=e.target.closest('td[data-ie]');if(!td||td.querySelector('.ie'))return;e.preventDefault();startInline(td);});

/* ---------- overview ---------- */
function shiftNode(d,s){const [y,m]=d.split('-');return S.hier.years?.[y]?.months?.[m]?.dates?.[d]?.shifts?.[s];}
function sectionCount(node,re){if(!node)return 0;return node.sections.filter(x=>re.test(x.name)).reduce((a,x)=>a+x.rows.length,0);}
function renderOverview(){
  const td=today(),long=new Date().toLocaleDateString([], {weekday:'long',day:'numeric',month:'long',year:'numeric'});
  const inc=(S.tables.Incidents?.rows||[]).filter(isOpen),hot=inc.filter(r=>['P1','P2'].includes(r.Priority)).sort((a,b)=>a.Priority.localeCompare(b.Priority));
  const soon=S.plans.tasks.filter(t=>t.date>=addDays(td,-1)&&t.date<=addDays(td,7)&&!['Cancelled'].includes(t.status)).sort((a,b)=>(a.computed.start||'').localeCompare(b.computed.start||''));
  const ki=(S.tables.Runbook?.rows||[]).filter(r=>truthy(r.Known_Issue));
  const attention=(S.tables.Asset_Master?.rows||[]).map(r=>{const w=[];['Warranty_End','OS_EOL'].forEach(k=>{if(r[k]){const d=(new Date(r[k])-new Date(td))/864e5;if(d<90)w.push(`${label(k)} ${d<0?'passed':'in '+Math.ceil(d)+' days'} (${r[k]})`);}});return w.length?{r,w}:null;}).filter(Boolean);
  const tiles=SHIFTS.map(s=>{const n=shiftNode(td,s);
    const body=n?`<div class="nums"><div><b>${sectionCount(n,/in work/i)}</b><span>in work</span></div><div><b>${sectionCount(n,/hold/i)}</b><span>on hold</span></div><div><b>${sectionCount(n,/scheduled/i)}</b><span>scheduled</span></div></div>`:`<div class="start">Not started yet. Select to open this shift.</div>`;
    return `<button class="tile" style="--c:${SHIFT_COLOR[s]}" data-act="open-shift" data-d="${td}" data-s="${s}"><h3>${s}<small>${s===currentShift()?'current shift':''}</small></h3>${body}</button>`;}).join('');
  const li=(t,s,lead='',act='')=>`<li>${lead}<div class="t">${act?`<button class="linkish" ${act}>${esc(t)}</button>`:esc(t)}<small>${esc(s)}</small></div></li>`;
  return `<div class="hello"><h2>${esc(long)}</h2><span class="st">${soon.length} execution plan(s) this week, ${inc.length} open incidents, ${hot.length} at P1/P2</span></div>
  <div class="sun">${tiles}</div>
  <div class="cols">
    <section class="panel"><h3>Execution plans: next 7 days</h3><ul class="list">${soon.length?soon.map(t=>li(`${t.customer}: ${t.name}`,
      `${t.computed.start} to ${t.computed.projected_end}${t.ticket_id?', '+t.ticket_id:''}, ${t.computed.done}/${t.computed.steps} steps${t.computed.overrun_min?', overruns window by '+fmtMin(t.computed.overrun_min):''}`,
      `<span class="stb ${t.status.replace(/\s/g,'')}">${esc(t.status)}</span>`,`data-act="open-task" data-id="${esc(t.id)}"`)).join(''):'<li class="st">Nothing scheduled.</li>'}</ul></section>
    <section class="panel"><h3>P1 / P2 in flight</h3><ul class="list">${hot.length?hot.map(r=>li(`${r.Ticket_ID||'(no ticket)'}  ${r.Summary}`,`${r.Customer} / ${r.Asset} / ${r.Status}`,`<span class="badge pri-${esc(r.Priority)}">${esc(r.Priority)}</span>`)).join(''):'<li class="st">No open P1 or P2 incidents.</li>'}</ul></section>
    <section class="panel"><h3>Known issues</h3><ul class="list">${ki.length?ki.slice(0,8).map(r=>li(`${r.Asset}: ${r.Title||r.Note}`,`${r.Customer}, ${r.Date}`,'',`data-act="goto-asset-rb" data-asset="${esc(r.Asset)}"`)).join(''):'<li class="st">No known issues recorded.</li>'}</ul></section>
    <section class="panel"><h3>CIs needing attention</h3><ul class="list">${attention.length?attention.slice(0,8).map(x=>li(`${x.r.Device} (${x.r.Customer})`,x.w.join('; '))).join(''):'<li class="st">No warranty or OS end-of-life dates within 90 days.</li>'}</ul></section>
    <section class="panel"><h3>Recent activity</h3><ul class="list">${S.audit.slice(0,7).map(a=>li(`${a.Action}: ${a.Target}`,`${a.User}, ${fmtTime(a.Timestamp)}`)).join('')||'<li class="st">Nothing yet.</li>'}</ul></section>
  </div>`;
}

/* ---------- runbook: Customer > Environment > Asset > entries ---------- */
function rbf(){return S.filters.runbook||(S.filters.runbook={customer:'',env:'',cat:'',q:'',ki:false});}
function runbookShell(){
  const f=rbf(),rows=S.tables.Runbook.rows;
  const customers=uniq(rows.map(r=>r.Customer)),envs=uniq(rows.map(r=>r.Environment)),cats=S.schema.Runbook.choices.Category;
  const sel=(k,list,all)=>`<select data-filter="${k}" data-scope="runbook" aria-label="${all}"><option value="">${all}</option>${list.map(c=>`<option ${c===f[k]?'selected':''}>${esc(c)}</option>`).join('')}</select>`;
  return `<h2>Runbook</h2><p class="sub">Every activity recorded against an asset, grouped by customer, environment and asset. Search finds past fixes across all notes, solutions and commands.</p>
  <div class="toolbar">${sel('customer',customers,'All customers')}${sel('env',envs,'All environments')}${sel('cat',cats,'All categories')}
    <label class="st"><input type="checkbox" data-filter="ki" data-scope="runbook" ${f.ki?'checked':''}> Known issues only</label>
    <div class="grow"><input type="search" data-filter="q" data-scope="runbook" placeholder="Search notes, solutions, commands, tickets" value="${esc(f.q)}" aria-label="Search runbook"></div>
    <button class="btn primary" data-act="add-row" data-sheet="Runbook">Add entry</button><button class="btn" data-act="columns" data-sheet="Runbook">Columns</button>
    ${moduleButtons('runbook')}<button class="btn ghost" data-act="expand" data-scope="rb">Expand all</button><button class="btn ghost" data-act="collapse" data-scope="rb">Collapse all</button>
  </div><div id="rbHost"></div>`;
}
function det(key,cls,summary,inner,extra=''){return `<details class="node ${cls}" data-key="${esc(key)}" ${S.open.has(key)?'open':''} ${extra}><summary>${summary}</summary>${inner}</details>`;}
function renderRunbook(){
  const host=$('#rbHost');if(!host)return;
  const f=rbf(),q=f.q.toLowerCase(),t=S.tables.Runbook;
  const rows=t.rows.filter(r=>(!f.customer||r.Customer===f.customer)&&(!f.env||r.Environment===f.env)&&(!f.cat||r.Category===f.cat)&&(!f.ki||truthy(r.Known_Issue))
    &&(!q||Object.values(r).some(v=>String(v).toLowerCase().includes(q))));
  const tree={};const put=(c,e,a)=>(((tree[c]=tree[c]||{})[e]=tree[c][e]||{})[a]=tree[c][e][a]||[]);
  rows.forEach(r=>put(r.Customer||'(no customer)',r.Environment||'(no environment)',r.Asset).push(r));
  if(S.rbFocus)put(S.rbFocus.c,S.rbFocus.e,S.rbFocus.a);
  const assets=S.tables.Asset_Master.rows;
  const visible=t.columns.filter(c=>!c.hidden&&!c.system&&!['Customer','Environment','Asset','Date','Category','Title','Ticket_Ref','Engineer','Known_Issue'].includes(c.key));
  const custs=Object.keys(tree).sort();
  if(!custs.length){host.innerHTML=`<div class="tablewrap"><div class="empty">${t.rows.length?'No entries match these filters.':'No runbook entries yet. Add one, or open a CI from the CMDB.'}</div></div>`;return;}
  host.innerHTML=custs.map(c=>{const envs=Object.keys(tree[c]).sort();
    const cn=envs.reduce((a,e)=>a+Object.values(tree[c][e]).reduce((x,l)=>x+l.length,0),0);
    return det('rb:c:'+c,'lvl0',`<span class="label">${esc(c)}</span><span class="meta">${cn} entries</span>`,
      envs.map(e=>{const as=Object.keys(tree[c][e]).sort();
        return det(`rb:e:${c}|${e}`,'',`<span class="label">${esc(e)}</span><span class="meta">${as.length} asset(s)</span>`,
          as.map(a=>{const list=tree[c][e][a].sort((x,y)=>String(y.Date).localeCompare(String(x.Date)));
            const inv=assets.find(x=>nk(x.Device)===nk(a));const kis=list.filter(r=>truthy(r.Known_Issue)).length;const key=`rb:a:${c}|${e}|${a}`;
            const sum=`<span class="label">${esc(a)}</span><span class="meta">${list.length} entries</span>${kis?`<span class="ki">${kis} known issue(s)</span>`:''}<span class="sp"></span>
              <button class="btn xs primary" data-act="rb-add" data-c="${esc(c==='(no customer)'?'':c)}" data-e="${esc(e==='(no environment)'?'':e)}" data-a="${esc(a)}">Add entry</button>
              ${inv?`<button class="btn xs" data-act="ssh" data-id="${esc(inv.Record_ID)}">SSH</button><button class="btn xs" data-act="rdp" data-id="${esc(inv.Record_ID)}">RDP</button><button class="btn xs" data-act="edit-row" data-sheet="Asset_Master" data-id="${esc(inv.Record_ID)}">Open asset</button>`:''}`;
            const info=inv?`<div class="rb-info">${esc([inv.OS,inv.OS_Version].filter(Boolean).join(' '))}${inv.Application?', '+esc(inv.Application):''}${inv.Platform_Type?', '+esc(inv.Platform_Type):''}${inv.Virtualization_Cloud?' ('+esc(inv.Virtualization_Cloud)+')':''}${inv.IP_Address?', '+esc(inv.IP_Address):''}${inv.Special_Care?`<div class="care-tag">${esc(inv.Special_Care)}</div>`:''}</div>`:'';
            return det(key,'rb-asset',sum,info+(list.length?list.map(r=>entryCard(r,visible)).join(''):'<p class="st" style="margin-left:22px">No entries for this asset yet.</p>'),`data-rbkey="${esc(nk(a))}"`);
          }).join(''));
      }).join(''));
  }).join('');
  if(S.rbFocus){const el=host.querySelector(`[data-rbkey="${nk(S.rbFocus.a)}"]`);if(el){el.querySelector('summary').classList.add('focus');requestAnimationFrame(()=>el.scrollIntoView({block:'start',behavior:'smooth'}));}S.rbFocus=null;}
}
function entryCard(r,cols){
  const ki=truthy(r.Known_Issue);
  return `<article class="entry ${ki?'ki-entry':''}"><header><span class="st">${esc(r.Date)}</span>${r.Category?`<span class="stb">${esc(r.Category)}</span>`:''}${ki?'<span class="ki">known issue</span>':''}
    <span class="t">${esc(r.Title||'')}</span>${r.Ticket_Ref?`<span class="badge pri-P4">${esc(r.Ticket_Ref)}</span>`:''}<span class="st">${esc(r.Engineer||'')}</span><span class="sp"></span>${rowActions('Runbook',r.Record_ID)}</header>
    ${cols.filter(c=>String(r[c.key]||'').trim()).map(c=>MONO.has(c.key)?`<div class="f"><b>${esc(c.label)}</b><pre class="cmd">${esc(r[c.key])}<button class="btn xs" data-act="copy-field" data-sheet="Runbook" data-id="${esc(r.Record_ID)}" data-key="${esc(c.key)}">Copy</button></pre></div>`
      :`<div class="f"><b>${esc(c.label)}</b><div>${esc(r[c.key])}</div></div>`).join('')}</article>`;
}
function gotoRunbook(device){
  const a=S.tables.Asset_Master.rows.find(x=>nk(x.Device)===nk(device));
  const rb=S.tables.Runbook.rows.find(r=>nk(r.Asset)===nk(device));
  const c=(rb&&rb.Customer)||a?.Customer||'(no customer)',e=(rb&&rb.Environment)||a?.Environment||'(no environment)';
  S.filters.runbook={customer:'',env:'',cat:'',q:'',ki:false};
  ['rb:c:'+c,`rb:e:${c}|${e}`,`rb:a:${c}|${e}|${device}`].forEach(k=>S.open.add(k));persistOpen();
  S.rbFocus={c,e,a:device};go('runbook');
}

/* ---------- execution plans: Year > Month > Date > Customer > Task > Phase > Steps ---------- */
function pf(){return S.filters.plans||(S.filters.plans={customer:'',status:'',q:''});}
function plansShell(){
  const f=pf(),customers=uniq(S.plans.tasks.map(t=>t.customer));
  return `<h2>Execution plans</h2><p class="sub">Year, month, day, customer and task. Inside each task, numbered phase headers (1, 2, 3) hold numbered steps (1.1, 1.2, 1.2.1). Planned times chain from the previous step unless you set one.</p>
  <div class="toolbar"><select data-filter="customer" data-scope="plans" aria-label="Customer"><option value="">All customers</option>${customers.map(c=>`<option ${c===f.customer?'selected':''}>${esc(c)}</option>`).join('')}</select>
    <select data-filter="status" data-scope="plans" aria-label="Status"><option value="">All statuses</option>${S.plans.task_statuses.map(c=>`<option ${c===f.status?'selected':''}>${esc(c)}</option>`).join('')}</select>
    <div class="grow"><input type="search" data-filter="q" data-scope="plans" placeholder="Search tasks, tickets, steps, commands" value="${esc(f.q)}" aria-label="Search plans"></div>
    <button class="btn primary" data-act="task-new">New task</button><button class="btn" data-act="templates">Templates</button>${moduleButtons('plans')}
    <button class="btn ghost" data-act="expand" data-scope="pl">Expand all</button><button class="btn ghost" data-act="collapse" data-scope="pl">Collapse all</button>
  </div><div id="plHost"></div><p class="st" style="margin-top:8px">Double-click a step cell to edit it in place. Grey times are calculated; bold times are set by hand.</p>`;
}
function taskMatches(t,f){
  if(f.customer&&t.customer!==f.customer)return false;if(f.status&&t.status!==f.status)return false;
  const q=f.q.toLowerCase();if(!q)return true;
  return [t.name,t.ticket_id,t.change_ref,t.customer,t.notes,t.assets].some(v=>String(v).toLowerCase().includes(q))||t.phases.some(p=>p.name.toLowerCase().includes(q)||p.steps.some(s=>[s.title,s.command,s.notes,s.owner].some(v=>String(v).toLowerCase().includes(q))));
}
const MONTHS=['January','February','March','April','May','June','July','August','September','October','November','December'];
function renderPlans(){
  const host=$('#plHost');if(!host)return;
  const f=pf(),tasks=S.plans.tasks.filter(t=>taskMatches(t,f));
  if(!tasks.length){host.innerHTML=`<div class="tablewrap"><div class="empty">${S.plans.tasks.length?'No tasks match these filters.':'No execution plans yet. Create a task, or import one from Excel.'}</div></div>`;return;}
  const tree={};tasks.forEach(t=>{const [y,m]=t.date.split('-');((((tree[y]=tree[y]||{})[m]=tree[y][m]||{})[t.date]=tree[y][m][t.date]||{})[t.customer]=tree[y][m][t.date][t.customer]||[]).push(t);});
  host.innerHTML=Object.keys(tree).sort().reverse().map(y=>det('pl:y:'+y,'lvl0',`<span class="label">${y}</span>`,
    Object.keys(tree[y]).sort().reverse().map(m=>det(`pl:m:${y}-${m}`,'',`<span class="label">${MONTHS[+m-1]}</span>`,
      Object.keys(tree[y][m]).sort().reverse().map(d=>{const day=new Date(d+'T00:00').toLocaleDateString([], {weekday:'long'});const n=Object.values(tree[y][m][d]).flat().length;
        return det('pl:d:'+d,'',`<span class="label">${d}</span><span class="meta">${day}, ${n} task(s)</span>`,
          Object.keys(tree[y][m][d]).sort().map(c=>det(`pl:c:${d}|${c}`,'',`<span class="label">${esc(c)}</span>`,
            tree[y][m][d][c].map(renderTask).join(''))).join(''));}).join(''))).join(''))).join('');
}
function stCls(s){return 'stb '+String(s).replace(/\s/g,'');}
function renderTask(t){
  const c=t.computed,key='pl:t:'+t.id;
  const sum=`<span class="label">${esc(t.name)}</span>${t.ticket_id?`<span class="badge pri-P4">${esc(t.ticket_id)}</span>`:''}<span class="${stCls(t.status)}">${esc(t.status)}</span>${t._access?`<span class="chip">${esc(t._owner||'')}${t._access==='partial'?': your steps':t._access==='shared'?': shared with you':''}</span>`:''}
    <span class="meta">${c.done}/${c.steps} steps, ${fmtMin(c.approx_total)}</span><div class="bar" style="width:80px"><i style="width:${c.progress}%"></i></div>${c.overrun_min?`<span class="overrun">overruns the window by ${fmtMin(c.overrun_min)}</span>`:''}`;
  const a=`data-task="${esc(t.id)}"`;
  const acc=t._access||'';  // isolation: owner, shared (whole plan) or partial (your team's steps only)
  const own=!acc||acc==='owner';
  const teamsOf=[...new Set(t.phases.map(ph=>ph.team).filter(Boolean))];
  const shareInfo=acc?`<div><span>Owner team</span><b>${esc(t._owner||t.owner_team||'—')}</b></div>
    <div><span>Shared whole with</span>${esc((t.shared_with||[]).join(', ')||'—')}</div>
    ${teamsOf.length?`<div><span>Phases given to</span>${esc(teamsOf.join(', '))}</div>`:''}`:'';
  const info=`<div class="task-info">${shareInfo}
    <div><span>Ticket ID</span>${esc(t.ticket_id||'—')}</div><div><span>Change ref</span>${esc(t.change_ref||'—')}</div>
    <div><span>Maintenance window</span>${t.window_start?esc(t.window_start)+' to '+esc(hm(t.window_end,t.window_start.slice(0,10))):'—'}</div>
    <div><span>Projected finish</span><b class="${c.overrun_min?'overrun':''}">${esc(c.projected_end)}</b></div>
    <div><span>Total approx time</span>${fmtMin(c.approx_total)}</div><div><span>Approver</span>${esc(t.approver||'—')}</div>
    ${t.assets?`<div><span>CIs</span>${esc(t.assets)}</div>`:''}${t.notes?`<div style="flex-basis:100%"><span>Notes</span>${esc(t.notes)}</div>`:''}
    <div style="flex-basis:100%" class="st">Created by ${esc(t.created_by||'?')}, last updated by ${esc(t.updated_by||'?')} ${esc(fmtTime(t.updated_at))}</div></div>`;
  const tools=!own?`<div class="task-tools"><button class="btn sm primary" data-act="run-open" ${a}>Run mode</button><button class="btn sm" data-act="task-print" ${a}>Print / PDF</button>
    <span class="st">${acc==='partial'?`Shared by ${esc(t._owner)}: you see the steps given to your team, and can update and complete them.`
      :`Shared with your team by ${esc(t._owner)}: you see the whole plan and can update the steps given to your team. The owning team changes the plan and signs off go/no-go.`}</span></div>`
    :`<div class="task-tools"><button class="btn sm primary" data-act="run-open" ${a}>Run mode</button><button class="btn sm" data-act="task-edit" ${a}>Edit task</button>${acc?`<button class="btn sm" data-act="task-share" ${a}>Share</button>`:''}
    <button class="btn sm" data-act="phase-add" ${a}>Add phase</button><button class="btn sm" data-act="step-cols" ${a}>Step columns</button>
    <button class="btn sm" data-act="task-clone" ${a}>Clone task</button><button class="btn sm" data-act="task-tpl" ${a}>Save as template</button>
    <button class="btn sm" data-act="task-print" ${a}>Print / PDF</button><button class="btn sm ghost" data-act="history" data-key="plan/${esc(t.id)}">History</button><button class="btn sm ghost" data-act="task-del" ${a}>Delete</button></div>`;
  return det(key,'',sum,`<div class="task${own?'':' ro'}">${info}${tools}${t.phases.map((ph,i)=>renderPhase(t,ph,i)).join('')||'<p class="st" style="padding:12px 16px">No phases yet. Add one, for example Pre-execution.</p>'}</div>`);
}
const SUB_KEYS=['command','rollback','expected'];
function renderPhase(t,ph,i){
  const a=`data-task="${esc(t.id)}" data-phase="${esc(ph.id)}"`;
  const gate=ph.checkpoint?(ph.signed_off_by?`<span class="gate ok">Go/no-go signed off by ${esc(ph.signed_off_by)} ${esc(hm(ph.signed_off_at,t.date))}</span><button class="btn xs ghost" data-act="phase-sign" data-undo="1" ${a}>Withdraw</button>`
    :`<span class="gate">Go/no-go checkpoint</span><button class="btn xs" data-act="phase-sign" ${a}>Sign off</button>`):'';
  const cols=t.step_columns.filter(c=>!c.hidden&&!SUB_KEYS.includes(c.key));
  const subs=t.step_columns.filter(c=>!c.hidden&&SUB_KEYS.includes(c.key));
  const showActual=cols.some(c=>c.key==='actual_end');
  const head=`<tr><th>No.</th>${cols.map(c=>`<th>${esc(c.label)}</th>${c.key==='actual_end'?'<th>Actual time</th>':''}`).join('')}<th></th></tr>`;
  const rows=ph.steps.map((st,si)=>{
    const parent=ph.steps[si+1]&&ph.steps[si+1].level>st.level;
    const sa=`${a} data-step="${esc(st.id)}"`;
    const cells=cols.map(c=>{const k=c.key;let v;
      if(k==='planned_start')v=st.planned_inherited?`<span class="inh" title="Calculated: previous step start + its approx time. Double-click to set a time.">${esc(hm(st.planned,t.date))}</span>`:`<span class="ovr" title="Set by hand. Clear it to follow the previous step again.">${esc(hm(st.planned,t.date))}</span>`;
      else if(k==='approx_min')v=st.approx_min===''?'':fmtMin(st.approx_min);
      else if(k==='actual_start'||k==='actual_end')v=esc(hm(st[k],t.date));
      else if(k==='status')v=`<span class="${stCls(st.status)}">${esc(st.status)}</span>`;
      else if(k.startsWith('x_'))v=`<div class="clip">${esc(st.custom[k]||'')}</div>`;
      else v=`<div class="clip">${esc(st[k]||'')}</div>`;
      const extra=k==='actual_end'&&showActual?`<td>${st.actual_min===''?'':fmtMin(st.actual_min)}${st.variance_min!==''&&st.variance_min!==0?` <span class="${st.variance_min>0?'bad':'st done'}">(${st.variance_min>0?'+':''}${fmtMin(st.variance_min)})</span>`:''}</td>`:'';
      return `<td class="${k==='title'?'ttl':''}" data-ie="s" ${sa} data-key="${esc(k)}">${v}</td>${extra}`;}).join('');
    const acts=`<td class="actions"><button class="btn xs" data-act="step-edit" ${sa}>Edit</button><button class="btn xs ghost" data-act="step-add" ${sa} title="Add a step below">Add</button><button class="btn xs ghost" data-act="step-add" data-child="1" ${sa} title="Add a sub-step">Sub</button>
      <button class="btn xs ghost" data-act="step-clone" ${sa}>Clone</button><button class="btn xs ghost" data-act="step-move" data-dir="up" ${sa} aria-label="Move up" title="Move up">&#8593;</button><button class="btn xs ghost" data-act="step-move" data-dir="down" ${sa} aria-label="Move down" title="Move down">&#8595;</button>
      <button class="btn xs ghost" data-act="step-move" data-dir="outdent" ${sa} aria-label="Outdent" title="Outdent">&#8592;</button><button class="btn xs ghost" data-act="step-move" data-dir="indent" ${sa} aria-label="Indent" title="Indent">&#8594;</button>
      <button class="btn xs ghost" data-act="step-del" ${sa}>Delete</button></td>`;
    const sub=subs.filter(c=>String(st[c.key]||'').trim());
    const subRow=sub.length?`<tr class="cmdrow"><td></td><td colspan="${cols.length+(showActual?1:0)+1}">${sub.map(c=>c.key==='expected'?`<div class="cmdlbl">${esc(c.label)}</div><div>${esc(st.expected)}</div>`
      :`${c.key==='command'?'':`<div class="cmdlbl">${esc(c.label)}</div>`}<pre class="cmd ${c.key==='rollback'?'rb':''}">${esc(st[c.key])}<button class="btn xs" data-act="copy-step" ${sa} data-key="${c.key}">Copy</button></pre>`).join('')}</td></tr>`:'';
    return `<tr class="stepr l${st.level} ${parent?'parent':''}"><td class="num">${esc(st.num)}</td>${cells}${acts}</tr>${subRow}`;
  }).join('');
  return `<section class="phase"><div class="phase-head"><span class="phase-no">${esc(ph.num)}</span><h4>${esc(ph.name)}</h4><span class="meta">${ph.steps.length} step(s), ${fmtMin(ph.approx_total)}</span>${gate}<span class="sp"></span>
    <button class="btn xs" data-act="step-add" ${a}>Add step</button><button class="btn xs ghost" data-act="phase-edit" ${a}>Rename</button>
    <button class="btn xs ghost" data-act="phase-move" data-dir="up" ${a} aria-label="Move phase up">&#8593;</button><button class="btn xs ghost" data-act="phase-move" data-dir="down" ${a} aria-label="Move phase down">&#8595;</button>
    <button class="btn xs ghost" data-act="phase-clone" ${a}>Clone</button><button class="btn xs ghost" data-act="phase-del" ${a}>Delete</button></div>
    ${ph.steps.length?`<div style="overflow:auto"><table class="grid steps"><thead>${head}</thead><tbody>${rows}</tbody></table></div>`:`<p class="st" style="padding:10px 16px">No steps in this phase yet.</p>`}</section>`;
}
function openTaskKeys(t){const [y,m]=t.date.split('-');['pl:y:'+y,`pl:m:${y}-${m}`,'pl:d:'+t.date,`pl:c:${t.date}|${t.customer}`,'pl:t:'+t.id].forEach(k=>S.open.add(k));persistOpen();}
function focusTask(id){const t=S.plans.tasks.find(x=>x.id===id);if(!t)return;openTaskKeys(t);S.filters.plans={customer:'',status:'',q:''};go('plans');
  requestAnimationFrame(()=>document.querySelector(`[data-key="pl:t:${id}"]`)?.scrollIntoView({block:'start',behavior:'smooth'}));}
function taskFields(t,isNew){
  const v=t||{date:today(),status:'Draft'};
  return fieldHtml({key:'date',label:'Date',type:'date',value:v.date,required:true})+fieldHtml({key:'customer',label:'Customer',value:v.customer,required:true,list:'dl-customers'})
    +fieldHtml({key:'name',label:'Task name',value:v.name,required:true,full:true})+fieldHtml({key:'ticket_id',label:'Ticket ID',value:v.ticket_id,hint:'Used with date and customer to detect duplicates.'})
    +fieldHtml({key:'change_ref',label:'Change ref',value:v.change_ref})+fieldHtml({key:'window_start',label:'Maintenance window start',type:'datetime',value:v.window_start,hint:'The first step starts here.'})
    +fieldHtml({key:'window_end',label:'Maintenance window end',type:'datetime',value:v.window_end})+fieldHtml({key:'approver',label:'Approver',value:v.approver})
    +fieldHtml({key:'status',label:'Status',type:'select',options:S.plans.task_statuses,value:v.status})+fieldHtml({key:'assets',label:'CIs',value:v.assets,full:true,hint:'Comma-separated CI names.'})
    +fieldHtml({key:'notes',label:'Notes',type:'textarea',value:v.notes})+datalists();
}
function taskModal(t){
  const tpl=t?'':`<div class="field full"><label for="tplSel">Start from</label><select id="tplSel" data-field="template">${S.plans.builtin.map(b=>`<option value="${b.id}" ${b.id==='change'?'selected':''}>${esc(b.name)}${b.phases.length?': '+esc(b.phases.join(', ')):''}</option>`).join('')}
    ${S.plans.templates.length?`<optgroup label="Saved templates">${S.plans.templates.map(x=>`<option value="${esc(x.id)}">${esc(x.name)} (${x.phases} phases, ${x.steps} steps)</option>`).join('')}</optgroup>`:''}</select></div>`;
  openModal({title:t?'Edit task':'New task',size:'wide',body:tpl+taskFields(t,!t),save:t?'Save changes':'Create task',onSave:async ov=>{
    const v=readFields(ov),template=v.template;delete v.template;
    if(!v.customer.trim()||!v.name.trim()||!v.date){toast('Date, customer and task name are required','error');return false;}
    if(t){await run(()=>planOp('task.update',{task_id:t.id,task:v}),'Task saved');return;}
    const r=await run(()=>dupSafe(a=>planOp('task.create',{task:v,template,allow_duplicate:a})),'Task created');
    focusTask(r.task_id);
  }});
}
function cloneTaskModal(t){
  openModal({title:'Clone task',body:`<p class="full st">Copies every phase, step, command and column. Actual times, statuses and sign-offs are cleared. Hand-set times and the maintenance window move with the new date.</p>`
    +fieldHtml({key:'date',label:'New date',type:'date',value:addDays(t.date,7),required:true})+fieldHtml({key:'customer',label:'Customer',value:t.customer,list:'dl-customers'})
    +fieldHtml({key:'name',label:'Task name',value:t.name,full:true})+fieldHtml({key:'ticket_id',label:'Ticket ID',value:t.ticket_id})+datalists(),
    save:'Clone task',onSave:async ov=>{const v=readFields(ov);const r=await run(()=>planOp('task.clone',{task_id:t.id,task:v}),'Task cloned');focusTask(r.task_id);}});
}
function phaseModal(t,ph){
  openModal({title:ph?'Edit phase':'Add phase',body:fieldHtml({key:'name',label:'Phase header',value:ph?ph.name:'',list:'dl-phases',full:true,hint:'For example: '+S.plans.presets.join(', ')+'.'})
    +fieldHtml({key:'checkpoint',type:'checkbox',label:'Go/no-go checkpoint',text:'Require a sign-off before the next phase can start',value:ph?ph.checkpoint:false,full:true})+datalists(),
    save:ph?'Save phase':'Add phase',onSave:async ov=>{const v=readFields(ov);if(!v.name.trim()){toast('Enter a phase name','error');return false;}
      if(ph)await run(()=>planOp('phase.update',{task_id:t.id,phase_id:ph.id,name:v.name,checkpoint:v.checkpoint}),'Phase saved');
      else await run(()=>planOp('phase.add',{task_id:t.id,name:v.name,checkpoint:v.checkpoint}),'Phase added');}});
}
function stepModal(t,ph,st,opts={}){
  const T={title:'text',owner:'text',planned_start:'datetime',approx_min:'text',actual_start:'datetime',actual_end:'datetime',status:'select',notes:'textarea',command:'mono',rollback:'mono',expected:'textarea'};
  const hint={planned_start:`Leave blank to follow the previous step${st?' (calculated: '+st.planned+')':''}.`,approx_min:'Minutes, or 1h30m, or 1:30.',command:'Shown under the step with a copy button.',rollback:'Back-out command if this step fails.'};
  const body=t.step_columns.map(c=>{const k=c.key,v=st?(k.startsWith('x_')?st.custom[k]:st[k]):(k==='status'?'Pending':(k==='owner'?S.user:''));
    return fieldHtml({key:k,label:c.label,type:T[k]||'text',options:k==='status'?S.plans.step_statuses:null,value:v===''||v==null?'':String(v),hint:hint[k],full:k==='title',required:k==='title'});}).join('');
  openModal({title:st?`Edit step ${st.num}`:(opts.child?'Add sub-step':'Add step'),size:'wide',body,save:st?'Save step':'Add step',onSave:async ov=>{
    const v=readFields(ov);if(!String(v.title||'').trim()&&!String(v.command||'').trim()){toast('Enter the step','error');return false;}
    if(st)await run(()=>planOp('step.update',{task_id:t.id,phase_id:ph.id,step_id:st.id,step:v}),'Step saved');
    else await run(()=>planOp('step.add',{task_id:t.id,phase_id:ph.id,after_step_id:opts.after||'',as_child:!!opts.child,step:v}),'Step added');}});
}
function templatesModal(){
  const list=S.plans.templates;
  openModal({title:'Plan templates',body:`<div class="full">${list.length?`<ul class="list">${list.map(x=>`<li><div class="t">${esc(x.name)}<small>${x.phases} phases, ${x.steps} steps, saved by ${esc(x.created_by)}</small></div><button class="btn xs danger" data-tdel="${esc(x.id)}">Delete</button></li>`).join('')}</ul>`
    :'<p class="st">No saved templates yet. Open a task and choose Save as template.</p>'}</div>`,
    onMount:ov=>ov.addEventListener('click',async e=>{const b=e.target.closest('[data-tdel]');if(!b)return;if(await confirmBox('Delete this template? Tasks already created from it are not affected.')){await run(()=>planOp('template.delete',{template_id:b.dataset.tdel}),'Template deleted');closeModal();templatesModal();}})});
}

/* ---------- run mode ---------- */
function openRun(tid){S.runTask=tid;openModal({title:'Run mode',size:'full',body:'<div class="full" id="runBody"></div>',onMount:ov=>ov.addEventListener('click',runClick),onClose:()=>{S.runTask=null;}});renderRun();}
function renderRun(){
  const el=$('#runBody');if(!el)return;const t=S.plans.tasks.find(x=>x.id===S.runTask);
  if(!t){el.innerHTML='<p class="st">This task no longer exists.</p>';return;}
  const c=t.computed;let blocked=false,cur=null;
  for(const ph of t.phases){for(const st of ph.steps){if(!['Done','Skipped'].includes(st.status)){cur=st.id;break;}}if(cur)break;}
  let html=`<div class="run-head"><b style="font-size:16px">${esc(t.name)}</b>${t.ticket_id?`<span class="badge pri-P4">${esc(t.ticket_id)}</span>`:''}<span>Window ${esc(t.window_start||'—')} to ${esc(t.window_end||'—')}</span>
    <span>Projected finish <b class="${c.overrun_min?'overrun':''}">${esc(c.projected_end)}</b></span><span>Now ${esc(nowLocal())}</span><span>${c.done}/${c.steps} done</span><div class="bar" style="width:220px"><i style="width:${c.progress}%"></i></div></div>`;
  t.phases.forEach(ph=>{
    html+=`<h3 style="margin:18px 0 4px"><span class="phase-no">${esc(ph.num)}</span> ${esc(ph.name)}</h3>`;
    ph.steps.forEach(st=>{const d=`data-phase="${esc(ph.id)}" data-step="${esc(st.id)}"`,dis=blocked?'disabled title="Waiting for a go/no-go sign-off"':'';
      const btns=st.status==='Pending'?`<button class="btn sm primary" data-run="start" ${d} ${dis}>Start</button><button class="btn sm" data-run="done" ${d} ${dis}>Done</button><button class="btn sm ghost" data-run="skip" ${d}>Skip</button>`
        :st.status==='In Progress'?`<button class="btn sm primary" data-run="done" ${d}>Done</button><button class="btn sm danger" data-run="fail" ${d}>Failed</button>`
        :`<span class="${stCls(st.status)}">${esc(st.status)}</span><button class="btn sm ghost" data-run="reset" ${d}>Reset</button>`;
      html+=`<div class="run-step ${st.id===cur?'cur':''}" style="margin-left:${st.level*18}px"><div class="n">${esc(st.num)}</div><div><div><b>${esc(st.title)}</b> <span class="st">${esc(st.owner||'')}</span></div>
        <div class="st">Planned ${esc(hm(st.planned,t.date))}, ${fmtMin(st.approx_min)}${st.actual_start?`; started ${esc(hm(st.actual_start,t.date))}`:''}${st.actual_end?`, ended ${esc(hm(st.actual_end,t.date))} (${fmtMin(st.actual_min)})`:''}</div>
        ${st.command?`<pre class="cmd">${esc(st.command)}<button class="btn xs" data-rcopy="command" ${d}>Copy</button></pre>`:''}${st.rollback?`<div class="cmdlbl">Rollback</div><pre class="cmd rb">${esc(st.rollback)}<button class="btn xs" data-rcopy="rollback" ${d}>Copy</button></pre>`:''}
        ${st.expected?`<div class="cmdlbl">Expected output</div><div>${esc(st.expected)}</div>`:''}</div><div class="btns">${btns}</div></div>`;});
    if(ph.checkpoint){html+=ph.signed_off_by?`<p><span class="gate ok">Go/no-go signed off by ${esc(ph.signed_off_by)} at ${esc(ph.signed_off_at)}</span> <button class="btn xs ghost" data-sign="undo" data-phase="${esc(ph.id)}">Withdraw</button></p>`
      :`<p><span class="gate">Go/no-go checkpoint: the next phase is locked until this is signed off.</span> <button class="btn sm primary" data-sign="1" data-phase="${esc(ph.id)}">Sign off: go</button></p>`;
      if(!ph.signed_off_by)blocked=true;}
  });
  el.innerHTML=html;
}
async function runClick(e){
  const b=e.target.closest('[data-run],[data-sign],[data-rcopy]');if(!b||b.disabled)return;
  const ds=b.dataset;
  if(ds.rcopy){const {st}=stepOf(S.runTask,ds.phase,ds.step);return st&&copyText(st[ds.rcopy]);}
  try{if(ds.run)await run(()=>planOp('step.action',{task_id:S.runTask,phase_id:ds.phase,step_id:ds.step,action:ds.run}));
      else await run(()=>planOp('phase.signoff',{task_id:S.runTask,phase_id:ds.phase,undo:ds.sign==='undo'}),ds.sign==='undo'?'Sign-off withdrawn':'Signed off');}catch(_){}
}

/* ---------- shift log ---------- */
function renderShiftLog(){
  const years=Object.keys(S.hier.years||{}).sort().reverse();
  const tb=`<h2>Shift log</h2><p class="sub">Year, month, day and shift, each with the seven standard handover sections. Download the template to fill shifts in Excel and import them; the Excel export imports back the same way.</p>
  <div class="toolbar"><input type="date" id="hDate" value="${esc(S.hDate)}" aria-label="Date"><select id="hShift" aria-label="Shift">${SHIFTS.map(s=>`<option ${s===S.hShift?'selected':''}>${s}</option>`).join('')}</select>
    <label class="st"><input type="checkbox" id="hCarry" ${S.carry?'checked':''}> Carry rows forward from the previous shift</label><button class="btn primary" data-act="h-open">Open shift</button><span class="grow"></span>
    ${moduleButtons('shiftlog')}<button class="btn" data-act="export-menu">Export</button>
    <button class="btn ghost" data-act="expand" data-scope="sl">Expand all</button><button class="btn ghost" data-act="collapse" data-scope="sl">Collapse all</button></div><div id="slHost">`;
  if(!years.length)return tb+`<div class="tablewrap"><div class="empty">No shifts recorded yet. Pick a date and shift above, then select Open shift.</div></div></div>`;
  return tb+years.map(y=>{const months=Object.keys(S.hier.years[y].months).sort().reverse();
    return det('y:'+y,'lvl0',`<span class="label">${y}</span><span class="meta">${months.length} month(s)</span><span class="sp"></span><button class="btn xs ghost" data-act="h-del-node" data-level="year" data-y="${y}">Delete year</button>`,
      months.map(m=>{const mn=S.hier.years[y].months[m],dates=Object.keys(mn.dates).sort().reverse();
        return det(`m:${y}-${m}`,'',`<span class="label">${esc(mn.name)}</span><span class="meta">${dates.length} day(s)</span><span class="sp"></span><button class="btn xs ghost" data-act="h-del-node" data-level="month" data-y="${y}" data-m="${m}">Delete month</button>`,
          dates.map(d=>{const dn=mn.dates[d],shifts=SHIFTS.filter(s=>dn.shifts[s]);
            return det('d:'+d,'',`<span class="label">${esc(d)}</span><span class="meta">${esc(dn.day)}, ${shifts.length} shift(s)</span><span class="sp"></span><button class="btn xs ghost" data-act="h-del-node" data-level="date" data-d="${d}">Delete day</button>`,
              shifts.map(s=>renderShift(d,s,dn.shifts[s])).join('')||'<p class="st">No shifts on this day.</p>');}).join(''));}).join(''));}).join('')+'</div>';
}
function renderShift(d,s,node){
  const rows=node.sections.reduce((a,x)=>a+x.rows.length,0);
  const sum=`<span class="label">${s}</span><span class="meta">${rows} row(s), opened by ${esc(node.created_by||'unknown')}</span><span class="sp"></span>
    ${node.cloned_from?`<span class="carried">cloned from ${esc(node.cloned_from)}</span>`:''}<button class="btn xs" data-act="h-clone-shift" data-d="${d}" data-s="${s}">Clone shift</button>
    <button class="btn xs" data-act="h-add-sec" data-d="${d}" data-s="${s}">Add section</button><button class="btn xs ghost" data-act="h-del-node" data-level="shift" data-d="${d}" data-s="${s}">Delete shift</button>`;
  return `<details class="node" data-key="s:${d}:${s}" ${S.open.has(`s:${d}:${s}`)?'open':''}><summary class="shiftsum" style="--c:${SHIFT_COLOR[s]}">${sum}</summary>${node.sections.map(sec=>renderSection(d,s,sec)).join('')}</details>`;
}
function renderSection(d,s,sec){
  const a=`data-d="${d}" data-s="${s}" data-sec="${esc(sec.id)}"`;
  const head=`<div class="sec-head"><h4>${esc(sec.name)}</h4><span class="meta">${sec.rows.length} row(s)</span><span class="sp"></span>
    <button class="btn xs primary" data-act="h-add-row" ${a}>Add row</button><button class="btn xs" data-act="h-cols" ${a}>Columns</button>${sec.template?'':`<button class="btn xs ghost" data-act="h-del-sec" ${a}>Delete section</button>`}</div>`;
  if(!sec.rows.length)return `<div class="sec">${head}<div class="empty">No rows in this section.</div></div>`;
  return `<div class="sec">${head}<div style="overflow:auto"><table class="grid"><thead><tr>${sec.columns.map(c=>`<th>${esc(label(c.name))}</th>`).join('')}<th></th></tr></thead><tbody>
    ${sec.rows.map(r=>`<tr>${sec.columns.map((c,i)=>`<td>${/^P[1-4]$/.test(r.cells[c.id]||'')?`<span class="badge pri-${esc(r.cells[c.id])}">${esc(r.cells[c.id])}</span>`:`<div class="clip">${esc(r.cells[c.id]||'')}</div>`}${i===0&&r.carried_from?`<div class="carried">from ${esc(r.carried_from)}</div>`:''}</td>`).join('')}
      <td class="actions"><button class="btn xs" data-act="h-edit-row" ${a} data-row="${esc(r.id)}">Edit</button><button class="btn xs ghost" data-act="h-clone-row" ${a} data-row="${esc(r.id)}">Clone</button><button class="btn xs ghost" data-act="h-del-row" ${a} data-row="${esc(r.id)}">Delete</button></td></tr>`).join('')}</tbody></table></div></div>`;
}
function findSection(ds){return shiftNode(ds.d,ds.s)?.sections.find(x=>x.id===ds.sec);}
function editHRow(ds){
  const sec=findSection(ds);if(!sec)return load();
  const row=ds.row?sec.rows.find(r=>r.id===ds.row):null;
  const body=sec.columns.map(c=>fieldHtml({key:c.id,label:label(c.name),type:/details|notes|mom|updates|playbook|task|next/i.test(c.name)?'textarea':'text',value:row?row.cells[c.id]:'',list:/client|customer/i.test(c.name)?'dl-customers':''})).join('')+datalists();
  openModal({title:`${row?'Edit':'Add'} row: ${sec.name}`,size:'wide',body,save:row?'Save changes':'Add row',onSave:async ov=>{
    const base={date:ds.d,shift:ds.s,section_id:ds.sec,cells:readFields(ov)};
    if(row)await run(()=>api('PUT','/api/hierarchy/row',{...base,row_id:row.id}),'Row saved');else await run(()=>api('POST','/api/hierarchy/row',base),'Row added');}});
}
function editSectionColumns(ds){
  const sec=findSection(ds);if(!sec)return load();
  columnEditor({title:'Columns: '+sec.name,cols:sec.columns.map(c=>({key:c.id,label:c.name,hidden:false,core:false})),
    note:'Rename or reorder columns without losing data. Removing a column deletes its values.',
    onSave:cols=>run(()=>api('PUT','/api/hierarchy/section/columns',{date:ds.d,shift:ds.s,section_id:ds.sec,columns:cols.map(c=>({id:c.key,name:c.label}))}),'Columns saved')});
}
function openShiftKeys(d,s){const [y,m]=d.split('-');['y:'+y,`m:${y}-${m}`,'d:'+d,`s:${d}:${s}`].forEach(k=>S.open.add(k));persistOpen();}
async function openShift(d,s,carry){
  const r=await run(()=>api('POST','/api/hierarchy/node',{level:'shift',date:d,shift:s,carry_forward:!!carry}));
  const [y,m]=d.split('-');['y:'+y,`m:${y}-${m}`,'d:'+d,`s:${d}:${s}`].forEach(k=>S.open.add(k));persistOpen();
  if(r.created)toast(r.carried_from?`Started ${s} shift with rows carried from ${r.carried_from}`:`Started ${s} shift`,'success');
  S.hDate=d;S.hShift=s;go('shiftlog');requestAnimationFrame(()=>document.querySelector(`[data-key="s:${d}:${s}"]`)?.scrollIntoView({block:'start',behavior:'smooth'}));
}

/* ---------- bulletins ---------- */
function showBulletin(id){
  const b=findRow('Bulletins',id);if(!b)return;
  const cols=S.tables.Bulletins.columns.filter(c=>!c.system&&!c.core&&String(b[c.key]||'').trim());
  const link=/^https?:\/\//i.test(b.Link||'')?`<p class="full"><a href="${esc(b.Link)}" target="_blank" rel="noopener noreferrer">${esc(b.Link)}</a></p>`:(b.Link?`<p class="full">${esc(b.Link)}</p>`:'');
  openModal({title:b.Title,size:'wide',body:`<p class="full"><span class="sev sev-${esc(b.Severity)}">${esc(b.Severity)}</span> ${truthy(b.Pinned)?'<span class="chip">pinned</span>':''} <span class="st">${b.Customers?'For '+esc(b.Customers)+'. ':''}From ${esc(b.Start)}${b.Expiry?' until '+esc(b.Expiry):''}. Posted by ${esc(b.Posted_By||'?')}.</span></p>
    <div class="full" style="white-space:pre-wrap;font-size:14.5px">${esc(b.Body||'No further details.')}</div>${link}${cols.map(c=>`<p class="full"><b>${esc(c.label)}:</b> ${esc(b[c.key])}</p>`).join('')}`,
    footExtra:`<button class="btn" data-bunread>Mark as unread</button><button class="btn" data-bedit>Edit</button>`,
    onMount:(ov,m)=>ov.addEventListener('click',async e=>{
      if(e.target.closest('[data-bunread]')){await run(()=>api('POST',`/api/bulletins/${encodeURIComponent(id)}/read`,{read:false}));closeModal(m);}
      if(e.target.closest('[data-bedit]')){closeModal(m);editRow('Bulletins',id);}})});
  if(!S.reads.has(id))api('POST',`/api/bulletins/${encodeURIComponent(id)}/read`,{read:true}).then(()=>{S.reads.add(id);renderTicker();}).catch(()=>{});
}

/* ---------- activity ---------- */
const MOD_TITLE={Asset_Master:'CMDB (configuration items)',Runbook:'Runbook',Incidents:'Incidents',Shift_Transitions:'Shift handovers',Bulletins:'Bulletins',plans:'Execution plans',shiftlog:'Shift log'};
async function renderActivity(el){
  el.innerHTML=`<h2>Activity</h2><p class="sub">Imports you can undo, and the full audit trail of who changed what and when.</p>
    <section class="panel" style="margin-bottom:18px"><h3>Imports</h3><div id="impHost"></div></section>
    <h3 style="margin:6px 0 10px">Audit trail</h3><div class="toolbar"><div class="grow"><input type="search" id="auditQ" placeholder="Search by person, area or record" aria-label="Search audit"></div></div><div id="auditHost"><p class="st">Loading…</p></div>`;
  $('#impHost').innerHTML=S.imports.length?`<div class="tablewrap" style="max-height:none"><table class="grid"><thead><tr><th>When</th><th>Who</th><th>File</th><th>Result</th><th></th></tr></thead><tbody>${S.imports.map(b=>`<tr><td class="st">${esc(fmtTime(b.at))}</td><td>${esc(b.by)}</td><td>${esc(b.file)}</td>
    <td>${(b.sections||[]).map(s=>`${esc(MOD_TITLE[s.module]||s.module)}: ${s.added} added, ${s.updated} updated, ${s.skipped} skipped`).join('<br>')}</td>
    <td class="actions">${b.undone?`<span class="st">Undone by ${esc(b.undone_by)}</span>`:(b.created_count||b.updated_count)?`<button class="btn xs" data-act="undo-import" data-id="${esc(b.id)}">Undo</button>`:''}</td></tr>`).join('')}</tbody></table></div>`:'<p class="st">No imports yet.</p>';
  let rows=[];try{rows=(await api('GET','/api/audit?limit=3000')).audit;}catch(e){toast(e.message,'error');}
  const draw=()=>{const q=($('#auditQ')?.value||'').toLowerCase(),host=$('#auditHost');if(!host)return;
    const list=rows.filter(a=>!q||Object.values(a).some(v=>String(v).toLowerCase().includes(q))).slice(0,1000);
    host.innerHTML=list.length?`<div class="tablewrap"><table class="grid"><thead><tr><th>When</th><th>Who</th><th>Area</th><th>Action</th><th>Record</th><th>Details</th></tr></thead><tbody>
      ${list.map(a=>`<tr><td class="st">${esc(fmtTime(a.Timestamp))}</td><td>${esc(a.User)}</td><td>${esc(a.Area)}</td><td>${esc(a.Action)}</td><td class="mono">${esc(a.Target)}</td><td><div class="clip">${esc(a.Details)}</div></td></tr>`).join('')}</tbody></table></div>`:'<div class="tablewrap"><div class="empty">No matching entries.</div></div>';};
  draw();$('#auditQ')?.addEventListener('input',draw);
}

/* ---------- import / templates / export / identity ---------- */
const IMPORT_MODULES=[['auto','Detect automatically'],['assets','CMDB (configuration items)'],['runbook','Runbook'],['plans','Execution plans'],['incidents','Incidents'],['handovers','Shift handovers'],['bulletins','Bulletins'],['shiftlog','Shift log']];
function importDialog(module='auto',file=null){
  let chosen=file;
  openModal({title:'Import data',save:'Preview import',
    body:`<p class="full st">Excel (.xlsx), CSV, notepad text (.txt, comma or tab separated) or JSON. Title rows above the headers are skipped automatically. Nothing is saved until you review the preview.</p>
      ${fieldHtml({key:'module',label:'What does the file contain?',type:'select',options:IMPORT_MODULES.map(x=>x[0]),value:module})}
      <div class="field"><label for="impFile">File</label><input type="file" id="impFile" accept=".xlsx,.xlsm,.csv,.txt,.tsv,.json"><span class="hint" id="impName">${chosen?'Selected: '+esc(chosen.name):''}</span></div>
      <div class="full st">Need the format? Download a template for the chosen data: <button class="btn xs" data-tpl="xlsx">Excel</button> <button class="btn xs" data-tpl="csv">CSV</button> <button class="btn xs" data-tpl="txt">Text (tab)</button> <button class="btn xs" data-tpl="json">JSON</button></div>`,
    onMount:ov=>{const sel=ov.querySelector('[data-field="module"]');[...sel.options].forEach(o=>{if(!o.value){o.remove();return;}const m=IMPORT_MODULES.find(x=>x[0]===o.value);if(m)o.textContent=m[1];});
      ov.querySelector('#impFile').addEventListener('change',e=>{chosen=e.target.files[0];ov.querySelector('#impName').textContent=chosen?'Selected: '+chosen.name:'';});
      ov.addEventListener('click',e=>{const b=e.target.closest('[data-tpl]');if(!b)return;const m=sel.value;if(m==='auto')return toast('Choose the kind of data first','error');download(`/api/template/${m}?format=${b.dataset.tpl}`);});},
    onSave:async ov=>{if(!chosen){toast('Choose a file first','error');return false;}
      const fd=new FormData();fd.append('file',chosen,chosen.name);fd.append('module',readFields(ov).module);
      let d;try{d=await api('POST','/api/import/preview',fd);}catch(e){toast(e.message,'error');return false;}
      setTimeout(()=>previewDialog(d),0);}});
}
const PV_LABEL={new:'New',duplicate:'Duplicate',duplicate_in_file:'Repeated in file',invalid:'Invalid',replace:'Replace'};
const PV_OPTS={new:[['add','Add'],['skip','Skip']],duplicate:[['skip','Skip'],['update','Update existing'],['keep','Keep both']],duplicate_in_file:[['skip','Skip'],['keep','Keep both']],replace:[['replace','Replace'],['skip','Skip']]};
function previewDialog(d){
  const body=`<div class="full"><p class="st" style="margin-bottom:10px">Review what will happen to each record in <b>${esc(d.filename)}</b>. Duplicates are matched on key fields (for example Serial No, or Customer and Device; Ticket ID for plans).</p>
    ${d.sections.map((s,si)=>{const dups=s.items.some(i=>i.status==='duplicate');
      return `<div class="pv-sec"><header><b>${esc(s.title)}</b><span class="st">from ${esc(s.source)}${s.header_row?`, headers found on row ${s.header_row}`:''}</span>
        ${Object.entries(s.counts).map(([k,v])=>`<span class="chip ${k}">${v} ${esc((PV_LABEL[k]||k).toLowerCase())}</span>`).join('')}
        ${s.new_columns.length?`<span class="st warn">New columns will be added: ${esc(s.new_columns.join(', '))}</span>`:''}<span class="sp" style="flex:1"></span>
        ${dups?`<label class="st">All duplicates: <select data-secdef="${si}"><option value="skip">Skip</option><option value="update">Update existing</option><option value="keep">Keep both</option></select></label>`:''}</header>
        <div style="max-height:44vh;overflow:auto"><table class="grid"><thead><tr><th>Status</th><th>Record</th><th>Details</th><th>Action</th></tr></thead><tbody>
        ${s.items.map(it=>`<tr><td><span class="chip ${it.status}">${esc(PV_LABEL[it.status]||it.status)}</span></td><td>${esc(it.label)}</td>
          <td class="st">${esc((it.errors||[]).join('; '))}${it.match_label?`Matches ${esc(it.match_label)}`:''}${it.changes&&it.changes.length?`. Would change: ${esc(it.changes.join(', '))}`:''}</td>
          <td>${PV_OPTS[it.status]?`<select data-dec="${si}:${it.i}" data-status="${it.status}">${PV_OPTS[it.status].map(([v,l])=>`<option value="${v}">${l}</option>`).join('')}</select>`:'<span class="st">Skipped</span>'}</td></tr>`).join('')}
        </tbody></table></div></div>`;}).join('')}
    ${d.ignored&&d.ignored.length?`<p class="st warn">Not imported: ${esc(d.ignored.join('; '))}</p>`:''}</div>`;
  openModal({title:'Import preview',size:'full',body,save:'Import',
    onMount:ov=>ov.addEventListener('change',e=>{const sd=e.target.dataset.secdef;if(sd===undefined)return;
      ov.querySelectorAll(`[data-dec^="${sd}:"][data-status="duplicate"]`).forEach(s=>{s.value=e.target.value;});}),
    onSave:async ov=>{const decisions={};ov.querySelectorAll('[data-dec]').forEach(s=>{const [si,i]=s.dataset.dec.split(':');(decisions[si]=decisions[si]||{})[i]=s.value;});
      const r=await run(()=>api('POST','/api/import/commit',{token:d.token,decisions}));
      const b=r.batch,msg=b.sections.map(s=>`${MOD_TITLE[s.module]||s.module}: ${s.added} added, ${s.updated} updated, ${s.skipped} skipped`).join('. ');
      toast('Imported. '+msg,'success',{label:'Undo',fn:()=>undoImport(b.id)});}});
}
async function undoImport(id){if(!await confirmBox('Undo this import? Added records are removed and updated records go back to how they were.','Undo import',true))return;
  const r=await run(()=>api('POST',`/api/imports/${encodeURIComponent(id)}/undo`,{}));toast(`Import undone: ${r.removed} removed, ${r.restored} restored`,'success');}
function templateDialog(module){
  openModal({title:'Download a template',body:`<p class="full st">The template uses the current column names, including any you renamed or added, with an example row and an instructions sheet (Excel).</p>
    ${fieldHtml({key:'module',label:'Data',type:'select',options:IMPORT_MODULES.filter(x=>x[0]!=='auto').map(x=>x[0]),value:module,full:true})}
    <div class="full"><button class="btn primary" data-tpl="xlsx">Excel (.xlsx)</button> <button class="btn" data-tpl="csv">CSV</button> <button class="btn" data-tpl="txt">Notepad text (tab separated)</button> <button class="btn" data-tpl="json">JSON</button></div>`,
    onMount:ov=>{const sel=ov.querySelector('[data-field="module"]');[...sel.options].forEach(o=>{if(!o.value){o.remove();return;}const m=IMPORT_MODULES.find(x=>x[0]===o.value);if(m)o.textContent=m[1];});
      ov.addEventListener('click',e=>{const b=e.target.closest('[data-tpl]');if(b)download(`/api/template/${sel.value}?format=${b.dataset.tpl}`);});}});
}
async function download(path){
  try{const r=await api('GET',path,undefined,{raw:true}),blob=await r.blob();
    const name=((r.headers.get('Content-Disposition')||'').match(/filename="([^"]+)"/)||[])[1]||'download';
    const a=document.createElement('a');a.href=URL.createObjectURL(blob);a.download=name;document.body.appendChild(a);a.click();
    setTimeout(()=>{URL.revokeObjectURL(a.href);a.remove();},500);toast('Downloaded '+name,'success');}
  catch(e){toast(e.message,'error');}
}
/* ---------- event wiring ---------- */
function taskOf(ds){return S.plans.tasks.find(t=>t.id===ds.task);}
const ACT={
  'nav':ds=>go(ds.view),'export-xlsx':()=>download('/api/export/xlsx'),'export-json':()=>download('/api/export/json'),
  'export-menu':()=>exportMenu(),'palette':()=>openPalette(),'theme':()=>toggleTheme(),'logout':()=>logout(),'account':()=>accountModal(),
  'side-toggle':()=>{const app=$('#app');if(innerWidth<=900)app.classList.toggle('mobile-open');else{app.classList.toggle('collapsed');localStorage.setItem('msp.collapsed',app.classList.contains('collapsed')?'1':'');}},
  'ssh':ds=>sshDialog({asset_id:ds.id}),'rdp':ds=>rdpDialog(ds.id),'term-new':()=>sshDialog({}),
  'task-print':ds=>printTask(taskOf(ds)),'h-clone-shift':ds=>cloneShiftModal(ds),
  'import':ds=>importDialog(ds.module||'auto'),'template':ds=>templateDialog(ds.module),'bulletin':ds=>showBulletin(ds.id),
  'add-row':ds=>editRow(ds.sheet,null),'edit-row':ds=>editRow(ds.sheet,ds.id),'clone-row':ds=>cloneRow(ds.sheet,ds.id),
  'del-row':async ds=>{if(await confirmBox(`Delete this ${S.schema[ds.sheet].noun}? It stays in the version history and can be restored.`))await run(()=>api('DELETE',`/api/table/${ds.sheet}/${encodeURIComponent(ds.id)}`),'Deleted');},
  'history':ds=>showHistory(ds.key),'columns':ds=>tableColumns(ds.sheet),
  'copy-field':ds=>{const r=findRow(ds.sheet,ds.id);r&&copyText(r[ds.key]);},
  'goto-runbook':ds=>{const r=findRow('Asset_Master',ds.id);if(r)gotoRunbook(r.Device);},'goto-asset-rb':ds=>gotoRunbook(ds.asset),
  'rb-add':ds=>editRow('Runbook',null,{Customer:ds.c,Environment:ds.e,Asset:ds.a}),
  'expand':ds=>document.querySelectorAll('details.node').forEach(d=>{if(d.dataset.key&&(ds.scope==='sl'?!/^(rb|pl):/.test(d.dataset.key):d.dataset.key.startsWith(ds.scope+':')))d.open=true;}),
  'collapse':ds=>document.querySelectorAll('details.node').forEach(d=>{if(d.dataset.key&&(ds.scope==='sl'?!/^(rb|pl):/.test(d.dataset.key):d.dataset.key.startsWith(ds.scope+':')))d.open=false;}),
  'open-task':ds=>focusTask(ds.id),'task-new':()=>taskModal(null),'templates':()=>templatesModal(),
  'task-edit':ds=>taskModal(taskOf(ds)),'task-clone':ds=>cloneTaskModal(taskOf(ds)),'run-open':ds=>openRun(ds.task),
  'task-tpl':ds=>{const t=taskOf(ds);openModal({title:'Save as template',body:fieldHtml({key:'name',label:'Template name',value:t.name,full:true,hint:'Phases, steps, commands and columns are saved. Times and statuses are not.'}),save:'Save template',
    onSave:async ov=>{await run(()=>planOp('template.save',{task_id:t.id,name:readFields(ov).name}),'Template saved');}});},
  'task-del':async ds=>{const t=taskOf(ds);if(await confirmBox(`Delete the task "${t.name}" on ${t.date}? It stays in the version history and can be restored.`))await run(()=>planOp('task.delete',{task_id:t.id}),'Task deleted');},
  'step-cols':ds=>{const t=taskOf(ds);columnEditor({title:'Step columns: '+t.name,cols:t.step_columns,note:'Rename, reorder, show or hide step columns, or add your own. Command, rollback and expected output show beneath each step.',
    onSave:cols=>run(()=>planOp('task.columns',{task_id:t.id,columns:cols}),'Step columns saved')});},
  'phase-add':ds=>phaseModal(taskOf(ds),null),'phase-edit':ds=>{const t=taskOf(ds);phaseModal(t,t.phases.find(p=>p.id===ds.phase));},
  'phase-move':ds=>run(()=>planOp('phase.move',{task_id:ds.task,phase_id:ds.phase,dir:ds.dir})),
  'phase-clone':ds=>run(()=>planOp('phase.clone',{task_id:ds.task,phase_id:ds.phase}),'Phase cloned'),
  'phase-del':async ds=>{if(await confirmBox('Delete this phase and all of its steps?'))await run(()=>planOp('phase.delete',{task_id:ds.task,phase_id:ds.phase}),'Phase deleted');},
  'phase-sign':ds=>run(()=>planOp('phase.signoff',{task_id:ds.task,phase_id:ds.phase,undo:!!ds.undo}),ds.undo?'Sign-off withdrawn':'Signed off'),
  'step-add':ds=>{const {t,ph}=stepOf(ds.task,ds.phase,ds.step);stepModal(t,ph,null,{after:ds.step||'',child:!!ds.child});},
  'step-edit':ds=>{const {t,ph,st}=stepOf(ds.task,ds.phase,ds.step);stepModal(t,ph,st);},
  'step-clone':ds=>run(()=>planOp('step.clone',{task_id:ds.task,phase_id:ds.phase,step_id:ds.step}),'Step cloned'),
  'step-move':ds=>run(()=>planOp('step.move',{task_id:ds.task,phase_id:ds.phase,step_id:ds.step,dir:ds.dir})),
  'step-del':async ds=>{if(await confirmBox('Delete this step and its sub-steps?'))await run(()=>planOp('step.delete',{task_id:ds.task,phase_id:ds.phase,step_id:ds.step}),'Step deleted');},
  'copy-step':ds=>{const {st}=stepOf(ds.task,ds.phase,ds.step);st&&copyText(st[ds.key]);},
  'undo-import':ds=>undoImport(ds.id),
  'open-shift':ds=>{if(S.rank<2&&!shiftNode(ds.d,ds.s))return toast('This shift has not been started yet.','info');if(S.rank<2){openShiftKeys(ds.d,ds.s);return go('shiftlog');}openShift(ds.d,ds.s,true);},
  'h-open':()=>{S.hDate=$('#hDate').value;S.hShift=$('#hShift').value;S.carry=$('#hCarry').checked;if(!S.hDate)return toast('Pick a date first','error');return openShift(S.hDate,S.hShift,S.carry);},
  'h-add-row':ds=>editHRow(ds),'h-edit-row':ds=>editHRow(ds),'h-cols':ds=>editSectionColumns(ds),
  'h-clone-row':ds=>run(()=>api('POST','/api/hierarchy/row/clone',{date:ds.d,shift:ds.s,section_id:ds.sec,row_id:ds.row}),'Row cloned'),
  'h-del-row':async ds=>{if(await confirmBox('Delete this row?'))await run(()=>api('DELETE','/api/hierarchy/row',{date:ds.d,shift:ds.s,section_id:ds.sec,row_id:ds.row}),'Row deleted');},
  'h-del-sec':async ds=>{if(await confirmBox('Delete this section and all of its rows?'))await run(()=>api('DELETE','/api/hierarchy/section',{date:ds.d,shift:ds.s,section_id:ds.sec}),'Section deleted');},
  'h-add-sec':ds=>openModal({title:'Add section',save:'Add section',body:fieldHtml({key:'name',label:'Section name',full:true})+fieldHtml({key:'columns',label:'Columns (comma separated)',value:'Item, Details, Owner, Status',full:true}),
    onSave:async ov=>{const v=readFields(ov);if(!v.name.trim()){toast('Enter a section name','error');return false;}
      await run(()=>api('POST','/api/hierarchy/section',{date:ds.d,shift:ds.s,name:v.name,columns:v.columns.split(',').map(x=>x.trim()).filter(Boolean)}),'Section added');}}),
  'h-del-node':async ds=>{const what=ds.level==='shift'?`the ${ds.s} shift on ${ds.d}`:ds.level==='date'?`every shift on ${ds.d}`:ds.level==='month'?`all of ${ds.y}-${ds.m}`:`all of ${ds.y}`;
    if(await confirmBox(`Delete ${what}? Everything under it is removed.`))await run(()=>api('DELETE','/api/hierarchy/node',{level:ds.level,year:ds.y,month:ds.m,date:ds.d,shift:ds.s}),'Deleted');},
};
/* ---------- shift roster (served by msp_roster.py) ---------- */
const RO={id:localStorage.getItem('msp.roster')||'',start:today(),mode:'week',layout:'',q:'',hide:{},data:null};
const RO_DAYS=['Mon','Tue','Wed','Thu','Fri','Sat','Sun'];
function roDate(d,n){const x=new Date(d+'T12:00:00');x.setDate(x.getDate()+n);return x.toLocaleDateString('en-CA');}
function roMonth(d,n){const x=new Date(d+'T12:00:00');x.setDate(1);x.setMonth(x.getMonth()+n);return x.toLocaleDateString('en-CA');}
function roColor(c){const s=(RO.data?.config.shifts||[]).find(x=>x.code===c);return s?.color||(c==='LEAVE'?'#F0B4B0':'');}
function roDaysText(days){if(days.length===7)return 'Mon-Sun';const s=[...days].sort((a,b)=>a-b);
  if(s.length>1&&s.every((v,i)=>i===0||v===s[i-1]+1))return RO_DAYS[s[0]]+'-'+RO_DAYS[s[s.length-1]];return s.map(i=>RO_DAYS[i]).join(',');}
function roParseDays(t){const out=new Set();String(t||'').split(',').map(x=>x.trim()).filter(Boolean).forEach(part=>{
  const [a,b]=part.split('-').map(x=>RO_DAYS.findIndex(d=>d.toLowerCase()===x.trim().slice(0,3).toLowerCase()));
  if(a<0||(part.includes('-')&&b<0))throw new Error('Unknown day in "'+part+'". Use Mon, Tue ... Sun, or a range such as Mon-Fri.');
  if(part.includes('-')){for(let i=a;;i=(i+1)%7){out.add(i);if(i===b)break;}}else out.add(a);});
  if(!out.size)throw new Error('Each shift needs at least one working day.');return [...out].sort((a,b)=>a-b);}
function roLines(t){return String(t||'').split('\n').map(x=>x.trim()).filter(x=>x&&!x.startsWith('#'));}
function roCells(line){return line.split('|').map(x=>x.trim());}
function roNewId(){return 'mem_'+Math.random().toString(16).slice(2,12);}

async function renderRoster(el){
  el.innerHTML=`<h2>Shift roster</h2><p class="sub">Who works which shift, with leave, swaps and cover, on-call and public holidays. Each team has its own roster.</p><div id="roHost"><p class="st">Loading…</p></div>`;
  try{
    if(!RO.id){const l=await api('GET','/api/roster');RO.id=l.rosters[0]?.id||'';RO.meta=l;}
    if(!RO.id){$('#roHost').innerHTML=roEmpty();return;}
    RO.data=await api('GET',`/api/roster/${encodeURIComponent(RO.id)}/view?start=${encodeURIComponent(RO.start)}&mode=${RO.mode}`);
  }catch(e){
    if(e.status===404&&!/not installed/i.test(e.message)&&RO.id){RO.id='';localStorage.removeItem('msp.roster');return renderRoster(el);}
    $('#roHost').innerHTML=`<div class="empty">${esc(e.message)}</div>`;return;
  }
  localStorage.setItem('msp.roster',RO.id);
  if(!RO.layout)RO.layout=RO.data.roster.layout||'by_person';
  roDraw();
}
function roEmpty(){return `<div class="tablewrap"><div class="empty">No rosters yet. A Lead can create one from a template, or import an existing roster sheet (Excel, CSV or text).
  <div class="toolbar" style="justify-content:center;margin-top:12px"><button class="btn primary" data-act="ro-new">New roster</button><button class="btn" data-act="ro-import">Import roster sheet</button><button class="btn ghost" data-act="ro-deleted">Deleted rosters</button></div></div></div>`;}
function roDraw(){
  const d=RO.data,c=d.config,host=$('#roHost');if(!host)return;
  const errs=d.checks.filter(i=>i.level==='error').length,hols=Object.keys(d.holidays).length;
  const opt=(v,l,cur)=>`<option value="${esc(v)}" ${v===cur?'selected':''}>${esc(l)}</option>`;
  const chips=c.shifts.map(s=>`<span class="ro-chip ${RO.hide[s.code]?'off':''}" data-ro-chip="${esc(s.code)}" style="background:${esc(s.color||'#ccc')}">${esc(s.code)} ${esc(s.name)} ${esc(s.start)}-${esc(s.end)}</span>`).join('')
    +`<span class="ro-chip ${RO.hide.OFF?'off':''}" data-ro-chip="OFF" style="background:#d9dee8">Off</span><span class="ro-chip ${RO.hide.LEAVE?'off':''}" data-ro-chip="LEAVE" style="background:#F0B4B0">Leave</span>`;
  host.innerHTML=`<div class="toolbar">
      <select data-ro="pick" aria-label="Roster">${d.rosters.map(r=>opt(r.id,`${r.name} (${r.team})${r.pending?', '+r.pending+' pending':''}`,RO.id)).join('')}</select>
      <button class="btn" data-act="ro-nav" data-n="-1" aria-label="Previous">&#8592;</button><button class="btn" data-act="ro-nav" data-n="0">Today</button><button class="btn" data-act="ro-nav" data-n="1" aria-label="Next">&#8594;</button>
      <select data-ro="mode" aria-label="Period">${opt('week','Week',RO.mode)}${opt('month','Month',RO.mode)}</select>
      <select data-ro="layout" aria-label="Layout">${opt('by_shift','By shift',RO.layout)}${opt('by_person','By person',RO.layout)}</select>
      <div class="grow"><input type="search" data-ro="q" value="${esc(RO.q)}" placeholder="Search this roster: people, dates, shifts, holidays" aria-label="Search this roster (use the top search box for everything else)"></div></div>
    <div class="toolbar"><button class="btn primary" data-act="ro-leave">Request leave</button><button class="btn" data-act="ro-swap">Swap or cover</button>
      <span class="grow"></span><button class="btn" data-act="ro-export" data-f="xlsx">Excel</button><button class="btn" data-act="ro-export" data-f="html">Report</button>
      <button class="btn" data-act="ro-holidays">Holidays</button><button class="btn" data-act="ro-settings">Settings</button><button class="btn" data-act="ro-template">Template</button><button class="btn" data-act="ro-import">Import sheet</button><button class="btn" data-act="ro-clone">Clone</button><button class="btn" data-act="ro-new">New roster</button><button class="btn ghost" data-act="ro-deleted">Deleted rosters</button></div>
    <p class="st" style="margin:0 0 10px">${esc(d.start)} to ${esc(d.end)} &middot; ${esc(c.team)} team &middot; <b style="color:${errs?'var(--p1)':'var(--ok)'}">${errs?errs+' problem(s) to fix':'every shift is staffed and every rule is met'}</b>${hols?` &middot; ${hols} public holiday(s)`:''} &middot; version ${d.history.length+1}</p>
    <div class="ro-legend">${chips}</div>
    <div class="tablewrap" id="roTable">${RO.layout==='by_shift'?roByShift(d):roByPerson(d)}</div>
    <div class="cols" style="margin-top:16px">${roChecks(d)}${roRequests(d)}${roHistory(d)}</div>`;
  roFilter();
}
function roByShift(d){
  const sh=d.config.shifts;
  const head=`<thead><tr><th>Day</th>${sh.map(s=>`<th>${esc(s.name)} <span class="ro-n">${esc(s.start)}-${esc(s.end)}</span></th>`).join('')}<th>Week off</th><th>Leave</th><th>On-call</th></tr></thead>`;
  const rows=d.days.map(day=>{const we=day.weekday==='Sat'||day.weekday==='Sun';
    if(day.has_data===false)return `<tr class="ro-nodata"><th>${esc(day.weekday)} ${esc(day.date)}${day.holiday?`<span class="ro-hb">${esc(day.holiday)}</span>`:''}</th><td colspan="${sh.length+3}" class="st">No roster for this day${d.can.edit?': import a sheet for it, or set people with the by-person view':''}</td></tr>`;
    const cells=sh.map(s=>{const p=day.shifts[s.code]||[],need=d.need[day.date]?.[s.code],short=need!=null&&p.length<need;
      const ed=d.can.edit&&need!=null?` data-act="ro-shiftcell" data-d="${esc(day.date)}" data-code="${esc(s.code)}" title="Select to choose who works ${esc(s.name)} on ${esc(day.date)}"`:'';
      return `<td data-c="${esc(s.code)}" class="${short?'ro-short':''}${ed?' ro-edit':''}"${ed} style="background:${esc(s.color||'#888888')}22">${esc(p.join(', '))}<span class="ro-n">(${p.length}${need?'/'+need:''})</span>${s.log_shift&&need!=null?` <button class="linkish ro-log" data-act="ro-log" data-d="${esc(day.date)}" data-s="${esc(s.log_shift)}" title="Open this shift in the shift log">log</button>`:''}</td>`;}).join('');
    return `<tr class="${day.holiday?'ro-hol':we?'ro-we':''}"><th>${esc(day.weekday)} ${esc(day.date)}${day.holiday?`<span class="ro-hb">${esc(day.holiday)}</span>`:''}</th>${cells}<td data-c="OFF">${esc(day.off.join(', '))}</td><td data-c="LEAVE">${esc(day.leave.join(', '))}</td><td>${esc(Object.entries(day.oncall).map(([k,v])=>k+': '+v).join(', '))}</td></tr>`;}).join('');
  return `<table class="grid">${head}<tbody>${rows}</tbody></table>`;
}
function roByPerson(d){
  const edit=d.can.edit;
  const head=`<thead><tr><th>Member</th>${d.days.map(day=>`<th class="ro-cell${day.holiday?' ro-holh':''}" ${day.holiday?`title="${esc(day.holiday)}"`:''}>${esc(day.weekday)}<br>${esc(day.date.slice(8))}/${esc(day.date.slice(5,7))}</th>`).join('')}<th>Shifts</th></tr></thead>`;
  const rows=d.grid.map(m=>{let n=0;const cells=m.cells.map((x,i)=>{const code=x.code||'';if(d.config.shifts.some(s=>s.code===code))n++;
      return `<td class="ro-cell${edit?' ro-edit':''}" data-c="${esc(code)}" title="${esc(x.source||'')}" style="background:${esc(roColor(code)||'transparent')}" ${edit?`data-act="ro-cell" data-m="${esc(m.id)}" data-d="${esc(d.days[i].date)}" data-code="${esc(code)}"`:''}>${esc(code)}</td>`;}).join('');
    return `<tr><th>${esc(m.name)}${m.id===d.me?' <span class="chip">you</span>':''}<span class="ro-n">${esc(m.role)}</span></th>${cells}<td><b>${n}</b></td></tr>`;}).join('');
  return `<table class="grid">${head}<tbody>${rows||`<tr><td colspan="${d.days.length+2}"><div class="empty">No members yet. A Lead adds them in Settings.</div></td></tr>`}</tbody></table>`;
}
function roChecks(d){
  const g=d.checks_grouped||[];
  const li=g.slice(0,60).map(i=>`<li><span class="${i.level==='error'?'badge pri-P1':'badge pri-P3'}">${esc(i.level)}</span><div class="t">${esc(i.message)}<small>${esc(i.dates.join(', '))}${i.count>1?` (${i.count} days)`:''} &middot; ${esc(i.rule)}</small></div></li>`).join('');
  return `<section class="panel"><h3>Checks <span class="n">${d.checks.length||''}</span></h3><ul class="list">${li||'<li class="st">No problems found.</li>'}</ul></section>`;
}
function roRequests(d){
  const edit=d.can.edit,me=nk(d.user||S.me.username),admin=d.can.admin;
  const label={awaiting_partner:'waiting for partner',pending:'waiting for approver'};
  const li=d.requests.slice(0,40).map(x=>{const k=x.kind==='leave'?'leave':'swap',mine=nk(x.requested_by)===me;
    const what=x.kind==='leave'?`Leave: ${x.who}, ${x.start}${x.end!==x.start?' to '+x.end:''} (${x.type})`:`${x.kind==='cover'?'Cover':'Swap'}: ${x.who}, ${x.date_a}${x.date_b&&x.date_b!==x.date_a?' / '+x.date_b:''}`;
    const warn=(x.warnings||[]).length?`<small style="color:var(--care)">${esc(x.warnings.join(' '))}</small>`:'';
    const who=x.status==='awaiting_partner'?`${x.partner} has to accept first`:x.status==='pending'?(x.approver?`Assigned to ${x.approver}`:'Any Lead can approve'):'';
    let btns='';
    if(x.status==='awaiting_partner'&&nk(x.partner_user)===me)btns+=`<button class="btn xs primary" data-act="ro-consent" data-id="${esc(x.id)}" data-ok="1">Accept</button><button class="btn xs" data-act="ro-consent" data-id="${esc(x.id)}" data-ok="">Decline</button>`;
    if(x.status==='pending'&&edit&&(!x.approver||nk(x.approver)===me||admin))btns+=`<button class="btn xs primary" data-act="ro-decide" data-k="${k}" data-id="${esc(x.id)}" data-ok="1">Approve</button><button class="btn xs" data-act="ro-decide" data-k="${k}" data-id="${esc(x.id)}" data-ok="">Reject</button>`;
    if(['awaiting_partner','pending'].includes(x.status)&&(edit||mine))btns+=`<button class="btn xs ghost" data-act="ro-cancel" data-k="${k}" data-id="${esc(x.id)}">Cancel</button>`;
    const cls=x.status==='approved'?'pri-P4':['pending','awaiting_partner'].includes(x.status)?'pri-P3':'pri-P2';
    return `<li><span class="badge ${cls}">${esc(label[x.status]||x.status)}</span><div class="t">${esc(what)}<small>by ${esc(x.requested_by||'')}${x.note?': '+esc(x.note):''}${who?' &middot; '+esc(who):''}</small>${warn}</div><span class="sp"></span>${btns}</li>`;}).join('');
  const open=d.requests.filter(x=>['pending','awaiting_partner'].includes(x.status)).length;
  return `<section class="panel"><h3>Leave, swaps and cover <span class="n">${open||''}</span></h3><ul class="list">${li||'<li class="st">No requests yet.</li>'}</ul></section>`;
}
function roHistory(d){
  const li=[...d.history].reverse().map(h=>`<li><div class="t">${esc(h.action)}<small>${esc(fmtTime(h.at))} by ${esc(h.by)}</small></div><span class="sp"></span><button class="btn xs" data-act="ro-restore" data-i="${h.index}">Restore</button></li>`).join('');
  return `<section class="panel"><h3>Versions <span class="n">${d.history.length||''}</span></h3><ul class="list">${li||'<li class="st">No earlier versions yet.</li>'}</ul></section>`;
}
function roFilter(){
  const q=RO.q.trim().toLowerCase();
  document.querySelectorAll('#roTable tbody tr').forEach(tr=>{tr.style.display=!q||tr.textContent.toLowerCase().includes(q)?'':'none';});
  document.querySelectorAll('#roTable [data-c]').forEach(td=>td.classList.toggle('ro-dim',!!RO.hide[td.dataset.c]));
}
document.addEventListener('change',e=>{const k=e.target.dataset?.ro;if(!k||k==='q')return;
  if(k==='pick'){RO.id=e.target.value;RO.layout='';}else if(k==='mode')RO.mode=e.target.value;else if(k==='layout'){RO.layout=e.target.value;return roDraw();}
  renderRoster($('#view'));});
document.addEventListener('input',e=>{if(e.target.dataset?.ro==='q'){RO.q=e.target.value;roFilter();}});
document.addEventListener('click',e=>{const ch=e.target.closest('[data-ro-chip]');if(!ch)return;const k=ch.dataset.roChip;RO.hide[k]=!RO.hide[k];ch.classList.toggle('off',RO.hide[k]);roFilter();});
async function roReload(msg){if(msg)toast(msg,'success');await renderRoster($('#view'));}
function roMembers(onlyMe){const d=RO.data;return d.grid.filter(m=>!onlyMe||m.id===d.me).map(m=>m.name);}
function roIdOf(name){const m=RO.data.grid.find(x=>x.name===name);if(!m)throw new Error('Choose a person from the list.');return m.id;}
function roApproverField(d,mode){
  const list=d.config.approvers||[];
  if(d.config.approvals[mode]!=='lead'||d.can.edit||!list.length)return '';
  return fieldHtml({key:'approver',label:'Send to approver',type:'select',options:list,value:list[0],required:true});
}
function roLeaveDialog(){
  const d=RO.data,lead=d.can.edit;
  if(!lead&&!d.me)return toast('You are not on this roster yet. A Lead can add you in Settings with your user name ('+S.me.username+').','error');
  const me=d.grid.find(m=>m.id===d.me)?.name||'';
  openModal({title:'Request leave',save:'Send request',body:`<div class="form">${fieldHtml({key:'who',label:'Person',type:'select',options:roMembers(!lead),value:me,required:true})}
    ${fieldHtml({key:'type',label:'Type',type:'select',options:d.config.leave_types||['Annual'],value:(d.config.leave_types||['Annual'])[0]})}
    ${fieldHtml({key:'start',label:'From',type:'date',value:today(),required:true})}${fieldHtml({key:'end',label:'To',type:'date',value:today(),required:true})}
    ${roApproverField(d,'leave')}${fieldHtml({key:'note',label:'Note',type:'textarea'})}<p class="hint full">${d.config.approvals.leave==='lead'&&!lead?((d.config.approvers||[]).length?'The approver you choose decides.':'Any Lead can approve leave for this roster.'):'This is recorded straight away.'}</p></div>`,
    onSave:async ov=>{const v=readFields(ov);const r=await api('POST',`/api/roster/${encodeURIComponent(RO.id)}/leave`,{member_id:roIdOf(v.who),start:v.start,end:v.end,type:v.type,note:v.note,approver:v.approver||null});
      await roReload(r.request.status==='pending'?`Leave requested: waiting for ${r.request.approver||'a Lead'}`:'Leave recorded');}});
}
function roSwapDialog(){
  const d=RO.data,lead=d.can.edit;
  if(!lead&&!d.me)return toast('You are not on this roster yet. A Lead can add you in Settings with your user name ('+S.me.username+').','error');
  const me=d.grid.find(m=>m.id===d.me)?.name||'';
  openModal({title:'Swap or cover a shift',save:'Send request',body:`<div class="form">${fieldHtml({key:'kind',label:'Kind',type:'select',options:['swap','cover'],value:'swap',hint:'Swap: you exchange shifts. Cover: the other person takes your shift and you are off.'})}
    ${fieldHtml({key:'a',label:'Giving up the shift',type:'select',options:roMembers(!lead),value:me,required:true})}${fieldHtml({key:'date_a',label:'On',type:'date',value:today(),required:true})}
    ${fieldHtml({key:'b',label:'Other person',type:'select',options:roMembers(false).filter(n=>n!==me),required:true})}${fieldHtml({key:'date_b',label:'Their shift on (swap on two days only)',type:'date'})}
    ${roApproverField(d,'swap')}${fieldHtml({key:'note',label:'Note',type:'textarea'})}<p class="hint full">The other person is asked to accept first${d.config.approvals.swap==='lead'&&!lead?', then the approver decides':''}. The roster checks rest hours and other rules, and shows anything this would break.</p></div>`,
    onSave:async ov=>{const v=readFields(ov);const r=await api('POST',`/api/roster/${encodeURIComponent(RO.id)}/swap`,{kind:v.kind,member_a:roIdOf(v.a),date_a:v.date_a,member_b:roIdOf(v.b),date_b:v.kind==='swap'?(v.date_b||null):null,note:v.note,approver:v.approver||null});
      if((r.request.warnings||[]).length)toast('Check: '+r.request.warnings.join(' '),'error');
      const st=r.request.status;await roReload(st==='awaiting_partner'?`Request sent: waiting for ${r.request.partner} to accept`:st==='pending'?`Request sent: waiting for ${r.request.approver||'a Lead'}`:'Recorded');}});
}
function roCellDialog(ds){
  const d=RO.data,m=d.grid.find(x=>x.id===ds.m),codes=[...d.config.shifts.map(s=>s.code),'OFF'];
  openModal({title:`${m?m.name:''}: ${ds.d}`,save:'Save',body:`<div class="form">${fieldHtml({key:'code',label:'Shift',type:'select',options:[...codes,'Back to rotation'],value:ds.code&&codes.includes(ds.code)?ds.code:'',hint:'This changes one day only. Leave and swaps stay as they are.'})}</div>`,
    onSave:async ov=>{const v=readFields(ov);await api('POST',`/api/roster/${encodeURIComponent(RO.id)}/override`,{member_id:ds.m,date:ds.d,code:v.code==='Back to rotation'||!v.code?null:v.code});await roReload('Shift changed');}});
}
function roHolidayDialog(){
  const c=RO.data.config,h=c.holidays||{region:'',subdivision:'',custom:[],ignore:[]};
  const regions=Object.entries(RO.data.regions);
  openModal({title:'Public holidays',save:'Save',size:'wide',body:`<div class="form">
    <div class="field"><label>Region</label><select data-field="region"><option value="">None (company list only)</option>${regions.map(([k,v])=>`<option value="${esc(k)}" ${k===h.region?'selected':''}>${esc(v)}</option>`).join('')}</select></div>
    ${fieldHtml({key:'subdivision',label:'State (optional)',value:h.subdivision||'',list:'dl-subdiv',placeholder:'For example KA, MH, TN, or CA, NY',hint:'Adds that state\'s holidays to the national list.'})}<datalist id="dl-subdiv"></datalist>
    ${fieldHtml({key:'custom',label:'Company holidays: one per line, DATE | Name',type:'mono',value:(h.custom||[]).map(x=>x.date+' | '+x.name).join('\n')})}
    ${fieldHtml({key:'ignore',label:'Region holidays your company works: one date per line',type:'mono',value:(h.ignore||[]).join('\n')})}
    <p class="hint full">${RO.data.holidays_available?'Region dates come from the open-source holidays list. Festival dates can differ from your company calendar: add or ignore days above.':'The region list needs the holidays package on the server (pip install holidays). The company list works without it.'}</p>
    <div class="full"><b>In view now:</b> ${Object.entries(RO.data.holidays).map(([d,n])=>esc(d+' '+n)).join(', ')||'none'}</div></div>`,
    onMount:ov=>{const sel=ov.querySelector('[data-field=region]');const fill=async()=>{try{const r=await api('GET','/api/roster/subdivisions?region='+encodeURIComponent(sel.value));
      ov.querySelector('#dl-subdiv').innerHTML=r.subdivisions.map(x=>`<option value="${esc(x)}">`).join('');}catch(_){}};sel.addEventListener('change',fill);fill();},
    onSave:async ov=>{const v=readFields(ov);
      const custom=roLines(v.custom).map(l=>{const [date,...name]=roCells(l);if(!/^[0-9]{4}-[0-9]{2}-[0-9]{2}$/.test(date)||!name.join(' ').trim())throw new Error('Company holidays: use one line per day as 2026-11-10 | Name.');return {date,name:name.join(' ').trim()};});
      const ignore=roLines(v.ignore);
      await api('PUT',`/api/roster/${encodeURIComponent(RO.id)}`,{changes:{holidays:{region:v.region,subdivision:(v.subdivision||'').toUpperCase(),custom,ignore}}});await roReload('Holidays saved');}});
}
function roSettingsDialog(){
  const c=RO.data.config,r=c.rules||{},oc=c.oncall||{},names=Object.fromEntries((c.members||[]).map(m=>[m.id,m.name]));
  const shifts=c.shifts.map(s=>[s.code,s.name,s.start,s.end,roDaysText(s.days),s.min_staff,s.holiday_min_staff??'',s.color||'',s.log_shift||''].join(' | ')).join('\n');
  const members=(c.members||[]).map(m=>[m.name,m.user||'',m.role||'Engineer',m.offset||0].join(' | ')).join('\n');
  const order=Object.entries(oc.order||{}).map(([k,v])=>k+': '+v.map(id=>names[id]||id).join(', ')).join('\n');
  const esc2=(oc.escalation||[]).map(e=>[e.level,e.name,e.contact].join(' | ')).join('\n');
  openModal({title:'Roster settings',save:'Save settings',size:'wide',footExtra:'<button class="btn danger" data-ro-del>Delete roster</button><span class="sp"></span>',body:`<div class="form">
    ${fieldHtml({key:'name',label:'Roster name',value:c.name,required:true})}${fieldHtml({key:'team',label:'Team',value:c.team,required:true})}
    ${fieldHtml({key:'cycle',label:'Cycle',type:'select',options:['daily','weekly','fortnightly','monthly','custom'],value:c.cycle.type})}${fieldHtml({key:'length',label:'Custom cycle length (days)',value:String(c.cycle.length_days||'')})}
    ${fieldHtml({key:'rotation',label:'Rotation: shift codes and OFF, in order',value:c.rotation.sequence.join(' '),full:true,hint:'Each person walks this sequence, starting at their offset. Its length must match the cycle (any length for monthly).'})}
    ${fieldHtml({key:'shifts',label:'Shifts: CODE | Name | Start | End | Days | Min staff | Holiday min | Colour | Shift log',type:'mono',value:shifts})}
    ${fieldHtml({key:'members',label:'Members: Name | user name | Role | Offset (days into the rotation)',type:'mono',value:members})}
    ${fieldHtml({key:'rest',label:'Minimum rest between shifts (hours)',value:String(r.min_rest_hours??8)})}${fieldHtml({key:'maxrun',label:'Maximum days in a row (empty for no limit)',value:String(r.max_consecutive_days??'')})}
    ${fieldHtml({key:'forbid',label:'Not allowed back to back (for example N>M, N>D)',value:(r.forbidden_sequences||[]).map(p=>p.join('>')).join(', ')})}
    ${fieldHtml({key:'leave',label:'Leave approval',type:'select',options:['lead','self'],value:c.approvals.leave})}${fieldHtml({key:'swap',label:'Swap and cover approval',type:'select',options:['lead','self'],value:c.approvals.swap})}
    ${fieldHtml({key:'approvers',label:'Approvers: user names, comma separated (empty: any Lead)',value:(c.approvers||[]).join(', '),full:true,hint:'Requests go to the approver the person chooses (the first one by default). Only that approver, or an Admin, can decide.'})}
    ${fieldHtml({key:'types',label:'Leave types (comma separated)',value:(c.leave_types||[]).join(', ')})}${fieldHtml({key:'layout',label:'Opens in layout',type:'select',options:['by_shift','by_person'],value:c.layout||'by_person'})}
    ${fieldHtml({key:'shared',label:'Members',type:'checkbox',text:'Allow people to be in other rosters too',value:!!r.allow_shared_members})}${fieldHtml({key:'skip',label:'Inactive days',type:'checkbox',text:'A shift that does not run that weekday becomes a day off',value:r.skip_inactive_days!==false})}
    ${fieldHtml({key:'oc',label:'On-call',type:'checkbox',text:'Run an on-call rota',value:!!oc.enabled})}${fieldHtml({key:'ocdays',label:'On-call changes every (days)',value:String(oc.rotation_days||7)})}
    ${fieldHtml({key:'ocroles',label:'On-call roles (comma separated)',value:(oc.roles||['Primary','Secondary']).join(', ')})}${fieldHtml({key:'ocstart',label:'On-call rota starts',type:'date',value:oc.start_date||c.start_date})}
    ${fieldHtml({key:'ocorder',label:'On-call order: Role: Name, Name, ...',type:'mono',value:order})}${fieldHtml({key:'ocesc',label:'Escalation: Level | Name | Contact',type:'mono',value:esc2})}</div>`,
    onMount:(ov,m)=>{ov.querySelector('[data-ro-del]').addEventListener('click',async()=>{if(!await confirmBox(`Delete the roster "${c.name}"? It disappears from the roster page. A Lead can bring it back later from Deleted rosters, exactly as it is now.`))return;
      try{await api('DELETE',`/api/roster/${encodeURIComponent(RO.id)}`);closeModal(m);RO.id='';localStorage.removeItem('msp.roster');await roReload('Roster deleted');}catch(e){toast(e.message,'error');}});},
    onSave:async ov=>{const v=readFields(ov);
      const shiftsOut=roLines(v.shifts).map(l=>{const x=roCells(l);if(x.length<6)throw new Error('Each shift line needs at least: CODE | Name | Start | End | Days | Min staff');
        const old=c.shifts.find(s=>s.code===x[0].toUpperCase())||{};return {...old,code:x[0].toUpperCase(),name:x[1],start:x[2],end:x[3],days:roParseDays(x[4]),min_staff:parseInt(x[5]||'0',10),
          holiday_min_staff:x[6]===''||x[6]===undefined?null:parseInt(x[6],10),color:x[7]||old.color||'',log_shift:x[8]||old.log_shift||x[1]};});
      const byName=Object.fromEntries((c.members||[]).map(m=>[m.name.toLowerCase(),m]));
      const membersOut=roLines(v.members).map(l=>{const x=roCells(l);if(!x[0])throw new Error('Each member line needs a name.');const old=byName[x[0].toLowerCase()];
        return {id:old?old.id:roNewId(),name:x[0],user:x[1]||'',role:x[2]||'Engineer',offset:parseInt(x[3]||'0',10)||0};});
      const ids=Object.fromEntries(membersOut.map(m=>[m.name.toLowerCase(),m.id]));
      const orderOut={};roLines(v.ocorder).forEach(l=>{const i=l.indexOf(':');if(i<0)throw new Error('On-call order lines look like  Primary: Name, Name');
        orderOut[l.slice(0,i).trim()]=l.slice(i+1).split(',').map(x=>x.trim()).filter(Boolean).map(n=>{const id=ids[n.toLowerCase()];if(!id)throw new Error('On-call order: '+n+' is not a member.');return id;});});
      const escOut=roLines(v.ocesc).map(l=>{const x=roCells(l);return {level:parseInt(x[0]||'1',10)||1,name:x[1]||'',contact:x[2]||''};});
      const forbid=String(v.forbid||'').split(',').map(x=>x.trim()).filter(Boolean).map(x=>x.split('>').map(y=>y.trim().toUpperCase()));
      const changes={name:v.name,team:v.team,layout:v.layout,cycle:{type:v.cycle,length_days:v.cycle==='custom'?parseInt(v.length||'0',10):undefined},
        rotation:{sequence:String(v.rotation||'').split(/[\s,]+/).filter(Boolean).map(x=>x.toUpperCase())},shifts:shiftsOut,members:membersOut,
        rules:{...(c.rules||{}),min_rest_hours:parseFloat(v.rest||'0'),max_consecutive_days:v.maxrun?parseInt(v.maxrun,10):null,forbidden_sequences:forbid,allow_shared_members:v.shared,skip_inactive_days:v.skip},
        approvals:{leave:v.leave,swap:v.swap},approvers:String(v.approvers||'').split(',').map(x=>x.trim().toLowerCase()).filter(Boolean),leave_types:String(v.types||'').split(',').map(x=>x.trim()).filter(Boolean),
        oncall:{...oc,enabled:v.oc,rotation_days:parseInt(v.ocdays||'7',10)||7,roles:String(v.ocroles||'').split(',').map(x=>x.trim()).filter(Boolean),start_date:v.ocstart||c.start_date,order:orderOut,escalation:escOut}};
      await api('PUT',`/api/roster/${encodeURIComponent(RO.id)}`,{changes});RO.layout='';await roReload('Roster saved');}});
}
function roNewDialog(){
  const meta=RO.meta||RO.data||{};
  const load=meta.templates?Promise.resolve(meta):api('GET','/api/roster');
  load.then(m=>{RO.meta=m;
    openModal({title:'New roster',save:'Create roster',size:'wide',body:`<div class="form">${fieldHtml({key:'name',label:'Roster name',required:true,placeholder:'For example Unix SA follow-the-sun'})}${fieldHtml({key:'team',label:'Team',required:true,placeholder:'Unix, Windows, Database, Boomi, NOC'})}
      <div class="field full"><label>Starting template</label><select data-field="template">${Object.entries(m.templates).map(([k,v])=>`<option value="${esc(k)}">${esc(k==='three_shift'?'Three-shift, by shift':'Standard, by person')}: ${esc(v)}</option>`).join('')}</select></div>
      ${fieldHtml({key:'start',label:'Rota starts',type:'date',value:today()})}<div class="field"><label>Public holidays</label><select data-field="region"><option value="">None for now</option>${Object.entries(m.regions).map(([k,v])=>`<option value="${esc(k)}">${esc(v)}</option>`).join('')}</select></div>
      ${fieldHtml({key:'members',label:'Members: Name | user name | Role | Offset (optional; you can add them later)',type:'mono'})}
      <p class="hint full">Everything can be changed afterwards in Settings: shifts and times, rotation, rules, approvals, on-call and holidays.</p></div>`,
      onSave:async ov=>{const v=readFields(ov);
        const members=roLines(v.members).map(l=>{const x=roCells(l);return {name:x[0],user:x[1]||'',role:x[2]||'Engineer',offset:parseInt(x[3]||'0',10)||0};});
        const r=await api('POST','/api/roster',{name:v.name,team:v.team,template:v.template,start_date:v.start,holidays:{region:v.region},members});
        RO.id=r.roster.id;RO.layout='';await roReload('Roster created');}});}).catch(e=>toast(e.message,'error'));
}
function roImportDialog(){
  openModal({title:'Import a roster sheet',save:'Import',body:`<div class="form"><div class="field full"><label>File (Excel, CSV or tab-separated text)</label><input type="file" data-ro-file accept=".xlsx,.xlsm,.csv,.txt,.tsv"></div>
    ${fieldHtml({key:'name',label:'Roster name',required:true})}${fieldHtml({key:'team',label:'Team',required:true})}
    <p class="hint full">Either layout works: one row per day with a column per shift (plus Week off and Leave), or one row per person with a column per date. Shift times are read from the header or a Shifts sheet.</p>
    <div class="full">Need a blank sheet? <button class="btn xs" data-ro-tpl="by_shift|xlsx">By shift (Excel)</button> <button class="btn xs" data-ro-tpl="by_shift|csv">By shift (CSV)</button> <button class="btn xs" data-ro-tpl="by_person|xlsx">By person (Excel)</button></div></div>`,
    onMount:ov=>ov.querySelectorAll('[data-ro-tpl]').forEach(b=>b.addEventListener('click',()=>{const [l,f]=b.dataset.roTpl.split('|');download(`/api/roster/template?layout=${l}&format=${f}&month=${today().slice(0,7)}`);})),
    onSave:async ov=>{const f=ov.querySelector('[data-ro-file]').files[0];if(!f)throw new Error('Choose a file first.');const v=readFields(ov);
      const data=await new Promise((res,rej)=>{const fr=new FileReader();fr.onload=()=>res(String(fr.result).split(',')[1]||'');fr.onerror=()=>rej(new Error('The file could not be read.'));fr.readAsDataURL(f);});
      const r=await api('POST','/api/roster/import',{filename:f.name,data,name:v.name,team:v.team});
      RO.id=r.roster.id;RO.layout='';await roReload(`Imported ${r.roster.members} people`);
      if(r.warnings.length)openModal({title:'Things to check in the imported sheet',body:`<ul class="list">${r.warnings.map(w=>`<li>${esc(w)}</li>`).join('')}</ul><p class="hint">Imported shifts start with a minimum of 1 person. Set your real numbers in Settings.</p>`});}});
}
function roShiftCellDialog(ds){
  const d=RO.data,s=d.config.shifts.find(x=>x.code===ds.code),day=d.days.find(x=>x.date===ds.d);if(!s||!day)return;
  const on=new Set(day.shifts[s.code]||[]),need=d.need[ds.d]?.[s.code];
  const where=n=>{for(const [c,ppl] of Object.entries(day.shifts))if(ppl.includes(n))return c;return day.leave.includes(n)?'LEAVE':'OFF';};
  openModal({title:`${s.name} on ${day.weekday} ${ds.d}`,save:'Save',body:`<p class="st full" style="margin:0 0 8px">Tick who works ${esc(s.name)} (${esc(s.start)}-${esc(s.end)}) this day${need?`: ${need} needed`:''}. Anyone you untick gets the day off. Only this day changes.</p>
    <div class="ro-picks">${d.grid.map(m=>{const w=where(m.name);return `<label class="st"><input type="checkbox" data-m="${esc(m.id)}" data-was="${esc(w)}" ${on.has(m.name)?'checked':''}> ${esc(m.name)} <span class="ro-n">${w===s.code?'':'now '+esc(w)}</span></label>`;}).join('')}</div>`,
    onSave:async ov=>{const changes=[];ov.querySelectorAll('[data-m]').forEach(cb=>{const was=cb.dataset.was;
        if(cb.checked&&was!==s.code)changes.push({member_id:cb.dataset.m,date:ds.d,code:s.code});
        if(!cb.checked&&was===s.code)changes.push({member_id:cb.dataset.m,date:ds.d,code:'OFF'});});
      if(!changes.length)return;
      if(changes.some(c=>ov.querySelector(`[data-m="${c.member_id}"]`).dataset.was==='LEAVE')&&!await confirmBox('Someone you ticked is on approved leave that day. Put them on the shift anyway? Their leave stays recorded.','Continue',false))throw Object.assign(new Error(''),{shown:true});
      await api('POST',`/api/roster/${encodeURIComponent(RO.id)}/overrides`,{changes});await roReload(`${changes.length} change(s) saved`);}});
}
function roCloneDialog(){
  const c=RO.data.config,shared=!!(c.rules||{}).allow_shared_members;
  openModal({title:'Clone this roster',save:'Clone',body:`<div class="form">${fieldHtml({key:'name',label:'New roster name',value:c.name+' (copy)',required:true})}${fieldHtml({key:'team',label:'Team',value:c.team,required:true})}
    ${fieldHtml({key:'members',label:'Members',type:'checkbox',text:'Copy the members and on-call order too',value:false})}
    <p class="hint full">Copies the shifts, rotation, rules, approvals, leave types, holidays and on-call settings. Leave, swaps, one-day changes and versions are not copied.${shared?'':' Members are dedicated to one roster here, so leave them out, or allow shared members in Settings first.'}</p></div>`,
    onSave:async ov=>{const v=readFields(ov);const r=await api('POST',`/api/roster/${encodeURIComponent(RO.id)}/clone`,{name:v.name,team:v.team,members:v.members});
      RO.id=r.roster.id;RO.layout='';await roReload('Roster cloned: add its members in Settings');}});
}
async function roDeletedDialog(){
  let list;try{list=(await api('GET','/api/roster/deleted')).deleted;}catch(e){return toast(e.message,'error');}
  openModal({title:'Deleted rosters',body:list.length?`<ul class="list">${list.map(x=>`<li><div class="t">${esc(x.name)} (${esc(x.team)})<small>${x.members} member(s), deleted ${esc(fmtTime(x.deleted_at))} by ${esc(x.deleted_by)}</small></div><span class="sp"></span><button class="btn xs primary" data-ro-undel="${esc(x.id)}">Restore</button></li>`).join('')}</ul>`:'<p class="st">No deleted rosters.</p>',
    onMount:(ov,m)=>ov.querySelectorAll('[data-ro-undel]').forEach(b=>b.addEventListener('click',async()=>{
      try{const r=await api('POST',`/api/roster/${encodeURIComponent(b.dataset.roUndel)}/undelete`,{});closeModal(m);RO.id=r.roster.id;RO.layout='';await roReload('Roster restored');}
      catch(e){toast(e.message,'error');}}))});
}
async function fillMyShifts(){
  const box=$('#myShifts');if(!box)return;
  try{const d=await api('GET','/api/roster/mine?days=7',undefined,{bg:true});
    const todo=(d.todo||[]).map(t=>`<li><span class="badge pri-P3">${t.action==='consent'?'your OK':'approve'}</span><div class="t">${esc(t.what)}<small>${esc(t.roster)}, asked by ${esc(t.by)}${t.warnings.length?': '+esc(t.warnings.join(' ')):''}</small></div><span class="sp"></span>`
      +(t.action==='consent'?`<button class="btn xs primary" data-act="ro-consent" data-r="${esc(t.roster_id)}" data-id="${esc(t.id)}" data-ok="1">Accept</button><button class="btn xs" data-act="ro-consent" data-r="${esc(t.roster_id)}" data-id="${esc(t.id)}" data-ok="">Decline</button>`
        :`<button class="btn xs primary" data-act="ro-decide" data-r="${esc(t.roster_id)}" data-k="${esc(t.kind)}" data-id="${esc(t.id)}" data-ok="1">Approve</button><button class="btn xs" data-act="ro-decide" data-r="${esc(t.roster_id)}" data-k="${esc(t.kind)}" data-id="${esc(t.id)}" data-ok="">Reject</button>`)+'</li>').join('');
    box.querySelector('h3').innerHTML='Your shifts, next 7 days'+(todo?` <span class="n hot">${d.todo.length} waiting for you</span>`:'');
    box.querySelector('ul').innerHTML=todo+d.shifts.filter(x=>x.code!=='OFF').map(x=>`<li><span class="ro-chip" style="background:${x.code==='LEAVE'?'#F0B4B0':'#d9dee8'}">${esc(x.code)}</span><div class="t">${esc(x.weekday)} ${esc(x.date)}: ${esc(x.shift)}${x.start?' '+esc(x.start)+'-'+esc(x.end):''}<small>${esc(x.roster)}${x.oncall.length?', on-call '+esc(x.oncall.join(', ')):''}${x.holiday?', '+esc(x.holiday):''}</small></div>${x.shift_log&&x.code!=='LEAVE'?`<span class="sp"></span><button class="linkish" data-act="ro-log" data-d="${esc(x.date)}" data-s="${esc(x.shift_log.split('/').pop())}">Shift log</button>`:''}</li>`).join('')
      ||'<li class="st">No roster shifts in the next 7 days. A Lead adds you to a roster with your user name.</li>';
  }catch(e){box.remove();}
}
Object.assign(ACT,{
  'ro-nav':ds=>{const n=+ds.n;RO.start=n===0?today():RO.mode==='week'?roDate(RO.start,7*n):roMonth(RO.start,n);renderRoster($('#view'));},
  'ro-log':ds=>go(`shiftlog=${ds.d}/${ds.s}`),
  'ro-leave':()=>roLeaveDialog(),'ro-swap':()=>roSwapDialog(),'ro-cell':ds=>roCellDialog(ds),
  'ro-decide':async ds=>{try{await api('POST',`/api/roster/${encodeURIComponent(ds.r||RO.id)}/decide`,{kind:ds.k,id:ds.id,approve:!!ds.ok});
    if(S.view==='mywork'){toast(ds.ok?'Approved':'Rejected','success');fillMyShifts();}else await roReload(ds.ok?'Approved':'Rejected');}catch(e){toast(e.message,'error');}},
  'ro-consent':async ds=>{try{const r=await api('POST',`/api/roster/${encodeURIComponent(ds.r||RO.id)}/consent`,{id:ds.id,accept:!!ds.ok});
    const msg=!ds.ok?'Declined':r.request.status==='pending'?`Accepted: now waiting for ${r.request.approver||'a Lead'}`:'Accepted: the swap is done';
    if(S.view==='mywork'){toast(msg,'success');fillMyShifts();}else await roReload(msg);}catch(e){toast(e.message,'error');}},
  'ro-template':()=>openModal({title:'Download a blank roster template',body:`<p>Fill it in with Excel, then use Import sheet. Every day of the month you pick is already listed: fill in every day, because a day with nobody on it counts as having no roster yet. Two shaded example rows show the format.</p>
    <div class="form">${fieldHtml({key:'month',label:'Month',type:'month',value:today().slice(0,7)})}</div>
    <div class="toolbar"><button class="btn primary" data-ro-tpl="by_shift|xlsx">By shift (Excel)</button><button class="btn" data-ro-tpl="by_shift|csv">By shift (CSV)</button><button class="btn" data-ro-tpl="by_person|xlsx">By person (Excel)</button></div>
    <p class="hint">By shift: one row per day, names under each shift, plus Week off and Leave. By person: one row per person, a shift code under each date.</p>`,
    onMount:ov=>ov.querySelectorAll('[data-ro-tpl]').forEach(b=>b.addEventListener('click',()=>{const [l,f]=b.dataset.roTpl.split('|');const mo=(ov.querySelector('[data-field=month]')?.value||today().slice(0,7));download(`/api/roster/template?layout=${l}&format=${f}&month=${encodeURIComponent(mo)}`);}))}),
  'ro-cancel':async ds=>{if(!await confirmBox('Cancel this request?','Cancel request'))return;try{await api('POST',`/api/roster/${encodeURIComponent(RO.id)}/cancel`,{kind:ds.k,id:ds.id});await roReload('Cancelled');}catch(e){toast(e.message,'error');}},
  'ro-restore':async ds=>{if(!await confirmBox('Restore this earlier version of the roster? The current one is kept as a version too.','Restore',false))return;
    try{await api('POST',`/api/roster/${encodeURIComponent(RO.id)}/restore`,{index:+ds.i});await roReload('Version restored');}catch(e){toast(e.message,'error');}},
  'ro-export':ds=>download(`/api/roster/${encodeURIComponent(RO.id)}/export?format=${ds.f}&start=${RO.data.start}&end=${RO.data.end}&layout=${RO.layout}`),
  'ro-holidays':()=>roHolidayDialog(),'ro-settings':()=>roSettingsDialog(),'ro-new':()=>roNewDialog(),'ro-import':()=>roImportDialog(),
  'ro-shiftcell':ds=>roShiftCellDialog(ds),'ro-clone':()=>roCloneDialog(),'ro-deleted':()=>roDeletedDialog()});
/* ---------- built-in assistant (served by msp_ai.py) and the on-shift strip ---------- */
const ASK={log:[],busy:false,mode:'search',model:''};
const ASK_ICON='<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M12 3l1.8 4.6L18.5 9l-4.7 1.6L12 15l-1.8-4.4L5.5 9l4.7-1.4z"/><path d="M19 15l.9 2.1L22 18l-2.1.9L19 21l-.9-2.1L16 18l2.1-.9z"/></svg>';
const ASK_SUGG=['Who is on shift now?','Who is on call this week?','Open P1 incidents','What changes are planned this week?','Known issues','How do I fix ASM diskgroup full?'];
function askEl(){
  let el=$('#askPanel');if(el)return el;
  document.body.insertAdjacentHTML('beforeend',`<div class="ask-scrim" id="askScrim"></div>
  <aside class="ask" id="askPanel" role="dialog" aria-modal="false" aria-labelledby="askTitle">
    <header><div class="mark">${ASK_ICON}</div><div><h3 id="askTitle">Ask MSP Ops Center</h3><small>Answers from this tool's records, on this server only</small></div><span class="sp"></span>
      <span class="mode" id="askMode">Search</span><button class="iconbtn" id="askClose" aria-label="Close"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M6 6l12 12M18 6 6 18"/></svg></button></header>
    <div class="log" id="askLog" aria-live="polite"></div>
    <form id="askForm"><div class="box"><textarea id="askInput" rows="1" maxlength="500" placeholder="Ask about devices, fixes, shifts, changes, incidents…" aria-label="Your question"></textarea>
      <button class="send" type="submit" aria-label="Ask">${'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M5 12h14M13 6l6 6-6 6"/></svg>'}</button></div>
      <p class="fine">Answers only use records you can see in MSP Ops Center, and show their sources. Check anything important before acting on it.</p></form>
  </aside>`);
  el=$('#askPanel');
  $('#askClose').addEventListener('click',askClose);$('#askScrim').addEventListener('click',askClose);
  $('#askForm').addEventListener('submit',e=>{e.preventDefault();askSend($('#askInput').value);});
  const ta=$('#askInput');ta.addEventListener('keydown',e=>{if(e.key==='Enter'&&!e.shiftKey){e.preventDefault();askSend(ta.value);}});
  ta.addEventListener('input',()=>{ta.style.height='auto';ta.style.height=Math.min(ta.scrollHeight,140)+'px';});
  el.addEventListener('click',e=>{
    const s=e.target.closest('[data-ask-sugg]');if(s)return askSend(s.dataset.askSugg);
    const c=e.target.closest('[data-ask-copy]');if(c){navigator.clipboard?.writeText(c.parentElement.firstChild.textContent).then(()=>toast('Copied','success'));return;}
    const n=e.target.closest('[data-ask-src]');if(n){const card=n.closest('.a');const src=card?.querySelector(`.src[data-n="${n.dataset.askSrc}"]`);
      if(n.classList.contains('cite')&&src){card.querySelectorAll('.src.hl').forEach(x=>x.classList.remove('hl'));src.classList.add('hl');src.scrollIntoView({block:'nearest'});return;}
      askGo(n.dataset.link);}
  });
  askRender();return el;
}
function askOpen(q){
  askEl();$('#askPanel').classList.add('on');$('#askScrim').classList.add('on');
  api('GET','/api/ai/status',undefined,{bg:true}).then(st=>{ASK.mode=st.mode;ASK.model=st.model;const m=$('#askMode');
    m.textContent=st.mode==='model'?'Local model: '+(st.model||'not set'):st.mode==='off'?'Switched off':'Search';m.className='mode '+st.mode;}).catch(()=>{});
  setTimeout(()=>$('#askInput')?.focus(),60);if(q)askSend(q);
}
function askClose(){$('#askPanel')?.classList.remove('on');$('#askScrim')?.classList.remove('on');}
function askGo(link){
  if(!link)return;askClose();
  const i=link.indexOf('='),view=i<0?link:link.slice(0,i),p=i<0?'':link.slice(i+1);
  if(view==='plans'&&p)return focusTask(p);
  if(view==='runbook'&&p)return gotoRunbook(p);
  if(['Asset_Master','Incidents'].includes(view)&&p){const f=tf(view);f.q=p;f.customer='';if(view==='Incidents')f.open=false;return go(view);}
  go(link);
}
function askInline(text){
  return esc(text).replace(/\[(\d{1,2})\]/g,(m,n)=>`<button class="cite" type="button" data-ask-src="${n}" aria-label="Source ${n}">${n}</button>`);
}
function askAnswerHtml(a){
  const blocks=a.blocks.map(b=>b.type==='list'?`<ul>${b.items.map(i=>`<li>${askInline(i)}</li>`).join('')}</ul>`
    :b.type==='code'?`<pre>${esc(b.text)}<button class="btn xs" type="button" data-ask-copy>Copy</button></pre>`:`<p>${askInline(b.text)}</p>`).join('');
  const srcs=a.sources.length?`<div class="srcs"><span class="h">Sources</span>${a.sources.map(s=>`<button class="src" type="button" data-n="${s.n}" data-ask-src="${s.n}" data-link="${esc(s.link)}">
    <span class="num">${s.n}</span><span><span class="kind">${esc(s.label)}</span><span class="ttl">${esc(s.title)}</span>${s.where?`<span class="whr">${esc(s.where)}</span>`:''}</span></button>`).join('')}</div>`:'';
  const notes=(a.notes||[]).map(n=>`<div class="note">${esc(n)}</div>`).join('');
  const how=a.mode==='model'?`Written by the local model from ${a.sources.length} record(s)`:`Searched ${a.records} records`;
  return `<div class="a${a.found?'':' nf'}">${blocks}${srcs}${notes}<div class="meta">${how} &middot; ${(a.took_ms/1000).toFixed(a.took_ms<1000?2:1)} s</div></div>`;
}
function askRender(){
  const log=$('#askLog');if(!log)return;
  if(!ASK.log.length){log.innerHTML=`<div class="welcome"><h4>What do you need to know?</h4><p>Ask in plain words: who is on shift, how something was fixed before, what is planned, which incidents are open, or anything about a device.</p>
    <div class="sugg">${ASK_SUGG.map(q=>`<button type="button" data-ask-sugg="${esc(q)}">${esc(q)}</button>`).join('')}</div></div>`;return;}
  log.innerHTML=ASK.log.map(t=>`<div class="q">${esc(t.q)}</div>`+(t.a?askAnswerHtml(t.a):t.err?`<div class="a nf"><p>${esc(t.err)}</p></div>`
    :`<div class="a"><div class="think"><i></i><i></i><i></i><span>${ASK.mode==='model'?'Finding records and writing the answer…':'Searching the records…'}</span></div></div>`)).join('');
  log.scrollTop=log.scrollHeight;
}
async function askSend(q){
  q=String(q||'').trim();if(!q||ASK.busy)return;
  const ta=$('#askInput');if(ta){ta.value='';ta.style.height='auto';}
  const turn={q};ASK.log.push(turn);ASK.busy=true;$('#askForm .send').disabled=true;askRender();
  try{turn.a=(await api('POST','/api/ai/ask',{question:q})).answer;}catch(e){turn.err=e.message;}
  ASK.busy=false;const b=$('#askForm .send');if(b)b.disabled=false;ASK.log=ASK.log.slice(-30);askRender();
}
async function fillOverviewRoster(){
  try{const d=await api('GET','/api/roster/now',undefined,{bg:true});if(!d.roster)return;
    document.querySelectorAll('.sun .tile[data-s]').forEach(t=>{const ppl=(d.today||{})[t.dataset.s];if(!ppl)return;
      t.querySelector('.rost')?.remove();t.insertAdjacentHTML('beforeend',`<div class="rost">On the roster: <b>${esc(ppl.join(', ')||'nobody')}</b></div>`);});
  }catch(_){}
}
async function fillOnShift(){
  const el=$('#onShift');if(!el||!S.me)return;
  try{const d=await api('GET','/api/roster/now',undefined,{bg:true});const s=d.now?.shifts?.[0];
    if(!s){el.classList.remove('show');return;}
    const h=Math.floor(s.minutes_left/60),m=s.minutes_left%60,people=s.people.length?s.people.join(', '):'nobody assigned';
    el.style.setProperty('--sc',s.color||'var(--accent)');
    el.innerHTML=`<span class="pulse"></span><b>${esc(s.name)}</b><span class="who">${esc(people)}</span><span class="left">${h?h+' h ':''}${m} m left</span>`;
    el.title=`${d.roster.name}: ${s.name} ${s.start} to ${s.end}${Object.keys(d.now.oncall||{}).length?'. On call: '+Object.entries(d.now.oncall).map(([k,v])=>k+' '+v).join(', '):''}`;
    el.classList.add('show');
  }catch(_){el.classList.remove('show');}
}
setInterval(()=>{if(S.me&&!document.hidden)fillOnShift();},60000);
document.addEventListener('keydown',e=>{
  if((e.ctrlKey||e.metaKey)&&!e.altKey&&e.key.toLowerCase()==='j'&&S.me){e.preventDefault();$('#askPanel')?.classList.contains('on')?askClose():askOpen();}
  else if(e.key==='Escape'&&$('#askPanel')?.classList.contains('on')&&!S.modals.length)askClose();
});
/* ---------- teams, customers and projects (served by msp_org.py) ---------- */
const TM={data:null,people:null};
async function renderTeams(el){
  el.innerHTML=`<h2>Teams</h2><p class="sub">Who belongs to which team, which customers each team serves, and which services each person works on.</p><div id="tmHost"><p class="st">Loading…</p></div>`;
  try{TM.data=await api('GET','/api/org');}catch(e){$('#tmHost').innerHTML=`<div class="empty">${esc(e.message)}</div>`;return;}
  tmDraw();
}
function tmProjName(pid){for(const c of TM.data.customers){const p=c.projects.find(x=>x.id===pid);if(p)return `${c.name}: ${p.name}`;}return pid;}
function tmDraw(){
  const d=TM.data,host=$('#tmHost');if(!host)return;
  const admin=d.is_admin;
  const intro=admin?`<p class="st" style="margin:0 0 12px">As administrator you create teams and customers, name each team's leads, and choose which teams serve each customer. Team leads then manage their own members and service assignments. Administrators are not team members, so they do not see team data.</p>`
    :d.leads_any?`<p class="st" style="margin:0 0 12px">You lead the teams marked below: add or remove members, create accounts for new engineers, and assign people to services.</p>`:'';
  const tools=admin?`<div class="toolbar"><button class="btn primary" data-act="tm-team">New team</button></div>`:'';
  const teams=d.teams.map(t=>{
    const leads=t.leads.map(l=>`<span class="chip">${esc(l.name)}</span>`).join(' ')||'<span class="st">No lead yet</span>';
    const custs=t.customers.map(cid=>{const c=d.customers.find(x=>x.id===cid);return c?`<span class="chip">${esc(c.name)}${c.primary_team===t.id?' <b style="color:var(--accent)">primary</b>':''}</span>`:'';}).join(' ')||'<span class="st">No customers yet</span>';
    const rows=t.members.map(m=>`<tr><td><b>${esc(m.name)}</b>${t.leads.some(l=>l.username===m.username)?' <span class="chip">lead</span>':''}${m.disabled?' <span class="chip">disabled</span>':''}<div class="st">${esc(m.username)}</div></td>
      <td>${esc(m.role)}</td><td>${m.projects.map(p=>`<span class="chip">${esc(tmProjName(p))}</span>`).join(' ')||'<span class="st">None</span>'}</td>
      <td style="white-space:nowrap">${t.can_manage?`<button class="btn xs" data-act="tm-proj" data-t="${esc(t.id)}" data-u="${esc(m.username)}">Services</button>${t.leads.some(l=>l.username===m.username)?'':` <button class="btn xs ghost" data-act="tm-rm" data-t="${esc(t.id)}" data-u="${esc(m.username)}">Remove</button>`}`:''}</td></tr>`).join('');
    return `<section class="panel" style="margin-bottom:16px"><h3>${esc(t.name)} <span class="n">${t.members.length}</span>${t.you_lead?' <span class="chip">you lead</span>':''}</h3>
      ${t.description?`<p class="st" style="margin:0 0 8px">${esc(t.description)}</p>`:''}
      <div style="display:flex;gap:18px;flex-wrap:wrap;margin:0 0 10px"><div><div class="st">Leads</div>${leads}</div><div><div class="st">Customers</div>${custs}</div></div>
      ${t.can_manage?`<div class="toolbar"><button class="btn primary" data-act="tm-add" data-t="${esc(t.id)}">Add member</button><button class="btn" data-act="tm-edit" data-t="${esc(t.id)}">Edit team</button>${admin?`<button class="btn ghost" data-act="tm-del" data-t="${esc(t.id)}">Delete team</button>`:''}</div>`:''}
      <div class="tablewrap"><table class="grid"><thead><tr><th>Member</th><th>Role</th><th>Services</th><th></th></tr></thead><tbody>${rows||'<tr><td colspan="4"><div class="empty">No members yet.</div></td></tr>'}</tbody></table></div></section>`;}).join('');
  const custs=d.customers.map(c=>`<li><div class="t"><b>${esc(c.name)}</b>${c.code?` <span class="st">${esc(c.code)}</span>`:''}
      <div style="margin-top:4px">${c.teams.map(tid=>`<span class="chip">${esc(d.team_names[tid]||tid)}${c.primary_team===tid?' <b style="color:var(--accent)">primary</b>':''}</span>`).join(' ')||'<span class="st">No teams assigned</span>'}</div>
      <small>Services: ${c.projects.map(p=>esc(p.name)).join(', ')||'none yet'}</small></div></li>`).join('');
  host.innerHTML=intro+tools+`<div class="cols" style="grid-template-columns:minmax(0,2fr) minmax(280px,1fr);align-items:start">
    <div>${teams||`<div class="tablewrap"><div class="empty">${admin?'No teams yet. Create the first team, then name its lead.':'You are not in a team yet. A team lead or administrator can add you.'}</div></div>`}</div>
    <section class="panel"><h3>Customers <span class="n">${d.customers.length||''}</span></h3><ul class="list">${custs||'<li class="st">No customers yet.</li>'}</ul>
      <a class="btn xs" href="#customers" style="margin-top:8px">${admin?'Manage customers and services':'Open Customers and services'}</a></section></div>`;
}
async function renderCustomers(el){
  el.innerHTML=`<h2>Customers and services</h2><p class="sub">The customers you serve, the services you deliver to each (agreed in service level agreements), the teams that serve them, and the primary team that owns each customer's core CI data. ITIL 4: service relationship management.</p><div id="cuHost"><p class="st">Loading…</p></div>`;
  try{TM.data=await api('GET','/api/org');}catch(e){$('#cuHost').innerHTML=`<div class="empty">${esc(e.message)}</div>`;return;}
  const d=TM.data,admin=d.is_admin;
  const people={};d.teams.forEach(t=>t.members.forEach(m=>m.projects.forEach(p=>{(people[p]=people[p]||new Set()).add(m.name);})));
  const cards=d.customers.map(c=>`<section class="panel"><h3>${esc(c.name)}${c.code?` <span class="n">${esc(c.code)}</span>`:''}</h3>
    <div class="st" style="margin:0 0 4px">Teams serving this customer</div>
    <div style="margin:0 0 10px">${c.teams.map(tid=>`<span class="chip">${esc(d.team_names[tid]||tid)}${c.primary_team===tid?' <b style="color:var(--accent)">primary</b>':''}</span>`).join(' ')||'<span class="st">None yet</span>'}</div>
    <div class="st" style="margin:0 0 4px">Services</div>
    <ul class="list">${c.projects.map(p=>`<li><div class="t">${esc(p.name)}<small>${people[p.id]?`${people[p.id].size} people: ${esc([...people[p.id]].slice(0,4).join(', '))}${people[p.id].size>4?'…':''}`:'Nobody assigned yet'}</small></div>
      ${admin?`<span class="sp"></span><button class="btn xs ghost" data-act="tm-pdel" data-c="${esc(c.id)}" data-p="${esc(p.id)}">Delete</button>`:''}</li>`).join('')||'<li class="st">No services yet.</li>'}</ul>
    ${admin?`<div class="toolbar" style="margin-top:10px"><button class="btn xs" data-act="tm-cedit" data-c="${esc(c.id)}">Edit</button><button class="btn xs" data-act="tm-padd" data-c="${esc(c.id)}">Add service</button><button class="btn xs ghost" data-act="tm-cdel" data-c="${esc(c.id)}">Delete</button></div>`:''}</section>`).join('');
  $('#cuHost').innerHTML=(admin?`<div class="toolbar"><button class="btn primary" data-act="tm-cust">New customer</button></div>`:'')
    +(cards?`<div class="cols" style="grid-template-columns:repeat(auto-fill,minmax(340px,1fr))">${cards}</div>`
      :`<div class="tablewrap"><div class="empty">${admin?'No customers yet. Create the first one, then choose the teams that serve it and add its services.':'None of your teams serves a customer yet.'}</div></div>`);
}
async function cuReload(msg){if(msg)toast(msg,'success');TM.people=null;if(S.view==='customers')await renderCustomers($('#view'));else await renderTeams($('#view'));}
async function tmPeople(){if(!TM.people)TM.people=(await api('GET','/api/org/people')).people;return TM.people;}
async function tmReload(msg){TM.people=null;if(msg)toast(msg,'success');await renderTeams($('#view'));}
function tmChecks(name,items,checked){return `<div class="full" style="display:grid;gap:6px">${items.map(([v,l])=>`<label style="display:flex;gap:8px;align-items:center"><input type="checkbox" data-${name}="${esc(v)}" ${checked.includes(v)?'checked':''}> ${esc(l)}</label>`).join('')||'<span class="st">Nothing to choose yet.</span>'}</div>`;}
function tmPicked(ov,name){return [...ov.querySelectorAll(`[data-${name}]:checked`)].map(x=>x.getAttribute('data-'+name));}
async function tmTeamDialog(tid){
  const t=tid?TM.data.teams.find(x=>x.id===tid):null,admin=TM.data.is_admin;
  let leads=[];if(admin){try{leads=(await api('GET','/api/admin/users')).users.filter(u=>u.role==='Lead'&&!u.disabled);}catch(e){return toast(e.message,'error');}}
  openModal({title:t?`Edit ${t.name}`:'New team',save:t?'Save':'Create team',body:`<div class="form">
    ${admin?fieldHtml({key:'name',label:'Team name',value:t?.name||'',required:true,placeholder:'For example DBA, Unix SA, Windows, Boomi, NOC'}):''}
    ${fieldHtml({key:'description',label:'Description',type:'textarea',value:t?.description||''})}
    ${admin?`<div class="field full"><label>Team leads (people with the Lead role)</label>${tmChecks('lead',leads.map(u=>[u.username,`${u.name} (${u.username})`]),(t?.leads||[]).map(l=>l.username))}
      <p class="hint">Give someone the Lead role in Admin, Users first. A person can lead several teams.</p></div>`:''}</div>`,
    onSave:async ov=>{const v=readFields(ov);const body=admin?{name:v.name,description:v.description,leads:tmPicked(ov,'lead')}:{description:v.description};
      if(t)await api('PUT',`/api/org/teams/${t.id}`,body);else await api('POST','/api/org/teams',body);await tmReload(t?'Team saved':'Team created');}});
}
async function tmAddDialog(tid){
  const t=TM.data.teams.find(x=>x.id===tid);let people;try{people=await tmPeople();}catch(e){return toast(e.message,'error');}
  const avail=people.filter(p=>!p.disabled&&!t.members.some(m=>m.username.toLowerCase()===p.username.toLowerCase()));
  openModal({title:`Add a member to ${t.name}`,save:'Add member',size:'wide',body:`<div class="form">
    ${fieldHtml({key:'mode',label:'Who',type:'select',options:['An existing user','A new user'],value:avail.length?'An existing user':'A new user'})}
    <div class="full" data-tm-existing>${fieldHtml({key:'username',label:'User',type:'select',options:avail.map(p=>`${p.name} (${p.username})`),value:''})}
      <p class="hint">People already in other teams can be added here too.</p></div>
    <div class="full form" data-tm-new hidden>${fieldHtml({key:'nu',label:'User name',placeholder:'For example priya.sharma'})}${fieldHtml({key:'nn',label:'Full name'})}
      ${fieldHtml({key:'ne',label:'Email (needed for Microsoft 365 sign-in)'})}${fieldHtml({key:'na',label:'Sign-in',type:'select',options:['local','entra','both'],value:'local'})}
      ${fieldHtml({key:'nr',label:'Role',type:'select',options:['Engineer','Viewer'],value:'Engineer',hint:'Only an administrator can grant Lead or Admin.'})}
      <p class="hint full">A one-time password is shown after saving for local sign-in. Give it to the person; they choose their own at first sign-in.</p></div></div>`,
    onMount:ov=>{const sel=ov.querySelector('[data-field=mode]');const sync=()=>{const nw=sel.value==='A new user';ov.querySelector('[data-tm-existing]').hidden=nw;ov.querySelector('[data-tm-new]').hidden=!nw;};sel.addEventListener('change',sync);sync();},
    onSave:async ov=>{const v=readFields(ov);let body;
      if(v.mode==='A new user'){if(!v.nu)throw new Error('Enter a user name.');body={create:{username:v.nu,name:v.nn||v.nu,email:v.ne,auth:v.na,role:v.nr}};}
      else{const p=avail.find(x=>`${x.name} (${x.username})`===v.username);if(!p)throw new Error('Choose a user.');body={username:p.username};}
      const r=await api('POST',`/api/org/teams/${t.id}/members`,body);await tmReload('Member added');
      if(r.temp_password)openModal({title:'One-time password',body:`<p>Give this password to <b>${esc(r.member.username)}</b> in person or by a secure channel. They must choose a new one at first sign-in.</p><pre class="mono" style="font-size:16px">${esc(r.temp_password)}</pre><p class="hint">It is not shown again.</p>`});}});
}
function tmProjDialog(tid,username){
  const t=TM.data.teams.find(x=>x.id===tid),m=t.members.find(x=>x.username===username);
  const items=t.customers.flatMap(cid=>{const c=TM.data.customers.find(x=>x.id===cid);return c?c.projects.filter(p=>p.active!==false).map(p=>[p.id,`${c.name}: ${p.name}`]):[];});
  openModal({title:`Services for ${m.name}`,save:'Save',body:`<p class="st">Services of the customers ${esc(t.name)} serves.</p>${tmChecks('prj',items,m.projects)}`,
    onSave:async ov=>{await api('PUT',`/api/org/teams/${t.id}/members/${encodeURIComponent(username)}`,{projects:tmPicked(ov,'prj')});await tmReload('Services saved');}});
}
function tmCustDialog(cid){
  const c=cid?TM.data.customers.find(x=>x.id===cid):null;const teams=TM.data.teams.map(t=>[t.id,t.name]);
  openModal({title:c?`Edit ${c.name}`:'New customer',save:c?'Save':'Create customer',body:`<div class="form">
    ${fieldHtml({key:'name',label:'Customer name',value:c?.name||'',required:true})}${fieldHtml({key:'code',label:'Short code (optional)',value:c?.code||'',placeholder:'For example ACME'})}
    ${c?`<div class="field full"><label>Teams that serve this customer</label>${tmChecks('team',teams,c.teams)}</div>
    <div class="field full"><label>Primary team (owns the core CI data, such as host name and IP)</label><select data-field="primary"><option value="">None yet</option>${teams.filter(([id])=>c.teams.includes(id)).map(([id,n])=>`<option value="${esc(id)}" ${c.primary_team===id?'selected':''}>${esc(n)}</option>`).join('')}</select>
      <p class="hint">Save the team list first if the team you want is not offered here yet.</p></div>`:'<p class="hint full">After creating it, choose its teams and primary team with Edit, and add its services.</p>'}</div>`,
    onSave:async ov=>{const v=readFields(ov);
      if(c){const chosen=tmPicked(ov,'team');await api('PUT',`/api/org/customers/${c.id}`,{name:v.name,code:v.code,teams:chosen,primary_team:chosen.includes(v.primary)?v.primary:null});}
      else await api('POST','/api/org/customers',{name:v.name,code:v.code});await cuReload(c?'Customer saved':'Customer created');}});
}
Object.assign(ACT,{
  'tm-team':()=>tmTeamDialog(),'tm-edit':ds=>tmTeamDialog(ds.t),'tm-add':ds=>tmAddDialog(ds.t),'tm-proj':ds=>tmProjDialog(ds.t,ds.u),
  'tm-cust':()=>tmCustDialog(),'tm-cedit':ds=>tmCustDialog(ds.c),
  'tm-rm':async ds=>{if(!await confirmBox(`Remove ${ds.u} from this team? Their service assignments in this team are removed too. Their account stays.`,'Remove'))return;
    try{await api('DELETE',`/api/org/teams/${ds.t}/members/${encodeURIComponent(ds.u)}`);await tmReload('Removed from the team');}catch(e){toast(e.message,'error');}},
  'tm-del':async ds=>{const t=TM.data.teams.find(x=>x.id===ds.t);if(!await confirmBox(`Delete the team "${t.name}"? Members keep their accounts. You can restore it from the team setup history.`,'Delete team'))return;
    try{await api('DELETE',`/api/org/teams/${ds.t}`);await tmReload('Team deleted');}catch(e){toast(e.message,'error');}},
  'tm-cdel':async ds=>{const c=TM.data.customers.find(x=>x.id===ds.c);if(!await confirmBox(`Delete the customer "${c.name}" and its services?`,'Delete customer'))return;
    try{await api('DELETE',`/api/org/customers/${ds.c}`);await cuReload('Customer deleted');}catch(e){toast(e.message,'error');}},
  'tm-padd':ds=>openModal({title:'Add a service',save:'Add service',body:`<div class="form">${fieldHtml({key:'name',label:'Service name',required:true,placeholder:'For example ERP platform, Partner EDI'})}<p class="hint full">A service is what you deliver to this customer and agree in a service level agreement (ITIL 4 service relationship management).</p></div>`,
    onSave:async ov=>{await api('POST',`/api/org/customers/${ds.c}/projects`,{name:readFields(ov).name});await cuReload('Service added');}}),
  'tm-pdel':async ds=>{if(!await confirmBox('Delete this service? People assigned to it lose that assignment.','Delete service'))return;
    try{await api('DELETE',`/api/org/customers/${ds.c}/projects/${ds.p}`);await cuReload('Service deleted');}catch(e){toast(e.message,'error');}}});
Object.assign(ACT,{'task-share':async ds=>{
  const t=S.plans.tasks.find(x=>x.id===ds.task);if(!t)return;
  let org;try{org=await api('GET','/api/org');}catch(e){return toast(e.message,'error');}
  const cust=org.customers.find(c=>c.name===t.customer);
  const serving=cust?cust.teams.map(id=>org.team_names[id]).filter(Boolean):Object.values(org.team_names);
  const mine=(S.myTeams||[]).map(x=>x.name);
  const opts=n=>['',...serving].map(x=>`<option value="${esc(x)}" ${x===(n||'')?'selected':''}>${esc(x||'Owning team')}</option>`).join('');
  openModal({title:`Share: ${t.name}`,save:'Save sharing',size:'wide',body:`<div class="form">
    <div class="field"><label>Owner team</label><select data-field="owner">${mine.map(n=>`<option ${n===(t._owner||t.owner_team)?'selected':''}>${esc(n)}</option>`).join('')}</select>
      <p class="hint">The owning team changes the plan and signs off go/no-go.</p></div>
    <div class="field full"><label>Share the whole plan with</label>${serving.filter(n=>n!==(t._owner||t.owner_team)).map(n=>`<label style="display:flex;gap:8px;align-items:center"><input type="checkbox" data-share="${esc(n)}" ${(t.shared_with||[]).includes(n)?'checked':''}> ${esc(n)}</label>`).join('')||'<span class="st">No other team serves this customer.</span>'}</div>
    <div class="field full"><label>Give phases to other teams (partial sharing)</label>
      ${t.phases.map(ph=>`<div style="display:flex;gap:10px;align-items:center;margin:4px 0"><span style="min-width:200px">${esc(ph.name)}</span><select data-phase="${esc(ph.id)}">${opts(ph.team)}</select></div>`).join('')}
      <p class="hint">A team given a phase sees the plan with just its own phases and steps, and can update and complete them. Teams serving ${esc(t.customer||'this customer')} are listed.</p></div></div>`,
    onSave:async ov=>{
      await planOp('task.share',{task_id:t.id,owner_team:ov.querySelector('[data-field=owner]').value,shared_with:[...ov.querySelectorAll('[data-share]:checked')].map(x=>x.dataset.share)});
      for(const sel of ov.querySelectorAll('[data-phase]')){const ph=t.phases.find(x=>x.id===sel.dataset.phase);if((ph.team||'')!==sel.value)await planOp('phase.team',{task_id:t.id,phase_id:ph.id,team:sel.value});}
      toast('Sharing saved','success');await load();}});}});
Object.assign(ACT,{'team-claim':async()=>{if(!await confirmBox('Move the older shift log and handovers into this team? Other teams will not see them.','Move'))return;
  try{const r=(await api('POST','/api/org/claim-legacy',{})).moved;toast(`Moved ${r.shift_days} day(s) and ${r.handovers} handover(s)`+(r.kept?`; ${r.kept} day(s) already existed and were kept`:''),'success');await load();}catch(e){toast(e.message,'error');}}});
Object.assign(ACT,{'ask-open':()=>askOpen(),'onshift':()=>{RO.start=today();RO.mode='week';go('roster');}});
document.addEventListener('click',async e=>{
  const el=e.target.closest('[data-act]');if(!el||el.closest('#modalRoot'))return;
  if(el.closest('summary')||el.tagName==='A')e.preventDefault();
  const fn=ACT[el.dataset.act];if(!fn)return;
  if(actRank(el)>S.rank)return toast('Your role cannot do this.','error');
  try{await fn({...el.dataset});}catch(err){}
});
document.addEventListener('toggle',e=>{const d=e.target;if(!(d instanceof HTMLDetailsElement)||!d.dataset.key)return;if(d.open)S.open.add(d.dataset.key);else S.open.delete(d.dataset.key);persistOpen();},true);
document.addEventListener('input',e=>{const k=e.target.dataset?.filter,scope=e.target.dataset?.scope;if(!k||!scope)return;
  const f=scope==='plans'?pf():scope==='runbook'?rbf():tf(scope);f[k]=e.target.type==='checkbox'?e.target.checked:e.target.value;renderHost(scope);});
document.addEventListener('change',e=>{if(e.target.dataset?.filter&&e.target.tagName==='SELECT')e.target.dispatchEvent(new Event('input',{bubbles:true}));});
let dragDepth=0;const drop=$('#drop');
window.addEventListener('dragenter',e=>{if([...(e.dataTransfer?.types||[])].includes('Files')){dragDepth++;drop.classList.add('on');}});
window.addEventListener('dragleave',()=>{dragDepth=Math.max(0,dragDepth-1);if(!dragDepth)drop.classList.remove('on');});
window.addEventListener('dragover',e=>e.preventDefault());
window.addEventListener('drop',e=>{e.preventDefault();dragDepth=0;drop.classList.remove('on');const f=e.dataTransfer?.files?.[0];if(!f)return;
  if(S.rank<3)return toast('Importing files needs the Lead role.','error');
  const mod=S.schema[S.view]?.module||({plans:'plans',runbook:'runbook',shiftlog:'shiftlog'}[S.view])||'auto';importDialog(mod,f);});
window.addEventListener('hashchange',()=>{if(!S.me)return;const v=location.hash.slice(1);if(v&&v!==S.view)go(v);});  // not signed in yet: the page opens the requested screen after sign-in
</script>
<script>
"use strict";
document.documentElement.dataset.theme=localStorage.getItem('msp.theme')||'dark';

/* ---------- role-aware controls ---------- */
const ACT_RANK={'add-row':2,'clone-row':2,'del-row':3,'columns':3,'import':3,'template':2,'export-menu':2,'export-xlsx':2,'export-json':2,
  'rb-add':2,'task-new':2,'templates':3,'run-open':2,'task-edit':2,'phase-add':2,'step-cols':3,'task-clone':2,'task-tpl':3,'task-del':3,
  'phase-sign':3,'phase-edit':2,'phase-move':2,'phase-clone':2,'phase-del':3,'step-add':2,'step-edit':2,'step-clone':2,'step-move':2,'step-del':2,
  'undo-import':3,'h-open':2,'h-add-row':2,'h-edit-row':2,'h-clone-row':2,'h-del-row':2,'h-cols':3,'h-del-sec':3,'h-add-sec':2,'h-del-node':3,
  'h-clone-shift':2,'ssh':2,'rdp':2,'term-new':2,
  'ro-leave':2,'ro-swap':2,'ro-cancel':2,'ro-export':2,'ro-decide':3,'ro-restore':3,'ro-cell':3,'ro-holidays':3,'ro-settings':3,'ro-import':3,'ro-new':3,'ro-shiftcell':3,'ro-clone':3,'ro-deleted':3,'ro-consent':2,'ro-template':2,'ask-open':1,'onshift':1,'team-claim':3,'task-share':2,'tm-team':4,'tm-cust':4,'tm-cedit':4,'tm-cdel':4,'tm-padd':4,'tm-pdel':4,'tm-del':4,'tm-edit':3,'tm-add':3,'tm-proj':3,'tm-rm':3};
function actRank(el){let r=ACT_RANK[el.dataset.act]||1;if(el.dataset.sheet==='Bulletins'&&['add-row','clone-row'].includes(el.dataset.act))r=3;return r;}
function applyRoles(root){if(!S.rank||!root.querySelectorAll)return;
  if(root.dataset&&root.dataset.act)root.classList.toggle('hidden',actRank(root)>S.rank);
  root.querySelectorAll('[data-act]').forEach(el=>el.classList.toggle('hidden',actRank(el)>S.rank));}
new MutationObserver(ms=>{for(const m of ms)m.addedNodes.forEach(n=>{if(n.nodeType===1)applyRoles(n);});}).observe(document.body,{childList:true,subtree:true});

/* ---------- sign-in ---------- */
function applyAuth(st){S.me=st.user;S.rank=st.user.rank;S.csrf=st.csrf;S.user=st.user.name;S.caps={terminals:st.terminals,guac:st.guacamole};
  S.roleHelp=st.role_help||{};S.roles=st.roles||[];document.body.className='r'+S.rank;}
async function boot(){
  const err=new URLSearchParams(location.search).get('auth_error');
  if(err)history.replaceState(null,'',location.pathname+location.hash);
  let st;
  try{st=await api('GET','/api/auth/status',undefined,{noAuth:true});}
  catch(e){document.body.innerHTML=`<p style="padding:40px;font-family:sans-serif">Cannot reach the MSP Ops Center server: ${esc(e.message)}</p>`;return;}
  if(!st.authenticated)return showLogin(st,err||st.note||'');
  applyAuth(st);
  if(st.user.must_change)return showLogin(st,'',true);
  await startApp();
}
async function startApp(){
  $('#login').classList.add('hidden');$('#login').innerHTML='';$('#app').classList.remove('hidden');
  if(localStorage.getItem('msp.collapsed'))$('#app').classList.add('collapsed');
  await load();
  const h=location.hash.slice(1);
  if(h.startsWith('terminal-new=')){let p={};try{p=JSON.parse(decodeURIComponent(h.slice(13)));}catch(_){} go('terminals');sshDialog(p);}
  else go(h.startsWith('shiftlog=')?h:S.view);
  fillOnShift();
  if(!S.started){S.started=true;setInterval(poll,15000);setInterval(renderTicker,60000);
    document.addEventListener('visibilitychange',()=>{if(!document.hidden)poll();});}
}
let endingSession=false;
function sessionEnded(){
  if(endingSession)return;endingSession=true;S.csrf='';while(S.modals.length)closeModal();
  api('GET','/api/auth/status',undefined,{noAuth:true}).then(st=>showLogin(st,'Your session has ended (signed out, or idle for too long). Please sign in again.'))
    .catch(()=>{}).finally(()=>{endingSession=false;});
}
function showLogin(st,msg,mustChange){
  $('#app').classList.add('hidden');const L=$('#login');L.classList.remove('hidden');
  const note=msg?`<div class="msg">${esc(msg)}</div>`:'';
  let card;
  if(mustChange)card=`<h2>Choose a new password</h2><p class="st">Your account has a temporary password. Choose your own to continue (at least 10 characters with letters and numbers).</p>${note}
    <input type="password" name="current" placeholder="Temporary password" autocomplete="current-password" required aria-label="Temporary password">
    <input type="password" name="new" placeholder="New password" autocomplete="new-password" required aria-label="New password">
    <input type="password" name="confirm" placeholder="Repeat the new password" autocomplete="new-password" required aria-label="Repeat new password">
    <button class="btn primary" type="submit">Save password and continue</button><button class="btn ghost" type="button" data-lout>Sign out</button>`;
  else if(st.setup_required)card=`<h2>Create the administrator</h2><p class="st">First start. Enter the setup code shown in the server window (also in FIRST_RUN_SETUP_CODE.txt next to the program).</p>${note}
    <input name="code" placeholder="Setup code, e.g. 1A2B-3C4D-5E6F" required aria-label="Setup code" autocomplete="off">
    <input name="name" placeholder="Your full name" required aria-label="Full name" autocomplete="name">
    <input name="username" placeholder="User name" required aria-label="User name" autocomplete="username">
    <input type="email" name="email" placeholder="Email (optional, used for Microsoft 365 sign-in)" aria-label="Email" autocomplete="email">
    <input type="password" name="password" placeholder="Password (10+ characters, letters and numbers)" required aria-label="Password" autocomplete="new-password">
    <input type="password" name="confirm" placeholder="Repeat the password" required aria-label="Repeat password" autocomplete="new-password">
    <button class="btn primary" type="submit">Create administrator</button>`;
  else card=`<h2>Sign in</h2>${note}
    <input name="username" placeholder="User name or email" required aria-label="User name" autocomplete="username" autofocus>
    <input type="password" name="password" placeholder="Password" required aria-label="Password" autocomplete="current-password">
    <button class="btn primary" type="submit">Sign in</button>
    ${st.methods&&st.methods.entra?`<div class="or">or</div><a class="btn ms" href="/auth/entra/login?next=${encodeURIComponent('/'+location.hash)}"><i aria-hidden="true"><b style="background:#F25022"></b><b style="background:#7FBA00"></b><b style="background:#00A4EF"></b><b style="background:#FFB900"></b></i>Sign in with Microsoft 365</a>`:''}
    <p class="st">Forgotten your password? Ask an administrator to reset it.</p>`;
  L.innerHTML=`<div class="art"><div class="brand"><div class="logo" aria-hidden="true">M</div><div><b>MSP Ops Center</b><span>ITIL v4 service operations</span></div></div>
    <div><h1>Plans, runbooks, shifts and terminals for the whole team.</h1><p>Every change and every terminal session is recorded against the person who made it.</p></div>
    <p style="position:relative;z-index:1;color:#7486A6;font-size:12px">${esc(location.host)}</p></div>
    <div class="form"><form class="card" id="loginForm" novalidate>${card}</form></div>`;
  const f=$('#loginForm');
  f.querySelector('[data-lout]')?.addEventListener('click',logout);
  f.querySelector('input')?.focus();
  f.addEventListener('submit',async e=>{
    e.preventDefault();const v=Object.fromEntries(new FormData(f).entries());const btn=f.querySelector('[type=submit]');btn.disabled=true;
    const fail=m=>{f.querySelector('.msg')?.remove();btn.insertAdjacentHTML('beforebegin',`<div class="msg">${esc(m)}</div>`);btn.disabled=false;};
    try{
      if(mustChange){if(v.new!==v.confirm)return fail('The two new passwords do not match.');
        await api('POST','/api/me/password',{current:v.current,new:v.new});}
      else if(st.setup_required){if(v.password!==v.confirm)return fail('The two passwords do not match.');
        await api('POST','/api/auth/setup',{code:v.code,name:v.name,username:v.username,email:v.email,password:v.password},{noAuth:true});}
      else await api('POST','/api/auth/login',{username:v.username,password:v.password},{noAuth:true});
      const s2=await api('GET','/api/auth/status',undefined,{noAuth:true});
      applyAuth(s2);
      if(s2.user.must_change)return showLogin(s2,'',true);
      await startApp();
    }catch(err){fail(err.message);}
  });
}
async function logout(){
  if(S.terms.some(t=>t.state==='connected')&&!await confirmBox('Signing out closes your open terminal sessions. Continue?','Sign out',false))return;
  S.terms.forEach(t=>{try{t.ws&&t.ws.close();}catch(_){}});
  try{await api('POST','/api/auth/logout',{},{noAuth:true});}catch(_){}
  location.href='/';
}
function toggleTheme(){const t=document.documentElement.dataset.theme==='light'?'dark':'light';document.documentElement.dataset.theme=t;localStorage.setItem('msp.theme',t);}
function accountModal(){
  const me=S.me;
  openModal({title:'Your account',body:`<div class="full"><p><b>${esc(me.name)}</b> <span class="st">(${esc(me.username)})</span></p>
    <p class="st">Role: <b>${esc(me.role)}</b>. ${esc(S.roleHelp[me.role]||'')}</p><p class="st">Sign-in: ${esc({local:'user name and password',entra:'Microsoft 365',both:'password or Microsoft 365'}[me.auth]||me.auth)}.${me.last_login?' Last sign-in '+esc(fmtTime(me.last_login))+'.':''}</p></div>`,
    footExtra:`${me.auth!=='entra'?'<button class="btn" data-cpw>Change password</button>':''}<button class="btn" data-theme2>Switch light / dark</button>`,
    onMount:(ov,m)=>ov.addEventListener('click',e=>{if(e.target.closest('[data-cpw]')){closeModal(m);changePasswordModal();}if(e.target.closest('[data-theme2]'))toggleTheme();})});
}
function changePasswordModal(){
  openModal({title:'Change password',save:'Change password',body:fieldHtml({key:'current',label:'Current password',type:'password',full:true})
    +fieldHtml({key:'new',label:'New password',type:'password',full:true,hint:'At least 10 characters, with letters and numbers.'})+fieldHtml({key:'confirm',label:'Repeat new password',type:'password',full:true}),
    onSave:async ov=>{const v=readFields(ov);if(v.new!==v.confirm){toast('The two new passwords do not match','error');return false;}
      await api('POST','/api/me/password',{current:v.current,new:v.new});toast('Password changed. Your other sessions were signed out.','success');}});
}
function exportMenu(){
  openModal({title:'Export',body:`<p class="full st">Download the live state of the whole system. The Excel workbook holds every table plus the execution plans and the shift log, and any of its sheets can be imported again. The JSON backup restores everything at once.</p>
    <div class="full"><button class="btn primary" data-ex="xlsx">Excel workbook</button> <button class="btn" data-ex="json">JSON backup</button></div>`,
    onMount:(ov,m)=>ov.addEventListener('click',e=>{const b=e.target.closest('[data-ex]');if(b){closeModal(m);download('/api/export/'+b.dataset.ex);}})});
}

/* ---------- my work ---------- */
function mine(v){const n=nk(v);return !!n&&!!S.me&&(n===nk(S.me.name)||n===nk(S.me.username));}
function myWork(){
  const td=today(),from=addDays(td,-1),to=addDays(td,7),steps=[];
  S.plans.tasks.forEach(t=>{if(t.date<from||t.date>to||['Completed','Cancelled'].includes(t.status))return;
    t.phases.forEach(ph=>ph.steps.forEach(st=>{if(mine(st.owner)&&!['Done','Skipped'].includes(st.status))steps.push({t,ph,st});}));});
  steps.sort((a,b)=>String(a.st.planned).localeCompare(String(b.st.planned)));
  const gates=S.rank>=3?S.plans.tasks.filter(t=>t.date>=from&&t.date<=to&&['Scheduled','In Progress'].includes(t.status))
    .flatMap(t=>t.phases.filter(p=>p.checkpoint&&!p.signed_off_by).map(p=>({t,p}))):[];
  const incs=(S.tables.Incidents?.rows||[]).filter(r=>isOpen(r)&&mine(r.Engineer));
  const unread=activeBulletins().filter(b=>!S.reads.has(b.Record_ID));
  return {steps,gates,incs,unread};
}
function myWorkCount(){if(!S.me||!S.plans)return 0;const w=myWork();return w.steps.length+w.gates.length+w.incs.length;}
function renderMyWork(){
  const w=myWork(),h=new Date().getHours(),greet=h<12?'Good morning':h<18?'Good afternoon':'Good evening';
  const first=(S.me.name||S.me.username).split(/\s+/)[0];
  const li=(t,s,lead='',act='')=>`<li>${lead}<div class="t">${act?`<button class="linkish" ${act}>${esc(t)}</button>`:esc(t)}<small>${esc(s)}</small></div></li>`;
  const mineAudit=S.audit.filter(a=>nk(a.User)===nk(S.me.username)).slice(0,8);
  const panel=(title,n,body,empty)=>`<section class="panel"><h3>${esc(title)}${n?` <span class="n">${n}</span>`:''}</h3><ul class="list">${body||`<li class="st">${esc(empty)}</li>`}</ul></section>`;
  return `<div class="hello"><h2>${greet}, ${esc(first)}</h2><span class="st">${new Date().toLocaleDateString([], {weekday:'long',day:'numeric',month:'long'})}. You are signed in as ${esc(S.me.role)}.</span></div>
  <div class="toolbar"><button class="btn primary" data-act="task-new">New execution plan</button><button class="btn" data-act="term-new">New SSH session</button>
    <button class="btn" data-act="rb-add" data-c="" data-e="" data-a="">Add runbook entry</button><button class="btn" data-act="open-shift" data-d="${today()}" data-s="${currentShift()}">Open current shift</button>
    <button class="btn ghost" data-act="palette">Search (Ctrl K)</button></div>
  <div class="cols">
    ${panel('Your steps in the next 7 days',w.steps.length,w.steps.slice(0,12).map(x=>li(`${x.st.num} ${x.st.title}`,`${x.t.customer}: ${x.t.name}, planned ${x.st.planned}`,`<span class="${stCls(x.st.status)}">${esc(x.st.status)}</span>`,`data-act="open-task" data-id="${esc(x.t.id)}"`)).join(''),'No steps are assigned to you. Steps count as yours when the Owner matches your name.')}
    ${S.rank>=3?panel('Waiting for your go/no-go',w.gates.length,w.gates.map(x=>li(`${x.t.name}: ${x.p.num}. ${x.p.name}`,`${x.t.date}, ${x.t.customer}`,'',`data-act="open-task" data-id="${esc(x.t.id)}"`)).join(''),'No checkpoints waiting.'):''}
    <section class="panel" id="myShifts"><h3>Your shifts, next 7 days</h3><ul class="list"><li class="st">Loading…</li></ul></section>
    ${panel('Your open incidents',w.incs.length,w.incs.map(r=>li(`${r.Ticket_ID} ${r.Summary}`,`${r.Customer}, ${r.Status}`,`<span class="badge pri-${esc(r.Priority)}">${esc(r.Priority)}</span>`)).join(''),'Nothing assigned to you.')}
    ${panel('Unread bulletins',w.unread.length,w.unread.map(b=>li(b.Title,b.Customers||'',`<span class="sev sev-${esc(b.Severity)}">${esc(b.Severity)}</span>`,`data-act="bulletin" data-id="${esc(b.Record_ID)}"`)).join(''),'You are up to date.')}
    ${panel('Your recent changes','',mineAudit.map(a=>li(`${a.Action}: ${a.Target}`,fmtTime(a.Timestamp))).join(''),'No changes yet.')}
  </div>`;
}

/* ---------- search palette (Ctrl+K) ---------- */
function paletteItems(){
  const items=VIEWS.filter(v=>(v[3]||1)<=S.rank).map(v=>({k:'Go to',t:v[1],run:()=>go(v[0])}));
  if(S.rank>=2)items.push({k:'Action',t:'New execution plan task',run:()=>{go('plans');taskModal(null);}},{k:'Action',t:'New SSH session',run:()=>sshDialog({})},
    {k:'Action',t:'Add asset',run:()=>{go('Asset_Master');editRow('Asset_Master',null);}},{k:'Action',t:'Add runbook entry',run:()=>{go('runbook');editRow('Runbook',null);}});
  items.push({k:'Action',t:'Switch light / dark theme',run:toggleTheme});
  (S.tables.Asset_Master?.rows||[]).forEach(a=>items.push({k:'CI',t:a.Device,s:[a.Customer,a.IP_Address,a.Application].filter(Boolean).join(', '),run:()=>gotoRunbook(a.Device)}));
  S.plans.tasks.forEach(t=>items.push({k:'Plan',t:t.name,s:[t.date,t.customer,t.ticket_id].filter(Boolean).join(', '),run:()=>focusTask(t.id)}));
  (S.tables.Incidents?.rows||[]).forEach(r=>items.push({k:'Incident',t:`${r.Ticket_ID} ${r.Summary}`,s:r.Customer,run:()=>{go('Incidents');editRow('Incidents',r.Record_ID);}}));
  (S.tables.Runbook?.rows||[]).forEach(r=>items.push({k:'Runbook',t:r.Title||String(r.Note).slice(0,70),s:`${r.Asset}, ${r.Date}`,run:()=>gotoRunbook(r.Asset)}));
  (S.tables.Bulletins?.rows||[]).forEach(b=>items.push({k:'Bulletin',t:b.Title,s:b.Severity,run:()=>showBulletin(b.Record_ID)}));
  return items;
}
function openPalette(){
  if(document.querySelector('.palette'))return;
  const ov=document.createElement('div');ov.className='overlay';
  ov.innerHTML=`<div class="palette" role="dialog" aria-label="Search"><input placeholder="Search devices, tickets, plans, runbook fixes, pages…" aria-label="Search"><ul role="listbox"></ul></div>`;
  $('#modalRoot').appendChild(ov);const m={el:ov,o:{}};S.modals.push(m);
  const inp=ov.querySelector('input'),ul=ov.querySelector('ul'),all=paletteItems();let sel=0,list=[];
  const draw=()=>{const q=inp.value.toLowerCase().split(/\s+/).filter(Boolean);
    list=all.filter(i=>q.every(w=>(i.t+' '+(i.s||'')+' '+i.k).toLowerCase().includes(w))).slice(0,60);sel=Math.min(sel,Math.max(0,list.length-1));
    ul.innerHTML=list.map((i,n)=>`<li class="${n===sel?'on':''}" data-i="${n}" role="option"><span class="k">${esc(i.k)}</span><span>${esc(i.t)}</span>${i.s?`<small>${esc(i.s)}</small>`:''}</li>`).join('')||'<li class="st">No matches.</li>';
    ul.querySelector('.on')?.scrollIntoView({block:'nearest'});};
  const pick=n=>{const it=list[n];if(!it)return;closeModal(m);it.run();};
  inp.addEventListener('input',()=>{sel=0;draw();});
  inp.addEventListener('keydown',e=>{if(e.key==='ArrowDown'){sel=Math.min(sel+1,list.length-1);draw();e.preventDefault();}
    else if(e.key==='ArrowUp'){sel=Math.max(sel-1,0);draw();e.preventDefault();}else if(e.key==='Enter'){pick(sel);e.preventDefault();}});
  ul.addEventListener('click',e=>{const li=e.target.closest('[data-i]');if(li)pick(+li.dataset.i);});
  ov.addEventListener('click',e=>{if(e.target===ov)closeModal(m);});
  draw();inp.focus();
}
document.addEventListener('keydown',e=>{if((e.ctrlKey||e.metaKey)&&e.key.toLowerCase()==='k'&&S.started&&!$('#app').classList.contains('hidden')){e.preventDefault();openPalette();}});

/* ---------- terminals ---------- */
const XTERM={js:['/static/vendor/xterm.js','sha384-/nfmYPUzWMS6v2atn8hbljz7NE0EI1iGx34lJaNzyVjWGDzMv+ciUZUeJpKA3Glc'],
  css:['/static/vendor/xterm.css','sha384-LJcOxlx9IMbNXDqJ2axpfEQKkAYbFjJfhXexLfiRJhjDU81mzgkiQq8rkV0j6dVh'],
  fit:['/static/vendor/xterm-addon-fit.js','sha384-AQLWHRKAgdTxkolJcLOELg4E9rE89CPE2xMy3tIRFn08NcGKPTsELdvKomqji+DL']};
let xtermP=null;
function loadXterm(){
  if(xtermP)return xtermP;
  if(window.MSP_NO_XTERM)return (xtermP=Promise.resolve(false));
  xtermP=new Promise(res=>{let finished=false;const done=ok=>{if(!finished){finished=true;res(ok&&typeof window.Terminal==='function');}};
    const add=(tag,attrs)=>{const el=document.createElement(tag);Object.assign(el,attrs);el.crossOrigin='anonymous';document.head.appendChild(el);return el;};
    add('link',{rel:'stylesheet',href:XTERM.css[0],integrity:XTERM.css[1]});
    const s1=add('script',{src:XTERM.js[0],integrity:XTERM.js[1]});
    s1.onload=()=>{const s2=add('script',{src:XTERM.fit[0],integrity:XTERM.fit[1]});s2.onload=()=>done(true);s2.onerror=()=>done(true);};
    s1.onerror=()=>done(false);setTimeout(()=>done(false),8000);});
  return xtermP;
}
const stripAnsi=s=>String(s).replace(/\x1b\[[0-9;?]*[ -\/]*[@-~]|\x1b\][^\x07\x1b]*(\x07|\x1b\\)|\x1b[()][A-Z0-9]|\r/g,'');
function makeWidget(pane,useX){
  if(useX){
    const box=document.createElement('div');box.className='xt';pane.appendChild(box);
    const term=new window.Terminal({cursorBlink:true,fontFamily:'"IBM Plex Mono",Consolas,monospace',fontSize:13,scrollback:5000,theme:{background:'#05080F'}});
    const FA=window.FitAddon&&(window.FitAddon.FitAddon||window.FitAddon);let fit=null;if(typeof FA==='function'){fit=new FA();term.loadAddon(fit);}
    term.open(box);
    return {write:d=>term.write(d),onData:cb=>term.onData(cb),fit:()=>{try{fit&&fit.fit();}catch(_){}},size:()=>({cols:term.cols,rows:term.rows}),focus:()=>term.focus(),dispose:()=>term.dispose()};
  }
  const wrap=document.createElement('div');wrap.className='fallback';
  wrap.innerHTML=`<pre aria-live="polite"></pre><div class="row"><input aria-label="Command" placeholder="Type a command and press Enter"><button class="btn sm" data-k="tab">Tab</button><button class="btn sm" data-k="c">Ctrl+C</button></div>`;
  pane.appendChild(wrap);const pre=wrap.querySelector('pre'),inp=wrap.querySelector('input');let cb=()=>{};
  inp.addEventListener('keydown',e=>{if(e.key==='Enter'){cb(inp.value+'\r');inp.value='';e.preventDefault();}});
  wrap.addEventListener('click',e=>{const b=e.target.closest('[data-k]');if(b){cb(b.dataset.k==='tab'?inp.value+'\t':'\x03');inp.value='';}});
  return {write:d=>{pre.textContent=(pre.textContent+stripAnsi(d)).slice(-200000);pre.scrollTop=pre.scrollHeight;},onData:f=>{cb=f;},fit:()=>{},
    size:()=>({cols:120,rows:32}),focus:()=>inp.focus(),dispose:()=>{}};
}
function renderTerminals(){
  const host=$('#termHost');
  if(!host.dataset.ready){
    host.dataset.ready='1';
    host.innerHTML=`<div class="toolbar" style="margin:0"><div><h2>Terminals</h2><p class="sub" style="margin:0">SSH sessions open here in tabs. Every session is recorded for audit; typing after a password prompt is masked in the recording.</p></div><span class="grow"></span>
      <button class="btn primary" data-act="term-new">New SSH session</button></div><div class="term-tabs" id="termTabs" role="tablist"></div><div class="term-panes" id="termPanes"></div>`;
    host.addEventListener('click',e=>{
      const x=e.target.closest('[data-tclose]');if(x){e.stopPropagation();closeTerm(x.dataset.tclose);return;}
      const tab=e.target.closest('[data-term]');if(tab){S.activeTerm=tab.dataset.term;renderTerminals();}
    });
  }
  renderTermTabs();
  $('#termPanes').querySelectorAll('.term-pane').forEach(p=>p.classList.toggle('on',p.dataset.pane===S.activeTerm));
  let empty=$('#termPanes .term-empty');
  if(!S.terms.length){if(!empty){empty=document.createElement('div');empty.className='term-empty';empty.innerHTML=`<div><h3>No terminal sessions open</h3><p class="st">Open one from a CI in the CMDB (SSH button), or start a new session.</p><button class="btn primary" data-act="term-new">New SSH session</button></div>`;$('#termPanes').appendChild(empty);}}
  else if(empty)empty.remove();
  const t=S.terms.find(x=>x.id===S.activeTerm);if(t){setTimeout(()=>{t.w.fit();sendResize(t);t.w.focus();},30);}
}
function renderTermTabs(){
  const tabs=$('#termTabs');if(tabs)tabs.innerHTML=S.terms.map(t=>`<div class="term-tab ${t.id===S.activeTerm?'on':''}" data-term="${esc(t.id)}" role="tab" tabindex="0" aria-selected="${t.id===S.activeTerm}"><span class="sdot ${esc(t.state)}"></span>${esc(t.title)}<button class="x" data-tclose="${esc(t.id)}" aria-label="Close ${esc(t.title)}">&times;</button></div>`).join('');
  renderChrome();
}
function sendResize(t){if(t.ws&&t.ws.readyState===1&&t.state==='connected'){const z=t.w.size();t.ws.send(JSON.stringify({t:'resize',cols:z.cols,rows:z.rows}));}}
let resizeTimer=null;window.addEventListener('resize',()=>{clearTimeout(resizeTimer);resizeTimer=setTimeout(()=>{const t=S.terms.find(x=>x.id===S.activeTerm);if(t&&S.view==='terminals'){t.w.fit();sendResize(t);}},150);});
window.addEventListener('beforeunload',e=>{if(S.terms.some(t=>t.state==='connected')){e.preventDefault();e.returnValue='';}});
function closeTerm(id){
  const t=S.terms.find(x=>x.id===id);if(!t)return;
  try{if(t.ws&&t.ws.readyState===1)t.ws.send(JSON.stringify({t:'close'}));t.ws&&t.ws.close();}catch(_){}
  try{t.w.dispose();}catch(_){} t.pane.remove();S.terms=S.terms.filter(x=>x.id!==id);
  if(S.activeTerm===id)S.activeTerm=S.terms.length?S.terms[S.terms.length-1].id:null;
  renderTerminals();
}
function credFields(){
  return fieldHtml({key:'auth',label:'Sign in with',type:'select',options:['password','key'],value:'password'})
    +fieldHtml({key:'password',label:'Password',type:'password'})
    +`<div class="field full hidden" data-keyf>${fieldHtml({key:'key',label:'Private key (OpenSSH or PEM)',type:'mono'}).replace(/^<div class="field full">|<\/div>$/g,'')}</div>`
    +`<div class="field hidden" data-keyf>${fieldHtml({key:'passphrase',label:'Key passphrase (if any)',type:'password'}).replace(/^<div class="field">|<\/div>$/g,'')}</div>`;
}
function wireCreds(ov){const sel=ov.querySelector('[data-field="auth"]');[...sel.options].forEach(o=>o.textContent=o.value==='key'?'Private key':'Password');
  const sync=()=>{const k=sel.value==='key';ov.querySelectorAll('[data-keyf]').forEach(x=>x.classList.toggle('hidden',!k));ov.querySelector('[data-field="password"]').closest('.field').classList.toggle('hidden',k);};
  sel.addEventListener('change',sync);sync();}
function credsOf(v){return v.auth==='key'?{password:'',key:v.key,passphrase:v.passphrase}:{password:v.password,key:'',passphrase:''};}
function sshDialog(p={}){
  if(!S.caps.terminals)return toast('SSH terminals are not available yet: the server needs the paramiko package (Start_MSP_Ops_Center.bat installs it).','error');
  const a=p.asset_id?findRow('Asset_Master',p.asset_id):null;
  const host=p.host||(a&&(a.IP_Address||a.Hostname_FQDN||a.Device))||'';
  const hosts=uniq((S.tables.Asset_Master?.rows||[]).flatMap(r=>[r.Device,r.IP_Address,r.Hostname_FQDN]));
  openModal({title:'New SSH session'+(a?': '+a.Device:''),save:'Connect',
    body:`<datalist id="dl-hosts">${hosts.map(h=>`<option value="${esc(h)}">`).join('')}</datalist>`
      +fieldHtml({key:'host',label:'Host or IP',value:host,list:'dl-hosts',required:true})+fieldHtml({key:'port',label:'Port',value:String(p.port||22)})
      +fieldHtml({key:'username',label:'User name on the server',value:p.username||'',required:true,full:true})+credFields()
      +fieldHtml({key:'newtab',type:'checkbox',label:'Where',text:'Open in a separate browser tab (you will enter the password there)',full:true})
      +`<p class="full st">Your password or key is sent once to sign in and is never stored. The session is recorded for audit.</p>`,
    onMount:wireCreds,
    onSave:async ov=>{const v=readFields(ov);
      if(!v.host.trim()||!v.username.trim()){toast('Enter the host and user name','error');return false;}
      if(v.newtab){window.open(location.pathname+'#terminal-new='+encodeURIComponent(JSON.stringify({asset_id:(a&&v.host===host)?p.asset_id:'',host:v.host,port:v.port,username:v.username})),'_blank');return;}
      await openTerminal({asset_id:(a&&v.host===host)?p.asset_id:'',host:v.host,port:v.port,username:v.username},credsOf(v));}});
}
async function openTerminal(p,creds){
  const r=await api('POST','/api/terminal/sessions',p);
  const useX=await loadXterm();
  const T={id:r.sid,title:`${r.username}@${r.label||r.host}`,state:'connecting',creds};
  S.terms.push(T);S.activeTerm=T.id;
  if(S.view!=='terminals')go('terminals');else renderTerminals();
  const pane=document.createElement('div');pane.className='term-pane on';pane.dataset.pane=T.id;$('#termPanes').appendChild(pane);T.pane=pane;
  pane.insertAdjacentHTML('beforeend','<span class="term-banner">Recorded session</span>');
  T.w=makeWidget(pane,useX);renderTerminals();
  const ws=new WebSocket(`${location.protocol==='https:'?'wss':'ws'}://${location.host}/ws/terminal?sid=${encodeURIComponent(r.sid)}&token=${encodeURIComponent(r.token)}`);
  T.ws=ws;
  const sendAuth=(extra={})=>{const z=T.w.size();ws.send(JSON.stringify({t:'auth',...(T.creds||{}),...extra,cols:z.cols,rows:z.rows}));};
  ws.onopen=()=>{T.w.write(`\x1b[90mConnecting to ${r.host}:${r.port} as ${r.username} ...\x1b[0m\r\n`);sendAuth();};
  ws.onmessage=ev=>{let m;try{m=JSON.parse(ev.data);}catch(_){return;}
    if(m.t==='out')T.w.write(m.d);
    else if(m.t==='status')T.w.write(`\x1b[90m${m.m}\x1b[0m\r\n`);
    else if(m.t==='hostkey')confirmBox(`First connection to ${r.host}. Its host key is:\n\n${m.type}  ${m.fingerprint}\n\nCompare this with the fingerprint from the server owner before trusting it. Once trusted, a changed key will be blocked.`,'Trust and connect',false)
      .then(ok=>{if(ok)sendAuth({trust:m.fingerprint});else{T.w.write('\r\nHost key not trusted. Connection cancelled.\r\n');T.state='error';renderTermTabs();try{ws.close();}catch(_){}}});
    else if(m.t==='connected'){T.state='connected';T.creds=null;renderTermTabs();T.w.fit();sendResize(T);T.w.focus();}
    else if(m.t==='error'){T.w.write(`\r\n\x1b[31m${m.m}\x1b[0m\r\n`);if(m.retry&&T.state!=='connected')retryCreds(T,sendAuth);else{T.state='error';renderTermTabs();}}
    else if(m.t==='closed'){T.state='closed';T.w.write(`\r\n\x1b[90m${m.m||'Session closed.'}\x1b[0m\r\n`);renderTermTabs();}
  };
  ws.onclose=()=>{if(['connected','connecting'].includes(T.state)){T.state='closed';T.w.write('\r\n\x1b[90mDisconnected.\x1b[0m\r\n');renderTermTabs();}};
  T.w.onData(d=>{if(ws.readyState===1&&T.state==='connected')ws.send(JSON.stringify({t:'in',d}));});
}
function retryCreds(T,sendAuth){
  openModal({title:'Sign in again: '+T.title,save:'Try again',body:credFields(),onMount:wireCreds,
    onSave:async ov=>{T.creds=credsOf(readFields(ov));sendAuth();},onClose:r=>{if(r!=='saved'){T.state='error';renderTermTabs();try{T.ws.close();}catch(_){}}}});
}
function rdpDialog(id){
  const a=findRow('Asset_Master',id);if(!a)return;
  const target=a.IP_Address||a.Hostname_FQDN||a.Device;
  openModal({title:'Remote Desktop: '+a.Device,size:'wide',
    body:`<p class="full st">Target: <b>${esc(target)}</b>${a.Customer?' ('+esc(a.Customer)+')':''}${a.Special_Care?`<br><span class="care-tag">${esc(a.Special_Care)}</span>`:''}</p>`
      +fieldHtml({key:'username',label:'User name (optional)',value:''})+fieldHtml({key:'domain',label:'Domain (optional, in-browser only)',value:''})
      +`<div class="full panel"><h3>Remote Desktop app</h3><p class="st">Downloads a .rdp file that opens Windows Remote Desktop on your computer, which must be able to reach the server. The launch is logged; the session itself is not recorded.</p><button class="btn" data-rdp="file">Open in Remote Desktop app</button></div>
      <div class="full panel"><h3>In this browser</h3><p class="st">${S.caps.guac?'Opens the desktop in a new browser tab through the Guacamole gateway, which can record the session.':'Not set up yet. An administrator adds the Guacamole address and key in msp_settings.ini.'}</p><button class="btn primary" data-rdp="guac" ${S.caps.guac?'':'disabled'}>Open in browser</button></div>`,
    onMount:(ov,m)=>ov.addEventListener('click',async e=>{const b=e.target.closest('[data-rdp]');if(!b||b.disabled)return;const v=readFields(ov);
      if(b.dataset.rdp==='file'){download(`/api/rdp/file?asset_id=${encodeURIComponent(id)}&username=${encodeURIComponent(v.username)}`);closeModal(m);return;}
      const w=window.open('about:blank','_blank');
      try{const r=await api('POST','/api/rdp/guac',{asset_id:id,username:v.username,domain:v.domain});
        if(w){w.opener=null;w.location.href=r.url;}else location.assign(r.url);closeModal(m);}
      catch(err){if(w)w.close();toast(err.message,'error');}})});
}

/* ---------- shift clone + print ---------- */
function cloneShiftModal(ds){
  const idx=SHIFTS.indexOf(ds.s),next=idx<2?{d:ds.d,s:SHIFTS[idx+1]}:{d:addDays(ds.d,1),s:'Morning'};
  openModal({title:`Clone the ${ds.s} shift of ${ds.d}`,save:'Clone shift',
    body:fieldHtml({key:'to_date',label:'Into date',type:'date',value:next.d})+fieldHtml({key:'to_shift',label:'Into shift',type:'select',options:SHIFTS,value:next.s})
      +fieldHtml({key:'mode',label:'What to copy',type:'select',options:['all','structure'],value:'all',full:true,hint:'All rows copies every section with its rows. Sections only copies sections and columns with empty tables.'})
      +fieldHtml({key:'overwrite',type:'checkbox',label:'If that shift already exists',text:'Replace it',full:true}),
    onMount:ov=>{[...ov.querySelector('[data-field="mode"]').options].forEach(o=>o.textContent=o.value==='all'?'All sections and rows':'Sections and columns only');},
    onSave:async ov=>{const v=readFields(ov);
      await run(()=>api('POST','/api/hierarchy/shift/clone',{date:ds.d,shift:ds.s,to_date:v.to_date,to_shift:v.to_shift,mode:v.mode,overwrite:v.overwrite}),'Shift cloned');
      openShiftKeys(v.to_date,v.to_shift);render();}});
}
function printTask(t){
  const rows=t.phases.map(ph=>`<tr class="ph"><td colspan="6">${esc(ph.num)}. ${esc(ph.name)}${ph.checkpoint?' (go/no-go checkpoint'+(ph.signed_off_by?', signed off by '+esc(ph.signed_off_by):'')+')':''}</td></tr>`+
    ph.steps.map(st=>`<tr><td>${esc(st.num)}</td><td style="padding-left:${6+st.level*14}px">${esc(st.title)}${st.command?`<pre>${esc(st.command)}</pre>`:''}${st.rollback?`<div class="rb">Rollback:<pre>${esc(st.rollback)}</pre></div>`:''}</td><td>${esc(st.owner)}</td><td>${esc(st.planned)}</td><td>${esc(st.approx_min===''?'':fmtMin(st.approx_min))}</td><td>${esc(st.status)}</td></tr>`).join('')).join('');
  const html=`<!DOCTYPE html><html><head><meta charset="utf-8"><title>${esc(t.name)}</title><style>body{font:12px Arial,sans-serif;color:#111;margin:24px}h1{font-size:18px;margin:0 0 4px}
    .meta{color:#444;margin-bottom:12px}table{width:100%;border-collapse:collapse}td,th{border:1px solid #bbb;padding:5px;vertical-align:top;text-align:left}th{background:#eef1f6}
    tr.ph td{background:#dfe8f3;font-weight:bold}pre{margin:4px 0 0;font:11px Consolas,monospace;white-space:pre-wrap;background:#f5f5f5;padding:4px}.rb{color:#a33}</style></head><body>
    <h1>${esc(t.name)}</h1><div class="meta">${esc(t.customer)}, ${esc(t.date)}${t.ticket_id?', ticket '+esc(t.ticket_id):''}${t.change_ref?', change '+esc(t.change_ref):''}<br>
    Window ${esc(t.window_start||'-')} to ${esc(t.window_end||'-')}; projected finish ${esc(t.computed.projected_end)}; total ${fmtMin(t.computed.approx_total)}. Status ${esc(t.status)}. Printed ${esc(nowLocal())} by ${esc(S.me.name)}.</div>
    <table><thead><tr><th>No.</th><th>Step</th><th>Owner</th><th>Planned</th><th>Approx</th><th>Status</th></tr></thead><tbody>${rows}</tbody></table></body></html>`;
  const w=window.open('','_blank');if(!w)return toast('Allow pop-ups for this site to print.','error');
  w.document.write(html);w.document.close();w.focus();setTimeout(()=>w.print(),300);
}

/* ---------- admin ---------- */
const ADMIN_TABS=[['users','Users and roles'],['sessions','Active sessions'],['recordings','Terminal recordings'],['hosts','Trusted SSH hosts'],['settings','Settings'],['health','Backups and health']];
async function renderAdmin(el){
  el.innerHTML=`<h2>Admin</h2><p class="sub">Only administrators can see this area. Every change here is written to the audit trail.</p>
    <div class="subtabs" role="tablist">${ADMIN_TABS.map(([k,l])=>`<button class="${S.adminTab===k?'on':''}" data-atab="${k}" role="tab">${esc(l)}</button>`).join('')}</div><div id="adminHost"><p class="st">Loading…</p></div>`;
  el.querySelector('.subtabs').addEventListener('click',e=>{const b=e.target.closest('[data-atab]');if(b){S.adminTab=b.dataset.atab;renderAdmin(el);}});
  const host=$('#adminHost');
  try{await ADMIN[S.adminTab](host);}catch(e){host.innerHTML=`<div class="empty">${esc(e.message)}</div>`;}
}
const AUTH_LABEL={local:'Password',entra:'Microsoft 365',both:'Password or Microsoft 365'};
function userModal(u,roles){
  const isNew=!u;u=u||{username:'',name:'',email:'',role:'Engineer',auth:'local'};
  openModal({title:isNew?'Add user':'Edit '+u.username,save:isNew?'Add user':'Save user',
    body:fieldHtml({key:'username',label:'User name',value:u.username,required:true,hint:isNew?'Letters, numbers, dot, dash, underscore or @.':'User names cannot be changed.'})
      +fieldHtml({key:'name',label:'Full name',value:u.name})+fieldHtml({key:'email',label:'Email / Microsoft 365 UPN',value:u.email,full:true})
      +fieldHtml({key:'role',label:'Role',type:'select',options:roles,value:u.role})+fieldHtml({key:'auth',label:'Signs in with',type:'select',options:['local','entra','both'],value:u.auth})
      +(isNew?fieldHtml({key:'password',label:'Password (optional)',type:'password',full:true,hint:'Leave empty to generate a temporary password that must be changed at first sign-in.'}):'')
      +`<div class="full st" id="roleHelp"></div>`,
    onMount:ov=>{if(!isNew)ov.querySelector('[data-field="username"]').disabled=true;
      [...ov.querySelector('[data-field="auth"]').options].forEach(o=>o.textContent=AUTH_LABEL[o.value]||'—');
      const rs=ov.querySelector('[data-field="role"]'),help=()=>{ov.querySelector('#roleHelp').textContent=S.roleHelp[rs.value]||'';};rs.addEventListener('change',help);help();},
    onSave:async ov=>{const v=readFields(ov);
      if(isNew){const r=await api('POST','/api/admin/users',v);
        if(r.temp_password)setTimeout(()=>showTempPassword(r.user.username,r.temp_password),0);else toast('User added','success');}
      else{await api('PUT','/api/admin/users/'+encodeURIComponent(u.username),{name:v.name,email:v.email,role:v.role,auth:v.auth});toast('User saved','success');}
      ADMIN.users($('#adminHost'));}});
}
function showTempPassword(user,pw){
  openModal({title:'Temporary password for '+user,body:`<p class="full">Give this password to <b>${esc(user)}</b> in a private message. They must choose their own password the first time they sign in. It is shown only once.</p>
    <pre class="cmd full" style="font-size:16px">${esc(pw)}<button class="btn xs" data-cp>Copy</button></pre>`,
    onMount:ov=>ov.querySelector('[data-cp]').addEventListener('click',()=>copyText(pw))});
}
const ADMIN={
  users:async host=>{
    const d=await api('GET','/api/admin/users');S.roleHelp=d.role_help;
    host.innerHTML=`<div class="toolbar"><button class="btn primary" data-au="add">Add user</button><span class="grow"></span></div>
      <div class="tablewrap" style="max-height:none"><table class="grid"><thead><tr><th>User</th><th>Name</th><th>Email / UPN</th><th>Role</th><th>Signs in with</th><th>Status</th><th>Last sign-in</th><th></th></tr></thead><tbody>
      ${d.users.map(u=>`<tr><td><b>${esc(u.username)}</b>${u.username===S.me.username?' <span class="chip">you</span>':''}</td><td>${esc(u.name)}</td><td>${esc(u.email)}</td>
        <td><select data-urole="${esc(u.username)}" aria-label="Role of ${esc(u.username)}">${d.roles.map(r=>`<option ${r===u.role?'selected':''}>${r}</option>`).join('')}</select></td>
        <td>${esc(AUTH_LABEL[u.auth]||u.auth)}${u.must_change?' <span class="chip">must change password</span>':''}</td>
        <td>${u.disabled?'<span class="bad">Disabled</span>':'<span class="st done">Active</span>'}</td><td class="st">${esc(fmtTime(u.last_login))||'never'}</td>
        <td class="actions"><button class="btn xs" data-au="edit" data-u="${esc(u.username)}">Edit</button>${u.auth!=='entra'?`<button class="btn xs" data-au="reset" data-u="${esc(u.username)}">Reset password</button>`:''}
          <button class="btn xs ghost" data-au="toggle" data-u="${esc(u.username)}">${u.disabled?'Enable':'Disable'}</button><button class="btn xs ghost" data-au="del" data-u="${esc(u.username)}">Delete</button></td></tr>`).join('')}
      </tbody></table></div>
      <div class="cols" style="margin-top:16px">${d.roles.map(r=>`<section class="panel"><h3>${esc(r)}</h3><p class="st" style="margin:0">${esc(d.role_help[r])}</p></section>`).join('')}</div>`;
    host.onchange=async e=>{const s=e.target.closest('[data-urole]');if(!s)return;
      try{await api('PUT','/api/admin/users/'+encodeURIComponent(s.dataset.urole),{role:s.value});toast('Role changed','success');}catch(err){toast(err.message,'error');}ADMIN.users(host);};
    host.onclick=async e=>{const b=e.target.closest('[data-au]');if(!b)return;const u=d.users.find(x=>x.username===b.dataset.u),un=encodeURIComponent(b.dataset.u||'');
      try{
        if(b.dataset.au==='add')return userModal(null,d.roles);
        if(b.dataset.au==='edit')return userModal(u,d.roles);
        if(b.dataset.au==='reset'){if(!await confirmBox(`Reset the password of ${u.username}? They are signed out and must choose a new password.`,'Reset password'))return;
          const r=await api('POST',`/api/admin/users/${un}/reset`,{});showTempPassword(u.username,r.temp_password);}
        if(b.dataset.au==='toggle'){await api('PUT',`/api/admin/users/${un}`,{disabled:!u.disabled});toast(u.disabled?'User enabled':'User disabled and signed out','success');}
        if(b.dataset.au==='del'){if(!await confirmBox(`Delete ${u.username}? Their past changes stay in the audit trail.`))return;await api('DELETE',`/api/admin/users/${un}`);toast('User deleted','success');}
        ADMIN.users(host);
      }catch(err){toast(err.message,'error');}};
  },
  sessions:async host=>{
    const d=await api('GET','/api/admin/sessions');
    host.innerHTML=`<h3 style="margin-bottom:8px">Signed-in browsers</h3><div class="tablewrap" style="max-height:none;margin-bottom:18px"><table class="grid"><thead><tr><th>User</th><th>From</th><th>Method</th><th>Since</th><th>Idle</th><th></th></tr></thead><tbody>
      ${d.web.map(s=>`<tr><td>${esc(s.user)}</td><td>${esc(s.ip)}</td><td>${esc(s.method)}</td><td class="st">${esc(fmtTime(s.created))}</td><td class="st">${fmtMin(Math.floor(s.idle_seconds/60))}</td>
        <td class="actions"><button class="btn xs" data-ks="${esc(s.handle)}">Sign out</button></td></tr>`).join('')}</tbody></table></div>
      <h3 style="margin-bottom:8px">Live terminal sessions</h3>${d.terminals.length?`<div class="tablewrap" style="max-height:none"><table class="grid"><thead><tr><th>User</th><th>Target</th><th>State</th><th>Started</th><th>From</th><th></th></tr></thead><tbody>
      ${d.terminals.map(t=>`<tr><td>${esc(t.user)}</td><td>${esc(t.username)}@${esc(t.host)}:${t.port}</td><td><span class="stb ${esc(t.state)}">${esc(t.state)}</span></td><td class="st">${esc(fmtTime(t.started))}</td><td>${esc(t.ip)}</td>
        <td class="actions"><button class="btn xs danger" data-kt="${esc(t.sid)}">End session</button></td></tr>`).join('')}</tbody></table></div>`:'<p class="st">No terminal sessions are running.</p>'}`;
    host.onclick=async e=>{const s=e.target.closest('[data-ks]'),t=e.target.closest('[data-kt]');
      try{if(s){await api('DELETE','/api/admin/sessions/'+encodeURIComponent(s.dataset.ks));toast('Signed out','success');}
        if(t){if(!await confirmBox('End this terminal session now? The user is disconnected.','End session'))return;await api('POST',`/api/admin/terminals/${encodeURIComponent(t.dataset.kt)}/kill`,{});toast('Session ended','success');}
        if(s||t)ADMIN.sessions(host);}catch(err){toast(err.message,'error');}};
  },
  recordings:async host=>{
    const d=await api('GET','/api/admin/recordings');
    host.innerHTML=`<div class="toolbar"><div class="grow"><input type="search" id="recQ" placeholder="Filter by user, host, type or date" aria-label="Filter recordings"></div></div><div id="recList"></div>
      <p class="st" style="margin-top:8px">SSH transcripts are stored in the hidden .recordings folder of the central storage. Remote Desktop app sessions are logged but cannot be recorded; Guacamole sessions are recorded on the Guacamole server.</p>`;
    const draw=()=>{const q=($('#recQ').value||'').toLowerCase();const list=d.sessions.filter(s=>!q||Object.values(s).some(v=>String(v).toLowerCase().includes(q)));
      $('#recList').innerHTML=list.length?`<div class="tablewrap"><table class="grid"><thead><tr><th>Started</th><th>Type</th><th>User</th><th>Target</th><th>Status</th><th>Ended</th><th>Size</th><th></th></tr></thead><tbody>
        ${list.map(s=>`<tr><td class="st">${esc(fmtTime(s.started))}</td><td>${esc(s.type)}</td><td>${esc(s.user)} <span class="st">(${esc(s.role)}, ${esc(s.ip)})</span></td><td>${esc(s.username)}@${esc(s.host)}:${s.port}${s.asset?` <span class="st">${esc(s.asset)}</span>`:''}</td>
          <td><span class="stb ${esc(s.status)}">${esc(s.status)}</span>${s.note?`<div class="st">${esc(s.note)}</div>`:''}</td><td class="st">${esc(fmtTime(s.ended))}</td><td class="st">${s.bytes?Math.ceil(s.bytes/1024)+' KB':''}</td>
          <td class="actions">${s.file?`<button class="btn xs" data-rv="${esc(s.id)}">View</button><button class="btn xs ghost" data-rd="${esc(s.id)}">Download</button>`:''}</td></tr>`).join('')}</tbody></table></div>`:'<div class="tablewrap"><div class="empty">No terminal sessions recorded yet.</div></div>';};
    draw();$('#recQ').addEventListener('input',draw);
    host.onclick=async e=>{const v=e.target.closest('[data-rv]'),dl=e.target.closest('[data-rd]');
      if(dl)return download(`/api/admin/recordings/${encodeURIComponent(dl.dataset.rd)}/download`);
      if(!v)return;try{const r=await api('GET',`/api/admin/recordings/${encodeURIComponent(v.dataset.rv)}`);showRecording(r);}catch(err){toast(err.message,'error');}};
  },
  hosts:async host=>{
    const d=await api('GET','/api/admin/known-hosts');
    host.innerHTML=`<p class="sub">SSH host keys trusted on first connection. If a server is rebuilt its key changes and connections are blocked until you remove the old entry here (after checking the new fingerprint with the server owner).</p>
      ${d.hosts.length?`<div class="tablewrap" style="max-height:none"><table class="grid"><thead><tr><th>Host</th><th>Key type</th><th>Fingerprint</th><th>Trusted by</th><th>When</th><th></th></tr></thead><tbody>
      ${d.hosts.map(h=>`<tr><td>${esc(h.host)}</td><td>${esc(h.type)}</td><td class="mono">${esc(h.fingerprint)}</td><td>${esc(h.added_by)}</td><td class="st">${esc(fmtTime(h.added_at))}</td>
        <td class="actions"><button class="btn xs danger" data-hk="${esc(h.host)}">Remove</button></td></tr>`).join('')}</tbody></table></div>`:'<div class="tablewrap"><div class="empty">No SSH hosts trusted yet.</div></div>'}`;
    host.onclick=async e=>{const b=e.target.closest('[data-hk]');if(!b)return;if(!await confirmBox(`Remove the trusted key of ${b.dataset.hk}? The next connection asks to trust its key again.`,'Remove'))return;
      try{await api('DELETE','/api/admin/known-hosts',{host:b.dataset.hk});toast('Removed','success');ADMIN.hosts(host);}catch(err){toast(err.message,'error');}};
  },
  settings:async host=>{
    const d=await api('GET','/api/admin/settings'),s=d.settings;
    const F=[['session_idle_minutes','Sign out after this many idle minutes','number'],['session_max_hours','Maximum session length (hours)','number'],
      ['lockout_attempts','Failed sign-ins before lockout','number'],['lockout_minutes','Lockout duration (minutes)','number'],
      ['entra_default_role','Role for new Microsoft 365 users','role','Empty = only users an admin has added can sign in.'],
      ['terminal_restrict_to_inventory','Terminals only to CIs in the CMDB','bool'],
      ['terminal_allowed_networks','Also allow these hosts / networks','text','Comma separated, e.g. 10.20.0.0/16, jump01.contoso.local'],
      ['terminal_max_per_user','Open terminals per person','number'],['terminal_idle_minutes','Close idle terminals after (minutes)','number'],
      ['terminal_record','Record terminal sessions','bool'],['auto_backup','Daily automatic backup','bool'],['backup_keep','Backups to keep','number'],
      ['ai_mode','Assistant','sel','off: no Ask button. search: answers built from the records. model: a model on your own network words the answers.',['off','search','model']],
      ['ai_api','Model server type','sel','ollama for Ollama; openai for llama.cpp, vLLM, LM Studio or other OpenAI-compatible servers.',['ollama','openai']],
      ['ai_endpoint','Model server address','text','This server or your private network only, e.g. http://127.0.0.1:11434. Public addresses are refused.'],
      ['ai_model','Model name','text','As your model server names it, e.g. llama3.2:3b or qwen2.5:7b-instruct.'],
      ['ai_timeout_seconds','Model time limit (seconds)','number','After this, the answer is built from the records instead.']];
    host.innerHTML=`<div class="panel" style="max-width:780px"><div class="body" style="display:grid;grid-template-columns:1fr 1fr;gap:12px">${F.map(([k,l,t,h])=>
      t==='bool'?fieldHtml({key:k,type:'checkbox',label:l,text:'On',value:s[k],hint:h}):t==='role'?fieldHtml({key:k,label:l,type:'select',options:S.roles,value:s[k],hint:h})
      :t==='sel'?fieldHtml({key:k,label:l,type:'select',options:F.find(x=>x[0]===k)[4],value:s[k],hint:h})
      :fieldHtml({key:k,label:l,value:String(s[k]),hint:h,full:t==='text'})).join('')}</div><div class="toolbar" style="margin:14px 0 0"><button class="btn primary" data-save>Save settings</button><button class="btn" data-aitest>Test the model server</button><span class="st" data-aires></span></div></div>`;
    host.onclick=async e=>{
      if(e.target.closest('[data-aitest]')){const out=host.querySelector('[data-aires]');out.textContent='Testing…';
        const v=readFields(host);v.ai_timeout_seconds=parseInt(v.ai_timeout_seconds,10)||60;
        try{const r=(await api('POST','/api/admin/ai/test',{settings:v})).status;out.textContent=r.detail+(r.models&&r.models.length?' Models on that server: '+r.models.join(', '):'');}
        catch(err){out.textContent=err.message;}return;}
      if(!e.target.closest('[data-save]'))return;const v=readFields(host);F.forEach(([k,,t])=>{if(t==='number')v[k]=parseInt(v[k],10);});
      try{await api('PUT','/api/admin/settings',{settings:v});toast('Settings saved','success');ADMIN.settings(host);}catch(err){toast(err.message,'error');}};
  },
  health:async host=>{
    const d=await api('GET','/api/admin/health'),c=d.checks;
    const chk=(ok,t,s)=>`<div class="check"><i class="${ok?'ok':'no'}"></i><div><b>${esc(t)}</b><div class="st">${esc(s)}</div></div></div>`;
    host.innerHTML=`<div class="checks">${chk(c.https,'HTTPS',c.https?'Traffic is encrypted.':'Not encrypted. Set tls_cert / tls_key or use an HTTPS proxy before team use.')}
      ${chk(c.ssh_terminals,'SSH terminals',c.ssh_terminals?'Ready.':'Install the paramiko package.')}${chk(c.guacamole,'In-browser RDP',c.guacamole?(c.guacamole_recording?'Guacamole ready, recording on.':'Guacamole ready, recording path not set.'):'Guacamole not configured.')}
      ${chk(c.microsoft_signin||c.proxy_identity,'Microsoft 365 sign-in',c.microsoft_signin?'Enabled.':c.proxy_identity?'Through the proxy / App Service.':'Off. Users sign in with passwords.')}
      ${chk(!d.status.last_error,'Central storage',d.status.last_error||`${c.storage.label}: ${c.storage.location}`)}${chk(c.assistant,'Assistant',c.assistant?(c.assistant_mode==='off'?'Installed, switched off in Settings.':c.assistant_mode==='model'?'On, using a model on your own network.':'On, answering from the records (no model needed).'):'Not installed: put msp_ai.py next to msp_runbook.py.')}${chk(c.local_assets,'Offline assets',c.local_assets?'Fonts and the terminal load from this server; nothing is fetched from the internet.':'The static folder is missing: copy it next to msp_runbook.py.')}${chk(c.roster,'Shift roster',c.roster?(c.roster_holidays?'Ready, with US and India public holidays.':'Ready. Region holidays need: pip install holidays.'):'Not installed: put msp_roster.py next to msp_runbook.py.')}${chk(true,'Activity',`${c.web_sessions} signed-in browser(s), ${c.live_terminals} live terminal(s). Python ${c.python}.`)}</div>
      <div class="toolbar"><h3 style="margin-right:auto">Backups</h3><button class="btn primary" data-bk>Back up now</button></div>
      ${d.backups.length?`<div class="tablewrap" style="max-height:none"><table class="grid"><thead><tr><th>When</th><th>File</th><th>By</th><th>Size</th></tr></thead><tbody>${d.backups.map(b=>`<tr><td class="st">${esc(fmtTime(b.at))}</td><td class="mono">${esc(b.file)}</td><td>${esc(b.by)}${b.auto?' (automatic)':''}</td><td class="st">${Math.ceil(b.bytes/1024)} KB</td></tr>`).join('')}</tbody></table></div>`:'<p class="st">No backups yet. A backup is taken automatically once a day while the server runs.</p>'}`;
    host.onclick=async e=>{if(!e.target.closest('[data-bk]'))return;try{const r=await api('POST','/api/admin/backup',{});toast('Backup saved: '+r.backup.file,'success');ADMIN.health(host);}catch(err){toast(err.message,'error');}};
  },
};
function showRecording(r){
  const s=r.session,lines=r.content.split('\n');
  openModal({title:`Recording: ${s.user} on ${s.username}@${s.host}`,size:'full',
    body:`<div class="full toolbar" style="margin:0"><input type="search" id="rq" placeholder="Show only lines containing…" aria-label="Filter lines" style="flex:1">
      <span class="st">${s.started?esc(fmtTime(s.started)):''} to ${s.ended?esc(fmtTime(s.ended)):'still running'}. Integrity: ${r.integrity_ok===true?'<span class="st done">verified (SHA-256 matches)</span>':r.integrity_ok===false?'<span class="bad">CHANGED since the session ended</span>':'not yet sealed'}</span></div>
      <pre class="rec full" id="rtext"></pre>`,
    footExtra:`<button class="btn" data-dl>Download</button>`,
    onMount:ov=>{const pre=ov.querySelector('#rtext'),q=ov.querySelector('#rq');
      const draw=()=>{const t=q.value.toLowerCase();const shown=t?lines.filter(l=>l.toLowerCase().includes(t)):lines;
        pre.innerHTML=shown.map(l=>t?esc(l).replace(new RegExp(esc(t).replace(/[.*+?^${}()|[\]\\]/g,'\\$&'),'gi'),m=>`<mark>${m}</mark>`):esc(l)).join('\n');};
      q.addEventListener('input',draw);draw();ov.querySelector('[data-dl]').addEventListener('click',()=>download(`/api/admin/recordings/${encodeURIComponent(s.id)}/download`));}});
}
boot();
</script>
</body>
</html>
"""


if __name__ == "__main__":
    sys.exit(main())
