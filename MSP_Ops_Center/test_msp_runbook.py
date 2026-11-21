#!/usr/bin/env python3
"""
Automated test suite for msp_runbook.py (version 3)

Starts real HTTP servers on random ports:
  * the application with local storage in a temporary folder, and
  * the application with SharePoint storage against an in-process fake Microsoft Graph.

Run:  python -m unittest -v test_msp_runbook.py
"""
import csv
import hashlib
from http.client import HTTPConnection
import io
import json
import os
import re
import shutil
import socket
import struct
import base64
import datetime as dt
import sys
import tempfile
import threading
import unittest
import urllib.error
import urllib.parse
import urllib.request
import time
import uuid
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

import openpyxl

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import msp_runbook as app  # noqa: E402


# --------------------------------------------------------------------------------------
# helpers: a signed-in HTTP client per test user (session cookie + CSRF token)
# --------------------------------------------------------------------------------------
app.PBKDF2_ITERATIONS = 1000          # fast hashing for tests only
ADMIN_PW = "Adm1n-Passw0rd!"
USER_PW = "Corr3ct-Horse-Battery"
CLIENTS = {}


def slug(name):
    return re.sub(r"[^a-z0-9]+", ".", name.lower()).strip(".") or "user"


class Client:
    def __init__(self, base):
        u = urllib.parse.urlsplit(base)
        self.base, self.host, self.port = base, u.hostname, u.port
        self.cookie, self.csrf = "", ""

    def request(self, method, path, body=None, headers=None, ctype=None):
        data, hdrs = None, {}
        if isinstance(body, (dict, list)):
            data = json.dumps(body).encode()
            hdrs["Content-Type"] = "application/json"
        elif isinstance(body, bytes):
            data = body
            if ctype:
                hdrs["Content-Type"] = ctype
        if self.cookie:
            hdrs["Cookie"] = f"msp_session={self.cookie}"
        if method not in ("GET", "HEAD") and self.csrf:
            hdrs["X-CSRF-Token"] = self.csrf
        hdrs.update(headers or {})
        conn = HTTPConnection(self.host, self.port, timeout=30)
        try:
            conn.request(method, path, body=data, headers=hdrs)
            resp = conn.getresponse()
            raw = resp.read()
            for c in resp.headers.get_all("Set-Cookie") or []:
                m = re.match(r"msp_session=([^;]*)", c)
                if m:
                    self.cookie = m.group(1)
            return resp.status, resp.headers, raw
        finally:
            conn.close()

    def json(self, method, path, body=None, headers=None):
        st, _, raw = self.request(method, path, body, headers)
        return st, json.loads(raw.decode() or "{}")

    def login(self, username, password):
        st, d = self.json("POST", "/api/auth/login", {"username": username, "password": password})
        if st == 200:
            self.csrf = d["csrf"]
        return st, d


def setup_admin(base, server):
    c = Client(base)
    st, d = c.json("POST", "/api/auth/setup", {"code": server.setup_code, "username": "admin",
                                               "name": "Test Engineer", "password": ADMIN_PW})
    assert st == 200, d
    c.csrf = d["csrf"]
    CLIENTS[(base, "Test Engineer")] = c
    return c


def client_for(base, name="Test Engineer", role="Lead"):
    key = (base, name)
    if key in CLIENTS:
        return CLIENTS[key]
    admin = CLIENTS[(base, "Test Engineer")]
    uname = slug(name)
    st, d = admin.json("POST", "/api/admin/users", {"username": uname, "name": name, "role": role, "password": USER_PW})
    assert st in (200, 409), d
    c = Client(base)
    st, d = c.login(uname, USER_PW)
    assert st == 200, d
    CLIENTS[key] = c
    return c


def http(base, method, path, body=None, headers=None, ctype=None, client=None):
    c = client or CLIENTS.get((base, "Test Engineer")) or Client(base)
    return c.request(method, path, body, headers, ctype)


def jcall(base, method, path, body=None, user="Test Engineer", headers=None):
    st, _, raw = client_for(base, user).request(method, path, body, headers)
    return st, json.loads(raw.decode() or "{}")


def multipart(fields, files):
    boundary = "----msptest" + uuid.uuid4().hex
    out = io.BytesIO()
    for k, v in fields.items():
        out.write(f"--{boundary}\r\nContent-Disposition: form-data; name=\"{k}\"\r\n\r\n{v}\r\n".encode())
    for field, (filename, data) in files.items():
        out.write(f"--{boundary}\r\nContent-Disposition: form-data; name=\"{field}\"; filename=\"{filename}\"\r\n"
                  f"Content-Type: application/octet-stream\r\n\r\n".encode())
        out.write(data)
        out.write(b"\r\n")
    out.write(f"--{boundary}--\r\n".encode())
    return f"multipart/form-data; boundary={boundary}", out.getvalue()


def start(server):
    threading.Thread(target=server.serve_forever, daemon=True).start()
    return f"http://127.0.0.1:{server.server_address[1]}"


def stop(server):
    server.shutdown()
    server.server_close()


def xlsx_bytes(sheets):
    """sheets: {title: [rows...]} -> workbook bytes"""
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for title, rows in sheets.items():
        ws = wb.create_sheet(title)
        for r in rows:
            ws.append(r)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


class LocalServerTestCase(unittest.TestCase):
    config_overrides = {}

    @classmethod
    def setUpClass(cls):
        cls.tmp = tempfile.mkdtemp(prefix="msp_test_")
        opts = dict(host="127.0.0.1", port=0, storage="local", data_dir=cls.tmp, refresh_interval=0, quiet=True,
                    seed=True, backup_thread=False, setup_code_file=os.path.join(cls.tmp, "setup_code.txt"))
        opts.update(cls.config_overrides)
        cls.server = app.create_server(app.Config(**opts))
        cls.base = start(cls.server)
        cls.admin = setup_admin(cls.base, cls.server)

    @classmethod
    def tearDownClass(cls):
        stop(cls.server)
        shutil.rmtree(cls.tmp, ignore_errors=True)

    def api(self, method, path, body=None, user="Test Engineer", expect=200):
        st, data = jcall(self.base, method, path, body, user)
        self.assertEqual(st, expect, f"{method} {path} -> {st}: {data}")
        return data

    def op(self, op, expect=200, user="Test Engineer", **body):
        return self.api("POST", "/api/plans/op", {"op": op, **body}, user=user, expect=expect)

    def task(self, tid):
        return next(t for t in self.api("GET", "/api/plans")["plans"]["tasks"] if t["id"] == tid)

    def table(self, sheet):
        return self.api("GET", f"/api/table/{sheet}")["table"]

    def upload(self, filename, data, module="auto", expect=200):
        ctype, body = multipart({"module": module}, {"file": (filename, data)})
        st, _, raw = http(self.base, "POST", "/api/import/preview", body, None, ctype,
                          client=client_for(self.base, "Importer"))
        out = json.loads(raw)
        self.assertEqual(st, expect, out)
        return out


# --------------------------------------------------------------------------------------
class T01ServerBasics(LocalServerTestCase):
    def test_index_and_health(self):
        st, hdrs, body = http(self.base, "GET", "/")
        self.assertEqual(st, 200)
        self.assertIn("text/html", hdrs["Content-Type"])
        self.assertIn("<title>MSP Ops Center</title>", body.decode())
        self.assertEqual(http(self.base, "GET", "/healthz")[0], 200)
        status = self.api("GET", "/api/status")
        self.assertEqual(status["storage"]["kind"], "local")
        for name in (app.XLSX_NAME, app.PLANS_NAME, app.JSON_NAME, app.META_NAME, app.USERS_NAME):
            self.assertTrue(os.path.exists(os.path.join(self.tmp, name)), name)

    def test_cache_headers_everywhere(self):
        for path in ("/", "/api/status?_ts=1", "/api/state", "/api/missing", "/api/template/assets?format=csv"):
            _, hdrs, _ = http(self.base, "GET", path)
            self.assertEqual(hdrs["Cache-Control"], "no-cache, no-store, must-revalidate, max-age=0", path)
            self.assertEqual(hdrs["Pragma"], "no-cache", path)
            self.assertEqual(hdrs["Expires"], "0", path)
        self.assertIn("'_ts='+Date.now()", http(self.base, "GET", "/")[2].decode())

    def test_state_payload(self):
        s = self.api("GET", "/api/state")
        for k in ("tables", "plans", "hierarchy", "audit", "revision", "my_reads", "imports", "schema"):
            self.assertIn(k, s)
        self.assertEqual(set(s["tables"]), set(app.SHEET_ORDER))
        self.assertTrue(s["plans"]["tasks"])  # seeded plan


# --------------------------------------------------------------------------------------
class T02ExecutionPlans(LocalServerTestCase):
    def new_task(self, **over):
        task = {"date": "2026-12-06", "customer": "Initech", "name": "SAN firmware upgrade", "ticket_id": "CHG-9001",
                "window_start": "2026-12-06 22:00", "window_end": "2026-12-07 01:00", **over}
        return self.op("task.create", task=task, template="change")["task_id"]

    def test_numbering_chained_times_and_overrides(self):
        tid = self.new_task()
        t = self.task(tid)
        self.assertEqual([p["name"] for p in t["phases"]], ["Pre-execution", "During outage", "Post execution"])
        self.assertEqual([p["num"] for p in t["phases"]], ["1", "2", "3"])
        p1, p2 = t["phases"][0]["id"], t["phases"][1]["id"]
        s1 = self.op("step.add", task_id=tid, phase_id=p1, step={"title": "Backup config", "approx_min": "30"})["step_id"]
        s2 = self.op("step.add", task_id=tid, phase_id=p1, after_step_id=s1,
                     step={"Task / step": "Health check", "Approx (min)": "15", "command": "show health"})["step_id"]
        s21 = self.op("step.add", task_id=tid, phase_id=p1, after_step_id=s2, as_child=True,
                      step={"title": "Check paths", "approx_min": "1h"})["step_id"]
        self.op("step.add", task_id=tid, phase_id=p2, step={"title": "Upgrade controller A", "approx_min": "1:30"})
        t = self.task(tid)
        steps1 = t["phases"][0]["steps"]
        self.assertEqual([s["num"] for s in steps1], ["1.1", "1.2", "1.2.1"])
        self.assertEqual(t["phases"][1]["steps"][0]["num"], "2.1")
        self.assertEqual([s["planned"] for s in steps1], ["2026-12-06 22:00", "2026-12-06 22:30", "2026-12-06 22:45"])
        self.assertTrue(all(s["planned_inherited"] for s in steps1))
        self.assertEqual(t["phases"][1]["steps"][0]["planned"], "2026-12-06 23:45")
        self.assertEqual(t["phases"][1]["steps"][0]["approx_min"], 90)
        self.assertEqual(t["computed"]["approx_total"], 195)
        self.assertEqual(t["computed"]["projected_end"], "2026-12-07 01:15")
        self.assertEqual(t["computed"]["overrun_min"], 15)
        # override one line: the following lines continue from it
        self.op("step.update", task_id=tid, phase_id=p1, step_id=s2, step={"planned_start": "2026-12-06T23:00"})
        t = self.task(tid)
        st = t["phases"][0]["steps"]
        self.assertEqual((st[1]["planned"], st[1]["planned_inherited"]), ("2026-12-06 23:00", False))
        self.assertEqual(st[2]["planned"], "2026-12-06 23:15")
        self.assertEqual(t["phases"][1]["steps"][0]["planned"], "2026-12-07 00:15")
        self.op("step.update", task_id=tid, phase_id=p1, step_id=s2, step={"planned_start": ""})
        self.assertTrue(self.task(tid)["phases"][0]["steps"][1]["planned_inherited"])
        self.op("step.update", task_id=tid, phase_id=p1, step_id=s2, step={"approx_min": "ten"}, expect=400)
        # move / indent / outdent keep blocks together
        self.op("step.move", task_id=tid, phase_id=p1, step_id=s1, dir="down")
        st = self.task(tid)["phases"][0]["steps"]
        self.assertEqual([s["id"] for s in st], [s2, s21, s1])
        self.assertEqual([s["num"] for s in st], ["1.1", "1.1.1", "1.2"])
        self.op("step.move", task_id=tid, phase_id=p1, step_id=s1, dir="indent")
        self.assertEqual(self.task(tid)["phases"][0]["steps"][2]["num"], "1.1.2")
        self.op("step.move", task_id=tid, phase_id=p1, step_id=s1, dir="outdent")
        self.op("step.move", task_id=tid, phase_id=p1, step_id=s2, dir="up", expect=400)
        self.op("step.move", task_id=tid, phase_id=p1, step_id=s2, dir="outdent", expect=400)
        # clone a block (step + its sub-step)
        self.op("step.clone", task_id=tid, phase_id=p1, step_id=s2)
        st = self.task(tid)["phases"][0]["steps"]
        self.assertEqual([s["num"] for s in st], ["1.1", "1.1.1", "1.2", "1.2.1", "1.3"])
        self.assertEqual(st[2]["title"], "Health check")
        self.assertNotEqual(st[2]["id"], s2)
        # delete a block
        self.op("step.delete", task_id=tid, phase_id=p1, step_id=st[2]["id"])
        self.assertEqual(len(self.task(tid)["phases"][0]["steps"]), 3)

    def test_duplicate_task_detection(self):
        self.new_task(ticket_id="CHG-DUP")
        st, data = jcall(self.base, "POST", "/api/plans/op", {"op": "task.create", "template": "blank", "task": {
            "date": "2026-12-06", "customer": "initech", "name": "Other name", "ticket_id": "chg-dup"}})
        self.assertEqual(st, 409)
        self.assertIn("already exists", data["error"])
        self.op("task.create", template="blank", allow_duplicate=True, task={
            "date": "2026-12-06", "customer": "Initech", "name": "Other", "ticket_id": "CHG-DUP"})
        self.op("task.create", task={"date": "2026-12-06", "name": "x"}, expect=400)  # no customer

    def test_checkpoints_run_actions_and_status(self):
        tid = self.new_task(ticket_id="CHG-RUN")
        t = self.task(tid)
        p1, p2 = t["phases"][0], t["phases"][1]
        self.assertTrue(p1["checkpoint"])
        a = self.op("step.add", task_id=tid, phase_id=p1["id"], step={"title": "Pre-check", "approx_min": 5})["step_id"]
        b = self.op("step.add", task_id=tid, phase_id=p2["id"], step={"title": "Do it", "approx_min": 5})["step_id"]
        self.op("step.action", task_id=tid, phase_id=p1["id"], step_id=a, action="start")
        self.assertEqual(self.task(tid)["status"], "In Progress")
        self.op("step.action", task_id=tid, phase_id=p1["id"], step_id=a, action="done")
        st, data = jcall(self.base, "POST", "/api/plans/op", {"op": "step.action", "task_id": tid,
                                                              "phase_id": p2["id"], "step_id": b, "action": "start"})
        self.assertEqual(st, 409)
        self.assertIn("go/no-go", data["error"].lower())
        self.op("phase.signoff", task_id=tid, phase_id=p1["id"], user="Lead Lee")
        self.assertEqual(self.task(tid)["phases"][0]["signed_off_by"], "lead.lee")
        self.op("step.update", task_id=tid, phase_id=p2["id"], step_id=b,
                step={"actual_start": "2026-12-06 22:10", "actual_end": "2026-12-06 22:22", "status": "Done"})
        sb = self.task(tid)["phases"][1]["steps"][0]
        self.assertEqual((sb["actual_min"], sb["variance_min"]), (12, 7))
        self.op("step.action", task_id=tid, phase_id=p2["id"], step_id=b, action="done")
        self.assertEqual(self.task(tid)["status"], "Completed")
        self.op("step.action", task_id=tid, phase_id=p2["id"], step_id=b, action="reset")
        self.assertEqual(self.task(tid)["status"], "In Progress")
        self.op("phase.signoff", task_id=tid, phase_id=p2["id"], expect=400)  # not a checkpoint

    def test_step_columns_rename_hide_add_remove(self):
        tid = self.new_task(ticket_id="CHG-COLS")
        t = self.task(tid)
        pid = t["phases"][0]["id"]
        sid = self.op("step.add", task_id=tid, phase_id=pid, step={"title": "x", "owner": "Raj"})["step_id"]
        cols = [dict(c) for c in t["step_columns"]]
        for c in cols:
            if c["key"] == "owner":
                c["label"] = "Engineer"
            if c["key"] == "rollback":
                c["hidden"] = True
        cols.append({"label": "Approved by"})
        self.op("task.columns", task_id=tid, columns=cols)
        t = self.task(tid)
        by_key = {c["key"]: c for c in t["step_columns"]}
        self.assertEqual(by_key["owner"]["label"], "Engineer")
        self.assertTrue(by_key["rollback"]["hidden"])
        custom = next(c for c in t["step_columns"] if not c["core"])
        self.op("step.update", task_id=tid, phase_id=pid, step_id=sid, step={"Approved by": "CAB", "Engineer": "Mia"})
        st = self.task(tid)["phases"][0]["steps"][0]
        self.assertEqual((st["custom"][custom["key"]], st["owner"]), ("CAB", "Mia"))
        res = self.op("task.columns", task_id=tid, columns=[{"key": "title", "label": "Step"}])
        self.assertEqual(res["removed"], ["Approved by"])
        t = self.task(tid)
        self.assertEqual(len(t["step_columns"]), len(app.STEP_CORE))  # core columns always come back
        self.assertEqual(t["phases"][0]["steps"][0]["custom"], {})
        self.op("task.columns", task_id=tid, columns=[{"label": "A"}, {"label": "a"}], expect=400)

    def test_phases_clone_task_templates(self):
        tid = self.new_task(ticket_id="CHG-PH", date="2026-12-13", window_start="2026-12-13 21:00",
                            window_end="2026-12-13 23:00")
        t = self.task(tid)
        p1 = t["phases"][0]["id"]
        self.op("step.add", task_id=tid, phase_id=p1, step={"title": "Step A", "approx_min": 10,
                                                            "planned_start": "2026-12-13 21:30"})
        newp = self.op("phase.add", task_id=tid, name="Pre-live", after_phase_id=p1)["phase_id"]
        self.assertEqual([p["name"] for p in self.task(tid)["phases"]][:2], ["Pre-execution", "Pre-live"])
        self.op("phase.move", task_id=tid, phase_id=newp, dir="up")
        self.assertEqual(self.task(tid)["phases"][0]["name"], "Pre-live")
        self.op("phase.move", task_id=tid, phase_id=newp, dir="up", expect=400)
        self.op("phase.clone", task_id=tid, phase_id=p1)
        self.assertIn("Pre-execution (copy)", [p["name"] for p in self.task(tid)["phases"]])
        self.op("phase.update", task_id=tid, phase_id=newp, name="During cut-over", checkpoint=True)
        self.op("phase.delete", task_id=tid, phase_id=newp)
        self.assertNotIn("During cut-over", [p["name"] for p in self.task(tid)["phases"]])
        sid = self.task(tid)["phases"][0]["steps"][0]["id"]
        self.op("step.action", task_id=tid, phase_id=p1, step_id=sid, action="skip")
        cid = self.op("task.clone", task_id=tid, task={"date": "2026-12-20", "ticket_id": "CHG-PH2"})["task_id"]
        c = self.task(cid)
        self.assertEqual((c["date"], c["ticket_id"], c["status"]), ("2026-12-20", "CHG-PH2", "Draft"))
        self.assertEqual((c["window_start"], c["window_end"]), ("2026-12-20 21:00", "2026-12-20 23:00"))
        st = c["phases"][0]["steps"][0]
        self.assertEqual((st["status"], st["planned_start"], st["title"]), ("Pending", "2026-12-20 21:30", "Step A"))
        self.assertNotEqual(st["id"], sid)
        same = self.op("task.clone", task_id=tid)["task_id"]
        self.assertTrue(self.task(same)["name"].endswith("(copy)"))
        tpl = self.op("template.save", task_id=tid, name="SAN upgrade")["template_id"]
        nid = self.op("task.create", template=tpl, task={"date": "2027-01-10", "customer": "Initech",
                                                         "name": "SAN upgrade site B"})["task_id"]
        n = self.task(nid)
        self.assertEqual((n["phases"][0]["steps"][0]["title"], n["phases"][0]["steps"][0]["status"]),
                         ("Step A", "Pending"))
        self.op("template.delete", template_id=tpl)
        self.op("task.create", template=tpl, task={"date": "2027-01-11", "customer": "x", "name": "y"}, expect=404)

    def test_plan_history_restore_and_delete(self):
        tid = self.new_task(ticket_id="CHG-HIST", name="Original name")
        self.op("task.update", task_id=tid, task={"name": "Renamed"})
        hist = self.api("GET", f"/api/history?key=plan/{tid}")["versions"]
        self.assertEqual(hist[0]["action"], "Update task")
        self.assertIn({"field": "name", "old": "Original name", "new": "Renamed"}, hist[0]["changes"])
        self.api("POST", "/api/history/restore", {"key": f"plan/{tid}", "v": hist[-1]["v"]})
        self.assertEqual(self.task(tid)["name"], "Original name")
        self.op("task.delete", task_id=tid)
        self.assertFalse(any(t["id"] == tid for t in self.api("GET", "/api/plans")["plans"]["tasks"]))
        last = self.api("GET", f"/api/history?key=plan/{tid}")["versions"][0]
        self.assertTrue(last["deleted"])
        self.api("POST", "/api/history/restore", {"key": f"plan/{tid}", "v": last["v"]})
        self.assertEqual(self.task(tid)["name"], "Original name")


# --------------------------------------------------------------------------------------
class T03TablesRunbookBulletins(LocalServerTestCase):
    def test_asset_crud_labels_duplicates_clone(self):
        row = self.api("POST", "/api/table/assets", {"row": {
            "Customer": "Umbrella", "Serial No": "SN-100", "Device": "umb-iis01", "Platform type": "vmware",
            "OS": "Windows Server", "OS version": "2019", "CPU": "4", "Memory": "16 GB", "Warranty end": "31/03/2027"}})["row"]
        self.assertEqual((row["Serial_No"], row["Platform_Type"], row["Warranty_End"], row["OS_Version"], row["Version"]),
                         ("SN-100", "Virtualization", "2027-03-31", "2019", "1"))
        rid = row["Record_ID"]
        st, data = jcall(self.base, "POST", "/api/table/assets", {"row": {"Device": "other", "Serial_No": "sn100"}})
        self.assertEqual(st, 409)
        self.assertIn("duplicate", data["error"].lower())
        self.api("POST", "/api/table/assets", {"row": {"Device": "umb-iis01", "Customer": "Umbrella"}}, expect=409)
        self.api("POST", "/api/table/assets", {"row": {"Device": "other", "Serial_No": "SN-100"},
                                               "allow_duplicate": True})
        # editing a key field into a duplicate is caught; editing other fields is not blocked
        other = next(r for r in self.table("Asset_Master")["rows"] if r["Device"] == "other")
        self.api("PUT", f"/api/table/assets/{other['Record_ID']}", {"row": {"Device": "umb-iis01", "Serial_No": "",
                                                                             "Customer": "Umbrella"}}, expect=409)
        upd = self.api("PUT", f"/api/table/assets/{rid}", {"row": {"Memory": "32 GB"}})["row"]
        self.assertEqual((upd["Memory"], upd["Version"], upd["CPU"]), ("32 GB", "2", "4"))
        clone = self.api("POST", f"/api/table/assets/{rid}/clone", {})["row"]
        self.assertEqual((clone["Serial_No"], clone["Device"], clone["Memory"]), ("", "umb-iis01 (copy)", "32 GB"))
        self.assertNotEqual(clone["Record_ID"], rid)
        clone2 = self.api("POST", f"/api/table/assets/{rid}/clone", {"row": {"Device": "umb-iis02"}})["row"]
        self.assertEqual(clone2["Device"], "umb-iis02")
        self.api("POST", "/api/table/assets", {"row": {"Customer": "x"}}, expect=400)
        self.api("POST", "/api/table/assets", {"row": {"Device": "d", "Warranty_End": "soon"}}, expect=400)

    def test_table_columns_persist_with_renamed_headers(self):
        t = self.table("Asset_Master")
        cols = [dict(c) for c in t["columns"] if not c["system"]]
        for c in cols:
            if c["key"] == "Device":
                c["label"] = "Server name"
            if c["key"] == "Role":
                c["hidden"] = False
        cols.append({"label": "Licence key"})
        self.api("PUT", "/api/table/Asset_Master/columns", {"columns": cols})
        t = self.table("Asset_Master")
        self.assertEqual(next(c for c in t["columns"] if c["key"] == "Device")["label"], "Server name")
        lic = next(c for c in t["columns"] if c["label"] == "Licence key")
        self.assertFalse(lic["core"])
        rid = self.api("POST", "/api/table/assets", {"row": {"Server name": "lic-host",
                                                             "Licence key": "ABC-123"}})["row"]["Record_ID"]
        wb = openpyxl.load_workbook(os.path.join(self.tmp, app.XLSX_NAME))
        headers = [c.value for c in wb["Asset_Master"][2]]
        self.assertIn("Server name", headers)
        self.assertIn("Licence key", headers)
        store2 = app.DataStore(app.LocalStorage(self.tmp), refresh_interval=0)
        row = next(r for r in store2.state["xlsx"]["tables"]["Asset_Master"]["rows"] if r["Record_ID"] == rid)
        self.assertEqual((row["Device"], row[lic["key"]]), ("lic-host", "ABC-123"))
        res = self.api("PUT", "/api/table/Asset_Master/columns", {"columns": [{"key": "Device", "label": "Device"}]})
        self.assertIn("Licence key", res["removed"])
        t = self.table("Asset_Master")
        self.assertTrue(all(c["key"] in app.schema_keys("Asset_Master") for c in t["columns"]))
        self.api("PUT", "/api/table/Asset_Master/columns", {"columns": [{"label": "X"}, {"label": "x"}]}, expect=400)

    def test_runbook_autofill_and_validation(self):
        self.api("POST", "/api/table/assets", {"row": {"Customer": "Hooli", "Device": "hooli-db01",
                                                       "Environment": "UAT", "Serial_No": "HOOLI-1"}})
        e = self.api("POST", "/api/table/runbook", {"row": {"Asset": "HOOLI-DB01", "Title": "Restarted listener",
                                                            "Solution": "lsnrctl start", "Commands": "lsnrctl status"}})["row"]
        self.assertEqual((e["Customer"], e["Environment"], e["Category"], e["Known_Issue"]),
                         ("Hooli", "UAT", "Information", "No"))
        self.assertTrue(e["Date"])
        self.api("POST", "/api/table/runbook", {"row": {"Title": "no asset"}}, expect=400)
        self.api("POST", "/api/table/runbook", {"row": {"Asset": "x"}}, expect=400)
        self.api("POST", "/api/table/runbook", {"row": {"Asset": "hooli-db01", "Title": "Restarted listener",
                                                        "Date": e["Date"], "Customer": "Hooli"}}, expect=409)

    def test_incident_art(self):
        st, data = jcall(self.base, "POST", "/api/table/incidents", {"row": {"Ticket_ID": "INC-1", "Action": "x"}})
        self.assertEqual(st, 400)
        self.assertIn("ART", data["error"])
        row = self.api("POST", "/api/table/incidents", {"row": {"Ticket_ID": "INC-77", "Priority": "p2", "Action": "a",
                                                                "Reason": "r", "Technical_Tracking": "t"}})["row"]
        self.assertEqual((row["Priority"], row["Status"]), ("P2", "New"))
        self.api("PUT", f"/api/table/incidents/{row['Record_ID']}", {"row": {"Priority": "P9"}}, expect=400)

    def test_bulletins_and_read_receipts(self):
        b = self.api("POST", "/api/table/bulletins", {"row": {"Title": "DC power test", "Severity": "critical",
                                                              "Expiry": "2099-01-01 10:00"}}, user="Lead")["row"]
        self.assertEqual((b["Severity"], b["Posted_By"]), ("Critical", "lead"))
        self.assertTrue(b["Start"])
        bid = b["Record_ID"]
        self.assertNotIn(bid, self.api("GET", "/api/state", user="Ana")["my_reads"])
        self.api("POST", f"/api/bulletins/{bid}/read", {}, user="Ana")
        self.assertIn(bid, self.api("GET", "/api/state", user="Ana")["my_reads"])
        self.assertNotIn(bid, self.api("GET", "/api/state", user="Ben")["my_reads"])
        self.api("POST", f"/api/bulletins/{bid}/read", {"read": False}, user="Ana")
        self.assertNotIn(bid, self.api("GET", "/api/state", user="Ana")["my_reads"])
        self.api("POST", "/api/table/bulletins", {"row": {"Title": "x", "Start": "2026-01-02 10:00",
                                                          "Expiry": "2026-01-01 10:00"}}, expect=400)
        self.api("POST", "/api/bulletins/R-missing/read", {}, expect=404)

    def test_row_history_and_restore_deleted(self):
        rid = self.api("POST", "/api/table/handovers", {"row": {"Date": "2026-11-06", "Shift": "night",
                                                                "Region": "APAC Night", "Notes": "first"}})["row"]["Record_ID"]
        self.api("PUT", f"/api/table/handovers/{rid}", {"row": {"Notes": "second"}}, user="Kim")
        h = self.api("GET", f"/api/history?key=table/Shift_Transitions/{rid}")["versions"]
        self.assertEqual([v["v"] for v in h], [2, 1])
        self.assertEqual(h[0]["by"], "kim")
        self.assertIn({"field": "Notes", "old": "first", "new": "second"}, h[0]["changes"])
        self.api("POST", "/api/history/restore", {"key": f"table/Shift_Transitions/{rid}", "v": 1})
        row = next(r for r in self.table("Shift_Transitions")["rows"] if r["Record_ID"] == rid)
        self.assertEqual((row["Notes"], row["Version"]), ("first", "3"))
        self.api("DELETE", f"/api/table/handovers/{rid}")
        last = self.api("GET", f"/api/history?key=table/Shift_Transitions/{rid}")["versions"][0]
        self.api("POST", "/api/history/restore", {"key": f"table/Shift_Transitions/{rid}", "v": last["v"]})
        self.assertTrue(any(r["Record_ID"] == rid for r in self.table("Shift_Transitions")["rows"]))
        self.api("POST", "/api/history/restore", {"key": "table/Shift_Transitions/none", "v": 1}, expect=404)


# --------------------------------------------------------------------------------------
class T04Imports(LocalServerTestCase):
    def commit(self, token, decisions=None, default_duplicate="skip"):
        return self.api("POST", "/api/import/commit", {"token": token, "decisions": decisions or {},
                                                       "default_duplicate": default_duplicate}, user="Importer")["batch"]

    def test_xlsx_preview_duplicates_commit_and_undo(self):
        existing = next(r for r in self.table("Asset_Master")["rows"] if r["Serial_No"] == "CZJ8340XQ1")
        data = xlsx_bytes({"Server Inventory": [
            ["ACME MANAGED SERVICES - CONFIDENTIAL"], [], ["Generated:", "2026-11-01"],
            ["Client", "Serial Number", "Host Name", "Env", "Operating System", "OS Version", "RAM", "Rack_Location"],
            ["Acme Corp", "CZJ8340XQ1", "acme-orarac01", "Production", "Oracle Linux", "8.10", "1 TB", "R12-U07"],
            ["Initech", "INI-1", "ini-sol01", "Production", "Solaris", "11.4", "64 GB", "R3-U01"],
            ["Initech", "INI-1", "ini-sol01-dup", "Production", "Solaris", "11.4", "64 GB", "R3-U02"],
            ["Initech", "", "", "UAT", "Linux", "9", "8 GB", ""],
        ]})
        pv = self.upload("servers.xlsx", data)
        sec = pv["sections"][0]
        self.assertEqual((sec["module"], sec["header_row"]), ("Asset_Master", 4))
        self.assertEqual(sec["new_columns"], ["Rack_Location"])
        self.assertEqual([it["status"] for it in sec["items"]], ["duplicate", "new", "duplicate_in_file", "invalid"])
        dup = sec["items"][0]
        self.assertEqual(dup["match_id"], existing["Record_ID"])
        self.assertIn("Memory", dup["changes"])
        batch = self.commit(pv["token"], {"0": {"0": "update"}})
        self.assertEqual(batch["sections"][0], {"module": "Asset_Master", "source": "Server Inventory",
                                                "added": 1, "updated": 1, "skipped": 2})
        t = self.table("Asset_Master")
        rack = next(c for c in t["columns"] if c["label"] == "Rack_Location")
        upd = next(r for r in t["rows"] if r["Record_ID"] == existing["Record_ID"])
        self.assertEqual((upd["Memory"], upd["OS_Version"], upd[rack["key"]]), ("1 TB", "8.10", "R12-U07"))
        self.assertEqual(upd["Special_Care"], existing["Special_Care"])  # untouched fields kept
        self.assertTrue(any(r["Device"] == "ini-sol01" for r in t["rows"]))
        self.api("POST", "/api/import/commit", {"token": pv["token"]}, expect=404)  # single use
        res = self.api("POST", f"/api/imports/{batch['id']}/undo", {})
        self.assertEqual((res["removed"], res["restored"]), (1, 1))
        t = self.table("Asset_Master")
        self.assertFalse(any(r["Device"] == "ini-sol01" for r in t["rows"]))
        self.assertEqual(next(r for r in t["rows"] if r["Record_ID"] == existing["Record_ID"])["Memory"],
                         existing["Memory"])
        self.api("POST", f"/api/imports/{batch['id']}/undo", {}, expect=409)
        self.assertTrue(self.api("GET", "/api/imports")["imports"][0]["undone"])

    def test_csv_and_tab_text_and_json(self):
        buf = io.StringIO()
        w = csv.writer(buf)
        w.writerow(["Customer", "Environment", "Asset", "Date", "Category", "Title", "Problem", "Resolution", "Commands"])
        w.writerow(["Acme Corp", "Production", "acme-boomi01", "05/11/2026", "Incident", "Atom down", "Listener hung",
                    "Restarted atom", "systemctl restart atom"])
        pv = self.upload("notes.csv", buf.getvalue().encode())
        sec = pv["sections"][0]
        self.assertEqual(sec["module"], "Runbook")
        self.assertEqual(sec["items"][0]["status"], "new")
        self.commit(pv["token"])
        e = next(r for r in self.table("Runbook")["rows"] if r["Title"] == "Atom down")
        self.assertEqual((e["Date"], e["Note"], e["Solution"]), ("2026-11-05", "Listener hung", "Restarted atom"))
        txt = "Serial No\tDevice\tCustomer\tOS\r\nTXT-1\ttxt-host01\tGlobex\tRHEL\r\n"
        pv = self.upload("assets.txt", txt.encode("utf-8"))
        self.assertEqual(pv["sections"][0]["module"], "Asset_Master")
        self.commit(pv["token"])
        self.assertTrue(any(r["Device"] == "txt-host01" for r in self.table("Asset_Master")["rows"]))
        js = json.dumps([{"Title": "Change freeze", "Severity": "Advisory", "Body": "No changes this weekend"}])
        pv = self.upload("notice.json", js.encode(), module="bulletins")
        self.assertEqual(pv["sections"][0]["module"], "Bulletins")
        self.commit(pv["token"])
        self.assertTrue(any(r["Title"] == "Change freeze" for r in self.table("Bulletins")["rows"]))
        self.upload("junk.csv", b"foo,bar\n1,2\n", expect=400)
        self.upload("x.docx", b"PK..", expect=400)
        self.upload("bad.xlsx", b"not a zip", expect=400)

    def test_plans_import_levels_and_duplicates(self):
        rows = [["Date", "Customer", "Task Name", "Ticket ID", "Window Start", "Phase", "Checkpoint", "Step No", "Step",
                 "Owner", "Approx Min", "Command"],
                ["2027-02-07", "Globex", "AG failover test", "CHG-555", "2027-02-07 20:00", "Pre-execution", "Yes",
                 "1.1", "Check AG health", "DBA", "15", "SELECT * FROM sys.dm_hadr_database_replica_states"],
                ["2027-02-07", "Globex", "AG failover test", "CHG-555", "", "Pre-execution", "", "1.1.1",
                 "Confirm sync", "DBA", "1h30m", ""],
                ["2027-02-07", "Globex", "AG failover test", "CHG-555", "", "During cut-over", "", "2.1",
                 "Fail over", "DBA", "0:10", "ALTER AVAILABILITY GROUP ag1 FAILOVER;"],
                ["2027-02-08", "Globex", "Bad", "CHG-556", "", "x", "", "", "Step", "", "lots", ""]]
        pv = self.upload("plan.xlsx", xlsx_bytes({"Execution Plan": rows}))
        sec = pv["sections"][0]
        self.assertEqual(sec["module"], "plans")
        self.assertEqual([i["status"] for i in sec["items"]], ["new", "invalid"])
        self.commit(pv["token"])
        t = next(t for t in self.api("GET", "/api/plans")["plans"]["tasks"] if t["ticket_id"] == "CHG-555")
        self.assertEqual([p["name"] for p in t["phases"]], ["Pre-execution", "During cut-over"])
        self.assertTrue(t["phases"][0]["checkpoint"])
        self.assertEqual([s["num"] for s in t["phases"][0]["steps"]], ["1.1", "1.1.1"])
        self.assertEqual(t["phases"][0]["steps"][1]["approx_min"], 90)
        self.assertEqual(t["phases"][1]["steps"][0]["planned"], "2027-02-07 21:45")
        pv = self.upload("plan.xlsx", xlsx_bytes({"Execution Plan": rows[:4]}))
        it = pv["sections"][0]["items"][0]
        self.assertEqual((it["status"], it["match_id"]), ("duplicate", t["id"]))
        self.commit(pv["token"], default_duplicate="update")
        self.assertEqual(sum(1 for x in self.api("GET", "/api/plans")["plans"]["tasks"] if x["ticket_id"] == "CHG-555"), 1)

    def test_templates_download_and_reimport(self):
        for module in ("assets", "runbook", "incidents", "handovers", "bulletins", "plans"):
            for fmt in ("xlsx", "csv", "txt", "json"):
                st, hdrs, raw = http(self.base, "GET", f"/api/template/{module}?format={fmt}")
                self.assertEqual(st, 200, (module, fmt))
                self.assertIn("attachment", hdrs["Content-Disposition"])
                if fmt == "xlsx":
                    self.assertIn("Instructions", openpyxl.load_workbook(io.BytesIO(raw)).sheetnames)
        _, _, raw = http(self.base, "GET", "/api/template/assets?format=xlsx")
        pv = self.upload("Asset_Master_template.xlsx", raw)
        self.assertEqual((pv["sections"][0]["module"], pv["sections"][0]["header_row"]), ("Asset_Master", 2))
        self.assertEqual(pv["sections"][0]["items"][0]["status"], "duplicate")  # example row = seeded asset
        _, _, raw = http(self.base, "GET", "/api/template/plans?format=csv")
        pv = self.upload("plans.csv", raw)
        self.assertEqual(pv["sections"][0]["module"], "plans")
        self.assertEqual(pv["sections"][0]["items"][0]["status"], "duplicate")  # example = seeded plan
        self.assertEqual(http(self.base, "GET", "/api/template/assets?format=pdf")[0], 400)

    def test_full_json_backup_round_trip(self):
        _, _, raw = http(self.base, "GET", "/api/export/json")
        self.assertTrue({"tables", "plans", "hierarchy", "audit"} <= set(json.loads(raw)))
        pv = self.upload("backup.json", raw)
        modules = {s["module"] for s in pv["sections"]}
        self.assertTrue({"Asset_Master", "plans", "shiftlog"} <= modules)
        assets = next(s for s in pv["sections"] if s["module"] == "Asset_Master")
        self.assertTrue(all(i["status"] == "duplicate" for i in assets["items"]))
        before = len(self.table("Asset_Master")["rows"])
        self.commit(pv["token"])
        self.assertEqual(len(self.table("Asset_Master")["rows"]), before)


# --------------------------------------------------------------------------------------
class T05ShiftLog(LocalServerTestCase):
    def sec(self, d, s, name):
        h = self.api("GET", "/api/hierarchy")["hierarchy"]
        y, m, _ = d.split("-")
        return next(x for x in h["years"][y]["months"][m]["dates"][d]["shifts"][s]["sections"] if x["name"] == name)

    def test_crud_clone_and_columns(self):
        self.assertTrue(self.api("POST", "/api/hierarchy/node", {"level": "year", "year": 2027})["created"])
        self.assertTrue(self.api("POST", "/api/hierarchy/node", {"level": "month", "year": 2027, "month": 3})["created"])
        self.assertTrue(self.api("POST", "/api/hierarchy/node", {"level": "shift", "date": "2027-03-15",
                                                                 "shift": "night"})["created"])
        sec = self.sec("2027-03-15", "Night", "In Work - All Clients")
        self.assertEqual(len(self.api("GET", "/api/hierarchy")["hierarchy"]["years"]["2027"]["months"]["03"]["dates"]
                             ["2027-03-15"]["shifts"]["Night"]["sections"]), 7)
        loc = {"date": "2027-03-15", "shift": "Night", "section_id": sec["id"]}
        row = self.api("POST", "/api/hierarchy/row", {**loc, "cells": {"ticket": "INC-1", "Client": "Acme"}})["row"]
        clone = self.api("POST", "/api/hierarchy/row/clone", {**loc, "row_id": row["id"],
                                                              "cells": {"Ticket": "INC-2"}})["row"]
        sec = self.sec("2027-03-15", "Night", "In Work - All Clients")
        ids = {c["name"]: c["id"] for c in sec["columns"]}
        self.assertEqual([r["cells"][ids["Ticket"]] for r in sec["rows"]], ["INC-1", "INC-2"])
        self.assertEqual(sec["rows"][1]["cells"][ids["Client"]], "Acme")
        self.api("PUT", "/api/hierarchy/row", {**loc, "row_id": clone["id"], "cells": {"Status": "Done"}})
        cols = [{"id": ids["Client"], "name": "Customer"}, {"id": ids["Ticket"], "name": "Ticket"}, {"name": "Added"}]
        res = self.api("PUT", "/api/hierarchy/section/columns", {**loc, "columns": cols})
        self.assertIn("Status", res["removed"])
        self.assertEqual(self.sec("2027-03-15", "Night", "In Work - All Clients")["rows"][0]["cells"][ids["Client"]], "Acme")
        self.api("DELETE", "/api/hierarchy/row", {**loc, "row_id": row["id"]})
        self.api("DELETE", "/api/hierarchy/row", {**loc, "row_id": row["id"]}, expect=404)
        self.api("POST", "/api/hierarchy/node", {"level": "shift", "date": "2027-03-15", "shift": "Lunch"}, expect=400)
        nxt = self.api("POST", "/api/hierarchy/node", {"level": "shift", "date": "2027-03-16", "shift": "Morning",
                                                       "carry_forward": True})
        self.assertEqual(nxt["carried_from"], "2027-03-15 Night")
        self.api("DELETE", "/api/hierarchy/node", {"level": "year", "year": "2027"})


# --------------------------------------------------------------------------------------
class T06PersistenceExports(LocalServerTestCase):
    def test_external_excel_edit_and_formula_safety(self):
        path = os.path.join(self.tmp, app.XLSX_NAME)
        wb = openpyxl.load_workbook(path)
        ws = wb["Asset_Master"]
        headers = [c.value for c in ws[2]]
        row = [""] * len(headers)
        row[headers.index("CI name")] = "ext-edit01"  # the Device column is headed with its CMDB label
        ws.append(row)
        wb.save(path)
        ext = next(r for r in self.table("Asset_Master")["rows"] if r["Device"] == "ext-edit01")
        self.assertTrue(ext["Record_ID"])  # row added in Excel without an ID gets one
        self.api("PUT", f"/api/table/assets/{ext['Record_ID']}", {"row": {"Role": '=HYPERLINK("http://evil")'}})
        wb = openpyxl.load_workbook(path)
        cell = next(c for r in wb["Asset_Master"].iter_rows() for c in r
                    if isinstance(c.value, str) and c.value.startswith("=HYPERLINK"))
        self.assertEqual(cell.data_type, "s")

    def test_export_workbook_and_audit(self):
        self.api("POST", "/api/table/handovers", {"row": {"Date": "2026-11-07", "Shift": "Morning"}}, user="Zoë Müller")
        self.assertEqual(self.api("GET", "/api/audit?limit=5")["audit"][0]["User"], "zo.m.ller")
        st, _, raw = http(self.base, "GET", "/api/export/xlsx")
        self.assertEqual(st, 200)
        wb = openpyxl.load_workbook(io.BytesIO(raw))
        for s in app.SHEET_ORDER + ["Execution_Plans", "Shift_Log", "Audit_Trail"]:
            self.assertIn(s, wb.sheetnames)
        plan_headers = [c.value for c in wb["Execution_Plans"][2]]
        self.assertIn("Planned_Calculated", plan_headers)
        self.assertIn("Step_No", plan_headers)
        pv = self.upload("export.xlsx", raw)  # the exported workbook re-imports
        self.assertIn("plans", {s["module"] for s in pv["sections"]})


# --------------------------------------------------------------------------------------
# Fake Microsoft Graph (token endpoint, site/drive lookup, folders, ETag uploads)
# --------------------------------------------------------------------------------------
class FakeGraph:
    SITE_ID = "contoso.sharepoint.com,11111111-aaaa,22222222-bbbb"

    def __init__(self):
        self.files, self.folders, self.lock = {}, set(), threading.Lock()
        fake = self

        class H(BaseHTTPRequestHandler):
            def log_message(self, *a):
                pass

            def _send(self, status, body=b"", ctype="application/json"):
                if isinstance(body, (dict, list)):
                    body = json.dumps(body).encode()
                self.send_response(status)
                self.send_header("Content-Type", ctype)
                self.send_header("Content-Length", str(len(body)))
                self.end_headers()
                self.wfile.write(body)

            def _auth(self):
                if self.headers.get("Authorization") != "Bearer fake-token":
                    self._send(401, {"error": {"code": "InvalidAuthenticationToken", "message": "no"}})
                    return False
                return True

            def do_POST(self):
                body = self.rfile.read(int(self.headers.get("Content-Length") or 0))
                path = urllib.parse.unquote(urllib.parse.urlsplit(self.path).path)
                if path.endswith("/oauth2/v2.0/token"):
                    if urllib.parse.parse_qs(body.decode()).get("client_secret") != ["shh"]:
                        return self._send(401, {"error": "invalid_client"})
                    return self._send(200, {"access_token": "fake-token", "expires_in": 3600})
                if not self._auth():
                    return
                if path.endswith("/children"):
                    name = json.loads(body)["name"]
                    parent = path.split("/root", 1)[1].removesuffix("/children").strip(":/")
                    with fake.lock:
                        fake.folders.add(f"{parent}/{name}".strip("/"))
                    return self._send(201, {"name": name, "folder": {}})
                self._send(404, {})

            def do_GET(self):
                path = urllib.parse.unquote(urllib.parse.urlsplit(self.path).path)
                if path.startswith("/download/"):
                    with fake.lock:
                        f = fake.files.get(path[len("/download/"):])
                    return self._send(200, f[0], "application/octet-stream") if f else self._send(404)
                if not self._auth():
                    return
                if path == "/v1.0/sites/contoso.sharepoint.com:/sites/MSPOps":
                    return self._send(200, {"id": FakeGraph.SITE_ID})
                if path == f"/v1.0/sites/{FakeGraph.SITE_ID}/drives":
                    return self._send(200, {"value": [{"id": "drv-docs", "name": "Documents"}]})
                if path.startswith("/v1.0/drives/drv-docs/root:/"):
                    key = path[len("/v1.0/drives/drv-docs/root:/"):].rstrip(":")
                    with fake.lock:
                        f, is_folder = fake.files.get(key), key in fake.folders
                    if is_folder:
                        return self._send(200, {"name": key, "folder": {}})
                    if not f:
                        return self._send(404, {"error": {"code": "itemNotFound", "message": "nope"}})
                    host = f"http://127.0.0.1:{self.server.server_address[1]}"
                    return self._send(200, {"name": key, "eTag": f[1],
                                            "@microsoft.graph.downloadUrl": f"{host}/download/{urllib.parse.quote(key)}"})
                self._send(404, {"error": {"code": "itemNotFound", "message": path}})

            def do_PUT(self):
                if not self._auth():
                    return
                parts = urllib.parse.urlsplit(self.path)
                key = urllib.parse.unquote(parts.path)[len("/v1.0/drives/drv-docs/root:/"):].removesuffix(":/content")
                body = self.rfile.read(int(self.headers.get("Content-Length") or 0))
                with fake.lock:
                    cur = fake.files.get(key)
                    if "conflictBehavior=fail" in urllib.parse.unquote(parts.query) and cur:
                        return self._send(409, {"error": {"code": "nameAlreadyExists"}})
                    match = self.headers.get("If-Match")
                    if match and (not cur or cur[1] != match):
                        return self._send(412, {"error": {"code": "resourceModified"}})
                    etag = f'"{{{uuid.uuid4()}}},1"'
                    fake.files[key] = [body, etag]
                self._send(201, {"name": key, "eTag": etag})

        self.server = ThreadingHTTPServer(("127.0.0.1", 0), H)
        self.base = start(self.server)


class T08SharePoint(unittest.TestCase):
    def setUp(self):
        self.fake = FakeGraph()

    def tearDown(self):
        stop(self.fake.server)

    def cfg(self, **over):
        opts = dict(host="127.0.0.1", port=0, storage="sharepoint", quiet=True, seed=True, refresh_interval=3600,
                    backup_thread=False, setup_code_file=os.path.join(tempfile.gettempdir(), f"sc_{uuid.uuid4().hex}.txt"),
                    sp_tenant_id="t", sp_client_id="c", sp_client_secret="shh", sp_site_host="contoso.sharepoint.com",
                    sp_site_path="/sites/MSPOps", sp_folder="Ops/MSP-Runbook", graph_base=self.fake.base + "/v1.0",
                    auth_base=self.fake.base)
        opts.update(over)
        return app.Config(**opts)

    def test_bad_secret(self):
        with self.assertRaises(app.StorageError) as cm:
            app.create_server(self.cfg(sp_client_secret="wrong"))
        self.assertIn("sign-in failed", cm.exception.message)

    def test_all_files_central_and_conflict_retry(self):
        server = app.create_server(self.cfg())
        base = start(server)
        setup_admin(base, server)
        try:
            for name in (app.XLSX_NAME, app.PLANS_NAME, app.JSON_NAME, app.META_NAME, app.USERS_NAME):
                self.assertIn("Ops/MSP-Runbook/" + name, self.fake.files)
            key = "Ops/MSP-Runbook/" + app.PLANS_NAME
            doc = json.loads(self.fake.files[key][0])
            tid = doc["tasks"][0]["id"]
            doc["tasks"][0]["notes"] = "edited elsewhere"  # someone else's save lands first
            with self.fake.lock:
                self.fake.files[key] = [json.dumps(doc).encode(), '"{external},9"']
            st, data = jcall(base, "POST", "/api/plans/op", {"op": "task.update", "task_id": tid,
                                                             "task": {"approver": "CAB-2"}})
            self.assertEqual(st, 200, data)
            stored = json.loads(self.fake.files[key][0])["tasks"][0]
            self.assertEqual((stored["notes"], stored["approver"]), ("edited elsewhere", "CAB-2"))
            st, data = jcall(base, "POST", "/api/table/assets", {"row": {"Device": "sp-host01"}})
            self.assertEqual(st, 200, data)
            wb = openpyxl.load_workbook(io.BytesIO(self.fake.files["Ops/MSP-Runbook/" + app.XLSX_NAME][0]))
            self.assertIn("sp-host01", [c.value for c in wb["Asset_Master"]["D"]])
            meta = json.loads(self.fake.files["Ops/MSP-Runbook/" + app.META_NAME][0])
            self.assertTrue(any(k.startswith("table/Asset_Master/") for k in meta["history"]))
            self.assertEqual(jcall(base, "GET", "/api/status")[1]["storage"]["kind"], "sharepoint")
        finally:
            stop(server)





# ======================================================================================
# Version 4: security, roles, Microsoft sign-in, terminals, RDP, admin, shift clone, backups
# ======================================================================================
class T09SecurityAuth(LocalServerTestCase):
    def test_everything_needs_sign_in(self):
        anon = Client(self.base)
        self.assertEqual(anon.request("GET", "/")[0], 200)  # the page itself shows the sign-in screen
        for method, path in [("GET", "/api/state"), ("GET", "/api/status"), ("GET", "/api/table/assets"),
                             ("POST", "/api/table/assets"), ("POST", "/api/plans/op"), ("GET", "/api/export/xlsx"),
                             ("GET", "/api/admin/users"), ("POST", "/api/import/preview"), ("GET", "/api/history?key=x"),
                             ("POST", "/api/terminal/sessions"), ("GET", "/api/rdp/file"), ("POST", "/api/me/password")]:
            st, _, _ = anon.request(method, path, {} if method != "GET" else None)
            self.assertEqual(st, 401, (method, path))
        st, d = anon.json("GET", "/api/auth/status")
        self.assertEqual((st, d["authenticated"], d["setup_required"]), (200, False, False))
        self.assertNotIn("csrf", d)

    def test_setup_cannot_be_repeated(self):
        st, d = Client(self.base).json("POST", "/api/auth/setup", {"code": "X", "username": "evil", "password": "Evil-Passw0rd"})
        self.assertEqual(st, 409)

    def test_login_generic_errors_and_lockout(self):
        client_for(self.base, "Lock Me", "Viewer")
        c = Client(self.base)
        st1, d1 = c.login("lock.me", "wrong-password1")
        st2, d2 = c.login("nobody.here", "wrong-password1")
        self.assertEqual((st1, st2), (401, 401))
        self.assertEqual(d1["error"], d2["error"])  # no user enumeration
        for _ in range(4):
            c.login("lock.me", "wrong-password1")
        st, d = c.login("lock.me", USER_PW)  # correct password, but locked out now
        self.assertEqual(st, 429, d)
        self.server.throttle.reset("lock.me")
        self.assertEqual(c.login("lock.me", USER_PW)[0], 200)

    def test_login_requires_json(self):
        st, _, _ = Client(self.base).request("POST", "/api/auth/login", b"username=admin&password=x",
                                             ctype="application/x-www-form-urlencoded")
        self.assertEqual(st, 415)

    def test_cookie_flags_and_session_rotation(self):
        c = Client(self.base)
        st, hdrs, _ = c.request("POST", "/api/auth/login", {"username": "admin", "password": ADMIN_PW})
        cookie = hdrs.get("Set-Cookie")
        self.assertIn("HttpOnly", cookie)
        self.assertIn("SameSite=Lax", cookie)
        self.assertNotIn("Secure", cookie)  # plain http in tests; added automatically with HTTPS
        first = c.cookie
        c.login("admin", ADMIN_PW)
        self.assertNotEqual(first, c.cookie)  # new session id at every sign-in
        old = Client(self.base)
        old.cookie = first
        self.assertEqual(old.request("GET", "/api/state")[0], 401)  # previous id is dead

    def test_csrf_required_for_changes(self):
        c = client_for(self.base, "Csrf User", "Engineer")
        saved = c.csrf
        c.csrf = ""
        self.assertEqual(c.json("POST", "/api/table/assets", {"row": {"Device": "csrf1"}})[0], 403)
        c.csrf = "forged"
        self.assertEqual(c.json("POST", "/api/table/assets", {"row": {"Device": "csrf1"}})[0], 403)
        c.csrf = saved
        self.assertEqual(c.json("POST", "/api/table/assets", {"row": {"Device": "csrf1"}})[0], 200)

    def test_security_headers(self):
        for path in ("/", "/api/state", "/api/missing"):
            _, h, _ = http(self.base, "GET", path)
            csp = h["Content-Security-Policy"]
            for part in ("default-src 'self'", "object-src 'none'", "frame-ancestors 'self'", "base-uri 'self'"):
                self.assertIn(part, csp)
            script_src = re.search(r"script-src ([^;]+)", csp).group(1)
            self.assertNotIn("unsafe-inline", script_src)  # inline scripts only with the per-request nonce
            self.assertEqual(h["X-Content-Type-Options"], "nosniff")
            self.assertEqual(h["Referrer-Policy"], "same-origin")
            self.assertNotIn("Python", h.get("Server", ""))
        st, h, raw = http(self.base, "GET", "/")
        page = raw.decode()
        nonce = re.search(r"'nonce-([^']+)'", h["Content-Security-Policy"]).group(1)
        self.assertEqual(page.count(f'<script nonce="{nonce}">'), 2)
        self.assertEqual(page.count("<script>"), 0)
        self.assertNotEqual(nonce, re.search(r"'nonce-([^']+)'", http(self.base, "GET", "/")[1]["Content-Security-Policy"]).group(1))
        self.assertEqual(page.count('integrity="'), 0)  # SRI is set from JS when xterm loads:
        self.assertIn("sha384-/nfmYPUzWMS6v2atn8hbljz7NE0EI1iGx34lJaNzyVjWGDzMv+ciUZUeJpKA3Glc", page)

    def test_password_policy_and_change(self):
        c = client_for(self.base, "Pw User", "Viewer")
        for bad in ("short1", "alllettersonly", "1234567890", "pw.user-Passw0rd"):
            st, d = c.json("POST", "/api/me/password", {"current": USER_PW, "new": bad})
            self.assertEqual(st, 400, bad)
        self.assertEqual(c.json("POST", "/api/me/password", {"current": "wrong", "new": "Fresh-Passw0rd-9"})[0], 400)
        self.assertEqual(c.json("POST", "/api/me/password", {"current": USER_PW, "new": "Fresh-Passw0rd-9"})[0], 200)
        self.assertEqual(Client(self.base).login("pw.user", USER_PW)[0], 401)
        self.assertEqual(Client(self.base).login("pw.user", "Fresh-Passw0rd-9")[0], 200)

    def test_temp_password_forces_change(self):
        st, d = self.admin.json("POST", "/api/admin/users", {"username": "newbie", "role": "Engineer"})
        temp = d["temp_password"]
        self.assertTrue(temp)
        c = Client(self.base)
        st, d = c.login("newbie", temp)
        self.assertTrue(d["user"]["must_change"])
        self.assertEqual(c.json("GET", "/api/state")[0], 403)
        self.assertEqual(c.json("POST", "/api/me/password", {"current": temp, "new": "Fresh-Start-2026"})[0], 200)
        self.assertEqual(c.json("GET", "/api/state")[0], 200)

    def test_disabled_user_and_last_admin(self):
        c = client_for(self.base, "Temp Worker", "Engineer")
        self.assertEqual(c.json("GET", "/api/state")[0], 200)
        self.assertEqual(self.admin.json("PUT", "/api/admin/users/temp.worker", {"disabled": True})[0], 200)
        self.assertEqual(c.json("GET", "/api/state")[0], 401)
        self.assertEqual(Client(self.base).login("temp.worker", USER_PW)[0], 401)
        self.assertEqual(self.admin.json("PUT", "/api/admin/users/admin", {"role": "Lead"})[0], 400)
        self.assertEqual(self.admin.json("DELETE", "/api/admin/users/admin")[0], 400)

    def test_path_traversal_and_safe_redirects(self):
        st = app.LocalStorage(self.tmp)
        for bad in ("../x.txt", "a/../../x", "C:/Windows/x", "/etc/passwd"):
            with self.assertRaises(app.StorageError):
                st.write(bad, b"x", None)
        self.assertEqual(self.admin.json("GET", "/api/admin/recordings/..%2F..%2Fmsp_users")[0], 404)
        for bad in ("//evil.com", "https://evil.com", "/\\evil.com", "javascript:alert(1)", "/ok\r\nSet-Cookie:x"):
            self.assertEqual(app.safe_next(bad), "/")
        self.assertEqual(app.safe_next("/#plans"), "/#plans")

    def test_body_limits_and_bad_json(self):
        self.assertEqual(http(self.base, "POST", "/api/table/assets", b"{not json", ctype="application/json")[0], 400)
        st, _, _ = http(self.base, "POST", "/api/table/assets", b"x", headers={"Content-Length": str(app.MAX_BODY + 1)},
                        ctype="application/json")
        self.assertEqual(st, 413)

    def test_password_hashing(self):
        h = app.hash_password("Some-Passw0rd")
        self.assertTrue(h.startswith("pbkdf2_sha256$"))
        self.assertTrue(app.verify_password("Some-Passw0rd", h))
        self.assertFalse(app.verify_password("some-passw0rd", h))
        self.assertNotEqual(h, app.hash_password("Some-Passw0rd"))  # salted
        users = json.load(open(os.path.join(self.tmp, app.USERS_NAME)))
        self.assertNotIn(ADMIN_PW, json.dumps(users))


class T09bSecurityHardening(LocalServerTestCase):
    def test_password_spraying_limited_per_ip(self):
        c = Client(self.base)
        for i in range(app.LoginThrottle.IP_LIMIT):
            c.login(f"spray{i}", "Wrong-passw0rd")
        st, d = c.login("admin", ADMIN_PW)  # even the right password is refused from that address now
        self.assertEqual(st, 429, d)
        self.server.throttle.ip_fails.clear()
        self.assertEqual(c.login("admin", ADMIN_PW)[0], 200)

    def test_admin_never_sees_session_ids(self):
        victim = client_for(self.base, "Victim Vic", "Engineer")
        st, d = self.admin.json("GET", "/api/admin/sessions")
        blob = json.dumps(d)
        self.assertNotIn(victim.cookie, blob)
        self.assertNotIn(self.admin.cookie, blob)
        handle = next(s["handle"] for s in d["web"] if s["user"] == "victim.vic")
        self.assertEqual(self.admin.json("DELETE", f"/api/admin/sessions/{handle}")[0], 200)
        self.assertEqual(victim.json("GET", "/api/state")[0], 401)
        own = next(s["handle"] for s in self.admin.json("GET", "/api/admin/sessions")[1]["web"] if s["user"] == "admin")
        self.assertEqual(self.admin.json("DELETE", f"/api/admin/sessions/{own}")[0], 400)

    def test_only_http_urls_are_fetched(self):
        with self.assertRaises(app.StorageError):
            app.http_open(urllib.request.Request("file:///etc/passwd"), 5)


class T10RoleMatrix(LocalServerTestCase):
    def test_permissions_by_role(self):
        v, e = client_for(self.base, "Vera Viewer", "Viewer"), client_for(self.base, "Eli Engineer", "Engineer")
        lead = client_for(self.base, "Lou Lead", "Lead")
        asset = self.admin.json("POST", "/api/table/assets", {"row": {"Device": "role-host", "Serial_No": "ROLE1"}})[1]["row"]
        rid = asset["Record_ID"]
        bid = self.admin.json("GET", "/api/table/bulletins")[1]["table"]["rows"][0]["Record_ID"]
        tid = self.admin.json("GET", "/api/plans")[1]["plans"]["tasks"][0]["id"]
        cases = [  # (client, method, path, body, expected)
            (v, "GET", "/api/state", None, 200), (v, "POST", f"/api/bulletins/{bid}/read", {}, 200),
            (v, "POST", "/api/table/assets", {"row": {"Device": "v1"}}, 403),
            (v, "PUT", f"/api/table/assets/{rid}", {"row": {"Role": "x"}}, 403),
            (v, "GET", "/api/export/json", None, 403), (v, "GET", "/api/template/assets", None, 403),
            (v, "POST", "/api/plans/op", {"op": "task.update", "task_id": tid, "task": {"notes": "x"}}, 403),
            (e, "PUT", f"/api/table/assets/{rid}", {"row": {"Role": "web"}}, 200),
            (e, "POST", "/api/plans/op", {"op": "task.update", "task_id": tid, "task": {"notes": "eng"}}, 200),
            (e, "DELETE", f"/api/table/assets/{rid}", None, 403),
            (e, "POST", "/api/table/bulletins", {"row": {"Title": "eng post"}}, 403),
            (e, "PUT", "/api/table/assets/columns", {"columns": []}, 403),
            (e, "POST", "/api/plans/op", {"op": "phase.signoff", "task_id": tid, "phase_id": "x"}, 403),
            (e, "POST", "/api/plans/op", {"op": "task.delete", "task_id": tid}, 403),
            (e, "POST", "/api/history/restore", {"key": "x", "v": 1}, 403),
            (e, "POST", "/api/imports/x/undo", {}, 403), (e, "GET", "/api/admin/users", None, 403),
            (e, "GET", "/api/export/json", None, 200),
            (lead, "POST", "/api/table/bulletins", {"row": {"Title": "lead post"}}, 200),
            (lead, "DELETE", f"/api/table/assets/{rid}", None, 200),
            (lead, "GET", "/api/admin/users", None, 403), (lead, "GET", "/api/admin/recordings", None, 403),
            (self.admin, "GET", "/api/admin/users", None, 200), (self.admin, "GET", "/api/admin/health", None, 200),
        ]
        for c, method, path, body, expected in cases:
            st, d = c.json(method, path, body)
            self.assertEqual(st, expected, f"{method} {path} -> {st} {d}")
            if expected == 403 and "role" in d.get("error", ""):
                self.assertIn("needs", d["error"])
        st, d = e.json("GET", "/api/state")
        self.assertEqual(d["user"]["role"], "Engineer")


class T11MicrosoftSignIn(unittest.TestCase):
    """OpenID Connect code flow with PKCE against a fake Microsoft authority."""

    def setUp(self):
        self.claims = {}
        test = self

        class H(BaseHTTPRequestHandler):
            def log_message(self, *a):
                pass

            def do_POST(self):
                form = urllib.parse.parse_qs(self.rfile.read(int(self.headers["Content-Length"])).decode())
                test.token_form = form
                payload = base64.urlsafe_b64encode(json.dumps(test.claims).encode()).decode().rstrip("=")
                body = json.dumps({"id_token": f"eyJhbGciOiJub25lIn0.{payload}.sig"}).encode()
                self.send_response(200)
                self.send_header("Content-Type", "application/json")
                self.send_header("Content-Length", str(len(body)))
                self.end_headers()
                self.wfile.write(body)

        self.authority = ThreadingHTTPServer(("127.0.0.1", 0), H)
        auth_base = start(self.authority)
        self.tmp = tempfile.mkdtemp()
        cfg = app.Config(port=0, data_dir=self.tmp, quiet=True, backup_thread=False, refresh_interval=0,
                         setup_code_file=os.path.join(self.tmp, "c.txt"), entra_enabled=True,
                         entra_tenant_id="contoso-tid", entra_client_id="app-123", entra_client_secret="s3",
                         public_url="https://ops.contoso.com", auth_base=auth_base)
        self.server = app.create_server(cfg)
        self.base = start(self.server)
        self.admin = setup_admin(self.base, self.server)

    def tearDown(self):
        stop(self.server)
        stop(self.authority)
        shutil.rmtree(self.tmp, ignore_errors=True)

    def flow(self, upn="ana@contoso.com", nonce_ok=True, tid="contoso-tid"):
        c = Client(self.base)
        st, h, _ = c.request("GET", "/auth/entra/login?next=/%23plans")
        self.assertEqual(st, 302)
        q = urllib.parse.parse_qs(urllib.parse.urlsplit(h["Location"]).query)
        self.assertEqual((q["code_challenge_method"][0], q["client_id"][0]), ("S256", "app-123"))
        self.assertEqual(q["redirect_uri"][0], "https://ops.contoso.com/auth/entra/callback")
        self.claims = {"aud": "app-123", "tid": tid, "nonce": q["nonce"][0] if nonce_ok else "other",
                       "exp": time.time() + 600, "preferred_username": upn, "name": "Ana Analyst"}
        st, h, _ = c.request("GET", f"/auth/entra/callback?code=abc&state={q['state'][0]}")
        return c, st, h["Location"]

    def test_unknown_user_denied_then_allowed(self):
        c, st, loc = self.flow()
        self.assertEqual(st, 302)
        self.assertTrue(loc.startswith("/?auth_error="))
        self.assertIn("code_verifier", self.token_form)  # PKCE verifier sent
        self.admin.json("POST", "/api/admin/users", {"username": "ana", "email": "ana@contoso.com", "role": "Engineer",
                                                     "auth": "entra"})
        c, st, loc = self.flow()
        self.assertEqual(loc, "/#plans")
        st, d = c.json("GET", "/api/auth/status")
        self.assertEqual((d["authenticated"], d["user"]["username"], d["user"]["role"]), (True, "ana", "Engineer"))
        # a Microsoft-only account cannot sign in with a password
        self.assertEqual(Client(self.base).login("ana", "whatever-1234")[0], 401)

    def test_token_checks_and_auto_create(self):
        self.admin.json("POST", "/api/admin/users", {"username": "bob", "email": "bob@contoso.com", "role": "Viewer",
                                                     "auth": "entra"})
        _, _, loc = self.flow("bob@contoso.com", nonce_ok=False)
        self.assertIn("auth_error", loc)
        _, _, loc = self.flow("bob@contoso.com", tid="other-tenant")
        self.assertIn("auth_error", loc)
        c = Client(self.base)
        st, _, _ = c.request("GET", "/auth/entra/callback?code=abc&state=forged")
        self.assertEqual(st, 302)
        self.admin.json("PUT", "/api/admin/settings", {"settings": {"entra_default_role": "Viewer"}})
        c, st, loc = self.flow("new.person@contoso.com")
        self.assertEqual(loc, "/#plans")
        self.assertEqual(c.json("GET", "/api/auth/status")[1]["user"]["role"], "Viewer")


class T12ProxyIdentity(LocalServerTestCase):
    config_overrides = {"trust_proxy_user": True}

    def test_easy_auth_header_maps_to_user(self):
        self.admin.json("POST", "/api/admin/users", {"username": "asha", "email": "asha@contoso.com", "role": "Lead",
                                                     "auth": "entra"})
        c = Client(self.base)
        st, d = c.json("GET", "/api/auth/status", headers={"X-MS-CLIENT-PRINCIPAL-NAME": "asha@contoso.com"})
        self.assertEqual((d["authenticated"], d["user"]["username"]), (True, "asha"))
        self.assertTrue(c.cookie)
        st, d = Client(self.base).json("GET", "/api/state", headers={"X-MS-CLIENT-PRINCIPAL-NAME": "stranger@x.com"})
        self.assertEqual(st, 403)


# ---- terminals: a real SSH server built with paramiko ----------------------------------
import paramiko  # noqa: E402


class FakeSSHServer:
    def __init__(self, password="pw"):
        self.key = paramiko.RSAKey.generate(2048)
        self.password = password
        self.sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        self.sock.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
        self.sock.bind(("127.0.0.1", 0))
        self.sock.listen(5)
        self.port = self.sock.getsockname()[1]
        self.running = True
        threading.Thread(target=self.accept_loop, daemon=True).start()

    def accept_loop(self):
        while self.running:
            try:
                conn, _ = self.sock.accept()
            except OSError:
                return
            threading.Thread(target=self.handle, args=(conn,), daemon=True).start()

    def handle(self, conn):
        outer = self

        class Srv(paramiko.ServerInterface):
            def __init__(self):
                self.shell = threading.Event()

            def check_channel_request(self, kind, chanid):
                return paramiko.OPEN_SUCCEEDED if kind == "session" else paramiko.OPEN_FAILED_ADMINISTRATIVELY_PROHIBITED

            def check_auth_password(self, username, password):
                return paramiko.AUTH_SUCCESSFUL if password == outer.password else paramiko.AUTH_FAILED

            def get_allowed_auths(self, username):
                return "password"

            def check_channel_pty_request(self, *a):
                return True

            def check_channel_shell_request(self, channel):
                self.shell.set()
                return True

            def check_channel_window_change_request(self, *a):
                return True

        try:
            t = paramiko.Transport(conn)
            t.add_server_key(self.key)
            srv = Srv()
            t.start_server(server=srv)
            chan = t.accept(20)
            if chan is None:
                return
            srv.shell.wait(10)
            chan.send(b"Welcome to fake-ssh\r\n$ ")
            line, secret = "", False
            while True:
                data = chan.recv(1024)
                if not data:
                    break
                for ch in data.decode():
                    if ch == "\r":
                        chan.send(b"\r\n")
                        if secret:
                            chan.send(b"sudo: authenticated\r\n$ ")
                            secret = False
                        elif line == "sudo -v":
                            chan.send(b"[sudo] password for ops: ")
                            secret = True
                        elif line == "exit":
                            chan.send(b"logout\r\n")
                            chan.close()
                            return
                        else:
                            chan.send(f"ran: {line}\r\n$ ".encode())
                        line = ""
                    else:
                        line += ch
                        if not secret:
                            chan.send(ch.encode())
        except Exception:
            pass

    def stop(self):
        self.running = False
        self.sock.close()


class WSClient:
    def __init__(self, base, path, cookie, origin=None):
        u = urllib.parse.urlsplit(base)
        self.sock = socket.create_connection((u.hostname, u.port), timeout=20)
        key = base64.b64encode(os.urandom(16)).decode()
        origin = origin or f"http://{u.hostname}:{u.port}"
        req = (f"GET {path} HTTP/1.1\r\nHost: {u.hostname}:{u.port}\r\nUpgrade: websocket\r\nConnection: Upgrade\r\n"
               f"Sec-WebSocket-Key: {key}\r\nSec-WebSocket-Version: 13\r\nOrigin: {origin}\r\n"
               f"Cookie: msp_session={cookie}\r\n\r\n")
        self.sock.sendall(req.encode())
        self.buf = b""
        while b"\r\n\r\n" not in self.buf:
            chunk = self.sock.recv(4096)
            if not chunk:
                break
            self.buf += chunk
        head, _, self.buf = self.buf.partition(b"\r\n\r\n")
        self.status = int(head.split(b" ")[1])
        self.accept_ok = app.ws_accept_key(key).encode() in head

    def send(self, obj):
        payload = json.dumps(obj).encode()
        mask = os.urandom(4)
        n = len(payload)
        hdr = bytes([0x81]) + (bytes([0x80 | n]) if n < 126 else bytes([0x80 | 126]) + struct.pack(">H", n))
        self.sock.sendall(hdr + mask + bytes(b ^ mask[i % 4] for i, b in enumerate(payload)))

    def _read(self, n):
        while len(self.buf) < n:
            chunk = self.sock.recv(65536)
            if not chunk:
                raise ConnectionError("closed")
            self.buf += chunk
        out, self.buf = self.buf[:n], self.buf[n:]
        return out

    def recv(self):
        b1, b2 = self._read(2)
        n = b2 & 0x7F
        if n == 126:
            n = struct.unpack(">H", self._read(2))[0]
        elif n == 127:
            n = struct.unpack(">Q", self._read(8))[0]
        data = self._read(n)
        if b1 & 0x0F == 0x8:
            raise ConnectionError("closed by server")
        return json.loads(data.decode())

    def until(self, pred, timeout=20):
        end, seen = time.time() + timeout, []
        self.sock.settimeout(timeout)
        while time.time() < end:
            m = self.recv()
            seen.append(m)
            if pred(m, seen):
                return m, seen
        raise AssertionError(f"condition not met; got {seen}")

    def close(self):
        self.sock.close()


class T13Terminals(LocalServerTestCase):
    @classmethod
    def setUpClass(cls):
        super().setUpClass()
        cls.ssh = FakeSSHServer()
        cls.asset = cls.admin.json("POST", "/api/table/assets", {"row": {"Device": "fake-ssh", "IP_Address": "127.0.0.1",
                                                                         "Customer": "Lab"}})[1]["row"]

    @classmethod
    def tearDownClass(cls):
        cls.ssh.stop()
        super().tearDownClass()

    def open(self, client):
        st, d = client.json("POST", "/api/terminal/sessions", {"asset_id": self.asset["Record_ID"], "port": self.ssh.port,
                                                               "username": "ops"})
        self.assertEqual(st, 200, d)
        ws = WSClient(self.base, f"/ws/terminal?sid={d['sid']}&token={d['token']}", client.cookie)
        self.assertEqual((ws.status, ws.accept_ok), (101, True))
        return d, ws

    def test_full_session_recording_and_masking(self):
        eng = client_for(self.base, "Term Eng", "Engineer")
        d, ws = self.open(eng)
        ws.send({"t": "auth", "password": "pw", "cols": 100, "rows": 30})
        m, _ = ws.until(lambda m, s: m["t"] == "hostkey")
        self.assertTrue(m["fingerprint"].startswith("SHA256:"))
        ws.send({"t": "auth", "password": "pw", "trust": m["fingerprint"], "cols": 100, "rows": 30})
        ws.until(lambda m, s: m["t"] == "connected")
        ws.until(lambda m, s: "Welcome" in "".join(x.get("d", "") for x in s))
        ws.send({"t": "in", "d": "uptime\r"})
        ws.until(lambda m, s: "ran: uptime" in "".join(x.get("d", "") for x in s))
        ws.send({"t": "resize", "cols": 140, "rows": 40})
        ws.send({"t": "in", "d": "sudo -v\r"})
        ws.until(lambda m, s: "password for ops" in "".join(x.get("d", "") for x in s))
        ws.send({"t": "in", "d": "TopSecret-99\r"})
        ws.until(lambda m, s: "authenticated" in "".join(x.get("d", "") for x in s))
        ws.send({"t": "in", "d": "exit\r"})
        ws.until(lambda m, s: m["t"] == "closed")
        ws.close()
        for _ in range(50):  # the server finalises the recording asynchronously
            recs = self.admin.json("GET", "/api/admin/recordings")[1]["sessions"]
            rec = next((r for r in recs if r["id"] == d["sid"]), None)
            if rec and rec["status"] == "ended":
                break
            time.sleep(0.1)
        self.assertEqual((rec["user"], rec["host"], rec["username"], rec["type"]), ("term.eng", "127.0.0.1", "ops", "ssh"))
        self.assertTrue(rec["file"].startswith(".recordings/"))
        st, r = self.admin.json("GET", f"/api/admin/recordings/{d['sid']}")
        text = r["content"]
        self.assertTrue(r["integrity_ok"])
        self.assertIn(">>> uptime", text)
        self.assertIn("ran: uptime", text)
        self.assertIn("[hidden input]", text)
        self.assertNotIn("TopSecret-99", text)
        self.assertNotIn("\x1b", text)
        raw = open(os.path.join(self.tmp, rec["file"]), "rb").read()
        self.assertEqual(hashlib.sha256(raw).hexdigest(), rec["sha256"])
        # only admins can read recordings
        self.assertEqual(eng.json("GET", f"/api/admin/recordings/{d['sid']}")[0], 403)
        # the host key is now trusted, so a second session connects straight away
        d2, ws2 = self.open(eng)
        ws2.send({"t": "auth", "password": "pw"})
        ws2.until(lambda m, s: m["t"] == "connected")
        ws2.close()

    def test_wrong_password_retry_and_changed_host_key(self):
        eng = client_for(self.base, "Term Eng2", "Engineer")
        hid = f"127.0.0.1:{self.ssh.port}"
        self.server.store.trust_host(hid, {"type": "ssh-rsa", "fingerprint": "SHA256:not-the-real-key"}, "x")
        d, ws = self.open(eng)
        ws.send({"t": "auth", "password": "pw"})
        m, _ = ws.until(lambda m, s: m["t"] == "error")
        self.assertIn("CHANGED", m["m"])
        ws.close()
        self.admin.json("DELETE", "/api/admin/known-hosts", {"host": hid})
        d, ws = self.open(eng)
        ws.send({"t": "auth", "password": "pw"})
        m, _ = ws.until(lambda m, s: m["t"] == "hostkey")
        ws.send({"t": "auth", "password": "bad", "trust": m["fingerprint"]})
        m, _ = ws.until(lambda m, s: m["t"] == "error")
        self.assertTrue(m.get("retry"))
        ws.send({"t": "auth", "password": "pw"})
        ws.until(lambda m, s: m["t"] == "connected")
        # an admin can end it
        live = self.admin.json("GET", "/api/admin/sessions")[1]["terminals"]
        self.assertTrue(any(t["sid"] == d["sid"] for t in live))
        self.assertEqual(self.admin.json("POST", f"/api/admin/terminals/{d['sid']}/kill", {})[0], 200)
        ws.until(lambda m, s: m["t"] == "error" and "administrator" in m["m"])
        ws.close()

    def test_terminal_guards(self):
        eng = client_for(self.base, "Term Eng3", "Engineer")
        viewer = client_for(self.base, "Term Viewer", "Viewer")
        self.assertEqual(viewer.json("POST", "/api/terminal/sessions", {"host": "127.0.0.1", "username": "x"})[0], 403)
        st, d = eng.json("POST", "/api/terminal/sessions", {"host": "10.99.99.99", "username": "x"})
        self.assertEqual(st, 403)
        self.assertIn("not a configuration item in the CMDB", d["error"])
        for bad in ({"host": "a;rm -rf /", "username": "x"}, {"host": "127.0.0.1", "username": "a b"},
                    {"host": "127.0.0.1", "username": "x", "port": 70000}):
            self.assertEqual(eng.json("POST", "/api/terminal/sessions", bad)[0], 400, bad)
        self.admin.json("PUT", "/api/admin/settings", {"settings": {"terminal_allowed_networks": "10.99.0.0/16"}})
        self.assertEqual(eng.json("POST", "/api/terminal/sessions", {"host": "10.99.99.99", "username": "x"})[0], 200)
        self.admin.json("PUT", "/api/admin/settings", {"settings": {"terminal_allowed_networks": ""}})
        # websocket: wrong token, other user's session, cross-site origin, no session
        st, d = eng.json("POST", "/api/terminal/sessions", {"asset_id": self.asset["Record_ID"], "port": self.ssh.port,
                                                            "username": "ops"})
        ws = WSClient(self.base, f"/ws/terminal?sid={d['sid']}&token=wrong", eng.cookie)
        self.assertIn("not valid", ws.recv()["m"])
        ws.close()
        other = client_for(self.base, "Term Eng4", "Engineer")
        ws = WSClient(self.base, f"/ws/terminal?sid={d['sid']}&token={d['token']}", other.cookie)
        self.assertIn("not valid", ws.recv()["m"])
        ws.close()
        for cookie, origin in ((eng.cookie, "https://evil.example"), ("", None)):
            w = WSClient(self.base, f"/ws/terminal?sid={d['sid']}&token={d['token']}", cookie, origin=origin)
            self.assertEqual(w.status, 403)
            w.close()

    def test_rdp_file_and_guacamole_token(self):
        eng = client_for(self.base, "Rdp Eng", "Engineer")
        st, h, raw = http(self.base, "GET", f"/api/rdp/file?asset_id={self.asset['Record_ID']}&username=ops%0D%0Ax",
                          client=eng)
        self.assertEqual(st, 200)
        text = raw.decode("utf-16")
        self.assertIn("full address:s:127.0.0.1", text)
        self.assertIn("username:s:opsx", text)  # CR/LF stripped: no injected .rdp settings
        self.assertRegex(h["Content-Disposition"], r'^attachment; filename="[A-Za-z0-9._\-]+"$')
        self.assertEqual(eng.json("POST", "/api/rdp/guac", {"asset_id": self.asset["Record_ID"]})[0], 501)
        recs = self.admin.json("GET", "/api/admin/recordings")[1]["sessions"]
        self.assertTrue(any(r["type"] == "rdp-file" and r["user"] == "rdp.eng" for r in recs))
        # Guacamole JSON-auth token: verify by decrypting it the way Guacamole does
        from cryptography.hazmat.primitives import padding
        from cryptography.hazmat.primitives.ciphers import Cipher, algorithms, modes
        key_hex = "4c0b569e4c96df157eee1b65dd0e4d41"
        token = app.guac_encrypted_token(key_hex, {"username": "u", "connections": {}})
        key = bytes.fromhex(key_hex)
        dec = Cipher(algorithms.AES(key), modes.CBC(b"\0" * 16)).decryptor()
        unpad = padding.PKCS7(128).unpadder()
        plain = unpad.update(dec.update(base64.b64decode(token)) + dec.finalize()) + unpad.finalize()
        sig, data = plain[:32], plain[32:]
        import hmac as _hmac
        self.assertEqual(sig, _hmac.new(key, data, hashlib.sha256).digest())
        self.assertEqual(json.loads(data)["username"], "u")


class T14GuacamoleConfigured(LocalServerTestCase):
    config_overrides = {"guac_url": "https://guac.contoso.local/guacamole", "guac_secret": "4c0b569e4c96df157eee1b65dd0e4d41",
                        "guac_recording_path": "${HISTORY_PATH}/${HISTORY_UUID}"}

    def test_guacamole_launch(self):
        a = self.admin.json("POST", "/api/table/assets", {"row": {"Device": "win01", "IP_Address": "10.0.0.5"}})[1]["row"]
        st, d = self.admin.json("POST", "/api/rdp/guac", {"asset_id": a["Record_ID"], "username": "adm", "domain": "CORP"})
        self.assertEqual(st, 200, d)
        self.assertTrue(d["url"].startswith("https://guac.contoso.local/guacamole/?data="))
        recs = self.admin.json("GET", "/api/admin/recordings")[1]["sessions"]
        self.assertIn("Guacamole recording", next(r for r in recs if r["type"] == "rdp-guac")["note"])


class T15AdminShiftBackupConfig(LocalServerTestCase):
    def test_admin_user_lifecycle(self):
        st, d = self.admin.json("POST", "/api/admin/users", {"username": "x y", "role": "Viewer"})
        self.assertEqual(st, 400)
        st, d = self.admin.json("POST", "/api/admin/users", {"username": "carol", "role": "God"})
        self.assertEqual(st, 400)
        st, d = self.admin.json("POST", "/api/admin/users", {"username": "carol", "name": "Carol", "role": "Engineer",
                                                             "email": "carol@contoso.com"})
        self.assertEqual(st, 200)
        self.assertNotIn("password", d["user"])
        self.assertEqual(self.admin.json("POST", "/api/admin/users", {"username": "CAROL"})[0], 409)
        c = Client(self.base)
        c.login("carol", d["temp_password"])
        c.json("POST", "/api/me/password", {"current": d["temp_password"], "new": "Blue-Ocean-2026"})
        self.assertEqual(c.json("GET", "/api/state")[0], 200)
        st, r = self.admin.json("POST", "/api/admin/users/carol/reset", {})
        self.assertEqual(c.json("GET", "/api/state")[0], 401)  # reset signs the user out
        self.assertEqual(self.admin.json("DELETE", "/api/admin/users/carol")[0], 200)
        users = [u["username"] for u in self.admin.json("GET", "/api/admin/users")[1]["users"]]
        self.assertNotIn("carol", users)
        audit = self.admin.json("GET", "/api/audit?limit=50")[1]["audit"]
        self.assertTrue(any(a["Action"] == "Create user" and a["Target"] == "carol" for a in audit))
        sessions = self.admin.json("GET", "/api/admin/sessions")[1]["web"]
        self.assertTrue(sessions)
        st, s = self.admin.json("PUT", "/api/admin/settings", {"settings": {"session_idle_minutes": 1, "bogus": 1}})
        self.assertEqual(s["settings"]["session_idle_minutes"], 5)  # clamped to a safe minimum
        self.assertNotIn("bogus", s["settings"])

    def test_idle_timeout(self):
        c = client_for(self.base, "Idle Ivy", "Viewer")
        self.assertEqual(c.json("GET", "/api/state")[0], 200)
        sess = next(v for v in self.server.sessions.items.values() if v["user"] == "idle.ivy")
        sess["last"] -= 31 * 60
        self.assertEqual(c.json("GET", "/api/state")[0], 401)

    def test_clone_whole_shift(self):
        a = self.admin
        a.json("POST", "/api/hierarchy/node", {"level": "shift", "date": "2027-05-03", "shift": "Morning"})
        h = a.json("GET", "/api/hierarchy")[1]["hierarchy"]
        sec = h["years"]["2027"]["months"]["05"]["dates"]["2027-05-03"]["shifts"]["Morning"]["sections"][4]
        a.json("POST", "/api/hierarchy/row", {"date": "2027-05-03", "shift": "Morning", "section_id": sec["id"],
                                              "cells": {"Ticket": "INC-1"}})
        st, d = a.json("POST", "/api/hierarchy/shift/clone", {"date": "2027-05-03", "shift": "Morning",
                                                              "to_date": "2027-05-03", "to_shift": "Night"})
        self.assertEqual((st, d["rows"]), (200, 1))
        self.assertEqual(a.json("POST", "/api/hierarchy/shift/clone", {"date": "2027-05-03", "shift": "Morning",
                                                                       "to_shift": "Night"})[0], 409)
        st, d = a.json("POST", "/api/hierarchy/shift/clone", {"date": "2027-05-03", "shift": "Morning",
                                                              "to_date": "2027-06-01", "to_shift": "Afternoon",
                                                              "mode": "structure"})
        self.assertEqual(d["rows"], 0)
        h = a.json("GET", "/api/hierarchy")[1]["hierarchy"]
        night = h["years"]["2027"]["months"]["05"]["dates"]["2027-05-03"]["shifts"]["Night"]
        self.assertEqual(night["cloned_from"], "2027-05-03 Morning")
        self.assertEqual(night["sections"][4]["rows"][0]["carried_from"], "2027-05-03 Morning")
        self.assertEqual(len(h["years"]["2027"]["months"]["06"]["dates"]["2027-06-01"]["shifts"]["Afternoon"]["sections"]), 7)
        self.assertEqual(a.json("POST", "/api/hierarchy/shift/clone", {"date": "2027-05-03", "shift": "Morning"})[0], 400)

    def test_backups(self):
        self.admin.json("PUT", "/api/admin/settings", {"settings": {"backup_keep": 2}})
        files = [self.admin.json("POST", "/api/admin/backup", {})[1]["backup"]["file"] for _ in range(3)]
        self.assertEqual(len(set(files)), 3)
        self.assertFalse(os.path.exists(os.path.join(self.tmp, files[0])))  # pruned (keep = 2)
        self.assertTrue(os.path.exists(os.path.join(self.tmp, files[-1])))
        backup = json.load(open(os.path.join(self.tmp, files[-1])))
        self.assertIn("tables", backup)
        self.assertNotIn("users", backup)  # password hashes never go into exports or backups
        self.assertIsNone(self.server.store.auto_backup_if_due())  # already one today
        health = self.admin.json("GET", "/api/admin/health")[1]
        self.assertEqual(len(health["backups"]), 2)

    def test_settings_file(self):
        ini = os.path.join(self.tmp, "msp_settings.ini")
        with open(ini, "w") as fh:
            fh.write(app.SETTINGS_TEMPLATE.replace("port = 8080", "port = 9123").replace("mode = local", "mode = local")
                     .replace("[guacamole]\n; In-browser", "[guacamole]\n; In-browser").replace("url =\n", "url = https://g.example\n", 1))
        cfg = app.config_from_args(["--config", ini])
        self.assertEqual((cfg.port, cfg.storage, cfg.host), (9123, "local", "127.0.0.1"))
        self.assertTrue(cfg.data_dir.startswith(self.tmp))
        self.assertTrue(cfg.open_browser)
        self.assertEqual(app.config_from_args(["--config", ini, "--port", "9200"]).port, 9200)


class T16ShiftLogTemplateImportExport(LocalServerTestCase):
    def shift(self, date, shift):
        h = self.api("GET", "/api/hierarchy")["hierarchy"]
        try:
            return h["years"][date[:4]]["months"][date[5:7]]["dates"][date]["shifts"][shift]
        except KeyError:
            return None

    def cells(self, sec):
        names = {c["id"]: c["name"] for c in sec["columns"]}
        return [{names[k]: v for k, v in r["cells"].items() if k in names} for r in sec["rows"]]

    def commit(self, token, decisions=None, default_duplicate="skip"):
        return self.api("POST", "/api/import/commit", {"token": token, "decisions": decisions or {},
                                                       "default_duplicate": default_duplicate})["batch"]

    def test_template_formats(self):
        for fmt in ("xlsx", "csv", "txt", "json"):
            st, h, raw = http(self.base, "GET", f"/api/template/shiftlog?format={fmt}")
            self.assertEqual(st, 200, fmt)
            self.assertIn("Shift_Log_template", h["Content-Disposition"])
        wb = openpyxl.load_workbook(io.BytesIO(raw if False else http(self.base, "GET", "/api/template/shiftlog")[2]))
        self.assertEqual(wb.sheetnames, ["Shift_Log", "Instructions"])
        headers = [c.value for c in wb["Shift_Log"][2]]
        self.assertEqual(headers[:3], ["Date", "Shift", "Section"])
        for col in ("Check", "Event", "Blocked_On", "Current_Level", "Next_Step", "MOM", "Playbook"):
            self.assertIn(col, headers)
        self.assertEqual(wb["Shift_Log"].max_row - 2, 7)  # one example per standard section
        notes = " ".join(str(c.value) for r in wb["Instructions"].iter_rows() for c in r if c.value)
        self.assertIn("Section: Scheduled", notes)

    def test_template_round_trip_duplicates_custom_sections_and_undo(self):
        raw = http(self.base, "GET", "/api/template/shiftlog")[2]
        wb = openpyxl.load_workbook(io.BytesIO(raw))
        ws = wb["Shift_Log"]
        for r in range(3, ws.max_row + 1):  # move the examples to a fixed date
            ws.cell(row=r, column=1, value="2027-08-02")
        buf = io.BytesIO()
        wb.save(buf)
        pv = self.upload("my_shift.xlsx", buf.getvalue())
        sec = pv["sections"][0]
        self.assertEqual((sec["module"], sec["header_row"]), ("shiftlog", 2))
        self.assertEqual([i["status"] for i in sec["items"]], ["new"] * 7)
        self.assertIn("creates the shift", sec["items"][0]["label"])
        batch = self.commit(pv["token"])
        self.assertEqual(batch["sections"][0]["added"], 7)
        node = self.shift("2027-08-02", "Morning")
        self.assertEqual(len(node["sections"]), 7)
        inwork = next(x for x in node["sections"] if x["name"] == "In Work - All Clients")
        self.assertEqual(self.cells(inwork)[0]["Ticket"], "INC-104233")
        self.assertEqual(self.cells(inwork)[0]["Next_Step"], "Add 200 GB LUN")
        sched = next(x for x in node["sections"] if x["name"] == "Scheduled")
        self.assertNotIn("Ticket", self.cells(sched)[0])  # other sections' columns stay out
        # the same file again: every line is a duplicate; skip changes nothing, update edits in place
        pv = self.upload("my_shift.xlsx", buf.getvalue())
        self.assertEqual([i["status"] for i in pv["sections"][0]["items"]], ["duplicate"] * 7)
        self.assertEqual(self.commit(pv["token"])["sections"][0]["skipped"], 7)
        self.assertEqual(len(next(x for x in self.shift("2027-08-02", "Morning")["sections"]
                                  if x["name"] == "Scheduled")["rows"]), 1)
        # CSV: custom section, extra column, invalid lines, repeated line
        csv_text = ("Date,Shift,Section,Item,Details,Vendor case\r\n"
                    "02/08/2027,night,Vendor escalations,Storage array,Controller B fault,SR-77\r\n"
                    "2027-08-02,Night,In Work - All Clients,,,\r\n"
                    "2027-08-02,Lunch,Scheduled,x,,\r\n"
                    "2027-08-02,Night,,x,,\r\n"
                    "2027-08-02,Night,Vendor escalations,Storage array,Controller B fault,SR-77\r\n")
        pv = self.upload("night.csv", csv_text.encode())
        sec = pv["sections"][0]
        self.assertEqual(sec["module"], "shiftlog")
        self.assertEqual([i["status"] for i in sec["items"]], ["new", "invalid", "invalid", "invalid", "duplicate_in_file"])
        self.assertIn("Vendor escalations: Vendor case", sec["new_columns"])
        batch2 = self.commit(pv["token"])
        night = self.shift("2027-08-02", "Night")
        esc = next(x for x in night["sections"] if x["name"] == "Vendor escalations")
        self.assertFalse(esc["template"])
        self.assertEqual(self.cells(esc), [{"Item": "Storage array", "Details": "Controller B fault", "Vendor case": "SR-77"}])
        # undo removes the rows and the shift the import created
        self.assertEqual(self.api("POST", f"/api/imports/{batch2['id']}/undo", {})["removed"], 1)
        self.assertIsNone(self.shift("2027-08-02", "Night"))
        self.assertEqual(self.api("POST", f"/api/imports/{batch['id']}/undo", {})["removed"], 7)
        self.assertIsNone(self.shift("2027-08-02", "Morning"))

    def test_export_reimports_and_updates(self):
        self.api("POST", "/api/hierarchy/node", {"level": "shift", "date": "2027-09-01", "shift": "Afternoon"})
        node = self.shift("2027-09-01", "Afternoon")
        hold = next(x for x in node["sections"] if x["name"] == "On-Hold OR Customer Hold")
        self.api("POST", "/api/hierarchy/row", {"date": "2027-09-01", "shift": "Afternoon", "section_id": hold["id"],
                                                "cells": {"Ticket": "INC-5", "Blocked_On": "Vendor"}})
        raw = http(self.base, "GET", "/api/export/xlsx")[2]
        wb = openpyxl.load_workbook(io.BytesIO(raw))
        ws = wb["Shift_Log"]
        headers = [c.value for c in ws[2]]
        line = next(r for r in ws.iter_rows(min_row=3) if r[headers.index("Ticket")].value == "INC-5")
        self.assertEqual(line[headers.index("Day")].value, "Wednesday")
        self.assertTrue(line[headers.index("Row_ID")].value)
        line[headers.index("Blocked_On")].value = "Customer CAB"  # edit the exported file
        for name in list(wb.sheetnames):
            if name != "Shift_Log":
                del wb[name]
        buf = io.BytesIO()
        wb.save(buf)
        pv = self.upload("export_edit.xlsx", buf.getvalue())
        item = next(i for i in pv["sections"][0]["items"] if "INC-5" in i["label"] or i.get("changes"))
        self.assertEqual(item["status"], "duplicate")
        self.assertEqual(item["changes"], ["Blocked_On"])
        self.commit(pv["token"], default_duplicate="update")
        hold = next(x for x in self.shift("2027-09-01", "Afternoon")["sections"] if x["name"] == "On-Hold OR Customer Hold")
        self.assertEqual(self.cells(hold), [{"Ticket": "INC-5", "Blocked_On": "Customer CAB"}])

    def test_tab_text_and_json_lines(self):
        txt = "Date\tShift\tSection\tCheck\tStatus\r\n2027-10-05\tMorning\tShift Handover Overview\tAV alerts\tGreen\r\n"
        pv = self.upload("notes.txt", txt.encode())
        self.assertEqual(pv["sections"][0]["module"], "shiftlog")
        self.commit(pv["token"])
        js = json.dumps([{"Date": "2027-10-05", "Shift": "Morning", "Section": "Weekly Tasks", "Task": "Patch report"}])
        pv = self.upload("tasks.json", js.encode(), module="shiftlog")
        self.commit(pv["token"])
        node = self.shift("2027-10-05", "Morning")
        got = {x["name"]: self.cells(x) for x in node["sections"] if x["rows"]}
        self.assertEqual(got, {"Shift Handover Overview": [{"Check": "AV alerts", "Status": "Green"}],
                               "Weekly Tasks": [{"Task": "Patch report"}]})


# --------------------------------------------------------------------------------------
class T17ShiftRoster(LocalServerTestCase):
    """The roster module (msp_roster.py) plugged into the app through /api/roster."""

    SHEET = ("Day\tDATE\tMORNING SHIFT\tAFTERNOON SHIFT\tNIGHT SHIFT\tWEEK OFF\tLEAVE\n"
             "\t\t6AM TO 3PM\t1PM TO 10PM\t9:30 PM TO 6:30 AM\t\t\n"
             "Monday\t2-Nov-26\tALPHA, BRAVO\tCHARLIE\tDELTA\tECHO\t\n"
             "Tuesday\t3-Nov-26\tECHO, BRAVO\tALPHA\tCHARLIE\tDELTA\t\n")

    @classmethod
    def setUpClass(cls):
        super().setUpClass()
        cls.viewer = client_for(cls.base, "Rita Viewer", "Viewer")
        cls.eng = client_for(cls.base, "Eli Engineer", "Engineer")
        cls.eng2 = client_for(cls.base, "Ema Engineer", "Engineer")
        cls.lead = client_for(cls.base, "Lou Lead", "Lead")
        st, d = cls.lead.json("POST", "/api/roster", {
            "name": "NOC 24x7", "team": "NOC", "template": "three_shift", "start_date": "2026-11-02",
            "holidays": {"region": "", "custom": [{"date": "2026-11-04", "name": "Company day"}]},
            "members": [{"name": "Eli Engineer", "user": "eli.engineer", "offset": 0},
                        {"name": "Ema Engineer", "user": "ema.engineer", "offset": 7},
                        {"name": "Ola Other", "user": "ola.other", "offset": 14}]})
        assert st == 200, d
        cls.rid = d["roster"]["id"]

    def view(self, client=None, start="2026-11-02", mode="week"):
        st, d = (client or self.lead).json("GET", f"/api/roster/{self.rid}/view?start={start}&mode={mode}")
        self.assertEqual(st, 200, d)
        return d

    def member(self, name):
        return next(m["id"] for m in self.view()["grid"] if m["name"] == name)

    def test_01_list_and_view_for_viewer(self):
        st, d = self.viewer.json("GET", "/api/roster")
        self.assertEqual(st, 200)
        self.assertIn("NOC 24x7", [r["name"] for r in d["rosters"]])
        v = self.view(self.viewer)
        self.assertEqual((v["start"], v["end"]), ("2026-11-02", "2026-11-08"))
        self.assertEqual(len(v["days"]), 7)
        self.assertEqual(v["holidays"], {"2026-11-04": "Company day"})
        self.assertFalse(v["can"]["edit"])
        self.assertEqual(self.view(mode="month")["end"], "2026-11-30")

    def test_02_viewer_cannot_request(self):
        st, d = self.viewer.json("POST", f"/api/roster/{self.rid}/leave",
                                 {"member_id": self.member("Eli Engineer"), "start": "2026-11-20", "end": "2026-11-20"})
        self.assertEqual(st, 403, d)

    def test_03_engineer_leave_needs_lead(self):
        eli = self.member("Eli Engineer")
        st, d = self.eng.json("POST", f"/api/roster/{self.rid}/leave", {"member_id": eli, "start": "2026-11-05", "end": "2026-11-06"})
        self.assertEqual(st, 200, d)
        self.assertEqual(d["request"]["status"], "pending")
        rq = d["request"]["id"]
        st, d = self.eng.json("POST", f"/api/roster/{self.rid}/leave", {"member_id": self.member("Ema Engineer"),
                                                                          "start": "2026-11-05", "end": "2026-11-05"})
        self.assertEqual(st, 403, d)  # only their own leave
        st, d = self.eng.json("POST", f"/api/roster/{self.rid}/decide", {"kind": "leave", "id": rq, "approve": True})
        self.assertEqual(st, 403, d)
        st, d = self.lead.json("POST", f"/api/roster/{self.rid}/decide", {"kind": "leave", "id": rq, "approve": True})
        self.assertEqual((st, d["request"]["status"]), (200, "approved"))
        day = next(x for x in self.view()["days"] if x["date"] == "2026-11-05")
        self.assertIn("Eli Engineer", day["leave"])

    def test_04_swap_with_rule_check(self):
        st, d = self.eng.json("POST", f"/api/roster/{self.rid}/swap", {"member_a": self.member("Eli Engineer"), "date_a": "2026-11-03",
                                                                        "member_b": self.member("Ema Engineer"), "kind": "swap"})
        self.assertEqual(st, 200, d)
        self.assertEqual(d["request"]["status"], "awaiting_partner")  # Ema must accept first
        self.assertIsInstance(d["request"]["warnings"], list)
        st, d2 = self.eng.json("POST", f"/api/roster/{self.rid}/cancel", {"kind": "swap", "id": d["request"]["id"]})
        self.assertEqual((st, d2["request"]["status"]), (200, "cancelled"))

    def test_05_lead_changes_one_cell_and_restores(self):
        eli = self.member("Eli Engineer")
        st, _ = self.eng.json("POST", f"/api/roster/{self.rid}/override", {"member_id": eli, "date": "2026-11-07", "code": "N"})
        self.assertEqual(st, 403)
        st, _ = self.lead.json("POST", f"/api/roster/{self.rid}/override", {"member_id": eli, "date": "2026-11-07", "code": "N"})
        self.assertEqual(st, 200)
        cell = next(m for m in self.view()["grid"] if m["id"] == eli)["cells"][5]
        self.assertEqual(cell, {"code": "N", "source": "override"})
        hist = self.view()["history"]
        self.assertEqual(hist[-1]["action"], "Change shift")
        st, _ = self.lead.json("POST", f"/api/roster/{self.rid}/restore", {"index": hist[-1]["index"]})
        self.assertEqual(st, 200)
        self.assertNotEqual(next(m for m in self.view()["grid"] if m["id"] == eli)["cells"][5]["source"], "override")

    def test_06_settings_update_and_validation(self):
        v = self.view()
        shifts = v["config"]["shifts"]
        shifts[0]["holiday_min_staff"] = 0
        st, d = self.lead.json("PUT", f"/api/roster/{self.rid}", {"changes": {"shifts": shifts, "approvals": {"leave": "self", "swap": "lead"}}})
        self.assertEqual(st, 200, d)
        st, d = self.lead.json("PUT", f"/api/roster/{self.rid}", {"changes": {"cycle": {"type": "weekly"}}})
        self.assertEqual(st, 400, d)  # 21-day rotation does not fit a weekly cycle
        self.assertIn("weekly cycle", d["error"])
        st, d = self.eng.json("PUT", f"/api/roster/{self.rid}", {"changes": {"name": "x"}})
        self.assertEqual(st, 403, d)
        self.lead.json("PUT", f"/api/roster/{self.rid}", {"changes": {"approvals": {"leave": "lead", "swap": "lead"}}})

    def test_07_members_dedicated_to_one_roster(self):
        st, d = self.lead.json("POST", "/api/roster", {"name": "Windows", "team": "Windows", "template": "standard",
                                                       "members": [{"name": "Eli Engineer", "user": "eli.engineer"}]})
        self.assertEqual(st, 409, d)
        self.assertIn("already belongs", d["error"])

    def test_08_exports(self):
        st, hdrs, raw = self.eng.request("GET", f"/api/roster/{self.rid}/export?format=xlsx&start=2026-11-02&end=2026-11-08")
        self.assertEqual(st, 200)
        wb = openpyxl.load_workbook(io.BytesIO(raw))
        self.assertIn("By shift", wb.sheetnames)
        st, hdrs, raw = self.eng.request("GET", f"/api/roster/{self.rid}/export?format=html&start=2026-11-02&end=2026-11-08")
        self.assertEqual(st, 200)
        self.assertIn("text/html", hdrs["Content-Type"])
        self.assertIn(b"Company day", raw)
        st, _, _ = self.viewer.request("GET", f"/api/roster/{self.rid}/export?format=xlsx")
        self.assertEqual(st, 403)

    def test_09_import_sheet(self):
        data = base64.b64encode(self.SHEET.encode()).decode()
        st, d = self.lead.json("POST", "/api/roster/import", {"filename": "nov.txt", "data": data, "team": "Ops", "name": "Ops Nov"})
        self.assertEqual(st, 200, d)
        self.assertEqual((d["roster"]["members"], d["layout"]), (5, "by_shift"))
        st, d = self.lead.json("POST", "/api/roster/import", {"filename": "nov.exe", "data": data, "team": "Ops", "name": "x"})
        self.assertEqual(st, 400, d)
        st, d = self.eng.json("POST", "/api/roster/import", {"filename": "nov.txt", "data": data, "team": "Ops", "name": "y"})
        self.assertEqual(st, 403, d)

    def test_10_my_shifts_audit_health_backup(self):
        st, d = self.eng.json("GET", "/api/roster/mine?days=31")
        self.assertEqual(st, 200)
        self.assertTrue(all(x["roster"] == "NOC 24x7" for x in d["shifts"]))
        st, d = self.lead.json("GET", "/api/audit")
        self.assertTrue(any(a.get("Area") == "Roster" for a in d["audit"]))
        st, d = self.admin.json("GET", "/api/admin/health")
        self.assertTrue(d["checks"]["roster"])
        st, _, raw = self.admin.request("GET", "/api/export/json")
        self.assertIn(self.rid, json.loads(raw)["rosters"])

    def test_13_consent_then_assigned_approver(self):
        kim = client_for(self.base, "Kim Approver", "Lead")
        st, _ = self.lead.json("PUT", f"/api/roster/{self.rid}", {"changes": {"approvers": ["kim.approver"]}})
        self.assertEqual(st, 200)
        try:
            st, d = self.eng.json("POST", f"/api/roster/{self.rid}/swap", {"member_a": self.member("Eli Engineer"), "date_a": "2026-11-12",
                                                                            "member_b": self.member("Ema Engineer"), "kind": "cover"})
            self.assertEqual((st, d["request"]["status"], d["request"]["partner"]), (200, "awaiting_partner", "Ema Engineer"))
            rq = d["request"]["id"]
            st, _ = self.lead.json("POST", f"/api/roster/{self.rid}/decide", {"kind": "swap", "id": rq, "approve": True})
            self.assertEqual(st, 409)  # partner has not accepted yet
            st, d = self.eng.json("POST", f"/api/roster/{self.rid}/consent", {"id": rq, "accept": True})
            self.assertEqual(st, 403)  # only Ema can accept
            st, mine = self.eng2.json("GET", "/api/roster/mine")
            self.assertIn(rq, [t["id"] for t in mine["todo"] if t["action"] == "consent"])
            st, d = self.eng2.json("POST", f"/api/roster/{self.rid}/consent", {"id": rq, "accept": True})
            self.assertEqual((st, d["request"]["status"], d["request"]["approver"]), (200, "pending", "kim.approver"))
            st, d = self.lead.json("POST", f"/api/roster/{self.rid}/decide", {"kind": "swap", "id": rq, "approve": True})
            self.assertEqual(st, 403)  # Lou is a Lead, but the request is assigned to Kim
            self.assertIn("assigned to kim.approver", d["error"])
            st, mine = kim.json("GET", "/api/roster/mine")
            self.assertIn(rq, [t["id"] for t in mine["todo"] if t["action"] == "approve"])
            st, d = kim.json("POST", f"/api/roster/{self.rid}/decide", {"kind": "swap", "id": rq, "approve": True})
            self.assertEqual((st, d["request"]["status"]), (200, "approved"))
            st, d = self.eng.json("POST", f"/api/roster/{self.rid}/leave", {"member_id": self.member("Eli Engineer"), "start": "2026-11-25",
                                                                             "end": "2026-11-25", "approver": "someone.else"})
            self.assertEqual(st, 400, d)
        finally:
            self.lead.json("PUT", f"/api/roster/{self.rid}", {"changes": {"approvers": []}})

    def test_14_template_download(self):
        st, hdrs, raw = self.eng.request("GET", "/api/roster/template?layout=by_shift&format=xlsx")
        self.assertEqual(st, 200)
        wb = openpyxl.load_workbook(io.BytesIO(raw))
        self.assertEqual(wb.sheetnames, ["Roster", "Instructions"])
        self.assertEqual(wb["Roster"]["C2"].value, "6AM TO 3PM")
        st, hdrs, raw = self.eng.request("GET", "/api/roster/template?layout=by_shift&format=csv")
        self.assertEqual(st, 200)
        self.assertIn("text/csv", hdrs["Content-Type"])
        st, _, raw = self.eng.request("GET", "/api/roster/template?layout=by_person&format=xlsx")
        self.assertEqual(st, 200)
        self.assertIn("Shifts", openpyxl.load_workbook(io.BytesIO(raw)).sheetnames)
        st, _, _ = self.viewer.request("GET", "/api/roster/template?layout=by_shift&format=xlsx")
        self.assertEqual(st, 403)
        st, _, _ = self.eng.request("GET", "/api/roster/template?layout=sideways")
        self.assertEqual(st, 400)

    def test_15_partly_filled_month(self):
        sheet = ("Day\tDATE\tMORNING SHIFT\tAFTERNOON SHIFT\tNIGHT SHIFT\tWEEK OFF\tLEAVE\n"
                 "\t\t6AM TO 3PM\t1PM TO 10PM\t9:30 PM TO 6:30 AM\t\t\n"
                 + "".join(f"Day\t{d}-Nov-26\t" + ("PIA, QUIN\tRUE\tSAM\tTOVE\t\n" if 9 <= d <= 22 else "\t\t\t\t\n") for d in range(1, 31)))
        st, d = self.lead.json("POST", "/api/roster/import", {"filename": "nov.txt", "data": base64.b64encode(sheet.encode()).decode(),
                                                                "team": "SA", "name": "SA Nov"})
        self.assertEqual(st, 200, d)
        self.assertTrue(d["warnings"][0].startswith("The sheet covers 2026-11-09 to 2026-11-22"))
        rid = d["roster"]["id"]
        st, v = self.lead.json("GET", f"/api/roster/{rid}/view?start=2026-11-15&mode=month")
        errors = [c for c in v["checks"] if c["rule"] == "coverage"]
        self.assertEqual(errors, [])
        self.assertEqual(sum(1 for x in v["days"] if not x["has_data"]), 16)
        self.assertIsNone(v["need"]["2026-11-05"]["M"])
        self.assertIn("checks_grouped", v)
        st, hdrs, raw = self.eng.request("GET", "/api/roster/template?layout=by_shift&format=csv&month=2026-02")
        self.assertEqual(raw.decode("utf-8-sig").count("-Feb-26"), 28)
        self.lead.json("DELETE", f"/api/roster/{rid}")

    def test_16_excel_upload_needs_no_temporary_file(self):
        """Reported on Windows: the roster upload failed because a temporary file could not be deleted."""
        demo = os.path.join(os.path.dirname(os.path.abspath(app.__file__)), "Demo_Data", "08_Shift_Roster_Nov_Dec_2026.xlsx")
        with open(demo, "rb") as fh:
            data = base64.b64encode(fh.read()).decode()
        import tempfile as tf
        real = tf.mkstemp
        tf.mkstemp = lambda *a, **k: (_ for _ in ()).throw(AssertionError("no temporary file should be needed"))
        try:
            st, d = self.lead.json("POST", "/api/roster/import", {"filename": "08_Shift_Roster_Nov_Dec_2026.xlsx", "data": data,
                                                                   "team": "Demo", "name": "Demo upload"})
        finally:
            tf.mkstemp = real
        self.assertEqual(st, 200, d)
        self.assertEqual((d["roster"]["members"], d["layout"]), (10, "by_shift"))
        self.lead.json("DELETE", f"/api/roster/{d['roster']['id']}")
        st, d = self.lead.json("POST", "/api/roster/import", {"filename": "x.xlsx", "data": base64.b64encode(b"hello").decode(),
                                                               "team": "Demo", "name": "Bad"})
        self.assertEqual(st, 400, d)
        self.assertIn("not an Excel workbook", d["error"])

    def test_11_page_has_roster(self):
        st, _, body = http(self.base, "GET", "/")
        self.assertIn("['roster','Shift roster','roster']", body.decode())

    def test_12_app_runs_without_roster_module(self):
        saved = self.server.rosters
        self.server.rosters = None
        try:
            st, d = self.lead.json("GET", "/api/roster")
            self.assertEqual(st, 404)
            self.assertIn("not installed", d["error"])
            st, _ = self.lead.json("GET", "/api/state")
            self.assertEqual(st, 200)
        finally:
            self.server.rosters = saved


# --------------------------------------------------------------------------------------
class T18AssistantAndOffline(LocalServerTestCase):
    """The built-in assistant (msp_ai.py) and the offline page: nothing loads from the internet."""

    @classmethod
    def setUpClass(cls):
        super().setUpClass()
        cls.viewer = client_for(cls.base, "Vera Viewer", "Viewer")
        cls.lead = client_for(cls.base, "Leo Lead", "Lead")

    def test_01_page_loads_nothing_from_the_internet(self):
        st, hdrs, body = http(self.base, "GET", "/")
        page = body.decode()
        for host in ("fonts.googleapis.com", "fonts.gstatic.com", "cdn.jsdelivr.net", "unpkg.com"):
            self.assertNotIn(host, page)
            self.assertNotIn(host, hdrs.get("Content-Security-Policy", ""))
        csp = hdrs["Content-Security-Policy"]
        self.assertNotIn("https:", csp)
        self.assertIn("/static/fonts/ibm-plex-sans-latin-400-normal.woff2", page)
        self.assertIn("'/static/vendor/xterm.js'", page)
        self.assertIn('data-act="ask-open"', page)

    def test_02_static_files_are_served_safely(self):
        st, hdrs, body = http(self.base, "GET", "/static/vendor/xterm.js")
        self.assertEqual(st, 200)
        self.assertIn("javascript", hdrs["Content-Type"])
        import hashlib
        digest = "sha384-" + base64.b64encode(hashlib.sha384(body).digest()).decode()
        self.assertIn(digest, app.HTML_PAGE)  # the integrity check in the page matches the shipped file
        st, hdrs, _ = http(self.base, "GET", "/static/fonts/ibm-plex-mono-latin-400-normal.woff2")
        self.assertEqual((st, hdrs["Content-Type"]), (200, "font/woff2"))
        for bad in ("/static/../msp_runbook.py", "/static/%2e%2e/msp_users.json", "/static/vendor/../../msp_users.json",
                    "/static/vendor/", "/static/x.py", "/static/fonts/missing.woff2", "/static//etc/passwd"):
            st, _, _ = http(self.base, "GET", bad)
            self.assertIn(st, (400, 404), bad)

    def test_03_viewer_can_ask_and_it_is_audited(self):
        st, d = self.viewer.json("POST", "/api/ai/ask", {"question": "how do I fix ASM diskgroup full"})
        self.assertEqual(st, 200, d)
        a = d["answer"]
        self.assertTrue(a["found"])
        self.assertEqual(a["sources"][0]["label"], "Runbook")
        self.assertIn("Re-register the credential", a["text"])
        st, au = self.lead.json("GET", "/api/audit")
        self.assertTrue(any(x.get("Area") == "Assistant" and "ASM diskgroup" in x.get("Target", "") for x in au["audit"]))

    def test_04_bad_questions_and_status(self):
        st, d = self.viewer.json("POST", "/api/ai/ask", {"question": ""})
        self.assertEqual(st, 400)
        st, d = self.viewer.json("GET", "/api/ai/status")
        self.assertEqual((st, d["mode"]), (200, "search"))
        st, _, _ = Client(self.base).request("POST", "/api/ai/ask", {"question": "hi"})  # not signed in
        self.assertEqual(st, 401)

    def test_05_rate_limit(self):
        client = client_for(self.base, "Rae Rapid", "Viewer")
        codes = [client.json("POST", "/api/ai/ask", {"question": "known issues"})[0] for _ in range(21)]
        self.assertEqual(codes[:20], [200] * 20)
        self.assertEqual(codes[20], 429)

    def test_06_settings_switch_off_and_refuse_public_model(self):
        st, d = self.admin.json("PUT", "/api/admin/settings", {"settings": {"ai_mode": "model", "ai_endpoint": "http://8.8.8.8:11434"}})
        self.assertEqual(st, 400, d)
        self.assertIn("public address", d["error"])
        st, d = self.admin.json("PUT", "/api/admin/settings", {"settings": {"ai_mode": "off"}})
        self.assertEqual(st, 200, d)
        try:
            st, d = self.viewer.json("POST", "/api/ai/ask", {"question": "known issues"})
            self.assertEqual(st, 403)
            self.assertIn("switched off", d["error"])
        finally:
            self.admin.json("PUT", "/api/admin/settings", {"settings": {"ai_mode": "search"}})
        st, d = self.viewer.json("POST", "/api/admin/ai/test", {"settings": {"ai_endpoint": "http://127.0.0.1:9"}})
        self.assertEqual(st, 403)  # admin only
        st, d = self.admin.json("POST", "/api/admin/ai/test", {"settings": {"ai_endpoint": "http://127.0.0.1:9", "ai_model": "x"}})
        self.assertEqual(st, 200)
        self.assertIn("did not answer", d["status"]["detail"])

    def test_07_who_is_on_now_and_health(self):
        st, d = self.lead.json("POST", "/api/roster", {"name": "Now 24x7", "team": "NOC", "template": "three_shift",
                                                       "start_date": (dt.date.today() - dt.timedelta(days=30)).isoformat(),
                                                       "members": [{"name": f"P{i}", "offset": o} for i, o in enumerate([0, 1, 6, 11, 16])]})
        self.assertEqual(st, 200, d)
        st, d = self.viewer.json("GET", "/api/roster/now")
        self.assertEqual(st, 200)
        self.assertGreaterEqual(len(d["now"]["shifts"]), 1)  # three shifts cover 24 hours, so one is always running
        self.assertIn("minutes_left", d["now"]["shifts"][0])
        st, h = self.admin.json("GET", "/api/admin/health")
        self.assertTrue(h["checks"]["assistant"])
        self.assertTrue(h["checks"]["local_assets"])

    def test_08_app_runs_without_the_assistant(self):
        saved = self.server.assistant
        self.server.assistant = None
        try:
            st, d = self.viewer.json("POST", "/api/ai/ask", {"question": "known issues"})
            self.assertEqual(st, 404)
            self.assertIn("not installed", d["error"])
            st, _ = self.viewer.json("GET", "/api/state")
            self.assertEqual(st, 200)
        finally:
            self.server.assistant = saved


# --------------------------------------------------------------------------------------
class T19NewInterface(LocalServerTestCase):
    """The new front end (frontend/, built to static/app) served at /app/."""

    def test_01_page_and_assets(self):
        st, hdrs, body = Client(self.base).request("GET", "/app/")
        self.assertEqual(st, 200)
        page = body.decode()
        self.assertIn('<div id="root"></div>', page)
        csp = hdrs["Content-Security-Policy"]
        self.assertIn("script-src 'self'", csp)
        self.assertNotIn("https:", csp)
        js = re.search(r'src="(/app/assets/[^"]+\.js)"', page).group(1)
        css = re.search(r'href="(/app/assets/[^"]+\.css)"', page).group(1)
        st, h, _ = Client(self.base).request("GET", js)
        self.assertEqual((st, "javascript" in h["Content-Type"]), (200, True))
        st, h, raw = Client(self.base).request("GET", css)
        self.assertEqual(st, 200)
        self.assertIn(b"/static/fonts/ibm-plex-sans-latin-400-normal.woff2", raw)  # fonts load from this server

    def test_02_no_escape_from_the_app_folder(self):
        for bad in ("/app/../msp_users.json", "/app/assets/../../msp_runbook.py", "/app/%2e%2e/%2e%2e/msp_users.json",
                    "/app/assets/", "/app/index.py", "/app/a/b/c/d/e.js"):
            st, _, _ = Client(self.base).request("GET", bad)
            self.assertIn(st, (400, 404), bad)

    def test_03_missing_build_is_explained(self):
        saved = app.STATIC_DIR
        app.STATIC_DIR = tempfile.mkdtemp()
        try:
            st, _, body = Client(self.base).request("GET", "/app/")
            self.assertEqual(st, 404)
            self.assertIn(b"not built yet", body)
        finally:
            app.STATIC_DIR = saved

    def test_04_classic_view_links_to_it(self):
        st, _, body = http(self.base, "GET", "/")
        self.assertIn('href="/app/"', body.decode())


# --------------------------------------------------------------------------------------
class T20Teams(LocalServerTestCase):
    """Teams, customers and projects (msp_org.py) through the API, with the permission rules."""

    @classmethod
    def setUpClass(cls):
        super().setUpClass()
        cls.lou = client_for(cls.base, "Lou Lead", "Lead")
        cls.kim = client_for(cls.base, "Kim Lead", "Lead")
        cls.eli = client_for(cls.base, "Eli Engineer", "Engineer")
        st, d = cls.admin.json("POST", "/api/org/teams", {"name": "DBA", "leads": ["lou.lead"]})
        assert st == 200, d
        cls.dba = d["team"]["id"]
        cls.sa = cls.admin.json("POST", "/api/org/teams", {"name": "Unix SA", "leads": ["kim.lead"]})[1]["team"]["id"]
        cls.acme = cls.admin.json("POST", "/api/org/customers", {"name": "Acme Corp"})[1]["customer"]["id"]
        cls.prj = cls.admin.json("POST", f"/api/org/customers/{cls.acme}/projects", {"name": "ERP"})[1]["project"]["id"]
        cls.admin.json("PUT", f"/api/org/customers/{cls.acme}", {"teams": [cls.dba], "primary_team": cls.dba})

    def test_01_admin_only_setup(self):
        for client in (self.lou, self.eli):
            st, _ = client.json("POST", "/api/org/customers", {"name": "Nope"})
            self.assertEqual(st, 403)
        st, d = self.lou.json("POST", "/api/org/teams", {"name": "Nope"})
        self.assertEqual(st, 403)

    def test_02_lead_adds_creates_and_assigns(self):
        st, d = self.lou.json("POST", f"/api/org/teams/{self.dba}/members", {"username": "eli.engineer"})
        self.assertEqual(st, 200, d)
        st, d = self.lou.json("POST", f"/api/org/teams/{self.dba}/members", {"create": {"username": "neha.dba", "name": "Neha"}})
        self.assertEqual(st, 200, d)
        self.assertTrue(d["temp_password"])
        st, users = self.admin.json("GET", "/api/admin/users")
        self.assertEqual(next(u for u in users["users"] if u["username"] == "neha.dba")["role"], "Engineer")
        st, d = self.lou.json("POST", f"/api/org/teams/{self.dba}/members", {"create": {"username": "boss", "role": "Admin"}})
        self.assertEqual(st, 400, d)
        st, d = self.lou.json("PUT", f"/api/org/teams/{self.dba}/members/eli.engineer", {"projects": [self.prj]})
        self.assertEqual((st, d["projects"]), (200, [self.prj]))

    def test_03_lead_cannot_touch_other_teams(self):
        st, d = self.lou.json("POST", f"/api/org/teams/{self.sa}/members", {"username": "eli.engineer"})
        self.assertEqual(st, 403, d)
        st, d = self.lou.json("POST", f"/api/org/teams/{self.sa}/members", {"create": {"username": "sneaky", "name": "S"}})
        self.assertEqual(st, 403, d)
        st, users = self.admin.json("GET", "/api/admin/users")
        self.assertFalse(any(u["username"] == "sneaky" for u in users["users"]))  # no account was created

    def test_04_views_and_people_list(self):
        self.lou.json("POST", f"/api/org/teams/{self.dba}/members", {"username": "eli.engineer"})
        st, v = self.eli.json("GET", "/api/org")
        self.assertEqual([t["name"] for t in v["teams"]], ["DBA"])
        self.assertFalse(v["teams"][0]["can_manage"])
        st, _ = self.eli.json("GET", "/api/org/people")
        self.assertEqual(st, 403)
        st, d = self.lou.json("GET", "/api/org/people")
        self.assertEqual(st, 200)
        self.assertFalse(any(p["role"] == "Admin" for p in d["people"]))

    def test_05_admin_not_a_member_and_audit(self):
        st, d = self.admin.json("POST", f"/api/org/teams/{self.dba}/members", {"username": "admin"})
        self.assertEqual(st, 400)
        st, a = self.admin.json("GET", "/api/audit")
        self.assertTrue(any(x.get("Area") == "Teams" for x in a["audit"]))

    def test_06_deleting_a_user_removes_them_from_teams(self):
        self.lou.json("POST", f"/api/org/teams/{self.dba}/members", {"create": {"username": "temp.user", "name": "Temp"}})
        st, _ = self.admin.json("DELETE", "/api/admin/users/temp.user")
        self.assertEqual(st, 200)
        st, v = self.admin.json("GET", "/api/org")
        self.assertFalse(any(m["username"] == "temp.user" for t in v["teams"] for m in t["members"]))

    def test_07_backup_and_page(self):
        st, _, raw = self.admin.request("GET", "/api/export/json")
        self.assertIn(self.dba, json.loads(raw)["org"]["teams"])
        st, _, body = http(self.base, "GET", "/")
        self.assertIn("['teams','Teams','users']", body.decode())


# --------------------------------------------------------------------------------------
class T21PortClash(unittest.TestCase):
    """Reported on Windows: the browser showed another program's 'Directory listing' on port 8080."""

    def test_busy_port_moves_to_the_next_free_one(self):
        import socket as so
        busy = so.socket(so.AF_INET, so.SOCK_STREAM)
        busy.bind(("127.0.0.1", 0))
        busy.listen(1)
        port = busy.getsockname()[1]
        tmp = tempfile.mkdtemp()
        try:
            self.assertFalse(app.port_free("127.0.0.1", port))
            srv = app.create_server(app.Config(host="127.0.0.1", port=port, storage="local", data_dir=tmp, refresh_interval=0,
                                               quiet=True, seed=False, backup_thread=False, setup_code_file=os.path.join(tmp, "c.txt")))
            try:
                self.assertNotEqual(srv.server_address[1], port)
                self.assertIn(f"Port {port} is already used by another program", srv.port_note)
            finally:
                srv.server_close()
        finally:
            busy.close()

    def test_exclusive_on_windows_and_exact_address_shown(self):
        self.assertEqual(app.AppServer.allow_reuse_address, os.name != "nt")
        src = open(app.__file__, encoding="utf-8").read()
        self.assertIn('shown = "127.0.0.1" if cfg.host in', src)  # never 'localhost', which may be another program on ::1


# --------------------------------------------------------------------------------------
class T22TeamShiftLogAndHandovers(LocalServerTestCase):
    """Isolation step 1: once teams exist, each team has its own shift log and handovers."""

    @classmethod
    def setUpClass(cls):
        super().setUpClass()
        # data from before teams existed
        cls.admin.json("POST", "/api/hierarchy/node", {"level": "shift", "date": "2026-10-01", "shift": "Morning"})
        st, d = cls.admin.json("POST", "/api/table/Shift_Transitions", {"Date": "2026-10-01", "Shift": "Morning", "Notes": "old handover"})
        assert st == 200, d
        cls.la = client_for(cls.base, "Lena A", "Lead")
        cls.lb = client_for(cls.base, "Liam B", "Lead")
        cls.ea = client_for(cls.base, "Erin A", "Engineer")
        cls.eb = client_for(cls.base, "Ezra B", "Engineer")
        cls.em = client_for(cls.base, "Mo Both", "Engineer")
        cls.ta = cls.admin.json("POST", "/api/org/teams", {"name": "Team A", "leads": ["lena.a"]})[1]["team"]["id"]
        cls.tb = cls.admin.json("POST", "/api/org/teams", {"name": "Team B", "leads": ["liam.b"]})[1]["team"]["id"]
        for team, lead, members in ((cls.ta, cls.la, ["erin.a", "mo.both"]), (cls.tb, cls.lb, ["ezra.b", "mo.both"])):
            for m in members:
                lead.json("POST", f"/api/org/teams/{team}/members", {"username": m})
        time.sleep(3.2)  # the team file is cached for 3 seconds

    def state(self, client, team=""):
        st, _, raw = client.request("GET", "/api/state", headers={"X-Team": team} if team else None)
        self.assertEqual(st, 200)
        return json.loads(raw)

    @staticmethod
    def days(d):
        return sorted(dd for y in d["hierarchy"]["years"].values() for m in y["months"].values() for dd in m["dates"])

    def test_01_each_team_has_its_own_shift_log(self):
        st, _ = self.ea.json("POST", "/api/hierarchy/node", {"level": "shift", "date": "2026-11-20", "shift": "Night"})
        self.assertEqual(st, 200)
        self.assertIn("2026-11-20", self.days(self.state(self.ea)))
        self.assertNotIn("2026-11-20", self.days(self.state(self.eb)))
        self.assertIn("2026-11-20", self.days(self.state(self.em, self.ta)))   # in both teams: switches with the header
        self.assertNotIn("2026-11-20", self.days(self.state(self.em, self.tb)))

    def test_02_handovers_are_team_only_and_hidden_from_others(self):
        st, d = self.ea.json("POST", "/api/table/Shift_Transitions", {"Date": "2026-11-20", "Shift": "Night", "Notes": "A only"})
        self.assertEqual(st, 200, d)
        rid = d["row"]["Record_ID"] if "row" in d else d["Record_ID"]
        rows = lambda c, t="": [r for r in self.state(c, t)["tables"]["Shift_Transitions"]["rows"] if r.get("Notes") == "A only"]  # noqa: E731
        self.assertEqual(len(rows(self.ea)), 1)
        self.assertEqual(rows(self.eb), [])
        self.assertEqual(rows(self.eb, self.ta), [])  # asking for another team's id is ignored
        for who, method, path in ((self.eb, "PUT", f"/api/table/Shift_Transitions/{rid}"),
                                  (self.lb, "DELETE", f"/api/table/Shift_Transitions/{rid}"),  # deleting needs a Lead
                                  (self.eb, "GET", f"/api/history?key=table/Shift_Transitions/{rid}")):
            st, _ = who.json(method, path, {"Notes": "changed"} if method == "PUT" else None)
            self.assertEqual(st, 404, (method, path))  # another team's record does not exist for you
        self.assertEqual(rows(self.ea)[0]["Notes"], "A only")

    def test_03_admin_sees_no_team_data(self):
        d = self.state(self.admin)
        self.assertTrue(d["team_scope"])
        self.assertIsNone(d["team"])
        self.assertEqual(d["tables"]["Shift_Transitions"]["rows"], [])
        self.assertEqual(self.days(d), [])
        st, _ = self.admin.json("POST", "/api/hierarchy/node", {"level": "shift", "date": "2026-11-21", "shift": "Morning"})
        self.assertEqual(st, 403)

    def test_04_old_data_moves_into_a_team_by_its_lead_only(self):
        d = self.state(self.lb)
        before = d["legacy"]
        self.assertGreaterEqual(before["shift_days"], 1)
        self.assertGreaterEqual(before["handovers"], 1)  # ours, plus any built-in sample data
        st, _ = self.eb.json("POST", "/api/org/claim-legacy", {})
        self.assertEqual(st, 403)  # engineers cannot
        st, d = self.lb.json("POST", "/api/org/claim-legacy", {})
        self.assertEqual((st, d["moved"]["shift_days"], d["moved"]["handovers"]), (200, before["shift_days"], before["handovers"]))
        self.assertEqual(self.state(self.lb)["legacy"], {"shift_days": 0, "handovers": 0})
        self.assertIn("2026-10-01", self.days(self.state(self.eb)))
        self.assertNotIn("2026-10-01", self.days(self.state(self.ea)))
        self.assertTrue(any(r.get("Notes") == "old handover" for r in self.state(self.eb)["tables"]["Shift_Transitions"]["rows"]))

    def test_05_shift_log_import_and_undo_stay_in_the_team(self):
        csv_text = "Date,Shift,Section,Check,Client,Status,Details,Owner\n2026-11-25,Morning,Shift Handover Overview,Backups,All,Green,ok,Erin\n"
        ctype, body = multipart({"module": "auto"}, {"file": ("shift.csv", csv_text.encode())})
        st, _, raw = self.la.request("POST", "/api/import/preview", body, ctype=ctype)  # importing needs a Lead
        pv = json.loads(raw)
        self.assertEqual(st, 200, pv)
        st, b = self.la.json("POST", "/api/import/commit", {"token": pv["token"], "decisions": {}, "default_duplicate": "skip"})
        self.assertEqual(st, 200, b)
        self.assertIn("2026-11-25", self.days(self.state(self.ea)))
        self.assertNotIn("2026-11-25", self.days(self.state(self.eb)))
        batch = b.get("batch", b).get("id")
        st, _, _ = self.em.request("GET", "/api/state", headers={"X-Team": self.tb})  # someone else working in team B meanwhile
        st, u = self.la.json("POST", f"/api/imports/{batch}/undo", {})
        self.assertEqual(st, 200, u)
        y = self.state(self.ea)["hierarchy"]["years"].get("2026", {}).get("months", {}).get("11", {}).get("dates", {})
        self.assertNotIn("Morning", (y.get("2026-11-25") or {}).get("shifts", {}))  # the imported shift is gone from team A

    def test_06_customers_and_services_page(self):
        st, _, body = http(self.base, "GET", "/")
        page = body.decode()
        self.assertIn("['customers','Customers and services','briefcase']", page)
        self.assertIn("function teamBar()", page)


if __name__ == "__main__":
    unittest.main(verbosity=2)
