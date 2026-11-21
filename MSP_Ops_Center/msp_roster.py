#!/usr/bin/env python3
"""MSP Ops Center: shift roster module (version 1, standalone).

A configurable shift roster for one team per roster:

* Cycle: daily, weekly, fortnightly, monthly (restarts on the 1st) or a custom
  number of days.
* Shift patterns: any code, name, start and end time (overnight allowed),
  working weekdays, and minimum / maximum people.
* Rotation: a sequence of shift codes and OFF days that each member walks
  through, staggered by a per-member offset.
* Rules: minimum rest hours between shifts, forbidden shift-to-shift
  sequences, maximum consecutive working days, and whether members may be
  shared with other rosters.
* Leave, swaps and cover: each can need Lead approval or be self-recorded.
* On-call: any number of roles (Primary, Secondary, ...), rotating every N
  days, skipping people on leave, plus an escalation list.
* Views: day / week / month calendar, per-person schedule, coverage check.
* Exports: Excel workbook (with live COUNTIF coverage formulas) and a
  printable HTML page.

Every change is written with optimistic concurrency and the previous state of
the roster is kept as a version (30 per roster) that can be restored.

It uses the same storage interface as msp_runbook.py (``read(name)`` returns
``(bytes, version)``; ``write(name, data, expected_version)``), so it can share
that app's local folder or SharePoint storage. On its own it uses FileStorage.

Command line:
    python msp_roster.py demo   --out DIR      # sample roster + exports
    python msp_roster.py export --data FILE --roster ID --start 2026-11-01 --end 2026-11-30 --out DIR
"""
from __future__ import annotations

import argparse
import copy
import datetime as dt
import html
import io
import json
import os
import re
import threading
import uuid
import zipfile

MODULE_VERSION = "1.0.0"
ROSTER_FILE = "msp_roster.json"
MAX_VERSIONS = 30
MAX_AUDIT = 2000
MAX_RANGE_DAYS = 400
OFF = "OFF"
LEAVE = "LEAVE"
RESERVED_CODES = {OFF, LEAVE}
CYCLE_TYPES = ("daily", "weekly", "fortnightly", "monthly", "custom")
CYCLE_DAYS = {"daily": 1, "weekly": 7, "fortnightly": 14}
APPROVAL_MODES = ("lead", "self")
ROLE_RANK = {"Viewer": 1, "Engineer": 2, "Lead": 3, "Admin": 4}
WEEKDAYS = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]
TIME_RE = re.compile(r"^([01]\d|2[0-3]):([0-5]\d)$")
CODE_RE = re.compile(r"^[A-Z0-9][A-Z0-9_-]{0,7}$")
COLOR_RE = re.compile(r"^#[0-9A-Fa-f]{6}$")


class RosterError(Exception):
    """A request that cannot be carried out; status follows HTTP meaning."""

    def __init__(self, status: int, message: str):
        super().__init__(message)
        self.status = status
        self.message = message


class ConflictError(Exception):
    """The stored file changed since it was read."""


def now_iso() -> str:
    return dt.datetime.now().astimezone().isoformat(timespec="seconds")


def new_id(prefix: str) -> str:
    return prefix + uuid.uuid4().hex[:10]


def parse_date(value) -> dt.date:
    if isinstance(value, dt.datetime):
        return value.date()
    if isinstance(value, dt.date):
        return value
    try:
        return dt.date.fromisoformat(str(value).strip())
    except ValueError:
        raise RosterError(400, f"'{value}' is not a date (use YYYY-MM-DD).")


def parse_time(value: str) -> dt.time:
    m = TIME_RE.match(str(value or "").strip())
    if not m:
        raise RosterError(400, f"'{value}' is not a time (use HH:MM, 24-hour).")
    return dt.time(int(m.group(1)), int(m.group(2)))


def daterange(start: dt.date, end: dt.date):
    if end < start:
        raise RosterError(400, "The end date is before the start date.")
    if (end - start).days + 1 > MAX_RANGE_DAYS:
        raise RosterError(400, f"A range can be at most {MAX_RANGE_DAYS} days.")
    day = start
    while day <= end:
        yield day
        day += dt.timedelta(days=1)


def norm(value) -> str:
    return re.sub(r"[^a-z0-9]", "", str(value or "").lower())


def rank(actor) -> int:
    return ROLE_RANK.get((actor or {}).get("role"), 0)


def actor_name(actor) -> str:
    return str((actor or {}).get("user") or (actor or {}).get("name") or "unknown")


# --------------------------------------------------------------------------- defaults


TEMPLATES = {
    "standard": "Morning 06:30-15:30, Afternoon 13:30-21:30, Night 21:00-07:00, Day 09:00-18:00 (weekdays)",
    "three_shift": "Morning 06:00-15:00, Afternoon 13:00-22:00, Night 21:30-06:30, with Week off and Leave",
}


def default_roster(team: str, name: str, start_date: str, template: str = "standard") -> dict:
    """A starting point; every part can be changed afterwards.

    standard:    the four-shift pattern (Morning, Afternoon, Night, Day).
    three_shift: the 6 AM / 1 PM / 9:30 PM pattern, laid out by shift (day rows,
                 shift columns, Week off, Leave), as used by teams that plan by day.
    """
    if template not in TEMPLATES:
        raise RosterError(400, f"Template must be one of: {', '.join(TEMPLATES)}.")
    r = _standard_roster(team, name, start_date)
    if template == "three_shift":
        r["template"] = "three_shift"
        r["layout"] = "by_shift"
        r["shifts"] = [
            {"code": "M", "name": "Morning", "start": "06:00", "end": "15:00", "days": [0, 1, 2, 3, 4, 5, 6],
             "min_staff": 2, "max_staff": None, "color": "#F2C76E", "log_shift": "Morning"},
            {"code": "A", "name": "Afternoon", "start": "13:00", "end": "22:00", "days": [0, 1, 2, 3, 4, 5, 6],
             "min_staff": 2, "max_staff": None, "color": "#7AA7FF", "log_shift": "Afternoon"},
            {"code": "N", "name": "Night", "start": "21:30", "end": "06:30", "days": [0, 1, 2, 3, 4, 5, 6],
             "min_staff": 2, "max_staff": None, "color": "#A391F5", "log_shift": "Night"},
        ]
        r["rules"]["forbidden_sequences"] = [["N", "M"]]
    return r


def _standard_roster(team: str, name: str, start_date: str) -> dict:
    return {
        "template": "standard",
        "layout": "by_person",
        "id": new_id("ros_"),
        "name": name,
        "team": team,
        "start_date": str(parse_date(start_date)),
        "cycle": {"type": "custom", "length_days": 21},
        "shifts": [
            {"code": "M", "name": "Morning", "start": "06:30", "end": "15:30", "days": [0, 1, 2, 3, 4, 5, 6],
             "min_staff": 1, "max_staff": None, "color": "#F2C76E", "log_shift": "Morning"},
            {"code": "A", "name": "Afternoon", "start": "13:30", "end": "21:30", "days": [0, 1, 2, 3, 4, 5, 6],
             "min_staff": 1, "max_staff": None, "color": "#7AA7FF", "log_shift": "Afternoon"},
            {"code": "N", "name": "Night", "start": "21:00", "end": "07:00", "days": [0, 1, 2, 3, 4, 5, 6],
             "min_staff": 1, "max_staff": None, "color": "#A391F5", "log_shift": "Night"},
            {"code": "D", "name": "Day", "start": "09:00", "end": "18:00", "days": [0, 1, 2, 3, 4],
             "min_staff": 0, "max_staff": None, "color": "#43C08F", "log_shift": "Morning"},
        ],
        "rotation": {"sequence": ["M"] * 5 + [OFF] * 2 + ["A"] * 5 + [OFF] * 2 + ["N"] * 5 + [OFF] * 2},
        "rules": {"min_rest_hours": 8, "forbidden_sequences": [["N", "M"], ["N", "D"]],
                  "max_consecutive_days": 6, "skip_inactive_days": True, "allow_shared_members": False},
        "members": [],
        "approvals": {"leave": "lead", "swap": "lead"},
        "approvers": [],
        "leave_types": ["Annual", "Sick", "Training", "Comp off", "Other"],
        "holidays": {"region": "", "subdivision": "", "custom": [], "ignore": []},
        "oncall": {"enabled": False, "roles": ["Primary", "Secondary"], "rotation_days": 7,
                   "start_date": str(parse_date(start_date)), "order": {}, "escalation": [], "overrides": {}},
        "overrides": {},
        "leaves": [],
        "swaps": [],
        "created_at": now_iso(),
        "updated_at": now_iso(),
    }


# --------------------------------------------------------------------------- validation


def cycle_length(r: dict):
    c = r.get("cycle") or {}
    kind = c.get("type")
    if kind in CYCLE_DAYS:
        return CYCLE_DAYS[kind]
    if kind == "custom":
        return int(c.get("length_days") or 0)
    return None  # monthly: restarts on the 1st, any sequence length


def shift_map(r: dict) -> dict:
    return {s["code"]: s for s in r.get("shifts", [])}


def validate_roster(r: dict) -> None:
    """Raise RosterError listing every problem with the configuration."""
    errs = []
    if not str(r.get("name") or "").strip():
        errs.append("The roster needs a name.")
    if not str(r.get("team") or "").strip():
        errs.append("The roster needs a team.")
    try:
        parse_date(r.get("start_date"))
    except RosterError as e:
        errs.append("Start date: " + e.message)
    cyc = r.get("cycle") or {}
    if cyc.get("type") not in CYCLE_TYPES:
        errs.append(f"Cycle type must be one of: {', '.join(CYCLE_TYPES)}.")
    elif cyc["type"] == "custom":
        try:
            n = int(cyc.get("length_days"))
            if not 1 <= n <= 366:
                raise ValueError
        except (TypeError, ValueError):
            errs.append("A custom cycle needs a length between 1 and 366 days.")
    codes = set()
    for i, s in enumerate(r.get("shifts") or []):
        label = f"Shift {i + 1} ({s.get('name') or s.get('code') or '?'})"
        code = str(s.get("code") or "")
        if not CODE_RE.match(code):
            errs.append(f"{label}: the code must be 1 to 8 capital letters or digits.")
        elif code in RESERVED_CODES:
            errs.append(f"{label}: '{code}' is reserved.")
        elif code in codes:
            errs.append(f"{label}: the code '{code}' is used twice.")
        codes.add(code)
        if not str(s.get("name") or "").strip():
            errs.append(f"{label}: needs a name.")
        for key in ("start", "end"):
            try:
                parse_time(s.get(key))
            except RosterError as e:
                errs.append(f"{label} {key}: {e.message}")
        days = s.get("days")
        if not isinstance(days, list) or not days or any(not isinstance(d, int) or not 0 <= d <= 6 for d in days):
            errs.append(f"{label}: working days must be a list of weekday numbers 0 (Mon) to 6 (Sun).")
        try:
            mn = int(s.get("min_staff", 0))
            if mn < 0:
                raise ValueError
            mx = s.get("max_staff")
            if mx is not None and int(mx) < mn:
                errs.append(f"{label}: maximum staff is below the minimum.")
        except (TypeError, ValueError):
            errs.append(f"{label}: staff numbers must be whole numbers, 0 or more.")
        if s.get("color") and not COLOR_RE.match(str(s["color"])):
            errs.append(f"{label}: colour must look like #RRGGBB.")
    if not codes:
        errs.append("Define at least one shift pattern.")
    seq = (r.get("rotation") or {}).get("sequence") or []
    if not seq:
        errs.append("The rotation needs at least one entry.")
    bad = sorted({c for c in seq if c != OFF and c not in codes})
    if bad:
        errs.append(f"The rotation uses unknown shift codes: {', '.join(bad)}.")
    n = cycle_length(r) if cyc.get("type") in CYCLE_TYPES else None
    if n and seq and len(seq) != n:
        errs.append(f"The rotation has {len(seq)} entries but the {cyc.get('type')} cycle is {n} days long.")
    rules = r.get("rules") or {}
    try:
        h = float(rules.get("min_rest_hours", 0))
        if not 0 <= h <= 48:
            raise ValueError
    except (TypeError, ValueError):
        errs.append("Minimum rest must be between 0 and 48 hours.")
    mc = rules.get("max_consecutive_days")
    if mc is not None:
        try:
            if not 1 <= int(mc) <= 31:
                raise ValueError
        except (TypeError, ValueError):
            errs.append("Maximum consecutive days must be between 1 and 31, or empty.")
    for pair in rules.get("forbidden_sequences") or []:
        if not (isinstance(pair, list) and len(pair) == 2 and all(p in codes for p in pair)):
            errs.append(f"Forbidden sequence {pair} must be two known shift codes.")
    ids = set()
    for m in r.get("members") or []:
        if not m.get("id") or m["id"] in ids:
            errs.append("Every member needs a unique id.")
        ids.add(m.get("id"))
        if not str(m.get("name") or "").strip():
            errs.append("Every member needs a name.")
        try:
            if int(m.get("offset", 0)) < 0:
                raise ValueError
        except (TypeError, ValueError):
            errs.append(f"Member {m.get('name')}: the offset must be 0 or more.")
    hol = r.get("holidays") or {}
    if hol.get("region") and hol["region"] not in HOLIDAY_REGIONS:
        errs.append(f"Holiday region must be one of: {', '.join(HOLIDAY_REGIONS)} (or empty).")
    if hol.get("subdivision") and not hol.get("region"):
        errs.append("Choose a holiday region before a state or subdivision.")
    for x in hol.get("custom") or []:
        try:
            parse_date((x or {}).get("date"))
            if not str((x or {}).get("name") or "").strip():
                raise RosterError(400, "")
        except RosterError:
            errs.append("Each custom holiday needs a date (YYYY-MM-DD) and a name.")
            break
    for x in hol.get("ignore") or []:
        try:
            parse_date(x)
        except RosterError:
            errs.append(f"Ignored holiday '{x}' is not a date.")
    for s in r.get("shifts") or []:
        hm = s.get("holiday_min_staff")
        if hm is not None:
            try:
                if int(hm) < 0:
                    raise ValueError
            except (TypeError, ValueError):
                errs.append(f"Shift {s.get('name')}: holiday minimum staff must be 0 or more.")
    if r.get("layout", "by_person") not in ("by_person", "by_shift"):
        errs.append("Layout must be by_person or by_shift.")
    apv = r.get("approvers", [])
    if not isinstance(apv, list) or any(not isinstance(x, str) or not x.strip() for x in apv):
        errs.append("Approvers must be a list of user names.")
    ap = r.get("approvals") or {}
    for key in ("leave", "swap"):
        if ap.get(key) not in APPROVAL_MODES:
            errs.append(f"Approval for {key} must be 'lead' or 'self'.")
    oc = r.get("oncall") or {}
    if oc.get("enabled"):
        roles = oc.get("roles") or []
        if not roles or any(not str(x).strip() for x in roles):
            errs.append("On-call needs at least one role name.")
        try:
            if int(oc.get("rotation_days", 0)) < 1:
                raise ValueError
        except (TypeError, ValueError):
            errs.append("On-call rotation must be 1 day or more.")
        for role, order in (oc.get("order") or {}).items():
            unknown = [x for x in order if x not in ids]
            if unknown:
                errs.append(f"On-call order for {role} names unknown members.")
    if errs:
        raise RosterError(400, " ".join(errs))


# --------------------------------------------------------------------------- holidays


HOLIDAY_REGIONS = {"US": "United States", "IN": "India"}


def holidays_available() -> bool:
    try:
        import holidays  # noqa: F401
        return True
    except ImportError:
        return False


def holiday_subdivisions(region: str) -> list:
    """State / subdivision codes the region supports (for example KA, MH, TN for India; CA, NY for the US)."""
    if region not in HOLIDAY_REGIONS or not holidays_available():
        return []
    import holidays
    return sorted(holidays.country_holidays(region).subdivisions)


def holidays_for(r: dict, start, end) -> dict:
    """{date: name} of public holidays in the range: region list, plus custom days, minus ignored days.

    Region dates come from the open-source 'holidays' package (pip install holidays). Company
    holiday lists often differ (especially lunar festival dates in India), so use 'custom' to add
    days and 'ignore' to remove them.
    """
    start, end = parse_date(start), parse_date(end)
    cfg = r.get("holidays") or {}
    out = {}
    region = cfg.get("region")
    if region and holidays_available():
        import holidays
        try:
            cal = holidays.country_holidays(region, subdiv=cfg.get("subdivision") or None,
                                            years=range(start.year, end.year + 1))
        except NotImplementedError:
            raise RosterError(400, f"'{cfg.get('subdivision')}' is not a subdivision of {region}.")
        for d, name in cal.items():
            if start <= d <= end:
                out[str(d)] = name
    for x in cfg.get("custom") or []:
        d = parse_date(x["date"])
        if start <= d <= end:
            out[str(d)] = str(x["name"])
    for x in cfg.get("ignore") or []:
        out.pop(str(parse_date(x)), None)
    return dict(sorted(out.items()))


# --------------------------------------------------------------------------- schedule


def _member(r: dict, member_id: str) -> dict:
    for m in r.get("members", []):
        if m["id"] == member_id:
            return m
    raise RosterError(404, "That member is not in this roster.")


def rotation_code(r: dict, day: dt.date, offset: int):
    seq = r["rotation"]["sequence"]
    start = parse_date(r["start_date"])
    if day < start:
        return None
    if r["cycle"]["type"] == "monthly":
        return seq[((day.day - 1) + offset) % len(seq)]
    return seq[((day - start).days + offset) % len(seq)]


def _approved(items):
    return [x for x in items if x.get("status") == "approved"]


def generate(r: dict, start, end) -> dict:
    """{date: {member_id: {"code": ..., "source": ...}}} for the range."""
    start, end = parse_date(start), parse_date(end)
    shifts = shift_map(r)
    skip = (r.get("rules") or {}).get("skip_inactive_days", True)
    imported = bool(r.get("imported"))
    ovs = r.get("overrides") or {}
    sched = {}
    for day in daterange(start, end):
        row = {}
        if imported and not ovs.get(str(day)):
            sched[str(day)] = row  # an imported roster has no data for a day its sheets never listed
            continue
        for m in r.get("members", []):
            code = OFF if imported else rotation_code(r, day, int(m.get("offset", 0)))
            if code is None:
                continue
            source = "rotation"
            if code in shifts and skip and day.weekday() not in shifts[code]["days"]:
                code, source = OFF, "inactive day"
            row[m["id"]] = {"code": code, "source": source}
        for mid, code in (r.get("overrides") or {}).get(str(day), {}).items():
            if mid in row:
                row[mid] = {"code": code, "source": "override"}
        sched[str(day)] = row
    for sw in sorted(_approved(r.get("swaps", [])), key=lambda x: x.get("decided_at") or x.get("requested_at") or ""):
        _apply_swap(sched, sw)
    for lv in _approved(r.get("leaves", [])):
        for day in daterange(max(parse_date(lv["start"]), start), min(parse_date(lv["end"]), end)) \
                if parse_date(lv["end"]) >= start and parse_date(lv["start"]) <= end else []:
            row = sched.get(str(day), {})
            if lv["member_id"] in row:
                row[lv["member_id"]] = {"code": LEAVE, "source": "leave: " + lv.get("type", "Leave")}
    return sched


def _apply_swap(sched: dict, sw: dict) -> None:
    a, b = sw["member_a"], sw["member_b"]
    if sw.get("kind") == "cover":
        row = sched.get(sw["date_a"])
        if row and a in row and b in row:
            row[b] = {"code": row[a]["code"], "source": "covering " + a}
            row[a] = {"code": OFF, "source": "covered by " + b}
        return
    for d in {sw["date_a"], sw["date_b"]}:
        row = sched.get(d)
        if row and a in row and b in row:
            ca, cb = row[a]["code"], row[b]["code"]
            row[a] = {"code": cb, "source": "swap"}
            row[b] = {"code": ca, "source": "swap"}


def required_staff(shift: dict, day, hols: dict) -> int:
    """Minimum people for a shift on a day: the holiday minimum on a public holiday, if one is set."""
    if str(day) in hols and shift.get("holiday_min_staff") is not None:
        return int(shift["holiday_min_staff"])
    return int(shift.get("min_staff", 0))


def _shift_window(shift: dict, day: dt.date):
    s, e = parse_time(shift["start"]), parse_time(shift["end"])
    begin = dt.datetime.combine(day, s)
    finish = dt.datetime.combine(day, e)
    if finish <= begin:
        finish += dt.timedelta(days=1)
    return begin, finish


def check(r: dict, sched: dict, other_rosters=None) -> list:
    """Coverage and rule problems as a list of dicts (level, date, member, rule, message)."""
    issues = []
    shifts = shift_map(r)
    rules = r.get("rules") or {}
    names = {m["id"]: m["name"] for m in r.get("members", [])}
    days = sorted(sched)
    hols = holidays_for(r, days[0], days[-1]) if days else {}
    for d in days:
        if not sched[d]:
            continue  # no roster data that day (before the rota starts, or outside an imported sheet)
        wd = parse_date(d).weekday()
        for code, s in shifts.items():
            if wd not in s["days"]:
                continue
            n = sum(1 for v in sched[d].values() if v["code"] == code)
            need = required_staff(s, d, hols)
            if n < need:
                issues.append({"level": "error", "date": d, "member": None, "rule": "coverage",
                               "message": f"{s['name']} has {n} of {need} people" + (f" ({hols[d]})." if d in hols else ".")})
            if s.get("max_staff") is not None and n > int(s["max_staff"]):
                issues.append({"level": "warning", "date": d, "member": None, "rule": "coverage",
                               "message": f"{s['name']} has {n} people, above the maximum of {s['max_staff']}."})
    min_rest = float(rules.get("min_rest_hours", 0))
    forbidden = {tuple(p) for p in rules.get("forbidden_sequences") or []}
    max_run = rules.get("max_consecutive_days")
    for mid, name in names.items():
        prev = None  # (day, code, finish)
        run = 0
        prev_day = None
        for d in days:
            cell = sched[d].get(mid)
            if not cell:
                prev, run, prev_day = None, 0, None
                continue
            code = cell["code"]
            day = parse_date(d)
            if code in shifts:
                begin, finish = _shift_window(shifts[code], day)
                run = run + 1 if prev_day == day - dt.timedelta(days=1) else 1
                prev_day = day
                if max_run and run > int(max_run):
                    issues.append({"level": "error", "date": d, "member": mid, "rule": "consecutive days",
                                   "message": f"{name} works {run} days in a row (limit {max_run})."})
                if prev:
                    gap = (begin - prev[2]).total_seconds() / 3600
                    if gap < min_rest:
                        issues.append({"level": "error", "date": d, "member": mid, "rule": "rest",
                                       "message": f"{name} has {gap:.1f} h rest before {shifts[code]['name']} (minimum {min_rest:g} h)."})
                    if prev[0] == day - dt.timedelta(days=1) and (prev[1], code) in forbidden:
                        issues.append({"level": "error", "date": d, "member": mid, "rule": "sequence",
                                       "message": f"{name}: {shifts[prev[1]]['name']} followed by {shifts[code]['name']} is not allowed."})
                prev = (day, code, finish)
            else:
                run, prev_day = 0, None
    if other_rosters is not None and not rules.get("allow_shared_members"):
        mine = {norm(m.get("user") or m["name"]): m["id"] for m in r.get("members", [])}
        for o in other_rosters:
            if o.get("id") == r.get("id"):
                continue
            for m in o.get("members", []):
                key = norm(m.get("user") or m["name"])
                if key in mine:
                    issues.append({"level": "warning", "date": None, "member": mine[key], "rule": "shared member",
                                   "message": f"{names[mine[key]]} is also in the roster '{o.get('name')}'."})
    return issues


# --------------------------------------------------------------------------- views


def oncall_for(r: dict, day) -> dict:
    """{role: {"member_id", "name", "substitute": bool}} for one day."""
    day = parse_date(day)
    oc = r.get("oncall") or {}
    if not oc.get("enabled"):
        return {}
    names = {m["id"]: m["name"] for m in r.get("members", [])}
    on_leave = {lv["member_id"] for lv in _approved(r.get("leaves", []))
                if parse_date(lv["start"]) <= day <= parse_date(lv["end"])}
    start = parse_date(oc.get("start_date") or r["start_date"])
    k = (day - start).days // int(oc.get("rotation_days") or 7)
    taken, out = set(), {}
    forced = (oc.get("overrides") or {}).get(str(day), {})
    for role in oc.get("roles", []):
        if role in forced and forced[role] in names:
            mid = forced[role]
            out[role] = {"member_id": mid, "name": names[mid], "substitute": False, "override": True}
            taken.add(mid)
            continue
        order = [x for x in (oc.get("order") or {}).get(role, []) if x in names]
        if not order:
            continue
        first = order[k % len(order)]
        pick = None
        for i in range(len(order)):
            cand = order[(k + i) % len(order)]
            if cand not in on_leave and cand not in taken:
                pick = cand
                break
        if pick:
            out[role] = {"member_id": pick, "name": names[pick], "substitute": pick != first, "override": False}
            taken.add(pick)
    return out


def on_now(r: dict, when=None) -> dict:
    """The shifts running at a moment (server time), who is on them, and who is on call.

    A night shift that started yesterday and ends this morning is found from yesterday's schedule.
    """
    when = when or dt.datetime.now()
    names = {m["id"]: m["name"] for m in r.get("members", [])}
    running = []
    for back in (1, 0):
        day = when.date() - dt.timedelta(days=back)
        row = generate(r, day, day)[str(day)]
        if not row:
            continue
        for s in r.get("shifts", []):
            if day.weekday() not in s.get("days", range(7)):
                continue  # the shift does not run on that weekday (for example a weekday-only Day shift)
            begin, finish = _shift_window(s, day)
            if begin <= when < finish:
                running.append({"code": s["code"], "name": s["name"], "start": s["start"], "end": s["end"], "date": str(day),
                                "color": s.get("color", ""), "minutes_left": int((finish - when).total_seconds() // 60),
                                "people": [names[mid] for mid, cell in row.items() if cell["code"] == s["code"]]})
    return {"at": when.isoformat(timespec="minutes"), "shifts": running,
            "oncall": {k: v["name"] for k, v in oncall_for(r, when.date()).items()}}


def calendar(r: dict, start, mode: str = "week") -> dict:
    """Day, week (Mon-Sun) or month view of who works which shift."""
    start = parse_date(start)
    if mode == "day":
        a, b = start, start
    elif mode == "week":
        a = start - dt.timedelta(days=start.weekday())
        b = a + dt.timedelta(days=6)
    elif mode == "month":
        a = start.replace(day=1)
        b = (a.replace(day=28) + dt.timedelta(days=4)).replace(day=1) - dt.timedelta(days=1)
    else:
        raise RosterError(400, "View must be day, week or month.")
    sched = generate(r, a, b)
    shifts = shift_map(r)
    names = {m["id"]: m["name"] for m in r.get("members", [])}
    hols = holidays_for(r, a, b)
    days = []
    for d in sorted(sched):
        entry = {"date": d, "weekday": WEEKDAYS[parse_date(d).weekday()], "holiday": hols.get(d, ""),
                 "has_data": bool(sched[d]), "shifts": {c: [] for c in shifts}, "off": [], "leave": [], "oncall": {}}
        for mid, cell in sched[d].items():
            c = cell["code"]
            if c in shifts:
                entry["shifts"][c].append(names[mid])
            elif c == LEAVE:
                entry["leave"].append(names[mid])
            elif not (r.get("imported") and cell["source"] == "rotation"):
                entry["off"].append(names[mid])
        entry["oncall"] = {k: v["name"] for k, v in oncall_for(r, d).items()}
        days.append(entry)
    return {"roster": r["name"], "mode": mode, "start": str(a), "end": str(b), "days": days}


def person_view(r: dict, member_id: str, start, end) -> list:
    m = _member(r, member_id)
    shifts = shift_map(r)
    sched = generate(r, start, end)
    hols = holidays_for(r, start, end)
    out = []
    for d in sorted(sched):
        cell = sched[d].get(m["id"])
        if not cell:
            continue
        s = shifts.get(cell["code"])
        roles = [role for role, v in oncall_for(r, d).items() if v["member_id"] == m["id"]]
        out.append({"date": d, "weekday": WEEKDAYS[parse_date(d).weekday()], "code": cell["code"],
                    "shift": s["name"] if s else cell["code"].title(), "start": s["start"] if s else "",
                    "end": s["end"] if s else "", "source": cell["source"], "oncall": roles, "holiday": hols.get(d, ""),
                    "shift_log": shiftlog_link(d, s) if s else ""})
    return out


def shiftlog_link(day, shift: dict) -> str:
    """Address of that day's shift in the main app's shift log."""
    target = shift.get("log_shift") or shift.get("name") or ""
    return f"/#shiftlog={parse_date(day)}/{target}"


# --------------------------------------------------------------------------- workflow (pure)


def _require(actor, need: int, what: str):
    if rank(actor) < need:
        raise RosterError(403, f"Your role cannot {what}.")


def _is_self(r: dict, actor, member_id: str) -> bool:
    m = _member(r, member_id)
    return bool(m.get("user")) and norm(m.get("user")) == norm((actor or {}).get("user"))


OPEN_STATES = ("awaiting_partner", "pending", "approved")


def _pick_approver(r: dict, approver):
    """The approver a request is assigned to: one of the roster's approvers, or None for any Lead."""
    names = [x for x in (r.get("approvers") or []) if str(x).strip()]
    if not names:
        return None
    if not approver:
        return names[0]
    for x in names:
        if norm(x) == norm(approver):
            return x
    raise RosterError(400, f"'{approver}' is not an approver for this roster. Choose one of: {', '.join(names)}.")


def request_leave(r: dict, actor, member_id, start, end, kind="Annual", note="", approver=None) -> dict:
    _require(actor, 2, "request leave")
    if rank(actor) < 3 and not _is_self(r, actor, member_id):
        raise RosterError(403, "Engineers can only request leave for themselves.")
    a, b = parse_date(start), parse_date(end)
    list(daterange(a, b))
    if kind not in r.get("leave_types", []):
        raise RosterError(400, f"Leave type must be one of: {', '.join(r.get('leave_types', []))}.")
    for lv in r.get("leaves", []):
        if lv["member_id"] == member_id and lv["status"] in ("pending", "approved") \
                and parse_date(lv["start"]) <= b and a <= parse_date(lv["end"]):
            raise RosterError(409, "This overlaps leave that is already requested or approved.")
    auto = r["approvals"]["leave"] == "self" or rank(actor) >= 3
    lv = {"id": new_id("lv_"), "member_id": member_id, "start": str(a), "end": str(b), "type": kind,
          "note": str(note)[:500], "status": "approved" if auto else "pending",
          "approver": None if auto else _pick_approver(r, approver),
          "requested_by": actor_name(actor), "requested_at": now_iso(),
          "decided_by": actor_name(actor) if auto else None, "decided_at": now_iso() if auto else None}
    r.setdefault("leaves", []).append(lv)
    return lv


def request_swap(r: dict, actor, member_a, date_a, member_b, date_b=None, kind="swap", note="", approver=None) -> dict:
    """Swap or cover. If the other person has a user name, they must accept first; then the approver decides
    (unless swaps are self-recorded or a Lead made the request)."""
    _require(actor, 2, "request a swap")
    if kind not in ("swap", "cover"):
        raise RosterError(400, "Kind must be swap or cover.")
    if member_a == member_b:
        raise RosterError(400, "Choose two different people.")
    partner = _member(r, member_b)
    if rank(actor) < 3 and not _is_self(r, actor, member_a):
        raise RosterError(403, "Engineers can only swap their own shifts.")
    da = parse_date(date_a)
    db = parse_date(date_b) if (date_b and kind == "swap") else da
    for d in {da, db}:
        row = generate(r, d, d)[str(d)]
        if member_a not in row or member_b not in row:
            raise RosterError(400, f"There is no roster for {d} yet, so there is no shift to swap. "
                                   "Import or fill in that period first.")
    sw = {"id": new_id("sw_"), "kind": kind, "member_a": member_a, "date_a": str(da), "member_b": member_b,
          "date_b": str(db), "note": str(note)[:500], "requested_by": actor_name(actor), "requested_at": now_iso()}
    lo, hi = min(da, db) - dt.timedelta(days=2), max(da, db) + dt.timedelta(days=2)
    trial = copy.deepcopy(r)
    trial.setdefault("swaps", []).append(dict(sw, status="approved", decided_at="~"))
    after = {(i["date"], i["member"], i["rule"]) for i in check(trial, generate(trial, lo, hi)) if i["member"] in (member_a, member_b)}
    before = {(i["date"], i["member"], i["rule"]) for i in check(r, generate(r, lo, hi)) if i["member"] in (member_a, member_b)}
    new_issues = [i["message"] for i in check(trial, generate(trial, lo, hi))
                  if (i["date"], i["member"], i["rule"]) in after - before]
    needs_lead = not (r["approvals"]["swap"] == "self" or rank(actor) >= 3)
    partner_user = str(partner.get("user") or "").strip()
    needs_consent = bool(partner_user) and norm(partner_user) != norm((actor or {}).get("user"))
    status = "awaiting_partner" if needs_consent else ("pending" if needs_lead else "approved")
    sw.update(status=status, warnings=new_issues, needs_lead=needs_lead, partner=partner["name"], partner_user=partner_user,
              approver=_pick_approver(r, approver) if needs_lead else None,
              consent_by=None if needs_consent else (actor_name(actor) if partner_user else None),
              decided_by=actor_name(actor) if status == "approved" else None,
              decided_at=now_iso() if status == "approved" else None)
    r.setdefault("swaps", []).append(sw)
    return sw


def consent(r: dict, actor, req_id: str, accept: bool, note="") -> dict:
    """The other person in a swap or cover agrees (or declines) before it goes to the approver."""
    for it in r.get("swaps", []):
        if it["id"] == req_id:
            if it["status"] != "awaiting_partner":
                raise RosterError(409, f"This request is already {it['status'].replace('_', ' ')}.")
            if norm(it.get("partner_user")) != norm((actor or {}).get("user")):
                raise RosterError(403, f"Only {it.get('partner') or 'the other person'} can accept or decline this.")
            it.update(consent_by=actor_name(actor), consent_at=now_iso(), consent_note=str(note)[:500])
            if not accept:
                it["status"] = "declined"
            elif it.get("needs_lead"):
                it["status"] = "pending"
            else:
                it.update(status="approved", decided_by=actor_name(actor) + " (agreed)", decided_at=now_iso())
            return it
    raise RosterError(404, "Request not found.")


def decide(r: dict, actor, kind: str, req_id: str, approve: bool, note="") -> dict:
    _require(actor, 3, "approve or reject requests")
    items = r.get("leaves" if kind == "leave" else "swaps", [])
    for it in items:
        if it["id"] == req_id:
            if it["status"] == "awaiting_partner":
                raise RosterError(409, f"Waiting for {it.get('partner') or 'the other person'} to accept first.")
            if it["status"] != "pending":
                raise RosterError(409, f"This request is already {it['status']}.")
            if it.get("approver") and rank(actor) < 4 and norm(it["approver"]) != norm((actor or {}).get("user")):
                raise RosterError(403, f"This request is assigned to {it['approver']}.")
            it.update(status="approved" if approve else "rejected", decided_by=actor_name(actor),
                      decided_at=now_iso(), decision_note=str(note)[:500])
            return it
    raise RosterError(404, "Request not found.")


def cancel(r: dict, actor, kind: str, req_id: str) -> dict:
    items = r.get("leaves" if kind == "leave" else "swaps", [])
    for it in items:
        if it["id"] == req_id:
            if it["status"] in ("cancelled", "rejected", "declined"):
                raise RosterError(409, f"This request is already {it['status']}.")
            if rank(actor) < 3 and norm(it.get("requested_by")) != norm(actor_name(actor)):
                raise RosterError(403, "Only the requester or a Lead can cancel this.")
            it.update(status="cancelled", cancelled_by=actor_name(actor), cancelled_at=now_iso())
            return it
    raise RosterError(404, "Request not found.")


def waiting_for(r: dict, actor) -> list:
    """Requests in this roster that wait for this person: their consent, or their approval."""
    me = norm((actor or {}).get("user"))
    out = []
    for it in r.get("leaves", []) + r.get("swaps", []):
        kind = "leave" if it["id"].startswith("lv_") else "swap"
        if it["status"] == "awaiting_partner" and norm(it.get("partner_user")) == me:
            out.append(dict(it, kind=it.get("kind", kind), action="consent"))
        elif it["status"] == "pending" and rank(actor) >= 3 and (
                (it.get("approver") and norm(it["approver"]) == me) or (not it.get("approver"))):
            out.append(dict(it, kind=it.get("kind", kind), action="approve"))
    return out


def set_override(r: dict, actor, member_id, day, code) -> None:
    _require(actor, 3, "change the roster")
    _member(r, member_id)
    if code not in shift_map(r) and code != OFF and code is not None:
        raise RosterError(400, f"Unknown shift code '{code}'.")
    d = str(parse_date(day))
    ov = r.setdefault("overrides", {})
    if code is None:
        ov.get(d, {}).pop(member_id, None)
        if d in ov and not ov[d]:
            del ov[d]
    else:
        ov.setdefault(d, {})[member_id] = code


def add_member(r: dict, actor, name, role="Engineer", user="", offset=0, other_rosters=()) -> dict:
    _require(actor, 3, "change the roster")
    if not str(name).strip():
        raise RosterError(400, "The member needs a name.")
    key = norm(user or name)
    if any(norm(m.get("user") or m["name"]) == key for m in r.get("members", [])):
        raise RosterError(409, f"{name} is already in this roster.")
    if not (r.get("rules") or {}).get("allow_shared_members"):
        for o in other_rosters:
            if o.get("id") != r.get("id") and any(norm(m.get("user") or m["name"]) == key for m in o.get("members", [])):
                raise RosterError(409, f"{name} already belongs to the roster '{o.get('name')}'. "
                                       "Allow shared members in this roster's rules to add them here too.")
    m = {"id": new_id("mem_"), "name": str(name).strip(), "role": role, "user": str(user or "").strip(), "offset": int(offset)}
    r.setdefault("members", []).append(m)
    return m


# --------------------------------------------------------------------------- storage


class FileStorage:
    """Files in one folder, atomic writes, version check. Same interface as msp_runbook storage."""

    def __init__(self, data_dir: str):
        self.data_dir = os.path.abspath(data_dir)
        os.makedirs(self.data_dir, exist_ok=True)
        self._lock = threading.Lock()

    def _path(self, name: str) -> str:
        if "/" in name or "\\" in name or name in ("", ".", "..") or ":" in name:
            raise RosterError(400, "Invalid file name.")
        return os.path.join(self.data_dir, name)

    def version(self, name: str):
        try:
            st = os.stat(self._path(name))
        except FileNotFoundError:
            return None
        return f"{st.st_mtime_ns}:{st.st_size}"

    def read(self, name: str):
        try:
            with open(self._path(name), "rb") as fh:
                return fh.read(), self.version(name)
        except FileNotFoundError:
            return None, None

    def write(self, name: str, data: bytes, expected):
        path = self._path(name)
        with self._lock:
            if self.version(name) != expected:
                raise ConflictError(name)
            tmp = f"{path}.{uuid.uuid4().hex[:8]}.tmp"
            with open(tmp, "wb") as fh:
                fh.write(data)
                fh.flush()
                os.fsync(fh.fileno())
            os.replace(tmp, path)
            return self.version(name)


class RosterStore:
    """Loads and saves all rosters; every change keeps the previous roster as a version."""

    def __init__(self, storage, name: str = ROSTER_FILE):
        self.storage = storage
        self.name = name

    def load(self):
        data, ver = self.storage.read(self.name)
        if not data:
            return {"format": 1, "rosters": {}, "versions": {}, "audit": []}, ver
        state = json.loads(data.decode("utf-8"))
        state.setdefault("rosters", {})
        state.setdefault("versions", {})
        state.setdefault("audit", [])
        return state, ver

    def rosters(self) -> list:
        return list(self.load()[0]["rosters"].values())

    def get(self, roster_id: str) -> dict:
        r = self.load()[0]["rosters"].get(roster_id)
        if not r:
            raise RosterError(404, "Roster not found.")
        return r

    def _mutate(self, actor, action: str, roster_id, fn):
        for _ in range(5):
            state, ver = self.load()
            before = copy.deepcopy(state["rosters"].get(roster_id)) if roster_id else None
            result, rid = fn(state)
            if rid in state["rosters"]:
                r = state["rosters"][rid]
                validate_roster(r)
                r["updated_at"] = now_iso()
            if before is not None:
                vs = state["versions"].setdefault(rid, [])
                vs.append({"at": now_iso(), "by": actor_name(actor), "action": action, "roster": before})
                del vs[:-MAX_VERSIONS]
            state["audit"].append({"at": now_iso(), "by": actor_name(actor), "action": action, "roster": rid})
            del state["audit"][:-MAX_AUDIT]
            try:
                self.storage.write(self.name, json.dumps(state, indent=1).encode("utf-8"), ver)
                return result
            except Exception as exc:  # any storage's ConflictError: reload and re-apply
                if type(exc).__name__ != "ConflictError":
                    raise
        raise RosterError(409, "The roster kept changing while saving; please try again.")

    # -- configuration
    def create(self, actor, roster: dict) -> dict:
        _require(actor, 3, "create a roster")
        roster = copy.deepcopy(roster)
        roster.setdefault("id", new_id("ros_"))
        validate_roster(roster)

        def fn(state):
            if roster["id"] in state["rosters"]:
                raise RosterError(409, "A roster with this id already exists.")
            state["rosters"][roster["id"]] = roster
            return roster, roster["id"]
        return self._mutate(actor, "Create roster", None, fn)

    def update(self, actor, roster_id: str, changes: dict) -> dict:
        _require(actor, 3, "change the roster")
        locked = {"id", "leaves", "swaps", "created_at"}

        def fn(state):
            r = state["rosters"].get(roster_id)
            if not r:
                raise RosterError(404, "Roster not found.")
            for k, v in changes.items():
                if k not in locked:
                    r[k] = copy.deepcopy(v)
            return r, roster_id
        return self._mutate(actor, "Update roster", roster_id, fn)

    def delete(self, actor, roster_id: str) -> None:
        _require(actor, 3, "delete a roster")

        def fn(state):
            if not state["rosters"].pop(roster_id, None):
                raise RosterError(404, "Roster not found.")
            return None, roster_id
        return self._mutate(actor, "Delete roster", roster_id, fn)

    def _on(self, actor, roster_id, action, op):
        def fn(state):
            r = state["rosters"].get(roster_id)
            if not r:
                raise RosterError(404, "Roster not found.")
            others = [o for k, o in state["rosters"].items() if k != roster_id]
            return op(r, others), roster_id
        return self._mutate(actor, action, roster_id, fn)

    def add_member(self, actor, roster_id, name, role="Engineer", user="", offset=0):
        return self._on(actor, roster_id, "Add member", lambda r, o: add_member(r, actor, name, role, user, offset, o))

    def request_leave(self, actor, roster_id, member_id, start, end, kind="Annual", note="", approver=None):
        return self._on(actor, roster_id, "Request leave",
                        lambda r, o: request_leave(r, actor, member_id, start, end, kind, note, approver))

    def request_swap(self, actor, roster_id, member_a, date_a, member_b, date_b=None, kind="swap", note="", approver=None):
        return self._on(actor, roster_id, "Request " + kind,
                        lambda r, o: request_swap(r, actor, member_a, date_a, member_b, date_b, kind, note, approver))

    def consent(self, actor, roster_id, req_id, accept, note=""):
        return self._on(actor, roster_id, ("Accept " if accept else "Decline ") + "swap",
                        lambda r, o: consent(r, actor, req_id, accept, note))

    def decide(self, actor, roster_id, kind, req_id, approve, note=""):
        return self._on(actor, roster_id, ("Approve " if approve else "Reject ") + kind,
                        lambda r, o: decide(r, actor, kind, req_id, approve, note))

    def cancel(self, actor, roster_id, kind, req_id):
        return self._on(actor, roster_id, "Cancel " + kind, lambda r, o: cancel(r, actor, kind, req_id))

    def set_override(self, actor, roster_id, member_id, day, code):
        return self._on(actor, roster_id, "Change shift", lambda r, o: set_override(r, actor, member_id, day, code))

    def import_new(self, actor, source, team: str, name: str) -> dict:
        """Create a roster from a by-shift sheet. Returns {"roster", "warnings"}."""
        parsed = parse_roster_file(source)
        roster = roster_from_by_shift(parsed, team, name, actor)
        created = self.create(actor, roster)
        return {"roster": created, "warnings": import_notes(created) + parsed["warnings"], "layout": parsed.get("layout"),
                "sheet": parsed.get("sheet")}

    def import_into(self, actor, roster_id, source) -> dict:
        """Apply a roster sheet (either layout) to an existing roster; shifts and people match by name."""
        parsed = parse_roster_file(source)
        res = self._on(actor, roster_id, "Import roster sheet", lambda r, o: apply_by_shift(r, actor, parsed))
        res["warnings"] = parsed["warnings"]
        return res

    def set_overrides(self, actor, roster_id, changes: list) -> int:
        """Several one-day changes saved as a single version: [{member_id, date, code or None}]."""
        _require(actor, 3, "change the roster")
        if not isinstance(changes, list) or not changes:
            raise RosterError(400, "Nothing to change.")
        if len(changes) > 500:
            raise RosterError(400, "Too many changes at once (500 at most).")

        def op(r, o):
            for c in changes:
                set_override(r, actor, (c or {}).get("member_id"), (c or {}).get("date"), (c or {}).get("code"))
            return len(changes)
        return self._on(actor, roster_id, f"Change {len(changes)} shift(s)", op)

    def clone(self, actor, roster_id, name: str, team: str, copy_members: bool = False) -> dict:
        """A new roster with the same shifts, rotation, rules, approvals and holidays.

        Members (and the on-call order) are copied only when asked, and only if this roster
        allows shared members, because people are dedicated to one team by default.
        Leave, swaps, one-day changes and versions are never copied.
        """
        _require(actor, 3, "clone a roster")
        src = self.get(roster_id)
        r = copy.deepcopy(src)
        r.update(id=new_id("ros_"), name=re.sub(r"\s+", " ", str(name or "")).strip() or f"{src['name']} (copy)",
                 team=re.sub(r"\s+", " ", str(team or "")).strip() or src["team"], leaves=[], swaps=[], overrides={},
                 created_at=now_iso(), updated_at=now_iso(), cloned_from=src["id"])
        r["oncall"] = dict(r.get("oncall") or {}, overrides={})
        if copy_members:
            if not (r.get("rules") or {}).get("allow_shared_members"):
                raise RosterError(409, "Members are dedicated to one roster here, so they cannot be copied. "
                                       "Clone without members, or allow shared members in this roster's rules first.")
            id_map = {m["id"]: new_id("mem_") for m in r.get("members", [])}
            for m in r["members"]:
                m["id"] = id_map[m["id"]]
            r["oncall"]["order"] = {k: [id_map[x] for x in v if x in id_map] for k, v in (r["oncall"].get("order") or {}).items()}
        else:
            r["members"] = []
            r["oncall"]["order"] = {}
        return self.create(actor, r)

    def deleted(self) -> list:
        """Rosters that were deleted but still have saved versions (newest first)."""
        state = self.load()[0]
        out = []
        for rid, vs in state["versions"].items():
            if rid not in state["rosters"] and vs and vs[-1].get("action") == "Delete roster" and vs[-1].get("roster"):
                r = vs[-1]["roster"]
                out.append({"id": rid, "name": r.get("name"), "team": r.get("team"), "members": len(r.get("members", [])),
                            "deleted_at": vs[-1]["at"], "deleted_by": vs[-1]["by"]})
        return sorted(out, key=lambda x: x["deleted_at"], reverse=True)

    def undelete(self, actor, roster_id: str) -> dict:
        """Bring back a deleted roster exactly as it was before it was deleted."""
        _require(actor, 3, "restore a deleted roster")

        def fn(state):
            if roster_id in state["rosters"]:
                raise RosterError(409, "That roster is not deleted.")
            vs = state["versions"].get(roster_id) or []
            if not vs or vs[-1].get("action") != "Delete roster":
                raise RosterError(404, "No deleted roster with that id.")
            r = copy.deepcopy(vs[-1]["roster"])
            if not (r.get("rules") or {}).get("allow_shared_members"):
                mine = {norm(m.get("user") or m["name"]): m["name"] for m in r.get("members", [])}
                for o in state["rosters"].values():
                    for m in o.get("members", []):
                        if norm(m.get("user") or m["name"]) in mine:
                            raise RosterError(409, f"{m['name']} now belongs to the roster '{o['name']}'. "
                                                   "Remove them there first, or allow shared members.")
            state["rosters"][roster_id] = r
            return r, roster_id
        return self._mutate(actor, "Restore deleted roster", None, fn)

    def history(self, roster_id) -> list:
        vs = self.load()[0]["versions"].get(roster_id, [])
        return [{"index": i, "at": v["at"], "by": v["by"], "action": v["action"]} for i, v in enumerate(vs)]

    def restore(self, actor, roster_id, index: int) -> dict:
        _require(actor, 3, "restore a version")

        def fn(state):
            vs = state["versions"].get(roster_id, [])
            if not 0 <= index < len(vs):
                raise RosterError(404, "Version not found.")
            state["rosters"][roster_id] = copy.deepcopy(vs[index]["roster"])
            return state["rosters"][roster_id], roster_id
        return self._mutate(actor, f"Restore version {index}", roster_id, fn)


# --------------------------------------------------------------------------- exports


def _col(n: int) -> str:
    s = ""
    while n:
        n, rem = divmod(n - 1, 26)
        s = chr(65 + rem) + s
    return s


def export_excel(r: dict, start, end) -> bytes:
    """Workbook: Calendar, Coverage (COUNTIF formulas), On-call, Requests, Shifts, Rules, Legend."""
    from openpyxl import Workbook
    from openpyxl.comments import Comment
    from openpyxl.styles import Alignment, Font, PatternFill

    start, end = parse_date(start), parse_date(end)
    sched = generate(r, start, end)
    days = sorted(sched)
    shifts = shift_map(r)
    members = r.get("members", [])
    hols = holidays_for(r, start, end)
    F = "Arial"
    bold = Font(name=F, bold=True)
    norm_font = Font(name=F)
    blue = Font(name=F, color="0000FF")
    center = Alignment(horizontal="center", vertical="center")
    hol_fill = PatternFill("solid", fgColor="FFD9A8")
    wb = Workbook()

    cal = wb.active
    cal.title = "Calendar"
    cal["A1"] = f"{r['name']} ({r['team']}): {start} to {end}"
    cal["A1"].font = Font(name=F, bold=True, size=13)
    cal["A2"], cal["B2"] = "Member", "Role"
    for j, d in enumerate(days):
        c = cal.cell(row=2, column=3 + j, value=parse_date(d))
        c.number_format = "ddd dd-mmm"
        c.font, c.alignment = bold, center
        if d in hols:
            c.fill = hol_fill
            c.comment = Comment("Holiday: " + hols[d], "MSP roster")
        cal.column_dimensions[_col(3 + j)].width = 10
    cal["A2"].font = cal["B2"].font = bold
    for i, m in enumerate(members):
        row = 3 + i
        cal.cell(row=row, column=1, value=m["name"]).font = norm_font
        cal.cell(row=row, column=2, value=m.get("role", "")).font = norm_font
        for j, d in enumerate(days):
            cell = sched[d].get(m["id"])
            code = cell["code"] if cell else ""
            c = cal.cell(row=row, column=3 + j, value=code)
            c.font, c.alignment = norm_font, center
            if code in shifts and shifts[code].get("color"):
                c.fill = PatternFill("solid", fgColor=shifts[code]["color"].lstrip("#"))
            elif code == LEAVE:
                c.fill = PatternFill("solid", fgColor="F0B4B0")
            if cell and cell["source"] not in ("rotation", "inactive day"):
                c.comment = Comment(cell["source"], "MSP roster")
    cal.column_dimensions["A"].width = 22
    cal.column_dimensions["B"].width = 12
    cal.freeze_panes = "C3"
    last = 2 + max(len(members), 1)

    bs = wb.create_sheet("By shift")
    shift_cols = list(shifts.values())
    bs.append(["Day", "Date"] + [f"{s['name'].upper()} SHIFT" for s in shift_cols] + ["WEEK OFF", "LEAVE", "HOLIDAY"])
    bs.append(["", ""] + [f"{s['start']} to {s['end']}" for s in shift_cols] + ["", "", ""])
    for c in bs[1]:
        c.font = bold
    for c in bs[2]:
        c.font = Font(name=F, italic=True)
    for row in by_shift_rows(r, start, end):
        bs.append([row["day"], parse_date(row["date"])] + [row["shifts"][s["code"]] for s in shift_cols]
                  + [row["off"], row["leave"], row["holiday"]])
        bs.cell(row=bs.max_row, column=2).number_format = "d-mmm-yy"
        if row["holiday"]:
            for k in (1, 2, len(shift_cols) + 5):
                bs.cell(row=bs.max_row, column=k).fill = hol_fill
    bs.column_dimensions["A"].width = 11
    bs.column_dimensions["B"].width = 11
    for j in range(len(shift_cols) + 2):
        bs.column_dimensions[_col(3 + j)].width = 34
    bs.column_dimensions[_col(len(shift_cols) + 5)].width = 26
    for row_cells in bs.iter_rows(min_row=3):
        for c in row_cells:
            c.alignment = Alignment(wrap_text=True, vertical="top")
    bs.freeze_panes = "C3"
    if r.get("layout") == "by_shift":
        wb.move_sheet(bs, offset=-1)
        wb.active = 0

    cov = wb.create_sheet("Coverage")
    cov["A1"] = "Coverage: Staffed is counted from the Calendar sheet; Required comes from each shift's minimum staff."
    cov["A1"].font = Font(name=F, italic=True)
    cov["A2"] = "Shift"
    cov["B2"] = "Line"
    cov["A2"].font = cov["B2"].font = bold
    for j, d in enumerate(days):
        c = cov.cell(row=2, column=3 + j, value=parse_date(d))
        c.number_format = "ddd dd-mmm"
        c.font, c.alignment = bold, center
    row = 3
    for code, s in shifts.items():
        for k, label in enumerate(("Staffed", "Required", "Status")):
            cov.cell(row=row + k, column=1, value=f"{s['name']} ({code})").font = norm_font
            cov.cell(row=row + k, column=2, value=label).font = norm_font
        for j, d in enumerate(days):
            col = _col(3 + j)
            cov[f"{col}{row}"] = f'=COUNTIF(Calendar!{col}$3:{col}${last},"{code}")'
            req = required_staff(s, d, hols) if (parse_date(d).weekday() in s["days"] and sched[d]) else 0
            cov[f"{col}{row + 1}"] = req
            cov[f"{col}{row + 1}"].font = blue
            cov[f"{col}{row + 2}"] = f'=IF({col}{row + 1}=0,"-",IF({col}{row}<{col}{row + 1},"SHORT","OK"))'
            for k in range(3):
                cov[f"{col}{row + k}"].alignment = center
                if k != 1:
                    cov[f"{col}{row + k}"].font = norm_font
        row += 4
    cov.column_dimensions["A"].width = 22
    cov.freeze_panes = "C3"

    oc = wb.create_sheet("On-call")
    roles = (r.get("oncall") or {}).get("roles", []) if (r.get("oncall") or {}).get("enabled") else []
    oc.append(["Date", "Day"] + roles + ["Notes"])
    for c in oc[1]:
        c.font = bold
    for d in days:
        who = oncall_for(r, d)
        notes = "; ".join(f"{k}: substitute (rota person on leave)" for k, v in who.items() if v["substitute"])
        oc.append([parse_date(d), WEEKDAYS[parse_date(d).weekday()]] + [who.get(k, {}).get("name", "") for k in roles] + [notes])
    for row_cells in oc.iter_rows(min_row=2):
        row_cells[0].number_format = "yyyy-mm-dd"
        for c in row_cells:
            c.font = norm_font
    esc = (r.get("oncall") or {}).get("escalation") or []
    if esc:
        oc.append([])
        oc.append(["Escalation", "Level", "Name", "Contact"])
        for e in esc:
            oc.append(["", e.get("level", ""), e.get("name", ""), e.get("contact", "")])
    oc.column_dimensions["A"].width = 12

    rq = wb.create_sheet("Requests")
    names = {m["id"]: m["name"] for m in members}
    rq.append(["Kind", "Person", "From", "To / other date", "Other person", "Type", "Status", "Requested by", "Decided by", "Note"])
    for c in rq[1]:
        c.font = bold
    for lv in r.get("leaves", []):
        rq.append(["Leave", names.get(lv["member_id"], "?"), lv["start"], lv["end"], "", lv.get("type", ""),
                   lv["status"], lv.get("requested_by", ""), lv.get("decided_by") or "", lv.get("note", "")])
    for sw in r.get("swaps", []):
        rq.append([sw["kind"].title(), names.get(sw["member_a"], "?"), sw["date_a"], sw["date_b"], names.get(sw["member_b"], "?"),
                   "", sw["status"], sw.get("requested_by", ""), sw.get("decided_by") or "", sw.get("note", "")])

    sh = wb.create_sheet("Shifts")
    sh.append(["Code", "Name", "Start", "End", "Working days", "Min staff", "Max staff", "Shift log"])
    for c in sh[1]:
        c.font = bold
    for s in shifts.values():
        sh.append([s["code"], s["name"], s["start"], s["end"], ", ".join(WEEKDAYS[d] for d in s["days"]),
                   s.get("min_staff", 0), s.get("max_staff") if s.get("max_staff") is not None else "", s.get("log_shift", "")])

    ru = wb.create_sheet("Rules")
    rules = r.get("rules") or {}
    cyc = r["cycle"]
    for k, v in [("Cycle", cyc["type"] + (f" ({cyc.get('length_days')} days)" if cyc["type"] == "custom" else "")),
                 ("Rotation", " ".join(r["rotation"]["sequence"])),
                 ("Minimum rest (hours)", rules.get("min_rest_hours")),
                 ("Not allowed back to back", ", ".join(f"{a} then {b}" for a, b in rules.get("forbidden_sequences") or [])),
                 ("Max consecutive working days", rules.get("max_consecutive_days") or "no limit"),
                 ("Shared members allowed", "Yes" if rules.get("allow_shared_members") else "No"),
                 ("Leave approval", "Lead approves" if r["approvals"]["leave"] == "lead" else "Self-recorded"),
                 ("Swap approval", "Lead approves" if r["approvals"]["swap"] == "lead" else "Self-recorded")]:
        ru.append([k, v])
    for c in ru["A"]:
        c.font = bold

    hs = wb.create_sheet("Holidays")
    cfg = r.get("holidays") or {}
    region = cfg.get("region") or ""
    hs.append(["Date", "Day", "Holiday", "Source"])
    for c in hs[1]:
        c.font = bold
    custom = {str(parse_date(x["date"])) for x in cfg.get("custom") or []}
    for d, name in hols.items():
        hs.append([parse_date(d), WEEKDAYS[parse_date(d).weekday()], name,
                   "Company list" if d in custom else f"{HOLIDAY_REGIONS.get(region, region)}"
                   + (f" ({cfg.get('subdivision')})" if cfg.get("subdivision") else "")])
        hs.cell(row=hs.max_row, column=1).number_format = "yyyy-mm-dd"
    if not hols:
        hs.append(["", "", "No holidays in this period" if (region or custom) else "No holiday region or company list set", ""])
    hs.column_dimensions["A"].width = 12
    hs.column_dimensions["C"].width = 34
    hs.column_dimensions["D"].width = 22

    lg = wb.create_sheet("Legend")
    lg.append(["Code", "Meaning"])
    for c in lg[1]:
        c.font = bold
    for s in shifts.values():
        lg.append([s["code"], f"{s['name']} {s['start']}-{s['end']}"])
        if s.get("color"):
            lg.cell(row=lg.max_row, column=1).fill = PatternFill("solid", fgColor=s["color"].lstrip("#"))
    lg.append([OFF, "Day off"])
    lg.append([LEAVE, "Approved leave"])
    lg.append(["", "Public holiday (orange date heading; name in the cell note)"])
    lg.cell(row=lg.max_row, column=1).fill = hol_fill
    lg.cell(row=lg.max_row, column=1).fill = PatternFill("solid", fgColor="F0B4B0")
    lg.append([])
    lg.append(["Blue numbers on the Coverage sheet are inputs (each shift's minimum staff). A cell comment on the Calendar marks a swap, cover, override or leave."])

    for ws in wb.worksheets:
        for row_cells in ws.iter_rows():
            for c in row_cells:
                if c.font is None or c.font.name != F:
                    c.font = Font(name=F, bold=c.font.bold if c.font else False, color=c.font.color if c.font else None)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


REPORT_CSS = """
:root{--ink:#14213d;--muted:#5b6b85;--line:#dde3ee;--bg:#f5f7fb;--card:#fff;--acc:#0e9f8e;--acc2:#2f6fdb;--bad:#d63c3c;--warn:#c27a00}
*{box-sizing:border-box}body{margin:0;font:14px/1.45 Arial,Helvetica,sans-serif;color:var(--ink);background:var(--bg)}
header.top{background:linear-gradient(120deg,#0b1a33,#11425a 55%,#0e9f8e);color:#fff;padding:22px 28px 18px}
header.top h1{margin:0 0 4px;font-size:24px}header.top .meta{opacity:.85;font-size:13px}
.ver{display:inline-block;margin-left:10px;padding:2px 10px;border-radius:12px;background:rgba(255,255,255,.18);font-size:12px}
.cards{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:12px;padding:16px 28px 4px}
.card{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:12px 14px}
.card b{display:block;font-size:24px}.card span{color:var(--muted);font-size:12px}.card.bad b{color:var(--bad)}.card.ok b{color:var(--acc)}
.toolbar{position:sticky;top:0;z-index:5;display:flex;flex-wrap:wrap;gap:10px;align-items:center;padding:12px 28px;background:var(--bg);border-bottom:1px solid var(--line)}
.tabs{display:flex;gap:4px;flex-wrap:wrap}.tab{border:1px solid var(--line);background:#fff;border-radius:18px;padding:6px 14px;cursor:pointer;font:inherit;color:var(--ink)}
.tab.on{background:var(--ink);color:#fff;border-color:var(--ink)}
#q{flex:1;min-width:200px;padding:8px 12px;border:1px solid var(--line);border-radius:18px;font:inherit}
.chip{display:inline-flex;align-items:center;gap:6px;padding:4px 10px;border-radius:14px;border:1px solid rgba(0,0,0,.12);cursor:pointer;font-size:12px;font-weight:bold;user-select:none}
.chip.off{opacity:.3;text-decoration:line-through}#count{color:var(--muted);font-size:12px}
main{padding:12px 28px 28px}.pane{display:none}.pane.on{display:block}
.wrap{overflow-x:auto;background:#fff;border:1px solid var(--line);border-radius:12px}
table{border-collapse:collapse;width:100%}th,td{border-bottom:1px solid var(--line);padding:7px 9px;text-align:left;vertical-align:top}
thead th{position:sticky;top:0;background:#eef2f8;font-size:12px;text-transform:uppercase;letter-spacing:.03em;white-space:nowrap}
th[data-sort]{cursor:pointer}th[data-sort]:after{content:" \\2195";opacity:.35}th[data-dir=asc]:after{content:" \\2191";opacity:1}th[data-dir=desc]:after{content:" \\2193";opacity:1}
tr.we td,tr.we th{background:#fbf6ea}tr.hol td,tr.hol th{background:#ffe7c7}.hb{display:block;margin-top:2px;font-size:11px;font-weight:bold;color:#a35400}tbody tr:hover td{background:#eef7f6}
td.c{text-align:center;font-weight:bold;min-width:42px}td.dim,.dim{opacity:.12}
.n{font-size:11px;color:var(--muted);margin-left:4px}.short{box-shadow:inset 0 0 0 2px var(--bad)}.short .n{color:var(--bad);font-weight:bold}
.pill{display:inline-block;padding:1px 8px;border-radius:10px;font-size:12px;font-weight:bold}
.lv-error{color:var(--bad)}.lv-warning{color:var(--warn)}.st-approved{background:#dff5ee;color:#08765f}.st-pending{background:#fff3d6;color:#8a5a00}
.st-rejected,.st-cancelled{background:#f3e1e1;color:#a33}.legend{display:flex;flex-wrap:wrap;gap:8px;margin:0 0 10px}
mark{background:#ffe27a;padding:0}footer{color:var(--muted);font-size:12px;padding:0 28px 24px}
@media print{.toolbar{display:none}.pane{display:block!important;break-before:page}.pane h2{display:block}body{background:#fff}.wrap{border:0}thead th{position:static}}
.pane h2{font-size:16px;margin:4px 0 10px}
"""

REPORT_JS = """
(function(){
var $=function(s){return document.querySelector(s)},$$=function(s){return Array.prototype.slice.call(document.querySelectorAll(s))};
var q='',hide={};
function pane(){return $('.pane.on')}
function apply(){var p=pane();if(!p)return;var rows=$$('.pane.on tbody tr'),shown=0;
 rows.forEach(function(tr){var hit=!q||tr.textContent.toLowerCase().indexOf(q)>=0;tr.style.display=hit?'':'none';if(hit)shown++;});
 $$('[data-c]').forEach(function(el){el.classList.toggle('dim',!!hide[el.getAttribute('data-c')]);});
 $$('.pane.on mark').forEach(function(m){m.parentNode.replaceChild(document.createTextNode(m.textContent),m);});
 $$('.pane.on tbody td, .pane.on tbody th').forEach(function(td){td.normalize();});
 if(q){$$('.pane.on tbody td, .pane.on tbody th').forEach(function(td){var w=document.createTreeWalker(td,NodeFilter.SHOW_TEXT,null),n,list=[];
  while((n=w.nextNode()))if(!(n.parentNode.classList&&n.parentNode.classList.contains('n')))list.push(n);
  list.forEach(function(node){var t=node.nodeValue,i=t.toLowerCase().indexOf(q);if(i<0)return;var f=document.createDocumentFragment();
   f.appendChild(document.createTextNode(t.slice(0,i)));var m=document.createElement('mark');m.textContent=t.slice(i,i+q.length);f.appendChild(m);
   f.appendChild(document.createTextNode(t.slice(i+q.length)));node.parentNode.replaceChild(f,node);});});}
 $('#count').textContent=shown+' of '+rows.length+' rows';}
$$('.tab').forEach(function(b){b.addEventListener('click',function(){$$('.tab').forEach(function(x){x.classList.toggle('on',x===b)});
 $$('.pane').forEach(function(p){p.classList.toggle('on',p.id===b.getAttribute('data-p'))});apply();});});
$('#q').addEventListener('input',function(e){q=e.target.value.trim().toLowerCase();apply();});
$$('.chip[data-code]').forEach(function(c){c.addEventListener('click',function(){var k=c.getAttribute('data-code');hide[k]=!hide[k];c.classList.toggle('off',hide[k]);apply();});});
$$('th[data-sort]').forEach(function(th){th.addEventListener('click',function(){var tb=th.closest('table').tBodies[0],i=Array.prototype.indexOf.call(th.parentNode.children,th);
 var dir=th.getAttribute('data-dir')==='asc'?'desc':'asc';$$('th[data-dir]').forEach(function(x){x.removeAttribute('data-dir')});th.setAttribute('data-dir',dir);
 var rows=Array.prototype.slice.call(tb.rows);rows.sort(function(a,b){var x=a.cells[i].getAttribute('data-v')||a.cells[i].textContent,y=b.cells[i].getAttribute('data-v')||b.cells[i].textContent;
 var nx=parseFloat(x),ny=parseFloat(y),r=(!isNaN(nx)&&!isNaN(ny)&&/^-?\\d+(\\.\\d+)?$/.test(x)&&/^-?\\d+(\\.\\d+)?$/.test(y))?nx-ny:x.localeCompare(y);return dir==='asc'?r:-r;});
 rows.forEach(function(r){tb.appendChild(r)});});});
apply();})();
"""


def export_html(r: dict, start, end, layout: str = None, history=None) -> str:
    """A self-contained report page: tabs per view, search, sorting, shift filters, checks and versions.

    Every value is escaped and rendered on the server, so the page reads and prints without scripts;
    the small inline script only adds tabs, search, sorting and filters.
    """
    e = html.escape
    start, end = parse_date(start), parse_date(end)
    layout = layout or r.get("layout") or "by_person"
    if layout not in ("by_person", "by_shift"):
        raise RosterError(400, "Layout must be by_person or by_shift.")
    sched = generate(r, start, end)
    days = sorted(sched)
    shifts = shift_map(r)
    cols = list(shifts.values())
    members = r.get("members", [])
    issues = check(r, sched)
    history = history or []
    hols = holidays_for(r, start, end)
    color = lambda c: (shifts[c].get("color") or "#dddddd") if c in shifts else ("#F0B4B0" if c == LEAVE else "#eef1f5")
    weekend = lambda d: parse_date(d).weekday() >= 5

    # By shift: one row per day, people per shift, with staffed count and shortfalls
    bs = []
    for x in by_shift_rows(r, start, end):
        wd = parse_date(x["date"]).weekday()
        tds = []
        for s in cols:
            names = x["shifts"][s["code"]]
            n = len([p for p in names.split(", ") if p])
            active = wd in s["days"]
            need = required_staff(s, x["date"], hols)
            short = active and n < need
            tds.append(f'<td data-c="{e(s["code"])}" class="{"short" if short else ""}" style="background:{e(color(s["code"]))}40" '
                       f'data-v="{n:03d}">{e(names)}<span class=n>({n}{"/" + str(need) if active and need else ""})</span></td>')
        hb = f'<span class=hb>{e(x["holiday"])}</span>' if x["holiday"] else ""
        bs.append(f'<tr class="{"hol" if x["holiday"] else ("we" if wd >= 5 else "")}"><th data-v="{e(x["date"])}">{e(x["day"])}{hb}</th><td data-v="{e(x["date"])}">{e(x["date"])}</td>'
                  + "".join(tds) + f'<td data-c="{OFF}">{e(x["off"])}</td><td data-c="{LEAVE}">{e(x["leave"])}</td></tr>')
    by_shift_table = ('<table><thead><tr><th data-sort>Day</th><th data-sort>Date</th>'
                      + "".join(f'<th data-sort>{e(s["name"])} <span class=n>{e(s["start"])}-{e(s["end"])}</span></th>' for s in cols)
                      + "<th data-sort>Week off</th><th data-sort>Leave</th></tr></thead><tbody>" + "".join(bs) + "</tbody></table>")

    # By person: one row per person, a coloured cell per day, plus totals
    bp = []
    for m in members:
        cells, tally = [], {}
        for d in days:
            cell = sched[d].get(m["id"])
            code = cell["code"] if cell else ""
            tally[code] = tally.get(code, 0) + 1
            tip = cell["source"] if cell else ""
            cells.append(f'<td class=c data-c="{e(code)}" title="{e(tip)}" style="background:{e(color(code))}">{e(code)}</td>')
        work = sum(v for k, v in tally.items() if k in shifts)
        on_hol = sum(1 for d in days if d in hols and (sched[d].get(m["id"]) or {}).get("code") in shifts)
        bp.append(f'<tr><th>{e(m["name"])}</th><td>{e(m.get("role", ""))}</td>' + "".join(cells)
                  + f'<td data-v="{work:03d}"><b>{work}</b></td><td data-v="{on_hol:03d}">{on_hol}</td><td data-v="{tally.get(LEAVE, 0):03d}">{tally.get(LEAVE, 0)}</td></tr>')
    by_person_table = ('<table><thead><tr><th data-sort>Member</th><th data-sort>Role</th>'
                       + "".join(f'<th data-sort class=c{" style=background:#ffd9a8" if d in hols else (" style=background:#f6ecd2" if weekend(d) else "")}'
                               f'{" title=" + chr(34) + e(hols[d]) + chr(34) if d in hols else ""}>{e(WEEKDAYS[parse_date(d).weekday()])}<br>{e(d[8:])}/{e(d[5:7])}</th>' for d in days)
                       + "<th data-sort>Shifts</th><th data-sort>On holidays</th><th data-sort>Leave</th></tr></thead><tbody>" + "".join(bp) + "</tbody></table>")

    chk = "".join(f'<tr><td data-v="{e(i["level"])}"><span class="lv-{e(i["level"])}">{e(i["level"].title())}</span></td><td>{e(i["date"] or "")}</td>'
                  f'<td>{e(i["rule"])}</td><td>{e(i["message"])}</td></tr>' for i in issues)
    checks_table = ('<table><thead><tr><th data-sort>Level</th><th data-sort>Date</th><th data-sort>Rule</th><th data-sort>Finding</th></tr></thead><tbody>'
                    + (chk or '<tr><td colspan=4>No problems found: every shift is staffed and every rule is met.</td></tr>') + "</tbody></table>")

    oc_table = "<p>On-call is not switched on for this roster.</p>"
    if (r.get("oncall") or {}).get("enabled"):
        roles = r["oncall"]["roles"]
        rows = []
        for d in days:
            who = oncall_for(r, d)
            rows.append(f'<tr class="{"we" if weekend(d) else ""}"><td data-v="{e(d)}">{e(d)}</td><td>{e(WEEKDAYS[parse_date(d).weekday()])}</td>'
                        + "".join(f'<td>{e(who.get(x, {}).get("name", ""))}{" <span class=n>(substitute)</span>" if who.get(x, {}).get("substitute") else ""}</td>' for x in roles) + "</tr>")
        esc = "".join(f'<li>Level {e(str(x.get("level", "")))}: {e(x.get("name", ""))} {e(x.get("contact", ""))}</li>' for x in r["oncall"].get("escalation") or [])
        oc_table = ('<table><thead><tr><th data-sort>Date</th><th data-sort>Day</th>' + "".join(f"<th data-sort>{e(x)}</th>" for x in roles)
                    + "</tr></thead><tbody>" + "".join(rows) + "</tbody></table>" + (f"<h2>Escalation</h2><ul>{esc}</ul>" if esc else ""))

    names = {m["id"]: m["name"] for m in members}
    rq = [f'<tr><td>Leave</td><td>{e(names.get(x["member_id"], "?"))}</td><td data-v="{e(x["start"])}">{e(x["start"])} to {e(x["end"])}</td><td>{e(x.get("type", ""))}</td>'
          f'<td><span class="pill st-{e(x["status"])}">{e(x["status"])}</span></td><td>{e(x.get("requested_by", ""))}</td><td>{e(x.get("note", ""))}</td></tr>'
          for x in r.get("leaves", [])]
    rq += [f'<tr><td>{e(x["kind"].title())}</td><td>{e(names.get(x["member_a"], "?"))} and {e(names.get(x["member_b"], "?"))}</td>'
           f'<td data-v="{e(x["date_a"])}">{e(x["date_a"])}{"" if x["date_b"] == x["date_a"] else " / " + e(x["date_b"])}</td><td>{e("; ".join(x.get("warnings") or []))}</td>'
           f'<td><span class="pill st-{e(x["status"])}">{e(x["status"])}</span></td><td>{e(x.get("requested_by", ""))}</td><td>{e(x.get("note", ""))}</td></tr>'
           for x in r.get("swaps", [])]
    req_table = ('<table><thead><tr><th data-sort>Kind</th><th data-sort>People</th><th data-sort>Dates</th><th data-sort>Type / checks</th>'
                 '<th data-sort>Status</th><th data-sort>Requested by</th><th>Note</th></tr></thead><tbody>'
                 + ("".join(rq) or '<tr><td colspan=7>No leave or swap requests.</td></tr>') + "</tbody></table>")

    hist_rows = "".join(f'<tr><td data-v="{i:04d}">{i + 1}</td><td>{e(h.get("at", ""))}</td><td>{e(h.get("by", ""))}</td><td>{e(h.get("action", ""))}</td></tr>'
                        for i, h in enumerate(history))
    hist_table = ('<table><thead><tr><th data-sort>Version</th><th data-sort>Saved at</th><th data-sort>By</th><th data-sort>Change</th></tr></thead><tbody>'
                  + (hist_rows or '<tr><td colspan=4>No earlier versions yet.</td></tr>') + "</tbody></table>"
                  + '<p class=n>Each row is a saved earlier version. A Lead can restore any of them; the restore is itself kept as a new version.</p>')

    cfg = r.get("holidays") or {}
    custom = {str(parse_date(x["date"])) for x in cfg.get("custom") or []}
    src = HOLIDAY_REGIONS.get(cfg.get("region") or "", "") + (f" ({cfg.get('subdivision')})" if cfg.get("subdivision") else "")
    hol_rows = "".join(f'<tr><td data-v="{e(d)}">{e(d)}</td><td>{e(WEEKDAYS[parse_date(d).weekday()])}</td><td><b>{e(n)}</b></td>'
                       f'<td>{e("Company list" if d in custom else src)}</td><td>'
                       + e(", ".join(f"{sh['name']}: {len([p for p in row['shifts'][sh['code']].split(', ') if p])}" for sh in cols for row in [next(y for y in by_shift_rows(r, d, d))]))
                       + "</td></tr>" for d, n in hols.items())
    note = "" if holidays_available() or not cfg.get("region") else '<p class=n>The holidays package is not installed, so only the company list is shown (pip install holidays).</p>'
    hol_table = ('<table><thead><tr><th data-sort>Date</th><th data-sort>Day</th><th data-sort>Holiday</th><th data-sort>Source</th><th>Staffed</th></tr></thead><tbody>'
                 + (hol_rows or '<tr><td colspan=5>No public holidays in this period' + ("" if cfg.get("region") or custom else ": no region or company list is set") + ".</td></tr>")
                 + "</tbody></table>" + note)
    errors = sum(1 for i in issues if i["level"] == "error")
    leave_days = sum(1 for d in days for v in sched[d].values() if v["code"] == LEAVE)
    version_no = len(history) + 1
    last = history[-1] if history else None
    cards = (f'<div class=card><b>{len(members)}</b><span>people</span></div><div class=card><b>{len(days)}</b><span>days</span></div>'
             f'<div class=card><b>{len(cols)}</b><span>shift patterns</span></div><div class=card><b>{leave_days}</b><span>leave days</span></div>'
             f'<div class=card><b>{len(hols)}</b><span>public holidays</span></div>'
             f'<div class="card {"bad" if errors else "ok"}"><b>{errors}</b><span>problems to fix</span></div>'
             f'<div class=card><b>v{version_no}</b><span>version</span></div>')
    chips = "".join(f'<span class=chip data-code="{e(s["code"])}" style="background:{e(color(s["code"]))}">{e(s["code"])} {e(s["name"])}</span>' for s in cols)
    chips += f'<span class=chip data-code="{OFF}" style="background:{color(OFF)}">Off</span><span class=chip data-code="{LEAVE}" style="background:{color(LEAVE)}">Leave</span>'
    first = "by_shift" if layout == "by_shift" else "by_person"
    panes = [("by_shift", "By shift", by_shift_table), ("by_person", "By person", by_person_table), ("checks", f"Checks ({len(issues)})", checks_table),
             ("oncall", "On-call", oc_table), ("holidays", f"Holidays ({len(hols)})", hol_table), ("requests", f"Requests ({len(rq)})", req_table),
             ("history", f"Versions ({len(history)})", hist_table)]
    panes.sort(key=lambda p: p[0] != first)
    tabs = "".join(f'<button class="tab{" on" if i == 0 else ""}" data-p="{p[0]}">{e(p[1])}</button>' for i, p in enumerate(panes))
    body = "".join(f'<section class="pane{" on" if i == 0 else ""}" id="{p[0]}"><h2>{e(p[1])}</h2><div class=wrap>{p[2]}</div></section>' for i, p in enumerate(panes))
    meta = (f'{e(r["team"])} team &middot; {e(str(start))} to {e(str(end))} &middot; {e(r["cycle"]["type"])} cycle'
            + (f' &middot; holidays: {e(src)}' if src else ""))
    ver = f'<span class=ver>Version {version_no}' + (f' &middot; last change {e(last["at"])} by {e(last["by"])}' if last else "") + "</span>"
    return (f'<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">'
            f'<title>{e(r["name"])} roster</title><style>{REPORT_CSS}</style></head><body>'
            f'<header class=top><h1>{e(r["name"])}</h1><div class=meta>{meta}{ver}</div></header><div class=cards>{cards}</div>'
            f'<div class=toolbar><div class=tabs>{tabs}</div><input id=q type=search placeholder="Search people, dates, shifts, notes" aria-label="Search">'
            f'<div class=legend>{chips}</div><span id=count></span></div><main>{body}</main>'
            f'<footer>Generated {e(now_iso())} by MSP Ops Center roster {MODULE_VERSION}. Click a column heading to sort; click a shift to hide or show it.</footer>'
            f'<script>{REPORT_JS}</script></body></html>')


# --------------------------------------------------------------------------- by-shift layout (day rows, shift columns)


def by_shift_rows(r: dict, start, end) -> list:
    """Rows of the by-shift layout: Day, Date, one column per shift, Week off, Leave."""
    sched = generate(r, start, end)
    shifts = shift_map(r)
    names = {m["id"]: m["name"] for m in r.get("members", [])}
    hols = holidays_for(r, start, end)
    rows = []
    for d in sorted(sched):
        by = {c: [] for c in shifts}
        off, leave = [], []
        for m in r.get("members", []):
            cell = sched[d].get(m["id"])
            if not cell:
                continue
            if cell["code"] in shifts:
                by[cell["code"]].append(names[m["id"]])
            elif cell["code"] == LEAVE:
                leave.append(names[m["id"]])
            elif not (r.get("imported") and cell["source"] == "rotation"):
                off.append(names[m["id"]])  # on an imported sheet, unlisted people are not "week off"
        rows.append({"day": WEEKDAYS_LONG[parse_date(d).weekday()], "date": d, "holiday": hols.get(d, ""), "has_data": bool(sched[d]),
                     "shifts": {c: ", ".join(v) for c, v in by.items()},
                     "off": ", ".join(off), "leave": ", ".join(leave)})
    return rows


WEEKDAYS_LONG = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
_TIME_WORD = re.compile(r"^\s*(\d{1,2})(?::(\d{2}))?\s*([AaPp][Mm])?\s*$")
_DATE_FORMATS = ("%d-%b-%y", "%d-%b-%Y", "%Y-%m-%d", "%d/%m/%Y", "%d/%m/%y", "%d %b %Y", "%d %B %Y")


def _time_from_words(text: str):
    m = _TIME_WORD.match(text)
    if not m:
        return None
    h, mi, ap = int(m.group(1)), int(m.group(2) or 0), (m.group(3) or "").upper()
    if ap == "PM" and h != 12:
        h += 12
    if ap == "AM" and h == 12:
        h = 0
    if not (0 <= h <= 23 and 0 <= mi <= 59):
        return None
    return f"{h:02d}:{mi:02d}"


def _date_from_cell(value):
    if isinstance(value, (dt.date, dt.datetime)):
        return value.date() if isinstance(value, dt.datetime) else value
    text = str(value or "").strip()
    for fmt in _DATE_FORMATS:
        try:
            return dt.datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    return None


def _split_names(cell) -> list:
    return [p.strip() for p in re.split(r"[,;\n]+", str(cell or "")) if p.strip()]


def _workbook_rows(source) -> dict:
    """{sheet title: rows} from an Excel path or the file's bytes. The workbook is always closed afterwards:
    a read-only workbook keeps its file open, and on Windows an open file cannot be deleted or replaced."""
    from openpyxl import load_workbook
    wb = load_workbook(io.BytesIO(bytes(source)) if isinstance(source, (bytes, bytearray)) else source,
                       data_only=True, read_only=True)
    try:
        return {ws.title: [[("" if c is None else c) for c in row] for row in ws.iter_rows(values_only=True)]
                for ws in wb.worksheets}
    finally:
        wb.close()


def _is_xlsx_bytes(source) -> bool:
    return isinstance(source, (bytes, bytearray)) and bytes(source[:4]) == b"PK\x03\x04"


def _rows_from_input(source) -> list:
    """Accept pasted text (tabs or commas inside quotes), a CSV/TSV/XLSX path, or a list of rows."""
    if isinstance(source, list):
        return [[("" if c is None else c) for c in row] for row in source]
    if isinstance(source, str) and os.path.isfile(source):
        if source.lower().endswith((".xlsx", ".xlsm")):
            return next(iter(_workbook_rows(source).values()), [])
        with open(source, encoding="utf-8-sig") as fh:
            source = fh.read()
    import csv
    text = str(source)
    delim = "\t" if "\t" in text else ","
    return [row for row in csv.reader(io.StringIO(text), delimiter=delim)]


def parse_by_shift(source) -> dict:
    """Read a by-shift roster (Day | Date | shift columns | Week off | Leave).

    Returns {"shifts": [{"name", "start", "end"}], "days": [{"date", "shifts": {name: [people]},
    "off": [...], "leave": [...]}], "warnings": [...]}. Repeated header rows and blank rows are skipped.
    """
    rows = _rows_from_input(source)
    warnings = []
    head_i, cols = None, {}
    for i, row in enumerate(rows):
        labels = [str(c or "").strip().upper() for c in row]
        if any("DATE" == x for x in labels) and any("SHIFT" in x or x in ("MORNING", "AFTERNOON", "NIGHT") for x in labels):
            head_i = i
            for j, x in enumerate(labels):
                if x == "DATE":
                    cols["date"] = j
                elif "WEEK" in x and "OFF" in x or x in ("OFF", "DAY OFF", "WEEKOFF"):
                    cols["off"] = j
                elif "LEAVE" in x:
                    cols["leave"] = j
                elif "SHIFT" in x or x in ("MORNING", "AFTERNOON", "NIGHT", "GENERAL"):
                    name = x.replace("SHIFT", "").strip().title() or f"Shift {j}"
                    cols.setdefault("shifts", []).append((j, name))
            break
    if head_i is None or "date" not in cols or not cols.get("shifts"):
        raise RosterError(400, "Could not find the header row: it needs a DATE column and shift columns.")
    shifts = []
    times_row = rows[head_i + 1] if head_i + 1 < len(rows) else []
    for j, name in cols["shifts"]:
        cell = str(times_row[j] if j < len(times_row) else "").upper()
        parts = re.split(r"\s+TO\s+|\s*-\s*", cell)
        start = _time_from_words(parts[0]) if len(parts) == 2 else None
        end = _time_from_words(parts[1]) if len(parts) == 2 else None
        if not (start and end):
            warnings.append(f"No times found for {name}; set them on the shift before using rest checks.")
        shifts.append({"name": name, "start": start or "00:00", "end": end or "00:00", "column": j})
    days, seen = [], {}
    for i, row in enumerate(rows[head_i + 1:], start=head_i + 2):
        if not any(str(c or "").strip() for c in row):
            continue
        d = _date_from_cell(row[cols["date"]] if cols["date"] < len(row) else "")
        if not d:
            continue  # the times row, a repeated header, or a note
        if str(d) in seen:
            warnings.append(f"{d} appears twice (rows {seen[str(d)]} and {i}); the later row is used.")
            days = [x for x in days if x["date"] != str(d)]
        seen[str(d)] = i
        get = lambda j: row[j] if j is not None and j < len(row) else ""
        entry = {"date": str(d), "row": i, "shifts": {s["name"]: _split_names(get(s["column"])) for s in shifts},
                 "off": _split_names(get(cols.get("off"))), "leave": _split_names(get(cols.get("leave")))}
        placed = {}
        for where, people in [(s, p) for s, p in entry["shifts"].items()] + [("Week off", entry["off"]), ("Leave", entry["leave"])]:
            for p in people:
                key = norm(p)
                if key in placed:
                    warnings.append(f"{d}: {p} is listed under both {placed[key]} and {where}.")
                placed[key] = where
        days.append(entry)
    if not days:
        raise RosterError(400, "No dated rows found under the header.")
    everyone = sorted({p for e in days for p in sum(e["shifts"].values(), []) + e["off"] + e["leave"]}, key=str.upper)
    warnings += similar_names(everyone)
    gaps = {}
    for e in sorted(days, key=lambda x: x["date"]):
        listed = {norm(p) for p in sum(e["shifts"].values(), []) + e["off"] + e["leave"]}
        for p in everyone:
            if norm(p) not in listed:
                gaps.setdefault(p, []).append(parse_date(e["date"]))
    for p, ds in gaps.items():
        spans, a = [], ds[0]
        for prev, cur in zip(ds, ds[1:] + [None]):
            if cur is None or (cur - prev).days != 1:
                spans.append(str(a) if a == prev else f"{a} to {prev}")
                a = cur
        warnings.append(f"{p} is not listed on {len(ds)} day(s): {', '.join(spans)}.")
    return {"shifts": [{k: v for k, v in s.items() if k != "column"} for s in shifts], "days": days,
            "people": everyone, "warnings": warnings}


PERSON_HEADERS = ("MEMBER", "NAME", "EMPLOYEE", "ENGINEER", "PERSON")


def parse_by_person(source, shift_rows=None) -> dict:
    """Read a by-person roster (Member | Role | one column per date, cells holding shift codes).

    shift_rows: rows of a "Shifts" sheet (Code, Name, Start, End) to learn names and times.
    Returns the same structure as parse_by_shift, so either can create a roster.
    """
    rows = _rows_from_input(source)
    head_i = None
    for i, row in enumerate(rows):
        first = str(row[0] if row else "").strip().upper()
        if first in PERSON_HEADERS and sum(1 for c in row[1:] if _date_from_cell(c)) >= 1:
            head_i = i
            break
    if head_i is None:
        raise RosterError(400, "Could not find the header row: it needs a Member (or Name) column and date columns.")
    header = rows[head_i]
    role_col = next((j for j, c in enumerate(header) if str(c or "").strip().upper() == "ROLE"), None)
    date_cols = [(j, _date_from_cell(c)) for j, c in enumerate(header) if j > 0 and _date_from_cell(c)]
    known = {}
    for row in shift_rows or []:
        code = str(row[0] if row else "").strip().upper()
        if code and code != "CODE" and len(row) >= 4:
            start = _time_from_words(str(row[2] or "")) or str(row[2] or "").strip()
            end = _time_from_words(str(row[3] or "")) or str(row[3] or "").strip()
            known[code] = {"name": str(row[1] or code).strip(), "start": start, "end": end}
    warnings, people, roles = [], [], {}
    cells = {}  # (date, person) -> code
    for row in rows[head_i + 1:]:
        person = re.sub(r"\s+", " ", str(row[0] if row else "")).strip()
        if not person:
            continue
        people.append(person)
        if role_col is not None and role_col < len(row) and str(row[role_col] or "").strip():
            roles[person] = str(row[role_col]).strip()
        for j, d in date_cols:
            code = str(row[j] if j < len(row) else "").strip().upper()
            if code:
                cells[(str(d), person)] = code
    if not people:
        raise RosterError(400, "No people found under the header.")
    codes = sorted({c for c in cells.values() if c not in RESERVED_CODES})
    shifts = []
    for c in codes:
        k = known.get(c)
        if k and TIME_RE.match(k["start"]) and TIME_RE.match(k["end"]):
            shifts.append({"name": k["name"], "start": k["start"], "end": k["end"], "code": c})
        else:
            warnings.append(f"No times found for shift code {c}; set them on the shift before using rest checks.")
            shifts.append({"name": (k or {}).get("name") or c, "start": "00:00", "end": "00:00", "code": c})
    name_of = {s["code"]: s["name"] for s in shifts}
    days = []
    for _, d in date_cols:
        e = {"date": str(d), "shifts": {s["name"]: [] for s in shifts}, "off": [], "leave": []}
        for p in people:
            code = cells.get((str(d), p))
            if code == LEAVE:
                e["leave"].append(p)
            elif code == OFF:
                e["off"].append(p)
            elif code in name_of:
                e["shifts"][name_of[code]].append(p)
        days.append(e)
    warnings += similar_names(sorted(set(people), key=str.upper))
    return {"shifts": [{k: v for k, v in s.items() if k != "code"} for s in shifts], "days": days,
            "people": people, "roles": roles, "warnings": warnings, "layout": "by_person"}


def parse_roster_file(source) -> dict:
    """Read either layout from a file path, an Excel file's bytes, or pasted text, choosing the sheet and layout itself."""
    if _is_xlsx_bytes(source) or (isinstance(source, str) and os.path.isfile(source) and source.lower().endswith((".xlsx", ".xlsm"))):
        try:
            sheet_rows = _workbook_rows(source)
        except (OSError, KeyError, ValueError, zipfile.BadZipFile) as exc:
            raise RosterError(400, f"The Excel file could not be read ({exc.__class__.__name__}). Save it again as .xlsx and retry.")
        order = list(sheet_rows)  # the workbook's own order: its first sheet is its main layout
        errors = []
        for title in order:
            for parser in ("by_shift", "by_person"):
                try:
                    if parser == "by_shift":
                        parsed = parse_by_shift(sheet_rows[title])
                        parsed["layout"] = "by_shift"
                    else:
                        parsed = parse_by_person(sheet_rows[title], sheet_rows.get("Shifts"))
                    parsed["sheet"] = title
                    return parsed
                except RosterError as exc:
                    errors.append(f"{title}: {exc.message}")
        raise RosterError(400, "No roster found in this workbook. Expected a by-shift sheet (DATE and shift "
                               "columns) or a by-person sheet (Member and date columns).")
    try:
        parsed = parse_by_shift(source)
        parsed["layout"] = "by_shift"
        return parsed
    except RosterError:
        return parse_by_person(source)


def similar_names(names: list) -> list:
    """Warn about names that may be one person typed two ways (MUKKU / NMUKKU)."""
    import difflib
    out = []
    keys = [(n, norm(n)) for n in names]
    for i, (a, ka) in enumerate(keys):
        for b, kb in keys[i + 1:]:
            if ka == kb:
                continue
            close = difflib.SequenceMatcher(None, ka, kb).ratio() >= 0.85
            shifted = len(ka) > 3 and len(kb) > 3 and (ka.endswith(kb) or kb.endswith(ka)) and abs(len(ka) - len(kb)) <= 2
            if close or shifted:
                out.append(f"'{a}' and '{b}' may be the same person.")
    return out


def roster_from_by_shift(parsed: dict, team: str, name: str, actor) -> dict:
    """A new roster built from an imported by-shift sheet: its shifts, its people, and every listed day."""
    _require(actor, 3, "import a roster")
    team, name = re.sub(r"\s+", " ", str(team)).strip(), re.sub(r"\s+", " ", str(name)).strip()
    first = min(e["date"] for e in parsed["days"])
    r = default_roster(team, name, first, "three_shift")
    if parsed.get("layout") == "by_person":
        r["layout"] = "by_person"
    codes = []
    r["shifts"] = []
    for s in parsed["shifts"]:
        base = re.sub(r"[^A-Z0-9]", "", s["name"].upper())[:1] or "S"
        code, k = base, 2
        while code in codes or code in RESERVED_CODES:
            code, k = f"{base}{k}", k + 1
        codes.append(code)
        worked = sorted({parse_date(e["date"]).weekday() for e in parsed["days"] if e["shifts"].get(s["name"])})
        r["shifts"].append({"code": code, "name": s["name"], "start": s["start"], "end": s["end"],
                            "days": worked or [0, 1, 2, 3, 4, 5, 6], "min_staff": 1, "max_staff": None,
                            "color": {"M": "#F2C76E", "A": "#7AA7FF", "N": "#A391F5"}.get(code, "#43C08F"),
                            "log_shift": s["name"] if s["name"] in ("Morning", "Afternoon", "Night") else "Morning"})
    r["rules"]["forbidden_sequences"] = [["N", "M"]] if {"N", "M"} <= set(codes) else []
    r["cycle"] = {"type": "custom", "length_days": 1}
    r["rotation"] = {"sequence": [OFF]}  # imported days are explicit; anything outside them is off
    r["imported"] = True
    roles = parsed.get("roles") or {}
    r["members"] = [{"id": new_id("mem_"), "name": p, "role": roles.get(p, "Engineer"), "user": "", "offset": 0}
                    for p in parsed["people"]]
    apply_by_shift(r, actor, parsed)
    data_days = [e for e in parsed["days"] if any(e["shifts"].values()) or e["off"] or e["leave"]]
    for s in r["shifts"]:
        counts = [len(e["shifts"].get(s["name"]) or []) for e in data_days
                  if parse_date(e["date"]).weekday() in s["days"]]
        s["min_staff"] = max(1, min(counts)) if counts else 1  # the fewest people the sheet ever puts on it
    validate_roster(r)
    return r


def import_notes(r: dict) -> list:
    """Plain notes about an imported roster: the period it covers and the minimum staff it was given."""
    days = sorted(d for d, v in (r.get("overrides") or {}).items() if v)
    notes = []
    if days:
        notes.append(f"The sheet covers {days[0]} to {days[-1]} ({len(days)} day(s) with people on them). "
                     "Days outside that show as \"No roster for this day\" and are not counted as problems.")
    mins = ", ".join(f"{s['name']} {s['min_staff']}" for s in r.get("shifts", []))
    if mins:
        notes.append(f"Minimum people per shift were set from your sheet: {mins}. Change them in Settings if your real minimum differs.")
    return notes


def group_issues(issues: list) -> list:
    """Collapse repeated findings into one line each with date ranges (Morning has 0 of 1 people: 1-8 Nov)."""
    groups = {}
    for i in issues:
        key = (i["level"], i["rule"], i["message"])
        groups.setdefault(key, []).append(i.get("date"))
    out = []
    for (level, rule, message), dates in groups.items():
        ds = sorted({parse_date(d) for d in dates if d})
        spans, a = [], None
        for k, d in enumerate(ds):
            if a is None:
                a = d
            if k == len(ds) - 1 or (ds[k + 1] - d).days != 1:
                spans.append(str(a) if a == d else f"{a} to {d}")
                a = None
        out.append({"level": level, "rule": rule, "message": message, "count": len(dates), "dates": spans,
                    "first": str(ds[0]) if ds else ""})
    return sorted(out, key=lambda g: (g["level"] != "error", g["first"], g["message"]))


def apply_by_shift(r: dict, actor, parsed: dict) -> dict:
    """Write an imported sheet into an existing roster as day-by-day entries and approved leave."""
    _require(actor, 3, "import a roster")
    by_name = {norm(m["name"]): m["id"] for m in r.get("members", [])}
    code_of = {s["name"].lower(): s["code"] for s in r.get("shifts", [])}
    unknown, written = set(), 0
    for e in parsed["days"]:
        ov = r.setdefault("overrides", {}).setdefault(e["date"], {})
        for sname, people in e["shifts"].items():
            code = code_of.get(sname.lower())
            if not code:
                raise RosterError(400, f"The roster has no shift named '{sname}'.")
            for p in people:
                mid = by_name.get(norm(p))
                if mid:
                    ov[mid] = code
                    written += 1
                else:
                    unknown.add(p)
        for p in e["off"]:
            mid = by_name.get(norm(p))
            if mid:
                ov[mid] = OFF
                written += 1
            else:
                unknown.add(p)
        for p in e["leave"]:
            mid = by_name.get(norm(p))
            if not mid:
                unknown.add(p)
                continue
            ov.setdefault(mid, OFF)  # marks the day as covered by the sheet; the leave below shows instead
            r.setdefault("leaves", []).append({
                "id": new_id("lv_"), "member_id": mid, "start": e["date"], "end": e["date"], "type": "Other",
                "note": "Imported from roster sheet", "status": "approved", "requested_by": actor_name(actor),
                "requested_at": now_iso(), "decided_by": actor_name(actor), "decided_at": now_iso()})
            written += 1
    for d in [d for d, v in r.get("overrides", {}).items() if not v]:
        del r["overrides"][d]  # a dated row with nobody on it is not data
    return {"entries": written, "unknown_people": sorted(unknown)}


# --------------------------------------------------------------------------- blank templates


def template_file(layout: str = "by_shift", fmt: str = "xlsx", start=None, month=None):
    """A blank roster sheet to fill in and import: (bytes, content type, file name).

    by_shift:  one row per day; names under each shift, Week off and Leave (comma separated).
    by_person: one row per person; a shift code in each date column, with a Shifts sheet for codes and times.
    The two example rows show the format; replace them with your own people.
    """
    if layout not in ("by_shift", "by_person"):
        raise RosterError(400, "Layout must be by_shift or by_person.")
    if fmt not in ("xlsx", "csv"):
        raise RosterError(400, "Format must be xlsx or csv.")
    if month or not start:
        # a whole calendar month (default: this month), so an import covers every day of it
        try:
            y, mth = (int(x) for x in str(month or dt.date.today().strftime("%Y-%m")).split("-")[:2])
            first = dt.date(y, mth, 1)
        except ValueError:
            raise RosterError(400, "Month must look like 2026-11.")
        nxt = (first.replace(day=28) + dt.timedelta(days=4)).replace(day=1)
        days = [first + dt.timedelta(days=i) for i in range((nxt - first).days)]
    else:
        first = parse_date(start)
        first = first - dt.timedelta(days=first.weekday())
        days = [first + dt.timedelta(days=i) for i in range(14)]
    base = default_roster("Team", "Template", str(first), "three_shift" if layout == "by_shift" else "standard")
    shifts = base["shifts"]
    ex1, ex2 = "Asha Rao", "Ben Cole"
    if layout == "by_shift":
        head = ["Day", "DATE"] + [f"{x['name'].upper()} SHIFT" for x in shifts] + ["WEEK OFF", "LEAVE"]
        times = ["", ""] + [f"{_ampm(x['start'])} TO {_ampm(x['end'])}" for x in shifts] + ["", ""]
        rows = []
        for i, d in enumerate(days):
            cells = [""] * (len(shifts) + 2)
            if i == 0:
                cells[0], cells[len(shifts)] = ex1, ex2
            elif i == 1:
                cells[1], cells[len(shifts) + 1] = ex1, ex2
            rows.append([WEEKDAYS_LONG[d.weekday()], d.strftime("%d-%b-%y")] + cells)
        if fmt == "csv":
            import csv
            buf = io.StringIO()
            w = csv.writer(buf)
            for row in [head, times] + rows:
                w.writerow(row)
            return buf.getvalue().encode("utf-8-sig"), "text/csv; charset=utf-8", "roster_template_by_shift.csv"
        sheet_rows, title = [head, times] + rows, "Roster"
    else:
        if fmt == "csv":
            raise RosterError(400, "The by-person template is Excel only (it needs its Shifts sheet).")
        cycle = [x["code"] for x in shifts if x["code"] != "D"] + [OFF]
        sheet_rows = [["Template roster: one row per person, a shift code in each date column"],
                      ["Member", "Role"] + days,
                      [ex1, "Engineer"] + [cycle[i % len(cycle)] for i in range(len(days))],
                      [ex2, "Engineer"] + [cycle[(i + 2) % len(cycle)] for i in range(len(days))]]
        title = "Calendar"
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    F = "Arial"
    wb = Workbook()
    ws = wb.active
    ws.title = title
    for row in sheet_rows:
        ws.append(row)
    example_fill = PatternFill("solid", fgColor="FFF2CC")
    for row_cells in ws.iter_rows():
        for c in row_cells:
            c.font = Font(name=F)
            if isinstance(c.value, dt.date):
                c.number_format = "ddd d-mmm"
    head_row = 1 if layout == "by_shift" else 2
    for c in ws[head_row]:
        c.font = Font(name=F, bold=True)
    if layout == "by_shift":
        for c in ws[2]:
            c.font = Font(name=F, italic=True)
        for r_i in (3, 4):
            for c in ws[r_i]:
                c.fill = example_fill
        ws.column_dimensions["A"].width = 12
        ws.column_dimensions["B"].width = 12
        for col in "CDEFG":
            ws.column_dimensions[col].width = 30
        for row_cells in ws.iter_rows(min_row=3):
            for c in row_cells:
                c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.freeze_panes = "C3"
    else:
        ws["A1"].font = Font(name=F, bold=True, size=12)
        for r_i in (3, 4):
            for c in ws[r_i]:
                c.fill = example_fill
        ws.column_dimensions["A"].width = 22
        ws.freeze_panes = "C3"
        sh = wb.create_sheet("Shifts")
        sh.append(["Code", "Name", "Start", "End"])
        for x in shifts:
            sh.append([x["code"], x["name"], x["start"], x["end"]])
        sh.append([OFF, "Day off", "", ""])
        sh.append([LEAVE, "Leave", "", ""])
        for row_cells in sh.iter_rows():
            for c in row_cells:
                c.font = Font(name=F, bold=c.row == 1)
    ins = wb.create_sheet("Instructions")
    lines = (["How to fill in this roster (by shift)",
              "1. Keep the header row and the times row. Change the shift names or times in those two rows if yours differ.",
              "2. One row per day, already filled in for the whole month. Type any extra dates as 9-Nov-26 (or 2026-11-09).",
              "   Fill in every day: a day with nobody listed is treated as having no roster yet.",
              "3. Under each shift, list the people working it, separated by commas: Asha Rao, Ben Cole.",
              "4. List people who are off under WEEK OFF, and people on leave under LEAVE.",
              "5. The two shaded rows are examples: replace them with your own people.",
              "6. Import it from the Shift roster page: Import sheet. Names must be spelled the same way every day."]
             if layout == "by_shift" else
             ["How to fill in this roster (by person)",
              "1. One row per person: name in column A, role in column B.",
              "2. Row 2 holds the dates for the whole month. Fill in every day: an empty day is treated as having no roster yet.",
              "3. In each date cell, type a shift code from the Shifts sheet (M, A, N, D), OFF for a day off, or LEAVE.",
              "4. Change codes, names or times on the Shifts sheet if yours differ.",
              "5. The two shaded rows are examples: replace them with your own people.",
              "6. Import it from the Shift roster page: Import sheet."])
    for i, t in enumerate(lines):
        ins.append([t])
        ins.cell(row=i + 1, column=1).font = Font(name=F, bold=i == 0, size=12 if i == 0 else 10)
    ins.column_dimensions["A"].width = 110
    buf = io.BytesIO()
    wb.save(buf)
    return (buf.getvalue(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            f"roster_template_{layout}.xlsx")


def _ampm(hhmm: str) -> str:
    h, m = (int(x) for x in hhmm.split(":"))
    suffix = "AM" if h < 12 else "PM"
    h12 = h % 12 or 12
    return f"{h12}:{m:02d} {suffix}" if m else f"{h12}{suffix}"  # 9:30 PM, 6AM, as in common roster sheets


# --------------------------------------------------------------------------- command line


ADMIN = {"user": "cli", "role": "Admin"}


def build_demo(store: RosterStore, template: str = "standard") -> dict:
    if template == "three_shift":
        r = default_roster("NOC", "NOC three-shift", "2026-11-02", "three_shift")
        r["holidays"] = {"region": "IN", "subdivision": "KA", "custom": [], "ignore": []}
        r["rotation"]["sequence"] = ["M"] * 5 + [OFF] * 2 + ["A"] * 5 + [OFF] * 2 + ["N"] * 5 + [OFF] * 2
        r = store.create(ADMIN, r)
        names = ["Asha Rao", "Ben Cole", "Chen Li", "Dina Shah", "Eli Moss", "Faye Kim", "Gus Hart", "Hana Ito",
                 "Ivan Popov", "Jade Wu"]
        offs = [0, 1, 6, 11, 16, 3, 8, 13, 18, 20]
        ids = [store.add_member(ADMIN, r["id"], n, "Engineer", n.split()[0].lower(), o)["id"] for n, o in zip(names, offs)]
        store.request_leave(ADMIN, r["id"], ids[2], "2026-11-09", "2026-11-10", "Annual", "Planned leave")
        return store.get(r["id"])
    r = default_roster("Unix", "Unix SA follow-the-sun", "2026-11-02")
    r["holidays"] = {"region": "US", "subdivision": "", "custom": [], "ignore": []}
    r["oncall"].update(enabled=True, escalation=[{"level": 1, "name": "Unix team lead", "contact": "+44 20 7946 0001"},
                                                 {"level": 2, "name": "Infrastructure manager", "contact": "+44 20 7946 0002"}])
    r = store.create(ADMIN, r)
    ids = [store.add_member(ADMIN, r["id"], n, role, u, off)["id"] for n, role, u, off in
           [("Mia Silva", "Engineer", "mia.s", 0), ("Ravi Nair", "Engineer", "ravi.n", 1), ("Ana Ito", "Engineer", "ana", 6),
            ("Tom Berg", "Engineer", "tom.b", 11), ("Lena Fox", "Engineer", "lena.f", 16)]]
    store.update(ADMIN, r["id"], {"oncall": dict(store.get(r["id"])["oncall"], order={"Primary": ids, "Secondary": ids[1:] + ids[:1]})})
    store.request_leave(ADMIN, r["id"], ids[1], "2026-11-16", "2026-11-18", "Annual", "Family event")
    store.request_swap(ADMIN, r["id"], ids[0], "2026-11-10", ids[3], "2026-11-12", "swap", "Doctor appointment")
    return store.get(r["id"])


def _find_member(r: dict, who: str) -> str:
    key = norm(who)
    for m in r.get("members", []):
        if m["id"] == who or norm(m["name"]) == key or (m.get("user") and norm(m["user"]) == key):
            return m["id"]
    raise RosterError(404, f"No member called '{who}' in this roster.")


def _actor(a) -> dict:
    return {"user": a.as_user or os.environ.get("USERNAME") or os.environ.get("USER") or "cli", "role": a.as_role}


def _print_requests(r: dict) -> None:
    names = {m["id"]: m["name"] for m in r.get("members", [])}
    rows = [("leave", x["id"], names.get(x["member_id"], "?"), f"{x['start']} to {x['end']}", x.get("type", ""), x["status"])
            for x in r.get("leaves", [])]
    rows += [(x["kind"], x["id"], f"{names.get(x['member_a'], '?')} / {names.get(x['member_b'], '?')}",
              x["date_a"] if x["date_a"] == x["date_b"] else f"{x['date_a']} / {x['date_b']}",
              "; ".join(x.get("warnings") or []) or "-", x["status"]) for x in r.get("swaps", [])]
    if not rows:
        print("No requests.")
    for row in rows:
        print("  ".join(str(c) for c in row))


def main(argv=None):
    p = argparse.ArgumentParser(description="MSP Ops Center shift roster")
    sub = p.add_subparsers(dest="cmd", required=True)
    d = sub.add_parser("demo", help="create a sample roster and export November 2026")
    d.add_argument("--out", default=".")
    d.add_argument("--template", choices=sorted(TEMPLATES), default="standard")
    im = sub.add_parser("import", help="create a roster from a by-shift sheet (.xlsx, .csv or tab-separated .txt)")
    im.add_argument("--file", required=True)
    im.add_argument("--team", required=True)
    im.add_argument("--name", required=True)
    im.add_argument("--data", default=".", help="folder for msp_roster.json")
    im.add_argument("--out", default=".")
    x = sub.add_parser("export", help="export a roster to Excel and printable HTML")
    x.add_argument("--data", required=True, help="folder holding msp_roster.json")
    x.add_argument("--roster", required=True)
    x.add_argument("--start", required=True)
    x.add_argument("--end", required=True)
    x.add_argument("--out", default=".")
    def roster_cmd(name, help_text):
        c = sub.add_parser(name, help=help_text)
        c.add_argument("--data", default=".", help="folder holding msp_roster.json")
        c.add_argument("--roster", required=True, help="roster id or name")
        c.add_argument("--as-user", default="", help="who is acting (default: your login name)")
        c.add_argument("--as-role", default="Lead", choices=sorted(ROLE_RANK), help="role of the person acting")
        return c
    lv = roster_cmd("leave", "request leave (approved at once if leave is self-recorded or you are a Lead)")
    lv.add_argument("--member", required=True)
    lv.add_argument("--start", required=True)
    lv.add_argument("--end", required=True)
    lv.add_argument("--type", default="Annual")
    lv.add_argument("--note", default="")
    sw = roster_cmd("swap", "swap shifts between two people, or have one cover the other")
    sw.add_argument("--member", required=True, help="the person giving up the shift")
    sw.add_argument("--date", required=True)
    sw.add_argument("--with", dest="other", required=True, help="the other person")
    sw.add_argument("--other-date", default=None, help="for a swap on two different days")
    sw.add_argument("--cover", action="store_true", help="the other person covers; no shift is given back")
    sw.add_argument("--note", default="")
    for verb in ("approve", "reject"):
        c = roster_cmd(verb, f"{verb} a pending leave or swap request (Lead or Admin)")
        c.add_argument("--request", required=True, help="request id from the requests command")
        c.add_argument("--note", default="")
    roster_cmd("requests", "list leave and swap requests with their status")
    hd = roster_cmd("holidays", "set the holiday region and company holidays, and list the holidays")
    hd.add_argument("--region", choices=["", *HOLIDAY_REGIONS], default=None, help="US or IN; empty to clear")
    hd.add_argument("--subdivision", default=None, help="state code, for example KA, MH, TN (India) or CA, NY (US)")
    hd.add_argument("--add", action="append", default=[], metavar="DATE=NAME", help="add a company holiday")
    hd.add_argument("--ignore", action="append", default=[], metavar="DATE", help="drop a region holiday")
    hd.add_argument("--year", type=int, default=None, help="year to list (default: this year)")
    a = p.parse_args(argv)
    if a.cmd in ("leave", "swap", "approve", "reject", "requests", "holidays"):
        store = RosterStore(FileStorage(a.data))
        all_r = store.rosters()
        match = [x for x in all_r if x["id"] == a.roster or norm(x["name"]) == norm(a.roster)]
        if not match:
            raise RosterError(404, f"No roster '{a.roster}' in {a.data}.")
        r, actor = match[0], _actor(a)
        if a.cmd == "leave":
            x = store.request_leave(actor, r["id"], _find_member(r, a.member), a.start, a.end, a.type, a.note)
            print(f"Leave {x['id']}: {x['status']}.")
        elif a.cmd == "swap":
            x = store.request_swap(actor, r["id"], _find_member(r, a.member), a.date, _find_member(r, a.other),
                                   a.other_date, "cover" if a.cover else "swap", a.note)
            print(f"{x['kind'].title()} {x['id']}: {x['status']}.")
            for w in x.get("warnings") or []:
                print("Check:", w)
        elif a.cmd in ("approve", "reject"):
            kind = "leave" if a.request.startswith("lv_") else "swap"
            x = store.decide(actor, r["id"], kind, a.request, a.cmd == "approve", a.note)
            print(f"{a.request}: {x['status']}.")
        elif a.cmd == "requests":
            _print_requests(r)
        else:
            if a.region is not None or a.subdivision is not None or a.add or a.ignore:
                cfg = dict(r.get("holidays") or {"region": "", "subdivision": "", "custom": [], "ignore": []})
                if a.region is not None:
                    cfg["region"] = a.region
                if a.subdivision is not None:
                    cfg["subdivision"] = a.subdivision.upper()
                for item in a.add:
                    if "=" not in item:
                        raise RosterError(400, "Use --add DATE=NAME, for example --add 2026-11-10=\"Company day\".")
                    d, n = item.split("=", 1)
                    cfg["custom"] = [x for x in cfg.get("custom", []) if x["date"] != str(parse_date(d))] + [{"date": str(parse_date(d)), "name": n.strip()}]
                cfg["ignore"] = sorted(set(cfg.get("ignore", [])) | {str(parse_date(d)) for d in a.ignore})
                r = store.update(actor, r["id"], {"holidays": cfg})
            if (r.get("holidays") or {}).get("region") and not holidays_available():
                print("Note: install the region list with: pip install holidays")
            year = a.year or dt.date.today().year
            for d, n in holidays_for(r, f"{year}-01-01", f"{year}-12-31").items():
                print(d, WEEKDAYS[parse_date(d).weekday()], n)
        return
    if a.cmd == "demo":
        store = RosterStore(FileStorage(a.out))
        r = build_demo(store, a.template)
        start, end = "2026-11-02", "2026-11-29"
    elif a.cmd == "import":
        store = RosterStore(FileStorage(a.data))
        if not os.path.isfile(a.file):
            raise RosterError(404, f"File not found: {a.file}")
        res = store.import_new(ADMIN, a.file, a.team, a.name)
        r = res["roster"]
        print(f"Read the {res['layout'].replace('_', ' ')} layout" + (f" from sheet '{res['sheet']}'" if res.get("sheet") else "") + ".")
        for w in res["warnings"]:
            print("Check:", w)
        dates = sorted(r["overrides"])
        start, end = dates[0], dates[-1]
    else:
        store = RosterStore(FileStorage(a.data))
        r = store.get(a.roster)
        start, end = a.start, a.end
    os.makedirs(a.out, exist_ok=True)
    base = os.path.join(a.out, "roster_" + norm(r["name"])[:30])
    with open(base + ".xlsx", "wb") as fh:
        fh.write(export_excel(r, start, end))
    with open(base + ".html", "w", encoding="utf-8") as fh:
        fh.write(export_html(r, start, end, history=store.history(r["id"])))
    issues = check(r, generate(r, start, end))
    print(f"Roster '{r['name']}' ({r['id']}): {len(issues)} check finding(s).")
    print("Wrote", base + ".xlsx", "and", base + ".html")


if __name__ == "__main__":
    import sys
    try:
        main()
    except RosterError as exc:
        print("Error:", exc.message, file=sys.stderr)
        sys.exit(2)
    except FileNotFoundError as exc:
        print("Error: file not found:", exc.filename, file=sys.stderr)
        sys.exit(2)
