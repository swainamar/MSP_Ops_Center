#!/usr/bin/env python3
"""MSP Ops Center: teams, customers and projects (version 1).

Who does what:
    Admin       creates customers, their projects and teams; assigns team leads; assigns teams to customers
                and marks one team per customer as primary (it owns the core CI data). Admins are never
                team members, so they gain no access to team data through a team.
    Team lead   (a user with the Lead role, named as a lead of the team) adds and removes members, creates
                Engineer or Viewer users for the team, and assigns members to projects of the customers the
                team serves. A person may lead several teams and belong to several teams.
    Everyone    sees the teams they belong to or lead, and the customers and projects those teams serve.

Stored in msp_org.json through the same storage interface as the rest of the app (local folder or
SharePoint), with optimistic concurrency, an audit trail and 30 restorable versions.
"""
from __future__ import annotations

import copy
import datetime as dt
import json
import re
import uuid

ORG_FILE = "msp_org.json"
MAX_VERSIONS = 30
MAX_AUDIT = 3000
ROLE_RANK = {"Viewer": 1, "Engineer": 2, "Lead": 3, "Admin": 4}
NAME_RE = re.compile(r"^[^\x00-\x1f<>]{1,80}$")
CODE_RE = re.compile(r"^[A-Z0-9][A-Z0-9_-]{0,11}$")


class OrgError(Exception):
    def __init__(self, status: int, message: str):
        super().__init__(message)
        self.status = status
        self.message = message


def now_iso() -> str:
    return dt.datetime.now().astimezone().isoformat(timespec="seconds")


def new_id(prefix: str) -> str:
    return prefix + uuid.uuid4().hex[:10]


def norm(v) -> str:
    return str(v or "").strip().lower()


def clean_name(value, what: str) -> str:
    name = re.sub(r"\s+", " ", str(value or "")).strip()
    if not NAME_RE.fullmatch(name):
        raise OrgError(400, f"The {what} needs a name of 1 to 80 characters, without < or >.")
    return name


def empty_state() -> dict:
    return {"format": 1, "teams": {}, "customers": {}, "versions": [], "audit": []}


# --------------------------------------------------------------------------- queries


def is_admin(actor) -> bool:
    return (actor or {}).get("role") == "Admin"


def leads_team(state: dict, user: str, team_id: str) -> bool:
    t = state["teams"].get(team_id)
    return bool(t) and norm(user) in [norm(x) for x in t.get("leads", [])]


def member_of(state: dict, user: str, team_id: str) -> bool:
    t = state["teams"].get(team_id)
    return bool(t) and norm(user) in t.get("members", {})


def teams_of(state: dict, user: str) -> list:
    """Teams the person belongs to or leads, as ids."""
    u = norm(user)
    return [tid for tid, t in state["teams"].items() if u in t.get("members", {}) or u in [norm(x) for x in t.get("leads", [])]]


def customers_of_team(state: dict, team_id: str) -> list:
    return [cid for cid, c in state["customers"].items() if team_id in c.get("teams", [])]


def view(state: dict, actor: dict, users: dict) -> dict:
    """What this person may see. `users` is {username_lower: public user} from the app, for display names."""
    me = norm(actor.get("user"))
    admin = is_admin(actor)
    team_ids = list(state["teams"]) if admin else teams_of(state, me)
    cust_ids = list(state["customers"]) if admin else sorted({c for t in team_ids for c in customers_of_team(state, t)})

    def person(username):
        u = users.get(norm(username)) or {}
        return {"username": u.get("username", username), "name": u.get("name", username), "role": u.get("role", ""),
                "disabled": bool(u.get("disabled")), "exists": bool(u)}

    teams = []
    for tid in team_ids:
        t = state["teams"][tid]
        lead_here = admin or leads_team(state, me, tid)
        members = [dict(person(un), projects=m.get("projects", []), added_by=m.get("added_by", ""), added_at=m.get("added_at", ""))
                   for un, m in sorted(t.get("members", {}).items())]
        teams.append({"id": tid, "name": t["name"], "description": t.get("description", ""),
                      "leads": [person(x) for x in t.get("leads", [])], "members": members,
                      "customers": customers_of_team(state, tid), "can_manage": lead_here, "you_lead": leads_team(state, me, tid)})
    customers = []
    for cid in cust_ids:
        c = state["customers"][cid]
        customers.append({"id": cid, "name": c["name"], "code": c.get("code", ""), "teams": [t for t in c.get("teams", []) if t in state["teams"]],
                          "primary_team": c.get("primary_team"), "projects": [dict(p, id=pid) for pid, p in c.get("projects", {}).items()]})
    return {"teams": sorted(teams, key=lambda t: t["name"].lower()), "customers": sorted(customers, key=lambda c: c["name"].lower()),
            "is_admin": admin, "leads_any": any(t["you_lead"] for t in teams),
            "team_names": {tid: t["name"] for tid, t in state["teams"].items()}}


# --------------------------------------------------------------------------- changes (pure, on a state dict)


def _admin(actor):
    if not is_admin(actor):
        raise OrgError(403, "Only an administrator can do this.")


def _team(state, tid):
    t = state["teams"].get(tid)
    if not t:
        raise OrgError(404, "Team not found.")
    return t


def _customer(state, cid):
    c = state["customers"].get(cid)
    if not c:
        raise OrgError(404, "Customer not found.")
    return c


def _unique(items: dict, name: str, what: str, skip=None):
    if any(k != skip and norm(v["name"]) == norm(name) for k, v in items.items()):
        raise OrgError(409, f"A {what} called '{name}' already exists.")


def _check_leads(leads, users: dict) -> list:
    out = []
    for x in leads or []:
        u = users.get(norm(x))
        if not u:
            raise OrgError(400, f"There is no user '{x}'.")
        if u.get("role") != "Lead":
            raise OrgError(400, f"{u.get('name', x)} needs the Lead role before they can lead a team (Admin, Users).")
        if norm(u["username"]) not in [norm(y) for y in out]:
            out.append(u["username"])
    return out


def create_team(state, actor, data, users) -> dict:
    _admin(actor)
    name = clean_name(data.get("name"), "team")
    _unique(state["teams"], name, "team")
    tid = new_id("team_")
    state["teams"][tid] = {"name": name, "description": str(data.get("description") or "")[:300],
                           "leads": _check_leads(data.get("leads"), users), "members": {}, "created_at": now_iso(),
                           "created_by": actor["user"]}
    for lead in state["teams"][tid]["leads"]:
        state["teams"][tid]["members"].setdefault(norm(lead), {"projects": [], "added_by": actor["user"], "added_at": now_iso()})
    return dict(state["teams"][tid], id=tid)


def update_team(state, actor, tid, data, users) -> dict:
    t = _team(state, tid)
    if not (is_admin(actor) or leads_team(state, actor["user"], tid)):
        raise OrgError(403, "Only the team's leads or an administrator can change it.")
    if "description" in data:
        t["description"] = str(data.get("description") or "")[:300]
    admin_only = {k for k in ("name", "leads") if k in data}
    if admin_only:
        _admin(actor)
        if "name" in data:
            name = clean_name(data["name"], "team")
            _unique(state["teams"], name, "team", skip=tid)
            t["name"] = name
        if "leads" in data:
            t["leads"] = _check_leads(data["leads"], users)
            for lead in t["leads"]:
                t["members"].setdefault(norm(lead), {"projects": [], "added_by": actor["user"], "added_at": now_iso()})
    return dict(t, id=tid)


def delete_team(state, actor, tid) -> None:
    _admin(actor)
    _team(state, tid)
    del state["teams"][tid]
    for c in state["customers"].values():
        c["teams"] = [x for x in c.get("teams", []) if x != tid]
        if c.get("primary_team") == tid:
            c["primary_team"] = None


def add_member(state, actor, tid, username, users) -> dict:
    t = _team(state, tid)
    if not (is_admin(actor) or leads_team(state, actor["user"], tid)):
        raise OrgError(403, "Only the team's leads or an administrator can add members.")
    u = users.get(norm(username))
    if not u:
        raise OrgError(404, f"There is no user '{username}'.")
    if u.get("role") == "Admin":
        raise OrgError(400, "Administrators manage people and settings and are not team members (separation of duties).")
    if norm(username) in t["members"]:
        raise OrgError(409, f"{u.get('name', username)} is already in {t['name']}.")
    t["members"][norm(username)] = {"projects": [], "added_by": actor["user"], "added_at": now_iso()}
    return {"username": u["username"], "team": t["name"]}


def remove_member(state, actor, tid, username) -> None:
    t = _team(state, tid)
    if not (is_admin(actor) or leads_team(state, actor["user"], tid)):
        raise OrgError(403, "Only the team's leads or an administrator can remove members.")
    if norm(username) not in t["members"]:
        raise OrgError(404, "That person is not in this team.")
    if norm(username) in [norm(x) for x in t.get("leads", [])]:
        raise OrgError(409, "A team lead stays a member. An administrator can remove them as lead first.")
    del t["members"][norm(username)]


def set_member_projects(state, actor, tid, username, project_ids) -> list:
    t = _team(state, tid)
    if not (is_admin(actor) or leads_team(state, actor["user"], tid)):
        raise OrgError(403, "Only the team's leads or an administrator can assign projects.")
    m = t["members"].get(norm(username))
    if m is None:
        raise OrgError(404, "That person is not in this team.")
    allowed = {pid for cid in customers_of_team(state, tid) for pid in state["customers"][cid].get("projects", {})}
    chosen = list(dict.fromkeys(project_ids or []))
    bad = [p for p in chosen if p not in allowed]
    if bad:
        raise OrgError(400, "Members can only be assigned to projects of customers this team serves.")
    m["projects"] = chosen
    return chosen


def create_customer(state, actor, data) -> dict:
    _admin(actor)
    name = clean_name(data.get("name"), "customer")
    _unique(state["customers"], name, "customer")
    code = str(data.get("code") or "").strip().upper()
    if code and not CODE_RE.fullmatch(code):
        raise OrgError(400, "A customer code is up to 12 capital letters, digits, dash or underscore.")
    cid = new_id("cus_")
    state["customers"][cid] = {"name": name, "code": code, "teams": [], "primary_team": None, "projects": {},
                               "created_at": now_iso(), "created_by": actor["user"]}
    return dict(state["customers"][cid], id=cid)


def update_customer(state, actor, cid, data) -> dict:
    _admin(actor)
    c = _customer(state, cid)
    if "name" in data:
        name = clean_name(data["name"], "customer")
        _unique(state["customers"], name, "customer", skip=cid)
        c["name"] = name
    if "code" in data:
        code = str(data.get("code") or "").strip().upper()
        if code and not CODE_RE.fullmatch(code):
            raise OrgError(400, "A customer code is up to 12 capital letters, digits, dash or underscore.")
        c["code"] = code
    if "teams" in data:
        teams = list(dict.fromkeys(data.get("teams") or []))
        unknown = [x for x in teams if x not in state["teams"]]
        if unknown:
            raise OrgError(400, "One of those teams no longer exists.")
        removed = set(c.get("teams", [])) - set(teams)
        c["teams"] = teams
        if c.get("primary_team") not in teams:
            c["primary_team"] = None
        gone = set(c.get("projects", {}))
        for tid in removed:  # people keep no project access through a team that no longer serves this customer
            for m in state["teams"][tid]["members"].values():
                m["projects"] = [p for p in m.get("projects", []) if p not in gone]
    if "primary_team" in data:
        p = data.get("primary_team") or None
        if p is not None and p not in c.get("teams", []):
            raise OrgError(400, "The primary team must be one of the customer's teams.")
        c["primary_team"] = p
    return dict(c, id=cid)


def delete_customer(state, actor, cid) -> None:
    _admin(actor)
    c = _customer(state, cid)
    gone = set(c.get("projects", {}))
    del state["customers"][cid]
    for t in state["teams"].values():
        for m in t["members"].values():
            m["projects"] = [p for p in m.get("projects", []) if p not in gone]


def save_project(state, actor, cid, data, pid=None) -> dict:
    _admin(actor)
    c = _customer(state, cid)
    name = clean_name(data.get("name"), "project")
    if any(k != pid and norm(p["name"]) == norm(name) for k, p in c["projects"].items()):
        raise OrgError(409, f"{c['name']} already has a project called '{name}'.")
    if pid is None:
        pid = new_id("prj_")
        c["projects"][pid] = {"name": name, "active": True, "created_at": now_iso()}
    else:
        if pid not in c["projects"]:
            raise OrgError(404, "Project not found.")
        c["projects"][pid]["name"] = name
        if "active" in data:
            c["projects"][pid]["active"] = bool(data["active"])
    return dict(c["projects"][pid], id=pid)


def delete_project(state, actor, cid, pid) -> None:
    _admin(actor)
    c = _customer(state, cid)
    if c["projects"].pop(pid, None) is None:
        raise OrgError(404, "Project not found.")
    for t in state["teams"].values():
        for m in t["members"].values():
            m["projects"] = [p for p in m.get("projects", []) if p != pid]


def forget_user(state, username) -> int:
    """A deleted user leaves every team and lead list. Returns how many teams changed."""
    n = 0
    for t in state["teams"].values():
        before = (len(t["members"]), len(t.get("leads", [])))
        t["members"].pop(norm(username), None)
        t["leads"] = [x for x in t.get("leads", []) if norm(x) != norm(username)]
        n += before != (len(t["members"]), len(t["leads"]))
    return n


# --------------------------------------------------------------------------- storage


class OrgStore:
    """Loads and saves msp_org.json; every change keeps the previous state as a version."""

    def __init__(self, storage, name: str = ORG_FILE):
        self.storage = storage
        self.name = name

    def load(self):
        data, ver = self.storage.read(self.name)
        if not data:
            return empty_state(), ver
        state = json.loads(data.decode("utf-8"))
        for k, v in empty_state().items():
            state.setdefault(k, v)
        return state, ver

    def mutate(self, actor, action: str, target: str, fn):
        """Apply fn(state) and save; retry on a write conflict. Returns fn's result."""
        for _ in range(5):
            state, ver = self.load()
            before = {"teams": copy.deepcopy(state["teams"]), "customers": copy.deepcopy(state["customers"])}
            result = fn(state)
            state["versions"].append({"at": now_iso(), "by": actor["user"], "action": action, "target": target, "state": before})
            del state["versions"][:-MAX_VERSIONS]
            state["audit"].append({"at": now_iso(), "by": actor["user"], "action": action, "target": target})
            del state["audit"][:-MAX_AUDIT]
            try:
                self.storage.write(self.name, json.dumps(state, indent=1).encode("utf-8"), ver)
                return result
            except Exception as exc:  # any storage's ConflictError: reload and re-apply
                if type(exc).__name__ != "ConflictError":
                    raise
        raise OrgError(409, "The team setup kept changing while saving; please try again.")

    def history(self) -> list:
        vs = self.load()[0]["versions"]
        return [{"index": i, "at": v["at"], "by": v["by"], "action": v["action"], "target": v.get("target", "")} for i, v in enumerate(vs)]

    def restore(self, actor, index: int):
        _admin(actor)

        def fn(state):
            vs = state["versions"]
            if not 0 <= index < len(vs):
                raise OrgError(404, "Version not found.")
            state["teams"] = copy.deepcopy(vs[index]["state"]["teams"])
            state["customers"] = copy.deepcopy(vs[index]["state"]["customers"])
        self.mutate(actor, "Restore version", str(index), fn)
