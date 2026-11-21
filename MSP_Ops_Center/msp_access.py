#!/usr/bin/env python3
"""MSP Ops Center: team access rules (isolation bundle A).

All "who may see / change what" decisions for team-owned data live here, as plain functions with no web or
storage code, so they can be read, tested and diagnosed in one place. msp_runbook.py calls them.

CMDB (Asset_Master)
    * The owning team (column "Owner team"; empty means the customer's primary team) sees and edits the whole CI.
    * Another team sees the CI only if it serves that customer and the CI is shared with it:
        Sharing = "Fully shared"  -> every field except other teams' "(team only)" columns, read-only
        Sharing = "Shared core"   -> only the fields named in "Shared fields" plus identity fields, read-only
        Sharing = "Team only"     -> not visible
      "Shared with teams" empty means: every team serving that customer.
    * A "<Team> details (team only)" column is visible only to members of that team, and that team may fill it
      in on any CI it can see (for example DBA adds database notes to a host owned by Unix SA).
Runbook
    * Each entry belongs to "Owner team" (empty: the customer's primary team) and can be shared with named teams
      ("Shared with teams"); shared teams read it, only the owner team changes it.
Roster
    * Everyone sees every roster; only a lead of the roster's team changes it.
Audit log
    * Administrators see everything; a team lead sees entries made by members of the teams they lead;
      everyone else sees only their own entries.
Administrators belong to no team, so they see no team data (separation of duties).
"""
from __future__ import annotations

import re

OWNER, SHARING, SHARED_WITH, SHARED_FIELDS = "Owner team", "Sharing", "Shared with teams", "Shared fields"
FULLY, CORE, TEAM_ONLY = "Fully shared", "Shared core", "Team only"
TEAM_ONLY_SUFFIX = " details (team only)"
SCOPED_SHEETS = ("Asset_Master", "Runbook")
ALWAYS_VISIBLE = {"Record_ID", "Version", "Created_By", "Created_At", "Updated_By", "Updated_At",
                  "Customer", "Device", "CI_Type", "Environment", "Status"}  # identity of a shared CI


def norm(v) -> str:
    return re.sub(r"\s+", " ", str(v or "")).strip().lower()


def split_list(v) -> list:
    return [x.strip() for x in re.split(r"[,;]", str(v or "")) if x.strip()]


class Access:
    """One person's view of team-owned data, built per request from the team file (msp_org.json)."""

    def __init__(self, org_state: dict, username: str, role: str):
        self.role = role
        self.admin = role == "Admin"
        self.user = norm(username)
        teams = org_state.get("teams", {})
        self.team_names = {tid: t["name"] for tid, t in teams.items()}
        self.mine = set() if self.admin else {
            norm(t["name"]) for tid, t in teams.items()
            if self.user in t.get("members", {}) or self.user in [norm(x) for x in t.get("leads", [])]}
        self.leads = set() if self.admin else {norm(t["name"]) for t in teams.values() if self.user in [norm(x) for x in t.get("leads", [])]}
        self.lead_members = {norm(m) for t in teams.values() if norm(t["name"]) in self.leads for m in t.get("members", {})}
        self.serving, self.primary = {}, {}
        for c in org_state.get("customers", {}).values():
            names = {norm(self.team_names[t]) for t in c.get("teams", []) if t in self.team_names}
            self.serving[norm(c["name"])] = names
            if c.get("primary_team") in self.team_names:
                self.primary[norm(c["name"])] = norm(self.team_names[c["primary_team"]])

    # -- helpers ------------------------------------------------------------------------------------
    @staticmethod
    def keys(columns: list) -> dict:
        """label (lower case) -> column key; works for standard and custom columns alike."""
        return {norm(c["label"]): c["key"] for c in columns}

    def owner_of(self, row: dict, k: dict) -> str:
        own = norm(row.get(k.get(norm(OWNER), ""), ""))
        return own or self.primary.get(norm(row.get("Customer")), "")

    def team_only_columns(self, columns: list) -> dict:
        """column key -> team name (lower case) for every '<Team> details (team only)' column."""
        out = {}
        for c in columns:
            lab = str(c["label"])
            if lab.lower().endswith(TEAM_ONLY_SUFFIX):
                out[c["key"]] = norm(lab[: -len(TEAM_ONLY_SUFFIX)])
        return out

    # -- reading ------------------------------------------------------------------------------------
    def view_row(self, sheet: str, row: dict, columns: list):
        """The row as this person may see it: the full row, a reduced copy, or None (not visible)."""
        if sheet not in SCOPED_SHEETS:
            return row
        k = self.keys(columns)
        owner = self.owner_of(row, k)
        team_cols = self.team_only_columns(columns)
        if owner and owner in self.mine:
            return self._hide_team_cols(dict(row), team_cols)
        cust = norm(row.get("Customer"))
        if sheet == "Runbook":
            shared = {norm(x) for x in split_list(row.get(k.get(norm(SHARED_WITH), ""), ""))}
            if self.mine & shared:
                out = self._hide_team_cols(dict(row), team_cols)
                out["_view_only"], out["_owner"] = True, self.team_label(owner)
                return out
            return None
        sharing = norm(row.get(k.get(norm(SHARING), ""), "")) or norm(CORE)
        if sharing == norm(TEAM_ONLY):
            return None
        shared_with = {norm(x) for x in split_list(row.get(k.get(norm(SHARED_WITH), ""), ""))} or self.serving.get(cust, set())
        if not (self.mine & shared_with & self.serving.get(cust, set())):
            return None
        out = self._hide_team_cols(dict(row), team_cols)
        if sharing != norm(FULLY):
            allowed = set(ALWAYS_VISIBLE) | {k.get(norm(OWNER)), k.get(norm(SHARING))}
            for want in split_list(row.get(k.get(norm(SHARED_FIELDS), ""), "")):
                w = norm(want)
                allowed |= {c["key"] for c in columns if norm(c["label"]) == w or norm(c["label"]).startswith(w + " ")
                            or norm(c["label"]).startswith(w + " /")}
            allowed |= set(team_cols)  # your own team-only column stays (hidden above if it is not yours)
            for key in list(out):
                if key not in allowed and not key.startswith("_"):
                    out[key] = ""
        out["_view_only"], out["_owner"] = True, self.team_label(owner)
        return out

    def _hide_team_cols(self, row: dict, team_cols: dict) -> dict:
        for key, team in team_cols.items():
            if team not in self.mine and key in row:
                row[key] = ""
        return row

    def can_reach_host(self, row: dict, columns: list) -> bool:
        """SSH/RDP to a CI: the owning team, or a team it is shared with (shared core includes remote access)."""
        return self.view_row("Asset_Master", row, columns) is not None

    # -- changing -----------------------------------------------------------------------------------
    def check_change(self, sheet: str, row: dict, columns: list, changed: list) -> None:
        """Raise PermissionError with a plain reason if this person may not make this change."""
        if sheet not in SCOPED_SHEETS:
            return
        k = self.keys(columns)
        owner = self.owner_of(row, k)
        if owner and owner in self.mine:
            return
        if self.view_row(sheet, row, columns) is None:
            raise LookupError("not visible")
        team_cols = self.team_only_columns(columns)
        own_cols = {key for key, team in team_cols.items() if team in self.mine}
        if changed and set(changed) <= own_cols:
            return  # filling in your own team's details on a CI shared with you
        raise PermissionError(f"Only the owning team ({self.team_label(owner)}) can change this. You can view it because it is shared with your team.")

    def default_owner(self, sheet: str, values: dict, columns: list) -> str:
        """Owner team for a new record: as given if it is one of yours, else your team serving that customer."""
        if sheet not in SCOPED_SHEETS:
            return ""
        k = self.keys(columns)
        given = norm(values.get(k.get(norm(OWNER), ""), ""))
        if given and given in self.mine:
            return self.team_label(given)
        if given:
            raise PermissionError(f"You can only create records for your own teams, not for {values.get(k.get(norm(OWNER)))}.")
        if not self.mine:
            raise PermissionError("Records belong to teams, and you are not in one. A team lead can add you to their team.")
        serving = self.serving.get(norm(values.get("Customer")), set())
        pick = sorted(self.mine & serving) or sorted(self.mine)
        return self.team_label(pick[0])

    def team_label(self, name_lower: str) -> str:
        return next((n for n in self.team_names.values() if norm(n) == name_lower), name_lower)

    # -- execution plans ----------------------------------------------------------------------------
    # The owning team runs the plan. "Shared with" teams see all of it. A phase or step given to another team
    # (partial sharing) lets that team see the plan with just its own phases and steps. Other teams edit and
    # complete their own steps; the plan's structure and go/no-go sign-off stay with the owning team.
    STEP_OPS = ("step.update", "step.action")

    def plan_owner(self, task: dict) -> str:
        return norm(task.get("owner_team")) or self.primary.get(norm(task.get("customer")), "")

    @staticmethod
    def step_team(task: dict, phase: dict, step: dict) -> str:
        return norm(step.get("team")) or norm(phase.get("team")) or ""

    def plan_mode(self, task: dict):
        """owner, shared (whole plan), partial (only your phases and steps) or None."""
        if not task:
            return None
        owner = self.plan_owner(task)
        if owner and owner in self.mine:
            return "owner"
        if self.mine & {norm(x) for x in task.get("shared_with") or []}:
            return "shared"
        for ph in task.get("phases", []):
            for st in ph.get("steps", []):
                if self.step_team(task, ph, st) in self.mine:
                    return "partial"
        return None

    def plan_view(self, task: dict):
        mode = self.plan_mode(task)
        if mode is None:
            return None
        out = dict(task)
        out["_access"], out["_owner"] = mode, self.team_label(self.plan_owner(task))
        if mode == "partial":
            phases = []
            for ph in task.get("phases", []):
                steps = [st for st in ph.get("steps", []) if self.step_team(task, ph, st) in self.mine]
                if steps:
                    phases.append(dict(ph, steps=steps))
            out["phases"] = phases
        if mode != "owner":  # mark the steps this person's team may edit
            out["phases"] = [dict(ph, steps=[dict(st, _mine=self.step_team(task, ph, st) in self.mine) for st in ph.get("steps", [])])
                             for ph in out["phases"]]
        return out

    def check_plan_op(self, op: str, task, p: dict) -> None:
        """Raise LookupError (hidden) or PermissionError (visible, not yours) for a plan operation."""
        if op in ("task.create", "template.save", "template.delete"):
            if not self.mine:
                raise PermissionError("Execution plans belong to teams, and you are not in one. A team lead can add you to their team.")
            return
        if task is None:
            return  # the operation itself reports a missing plan
        mode = self.plan_mode(task)
        if mode is None:
            raise LookupError("not visible")
        if mode == "owner":
            if op == "task.share" and p.get("owner_team") and norm(p["owner_team"]) not in self.mine:
                raise PermissionError("You can hand a plan only to one of your own teams.")
            return
        if op in self.STEP_OPS:
            ph = next((x for x in task.get("phases", []) if x.get("id") == p.get("phase_id")), None)
            st = next((x for x in (ph or {}).get("steps", []) if x.get("id") == p.get("step_id")), None)
            if ph is not None and st is not None and self.step_team(task, ph, st) in self.mine:
                return
            raise PermissionError(f"This step belongs to the owning team ({self.team_label(self.plan_owner(task))}). "
                                  "You can update the steps given to your team.")
        if op == "phase.signoff":
            raise PermissionError(f"Go/no-go is signed off by the owning team ({self.team_label(self.plan_owner(task))}).")
        raise PermissionError(f"Only the owning team ({self.team_label(self.plan_owner(task))}) changes this plan. "
                              "You can update the steps given to your team.")

    def plan_default_owner(self, task: dict) -> str:
        given = norm(task.get("owner_team"))
        if given and given in self.mine:
            return self.team_label(given)
        if not self.mine:
            raise PermissionError("Execution plans belong to teams, and you are not in one.")
        pick = sorted(self.mine & self.serving.get(norm(task.get("customer")), set())) or sorted(self.mine)
        return self.team_label(pick[0])

    # -- roster and audit ---------------------------------------------------------------------------
    def can_manage_roster(self, roster_team: str) -> bool:
        """Only a lead of the roster's team. A roster for a team that does not exist stays with any Lead."""
        team = norm(roster_team)
        if team not in {norm(n) for n in self.team_names.values()}:
            return self.role in ("Lead", "Admin")
        return team in self.leads

    def audit_visible(self, entry: dict) -> bool:
        if self.admin:
            return True
        who = norm(entry.get("User"))
        return who == self.user or (bool(self.leads) and who in self.lead_members)
