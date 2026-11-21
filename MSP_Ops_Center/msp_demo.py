#!/usr/bin/env python3
"""MSP Ops Center: load a complete, connected demo (for presentations and trials only).

Creates, in the data folder named in msp_settings.ini:
  * demo users: an administrator, five team leads, engineers and operators (operators are read-only Viewers)
  * teams (Unix SA, Windows, DBA, Boomi, NOC), customers with projects, the teams serving each customer,
    primary teams, members and their project assignments
  * the eight files in Demo_Data (CMDB, runbook, incidents, handovers, bulletins, execution plans, shift log, roster)
  * the roster linked to the users' logins, with approvers, an on-call rota, escalation and India (Karnataka) holidays
  * Demo_Data/09_Demo_Logins_and_Teams.xlsx, a one-page reference for the presenter

Safe to run again: anything that already exists is kept, never duplicated. Every demo user has the same
password (printed at the end). Use it only on a demo copy, never on a live installation.

    python msp_demo.py                 # uses msp_settings.ini next to this file
    python msp_demo.py --config other.ini
"""
from __future__ import annotations

import argparse
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

import msp_runbook as app  # noqa: E402
import msp_org as O  # noqa: E402
import msp_roster as R  # noqa: E402

# Deliberately fixed and published so a presenter can sign in as anyone; refused on live data (see load).
DEMO_PASSWORD = os.environ.get("MSP_DEMO_PASSWORD") or "Demo-Ops-2026!"  # nosec B105 - demo-only accounts
ADMIN = {"user": "admin", "role": "Admin"}
DATA_DIR = os.path.join(HERE, "Demo_Data")

import random  # noqa: E402
import msp_demo_data as G  # noqa: E402

PEOPLE = G.PEOPLE  # username, full name, role, teams (first = home team)



def log(msg):
    print("  " + msg)


class LiveDataError(Exception):
    pass


def load(config, force: bool = False, password=None) -> dict:
    """Load everything into the configured data folder. Returns counts for the summary.

    Refuses when the folder already has people who are not demo users (a live installation), unless force is set."""
    global DEMO_PASSWORD
    if password:
        DEMO_PASSWORD = password
    storage = app.build_storage(config)
    store = app.DataStore(storage, refresh_interval=0, seed=False)
    demo_names = {"admin"} | {p[0] for p in PEOPLE}
    others = sorted(u["username"] for k, u in store.state["users"]["users"].items() if k not in demo_names and u["role"] != "Admin")
    if others and not force:
        raise LiveDataError(f"This data folder already has {len(others)} real user(s), for example {others[0]}. The demo adds "
                            "accounts with a shared, published password, so it is meant for a demo copy only. "
                            "Point msp_settings.ini at an empty folder, or run with --force if you are sure.")
    org = O.OrgStore(storage)
    rosters = R.RosterStore(storage)
    rng = random.Random(G.RNG_SEED)  # nosec B311 - repeatable demo data, not security
    world = G.build(random.Random(G.RNG_SEED))  # nosec B311 - repeatable demo data, not security
    out = {"users_new": 0, "users_kept": 0}

    # 1. users
    if not store.state["users"]["users"]:
        store.create_user({"username": "admin", "name": "Ada Admin", "role": "Admin", "password": DEMO_PASSWORD}, "demo-loader", setup=True)
        log("Created the administrator: admin")
    admin_name = "admin" if store.get_user("admin") else next(u for u in store.state["users"]["users"].values() if u["role"] == "Admin")["username"]
    actor = {"user": admin_name, "role": "Admin"}
    for un, name, role, _teams in PEOPLE:
        if store.get_user(un):
            out["users_kept"] += 1
            continue
        store.create_user({"username": un, "name": name, "role": role, "password": DEMO_PASSWORD}, admin_name)
        out["users_new"] += 1
    log(f"Users: {out['users_new']} created, {out['users_kept']} already there")
    users = {k: app.public_user(u) for k, u in store.state["users"]["users"].items()}

    # 2. teams, customers, projects, members and project assignments
    def setup(state):
        tid = {t["name"]: k for k, t in state["teams"].items()}
        for name, (text, _kind) in G.TEAMS.items():
            if name not in tid:
                tid[name] = O.create_team(state, actor, {"name": name, "description": text,
                                                         "leads": [p[0] for p in G.leads(name)]}, users)["id"]
        cid = {c["name"]: k for k, c in state["customers"].items()}
        for cname, code, _ind, projects, teams, primary in G.CUSTOMERS:
            if cname not in cid:
                cid[cname] = O.create_customer(state, actor, {"name": cname, "code": code})["id"]
            have = {p["name"] for p in state["customers"][cid[cname]]["projects"].values()}
            for pn in projects:
                if pn not in have:
                    O.save_project(state, actor, cid[cname], {"name": pn})
            O.update_customer(state, actor, cid[cname], {"teams": sorted(set(state["customers"][cid[cname]]["teams"]) | {tid[t] for t in teams}),
                                                         "primary_team": state["customers"][cid[cname]].get("primary_team") or tid[primary]})
        for un, _n, _role, teams in PEOPLE:
            for t in teams:
                if un not in state["teams"][tid[t]]["members"]:
                    O.add_member(state, actor, tid[t], un, users)
                m = state["teams"][tid[t]]["members"][un]
                if not m.get("projects"):  # each person works on one to three of the team's projects
                    allowed = sorted(k for c in O.customers_of_team(state, tid[t]) for k in state["customers"][c]["projects"])
                    m["projects"] = rng.sample(allowed, min(len(allowed), rng.randint(1, 3)))
        return {"teams": len(state["teams"]), "customers": len(state["customers"]),
                "projects": sum(len(c["projects"]) for c in state["customers"].values())}
    out.update(org.mutate(actor, "Load demo teams", "demo", setup))
    log(f"Teams: {out['teams']}, customers: {out['customers']}, projects: {out['projects']}")

    # 3. Excel files (written from the generated world, then imported exactly like a person would)
    org_state = org.load()[0]
    team_id = {t["name"]: k for k, t in org_state["teams"].items()}
    for f, team in write_files(store, world):
        if team:  # a team's own shift log and handovers are imported as that team
            store.set_scope(True, team_id[team], [{"id": team_id[team], "name": team}])
        try:
            with open(os.path.join(DATA_DIR, f), "rb") as fh:
                pv = store.import_preview(f, fh.read(), "auto", admin_name)
            store.import_commit(pv["token"], {}, "skip", admin_name)
        finally:
            store.set_scope(False)
        log(f"Imported {f}" + (f" into the {team} team" if team else ""))

    # 3b. plan ownership: each step's team is in its notes ("DBA step"); the shared upgrade is owned by DBA,
    #     with its Unix SA and Boomi steps given to those teams; Acme patching is shared whole with NOC.
    def own(doc, ctx):
        for t in doc["tasks"]:
            teams = [st.get("notes", "").replace(" step", "") for ph in t["phases"] for st in ph["steps"]]
            owner = "DBA" if t.get("change_ref") == "CHG-4101" else (teams[0] if teams else "")
            t["owner_team"] = owner
            t["shared_with"] = ["NOC"] if t.get("change_ref") == "CHG-4102" else []
            for ph in t["phases"]:
                for st in ph["steps"]:
                    team = st.get("notes", "").replace(" step", "")
                    st["team"] = team if team and team != owner else ""
        ctx["details"] = "demo plan ownership"
    store.mutate("plans", own, admin_name, "Execution plans", "Demo ownership", "demo")
    log("Execution plans: owners and shared steps set")

    # 4. one roster per team, linked to logins
    out["rosters"] = []
    have = {r["name"] for r in rosters.rosters()}
    for team, (_text, kind) in G.TEAMS.items():
        name = f"{team} roster"
        if name in have:
            log(f"Roster '{name}' already there")
            continue
        out["rosters"].append(build_roster(rosters, actor, team, kind, name, rng))
        log(f"Roster '{name}' created with {len(G.home_members(team))} people")
    for r in rosters.rosters():  # each team's roster as a spreadsheet too (by shift and by person)
        if r["name"].endswith(" roster") and r["team"] in G.TEAMS:
            with open(os.path.join(DATA_DIR, f"10_Roster_{r['team'].replace(' ', '_')}.xlsx"), "wb") as fh:
                fh.write(R.export_excel(r, "2026-11-01", "2026-12-31"))
    out["sheet"] = write_logins_sheet()
    return out


def build_roster(rosters, actor, team, kind, name, rng):
    start = "2026-10-05"
    r = R.default_roster(team, name, start, "three_shift" if kind == "24x7" else "standard")
    if kind == "business":  # Boomi: weekday day and evening shifts, weekends off, on-call covers the rest
        r["shifts"] = [{"code": "D", "name": "Day", "start": "09:00", "end": "18:00", "days": [0, 1, 2, 3, 4], "min_staff": 3,
                        "max_staff": None, "color": "#43C08F", "log_shift": "Morning"},
                       {"code": "E", "name": "Evening", "start": "14:00", "end": "23:00", "days": [0, 1, 2, 3, 4], "min_staff": 2,
                        "max_staff": None, "color": "#7AA7FF", "log_shift": "Afternoon"}]
        r["cycle"] = {"type": "custom", "length_days": 14}
        r["rotation"] = {"sequence": ["D"] * 5 + ["OFF"] * 2 + ["E"] * 5 + ["OFF"] * 2}
        r["rules"]["forbidden_sequences"] = []
        offsets = [0, 7]
    else:
        for s in r["shifts"]:
            s["min_staff"] = 2
        offsets = [0, 1, 6, 11, 16, 3, 8, 13, 18, 20]
    people = G.home_members(team)
    lead_users = [p[0] for p in G.leads(team)]
    r["members"] = [{"id": R.new_id("mem_"), "name": p[1], "role": p[2], "user": p[0], "offset": offsets[i % len(offsets)]}
                    for i, p in enumerate(people)]
    r["approvers"] = lead_users
    r["rules"]["max_consecutive_days"] = None
    r["holidays"] = {"region": "IN", "subdivision": "KA", "custom": [{"date": "2026-11-27", "name": "Company foundation day"}], "ignore": []}
    eng_ids = [m["id"] for m in r["members"] if m["role"] == "Engineer"]
    lead_ids = [m["id"] for m in r["members"] if m["role"] == "Lead"]
    r["oncall"] = {"enabled": True, "roles": ["Primary", "Secondary"], "rotation_days": 7, "start_date": "2026-11-02",
                   "order": {"Primary": eng_ids, "Secondary": lead_ids or eng_ids[::-1]},
                   "escalation": [{"level": 1, "name": f"{G.leads(team)[0][1]}, {team} lead", "contact": f"ext {4100 + list(G.TEAMS).index(team)}"},
                                  {"level": 2, "name": "Service delivery manager", "contact": "ext 4000"}], "overrides": {}}
    r = rosters.create(actor, r)
    # some approved leave, one leave and one swap waiting, so approvals and consent can be shown
    engs = [m for m in r["members"] if m["role"] == "Engineer"]
    for m in rng.sample(engs, 2):
        d = R.parse_date("2026-11-16") + R.dt.timedelta(days=rng.randint(0, 25))
        rosters.request_leave({"user": lead_users[0], "role": "Lead"}, r["id"], m["id"], d, d + R.dt.timedelta(days=rng.randint(0, 2)), "Annual", "Planned leave")
    asker = engs[0]
    try:
        rosters.request_leave({"user": asker["user"], "role": "Engineer"}, r["id"], asker["id"], "2026-12-21", "2026-12-23", "Annual", "Family visit")
        rosters.request_swap({"user": engs[1]["user"], "role": "Engineer"}, r["id"], engs[1]["id"], "2026-11-20", engs[2]["id"], None, "swap", "Doctor appointment")
    except R.RosterError:
        pass  # a clash with generated leave is fine: the demo still has other requests
    return r["id"]


def write_files(store, world) -> list:
    """Write 01-07 from the app's own templates (exact headers), adding the CMDB sharing columns."""
    import io
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment
    specs = [("assets", "01_CMDB_Configuration_Items.xlsx", "CMDB: configuration items, with owner team and sharing model - DEMO DATA", world["cis"], G.SHARE_COLUMNS),
             ("runbook", "02_Runbook.xlsx", "Runbook - DEMO DATA", world["runbook"], []),
             ("incidents", "03_Incidents.xlsx", "Incidents (ART) - DEMO DATA", world["incidents"], []),
             ("bulletins", "05_Bulletins.xlsx", "Bulletins - DEMO DATA", world["bulletins"], []),
             ("plans", "06_Execution_Plans.xlsx", "Execution plans - DEMO DATA", world["plans"], []),
             ]
    for team in G.TEAMS:  # each team's own handovers and shift log
        tag = team.replace(" ", "_")
        specs.append(("handovers", f"04_Shift_Handovers_{tag}.xlsx", f"Shift handovers: {team} team only - DEMO DATA",
                      [r for r in world["handovers"] if r["Team"] == team], [], team))
        specs.append(("shiftlog", f"07_Shift_Log_{tag}.xlsx", f"Shift log: {team} team only - DEMO DATA",
                      [{k: v for k, v in r.items() if k != "Team"} for r in world["shiftlog"] if r["Team"] == team], [], team))
    for old in ("04_Shift_Handovers.xlsx", "07_Shift_Log.xlsx"):
        if os.path.exists(os.path.join(DATA_DIR, old)):
            os.remove(os.path.join(DATA_DIR, old))
    names = []
    for spec in specs:
        module, fname, title, rows, extra = spec[:5]
        team = spec[5] if len(spec) > 5 else None
        data, _, _ = store.template_file(module, "xlsx")
        wb = openpyxl.load_workbook(io.BytesIO(data))
        ws = wb.worksheets[0]
        head = [c.value for c in ws[2]]
        extra = [h for h in extra if h not in head]  # ownership columns are standard now; only team-only columns are extra
        head += extra
        unknown = sorted({k for r in rows for k in r} - set(head))
        if unknown:
            raise ValueError(f"{fname}: unknown columns {unknown}")
        for j, h in enumerate(extra, start=ws.max_column + 1):
            ws.cell(row=2, column=j, value=h)
        ws.delete_rows(3, ws.max_row)
        ws["A1"] = title
        for r in rows:
            ws.append([r.get(h) for h in head])
        for row in ws.iter_rows(min_row=1):
            for c in row:
                c.font = Font(name="Arial", size=12 if c.row == 1 else 10, bold=c.row <= 2)
                if c.row > 2:
                    c.alignment = Alignment(vertical="top")
                    if c.row % 2 == 0:
                        c.fill = PatternFill("solid", fgColor="EEF4F7")
        for i, h in enumerate(head, start=1):
            ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = max(10, min(len(str(h)) + 4, 40))
        ws.freeze_panes = "A3"
        wb.save(os.path.join(DATA_DIR, fname))
        names.append((fname, team))
    return names


def write_logins_sheet() -> str:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Demo logins"
    bold, head = Font(name="Arial", bold=True), PatternFill("solid", fgColor="DDEBF7")
    ws.append(["DEMO ONLY: every account below shares the same password. Never use these accounts on a live installation."])
    ws["A1"].font = Font(name="Arial", bold=True, color="B42318")
    ws.append([f"Password for every account: {DEMO_PASSWORD}"])
    ws["A2"].font = bold
    ws.append([])
    ws.append(["User name", "Name", "Role", "Teams", "What they do", "Try this in the demo"])
    for c in ws[4]:
        c.font, c.fill = bold, head
    tips = {"admin": "Teams page: all teams, customers and primary teams. Admin: users and settings.",
            "priya.sharma": "DBA lead: My work, DBA roster approvals, the shared Harbourline upgrade plan.",
            "anjali.nair": "NOC lead and Windows member: Teams (add a member), NOC roster, on-call.",
            "rahul.verma": "Engineer in two teams (Windows and NOC).",
            "vikram.rao": "Boomi lead: business-hours roster with evening shift.",
            "tom.berg": "Read-only operator in Unix SA."}
    ws.append(["admin", "Ada Admin", "Admin", "(not a team member)", "Sets up teams, customers and users", tips["admin"]])
    for un, name, role, teams in PEOPLE:
        ws.append([un, name, role, ", ".join(teams) or "(none)", G.ROLE_NOTE[role], tips.get(un, "")])
    ws.append([])
    ws.append(["Customer", "Code", "Industry", "Projects", "Teams serving it", "Primary team (owns core CI data)"])
    for c in ws[ws.max_row]:
        c.font, c.fill = bold, head
    for cname, code, industry, projects, teams, primary in G.CUSTOMERS:
        ws.append([cname, code, industry, ", ".join(projects), ", ".join(teams), primary])
    for row in ws.iter_rows(min_row=4):
        for c in row:
            if c.font != bold:
                c.font = Font(name="Arial")
            c.alignment = Alignment(vertical="top", wrap_text=True)
    for col, w in zip("ABCDEF", (16, 18, 10, 26, 36, 60)):
        ws.column_dimensions[col].width = w
    path = os.path.join(DATA_DIR, "09_Demo_Logins_and_Teams.xlsx")
    wb.save(path)
    return path


def main(argv=None):
    p = argparse.ArgumentParser(description="Load the MSP Ops Center demo (users, teams, customers, data, roster).")
    p.add_argument("--config", default="")
    p.add_argument("--force", action="store_true", help="load even if the folder already has real users")
    p.add_argument("--password", default="", help="password for every demo account (default: the published demo password)")
    a, rest = p.parse_known_args(argv)
    config = app.config_from_args((["--config", a.config] if a.config else []) + rest)
    print("MSP Ops Center demo loader")
    print(f"  Data folder: {app.build_storage(config).describe()['location']}")
    try:
        res = load(config, force=a.force, password=a.password or None)
    except LiveDataError as exc:
        print(f"\nStopped: {exc}")
        return 3
    except (app.ApiError, O.OrgError, R.RosterError) as exc:
        print(f"\nThe demo could not be loaded: {exc.message}")
        return 2
    print(f"\nDone. Every demo account uses the password: {DEMO_PASSWORD}")
    print("  Try: admin (setup), priya.sharma (DBA lead), anjali.nair (NOC lead), rahul.verma (engineer), tom.berg (operator)")
    print(f"  A list of all logins and teams: {res['sheet']}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
