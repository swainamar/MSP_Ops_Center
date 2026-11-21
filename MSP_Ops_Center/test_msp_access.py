#!/usr/bin/env python3
"""Tests for msp_access.py (isolation bundle A). Run: python -m unittest test_msp_access -v

Part 1 checks the rules as plain functions. Part 2 loads the demo into a temporary folder, starts the server,
and checks every rule over the API as different people."""
import contextlib
import io
import json
import os
import shutil
import tempfile
import threading
import unittest
import urllib.error
import urllib.request
import http.cookiejar

import msp_access as A

ORG = {"teams": {"t1": {"name": "Unix SA", "leads": ["arun"], "members": {"arun": {}, "suresh": {}, "chen": {}}},
                 "t2": {"name": "DBA", "leads": ["priya"], "members": {"priya": {}, "sneha": {}, "chen": {}}},
                 "t3": {"name": "Boomi", "leads": ["vikram"], "members": {"vikram": {}, "omar": {}}}},
       "customers": {"c1": {"name": "Acme Corp", "teams": ["t1", "t2", "t3"], "primary_team": "t1"},
                     "c2": {"name": "Initech", "teams": ["t1"], "primary_team": "t1"}}}
COLS = [{"key": k, "label": l} for k, l in [("Record_ID", "Record ID"), ("Customer", "Customer"), ("Device", "CI name"),
        ("CI_Type", "CI type"), ("Hostname_FQDN", "Hostname / FQDN"), ("IP_Address", "IP address"), ("OS", "OS"),
        ("Memory", "Memory"), ("Owner_Team", "Owner team"), ("Sharing", "Sharing"), ("Shared_With", "Shared with teams"),
        ("Shared_Fields", "Shared fields"), ("c_sa", "Unix SA details (team only)"), ("c_dba", "DBA details (team only)")]]
CORE = "CI name, Hostname, IP address, Remote access (SSH/RDP)"


def ci(**kw):
    row = {"Record_ID": "R1", "Customer": "Acme Corp", "Device": "acme-db01", "CI_Type": "Server", "Hostname_FQDN": "acme-db01.example",
           "IP_Address": "10.0.0.5", "OS": "Oracle Linux", "Memory": "512 GB", "Owner_Team": "Unix SA", "Sharing": "Shared core",
           "Shared_With": "DBA", "Shared_Fields": CORE, "c_sa": "multipath, HugePages", "c_dba": "SID ACMEPRD"}
    row.update(kw)
    return row


def who(name, role="Engineer"):
    return A.Access(ORG, name, role)


class T01Cmdb(unittest.TestCase):
    def test_owner_sees_everything_except_other_teams_notes(self):
        v = who("suresh").view_row("Asset_Master", ci(), COLS)
        self.assertEqual((v["OS"], v["Memory"], v["c_sa"]), ("Oracle Linux", "512 GB", "multipath, HugePages"))
        self.assertEqual(v["c_dba"], "")  # DBA's own notes stay with DBA
        self.assertNotIn("_view_only", v)

    def test_shared_core_shows_only_the_shared_fields(self):
        v = who("sneha").view_row("Asset_Master", ci(), COLS)
        self.assertTrue(v["_view_only"])
        self.assertEqual((v["Device"], v["Hostname_FQDN"], v["IP_Address"]), ("acme-db01", "acme-db01.example", "10.0.0.5"))
        self.assertEqual((v["OS"], v["Memory"], v["c_sa"]), ("", "", ""))
        self.assertEqual(v["c_dba"], "SID ACMEPRD")  # their own team-only column

    def test_fully_shared_and_team_only(self):
        v = who("sneha").view_row("Asset_Master", ci(Sharing="Fully shared"), COLS)
        self.assertEqual((v["OS"], v["c_sa"]), ("Oracle Linux", ""))
        self.assertIsNone(who("sneha").view_row("Asset_Master", ci(Sharing="Team only"), COLS))

    def test_not_shared_with_you_or_customer_not_served(self):
        self.assertIsNone(who("omar").view_row("Asset_Master", ci(), COLS))  # shared with DBA only
        self.assertIsNotNone(who("omar").view_row("Asset_Master", ci(Shared_With=""), COLS))  # empty = every team serving Acme
        self.assertIsNone(who("sneha").view_row("Asset_Master", ci(Customer="Initech", Shared_With="DBA"), COLS))  # DBA does not serve Initech

    def test_empty_owner_means_primary_team(self):
        self.assertNotIn("_view_only", who("arun").view_row("Asset_Master", ci(Owner_Team=""), COLS))

    def test_admin_sees_nothing(self):
        self.assertIsNone(who("ada", "Admin").view_row("Asset_Master", ci(Shared_With=""), COLS))

    def test_changes(self):
        who("suresh").check_change("Asset_Master", ci(), COLS, ["OS"])
        with self.assertRaises(PermissionError):
            who("sneha").check_change("Asset_Master", ci(), COLS, ["OS"])
        who("sneha").check_change("Asset_Master", ci(), COLS, ["c_dba"])  # DBA fills in its own notes on a shared CI
        with self.assertRaises(LookupError):
            who("omar").check_change("Asset_Master", ci(), COLS, ["OS"])  # not visible at all
        with self.assertRaises(PermissionError):
            who("sneha").check_change("Asset_Master", ci(), COLS, None)  # delete or clone

    def test_new_records(self):
        self.assertEqual(who("sneha").default_owner("Asset_Master", {"Customer": "Acme Corp"}, COLS), "DBA")
        self.assertEqual(who("chen").default_owner("Asset_Master", {"Customer": "Initech"}, COLS), "Unix SA")  # the team that serves it
        with self.assertRaises(PermissionError):
            who("sneha").default_owner("Asset_Master", {"Customer": "Acme Corp", "Owner_Team": "Unix SA"}, COLS)
        with self.assertRaises(PermissionError):
            who("ada", "Admin").default_owner("Asset_Master", {"Customer": "Acme Corp"}, COLS)


class T02RunbookRosterAudit(unittest.TestCase):
    RB = [{"key": k, "label": l} for k, l in [("Record_ID", "Record ID"), ("Customer", "Customer"), ("Title", "Title"),
                                              ("Owner_Team", "Owner team"), ("Shared_With", "Shared with teams")]]

    def test_runbook(self):
        row = {"Record_ID": "R9", "Customer": "Acme Corp", "Title": "Fix", "Owner_Team": "DBA", "Shared_With": ""}
        self.assertIsNone(who("suresh").view_row("Runbook", row, self.RB))
        v = who("suresh").view_row("Runbook", dict(row, Shared_With="Unix SA"), self.RB)
        self.assertTrue(v["_view_only"])
        self.assertEqual(v["Title"], "Fix")  # shared runbook entries are read in full
        with self.assertRaises(PermissionError):
            who("suresh").check_change("Runbook", dict(row, Shared_With="Unix SA"), self.RB, ["Title"])

    def test_roster(self):
        self.assertTrue(who("priya", "Lead").can_manage_roster("DBA"))
        self.assertFalse(who("arun", "Lead").can_manage_roster("DBA"))
        self.assertFalse(who("sneha").can_manage_roster("DBA"))
        self.assertTrue(who("arun", "Lead").can_manage_roster("Old roster team"))  # not a team: any Lead, as before

    def test_audit(self):
        e = lambda u: {"User": u}  # noqa: E731
        self.assertTrue(who("ada", "Admin").audit_visible(e("anyone")))
        self.assertTrue(who("priya", "Lead").audit_visible(e("sneha")))
        self.assertTrue(who("priya", "Lead").audit_visible(e("chen")))  # chen is in DBA too
        self.assertFalse(who("priya", "Lead").audit_visible(e("omar")))
        self.assertTrue(who("sneha").audit_visible(e("sneha")))
        self.assertFalse(who("sneha").audit_visible(e("priya")))


def plan(**kw):
    t = {"id": "t1", "customer": "Acme Corp", "owner_team": "DBA", "shared_with": [], "name": "DB upgrade",
         "phases": [{"id": "p1", "name": "Pre", "team": "", "steps": [{"id": "s1", "team": ""}, {"id": "s2", "team": "Boomi"}]},
                    {"id": "p2", "name": "Upgrade", "team": "Unix SA", "steps": [{"id": "s3", "team": ""}]}]}
    t.update(kw)
    return t


class T04Plans(unittest.TestCase):
    def test_modes(self):
        self.assertEqual(who("sneha").plan_mode(plan()), "owner")
        self.assertEqual(who("omar").plan_mode(plan()), "partial")     # one step given to Boomi
        self.assertEqual(who("suresh").plan_mode(plan()), "partial")   # a phase given to Unix SA
        self.assertIsNone(who("omar").plan_mode(plan(phases=[])))
        self.assertEqual(who("omar").plan_mode(plan(shared_with=["Boomi"])), "shared")
        self.assertIsNone(who("ada", "Admin").plan_mode(plan()))

    def test_partial_view_shows_only_your_steps(self):
        v = who("omar").plan_view(plan())
        self.assertEqual([(ph["name"], [st["id"] for st in ph["steps"]]) for ph in v["phases"]], [("Pre", ["s2"])])
        self.assertTrue(v["phases"][0]["steps"][0]["_mine"])
        full = who("omar").plan_view(plan(shared_with=["Boomi"]))
        self.assertEqual(sum(len(ph["steps"]) for ph in full["phases"]), 3)  # shared whole: everything, marked
        self.assertEqual([st["_mine"] for ph in full["phases"] for st in ph["steps"]], [False, True, False])

    def test_operations(self):
        who("omar").check_plan_op("step.action", plan(), {"phase_id": "p1", "step_id": "s2"})   # his own step
        with self.assertRaises(PermissionError):
            who("omar").check_plan_op("step.update", plan(), {"phase_id": "p1", "step_id": "s1"})  # the owner's step
        for op in ("task.update", "phase.add", "step.delete", "phase.signoff", "task.share", "task.delete"):
            with self.assertRaises(PermissionError, msg=op):
                who("omar").check_plan_op(op, plan(), {"phase_id": "p1", "step_id": "s2"})
        with self.assertRaises(LookupError):
            who("arun").check_plan_op("task.update", plan(phases=[]), {})  # hidden
        who("sneha").check_plan_op("phase.signoff", plan(), {"phase_id": "p1"})  # the owner signs off
        with self.assertRaises(PermissionError):
            who("sneha").check_plan_op("task.share", plan(), {"owner_team": "Boomi"})  # only to your own teams

    def test_new_plan_owner(self):
        self.assertEqual(who("sneha").plan_default_owner({"customer": "Acme Corp"}), "DBA")
        self.assertEqual(who("chen").plan_default_owner({"customer": "Initech"}), "Unix SA")


class T03OverTheApi(unittest.TestCase):
    """The demo world (5 teams, 10 customers, 75 CIs) behind a real server."""

    @classmethod
    def setUpClass(cls):
        import msp_runbook as app
        import msp_demo as D
        cls.tmp = tempfile.mkdtemp()
        cfg = app.Config(host="127.0.0.1", port=0, storage="local", data_dir=cls.tmp, refresh_interval=0, quiet=True,
                         seed=False, backup_thread=False, setup_code_file=os.path.join(cls.tmp, "c.txt"))
        with contextlib.redirect_stdout(io.StringIO()):
            D.load(cfg)
        cls.srv = app.create_server(cfg)
        threading.Thread(target=cls.srv.serve_forever, daemon=True).start()
        cls.base = f"http://127.0.0.1:{cls.srv.server_address[1]}"

    @classmethod
    def tearDownClass(cls):
        cls.srv.shutdown()
        cls.srv.server_close()
        shutil.rmtree(cls.tmp)

    def login(self, user):
        op = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(http.cookiejar.CookieJar()))
        r = urllib.request.Request(self.base + "/api/auth/login", data=json.dumps({"username": user, "password": "Demo-Ops-2026!"}).encode(),
                                   method="POST", headers={"Content-Type": "application/json"})
        return op, json.loads(op.open(r).read())["csrf"]

    def call(self, s, method, path, body=None):
        op, csrf = s
        r = urllib.request.Request(self.base + path, data=json.dumps(body).encode() if body is not None else None, method=method,
                                   headers={"Content-Type": "application/json", "X-CSRF-Token": csrf})
        try:
            return 200, json.loads(op.open(r).read())
        except urllib.error.HTTPError as e:
            return e.code, json.loads(e.read())

    def cmdb(self, s):
        return {r["Device"]: r for r in self.call(s, "GET", "/api/state")[1]["tables"]["Asset_Master"]["rows"]}

    def test_cmdb_views_differ_by_team(self):
        sa, dba, boomi = self.login("suresh.babu"), self.login("sneha.iyer"), self.login("omar.khan")
        a, d, b = self.cmdb(sa), self.cmdb(dba), self.cmdb(boomi)
        host = a["acme-lnx-db01"]
        self.assertTrue(host["Memory"] and "_view_only" not in host)       # Unix SA owns it: everything
        self.assertEqual(d["acme-lnx-db01"]["Memory"], "")                 # DBA: shared core only
        self.assertEqual(d["acme-lnx-db01"]["IP_Address"], host["IP_Address"])
        self.assertIn("acme-db-ora01", d)                                  # DBA's own database
        self.assertNotIn("acme-db-test01", a)                               # team-only test database: DBA only
        self.assertNotIn("bpl-boomi01", a)                                  # Bluepeak: only Boomi serves it
        self.assertIn("bpl-boomi01", b)
        self.assertEqual(self.cmdb(self.login("admin")), {})                # admin: no team data

    def test_changes_are_for_the_owner(self):
        sa, dba = self.login("suresh.babu"), self.login("sneha.iyer")
        rid = self.cmdb(sa)["acme-lnx-db01"]["Record_ID"]
        st, d = self.call(dba, "PUT", f"/api/table/Asset_Master/{rid}", {"Memory": "1 TB"})
        self.assertEqual(st, 403, d)
        self.assertIn("Only the owning team (Unix SA)", d["error"])
        st, d = self.call(sa, "PUT", f"/api/table/Asset_Master/{rid}", {"Notes": "", "Memory": "768 GB"})
        self.assertEqual(st, 200, d)
        hidden = self.cmdb(self.login("priya.sharma"))["acme-db-test01"]["Record_ID"]
        st, _ = self.call(sa, "GET", f"/api/history?key=table/Asset_Master/{hidden}")
        self.assertEqual(st, 404)  # not even the history of a CI you cannot see

    def test_audit_log_scope(self):
        st, d = self.call(self.login("sneha.iyer"), "GET", "/api/state")
        self.assertTrue(all(a["User"] == "sneha.iyer" for a in d["audit"]))
        st, d = self.call(self.login("admin"), "GET", "/api/audit")
        self.assertTrue(len({a["User"] for a in d["audit"]}) >= 1)

    def test_roster_changes_by_the_owning_lead_only(self):
        st, d = self.call(self.login("arun.kumar"), "GET", "/api/roster")
        dba_roster = next(r for r in d["rosters"] if r["team"] == "DBA")
        self.assertEqual(len(d["rosters"]), 5)  # everyone sees every roster
        st, d = self.call(self.login("arun.kumar"), "PUT", f"/api/roster/{dba_roster['id']}", {"changes": {"name": "Hijacked"}})
        self.assertEqual(st, 403, d)
        st, d = self.call(self.login("priya.sharma"), "PUT", f"/api/roster/{dba_roster['id']}", {"changes": {"name": "DBA roster"}})
        self.assertEqual(st, 200, d)


class T05PlansOverTheApi(T03OverTheApi):
    """Execution plan isolation on the demo (CHG-4101 owned by DBA, with Unix SA and Boomi steps)."""

    def plans(self, s):
        return {t["change_ref"]: t for t in self.call(s, "GET", "/api/state")[1]["plans"]["tasks"]}

    def test_cmdb_views_differ_by_team(self):  # covered by T03; not repeated here
        pass

    test_changes_are_for_the_owner = test_audit_log_scope = test_roster_changes_by_the_owning_lead_only = test_cmdb_views_differ_by_team

    def test_plan_views_and_rights(self):
        sneha, omar, karthik = self.login("sneha.iyer"), self.login("omar.khan"), self.login("karthik.reddy")
        mine = self.plans(sneha)["CHG-4101"]
        self.assertEqual(mine["_access"], "owner")
        boomi = self.plans(omar)["CHG-4101"]
        self.assertEqual(boomi["_access"], "partial")
        steps = [(ph["id"], st) for ph in boomi["phases"] for st in ph["steps"]]
        self.assertTrue(steps and all(st["_mine"] for _, st in steps))
        self.assertLess(len(steps), sum(len(ph["steps"]) for ph in mine["phases"]))
        st, d = self.call(omar, "POST", "/api/plans/op", {"op": "step.update", "task_id": boomi["id"], "phase_id": steps[0][0],
                                                          "step_id": steps[0][1]["id"], "step": {"notes": "Listeners stopped at 22:05"}})
        self.assertEqual(st, 200, d)
        st, d = self.call(omar, "POST", "/api/plans/op", {"op": "task.update", "task_id": boomi["id"], "task": {"name": "x"}})
        self.assertEqual(st, 403)
        self.assertIn("Only the owning team (DBA)", d["error"])
        self.assertNotIn("CHG-4101", self.plans(karthik))
        st, _ = self.call(karthik, "POST", "/api/plans/op", {"op": "task.update", "task_id": boomi["id"], "task": {"name": "x"}})
        self.assertEqual(st, 404)
        st, _ = self.call(omar, "GET", f"/api/history?key=plan/{boomi['id']}")
        self.assertEqual(st, 404)  # partial: the full history would show the other teams' steps
        self.assertEqual(self.plans(self.login("dinesh.yadav"))["CHG-4102"]["_access"], "shared")
        self.assertEqual(self.plans(self.login("admin")), {})

    def test_sharing_by_the_owner(self):
        priya = self.login("priya.sharma")
        t = self.plans(priya)["CHG-4103"]
        st, d = self.call(priya, "POST", "/api/plans/op", {"op": "task.share", "task_id": t["id"], "shared_with": ["NOC"]})
        self.assertEqual(st, 200, d)
        self.assertEqual(self.plans(self.login("dinesh.yadav"))["CHG-4103"]["_access"], "shared")
        st, d = self.call(priya, "POST", "/api/plans/op", {"op": "task.share", "task_id": t["id"], "owner_team": "Boomi"})
        self.assertEqual(st, 403)  # only to your own teams
        st, d = self.call(priya, "POST", "/api/plans/op", {"op": "task.create", "task": {"name": "New DBA plan", "customer": "Globex", "date": "2026-12-01"}})
        self.assertEqual(st, 200, d)
        self.assertEqual(self.plans(priya).get("", {}).get("_access", "owner"), "owner")
        newest = [x for x in self.call(priya, "GET", "/api/state")[1]["plans"]["tasks"] if x["name"] == "New DBA plan"][0]
        self.assertEqual(newest["owner_team"], "DBA")


if __name__ == "__main__":
    unittest.main(verbosity=2)
