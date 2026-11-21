#!/usr/bin/env python3
"""Tests for msp_demo.py. Run: python -m unittest test_msp_demo -v"""
import contextlib
import io
import os
import shutil
import tempfile
import unittest

import msp_demo as D
import msp_org as O
import msp_roster as R
import msp_runbook as app


class T01DemoLoader(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp()
        self.cfg = app.Config(host="127.0.0.1", port=0, storage="local", data_dir=self.tmp, refresh_interval=0, quiet=True,
                              seed=False, backup_thread=False, setup_code_file=os.path.join(self.tmp, "c.txt"))

    def tearDown(self):
        shutil.rmtree(self.tmp)

    def run_loader(self):
        with contextlib.redirect_stdout(io.StringIO()):
            return D.load(self.cfg)

    def test_complete_connected_demo_and_safe_to_repeat(self):
        self.run_loader()
        self.run_loader()  # a second run must not duplicate anything
        storage = app.build_storage(self.cfg)
        st = app.DataStore(storage, refresh_interval=0, seed=False)
        users = st.state["users"]["users"]
        self.assertEqual(len(users), len(D.PEOPLE) + 1)
        self.assertEqual({u["role"] for u in users.values()}, {"Admin", "Lead", "Engineer", "Viewer"})
        snap = st.snapshot("admin")
        rows = {k: len(v["rows"]) for k, v in snap["tables"].items()}
        self.assertGreaterEqual(rows["Asset_Master"], 70)
        self.assertGreaterEqual(rows["Runbook"], 30)
        self.assertGreaterEqual(len(snap["plans"]["tasks"]), 8)
        labels = {c["label"] for c in snap["tables"]["Asset_Master"]["columns"]}
        self.assertTrue({"Owner team", "Sharing", "Shared with teams", "Shared fields", "DBA details (team only)"} <= labels)
        org = O.OrgStore(storage).load()[0]
        self.assertEqual(len(org["customers"]), 10)
        for t in org["teams"].values():
            self.assertGreaterEqual(len(t["members"]), 10, t["name"])
        self.assertFalse(any("admin" in t["members"] for t in org["teams"].values()))
        for c in org["customers"].values():
            self.assertIn(c["primary_team"], c["teams"])
        rosters = R.RosterStore(storage).rosters()
        self.assertEqual(sorted(r["team"] for r in rosters), sorted(t["name"] for t in org["teams"].values()))
        for r in rosters:
            self.assertGreaterEqual(len(r["members"]), 10, r["name"])
            self.assertTrue(all(m["user"] in users for m in r["members"]), r["name"])  # every roster member can sign in
            self.assertTrue(any(x["status"] == "pending" for x in r["leaves"]), r["name"])  # something to approve
        # every CI owner team is a real team, and team-only CIs are shared with nobody
        teams = {t["name"] for t in org["teams"].values()}
        cis = snap["tables"]["Asset_Master"]["rows"]
        key = {c["label"]: c["key"] for c in snap["tables"]["Asset_Master"]["columns"]}
        self.assertTrue({r[key["Owner team"]] for r in cis} <= teams)
        self.assertTrue(all(not r[key["Shared with teams"]] for r in cis if r[key["Sharing"]] == "Team only"))
        names = {u["name"] for u in users.values()}
        self.assertTrue({r["Engineer"] for r in snap["tables"]["Incidents"]["rows"]} <= names)
        self.assertTrue({s.get("owner") for t in snap["plans"]["tasks"] for p in t["phases"] for s in p["steps"]} <= names)
        for f in ("09_Demo_Logins_and_Teams.xlsx", "10_Roster_DBA.xlsx"):
            self.assertTrue(os.path.isfile(os.path.join(D.DATA_DIR, f)), f)


class T02LiveDataGuard(unittest.TestCase):
    def test_refuses_a_folder_with_real_users(self):
        tmp = tempfile.mkdtemp()
        try:
            cfg = app.Config(host="127.0.0.1", port=0, storage="local", data_dir=tmp, refresh_interval=0, quiet=True, seed=False,
                             backup_thread=False, setup_code_file=os.path.join(tmp, "c.txt"))
            st = app.DataStore(app.build_storage(cfg), refresh_interval=0, seed=False)
            st.create_user({"username": "admin", "name": "Admin", "role": "Admin", "password": "Real-Adm1n-Pass!"}, "setup", setup=True)
            st.create_user({"username": "real.person", "name": "Real Person", "role": "Engineer", "password": "Real-Pass-12345!"}, "admin")
            with contextlib.redirect_stdout(io.StringIO()):
                with self.assertRaises(D.LiveDataError):
                    D.load(cfg)
                D.load(cfg, force=True)  # only when the person explicitly insists
        finally:
            shutil.rmtree(tmp)


if __name__ == "__main__":
    unittest.main(verbosity=2)
