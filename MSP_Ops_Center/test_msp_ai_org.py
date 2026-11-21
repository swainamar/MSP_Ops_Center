#!/usr/bin/env python3
"""Tests for msp_ai_org.py and msp_ai_help.py, on the demo world behind a real server.
Run: python -m unittest test_msp_ai_org -v"""
import contextlib
import http.cookiejar
import io
import json
import os
import shutil
import tempfile
import threading
import unittest
import urllib.request

import msp_ai_help as H


class T01Guides(unittest.TestCase):
    def test_every_guide_is_reachable_by_its_title(self):
        for g in H.GUIDES:
            found = H.best_guide("How do I " + g[0].lower() + "?")
            self.assertIsNotNone(found, g[0])
            self.assertEqual(found[0], g[0])

    def test_not_a_how_to_question(self):
        self.assertIsNone(H.best_guide("acme-lnx-db01 memory"))


class T02AskOverTheApi(unittest.TestCase):
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

    def ask(self, user, q):
        op = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(http.cookiejar.CookieJar()))
        r = urllib.request.Request(self.base + "/api/auth/login", data=json.dumps({"username": user, "password": "Demo-Ops-2026!"}).encode(),
                                   method="POST", headers={"Content-Type": "application/json"})
        csrf = json.loads(op.open(r).read())["csrf"]
        r = urllib.request.Request(self.base + "/api/ai/ask", data=json.dumps({"question": q}).encode(), method="POST",
                                   headers={"Content-Type": "application/json", "X-CSRF-Token": csrf})
        return json.loads(op.open(r).read())["answer"]

    def test_the_reported_question(self):
        a = self.ask("vikram.rao", "Who are the team members are part of Boomi team?")
        self.assertIn("The Boomi team has 11 people", a["text"])
        self.assertIn("Omar Khan (omar.khan)", a["text"])
        self.assertIn("Operators (read-only):", a["text"])
        self.assertEqual(a["sources"][0]["label"], "Team")
        self.assertNotIn("Harbourline DB upgrade", a["text"])  # the unrelated bulletin it used to show

    def test_isolation_of_member_lists(self):
        self.assertIn("The Boomi team has", self.ask("admin", "who is in the Boomi team")["text"])  # admins see every team
        a = self.ask("sneha.iyer", "who is in the Boomi team")
        self.assertFalse(a["found"])
        self.assertIn("You are not in the Boomi team", a["text"])
        self.assertNotIn("omar.khan", a["text"])

    def test_leads_my_teams_and_customers(self):
        self.assertIn("NOC is led by Anjali Nair, Harish Kumar", self.ask("rahul.verma", "who leads the NOC team?")["text"])
        a = self.ask("rahul.verma", "which teams am I in?")["text"]
        self.assertIn("NOC", a)
        self.assertIn("Windows", a)
        a = self.ask("priya.sharma", "which teams serve Harbourline Retail?")["text"]
        self.assertIn("Primary team (owns the core CI data): Unix SA", a)
        self.assertIn("Services: Store systems, Online shop, Integration", a)

    def test_how_to_beats_my_team(self):
        a = self.ask("anjali.nair", "how to add a new engineer to my team")
        self.assertEqual(a["sources"][0]["label"], "How-to")
        self.assertIn("Add a person to your team", a["text"])
        self.assertIn("Add member", a["text"])

    def test_search_index_is_per_person(self):
        """Since isolation, two people must never share a cached index (it would leak records)."""
        sa = self.ask("suresh.babu", "acme-db-test01")       # DBA's team-only test database
        dba = self.ask("sneha.iyer", "acme-db-test01")
        self.assertFalse(any("acme-db-test01" == s["title"] for s in sa["sources"]))
        self.assertTrue(any("acme-db-test01" == s["title"] for s in dba["sources"]))
        sa_again = self.ask("suresh.babu", "acme-db-test01")
        self.assertFalse(any("acme-db-test01" == s["title"] for s in sa_again["sources"]))


if __name__ == "__main__":
    unittest.main(verbosity=2)
