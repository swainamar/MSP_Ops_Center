#!/usr/bin/env python3
"""Tests for msp_org.py (teams, customers, projects). Run: python -m unittest test_msp_org -v"""
import shutil
import tempfile
import unittest

import msp_org as O
import msp_roster as R  # its FileStorage has the same interface as the app's storage

ADMIN = {"user": "ada", "role": "Admin"}
LOU = {"user": "lou.lead", "role": "Lead"}
KIM = {"user": "kim.lead", "role": "Lead"}
ELI = {"user": "eli.eng", "role": "Engineer"}
USERS = {"ada": {"username": "ada", "name": "Ada", "role": "Admin"}, "lou.lead": {"username": "lou.lead", "name": "Lou", "role": "Lead"},
         "kim.lead": {"username": "kim.lead", "name": "Kim", "role": "Lead"}, "eli.eng": {"username": "eli.eng", "name": "Eli", "role": "Engineer"},
         "mia.s": {"username": "mia.s", "name": "Mia", "role": "Engineer"}}


def setup_state():
    st = O.empty_state()
    dba = O.create_team(st, ADMIN, {"name": "DBA", "leads": ["lou.lead"]}, USERS)["id"]
    sa = O.create_team(st, ADMIN, {"name": "Unix SA", "leads": ["kim.lead"]}, USERS)["id"]
    acme = O.create_customer(st, ADMIN, {"name": "Acme Corp", "code": "acme"})["id"]
    p1 = O.save_project(st, ADMIN, acme, {"name": "ERP support"})["id"]
    ini = O.create_customer(st, ADMIN, {"name": "Initech"})["id"]
    p2 = O.save_project(st, ADMIN, ini, {"name": "Core banking"})["id"]
    O.update_customer(st, ADMIN, acme, {"teams": [dba, sa], "primary_team": sa})
    O.update_customer(st, ADMIN, ini, {"teams": [sa], "primary_team": sa})
    return st, dba, sa, acme, ini, p1, p2


class T01Admin(unittest.TestCase):
    def test_only_admin_creates_teams_customers_projects(self):
        st, dba, sa, acme, *_ = setup_state()
        for fn in (lambda: O.create_team(st, LOU, {"name": "X"}, USERS), lambda: O.create_customer(st, LOU, {"name": "X"}),
                   lambda: O.save_project(st, LOU, acme, {"name": "X"}), lambda: O.update_customer(st, LOU, acme, {"teams": []}),
                   lambda: O.delete_team(st, LOU, dba)):
            with self.assertRaises(O.OrgError) as e:
                fn()
            self.assertEqual(e.exception.status, 403)

    def test_names_codes_and_duplicates(self):
        st, *_ = setup_state()
        self.assertEqual(next(iter(st["customers"].values()))["code"], "ACME")
        for bad in ({"name": "dba"}, {"name": ""}, {"name": "<b>x</b>"}):
            with self.assertRaises(O.OrgError):
                O.create_team(st, ADMIN, bad, USERS)
        with self.assertRaises(O.OrgError):
            O.create_customer(st, ADMIN, {"name": "New", "code": "has space"})

    def test_leads_need_the_lead_role_and_become_members(self):
        st, dba, *_ = setup_state()
        with self.assertRaises(O.OrgError) as e:
            O.update_team(st, ADMIN, dba, {"leads": ["eli.eng"]}, USERS)
        self.assertIn("needs the Lead role", e.exception.message)
        O.update_team(st, ADMIN, dba, {"leads": ["lou.lead", "kim.lead"]}, USERS)  # one person can lead several teams
        self.assertIn("kim.lead", st["teams"][dba]["members"])
        self.assertEqual(O.teams_of(st, "kim.lead"), list(st["teams"]))

    def test_primary_team_must_serve_the_customer(self):
        st, dba, sa, acme, ini, *_ = setup_state()
        with self.assertRaises(O.OrgError):
            O.update_customer(st, ADMIN, ini, {"primary_team": dba})
        O.update_customer(st, ADMIN, acme, {"teams": [dba]})
        self.assertIsNone(st["customers"][acme]["primary_team"])  # Unix SA no longer serves Acme


class T02TeamLeads(unittest.TestCase):
    def test_lead_manages_only_own_team(self):
        st, dba, sa, *_ = setup_state()
        O.add_member(st, LOU, dba, "eli.eng", USERS)
        with self.assertRaises(O.OrgError):
            O.add_member(st, LOU, sa, "mia.s", USERS)
        with self.assertRaises(O.OrgError):
            O.add_member(st, ELI, dba, "mia.s", USERS)
        with self.assertRaises(O.OrgError):
            O.update_team(st, LOU, dba, {"leads": ["lou.lead", "kim.lead"]}, USERS)
        O.update_team(st, LOU, dba, {"description": "Databases"}, USERS)
        self.assertEqual(st["teams"][dba]["description"], "Databases")

    def test_admins_are_never_members(self):
        st, dba, *_ = setup_state()
        with self.assertRaises(O.OrgError) as e:
            O.add_member(st, ADMIN, dba, "ada", USERS)
        self.assertIn("separation of duties", e.exception.message)

    def test_people_in_several_teams(self):
        st, dba, sa, *_ = setup_state()
        O.add_member(st, LOU, dba, "mia.s", USERS)
        O.add_member(st, KIM, sa, "mia.s", USERS)
        self.assertEqual(sorted(O.teams_of(st, "mia.s")), sorted([dba, sa]))
        with self.assertRaises(O.OrgError):
            O.add_member(st, LOU, dba, "mia.s", USERS)

    def test_projects_only_from_customers_the_team_serves(self):
        st, dba, sa, acme, ini, p1, p2 = setup_state()
        O.add_member(st, LOU, dba, "eli.eng", USERS)
        self.assertEqual(O.set_member_projects(st, LOU, dba, "eli.eng", [p1]), [p1])
        with self.assertRaises(O.OrgError):
            O.set_member_projects(st, LOU, dba, "eli.eng", [p2])  # Initech is served by Unix SA only

    def test_lead_cannot_be_removed_as_member(self):
        st, dba, *_ = setup_state()
        with self.assertRaises(O.OrgError):
            O.remove_member(st, LOU, dba, "lou.lead")


class T03Cleanup(unittest.TestCase):
    def test_deleting_things_removes_assignments(self):
        st, dba, sa, acme, ini, p1, p2 = setup_state()
        O.add_member(st, LOU, dba, "eli.eng", USERS)
        O.set_member_projects(st, LOU, dba, "eli.eng", [p1])
        O.delete_project(st, ADMIN, acme, p1)
        self.assertEqual(st["teams"][dba]["members"]["eli.eng"]["projects"], [])
        O.delete_team(st, ADMIN, sa)
        self.assertEqual(st["customers"][acme]["teams"], [dba])
        self.assertIsNone(st["customers"][ini]["primary_team"])
        self.assertEqual(O.forget_user(st, "lou.lead"), 1)
        self.assertEqual(st["teams"][dba]["leads"], [])


class T04Views(unittest.TestCase):
    def test_who_sees_what(self):
        st, dba, sa, acme, ini, *_ = setup_state()
        O.add_member(st, LOU, dba, "eli.eng", USERS)
        v = O.view(st, ELI, USERS)
        self.assertEqual([t["name"] for t in v["teams"]], ["DBA"])
        self.assertEqual([c["name"] for c in v["customers"]], ["Acme Corp"])  # not Initech
        self.assertFalse(v["teams"][0]["can_manage"])
        v = O.view(st, LOU, USERS)
        self.assertTrue(v["teams"][0]["can_manage"] and v["leads_any"])
        v = O.view(st, ADMIN, USERS)
        self.assertEqual(len(v["teams"]), 2)
        self.assertEqual(O.view(st, {"user": "mia.s", "role": "Engineer"}, USERS)["teams"], [])


class T05Store(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp()
        self.store = O.OrgStore(R.FileStorage(self.tmp))

    def tearDown(self):
        shutil.rmtree(self.tmp)

    def test_versions_restore_and_conflict_retry(self):
        tid = self.store.mutate(ADMIN, "Create team", "DBA", lambda st: O.create_team(st, ADMIN, {"name": "DBA"}, USERS))["id"]
        self.store.mutate(ADMIN, "Rename", tid, lambda st: O.update_team(st, ADMIN, tid, {"name": "Databases"}, USERS))
        self.assertEqual(self.store.load()[0]["teams"][tid]["name"], "Databases")
        h = self.store.history()
        self.store.restore(ADMIN, h[-1]["index"])
        self.assertEqual(self.store.load()[0]["teams"][tid]["name"], "DBA")
        with self.assertRaises(O.OrgError):
            self.store.restore(LOU, 0)
        real, calls = self.store.storage.write, {"n": 0}

        def flaky(name, data, expected):
            calls["n"] += 1
            if calls["n"] == 1:
                raise R.ConflictError(name)
            return real(name, data, expected)
        self.store.storage.write = flaky
        self.store.mutate(ADMIN, "Describe", tid, lambda st: O.update_team(st, ADMIN, tid, {"description": "x"}, USERS))
        self.assertEqual(calls["n"], 2)


if __name__ == "__main__":
    unittest.main(verbosity=2)
