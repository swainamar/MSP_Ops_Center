#!/usr/bin/env python3
"""Tests for msp_roster.py. Run: python -m unittest test_msp_roster -v"""
import copy
import datetime as dt
import io
import os
import shutil
import tempfile
import unittest

import msp_roster as R

ADMIN = {"user": "ada", "role": "Admin"}
LEAD = {"user": "raj.k", "role": "Lead"}
MIA = {"user": "mia.s", "role": "Engineer"}
VIEWER = {"user": "vic", "role": "Viewer"}


def small_roster(**over):
    r = R.default_roster("Unix", "Unix 24x7", "2026-11-02")
    r["members"] = [{"id": f"m{i}", "name": n, "role": "Engineer", "user": u, "offset": o}
                    for i, (n, u, o) in enumerate([("Mia Silva", "mia.s", 0), ("Ravi Nair", "ravi.n", 1), ("Ana Ito", "ana", 6),
                                                   ("Tom Berg", "tom.b", 11), ("Lena Fox", "lena.f", 16)])]
    r.update(over)
    return r


class T01Validation(unittest.TestCase):
    def test_default_roster_is_valid(self):
        R.validate_roster(small_roster())

    def test_bad_time_and_code_rejected(self):
        r = small_roster()
        r["shifts"][0]["start"] = "25:00"
        r["shifts"][1]["code"] = "off"
        with self.assertRaises(R.RosterError) as e:
            R.validate_roster(r)
        self.assertIn("not a time", e.exception.message)
        self.assertIn("capital letters", e.exception.message)

    def test_rotation_must_match_cycle_and_codes(self):
        r = small_roster(cycle={"type": "weekly"})
        with self.assertRaises(R.RosterError) as e:
            R.validate_roster(r)
        self.assertIn("entries but the weekly cycle", e.exception.message)
        r = small_roster()
        r["rotation"]["sequence"][0] = "X"
        with self.assertRaises(R.RosterError):
            R.validate_roster(r)

    def test_monthly_accepts_any_length(self):
        r = small_roster(cycle={"type": "monthly"})
        r["rotation"]["sequence"] = ["M", "A", "OFF"]
        R.validate_roster(r)

    def test_approval_mode_checked(self):
        r = small_roster(approvals={"leave": "maybe", "swap": "lead"})
        with self.assertRaises(R.RosterError):
            R.validate_roster(r)


class T02Cycles(unittest.TestCase):
    def test_custom_cycle_and_offsets(self):
        r = small_roster()
        s = R.generate(r, "2026-11-02", "2026-11-08")
        self.assertEqual(s["2026-11-02"]["m0"]["code"], "M")
        self.assertEqual(s["2026-11-07"]["m0"]["code"], "OFF")
        self.assertEqual(s["2026-11-02"]["m2"]["code"], "OFF")  # offset 6: last day of first off block
        self.assertEqual(s["2026-11-03"]["m2"]["code"], "A")

    def test_weekly_and_daily(self):
        r = small_roster(cycle={"type": "weekly"})
        r["rotation"]["sequence"] = ["D", "D", "D", "D", "D", "OFF", "OFF"]
        R.validate_roster(r)
        s = R.generate(r, "2026-11-02", "2026-11-15")
        self.assertEqual([s[d]["m0"]["code"] for d in sorted(s)][:7], ["D"] * 5 + ["OFF"] * 2)
        r = small_roster(cycle={"type": "daily"})
        r["rotation"]["sequence"] = ["M"]
        R.validate_roster(r)
        self.assertTrue(all(v["m0"]["code"] == "M" for v in R.generate(r, "2026-11-02", "2026-11-05").values()))

    def test_monthly_restarts_on_first(self):
        r = small_roster(cycle={"type": "monthly"})
        r["rotation"]["sequence"] = ["M", "A", "N"]
        s = R.generate(r, "2026-11-29", "2026-12-02")
        self.assertEqual(s["2026-12-01"]["m0"]["code"], "M")
        self.assertEqual(s["2026-11-30"]["m0"]["code"], "N")  # day 30: index 29 % 3 = 2

    def test_inactive_weekday_becomes_off(self):
        r = small_roster(cycle={"type": "weekly"})
        r["rotation"]["sequence"] = ["D"] * 7
        s = R.generate(r, "2026-11-07", "2026-11-07")  # Saturday
        self.assertEqual(s["2026-11-07"]["m0"], {"code": "OFF", "source": "inactive day"})

    def test_nothing_before_start(self):
        self.assertEqual(R.generate(small_roster(), "2026-11-01", "2026-11-01")["2026-11-01"], {})


class T03Checks(unittest.TestCase):
    def test_full_coverage_has_no_findings(self):
        r = small_roster()
        self.assertEqual(R.check(r, R.generate(r, "2026-11-02", "2026-12-06")), [])

    def test_coverage_shortfall(self):
        r = small_roster()
        r["members"] = r["members"][:3]
        msgs = [i for i in R.check(r, R.generate(r, "2026-11-02", "2026-11-15")) if i["rule"] == "coverage"]
        self.assertTrue(msgs)
        self.assertTrue(all(i["level"] == "error" for i in msgs))

    def test_overnight_rest_and_forbidden_sequence(self):
        r = small_roster()
        r["overrides"] = {"2026-11-02": {"m0": "N"}, "2026-11-03": {"m0": "M"}}
        rules = {i["rule"] for i in R.check(r, R.generate(r, "2026-11-02", "2026-11-03")) if i["member"] == "m0"}
        self.assertIn("rest", rules)       # Night ends 07:00, Morning starts 06:30
        self.assertIn("sequence", rules)   # N then M is forbidden

    def test_max_consecutive_days(self):
        r = small_roster()
        r["overrides"] = {"2026-11-07": {"m0": "M"}, "2026-11-08": {"m0": "M"}}
        found = [i for i in R.check(r, R.generate(r, "2026-11-02", "2026-11-08")) if i["rule"] == "consecutive days"]
        self.assertTrue(found)

    def test_shared_member_flagged(self):
        a, b = small_roster(), small_roster()
        b["id"], b["name"] = "other", "Windows 24x7"
        found = [i for i in R.check(a, R.generate(a, "2026-11-02", "2026-11-02"), [b]) if i["rule"] == "shared member"]
        self.assertEqual(len(found), 5)
        a["rules"]["allow_shared_members"] = True
        self.assertFalse([i for i in R.check(a, R.generate(a, "2026-11-02", "2026-11-02"), [b]) if i["rule"] == "shared member"])


class T04Workflow(unittest.TestCase):
    def test_leave_needs_lead_when_configured(self):
        r = small_roster()
        lv = R.request_leave(r, MIA, "m0", "2026-11-03", "2026-11-04", "Annual")
        self.assertEqual(lv["status"], "pending")
        self.assertEqual(R.generate(r, "2026-11-03", "2026-11-03")["2026-11-03"]["m0"]["code"], "M")
        with self.assertRaises(R.RosterError) as e:
            R.decide(r, MIA, "leave", lv["id"], True)
        self.assertEqual(e.exception.status, 403)
        R.decide(r, LEAD, "leave", lv["id"], True)
        self.assertEqual(R.generate(r, "2026-11-03", "2026-11-03")["2026-11-03"]["m0"]["code"], "LEAVE")

    def test_self_recorded_leave(self):
        r = small_roster(approvals={"leave": "self", "swap": "self"})
        lv = R.request_leave(r, MIA, "m0", "2026-11-03", "2026-11-03", "Sick")
        self.assertEqual(lv["status"], "approved")

    def test_engineer_only_for_self_viewer_never(self):
        r = small_roster()
        with self.assertRaises(R.RosterError):
            R.request_leave(r, MIA, "m1", "2026-11-03", "2026-11-03")
        with self.assertRaises(R.RosterError):
            R.request_leave(r, VIEWER, "m0", "2026-11-03", "2026-11-03")

    def test_overlapping_leave_rejected(self):
        r = small_roster()
        R.request_leave(r, MIA, "m0", "2026-11-03", "2026-11-05")
        with self.assertRaises(R.RosterError) as e:
            R.request_leave(r, MIA, "m0", "2026-11-05", "2026-11-06")
        self.assertEqual(e.exception.status, 409)

    def test_swap_exchanges_and_warns(self):
        r = small_roster()
        before = R.generate(r, "2026-11-03", "2026-11-03")["2026-11-03"]
        sw = R.request_swap(r, MIA, "m0", "2026-11-03", "m2")
        self.assertEqual(sw["status"], "awaiting_partner")  # Ana must agree first
        with self.assertRaises(R.RosterError):
            R.decide(r, LEAD, "swap", sw["id"], True)
        R.consent(r, {"user": "ana", "role": "Engineer"}, sw["id"], True)
        self.assertEqual(sw["status"], "pending")
        R.decide(r, LEAD, "swap", sw["id"], True)
        after = R.generate(r, "2026-11-03", "2026-11-03")["2026-11-03"]
        self.assertEqual(after["m0"]["code"], before["m2"]["code"])
        self.assertEqual(after["m2"]["code"], before["m0"]["code"])
        self.assertIsInstance(sw["warnings"], list)

    def test_cover(self):
        r = small_roster(approvals={"leave": "lead", "swap": "self"})
        day = "2026-11-07"
        orig = R.generate(r, day, day)[day]
        off_person = next(k for k, v in orig.items() if v["code"] == "OFF")
        worker = next(k for k, v in orig.items() if v["code"] == "M")
        sw = R.request_swap(r, LEAD, worker, day, off_person, kind="cover")
        partner = next(m for m in r["members"] if m["id"] == off_person)
        R.consent(r, {"user": partner["user"], "role": "Engineer"}, sw["id"], True)
        self.assertEqual(sw["status"], "approved")  # a Lead asked, so no further approval
        after = R.generate(r, day, day)[day]
        self.assertEqual(after[off_person]["code"], "M")
        self.assertEqual(after[worker]["code"], "OFF")

    def test_cancel_and_reject(self):
        r = small_roster()
        lv = R.request_leave(r, MIA, "m0", "2026-11-03", "2026-11-03")
        R.cancel(r, MIA, "leave", lv["id"])
        self.assertEqual(lv["status"], "cancelled")
        lv2 = R.request_leave(r, MIA, "m0", "2026-11-10", "2026-11-10")
        R.decide(r, LEAD, "leave", lv2["id"], False, "Short staffed")
        self.assertEqual(lv2["status"], "rejected")
        with self.assertRaises(R.RosterError):
            R.decide(r, LEAD, "leave", lv2["id"], True)


class T05Oncall(unittest.TestCase):
    def setUp(self):
        self.r = small_roster()
        ids = [m["id"] for m in self.r["members"]]
        self.r["oncall"].update(enabled=True, rotation_days=7, order={"Primary": ids, "Secondary": ids[1:] + ids[:1]},
                                escalation=[{"level": 1, "name": "Lead", "contact": "x"}])

    def test_weekly_rotation(self):
        self.assertEqual(R.oncall_for(self.r, "2026-11-02")["Primary"]["member_id"], "m0")
        self.assertEqual(R.oncall_for(self.r, "2026-11-09")["Primary"]["member_id"], "m1")
        self.assertEqual(R.oncall_for(self.r, "2026-11-02")["Secondary"]["member_id"], "m1")

    def test_leave_skips_to_next_and_no_double_booking(self):
        lv = R.request_leave(self.r, LEAD, "m0", "2026-11-02", "2026-11-03")
        self.assertEqual(lv["status"], "approved")
        oc = R.oncall_for(self.r, "2026-11-02")
        self.assertEqual(oc["Primary"]["member_id"], "m1")
        self.assertTrue(oc["Primary"]["substitute"])
        self.assertNotEqual(oc["Secondary"]["member_id"], oc["Primary"]["member_id"])

    def test_override(self):
        self.r["oncall"]["overrides"] = {"2026-11-04": {"Primary": "m4"}}
        self.assertEqual(R.oncall_for(self.r, "2026-11-04")["Primary"]["member_id"], "m4")

    def test_disabled(self):
        self.r["oncall"]["enabled"] = False
        self.assertEqual(R.oncall_for(self.r, "2026-11-02"), {})


class T06Views(unittest.TestCase):
    def test_week_and_month_ranges(self):
        r = small_roster()
        w = R.calendar(r, "2026-11-05", "week")
        self.assertEqual((w["start"], w["end"]), ("2026-11-02", "2026-11-08"))
        m = R.calendar(r, "2026-11-15", "month")
        self.assertEqual((m["start"], m["end"]), ("2026-11-01", "2026-11-30"))
        self.assertEqual(len(R.calendar(r, "2026-11-05", "day")["days"]), 1)

    def test_person_view_has_shift_log_link(self):
        r = small_roster()
        rows = R.person_view(r, "m0", "2026-11-02", "2026-11-03")
        self.assertEqual(rows[0]["shift"], "Morning")
        self.assertEqual(rows[0]["shift_log"], "/#shiftlog=2026-11-02/Morning")

    def test_range_limit(self):
        with self.assertRaises(R.RosterError):
            R.generate(small_roster(), "2026-01-01", "2027-12-31")


class T07StoreAndVersions(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp()
        self.store = R.RosterStore(R.FileStorage(self.tmp))

    def tearDown(self):
        shutil.rmtree(self.tmp)

    def test_create_requires_lead(self):
        with self.assertRaises(R.RosterError):
            self.store.create(MIA, R.default_roster("Unix", "x", "2026-11-02"))

    def test_versions_and_restore(self):
        r = self.store.create(ADMIN, R.default_roster("Unix", "Unix", "2026-11-02"))
        self.store.update(ADMIN, r["id"], {"name": "Renamed"})
        self.assertEqual(self.store.get(r["id"])["name"], "Renamed")
        hist = self.store.history(r["id"])
        self.assertEqual(hist[-1]["action"], "Update roster")
        self.store.restore(ADMIN, r["id"], hist[-1]["index"])
        self.assertEqual(self.store.get(r["id"])["name"], "Unix")

    def test_invalid_change_not_saved(self):
        r = self.store.create(ADMIN, R.default_roster("Unix", "Unix", "2026-11-02"))
        with self.assertRaises(R.RosterError):
            self.store.update(ADMIN, r["id"], {"cycle": {"type": "weekly"}})
        self.assertEqual(self.store.get(r["id"])["cycle"]["type"], "custom")

    def test_dedicated_members(self):
        a = self.store.create(ADMIN, R.default_roster("Unix", "Unix", "2026-11-02"))
        b = self.store.create(ADMIN, R.default_roster("Windows", "Windows", "2026-11-02"))
        self.store.add_member(ADMIN, a["id"], "Mia Silva", user="mia.s")
        with self.assertRaises(R.RosterError) as e:
            self.store.add_member(ADMIN, b["id"], "Mia Silva", user="mia.s")
        self.assertIn("already belongs", e.exception.message)
        rb = self.store.get(b["id"])
        rb["rules"]["allow_shared_members"] = True
        self.store.update(ADMIN, b["id"], {"rules": rb["rules"]})
        self.store.add_member(ADMIN, b["id"], "Mia Silva", user="mia.s")

    def test_conflict_retry(self):
        r = self.store.create(ADMIN, R.default_roster("Unix", "Unix", "2026-11-02"))
        real_write = self.store.storage.write
        calls = {"n": 0}

        def flaky(name, data, expected):
            calls["n"] += 1
            if calls["n"] == 1:
                raise R.ConflictError(name)
            return real_write(name, data, expected)
        self.store.storage.write = flaky
        self.store.update(ADMIN, r["id"], {"name": "After retry"})
        self.assertEqual(self.store.get(r["id"])["name"], "After retry")
        self.assertEqual(calls["n"], 2)

    def test_storage_path_guard(self):
        with self.assertRaises(R.RosterError):
            R.FileStorage(self.tmp).read("../x.json")


class T08Exports(unittest.TestCase):
    def test_excel_sheets_and_formulas(self):
        import openpyxl
        r = small_roster()
        r["oncall"].update(enabled=True, order={"Primary": ["m0", "m1"]})
        data = R.export_excel(r, "2026-11-02", "2026-11-08")
        wb = openpyxl.load_workbook(io.BytesIO(data))
        self.assertEqual(wb.sheetnames, ["Calendar", "By shift", "Coverage", "On-call", "Requests", "Shifts", "Rules", "Holidays", "Legend"])
        self.assertTrue(str(wb["Coverage"]["C3"].value).startswith("=COUNTIF(Calendar!C$3:C$7"))
        self.assertEqual(wb["Calendar"]["C3"].value, "M")

    def test_html_escapes(self):
        r = small_roster()
        r["members"][0]["name"] = "<img src=x onerror=alert(1)>"
        page = R.export_html(r, "2026-11-02", "2026-11-03")
        self.assertNotIn("<img src=x", page)
        self.assertIn("&lt;img src=x", page)


SHEET = """Day\tDATE\tMORNING SHIFT\tAFTERNOON SHIFT\tNIGHT SHIFT\tWEEK OFF\tLEAVE
\t\t6AM TO 3PM\t1PM TO 10PM\t9:30 PM TO 6:30 AM\t\t
\t\t\t\t\t\t
Friday\t1-Apr-22\tALPHA, BRAVO\tCHARLIE, DELTA\tECHO, FOXTROT\tGOLF, HOTEL\tINDIA
Saturday\t2-Apr-22\tGOLF, HOTEL\tALPHA, BRAVO\tCHARLIE, DELTA\tECHO, FOXTROT, INDIA\t
\t\t\t\t\t\t
Sunday\t3-Apr-22\tECHO, INDIA\tGOLF, HOTEL\tALPHA, BRAVO\tCHARLIE, DELTA, FOXTROT, XRAY\t
"""


class T09ByShiftTemplate(unittest.TestCase):
    def test_three_shift_template(self):
        r = R.default_roster("NOC", "NOC", "2026-11-02", "three_shift")
        R.validate_roster(r)
        times = {s["code"]: (s["start"], s["end"]) for s in r["shifts"]}
        self.assertEqual(times, {"M": ("06:00", "15:00"), "A": ("13:00", "22:00"), "N": ("21:30", "06:30")})
        self.assertEqual(r["layout"], "by_shift")
        with self.assertRaises(R.RosterError):
            R.default_roster("NOC", "NOC", "2026-11-02", "nope")

    def test_parse_sheet(self):
        p = R.parse_by_shift(SHEET)
        self.assertEqual([s["name"] for s in p["shifts"]], ["Morning", "Afternoon", "Night"])
        self.assertEqual(p["shifts"][2], {"name": "Night", "start": "21:30", "end": "06:30"})
        self.assertEqual([d["date"] for d in p["days"]], ["2022-04-01", "2022-04-02", "2022-04-03"])
        self.assertEqual(p["days"][0]["leave"], ["INDIA"])
        self.assertIn("XRAY is not listed on 2 day(s): 2022-04-01 to 2022-04-02.", p["warnings"])

    def test_similar_names_and_double_listing(self):
        self.assertTrue(R.similar_names(["MUKKU", "NMUKKU", "ALPHA"]))
        self.assertFalse(R.similar_names(["ALPHA", "BRAVO"]))
        bad = SHEET.replace("GOLF, HOTEL\tINDIA", "GOLF, HOTEL, ALPHA\tINDIA")
        self.assertTrue(any("listed under both" in w for w in R.parse_by_shift(bad)["warnings"]))

    def test_no_header(self):
        with self.assertRaises(R.RosterError):
            R.parse_by_shift("just\ttext\n1\t2\n")

    def test_import_new_and_round_trip(self):
        tmp = tempfile.mkdtemp()
        try:
            st = R.RosterStore(R.FileStorage(tmp))
            res = st.import_new(LEAD, SHEET, "NOC", "April")
            r = res["roster"]
            self.assertEqual(len(r["members"]), 10)
            s = R.generate(r, "2022-04-01", "2022-04-03")
            alpha = next(m["id"] for m in r["members"] if m["name"] == "ALPHA")
            self.assertEqual([s[d][alpha]["code"] for d in sorted(s)], ["M", "A", "N"])
            rows = R.by_shift_rows(r, "2022-04-01", "2022-04-01")
            self.assertEqual(rows[0]["shifts"]["M"], "ALPHA, BRAVO")
            self.assertEqual(rows[0]["leave"], "INDIA")
            self.assertNotIn("XRAY", rows[0]["off"])  # not on the sheet that day, so not shown as week off
            text = "Day\tDATE\tMORNING SHIFT\tAFTERNOON SHIFT\tNIGHT SHIFT\tWEEK OFF\tLEAVE\n\t\t6AM TO 3PM\t1PM TO 10PM\t9:30 PM TO 6:30 AM\t\t\n"
            for x in R.by_shift_rows(r, "2022-04-01", "2022-04-03"):
                d = R.parse_date(x["date"]).strftime("%d-%b-%y")
                text += "\t".join([x["day"], d, x["shifts"]["M"], x["shifts"]["A"], x["shifts"]["N"], x["off"], x["leave"]]) + "\n"
            again = R.parse_by_shift(text)
            first = R.parse_by_shift(SHEET)
            for a, b in zip(first["days"], again["days"]):
                self.assertEqual({k: sorted(v) for k, v in a["shifts"].items()}, {k: sorted(v) for k, v in b["shifts"].items()})
                self.assertEqual(sorted(a["leave"]), sorted(b["leave"]))
        finally:
            shutil.rmtree(tmp)

    def test_engineer_cannot_import(self):
        with self.assertRaises(R.RosterError):
            R.roster_from_by_shift(R.parse_by_shift(SHEET), "NOC", "x", MIA)

    def test_both_layouts_export(self):
        r = R.default_roster("NOC", "NOC", "2026-11-02", "three_shift")
        r["members"] = small_roster()["members"]
        page = R.export_html(r, "2026-11-02", "2026-11-04")
        self.assertIn('<section class="pane on" id="by_shift">', page)  # opens on its own layout
        self.assertIn("06:00-15:00", page)
        page2 = R.export_html(r, "2026-11-02", "2026-11-04", "by_person", history=[{"at": "t", "by": "raj.k", "action": "Update roster"}])
        self.assertIn('<section class="pane on" id="by_person">', page2)
        self.assertIn("Version 2", page2)
        self.assertIn("(1/2)", page)  # a shortfall is shown against the minimum
        with self.assertRaises(R.RosterError):
            R.export_html(r, "2026-11-02", "2026-11-04", "sideways")
        import openpyxl
        wb = openpyxl.load_workbook(io.BytesIO(R.export_excel(r, "2026-11-02", "2026-11-04")))
        self.assertEqual(wb.sheetnames[0], "By shift")  # this template opens on its own layout
        self.assertEqual(wb["By shift"]["C2"].value, "06:00 to 15:00")


class T10ImportWorkbooks(unittest.TestCase):
    """Re-import both exported layouts: the by-person case is the one reported as failing."""

    def setUp(self):
        self.tmp = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.tmp)

    def _export(self, roster, name):
        path = os.path.join(self.tmp, name)
        with open(path, "wb") as fh:
            fh.write(R.export_excel(roster, "2026-11-02", "2026-11-15"))
        return path

    def test_by_person_workbook(self):
        r = small_roster()
        r["members"][0]["role"] = "Lead"
        path = self._export(r, "person.xlsx")
        st = R.RosterStore(R.FileStorage(self.tmp))
        res = st.import_new(LEAD, path, "Ops", "\nApril ")  # stray line break, as from a PowerShell line wrap
        self.assertEqual(res["layout"], "by_person")
        self.assertEqual(res["sheet"], "Calendar")
        new = res["roster"]
        self.assertEqual(new["name"], "April")
        self.assertEqual({(s["code"], s["start"], s["end"]) for s in new["shifts"]},
                         {("M", "06:30", "15:30"), ("A", "13:30", "21:30"), ("N", "21:00", "07:00")})
        self.assertEqual(new["members"][0]["role"], "Lead")
        before, after = R.generate(r, "2026-11-02", "2026-11-15"), R.generate(new, "2026-11-02", "2026-11-15")
        ids = {m["name"]: m["id"] for m in new["members"]}
        for d in before:
            for m in r["members"]:
                self.assertEqual(before[d][m["id"]]["code"], after[d][ids[m["name"]]]["code"], (d, m["name"]))

    def test_by_shift_workbook(self):
        r = R.default_roster("NOC", "NOC", "2026-11-02", "three_shift")
        r["members"] = small_roster()["members"]
        path = self._export(r, "shift.xlsx")
        res = R.RosterStore(R.FileStorage(self.tmp)).import_new(LEAD, path, "NOC", "Nov")
        self.assertEqual((res["layout"], res["sheet"]), ("by_shift", "By shift"))

    def test_unknown_workbook(self):
        import openpyxl
        wb = openpyxl.Workbook()
        wb.active.append(["nothing", "here"])
        path = os.path.join(self.tmp, "x.xlsx")
        wb.save(path)
        with self.assertRaises(R.RosterError) as e:
            R.parse_roster_file(path)
        self.assertIn("No roster found", e.exception.message)

    def test_cli_reports_plain_error(self):
        import subprocess
        import sys
        out = subprocess.run([sys.executable, R.__file__, "import", "--file", os.path.join(self.tmp, "missing.xlsx"),
                              "--team", "a", "--name", "b", "--data", self.tmp, "--out", self.tmp],
                             capture_output=True, text=True)
        self.assertEqual(out.returncode, 2)
        self.assertIn("Error: File not found", out.stderr)
        self.assertNotIn("Traceback", out.stderr)


class T11Holidays(unittest.TestCase):
    def test_company_list_without_region(self):
        r = small_roster(holidays={"region": "", "subdivision": "", "ignore": [],
                                   "custom": [{"date": "2026-11-10", "name": "Company day"}]})
        R.validate_roster(r)
        self.assertEqual(R.holidays_for(r, "2026-11-01", "2026-11-30"), {"2026-11-10": "Company day"})

    @unittest.skipUnless(R.holidays_available(), "holidays package not installed")
    def test_us_and_india(self):
        us = small_roster(holidays={"region": "US", "subdivision": "", "custom": [], "ignore": []})
        self.assertEqual(R.holidays_for(us, "2026-11-01", "2026-11-30").get("2026-11-26"), "Thanksgiving Day")
        ind = small_roster(holidays={"region": "IN", "subdivision": "", "custom": [], "ignore": []})
        self.assertEqual(R.holidays_for(ind, "2026-01-01", "2026-01-31").get("2026-01-26"), "Republic Day")
        ka = small_roster(holidays={"region": "IN", "subdivision": "KA", "custom": [], "ignore": ["2026-01-26"]})
        jan = R.holidays_for(ka, "2026-01-01", "2026-01-31")
        self.assertIn("2026-01-14", jan)          # Makar Sankranti in Karnataka
        self.assertNotIn("2026-01-26", jan)       # ignored by the company
        self.assertIn("KA", R.holiday_subdivisions("IN"))
        bad = small_roster(holidays={"region": "IN", "subdivision": "ZZ", "custom": [], "ignore": []})
        with self.assertRaises(R.RosterError):
            R.holidays_for(bad, "2026-01-01", "2026-01-31")

    def test_validation(self):
        with self.assertRaises(R.RosterError):
            R.validate_roster(small_roster(holidays={"region": "UK", "subdivision": "", "custom": [], "ignore": []}))
        with self.assertRaises(R.RosterError):
            R.validate_roster(small_roster(holidays={"region": "", "subdivision": "KA", "custom": [], "ignore": []}))
        with self.assertRaises(R.RosterError):
            R.validate_roster(small_roster(holidays={"region": "", "subdivision": "", "custom": [{"date": "x"}], "ignore": []}))

    def test_holiday_minimum_staff(self):
        r = small_roster(holidays={"region": "", "subdivision": "", "ignore": [],
                                   "custom": [{"date": "2026-11-03", "name": "Company day"}]})
        r["shifts"][0]["holiday_min_staff"] = 3
        found = [i for i in R.check(r, R.generate(r, "2026-11-03", "2026-11-03")) if i["rule"] == "coverage"]
        self.assertTrue(any("Morning has 2 of 3 people (Company day)" in i["message"] for i in found))
        r["shifts"][0]["holiday_min_staff"] = 0
        r["members"] = [m for m in r["members"] if m["id"] != "m0"]  # nobody on Morning that day
        self.assertFalse([i for i in R.check(r, R.generate(r, "2026-11-03", "2026-11-03"))
                          if i["rule"] == "coverage" and "Morning" in i["message"]])

    def test_exports_show_holidays(self):
        import openpyxl
        r = small_roster(holidays={"region": "", "subdivision": "", "ignore": [],
                                   "custom": [{"date": "2026-11-04", "name": "Company day"}]})
        wb = openpyxl.load_workbook(io.BytesIO(R.export_excel(r, "2026-11-02", "2026-11-08")))
        self.assertEqual([c.value for c in wb["Holidays"][2]][2:], ["Company day", "Company list"])
        self.assertEqual(wb["By shift"].cell(row=5, column=9).value, "Company day")
        page = R.export_html(r, "2026-11-02", "2026-11-08")
        self.assertIn("Holidays (1)", page)
        self.assertIn("<span class=hb>Company day</span>", page)
        self.assertEqual(R.calendar(r, "2026-11-04", "day")["days"][0]["holiday"], "Company day")


class T12CommandLine(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp()
        R.main(["demo", "--out", self.tmp])
        self.r = R.RosterStore(R.FileStorage(self.tmp)).rosters()[0]

    def tearDown(self):
        shutil.rmtree(self.tmp)

    def run_cli(self, *args):
        import contextlib
        buf = io.StringIO()
        with contextlib.redirect_stdout(buf):
            R.main(list(args))
        return buf.getvalue()

    def test_leave_and_approval_roles(self):
        base = ["--data", self.tmp, "--roster", self.r["name"]]
        out = self.run_cli("leave", *base, "--member", "mia.s", "--start", "2026-11-23", "--end", "2026-11-23",
                           "--as-user", "mia.s", "--as-role", "Engineer")
        self.assertIn("pending", out)
        req = out.split()[1].rstrip(":")
        with self.assertRaises(R.RosterError):
            self.run_cli("approve", *base, "--request", req, "--as-user", "ana", "--as-role", "Engineer")
        self.assertIn("approved", self.run_cli("approve", *base, "--request", req, "--as-user", "raj.k"))

    def test_swap_cover_and_holidays(self):
        base = ["--data", self.tmp, "--roster", self.r["id"]]
        self.assertIn("Cover", self.run_cli("swap", *base, "--member", "Ana Ito", "--date", "2026-11-20",
                                            "--with", "Lena Fox", "--cover"))
        out = self.run_cli("holidays", *base, "--add", "2026-11-10=Company day", "--year", "2026")
        self.assertIn("2026-11-10 Tue Company day", out)
        with self.assertRaises(R.RosterError):
            self.run_cli("leave", *base, "--member", "nobody", "--start", "2026-11-23", "--end", "2026-11-23")


class T13ApprovalAndConsent(unittest.TestCase):
    def setUp(self):
        self.r = small_roster(approvers=["raj.k", "kim.l"])
        self.ana = {"user": "ana", "role": "Engineer"}

    def test_leave_assigned_to_approver(self):
        lv = R.request_leave(self.r, MIA, "m0", "2026-11-10", "2026-11-10", approver="kim.l")
        self.assertEqual(lv["approver"], "kim.l")
        with self.assertRaises(R.RosterError) as e:
            R.decide(self.r, LEAD, "leave", lv["id"], True)  # raj.k is a Lead but not the assigned approver
        self.assertIn("assigned to kim.l", e.exception.message)
        R.decide(self.r, {"user": "kim.l", "role": "Lead"}, "leave", lv["id"], True)
        self.assertEqual(lv["status"], "approved")

    def test_default_and_unknown_approver_and_admin(self):
        lv = R.request_leave(self.r, MIA, "m0", "2026-11-11", "2026-11-11")
        self.assertEqual(lv["approver"], "raj.k")  # the first approver by default
        with self.assertRaises(R.RosterError):
            R.request_leave(self.r, MIA, "m0", "2026-11-12", "2026-11-12", approver="nobody")
        R.decide(self.r, ADMIN, "leave", lv["id"], False)  # an Admin can always decide
        self.assertEqual(lv["status"], "rejected")

    def test_partner_consent_rules(self):
        sw = R.request_swap(self.r, MIA, "m0", "2026-11-03", "m2", approver="kim.l")
        self.assertEqual((sw["status"], sw["partner"]), ("awaiting_partner", "Ana Ito"))
        with self.assertRaises(R.RosterError):
            R.consent(self.r, {"user": "tom.b", "role": "Engineer"}, sw["id"], True)  # not the partner
        self.assertEqual([x["action"] for x in R.waiting_for(self.r, self.ana)], ["consent"])
        R.consent(self.r, self.ana, sw["id"], True)
        self.assertEqual([x["id"] for x in R.waiting_for(self.r, {"user": "kim.l", "role": "Lead"})], [sw["id"]])
        self.assertEqual(R.waiting_for(self.r, LEAD), [])  # assigned to kim.l, not raj.k
        sw2 = R.request_swap(self.r, MIA, "m0", "2026-11-04", "m2")
        R.consent(self.r, self.ana, sw2["id"], False)
        self.assertEqual(sw2["status"], "declined")
        self.assertEqual(R.generate(self.r, "2026-11-04", "2026-11-04")["2026-11-04"]["m0"]["code"], "M")

    def test_partner_without_user_needs_no_consent(self):
        self.r["members"][2]["user"] = ""
        sw = R.request_swap(self.r, MIA, "m0", "2026-11-03", "m2")
        self.assertEqual(sw["status"], "pending")


class T14CloneDeleteTemplates(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp()
        self.st = R.RosterStore(R.FileStorage(self.tmp))
        r = small_roster()
        r["id"] = "src"
        r["leaves"] = [{"id": "lv_1", "member_id": "m0", "start": "2026-11-03", "end": "2026-11-03", "type": "Annual", "status": "approved"}]
        self.src = self.st.create(ADMIN, r)

    def tearDown(self):
        shutil.rmtree(self.tmp)

    def test_clone_without_people(self):
        c = self.st.clone(LEAD, "src", "Unix December", "Unix")
        self.assertEqual((c["name"], c["cloned_from"]), ("Unix December", "src"))
        self.assertEqual((c["members"], c["leaves"], c["overrides"]), ([], [], {}))
        self.assertEqual(c["shifts"], self.src["shifts"])

    def test_clone_with_people_needs_sharing(self):
        with self.assertRaises(R.RosterError) as e:
            self.st.clone(LEAD, "src", "Copy", "Unix", copy_members=True)
        self.assertIn("dedicated to one roster", e.exception.message)
        rules = dict(self.src["rules"], allow_shared_members=True)
        self.st.update(ADMIN, "src", {"rules": rules})
        c = self.st.clone(LEAD, "src", "Copy", "Unix", copy_members=True)
        self.assertEqual([m["name"] for m in c["members"]], [m["name"] for m in self.src["members"]])
        self.assertNotEqual(c["members"][0]["id"], self.src["members"][0]["id"])
        with self.assertRaises(R.RosterError):
            self.st.clone(MIA, "src", "x", "Unix")

    def test_delete_and_restore(self):
        self.st.delete(LEAD, "src")
        gone = self.st.deleted()
        self.assertEqual([(x["id"], x["name"]) for x in gone], [("src", "Unix 24x7")])
        with self.assertRaises(R.RosterError):
            self.st.undelete(MIA, "src")
        self.st.undelete(LEAD, "src")
        self.assertEqual(self.st.get("src")["name"], "Unix 24x7")
        self.assertEqual(len(self.st.get("src")["leaves"]), 1)
        self.assertEqual(self.st.deleted(), [])

    def test_bulk_changes_are_one_version(self):
        before = len(self.st.history("src"))
        n = self.st.set_overrides(LEAD, "src", [{"member_id": "m2", "date": "2026-11-04", "code": "M"},
                                                {"member_id": "m0", "date": "2026-11-04", "code": "OFF"}])
        self.assertEqual(n, 2)
        self.assertEqual(len(self.st.history("src")), before + 1)
        day = R.generate(self.st.get("src"), "2026-11-04", "2026-11-04")["2026-11-04"]
        self.assertEqual((day["m2"]["code"], day["m0"]["code"]), ("M", "OFF"))
        for bad in ([], [{"member_id": "m0", "date": "2026-11-04", "code": "M"}] * 501):
            with self.assertRaises(R.RosterError):
                self.st.set_overrides(LEAD, "src", bad)
        with self.assertRaises(R.RosterError):
            self.st.set_overrides(MIA, "src", [{"member_id": "m0", "date": "2026-11-04", "code": "M"}])

    def test_templates_import_back(self):
        for layout, fmt in (("by_shift", "xlsx"), ("by_shift", "csv"), ("by_person", "xlsx")):
            data, ctype, name = R.template_file(layout, fmt, "2026-11-09")
            path = os.path.join(self.tmp, name)
            with open(path, "wb") as fh:
                fh.write(data)
            res = self.st.import_new(LEAD, path, "Ops", f"T {layout} {fmt}")
            self.assertEqual(res["layout"], layout, name)
            self.assertEqual(sorted(m["name"] for m in res["roster"]["members"]), ["Asha Rao", "Ben Cole"])
        night = next(s for s in R.parse_roster_file(os.path.join(self.tmp, "roster_template_by_shift.csv"))["shifts"] if s["name"] == "Night")
        self.assertEqual((night["start"], night["end"]), ("21:30", "06:30"))
        with self.assertRaises(R.RosterError):
            R.template_file("by_person", "csv")


class T15PartlyFilledSheets(unittest.TestCase):
    """A sheet that covers only part of the month (as reported): days outside it are 'no roster', not errors."""

    def setUp(self):
        import openpyxl
        self.tmp = tempfile.mkdtemp()
        data, _, _ = R.template_file("by_shift", "xlsx", month="2026-11")
        wb = openpyxl.load_workbook(io.BytesIO(data))
        ws = wb["Roster"]
        names = ["ALPHA", "BRAVO", "CHARLIE", "DELTA", "ECHO", "FOXTROT", "GOLF", "HOTEL"]
        for row in range(3, ws.max_row + 1):
            for c in range(3, 8):
                ws.cell(row=row, column=c).value = None
            day = int(str(ws.cell(row=row, column=2).value).split("-")[0])
            if 9 <= day <= 22:  # only 9-22 Nov filled in, the same rota every day
                ws.cell(row=row, column=3).value = "ALPHA, BRAVO"
                ws.cell(row=row, column=4).value = "CHARLIE, DELTA, ECHO"
                ws.cell(row=row, column=5).value = "FOXTROT, GOLF"
                ws.cell(row=row, column=6).value = "HOTEL"
        self.path = os.path.join(self.tmp, "nov.xlsx")
        wb.save(self.path)
        self.st = R.RosterStore(R.FileStorage(self.tmp))
        self.res = self.st.import_new(LEAD, self.path, "SA", "Nov2026 (SA)")
        rules = dict(self.res["roster"]["rules"], max_consecutive_days=None)  # the test sheet repeats one rota every day
        self.r = self.st.update(LEAD, self.res["roster"]["id"], {"rules": rules})

    def tearDown(self):
        shutil.rmtree(self.tmp)

    def test_month_template_lists_every_day(self):
        import openpyxl
        ws = openpyxl.load_workbook(self.path)["Roster"]
        self.assertEqual(ws.max_row - 2, 30)
        with self.assertRaises(R.RosterError):
            R.template_file("by_shift", "xlsx", month="November")

    def test_days_outside_the_sheet_are_not_problems(self):
        sched = R.generate(self.r, "2026-11-01", "2026-11-30")
        self.assertEqual(R.check(self.r, sched), [])
        cal = R.calendar(self.r, "2026-11-15", "month")
        empty = [d["date"][8:] for d in cal["days"] if not d["has_data"]]
        self.assertEqual(empty, [f"{i:02d}" for i in list(range(1, 9)) + list(range(23, 31))])
        self.assertEqual(next(d for d in cal["days"] if d["date"] == "2026-11-23")["off"], [])  # nobody listed as week off

    def test_import_notes_and_minimum_from_sheet(self):
        self.assertTrue(self.res["warnings"][0].startswith("The sheet covers 2026-11-09 to 2026-11-22 (14 day(s)"))
        mins = {s["name"]: s["min_staff"] for s in self.r["shifts"]}
        self.assertEqual(mins, {"Morning": 2, "Afternoon": 3, "Night": 2})

    def test_real_gap_inside_the_sheet_is_found_and_grouped(self):
        r = copy.deepcopy(self.r)
        alpha = next(m["id"] for m in r["members"] if m["name"] == "ALPHA")
        for d in ("2026-11-10", "2026-11-11", "2026-11-12"):
            r["overrides"][d][alpha] = "OFF"
        groups = R.group_issues(R.check(r, R.generate(r, "2026-11-01", "2026-11-30")))
        self.assertEqual([(g["message"], g["dates"], g["count"]) for g in groups],
                         [("Morning has 1 of 2 people.", ["2026-11-10 to 2026-11-12"], 3)])

    def test_no_swap_on_a_day_without_roster(self):
        a, b = self.r["members"][0]["id"], self.r["members"][1]["id"]
        with self.assertRaises(R.RosterError) as e:
            R.request_swap(copy.deepcopy(self.r), LEAD, a, "2026-11-08", b)
        self.assertIn("no roster for 2026-11-08", e.exception.message)

    def test_hand_edit_on_an_empty_day_creates_data(self):
        r = copy.deepcopy(self.r)
        R.set_override(r, LEAD, r["members"][0]["id"], "2026-11-05", "M")
        day = R.generate(r, "2026-11-05", "2026-11-05")["2026-11-05"]
        self.assertEqual(day[r["members"][0]["id"]]["code"], "M")

    def test_second_import_fills_the_gap(self):
        import openpyxl
        data, _, _ = R.template_file("by_shift", "xlsx", month="2026-11")
        wb = openpyxl.load_workbook(io.BytesIO(data))
        ws = wb["Roster"]
        for row in range(3, ws.max_row + 1):
            for c in range(3, 8):
                ws.cell(row=row, column=c).value = None
            if int(str(ws.cell(row=row, column=2).value).split("-")[0]) <= 8:
                ws.cell(row=row, column=3).value = "ALPHA, BRAVO"
                ws.cell(row=row, column=4).value = "CHARLIE, DELTA, ECHO"
                ws.cell(row=row, column=5).value = "FOXTROT, GOLF"
                ws.cell(row=row, column=6).value = "HOTEL"
        p2 = os.path.join(self.tmp, "early.xlsx")
        wb.save(p2)
        self.st.import_into(LEAD, self.r["id"], p2)
        cal = R.calendar(self.st.get(self.r["id"]), "2026-11-02", "week")
        self.assertTrue(all(d["has_data"] for d in cal["days"]))


class T16WindowsFileHandling(unittest.TestCase):
    """Reported on Windows: importing the roster demo file gave 'Unexpected server error'.
    Cause: the workbook kept the uploaded file open, and Windows cannot delete an open file."""

    DEMO = os.path.join(os.path.dirname(os.path.abspath(R.__file__)), "Demo_Data", "08_Shift_Roster_Nov_Dec_2026.xlsx")

    @staticmethod
    def handles(path):
        n = 0
        for f in os.listdir("/proc/self/fd"):
            try:
                n += os.path.realpath(os.readlink(f"/proc/self/fd/{f}")) == os.path.realpath(path)
            except OSError:
                pass
        return n

    @unittest.skipUnless(os.path.isdir("/proc/self/fd"), "needs /proc to count open files")
    def test_reading_a_workbook_releases_the_file(self):
        tmp = tempfile.mkdtemp()
        try:
            path = os.path.join(tmp, "r.xlsx")
            shutil.copy(self.DEMO, path)
            R.parse_roster_file(path)
            self.assertEqual(self.handles(path), 0)
            R._rows_from_input(path)
            self.assertEqual(self.handles(path), 0)
        finally:
            shutil.rmtree(tmp)

    def test_reads_the_upload_from_memory(self):
        with open(self.DEMO, "rb") as fh:
            p = R.parse_roster_file(fh.read())
        self.assertEqual((p["layout"], len(p["days"]), len(p["people"])), ("by_shift", 61, 10))

    def test_damaged_workbook_is_a_clear_error(self):
        with self.assertRaises(R.RosterError) as e:
            R.parse_roster_file(b"PK\x03\x04 not really a workbook")
        self.assertEqual(e.exception.status, 400)
        self.assertIn("could not be read", e.exception.message)


class T17OnNowWeekdays(unittest.TestCase):
    def test_weekday_only_shift_is_not_running_at_the_weekend(self):
        import datetime as dt
        r = R.default_roster("Boomi", "Boomi", "2026-10-05", "standard")
        r["shifts"] = [{"code": "D", "name": "Day", "start": "09:00", "end": "18:00", "days": [0, 1, 2, 3, 4], "min_staff": 1}]
        r["rotation"] = {"sequence": ["D"] * 7}
        r["cycle"] = {"type": "weekly"}
        r["members"] = [{"id": "m0", "name": "Omar", "role": "Engineer", "user": "omar", "offset": 0}]
        self.assertEqual(R.on_now(r, dt.datetime(2026, 11, 15, 11, 0))["shifts"], [])  # Sunday
        self.assertEqual(R.on_now(r, dt.datetime(2026, 11, 16, 11, 0))["shifts"][0]["people"], ["Omar"])  # Monday


if __name__ == "__main__":
    unittest.main(verbosity=2)
