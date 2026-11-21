#!/usr/bin/env python3
"""Assistant quality test: real questions from every level (operator, engineer, team lead, admin), each with what a
correct answer must contain and what it must NOT reveal (team isolation). Runs on the demo data.
Run: python -m unittest test_msp_ai_eval -v      The same questions are in Demo_Data/Assistant_Test_Questions.xlsx.
To add a question, add a line to Q: it is checked the next time the tests run."""
import contextlib
import http.cookiejar
import io
import json
import os
import shutil
import tempfile
import threading
import unittest
import urllib.error
import urllib.request

# id, area, level, sign in as, question, must contain (all, case-insensitive), must NOT contain, what a correct answer is
Q = [
 ("T01","Teams","Team lead","vikram.rao","Who are the team members of the Boomi team?",["Boomi team has 11 people","Omar Khan"],[],"Lists the 11 Boomi members by role, and the customers served"),
 ("T02","Teams","Engineer","sneha.iyer","Who leads the DBA team?",["Priya Sharma","Divya Menon"],[],"Names both DBA leads"),
 ("T03","Teams","Engineer","rahul.verma","Which teams am I in?",["NOC","Windows"],[],"Both of Rahul's teams"),
 ("T04","Isolation","Engineer","sneha.iyer","Who is in the Boomi team?",["not in the Boomi team"],["omar.khan","Omar Khan"],"Refuses: Sneha is not in Boomi; no member names"),
 ("T05","Teams","Admin","admin","Who is in the NOC team?",["NOC team has","Anjali Nair"],[],"Admins see every team"),
 ("C01","Customers","Team lead","priya.sharma","Which teams serve Harbourline Retail?",["Unix SA","Primary team"],[],"All five teams, Unix SA as primary, the 3 services"),
 ("C02","Customers","Team lead","karthik.reddy","What services does Globex have?",["Retail data platform","Cloud and web"],[],"Globex's two services"),
 ("C03","Customers","Engineer","omar.khan","Who is the primary team for Bluepeak Logistics?",["Boomi"],[],"Boomi is primary for Bluepeak"),
 ("R01","Roster","Operator","tom.berg","Who is on shift now?",["shift"],[],"The running shift and its people (depends on the time of day)"),
 ("R02","Roster","Engineer","sneha.iyer","Who is on call this week?",["Primary"],[],"DBA on-call: Primary and Secondary names"),
 ("R03","Roster","Engineer","sneha.iyer","When is my next shift?",["Sneha"],[],"Sneha's upcoming shifts from the DBA roster"),
 ("R04","Roster","Engineer","suresh.babu","Is 27 November a holiday?",["Company foundation day"],[],"The company holiday on 27 Nov"),
 ("M01","CMDB","Engineer","suresh.babu","What is the IP address of acme-lnx-db01?",["10.10.0.14"],[],"The host's IP (Unix SA owns it)"),
 ("M02","Isolation","Engineer","sneha.iyer","How much memory does acme-lnx-db01 have?",[],["512 GB"],"DBA sees only the shared core: memory must NOT be revealed"),
 ("M03","CMDB","Engineer","sneha.iyer","Who owns acme-db-ora01?",["DBA"],[],"Owner team DBA"),
 ("M04","CMDB","Engineer","dinesh.yadav","What special care is needed for acme-fw01?",["passive"],[],"HA pair: upgrade the passive unit first"),
 ("M05","Isolation","Engineer","suresh.babu","acme-db-test01",[],["acme-db-test01 ("],"DBA's team-only test database must not appear for Unix SA"),
 ("K01","Runbook","Engineer","sneha.iyer","How do I fix the archiver stuck when the FRA is full?",["credential"],[],"The runbook fix: re-register the credential, back up and purge archive logs"),
 ("K02","Runbook","Engineer","omar.khan","EDI listener stops after Java heap exhaustion",["heap"],[],"Raise the heap to 4 GB and restart the listener"),
 ("K03","Runbook","Team lead","priya.sharma","What are the known issues for Acme Corp?",["Known"],[],"Acme known issues the DBA team can see"),
 ("I01","Incidents","Team lead","priya.sharma","Show open P1 incidents",["P1"],[],"Open P1 incidents (or a clear 'none')"),
 ("P01","Plans","Engineer","sneha.iyer","What changes are planned this week?",["CHG-"],[],"Upcoming changes the DBA team can see"),
 ("P02","Plans","Engineer","sneha.iyer","What is the rollback for CHG-4101?",["Flashback"],[],"Rollback steps of the shared DB upgrade"),
 ("P03","Isolation","Engineer","omar.khan","What is the rollback for CHG-4102?",[],["CHG-4102"],"Omar (Boomi) cannot see Acme patching"),
 ("B01","Bulletins","Operator","ana.ito","Are there any active bulletins?",["OpenSSH"],[],"Active bulletins, critical OpenSSH first"),
 ("H01","How-to","Engineer","sneha.iyer","How do I request leave?",["Request leave"],[],"Steps on the Shift roster page"),
 ("H02","How-to","Engineer","suresh.babu","How do I share a CI with another team?",["Sharing","Shared with teams"],[],"Steps: Owner team, Sharing, Shared with teams, Shared fields"),
 ("H03","How-to","Team lead","anjali.nair","How do I add a new engineer to my team?",["Add member"],[],"Teams page, Add member, new user"),
 ("H04","How-to","Engineer","rahul.verma","How do I swap a shift with someone?",["Swap or cover"],[],"Swap or cover steps, partner accepts, approver decides"),
 ("H05","How-to","Team lead","priya.sharma","Where do I import a roster sheet?",["Import sheet"],[],"Template then Import sheet"),
 ("H06","How-to","Admin","admin","How do I set up the local AI model?",["Ollama"],[],"Install Ollama, Admin settings, test the server"),
 ("A01","Isolation","Admin","admin","What is in the DBA shift log?",[],["RMAN backups"],"Admin sees no team shift logs"),
]


class T01AssistantQuality(unittest.TestCase):
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
        cls.sessions = {}

    @classmethod
    def tearDownClass(cls):
        cls.srv.shutdown()
        cls.srv.server_close()
        shutil.rmtree(cls.tmp)

    def session(self, user):
        if user not in self.sessions:
            op = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(http.cookiejar.CookieJar()))
            r = urllib.request.Request(self.base + "/api/auth/login", data=json.dumps({"username": user, "password": "Demo-Ops-2026!"}).encode(),
                                       method="POST", headers={"Content-Type": "application/json"})
            self.sessions[user] = (op, json.loads(op.open(r).read())["csrf"])
        return self.sessions[user]

    def call(self, user, method, path, body=None):
        op, csrf = self.session(user)
        r = urllib.request.Request(self.base + path, data=json.dumps(body).encode() if body is not None else None, method=method,
                                   headers={"Content-Type": "application/json", "X-CSRF-Token": csrf})
        try:
            return 200, json.loads(op.open(r).read())
        except urllib.error.HTTPError as e:
            return e.code, json.loads(e.read())

    def test_every_question(self):
        failures = []
        for qid, area, level, user, q, must, never, desc in Q:
            a = self.call(user, "POST", "/api/ai/ask", {"question": q})[1]["answer"]
            text = (a["text"] + " " + " ".join(s["title"] + " " + s.get("where", "") for s in a["sources"])).lower()
            miss = [m for m in must if m.lower() not in text]
            leak = [n for n in never if n.lower() in text]
            if miss or leak:
                failures.append(f"{qid} [{user}] {q}: missing {miss} leaked {leak}")
        self.assertEqual(failures, [], "\n".join(failures))

    def test_incident_about_a_hidden_ci_cannot_be_changed(self):
        rows = self.call("sneha.iyer", "GET", "/api/state")[1]["tables"]["Incidents"]["rows"]
        hidden = next(r for r in rows if r["Asset"] == "acme-db-test01")       # DBA sees it
        st, _ = self.call("suresh.babu", "PUT", f"/api/table/Incidents/{hidden['Record_ID']}", {"Summary": "x"})
        self.assertEqual(st, 404)                                              # Unix SA: does not exist for you


if __name__ == "__main__":
    unittest.main(verbosity=2)
