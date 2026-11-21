#!/usr/bin/env python3
"""Security probes against a running MSP Ops Center (demo data). Run: python -m unittest test_msp_security -v
They attack the server the way an outsider or a curious insider would; every probe must be refused."""
import contextlib
import http.client
import io
import json
import os
import shutil
import tempfile
import threading
import unittest


class T01Probes(unittest.TestCase):
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
        cls.port = cls.srv.server_address[1]
        st, h, b = cls.req("POST", "/api/auth/login", json.dumps({"username": "sneha.iyer", "password": "Demo-Ops-2026!"}),
                           {"Content-Type": "application/json"})
        cls.cookie, cls.sid, cls.csrf = h.get("Set-Cookie", ""), h.get("Set-Cookie", "").split(";")[0], json.loads(b)["csrf"]

    @classmethod
    def tearDownClass(cls):
        cls.srv.shutdown()
        cls.srv.server_close()
        shutil.rmtree(cls.tmp)

    @classmethod
    def req(cls, method, path, body=None, headers=None):
        c = http.client.HTTPConnection("127.0.0.1", cls.port, timeout=15)
        c.request(method, path, body=body, headers=headers or {})
        r = c.getresponse()
        data = r.read()
        c.close()
        return r.status, dict(r.getheaders()), data

    def test_security_headers(self):
        st, h, _ = self.req("GET", "/")
        csp = h["Content-Security-Policy"]
        self.assertIn("default-src 'self'", csp)
        self.assertNotIn("https:", csp)            # nothing loads from outside
        self.assertNotIn("'unsafe-inline'", csp.split("script-src")[1].split(";")[0])  # no inline scripts
        self.assertIn("frame-ancestors 'self'", csp)
        self.assertEqual(h["X-Frame-Options"], "SAMEORIGIN")
        self.assertEqual(h["X-Content-Type-Options"], "nosniff")

    def test_session_cookie(self):
        self.assertIn("HttpOnly", self.cookie)
        self.assertIn("SameSite", self.cookie)

    def test_no_access_without_signing_in_or_csrf(self):
        self.assertEqual(self.req("GET", "/api/state")[0], 401)
        body, ct = json.dumps({"Title": "x"}), {"Content-Type": "application/json", "Cookie": self.sid}
        self.assertEqual(self.req("POST", "/api/table/Runbook", body, ct)[0], 403)
        self.assertEqual(self.req("POST", "/api/table/Runbook", body, dict(ct, **{"X-CSRF-Token": "wrong"}))[0], 403)

    def test_path_traversal(self):
        for p in ("/static/../msp_users.json", "/static/%2e%2e/msp_users.json", "/static/..%2fmsp_users.json",
                  "/static/%2e%2e%2f%2e%2e%2fetc/passwd", "/app/../msp_runbook.py", "/app/%2e%2e/msp_users.json",
                  "/static/fonts/..\\..\\msp_users.json", "/static//etc/passwd", "/app/assets/../../../msp_settings.ini"):
            st, _, data = self.req("GET", p)
            self.assertFalse(st == 200 and (b"password" in data.lower() or b"import " in data or b"root:" in data), p)

    def test_password_guessing_is_throttled(self):
        codes = {self.req("POST", "/api/auth/login", json.dumps({"username": "priya.sharma", "password": f"wrong{i}"}),
                          {"Content-Type": "application/json"})[0] for i in range(12)}
        self.assertIn(429, codes)

    def test_huge_request_is_refused(self):
        try:
            st = self.req("POST", "/api/ai/ask", b'{"q":"' + b"a" * (30 * 1024 * 1024) + b'"}',
                          {"Content-Type": "application/json", "Cookie": self.sid, "X-CSRF-Token": self.csrf})[0]
        except (BrokenPipeError, ConnectionResetError):
            st = 413  # the server closed the connection before reading the upload
        self.assertIn(st, (400, 413))

    def test_made_up_team_header_is_ignored(self):
        st, _, data = self.req("GET", "/api/state", headers={"Cookie": self.sid, "X-Team": "team_doesnotexist"})
        self.assertNotEqual(json.loads(data)["team"], "team_doesnotexist")


if __name__ == "__main__":
    unittest.main(verbosity=2)
