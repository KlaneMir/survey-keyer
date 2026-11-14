"""Security audit: regression tests for attacks from the audit list (the numbers match the audit report).

Each test fails when Survey Keyer is vulnerable to that attack. Items that do not apply have no test; the report
explains why (no HTTP/2, SAML, OAuth, GraphQL, 2FA, JWT, reverse proxy, cache, database, encryption, WebSocket
or user accounts in this program).

Run:  python -m unittest test_security_audit -v
"""
import ast
import gc
import http.client
import json
import os
import re
import subprocess
import sys
import tempfile
import threading
import unittest
import urllib.error
import urllib.request
import zipfile
from pathlib import Path

import openpyxl

import survey_keyer as A

HERE = Path(__file__).resolve().parent
SOURCE = (HERE / "survey_keyer.py").read_text(encoding="utf-8")
SPEC = {"title": "Audit", "rows": 5, "questions": [
    {"text": "Are you happy?", "kind": "single", "options": ["Yes", "No", "Other (please specify)"], "other": True},
    {"text": "What do you like?", "kind": "multi", "options": ["Park", "Pool"]},
    {"text": "Comments", "kind": "open"},
]}


def run_timed(code, *args, timeout=30):
    """Runs code in a separate Python (a runaway regex cannot hang the test run); returns (seconds, label)."""
    env = {**os.environ, "PYTHONDONTWRITEBYTECODE": "1", "PYTHONIOENCODING": "utf-8"}
    r = subprocess.run([sys.executable, "-c", code, *map(str, args)], capture_output=True, text=True,
                       timeout=timeout, cwd=HERE, env=env)
    if r.returncode:
        raise AssertionError(r.stderr)
    took, label = r.stdout.strip().split(" ", 1)
    return float(took), label


class Server(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.dir = tempfile.TemporaryDirectory()
        A.STATE_FILE = Path(cls.dir.name) / "state.json"                # leave the user's settings alone
        cls.path = Path(cls.dir.name) / "a.xlsx"
        A.build_survey_file(cls.path, A.clean_spec(SPEC))
        cls.app = A.App()
        cls.srv = A.start_server(cls.app, 18865)
        cls.port = cls.srv.server_address[1]
        cls.base = f"http://127.0.0.1:{cls.port}"
        threading.Thread(target=cls.srv.serve_forever, daemon=True).start()

    @classmethod
    def tearDownClass(cls):
        cls.srv.shutdown()
        cls.srv.server_close()
        cls.dir.cleanup()

    def call(self, path, body=None, headers=None, token=True):
        h = {"Content-Type": "application/json",
             **({"X-Survey-Keyer-Token": self.app.token} if token is True else {}), **(headers or {})}
        req = urllib.request.Request(self.base + path, data=None if body is None else json.dumps(body).encode(), headers=h)
        try:
            with urllib.request.urlopen(req, timeout=10) as r:
                return r.status, json.loads(r.read() or b"{}")
        except urllib.error.HTTPError as e:
            return e.code, json.loads(e.read() or b"{}")

    # ---------- 2. Long input / truncation ----------

    def test_02_token_is_compared_in_full(self):
        t = self.app.token
        for wrong in (t + "x", t[:-1], t[:20], t + "A" * 200):
            code, _ = self.call("/api/files", headers={"X-Survey-Keyer-Token": wrong}, token=False)
            self.assertEqual(code, 403, wrong)

    def test_02_oversized_request_is_refused_before_reading(self):
        conn = http.client.HTTPConnection("127.0.0.1", self.port, timeout=10)
        conn.putrequest("POST", "/api/files")
        conn.putheader("Content-Type", "application/json")
        conn.putheader("Content-Length", str(A.MAX_BODY + 1))
        conn.endheaders()                                               # the body is never sent
        r = conn.getresponse()
        self.assertEqual(r.status, 400)
        conn.close()

    def test_02_answer_longer_than_an_excel_cell_is_refused(self):
        code, s = self.call("/api/open", {"path": str(self.path)})
        self.assertEqual(code, 200)
        h = {"X-Survey-Keyer-Path": urllib.request.quote(s["path"]), "X-Survey-Keyer-Rev": s["rev"]}
        key = s["questions"][2]["key"]
        code, out = self.call("/api/save", {"seq": 3, "answers": {key: {"text": "x" * (A.MAX_CELL + 1)}}}, headers=h)
        self.assertEqual(code, 400, out)

    # ---------- 11. Spoofed client IP / forwarded headers ----------

    def test_11_forwarded_headers_do_not_bypass_host_and_origin_checks(self):
        spoof = {"X-Forwarded-Host": f"127.0.0.1:{self.port}", "X-Forwarded-For": "127.0.0.1",
                 "X-Real-IP": "127.0.0.1", "Forwarded": f"for=127.0.0.1;host=127.0.0.1:{self.port}"}
        self.assertEqual(self.call("/api/files", headers={"Host": "evil.example", **spoof})[0], 403)
        self.assertEqual(self.call("/api/files", {}, headers={"Origin": "https://evil.example", **spoof})[0], 403)

    def test_11_app_never_decides_on_client_ip_headers(self):
        low = SOURCE.lower()
        for marker in ("x-forwarded", "x-real-ip", "client_address"):
            self.assertNotIn(marker, low)

    # ---------- 13. Cryptographic randomness ----------

    def test_13_token_and_nonce_are_long_and_fresh(self):
        self.assertRegex(self.app.token, r"^[A-Za-z0-9_-]{43,}$")      # 32+ random bytes
        nonces = set()
        for _ in range(2):
            with urllib.request.urlopen(self.base + "/") as r:
                nonces.add(re.search(r"'nonce-([^']+)'", r.headers["Content-Security-Policy"]).group(1))
        self.assertEqual(len(nonces), 2, "the CSP nonce must change on every page load")
        self.assertTrue(all(len(n) >= 22 for n in nonces))

    # ---------- 17. Template injection ----------

    def test_17_survey_text_never_reaches_the_server_rendered_page(self):
        """The page is a fixed string with only the nonce filled in; survey text arrives as JSON and is escaped."""
        payload = "{{7*7}} __NONCE__ ${7*7} </script><script>alert(1)</script>"
        path = Path(self.dir.name) / "ssti.xlsx"
        A.build_survey_file(path, A.clean_spec({"title": payload, "rows": 2, "questions": [
            {"text": payload, "kind": "single", "options": [payload, "No"]}]}))
        self.assertEqual(self.call("/api/open", {"path": str(path)})[0], 200)
        with urllib.request.urlopen(self.base + "/") as r:
            page = r.read().decode()
        for part in ("{{7*7}}", "${7*7}", "alert(1)"):
            self.assertNotIn(part, page)
        self.assertNotIn("__NONCE__", page)                             # the only placeholder is always filled in


class StaticChecks(unittest.TestCase):
    # ---------- 13. Cryptographic randomness ----------

    def test_13_secrets_come_from_the_secrets_module(self):
        tree = ast.parse(SOURCE)
        fn = next(n for n in ast.walk(tree) if isinstance(n, ast.FunctionDef) and n.name == "app_token")
        self.assertIn("secrets.token_urlsafe(32)", ast.unparse(fn))
        # the random module only makes up demo answers
        for f in (n for n in ast.walk(tree) if isinstance(n, ast.FunctionDef)):
            uses = [x for x in ast.walk(f) if isinstance(x, ast.Attribute) and getattr(x.value, "id", "") == "random"]
            if uses:
                self.assertEqual(f.name, "demo_file", f"random used in {f.name}()")
        # Math.random only names a browser tab (not a secret: the server also requires the token)
        lines = [line for line in SOURCE.splitlines() if "Math.random" in line]
        self.assertTrue(lines and all("TAB = Math.random()" in line for line in lines), lines)

    # ---------- 14. Timing attacks ----------

    def test_14_secret_comparisons_are_constant_time(self):
        bad = []
        for n in ast.walk(ast.parse(SOURCE)):
            if not (isinstance(n, ast.Compare) and any(isinstance(op, (ast.Eq, ast.NotEq)) for op in n.ops)):
                continue
            sides = [n.left, *n.comparators]
            if any(isinstance(s, ast.Constant) for s in sides):
                continue
            names = {x.id.lower() for s in sides for x in ast.walk(s) if isinstance(x, ast.Name)}
            names |= {x.attr.lower() for s in sides for x in ast.walk(s) if isinstance(x, ast.Attribute)}
            if any(w in name for name in names for w in ("secret", "token", "proof", "nonce", "password")):
                bad.append(f"line {n.lineno}: {ast.unparse(n)}")
        self.assertEqual(bad, [])


class ReDoS(unittest.TestCase):
    # ---------- 7. ReDoS ----------

    def test_07_parsers_of_typed_answers_and_names_stay_fast(self):
        code = r'''
import time, survey_keyer as A
N = 100_000
cases = [("to_amount", A.to_amount, "1" * N + "x"), ("to_amount", A.to_amount, "1.000" * (N // 5) + "x"),
         ("to_amount", A.to_amount, "1,000" * (N // 5) + ".x"), ("to_amount", A.to_amount, "1" * N + "tys" + "x"),
         ("to_amount", A.to_amount, " " * N + "1"), ("to_amount", A.to_amount, "zł" * N),
         ("clean", A.clean, " \t" * N + "x"), ("safe_filename", A.safe_filename, "<>" * N + " ." * N),
         ("csv_safe", A.csv_safe, "=" * N)]
worst, which = 0.0, ""
for name, fn, arg in cases:
    t = time.perf_counter()
    try:
        fn(arg)
    except Exception:
        pass
    if time.perf_counter() - t > worst:
        worst, which = time.perf_counter() - t, name
print(f"{worst:.3f} {which}")
'''
        took, which = run_timed(code)
        self.assertLess(took, 2.0, f"slowest case ({which}) took {took} s")

    def test_07_hostile_workbook_text_is_read_quickly(self):
        """End to end: a workbook with 30,000-character headers and answer labels full of spaces and brackets."""
        with tempfile.TemporaryDirectory() as d:
            path = Path(d) / "hostile.xlsx"
            A.build_survey_file(path, A.clean_spec(SPEC))
            path.with_name("hostile.survey.json").unlink()              # force the layout detection
            wb = openpyxl.load_workbook(path)
            ws = wb["Survey"]
            hostile = " " * 15000 + "(" + " " * 15000 + "x"
            ws.cell(4, 1).value = "No" + " " * 30000 + "."
            for col in range(2, 6):
                ws.cell(5, col).value = hostile
            wb.save(path)
            code = r'''
import sys, time, survey_keyer as A
t = time.perf_counter()
try:
    A.Survey(sys.argv[1])
except Exception:
    pass
print(f"{time.perf_counter() - t:.3f} Survey()")
'''
            took, _ = run_timed(code, path)
        self.assertLess(took, 5.0, f"opening the hostile workbook took {took} s")

    def test_07_open_question_pattern_has_no_catastrophic_backtracking(self):
        """OPEN_RE on its own (today it only sees text that clean() has already collapsed)."""
        code = r'''
import time, survey_keyer as A
N = 100_000
cases = [("1,000 spaces", " " * 1000 + "x"), ("100,000 spaces", " " * N + "x"),
         ("bracket + spaces", "(" + " " * N + "x"), ("marker + spaces", "open question" + " " * N + "x"),
         ("brackets", "( " * (N // 2) + "x"), ("marker + brackets", "pytanie otwarte" + " )" * (N // 2) + "x")]
worst, which = 0.0, ""
for name, s in cases:
    t = time.perf_counter()
    A.OPEN_RE.sub("", s)
    if time.perf_counter() - t > worst:
        worst, which = time.perf_counter() - t, name
print(f"{worst:.3f} {which}")
'''
        try:
            took, which = run_timed(code, timeout=20)
        except subprocess.TimeoutExpired:
            self.fail("OPEN_RE ran for over 20 s on a long input (catastrophic backtracking)")
        self.assertLess(took, 1.0, f"OPEN_RE took {took} s ({which})")

    def test_07_open_question_marker_is_still_removed_the_same_way(self):
        for raw, want in [("Comments (open question)", "Comments"), ("Notes (pytanie otwarte)", "Notes"),
                          ("Notes ( Pytanie Otwarte ) ", "Notes"), ("Remarks open question", "Remarks"),
                          ("(open question)", ""), ("Open question", ""), ("Questions opened", "Questions opened")]:
            self.assertEqual(A.OPEN_RE.sub("", raw).strip(), want, raw)


class XXE(unittest.TestCase):
    # ---------- 18. XXE ----------

    MARKER = "AUDIT-SECRET-7f3a"

    def hostile_workbook(self, d):
        """A real workbook whose sheet (read by openpyxl) and shared strings (read by the app's own XML reader)
        declare an external entity pointing at a local file."""
        secret = Path(d) / "secret.txt"
        secret.write_text(self.MARKER, encoding="utf-8")
        clean = Path(d) / "clean.xlsx"
        wb = openpyxl.Workbook()
        wb.active["A1"] = "PLACEHOLDER"
        wb.save(clean)
        path = Path(d) / "xxe.xlsx"

        def with_dtd(text, root):
            dtd = f'<!DOCTYPE {root} [<!ENTITY xxe SYSTEM "{secret.as_uri()}">]>'
            m = re.match(r"<\?xml[^>]*\?>", text)
            return text[:m.end()] + dtd + text[m.end():] if m else dtd + text

        with zipfile.ZipFile(clean) as src, zipfile.ZipFile(path, "w") as out:
            for item in src.infolist():
                data = src.read(item.filename)
                if item.filename == "xl/worksheets/sheet1.xml":
                    data = with_dtd(data.decode("utf-8"), "worksheet").replace("PLACEHOLDER", "&xxe;").encode("utf-8")
                out.writestr(item, data)
            sst = ('<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">'
                   "<si><t>&xxe;</t></si></sst>")
            out.writestr("xl/sharedStrings.xml", with_dtd(sst, "sst"))
        return path

    def test_18_external_entities_are_never_resolved(self):
        # openpyxl keeps a rejected file open until garbage collection, so collect before the folder is removed
        self.addCleanup(gc.collect)
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as d:
            path = self.hostile_workbook(d)
            with zipfile.ZipFile(path) as z:
                self.assertIn("&xxe;", z.read("xl/worksheets/sheet1.xml").decode())
                self.assertIn("<!DOCTYPE", z.read("xl/sharedStrings.xml").decode())
                try:
                    seen = repr(A.shared_strings(z))                    # the app's own XML reader
                except Exception as e:  # noqa: BLE001 - refusing the file is a correct answer
                    seen = repr(e)
            self.assertNotIn(self.MARKER, seen)
            try:
                seen = repr(openpyxl.load_workbook(path).active["A1"].value)   # the parser the app opens files with
            except Exception as e:  # noqa: BLE001
                seen = repr(e)
            self.assertNotIn(self.MARKER, seen)
            try:
                seen = json.dumps(A.Survey(path).schema())
            except Exception as e:  # noqa: BLE001
                seen = repr(e)
            self.assertNotIn(self.MARKER, seen)


if __name__ == "__main__":
    unittest.main()
