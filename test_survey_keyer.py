"""Tests for the Survey Keyer app.  Run:  python -m unittest test_survey_keyer -v"""
import http.server
import json
import tempfile
import threading
import unittest
import urllib.error
import urllib.request
import zipfile
from pathlib import Path

import openpyxl

import survey_keyer as A

SPEC = {"title": "Test 2026", "rows": 5, "questions": [
    {"text": "Are you satisfied?", "kind": "single", "options": ["Yes", "No", "Other (please specify)"], "other": True},
    {"text": "What do you like?", "kind": "multi", "options": ["Park", "Swimming pool"]},
    {"text": "Comments", "kind": "open"},
    {"text": "How much do you spend?", "kind": "amount", "options": ["Amount"]},
]}


class Parsing(unittest.TestCase):
    def test_amounts(self):
        for raw, want in [("10 000 zł", 10000), ("1.500", 1500), ("1.500,50", 1500.5), ("12,50", 12.5),
                          ("2,5 tys.", 2500), ("5k", 5000), ("1,500.25", 1500.25), ("-30", -30), ("", None),
                          ("nie wiem", "nie wiem"), ("inf", "inf"), ("nan", "nan"), ("1e999", "1e999"),
                          ("1,500", 1500), ("1,234,567.8", 1234567.8), ("250 PLN", 250), ("0,500", 0.5),
                          ("don't know", "don't know")]:
            self.assertEqual(A.to_amount(raw), want, raw)

    def test_cell_xml_is_always_valid(self):
        x = A.cell_xml(2, 5, None, "a\x0bb<c>&")
        self.assertIn("ab&lt;c&gt;&amp;", x)
        with self.assertRaises(ValueError):
            A.cell_xml(2, 5, None, float("inf"))

    def test_force_recalc(self):
        self.assertIn('fullCalcOnLoad="1"', A.force_recalc('<workbook><sheets/></workbook>') or "")
        self.assertIn('<sheets></sheets><calcPr fullCalcOnLoad="1"/>', A.force_recalc("<workbook><sheets></sheets></workbook>"))
        self.assertIn('fullCalcOnLoad="1"', A.force_recalc('<calcPr calcId="1" fullCalcOnLoad="0"/>'))

    def test_workbook_xml_with_entities_is_rejected(self):
        """XML from an untrusted file: entity declarations (DTD) are rejected, not expanded (XXE, billion laughs)."""
        m = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
        r = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
        with tempfile.TemporaryDirectory() as d:
            path = Path(d) / "bad.xlsx"
            with zipfile.ZipFile(path, "w") as z:
                z.writestr("xl/workbook.xml", f'<?xml version="1.0"?><!DOCTYPE w [<!ENTITY e "x">]><workbook {m} {r}>'
                                              f'<sheets><sheet name="A&e;" sheetId="1" r:id="rId1"/></sheets></workbook>')
                z.writestr("xl/_rels/workbook.xml.rels", '<Relationships xmlns="http://schemas.openxmlformats.org/'
                           'package/2006/relationships"><Relationship Id="rId1" Target="worksheets/s.xml"/></Relationships>')
            with zipfile.ZipFile(path) as z, self.assertRaises(ValueError):
                A.sheet_part(z, "Ax")

    def test_csv_injection(self):
        self.assertEqual(A.csv_safe("=HYPERLINK(1)"), "'=HYPERLINK(1)")
        self.assertEqual(A.csv_safe("Yes"), "Yes")


class Workbook(unittest.TestCase):
    def setUp(self):
        self.dir = tempfile.TemporaryDirectory()
        self.path = Path(self.dir.name) / "survey.xlsx"
        A.build_survey_file(self.path, A.clean_spec(SPEC))

    def tearDown(self):
        self.dir.cleanup()

    def test_created_file_is_ready_for_entry(self):
        s = A.Survey(self.path)
        self.assertTrue(s.schema()["has_config"])
        self.assertEqual([q["kind"] for q in s.questions], ["single", "multi", "open", "amount"])
        self.assertEqual(sorted(s.rows), [1, 2, 3, 4, 5])
        self.assertTrue(s.questions[0]["options"][2]["text"])          # "Other" has a text field
        self.assertIn("Charts", openpyxl.load_workbook(self.path).sheetnames)

    def test_save_load_stats_and_new_row(self):
        s = A.Survey(self.path)
        k = [q["key"] for q in s.questions]
        s.save(1, {k[0]: {"chosen": [2], "texts": {"2": "Bike"}}, k[1]: {"chosen": [0, 1]},
                   k[2]: {"text": "More benches"}, k[3]: {"values": {"0": "1.500 zł"}}})       # Polish amount format
        s.save(2, {k[0]: {"chosen": [0]}, k[3]: {"values": {"0": "250"}}})
        self.assertEqual(s.load(1)[k[0]], {"chosen": [2], "texts": {"2": "Bike"}})
        st = s.stats()
        self.assertEqual(st["filled"], 2)
        self.assertEqual(st["questions"][3]["options"][0]["sum"], 1750)
        self.assertIn("No.,Question,Answer,Count,% of respondents,Sum,Average", s.stats_csv())
        with self.assertRaises(ValueError):
            s.save(3, {k[0]: {"chosen": [0, 1]}})                       # single choice: only one answer
        seq = s.new_row()
        s.refresh()
        self.assertEqual(seq, 6)
        self.assertIn(6, s.rows)
        ws = openpyxl.load_workbook(self.path)["Survey"]
        self.assertEqual(ws.cell(6, 2).value, "=COUNTA(B7:B12)")        # the total's range now covers the new row
        self.assertEqual(ws.cell(12, 2).border.left.style, "thin")      # the new row has borders like the previous one
        sheet = zipfile.ZipFile(self.path).read("xl/worksheets/sheet1.xml").decode()
        self.assertRegex(sheet, r'<c r="D6"[^>]*><f>COUNTA\(D7:D1[12]\)</f><v>1</v>')   # cached COUNTA result

    def test_file_changed_outside_app(self):
        s = A.Survey(self.path)
        wb = openpyxl.load_workbook(self.path)
        wb["Survey"].cell(7, 2, 1)                                      # someone typed an answer in Excel
        wb.save(self.path)
        s.refresh()
        self.assertIn(1, s.filled_ids())

    def test_wizard_text_never_becomes_formula(self):
        spec = A.clean_spec({**SPEC, "title": "=WEBSERVICE(\"http://x\")",
                             "questions": [{"text": "Q", "kind": "single", "options": ["=1+1", "No"]}]})
        A.build_survey_file(self.path.with_name("f.xlsx"), spec)
        ws = openpyxl.load_workbook(self.path.with_name("f.xlsx"))["Survey"]
        self.assertEqual((ws["A1"].value, ws["A1"].data_type), ("=WEBSERVICE(\"http://x\")", "s"))
        self.assertEqual((ws.cell(5, 2).value, ws.cell(5, 2).data_type), ("=1+1", "s"))

    def test_zip_bomb_is_rejected_before_loading(self):
        """A file that is too large once unpacked never reaches openpyxl (it would load everything into memory)."""
        path = self.path.with_name("three.xlsx")
        q = {"kind": "single", "options": ["Yes", "No"]}
        A.build_survey_file(path, A.clean_spec({"title": "T", "rows": 3, "questions": [
            {**q, "text": "A"}, {**q, "text": "B"}, {**q, "text": "C"}]}))
        path.with_name("three.survey.json").unlink()                    # forces a check of the file's contents
        self.assertTrue(A.looks_like_survey(path))
        old = A.MAX_PART
        A.MAX_PART = 2000                                               # a small limit instead of a multi-gigabyte file
        try:
            self.assertFalse(A.looks_like_survey(path))
            with self.assertRaisesRegex(ValueError, "once unpacked"):
                A.Survey(path)
        finally:
            A.MAX_PART = old

    def test_bad_config_does_not_block_opening(self):
        self.path.with_name("survey.survey.json").write_text("{broken", encoding="utf-8")
        s = A.Survey(self.path)
        self.assertFalse(s.schema()["has_config"])
        self.assertTrue(s.schema()["warning"])

    def test_layout_is_detected_with_english_and_polish_headers(self):
        """Workbooks the app writes (English headers) and older Polish workbooks are both understood."""
        self.path.with_name("survey.survey.json").unlink()              # no settings: everything comes from the sheet
        s = A.Survey(self.path)
        self.assertEqual(s.layout, {"sheet": "Survey", "header_row": 4, "option_row": 5, "first_row": 7})
        self.assertEqual(sorted(s.rows), [1, 2, 3, 4, 5])
        self.assertEqual(s.questions[2]["kind"], "open")                # "(open question)"
        self.assertTrue(s.questions[0]["options"][2]["text"])          # "Other (please specify)" gets a text field
        for no, total, open_label in (("No.", "Total", "(open question)"), ("L.p.", "Razem", "(pytanie otwarte)")):
            wb = openpyxl.Workbook()                                    # no merged cells and no surveys yet
            ws = wb.active
            for ref, value in {"A1": "Title", "A3": no, "B3": "Comments", "B4": open_label, "C3": "Age",
                               "C4": "under 30", "D4": "30 or over", "A5": total}.items():
                ws[ref] = value
            layout = A.detect_layout(wb)
            self.assertEqual((layout["header_row"], layout["option_row"], layout["first_row"]), (3, 4, 6), no)
            self.assertEqual([q["kind"] for q in A.parse_questions(ws, layout)], ["open", "single"], no)


class Demo(unittest.TestCase):
    def test_demo_survey_is_ready_and_filled(self):
        with tempfile.TemporaryDirectory() as d:
            path = A.demo_file(d)
            s = A.Survey(path)
            st = s.stats()
            self.assertTrue(s.schema()["has_config"])
            self.assertEqual(st["filled"], A.DEMO_FILLED)
            self.assertEqual([q["kind"] for q in s.questions], ["single", "multi", "open", "amount", "single"])
            mtime = path.stat().st_mtime
            self.assertEqual(A.demo_file(d), path)                          # second call: the same file, not overwritten
            self.assertEqual(path.stat().st_mtime, mtime)

    def test_demo_switch(self):
        self.assertTrue(A.is_demo(["--demo"]))
        self.assertFalse(A.is_demo(["file.xlsx"]) and not A.os.environ.get("DEMO"))


class Server(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.dir = tempfile.TemporaryDirectory()
        A.STATE_FILE = Path(cls.dir.name) / "state.json"                # leave the user's settings alone
        cls.path = Path(cls.dir.name) / "a.xlsx"
        A.build_survey_file(cls.path, A.clean_spec(SPEC))
        cls.app = A.App()
        cls.srv = A.start_server(cls.app, 18765)
        cls.base = f"http://127.0.0.1:{cls.srv.server_address[1]}"
        threading.Thread(target=cls.srv.serve_forever, daemon=True).start()

    @classmethod
    def tearDownClass(cls):
        cls.srv.shutdown()
        cls.srv.server_close()
        cls.dir.cleanup()

    def call(self, path, body=None, headers=None, token=True):
        h = {"Content-Type": "application/json", **({"X-Survey-Keyer-Token": self.app.token} if token else {}), **(headers or {})}
        req = urllib.request.Request(self.base + path, data=None if body is None else json.dumps(body).encode(), headers=h)
        try:
            with urllib.request.urlopen(req, timeout=10) as r:
                return r.status, json.loads(r.read() or b"{}")
        except urllib.error.HTTPError as e:
            return e.code, json.loads(e.read() or b"{}")

    def test_security(self):
        self.assertEqual(self.call("/api/open", {"path": str(self.path)}, token=False)[0], 403)
        self.assertEqual(self.call("/api/files", headers={"Host": "evil.example:80"})[0], 403)
        self.assertEqual(self.call("/api/files", {}, headers={"Origin": "https://evil.example"})[0], 403)
        self.assertEqual(self.call("/api/hello", token=False)[0], 200)
        with urllib.request.urlopen(self.base + "/") as r:
            self.assertIn("frame-ancestors 'none'", r.headers["Content-Security-Policy"])

    def test_page_does_not_reveal_token(self):
        """Any process on the computer (even another user's) can fetch the page, so it must not contain the key."""
        with urllib.request.urlopen(self.base + "/?t=anything") as r:
            self.assertNotIn(self.app.token, r.read().decode())

    def test_hello_proves_identity_without_revealing_token(self):
        code, out = self.call("/api/hello?n=abc", token=False)
        self.assertEqual(code, 200)
        self.assertNotIn(self.app.token, json.dumps(out))
        self.assertIs(A.port_owner(self.srv.server_address[1]), True)

    def test_port_owner_rejects_impostor_and_relay(self):
        """A foreign program on the port does not get the key: neither by replaying an old answer nor by relaying ours."""
        real = self.base

        class Fake(http.server.BaseHTTPRequestHandler):
            def do_GET(self):
                if self.path.startswith("/relay"):
                    with urllib.request.urlopen(real + self.path[len("/relay"):]) as r:
                        body = r.read()
                else:
                    body = json.dumps({"app": A.APP_ID}).encode()
                self.send_response(200)
                self.send_header("Content-Length", str(len(body)))
                self.end_headers()
                self.wfile.write(body)

            def log_message(self, *a):
                pass

        for mode in ("old", "relay"):
            class H(Fake):
                def do_GET(self, mode=mode):
                    if mode == "relay":
                        self.path = "/relay" + self.path
                    Fake.do_GET(self)
            fake = http.server.ThreadingHTTPServer(("127.0.0.1", 0), H)
            threading.Thread(target=fake.serve_forever, daemon=True).start()
            try:
                self.assertFalse(A.port_owner(fake.server_address[1]), mode)
            finally:
                fake.shutdown()
                fake.server_close()

    def test_create_ignores_path_sent_by_page(self):
        target = Path(self.dir.name) / "from_outside.xlsx"
        self.app.ask_save = lambda title, folder: None                  # the user cancelled the save dialog
        code, out = self.call("/api/create", {**SPEC, "path": str(target)})
        self.assertEqual((code, out), (200, {"cancelled": True}))
        self.assertFalse(target.exists())

    def test_flow_and_conflict(self):
        code, s = self.call("/api/open", {"path": str(self.path)})
        self.assertEqual(code, 200)
        h = {"X-Survey-Keyer-Path": urllib.request.quote(s["path"]), "X-Survey-Keyer-Rev": s["rev"]}
        self.assertEqual(self.call("/api/progress", headers=h)[0], 200)
        self.assertEqual(self.call("/api/progress", headers={**h, "X-Survey-Keyer-Rev": "stale"})[0], 409)
        k = s["questions"][0]["key"]
        self.assertEqual(self.call("/api/save", {"seq": 3, "answers": {k: {"chosen": [1]}}}, headers=h)[0], 200)
        self.assertEqual(self.call("/api/save", {"seq": "x", "answers": {}}, headers=h)[0], 400)
        self.app.ask_save = lambda title, folder: str(self.path)       # an existing file was picked in the save dialog
        code, out = self.call("/api/create", SPEC)
        self.assertEqual(code, 400)                                     # does not overwrite the existing file
        self.assertIn("already exists", out["error"])


if __name__ == "__main__":
    unittest.main()
