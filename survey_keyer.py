#!/usr/bin/env python3
"""
Survey Keyer - typing paper survey answers into Excel.

Run:        SurveyKeyer.exe   or   python survey_keyer.py ["file.xlsx"]   (requires: pip install -r requirements.txt)
Demo mode:  python survey_keyer.py --demo   (or DEMO=true): a sample survey with made-up answers

Sheet layout (detected automatically, can be corrected in the "Question settings" tab):
  question row (merged cells above the answers) -> answer row -> totals row
  -> one row per survey, its number in column A (header "No.", or "L.p." in Polish workbooks).

Saving changes ONLY the answer cells (charts, formulas and formatting stay intact).
Before the first save in a session, a copy of the file is made in the "backups" subfolder.
Question settings are saved next to the file: <name>.survey.json.
The "New survey" wizard creates an Excel file in this layout (with a charts sheet) and a printable form.

Security: the server listens only on 127.0.0.1, checks the Host header (DNS rebinding)
and requires a secret token in every /api/ request (other pages in the browser cannot change anything).
"""
import csv
import hashlib
import hmac
import html
import io
import json
import math
import os
import re
import secrets
import shutil
import sys
import threading
import time
import urllib.request
import warnings
import webbrowser
import zipfile
from datetime import datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, unquote, urlparse

warnings.filterwarnings("ignore")
try:
    import defusedxml.ElementTree as ET   # untrusted XML: rejects DTDs and entities (XXE, billion laughs); openpyxl uses it too
    import openpyxl
    from openpyxl.utils import column_index_from_string as col_index
    from openpyxl.utils import get_column_letter as col_letter
except ImportError:
    sys.exit("Missing libraries. Install them with:  python -m pip install -r requirements.txt")

APP_ID = "survey-keyer"
HOST, PORT = "127.0.0.1", 8765
FROZEN = getattr(sys, "frozen", False)
APP_DIR = Path(sys.executable if FROZEN else __file__).resolve().parent
STATE_FILE = Path(os.environ.get("SURVEY_KEYER_STATE") or Path.home() / ".survey_keyer.json")   # recent files, token

NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
REL_ID = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
# starts at "(" or the marker word, never at a run of spaces, and has no two adjacent \s*: no catastrophic backtracking
# (the leading spaces are removed by .strip() where it is used)
OPEN_RE = re.compile(r"(?:\(\s*)?(?:pytanie otwarte|open question)\s*(?:\)\s*)?$", re.I)
KINDS = ("single", "multi", "open", "amount")
XML_BAD = re.compile("[\x00-\x08\x0b\x0c\x0e-\x1f\ud800-\udfff\ufffe\uffff]")   # characters that corrupt an .xlsx file
MAX_BODY = 2 * 1024 * 1024          # largest request accepted from the browser
MAX_PART = 300 * 1024 * 1024        # largest unpacked sheet (protection against a "zip bomb")
MAX_CELL = 32767                    # character limit of an Excel cell
MONTHS = ("January", "February", "March", "April", "May", "June", "July", "August", "September", "October",
          "November", "December")   # fixed English names: strftime("%B") would follow the system locale
CSP = ("default-src 'none'; script-src 'nonce-%s'; style-src 'unsafe-inline'; img-src 'self' data:; "
       "connect-src 'self'; base-uri 'none'; form-action 'none'; frame-ancestors 'none'")


class Conflict(Exception):
    """The file or its settings changed in another window - the browser must reload its view."""


def clean(s):
    return re.sub(r"\s+", " ", str(s)).strip()


def xml_text(s):
    return XML_BAD.sub("", str(s))


def fail_box(msg):
    """Error message that also shows up in the .exe build, which has no console window."""
    try:
        import tkinter
        from tkinter import messagebox
        root = tkinter.Tk()
        root.withdraw()
        messagebox.showerror("Survey Keyer", msg)
        root.destroy()
    except Exception:
        print(msg)


# --------------------------------------------------------------------------
# Detecting the layout and the questions
# --------------------------------------------------------------------------
def detect_layout(wb):
    """Sheet and rows: questions / answers / first survey."""
    best = None
    for ws in wb.worksheets:
        counts = {}
        for r in ws.merged_cells.ranges:
            if r.min_row == r.max_row and r.max_col > r.min_col and r.min_row <= 15:
                counts[r.min_row] = counts.get(r.min_row, 0) + 1
        if counts:
            row, n = max(counts.items(), key=lambda kv: (kv[1], -kv[0]))
            if not best or n > best[2]:
                best = (ws.title, row, n)
    if best:
        sheet, header = best[0], best[1]
    else:                                         # no merged cells: the row with "No." / "L.p." / "nr" in column A
        sheet, header = wb.worksheets[0].title, 1
        for ws in wb.worksheets:
            hit = next((r for r in range(1, 16)
                        if re.match(r"(l\.?\s*p\.?|nr)\b|no\.|no$", clean(ws.cell(r, 1).value or "").lower())), None)
            if hit:
                sheet, header = ws.title, hit
                break
    ws = wb[sheet]
    option = header + 1
    first = None
    for r in range(option + 1, min(ws.max_row, option + 30) + 1):
        v = ws.cell(r, 1).value
        if isinstance(v, (int, float)) and not isinstance(v, bool):
            first = r
            break
    if first is None:
        nxt = clean(ws.cell(option + 1, 1).value or "").lower()
        first = option + 2 if nxt.startswith(("razem", "suma", "total", "sum")) else option + 1
    return {"sheet": sheet, "header_row": header, "option_row": option, "first_row": first}


def guess(text, labels):
    low = text.lower()
    if len(labels) <= 1 and (not labels or any(w in labels[0][1].lower() for w in ("otwarte", "open question"))):
        return "open"
    if any(w in low for w in ("kwot", "budżet", "budzet", "ile zł", "ile zl", "ile pln", "amount", "budget", "pln")):
        return "amount"
    if any(w in low for w in ("kilka", "wiele", "wielokrot", "wszystkie, które", "zaznacz wszystkie",
                              "several", "multiple", "all that apply")):
        return "multi"
    return "single"


def parse_questions(ws, layout):
    H, O = layout["header_row"], layout["option_row"]
    merged = sorted((r.min_col, r.max_col) for r in ws.merged_cells.ranges
                    if r.min_row == r.max_row == H)
    in_merge = {c for a, b in merged for c in range(a, b + 1)}
    groups = list(merged)
    # headers without merged cells: a question spans the following columns that have an answer and an empty header
    c, last = 2, ws.max_column
    while c <= last:
        if c not in in_merge and ws.cell(H, c).value not in (None, ""):
            end = c
            while (end + 1 <= last and end + 1 not in in_merge
                   and ws.cell(H, end + 1).value in (None, "")
                   and ws.cell(O, end + 1).value not in (None, "")):
                end += 1
            groups.append((c, end))
            c = end + 1
        else:
            c += 1
    questions = []
    for n, (a, b) in enumerate(sorted(groups), 1):
        labels = [(col, clean(ws.cell(O, col).value)) for col in range(a, b + 1)
                  if ws.cell(O, col).value not in (None, "") and clean(ws.cell(O, col).value)]
        text = re.sub(r"^\d+\s*[.)]\s*", "", clean(ws.cell(H, a).value or "")) or f"Question {n}"
        opts = [dict(col=col, label=OPEN_RE.sub("", lab).strip() or lab,
                     text=any(w in lab.lower() for w in ("otwarte", "open question"))
                     or lab.lower().startswith(("inne", "other")))
                for col, lab in labels]
        questions.append(dict(key=str(a), num=n, text=text, col=a, last_col=b,
                              options=opts, kind=guess(text, labels)))
    return questions


# --------------------------------------------------------------------------
# Reading and writing cells directly in the XML (the rest of the file stays untouched)
# --------------------------------------------------------------------------
def read_part(z, name):
    if z.getinfo(name).file_size > MAX_PART:
        raise ValueError("The sheet is too large to open safely.")
    return z.read(name)


def check_unpacked_size(path):
    """Zip bomb: the total unpacked size is checked before openpyxl loads the whole file into memory."""
    try:
        with zipfile.ZipFile(path) as z:
            total = sum(i.file_size for i in z.infolist())
    except (OSError, zipfile.BadZipFile):
        return                                      # not a zip or no access: openpyxl will report a clear error
    if total > MAX_PART:
        raise ValueError("The file is too large once unpacked to open safely.")


def put_text(ws, row, col, text):
    """User text is always stored as text: openpyxl turns a string starting with "=" into a formula."""
    cell = ws.cell(row, col, text)
    if isinstance(text, str) and text.startswith("="):
        cell.data_type = "s"
    return cell


def sheet_part(z, sheet_name):
    wb = ET.fromstring(read_part(z, "xl/workbook.xml"))
    rels = ET.fromstring(read_part(z, "xl/_rels/workbook.xml.rels"))
    sheets = wb.find("m:sheets", NS)
    rid = next((s.get(REL_ID) for s in sheets if s.get("name") == sheet_name), sheets[0].get(REL_ID))
    for r in rels:
        if r.get("Id") == rid:
            t = r.get("Target")
            return t.lstrip("/") if t.startswith("/") else "xl/" + t
    raise RuntimeError("Sheet not found in the file.")


def shared_strings(z):
    if "xl/sharedStrings.xml" not in z.namelist():
        return []
    root = ET.fromstring(read_part(z, "xl/sharedStrings.xml"))
    return ["".join(t.text or "" for t in si.iter("{%s}t" % NS["m"])) for si in root]


def read_all_rows(path, sheet_name):
    """{row: {column: value}} - a single read of the file for the whole sheet."""
    with zipfile.ZipFile(path) as z:
        sheet = read_part(z, sheet_part(z, sheet_name)).decode("utf-8")
        sst = shared_strings(z)
    out = {}
    for rm in re.finditer(r'<row r="(\d+)"[^>]*?(?:/>|>(.*?)</row>)', sheet, re.S):
        cells = {}
        for m in re.finditer(r'<c r="([A-Z]+)\d+"([^>]*?)(?:/>|>(.*?)</c>)', rm.group(2) or "", re.S):
            col, attrs, body = m.group(1), m.group(2), m.group(3) or ""
            t = re.search(r'\bt="(\w+)"', attrs)
            t = t.group(1) if t else "n"
            val = None
            if t == "inlineStr":
                val = html.unescape("".join(re.findall(r"<t[^>]*>(.*?)</t>", body, re.S)))
            else:
                v = re.search(r"<v>(.*?)</v>", body, re.S)
                if v:
                    val = sst[int(v.group(1))] if t == "s" else html.unescape(v.group(1))
            if val not in (None, ""):
                cells[col_index(col)] = val
        out[int(rm.group(1))] = cells
    return out


def cell_xml(col, row, style, value):
    ref = f"{col_letter(col)}{row}"
    s = f' s="{style}"' if style else ""
    if value is None:
        return f'<c r="{ref}"{s}/>'
    if isinstance(value, bool):
        value = int(value)
    if isinstance(value, (int, float)):
        if not math.isfinite(value):
            raise ValueError(f"Cell {ref}: invalid number.")
        return f'<c r="{ref}"{s}><v>{value}</v></c>'
    text = xml_text(value)
    if len(text) > MAX_CELL:
        raise ValueError(f"Cell {ref}: the text is too long (an Excel cell holds up to {MAX_CELL:,} characters).")
    return (f'<c r="{ref}"{s} t="inlineStr"><is><t xml:space="preserve">'
            f"{html.escape(text, quote=False)}</t></is></c>")


def refresh_sums(sheet, cols, sum_row, first_row):
    """Refreshes the cached SUM/COUNTA results in the totals row (for previews that do not recalculate).
    Excel recalculates everything on open anyway (fullCalcOnLoad)."""
    for col in cols:
        L = col_letter(col)
        total, filled = 0.0, 0
        for m in re.finditer(r'<c r="%s(\d+)"([^>]*?)(?<!/)>(.*?)</c>' % L, sheet, re.S):
            if int(m.group(1)) < first_row:
                continue
            t = re.search(r'\bt="(\w+)"', m.group(2))
            v = re.search(r"<v>(.*?)</v>", m.group(3), re.S)
            if (v and v.group(1)) or "<is>" in m.group(3):
                filled += 1
            if v and (not t or t.group(1) == "n"):
                try:
                    total += float(v.group(1))
                except ValueError:
                    pass
        total = int(total) if total == int(total) else round(total, 10)
        pat = re.compile(r'(<c r="%s%d")([^>]*)>(<f[^>]*?(?:/>|>(.*?)</f>))(?:<v>.*?</v>|<v/>)?' % (L, sum_row), re.S)

        def put(mm):
            fn = (mm.group(4) or "").lstrip("=").upper()
            if fn.startswith("COUNTA("):
                val = filled
            elif not fn or fn.startswith("SUM("):     # empty body = shared formula (usually SUM)
                val = total
            else:
                return mm.group(0)                     # some other formula - leave it, Excel will recalculate it
            attrs = re.sub(r'\st="\w+"', "", mm.group(2))
            return f"{mm.group(1)}{attrs}>{mm.group(3)}<v>{val}</v>"
        sheet = pat.sub(put, sheet, count=1)
    return sheet


def rewrite_zip(path, replacements):
    """Replaces the given parts of the .xlsx file; atomic write through a temporary file."""
    path = Path(path)
    tmp = path.with_name(path.name + ".tmp")
    try:
        with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, "w") as zout:
            for item in zin.infolist():
                data = replacements.get(item.filename)
                zout.writestr(item, data.encode("utf-8") if data is not None else zin.read(item.filename),
                              compress_type=item.compress_type)
        tmp.replace(path)
    except PermissionError:
        raise PermissionError("Cannot save the file - it is open in Excel or another program. "
                              "Close it and try again.")
    finally:
        tmp.unlink(missing_ok=True)


def force_recalc(wbxml):
    """Makes Excel recalculate the formulas (totals, charts) when the file is opened."""
    if re.search(r'fullCalcOnLoad="(1|true)"', wbxml):
        return wbxml
    if "<calcPr" in wbxml:
        wbxml = re.sub(r'\sfullCalcOnLoad="\w*"', "", wbxml)
        return re.sub(r"<calcPr\b", '<calcPr fullCalcOnLoad="1"', wbxml, count=1)
    # no <calcPr> - it must come after sheets/functionGroups/externalReferences/definedNames
    ends = [m.end() for m in re.finditer(
        r"</(?:sheets|functionGroups|externalReferences|definedNames)>"
        r"|<(?:sheets|functionGroups|externalReferences|definedNames)\b[^>]*/>", wbxml)]
    if not ends:
        return wbxml
    return wbxml[:max(ends)] + '<calcPr fullCalcOnLoad="1"/>' + wbxml[max(ends):]


def set_row_cells(row_text, row, values):
    for col, value in values.items():
        ref = f"{col_letter(col)}{row}"
        cm = re.search(r'<c r="%s"([^>]*?)(?:/>|>.*?</c>)' % ref, row_text, re.S)
        if cm:
            sm = re.search(r'\bs="(\d+)"', cm.group(1))
            row_text = (row_text[:cm.start()] + cell_xml(col, row, sm.group(1) if sm else None, value)
                        + row_text[cm.end():])
        elif value is not None:
            new = cell_xml(col, row, None, value)
            pos = next((o.start() for o in re.finditer(r'<c r="([A-Z]+)%d"' % row, row_text)
                        if col_index(o.group(1)) > col), None)
            if pos is not None:
                row_text = row_text[:pos] + new + row_text[pos:]
            elif row_text.endswith("/>"):
                row_text = row_text[:-2] + ">" + new + "</row>"
            else:
                end = row_text.rfind("</row>")
                row_text = row_text[:end] + new + row_text[end:]
    return row_text


def write_cells(path, sheet_name, row, values, sum_row, first_row):
    with zipfile.ZipFile(path) as z:
        part = sheet_part(z, sheet_name)
        sheet = read_part(z, part).decode("utf-8")
        wbxml = read_part(z, "xl/workbook.xml").decode("utf-8")
    m = re.search(r'<row r="%d"[^>]*?(?:/>|>.*?</row>)' % row, sheet, re.S)
    if not m:
        raise RuntimeError(f"Row {row} is missing from the sheet.")
    sheet = sheet[:m.start()] + set_row_cells(m.group(0), row, values) + sheet[m.end():]
    sheet = refresh_sums(sheet, values.keys(), sum_row, first_row)
    rewrite_zip(path, {part: sheet, "xl/workbook.xml": force_recalc(wbxml)})


def styled_copy(sheet, src_row, row, seq):
    """New row with the formatting (borders, height) copied from row src_row."""
    m = re.search(r'<row r="%d"([^>]*?)(?:/>|>(.*?)</row>)' % src_row, sheet, re.S)
    if not m:
        return f'<row r="{row}">{cell_xml(1, row, None, seq)}</row>'
    attrs = re.sub(r'\s(?:r|spans)="[^"]*"', "", m.group(1))
    cells = []
    for c in re.finditer(r'<c r="([A-Z]+)%d"([^>]*?)(?:/>|>.*?</c>)' % src_row, m.group(2) or "", re.S):
        st = re.search(r'\bs="(\d+)"', c.group(2))
        col = col_index(c.group(1))
        cells.append(cell_xml(col, row, st.group(1) if st else None, seq if col == 1 else None))
    if not any(c.startswith(f'<c r="A{row}"') for c in cells):
        cells.insert(0, cell_xml(1, row, None, seq))
    return f'<row r="{row}"{attrs}>{"".join(cells)}</row>'


def add_row(path, sheet_name, row, seq, sum_row, old_last):
    """Adds a survey row (its number in column A) and extends the SUM ranges in the totals row."""
    with zipfile.ZipFile(path) as z:
        part = sheet_part(z, sheet_name)
        sheet = read_part(z, part).decode("utf-8")
    m = re.search(r'<row r="%d"[^>]*?(?:/>|>.*?</row>)' % row, sheet, re.S)
    if m:
        sheet = sheet[:m.start()] + set_row_cells(m.group(0), row, {1: seq}) + sheet[m.end():]
    else:
        new = styled_copy(sheet, old_last, row, seq) if old_last else f'<row r="{row}">{cell_xml(1, row, None, seq)}</row>'
        nxt = next((x for x in re.finditer(r'<row r="(\d+)"', sheet) if int(x.group(1)) > row), None)
        pos = nxt.start() if nxt else sheet.index("</sheetData>")
        if not nxt and "<sheetData/>" in sheet:
            sheet = sheet.replace("<sheetData/>", f"<sheetData>{new}</sheetData>")
        else:
            sheet = sheet[:pos] + new + sheet[pos:]
    if old_last and row > old_last:              # SUM(B4:B109) -> SUM(B4:B110)
        sm = re.search(r'<row r="%d"[^>]*?>.*?</row>' % sum_row, sheet, re.S)
        if sm:
            fixed = re.sub(r"(<f[^>]*>)(.*?)(</f>)",
                           lambda f: f.group(1) + re.sub(r"(:\$?[A-Z]{1,3}\$?)%d\b" % old_last,
                                                         lambda g: g.group(1) + str(row), f.group(2)) + f.group(3),
                           sm.group(0), flags=re.S)
            sheet = sheet[:sm.start()] + fixed + sheet[sm.end():]
    sheet = re.sub(r'(<dimension ref="[A-Z]+\d+:[A-Z]+)(\d+)"',
                   lambda d: d.group(1) + str(max(int(d.group(2)), row)) + '"', sheet, count=1)
    rewrite_zip(path, {part: sheet})


def to_amount(s):
    """Polish and English formats: '10 000 zł' -> 10000, '1.500' -> 1500, '2,5 tys.' -> 2500, '12,50' -> 12.5,
    '1,500' -> 1500, '1,234,567.8' -> 1234567.8, '250 PLN' -> 250.
    Text that cannot be read as a number stays text."""
    s = xml_text(s or "").strip()
    if not s:
        return None
    t = re.sub(r"(?i)złotych|złote|zlotych|zł|zl|pln", "", s).rstrip(".")
    t = re.sub(r"[\s  ']", "", t).replace("−", "-")
    mult = 1
    m = re.fullmatch(r"(?i)(.*\d)(tys\.?|k)", t)
    if m:
        t, mult = m.group(1), 1000
    if re.fullmatch(r"-?\d{1,3}(\.\d{3})+(,\d+)?", t):     # 1.500 or 1.500,50 - dots separate thousands
        t = t.replace(".", "")
    elif re.fullmatch(r"-?[1-9]\d{0,2}(,\d{3})+(\.\d+)?", t):   # 1,500 or 1,500.50 - English notation
        t = t.replace(",", "")
    if not re.fullmatch(r"-?(\d+([.,]\d*)?|[.,]\d+)", t):
        return s
    f = float(t.replace(",", ".")) * mult
    if not math.isfinite(f):
        return s
    return int(f) if f == int(f) else round(f, 10)


def num(v):
    x = to_amount(str(v)) if v is not None else None
    return x if isinstance(x, (int, float)) else None


def csv_safe(v):
    """Text typed by the user must not turn into a formula in Excel (CSV injection)."""
    v = "" if v is None else str(v)
    return "'" + v if v[:1] in ("=", "+", "-", "@", "\t", "\r") else v


def csv_num(v):
    return str(v) if v not in (None, "") else ""


# --------------------------------------------------------------------------
# Survey
# --------------------------------------------------------------------------
class Survey:
    def __init__(self, path):
        self.path = Path(path).resolve()
        if self.path.suffix.lower() not in (".xlsx", ".xlsm"):
            raise ValueError("Only .xlsx files are supported. Open an .xls or .ods file in Excel and save it as .xlsx.")
        self.config_path = self.path.with_name(self.path.stem + ".survey.json")
        self.lock = threading.Lock()
        self.backed_up = False
        self._cache = None
        self._mtime = None
        self.cfg_warning = None
        self.reopen()

    def reopen(self):
        """(Re)loads the settings and the structure of the file."""
        self.cfg = {}
        if self.config_path.exists():
            try:
                cfg = json.loads(self.config_path.read_text(encoding="utf-8"))
                self.cfg = cfg if isinstance(cfg, dict) else {}
            except (OSError, ValueError):
                self.cfg_warning = (f"The settings file {self.config_path.name} is damaged - "
                                    "check the question types and save the settings again.")
        self.load_structure(self.cfg.get("_layout"))

    # ---- structure ----
    @staticmethod
    def clean_layout(layout, sheets):
        """Only valid layout entries (from the settings file or from the user)."""
        out = {}
        if not isinstance(layout, dict):
            return out
        for k in ("header_row", "option_row", "first_row"):
            v = layout.get(k)
            if v in (None, ""):
                continue
            try:
                v = int(v)
            except (TypeError, ValueError):
                raise ValueError("Row numbers must be whole numbers.")
            if not 1 <= v <= 1_048_576:
                raise ValueError("A row number must be between 1 and 1,048,576.")
            out[k] = v
        if layout.get("sheet") in sheets:
            out["sheet"] = layout["sheet"]
        return out

    def load_structure(self, layout=None):
        check_unpacked_size(self.path)
        try:
            wb = openpyxl.load_workbook(self.path)
        except Exception as e:
            raise ValueError(f"Could not open the file as an Excel workbook ({e}).")
        try:
            auto = detect_layout(wb)
            try:
                saved = self.clean_layout(layout, wb.sheetnames)
            except ValueError:
                saved = {}
            self.layout = {**auto, **saved}
            self.sheets = wb.sheetnames
            self.questions = parse_questions(wb[self.layout["sheet"]], self.layout)
        finally:
            wb.close()
        self._cache = None
        self._load_data()
        # how many surveys have several answers ticked - if any, the question defaults to multiple choice
        for q in self.questions:
            q["multi_seen"] = 0
            if q["kind"] in ("single", "multi") and len(q["options"]) > 1:
                for seq in self.rows:
                    d = self.row_data(seq)
                    if sum(1 for o in q["options"] if str(d.get(o["col"], "")).strip() not in ("", "0")) > 1:
                        q["multi_seen"] += 1
                if q["multi_seen"]:
                    q["kind"] = "multi"
            if q["kind"] in ("single", "multi"):          # text typed in an answer column -> text field
                for o in q["options"]:
                    if any(str(self.row_data(seq).get(o["col"], "1")).strip() not in ("1", "0", "")
                           for seq in self.rows):
                        o["text"] = True
        self._apply_question_cfg(self.cfg)

    def _apply_question_cfg(self, cfg):
        for q in self.questions:
            c = cfg.get(q["key"]) if isinstance(cfg, dict) else None
            if not isinstance(c, dict):
                continue
            if c.get("kind") in KINDS:
                q["kind"] = c["kind"]
            tc = c.get("text_cols") if isinstance(c.get("text_cols"), dict) else {}
            for o in q["options"]:
                if str(o["col"]) in tc:
                    o["text"] = bool(tc[str(o["col"])])
        self._make_rev()

    def _make_rev(self):
        """Fingerprint of the layout and settings - a browser with a different fingerprint has a stale form."""
        sig = [str(self.path), self.layout,
               [(q["key"], q["kind"], [(o["col"], o["label"], o["text"]) for o in q["options"]]) for q in self.questions]]
        self.rev = hashlib.sha1(json.dumps(sig, ensure_ascii=False, sort_keys=True).encode("utf-8"),
                                usedforsecurity=False).hexdigest()[:16]   # change fingerprint, not a security measure

    def set_config(self, cfg, layout=None):
        if layout:
            new = self.clean_layout(layout, self.sheets)
            if new != {k: self.layout[k] for k in new}:
                merged = {**self.layout, **new}
                if not merged["header_row"] < merged["option_row"] < merged["first_row"]:
                    raise ValueError("The rows must be in order: questions first, the answers below them, "
                                     "and the surveys further down.")
                self.cfg = {"_layout": merged}
                self.load_structure(merged)
                cfg = {}
        self._apply_question_cfg(cfg)
        self.cfg = {"_layout": self.layout}
        self.cfg.update({q["key"]: {"kind": q["kind"], "text": q["text"],
                                    "text_cols": {str(o["col"]): o["text"] for o in q["options"]}}
                         for q in self.questions})
        self.config_path.write_text(json.dumps(self.cfg, ensure_ascii=False, indent=1), encoding="utf-8")
        self.cfg_warning = None

    # ---- data ----
    def cols(self, q):
        return [o["col"] for o in q["options"]] or [q["col"]]

    def _load_data(self):
        try:
            self._mtime = self.path.stat().st_mtime_ns
        except OSError:
            raise ValueError(f"File not found: {self.path}")
        self._cache = read_all_rows(self.path, self.layout["sheet"])
        self.rows, self.dupes = {}, []
        for r in sorted(self._cache):
            v = self._cache[r].get(1)
            if r >= self.layout["first_row"] and v is not None and re.fullmatch(r"\s*\d+(\.0*)?\s*", str(v)):
                seq = int(float(v))
                if seq in self.rows:
                    self.dupes.append(seq)
                else:
                    self.rows[seq] = r

    def refresh(self):
        """File changed outside the app (e.g. in Excel) - reload the data before using it."""
        try:
            changed = self.path.stat().st_mtime_ns != self._mtime
        except OSError:
            raise ValueError(f"File not found: {self.path}. Has it been moved?")
        if changed or self._cache is None:
            self._load_data()

    def data(self):
        if self._cache is None:
            self._load_data()
        return self._cache

    def locked(self):
        """Excel locks an open file exclusively - saving would fail."""
        try:
            with open(self.path, "r+b"):
                return False
        except PermissionError:
            return True
        except OSError:
            return False

    def row_data(self, seq):
        return self.data().get(self.rows[seq], {})

    def is_filled(self, seq):
        d = self.row_data(seq)
        return any(c in d for q in self.questions for c in self.cols(q))

    def filled_ids(self):
        return [seq for seq in sorted(self.rows) if self.is_filled(seq)]

    def next_empty(self, after=0):
        ids = sorted(self.rows)
        for seq in [i for i in ids if i > after] + [i for i in ids if i <= after]:
            if not self.is_filled(seq):
                return seq
        return None

    def load(self, seq):
        d = self.row_data(seq)
        ans = {}
        for q in self.questions:
            if q["kind"] == "open":
                ans[q["key"]] = {"text": " ".join(str(d[c]) for c in self.cols(q) if c in d)}
            elif q["kind"] == "amount":
                ans[q["key"]] = {"values": {str(i): str(d[o["col"]]) for i, o in enumerate(q["options"])
                                            if o["col"] in d}}
            else:
                chosen, texts = [], {}
                for i, o in enumerate(q["options"]):
                    v = d.get(o["col"])
                    if v is None or str(v).strip() in ("", "0"):
                        continue
                    chosen.append(i)
                    if str(v).strip() != "1":
                        texts[str(i)] = str(v)
                ans[q["key"]] = {"chosen": chosen, "texts": texts}
        return ans

    # ---- saving ----
    def backup(self):
        folder = self.path.parent / "backups"
        folder.mkdir(exist_ok=True)
        dst = folder / f"{self.path.stem}_{datetime.now():%Y-%m-%d_%H%M%S}{self.path.suffix}"
        shutil.copy2(self.path, dst)
        self.backed_up = True
        return dst

    def _ensure_backup(self):
        if not self.backed_up:
            self.backup()

    def save(self, seq, ans):
        values = {}
        old = self.row_data(seq)
        for q in self.questions:
            a = ans.get(q["key"])
            a = a if isinstance(a, dict) else {}
            cols = self.cols(q)
            if q["kind"] == "open":
                values[cols[0]] = xml_text(a.get("text") or "").strip() or None
                for c in cols[1:]:
                    values[c] = None
            elif q["kind"] == "amount":
                vals = a.get("values") if isinstance(a.get("values"), dict) else {}
                for i, o in enumerate(q["options"]):
                    v = vals.get(str(i))
                    values[o["col"]] = to_amount(v if isinstance(v, str) else None)
            else:
                raw = a.get("chosen") if isinstance(a.get("chosen"), list) else []
                chosen = {c for c in raw if isinstance(c, int) and not isinstance(c, bool)}
                texts = a.get("texts") if isinstance(a.get("texts"), dict) else {}
                if q["kind"] == "single" and len(chosen) > 1:
                    raise ValueError(f"Question {q['num']}: only one answer is allowed - untick the others.")
                for i, o in enumerate(q["options"]):
                    t = xml_text(texts.get(str(i)) if isinstance(texts.get(str(i)), str) else "").strip()
                    prev = str(old.get(o["col"], "")).strip()
                    if not t and not o["text"] and prev not in ("", "0", "1"):
                        t = prev                          # do not replace previously typed text with "1"
                    values[o["col"]] = (t or 1) if i in chosen else None
        L = self.layout
        with self.lock:
            self._ensure_backup()
            write_cells(self.path, L["sheet"], self.rows[seq], values, L["option_row"] + 1, L["first_row"])
            self._cache = None

    def new_row(self):
        """New survey at the end of the sheet. Returns its number (column A)."""
        L = self.layout
        with self.lock:
            last_row = max(self.rows.values(), default=L["first_row"] - 1)
            row, seq = last_row + 1, max(self.rows, default=0) + 1
            if any(v not in (None, "") for v in self.data().get(row, {}).values()):
                raise ValueError(f"Row {row} below the surveys is not empty - cannot add a survey automatically.")
            self._ensure_backup()
            add_row(self.path, L["sheet"], row, seq, L["option_row"] + 1, last_row if self.rows else None)
            self.rows[seq] = row
            self._cache = None
        return seq

    # ---- statistics ----
    def stats(self):
        ids = self.filled_ids()
        n = len(ids)
        out = []
        for q in self.questions:
            rows = [self.row_data(seq) for seq in ids]
            item = {"num": q["num"], "text": q["text"], "kind": q["kind"], "options": [], "answers": []}
            if q["kind"] == "open":
                item["answers"] = [{"seq": seq, "text": " ".join(str(self.row_data(seq)[c]) for c in self.cols(q)
                                                                if c in self.row_data(seq))}
                                   for seq in ids if any(c in self.row_data(seq) for c in self.cols(q))]
                item["answered"] = len(item["answers"])
            else:
                answered = sum(1 for d in rows if any(o["col"] in d for o in q["options"]))
                for o in q["options"]:
                    vals = [d[o["col"]] for d in rows if o["col"] in d and str(d[o["col"]]).strip() not in ("", "0")]
                    entry = {"label": o["label"], "count": len(vals),
                             "pct": round(100 * len(vals) / answered, 1) if answered else 0}
                    if q["kind"] == "amount":
                        nums = [x for x in (num(v) for v in vals) if x is not None]
                        entry["sum"] = round(sum(nums), 2)
                        entry["avg"] = round(sum(nums) / len(nums), 2) if nums else 0
                        entry["not_numbers"] = len(vals) - len(nums)
                    if o["text"]:
                        entry["texts"] = [str(v) for v in vals if str(v).strip() != "1"]
                    item["options"].append(entry)
                item["answered"] = answered
            out.append(item)
        return {"filled": n, "total": len(self.rows), "questions": out}

    def stats_csv(self):
        buf = io.StringIO()
        w = csv.writer(buf)
        st = self.stats()
        w.writerow([csv_safe(f"Survey: {self.path.name}"), f"Filled in: {st['filled']}"])
        w.writerow(["No.", "Question", "Answer", "Count", "% of respondents", "Sum", "Average"])
        for q in st["questions"]:
            if q["kind"] == "open":
                for a in q["answers"]:
                    w.writerow([q["num"], csv_safe(q["text"]), csv_safe(a["text"]), "", "", "", ""])
                continue
            for o in q["options"]:
                w.writerow([q["num"], csv_safe(q["text"]), csv_safe(o["label"]), o["count"], csv_num(o["pct"]),
                            csv_num(o.get("sum")), csv_num(o.get("avg"))])
        return "﻿" + buf.getvalue()

    def schema(self):
        return {"file": self.path.name, "path": str(self.path), "questions": self.questions,
                "ids": sorted(self.rows), "filled": len(self.filled_ids()), "rev": self.rev,
                "has_config": self.config_path.exists() and not self.cfg_warning, "warning": self.cfg_warning,
                "dupes": sorted(set(self.dupes)), "layout": self.layout, "sheets": self.sheets}


# --------------------------------------------------------------------------
# Wizard: a new survey file in a layout the app understands right away
# --------------------------------------------------------------------------
def clean_spec(d):
    """Validates the wizard data. Returns {title, rows, charts, questions: [{text, kind, options, other}]}."""
    if not isinstance(d, dict):
        raise ValueError("Invalid survey data.")
    title = xml_text(clean(d.get("title") or ""))[:200] or "Survey"
    try:
        rows = int(d.get("rows") or 100)
    except (TypeError, ValueError):
        raise ValueError("The number of survey rows must be a number.")
    if not 1 <= rows <= 5000:
        raise ValueError("The number of survey rows must be between 1 and 5,000.")
    qs = d.get("questions")
    if not isinstance(qs, list) or not qs:
        raise ValueError("Add at least one question.")
    if len(qs) > 300:
        raise ValueError("A survey can have at most 300 questions.")
    out, ncols = [], 1
    for i, q in enumerate(qs, 1):
        if not isinstance(q, dict):
            raise ValueError(f"Question {i}: invalid data.")
        text = xml_text(clean(q.get("text") or ""))[:1000]
        kind = q.get("kind")
        if not text:
            raise ValueError(f"Question {i}: type the question text.")
        if kind not in KINDS:
            raise ValueError(f"Question {i}: choose a question type.")
        raw = q.get("options") if isinstance(q.get("options"), list) else []
        opts = [xml_text(clean(o))[:300] for o in raw if isinstance(o, str) and clean(o)] if kind != "open" else []
        if kind != "open" and not opts:
            raise ValueError(f"Question {i}: type at least " +
                             ("one amount field." if kind == "amount" else "one answer."))
        if len(opts) > 60:
            raise ValueError(f"Question {i}: at most 60 answers.")
        out.append({"text": text, "kind": kind, "options": opts,
                    "other": bool(q.get("other")) and kind in ("single", "multi")})
        ncols += max(1, len(opts))
    if ncols > 16000:
        raise ValueError("Too many answers - Excel has a column limit.")
    return {"title": title, "rows": rows, "charts": d.get("charts") is not False, "questions": out}


def build_survey_file(path, spec):
    """Sheet: title, question row (merged cells), answer row, totals row, numbered survey rows.
    Saves <name>.survey.json with the question types next to it, so the file is ready for data entry right away."""
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    H, O, SUMR, FIRST = 4, 5, 6, 7
    last = FIRST + spec["rows"] - 1
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Survey"
    line = Side(style="thin", color="B9B4A6")
    grid = Border(left=line, right=line, top=line, bottom=line)
    centre = Alignment(horizontal="center", vertical="center", wrap_text=True)
    put_text(ws, 1, 1, spec["title"]).font = Font(size=14, bold=True)
    now = datetime.now()
    ws["A2"] = f"Created on {now.day} {MONTHS[now.month - 1]} {now.year} in the Survey Keyer app. One row is one survey."
    ws["A2"].font = Font(italic=True, color="6B7268")
    ws.cell(H, 1, "No.")
    ws.cell(SUMR, 1, "Total")
    ws.column_dimensions["A"].width = 7
    cfg = {"_layout": {"sheet": ws.title, "header_row": H, "option_row": O, "first_row": FIRST}}
    charts, col, head_h, opt_h = [], 2, 30, 30
    for i, q in enumerate(spec["questions"], 1):
        labels = q["options"] or ["(open question)"]
        a, b = col, col + len(labels) - 1
        ws.cell(H, a, f"{i}. {q['text']}")
        if b > a:
            ws.merge_cells(start_row=H, start_column=a, end_row=H, end_column=b)
        widths = []
        for k, lab in enumerate(labels):
            c, L = a + k, col_letter(a + k)
            w = 40 if q["kind"] == "open" else max(10, min(24, len(lab) + 3))
            widths.append(w)
            ws.column_dimensions[L].width = w
            put_text(ws, O, c, lab)
            ws.cell(SUMR, c,f"={'SUM' if q['kind'] == 'amount' else 'COUNTA'}({L}{FIRST}:{L}{last})")
            opt_h = max(opt_h, 15 * math.ceil(len(lab) * 1.1 / (w - 1)) + 6)
        head_h = max(head_h, 15 * math.ceil((len(q["text"]) + 4) * 1.1 / sum(widths)) + 6)
        cfg[str(a)] = {"kind": q["kind"], "text": q["text"],
                       "text_cols": {str(a + k): q["other"] and k == len(labels) - 1 for k in range(len(labels))}}
        if q["kind"] in ("single", "multi"):
            charts.append((f"{i}. {q['text']}", a, b))
        col = b + 1
    for r in (H, O, SUMR):
        for c in range(1, col):
            cell = ws.cell(r, c)
            cell.border, cell.alignment = grid, centre
            cell.fill = PatternFill("solid", fgColor="E6F0E9" if r == H else "F5F3EE")
            cell.font = Font(bold=r != O, color="1D5C3E" if r == H else "1C1F1B")
    for r in range(FIRST, last + 1):
        ws.cell(r, 1, r - FIRST + 1).alignment = Alignment(horizontal="center")
        for c in range(1, col):
            ws.cell(r, c).border = grid
    ws.row_dimensions[H].height = min(160, head_h)
    ws.row_dimensions[O].height = min(120, opt_h)
    ws.freeze_panes = ws.cell(FIRST, 2)
    if spec["charts"] and charts:
        from openpyxl.chart import BarChart, Reference
        from openpyxl.chart.label import DataLabelList
        cs = wb.create_sheet("Charts")
        cs["A1"] = "Number of answers to each question (updates when the file is opened in Excel)"
        cs["A1"].font = Font(bold=True)
        top = 3
        for title, a, b in charts:
            ch = BarChart()
            ch.type, ch.legend = "bar", None
            ch.title = title if len(title) <= 110 else title[:107] + "…"
            ch.add_data(Reference(ws, min_col=a, max_col=b, min_row=SUMR, max_row=SUMR), from_rows=True)
            ch.set_categories(Reference(ws, min_col=a, max_col=b, min_row=O, max_row=O))
            ch.x_axis.scaling.orientation = "maxMin"        # answers in the same order as in the survey
            ch.x_axis.delete = ch.y_axis.delete = False
            ch.y_axis.majorGridlines = None
            ch.dataLabels = DataLabelList()
            ch.dataLabels.showVal = True
            ch.series[0].graphicalProperties.solidFill = "2B7A54"
            ch.series[0].graphicalProperties.line.noFill = True
            ch.width, ch.height = 18, max(5.5, 1.1 * (b - a + 1) + 3)
            cs.add_chart(ch, f"A{top}")
            top += math.ceil(ch.height / 0.53) + 2
    wb.save(path)
    Path(path).with_name(Path(path).stem + ".survey.json").write_text(
        json.dumps(cfg, ensure_ascii=False, indent=1), encoding="utf-8")


# --------------------------------------------------------------------------
# Demo mode: a sample survey with made-up data (no organisation files needed)
# --------------------------------------------------------------------------
DEMO_SPEC = {"title": "Residents' survey 2026 (sample, made-up data)", "rows": 60, "questions": [
    {"text": "How do you rate the green spaces in your area?", "kind": "single",
     "options": ["Very good", "Good", "Average", "Poor", "No opinion"]},
    {"text": "What do you use most often?", "kind": "multi",
     "options": ["Park", "Playground", "Cycle path", "Outdoor gym", "Other (please specify)"], "other": True},
    {"text": "What is missing in your area?", "kind": "open"},
    {"text": "How much do you spend per month on sports and cultural activities (PLN)?", "kind": "amount",
     "options": ["Sport", "Culture"]},
    {"text": "Age", "kind": "single", "options": ["18 or under", "19-35", "36-60", "over 60"]},
]}
DEMO_FILLED = 42
DEMO_TEXTS = ["More benches", "Better lighting on the paths", "A toilet in the park", "More litter bins",
              "Bike racks", "A fountain", "", "", ""]


def is_demo(argv):
    """Demo mode: the --demo argument or the environment variable DEMO=true."""
    return "--demo" in argv or (os.environ.get("DEMO") or "").strip().lower() in ("1", "true", "tak", "yes", "on")


def demo_file(folder=None):
    """Creates (once) a sample survey with a few dozen made-up answers and returns the path to the file."""
    import random
    folder = Path(folder or Path.home() / "Survey Keyer demo")
    folder.mkdir(parents=True, exist_ok=True)
    path = folder / "Sample survey (demo).xlsx"
    if path.exists():
        return path
    build_survey_file(path, clean_spec(DEMO_SPEC))
    s = Survey(path)
    k = [q["key"] for q in s.questions]
    rnd = random.Random(2026)   # nosec B311 (made-up answers, always the same; not cryptography)
    for seq in range(1, DEMO_FILLED + 1):
        multi = sorted(rnd.sample(range(5), rnd.randint(1, 3)))
        ans = {k[0]: {"chosen": [rnd.choices(range(5), weights=[3, 5, 3, 1, 1])[0]]},
               k[1]: {"chosen": multi},
               k[3]: {"values": {"0": str(rnd.choice([0, 50, 80, 120, 200])), "1": str(rnd.choice([0, 30, 60, 100]))}},
               k[4]: {"chosen": [rnd.randint(0, 3)]}}
        if 4 in multi:
            ans[k[1]]["texts"] = {"4": rnd.choice(["Sports field", "Library", "Swimming pool"])}
        text = rnd.choice(DEMO_TEXTS)
        if text:
            ans[k[2]] = {"text": text}
        s.save(seq, ans)
    return path


def safe_filename(title):
    name = re.sub(r'[<>:"/\\|?*\x00-\x1f]+', " ", title).strip(" .")
    return clean(name)[:80] or "Survey"


# --------------------------------------------------------------------------
# Files: search, recent files, the "Open" and "Save as" dialogs
# --------------------------------------------------------------------------
def load_state():
    try:
        st = json.loads(STATE_FILE.read_text(encoding="utf-8"))
        return st if isinstance(st, dict) else {}
    except Exception:
        return {}


def save_state(st):
    try:
        STATE_FILE.write_text(json.dumps(st, ensure_ascii=False), encoding="utf-8")
    except OSError:
        pass


def remember(path):
    st = load_state()
    recent = [p for p in st.get("recent", []) if isinstance(p, str) and p != str(path) and Path(p).exists()]
    st["recent"] = [str(path)] + recent[:7]
    save_state(st)


def app_token():
    """Persistent secret key of this installation: a browser tab keeps working after the app restarts."""
    st = load_state()
    tok = st.get("token")
    if not (isinstance(tok, str) and len(tok) >= 32):
        tok = st["token"] = secrets.token_urlsafe(32)
        save_state(st)
    return tok


def looks_like_survey(f):
    if f.with_name(f.stem + ".survey.json").exists():
        return True
    try:
        check_unpacked_size(f)
        wb = openpyxl.load_workbook(f, read_only=False)
        lay = detect_layout(wb)
        ws = wb[lay["sheet"]]
        heads = sum(1 for r in ws.merged_cells.ranges if r.min_row == r.max_row == lay["header_row"])
        wb.close()
        return heads >= 3
    except Exception:
        return False


def find_surveys(root, depth=3, limit=300):
    """Survey files in the projects folder - at most `depth` levels deep and `limit` checked files."""
    found, checked = [], 0
    # real folder names: the app's "backups", Polish "kopi..." (copies), "poprzednia wersja" (old version)
    skip = ("kopi", "backups", "poprzednia wersja", "graphify-out", "__pycache__", "node_modules")
    for dirpath, dirnames, filenames in os.walk(root, onerror=lambda e: None):
        rel = Path(dirpath).relative_to(root)
        dirnames[:] = [] if len(rel.parts) >= depth else sorted(
            d for d in dirnames if not d.startswith((".", "$", "~")) and not any(s in d.lower() for s in skip))
        for name in sorted(filenames):
            low = name.lower()
            if not low.endswith(".xlsx") or name.startswith("~$") or "kopia" in low:
                continue
            f = Path(dirpath) / name
            checked += 1
            try:
                if checked > limit or f.stat().st_size > 50 * 1024 * 1024:
                    continue
            except OSError:
                continue
            if looks_like_survey(f):
                found.append({"path": str(f), "name": str(f.relative_to(root))})
    return found


DIALOG_LOCK = threading.Lock()


def _dialog(fn):
    """System dialog (Open / Save as) on top of the browser; at most one at a time."""
    if not DIALOG_LOCK.acquire(blocking=False):
        raise ValueError("A file dialog is already open - look for it on the taskbar.")
    try:
        import tkinter
        root = tkinter.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        try:
            return fn(root)
        finally:
            root.destroy()
    finally:
        DIALOG_LOCK.release()


def ask_file():
    from tkinter import filedialog
    return _dialog(lambda root: filedialog.askopenfilename(
        parent=root, title="Choose a survey file",
        filetypes=[("Excel files", "*.xlsx *.xlsm"), ("All files", "*.*")]))


def ask_save(title, folder):
    from tkinter import filedialog
    return _dialog(lambda root: filedialog.asksaveasfilename(
        parent=root, title="Save the new survey as", initialdir=str(folder),
        initialfile=safe_filename(title) + ".xlsx", defaultextension=".xlsx",
        filetypes=[("Excel file", "*.xlsx")], confirmoverwrite=False))


# --------------------------------------------------------------------------
# Server
# --------------------------------------------------------------------------
class App:
    def __init__(self):
        self.survey = None                       # most recently opened file (for a new tab)
        self.surveys = {}                        # all open files - each tab works on its own
        self.tabs = {}                           # browser tab -> time of its last signal
        self.closing_at = None
        self.lock = threading.Lock()
        self.token = app_token()
        self.ask_save = ask_save                 # "Save as" dialog (replaced in tests)
        parent = APP_DIR.parent                  # projects folder above the app folder
        self.search_root = APP_DIR if parent in (Path.home(), Path(parent.anchor)) else parent
        self._found = None

    def get(self, path, reload=False):
        p = Path(str(path).strip().strip('"'))
        if not p.is_file():
            raise ValueError(f"File not found: {p}")
        key = str(p.resolve())
        with self.lock:
            s = self.surveys.get(key)
            if s is None:
                s = Survey(p)
                for old in list(self.surveys)[:-7]:          # keep at most 8 files in memory
                    if self.surveys[old] is not self.survey:
                        del self.surveys[old]
                self.surveys[key] = s
            elif reload:
                s.reopen()
        return s

    def open(self, path):
        self.survey = self.get(path, reload=True)
        remember(self.survey.path)
        return self.survey.schema()

    def default_dir(self):
        recent = [p for p in load_state().get("recent", []) if isinstance(p, str) and Path(p).parent.is_dir()]
        return Path(recent[0]).parent if recent else self.search_root

    def create(self, data):
        spec = clean_spec(data)
        p = self.ask_save(spec["title"], self.default_dir())   # the user picks the location, not the page
        if not p:
            return {"cancelled": True}
        p = Path(p)
        if p.suffix.lower() != ".xlsx":
            p = p.with_name(p.name + ".xlsx")
        if p.exists():
            raise ValueError(f"The file “{p.name}” already exists. Choose another name - the existing file stays unchanged.")
        if not p.parent.is_dir():
            raise ValueError(f"The folder {p.parent} does not exist.")
        build_survey_file(p, spec)
        self._found = None
        return self.open(p)

    def files(self, refresh=False, scan=True):
        if scan and (self._found is None or refresh):
            self._found = find_surveys(self.search_root)
        recent = [{"path": p, "name": Path(p).name, "dir": str(Path(p).parent)}
                  for p in load_state().get("recent", []) if isinstance(p, str) and Path(p).exists()]
        return {"recent": recent, "found": self._found if scan else None, "root": str(self.search_root)}

    def touch(self, tab):
        if tab:
            self.tabs[tab[:64]] = time.time()
            self.closing_at = None

    def bye(self, tab):
        """Tab closed or reloading - shut the app down unless some tab checks in shortly."""
        self.tabs.pop(tab[:64], None)
        if not self.tabs:
            self.closing_at = time.time() + 8


def make_handler(app):
    class H(BaseHTTPRequestHandler):
        server_version, sys_version = "Survey Keyer", ""

        def log_message(self, *a):
            pass

        def _send(self, body, ctype, code=200, extra=None):
            self.send_response(code)
            self.send_header("Content-Type", ctype)
            self.send_header("Content-Length", str(len(body)))
            self.send_header("Cache-Control", "no-store")
            self.send_header("X-Content-Type-Options", "nosniff")
            self.send_header("X-Frame-Options", "DENY")
            self.send_header("Referrer-Policy", "no-referrer")
            self.send_header("Cross-Origin-Resource-Policy", "same-origin")
            for k, v in (extra or {}).items():
                self.send_header(k, v)
            self.end_headers()
            self.wfile.write(body)

        def _json(self, obj, code=200):
            self._send(json.dumps(obj, ensure_ascii=False).encode("utf-8"), "application/json; charset=utf-8", code)

        def _trusted(self):
            """Only our own page: the right Host (protects against DNS rebinding) and Origin (when the browser sends one)."""
            port = self.server.server_address[1]
            hosts = {f"127.0.0.1:{port}", f"localhost:{port}"}
            origin = self.headers.get("Origin")
            return self.headers.get("Host") in hosts and (origin is None or origin in {"http://" + h for h in hosts})

        def _authed(self):
            given = self.headers.get("X-Survey-Keyer-Token", "")
            return hmac.compare_digest(given.encode("latin-1", "replace"), app.token.encode("latin-1"))

        def _read_raw(self):
            """The request body is always read before replying (otherwise Windows drops the connection); None = too large."""
            try:
                n = int(self.headers.get("Content-Length") or 0)
            except ValueError:
                return None
            if not 0 <= n <= MAX_BODY:
                return None
            return self.rfile.read(n) if n else b""

        @staticmethod
        def _body(raw):
            data = json.loads(raw or b"{}")
            if not isinstance(data, dict):
                raise ValueError("Invalid data in the request.")
            return data

        def _survey(self):
            p = self.headers.get("X-Survey-Keyer-Path")
            s = app.get(unquote(p)) if p else app.survey
            if s is None:
                raise ValueError("Choose a survey file first.")
            rev = self.headers.get("X-Survey-Keyer-Rev")
            if rev and rev != s.rev:
                raise Conflict("The questions or settings of this file have changed (in another window or in Excel). "
                               "The form will be reloaded.")
            s.refresh()
            return s

        def _seq(self, s, seq):
            try:
                seq = int(seq)
            except (TypeError, ValueError):
                raise ValueError("The survey number must be a number.")
            if seq not in s.rows:
                raise ValueError(f"There is no survey no. {seq} in the file.")
            return seq

        def do_GET(self):
            self._handle("GET")

        def do_POST(self):
            self._handle("POST")

        def _handle(self, method):
            u = urlparse(self.path)
            raw = self._read_raw() if method == "POST" else b""
            if raw is None:
                self.close_connection = True
                return self._json({"error": "Too much data in one request."}, 400)
            if not self._trusted():
                return self._json({"error": "Access denied."}, 403)
            if method == "GET" and u.path == "/":
                # no key in the page: any process on the computer can fetch it; the key arrives in the URL (?t=) from the app
                nonce = secrets.token_urlsafe(16)
                page = PAGE.replace("__NONCE__", nonce)
                return self._send(page.encode("utf-8"), "text/html; charset=utf-8",
                                  extra={"Content-Security-Policy": CSP % nonce})
            if u.path == "/api/hello":
                challenge = (parse_qs(u.query).get("n") or [""])[0][:64]
                return self._json({"app": APP_ID, "proof": hello_proof(app.token, challenge, self.server.server_address[1])})
            if not u.path.startswith("/api/"):
                return self._json({"error": "Not found"}, 404)
            if not self._authed():
                return self._json({"error": "This tab has no access to the app. Close it and open Survey Keyer again "
                                            "from the program icon.", "auth": True}, 403)
            tab = self.headers.get("X-Survey-Keyer-Tab", "")
            try:
                if method == "POST":
                    data = self._body(raw)
                    if u.path == "/api/bye":
                        app.bye(tab)
                        return self._json({"ok": True})
                    app.touch(tab)
                    out = self._post(u.path, data)
                else:
                    app.touch(tab)
                    out = self._get(u.path, parse_qs(u.query))
                if isinstance(out, tuple):                     # (body, type) - a file to download
                    return self._send(out[0], out[1])
                if out is None:
                    return self._json({"error": "Not found"}, 404)
                self._json(out)
            except Conflict as e:
                self._json({"error": str(e), "conflict": True}, 409)
            except KeyError as e:
                self._json({"error": f"Missing field {e} in the request."}, 400)
            except (ValueError, PermissionError, RuntimeError, OSError, zipfile.BadZipFile) as e:
                self._json({"error": str(e)}, 400)
            except Exception as e:
                self._json({"error": f"Unexpected error: {e}"}, 500)

        def _get(self, path, q):
            if path == "/api/ping":
                p = self.headers.get("X-Survey-Keyer-Path")
                s = app.surveys.get(str(Path(unquote(p)).resolve())) if p else None
                return {"ok": True, "locked": bool(s and s.locked())}
            if path == "/api/files":
                return app.files(q.get("refresh") == ["1"], q.get("scan") != ["0"])
            s = self._survey()
            if path == "/api/schema":
                return s.schema()
            if path == "/api/next":
                return {"seq": s.next_empty(self._int(q, "after"))}
            if path == "/api/state":
                seq = self._seq(s, q["seq"][0])
                return {"filled": s.is_filled(seq), "answers": s.load(seq)}
            if path == "/api/progress":
                return {"filled": s.filled_ids(), "total": len(s.rows), "ids": sorted(s.rows), "locked": s.locked()}
            if path == "/api/stats":
                return s.stats()
            if path == "/api/stats.csv":
                return s.stats_csv().encode("utf-8"), "text/csv; charset=utf-8"
            return None

        @staticmethod
        def _int(q, key):
            try:
                return int(q.get(key, ["0"])[0])
            except ValueError:
                return 0

        def _post(self, path, data):
            if path == "/api/quit":
                app.tabs.clear()
                app.closing_at = time.time()
                return {"ok": True}
            if path == "/api/browse":
                p = ask_file()
                return app.open(p) if p else {"cancelled": True}
            if path == "/api/open":
                return app.open(data["path"])
            if path == "/api/create":
                return app.create(data)
            s = self._survey()
            if path == "/api/config":
                cfg = data.get("config")
                s.set_config(cfg if isinstance(cfg, dict) else {}, data.get("layout"))
                return s.schema()
            if path == "/api/save":
                ans = data.get("answers")
                s.save(self._seq(s, data.get("seq")), ans if isinstance(ans, dict) else {})
                return {"ok": True, "filled": len(s.filled_ids())}
            if path == "/api/newrow":
                return {"seq": s.new_row(), "ids": sorted(s.rows)}
            if path == "/api/backup":
                return {"backup": str(s.backup())}
            return None

    return H


def hello_proof(token, challenge, port):
    """Proof that this installation runs on the port: HMAC of the key over the challenge and the port. The key never
    leaves the app, and the port in the proof stops a foreign program on another port from relaying the real app's answer."""
    return hmac.new(token.encode("utf-8"), f"{challenge}:{port}".encode("utf-8"), hashlib.sha256).hexdigest()


def port_owner(port):
    """True when this Survey Keyer installation answers on the port. Only then may the key and the file path be sent to it."""
    challenge = secrets.token_urlsafe(16)
    try:
        with urllib.request.urlopen(f"http://{HOST}:{port}/api/hello?n={challenge}", timeout=1) as r:  # nosec B310
            j = json.loads(r.read())
    except Exception:
        return False
    if not isinstance(j, dict) or j.get("app") != APP_ID:
        return False
    return hmac.compare_digest(str(j.get("proof", "")).encode("utf-8", "replace"),
                               hello_proof(app_token(), challenge, port).encode("utf-8"))


def open_in_running(path):
    """Second instance of the program (e.g. a file dropped onto SurveyKeyer.exe) - hand the file over to the running app."""
    try:
        req = urllib.request.Request(
            f"http://{HOST}:{PORT}/api/open", data=json.dumps({"path": str(Path(path).resolve())}).encode("utf-8"),
            headers={"X-Survey-Keyer-Token": app_token(), "Content-Type": "application/json"})
        urllib.request.urlopen(req, timeout=15).read()  # nosec B310 - fixed address http://127.0.0.1
    except Exception:
        pass


class Server(ThreadingHTTPServer):
    allow_reuse_address = os.name != "nt"     # on Windows SO_REUSEADDR would let another program take over the port
    daemon_threads = True


def start_server(app, start):
    for p in range(start, start + 30):
        try:
            return Server((HOST, p), make_handler(app))
        except OSError:
            continue
    raise RuntimeError("No free port.")


def watchdog(app, srv):
    """Shuts the app down once the last browser tab is closed (or the "Quit" button is pressed)."""
    prev = time.time()
    while True:
        time.sleep(2)
        now = time.time()
        if now - prev > 30:                       # the computer was asleep - give the tabs time to check in
            for t in list(app.tabs):
                app.tabs[t] = now
        prev = now
        for t, seen in list(app.tabs.items()):
            if now - seen > 300:                  # no signal for 5 minutes (e.g. the browser crashed)
                app.tabs.pop(t, None)
                if not app.tabs:
                    app.closing_at = now
        if app.closing_at and now >= app.closing_at and not app.tabs:
            srv.shutdown()
            return


def main():
    args = [a for a in sys.argv[1:] if a != "--demo"]
    try:
        # demo mode (--demo or DEMO=true): a sample survey with made-up data instead of an organisation's file
        target = str(demo_file()) if is_demo(sys.argv[1:]) else (args[0] if args else None)
        if port_owner(PORT):                         # the app is already running - just show the window
            if target:
                open_in_running(target)
            webbrowser.open(f"http://{HOST}:{PORT}/?t={app_token()}")
            return
        app = App()
        if target:
            app.open(target)
        srv = start_server(app, PORT)
        port = srv.server_address[1]
        url = f"http://{HOST}:{port}/"
        print(f"Survey Keyer: {url}  (close the tab or press Ctrl+C to quit)")
        threading.Thread(target=watchdog, args=(app, srv), daemon=True).start()
        if not os.environ.get("SURVEY_KEYER_NO_BROWSER"):    # variable for automated tests
            threading.Timer(0.6, lambda: webbrowser.open(f"{url}?t={app.token}")).start()
        srv.serve_forever()
    except KeyboardInterrupt:
        pass
    except Exception as e:
        fail_box(f"Could not start the app:\n{e}")


# --------------------------------------------------------------------------
# User interface
# --------------------------------------------------------------------------
PAGE = r"""<!doctype html><html lang="en"><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Survey Keyer</title>
<style>
:root{
 --paper:#f5f3ee; --sheet:#fffefb; --ink:#1c1f1b; --ink-2:#4a5048; --ink-3:#6b7268;
 --line:#dedad0; --line-2:#ebe8e0; --green:#1d5c3e; --green-2:#2b7a54; --green-wash:#e6f0e9;
 --amber:#8a5a00; --amber-wash:#fbf1dc; --red:#a3271c; --red-wash:#fbe9e6;
 --radius:6px; --shadow:0 1px 2px rgba(28,31,27,.06),0 4px 14px rgba(28,31,27,.06);
 --font:"Segoe UI Variable Text","Segoe UI",system-ui,sans-serif;
}
*{box-sizing:border-box}
html{color-scheme:light}
body{margin:0;background:var(--paper);color:var(--ink);font:15px/1.5 var(--font);-webkit-font-smoothing:antialiased}
::selection{background:#cfe3d6}
:focus-visible{outline:2px solid var(--green-2);outline-offset:2px}
button{font:inherit;color:inherit;cursor:pointer}
input,textarea,select{font:inherit;color:var(--ink);caret-color:var(--green)}
input::placeholder,textarea::placeholder{color:var(--ink-3);font-weight:400;opacity:1}
input[type=radio],input[type=checkbox]{accent-color:var(--green);width:17px;height:17px;margin:0;flex:none}
::-webkit-scrollbar{width:11px;height:11px}::-webkit-scrollbar-thumb{background:#cbc6b9;border-radius:6px;border:3px solid var(--paper)}
.num{font-variant-numeric:tabular-nums}
[hidden]{display:none!important}

/* top bar */
.bar{position:sticky;top:0;z-index:20;display:flex;align-items:center;gap:18px;padding:0 20px;height:56px;
 background:var(--sheet);border-bottom:1px solid var(--line)}
.brand{font-weight:650;letter-spacing:-.01em;display:flex;align-items:center;gap:9px;white-space:nowrap}
.brand svg{color:var(--green)}
.file{color:var(--ink-3);font-size:13.5px;max-width:28ch;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.tabs{display:flex;gap:2px;margin-left:8px}
.tab{border:0;background:none;padding:8px 12px;border-radius:var(--radius);color:var(--ink-2);font-size:14px}
.tab:hover{background:var(--line-2)}
.tab[aria-current=true]{color:var(--green);background:var(--green-wash);font-weight:600}
.spacer{flex:1}
.ghost{border:1px solid transparent;background:none;padding:7px 11px;border-radius:var(--radius);color:var(--ink-2);font-size:13.5px}
.ghost:hover{background:var(--line-2);color:var(--ink)}

/* buttons */
.btn{display:inline-flex;align-items:center;gap:8px;border:1px solid var(--line);background:var(--sheet);
 padding:8px 14px;border-radius:var(--radius);font-size:14px;transition:background .15s,border-color .15s}
.btn:hover{border-color:#c9c4b6;background:#fbfaf6}
.btn.primary{background:var(--green);border-color:var(--green);color:#fff;font-weight:600}
.btn.primary:hover{background:var(--green-2)}
.btn:disabled{opacity:.5;cursor:default}
kbd{font:600 11px/1 var(--font);padding:3px 5px;border-radius:4px;border:1px solid currentColor;opacity:.75}
.btn.primary kbd{border-color:rgba(255,255,255,.6)}

/* layout */
.view{max-width:1180px;margin:0 auto;padding:28px 24px 120px}
h1{font-size:24px;letter-spacing:-.02em;margin:0 0 6px;font-weight:650}
h2{font-size:17px;margin:0 0 10px;font-weight:620}
.lead{color:var(--ink-2);margin:0 0 24px;max-width:70ch}
.muted{color:var(--ink-3)}
.panel{background:var(--sheet);border:1px solid var(--line);border-radius:8px;box-shadow:var(--shadow)}

/* start */
.start{display:grid;grid-template-columns:minmax(0,1.4fr) minmax(0,1fr);gap:24px;align-items:start}
.list{list-style:none;margin:0;padding:6px}
.list button{width:100%;display:flex;flex-direction:column;align-items:flex-start;gap:1px;text-align:left;
 border:0;background:none;padding:10px 12px;border-radius:var(--radius)}
.list button:hover{background:var(--green-wash)}
.list b{font-weight:600}.list span{font-size:12.5px;color:var(--ink-3)}
.sect{padding:16px 18px 6px;border-bottom:1px solid var(--line-2);display:flex;justify-content:space-between;align-items:baseline}
.side-help{padding:18px 20px}
.side-help ol{margin:8px 0 0;padding-left:20px;color:var(--ink-2)}
.side-help li{margin:6px 0}
.pathrow{display:flex;gap:8px;padding:14px 18px 18px}
.pathrow input{flex:1;border:1px solid var(--line);border-radius:var(--radius);padding:8px 10px;background:#fff}

/* entry */
.entry{display:grid;grid-template-columns:230px minmax(0,1fr);gap:24px;align-items:start}
.nav{position:sticky;top:76px;max-height:calc(100vh - 170px);overflow:auto;padding:10px}
.nav .nav-title{font-size:12.5px;font-weight:600;color:var(--ink-3);margin:4px 6px 8px}
.qnav{display:grid;grid-template-columns:repeat(5,1fr);gap:5px}
.qnav button{border:1px solid var(--line);background:#fff;border-radius:5px;height:32px;font-size:13px;color:var(--ink-2)}
.qnav button.done{background:var(--green-wash);border-color:#bcd7c6;color:var(--green);font-weight:600}
.qnav button.active{outline:2px solid var(--green);outline-offset:1px}
.nav .keys{margin-top:14px;padding:10px 8px 4px;border-top:1px solid var(--line-2);font-size:12.5px;color:var(--ink-3)}
.nav .keys div{display:flex;justify-content:space-between;gap:8px;margin:5px 0}
.nav .keys kbd{color:var(--ink-2)}

.survey-head{display:flex;align-items:center;gap:14px;flex-wrap:wrap;margin-bottom:14px}
.stepper{display:flex;align-items:center;border:1px solid var(--line);border-radius:var(--radius);background:#fff;overflow:hidden}
.stepper button{border:0;background:none;width:34px;height:36px;color:var(--ink-2)}
.stepper button:hover{background:var(--line-2)}
.stepper input{width:62px;height:36px;border:0;border-left:1px solid var(--line);border-right:1px solid var(--line);
 text-align:center;font-weight:650;font-size:16px;-moz-appearance:textfield}
.stepper input::-webkit-inner-spin-button{display:none}
.badge{font-size:12.5px;padding:3px 9px;border-radius:20px;background:var(--line-2);color:var(--ink-2)}
.badge.ok{background:var(--green-wash);color:var(--green)}
.badge.warn{background:var(--amber-wash);color:var(--amber)}
.progress{margin-left:auto;display:flex;align-items:center;gap:10px;font-size:13px;color:var(--ink-3)}
.meter{width:140px;height:6px;border-radius:4px;background:var(--line-2);overflow:hidden}
.meter i{display:block;height:100%;background:var(--green);transition:width .4s cubic-bezier(.16,1,.3,1)}

.q{background:var(--sheet);border:1px solid var(--line);border-radius:8px;padding:16px 18px 12px;margin:0 0 12px;
 transition:border-color .15s,box-shadow .2s}
.q.active{border-color:var(--green-2);box-shadow:0 0 0 3px rgba(43,122,84,.12),var(--shadow)}
.q-title{display:flex;gap:12px;align-items:baseline;margin-bottom:10px}
.q-num{flex:none;font-weight:700;color:var(--green);min-width:24px}
.q-text{font-weight:600;max-width:75ch}
.q-kind{margin-left:auto;flex:none;font-size:12px;color:var(--ink-3);white-space:nowrap}
.opt{display:flex;align-items:center;gap:11px;padding:6px 8px;border-radius:5px;cursor:pointer}
.opt:hover{background:#f6f5f0}
.opt .key{flex:none;width:21px;height:21px;border-radius:4px;border:1px solid var(--line);display:grid;place-items:center;
 font-size:11.5px;font-weight:600;color:var(--ink-3);background:#fff}
.opt.on .key{background:var(--green);border-color:var(--green);color:#fff}
.opt.on{background:var(--green-wash)}
.opt input[type=text]{flex:1;min-width:140px;border:1px solid var(--line);border-radius:5px;padding:5px 8px;background:#fff}
.opt input[type=text]:disabled{background:#f3f2ee;border-style:dashed}
.amount{display:grid;grid-template-columns:minmax(160px,1fr) 170px;gap:6px 14px;align-items:center;padding:2px 8px}
.amount input{border:1px solid var(--line);border-radius:5px;padding:6px 9px;text-align:right;background:#fff}
.q textarea{width:100%;min-height:74px;border:1px solid var(--line);border-radius:6px;padding:8px 10px;resize:vertical;background:#fff}

.actions{position:fixed;left:0;right:0;bottom:0;z-index:20;background:rgba(255,254,251,.96);border-top:1px solid var(--line);
 display:flex;justify-content:center;gap:10px;padding:12px 20px}

/* settings */
.cfg-table{width:100%;border-collapse:collapse}
.cfg-table th{font-size:12.5px;font-weight:600;color:var(--ink-3);text-align:left;padding:10px 14px;border-bottom:1px solid var(--line)}
.cfg-table td{padding:12px 14px;border-bottom:1px solid var(--line-2);vertical-align:top}
.cfg-table tr:last-child td{border-bottom:0}
.seg{display:inline-flex;border:1px solid var(--line);border-radius:var(--radius);overflow:hidden;background:#fff}
.seg button{border:0;background:none;padding:6px 10px;font-size:13px;color:var(--ink-2);white-space:nowrap}
.seg button+button{border-left:1px solid var(--line)}
.seg button[aria-pressed=true]{background:var(--green);color:#fff;font-weight:600}
.chk{display:flex;gap:8px;align-items:center;font-size:13.5px;color:var(--ink-2);margin:3px 0}
.warn-note{margin-top:8px;font-size:12.5px;color:var(--amber);background:var(--amber-wash);padding:6px 9px;border-radius:5px;max-width:320px}
.layout-grid{display:flex;gap:16px;flex-wrap:wrap;padding:16px 18px}
.layout-grid label{display:flex;flex-direction:column;gap:4px;font-size:13px;color:var(--ink-2)}
.layout-grid input,.layout-grid select{border:1px solid var(--line);border-radius:5px;padding:6px 8px;background:#fff;width:120px}
.layout-grid select{width:180px}
details.panel>summary{padding:14px 18px;cursor:pointer;font-weight:600;list-style:none}
details.panel>summary::after{content:"Show";float:right;font-weight:400;color:var(--ink-3);font-size:13px}
details[open].panel>summary::after{content:"Hide"}

/* stats */
.stat{padding:18px 20px;margin-bottom:12px}
.stat-head{display:flex;gap:12px;align-items:baseline;margin-bottom:12px}
.stat-row{display:grid;grid-template-columns:minmax(180px,1.6fr) minmax(120px,2fr) 64px 64px;gap:14px;align-items:center;padding:5px 0;font-size:14px}
.stat-row.amt{grid-template-columns:minmax(180px,1.6fr) 70px 120px 120px}
.bar-track{height:10px;border-radius:5px;background:var(--line-2);overflow:hidden}
.bar-track i{display:block;height:100%;background:var(--green-2);border-radius:5px}
.r{text-align:right}
.answers{margin:8px 0 0;padding:0;list-style:none;max-height:260px;overflow:auto;border-top:1px solid var(--line-2)}
.answers li{padding:7px 2px;border-bottom:1px solid var(--line-2);font-size:14px;color:var(--ink-2)}
.answers li b{color:var(--ink-3);font-weight:500;margin-right:8px}

/* toast */
.toast{position:fixed;left:50%;bottom:84px;transform:translate(-50%,12px);opacity:0;z-index:30;pointer-events:none;
 background:var(--ink);color:#fff;padding:10px 16px;border-radius:7px;font-size:14px;box-shadow:0 8px 24px rgba(0,0,0,.18);
 transition:opacity .2s,transform .3s cubic-bezier(.16,1,.3,1);max-width:min(90vw,640px)}
.toast.show{opacity:1;transform:translate(-50%,0)}
.toast.err{background:var(--red)}
.empty{padding:40px;text-align:center;color:var(--ink-3)}

/* more buttons */
.btn.danger{background:var(--red);border-color:var(--red);color:#fff;font-weight:600}
.btn.danger:hover{background:#8c2117}
.btn svg,.ghost svg{flex:none}
.icon-btn{border:1px solid transparent;background:none;width:32px;height:32px;border-radius:var(--radius);
 display:inline-grid;place-items:center;color:var(--ink-3);transition:background .15s,color .15s}
.icon-btn:hover{background:var(--line-2);color:var(--ink)}
.icon-btn:disabled{opacity:.3;cursor:default;background:none}
.icon-btn.del:hover{background:var(--red-wash);color:var(--red)}
a.link,button.link{border:0;background:none;padding:0;color:var(--green);font:inherit;font-weight:600;text-decoration:underline;
 text-underline-offset:3px;text-decoration-thickness:1px;cursor:pointer}
button.link:hover{color:var(--green-2)}

/* banners */
.banner{position:sticky;top:56px;z-index:19;display:flex;align-items:center;justify-content:center;gap:10px;
 padding:9px 20px;font-size:14px;border-bottom:1px solid;text-align:center}
.banner.err{background:var(--red-wash);color:var(--red);border-color:#efc9c3}
.banner.warn{background:var(--amber-wash);color:var(--amber);border-color:#ecd8ad}
.banner svg{flex:none}

/* dialog */
dialog.dlg{border:1px solid var(--line);border-radius:10px;padding:22px 24px 18px;width:min(92vw,470px);
 background:var(--sheet);color:var(--ink);box-shadow:0 2px 6px rgba(28,31,27,.08),0 22px 60px rgba(28,31,27,.22)}
dialog.dlg[open]{animation:dlg-in .26s cubic-bezier(.16,1,.3,1)}
dialog.dlg::backdrop{background:rgba(28,31,27,.34)}
dialog.dlg h2{font-size:17.5px;margin:0 0 8px}
dialog.dlg p{margin:0 0 22px;color:var(--ink-2);white-space:pre-line}
.dlg-actions{display:flex;justify-content:flex-end;gap:8px;flex-wrap:wrap}
@keyframes dlg-in{from{opacity:0;transform:translateY(8px) scale(.985)}}

/* start: new survey */
.new-panel{padding:18px 20px;margin-bottom:16px;display:grid;gap:10px;justify-items:start}
.new-panel p{margin:0;color:var(--ink-2)}

/* creator */
.create-meta{display:grid;grid-template-columns:minmax(0,1fr) 190px;gap:14px 20px;padding:18px 20px;margin-bottom:18px;align-items:end}
.create-meta .chk{grid-column:1/-1;margin:0}
.field{display:flex;flex-direction:column;gap:5px;font-size:13px;color:var(--ink-2);font-weight:500}
.field input{border:1px solid var(--line);border-radius:var(--radius);padding:8px 10px;background:#fff;font-size:15px;font-weight:400}
.field input.title-in{font-size:17px;font-weight:600}
.cq .q-title{align-items:center;margin-bottom:12px}
.cq .q-title input{flex:1;min-width:0;border:1px solid var(--line);border-radius:var(--radius);padding:8px 10px;background:#fff;font-weight:600}
.cq-tools{display:flex;gap:2px;flex:none}
.cq-body{display:grid;gap:12px;padding:0 0 4px 36px}
.cq textarea{min-height:104px;max-height:340px;field-sizing:content;font-size:14.5px;line-height:1.65}
.cq-body .seg{justify-self:start}
.cq .hint{font-size:12.5px;color:var(--ink-3);margin:-6px 0 0}
.q.invalid{border-color:var(--red);box-shadow:0 0 0 3px rgba(163,39,28,.1)}
.err-note{font-size:13px;color:var(--red);font-weight:500}
.add-q{width:100%;justify-content:center;border-style:dashed;padding:12px;color:var(--ink-2)}
.add-q:hover{color:var(--green)}
.create-head{display:flex;align-items:baseline;gap:14px;flex-wrap:wrap}
.create-head .muted{font-size:13.5px}

/* stats: empty surveys */
.gaps{padding:14px 18px;margin-bottom:14px;font-size:14px;color:var(--ink-2);line-height:1.9}
.gaps b{color:var(--ink);font-weight:600;margin-right:6px}
.gaps .link{margin:0 4px}
.not-num{font-size:12.5px;color:var(--amber)}

@media (max-width:900px){.entry,.start{grid-template-columns:1fr}.nav{position:static;max-height:none}.file{display:none}
 .create-meta{grid-template-columns:1fr}.cq-body{padding-left:0}}
@media (prefers-reduced-motion:reduce){*{transition:none!important;animation:none!important}}

/* paper form (print only) */
#v-print{display:none}
@media print{
 @page{size:A4;margin:14mm 15mm 14mm}
 html,body{background:#fff}
 body>*:not(#v-print){display:none!important}
 #v-print{display:block;color:#111;font:10.5pt/1.45 "Segoe UI",Arial,sans-serif}
 .pf-head{display:flex;justify-content:space-between;align-items:flex-end;gap:16px;border-bottom:1.2pt solid #111;padding-bottom:6pt;margin-bottom:8pt}
 .pf-title{font-size:17pt;font-weight:700;margin:0;letter-spacing:-.01em;line-height:1.2}
 .pf-no{white-space:nowrap;font-size:10pt}
 .pf-hint{margin:0 0 12pt;font-size:9.5pt;color:#333}
 .pf-q{break-inside:avoid;margin:0 0 11pt}
 .pf-qt{font-weight:600;margin-bottom:4pt}
 .pf-qt .pf-k{font-weight:400;color:#444;font-size:9pt;margin-left:4pt}
 .pf-opts{list-style:none;margin:0;padding:0 0 0 14pt}
 .pf-opts.cols2{columns:2;column-gap:18pt}.pf-opts.cols3{columns:3;column-gap:14pt}
 .pf-opts li{display:flex;align-items:baseline;gap:7pt;padding:1.5pt 0;break-inside:avoid}
 .box{flex:none;display:inline-block;width:9pt;height:9pt;border:1pt solid #111;border-radius:1.5pt;transform:translateY(1pt)}
 .box.round{border-radius:50%}
 .pf-line{flex:1;min-width:90pt;border-bottom:.8pt solid #555;height:11pt}
 .pf-ruled{margin:0 0 0 14pt}
 .pf-ruled div{border-bottom:.8pt solid #777;height:17pt}
 .pf-amt{display:grid;grid-template-columns:auto 130pt;gap:5pt 12pt;padding-left:14pt;align-items:end}
 .pf-foot{margin-top:16pt;padding-top:6pt;border-top:.6pt solid #999;font-size:9pt;color:#444;text-align:center}
}
</style>

<header class="bar">
 <div class="brand"><svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="4" y="3" width="16" height="18" rx="2"/><path d="M8 8h8M8 12h8M8 16h5"/></svg>Survey Keyer</div>
 <span class="file" id="file"></span>
 <nav class="tabs" id="tabs" hidden>
  <button class="tab" data-view="entry">Data entry</button>
  <button class="tab" data-view="settings">Question settings</button>
  <button class="tab" data-view="stats">Statistics</button>
 </nav>
 <span class="spacer"></span>
 <button class="ghost" id="toFiles" hidden>Change file</button>
 <button class="ghost" id="quit">Quit</button>
</header>
<div class="banner err" id="offline" role="alert" hidden>
 <svg width="17" height="17" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M12 9v4M12 17h.01"/><path d="M10.3 3.9 2.4 18a2 2 0 0 0 1.7 3h15.8a2 2 0 0 0 1.7-3L13.7 3.9a2 2 0 0 0-3.4 0z"/></svg>
 <span id="offlineText">Lost connection to the Survey Keyer app. Start it again (SurveyKeyer.exe): this tab will reconnect by itself, and the answers you typed stay in the form.</span>
</div>
<div class="banner warn" id="lockedBar" role="status" hidden>
 <svg width="17" height="17" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><rect x="5" y="11" width="14" height="10" rx="2"/><path d="M8 11V7a4 4 0 0 1 8 0v4"/></svg>
 <span>This file is open in Excel. Close it in Excel before you save the survey, or saving will fail.</span>
</div>

<!-- start -->
<section class="view" id="v-start">
 <h1>Choose a survey file</h1>
 <p class="lead">The app reads the questions from an Excel sheet. The answers are saved straight into that file, and a backup copy is made before the first save.</p>
 <div class="start">
  <div class="panel">
   <div class="sect"><h2>Recently opened</h2></div>
   <ul class="list" id="recent"><li class="empty">None yet. Open a file below.</li></ul>
   <div class="sect"><h2>Surveys nearby</h2><button class="ghost" id="rescan">Search again</button></div>
   <ul class="list" id="found"><li class="empty">Looking for files…</li></ul>
   <div class="pathrow">
    <button class="btn primary" id="browse">Choose a file from disk…</button>
    <input id="path" placeholder="or paste the path to an .xlsx file and press Enter">
   </div>
  </div>
  <div>
  <div class="panel new-panel">
   <h2 style="margin:0">New survey</h2>
   <p>No file yet? Type the questions into the wizard and the app will prepare an Excel sheet with charts and a form to print.</p>
   <button class="btn primary" id="toCreate">Create a new survey</button>
  </div>
  <aside class="panel side-help">
   <h2>What the sheet should look like</h2>
   <ol>
    <li>A row of questions: each question cell is merged across its answers.</li>
    <li>Below it, a row of answers, one per column.</li>
    <li>A totals row (e.g. “Total”), optional.</li>
    <li>Then one row per survey, with its number in column A.</li>
   </ol>
   <p class="muted" style="margin:12px 0 0">If your layout is different, correct the row numbers in “Question settings” → “Sheet layout”. Save .xls and .ods files as .xlsx in Excel first.</p>
  </aside>
  </div>
 </div>
</section>

<!-- creator -->
<section class="view" id="v-create" hidden>
 <div class="create-head"><h1>New survey</h1><span class="muted num" id="c-sum"></span><span class="spacer"></span><button class="ghost" id="c-back">Close wizard</button></div>
 <p class="lead">Type the questions and answers. The app will create an Excel file in a layout ready for data entry, with a totals row and a charts sheet. The wizard remembers what you type, so you can stop and come back later.</p>
 <div class="panel create-meta">
  <label class="field">Survey title<input id="c-title" class="title-in" maxlength="200" placeholder="e.g. Residents' satisfaction survey 2026"></label>
  <label class="field">Number of surveys (rows)<input id="c-rows" type="number" min="1" max="5000" value="100" class="num"></label>
  <label class="chk"><input type="checkbox" id="c-charts" checked> Add a sheet with answer charts</label>
 </div>
 <div id="c-list"></div>
 <button class="btn add-q" id="c-add"><svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" aria-hidden="true"><path d="M12 5v14M5 12h14"/></svg>Add question</button>
</section>

<!-- settings -->
<section class="view" id="v-settings" hidden>
 <h1>Question settings</h1>
 <p class="lead">Check the type of each question. Tick the answers that need a text field, e.g. “Other (please specify)”. The settings are saved next to the file.</p>
 <details class="panel" id="layoutBox" style="margin-bottom:16px">
  <summary>Sheet layout</summary>
  <div class="layout-grid">
   <label>Sheet<select id="l-sheet"></select></label>
   <label>Question row<input id="l-header" type="number" min="1"></label>
   <label>Answer row<input id="l-option" type="number" min="1"></label>
   <label>First survey row<input id="l-first" type="number" min="1"></label>
  </div>
 </details>
 <div class="panel" style="padding:4px 8px"><table class="cfg-table"><thead><tr><th style="width:44px">No.</th><th>Question</th><th style="width:330px">Type</th><th style="width:300px">Text field for answer</th></tr></thead>
  <tbody id="cfg"></tbody></table></div>
</section>

<!-- entry -->
<section class="view" id="v-entry" hidden>
 <div class="entry">
  <aside class="panel nav">
   <h2 class="nav-title">Questions</h2>
   <div class="qnav" id="qnav"></div>
   <div class="keys">
    <div><span>Tick an answer</span><span><kbd>1</kbd>–<kbd>9</kbd></span></div>
    <div><span>Next question</span><kbd>Enter</kbd></div>
    <div><span>Previous / next</span><span><kbd>↑</kbd> <kbd>↓</kbd></span></div>
    <div><span>Clear question</span><kbd>0</kbd></div>
    <div><span>Save survey</span><kbd>Ctrl+Enter</kbd></div>
   </div>
  </aside>
  <div>
   <div class="survey-head">
    <span class="muted">Survey no.</span>
    <div class="stepper"><button id="prev" aria-label="Previous survey">‹</button><input id="seq" type="number" min="1" class="num"><button id="next" aria-label="Next survey">›</button></div>
    <span class="badge" id="seqState"></span>
    <button class="ghost" id="nextEmpty">Next empty</button>
    <button class="ghost" id="newRow">Add survey</button>
    <div class="progress"><span class="num" id="progressText"></span><div class="meter"><i id="meter"></i></div></div>
   </div>
   <div id="form"></div>
  </div>
 </div>
</section>

<!-- stats -->
<section class="view" id="v-stats" hidden>
 <div style="display:flex;align-items:baseline;gap:14px;flex-wrap:wrap">
  <h1>Statistics</h1><span class="muted num" id="statsInfo"></span><span class="spacer"></span>
  <button class="btn" id="csv">Download CSV</button>
 </div>
 <p class="lead">Percentages are based on the people who answered each question.</p>
 <div class="panel gaps" id="gaps" hidden></div>
 <div id="stats"></div>
</section>

<!-- paper form, visible only when printing -->
<section id="v-print" aria-hidden="true"></section>

<footer class="actions" id="a-entry" hidden>
 <button class="btn primary" id="save">Save and next <kbd>Ctrl+Enter</kbd></button>
 <button class="btn" id="clear">Clear form</button>
 <button class="btn" id="backup">Back up file</button>
</footer>
<footer class="actions" id="a-settings" hidden>
 <button class="btn primary" id="saveCfg">Save settings</button>
 <button class="btn" id="printForm">Print form</button>
 <button class="btn" id="copyToCreate">New file with these questions</button>
</footer>
<footer class="actions" id="a-create" hidden>
 <button class="btn primary" id="c-make">Create Excel file</button>
 <button class="btn" id="c-print">Print preview</button>
 <button class="btn" id="c-reset">Clear wizard</button>
</footer>
<div class="toast" id="toast" role="status" aria-live="polite"></div>
<dialog class="dlg" id="dlg" aria-labelledby="dlgTitle" aria-describedby="dlgMsg">
 <h2 id="dlgTitle"></h2><p id="dlgMsg"></p>
 <div class="dlg-actions"><button class="btn" id="dlgCancel"></button><button class="btn primary" id="dlgOk"></button></div>
</dialog>

<script nonce="__NONCE__">
const $ = id => document.getElementById(id);
/* the key arrives only in the URL opened by the app (?t=); the tab keeps it in sessionStorage and clears the URL */
const TOKEN = (() => {
  const t = new URLSearchParams(location.search).get('t');
  if (t) history.replaceState(null, '', location.pathname);
  try { if (t) sessionStorage.setItem('survey-keyer-token', t); return t || sessionStorage.getItem('survey-keyer-token') || ''; }
  catch(e){ return t || ''; }
})(), TAB = Math.random().toString(36).slice(2) + Date.now().toString(36);
const KIND = {single:'Single', multi:'Multiple', open:'Open', amount:'Amounts'};
const KIND_HINT = {single:'one answer', multi:'several allowed', open:'open answer', amount:'type amounts or numbers'};
const VIEWS = ['start','settings','entry','stats','create'];
let S = null, view = 'start', active = 0, filled = new Set(), dirty = false, formReady = false;
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const fmt = n => Number(n).toLocaleString('en-US');
const plural = (n, one, many) => n === 1 ? one : many;

let toastTimer;
function toast(msg, err){
  const t = $('toast'); t.textContent = msg; t.className = 'toast show' + (err ? ' err' : '');
  clearTimeout(toastTimer); toastTimer = setTimeout(() => t.className = 'toast', err ? 6500 : 2600);
}

/* confirmation dialog instead of the native confirm() */
function ask(msg, o = {}){
  return new Promise(res => {
    const d = $('dlg'), ok = $('dlgOk'), no = $('dlgCancel');
    $('dlgTitle').textContent = o.title || 'Confirm'; $('dlgMsg').textContent = msg;
    ok.textContent = o.ok || 'OK'; ok.className = 'btn ' + (o.danger ? 'danger' : 'primary');
    no.hidden = o.cancel === null; no.textContent = o.cancel || 'Cancel';
    const done = v => { d.close(); res(v); };
    ok.onclick = () => done(true); no.onclick = () => done(false);
    d.oncancel = e => { e.preventDefault(); done(false); };
    d.showModal(); (o.focusCancel && o.cancel !== null ? no : ok).focus();
  });
}

/* connection to the app */
function headers(){
  const h = {'Content-Type':'application/json', 'X-Survey-Keyer-Token':TOKEN, 'X-Survey-Keyer-Tab':TAB};
  if (S) { h['X-Survey-Keyer-Path'] = encodeURIComponent(S.path); h['X-Survey-Keyer-Rev'] = S.rev; }
  return h;
}
let offline = false, conflictOpen = false;
function setOffline(on, text){
  if (text) $('offlineText').textContent = text;
  if (on === offline && !text) return;
  offline = on; $('offline').hidden = !on;
  if (on) schedulePing(4000);
  if (!on) { $('offlineText').textContent = 'Lost connection to the Survey Keyer app. Start it again (SurveyKeyer.exe): this tab will reconnect by itself, and the answers you typed stay in the form.'; }
}
async function request(url, body){
  try { return await fetch(url, {method: body !== undefined ? 'POST' : 'GET', headers: headers(), body: body !== undefined ? JSON.stringify(body) : undefined}); }
  catch(e){ setOffline(true); throw new Error('No connection to the app. Start Survey Keyer again and try once more.'); }
}
async function api(url, body){
  const r = await request(url, body);
  let j = null; try { j = await r.json(); } catch(e){}
  if (r.status === 403 && j && j.auth) { setOffline(true, j.error); throw new Error(j.error); }
  setOffline(false);
  if (r.status === 409 && j && j.conflict) { onConflict(j.error); throw new Error(j.error); }
  if (!r.ok) throw new Error((j && j.error) || ('App error (' + r.status + ').'));
  return j;
}
async function onConflict(msg){
  if (conflictOpen) return; conflictOpen = true;
  await ask(msg, {title:'The file has changed', ok:'Reload', cancel:null});
  conflictOpen = false; dirty = false;
  const path = S.path; S = null;
  await openFile(path);
}
async function ping(){
  if (offline || !S) {
    try { const r = await request('/api/ping'); if (r.ok) { setOffline(false); if (S) refreshProgress().catch(() => {}); } } catch(e){}
    return;
  }
  try { const r = await request('/api/ping'); if (r.ok) $('lockedBar').hidden = !(await r.json()).locked || view !== 'entry'; } catch(e){}
}

function show(v){
  view = v;
  VIEWS.forEach(x => $('v-'+x).hidden = x !== v);
  $('a-entry').hidden = v !== 'entry'; $('a-settings').hidden = v !== 'settings'; $('a-create').hidden = v !== 'create';
  $('tabs').hidden = $('toFiles').hidden = !S;
  if (v !== 'entry') $('lockedBar').hidden = true;
  document.querySelectorAll('.tab').forEach(b => b.setAttribute('aria-current', b.dataset.view === v));
  if (v === 'settings') buildSettings();
  if (v === 'stats') loadStats();
  if (v === 'start') loadFiles();
  if (v === 'create') buildCreator();
  window.scrollTo(0, 0);
}

/* ---------- files ---------- */
function fileItem(f, sub){
  const li = document.createElement('li'), b = document.createElement('button');
  b.innerHTML = '<b>' + esc(f.name) + '</b><span>' + esc(sub) + '</span>';
  b.onclick = () => openFile(f.path); li.appendChild(b); return li;
}
async function loadFiles(refresh){
  try {
    const r = await api('/api/files?scan=0');
    $('recent').innerHTML = '';
    if (!r.recent.length) $('recent').innerHTML = '<li class="empty">None yet. Open a file below.</li>';
    r.recent.forEach(f => $('recent').appendChild(fileItem(f, f.dir)));
    const j = await api('/api/files' + (refresh ? '?refresh=1' : ''));
    $('found').innerHTML = '';
    if (!j.found.length) $('found').innerHTML = '<li class="empty">No surveys found in ' + esc(j.root) + '.</li>';
    j.found.forEach(f => $('found').appendChild(fileItem(f, 'in folder ' + j.root.split(/[\\/]/).pop())));
  } catch(e){ toast(e.message, true); }
}
async function openFile(path){
  try { const s = await api('/api/open', {path}); S = null; setSchema(s); afterOpen(); } catch(e){ toast(e.message, true); }
}
async function browse(){
  try { const s = await api('/api/browse', {}); if (s.cancelled) return; S = null; setSchema(s); afterOpen(); }
  catch(e){ toast(e.message, true); }
}
function setSchema(s){ S = s; $('file').textContent = s.file; $('file').title = s.path; }
async function afterOpen(){
  formReady = false; dirty = false; $('seq').dataset.cur = '';
  if (S.dupes && S.dupes.length) toast('Survey numbers repeat in column A: ' + S.dupes.join(', ') + '. Fix them in Excel, otherwise some rows will not be available.', true);
  if (!S.questions.length){ show('settings'); $('layoutBox').open = true; toast('No questions found. Check the sheet layout.', true); return; }
  if (!S.has_config){ show('settings'); toast(S.warning || 'Check the question types and save the settings.', !!S.warning); return; }
  await startEntry();
}
async function startEntry(seq){
  show('entry'); buildForm(); formReady = true;
  await refreshProgress();
  if (seq && S.ids.includes(seq)) await loadSeq(seq, true); else await gotoNextEmpty(0, true);
}

/* ---------- settings ---------- */
function buildSettings(){
  const L = S.layout;
  $('l-sheet').innerHTML = S.sheets.map(n => '<option' + (n === L.sheet ? ' selected' : '') + '>' + esc(n) + '</option>').join('');
  $('l-header').value = L.header_row; $('l-option').value = L.option_row; $('l-first').value = L.first_row;
  const tb = $('cfg'); tb.innerHTML = '';
  if (!S.questions.length) tb.innerHTML = '<tr><td colspan="4" class="empty">No questions in row ' + L.header_row + '. Change the sheet layout above and save.</td></tr>';
  S.questions.forEach(q => {
    const tr = document.createElement('tr'); tr.dataset.key = q.key; tr.dataset.kind = q.kind;
    const seg = '<div class="seg">' + Object.entries(KIND).map(([k, v]) => '<button type="button" data-k="'+k+'" aria-pressed="'+(k===q.kind)+'">'+v+'</button>').join('') + '</div>';
    const opts = q.options.length ? q.options.map(o => '<label class="chk"><input type="checkbox" data-col="'+o.col+'"'+(o.text?' checked':'')+'>'+esc(o.label)+'</label>').join('') : '<span class="muted">—</span>';
    const warn = q.multi_seen ? '<div class="warn-note" hidden>'+q.multi_seen+' '+plural(q.multi_seen,'survey has','surveys have')+' several answers ticked. With single choice, they need fixing.</div>' : '';
    tr.innerHTML = '<td class="num muted">'+q.num+'</td><td>'+esc(q.text)+'<div class="muted" style="font-size:12.5px;margin-top:3px">'+q.options.length+' '+plural(q.options.length,'answer','answers')+'</div></td><td>'+seg+warn+'</td><td>'+opts+'</td>';
    const syncWarn = () => { const w = tr.querySelector('.warn-note'); if (w) w.hidden = tr.dataset.kind !== 'single'; };
    tr.querySelectorAll('.seg button').forEach(b => b.onclick = () => {
      tr.dataset.kind = b.dataset.k; tr.querySelectorAll('.seg button').forEach(x => x.setAttribute('aria-pressed', x === b)); syncWarn();
    });
    syncWarn();
    tb.appendChild(tr);
  });
}
async function saveSettings(){
  const config = {};
  document.querySelectorAll('#cfg tr[data-key]').forEach(tr => {
    const text_cols = {}; tr.querySelectorAll('input[type=checkbox]').forEach(c => text_cols[c.dataset.col] = c.checked);
    config[tr.dataset.key] = {kind: tr.dataset.kind, text_cols};
  });
  const layout = {sheet: $('l-sheet').value, header_row: +$('l-header').value, option_row: +$('l-option').value, first_row: +$('l-first').value};
  const changed = ['sheet','header_row','option_row','first_row'].some(k => String(layout[k]) !== String(S.layout[k]));
  if (dirty && !(await ask('The form has unsaved answers. Saving the settings will clear them.', {title:'Save settings?', ok:'Save settings', cancel:'Go back', focusCancel:true}))) return;
  const cur = +$('seq').dataset.cur || 0;
  try { setSchema(await api('/api/config', {config, layout})); } catch(e){ toast(e.message, true); return; }
  dirty = false;
  if (changed){ formReady = false; buildSettings(); toast('Layout changed: found ' + S.questions.length + ' ' + plural(S.questions.length, 'question', 'questions') + '. Check the types and save again.'); return; }
  toast('Settings saved'); await startEntry(cur);
}

/* ---------- form ---------- */
function isOpen(q){ return q.kind === 'open' || !q.options.length; }
function buildForm(){
  const form = $('form'); form.innerHTML = ''; $('qnav').innerHTML = '';
  S.questions.forEach((q, qi) => {
    const box = document.createElement('section'); box.className = 'q'; box.dataset.i = qi;
    box.innerHTML = '<div class="q-title"><span class="q-num num">'+q.num+'</span><span class="q-text">'+esc(q.text)+'</span><span class="q-kind">'+KIND_HINT[isOpen(q)?'open':q.kind]+'</span></div>';
    if (isOpen(q)) box.insertAdjacentHTML('beforeend', '<textarea aria-label="Answer to question '+q.num+'"></textarea>');
    else if (q.kind === 'amount') {
      const g = document.createElement('div'); g.className = 'amount';
      q.options.forEach((o, i) => g.insertAdjacentHTML('beforeend', '<span>'+esc(o.label)+'</span><input inputmode="decimal" class="num" data-i="'+i+'" placeholder="amount / number">'));
      box.appendChild(g);
    } else q.options.forEach((o, i) => {
      const l = document.createElement('label'); l.className = 'opt';
      l.innerHTML = '<input type="'+(q.kind==='single'?'radio':'checkbox')+'" name="q'+q.key+'" value="'+i+'" hidden><span class="key">'+(i<9?i+1:'·')+'</span><span>'+esc(o.label)+'</span>' +
        (o.text ? '<input type="text" data-i="'+i+'" placeholder="type the answer…" disabled>' : '');
      box.appendChild(l);
    });
    box.addEventListener('focusin', () => setActive(qi, false));
    box.addEventListener('mousedown', () => setActive(qi, false));
    box.addEventListener('change', () => { syncBox(box); markDirty(); });
    box.addEventListener('input', markDirty);
    form.appendChild(box);
    const nb = document.createElement('button'); nb.textContent = q.num; nb.className = 'num'; nb.title = q.text;
    nb.onclick = () => setActive(qi, true); $('qnav').appendChild(nb);
  });
  active = 0; setActive(0, false); dirty = false;
}
function syncBox(box){
  box.querySelectorAll('label.opt').forEach(l => {
    const c = l.querySelector('input[type=radio],input[type=checkbox]'), t = l.querySelector('input[type=text]');
    l.classList.toggle('on', c.checked);
    if (t) t.disabled = !c.checked;
  });
  updateNav();
}
function answered(box){
  const q = S.questions[+box.dataset.i];
  if (isOpen(q)) return !!box.querySelector('textarea').value.trim();
  if (q.kind === 'amount') return [...box.querySelectorAll('input')].some(i => i.value.trim());
  return !!box.querySelector('input:checked');
}
function updateNav(){
  document.querySelectorAll('#form .q').forEach((box, i) => {
    const b = $('qnav').children[i]; b.classList.toggle('done', answered(box)); b.classList.toggle('active', i === active);
  });
}
function setActive(i, focus){
  const boxes = document.querySelectorAll('#form .q'); if (!boxes.length) return;
  active = Math.max(0, Math.min(boxes.length - 1, i));
  boxes.forEach((b, k) => b.classList.toggle('active', k === active));
  updateNav();
  if (focus){
    const box = boxes[active];
    box.scrollIntoView({block:'center', behavior:'smooth'});
    const field = box.querySelector('textarea, .amount input');
    if (field) field.focus({preventScroll:true}); else { if (document.activeElement) document.activeElement.blur(); }
  }
}
function pick(n){
  const box = document.querySelectorAll('#form .q')[active], q = S.questions[active];
  if (!box || isOpen(q) || q.kind === 'amount') return;
  const all = box.querySelectorAll('label.opt input[type=radio],label.opt input[type=checkbox]'), inp = all[n];
  if (!inp) return;
  if (q.kind === 'single') {        // also old data with several answers: only the chosen one stays
    all.forEach(c => { c.checked = false; c.type = 'radio'; });
    const k = box.querySelector('.q-kind'); k.textContent = KIND_HINT.single; k.style.color = '';
    inp.checked = true;
  } else inp.checked = !inp.checked;
  syncBox(box); markDirty();
  const t = inp.closest('label').querySelector('input[type=text]');
  if (t && inp.checked){ t.focus(); return; }
  if (q.kind === 'single') setActive(active + 1, true);
}
function clearActive(){
  const box = document.querySelectorAll('#form .q')[active]; if (!box) return;
  box.querySelectorAll('input[type=radio],input[type=checkbox]').forEach(i => i.checked = false);
  box.querySelectorAll('input[type=text],input[inputmode],textarea').forEach(i => i.value = '');
  syncBox(box); markDirty();
}
function markDirty(){ dirty = true; updateNav(); }
function collect(){
  const ans = {};
  document.querySelectorAll('#form .q').forEach(box => {
    const q = S.questions[+box.dataset.i];
    if (isOpen(q)) { ans[q.key] = {text: box.querySelector('textarea').value}; return; }
    if (q.kind === 'amount') { const values = {}; box.querySelectorAll('.amount input').forEach(t => { if (t.value.trim()) values[t.dataset.i] = t.value; }); ans[q.key] = {values}; return; }
    const chosen = [], texts = {};
    box.querySelectorAll('label.opt input:not([type=text]):checked').forEach(c => chosen.push(+c.value));
    box.querySelectorAll('label.opt input[type=text]').forEach(t => { if (t.value.trim()) texts[t.dataset.i] = t.value; });
    ans[q.key] = {chosen, texts};
  });
  return ans;
}
function fill(ans){
  buildForm();
  document.querySelectorAll('#form .q').forEach(box => {
    const q = S.questions[+box.dataset.i], a = ans[q.key]; if (!a) return;
    if (isOpen(q)) { box.querySelector('textarea').value = a.text || ''; return; }
    if (q.kind === 'amount') { Object.entries(a.values||{}).forEach(([i,v]) => { const x = box.querySelector('.amount input[data-i="'+i+'"]'); if (x) x.value = v; }); return; }
    if (q.kind === 'single' && (a.chosen||[]).length > 1) {   // old data: several answers - show all of them
      box.querySelectorAll('label.opt input[type=radio]').forEach(r => r.type = 'checkbox');
      box.querySelector('.q-kind').textContent = 'several ticked, keep one';
      box.querySelector('.q-kind').style.color = 'var(--amber)';
    }
    (a.chosen||[]).forEach(i => { const c = box.querySelectorAll('label.opt input:not([type=text])')[i]; if (c) c.checked = true; });
    Object.entries(a.texts||{}).forEach(([i,v]) => { const t = box.querySelector('label.opt input[type=text][data-i="'+i+'"]'); if (t) t.value = v; });
    syncBox(box);
  });
  dirty = false;
}
function setState(seq, isFilled){
  const b = $('seqState'); b.textContent = isFilled ? 'filled in, saving overwrites' : 'empty'; b.className = 'badge' + (isFilled ? ' warn' : ' ok');
}
async function refreshProgress(){
  const p = await api('/api/progress'); filled = new Set(p.filled); S.ids = p.ids;
  $('progressText').textContent = fmt(p.filled.length) + ' of ' + fmt(p.total) + ' filled in';
  $('meter').style.width = (p.total ? 100 * p.filled.length / p.total : 0) + '%';
  $('seq').max = Math.max(1, ...S.ids);
  $('lockedBar').hidden = !p.locked || view !== 'entry';
}
async function confirmLeave(){
  return !dirty || await ask('The answers typed into this survey have not been saved.', {title:'Discard changes?', ok:'Discard changes', cancel:'Back to survey', danger:true, focusCancel:true});
}
async function loadSeq(seq, force){
  if (!force && !(await confirmLeave())) { $('seq').value = $('seq').dataset.cur || ''; return; }
  try {
    const j = await api('/api/state?seq=' + encodeURIComponent(seq));
    $('seq').value = seq; $('seq').dataset.cur = seq; fill(j.answers); setState(seq, j.filled);
    window.scrollTo(0, 0); setActive(0, false);
  } catch(e){ toast(e.message, true); $('seq').value = $('seq').dataset.cur || ''; }
}
async function gotoNextEmpty(after, force){
  const j = await api('/api/next' + (after ? '?after=' + after : ''));
  if (j.seq == null){
    toast('All surveys in the file are filled in. Use “Add survey” to add another one.');
    if ($('seq').dataset.cur || !S.ids.length) return; await loadSeq(S.ids[0], true); return;
  }
  await loadSeq(j.seq, force);
}
function step(d){
  const ids = S.ids, cur = +$('seq').dataset.cur, i = ids.indexOf(cur);
  const n = ids[Math.max(0, Math.min(ids.length - 1, (i < 0 ? 0 : i) + d))]; if (n !== undefined && n !== cur) loadSeq(n);
}
function missing(ans){
  return S.questions.filter(q => { const a = ans[q.key];
    return a.text !== undefined ? !a.text.trim() : a.values ? !Object.keys(a.values).length : !a.chosen.length; }).map(q => q.num);
}
let saving = false;
async function save(){
  if (saving) return;
  const seq = +$('seq').dataset.cur; if (!seq) return;
  const ans = collect(), miss = missing(ans);
  if (miss.length === S.questions.length){ toast('The form is empty, nothing was saved.', true); return; }
  saving = true;
  try {
    const notes = [];
    if (miss.length) notes.push('No answer: ' + plural(miss.length, 'question ', 'questions ') + miss.join(', ') + '.');
    if (filled.has(seq)) notes.push('Survey no. ' + seq + ' was already filled in. Saving will replace the previous answers.');
    if (notes.length && !(await ask(notes.join('\n'), {title:'Save survey no. ' + seq + '?', ok:'Save', cancel:'Back to form'}))) return;
    $('save').disabled = true;
    await api('/api/save', {seq, answers: ans}); dirty = false;
    await refreshProgress(); toast('Saved survey no. ' + seq); await gotoNextEmpty(seq, true);
  }
  catch(e){ toast(e.message, true); }
  finally { saving = false; $('save').disabled = false; }
}
async function newRow(){
  if (!(await confirmLeave())) return;
  try { const j = await api('/api/newrow', {}); S.ids = j.ids; await refreshProgress(); await loadSeq(j.seq, true); toast('Added survey no. ' + j.seq); }
  catch(e){ toast(e.message, true); }
}

/* ---------- statistics ---------- */
function ranges(nums){
  const out = []; let a = null, b = null;
  nums.forEach(n => { if (b !== null && n === b + 1) { b = n; return; } if (a !== null) out.push([a, b]); a = b = n; });
  if (a !== null) out.push([a, b]); return out;
}
async function showGaps(){
  const p = await api('/api/progress'), done = new Set(p.filled), empty = p.ids.filter(i => !done.has(i)), g = $('gaps');
  g.hidden = !p.filled.length || !empty.length; if (g.hidden) return;
  const rs = ranges(empty), shown = rs.slice(0, 40);
  g.innerHTML = '<b>Not filled in (' + fmt(empty.length) + '):</b>' + shown.map(([a, b]) =>
    '<button class="link num" data-seq="' + a + '">' + (a === b ? a : a + '–' + b) + '</button>').join(',') +
    (rs.length > shown.length ? ' <span class="muted">and more…</span>' : '');
  g.querySelectorAll('button').forEach(x => x.onclick = () => goEntry(+x.dataset.seq));
}
async function goEntry(seq){
  if (!formReady) return startEntry(seq);
  show('entry'); await loadSeq(seq);
}
async function downloadCsv(){
  try {
    const r = await request('/api/stats.csv');
    if (!r.ok) { let j = null; try { j = await r.json(); } catch(e){} throw new Error((j && j.error) || 'Could not prepare the CSV file.'); }
    const url = URL.createObjectURL(await r.blob()), a = document.createElement('a');
    a.href = url; a.download = S.file.replace(/\.[^.]+$/, '') + '_stats.csv';
    document.body.appendChild(a); a.click(); a.remove(); setTimeout(() => URL.revokeObjectURL(url), 2000);
  } catch(e){ toast(e.message, true); }
}
async function loadStats(){
  const box = $('stats'); box.innerHTML = '<div class="empty">Calculating…</div>';
  try {
    const st = await api('/api/stats');
    $('statsInfo').textContent = fmt(st.filled) + ' of ' + fmt(st.total) + ' surveys filled in';
    showGaps().catch(() => {});
    if (!st.filled){ box.innerHTML = '<div class="panel empty">No surveys have been filled in yet.</div>'; return; }
    box.innerHTML = '';
    st.questions.forEach(q => {
      const el = document.createElement('div'); el.className = 'panel stat';
      let h = '<div class="stat-head"><span class="q-num num">'+q.num+'</span><span class="q-text">'+esc(q.text)+'</span><span class="q-kind num">'+fmt(q.answered)+' '+plural(q.answered,'answer','answers')+'</span></div>';
      if (q.kind === 'open') {
        h += q.answers.length ? '<ul class="answers">' + q.answers.map(a => '<li><b class="num">no. '+a.seq+'</b>'+esc(a.text)+'</li>').join('') + '</ul>' : '<div class="muted">No answers.</div>';
      } else if (q.kind === 'amount') {
        h += '<div class="stat-row amt muted" style="font-size:12.5px"><span></span><span class="r">Answers</span><span class="r">Sum</span><span class="r">Average</span></div>';
        h += q.options.map(o => '<div class="stat-row amt"><span>'+esc(o.label)+(o.not_numbers ? '<br><span class="not-num">'+o.not_numbers+' '+plural(o.not_numbers,'entry is not a number','entries are not numbers')+', skipped</span>' : '')+'</span><span class="r num">'+fmt(o.count)+'</span><span class="r num">'+fmt(o.sum)+'</span><span class="r num">'+fmt(o.avg)+'</span></div>').join('');
      } else {
        const max = Math.max(1, ...q.options.map(o => o.count));
        h += q.options.map(o => '<div class="stat-row"><span>'+esc(o.label)+'</span><span class="bar-track"><i style="width:'+(100*o.count/max)+'%"></i></span><span class="r num">'+fmt(o.count)+'</span><span class="r num muted">'+fmt(o.pct)+'%</span></div>' +
          (o.texts && o.texts.length ? '<ul class="answers" style="margin:0 0 8px">' + o.texts.map(t => '<li>'+esc(t)+'</li>').join('') + '</ul>' : '')).join('');
      }
      el.innerHTML = h; box.appendChild(el);
    });
  } catch(e){ box.innerHTML = ''; toast(e.message, true); }
}

/* ---------- new survey wizard ---------- */
const DRAFT = 'survey-keyer-creator';
const ICON = {
  up: '<path d="m18 15-6-6-6 6"/>', down: '<path d="m6 9 6 6 6-6"/>',
  copy: '<rect x="9" y="9" width="11" height="11" rx="2"/><path d="M5 15V5a2 2 0 0 1 2-2h10"/>',
  del: '<path d="M4 7h16M10 11v6M14 11v6M6 7l1 13h10l1-13M9 7V4h6v3"/>'
};
const tool = (k, label, disabled, cls) => '<button type="button" class="icon-btn' + (cls ? ' ' + cls : '') + '" data-act="' + k + '" title="' + label + '" aria-label="' + label + '"' + (disabled ? ' disabled' : '') +
  '><svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">' + ICON[k] + '</svg></button>';
let C = null, draftTimer;
const newQ = () => ({text:'', kind:'single', options:[], other:false});
const blankCreator = () => ({title:'', rows:100, charts:true, questions:[newQ()]});
const hasContent = c => !!c && (!!String(c.title || '').trim() || c.questions.some(q => q.text.trim() || q.options.length));
function loadDraft(){
  try { const d = JSON.parse(localStorage.getItem(DRAFT)); if (d && Array.isArray(d.questions) && d.questions.length) return d; } catch(e){}
  return null;
}
function saveDraft(){ clearTimeout(draftTimer); draftTimer = setTimeout(() => { try { localStorage.setItem(DRAFT, JSON.stringify(C)); } catch(e){} }, 300); }
function buildCreator(){
  if (!C) C = loadDraft() || blankCreator();
  $('c-title').value = C.title || ''; $('c-rows').value = C.rows || 100; $('c-charts').checked = C.charts !== false;
  renderQuestions();
}
function updateCreateSummary(){
  const n = C.questions.length, cols = C.questions.reduce((s, q) => s + (q.kind === 'open' ? 1 : Math.max(1, q.options.length)), 0);
  $('c-sum').textContent = n + ' ' + plural(n, 'question', 'questions') + ' · ' + cols + ' ' + plural(cols, 'column', 'columns') + ' in Excel';
}
function renderQuestions(focusIdx){
  const list = $('c-list'), n = C.questions.length; list.innerHTML = '';
  C.questions.forEach((q, i) => {
    const box = document.createElement('section'); box.className = 'q cq';
    const opts = q.kind === 'open' ? '<p class="hint" style="margin:0">The answer is typed in words: one wide column in Excel.</p>' :
      '<label class="field">' + (q.kind === 'amount' ? 'Fields for amounts or numbers, one per line' : 'Answers, one per line') +
      '<textarea class="c-opts" placeholder="' + (q.kind === 'amount' ? 'Amount' : 'Yes&#10;No&#10;Other (please specify)') + '">' + esc(q.options.join('\n')) + '</textarea></label>' +
      (q.kind === 'amount' ? '' : '<label class="chk"><input type="checkbox" class="c-other"' + (q.other ? ' checked' : '') + '> The last answer has a text field (e.g. “Other (please specify)”)</label>');
    box.innerHTML = '<div class="q-title"><span class="q-num num">' + (i + 1) + '</span>' +
      '<input class="c-text" maxlength="1000" placeholder="Question text" aria-label="Question text ' + (i + 1) + '" value="' + esc(q.text) + '">' +
      '<span class="cq-tools">' + tool('up', 'Move up', i === 0) + tool('down', 'Move down', i === n - 1) + tool('copy', 'Duplicate question') + tool('del', 'Delete question', false, 'del') + '</span></div>' +
      '<div class="cq-body"><div class="seg" role="group" aria-label="Question type">' +
      Object.entries(KIND).map(([k, v]) => '<button type="button" data-k="' + k + '" aria-pressed="' + (k === q.kind) + '">' + v + '</button>').join('') +
      '</div>' + opts + '<div class="err-note" hidden></div></div>';
    const ok = () => { box.classList.remove('invalid'); box.querySelector('.err-note').hidden = true; };
    box.querySelector('.c-text').oninput = e => { q.text = e.target.value; ok(); saveDraft(); };
    const ta = box.querySelector('.c-opts');
    if (ta) ta.oninput = () => { q.options = ta.value.split('\n').map(s => s.trim()).filter(Boolean); ok(); updateCreateSummary(); saveDraft(); };
    const other = box.querySelector('.c-other');
    if (other) other.onchange = () => { q.other = other.checked; saveDraft(); };
    box.querySelectorAll('.seg button').forEach(b => b.onclick = () => {
      if (q.kind === b.dataset.k) return;
      q.kind = b.dataset.k; if (q.kind === 'amount' && !q.options.length) q.options = ['Amount'];
      saveDraft(); renderQuestions(); list.children[i].querySelector('.seg button[aria-pressed=true]').focus();
    });
    box.querySelectorAll('.cq-tools button').forEach(b => b.onclick = async () => {
      const act = b.dataset.act, Q = C.questions;
      if (act === 'up' || act === 'down') { const j = act === 'up' ? i - 1 : i + 1; [Q[i], Q[j]] = [Q[j], Q[i]]; saveDraft(); renderQuestions(); list.children[j].querySelector('[data-act="' + act + '"]:not(:disabled), .c-text').focus(); return; }
      if (act === 'copy') { Q.splice(i + 1, 0, JSON.parse(JSON.stringify(q))); saveDraft(); renderQuestions(i + 1); return; }
      if ((q.text.trim() || q.options.length) && !(await ask('The question “' + (q.text.trim() || 'without text') + '” will be removed from the wizard.', {title:'Delete question?', ok:'Delete', danger:true, focusCancel:true}))) return;
      Q.splice(i, 1); if (!Q.length) Q.push(newQ()); saveDraft(); renderQuestions(Math.min(i, Q.length - 1));
    });
    list.appendChild(box);
  });
  updateCreateSummary();
  if (focusIdx !== undefined && list.children[focusIdx]) {
    const b = list.children[focusIdx]; b.scrollIntoView({block:'center', behavior:'smooth'}); b.querySelector('.c-text').focus({preventScroll:true});
  }
}
function validateCreator(){
  let first = -1;
  [...$('c-list').children].forEach((box, i) => {
    const q = C.questions[i];
    const msg = !q.text.trim() ? 'Type the question text.' : q.kind !== 'open' && !q.options.length ?
      (q.kind === 'amount' ? 'Type at least one amount field.' : 'Type at least one answer.') : '';
    box.classList.toggle('invalid', !!msg); const note = box.querySelector('.err-note'); note.textContent = msg; note.hidden = !msg;
    if (msg && first < 0) first = i;
  });
  return first;
}
async function makeFile(){
  const bad = validateCreator();
  if (bad >= 0) { const b = $('c-list').children[bad]; b.scrollIntoView({block:'center', behavior:'smooth'}); b.querySelector('.c-text').focus({preventScroll:true}); toast('Complete the highlighted question.', true); return; }
  const rows = +C.rows;
  if (!(rows >= 1 && rows <= 5000)) { $('c-rows').focus(); toast('The number of surveys (rows) must be between 1 and 5,000.', true); return; }
  if (!(await confirmLeave())) return;
  $('c-make').disabled = true;
  try {
    const s = await api('/api/create', {title: C.title, rows, charts: C.charts !== false, questions: C.questions});
    if (s.cancelled) return;
    try { localStorage.removeItem(DRAFT); } catch(e){}
    C = null; S = null; setSchema(s); toast('Created ' + s.file + '. You can start entering surveys.');
    await afterOpen();
  } catch(e){ toast(e.message, true); }
  finally { $('c-make').disabled = false; }
}
async function copyToCreator(){
  if (hasContent(loadDraft()) && !(await ask('The wizard already holds a survey you typed earlier. Replace it with the questions from this file?', {title:'Replace draft?', ok:'Replace', cancel:'Keep draft', focusCancel:true}))) { show('create'); return; }
  C = {title: S.file.replace(/\.[^.]+$/, '') + ' (new)', rows: Math.min(5000, S.ids.length || 100), charts: true,
       questions: S.questions.map(q => {
         const open = isOpen(q), last = q.options[q.options.length - 1];
         return {text: q.text, kind: open ? 'open' : q.kind, options: open ? [] : q.options.map(o => o.label),
                 other: !open && q.kind !== 'amount' && !!(last && last.text)};
       })};
  saveDraft(); show('create'); toast('Questions copied to the wizard. Change what you need and create the new file.');
}

/* ---------- printable form ---------- */
function printForm(m){
  const KH = {single:'one answer', multi:'tick all that apply'};
  let h = '<div class="pf-head"><h1 class="pf-title">' + esc(m.title) + '</h1><div class="pf-no">Survey no.: <span style="display:inline-block;width:62pt;border-bottom:.8pt solid #555"></span></div></div>' +
    '<p class="pf-hint">Mark your answers with an X. Answer open questions in your own words.</p>';
  m.questions.forEach((q, i) => {
    h += '<div class="pf-q"><div class="pf-qt">' + (i + 1) + '. ' + esc(q.text) + (KH[q.kind] ? '<span class="pf-k">(' + KH[q.kind] + ')</span>' : '') + '</div>';
    if (q.kind === 'open') h += '<div class="pf-ruled"><div></div><div></div><div></div></div>';
    else if (q.kind === 'amount') h += '<div class="pf-amt">' + q.options.map(o => '<span>' + esc(o.label) + '</span><span class="pf-line"></span>').join('') + '</div>';
    else {
      const short = q.options.every(o => o.label.length <= 24 && !o.text), n = q.options.length;
      h += '<ul class="pf-opts' + (short && n >= 6 ? ' cols3' : short && n >= 4 ? ' cols2' : '') + '">' + q.options.map(o =>
        '<li><span class="box' + (q.kind === 'single' ? ' round' : '') + '"></span><span>' + esc(o.label) + '</span>' + (o.text ? '<span class="pf-line"></span>' : '') + '</li>').join('') + '</ul>';
    }
    h += '</div>';
  });
  $('v-print').innerHTML = h + '<div class="pf-foot">Thank you for completing the survey.</div>';
  window.print();
}
const surveyPrintModel = () => ({title: S.file.replace(/\.[^.]+$/, ''), questions: S.questions.map(q => ({
  text: q.text, kind: isOpen(q) ? 'open' : q.kind, options: q.options.map(o => ({label: o.label, text: o.text}))}))});
const creatorPrintModel = () => ({title: C.title.trim() || 'Survey', questions: C.questions.filter(q => q.text.trim()).map(q => ({
  text: q.text, kind: q.kind, options: q.kind === 'open' ? [] : q.options.map((o, k) => ({label: o, text: q.other && k === q.options.length - 1}))}))});

/* ---------- keyboard and events ---------- */
document.addEventListener('keydown', e => {
  if (view !== 'entry' || $('dlg').open) return;
  if (e.ctrlKey && (e.key === 'Enter' || e.key.toLowerCase() === 's')) { e.preventDefault(); save(); return; }
  const el = document.activeElement, inField = el && (el.tagName === 'TEXTAREA' || (el.tagName === 'INPUT' && el.type !== 'radio' && el.type !== 'checkbox'));
  if (inField) {
    if (el.id === 'seq') { if (e.key === 'Enter') { e.preventDefault(); loadSeq(+el.value); } return; }
    if (e.key === 'Escape') { el.blur(); return; }
    if (e.key === 'Enter' && !(el.tagName === 'TEXTAREA' && e.shiftKey)) {
      e.preventDefault();
      const box = el.closest('.q'), inputs = box ? [...box.querySelectorAll('.amount input')] : [];
      const k = inputs.indexOf(el);
      if (k >= 0 && k < inputs.length - 1) inputs[k + 1].focus(); else { el.blur(); setActive(active + 1, true); }
    }
    return;
  }
  if (/^[1-9]$/.test(e.key)) { e.preventDefault(); pick(+e.key - 1); }
  else if (e.key === '0') { e.preventDefault(); clearActive(); }
  else if (e.key === 'Enter' || e.key === 'ArrowDown') { e.preventDefault(); setActive(active + 1, true); }
  else if (e.key === 'ArrowUp') { e.preventDefault(); setActive(active - 1, true); }
});
document.querySelectorAll('.tab').forEach(b => b.onclick = () => {   // the form stays as it was - nothing typed is lost
  const v = b.dataset.view; if (v === view) return;
  if (v === 'entry' && !formReady) { startEntry(); return; }
  show(v);
});
$('toFiles').onclick = async () => { if (await confirmLeave()) { dirty = false; show('start'); } };
$('rescan').onclick = () => { $('found').innerHTML = '<li class="empty">Looking for files…</li>'; loadFiles(true); };
$('browse').onclick = browse;
$('path').addEventListener('keydown', e => { if (e.key === 'Enter') openFile($('path').value); });
$('saveCfg').onclick = saveSettings; $('save').onclick = save; $('newRow').onclick = newRow;
$('clear').onclick = () => { buildForm(); markDirty(); toast('Form cleared. The file is unchanged.'); };
$('backup').onclick = async () => { try { const r = await api('/api/backup', {}); toast('Backup saved: backups\\' + r.backup.split(/[\\/]/).pop()); } catch(e){ toast(e.message, true); } };
$('prev').onclick = () => step(-1); $('next').onclick = () => step(1);
$('nextEmpty').onclick = async () => { if (await confirmLeave()) gotoNextEmpty(+$('seq').dataset.cur || 0, true); };
$('csv').onclick = downloadCsv;
$('printForm').onclick = () => printForm(surveyPrintModel());
$('copyToCreate').onclick = copyToCreator;
$('toCreate').onclick = () => show('create');
$('c-back').onclick = () => formReady ? show('entry') : show('start');
$('c-title').oninput = e => { C.title = e.target.value; saveDraft(); };
$('c-rows').oninput = e => { C.rows = e.target.value; saveDraft(); };
$('c-charts').onchange = e => { C.charts = e.target.checked; saveDraft(); };
$('c-add').onclick = () => { C.questions.push(newQ()); saveDraft(); renderQuestions(C.questions.length - 1); };
$('c-make').onclick = makeFile;
$('c-print').onclick = () => { const m = creatorPrintModel(); if (!m.questions.length) { toast('Type the text of at least one question first.', true); return; } printForm(m); };
$('c-reset').onclick = async () => {
  if (hasContent(C) && !(await ask('All questions typed into the wizard will be removed.', {title:'Clear wizard?', ok:'Clear', danger:true, focusCancel:true}))) return;
  C = blankCreator(); saveDraft(); buildCreator(); $('c-title').focus();
};
let pingTimer;
$('quit').onclick = async () => {
  if (!(await confirmLeave())) return;
  dirty = false; clearTimeout(pingTimer);
  try { await api('/api/quit', {}); } catch(e){}
  document.body.innerHTML = '<div class="view"><h1>App closed</h1><p class="lead">You can close this tab. All saved surveys are in the Excel file.</p></div>';
};
window.addEventListener('beforeunload', e => { if (dirty) { e.preventDefault(); e.returnValue = ''; } });
window.addEventListener('pagehide', () => { try { fetch('/api/bye', {method:'POST', keepalive:true, headers:headers(), body:'{}'}); } catch(e){} });
function schedulePing(ms){ clearTimeout(pingTimer); pingTimer = setTimeout(async () => { await ping(); schedulePing(offline ? 4000 : 15000); }, ms); }
schedulePing(15000);
(async () => {
  try { setSchema(await api('/api/schema')); afterOpen(); } catch(e){ show('start'); }
})();
</script></html>"""


if __name__ == "__main__":
    main()
