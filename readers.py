"""Reads what openpyxl cannot, into an ordinary openpyxl workbook, so everything after this works the same for every kind of file. No GUI.

.csv .tsv .txt  plain text, in whichever encoding opens it
.xls            xlrd, with number formats. Many a system's ".xls" is really an HTML table or a renamed .xlsx, so the content decides.
.xlsb .ods      python-calamine. It gives values without number formats, so 85% reads as 0.85.
"""

import csv
import datetime as dt
import html.parser
import io
import os
import re
import zipfile

import openpyxl

OPENPYXL_TYPES = (".xlsx", ".xlsm", ".xltx", ".xltm")
TEXT_TYPES = (".csv", ".tsv", ".txt")
OTHER_TYPES = (".xls", ".xlsb", ".ods") + TEXT_TYPES
ALL_TYPES = OPENPYXL_TYPES + OTHER_TYPES
DOCUMENT_TYPES = (".docx",)  # read by documents.py: words to line up with figures, not a table of figures
NOT_CHECKED_AS_ZIP = (".xls",) + TEXT_TYPES  # the rest are zip files, and a file that is not one is probably wrapped by document security
ENCODINGS = ("utf-8-sig", "cp949", "utf-16")
PLAIN_NUMBER = re.compile(r"-?(?:0|[1-9]\d*)(?:\.\d+)?")
GROUPED_NUMBER = re.compile(r"-?[1-9]\d{0,2}(?:,\d{3})+(?:\.\d+)?")
BAD_IN_TITLE = re.compile(r"[\\/*?:\[\]]")
MAX_TITLE = 31


class Protected(Exception):
    """The file needs a password."""


def _title(name, taken=()):
    title = BAD_IN_TITLE.sub(" ", name).strip()[:MAX_TITLE] or "Sheet"
    number = 2
    while title in taken:
        title = f"{title[:MAX_TITLE - 3]}_{number}"
        number += 1
    return title


def _number(value):
    return int(value) if isinstance(value, float) and value.is_integer() else value


def _text_cell(text):
    """A number if it is written as one. 1,234 is too, but 007 and 010-1234 are codes and stay as written."""
    text = text.strip()
    if PLAIN_NUMBER.fullmatch(text):
        return (float(text) if "." in text else int(text)), None
    if GROUPED_NUMBER.fullmatch(text):
        return (float(text.replace(",", "")) if "." in text else int(text.replace(",", ""))), "#,##0.##" if "." in text else "#,##0"
    return text or None, None


def _put_rows(ws, rows):
    for r, row in enumerate(rows, start=1):
        for c, raw in enumerate(row, start=1):
            value, fmt = _text_cell(raw) if isinstance(raw, str) else (raw, None)
            if value is not None:
                cell = ws.cell(r, c, value)
                if isinstance(value, str):
                    cell.data_type = "s"  # never a formula, whatever it starts with
                if fmt:
                    cell.number_format = fmt


def _decode(raw):
    if raw[:2] in (b"\xff\xfe", b"\xfe\xff"):
        return raw.decode("utf-16")
    for encoding in ENCODINGS[:2]:
        try:
            return raw.decode(encoding)
        except UnicodeDecodeError:
            continue
    return raw.decode("utf-16")


def _delimiter(text, ext):
    head = text[:4000].splitlines()[:5]
    counts = {mark: sum(line.count(mark) for line in head) for mark in ("\t", ",", ";", "|")}
    if ext == ".tsv" and counts["\t"]:
        return "\t"
    return max(counts, key=lambda mark: (counts[mark], mark == ","))


def _from_text(path, raw, ext):
    text = _decode(raw)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = _title(os.path.splitext(os.path.basename(path))[0])
    _put_rows(ws, csv.reader(io.StringIO(text), delimiter=_delimiter(text, ext)))
    return wb


class _Tables(html.parser.HTMLParser):
    """Rows of every <table>. Cell spans are spread out, so columns stay lined up."""

    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.tables = []
        self.row = None
        self.cell = None
        self.span = 1

    def handle_starttag(self, tag, attrs):
        if tag == "table":
            self.tables.append([])
        elif tag == "tr" and self.tables:
            self.row = []
        elif tag in ("td", "th") and self.row is not None:
            self.cell = []
            try:
                self.span = max(1, int(dict(attrs).get("colspan") or 1))
            except ValueError:
                self.span = 1

    def handle_endtag(self, tag):
        if tag in ("td", "th") and self.cell is not None:
            self.row.append(" ".join("".join(self.cell).split()))
            self.row.extend("" for _ in range(self.span - 1))
            self.cell = None
        elif tag == "tr" and self.row is not None:
            if self.tables and any(self.row):
                self.tables[-1].append(self.row)
            self.row = None

    def handle_data(self, data):
        if self.cell is not None:
            self.cell.append(data)


def _from_html(path, raw):
    parser = _Tables()
    parser.feed(_decode(raw))
    tables = [table for table in parser.tables if table]
    if not tables:
        raise ValueError("no table in this file")
    wb = openpyxl.Workbook()
    stem = os.path.splitext(os.path.basename(path))[0]
    for number, table in enumerate(tables):
        ws = wb.active if number == 0 else wb.create_sheet()
        ws.title = _title(stem if number == 0 else f"{stem} {number + 1}", [s.title for s in wb.worksheets if s is not ws])
        _put_rows(ws, table)
    return wb


def _from_xls(path):
    import xlrd
    try:
        book = xlrd.open_workbook(path, formatting_info=True)
    except xlrd.biffh.XLRDError as exc:
        if "encrypted" in str(exc).lower():
            raise Protected() from exc
        raise
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for sheet in book.sheets():
        ws = wb.create_sheet(_title(sheet.name, [s.title for s in wb.worksheets]))
        if sheet.visibility:
            ws.sheet_state = "hidden"
        for r in range(sheet.nrows):
            for c in range(sheet.ncols):
                kind, value = sheet.cell_type(r, c), sheet.cell_value(r, c)
                if kind in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK):
                    continue
                fmt = book.format_map[book.xf_list[sheet.cell_xf_index(r, c)].format_key].format_str
                if kind == xlrd.XL_CELL_DATE:
                    try:
                        value = xlrd.xldate.xldate_as_datetime(value, book.datemode)
                    except (ValueError, OverflowError):
                        pass
                elif kind == xlrd.XL_CELL_NUMBER:
                    value = _number(value)
                elif kind == xlrd.XL_CELL_BOOLEAN:
                    value = bool(value)
                elif kind == xlrd.XL_CELL_ERROR:
                    value = xlrd.error_text_from_code.get(value, "#ERROR")
                cell = ws.cell(r + 1, c + 1, value)
                if isinstance(value, str):
                    cell.data_type = "s"
                if fmt and fmt != "General":
                    cell.number_format = fmt
        for top, bottom, left, right in sheet.merged_cells:
            if bottom - top > 1 or right - left > 1:
                ws.merge_cells(start_row=top + 1, start_column=left + 1, end_row=bottom, end_column=right)
    return wb


def _from_calamine(path):
    from python_calamine import CalamineWorkbook
    book = CalamineWorkbook.from_path(path)
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for name in book.sheet_names:
        ws = wb.create_sheet(_title(name, [s.title for s in wb.worksheets]))
        sheet = book.get_sheet_by_name(name)
        for r, row in enumerate(sheet.to_python(skip_empty_area=False), start=1):
            for c, value in enumerate(row, start=1):
                if value in ("", None):
                    continue
                value = _number(value)
                if isinstance(value, dt.time):
                    value = value.isoformat()
                cell = ws.cell(r, c, value)
                if isinstance(value, str):
                    cell.data_type = "s"
                elif isinstance(value, (dt.datetime, dt.date)):
                    cell.number_format = "yyyy-mm-dd"
        for (top, left), (bottom, right) in getattr(sheet, "merged_cell_ranges", None) or []:
            ws.merge_cells(start_row=top + 1, start_column=left + 1, end_row=bottom + 1, end_column=right + 1)
    return wb


def load(path):
    """An openpyxl workbook of a file that is not one. Raises Protected, or whatever the reader raises on a damaged file."""
    ext = os.path.splitext(path)[1].lower()
    with open(path, "rb") as fh:
        raw = fh.read()
    if ext in TEXT_TYPES:
        return _from_text(path, raw, ext)
    if ext == ".xls":
        if raw[:8] == bytes.fromhex("d0cf11e0a1b11ae1"):
            return _from_xls(path)
        if zipfile.is_zipfile(path):  # a renamed .xlsx
            return openpyxl.load_workbook(io.BytesIO(raw), data_only=True)
        if b"<" in raw[:2000] and re.search(rb"<\s*(?:table|html)", raw[:20000], re.I):
            return _from_html(path, raw)
        return _from_text(path, raw, ".txt")  # a tab-separated export wearing the name .xls
    return _from_calamine(path)
