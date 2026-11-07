"""Reads what openpyxl cannot, into an ordinary openpyxl workbook, so everything after this works the same for every kind of file. No GUI.

.csv .tsv .txt  plain text, in whichever encoding opens it
.xls            xlrd, with number formats. Many a system's ".xls" is really an HTML table or a renamed .xlsx, so the content decides.
.xlsb .ods      python-calamine for the values. It gives no number formats, so they are read from the file's own style records: 85% stays 85%.
"""

import csv
import datetime as dt
import html.parser
import io
import os
import re
import struct
import xml.etree.ElementTree as ET  # expat fetches no external entities
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
MAX_SPAN = 1000  # a rowspan of 65534 written by mistake would otherwise fill the sheet
# .xlsb records (MS-XLSB 2.3.2): a row's start, the cell records that each begin with column and style, and the styles' formats
XLSB_ROW, XLSB_CELLS, XLSB_FORMAT, XLSB_XF, XLSB_BEGIN_CELL_XFS, XLSB_END_CELL_XFS, XLSB_SHEET = 0, frozenset(range(1, 12)), 44, 47, 617, 618, 156
XLSB_MERGE = 176  # BrtMergeCell: first row, last row, first column, last column, from 0
ODS = {"table": "urn:oasis:names:tc:opendocument:xmlns:table:1.0", "office": "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
       "style": "urn:oasis:names:tc:opendocument:xmlns:style:1.0", "number": "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0"}
ODS_VALUE_TYPES = ("float", "percentage", "currency", "date")
ODS_MAX_COLUMNS = 16384
ODS_MAX_REPEATED_ROWS = 1000  # a style carried down a million empty rows holds nothing to show
MAX_FORMATTED_CELLS = 2_000_000


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
    """Rows of every <table>, each cell as (text, rows it spans, columns it spans)."""

    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.tables = []
        self.row = None
        self.cell = None
        self.span = (1, 1)

    def handle_starttag(self, tag, attrs):
        if tag == "table":
            self.tables.append([])
        elif tag == "tr" and self.tables:
            self.row = []
        elif tag in ("td", "th") and self.row is not None:
            self.cell = []
            given = dict(attrs)
            self.span = tuple(min(MAX_SPAN, max(1, _whole(given.get(name)))) for name in ("rowspan", "colspan"))

    def handle_endtag(self, tag):
        if tag in ("td", "th") and self.cell is not None:
            self.row.append((" ".join("".join(self.cell).split()), *self.span))
            self.cell = None
        elif tag == "tr" and self.row is not None:
            if self.tables and self.row:
                self.tables[-1].append(self.row)
            self.row = None

    def handle_data(self, data):
        if self.cell is not None:
            self.cell.append(data)


def _whole(text):
    try:
        return int(str(text).strip())
    except ValueError:
        return 1


def _laid_out(table):
    """(rows of texts, merged ranges) for an HTML table, as Excel lays it out: a cell spanning rows or columns is merged, its text in the top left.
    A rowspan from the row above takes its place in the rows below, so their cells move right and stay under their headings."""
    grid, merges, taken, reach = [], [], set(), {}
    for r, cells in enumerate(table):
        row, c = [], 0
        for text, down, across in cells:
            while (r, c) in taken:
                c += 1
            for dr in range(down):
                reach[r + dr] = max(reach.get(r + dr, 0), c + across)  # how far each row is covered, kept as it goes: scanning taken for every row is slow on a long table
                for dc in range(across):
                    taken.add((r + dr, c + dc))
            row.append((c, text))
            if down > 1 or across > 1:
                merges.append((r + 1, c + 1, r + down, c + across))
            c += across
        grid.append(row)
    rows = []
    for r, row in enumerate(grid):
        width = max([c + 1 for c, _text in row] + [reach.get(r, 0)])
        line = [""] * width
        for c, text in row:
            line[c] = text
        rows.append(line)
    spanned = {row for top, _l, bottom, _r in merges for row in range(top, bottom + 1)}
    keep = [r for r, line in enumerate(rows) if any(line) or r + 1 in spanned]
    if len(keep) != len(rows):  # rows with nothing in them and no span through them: drop them, and move spans up with the rows
        moved = {old + 1: new + 1 for new, old in enumerate(keep)}
        rows = [rows[r] for r in keep]
        merges = [(moved[top], left, moved.get(bottom, moved[top]), right) for top, left, bottom, right in merges if top in moved]
    return rows, merges


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
        rows, merges = _laid_out(table)
        _put_rows(ws, rows)
        for top, left, bottom, right in merges:
            ws.merge_cells(start_row=top, start_column=left, end_row=bottom, end_column=right)
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


def _records(data):
    """(type, body) of each BIFF12 record in an .xlsb part: type and size are written 7 bits a byte, the high bit set while more follow."""
    at, end = 0, len(data)
    while at < end:
        kind = data[at] & 0x7F
        if data[at] & 0x80:
            kind |= (data[at + 1] & 0x7F) << 7
            at += 2
        else:
            at += 1
        size, shift = 0, 0
        for _ in range(4):
            byte = data[at]
            at += 1
            size |= (byte & 0x7F) << shift
            shift += 7
            if not byte & 0x80:
                break
        yield kind, data[at:at + size]
        at += size


def _wide(body, at):
    """An XLWideString starting at at: its text, and where it ends."""
    count = struct.unpack_from("<I", body, at)[0]
    start = at + 4
    return body[start:start + count * 2].decode("utf-16-le", "replace"), start + count * 2


def _xlsb_formats(path):
    """({sheet name: {(row, column): number format}}, {sheet name: [(top, left, bottom, right)]}) with 1-based places: the cells of an .xlsb whose format is not General, and its merged cells (MS-XLSB records)."""
    from openpyxl.styles.numbers import BUILTIN_FORMATS
    with zipfile.ZipFile(path) as package:
        names = set(package.namelist())
        codes, styles, inside = {}, [], False
        for kind, body in _records(package.read("xl/styles.bin")):
            if kind == XLSB_FORMAT:
                codes[struct.unpack_from("<H", body, 0)[0]] = _wide(body, 2)[0]
            elif kind in (XLSB_BEGIN_CELL_XFS, XLSB_END_CELL_XFS):
                inside = kind == XLSB_BEGIN_CELL_XFS
            elif kind == XLSB_XF and inside:
                number = struct.unpack_from("<H", body, 2)[0]
                styles.append(codes.get(number) or BUILTIN_FORMATS.get(number, "General"))
        rels = package.read("xl/_rels/workbook.bin.rels").decode("utf-8", "replace")
        targets = {}
        for tag in re.findall(r"<Relationship\b[^>]*>", rels):
            ident, target = re.search(r'\bId="([^"]+)"', tag), re.search(r'\bTarget="([^"]+)"', tag)
            if ident and target:
                targets[ident.group(1)] = target.group(1)
        found, merged = {}, {}
        for kind, body in _records(package.read("xl/workbook.bin")):
            if kind != XLSB_SHEET:
                continue
            relation, at = ("", 12) if struct.unpack_from("<I", body, 8)[0] == 0xFFFFFFFF else _wide(body, 8)
            name = _wide(body, at)[0]
            target = targets.get(relation, "").lstrip("/")
            part = target if target.startswith("xl/") else "xl/" + target
            if part not in names:
                continue
            cells, row, merges = {}, 0, []
            for record, data in _records(package.read(part)):
                if record == XLSB_MERGE and len(data) >= 16:
                    top, bottom, left, right = struct.unpack_from("<IIII", data, 0)
                    merges.append((top + 1, left + 1, bottom + 1, right + 1))
                elif record == XLSB_ROW and len(data) >= 4:
                    row = struct.unpack_from("<I", data, 0)[0]
                elif record in XLSB_CELLS and len(data) >= 8:
                    column, style = struct.unpack_from("<II", data, 0)
                    style &= 0xFFFFFF
                    code = styles[style] if style < len(styles) else "General"
                    if code != "General":
                        cells[(row + 1, column + 1)] = code
                        if len(cells) > MAX_FORMATTED_CELLS:
                            break
            found[name], merged[name] = cells, merges
        return found, merged


def _ods_code(style):
    """An Excel number format for an OpenDocument data style: 0.0% for a percentage with one decimal, #,##0 for grouped digits, yyyy-mm for a year and month. None where General serves."""
    kind = style.tag.rpartition("}")[2]
    if kind == "date-style":
        parts = []
        for part in style:
            name, long = part.tag.rpartition("}")[2], part.get(f"{{{ODS['number']}}}style") == "long"
            if name == "year":
                parts.append("yyyy" if long else "yy")
            elif name == "month":
                parts.append("mmm" if part.get(f"{{{ODS['number']}}}textual") == "true" else "mm" if long else "m")
            elif name == "day":
                parts.append("dd" if long else "d")
            elif name == "text":
                parts.append("".join(ch if ch in "-/. " else f"\\{ch}" for ch in (part.text or "")))
        return "".join(parts) or None
    number = style.find("number:number", ODS)
    if number is None or kind not in ("number-style", "percentage-style", "currency-style"):
        return None
    places = int(number.get(f"{{{ODS['number']}}}decimal-places") or 0)
    grouped = number.get(f"{{{ODS['number']}}}grouping") == "true" or kind == "currency-style"
    if kind == "number-style" and not grouped and number.get(f"{{{ODS['number']}}}decimal-places") is None:
        return None  # N0: what General shows
    code = ("#,##0" if grouped else "0") + ("." + "0" * places if places else "")
    return code + "%" if kind == "percentage-style" else code


def _ods_formats(path):
    """({sheet name: {(row, column): number format}}, {sheet name: [(top, left, bottom, right)]}) with 1-based places: the value cells of an .ods whose data style is not General, and its merged cells."""
    with zipfile.ZipFile(path) as package:
        roots = [ET.fromstring(package.read(part)) for part in ("styles.xml", "content.xml") if part in package.namelist()]
    codes, data_style, parent = {}, {}, {}
    for root in roots:
        for style in root.iter():
            name = style.get(f"{{{ODS['style']}}}name")
            if name is None:
                continue
            if style.tag.startswith(f"{{{ODS['number']}}}"):
                code = _ods_code(style)
                if code:
                    codes[name] = code
            elif style.tag == f"{{{ODS['style']}}}style" and style.get(f"{{{ODS['style']}}}family") == "table-cell":
                data_style[name] = style.get(f"{{{ODS['style']}}}data-style-name")
                parent[name] = style.get(f"{{{ODS['style']}}}parent-style-name")

    def code_of(cell_style):
        for _ in range(5):  # a style without a data style takes its parent's
            if cell_style is None:
                return None
            if data_style.get(cell_style):
                return codes.get(data_style[cell_style])
            cell_style = parent.get(cell_style)
        return None

    found, merged = {}, {}
    content = roots[-1] if roots else None
    for table in (content.iter(f"{{{ODS['table']}}}table") if content is not None else ()):
        defaults = []
        for column in table.iter(f"{{{ODS['table']}}}table-column"):
            defaults += [column.get(f"{{{ODS['table']}}}default-cell-style-name")] * min(int(column.get(f"{{{ODS['table']}}}number-columns-repeated") or 1), ODS_MAX_COLUMNS)
        cells, row, merges = {}, 0, []
        for line in table.iter(f"{{{ODS['table']}}}table-row"):
            repeat = int(line.get(f"{{{ODS['table']}}}number-rows-repeated") or 1)
            column, values = 0, []
            for cell in line:
                if cell.tag not in (f"{{{ODS['table']}}}table-cell", f"{{{ODS['table']}}}covered-table-cell"):
                    continue
                across = int(cell.get(f"{{{ODS['table']}}}number-columns-repeated") or 1)
                wide, tall = (min(int(cell.get(f"{{{ODS['table']}}}number-{kind}-spanned") or 1), MAX_SPAN) for kind in ("columns", "rows"))
                if wide > 1 or tall > 1:  # a merged cell: the cells it covers follow as covered-table-cell
                    merges.append((row + 1, column + 1, row + tall, column + wide))
                if cell.get(f"{{{ODS['office']}}}value-type") in ODS_VALUE_TYPES:
                    style = cell.get(f"{{{ODS['table']}}}style-name") or (defaults[column] if column < len(defaults) else None)
                    code = code_of(style)
                    if code:
                        values += [(column + k, code) for k in range(min(across, ODS_MAX_COLUMNS))]
                column += across
            if values:
                for k in range(min(repeat, ODS_MAX_REPEATED_ROWS)):
                    cells.update({(row + k + 1, c + 1): code for c, code in values})
            row += repeat
            if len(cells) > MAX_FORMATTED_CELLS:
                break
        found[table.get(f"{{{ODS['table']}}}name")] = cells
        merged[table.get(f"{{{ODS['table']}}}name")] = merges
    return found, merged


def _formats(path):
    """(number formats, merged cells) a reader of values leaves out, or nothing when the file's own records cannot be read: the values still are."""
    try:
        return _xlsb_formats(path) if path.lower().endswith(".xlsb") else _ods_formats(path)
    except (KeyError, ValueError, struct.error, ET.ParseError, zipfile.BadZipFile, IndexError):
        return {}, {}


def _from_calamine(path):
    from python_calamine import CalamineWorkbook
    book = CalamineWorkbook.from_path(path)
    formats, merged = _formats(path)
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for name in book.sheet_names:
        ws = wb.create_sheet(_title(name, [s.title for s in wb.worksheets]))
        sheet = book.get_sheet_by_name(name)
        shown = formats.get(name, {})
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
                elif (r, c) in shown and not isinstance(value, bool):
                    cell.number_format = shown[(r, c)]  # 85% as the file shows it, not 0.85
                elif isinstance(value, (dt.datetime, dt.date)):
                    cell.number_format = "yyyy-mm-dd"
        ranges = [(top + 1, left + 1, bottom + 1, right + 1) for (top, left), (bottom, right) in getattr(sheet, "merged_cell_ranges", None) or []]
        for top, left, bottom, right in ranges or merged.get(name, []):  # the reader gives merges for .xls only; .xlsb and .ods have them from their own records
            if bottom >= top and right >= left and (bottom, right) != (top, left):
                ws.merge_cells(start_row=top, start_column=left, end_row=bottom, end_column=right)
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
