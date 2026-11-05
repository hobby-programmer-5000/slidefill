"""SlideFill logic: fills {{placeholders}}, tables and charts of a PowerPoint template from an Excel workbook.

No GUI imports. run() is the single code path for both preview (output=None) and generation.
A value that cannot be determined never becomes a blank: its placeholder stays visible in the output
and the problem is reported as an issue.
"""

import copy
import datetime as dt
import fnmatch
import json
import math
import os
import re
import statistics
import unicodedata
import warnings
import zipfile
from collections import Counter
from dataclasses import dataclass, field, replace

import openpyxl
from lxml import etree
from openpyxl.utils import range_boundaries
from openpyxl.utils.cell import column_index_from_string, get_column_letter
from pptx import Presentation
from pptx.chart.data import CategoryChartData, XyChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.opc.package import XmlPart

import documents
import facts
import readers


class _NoPivotCaches(dict):
    def __missing__(self, key):
        return None


# A pivot table's cache holds every source record, read again for each sheet: one public workbook passed 6 GB here. What the table shows is stored in the sheet's cells,
# and nothing here ever saves a workbook it has read, so the caches are left unread.
openpyxl.reader.workbook.WorkbookParser.pivot_caches = property(lambda self: _NoPivotCaches())
from numfmt import format_value

warnings.filterwarnings("ignore", module="openpyxl")

APP_NAME = "SlideFill"
SETTINGS_FILE = "SlideFill.json"
INSIGHTS_FILE = "SlideFill.insights.json"
OLE_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"  # legacy .xls/.ppt, and also encrypted or many DRM-wrapped files

A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
P = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
A_P, A_R, A_T, A_BR, A_RPR = A + "p", A + "r", A + "t", A + "br", A + "rPr"

PLACEHOLDER = re.compile(r"\{\{\s*(.+?)\s*\}\}", re.S)
MARKER = re.compile(r"\{\{\s*(table|chart|표|차트)\s*:\s*(.+?)\s*\}\}", re.I | re.S)
COORD = re.compile(r"[A-Za-z]{1,3}[0-9]+(:[A-Za-z]{1,3}[0-9]+)?")
LATEST = re.compile(r"\s*(?:최신|latest)\s*(\d*)\s*:(.*)", re.I | re.S)
LATEST_COLUMNS = re.compile(r"[A-Za-z]{1,3}(\s*:\s*[A-Za-z]{1,3})?")
DATE_TEXT = re.compile(r"(\d{4})\s*[-/.년]\s*(\d{1,2})\s*(?:[-/.월]\s*(\d{1,2})\s*[.일]?)?\s*[월.]?\s*(?:\(\w+\))?"
                       r"\s*(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?")
DATE_DIGITS = re.compile(r"(\d{4})(\d{2})(\d{2})")
YEAR_TEXT = re.compile(r"((?:19|20)\d\d)\s*(?:년도?)?")
DATE_MARK = re.compile(r"\s*(?:\((?:p|r|e|est)\)|\*+)\s*$", re.I)  # Jul. 2026(p): preliminary and revised marks
MONTH_WORDS = {"jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "sept", "oct", "nov", "dec", "january", "february", "march", "april", "june",
               "july", "august", "september", "october", "november", "december"}
MONTH_NUMBERS = {name: i + 1 for i, name in enumerate("jan feb mar apr may jun jul aug sep oct nov dec".split())}
YEAR_MONTH_NAME = re.compile(r"(\d{4})[\s-]+([A-Za-z]{3,9})\.?")
MONTH_NAME_YEAR = re.compile(r"([A-Za-z]{3,9})\.?[\s,-]*(\d{4})")
MONTH_NAME_SHORT = re.compile(r"([A-Za-z]{3,9})\.?\s*[-']\s*(\d{2})")  # Aug-26. Aug 12 alone could be a day, so it is not read
KOREAN_SHORT_MONTH = re.compile(r"(\d{2})\s*년\s*(\d{1,2})\s*월")
QUARTER_YEAR_FIRST = re.compile(r"(\d{4})\s*(?:년\s*)?[-.]?\s*Q?([1-4])(?:\s*분기|(?<=Q[1-4]))", re.I)
QUARTER_Q_FIRST = re.compile(r"Q([1-4])\s*[-.']?\s*(\d{4}|\d{2})", re.I)
QUARTER_Q_AFTER = re.compile(r"([1-4])Q\s*[-.']?\s*(\d{4}|\d{2})", re.I)
DAY_MONTH_YEAR = re.compile(r"(\d{1,2})\s+([A-Za-z]{3,9})\.?,?\s+(\d{4})")
MONTH_DAY_YEAR = re.compile(r"([A-Za-z]{3,9})\.?\s+(\d{1,2}),\s*(\d{4})")
MONTH_QUARTER = re.compile(r"([A-Za-z]{3,9})\s+quarter\s+(\d{4})", re.I)
YEAR_SPAN = re.compile(r"((?:19|20)\d\d)\s*[-/–—]\s*(\d{2})")
FISCAL_YEAR = re.compile(r"(?:fy\s*)?((?:19|20)\d\d)\s*[–—]\s*(\d{2})", re.I)  # 2007–08, named by the year it ends. 2007-08 with a plain hyphen reads as a month
MONTH_ALONE = re.compile(r"0?(\d{1,2})\s*월")
NUMBER_TEXT = re.compile(r"\s*([+-]?\d[\d,]*(?:\.\d+)?(?:[eE][+-]?\d+)?)\s*(?:\([^()]*\)|\[[^\[\]]*\]|명|건|원|개|회|톤|대|호)?\s*")
AGGREGATE = re.compile(r"\s*(합계|sum|건수|count|평균|avg|average|mean|비율|rate)\s*\((.*)\)\s*", re.I | re.S)
AGGREGATE_KINDS = {"합계": "sum", "sum": "sum", "건수": "count", "count": "count", "평균": "mean", "avg": "mean", "average": "mean", "mean": "mean", "비율": "rate", "rate": "rate"}
YEAR_HEADING = re.compile(r"(?:year|연도|fy)?\s*((?:19|20)\d\d)\s*(?:년도?)?", re.I)
UNIT_HEADING = re.compile(r"\(.*\)|\$m|\$|단위.*|%", re.S)
RATE_EXPLICIT = re.compile(r"percentage change|% ?change|증감률|전년비|전년동월비", re.I)
RATE_TITLE = re.compile(r"(?:률|율|점수)\s*$")
SCHOOL_YEAR = re.compile(r"학년도|회계연도|fiscal|\bfy\b", re.I)
BOUNDED_LINE = re.compile(r"([A-Z]{1,3})(\d+):([A-Z]{1,3})(\d+)")
RANK = re.compile(r"\s*(순위값|순위|구성비|topvalue|top|share)\s*(\d*)\s*:(.*)", re.I | re.S)
RANK_NAMES = {"순위": "name", "top": "name", "순위값": "value", "topvalue": "value", "구성비": "share", "share": "share"}
TOTAL_WORDS = {"합계", "총계", "총합", "총합계", "계", "전체", "total", "grandtotal", "sum", "합"}
SUBTOTAL_WORDS = {"소계", "subtotal"}
TOTAL_ROW = re.compile(r"totals?(?![a-z])(?!.*\b(?:pty|ltd|inc|limited|llc|corp|plc|gmbh)\b)", re.I)  # Total revenue, not Total Tools Importing Pty Ltd or TotalEnergies
TOTAL_TOLERANCE = 0.005  # a stated total this close to the items' sum is taken to be their total, allowing for rounding
LATE_MINUTES = 60  # a newest day's last entry this much earlier than usual looks unfinished
SHARE_HEADER = "%"
AMOUNT_HEADING = re.compile(r"[증중]감(?![률율])")  # 증감 alone is an amount that went up or down, 증감률 is a rate
RATE_HEADING = re.compile(r"율|률|증감|전년비|전년동월비|지수.*증|점수|만족도|\bscores?\b|\bpercentage|% ?change|\bgrowth|\brates?\b|\bratios?\b", re.I)
HANGUL = re.compile(r"[가-힣]")
ID_CODE = re.compile(r"[A-Z0-9]{2,8}")  # J468: a dataset's code, not a name
MAX_LINES_FOR_A_REPEAT = 200  # how far along a row of headings to look for the same one again
HEADING_LOOK_AHEAD = 30
PIVOT_HEADING = re.compile(r"column labels|row labels|values|열 레이블|행 레이블|값", re.I)
GRAND_TOTAL = re.compile(r"(?:grand\s+)?total|총합계|합계|총계|전체", re.I)
UNIT_NOTE = re.compile(r"단위\s*[:：]|[(（\[]\s*단위|\bunits?\s*[:：]|[(（\[]\s*units?\b", re.I)  # not 유통단계별단위명, which is only a column's name
EARLIER_NAME = re.compile(r"전년|전월|전기|전분기|작년|지난|이전|직전|prior|previous|last\s+(?:year|month|quarter)", re.I)
UNIT_NOTE_LENGTH = 40
UNIT_WORD_LENGTH = 20
UNIT_SCAN_ROWS = 40  # a note may sit under a short table
UNIT_SCAN_COLUMNS = 12
PER_CENT_UNIT = re.compile(r"\s*(?:per\s?cent|percent(?:age)?|퍼센트|백분율|%)\s*", re.I)
AS_AT = re.compile(r"\bas\s+(?:at|of)\b|기준일|자료\s*기준|기준\s*[:：]", re.I)  # Data as at 30-Sep-2025
AS_AT_ROWS = 60
AS_AT_COLUMNS = 3
AS_AT_LENGTH = 200
DAY_FIRST = re.compile(r"\b(\d{1,2})[-\s]([A-Za-z]{3,9})[-\s]((?:19|20)\d\d)\b")
YEAR_FIRST = re.compile(r"((?:19|20)\d\d)\s*[-./년]\s*(\d{1,2})\s*[-./월]\s*(\d{1,2})")
MONTH_NAMES = ("jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec")
PERIOD_WORD = {"ko": "기간", "en": "Period"}
ITEM_WORD = {"ko": "항목", "en": "Item"}
TRAILING_TOTAL = re.compile(r"\s+(?:total|합계|총계)\s*$", re.I)  # the row under a date heading is the sheet's total, not part of what the dates are called
FISCAL_YEAR_ENDS = (6, 30)  # the file does not say. March, June and September are all in use; before the second calendar year's middle, a 2025-26 figure cannot be called complete
COUNT_WORD_KO = "건수"
MIN_HEADING_LENGTH = 2
OUTER_LEVEL_SHARE = 0.5  # labels merged in from the left on most rows are simply how this table is written
CALC = re.compile(r"\s*(전일대비|직전대비|전주대비|평균대비|최고일|최저일|해설|읽는법|prev|week|avg|peak|low|note|reading)\s*(\d*)\s*:(.*)", re.I | re.S)
INSIGHT = re.compile(r"\s*(?:통찰|insight)\s*:(.*)", re.I | re.S)
CALC_NAMES = {"전일대비": "prev", "직전대비": "prev", "전주대비": "week", "평균대비": "avg", "최고일": "peak", "최저일": "low", "해설": "해설", "읽는법": "읽는법"}
ILLEGAL_XML = re.compile("[\x00-\x08\x0b\x0c\x0e-\x1f￾￿]")
SMART_QUOTES = str.maketrans({"‘": "'", "’": "'", "“": '"', "”": '"'})

XY_CHARTS = {XL_CHART_TYPE.XY_SCATTER, XL_CHART_TYPE.XY_SCATTER_LINES, XL_CHART_TYPE.XY_SCATTER_LINES_NO_MARKERS,
             XL_CHART_TYPE.XY_SCATTER_SMOOTH, XL_CHART_TYPE.XY_SCATTER_SMOOTH_NO_MARKERS}
BUBBLE_CHARTS = {XL_CHART_TYPE.BUBBLE, XL_CHART_TYPE.BUBBLE_THREE_D_EFFECT}

# Built-in format ids that openpyxl does not know, as real Excel on Korean Windows displays them.
# Without these a date reads as a bare serial number. \/ is a literal slash, unlike / which follows the PC's locale.
KOREAN_BUILTIN_FORMATS = {
    27: 'yyyy"年" mm"月" dd"日"', 28: "mm-dd", 29: "mm-dd", 30: "mm-dd-yy", 31: 'yyyy"년" mm"월" dd"일"',
    32: 'h"시" mm"분"', 33: 'h"시" mm"분" ss"초"', 34: r"yyyy\/mm\/dd", 35: "yyyy-mm-dd", 36: 'yyyy"年" mm"月" dd"日"',
    50: 'yyyy"年" mm"月" dd"日"', 51: "mm-dd", 52: r"yyyy\/mm\/dd", 53: "yyyy-mm-dd", 54: "mm-dd",
    55: r"yyyy\/mm\/dd", 56: "yyyy-mm-dd", 57: 'yyyy"年" mm"月" dd"日"', 58: "mm-dd",
}

# Errors leave the placeholder, table or chart untouched. Warnings fill it and tell the user why it may look off.
ERRORS = {"missing_name", "missing_sheet", "bad_ref", "range_for_text", "name_not_cells", "multi_area",
          "formula_no_cache", "error_value", "marker_in_text", "marker_wrong_shape", "column_mismatch",
          "merged_resize", "chart_too_small", "chart_unsupported", "latest_bad", "latest_no_dates", "latest_ambiguous",
          "latest_duplicates", "latest_sheets_overlap", "latest_sheets_differ", "cross_no_items", "cross_not_additive", "calc_columns", "calc_no_base"}


class SlideFillError(Exception):
    """A file cannot be used at all. code is an i18n key; params fill its message."""

    def __init__(self, code, **params):
        super().__init__(code)
        self.code, self.params = code, params


class IssueError(Exception):
    def __init__(self, code, **params):
        super().__init__(code)
        self.issue = Issue(code, params)


@dataclass(frozen=True)
class Issue:
    code: str
    params: dict = field(default_factory=dict, hash=False)

    @property
    def is_error(self):
        return self.code in ERRORS


@dataclass
class Item:
    """One thing the template asks for, with every place it occurs."""
    kind: str  # text, table, chart, unsupported
    key: str
    value: str = ""
    source: str = ""
    wheres: list = field(default_factory=list)
    issues: list = field(default_factory=list)

    @property
    def status(self):
        if any(i.is_error for i in self.issues):
            return "error"
        return "warn" if self.issues else "ok"


@dataclass(frozen=True)
class InsightRequest:
    """What the second stage needs: the calculated results, and the sentence to use if nothing is chosen."""
    key: str
    lang: str
    facts: tuple
    fingerprint: str
    fallback: str
    source: str = ""


@dataclass
class Report:
    items: list = field(default_factory=list)
    output: str = ""
    requests: dict = field(default_factory=dict)

    def count(self, status):
        return sum(1 for item in self.items if item.status == status)


# ---------------------------------------------------------------- reading files

def _check_file(path, extensions, kind):
    ext = os.path.splitext(path)[1].lower()
    try:
        with open(path, "rb") as fh:
            head = fh.read(8)
    except FileNotFoundError:
        raise SlideFillError("file_missing", path=path)
    except PermissionError:
        raise SlideFillError("file_locked", path=path)
    except OSError as exc:
        raise SlideFillError("file_unreadable", path=path, detail=str(exc))
    if ext not in extensions:
        raise SlideFillError(f"wrong_type_{kind}", path=path)
    if kind == "excel" and ext in readers.NOT_CHECKED_AS_ZIP:
        return
    if head[:8] == OLE_MAGIC:
        raise SlideFillError("file_protected", path=path)
    if not zipfile.is_zipfile(path):
        raise SlideFillError("file_not_office", path=path)


def as_date(raw):
    """A datetime for a real Excel date or a date written as text or digits, otherwise None."""
    if isinstance(raw, dt.datetime):
        return raw
    if isinstance(raw, dt.date):
        return dt.datetime.combine(raw, dt.time())
    if isinstance(raw, (int, float)) and not isinstance(raw, bool) and float(raw).is_integer():
        raw = str(int(raw))
    if not isinstance(raw, str):
        return None
    raw = DATE_MARK.sub("", raw.strip())
    match = DATE_TEXT.fullmatch(raw) or DATE_DIGITS.fullmatch(raw)
    if match:
        year, month, day, *clock = list(match.groups()) + [None] * 3
        try:
            return dt.datetime(int(year), int(month), int(day or 1), *(int(part or 0) for part in clock[:3]))
        except ValueError:
            return None
    return _month_or_quarter(raw)


def _month_or_quarter(raw):
    """2026 Aug, Aug. 2026, August 2026, Aug-26, 26년 8월, 2026 Q3, Q3 2026, 3Q26, 2026년 3분기, 30 June 2017, June 30, 2017, March quarter 2026."""
    for pattern, day_at, month_at, year_at in ((DAY_MONTH_YEAR, 1, 2, 3), (MONTH_DAY_YEAR, 2, 1, 3)):
        match = pattern.fullmatch(raw)
        if match:
            month = MONTH_NUMBERS.get(match.group(month_at)[:3].lower())
            try:
                return dt.datetime(int(match.group(year_at)), month, int(match.group(day_at))) if month else None
            except ValueError:
                return None
    match = MONTH_QUARTER.fullmatch(raw)
    if match:
        month = MONTH_NUMBERS.get(match.group(1)[:3].lower())
        return dt.datetime(int(match.group(2)), month, 1) if month else None
    for pattern, year_at, month_at in ((YEAR_MONTH_NAME, 1, 2), (MONTH_NAME_YEAR, 2, 1), (MONTH_NAME_SHORT, 2, 1)):
        match = pattern.fullmatch(raw)
        if match:
            month = MONTH_NUMBERS.get(match.group(month_at)[:3].lower())
            year = int(match.group(year_at))
            return dt.datetime(year if year > 99 else 2000 + year, month, 1) if month else None
    match = KOREAN_SHORT_MONTH.fullmatch(raw)
    if match:
        return dt.datetime(2000 + int(match.group(1)), int(match.group(2)), 1) if 1 <= int(match.group(2)) <= 12 else None
    for pattern, year_at, quarter_at in ((QUARTER_YEAR_FIRST, 1, 2), (QUARTER_Q_FIRST, 2, 1), (QUARTER_Q_AFTER, 2, 1)):
        match = pattern.fullmatch(raw)
        if match:
            year = int(match.group(year_at))
            return dt.datetime(year if year > 99 else 2000 + year, (int(match.group(quarter_at)) - 1) * 3 + 1, 1)
    return None


def as_month(raw):
    """The month number of a month written with no year, such as 01월 or Jan, otherwise None."""
    if not isinstance(raw, str):
        return None
    match = MONTH_ALONE.fullmatch(raw.strip())
    if match:
        return int(match.group(1)) if 1 <= int(match.group(1)) <= 12 else None
    return MONTH_NUMBERS.get(raw.strip().rstrip(".")[:3].lower()) if raw.strip().rstrip(".").isalpha() and raw.strip().rstrip(".").lower() in MONTH_WORDS else None


def as_year(raw):
    """January 1 of a year written alone, such as 2021 or 2021년도, otherwise None."""
    if isinstance(raw, (int, float)) and not isinstance(raw, bool) and float(raw).is_integer():
        raw = str(int(raw))
    match = YEAR_TEXT.fullmatch(raw.strip()) if isinstance(raw, str) else None
    if match:
        return dt.datetime(int(match.group(1)), 1, 1)
    match = FISCAL_YEAR.fullmatch(raw.strip()) if isinstance(raw, str) else None
    if match:
        century = int(match.group(1)) // 100 * 100
        ending = century + int(match.group(2))
        return dt.datetime(ending if ending > int(match.group(1)) else ending + 100, 1, 1)
    return None


def as_number(raw):
    """A number, or the figure leading a text cell such as 1,234 or 362(37%) or 15건. None for anything else, including - and 78.5%."""
    if isinstance(raw, bool):
        return None
    if isinstance(raw, (int, float)):
        return raw
    match = NUMBER_TEXT.fullmatch(raw) if isinstance(raw, str) else None
    if not match:
        return None
    number = float(match.group(1).replace(",", ""))
    return int(number) if number.is_integer() else number


def as_part_of_date(raw):
    """A month or a day held in a column of its own: 9, 09, 9월, Sep. None for anything else."""
    if isinstance(raw, bool):
        return None
    if isinstance(raw, (int, float)) and float(raw).is_integer():
        return int(raw)
    if isinstance(raw, str):
        if raw.strip().isdigit():
            return int(raw.strip())
        return as_month(raw)
    return None


@dataclass(frozen=True)
class Aggregate:
    """합계(D:E), 합계(D*F), 건수(A), each with optional conditions such as B=민원."""
    kind: str  # sum, count, mean or rate
    word: str  # as written, for headers and sources
    lines: tuple
    product: bool
    filters: tuple
    numerator: tuple = None  # for a rate: (line, text) a row must match to count


@dataclass(frozen=True)
class Ref:
    """Which cells a key means. A line is a column, or a row when by_row. A position is a place along that line."""
    sheets: tuple
    by_row: bool
    first: tuple  # the date line, or the lines that name an item in a table without dates
    values: tuple  # line numbers and Aggregates
    start: int = 1
    end: int = 0  # 0 runs to the end of the sheet
    date_parts: tuple = ()  # 연도+월: the year, month and perhaps day lines of a date split across cells

    @property
    def date(self):
        return self.first[0]

    @property
    def aggregated(self):
        return any(isinstance(value, Aggregate) for value in self.values)

    @property
    def plain(self):
        """The value lines of a dated reference. 최신:일별!A:C means A then B and C, as it always has."""
        return tuple(self.first[1:]) + tuple(value for value in self.values if not isinstance(value, Aggregate))

    @property
    def label_lines(self):
        return self.first if self.values else self.first[:1]

    @property
    def specs(self):
        return self.values if self.values else tuple(self.first[1:])

    @property
    def spec_count(self):
        return len(self.first) - 1 + len(self.values)

    def at(self, position, line):
        return (line, position) if self.by_row else (position, line)

    def last(self, ws):
        return self.end or (ws.max_column if self.by_row else ws.max_row)


@dataclass
class Series:
    ref: Ref
    points: list
    newest: tuple  # (sheet number, position) of the newest row with any value, which may not be a number
    line: int  # the line whose cell format the figures use, or None when the figure is not any one cell
    header: str
    notes: list
    rate: bool = False  # a rate or a score: it moves in points
    scale: int = 1
    percent: bool = False  # the figure is a share, computed here, and shows as one
    cut_short: tuple = None  # (rows found, rows usual) when the newest period has fewer rows than usual
    per_cent_unit: bool = False  # the sheet says its figures are per cent, and the cells do not show it

    @property
    def label(self):
        """What the figure is called in a sentence. header stays as the sheet has it, for telling which language a file is in."""
        return facts.readable_name(self.header)

    def _cell(self, sheet, position):
        row, column = self.ref.at(position, self.line if self.line is not None else self.ref.date)
        return f"{self.ref.sheets[sheet].title}!{get_column_letter(column)}{row}"

    def cell_text(self, point):
        return self._cell(point.sheet, point.row)

    def newest_cell(self):
        return self._cell(*self.newest)

    def source(self, point):
        return f"{point.date.date().isoformat()} · {self.cell_text(point)}"

    def newest_issues(self, source):
        if self.line is None:
            return []
        return source.read(self.ref.sheets[self.newest[0]], *self.ref.at(self.newest[1], self.line))[3]

    def show(self, source):
        if self.percent:
            return lambda value: format_value(value, "0.0%")[0]
        if self.line is None:
            return facts.plain_number
        newest = self.points[-1]
        show = source._prose_show(self.ref.sheets[newest.sheet], *self.ref.at(newest.row, self.line))
        return _with_per_cent(show) if self.per_cent_unit else show


@dataclass
class Comparison:
    """A table of named items with no dates."""
    items: list
    header: str
    total: float
    before_total: float
    additive: bool
    show: object
    issues: list
    source: str
    label_header: str
    stated: bool
    rate: bool = False
    scale: int = 1

    @property
    def label(self):
        return facts.readable_name(self.header)

    def ranked(self):
        return facts.ranked_items(self.items)


def _split_top(spec):
    """Split on commas that are not inside brackets. 합계(D, B=민원),C gives two parts."""
    parts, depth, current = [], 0, ""
    for letter in spec:
        depth += (letter == "(") - (letter == ")")
        if letter == "," and depth == 0:
            parts.append(current.strip())
            current = ""
        else:
            current += letter
    return parts + [current.strip()]


def _line(text):
    """(line number, by_row) of A or 5."""
    text = text.replace("$", "").strip().upper()
    if text.isdigit():
        if int(text) < 1:
            raise ValueError(text)
        return int(text), True
    if not re.fullmatch(r"[A-Z]{1,3}", text):
        raise ValueError(text)
    return column_index_from_string(text), False


def _parse_part(text):
    """One piece of a reference: (lines or Aggregate, by_row or None, (start, end) or None). Raises ValueError."""
    match = AGGREGATE.fullmatch(text)
    if match:
        return _parse_aggregate(match)
    text = text.replace("$", "").strip().upper()
    match = BOUNDED_LINE.fullmatch(text)
    if match:
        col_a, row_a, col_b, row_b = match.groups()
        col_a, col_b, row_a, row_b = column_index_from_string(col_a), column_index_from_string(col_b), int(row_a), int(row_b)
        if col_a == col_b and row_a <= row_b:
            return (col_a,), False, (row_a, row_b)
        if row_a == row_b and col_a <= col_b:
            return (row_a,), True, (col_a, col_b)
        raise ValueError(text)
    first, mark, last = text.partition(":")
    start, by_row = _line(first)
    end, other = _line(last) if mark else (start, by_row)
    if by_row != other or end < start:
        raise ValueError(text)
    return tuple(range(start, end + 1)), by_row, None


def _parse_aggregate(match):
    word, inner = match.group(1), match.group(2)
    pieces = _split_top(inner)
    kind = AGGREGATE_KINDS[word.lower()]
    numerator = None
    if kind == "rate":  # 비율(B=Yes): the share of rows matching. Any further pieces narrow the rows it is a share of.
        name, mark, wanted = pieces[0].partition("=")
        if not mark or not wanted.strip():
            raise ValueError(pieces[0])
        numerator = (_line(name)[0], wanted.strip())
        pieces[0] = name
    filters, by_row = [], None
    for piece in pieces[1:]:
        name, mark, wanted = piece.partition("=")
        if not mark or not wanted.strip():
            raise ValueError(piece)
        line, way = _line(name)
        by_row = way if by_row is None else by_row
        if way != by_row:
            raise ValueError(piece)
        filters.append((line, wanted.strip()))
    main = pieces[0].strip()
    product = "*" in main
    if product:
        lines, ways = zip(*(_line(part) for part in main.split("*")))
        way = ways[0]
        if len(set(ways)) != 1:
            raise ValueError(main)
    else:
        first, mark, last = main.partition(":")
        start, way = _line(first)
        end, other = _line(last) if mark else (start, way)
        if way != other or end < start:
            raise ValueError(main)
        lines = tuple(range(start, end + 1))
    if by_row is not None and by_row != way:
        raise ValueError(main)
    return Aggregate(kind, word, tuple(lines), product, tuple(filters), numerator), way, None


def _aggregate_name(spec, names):
    """What a total, a mean or a share of rows is called in a sentence."""
    if spec.kind == "count":
        name = spec.word
    elif spec.kind == "rate":
        name = f"{spec.word} ({spec.numerator[1]})"
    else:
        name = (" × " if spec.product else " + ").join(names) if len(names) <= 3 else spec.word
        if spec.kind == "mean":
            name = f"{spec.word} {name}"
    if spec.filters:
        name += " (" + ", ".join(text for _line, text in spec.filters) + ")"
    return name


def _line_name(ref, line):
    return str(line) if ref.by_row else get_column_letter(line)


def _sheet_names(ref):
    return ref.sheets[0].title if len(ref.sheets) == 1 else f"{ref.sheets[0].title} … {ref.sheets[-1].title}"


def _with_per_cent(show):
    return lambda value: show(value) + "%"


def unit_note(ws, rows, columns):
    """(the words as written, the unit alone) of a note saying how a sheet's figures are measured. (단위: 백만원) says it in one cell; Unit: may leave it to the cell beside or below. ("", "") when there is none."""
    for row in range(1, min(ws.max_row, rows) + 1):
        for column in range(1, min(ws.max_column, columns) + 1):
            cell = ws._cells.get((row, column))  # only cells that exist: asking for rows makes a cell at every place
            if cell is None or not isinstance(cell.value, str):
                continue
            text = cell.value.strip()
            match = UNIT_NOTE.search(text)
            if not match or len(text) > UNIT_NOTE_LENGTH:
                continue
            rest = text[match.end():].strip(" )）]:：")
            if rest:
                return text, rest
            for place in ((row, column + 1), (row + 1, column)):
                beside = ws._cells.get(place)
                words = beside.value.strip() if beside is not None and isinstance(beside.value, str) else ""
                if words and len(words) <= UNIT_WORD_LENGTH:
                    return f"{text} {words}", words
    return "", ""


def _reads_backwards(first, second):
    """Two periods are read first to second. Sheets often put this year first, and 매출 then 전년 read that way turns a rise into a fall with nothing to show for it."""
    years = [re.search(r"(?:19|20)\d\d", name or "") for name in (first, second)]
    if all(years):
        return int(years[0].group(0)) > int(years[1].group(0))
    return bool(EARLIER_NAME.search(second or "")) and not EARLIER_NAME.search(first or "")


def in_per_cent(ws):
    """The sheet itself says its figures are per cent, not merely that a % appears somewhere on it."""
    if "_in_per_cent" not in ws.__dict__:
        ws.__dict__["_in_per_cent"] = PER_CENT_UNIT.fullmatch(unit_note(ws, UNIT_SCAN_ROWS, UNIT_SCAN_COLUMNS)[1]) is not None
    return ws.__dict__["_in_per_cent"]


def _stated_day(text):
    """A date written out in a sentence, or None."""
    words = DAY_FIRST.search(text)
    if words:
        month = next((place for place, name in enumerate(MONTH_NAMES, start=1) if words.group(2).casefold().startswith(name)), None)
        numbers = (int(words.group(3)), month, int(words.group(1))) if month else None
    else:
        digits = YEAR_FIRST.search(text)
        numbers = tuple(int(part) for part in digits.groups()) if digits else None
    try:
        return dt.date(*numbers) if numbers else None
    except ValueError:
        return None


def _clean_label(text):
    """A heading as a person reads it: 유\\n효\\n수\\n량 and 일 생 산 능 력 are written apart only to fill a tall or wide cell."""
    parts = str(text).split()
    return "".join(parts) if parts and all(len(part) == 1 for part in parts) else " ".join(parts)


def _merge_origin(ws, row, column):
    """The first cell of the merged range this cell is in, or None."""
    origins = ws.__dict__.get("_merge_origins")
    if origins is None:
        origins = {}
        for merged in ws.merged_cells.ranges:
            for r in range(merged.min_row, merged.max_row + 1):
                for c in range(merged.min_col, merged.max_col + 1):
                    origins[(r, c)] = (merged.min_row, merged.min_col)
        ws.__dict__["_merge_origins"] = origins
    return origins.get((row, column))


def _merged_value(ws, row, column):
    """The value a person sees in a cell: inside a merged range, that of its first cell."""
    return ws.cell(*(_merge_origin(ws, row, column) or (row, column))).value


def _rows_of_an_outer_level(ws, ref, labels):
    """Closing lines whose label is merged in from the column before: 무료, across 지역 and 시설, adds up the rows above it and is not one more 시설 beside them. Such lines at the top of a table are its main figures, and those stay."""
    if len(labels) != 1 or labels[0] < 2:
        return frozenset()
    _merged_value(ws, 1, 1)  # makes sure the map of merged cells exists
    origins = ws.__dict__["_merge_origins"]
    if not origins:
        return frozenset()
    line = labels[0]
    filled, outer, last_own = 0, set(), 0
    for position in range(ref.start, ref.last(ws) + 1):
        here = ref.at(position, line)
        if here in origins and origins[here] == origins.get(ref.at(position, line - 1)):
            outer.add(position)
            filled += 1
        elif ws.cell(*here).value not in (None, ""):
            filled += 1
            last_own = position
    closing = {position for position in outer if position > last_own}
    return frozenset(closing) if len(outer) < filled * OUTER_LEVEL_SHARE else frozenset()


def _rows_indented_under_others(ws, ref, labels, figures):
    """A pivot table's lines indented under the one above: schools under regions under terms, all in one column. Only the outermost can stand side by side, and only where more than one line sits at that level. A heading with no figure beside it is not a line of the table."""
    if len(labels) != 1:
        return frozenset()
    indents = {}
    for position in range(ref.start, ref.last(ws) + 1):
        cell = ws._cells.get(ref.at(position, labels[0]))
        beside = [ws._cells.get(ref.at(position, line)) for line in figures]
        if cell is not None and isinstance(cell.value, str) and cell.value.strip() and any(other is not None and as_number(other.value) is not None for other in beside):
            indents[position] = cell.alignment.indent or 0
    if len(set(indents.values())) < 2:
        return frozenset()
    outermost = min(indents.values())
    top = [position for position, indent in indents.items() if indent == outermost and "".join(str(ws._cells[ref.at(position, labels[0])].value).split()).casefold() not in TOTAL_WORDS and not TOTAL_ROW.match(str(ws._cells[ref.at(position, labels[0])].value).strip()) and not GRAND_TOTAL.fullmatch(str(ws._cells[ref.at(position, labels[0])].value).strip())]
    if len(top) < 2:
        return frozenset()
    return frozenset(position for position, indent in indents.items() if indent > outermost)


def _exact_exists(ws, ref, line, wanted):
    """True when some cell of the line is exactly this word, so 쌀 does not also take in 찹쌀. Otherwise a part is enough: nidirect finds nidirect Contact Centre calls."""
    known = ws.__dict__.setdefault("_exact_words", {})
    key = (ref.by_row, line, wanted)
    if key not in known:
        word = wanted.casefold()
        known[key] = any(str(_merged_value(ws, *ref.at(position, line))).strip().casefold() == word for position in range(ref.start, ref.last(ws) + 1))
    return known[key]


def _matches(ws, ref, position, line, wanted):
    value = _merged_value(ws, *ref.at(position, line))
    if value is None:
        return False
    if _exact_exists(ws, ref, line, wanted):
        return str(value).strip().casefold() == wanted.casefold()
    return wanted.casefold() in str(value).casefold()


def _aggregate_value(ws, ref, position, spec, from_text=None):
    """What one row adds to an Aggregate, or None when it adds nothing."""
    for line, wanted in spec.filters:
        if not _matches(ws, ref, position, line, wanted):
            return None
    cells = [ws.cell(*ref.at(position, line)) for line in spec.lines]
    if spec.kind == "rate":
        value = _merged_value(ws, *ref.at(position, spec.numerator[0]))
        return None if value in (None, "") else int(_matches(ws, ref, position, *spec.numerator))
    if spec.kind == "count":
        return 1 if any(cell.value not in (None, "") for cell in cells) else None
    numbers = [as_number(cell.value) for cell in cells]
    if from_text is not None:
        from_text.extend(cell for cell, number in zip(cells, numbers) if number is not None and isinstance(cell.value, str))
    if spec.product:
        return None if any(number is None for number in numbers) else math.prod(numbers)
    found = [number for number in numbers if number is not None]
    return sum(found) if found else None


def _years_only(ws, ref, position, specs):
    lines = [line for spec in specs for line in (spec.lines if isinstance(spec, Aggregate) else [spec])]
    return all(as_year(ws.cell(*ref.at(position, line)).value) is not None for line in lines)


def _plain_value(ws, ref, position, line, from_text):
    cell = ws.cell(*ref.at(position, line))
    number = as_number(cell.value)
    if number is not None and isinstance(cell.value, str):
        from_text.append(cell)
    return number


def _total_kind(ws, ref, position, before, chosen=()):
    """total or subtotal if a cell before the values says so. 합계 may sit in a column left of the names, as with merged cells. A column the reference chose a row by says which rows it wants, and says nothing about them being a total of the rest."""
    for line in range(1, before):
        if line in chosen:
            continue
        value = _merged_value(ws, *ref.at(position, line))
        if isinstance(value, str):
            word = "".join(value.split()).casefold()
            if word in SUBTOTAL_WORDS:
                return "subtotal"
            if word in TOTAL_WORDS or TOTAL_ROW.match(value.strip()) or word.endswith(("합계", "총계")):
                return "total"
    return None


def _usual_rows(days, groups):
    """The steady number of rows each date has, or None when it varies."""
    sizes = [groups[day][1] for day in days[:-1]][-12:]
    if len(sizes) < 5:
        return None
    usual, times = Counter(sizes).most_common(1)[0]
    return usual if times >= len(sizes) * 0.8 else None


def _group_notes(days, groups):
    """Warn when the newest day looks cut short: fewer rows than a steady number, or a last entry much earlier than usual."""
    notes = []
    newest = groups[days[-1]]
    usual = _usual_rows(days, groups)
    if usual is not None and newest[1] != usual:
        notes.append(Issue("latest_group_size", {"date": days[-1].date().isoformat(), "found": newest[1], "usual": usual}))
    ends = [groups[day][4] for day in days[:-1] if groups[day][4] is not None][-8:]
    if len(ends) >= 5 and newest[4] is not None and newest[4] + LATE_MINUTES < statistics.median(ends):
        usual = int(statistics.median(ends))
        notes.append(Issue("latest_partial_day", {"date": days[-1].date().isoformat(), "last": f"{newest[4] // 60:02d}:{newest[4] % 60:02d}", "usual": f"{usual // 60:02d}:{usual % 60:02d}"}))
    return notes


def is_latest_key(key):
    return LATEST.fullmatch(key) is not None


def is_rank_key(key):
    return RANK.fullmatch(key) is not None


def is_calc_key(key):
    return CALC.fullmatch(key) is not None


def is_insight_key(key):
    return INSIGHT.fullmatch(key) is not None


class ExcelSource:
    """Read-only view of a workbook: resolves keys and renders cells as Excel displays them."""

    def __init__(self, path):
        _check_file(path, readers.ALL_TYPES, "excel")
        self.path = path
        self._converted = None
        self.wb = self._load(True)
        self.thousands = False  # a template promises what Excel shows; a report laid out here should match its own sentences
        self._table_lang = "ko"
        self._formula_wb = None
        self._calculated = None
        self._sheets = {ws.title.casefold(): ws for ws in self.wb.worksheets}
        for ws in self.wb.worksheets:  # nobody can see the space at the end of a sheet's name when typing it
            self._sheets.setdefault(ws.title.strip().casefold(), ws)
        self._dated = {}
        self._made = {}

    def _load(self, data_only):
        try:
            if os.path.splitext(self.path)[1].lower() in readers.OTHER_TYPES:
                if self._converted is None:  # no formulas to tell apart from their values, so once will do
                    self._converted = readers.load(self.path)
                return self._converted
            return openpyxl.load_workbook(self.path, data_only=data_only)
        except PermissionError:
            raise SlideFillError("file_locked", path=self.path)
        except readers.Protected:
            raise SlideFillError("file_protected", path=self.path)
        except Exception as exc:  # the readers raise many types on damaged or unusual files
            raise SlideFillError("file_unreadable", path=self.path, detail=str(exc))

    def locate(self, key):
        """Turn a key into (worksheet, (min_col, min_row, max_col, max_row), source text). Raises IssueError."""
        key = key.translate(SMART_QUOTES).strip()
        if "!" in key:
            return self._locate_ref(key)
        return self._locate_name(key)

    def _sheet(self, name):
        if name.casefold() in self._sheets:  # a sheet's name can end in a space
            return self._sheets[name.casefold()]
        name = name.strip()
        if len(name) >= 2 and name[0] == name[-1] == "'":
            name = name[1:-1].replace("''", "'")
        ws = self._sheets.get(name.casefold())
        if ws is None:
            raise IssueError("missing_sheet", sheet=name)
        return ws

    def _locate_ref(self, key):
        sheet, _, coord = key.rpartition("!")
        ws = self._sheet(sheet)
        coord = coord.replace("$", "").strip()
        if not COORD.fullmatch(coord):
            raise IssueError("bad_ref", ref=key)
        try:
            bounds = range_boundaries(coord.upper())
        except ValueError:
            raise IssueError("bad_ref", ref=key)
        return ws, bounds, f"{ws.title}!{coord.upper()}"

    def _locate_name(self, key):
        wanted = key.casefold()
        names = [dn for name, dn in self.wb.defined_names.items() if name.casefold() == wanted]
        if not names:
            for ws in self.wb.worksheets:
                names += [dn for name, dn in ws.defined_names.items() if name.casefold() == wanted]
        if len(names) > 1:
            raise IssueError("multi_area", name=key)
        if names:
            return self._locate_defined(key, names[0])
        for ws in self.wb.worksheets:
            for name, ref in ws.tables.items():  # openpyxl gives (table name, range text) pairs here
                if name.casefold() == wanted:
                    return ws, range_boundaries(ref), f"{ws.title}!{ref}"
        raise IssueError("missing_name", name=key)

    def _locate_defined(self, key, defined):
        if defined.type != "RANGE":
            raise IssueError("name_not_cells", name=key)  # a constant or a formula such as OFFSET
        destinations = list(defined.destinations)
        if len(destinations) != 1:
            raise IssueError("multi_area", name=key)
        sheet, coord = destinations[0]
        ws = self._sheets.get(sheet.casefold())
        if ws is None:
            raise IssueError("missing_sheet", sheet=sheet)
        coord = coord.replace("$", "")
        return ws, range_boundaries(coord), f"{ws.title}!{coord}"

    def _cell_format(self, cell):
        fmt = cell.number_format
        if fmt in (None, "General"):  # openpyxl reports ids it does not know as General
            return KOREAN_BUILTIN_FORMATS.get(cell._style.numFmtId, "General")
        return fmt

    def _never_calculated(self):
        """True when no formula in the file has a stored result, i.e. it was never saved by Excel."""
        if self._calculated is None:
            self._calculated = False
            self._formula_wb = self._formula_wb or self._load(False)
            for ws in self._formula_wb.worksheets:
                for cell in list(ws._cells.values()):  # only cells that exist: iter_rows would make every one a sheet's range covers, and one reaching column XFC is billions
                    if cell.data_type == "f" and self.wb[ws.title][cell.coordinate].value is not None:
                        self._calculated = True
                        return False
        return not self._calculated

    def _has_formula(self, ws, row, col):
        self._formula_wb = self._formula_wb or self._load(False)
        return self._formula_wb[ws.title].cell(row, col).data_type == "f"

    def read(self, ws, row, col):
        """Return (raw value, display text, format, issues). Text is None when there is no trustworthy value."""
        cell = ws.cell(row, col)
        address = f"{ws.title}!{get_column_letter(col)}{row}"
        raw = cell.value
        if cell.data_type == "e":
            return raw, str(raw), "General", [Issue("error_value", {"cell": address, "error": str(raw)})]
        if raw is None:
            if self._has_formula(ws, row, col) and self._never_calculated():
                return None, None, "General", [Issue("formula_no_cache", {"cell": address})]
            return None, "", "General", []
        fmt = self._cell_format(cell)
        text, exact = format_value(raw, fmt, self.wb.epoch)
        issues = [] if exact else [Issue("format_fallback", {"cell": address, "format": fmt})]
        return raw, text, fmt, issues

    def _latest_request(self, key):
        """Parse 최신7:시트!A,C:E into (reference, row count). The first column holds the dates."""
        match = LATEST.fullmatch(key.translate(SMART_QUOTES))
        count = int(match.group(1) or 1)
        if count < 1:
            raise IssueError("latest_bad", key=key)
        return self._parse_ref(match.group(2), key), count

    def _named_sheets(self, name):
        """The sheet with this name, or every sheet matching * and ?, which a real sheet name cannot contain."""
        if "*" not in name and "?" not in name:
            return (self._sheet(name),)
        pattern = name.strip().strip("'").casefold()
        found = tuple(ws for ws in self.wb.worksheets if fnmatch.fnmatchcase(ws.title.casefold(), pattern))
        if not found:
            raise IssueError("missing_sheet", sheet=name.strip())
        return found

    def _parse_ref(self, body, key):
        """Parse 시트!A,C:E, 시트!5,7, 시트!A10:A400,C10:C400, 여러*!A,합계(D:E, B=민원) into a Ref."""
        sheet, mark, spec = body.rpartition("!")
        parts = _split_top(spec)
        if not mark or not all(parts):
            raise IssueError("latest_bad", key=key)
        date_parts = ()
        try:
            if "+" in parts[0]:  # 연도+월: the year and the month in separate columns
                pieces = [_line(piece) for piece in parts[0].split("+")]
                if not 2 <= len(pieces) <= 3 or len({way for _line_number, way in pieces}) != 1:
                    raise ValueError(parts[0])
                date_parts = tuple(line for line, _way in pieces)
                parts[0] = parts[0].split("+")[0]
            parsed = [_parse_part(part) for part in parts]
        except ValueError:
            raise IssueError("latest_bad", key=key)
        ways = {by_row for _value, by_row, _span in parsed if by_row is not None}
        spans = {span for _value, _by_row, span in parsed if span is not None}
        if len(ways) > 1 or len(spans) > 1 or isinstance(parsed[0][0], Aggregate):
            raise IssueError("latest_bad", key=key)
        values = []
        for value, _by_row, _span in parsed[1:]:
            values.extend([value] if isinstance(value, Aggregate) else value)
        start, end = next(iter(spans), (1, 0))
        return Ref(self._named_sheets(sheet), ways == {True}, tuple(parsed[0][0]), tuple(values), start, end, date_parts)

    def _dated_rows(self, ws, ref):
        """(date, position) for every row, or column, whose date cell is a date, so subtotal and total rows drop out.

        A bare year such as 2021 only counts as a date where most cells are years, since it is also an ordinary number.
        A month with no year, such as 01월, takes its year from the cell above it, which is how a merged year heading reads.
        """
        cache = (ws.title, ref.by_row, ref.date, ref.start, ref.end, ref.date_parts)
        if cache not in self._dated:
            dates, years, months = {}, {}, {}
            fiscal = None if ref.date_parts else self._fiscal_years(ws, ref)
            for position in range(ref.start, ref.last(ws) + 1):
                if fiscal is not None:
                    if position in fiscal:
                        dates[position] = fiscal[position]
                    continue
                if ref.date_parts:
                    date = self._date_from_parts(ws, ref, position)
                    if date is not None:
                        dates[position] = date
                    continue
                raw = ws.cell(*ref.at(position, ref.date)).value
                date = as_date(raw)
                if date is not None:
                    dates[position] = date
                elif (year := as_year(raw)) is not None:
                    years[position] = year
                elif (month := as_month(raw)) is not None:
                    months[position] = month
            for position, month in months.items():
                year = self._year_above(ws, ref, position)
                if year is not None:
                    dates[position] = year.replace(month=month)
            if len(set(years.values())) >= 2 and len(years) > len(dates):
                dates.update(years)
            self._dated[cache] = [(dates[position], position) for position in sorted(dates)]
        return self._dated[cache]

    @staticmethod
    def _fiscal_years(ws, ref):
        """{position: January 1 of the ending year} when a whole line reads 2011/12, 2012/13: financial years, not December 2011 and a month 13 that does not exist. Else None."""
        found, other = {}, 0
        for position in range(ref.start, ref.last(ws) + 1):
            raw = ws.cell(*ref.at(position, ref.date)).value
            match = YEAR_SPAN.fullmatch(raw.strip()) if isinstance(raw, str) else None
            if match and (int(match.group(1)) + 1) % 100 == int(match.group(2)):
                found[position] = dt.datetime(int(match.group(1)) + 1, 1, 1)
            elif as_date(raw) is not None:
                other += 1
        if len(set(found.values())) < 2 or other > len(found) * 0.2:
            return None
        return found

    @staticmethod
    def _date_from_parts(ws, ref, position):
        """A date split across cells, such as 2026 and 09. None unless every part reads and the date is real."""
        year = as_year(ws.cell(*ref.at(position, ref.date_parts[0])).value)
        others = [as_part_of_date(ws.cell(*ref.at(position, line)).value) for line in ref.date_parts[1:]]
        if year is None or None in others:
            return None
        try:
            return year.replace(month=others[0], day=others[1] if len(others) > 1 else 1)
        except ValueError:
            return None

    def _date_cell(self, ref, sheet, position):
        """(raw, text, format) of a row's date as a person should read it. A date split across cells is shown joined."""
        if ref.date_parts:
            date = self._date_from_parts(ref.sheets[sheet], ref, position)
            text = date.strftime("%Y-%m-%d" if len(ref.date_parts) == 3 else "%Y-%m")
            return date, text, "General"
        return self.read(ref.sheets[sheet], *ref.at(position, ref.date))[:3]

    def _date_header(self, ws, ref, top):
        return " ".join(self._header(ws, ref, top, line) for line in (ref.date_parts or (ref.date,)))

    @staticmethod
    def _year_above(ws, ref, position):
        """The year a month belongs to: a section heading such as Year 2026 in its own line, or a merged heading above it."""
        for back in range(position - 1, 0, -1):
            raw = ws.cell(*ref.at(back, ref.date)).value
            if raw in (None, "") or as_month(raw) is not None:
                continue
            match = YEAR_HEADING.fullmatch(str(raw).strip())
            if match:
                return dt.datetime(int(match.group(1)), 1, 1)
            break
        for line in range(ref.date - 1, max(ref.date - 4, 0), -1):
            for back in range(position, 0, -1):
                raw = ws.cell(*ref.at(back, line)).value
                if raw not in (None, ""):
                    year = as_year(raw)
                    if year is not None:
                        return year
                    break
        return None

    def _all_rows(self, ref):
        """(date, sheet number, position) for every dated row of every sheet, oldest first."""
        return sorted((date, i, position) for i, ws in enumerate(ref.sheets) for date, position in self._dated_rows(ws, ref))

    def _filled(self, ws, row, column):
        """A formula with no saved result counts as filled, so read() can report it instead of an older row being used."""
        value = ws.cell(row, column).value
        if value not in (None, ""):
            return True
        return value is None and self._has_formula(ws, row, column) and self._never_calculated()

    def _has_figure(self, ws, row, column):
        """A number, a date, or a cell read() must report: anything but text, which may only be a label waiting for figures."""
        cell = ws.cell(row, column)
        if cell.data_type == "e" or isinstance(cell.value, (dt.date, dt.time)) or as_number(cell.value) is not None:
            return True
        return cell.value is None and self._filled(ws, row, column)

    def _latest_rows(self, ref):
        """Dated rows, oldest first, that hold a value. A row dated ahead with nothing in it yet is not the latest.

        Where these columns hold figures, a row with only text is not filled: a queue name printed ahead for each coming month has no data yet.
        """
        rows = []
        for row in self._all_rows(ref):
            ws = ref.sheets[row[1]]
            others = ref.plain or self._other_lines(ws, ref)
            if any(self._filled(ws, *ref.at(row[2], line)) for line in others):
                rows.append(row)
        rows = [row for row in rows if any(self._has_figure(ref.sheets[row[1]], *ref.at(row[2], line)) for line in (ref.plain or self._other_lines(ref.sheets[row[1]], ref)))] or rows
        if not rows:
            raise IssueError("latest_no_dates", sheet=ref.sheets[0].title, column=_line_name(ref, ref.date))
        return rows

    def _other_lines(self, ws, ref):
        """Every line but the date. Across a row, that is the table under the dates: it ends at the first row with no figures, as another table may follow."""
        cache = ("others", ws.title, ref.by_row, ref.date, ref.start, ref.end)
        if cache not in self._dated:
            if not ref.by_row:
                lines = [line for line in range(1, ws.max_column + 1) if line != ref.date and line not in ref.date_parts]
            else:
                lines = []
                for line in range(ref.date + 1, ws.max_row + 1):
                    if not any(self._has_figure(ws, *ref.at(position, line)) for position in range(ref.start, ref.last(ws) + 1)):
                        break
                    lines.append(line)
            self._dated[cache] = lines
        return self._dated[cache]

    def _older_than_sheet(self, ref, used):
        """Warn when a newer row has figures elsewhere, so this value and the rest of the report may be different days."""
        for date, i, position in reversed(self._all_rows(ref)):
            if date <= used:
                break
            ws = ref.sheets[i]
            # A 0 printed ahead, as a formula on an empty month gives, is not a newer value.
            if any(self._has_figure(ws, *ref.at(position, line)) and ws.cell(*ref.at(position, line)).value != 0 for line in self._other_lines(ws, ref)):
                return [Issue("latest_older", {"used": used.date().isoformat(), "newest": date.date().isoformat()})]
        return []

    def _undated_groups(self, ref):
        """Warn when the date is only on the first row of each group, as with merged cells: the rows below it are not being read."""
        groups = 0
        for ws in ref.sheets:
            dated = [position for _date, position in self._dated_rows(ws, ref)]
            for position, following in zip(dated, dated[1:] + [ref.last(ws) + 1]):
                below = position + 1
                if below < following and ws.cell(*ref.at(below, ref.date)).value in (None, "") and any(self._has_figure(ws, *ref.at(below, line)) for line in ref.plain):
                    groups += 1
        return [Issue("latest_grouped", {"column": _line_name(ref, ref.date), "count": groups})] if groups >= 2 else []

    def _newer_text(self, ref, used):
        """Warn when rows newer than the one used say something other than a figure, such as (결측) or 집계중: the latest figure is older than the sheet."""
        rows = self._all_rows(ref)
        for line in ref.plain:
            newer = [ref.sheets[i].cell(*ref.at(position, line)).value for date, i, position in rows if date > used]
            words = [value for value in newer if isinstance(value, str) and value.strip() and as_number(value) is None]
            if words and any(as_number(ref.sheets[i].cell(*ref.at(position, line)).value) is not None for date, i, position in rows if date <= used):
                return [Issue("latest_text_newer", {"used": used.date().isoformat(), "count": len(words), "text": words[-1].strip()})]
        return []

    def _row_notes(self, ref, used):
        return self._older_than_sheet(ref, used) + self._newer_text(ref, used) + self._undated_groups(ref)

    def _as_of(self):
        """The day the data can be known to reach: when the file was last saved, or today if that is earlier."""
        saved = self.wb.properties.modified
        return min(saved.date(), dt.date.today()) if saved else dt.date.today()

    def _period_issues(self, points, date_header="", ws=None, ref=None):
        """Warn when the newest figure covers less than the figures it is compared with."""
        months = facts.partial_year(points)
        if months:
            return [Issue("latest_partial_year", {"date": points[-1].date.strftime("%Y-%m"), "months": months})]
        as_of = self._as_of()
        if ws is not None and not ref.date_parts and self._fiscal_years(ws, ref):  # 2025-26 is read as 2026, and it is over long before 2026 is
            if as_of >= dt.date(points[-1].date.year, *FISCAL_YEAR_ENDS):
                return []
            return [self._unfinished_issue(points, as_of)]
        # 학년도 2026 is not the calendar year 2026, so a calendar cannot say it is unfinished.
        if not SCHOOL_YEAR.search(date_header) and facts.unfinished(points, as_of):
            return [self._unfinished_issue(points, as_of)]
        return []

    def _unfinished_issue(self, points, as_of):
        params = {"date": points[-1].date.date().isoformat(), "saved": as_of.isoformat()}
        stated = self._stated_as_of()
        # The date the file itself gives says more than the day it was saved, if it falls where it could: another date in a note must not move the period.
        if stated and stated <= as_of and len(points) >= 2 and stated > points[-2].date.date():
            params["as_at"] = stated.isoformat()
        return Issue("latest_unfinished", params)

    def _prose_show(self, ws, row, column):
        """How a figure reads in a sentence: the cell's own format, but a cell with no format gets sensible digits rather than all of them."""
        fmt = self._cell_format(ws.cell(row, column))
        if fmt in ("General", "@", "0"):
            return facts.plain_number
        return lambda value: format_value(value, fmt, self.wb.epoch)[0]

    def _latest_date_text(self, ref):
        """{{최신:시트!A+B}}: the newest date of a date split across cells."""
        date, i, position = self._latest_rows(ref)[-1]
        return self._date_cell(ref, i, position)[1], f"{date.date().isoformat()} · {ref.sheets[i].title}!{_line_name(ref, ref.date)}", self._row_notes(ref, date)

    def _fully_filled(self, rows, ref):
        """The rows filled as a row of this sheet usually is. Asked only for the date, there is no column to judge by, and a month ahead with only its target written must not pass for the latest."""
        counts = []
        for _date, i, position in rows:
            ws = ref.sheets[i]
            counts.append(sum(self._has_figure(ws, *ref.at(position, line)) for line in self._other_lines(ws, ref)))
        usual = Counter(counts).most_common(1)[0][0]
        return [row for row, count in zip(rows, counts) if count >= usual] or rows

    def _latest_text(self, ref, key):
        rows = self._latest_rows(ref)
        if not ref.plain:
            rows = self._fully_filled(rows, ref)
        if len(ref.plain) > 1:
            raise IssueError("range_for_text", key=key, source=key)
        newest = rows[-1][0]
        line = ref.plain[-1] if ref.plain else ref.date
        same_day = [(i, position) for date, i, position in rows if date == newest]
        if len({self.read(ref.sheets[i], *ref.at(position, line))[1] for i, position in same_day}) > 1:
            raise IssueError("latest_ambiguous", date=newest.date().isoformat(), count=len(same_day))
        i, position = same_day[-1]
        row, column = ref.at(position, line)
        cell = f"{ref.sheets[i].title}!{get_column_letter(column)}{row}"
        # The date goes first: it is how you spot an old workbook, and the list may clip the end of this text.
        return ref.sheets[i], row, column, f"{newest.date().isoformat()} · {cell}", self._row_notes(ref, newest)

    def _header(self, ws, ref, top, line):
        """The nearest filled cell before the first data row or column, or the line's name when the sheet has no header."""
        names = [(position, _clean_label(text)) for position in range(top - 1, 0, -1)
                 if (text := self.read(ws, *ref.at(position, line))[1]) and "," not in text.replace(", ", "")[:6] and (as_number(ws.cell(*ref.at(position, line)).value) is None or as_year(ws.cell(*ref.at(position, line)).value))]
        if not names:
            return _line_name(ref, line)
        position, name = names[0]
        if "".join(name.split()).casefold() in TOTAL_WORDS | SUBTOTAL_WORDS:
            parent = self._heading_over(ws, ref, position - 1, line)
            if parent and "".join(parent.split()).casefold() in TOTAL_WORDS and "".join(name.split()).casefold() in TOTAL_WORDS:
                return parent  # 계 under 합계 is one total
            return f"{parent} {name}" if parent else name
        useful = [later for _place, later in names if not ID_CODE.fullmatch(later)]
        chosen = useful[0] if useful else name
        if UNIT_HEADING.fullmatch(chosen) and len(useful) > 1:  # (Billion Kilowatthours) says how it is measured, not what it is
            return f"{useful[1]} {chosen}"
        return self._with_merged_heading(ws, ref, next(place for place, later in names if later == chosen), line, chosen)

    def _with_merged_heading(self, ws, ref, position, line, name):
        """건수 under 일반수출신고, with another 건수 under 간이수출신고: alone it reads as the whole. Only a heading placed over this very column counts, so a sheet's title is not borrowed."""
        if position < 2:
            return name
        last = min(ws.max_column if ref.at(1, 2)[1] == 2 else ws.max_row, MAX_LINES_FOR_A_REPEAT)
        if not any(other != line and _clean_label(self.read(ws, *ref.at(position, other))[1] or "") == name for other in range(1, last + 1)):
            return name
        above = _merged_value(ws, *ref.at(position - 1, line))
        if above in (None, "") or as_number(above) is not None:
            return name
        above = _clean_label(str(above))
        return f"{above} {name}" if len(above) >= MIN_HEADING_LENGTH and above != name else name  # C over 제조업 is an industry's code letter

    def _heading_over(self, ws, ref, position, line):
        """A heading merged across several columns sits in the first of them, so a 소계 under it looks back along the row above. A title alone on that row is about the sheet, not this column."""
        if position < 1:
            return ""
        for back in range(line, 0, -1):
            text = self.read(ws, *ref.at(position, back))[1]
            if text:
                if PIVOT_HEADING.fullmatch(text.strip()):  # Column Labels says nothing about the figure
                    return ""
                alone =not any(self.read(ws, *ref.at(position, other))[1] for other in range(1, line + HEADING_LOOK_AHEAD) if other != back)
                if back != line and alone and (back == 1 or _merged_value(ws, *ref.at(position, line)) in (None, "")):
                    return ""
                return _clean_label(text)
        return ""

    def _latest_range(self, ref, count):
        rows = self._latest_rows(ref)
        picked = rows[-count:]
        issues = self._row_notes(ref, picked[-1][0])
        if len(picked) != count:
            issues.append(Issue("latest_fewer", {"wanted": count, "found": len(picked)}))
        gaps = [date for date, i, position in picked if any(ref.sheets[i].cell(*ref.at(position, line)).value in (None, "") for line in ref.plain)]
        if gaps and len(ref.plain) > 1:  # targets written for the months ahead: their actual figures show as blanks
            issues.append(Issue("latest_partial_rows", {"count": len(gaps), "first": gaps[0].date().isoformat()}))
        first_sheet = ref.sheets[picked[0][1]]
        top = self._dated_rows(first_sheet, ref)[0][1]
        lines = (ref.date, *ref.plain)
        names = [self._report_date_heading(first_sheet, ref, top)] + [self._report_heading(self._header(first_sheet, ref, top, line)) for line in ref.plain]
        table = [[(text, text, "General") for text in names]]
        for _date, i, position in picked:
            line_cells = [self._date_cell(ref, i, position)]
            for line in ref.plain:
                raw, text, fmt, cell_issues = self.read(ref.sheets[i], *ref.at(position, line))
                line_cells.append((raw, text, fmt))
                issues.extend(cell_issues)
            table.append(line_cells)
        letters = ",".join(_line_name(ref, line) for line in lines)
        first, last = picked[0][0].date().isoformat(), picked[-1][0].date().isoformat()
        dates = last if first == last else f"{first} ~ {last}"
        return table, f"{dates} · {_sheet_names(ref)}!{letters}", issues

    def _report_heading(self, name):
        """A template shows the sheet's own heading. A report laid out here says it as its sentences do."""
        return facts.readable_name(name) if self.thousands else name

    def _report_date_heading(self, ws, ref, top):
        """What a laid-out table calls its first column. Across a row, the cell before the dates names the row's items, not the dates. The same goes for a sheet's name written there."""
        name = self._date_header(ws, ref, top)
        if not self.thousands:
            return name
        if ref.by_row or name.casefold() == ws.title.casefold() or name == _line_name(ref, ref.date):
            return PERIOD_WORD[self._table_lang]
        return TRAILING_TOTAL.sub("", name)

    def _latest_sum_text(self, ref, count, key):
        """{{최신:시트!A,합계(D)}}: the newest date's total. Returns (text, source, issues)."""
        if count != 1 or len(ref.values) != 1:
            raise IssueError("range_for_text", key=key, source=key)
        found = self._series(ref)
        newest = found.points[-1]
        return found.show(self)(newest.value), found.source(newest), found.notes

    def _latest_sum_range(self, ref, count):
        """{{표:최신7:시트!A,합계(D)}}: date and total for the newest dates."""
        if len(ref.values) != 1:
            raise IssueError("calc_columns", key="")
        found = self._series(ref)
        picked = found.points[-count:]
        issues = list(found.notes)
        if len(picked) != count:
            issues.append(Issue("latest_fewer", {"wanted": count, "found": len(picked)}))
        top_ws = ref.sheets[picked[0].sheet]
        top = self._dated_rows(top_ws, ref)[0][1]
        names = (self._report_date_heading(top_ws, ref, top), self._report_heading(found.header))
        table = [[(text, text, "General") for text in names]]
        show = found.show(self)
        for point in picked:
            table.append([self._date_cell(ref, point.sheet, point.row), (point.value, show(point.value), "General")])
        first, last = picked[0].date.date().isoformat(), picked[-1].date.date().isoformat()
        dates = last if first == last else f"{first} ~ {last}"
        return table, f"{dates} · {_sheet_names(ref)}!{_line_name(ref, ref.date)},{ref.values[0].word}", issues

    def _series(self, ref):
        if ref not in self._made:
            self._same_lines_on_each_sheet(ref)
            self._made[ref] = self._summed_series(ref) if ref.aggregated else self._plain_series(ref)
        return self._made[ref]

    def _same_lines_on_each_sheet(self, ref):
        """A row or column can mean something else on another sheet, so several sheets must agree on what these lines are called."""
        if len(ref.sheets) < 2:
            return
        lines = ref.values[0].lines if ref.aggregated else ref.plain[:1]
        for line in lines:
            names = {}
            for ws in ref.sheets:
                dated = self._dated_rows(ws, ref)
                if dated:
                    names.setdefault(self._header(ws, ref, dated[0][1], line).casefold(), ws.title)
            if len(names) > 1:
                first, second = list(names.values())[:2]
                raise IssueError("latest_sheets_differ", line=_line_name(ref, line), first=first, second=second)

    def _plain_series(self, ref):
        """A date line and a value line as numeric points, oldest first."""
        rows = self._latest_rows(ref)
        line = ref.plain[0]
        if rows[-1][0] == (rows[-2][0] if len(rows) > 1 else None):
            raise IssueError("latest_ambiguous", date=rows[-1][0].date().isoformat(), count=sum(row[0] == rows[-1][0] for row in rows))
        points, from_text = [], []
        for date, i, position in rows:
            cell = ref.sheets[i].cell(*ref.at(position, line))
            number = as_number(cell.value)
            if number is not None:
                points.append(facts.Point(date, number, position, i))
                if isinstance(cell.value, str):
                    from_text.append(cell)
        # Anywhere in the series, not only at the newest: two tables on one sheet would otherwise be compared with each other.
        repeated = sorted(date for date, times in Counter(p.date for p in points).items() if times > 1)
        if repeated:
            raise IssueError("latest_duplicates", date=repeated[-1].date().isoformat(), count=len(repeated))
        points = self._with_period_names(points, ref)
        notes =[Issue("number_from_text", {"cell": f"{from_text[-1].parent.title}!{from_text[-1].coordinate}", "text": str(from_text[-1].value), "count": len(from_text)})] if from_text else []
        if len(points) >= 3 and points[-1].value == 0 and points[-2].value != 0:
            notes.append(Issue("latest_zero", {"date": points[-1].date.date().isoformat()}))
        newest = rows[-1]
        top_ws = ref.sheets[rows[0][1]]
        top = self._dated_rows(top_ws, ref)[0][1]
        header = self._header(top_ws, ref, top, line)
        found = Series(ref, points, (newest[1], newest[2]), line, header, notes)
        if not points:
            return found
        found.notes = self._row_notes(ref, points[-1].date) + self._period_issues(points, self._date_header(top_ws, ref, top), top_ws, ref) + notes
        newest_point = points[-1]
        fmt = self._cell_format(ref.sheets[newest_point.sheet].cell(*ref.at(newest_point.row, line)))
        found.rate = self._is_rate_series(top_ws, ref, top, line, header, fmt) or AMOUNT_HEADING.search(header) is not None
        found.scale = 100 if "%" in fmt else facts.AMOUNT if AMOUNT_HEADING.search(header) else 1
        found.per_cent_unit = "%" not in fmt and in_per_cent(top_ws)
        return found

    def _with_period_names(self, points, ref):
        """A financial year is read as a date so that it can be sorted, but a sentence must call it what the sheet does."""
        if not points or ref.date_parts or not self._fiscal_years(ref.sheets[points[0].sheet], ref):
            return points
        return [replace(point, label=f"{point.date.year - 1}-{point.date.year % 100:02d}") for point in points]

    def _summed_series(self, ref):
        """Rows added up per calendar day: the person asked for this by writing 합계( ) or 건수( ). Nothing is summed unasked."""
        spec = ref.values[0]
        if len(ref.values) != 1:
            raise IssueError("calc_columns", key=spec.word)
        groups, owners, from_text = {}, {}, []
        for i, ws in enumerate(ref.sheets):
            dated = dict((position, date) for date, position in self._dated_rows(ws, ref))
            current, current_position = None, None
            for position in range(ref.start, ref.last(ws) + 1):
                if position in dated:
                    current, current_position = dated[position], position
                elif ws.cell(*ref.at(position, ref.date)).value not in (None, ""):
                    current = None  # a row such as Total ends the group
                if current is None:
                    continue
                number = _aggregate_value(ws, ref, position, spec, from_text)
                if number is None:
                    continue
                day = dt.datetime.combine(current.date(), dt.time())
                if owners.setdefault(day, i) != i:
                    raise IssueError("latest_sheets_overlap", date=day.date().isoformat(), first=ref.sheets[owners[day]].title, second=ws.title)
                group = groups.setdefault(day, [0, 0, i, current_position, None])
                group[0] += number
                group[1] += 1
                if current.time() != dt.time():
                    moment = current.hour * 60 + current.minute
                    group[4] = max(group[4] or 0, moment)
        if not groups:
            raise IssueError("latest_no_dates", sheet=ref.sheets[0].title, column=_line_name(ref, ref.date))
        days = sorted(groups)
        points = [facts.Point(day, groups[day][0], groups[day][3], groups[day][2]) for day in days]
        top_ws = ref.sheets[groups[days[0]][2]]
        top = self._dated_rows(top_ws, ref)[0][1]
        names = [self._header(top_ws, ref, top, line) for line in spec.lines]
        name = _aggregate_name(spec, names)
        if spec.kind in ("mean", "rate"):
            points = [facts.Point(p.date, p.value / groups[p.date][1], p.row, p.sheet) for p in points]
        points = self._with_period_names(points, ref)
        newest = points[-1]
        found = Series(ref, points, (newest.sheet, newest.row), None if spec.kind in ("count", "rate") or spec.product else spec.lines[0], name, [])
        found.notes = self._period_issues(points, self._date_header(top_ws, ref, top), top_ws, ref) + _group_notes(days, groups)
        if len(points) >= 3 and points[-1].value == 0 and points[-2].value != 0:
            found.notes.append(Issue("latest_zero", {"date": points[-1].date.date().isoformat()}))
        found.percent = spec.kind == "rate"
        found.rate = spec.kind == "rate" or (spec.kind != "count" and self._is_rate_series(top_ws, ref, top, spec.lines[0], name, ""))
        found.scale = 100 if spec.kind == "rate" else 1
        found.per_cent_unit = spec.kind in ("sum", "mean") and in_per_cent(top_ws)
        usual = _usual_rows(days, groups)
        if usual is not None and groups[days[-1]][1] < usual and spec.kind in ("sum", "count"):
            found.cut_short = (groups[days[-1]][1], usual)
        if from_text:
            found.notes.append(Issue("number_from_text", {"cell": f"{from_text[-1].parent.title}!{from_text[-1].coordinate}", "text": str(from_text[-1].value), "count": len(from_text)}))
        return found

    @staticmethod
    def _is_rate_series(ws, ref, top, line, header, fmt):
        """True when the figures are rates, scores or changes already, so they move in points. Judged from the heading, the format, the sheet's name, and plain words above the table."""
        if "%" in fmt or RATE_HEADING.search(header) or RATE_TITLE.search(ws.title) or in_per_cent(ws):
            return True
        above =" ".join(str(ws.cell(*ref.at(position, other)).value or "") for position in range(top - 1, 0, -1) for other in (line, 1))
        return RATE_EXPLICIT.search(above) is not None

    def _series_words(self, found, lang):
        """(the report sentence, the reading facts, the facts for a model) for a series, with what a reader must know first put in the words themselves.

        Warnings stay in the app. Only the sentence reaches the slide, so a period that is not whole has to say so there.
        """
        points, show, label = found.points, found.show(self), found.label
        if found.cut_short:
            text = facts.cut_short_text(points, *found.cut_short, show, lang)
            head = f"{facts.point_text(points[-1].date, points, lang)} {label} {show(points[-1].value)}."
            fact = facts.Fact("read_cut_short", text, tuple(facts.NUMBER.findall(text)))
            return f"{head} {text}", [fact], [fact, facts.Fact("value", head, tuple(facts.NUMBER.findall(head)))]
        if len(points) >= 2 and any(issue.code == "latest_unfinished" for issue in found.notes):
            return self._unfinished_words(found, show, lang)
        if found.rate:
            note, read, pack = (facts.rate_sentence(points, label, show, lang, found.scale), facts.rate_readings(points, show, lang, found.scale),
                                facts.rate_pack(points, label, show, lang, found.scale))
        else:
            note, read, pack = facts.sentence(points, label, show, lang), facts.readings(points, show, lang), facts.fact_pack(points, label, show, lang)
        lead = self._leads(found, lang)
        if lead:
            fact = facts.Fact("read_lead", lead, tuple(facts.NUMBER.findall(lead)))
            return f"{lead} {note}", [fact] + list(read), [fact] + list(pack)
        return note, list(read), list(pack)

    def _unfinished_words(self, found, show, lang):
        """A period still running says what it holds so far, and the readings come from the periods that have ended. Three months of a year set against whole years would read as a collapse, however well warned."""
        points, label = found.points, found.label
        ended = points[:-1]
        now = points[-1]
        ko = lang == "ko"
        lead = self._leads(found, lang)
        head = f"{facts.point_text(now.date, points, lang)} {label} {show(now.value)}."
        last = f"{facts.point_text(ended[-1].date, points, lang)} {show(ended[-1].value)}"
        settled = f"끝난 기간 중 가장 최근은 {last}입니다." if ko else f"The latest period that has ended is {last}."
        read = [facts.Fact("read_lead", lead, tuple(facts.NUMBER.findall(lead))), facts.Fact("read_ended", settled, tuple(facts.NUMBER.findall(settled)))]
        read += facts.rate_readings(ended, show, lang, found.scale) if found.rate else facts.readings(ended, show, lang)
        value = facts.Fact("value", f"{facts.point_text(now.date, points, lang)} {label}: {show(now.value)}", tuple(facts.NUMBER.findall(show(now.value))))
        return f"{lead} {head}", read[1:], read + [value]

    def _stated_as_of(self):
        """The date a workbook says its data runs to, on any sheet: Data as at 30-Sep-2025. None when it does not say."""
        if not hasattr(self, "_as_at"):
            self._as_at = None
            for ws in self.wb.worksheets:
                for row in range(1, min(ws.max_row, AS_AT_ROWS) + 1):
                    for column in range(1, min(ws.max_column, AS_AT_COLUMNS) + 1):
                        cell = ws._cells.get((row, column))
                        text = cell.value if cell is not None and isinstance(cell.value, str) else ""
                        match = AS_AT.search(text) if len(text) <= AS_AT_LENGTH else None
                        day = _stated_day(text[match.end():]) if match else None
                        if day:
                            self._as_at = day
                            return day
        return self._as_at

    @staticmethod
    def _leads(found, lang):
        """A sentence for each warning that changes how the figure reads."""
        ko = lang == "ko"
        day = facts.point_text(found.points[-1].date, found.points, lang)
        parts = []
        for issue in found.notes:
            if issue.code == "latest_text_newer":
                parts.append(f"그 뒤 {issue.params['count']}개 기간은 '{issue.params['text']}'이라 {day} 값을 썼습니다." if ko else
                             f"The {issue.params['count']} periods after it say '{issue.params['text']}', so this uses {day}.")
            elif issue.code == "latest_unfinished":
                as_at = issue.params.get("as_at")
                if as_at:
                    parts.append(f"{day}은 자료 기준일 {as_at}에 아직 끝나지 않은 기간이라 값이 덜 쌓였을 수 있어, 다른 기간과 비교하지 않았습니다." if ko else
                                 f"{day} had not ended when the data was taken on {as_at}, so the value may be incomplete and is not compared with other periods.")
                else:
                    parts.append(f"{day}은 파일을 저장한 {issue.params['saved']}에 아직 끝나지 않은 기간이라 값이 덜 쌓였을 수 있어, 다른 기간과 비교하지 않았습니다." if ko else
                                 f"{day} had not ended when the file was saved on {issue.params['saved']}, so the value may be incomplete and is not compared with other periods.")
            elif issue.code == "latest_partial_day":
                parts.append(f"그날 마지막 기록이 {issue.params['last']}로 평소 {issue.params['usual']}보다 일러 값이 덜 쌓였을 수 있습니다." if ko else
                             f"The last entry that day is at {issue.params['last']}, while days usually run to {issue.params['usual']}, so the value may be incomplete.")
            elif issue.code == "latest_zero":
                parts.append("최신 값이 0입니다. 빈 기간에 미리 채운 0일 수 있습니다." if ko else "The newest value is 0, which may have been filled in ahead for an empty period.")
        return " ".join(parts)

    def _calc_text(self, key):
        """Work out {{전일대비:시트!A,B}} and friends. Returns (text or None, source, issues)."""
        match = CALC.fullmatch(key.translate(SMART_QUOTES))
        word, count = match.group(1), match.group(2)
        name = CALC_NAMES.get(word, word.lower())
        ref = self._parse_ref(match.group(3), key)
        if count and name not in ("avg", "peak", "low"):
            raise IssueError("latest_bad", key=key)
        if name in ("읽는법", "reading", "해설", "note") and self._is_comparison(ref):
            return self._comparison_text(ref, name)
        if ref.spec_count != 1:
            raise IssueError("calc_columns", key=key)
        found = self._series(ref)
        points = found.points
        if not points:
            raise IssueError("latest_no_dates", sheet=ref.sheets[0].title, column=_line_name(ref, ref.date))
        issues = found.newest_issues(self)
        if any(i.is_error for i in issues):
            return None, found.newest_cell(), issues
        newest = points[-1]
        issues = found.notes + issues
        when = newest.date.date().isoformat()
        cell = found.cell_text(newest)
        if name in ("peak", "low"):
            top = facts.extreme(points, int(count or 30), highest=name == "peak")
            return self._date_cell(ref, top.sheet, top.row)[1], f"{top.date.date().isoformat()} · {found.cell_text(top)}", issues
        if name in ("읽는법", "reading"):
            note, read, _pack = self._series_words(found, "ko" if name == "읽는법" else "en")
            read = [fact for fact in read if fact.text not in note] or read  # a rate's sentence already holds its change, and the slide shows both
            if not read:
                return None, f"{when} · {cell}", issues + [Issue("calc_no_base", {"date": when})]
            return " ".join(fact.text for fact in read), f"{when} · {cell}", issues
        if name in ("해설", "note"):
            return self._series_words(found, "ko" if name == "해설" else "en")[0], f"{when} · {cell}", issues
        if name == "avg":
            size = int(count or 7)
            base = facts.mean_before(points, size) if size >= 1 else None
            against = f"{points[-size - 1].date.date().isoformat()} ~ {points[-2].date.date().isoformat()}" if base is not None else ""
        else:
            before = facts.previous(points) if name == "prev" else facts.week_ago(points)
            base = before.value if before else None
            against = before.date.date().isoformat() if before else ""
        if found.rate and base is not None:
            lang = "ko" if HANGUL.search(word) else "en"
            return facts.points_text(newest.value - base, lang, found.scale), f"{when} ↔ {against} · {cell}", issues
        change = facts.ratio(newest.value, base) if base is not None else None
        if change is None:
            return None, f"{when} · {cell}", issues + [Issue("calc_no_base", {"date": when})]
        return facts.percent(change), f"{when} ↔ {against} · {cell}", issues

    def _is_comparison(self, ref):
        """No dates at all in the first line, and something to compare: a table of named items rather than a series."""
        return len(ref.specs) in (1, 2) and not self._all_rows(ref)

    def _comparison(self, ref):
        """The named items of a table without dates, with any stated total taken out. Returns a Comparison."""
        key = ("comparison", ref)  # the same cells can be read both ways, as dates or as named items
        if key not in self._made:
            self._made[key] = self._read_comparison(ref)
        return self._made[key]

    def _read_comparison(self, ref):
        if len(ref.sheets) != 1:
            raise IssueError("latest_bad", key=f"{_sheet_names(ref)}!")
        ws = ref.sheets[0]
        labels, specs = ref.label_lines, ref.specs
        lower = min(spec.lines[0] if isinstance(spec, Aggregate) else spec for spec in specs)
        items, counts, order, totals, unlabeled, from_text, first = {}, {}, [], [], 0, [], None
        heading = self._heading_row(ws, ref, labels, specs)
        outer = _rows_of_an_outer_level(ws, ref, labels)
        chosen = {line for spec in specs if isinstance(spec, Aggregate) for line, _word in spec.filters}
        indented = _rows_indented_under_others(ws, ref, labels, [spec.lines[0] if isinstance(spec, Aggregate) else spec for spec in specs])
        for position in range(ref.start, ref.last(ws) + 1):
            if position == heading or position in indented:
                continue
            texts_so_far = len(from_text)
            numbers = [_aggregate_value(ws, ref, position, spec, from_text) if isinstance(spec, Aggregate) else _plain_value(ws, ref, position, spec, from_text) for spec in specs]
            if all(number is None for number in numbers):
                continue
            if first is None and _years_only(ws, ref, position, specs):
                del from_text[texts_so_far:]
                continue  # 2023, 2024 heading the value columns, not an item
            first = first or position
            kind = _total_kind(ws, ref, position, lower, chosen)
            if kind == "total":
                totals.append(numbers)
                continue
            if kind == "subtotal" or position in outer:
                continue
            label = self._label(ws, ref, position, labels)
            if not label:
                unlabeled += 1
                continue
            if any(isinstance(spec, Aggregate) for spec in specs) and label in items:
                items[label] = [(a or 0) + (b or 0) if a is not None or b is not None else None for a, b in zip(items[label], numbers)]
                counts[label] = [c + (b is not None) for c, b in zip(counts[label], numbers)]
            else:
                key = label if label not in items else (label, position)
                items[key] = numbers
                counts[key] = [int(b is not None) for b in numbers]
                order.append(key)
        order = [key for key in order if items[key][-1] is not None]
        for key in order:  # a mean or a rate is worked out once every row of the item is in
            items[key] = [value / count if value is not None and isinstance(spec, Aggregate) and spec.kind in ("mean", "rate") else value
                          for value, count, spec in zip(items[key], counts[key], specs)]
        if len(order) < 2:
            raise IssueError("cross_no_items", sheet=ws.title, column=_line_name(ref, labels[0]))
        rows = [facts.Item(key if isinstance(key, str) else key[0], items[key][-1], items[key][0] if len(specs) == 2 else None) for key in order]
        first_line = specs[-1].lines[0] if isinstance(specs[-1], Aggregate) else specs[-1]
        header = self._header(ws, ref, first or ref.start, first_line)
        fmt_cell = ws.cell(*ref.at(first or ref.start, first_line))
        fmt = self._cell_format(fmt_cell)
        last_spec = specs[-1]
        is_rate_kind = isinstance(last_spec, Aggregate) and last_spec.kind == "rate"
        if isinstance(last_spec, Aggregate) and (last_spec.kind in ("mean", "rate", "count") or last_spec.filters):
            header = _aggregate_name(last_spec, [header])
        if is_rate_kind:
            fmt = "0.0%"
        issues = []
        if len(specs) == 2:
            first_name = self._header(ws, ref, first or ref.start, specs[0].lines[0] if isinstance(specs[0], Aggregate) else specs[0])
            second_name = self._header(ws, ref, first or ref.start, first_line)
            if _reads_backwards(first_name, second_name):
                issues.append(Issue("cross_backwards", {"first": first_name, "second": second_name}))
        if from_text:
            issues.append(Issue("number_from_text", {"cell": f"{ws.title}!{from_text[-1].coordinate}", "text": str(from_text[-1].value), "count": len(from_text)}))
        if unlabeled:
            issues.append(Issue("cross_unlabeled", {"count": unlabeled}))
        stated = totals[0] if totals else None
        if not stated and len(rows) >= 4:  # 합계 under a name this list does not know: the first or last item is the sum of all the others
            largest = max(rows, key=lambda item: item.value)
            if largest in (rows[0], rows[-1]) and largest.value > 0 and abs(sum(item.value for item in rows) - 2 * largest.value) <= largest.value * TOTAL_TOLERANCE:
                rows = [item for item in rows if item is not largest]
                stated = [largest.before, largest.value] if len(specs) == 2 else [largest.value]
                issues.append(Issue("cross_total_found", {"label": largest.label}))
        per_cent = not is_rate_kind and in_per_cent(ws)
        rate = is_rate_kind or per_cent or (isinstance(last_spec, Aggregate) and last_spec.kind == "mean") or facts.is_rate(f"{header} {ws.title}", fmt)
        levels = facts.top_level(rows)
        if levels:  # 자산 and its parts, numbered by level: only the top level can stand side by side, and none of it adds up to a whole
            rows = levels
            issues.append(Issue("cross_outline", {"count": len(rows)}))
        nested = facts.without_parts(rows) if not levels else None
        if nested:  # a row that is the sum of the rows under it: they are its parts, not its rivals
            rows = nested
            issues.append(Issue("cross_nested", {"count": len(rows)}))
        additive = not levels and not per_cent and facts.is_additive(f"{header} {ws.title}", fmt, [item.value for item in rows])
        total = before_total = None
        added = sum(item.value for item in rows)
        has_total = False
        if additive:
            total = added
            matching = next((line for line in totals if line[-1] is not None and abs(line[-1] - added) <= abs(line[-1]) * TOTAL_TOLERANCE), None)
            stated = matching or stated
            if stated and stated[-1] is not None:
                if matching or abs(stated[-1] - added) <= abs(stated[-1]) * TOTAL_TOLERANCE:
                    total = stated[-1]
                    has_total = True
                else:
                    issues.append(Issue("cross_total_differs", {"stated": facts.plain_number(stated[-1]), "added": facts.plain_number(added)}))
                    total = None
            else:
                issues.append(Issue("cross_no_total", {"added": facts.plain_number(added)}))
            if len(specs) == 2:
                before_total = sum(item.before for item in rows if item.before is not None)
                if stated and stated[0] is not None and total is not None:
                    before_total = stated[0]
        # A rate shown as 4 and 6 cannot explain +1.8 points, so a rate keeps its digits unless the cell is a real percentage.
        plain = fmt in ("General", "@") or (rate and "%" not in fmt)
        show = facts.plain_number if plain else (lambda value: format_value(value, fmt, self.wb.epoch)[0])
        if per_cent and "%" not in fmt:
            show = _with_per_cent(show)
        source = f"{ws.title}!{','.join(_line_name(ref, line) for line in labels)}," + ",".join(spec.word if isinstance(spec, Aggregate) else _line_name(ref, spec) for spec in specs) + f" · {len(rows)}"
        label_header = self._header(ws, ref, first or ref.start, labels[-1])
        label_header = "" if label_header == _line_name(ref, labels[-1]) else label_header  # no heading: blank, not a column letter
        return Comparison(rows, header, total, before_total, additive, show, issues, source, label_header, has_total, rate, 100 if "%" in fmt else 1)

    def _heading_row(self, ws, ref, labels, specs):
        """When rows are counted, the row of headings would count as one more item: a name that appears once among names that repeat."""
        if not any(isinstance(spec, Aggregate) and spec.kind in ("count", "rate") for spec in specs) or ref.start > 1 or ref.end:
            return None
        names = {}
        for position in range(1, ref.last(ws) + 1):
            label = self._label(ws, ref, position, labels)
            if label:
                names.setdefault(label, []).append(position)
        if not names:
            return None
        first = next(iter(names))
        return names[first][0] if len(names[first]) == 1 and max(len(places) for places in names.values()) > 1 else None

    def _label(self, ws, ref, position, lines):
        """The most specific name a row has: the last filled cell of its label columns."""
        for line in reversed(lines):
            value = _merged_value(ws, *ref.at(position, line))
            # A bare number is a figure or a code, not a name a reader could use.
            if value not in (None, "") and as_number(value) is None:
                return _clean_label(value) if isinstance(value, str) else self.read(ws, *ref.at(position, line))[1]
        return ""

    def _comparison_text(self, ref, name):
        found = self._comparison(ref)
        lang = "ko" if name in ("읽는법", "해설") else "en"
        if name in ("읽는법", "reading"):
            text = " ".join(fact.text for fact in facts.item_readings(found.items, found.show, lang, found.total, found.additive, found.before_total, limit=5, stated=found.stated, rate=found.rate, scale=found.scale))
        else:
            text = facts.item_sentence(found.items, found.label, found.show, lang, found.total, found.additive, found.before_total, found.stated, found.rate, found.scale)
        return text, found.source, found.issues

    def _rank_text(self, key):
        """{{순위2:시트!A,B}} the name, {{순위값2:...}} the value, {{구성비2:...}} the share of the second largest."""
        match = RANK.fullmatch(key.translate(SMART_QUOTES))
        word, wanted = RANK_NAMES[match.group(1).lower()], int(match.group(2) or 1)
        found = self._comparison_for_rank(match.group(3), key)
        ordered = found.ranked()
        if wanted < 1 or wanted > len(ordered):
            raise IssueError("latest_fewer", wanted=wanted, found=len(ordered))
        item = ordered[wanted - 1]
        if word == "name":
            return item.label, found.source, found.issues
        if word == "value":
            return found.show(item.value), found.source, found.issues
        if not found.additive or not found.total:
            raise IssueError("cross_not_additive", header=found.header)
        return format_value(item.value / found.total, "0.0%")[0], found.source, found.issues

    def _comparison_for_rank(self, body, key):
        ref = self._parse_ref(body, key)
        if len(ref.specs) != 1:
            raise IssueError("calc_columns", key=key)
        return self._comparison(ref)

    def _rank_range(self, key):
        """{{표:순위5:시트!A,B}} names and values, largest first. {{표:구성비5:...}} adds each one's share of the total."""
        match = RANK.fullmatch(key.translate(SMART_QUOTES))
        word, count = RANK_NAMES[match.group(1).lower()], int(match.group(2) or 5)
        found = self._comparison_for_rank(match.group(3), key)
        with_share = word == "share"
        if with_share and (not found.additive or not found.total):
            raise IssueError("cross_not_additive", header=found.header)
        ordered = found.ranked()
        picked = ordered[:count]
        label_header = found.label_header
        if self.thousands and not label_header:
            label_header = ITEM_WORD[self._table_lang]
        names = [label_header, self._report_heading(found.header)] + ([SHARE_HEADER] if with_share else [])
        table = [[(text, text, "General") for text in names]]
        for item in picked:
            line = [(item.label, item.label, "General"), (item.value, found.show(item.value), "General")]
            if with_share:
                share = item.value / found.total
                line.append((share, format_value(share, "0.0%")[0], "0.0%"))
            table.append(line)
        issues = list(found.issues)
        if len(picked) != count:
            issues.append(Issue("latest_fewer", {"wanted": count, "found": len(picked)}))
        return table, found.source, issues

    def insight_request(self, key, template_text=""):
        """First stage for {{통찰:시트!A,B}}. Returns (InsightRequest or None, source, issues)."""
        ref = self._parse_ref(INSIGHT.fullmatch(key.translate(SMART_QUOTES)).group(1), key)
        words = [template_text] + [sheet.title for sheet in self.wb.worksheets]
        if self._is_comparison(ref):
            found = self._comparison(ref)
            lang = facts.detect_language(words + [found.header])
            pack = _with_sheet(facts.item_pack(found.items, found.show, lang, found.total, found.additive, found.before_total, found.stated, found.rate, found.scale), ref, lang, found.label)
            fallback = facts.item_sentence(found.items, found.label, found.show, lang, found.total, found.additive, found.before_total, found.stated, found.rate, found.scale)
            return InsightRequest(key, lang, pack, facts.fingerprint(pack, lang), fallback, found.source), found.source, found.issues
        if ref.spec_count != 1:
            raise IssueError("calc_columns", key=key)
        found = self._series(ref)
        points = found.points
        if not points:
            raise IssueError("latest_no_dates", sheet=ref.sheets[0].title, column=_line_name(ref, ref.date))
        issues = found.newest_issues(self)
        if any(i.is_error for i in issues):
            return None, found.newest_cell(), issues
        newest = points[-1]
        lang = facts.detect_language(words + [found.header])
        note, read, pack = self._series_words(found, lang)
        pack = _with_sheet(pack, ref, lang, found.label, _period_of(pack, found.label))
        source = f"{newest.date.date().isoformat()} · {found.cell_text(newest)}"
        fallback = " ".join([note] + [fact.text for fact in read if fact.text not in note])  # a rate's sentence already holds its readings
        request = InsightRequest(key, lang, pack, facts.fingerprint(pack, lang), fallback, source)
        return request, source, found.notes + issues

    def text_value(self, key):
        """Resolve a text placeholder. Returns (text or None, source, issues)."""
        if is_calc_key(key):
            return self._calc_text(key)
        if is_rank_key(key):
            return self._rank_text(key)
        notes = []
        if is_latest_key(key):
            ref, count = self._latest_request(key)
            if ref.aggregated:
                return self._latest_sum_text(ref, count, key)
            if count > 1:
                raise IssueError("range_for_text", key=key, source=key)
            if ref.date_parts and not ref.plain:
                return self._latest_date_text(ref)
            ws, row, col, source, notes = self._latest_text(ref, key)
        else:
            ws, (col, row, max_col, max_row), source = self.locate(key)
            if (col, row) != (max_col, max_row):
                raise IssueError("range_for_text", key=key, source=source)
        _raw, text, _fmt, issues = self.read(ws, row, col)
        issues = notes + issues
        if any(i.is_error for i in issues):
            return None, source, issues
        if text == "":
            issues = issues + [Issue("empty_cell", {"cell": source})]
        return text, source, issues

    def range_cells(self, key):
        """Resolve a range. Returns (rows of (raw, text, format), source, issues). A latest or ranked key's first row is its header."""
        self._table_lang = "ko" if HANGUL.search(key.partition(":")[0]) else "en"  # the words a table adds itself follow the language it was asked for in, as its sentences do
        rows, source, issues = self._range_cells(key)
        return (_with_commas(rows) if self.thousands else rows), source, issues

    def _range_cells(self, key):
        if is_rank_key(key):
            return self._rank_range(key)
        if is_latest_key(key):
            ref, count = self._latest_request(key)
            return self._latest_sum_range(ref, count) if ref.aggregated else self._latest_range(ref, count)
        ws, (col, row, max_col, max_row), source = self.locate(key)
        rows, issues = [], []
        for r in range(row, max_row + 1):
            line = []
            for c in range(col, max_col + 1):
                raw, text, fmt, cell_issues = self.read(ws, r, c)
                if raw in (None, ""):  # 지역 merged over two rows, with the range starting on the second: a person sees the word there
                    origin = _merge_origin(ws, r, c)
                    if origin and not (row <= origin[0] <= max_row and col <= origin[1] <= max_col):
                        raw, text, fmt, cell_issues = self.read(ws, *origin)
                line.append((raw, text, fmt))
                issues.extend(cell_issues)
            rows.append(line)
        return rows, source, issues


# ---------------------------------------------------------------- writing text into slide XML

def _clean(text):
    return ILLEGAL_XML.sub("", text.replace("\r\n", "\n").replace("\r", "\n"))


def _write_run(run, text):
    """Set a run's text. Line breaks in the value become real line breaks with the same formatting."""
    lines = _clean(text).split("\n")
    run.find(A_T).text = lines[0]
    anchor = run
    for line in lines[1:]:
        br = etree.Element(A_BR)
        rpr = run.find(A_RPR)
        if rpr is not None:
            br.append(copy.deepcopy(rpr))
        anchor.addnext(br)
        extra = copy.deepcopy(run)
        extra.find(A_T).text = line
        br.addnext(extra)
        anchor = extra


def _run_groups(paragraph):
    """Yield lists of adjacent runs. Line breaks and fields split a paragraph into separate groups."""
    group = []
    for child in paragraph:
        if child.tag == A_R and child.find(A_T) is not None:
            group.append(child)
        elif group:
            yield group
            group = []
    if group:
        yield group


def replace_in_paragraph(paragraph, resolve):
    """Replace placeholders in a paragraph even when PowerPoint split them across runs.

    resolve(key) returns the new text or None to leave that placeholder as it is. The new text takes
    the formatting of the run where the placeholder starts.
    """
    for runs in list(_run_groups(paragraph)):
        texts = [r.find(A_T).text or "" for r in runs]
        joined = "".join(texts)
        replaced = {}
        emptied = set()
        for match in reversed(list(PLACEHOLDER.finditer(joined))):
            new = resolve(match.group(1))
            if new is None:
                continue
            start, end = match.span()
            pos = 0
            for i, text in enumerate(texts):
                run_start, run_end = pos, pos + len(text)
                pos = run_end
                if run_end <= start or run_start >= end:
                    continue
                left = text[:max(start - run_start, 0)]
                right = text[max(end - run_start, 0):] if end < run_end else ""
                middle = new if run_start <= start < run_end else ""
                texts[i] = left + middle + right
                replaced[i] = True
                if not texts[i]:
                    emptied.add(i)
        for i in replaced:
            if i in emptied:
                paragraph.remove(runs[i])
            else:
                _write_run(runs[i], texts[i])


def _iter_shapes(shapes, in_group=False):
    """Yield (shape, in_group) for every shape, looking inside groups."""
    for shape in shapes:
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from _iter_shapes(shape.shapes, True)
        else:
            yield shape, in_group


def _alt_text_element(shape):
    """The element holding a shape's alt text (PowerPoint's Alt Text pane writes the descr attribute)."""
    found = shape._element.xpath(".//p:cNvPr")
    return found[0] if found else None


def set_cell_text(tc, text):
    """Replace a table cell's text, keeping the formatting of its first run or its end-of-paragraph mark."""
    body = tc.txBody
    paragraphs = body.findall(A_P)
    paragraph = paragraphs[0]
    for extra in paragraphs[1:]:
        body.remove(extra)
    runs = paragraph.findall(A_R)
    for child in list(paragraph):
        if child.tag in (A_BR, A + "fld") or (child.tag == A_R and child is not (runs[0] if runs else None)):
            paragraph.remove(child)
    if runs:
        run = runs[0]
    else:
        run = etree.Element(A_R)
        end = paragraph.find(A + "endParaRPr")
        if end is not None:
            rpr = copy.deepcopy(end)
            rpr.tag = A_RPR
            run.append(rpr)
            end.addprevious(run)
        else:
            paragraph.append(run)
        etree.SubElement(run, A_T)
    _write_run(run, text)


# ---------------------------------------------------------------- tables and charts

def _renew_row_id(tr, used):
    for node in tr.iter():
        if etree.QName(node).localname == "rowId" and node.get("val"):
            value = max(used) + 1 if used else 10000
            used.add(value)
            node.set("val", str(value))


def _resize_rows(tbl, wanted):
    """Grow or shrink a table by copying a body row, so new rows look like existing ones."""
    trs = tbl.tr_lst
    total = tbl.tblPr is not None and tbl.tblPr.get("lastRow") == "1"
    used = {int(node.get("val")) for tr in trs for node in tr.iter()
            if etree.QName(node).localname == "rowId" and (node.get("val") or "").isdigit()}
    while len(trs) < wanted:
        source = trs[-2] if total and len(trs) >= 3 else trs[-1]
        new = copy.deepcopy(source)
        _renew_row_id(new, used)
        source.addnext(new)
        trs = tbl.tr_lst
    while len(trs) > wanted:
        tbl.remove(trs[-2] if total and wanted >= 2 and len(trs) >= 2 else trs[-1])
        trs = tbl.tr_lst


def _cell_height(tc, width):
    """Rough height of a cell's wrapped text, in EMU. Full-width characters take one em, others about half."""
    pr = tc.find(A + "tcPr")

    def margin(name, default):
        return int(pr.get(name, default)) if pr is not None else default

    room = max(width - margin("marL", 91440) - margin("marR", 91440), 1)
    sizes = [node.get("sz") for node in tc.iter(A_RPR, A + "endParaRPr") if node.get("sz")]
    size = int(sizes[0]) * 127 if sizes else 228600  # hundredths of a point to EMU; PowerPoint's default is 18pt
    lines = 0
    for paragraph in tc.txBody.findall(A_P):
        parts = [""]
        for child in paragraph:
            if child.tag == A_BR:
                parts.append("")
            elif child.tag == A_R:
                parts[-1] += child.findtext(A_T) or ""
        for part in parts:
            em = sum(1 if unicodedata.east_asian_width(c) in "WF" else 0.55 for c in part)
            lines += max(math.ceil(em * size / room), 1)
    return lines * size * 1.2 + margin("marT", 45720) + margin("marB", 45720)


def _drawn_height(tbl):
    """Table height once PowerPoint has grown each row to fit its text. The file only stores the minimum."""
    widths = [col.w for col in tbl.tblGrid.gridCol_lst]
    total = 0
    for tr in tbl.tr_lst:
        height = tr.h
        for i, tc in enumerate(tr.tc_lst):
            if not (tc.get("hMerge") or tc.get("vMerge") or tc.get("rowSpan")):
                height = max(height, _cell_height(tc, sum(widths[i:i + int(tc.get("gridSpan") or 1)])))
        total += height
    return total


def _can_resize_merged(tbl, wanted):
    """Rows come and go at the bottom, so cells merged down in a heading stay whole: a two-row heading is the usual company table. Merged in the body, a copied or removed row would tear a merge apart."""
    trs = tbl.tr_lst
    merged = [place for place, tr in enumerate(trs) if any(tc.get("rowSpan") or tc.get("vMerge") for tc in tr.tc_lst)]
    if not merged:
        return True
    heading = merged[-1] + 1
    total = 1 if tbl.tblPr is not None and tbl.tblPr.get("lastRow") == "1" else 0
    return merged == list(range(heading)) and len(trs) - heading - total >= 1 and wanted - heading - total >= 1


def fill_table(shape, rows, slide_height, in_group, skip=0):
    """Fill a table from its top-left cell, leaving the first skip rows as they are. Returns issues.

    On an error issue the table is left untouched.
    """
    tbl = shape.table._tbl
    n_cols = len(tbl.tblGrid.gridCol_lst)
    if len(rows[0]) != n_cols:
        return [Issue("column_mismatch", {"range": len(rows[0]), "table": n_cols})]
    if len(rows) != len(tbl.tr_lst) and not _can_resize_merged(tbl, len(rows)):
        return [Issue("merged_resize", {})]
    bottom_before = shape.top + _drawn_height(tbl) if not in_group else 0
    _resize_rows(tbl, len(rows))
    issues = []
    for tr, line in zip(tbl.tr_lst[skip:], rows[skip:]):
        for tc, (_raw, text, _fmt) in zip(tr.tc_lst, line):
            if tc.get("hMerge") or tc.get("vMerge"):
                if text:
                    issues.append(Issue("merged_value_lost", {"value": text}))
                continue
            set_cell_text(tc, text or "")
    if not in_group:
        shape.height = sum(tr.h for tr in tbl.tr_lst)
        if shape.top + _drawn_height(tbl) > slide_height:
            issues.append(Issue("table_overflow", {"rows": len(rows)}))
        covered = _shape_now_covered(shape, bottom_before, shape.top + _drawn_height(tbl))
        if covered:
            issues.append(Issue("table_covers", {"rows": len(rows), "name": covered}))
    return issues


def _shape_now_covered(table, bottom_before, bottom_after):
    """The name of a shape that was clear below a table and is under it now that it has grown. A table that grew last month is the usual cause, and the slide looks wrong with no error anywhere."""
    for other in table._parent:
        if other.shape_id == table.shape_id or None in (other.top, other.left, other.width):
            continue
        beside = other.left < table.left + table.width and other.left + other.width > table.left
        if beside and bottom_before <= other.top < bottom_after:
            return other.name
    return ""


def _chart_number_format(rows):
    """Number format for chart values. The XML writer in python-pptx does not escape it, so unsafe ones are dropped."""
    for line in rows[1:]:
        for raw, _text, fmt in line[1:]:
            if isinstance(raw, (int, float)) and not isinstance(raw, bool):
                return "General" if set(fmt) & set("&<>") else fmt
    return "General"


def _number_in_first_row(rows):
    """A number where a series name belongs usually means the range has no header row, so that row is lost.

    Plain years are a common real header and do not count. Returns the number's text, or None.
    """
    for raw, text, fmt in rows[0][1:]:
        if isinstance(raw, bool) or not isinstance(raw, (int, float)):
            continue
        if not (float(raw).is_integer() and 1900 <= raw <= 2100 and fmt in ("General", "0")):
            return text
    return None


def fill_chart(chart, rows):
    """Fill a chart from a range: first row holds series names, first column holds categories. Returns issues."""
    if chart.chart_type in BUBBLE_CHARTS:
        return [Issue("chart_unsupported", {})]
    if len(rows) < 2 or len(rows[0]) < 2:
        return [Issue("chart_too_small", {})]
    issues = []
    lost = _number_in_first_row(rows)
    if lost is not None:
        issues.append(Issue("chart_header_number", {"value": lost}))

    def number(cell):
        raw, text, _fmt = cell
        if isinstance(raw, bool) or not isinstance(raw, (int, float)):
            if raw is not None:
                issues.append(Issue("chart_not_number", {"value": text}))
            return None
        return raw

    number_format = _chart_number_format(rows)
    if chart.chart_type in XY_CHARTS:
        data = XyChartData(number_format=number_format)
        for c in range(1, len(rows[0])):
            series = data.add_series(rows[0][c][1] or "")
            for line in rows[1:]:
                x, y = number(line[0]), number(line[c])
                if x is not None and y is not None:
                    series.add_data_point(x, y)
    else:
        data = CategoryChartData(number_format=number_format)
        date_axis = bool(chart._chartSpace.xpath(".//c:dateAx"))
        categories = [line[0] for line in rows[1:]]
        if date_axis and all(isinstance(c[0], (dt.date, dt.datetime)) for c in categories):
            data.categories = [c[0].date() if isinstance(c[0], dt.datetime) else c[0] for c in categories]
        else:
            data.categories = [c[1] or "" for c in categories]
        for c in range(1, len(rows[0])):
            data.add_series(rows[0][c][1] or "", [number(line[c]) for line in rows[1:]])
    chart.replace_data(data)
    return issues


# ---------------------------------------------------------------- the run

class _Run:
    def __init__(self, source, prs, insights=None):
        self.source, self.prs = source, prs
        self.insights = insights or {}
        self.requests = {}
        self.template_text = None
        self.items = {}
        self.done_parts = set()
        self.filled_paragraphs = set()  # table cells we just wrote: their data must not be scanned for placeholders
        self.words_of_template()  # now, while they are its own words and not what has been filled in

    def item(self, kind, key, where):
        item = self.items.setdefault((kind, key.casefold()), Item(kind, key))
        if where not in item.wheres:
            item.wheres.append(where)
        return item

    @staticmethod
    def add_issues(item, issues):
        for issue in issues:
            if issue not in item.issues:
                item.issues.append(issue)

    def resolve_text(self, where):
        def resolve(key):
            item = self.item("text", key, where)
            if MARKER.fullmatch("{{" + key + "}}"):
                self.add_issues(item, [Issue("marker_in_text", {"key": key})])
                return None
            if is_insight_key(key):
                return self.resolve_insight(key, item)
            try:
                text, source, issues = self.source.text_value(key)
            except IssueError as exc:
                self.add_issues(item, [exc.issue])
                return None
            item.source = source
            item.value = text or ""
            self.add_issues(item, issues)
            return text
        return resolve

    def words_of_template(self):
        """The template's own words, for the language router. Placeholders are left out: they are mostly addresses."""
        if self.template_text is None:
            parts = [node.text or "" for slide in self.prs.slides for node in slide.part._element.iter(A_T)]
            self.template_text = PLACEHOLDER.sub(" ", "".join(parts))
        return self.template_text

    def resolve_insight(self, key, item):
        """Use the sentence the person chose for exactly these figures, or else the calculated sentence."""
        try:
            request, source, issues = self.source.insight_request(key, self.words_of_template())
        except IssueError as exc:
            self.add_issues(item, [exc.issue])
            return None
        item.source = source
        if request is None:
            self.add_issues(item, issues)
            return None
        self.requests[key] = request
        chosen = self.insights.get(key.casefold())
        if chosen and chosen.get("fingerprint") == request.fingerprint and chosen.get("text"):
            text = chosen["text"]
        else:
            text = request.fallback
            issues = [Issue("insight_stale" if chosen else "insight_not_chosen", {})] + issues
        item.value = text
        self.add_issues(item, issues)
        return text

    def fill_part_text(self, part, where):
        self.done_parts.add(part)
        resolve = self.resolve_text(where)
        for paragraph in part._element.iter(A_P):
            if paragraph not in self.filled_paragraphs:
                replace_in_paragraph(paragraph, resolve)

    def fill_marked_shape(self, shape, where, in_group):
        element = _alt_text_element(shape)
        descr = element.get("descr") if element is not None else None
        match = MARKER.search(descr or "")
        if not match:
            return
        kind = {"표": "table", "차트": "chart"}.get(match.group(1).lower(), match.group(1).lower())
        item = self.item(kind, match.group(2), where)
        fits = shape.has_table if kind == "table" and hasattr(shape, "has_table") else (
            shape.has_chart if kind == "chart" and hasattr(shape, "has_chart") else False)
        if not fits:
            self.add_issues(item, [Issue("marker_wrong_shape", {"kind": kind, "shape": shape.name})])
            return
        try:
            rows, source, issues = self.source.range_cells(match.group(2))
        except IssueError as exc:
            self.add_issues(item, [exc.issue])
            return
        item.source = source
        item.value = f"{len(rows) - 1} × {len(rows[0])}"  # rows of data: the first is the heading
        if kind == "chart":
            issues = [_chart_error_as_warning(i) for i in issues]
        else:
            issues = [i for i in issues if i.code != "error_value"] + [
                Issue("table_error_value", i.params) for i in issues if i.code == "error_value"]
        self.add_issues(item, issues)
        if any(i.is_error for i in issues):
            return
        # A latest range starts below the sheet's header, so a header the template already has is kept.
        skip = 0
        if kind == "table" and (is_latest_key(match.group(2)) or is_rank_key(match.group(2))):
            skip = int(any(t.text and t.text.strip() for t in shape.table._tbl.tr_lst[0].iter(A_T)))
        result = fill_chart(shape.chart, rows) if kind == "chart" else fill_table(
            shape, rows, self.prs.slide_height, in_group, skip)
        self.add_issues(item, result)
        if any(i.is_error for i in result):
            return
        if kind == "table":
            for tr in shape.table._tbl.tr_lst[skip:]:
                self.filled_paragraphs.update(tr.iter(A_P))
        del element.attrib["descr"]  # the marker has done its job; it should not ship in the finished report

    def walk_slide(self, slide, number):
        where = ("slide", number)
        charts = []
        for shape, in_group in _iter_shapes(slide.shapes):
            self.fill_marked_shape(shape, where, in_group)
            if getattr(shape, "has_chart", False):
                charts.append(shape.chart.part)
        self.fill_part_text(slide.part, where)
        for part in charts:
            self.fill_part_text(part, where)
        if slide.has_notes_slide:
            self.fill_part_text(slide.notes_slide.part, ("notes", number))

    def walk(self):
        for number, slide in enumerate(self.prs.slides, start=1):
            self.walk_slide(slide, number)
        for m, master in enumerate(self.prs.slide_masters, start=1):
            self.fill_part_text(master.part, ("master", m))
            for layout in master.slide_layouts:
                self.fill_part_text(layout.part, ("layout", layout.name))
        self.flag_unreachable()

    def flag_unreachable(self):
        """Placeholders in parts we do not edit, such as SmartArt, are reported rather than silently left."""
        for part in self.prs.part.package.iter_parts():
            name = str(part.partname)
            if part in self.done_parts or not name.startswith("/ppt/") or not name.endswith(".xml"):
                continue
            if isinstance(part, XmlPart) and not name.startswith(("/ppt/diagrams/", "/ppt/charts/")):
                continue
            if name.startswith("/ppt/diagrams/drawing"):  # PowerPoint's drawn copy of the SmartArt beside it, not a second place to fix
                continue
            if not PLACEHOLDER.search(part.blob.decode("utf-8", "ignore")):
                continue
            item = self.item("unsupported", name, ("part", name))
            self.add_issues(item, [Issue("unsupported_location", {"part": name, "slide": self.slide_holding(part)})])

    def slide_holding(self, part):
        """The number of the slide that uses a part, so a person can find it."""
        for number, slide in enumerate(self.prs.slides, start=1):
            if any(not rel.is_external and rel.target_part is part for rel in slide.part.rels.values()):
                return number
        return "?"


def _chart_error_as_warning(issue):
    """#N/A is a common way to leave a gap in a chart, so an error cell there is a warning, not a blocker."""
    return Issue("chart_not_number", {"value": issue.params["error"]}) if issue.code == "error_value" else issue


def open_template(path):
    _check_file(path, (".pptx",), "template")
    try:
        return Presentation(path)
    except PermissionError:
        raise SlideFillError("file_locked", path=path)
    except Exception as exc:
        raise SlideFillError("file_unreadable", path=path, detail=str(exc))


FILE_MARK = re.compile(r"@(\d+)@")
ALL_FILES = "@all@"
DOCUMENT_MARK = "@doc@"
DOCUMENT_LINKS = 6  # lines of a document's line-up that a slide can hold, and that a model is given
DOCUMENT_PARAGRAPHS = 5


def _read_document(path):
    _check_file(path, readers.DOCUMENT_TYPES, "document")
    try:
        return documents.read(path, os.path.splitext(os.path.basename(path))[0])
    except PermissionError:
        raise SlideFillError("file_locked", path=path)
    except documents.NotAWordFile as exc:
        raise SlideFillError("file_unreadable", path=path, detail=str(exc))
    except etree.XMLSyntaxError as exc:
        raise SlideFillError("file_unreadable", path=path, detail=str(exc))
FACTS_PER_FILE = 13  # the first facts are the readings; a model's memory is short, and several files have to fit
FACTS_WITH_DOCUMENT = 6
SHEET_TITLE_ROWS = 6
SHEET_TITLE_COLUMNS = 30
SHEET_TITLE_SHORTEST, SHEET_TITLE_LONGEST = 8, 120


class Sources:
    """Several workbooks behind one reader. A key says which with @1@ before its sheet."""

    def __init__(self, paths):
        workbooks = [path for path in paths if os.path.splitext(path)[1].lower() not in readers.DOCUMENT_TYPES]
        if not workbooks:
            raise SlideFillError("doc_needs_workbook")
        self.sources = [ExcelSource(path) for path in workbooks]
        self.names = [os.path.splitext(os.path.basename(path))[0] for path in workbooks]
        self.documents = [_read_document(path) for path in paths if path not in workbooks]
        self._links = None
        self._values = None

    def links(self):
        """What each document's dated entries line up with: [(document, links)]. suggest knows how a sheet is laid out, and it needs this module first, so it comes in here."""
        if self._links is None:
            import suggest
            self._links = [(document, documents.link_all(document.events, suggest.document_items(self, document.events))) for document in self.documents]
        return self._links

    def _all_values(self):
        if self._values is None:
            self._values = set()
            for source in self.sources:
                for ws in source.wb.worksheets:
                    # Only the cells that exist. Asking for rows makes a cell at every place, and a sheet whose range reaches column XFC has sixteen thousand to a row.
                    for cell in ws._cells.values():
                        value = cell.value
                        if isinstance(value, (int, float)) and not isinstance(value, bool):
                            self._values.update((float(value), round(float(value), 6)))
        return self._values

    def _document_text(self, key):
        lang = "ko" if key.strip().startswith(("읽는법", "해설")) else "en"
        number, _hash, paragraph = key.partition(DOCUMENT_MARK)[2].partition("#")
        document, links = self.links()[int(number or 0)]
        if paragraph:
            chosen = documents.key_paragraphs(document)
            return (chosen[int(paragraph) - 1] if int(paragraph) <= len(chosen) else None), document.name, []
        if key.strip().startswith(("해설", "note")):
            return documents.headline(document, links, lang), document.name, []
        found, missing = documents.check_figures(document, self._all_values())
        words = " ".join(part for part in (documents.caution(lang) if links else "", documents.figures_sentence(found, missing, lang)) if part)
        return words or None, document.name, []

    def _document_facts(self, lang):
        """What the model may read from the documents: each line-up with its figures, and failing those the document's own long paragraphs."""
        pack = []
        for number, (document, links) in enumerate(self.links(), start=1):
            for place, found in enumerate(documents.highlights(links, DOCUMENT_LINKS), start=1):
                period = documents.period_text(found.event.date, lang)
                text = (f"[{document.name}] {period}에 적힌 일: {found.event.text} 같은 항목 {found.item.name}의 전 {found.window}개 기간 평균 {found.before:,.0f}, 후 {found.window}개 기간 평균 {found.after:,.0f}, {documents.change_text(found)}." if lang == "ko"
                        else f"[{document.name}] Entered for {period}: {found.event.text} {found.item.name} averaged {found.before:,.0f} over the {found.window} periods before and {found.after:,.0f} over the {found.window} after, {documents.change_text(found)}.")
                pack.append(facts.Fact(f"doc{number}_link{place}", text, tuple(facts.NUMBER.findall(f"{found.before:,.0f} {found.after:,.0f} {documents.change_text(found)}"))))
            if not links:
                long = [text for text in document.paragraphs if len(text) >= documents.LONG_PARAGRAPH][:DOCUMENT_PARAGRAPHS]
                pack += [facts.Fact(f"doc{number}_text{place}", f"[{document.name}] {text}", tuple(facts.NUMBER.findall(text))) for place, text in enumerate(long, start=1)]
        return pack

    @property
    def thousands(self):
        return self.sources[0].thousands

    @thousands.setter
    def thousands(self, value):
        for source in self.sources:
            source.thousands = value

    def _route(self, key):
        match = FILE_MARK.search(key)
        number = int(match.group(1)) if match else 0
        if number >= len(self.sources):
            raise IssueError("missing_sheet", sheet=key)
        return number, self.sources[number], key[:match.start()] + key[match.end():] if match else key

    def _named(self, number, where):
        return f"{self.names[number]} · {where}" if where else self.names[number]

    def text_value(self, key):
        if DOCUMENT_MARK in key:
            return self._document_text(key)
        if ALL_FILES in key:
            word, _mark, references = key.partition(ALL_FILES)
            lang = "ko" if word.strip().startswith(("읽는법", "해설")) else "en"
            readings, _pack, _notes, issues, used, _lang = self._gather(references, word, "", lang)
            return " ".join(fact.text for fact in readings) or None, ", ".join(used), issues
        number, source, plain = self._route(key)
        text, where, issues = source.text_value(plain)
        return text, self._named(number, where), issues

    def _gather(self, references, word, template_text, lang=None):
        """Each file's readings, and what can be calculated across them. word is any prefix that reads a reference as an insight."""
        pack, notes, issues, used, entries = [], [], [], [], []
        for reference in references.split("|"):
            number, source, plain = self._route(reference)
            request, _where, found = source.insight_request(f"통찰:{plain}", template_text)
            issues += found
            if request is None:
                continue
            lang = lang or request.lang
            name = self.names[number]
            figure = _lead_figure(request)
            if figure and figure not in name and not _WHOLE.fullmatch(figure):  # a file called 대구 지하철 승차 whose figure is one station: saying only the file's name would claim the whole system
                name = f"{name} ({figure})"
            used.append(name)
            # With a document in play its line-up is the point, and a model follows whichever it is given the most of.
            pack += [replace(fact, id=f"file{number + 1}_{fact.id}", text=f"[{name}] {fact.text}") for fact in (_kept_with_document(request.facts) if self.documents else request.facts[:FACTS_PER_FILE])]
            notes.append(f"[{name}] {request.fallback}")
            entries.append((name,) + _lead_change(request, lang or request.lang))
        return facts.cross_file_readings(entries, lang or "ko"), pack, notes, issues, used, lang or "ko"

    def range_cells(self, key):
        if DOCUMENT_MARK in key:
            head, _mark, tail = key.partition(DOCUMENT_MARK)
            number, _colon, count = tail.partition(":")
            document, links = self.links()[int(number or 0)]
            return documents.table_rows(links, int(count or DOCUMENT_LINKS), "ko" if head.startswith("ko") else "en"), document.name, []
        number, source, plain = self._route(key)
        rows, where, issues = source.range_cells(plain)
        return rows, self._named(number, where), issues

    def insight_request(self, key, template_text=""):
        if ALL_FILES in key:
            return self._all_files_request(key, template_text)
        number, source, plain = self._route(key)
        request, where, issues = source.insight_request(plain, template_text)
        where = self._named(number, where)
        if request is None:
            return None, where, issues
        return InsightRequest(key, request.lang, request.facts, request.fingerprint, request.fallback, where), where, issues

    def _all_files_request(self, key, template_text):
        """What the model is given to write about the files together: each one's readings, marked with the file they came from."""
        word, _mark, references = key.partition(ALL_FILES)
        readings, pack, notes, issues, used, lang = self._gather(references, word, template_text)
        if not pack:
            raise IssueError("calc_columns", key=key)
        about = ""
        if len(used) >= 2:
            about = (f"이 보고서는 서로 다른 파일 {len(used)}개를 함께 봅니다: {', '.join(used)}. cross로 시작하는 항목이 파일을 함께 계산한 결과이니 이것을 먼저 쓰세요. 한 파일만 설명하지 마세요. 파일 사이의 인과나 상관은 계산 결과에 없으므로 쓰지 마세요." if lang == "ko"
                     else f"This report looks at {len(used)} different files together: {', '.join(used)}. Items starting with cross were calculated across the files, so use them first. Do not describe one file alone. Whether one file affects another is not in the results, so do not say it.")
        extra = self._document_facts(lang) if self.documents else []
        if extra:
            about += (" doc로 시작하는 항목은 문서에 적힌 일과 그 전후 수치입니다. 문장마다 doc 항목을 하나 이상 인용해, 문서에 적힌 일과 그 시기의 수치를 함께 쓰세요. 같은 시기라는 사실만 쓰고, 그 일 때문에 수치가 변했다고 쓰지 마세요." if lang == "ko"
                      else " Items starting with doc are what a document records, with the figures around that time. Cite at least one doc item in each sentence, and give what the document records together with the figures from that time. Say only that they came at the same time, never that one caused the other.")
        pack = tuple([facts.Fact("files", about, ())] + readings + extra + pack)
        where = ", ".join(used)
        return InsightRequest(key, lang, pack, facts.fingerprint(pack, lang), " ".join(notes), where), where, issues


_WHOLE = re.compile(r"합계|총계|전체|total|sum", re.I)  # the whole of the file: its name already says what


def _lead_figure(request):
    """The name of the figure a request is about, from its calculated value: 설화명곡 in 2026년 9월 설화명곡: 144,653."""
    fact = next((fact for fact in request.facts if fact.id == "value"), None)
    if fact is None or ":" not in fact.text:
        return ""
    return re.sub(r"^\s*(?:\d{4}년\s*(?:\d+월\s*(?:\d+일\s*(?:\(.\)\s*)?)?|\d분기\s*)?|\d+월\s*\d+일\s*(?:\(.\)\s*)?|\d{4}-\d{2}(?:-\d{2})?\s*|Q\d \d{4}\s*|\d{4}\s+|[A-Z][a-z]{2} \d{4}\s*|[A-Z][a-z]{2} \d+ (?:\(\w+\) )?)", "", fact.text.partition(":")[0]).strip()


def _lead_change(request, lang):
    """(percent change, period, basis) of a figure, from the facts already calculated for it. A change in points, or no change to give, is None: it cannot stand beside percentages."""
    found = {fact.id: fact for fact in request.facts}
    fact, basis = (found["year"], "year") if "year" in found else (found.get("prev"), "prev")
    match = re.search(r"([+-]?\d[\d,.]*)%", fact.text) if fact else None  # 0.0% has no sign, and no change is still a change to give
    period = next((fact.period for fact in request.facts if fact.period), "")  # as the figure's own sentence says it: yearly data has no month
    return (float(match.group(1).replace(",", "")) if match else None), _period_in(period, lang), basis


def _period_in(period, lang):
    """A file's period in the report's language: one English file among Korean ones must not make 2023 beside 2021년 8월."""
    if lang == "ko":
        year = re.fullmatch(r"(?:19|20)\d\d", period)
        month = re.fullmatch(r"([A-Z][a-z]{2}) ((?:19|20)\d\d)", period)
        quarter = re.fullmatch(r"Q(\d) ((?:19|20)\d\d)", period)
        if year:
            return f"{period}년"
        if month and month.group(1) in facts.MONTHS:
            return f"{month.group(2)}년 {facts.MONTHS.index(month.group(1)) + 1}월"
        if quarter:
            return f"{quarter.group(2)}년 {quarter.group(1)}분기"
        return period
    year = re.fullmatch(r"((?:19|20)\d\d)년", period)
    month = re.fullmatch(r"((?:19|20)\d\d)년 (\d{1,2})월", period)
    quarter = re.fullmatch(r"((?:19|20)\d\d)년 (\d)분기", period)
    if year:
        return year.group(1)
    if month:
        return f"{facts.MONTHS[int(month.group(2)) - 1]} {month.group(1)}"
    if quarter:
        return f"Q{quarter.group(2)} {quarter.group(1)}"
    return period


def _period_of(pack, header):
    """The period the newest value belongs to, as the value's own fact writes it: 2026년 8월 in 2026년 8월 내국인: 192."""
    fact = next((fact for fact in pack if fact.id == "value"), None)
    return fact.text.partition(":")[0].replace(header, "").strip() if fact else ""


def _value_of(pack):
    fact = next((fact for fact in pack if fact.id == "value"), None)
    return fact.text.partition(":")[2].strip() if fact else ""


def _kept_with_document(pack):
    """What a file keeps when a document shares the model's short memory: what it is, its current value, and its first readings."""
    must = [fact for fact in pack if fact.id in ("subject", "sheet", "value")]
    return must + [fact for fact in pack if fact not in must][:FACTS_WITH_DOCUMENT - len(must)]


def _with_sheet(pack, ref, lang, figure="", period=""):
    """What was counted and when, for the model and for the checks. A column headed 외국인 in a sheet of sales is not a number of foreigners, and only the sheet's name says so."""
    name = _sheet_names(ref)
    title = _sheet_title(ref.sheets[0])
    value = _value_of(pack)
    if lang != "ko" and figure == COUNT_WORD_KO:  # the notation follows the window's language; the sentence follows the data's
        figure = "the number of rows"
    if lang == "ko":
        now = f"이번 기간은 {period}이고 이번 값은 {value}입니다. " if period and value else ""
    else:
        now = f"The current period is {period} and its value is {value}. " if period and value else ""
    whole = _WHOLE.fullmatch(figure or "") is not None
    if lang == "ko":
        subject = f"{now}이 수치는 시트 전체의 합계입니다." if whole else f"{now}이 파일의 수치 이름은 {figure}이고 전체가 아닐 수 있으니 문장에 이 이름을 쓰세요."
        subject += " 다른 기간의 값은 비교 대상일 뿐입니다."
        sheet = f"이 수치가 있는 시트의 이름: {name}" + (f". 시트 제목: {title}" if title else "")
    else:
        subject = f"{now}This figure is the total of the sheet." if whole else f"{now}The figure in this file is called {figure}, which may not be the whole, so use that name in each sentence."
        subject += " Values for other periods are only for comparing."
        sheet = f"The sheet these figures are on is named: {name}" + (f". Its title says: {title}" if title else "")
    return (facts.Fact("subject", subject, (), figure, period), facts.Fact("sheet", sheet, ())) + tuple(pack)


def _sheet_title(ws):
    """A line alone at the top of a sheet, which often says what is being counted when the column headings do not: Quarterly Passenger Statistics."""
    for row in ws.iter_rows(min_row=1, max_row=SHEET_TITLE_ROWS, max_col=min(ws.max_column, SHEET_TITLE_COLUMNS), values_only=True):
        filled = [cell for cell in row if cell not in (None, "")]
        if len(filled) == 1 and isinstance(filled[0], str) and SHEET_TITLE_SHORTEST <= len(filled[0].strip()) <= SHEET_TITLE_LONGEST:
            return " ".join(filled[0].split())
    return ""


def _with_commas(rows):
    """493,434 for a number in General format, leaving the first column alone: years and codes live there."""
    done = []
    for row in rows:
        cells = [row[0]]
        for raw, text, fmt in row[1:]:
            if fmt in ("General", "0") and isinstance(raw, (int, float)) and not isinstance(raw, bool) and abs(raw) >= 1000:
                text = re.sub(r"^(-?)(\d+)", lambda m: f"{m.group(1)}{int(m.group(2)):,}", text)
            cells.append((raw, text, fmt))
        done.append(cells)
    return done


def run(template, excel, output=None, insights=None, source=None, thousands=False):
    """Fill the template. Without output this is the preview: same work, nothing written.

    insights holds sentences a person chose earlier. No model is called here, so preview and report always agree.
    source is a workbook already read from excel, so a large one need not be read twice in a row.
    thousands writes plain large numbers in tables with commas, for a report SlideFill laid out itself.
    """
    if source is None:
        source = Sources(excel) if isinstance(excel, (list, tuple)) else ExcelSource(excel)
    source.thousands = thousands
    prs = open_template(template)
    work = _Run(source, prs, insights)
    work.walk()
    report = Report(items=list(work.items.values()), requests=work.requests)
    if output:
        report.output = _save(prs, output)
    return report


def _save(prs, output):
    tmp = output + ".tmp"
    try:
        prs.save(tmp)
        os.replace(tmp, output)
    except OSError as exc:
        if os.path.exists(tmp):
            os.remove(tmp)
        raise SlideFillError("cannot_save", path=output, detail=str(exc))
    return output


def output_path(folder, template, today=None):
    """<template>_<date>.pptx in folder, with _2, _3... added so an existing or open file is never touched."""
    stem = os.path.splitext(os.path.basename(template))[0]
    date = (today or dt.date.today()).isoformat()
    candidate = os.path.join(folder, f"{stem}_{date}.pptx")
    n = 2
    while os.path.exists(candidate):
        candidate = os.path.join(folder, f"{stem}_{date}_{n}.pptx")
        n += 1
    return candidate


# ---------------------------------------------------------------- settings

def settings_path(base, name=SETTINGS_FILE):
    path = os.path.join(base, name)
    if os.access(path, os.W_OK) or (not os.path.exists(path) and os.access(base, os.W_OK)):
        return path
    alt = os.path.join(os.environ.get("APPDATA", base), APP_NAME)
    os.makedirs(alt, exist_ok=True)
    return os.path.join(alt, name)


def load_settings(path):
    """Read settings. A damaged file is moved aside to .broken instead of being overwritten."""
    try:
        with open(path, encoding="utf-8") as fh:
            data = json.load(fh)
        return data if isinstance(data, dict) else {}
    except FileNotFoundError:
        return {}
    except (ValueError, OSError):
        try:
            os.replace(path, path + ".broken")
        except OSError:
            pass
        return {}


def save_settings(path, data):
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as fh:
        json.dump(data, fh, ensure_ascii=False, indent=2)
    os.replace(tmp, path)
