"""SlideFill logic: fills {{placeholders}}, tables and charts of a PowerPoint template from an Excel workbook.

No GUI imports. run() is the single code path for both preview (output=None) and generation.
A value that cannot be determined never becomes a blank: its placeholder stays visible in the output
and the problem is reported as an issue.
"""

import calendar
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
from fractions import Fraction

import openpyxl
from lxml import etree
from openpyxl.utils import range_boundaries
from openpyxl.utils.cell import column_index_from_string, get_column_letter
from pptx import Presentation
from pptx.chart.data import CategoryChartData, XyChartData, BubbleChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.opc.package import XmlPart

import documents
import facts
import readers
import layouts
import native_vba_cases
from verification import automatic
from verification.kernel import sum_rationals


class _NoPivotCaches(dict):
    def __missing__(self, key):
        return None


# A pivot table's cache holds every source record, read again for each sheet: one public workbook passed 6 GB here. What the table shows is stored in the sheet's cells,
# and nothing here ever saves a workbook it has read, so the caches are left unread.
openpyxl.reader.workbook.WorkbookParser.pivot_caches = property(lambda self: _NoPivotCaches())
from numfmt import format_color, format_value
from numeric_text import unbacked_numbers

warnings.filterwarnings("ignore", module="openpyxl")

APP_NAME = "SlideFill"
SETTINGS_FILE = "SlideFill.json"
INSIGHTS_FILE = "SlideFill.insights.json"
OLE_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"  # legacy .xls/.ppt, and also encrypted or many DRM-wrapped files

A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
P = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
A_P, A_R, A_T, A_BR, A_RPR = A + "p", A + "r", A + "t", A + "br", A + "rPr"
DSP = "http://schemas.microsoft.com/office/drawing/2008/diagram"
RT_DIAGRAM_DRAWING = "http://schemas.microsoft.com/office/2007/relationships/diagramDrawing"

PLACEHOLDER = re.compile(r"\{\{\s*(.+?)\s*\}\}", re.S)
MARKER = re.compile(r"\{\{\s*(table|chart|표|차트)\s*:\s*(.+?)\s*\}\}", re.I | re.S)
COORD = re.compile(r"[A-Za-z]{1,3}[0-9]+(:[A-Za-z]{1,3}[0-9]+)?")
LATEST = re.compile(r"\s*(?:최신|latest)\s*(\d*)\s*:(.*)", re.I | re.S)
LATEST_COLUMNS = re.compile(r"[A-Za-z]{1,3}(\s*:\s*[A-Za-z]{1,3})?")
DATE_TEXT = re.compile(r"(\d{4})\s*[-/.년]\s*(\d{1,2})\s*(?:[-/.월]\s*(\d{1,2})\s*[.일]?)?\s*[월.]?\s*(?:\(\w+\))?"
                       r"\s*(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?")
DATE_DIGITS = re.compile(r"(\d{4})(\d{2})(\d{2})")
EMPTY_EDGE_ROWS, EMPTY_EDGE_COLUMNS = 1000, 100  # cells with only a format this far past the last value are dropped on loading
DATE_YEARS = (1800, 2199)  # the years a date written as text or digits may have
YEAR_TEXT = re.compile(r"((?:19|20)\d\d)\s*(?:년도?)?|\[\s*((?:19|20)\d\d)\s*\]")  # [2023] is how a catalogue writes a year it worked out
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
FISCAL_QUARTER_SPAN = re.compile(r"Q([1-4])\s*[-.' ]?\s*(\d{4}|\d{2})\s*[/–—-]\s*(\d{2})", re.I)
FISCAL_QUARTER_FY = re.compile(r"Q([1-4])\s*[-.' ]?\s*FY\s*'?(\d{4}|\d{2})|FY\s*'?(\d{4}|\d{2})\s*[-.' ]?\s*Q([1-4])", re.I)
DAY_MONTH_YEAR = re.compile(r"(\d{1,2})\s+([A-Za-z]{3,9})\.?,?\s+(\d{4})")
MONTH_DAY_YEAR = re.compile(r"([A-Za-z]{3,9})\.?\s+(\d{1,2}),\s*(\d{4})")
MONTH_QUARTER = re.compile(r"([A-Za-z]{3,9})\s+quarter\s+(\d{4})", re.I)
YEAR_SPAN = re.compile(r"((?:19|20)\d\d)\s*[-/–—]\s*(\d{2})")
FISCAL_YEAR = re.compile(r"(?:fy\s*)?((?:19|20)\d\d)\s*[–—]\s*(\d{2})", re.I)  # 2007–08, named by the year it ends. 2007-08 with a plain hyphen reads as a month
MONTH_ALONE = re.compile(r"0?(\d{1,2})\s*월")
DIGITS = r"(?:\d{1,3}(?:,\d{3})+|\d{1,3}(?:[ \u00a0\u202f]\d{3})+|\d+)"  # commas or spaces only where they group thousands: 1,2,3 is three figures in one cell, not 123, and 39 925 in a statistics yearbook is 39,925
HEADING_SCAN_CELLS = 20000  # a small sheet's cells are looked through at once; a large one's heading row only to this many columns
HEADING_SCAN_COLUMNS = 200
ONE_LABEL_SHARE = 0.8  # lines with one label among the label columns, for those columns to be levels of one outline
SHEET_WORD_ROWS = 20
SHEET_WORD_COLUMNS = 10
GROUPING = re.compile(r"[, \u00a0\u202f]")
NUMBER_TEXT = re.compile(r"\s*([+-]?" + DIGITS + r"(?:\.\d+)?(?:[eE][+-]?\d+)?)\s*(?:\([^()]*\)|\[[^\[\]]*\]|명|건|원|개|회|톤|대|호)?\s*")
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
TOTAL_WORDS = {"합계", "총계", "총합", "총합계", "계", "전체", "전국", "전산업", "total", "grandtotal", "sum", "합"}  # 전국 over the regions, 전산업 over the industries
SUBTOTAL_WORDS = {"소계", "subtotal"}
GROUP_TOTAL_WORDS = 4  # 경상 수입 계: a short name before 계 or 소계. A longer line is a sentence that happens to end so
DYNAMIC_TOKEN = re.compile(r"(?:'(?:[^']|'')+'|[^\s!(),:+\-'\"]+)![$A-Za-z0-9:]+(?=[\s,)+\-]|:[A-Za-z]|$)|[A-Za-z]+(?=\s*\()|\d+|[(),:+\-]")  # a sheet's cells, a function, a whole number, or a mark
DYNAMIC_MAX_SIDE = 200
DYNAMIC_MAX_CELLS = 1_048_576
MONTHS_ONLY_YEAR = 2001  # a year no sheet of months alone is about, only to put them in order; a sentence calls them as the sheet does
SHEET_FIGURE_LINES = 60  # how far along a month's line a figure is looked for
MONTHS_ONLY_MIN = 3
YEARS_ACROSS_LOOK = 40
SHEET_YEAR = re.compile(r"(?<!\d)((?:19|20)\d\d)(?!\d)(?:\s*년)?")
SHEET_YEAR_ROWS = 4
SHEET_YEAR_COLUMNS = 10
SHEET_YEAR_LENGTH = 60
NAME_LOOK_RIGHT = 3  # columns to the right of a name that may tell two lines of that name apart
PERIOD_COLUMN_LOOK = 500  # cells of a condition's column read to tell whether it holds periods
PERIOD_COLUMN_SHARE = 0.6  # its heading and a note or two below need not be dates
PERIOD_FILTER = re.compile(r"((?:18|19|20|21)\d\d)(?:\s*년|\s*[-./년]\s*(\d{1,2})\s*월?(?:\s*[-./]\s*(\d{1,2})\s*일?)?)?")  # 2022, 2022년, 2022-05 or 2022년 5월, kept by a condition
SUBTOTALS_ADD_UP = 0.001  # a grand total this close to the other total lines is their sum, not one more kind of thing
YEAR_IN_HEADING = re.compile(r"(?:19|20)\d\d(?:\s*[-–/]\s*\d{2,4})?")  # the year a column heading names, as 2024-25 or 2025
BUDGET_NAME = re.compile(r"예산|목표|계획|\b(?:budget|target|plan)(?:s|ned)?\b", re.I)  # what was planned for a period, set beside what happened in it
TOTAL_ROW = re.compile(r"totals?(?![a-z])(?!.*\b(?:pty|ltd|inc|limited|llc|corp|plc|gmbh)\b)", re.I)  # Total revenue, not Total Tools Importing Pty Ltd or TotalEnergies
TOTAL_TOLERANCE = 0.005  # a stated total this close to the items' sum is taken to be their total, allowing for rounding
LATE_MINUTES = 60  # a newest day's last entry this much earlier than usual looks unfinished
SHARE_HEADER = "%"
AMOUNT_HEADING = re.compile(r"[증중]감(?![률율])")  # 증감 alone is an amount that went up or down, 증감률 is a rate
RATE_HEADING = re.compile(facts.RATE_SYLLABLE + r"|증감|전년비|전년동월비|지수.*증|점수|만족도|\bscores?\b|\bpercentage|% ?change|\bgrowth|\brates?\b|\bratios?\b", re.I)
HANGUL = re.compile(r"[가-힣]")
ID_CODE = re.compile(r"[A-Z0-9]{2,8}")  # J468: a dataset's code, not a name
MAX_LINES_FOR_A_REPEAT = 200  # how far along a row of headings to look for the same one again
HEADING_LOOK_AHEAD = 30
PIVOT_HEADING = re.compile(r"column labels|row labels|values|열 레이블|행 레이블|값", re.I)
GRAND_TOTAL = re.compile(r"(?:grand\s+)?total|총합계|합계|총계|전체", re.I)
UNIT_NOTE = re.compile(r"단위\s*[:：]|[(（\[]\s*단위|\bunits?\s*[:：]|[(（\[]\s*units?\b", re.I)  # not 유통단계별단위명, which is only a column's name
YEAR_LINE = re.compile(r"연간|연평균|연계|연합계|년간|annual|yearly|full\s*year", re.I)
MONTH_LINE = re.compile(r"(?:0?[1-9]|1[0-2])\s*월|jan(?:uary)?|feb(?:ruary)?|mar(?:ch)?|apr(?:il)?|may|june?|july?|aug(?:ust)?|sep(?:t(?:ember)?)?|oct(?:ober)?|nov(?:ember)?|dec(?:ember)?", re.I)
MIN_MONTH_LINES = 6
PART_CODE = re.compile(r"([0-9A-Za-z]+(?:[-.][0-9A-Za-z]+)*)")  # 611 and 611-01 at the start of an account's name
BRACKETED_LOSS_DIGITS = 3
COLUMN_MARK = re.compile(r"\(\s*\d{1,2}\s*\)(?:\s*[=+\-−×*/÷]\s*(?:\(\s*\d{1,2}\s*\)|\d+))*")  # (1), (3)=(1)-(2): a column's number, not a figure
NEGATIVE_TEXT = re.compile(r"[△▵−]\s*(" + DIGITS + r"(?:\.\d+)?)|\(\s*(" + DIGITS + r"(?:\.\d+)?)\s*\)")
QUARTER_SCAN = 30
QUARTER_WORD = re.compile(r"Q\s*([1-4])|([1-4])\s*(?:분기|Q)", re.I)
QUARTER_HEADING = re.compile(r"분기|quarter|qtr|\bq\b", re.I)
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
PERIOD_WORD = {"ko": "기간", "en": "Period"}
ITEM_WORD = {"ko": "항목", "en": "Item"}
TRAILING_TOTAL = re.compile(r"\s+(?:total|합계|총계)\s*$", re.I)  # the row under a date heading is the sheet's total, not part of what the dates are called
FISCAL_YEAR_ENDS = (6, 30)  # internal July-start coordinate for sorting FY quarters, never evidence that a period is complete
_MONTH_NAME = r"(jan(?:uary)?|feb(?:ruary)?|mar(?:ch)?|apr(?:il)?|may|june?|july?|aug(?:ust)?|sept?(?:ember)?|oct(?:ober)?|nov(?:ember)?|dec(?:ember)?)"
_DAY = r"(?:\d{1,2}(?:st|nd|rd|th)?\s+)?"
FISCAL_END_WORDS = re.compile(rf"\b(?:(?:year|fy)\s+(?:ending|ended|ends|end)|(?:financial|fiscal)\s+year\s+to|fy\s+to)\s+(?:on\s+)?{_DAY}{_MONTH_NAME}\b", re.I)  # year ending 30 June, FY ended March 31; a bare "year to October" is a year to date
FISCAL_MONTH_SPAN = re.compile(rf"\b{_DAY}{_MONTH_NAME}\b(?:\s+(?:19|20)\d\d)?\s*(?:to|through|until|[-–—~])\s*{_DAY}{_MONTH_NAME}\b", re.I)  # July to June, 1 July 2024 to 30 June 2025
FISCAL_MONTH_SPAN_KO = re.compile(r"(\d{1,2})\s*월(?:\s*\d{1,2}\s*일)?\s*(?:부터|[~～\-–])\s*(?:익년|다음\s*해|이듬해)?\s*(\d{1,2})\s*월")  # 4월~3월, 7월 1일부터 6월 30일
FISCAL_END_KO = re.compile(r"(\d{1,2})\s*월\s*(?:말\s*)?결산")  # 3월 결산
FISCAL_WORD_ROWS = 30  # notes on the year sit at the top of a sheet or under its table
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
ERRORS.update({"verification_failed", "nonfinite_value", "chart_bubble_columns", "chart_bubble_size", "chart_bubble_empty"})


class Shown(str):
    """Text with a whole-value colour or explicit source-value spans; prose keeps its own colour."""

    def __new__(cls, text, color=None, spans=()):
        made = super().__new__(cls, text)
        made.color = color
        made.spans = tuple(spans)
        return made

    def __reduce__(self):  # copy and pickle keep the color
        return Shown, (str(self), self.color, self.spans)


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
    verification: dict = field(default_factory=dict)
    source_snapshot: object = field(default=None, repr=False)
    vba_cases: native_vba_cases.VbaCases | None = field(default=None, repr=False)

    @property
    def blocked(self):
        return self.verification.get("state") == "blocked"

    @property
    def incomplete(self):
        return self.verification.get("state") == "partial"

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
        if not DATE_YEARS[0] <= int(year) <= DATE_YEARS[1]:
            return None  # 1001.5 written as text is a figure, not May of the year 1001
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
    return _fiscal_quarter(raw)


def _fiscal_quarter(raw):
    """Q1 12/13, Q1 FY24 and FY24 Q1: a quarter of a year that spans two. Read as starting in July only to put them in order: a sentence says them as the sheet does."""
    match = FISCAL_QUARTER_SPAN.fullmatch(raw)
    if match:
        start, end = int(match.group(2)), int(match.group(3))
        if (start + 1) % 100 != end % 100:
            return None
        ending = (start if start > 99 else 2000 + start) + 1
        quarter = int(match.group(1))
    else:
        match = FISCAL_QUARTER_FY.fullmatch(raw)
        if not match:
            return None
        year = int(match.group(2) or match.group(3))
        ending = year if year > 99 else 2000 + year
        quarter = int(match.group(1) or match.group(4))
    month = 7 + (quarter - 1) * 3
    return dt.datetime(ending - 1, month, 1) if month <= 12 else dt.datetime(ending, month - 12, 1)


def _months_later(date, months):
    """The same day of the month, months later (or earlier when negative). Quarters start on the 1st, so the day always exists."""
    index = date.year * 12 + date.month - 1 + months
    return date.replace(year=index // 12, month=index % 12 + 1)


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
        return dt.datetime(int(match.group(1) or match.group(2)), 1, 1)
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
    if isinstance(raw,Fraction):
        return raw.numerator if raw.denominator==1 else raw
    if isinstance(raw, (int, float)):
        return raw if not isinstance(raw,float) or math.isfinite(raw) else None
    match = NUMBER_TEXT.fullmatch(raw) if isinstance(raw, str) else None
    if not match:
        return None
    exact = Fraction(GROUPING.sub("", match.group(1)))
    if exact.denominator==1:return exact.numerator
    try:number=float(exact)
    except OverflowError:return None
    return number if math.isfinite(number) else None


def _sheet_words(ws, rows=SHEET_WORD_ROWS, columns=SHEET_WORD_COLUMNS):
    """One vote for each cell of words in the top corner of a sheet, as one letter of the script it is mostly in: 회사 and Hyundai Motor Company count once each, so long English names do not outvote short Korean headings."""
    found = []
    for row in range(1, min(ws.max_row, rows) + 1):
        for column in range(1, min(ws.max_column, columns) + 1):
            cell = ws._cells.get((row, column))
            if cell is not None and isinstance(cell.value, str):
                found.append(cell.value)
    return _votes(found)


def fiscal_year_end(texts):
    """The month a financial year ends in, as the file's own words put it, or None when they say nothing or disagree.

    A span of months counts only when it covers a whole year (July to June). December is the calendar year and says
    nothing about a 2025-26 year, so it is not taken."""
    found = set()

    def month(name):
        return MONTH_NUMBERS[name[:3].lower()]

    def whole_year(start, end):
        return 1 <= start <= 12 and end == (start + 10) % 12 + 1

    for text in texts:
        found.update(month(m.group(1)) for m in FISCAL_END_WORDS.finditer(text))
        found.update(month(m.group(2)) for m in FISCAL_MONTH_SPAN.finditer(text) if whole_year(month(m.group(1)), month(m.group(2))))
        found.update(int(m.group(2)) for m in FISCAL_MONTH_SPAN_KO.finditer(text) if whole_year(int(m.group(1)), int(m.group(2))))
        found.update(int(m.group(1)) for m in FISCAL_END_KO.finditer(text) if 1 <= int(m.group(1)) <= 12)
    found.discard(12)
    return found.pop() if len(found) == 1 else None


def _fiscal_words(ws):
    """The sheet's name and the words at its top and bottom, where a note on the financial year is written."""
    cells = getattr(ws, "_cells", {})
    rows = set(range(1, min(ws.max_row, FISCAL_WORD_ROWS) + 1)) | set(range(max(ws.max_row - FISCAL_WORD_ROWS, 1), ws.max_row + 1))
    found = [ws.title]
    for row in sorted(rows):
        for column in range(1, min(ws.max_column, SHEET_WORD_COLUMNS * 3) + 1):
            cell = cells.get((row, column))
            if cell is not None and isinstance(cell.value, str):
                found.append(cell.value)
    return found


def _votes(texts):
    """Each text as one letter of the script it is mostly in, so a long name counts no more than a short one."""
    found = []
    for text in texts:
        korean, latin = len(HANGUL.findall(text)), len(facts.LATIN.findall(text))
        if korean or latin:
            found.append("가" if korean >= latin else "a")
    return found


def figure_number(raw):
    """What a cell holds as a figure: as_number, and also a loss written as text, △500 or (300) or −500. Only for cells already read as figures: elsewhere (1) is a numbered heading."""
    number = as_number(raw)
    if number is not None or not isinstance(raw, str):
        return number
    match = NEGATIVE_TEXT.fullmatch(raw.strip())
    if not match:
        return None
    if match.group(2) and len(re.sub(r"\D", "", match.group(2))) < BRACKETED_LOSS_DIGITS and not re.search(r"[,. \u00a0\u202f]", match.group(2)):
        return None  # (1) (2) (3) under the headings number the columns; a loss kept as text is (300) or (1,234)
    exact = -Fraction(GROUPING.sub("", match.group(1) or match.group(2)))
    if exact.denominator==1:return exact.numerator
    try:number=float(exact)
    except OverflowError:return None
    return number if math.isfinite(number) else None


def _holds_quarters(ws, ref, line):
    """The second part of a split date is a quarter when its heading or its words say so: 4 under 분기 is not April."""
    known = ws.__dict__.setdefault("_quarter_lines", {})
    key = (ref.by_row, line)
    if key not in known:
        words = [cell.value.strip() for cell in (ws._cells.get(ref.at(position, line)) for position in range(1, QUARTER_SCAN + 1)) if cell is not None and isinstance(cell.value, str)]
        known[key] = any(QUARTER_HEADING.search(word) or QUARTER_WORD.fullmatch(word) for word in words)
    return known[key]


def _unread_figure(raw):
    """Digits where a figure belongs that cannot be read as one, such as ▲500 or 14000 and 15000 in one cell: leaving them out changes the total, so it must not go unsaid. With letters in it, it is a heading or a note."""
    return isinstance(raw, str) and any(c.isdigit() for c in raw) and not any(c.isalpha() for c in raw) and figure_number(raw) is None and as_date(raw) is None and not COLUMN_MARK.fullmatch(raw.strip())


def as_part_of_date(raw):
    """A month or a day held in a column of its own: 9, 09, 9월, Sep. A quarter too: Q4, 4분기. None for anything else."""
    if isinstance(raw, bool):
        return None
    if isinstance(raw, (int, float)) and float(raw).is_integer():
        return int(raw)
    if isinstance(raw, str):
        if raw.strip().isdigit():
            return int(raw.strip())
        quarter = QUARTER_WORD.fullmatch(raw.strip())
        if quarter:
            return int(quarter.group(1) or quarter.group(2))
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
        return f"{_when(point)} · {self.cell_text(point)}"

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
    basis: str = ""  # budget when the older of two columns is what was planned, not an earlier period
    wholes: list = field(default_factory=list)  # total lines of unlike things, as Total income and Total expenses: facts.Item each
    origin_rows: dict = field(default_factory=dict)
    number_format: str = 'General'

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


def _dynamic_area(source, formula):
    """(sheet, (left, top, right, bottom)) of a name that grows with its data, such as OFFSET(표!$B$2,0,0,COUNTA(표!$B:$B)-1,1)
    or 표!$B$2:INDEX(표!$B:$B,COUNTA(표!$B:$B)). Only OFFSET, INDEX, COUNTA, COUNT, ROWS and COLUMNS over cells, with whole numbers, + and -; anything else is None."""
    tokens = DYNAMIC_TOKEN.findall(formula.strip().lstrip("="))
    if not tokens or "".join(tokens).replace(" ", "") != re.sub(r"\s+", "", formula.strip().lstrip("=")):
        return None
    at = [0]

    def peek():
        return tokens[at[0]] if at[0] < len(tokens) else ""

    def take(expected=None):
        token = peek()
        if expected is not None and token.upper() != expected:
            raise ValueError(formula)
        at[0] += 1
        return token

    def area(token):
        sheet, _mark, cells = token.rpartition("!")
        ws = source._sheets.get(sheet.strip("'").replace("''", "'").casefold())
        if ws is None:
            raise ValueError(sheet)
        first, _colon, last = cells.replace("$", "").partition(":")
        left, top = _cell_place(first, ws, True)
        right, bottom = _cell_place(last or first, ws, False)
        return ws, (left, top, right, bottom)

    def atom():
        token = take()
        if token.isdigit():
            return int(token)
        if token == "-":
            return -atom()
        if token == "(":
            value = expression()
            take(")")
            return value
        if "!" in token:
            value = area(token)
            if peek() == ":":  # 표!$B$2:INDEX(...): from one corner to the other
                take()
                other = atom()
                (ws, (left, top, _r, _b)), (_ws, (_l, _t, right, bottom)) = value, other
                return ws, (min(left, other[1][0]), min(top, other[1][1]), max(right, other[1][2]), max(bottom, other[1][3]))
            return value
        name = token.upper()
        take("(")
        args = [expression()]
        while peek() == ",":
            take()
            args.append(expression())
        take(")")
        if name in ("COUNTA", "COUNT"):
            ws, (left, top, right, bottom) = args[0]
            bottom = min(bottom, ws.max_row)
            stored = getattr(ws, "_cells", {})  # ws.cell would make an empty cell at every place it is asked about
            cells = (cell.value for row in range(top, bottom + 1) for column in range(left, right + 1) if (cell := stored.get((row, column))) is not None)
            return sum(raw not in (None, "") and (name == "COUNTA" or (isinstance(raw, (int, float)) and not isinstance(raw, bool))) for raw in cells)
        if name in ("ROWS", "COLUMNS"):
            _ws, (left, top, right, bottom) = args[0]
            return bottom - top + 1 if name == "ROWS" else right - left + 1
        if name == "OFFSET" and 3 <= len(args) <= 5:
            ws, (left, top, right, bottom) = args[0]
            down, across = args[1], args[2]
            height = args[3] if len(args) > 3 else bottom - top + 1
            width = args[4] if len(args) > 4 else right - left + 1
            if height < 1 or width < 1 or top + down < 1 or left + across < 1:
                raise ValueError(formula)
            return ws, (left + across, top + down, left + across + width - 1, top + down + height - 1)
        if name == "INDEX" and 2 <= len(args) <= 3:
            ws, (left, top, right, bottom) = args[0]
            row, column = args[1], args[2] if len(args) > 2 else 1
            if top == bottom and right > left and len(args) == 2:  # INDEX(B2:G2, 6): in one row the number counts across, as Excel reads it
                row, column = 1, args[1]
            if not (1 <= row <= bottom - top + 1 and 1 <= column <= right - left + 1):  # Excel gives #REF! outside the range
                raise ValueError(formula)
            return ws, (left + column - 1, top + row - 1, left + column - 1, top + row - 1)
        raise ValueError(name)

    def expression():
        value = atom()
        while peek() in ("+", "-"):
            sign = take()
            other = atom()
            if not isinstance(value, int) or not isinstance(other, int):
                raise ValueError(formula)
            value = value + other if sign == "+" else value - other
        return value

    try:
        found = expression()
    except (ValueError, IndexError, TypeError):
        return None
    if at[0] != len(tokens) or not isinstance(found, tuple):
        return None
    ws, (left, top, right, bottom) = found
    return (ws, (left, top, right, bottom)) if right - left < DYNAMIC_MAX_SIDE and bottom - top < DYNAMIC_MAX_CELLS else None


def _cell_place(text, ws, start):
    """(column, row) of A1, or of a whole column A (its first row when start, else its last filled row)."""
    match = re.fullmatch(r"([A-Za-z]{1,3})(\d*)", text)
    if not match:
        raise ValueError(text)
    column = column_index_from_string(match.group(1).upper())
    return column, int(match.group(2)) if match.group(2) else (1 if start else ws.max_row)


def _years_across(ws, ref, first):
    """Years on another line, before the months: 2017년 to 2025년 down the first column with 1월 to 12월 across the top. The years are the periods and the months their parts."""
    end = ws.max_row if ref.by_row else ws.max_column  # how far the lines beside the date line run
    for position in range(1, first):  # the label column before months laid across, or the heading rows above months laid down
        cells = [ws.cell(*ref.at(position, line)).value for line in range(ref.date + 1, min(end, ref.date + YEARS_ACROSS_LOOK) + 1)]
        if sum(as_year(cell) is not None for cell in cells if cell not in (None, "")) >= 2:
            return True
    return False


def _when(point):
    """A point's period for the place a figure came from: as the sheet names it when a date cannot (12월, 2024-25), else the date."""
    return point.label or point.date.date().isoformat()


def _months_in_order(months):
    """{position: month} that runs 1월, 2월, 3월 down a line, each month once: a year of months. 3월 and 1월 here and there are not."""
    if len(months) < MONTHS_ONLY_MIN:
        return False
    order = [months[position] for position in sorted(months)]
    return len(set(order)) == len(order) and order == sorted(order)


def _names_told_apart(ws, ref, labels, specs, order, places):
    """{old key: fuller name} for items whose name is written on more than one line: 2,001이상 under 일반용 and under 대중탕용.
    The words that tell them apart come from another column on the same lines, the label columns first, then the nearest to the left, then to the right."""
    by_name = {}
    for key in order:
        by_name.setdefault(key if isinstance(key, str) else key[0], []).append(key)
    repeated = [keys for keys in by_name.values() if len(keys) > 1]
    if not repeated:
        return {}
    figure_lines = {line for spec in specs for line in (spec.lines if isinstance(spec, Aggregate) else [spec])}
    first_label = min(labels)
    nearest = sorted((line for line in range(1, first_label + NAME_LOOK_RIGHT + 1) if line not in figure_lines and line not in labels),
                     key=lambda line: (line > first_label, abs(line - first_label)))
    taken = {key if isinstance(key, str) else key[0] for key in order}
    renamed = {}
    for keys in repeated:
        rows = [places.get(key) for key in keys]
        if None in rows:
            continue
        for line in [line for line in reversed(labels[:-1])] + nearest:
            words = []
            for row in rows:
                raw = _merged_value(ws, *ref.at(row, line))
                words.append(_clean_label(raw) if isinstance(raw, str) and raw.strip() and as_number(raw) is None else "")
            if all(words) and len(set(words)) == len(words):
                name = keys[0] if isinstance(keys[0], str) else keys[0][0]
                fuller = [f"{word} {name}" if line < first_label else f"{name} ({word})" for word in words]
                if not taken & set(fuller):
                    renamed.update(zip(keys, fuller))
                break
    return renamed


def _holds_periods(ref, line):
    """True when most filled cells of a condition's column, on the first sheet, read as dates, years or financial years."""
    ws = ref.sheets[0]
    size = min(ws.max_column if ref.by_row else ws.max_row, PERIOD_COLUMN_LOOK)
    cells = [raw for position in range(1, size + 1) if (raw := _merged_value(ws, *ref.at(position, line))) not in (None, "")]
    periods = sum(as_date(raw) is not None or as_year(raw) is not None or (isinstance(raw, str) and YEAR_SPAN.fullmatch(raw.strip()) is not None) for raw in cells)
    return bool(cells) and periods >= len(cells) * PERIOD_COLUMN_SHARE


def _filtered_period(ref, lang):
    """The period a condition such as B=2022-05 keeps a comparison to, in words: 2022년 5월, or May 2022. None when no condition is a period."""
    for value in ref.values:
        for line, text in (value.filters if isinstance(value, Aggregate) else ()):
            if not _holds_periods(ref, line):
                continue  # B=2100 on a column of product codes keeps a product, not a year
            span = YEAR_SPAN.fullmatch(str(text).strip())
            if span and (int(span.group(1)) + 1) % 100 == int(span.group(2)) % 100:
                return str(text).strip()  # 2025-26, a financial year, as the sheet writes it
            match = PERIOD_FILTER.fullmatch(str(text).strip())
            if not match:
                continue
            year, month, day = (int(part) if part else None for part in match.groups())
            if month is None:
                return f"{year}년" if lang == "ko" else str(year)
            if not 1 <= month <= 12:
                return None
            if day is None:
                return f"{year}년 {month}월" if lang == "ko" else f"{facts.MONTHS[month - 1]} {year}"
            return f"{year}년 {month}월 {day}일" if lang == "ko" else f"{day} {facts.MONTHS[month - 1]} {year}"
    return None


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


def _unit_notes(ws, rows, columns):
    """(row, the words as written, the unit alone) for each note saying how figures are measured. (단위: 백만원) says it in one cell; Unit: may leave it to the cell beside or below."""
    last_row, last_column = min(ws.max_row, rows), min(ws.max_column, columns)
    for row in range(1, last_row + 1):
        for column in range(1, last_column + 1):
            cell = ws._cells.get((row, column))  # only cells that exist: asking for rows makes a cell at every place
            if cell is None or not isinstance(cell.value, str):
                continue
            text = cell.value.strip()
            match = UNIT_NOTE.search(text)
            if not match or len(text) > UNIT_NOTE_LENGTH:
                continue
            rest = text[match.end():].strip(" )）]:：")
            if rest:
                yield row, text, rest
                continue
            for place in ((row, column + 1), (row + 1, column)):
                beside = ws._cells.get(place)
                words = beside.value.strip() if beside is not None and isinstance(beside.value, str) else ""
                if words and len(words) <= UNIT_WORD_LENGTH:
                    yield row, f"{text} {words}", words
                    break


def unit_note(ws, rows, columns):
    """(the words as written, the unit alone) of the first such note, or ("", "")."""
    first = next(_unit_notes(ws, rows, columns), None)
    return first[1:] if first else ("", "")


def _reads_backwards(first, second):
    """Two periods are read first to second. Sheets often put this year first, and 매출 then 전년 read that way turns a rise into a fall with nothing to show for it."""
    years = [re.search(r"(?:19|20)\d\d", name or "") for name in (first, second)]
    if all(years):
        return int(years[0].group(0)) > int(years[1].group(0))
    return bool(EARLIER_NAME.search(second or "")) and not EARLIER_NAME.search(first or "")


def in_per_cent(ws, row=None):
    """The note that covers this row says per cent, not merely that a % appears somewhere on the sheet. A note covers the rows under it. One with no figures under it is a footnote, and covers what is above. row None asks about the sheet as a whole, which only a footnote can answer."""
    if "_unit_notes" not in ws.__dict__:
        ws.__dict__["_unit_notes"] = list(_unit_notes(ws, UNIT_SCAN_ROWS, UNIT_SCAN_COLUMNS))
        ws.__dict__["_last_figure_row"] = max((place[0] for place, cell in ws._cells.items() if isinstance(cell.value, (int, float)) and not isinstance(cell.value, bool)), default=0) if ws.__dict__["_unit_notes"] else 0
    notes = ws.__dict__["_unit_notes"]
    above = [note for note in notes if row is not None and note[0] < row]
    note = above[-1] if above else next((note for note in notes if note[0] >= ws.__dict__["_last_figure_row"]), None)
    return note is not None and PER_CENT_UNIT.fullmatch(note[2]) is not None


def _year_beside_months(labels):
    """The one line for the whole year among lines for its months. Beside anything else, 연간 may be an item of its own, such as a yearly plan beside a monthly one."""
    names = [label for label in labels if isinstance(label, str)]
    years = [label for label in names if YEAR_LINE.fullmatch(label.strip())]
    months = sum(MONTH_LINE.fullmatch(label.strip()) is not None for label in names)
    return years[0] if len(years) == 1 and months >= MIN_MONTH_LINES else None



def _stated_day(text):
    """A date written out in a sentence, or None."""
    words = DAY_FIRST.search(text)
    if words:
        month = MONTH_NUMBERS.get(words.group(2)[:3].lower())
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


def _heading_spilled_over(ws, cell):
    """True when a cell sits inside a merged range that starts higher up, on a row of words only: a heading, not a name for this row."""
    origin = _merge_origin(ws, *cell)
    if origin is None or origin[0] >= cell[0]:
        return False
    row = origin[0]
    known = ws.__dict__.setdefault("_rows_of_words", {})  # asked for every line under a merged name, so each row is looked at once
    if row not in known:
        known[row] = not any(as_number(found.value) is not None for (r, _c), found in ws._cells.items() if r == row) if len(ws._cells) <= HEADING_SCAN_CELLS else not any(
            as_number(ws.cell(row, column).value) is not None for column in range(1, min(ws.max_column, HEADING_SCAN_COLUMNS) + 1))
    return known[row]


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


def _rows_in_inner_label_columns(ws, ref, labels, figures):
    """A table whose levels stand in columns of their own, one label to a line: 경북 in 시·도 with its 시·군·구 in the next column and their 행정동 in the one after, each line with its own sum. Only the outermost level that labels two lines or more is compared, as an indented pivot table is; adding all of them counts every hectare two or three times."""
    lines = {}
    for position in range(ref.start, ref.last(ws) + 1):
        beside = [ws._cells.get(ref.at(position, line)) for line in figures]
        if not any(other is not None and as_number(other.value) is not None for other in beside):
            continue
        named, said = [], None
        for index, line in enumerate(labels):
            cell = ws._cells.get(ref.at(position, line))
            if cell is not None and isinstance(cell.value, str) and cell.value.strip() and as_number(cell.value) is None:
                if "".join(cell.value.split()) != said:  # 관평로 written again in each of A, B and C is one name, not three levels
                    named.append(index)
                said = "".join(cell.value.split())
        if named:
            lines[position] = named
    if not lines or sum(len(named) == 1 for named in lines.values()) < len(lines) * ONE_LABEL_SHARE:
        return frozenset()
    # A column is a level only where it names lines of its own. 서울 beside 강남 on a group's first line names the group, and the line is 강남's, one of the districts;
    # 평촌대로 and 관평로 beside 동안로 on one switchboard's line, and nowhere else, are more roads for that line, not a level under 동안로.
    alone = {named[0] for named in lines.values() if len(named) == 1}
    level = {position: next((index for index in reversed(named) if index in alone), named[0]) for position, named in lines.items()}
    counts = Counter(level.values())
    outer = min((index for index, count in counts.items() if count >= 2), default=None)
    if outer is None or len(set(level.values())) < 2:
        return frozenset()
    return frozenset(position for position, index in level.items() if index > outer)


def _rows_indented_under_others(ws, ref, labels, figures):
    """A pivot table's lines indented under the one above: schools under regions under terms, all in one column. Only the outermost can stand side by side, and only where more than one line sits at that level. A heading with no figure beside it is not a line of the table."""
    if len(labels) > 1:
        return _rows_in_inner_label_columns(ws, ref, labels, figures)
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


def _parts_set_apart(ws, ref, labels, parts, kept):
    """True when the sheet itself marks the rows taken as parts: indented further, written after a space or a bullet, or the rows kept in bold while the parts are not. Rows that only add up may do so by chance, as 300 = 100 + 200, and then a person is asked."""
    def look(position):
        cell = next((found for line in labels if position is not None and (found := ws._cells.get(ref.at(position, line))) is not None and found.value not in (None, "")), None)
        if cell is None:
            return 0, False, False
        text = cell.value if isinstance(cell.value, str) else ""
        return cell.alignment.indent or 0, text[:1].isspace() or facts.PART_MARK.match(text.strip()) is not None, bool(cell.font is not None and cell.font.b)

    part_looks, kept_looks = [look(position) for position in parts], [look(position) for position in kept]
    if not part_looks:
        return True
    if _named_as_parts(ws, ref, labels, parts, kept):
        return True
    if max(indent for indent, _spaced, _bold in part_looks) > min(indent for indent, _spaced, _bold in kept_looks or [(0, False, False)]):
        return True
    if any(spaced for _indent, spaced, _bold in part_looks) and not all(spaced for _indent, spaced, _bold in kept_looks):
        return True
    return any(bold for _indent, _spaced, bold in kept_looks) and not any(bold for _indent, _spaced, bold in part_looks)


def _first_parts(rows, kept, row_of):
    """{parent, parts}: the first row taken as a sum and the names under it, so a person asked can see whether 서울 really is 부산 and 대구 together."""
    place = lambda item: row_of.get(id(item)) or 0  # noqa: E731
    parts = [item for item in rows if id(item) not in kept]
    above = [item for item in rows if id(item) in kept and place(item) < place(parts[0])]
    if not above:
        return {}
    parent = max(above, key=place)
    following = [item for item in rows if id(item) in kept and place(item) > place(parent)]
    end = min((place(item) for item in following), default=float("inf"))
    under = [item.label for item in parts if place(parent) < place(item) < end]
    return {"parent": parent.label, "parts": ", ".join(under[:4]) + (" …" if len(under) > 4 else "")}


def _named_as_parts(ws, ref, labels, parts, kept):
    """True when every part's name says whose part it is: 611-01 under 611, or 서울 강남 under 서울. The row it belongs to is the nearest row kept above it."""
    def label(position):
        for line in labels:
            cell = ws._cells.get(ref.at(position, line)) if position is not None else None
            if cell is not None and cell.value not in (None, ""):
                return str(cell.value).strip()
        return ""

    above = sorted(position for position in kept if position is not None)
    for part in parts:
        parent = max((position for position in above if part is not None and position < part), default=None)
        if parent is None or not _name_under(label(parent), label(part)):
            return False
    return True


def _name_under(parent, part):
    """611-01 가정용 under 611 급수수익, 1-2 under 1, 서울 강남 under 서울."""
    if not parent or not part:
        return False
    codes = [PART_CODE.match(text) for text in (parent, part)]
    if all(codes) and any(c.isdigit() for c in codes[0].group(1)):
        whole, piece = codes[0].group(1), codes[1].group(1)
        if len(piece) > len(whole) and piece.startswith(whole) and piece[len(whole)] in "-.":
            return True
    return len(part) > len(parent) and part.startswith(parent) and (part[len(parent)].isspace() or part[len(parent)] in "-·(/")


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
    year = as_year(wanted)
    if year is not None and as_year(value) == year:
        return True  # 2024 asked for finds [2024] and 2024년 as well: one year written three ways
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
    numbers = [figure_number(cell.value) for cell in cells]
    if from_text is not None:
        from_text.extend(cell for cell, number in zip(cells, numbers) if number is not None and isinstance(cell.value, str))
    # Exact decimal spellings of finite decoded operands. No binary rounding
    # between row products, row sums, group sums or the final mean division.
    numbers = [Fraction(str(number)) if number is not None else None for number in numbers]
    if spec.product:
        return None if any(number is None for number in numbers) else math.prod(numbers)
    found = [number for number in numbers if number is not None]
    return sum_rationals(found) if found else None


def _years_only(ws, ref, position, specs):
    lines = [line for spec in specs for line in (spec.lines if isinstance(spec, Aggregate) else [spec])]
    return all(as_year(ws.cell(*ref.at(position, line)).value) is not None for line in lines)


def _plain_value(ws, ref, position, line, from_text):
    cell = ws.cell(*ref.at(position, line))
    number = figure_number(cell.value)
    if number is not None and isinstance(cell.value, str):
        from_text.append(cell)
    return number


def _total_kind(ws, ref, position, before, chosen=()):
    """total or subtotal if a cell before the values says so. 합계 may sit in a column left of the names, as with merged cells. A column the reference chose a row by says which rows it wants, and says nothing about them being a total of the rest."""
    for line in range(1, before):
        if line in chosen:
            continue
        kind = total_word_kind(_merged_value(ws, *ref.at(position, line)))
        if kind:
            return kind
    return None


def _several_totals(names, totals):
    """The total lines of a table that totals unlike things, as an operating statement does with Total income and Total expenses, as facts.Item each. Empty when there is one whole.
    The lines between them are income, expenses and results, so ranking them against one another says nothing."""
    found, seen = [], set()
    for name, numbers in zip(names, totals):
        label = _clean_label(name) if isinstance(name, str) else ""
        if not label or numbers[-1] is None:
            continue
        bare = lambda text: "".join(text.split()).casefold() in TOTAL_WORDS  # noqa: E731
        if label.casefold() in seen or any(item.value == numbers[-1] and (bare(label) or bare(item.label)) for item in found):  # 합계 on top and 총계 below, both the same whole
            continue
        seen.add(label.casefold())
        found.append(facts.Item(label, numbers[-1], numbers[0] if len(numbers) == 2 else None))
    if len(found) < 2:
        return []
    # 서울 합계 30 and 부산 합계 12 under 총계 42 are groups of one kind, the grand total their sum: the usual comparison holds.
    # Income and expenses do not add up to a line of the table, however close they come (33,037 against 33,118).
    for whole in found:
        rest = sum(item.value for item in found if item is not whole)
        if abs(whole.value - rest) <= abs(whole.value) * SUBTOTALS_ADD_UP:
            return []
    return found


def total_word_kind(value):
    """subtotal or total when a cell's word says its row adds up others, else None."""
    if not isinstance(value, str):
        return None
    word = "".join(value.split()).casefold()
    last = value.split()[-1].casefold() if value.split() else ""
    if word in SUBTOTAL_WORDS or (last in SUBTOTAL_WORDS and len(value.split()) <= GROUP_TOTAL_WORDS):  # 서울 소계
        return "subtotal"
    if last == "계" and len(value.split()) <= GROUP_TOTAL_WORDS and len("".join(value.split()[:-1])) >= 2:  # 수입 계 and 지출 계, the total of each part of a statement. 회계 and 생계 run the 계 on, and so do 회  계 and 통  계 spaced out to fill their cells
        return "total"
    if word in TOTAL_WORDS or TOTAL_ROW.match(value.strip()) or word.endswith(("합계", "총계", "전체")):  # 공원관리사업소 전체 is the whole office
        return "total"
    if ":" in value and TOTAL_ROW.match(value.rpartition(":")[2].strip()):  # Administered Income: Total administered income, a statement's heading before its line
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


def _without_empty_edges(wb):
    """The workbook without cells far past the last value that hold only a format. A format given to a whole column leaves one on row 1,048,576, and every line read up to the sheet's last row would then run a million times.
    Only far past: a column of formulas Excel never worked out has no values either, and right beside the figures it must still be seen and warned about."""
    for ws in wb.worksheets:
        cells = getattr(ws, "_cells", None)
        if not cells:
            continue
        filled = {key for key, cell in cells.items() if cell.value is not None}
        last_row = max((row for row, _column in filled), default=0)
        last_column = max((column for _row, column in filled), default=0)
        for merged in ws.merged_cells.ranges:  # a merged heading reaches as far as it is drawn
            if (merged.min_row, merged.min_col) in filled:
                last_row, last_column = max(last_row, merged.max_row), max(last_column, merged.max_col)
        far_row = last_row + EMPTY_EDGE_ROWS if ws.max_row > last_row + EMPTY_EDGE_ROWS else ws.max_row
        far_column = last_column + EMPTY_EDGE_COLUMNS if ws.max_column > last_column + EMPTY_EDGE_COLUMNS else ws.max_column
        for key in [key for key in cells if key[0] > far_row or key[1] > far_column]:
            del cells[key]
    return wb


class ExcelSource:
    """Read-only view of a workbook: resolves keys and renders cells as Excel displays them."""

    def __init__(self, path):
        _check_file(path, readers.ALL_TYPES, "excel")
        self.path = path
        try:
            self._origin_sha256 = automatic.digest(path)
        except OSError:
            self._origin_sha256 = None
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
        self._months_only = set()  # date lines read as months with no year: their points are called as the sheet writes them
        self._made = {}
        self._fiscal_end = None

    def fiscal_end(self, year):
        """The stated end, or None. Missing or conflicting evidence must not imply June."""
        if self._fiscal_end is None:
            self._fiscal_end = fiscal_year_end([text for ws in self.wb.worksheets for text in _fiscal_words(ws)]) or 0
        month = self._fiscal_end
        return dt.date(year, month, calendar.monthrange(year, month)[1]) if month else None

    def _fiscal_quarter_shift(self):
        """Months to move a fiscal quarter by, from the July start it is read with to the start the workbook gives."""
        self.fiscal_end(2001)
        return (self._fiscal_end - FISCAL_YEAR_ENDS[0]) if self._fiscal_end else 0

    def _load(self, data_only):
        try:
            if os.path.splitext(self.path)[1].lower() in readers.OTHER_TYPES:
                if self._converted is None:  # no formulas to tell apart from their values, so once will do
                    self._converted = readers.load(self.path)
                return self._converted
            return _without_empty_edges(openpyxl.load_workbook(self.path, data_only=data_only))
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
            area = _dynamic_area(self, defined.attr_text or "")
            if area is None:
                raise IssueError("name_not_cells", name=key)  # a constant, or a formula other than OFFSET and INDEX over counts
            ws, (left, top, right, bottom) = area
            coord = f"{get_column_letter(left)}{top}" + (f":{get_column_letter(right)}{bottom}" if (left, top) != (right, bottom) else "")
            return ws, (left, top, right, bottom), f"{ws.title}!{coord}"
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
        if isinstance(raw,float) and not math.isfinite(raw):
            return raw,None,"General",[Issue("nonfinite_value",{"cell":address})]
        if cell.data_type == "e":
            return raw, str(raw), "General", [Issue("error_value", {"cell": address, "error": str(raw)})]
        if raw is None:
            if self._has_formula(ws, row, col) and self._never_calculated():
                return None, None, "General", [Issue("formula_no_cache", {"cell": address})]
            return None, "", "General", []
        fmt = self._cell_format(cell)
        text, exact = format_value(raw, fmt, self.wb.epoch)
        color = format_color(raw, fmt) if exact else None
        if color:
            text = Shown(text, color)
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
            alone = {position: month for position, month in months.items() if position not in dates}
            if not dates and _months_in_order(alone) and not _years_across(ws, ref, min(alone)):
                # 1월 to 12월 with no year over them: the year the sheet's title or name gives, or else the months as the sheet writes them
                year = self._sheet_year(ws)
                if year is not None and any(dt.date(year.year, month, 1) > self._as_of() and self._line_has_figure(ws, ref, position) for position, month in alone.items()):
                    year = None  # tourStat..._2021-07-28 saved that day with August to December filled: the name is when it was exported, not the year of its months
                for position, month in alone.items():
                    dates[position] = (year or dt.datetime(MONTHS_ONLY_YEAR, 1, 1)).replace(month=month)
                if year is None:
                    self._months_only.add(cache)
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
        if _holds_quarters(ws, ref, ref.date_parts[1]):
            return year.replace(month=(others[0] - 1) * 3 + 1) if len(others) == 1 and 1 <= others[0] <= 4 else None
        try:
            return year.replace(month=others[0], day=others[1] if len(others) > 1 else 1)
        except ValueError:
            return None

    def _date_cell(self, ref, sheet, position):
        """(raw, text, format) of a row's date as a person should read it. A date split across cells is shown joined."""
        if ref.date_parts:
            date = self._date_from_parts(ref.sheets[sheet], ref, position)
            if _holds_quarters(ref.sheets[sheet], ref, ref.date_parts[1]):
                return date, f"{date.year} Q{(date.month - 1) // 3 + 1}", "General"
            text = date.strftime("%Y-%m-%d" if len(ref.date_parts) == 3 else "%Y-%m")
            return date, text, "General"
        return self.read(ref.sheets[sheet], *ref.at(position, ref.date))[:3]

    def _date_header(self, ws, ref, top):
        return " ".join(self._header(ws, ref, top, line) for line in (ref.date_parts or (ref.date,)))

    @staticmethod
    def _line_has_figure(ws, ref, position):
        """True when a date's row (or column) holds a number anywhere along it."""
        size = ws.max_row if ref.by_row else ws.max_column
        return any(line != ref.date and as_number(ws.cell(*ref.at(position, line)).value) is not None for line in range(1, min(size, SHEET_FIGURE_LINES) + 1))

    @staticmethod
    def _sheet_year(ws):
        """The one year a sheet's title rows or its name give, such as 2025년 월별 실적 or 2025 Sales: what months with no year of their own belong to. None when there is none, or more than one."""
        found = set()
        for row in range(1, min(ws.max_row, SHEET_YEAR_ROWS) + 1):
            for column in range(1, min(ws.max_column, SHEET_YEAR_COLUMNS) + 1):
                raw = ws.cell(row, column).value
                if isinstance(raw, str) and len(raw) <= SHEET_YEAR_LENGTH:
                    found |= {int(year) for year in SHEET_YEAR.findall(raw)}
        if not found:
            found = {int(year) for year in SHEET_YEAR.findall(ws.title)}
        return dt.datetime(found.pop(), 1, 1) if len(found) == 1 else None

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
        if cell.data_type == "e" or isinstance(cell.value, (dt.date, dt.time)) or figure_number(cell.value) is not None:
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
        fiscal_years = ws is not None and not ref.date_parts and self._fiscal_years(ws, ref)
        fiscal_quarters = ws is not None and self._in_fiscal_quarters(points, ref)
        end = self.fiscal_end(points[-1].date.year) if fiscal_years or fiscal_quarters else None
        if (fiscal_years or fiscal_quarters) and end is None:
            return [Issue("fiscal_end_unknown")]
        if fiscal_years:
            if as_of >= end:
                return []
            return [self._unfinished_issue(points, as_of)]
        if ws is not None and self._in_fiscal_quarters(points, ref) and self._fiscal_quarter_shift():  # Q1 FY24 was read as starting in July
            shift = self._fiscal_quarter_shift()
            points = [replace(point, date=_months_later(point.date, shift)) for point in points]
        # 학년도 2026 is not the calendar year 2026, so a calendar cannot say it is unfinished.
        if not SCHOOL_YEAR.search(date_header) and facts.unfinished(points, as_of):
            return [self._unfinished_issue(points, as_of)]
        return []

    def _unfinished_issue(self, points, as_of):
        params = {"date": _when(points[-1]), "saved": as_of.isoformat()}
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

    @staticmethod
    def _day(ref, i, position, date):
        """A row's date for the place a figure came from: 12월 as the sheet writes it when the months have no year, else the date."""
        if date.year == MONTHS_ONLY_YEAR:
            return " ".join(str(ref.sheets[i].cell(*ref.at(position, ref.date)).value).split())
        return date.date().isoformat()

    def _latest_date_text(self, ref):
        """{{최신:시트!A+B}}: the newest date of a date split across cells."""
        date, i, position = self._latest_rows(ref)[-1]
        return self._date_cell(ref, i, position)[1], f"{self._day(ref, i, position, date)} · {ref.sheets[i].title}!{_line_name(ref, ref.date)}", self._row_notes(ref, date)

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
        return ref.sheets[i], row, column, f"{self._day(ref, i, position, newest)} · {cell}", self._row_notes(ref, newest)

    def _header(self, ws, ref, top, line):
        """The nearest filled cell before the first data row or column, or the line's name when the sheet has no header."""
        found = ((position, _clean_label(text)) for position in range(top - 1, 0, -1)
                 if (text := self.read(ws, *ref.at(position, line))[1]) and "," not in text.replace(", ", "")[:6] and (as_number(ws.cell(*ref.at(position, line)).value) is None or as_year(ws.cell(*ref.at(position, line)).value)))
        names, useful_seen = [], 0
        for place, text in found:  # the nearest and two that are not codes are all it uses: a table that starts on row 5,000 need not read every row above
            names.append((place, text))
            useful_seen += not ID_CODE.fullmatch(text)
            if useful_seen >= 2:
                break
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
        first, last = self._day(ref, picked[0][1], picked[0][2], picked[0][0]), self._day(ref, picked[-1][1], picked[-1][2], picked[-1][0])
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
            fmt='0.0%' if found.percent else self._cell_format(ref.sheets[found.points[-1].sheet].cell(*ref.at(found.points[-1].row,found.line))) if found.line is not None else 'General'
            if fmt in ('General','@','0'):fmt='General'
            if found.per_cent_unit and '%' not in fmt:fmt+='"%"'
            table.append([self._date_cell(ref, point.sheet, point.row), (point.value, show(point.value), fmt)])
        first, last = _when(picked[0]), _when(picked[-1])
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
        unread = [raw for date, i, position in self._all_rows(ref) if date <= rows[-1][0] and _unread_figure(raw := ref.sheets[i].cell(*ref.at(position, line)).value)] + self._unread_dates(ref)
        for date, i, position in rows:
            cell = ref.sheets[i].cell(*ref.at(position, line))
            number = figure_number(cell.value)
            if number is not None:
                points.append(facts.Point(date, Fraction(str(number)), position, i))
                if isinstance(cell.value, str):
                    from_text.append(cell)
        # Anywhere in the series, not only at the newest: two tables on one sheet would otherwise be compared with each other.
        repeated = sorted(date for date, times in Counter(p.date for p in points).items() if times > 1)
        if repeated:
            raise IssueError("latest_duplicates", date=repeated[-1].date().isoformat(), count=len(repeated))
        points = self._with_period_names(points, ref)
        notes = [Issue("number_from_text", {"cell": f"{from_text[-1].parent.title}!{from_text[-1].coordinate}", "text": str(from_text[-1].value), "count": len(from_text)})] if from_text else []
        if unread:
            notes.append(Issue("figures_not_read", {"count": len(unread), "text": str(unread[-1]).strip()}))
        if len(points) >= 3 and points[-1].value == 0 and points[-2].value != 0:
            notes.append(Issue("latest_zero", {"date": _when(points[-1])}))
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
        found.per_cent_unit = "%" not in fmt and in_per_cent(top_ws, ref.at(top, line)[0])
        return found

    def _unread_dates(self, ref):
        """Cells between the first and last date that hold digits but no date, such as 2026_: their rows are left out of every period, and that must not go unsaid."""
        if ref.date_parts:
            return []
        found = []
        for ws in ref.sheets:
            dated = self._dated_rows(ws, ref)
            if dated:
                at = {position for _date, position in dated}
                found += [raw for position in range(dated[0][1], dated[-1][1] + 1) if position not in at and _unread_figure(raw := ws.cell(*ref.at(position, ref.date)).value)]
        return found

    def _with_period_names(self, points, ref):
        """A financial year or its quarter is read as a date so that it can be sorted, but a sentence must call it what the sheet does."""
        if not points or ref.date_parts:
            return points
        if self._fiscal_years(ref.sheets[points[0].sheet], ref):
            return [replace(point, label=f"{point.date.year - 1}-{point.date.year % 100:02d}") for point in points]
        raws = [ref.sheets[point.sheet].cell(*ref.at(point.row, ref.date)).value for point in points]
        if any((ws.title, ref.by_row, ref.date, ref.start, ref.end, ref.date_parts) in self._months_only for ws in ref.sheets):
            return [replace(point, label=" ".join(str(raw).split())) for point, raw in zip(points, raws)]  # 12월, not December of a year no one wrote
        if self._in_fiscal_quarters(points, ref):
            return [replace(point, label=" ".join(raw.split())) for point, raw in zip(points, raws)]
        return points

    def _in_fiscal_quarters(self, points, ref):
        """True when every point is a quarter of a financial year, such as Q1 FY24."""
        if not points or ref.date_parts:
            return False
        raws = [ref.sheets[point.sheet].cell(*ref.at(point.row, ref.date)).value for point in points]
        return all(isinstance(raw, str) and _fiscal_quarter(DATE_MARK.sub("", raw.strip())) for raw in raws)

    def _summed_series(self, ref):
        """Rows added up per calendar day: the person asked for this by writing 합계( ) or 건수( ). Nothing is summed unasked."""
        spec = ref.values[0]
        if len(ref.values) != 1:
            raise IssueError("calc_columns", key=spec.word)
        groups, owners, from_text, unread = {}, {}, [], []
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
                    if spec.kind in ("sum", "mean"):
                        unread += [raw for raw in (ws.cell(*ref.at(position, line)).value for line in spec.lines) if _unread_figure(raw)]
                    continue
                day = dt.datetime.combine(current.date(), dt.time())
                if owners.setdefault(day, i) != i:
                    raise IssueError("latest_sheets_overlap", date=day.date().isoformat(), first=ref.sheets[owners[day]].title, second=ws.title)
                group = groups.setdefault(day, [0, 0, i, current_position, None])
                group[0] = sum_rationals((group[0],number))
                group[1] += 1
                if current.time() != dt.time():
                    moment = current.hour * 60 + current.minute
                    group[4] = max(group[4] or 0, moment)
        if not groups:
            raise IssueError("latest_no_dates", sheet=ref.sheets[0].title, column=_line_name(ref, ref.date))
        unread += self._unread_dates(ref)
        days = sorted(groups)
        points = [facts.Point(day, groups[day][0], groups[day][3], groups[day][2]) for day in days]
        top_ws = ref.sheets[groups[days[0]][2]]
        top = self._dated_rows(top_ws, ref)[0][1]
        names = [self._header(top_ws, ref, top, line) for line in spec.lines]
        name = _aggregate_name(spec, names)
        if spec.kind in ("mean", "rate"):
            points = [facts.Point(p.date, Fraction(p.value) / groups[p.date][1], p.row, p.sheet) for p in points]
        # Keep the exact aggregate through further analysis and text/table formatting.
        # Numeric chart data is converted only at the container boundary and checked there.
        points = self._with_period_names(points, ref)
        newest = points[-1]
        found = Series(ref, points, (newest.sheet, newest.row), None if spec.kind in ("count", "rate") or spec.product else spec.lines[0], name, [])
        found.notes = self._period_issues(points, self._date_header(top_ws, ref, top), top_ws, ref) + _group_notes(days, groups)
        if len(points) >= 3 and points[-1].value == 0 and points[-2].value != 0:
            found.notes.append(Issue("latest_zero", {"date": _when(points[-1])}))
        found.percent = spec.kind == "rate"
        first_fmt = self._cell_format(top_ws.cell(*ref.at(top, spec.lines[0]))) if spec.kind in ("sum", "mean") and spec.lines else ""
        found.rate = spec.kind == "rate" or (spec.kind != "count" and self._is_rate_series(top_ws, ref, top, spec.lines[0], name, first_fmt))
        found.scale = 100 if spec.kind == "rate" or "%" in first_fmt else 1  # 0.92 to 0.93 is one point, not 0.01
        found.per_cent_unit = spec.kind in ("sum", "mean") and in_per_cent(top_ws, ref.at(top, spec.lines[0])[0])
        usual = _usual_rows(days, groups)
        if usual is not None and groups[days[-1]][1] < usual and spec.kind in ("sum", "count"):
            found.cut_short = (groups[days[-1]][1], usual)
        if from_text:
            found.notes.append(Issue("number_from_text", {"cell": f"{from_text[-1].parent.title}!{from_text[-1].coordinate}", "text": str(from_text[-1].value), "count": len(from_text)}))
        if unread:
            found.notes.append(Issue("figures_not_read", {"count": len(unread), "text": str(unread[-1]).strip()}))
        return found

    @staticmethod
    def _is_rate_series(ws, ref, top, line, header, fmt):
        """True when the figures are rates, scores or changes already, so they move in points. Judged from the heading, the format, the sheet's name, and plain words above the table."""
        if "%" in fmt or RATE_HEADING.search(header) or RATE_TITLE.search(ws.title) or in_per_cent(ws, ref.at(top, line)[0]):
            return True
        above =" ".join(str(ws.cell(*ref.at(position, other)).value or "") for position in range(top - 1, 0, -1) for other in (line, 1))
        return RATE_EXPLICIT.search(above) is not None

    def _series_words(self, found, lang):
        note, read, pack = self._series_words_plain(found, lang)
        # Only an actual, unaggregated source cell can colour the current value.
        # Match the complete generated head, never every occurrence of its digits.
        if found.line is not None and not found.ref.aggregated and not found.percent and found.points:
            point = found.points[-1]
            raw, shown, _fmt, _issues = self.read(found.ref.sheets[point.sheet], *found.ref.at(point.row, found.line))
            color = getattr(shown, "color", None)
            value = found.show(self)(point.value)
            head = f"{facts.point_text(point.date, found.points, lang)} {found.label} {value}"
            if color and raw == point.value and note.count(head) == 1:
                end = note.index(head) + len(head)
                note = Shown(note, spans=((end - len(value), end, color),))
        return note, read, pack

    def _series_words_plain(self, found, lang):
        """(the report sentence, the reading facts, the facts for a model) for a series, with what a reader must know first put in the words themselves.

        Warnings stay in the app. Only the sentence reaches the slide, so a period that is not whole has to say so there.
        """
        points, show, label = found.points, found.show(self), found.label
        if any(issue.code == "fiscal_end_unknown" for issue in found.notes):
            lead = self._leads(found, lang)
            head = f"{facts.point_text(points[-1].date, points, lang)} {label} {show(points[-1].value)}."
            warning = facts.Fact("read_lead", lead, ())
            value = facts.Fact("value", f"{facts.point_text(points[-1].date, points, lang)} {label}: {show(points[-1].value)}", tuple(facts.NUMBER.findall(show(points[-1].value))))
            return f"{lead} {head}", [warning], [warning, value]
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
        earlier = facts.rate_readings(ended, show, lang, found.scale) if found.rate else facts.readings(ended, show, lang)
        # One statement: apart, -25.1% against the year before reads as being about the period still running.
        words = " ".join([settled] + [fact.text for fact in earlier])
        reading = facts.Fact("read_ended", words, tuple(facts.NUMBER.findall(words)), label, facts.point_text(ended[-1].date, points, lang))
        warning = facts.Fact("read_lead", lead, tuple(facts.NUMBER.findall(lead)))
        value = facts.Fact("value", f"{facts.point_text(now.date, points, lang)} {label}: {show(now.value)}", tuple(facts.NUMBER.findall(show(now.value))))
        return f"{lead} {head}", [reading], [warning, reading, value]

    def _stated_as_of(self):
        """The date a workbook says its data runs to, on any sheet: Data as at 30-Sep-2025. None when it does not say."""
        if not hasattr(self, "_as_at"):
            self._as_at = None
            for ws in self.wb.worksheets:
                columns = min(ws.max_column, AS_AT_COLUMNS)
                for row in range(1, min(ws.max_row, AS_AT_ROWS) + 1):
                    for column in range(1, columns + 1):
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
            if issue.code == "fiscal_end_unknown":
                parts.append("회계연도 종료월을 확정할 수 없어 기간 완료 여부를 판단하거나 다른 기간과 비교하지 않았습니다. 원본에 종료월을 명시하세요." if ko else
                             "The fiscal year end is unspecified or conflicting, so completion and period comparisons are not determined. State the year end in the workbook.")
            elif issue.code == "latest_text_newer":
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
        when = _when(newest)
        cell = found.cell_text(newest)
        if name in ("peak", "low"):
            top = facts.extreme(points, int(count or 30), highest=name == "peak")
            return self._date_cell(ref, top.sheet, top.row)[1], f"{_when(top)} · {found.cell_text(top)}", issues
        if name in ("읽는법", "reading"):
            note, read, _pack = self._series_words(found, "ko" if name == "읽는법" else "en")
            read = [fact for fact in read if fact.text not in note] or read  # a rate's sentence already holds its change, and the slide shows both
            if not read:
                return None, f"{when} · {cell}", issues + [Issue("calc_no_base", {"date": when})]
            return " ".join(fact.text for fact in read), f"{when} · {cell}", issues
        if name in ("해설", "note"):
            return self._series_words(found, "ko" if name == "해설" else "en")[0], f"{when} · {cell}", issues
        if any(issue.code == "fiscal_end_unknown" for issue in found.notes):
            raise IssueError("fiscal_end_unknown")
        if name == "avg":
            size = int(count or 7)
            base = facts.mean_before(points, size) if size >= 1 else None
            against = f"{_when(points[-size - 1])} ~ {_when(points[-2])}" if base is not None else ""
        else:
            before = facts.previous(points) if name == "prev" else facts.week_ago(points)
            base = before.value if before else None
            against = _when(before) if before else ""
        if found.rate and base is not None:
            lang = "ko" if HANGUL.search(word) else "en"
            return facts.points_text(facts.difference(newest.value,base), lang, found.scale), f"{when} ↔ {against} · {cell}", issues
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
        origin_positions = {}
        unread, places, total_names = [], {}, []
        heading = self._heading_row(ws, ref, labels, specs)
        outer = _rows_of_an_outer_level(ws, ref, labels)
        chosen = {line for spec in specs if isinstance(spec, Aggregate) for line, _word in spec.filters}
        indented = _rows_indented_under_others(ws, ref, labels, [spec.lines[0] if isinstance(spec, Aggregate) else spec for spec in specs])
        for position in range(ref.start, ref.last(ws) + 1):
            if position == heading or position in indented:
                continue
            texts_so_far = len(from_text)
            numbers = [_aggregate_value(ws, ref, position, spec, from_text) if isinstance(spec, Aggregate) else _plain_value(ws, ref, position, spec, from_text) for spec in specs]
            numbers=[Fraction(str(n)) if n is not None else None for n in numbers]
            if all(number is None for number in numbers):
                wanted = [line for spec in specs if not isinstance(spec, Aggregate) or spec.kind in ("sum", "mean") for line in (spec.lines if isinstance(spec, Aggregate) else [spec])]  # 건수( ) counts rows and reads no figure
                unread += [raw for raw in (ws.cell(*ref.at(position, line)).value for line in wanted) if _unread_figure(raw)]
                continue
            if first is None and _years_only(ws, ref, position, specs) and not _years_only(ws, ref, position + 1, specs):
                del from_text[texts_so_far:]
                continue  # 2023, 2024 heading the value columns, not an item. 1987 and 1994 under 건수 on row after row are counts
            first = first or position
            kind = _total_kind(ws, ref, position, lower, chosen)
            if kind == "total":
                totals.append(numbers)
                total_names.append(self._label(ws, ref, position, labels))
                continue
            if kind == "subtotal" or position in outer:
                continue
            label = self._label(ws, ref, position, labels)
            if not label:
                unlabeled += 1
                continue
            if any(isinstance(spec, Aggregate) for spec in specs) and label in items:
                origin_key = label
                places.setdefault(label, position)
                items[label] = [sum_rationals((a or 0,b or 0)) if a is not None or b is not None else None for a, b in zip(items[label], numbers)]
                counts[label] = [c + (b is not None) for c, b in zip(counts[label], numbers)]
            else:
                key = label if label not in items else (label, position)
                origin_key = key
                places[key] = position
                items[key] = numbers
                counts[key] = [int(b is not None) for b in numbers]
                order.append(key)
            origin_positions.setdefault(origin_key, []).append(position)
        order = [key for key in order if items[key][-1] is not None]
        named = _names_told_apart(ws, ref, labels, specs, order, places)
        if named:  # 2,001이상 of 일반용 and of 대중탕용: one name for two rates reads as one
            items = {named.get(key, key): numbers for key, numbers in items.items()}
            counts = {named.get(key, key): numbers for key, numbers in counts.items()}
            places = {named.get(key, key): position for key, position in places.items()}
            order = [named.get(key, key) for key in order]
            origin_positions = {named.get(key, key): positions for key, positions in origin_positions.items()}
        for key in order:  # a mean or a rate is worked out once every row of the item is in
            items[key] = [Fraction(value) / count if value is not None and isinstance(spec, Aggregate) and spec.kind in ("mean", "rate") else value
                          for value, count, spec in zip(items[key], counts[key], specs)]
        year_line = _year_beside_months(order)
        if year_line:  # 1월 to 12월 and 연간: the year is their whole, not a thirteenth month to rank them against
            order.remove(year_line)
            totals.append(items[year_line])
            total_names.append(year_line)
        if len(order) < 2:
            raise IssueError("cross_no_items", sheet=ws.title, column=_line_name(ref, labels[0]))
        rows = [facts.Item(key if isinstance(key, str) else key[0], items[key][-1], items[key][0] if len(specs) == 2 else None) for key in order]
        row_of = {id(item): places.get(key) for item, key in zip(rows, order)}
        audit_origins = {id(item): origin_positions.get(key, []) for item, key in zip(rows, order)}
        first_line = specs[-1].lines[0] if isinstance(specs[-1], Aggregate) else specs[-1]
        header = self._header(ws, ref, first or ref.start, first_line)
        fmt_cell = ws.cell(*ref.at(first or ref.start, first_line))
        fmt = self._cell_format(fmt_cell)
        last_spec = specs[-1]
        priced = facts.is_price(header) and not (isinstance(last_spec, Aggregate) and last_spec.kind == "sum")  # each row's fee: 합계( ) written out asks to add them anyway
        is_rate_kind = isinstance(last_spec, Aggregate) and last_spec.kind == "rate"
        if isinstance(last_spec, Aggregate) and (last_spec.kind in ("mean", "rate", "count") or last_spec.filters):
            header = _aggregate_name(last_spec, [header])
        if is_rate_kind:
            fmt = "0.0%"
        elif isinstance(last_spec, Aggregate) and last_spec.kind == "count":
            fmt = "General"  # rows counted by a date column are a number of rows, not a date: 2, not 01월 02일
        issues = [Issue("cross_total_found", {"label": year_line})] if year_line else []
        basis = ""
        if unread:
            issues.append(Issue("figures_not_read", {"count": len(unread), "text": str(unread[-1]).strip()}))
        if len(specs) == 2:
            first_name = self._header(ws, ref, first or ref.start, specs[0].lines[0] if isinstance(specs[0], Aggregate) else specs[0])
            second_name = self._header(ws, ref, first or ref.start, first_line)
            if _reads_backwards(first_name, second_name):
                issues.append(Issue("cross_backwards", {"first": first_name, "second": second_name}))
            planned = BUDGET_NAME.search(first_name)
            years = [YEAR_IN_HEADING.findall(name) for name in (first_name, second_name)]
            if planned and not BUDGET_NAME.search(second_name) and (not all(years) or years[0] == years[1]):
                # 2024-25 Published Budget beside 2024-25 Actual: the same year, what was planned and what happened. A 2026 budget beside 2025 actuals is not that
                basis = "budget" if re.search(r"예산|budget", planned.group(0), re.I) else "target" if re.search(r"목표|target", planned.group(0), re.I) else "plan"
        whole_lines = _several_totals(total_names, totals)
        if whole_lines:
            issues.append(Issue("cross_several_totals", {"count": len(whole_lines), "names": ", ".join(item.label for item in whole_lines[:3])}))
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
        per_cent = not is_rate_kind and in_per_cent(ws, ref.at(first or ref.start, first_line)[0])
        # A price is not added, but it moves in percent: +300포인트 says nothing about 1,000 to 1,300.
        rate = is_rate_kind or per_cent or (not priced and ((isinstance(last_spec, Aggregate) and last_spec.kind == "mean") or facts.is_rate(f"{header} {ws.title}", fmt)))
        levels = facts.top_level(rows)
        if levels:  # 자산 and its parts, numbered by level: only the top level can stand side by side, and none of it adds up to a whole
            rows = levels
            issues.append(Issue("cross_outline", {"count": len(rows)}))
        nested = facts.without_parts(rows) if not levels else None
        if nested:  # a row that is the sum of the rows under it: they are its parts, not its rivals
            kept = {id(item) for item in nested}
            marked = _parts_set_apart(ws, ref, labels, [row_of.get(id(item)) for item in rows if id(item) not in kept], [row_of.get(id(item)) for item in nested])
            params = {"count": len(nested), "by_sum": not marked}
            if not marked:
                params.update(_first_parts(rows, kept, row_of))
            rows = nested
            issues.append(Issue("cross_nested", params))
        additive = not levels and not per_cent and not priced and facts.is_additive(f"{header} {ws.title}", fmt, [item.value for item in rows])
        total = before_total = None
        added = facts.exact_total(item.value for item in rows)
        has_total = False
        if additive:
            total = added
            matching = next((line for line in totals if line[-1] is not None and abs(line[-1] - added) <= abs(line[-1]) * TOTAL_TOLERANCE), None)
            stated = matching or stated
            if stated and stated[-1] is not None:
                if matching or abs(stated[-1] - added) <= abs(stated[-1]) * TOTAL_TOLERANCE:
                    total = stated[-1]
                    has_total = True
                elif whole_lines:  # 수입 계 and 지출 계: no one line is the sum of everything listed, and the sentences give each total line instead
                    total = None
                else:
                    issues.append(Issue("cross_total_differs", {"stated": facts.plain_number(stated[-1]), "added": facts.plain_number(added)}))
                    total = None
            else:
                issues.append(Issue("cross_no_total", {"added": facts.plain_number(added)}))
            if len(specs) == 2:
                before_total = facts.exact_total(item.before for item in rows if item.before is not None)
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
        output_format='General' if plain else fmt
        if per_cent and '%' not in fmt:output_format+='"%"'
        return Comparison(rows, header, total, before_total, additive, show, issues, source, label_header, has_total, rate, 100 if "%" in fmt else 1, basis, whole_lines, audit_origins,output_format)

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
        """The most specific name a row has: the last filled cell of its label columns.
        A heading merged down over the first line of figures is not that line's name: 구 분 over A3:C5, with the sheet's total of 7,164 on row 5."""
        for line in reversed(lines):
            here = ref.at(position, line)
            if _heading_spilled_over(ws, here):
                continue
            value = _merged_value(ws, *here)
            # A bare number is a figure or a code, not a name a reader could use.
            if value not in (None, "") and as_number(value) is None:
                return _clean_label(value) if isinstance(value, str) else self.read(ws, *ref.at(position, line))[1]
        return ""

    def _comparison_text(self, ref, name):
        found = self._comparison(ref)
        lang = "ko" if name in ("읽는법", "해설") else "en"
        if name in ("읽는법", "reading"):
            text = " ".join(fact.text for fact in facts.item_readings(found.items, found.show, lang, found.total, found.additive, found.before_total, limit=5, stated=found.stated, rate=found.rate, scale=found.scale, basis=found.basis, wholes=found.wholes))
            period = _filtered_period(ref, lang)
            if period and text:  # items of one month under a sentence about one day: without it, 22 of 36 reads as the whole survey
                text = f"{period} 기준, {text}" if lang == "ko" else f"In {period}, {text[0].lower() + text[1:] if text.startswith('The ') else text}"
        else:
            text = facts.item_sentence(found.items, found.label, found.show, lang, found.total, found.additive, found.before_total, found.stated, found.rate, found.scale, found.basis, found.wholes)
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
        return format_value(facts.quotient(item.value,found.total), "0.0%")[0], found.source, found.issues

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
            fmt=found.number_format
            line = [(item.label, item.label, "General"), (item.value, found.show(item.value), fmt)]
            if with_share:
                share = facts.quotient(item.value,found.total)
                line.append((share, format_value(share, "0.0%")[0], "0.0%"))
            table.append(line)
        issues = list(found.issues)
        if len(picked) != count:
            issues.append(Issue("latest_fewer", {"wanted": count, "found": len(picked)}))
        return table, found.source, issues

    def insight_request(self, key, template_text=""):
        """First stage for {{통찰:시트!A,B}}. Returns (InsightRequest or None, source, issues)."""
        ref = self._parse_ref(INSIGHT.fullmatch(key.translate(SMART_QUOTES)).group(1), key)
        words = [template_text] + [vote for sheet in self.wb.worksheets for vote in _votes([sheet.title]) + _sheet_words(sheet)]  # sheets named Question1 to Question8 over Korean answers are Korean
        if self._is_comparison(ref):
            found = self._comparison(ref)
            lang = facts.detect_language(words + [found.header])
            pack = _with_sheet(facts.item_pack(found.items, found.show, lang, found.total, found.additive, found.before_total, found.stated, found.rate, found.scale, found.basis, found.wholes), ref, lang, found.label)
            fallback = facts.item_sentence(found.items, found.label, found.show, lang, found.total, found.additive, found.before_total, found.stated, found.rate, found.scale, found.basis, found.wholes)
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
        source = f"{_when(newest)} · {found.cell_text(newest)}"
        fallback = " ".join([note] + [fact.text for fact in read if fact.text not in note])  # a rate's sentence already holds its readings
        request = InsightRequest(key, lang, pack, facts.fingerprint(pack, lang), fallback, source)
        return request, source, found.notes + issues

    def text_value(self, key):
        """Resolve a text placeholder. Returns (text or None, source, issues)."""
        if key.startswith("layout:"):
            return self._layout_value(key, False)
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
        if key.startswith("layout:"):
            return self._layout_value(key, True)
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

    def _layout_value(self, key, grid):
        try:
            ws, view, start, stop, lang, part = layouts.resolve(self, key)
            if (part == "grid") != grid:
                raise ValueError("wrong layout reference kind")
        except ValueError:
            raise IssueError("bad_ref", ref=key)
        refs = sorted({ref for line in view.rows[start:stop] for token in line for ref in token.refs})
        where = f"{ws.title}!" + ",".join(refs)
        issues = [Issue("layout_review", {"detail": view.note(lang)})]
        if not grid:
            return view.note(lang), where, issues
        rows, notes = layouts.table(self, ws, view, start, stop, lang)
        return rows, where, issues + notes

    def audit_evidence(self, key, grid):
        """Source recipes, not production values. An independent reader evaluates them."""
        if key.startswith("layout:"):
            ws, view, start, stop, lang, _part = layouts.resolve(self, key)
            return {"kind": "layout", "sheet": ws.title, "view": view.id, "start": start, "stop": stop, "lang": lang,
                    "rows": [[token.certificate() for token in row] for row in view.rows[start:stop]]}
        def cell(ws, row, column, op="cell"):
            return {"op": op, "sheet": ws.title, "cell": ws.cell(row, column).coordinate}
        def date_cell(ws, ref, position, raw):
            if ref.date_parts:
                return {"op": "date_parts", "cells": [cell(ws, *ref.at(position, line)) for line in ref.date_parts],
                        "quarter": _holds_quarters(ws, ref, ref.date_parts[1])}
            return cell(ws, *ref.at(position, ref.date), "date" if isinstance(raw, (dt.date, dt.datetime)) else "cell")
        condition_domains={}
        def term(ws, ref, spec, positions):
            if not isinstance(spec, Aggregate):
                return cell(ws, *ref.at(positions[0], spec)) if len(positions) == 1 else None
            def condition(position, line, wanted):
                origin = _merge_origin(ws, *ref.at(position, line)) or ref.at(position, line)
                result = cell(ws, *origin, "date" if isinstance(ws.cell(*origin).value,(dt.date,dt.datetime)) else "cell")
                domain_key=(ws.title,line)
                if domain_key not in condition_domains:
                    if ref.last(ws)-ref.start+1>100000:raise ValueError('filter evidence exceeds independent coverage budget')
                    domain=[ws.cell(*(_merge_origin(ws, *ref.at(p, line)) or ref.at(p, line))) for p in range(ref.start, ref.last(ws) + 1)]
                    condition_domains[domain_key]=(tuple(c.coordinate for c in domain),tuple(c.coordinate for c in domain if isinstance(c.value,(dt.date,dt.datetime))))
                domain,date_domain=condition_domains[domain_key]
                result.update(wanted=wanted, domain=domain, date_domain=date_domain)
                return result
            return {"op": "aggregate", "kind": spec.kind, "product": spec.product,
                    "rows": [{"cells": [cell(ws, *ref.at(p, line)) for line in spec.lines],
                              "filters": [condition(p, line, wanted) for line, wanted in spec.filters],
                              "numerator": condition(p, *spec.numerator) if spec.numerator else None} for p in positions]}
        def with_domain(recipe,ref):
            # Choosing latest/ranked rows depends on candidates that may not appear in the output.
            # Audit their formula caches too; a stale date must not hide the true newest row.
            domain_key=('audit_domain',ref)
            if domain_key not in self._made:
                lines=set(ref.label_lines)|set(ref.date_parts)
                if ref.date is not None:lines.add(ref.date)
                for spec in ref.specs:
                    if isinstance(spec,Aggregate):
                        lines.update(spec.lines);lines.update(line for line,_word in spec.filters)
                        if spec.numerator:lines.add(spec.numerator[0])
                    else:lines.add(spec)
                domain=[{'op':'selection_domain','sheet':ws.title,'by_row':ref.by_row,
                         'lines':sorted(lines),'start':1,'stop':ref.last(ws)} for ws in ref.sheets]
                self._made[domain_key]=domain
            if len(recipe)>1 and recipe[1] and isinstance(recipe[1][0],dict):
                recipe[1][0]['selection_domain']=self._made[domain_key]
            return recipe
        if is_latest_key(key):
            ref, count = self._latest_request(key)
            recipe = [[None] * len(grid[0])]
            if not ref.aggregated:
                for _date, i, p in self._latest_rows(ref)[-count:]:
                    ws = ref.sheets[i]
                    recipe.append([date_cell(ws, ref, p, grid[len(recipe)][0][0])] + [cell(ws, *ref.at(p, line)) for line in ref.plain])
            else:
                for point in self._series(ref).points[-count:]:
                    ws = ref.sheets[point.sheet];dated = {p: date for date, p in self._dated_rows(ws, ref)}
                    positions = [];current = None
                    for p in range(ref.start, ref.last(ws) + 1):
                        if p in dated: current = dated[p]
                        elif ws.cell(*ref.at(p, ref.date)).value not in (None, ""): current = None
                        if current is not None and current.date() == point.date.date(): positions.append(p)
                    recipe.append([date_cell(ws, ref, point.row, grid[len(recipe)][0][0]), term(ws, ref, ref.values[0], positions)])
            return with_domain(recipe,ref)
        if is_rank_key(key):
            match = RANK.fullmatch(key.translate(SMART_QUOTES));ref = self._parse_ref(match.group(3), key)
            found = self._comparison_for_rank(match.group(3), key)
            recipe = [[None] * len(grid[0])]
            denominator = None
            if RANK_NAMES[match.group(1).lower()] == "share":
                parts = [term(ref.sheets[0], ref, ref.specs[-1], found.origin_rows.get(id(item), [])) for item in found.items]
                if not found.stated:
                    denominator = {"op": "sum_terms", "terms": parts} if all(parts) else {"op": "unavailable"}
                else:
                    ws = ref.sheets[0];spec = ref.specs[-1]
                    lower = min(s.lines[0] if isinstance(s, Aggregate) else s for s in ref.specs)
                    chosen = {line for s in ref.specs if isinstance(s, Aggregate) for line, _word in s.filters}
                    candidates = []
                    for p in range(ref.start, ref.last(ws) + 1):
                        if _total_kind(ws, ref, p, lower, chosen) == "total":
                            value = _aggregate_value(ws, ref, p, spec) if isinstance(spec, Aggregate) else figure_number(ws.cell(*ref.at(p, spec)).value)
                            if value == found.total: candidates.append(term(ws, ref, spec, [p]))
                    denominator = candidates[0] if len(candidates) == 1 else {"op": "unavailable"}
            for item in found.ranked()[:int(match.group(2) or 5)]:
                positions = found.origin_rows.get(id(item), [])
                if not positions: return None
                value = term(ref.sheets[0], ref, ref.specs[-1], positions)
                ws = ref.sheets[0]
                label = {"op": "label", "cells": [cell(ws, *(_merge_origin(ws, *ref.at(positions[0], line)) or ref.at(positions[0], line))) for line in ref.label_lines]}
                figure_lines = {line for spec in ref.specs for line in (spec.lines if isinstance(spec, Aggregate) else [spec])}
                first_label = min(ref.label_lines)
                label["context"] = [{**cell(ws, *(_merge_origin(ws, *ref.at(positions[0], line)) or ref.at(positions[0], line))),
                                     "before": line < first_label}
                                    for line in range(1, first_label + NAME_LOOK_RIGHT + 1) if line not in figure_lines]
                recipe.append([label, value] + ([{"op": "ratio", "numerator": value, "denominator": denominator}] if denominator is not None else []))
            return with_domain(recipe,ref)
        ws, (left, top, right, bottom), _where = self.locate(key)
        recipe = []
        for r in range(top, bottom + 1):
            line = []
            for c in range(left, right + 1):
                rr, cc = r, c
                if ws.cell(r, c).value in (None, ""):
                    anchor = _merge_origin(ws, r, c)
                    if anchor and not (top <= anchor[0] <= bottom and left <= anchor[1] <= right): rr, cc = anchor
                raw = grid[r-top][c-left][0]
                line.append(cell(ws, rr, cc, "date" if isinstance(raw, (dt.date, dt.datetime)) else "cell"))
            recipe.append(line)
        return recipe

    def text_evidence(self, key):
        """Explicit input values for a text fill; saved wording is checked separately."""
        expected = None
        try:
            if key.startswith('layout:'):
                table_key=key.rpartition('!')[0]+'!grid'
                grid, _where, _issues = self.range_cells(table_key)
                expected=self._layout_value(key,False)[0]
                return [{'source':self,'grid':grid,'recipe':self.audit_evidence(table_key,grid),'expected':expected}]
            elif is_calc_key(key) or is_insight_key(key):
                match = (CALC if is_calc_key(key) else INSIGHT).fullmatch(key.translate(SMART_QUOTES))
                body = match.group(3) if is_calc_key(key) else match.group(1)
                ref = self._parse_ref(body,key)
                if self._is_comparison(ref):
                    count = len(self._comparison(ref).items)
                    table_key = 'top'+str(max(count,1))+':'+body
                else:
                    count = len(self._series(ref).points)
                    if count>100000:raise ValueError('text calculation exceeds independent evidence budget')
                    table_key = 'latest'+str(max(count,1))+':'+body
                grid, _where, _issues = self.range_cells(table_key)
                evidence={'source':self,'grid':grid,'recipe':self.audit_evidence(table_key,grid),'expected':None}
                if is_calc_key(key):
                    word=match.group(1);name=CALC_NAMES.get(word,word.lower())
                    if name in ('prev','week','avg','peak','low'):
                        found=self._series(ref)
                        evidence['calculation']={'name':name,'count':int(match.group(2) or (30 if name in ('peak','low') else 7)),
                            'rate':found.rate,'scale':found.scale,'lang':'ko' if HANGUL.search(word) else 'en',
                            'date_parts':len(ref.date_parts),
                            'quarter':bool(ref.date_parts and _holds_quarters(ref.sheets[0],ref,ref.date_parts[1]))}
                return [evidence]
            elif is_rank_key(key):
                grid, _where, _issues = self.range_cells(key)
                match=RANK.fullmatch(key.translate(SMART_QUOTES))
                word=RANK_NAMES[match.group(1).lower()]
                expected=grid[-1][0 if word=='name' else 2 if word=='share' else 1][1]
            elif is_latest_key(key):
                grid, _where, _issues = self.range_cells(key)
                expected=grid[-1][-1][1]
            else:
                ws,(left,top,right,bottom),_where=self.locate(key)
                if (left,top)!=(right,bottom):return []
                raw,text,fmt,_issues=self.read(ws,top,left)
                grid=[[(raw,text,fmt)]];expected=text
            return [{'source':self,'grid':grid,'recipe':self.audit_evidence(key,grid),'expected':expected}]
        except (IssueError,ValueError,IndexError) as exc:
            return [{'source':self,'grid':[],'recipe':[],
                     'unavailable':'text source evidence is unavailable: '+str(exc)}]


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
        sizes = [len(text) for text in texts]  # the original lengths: matches are found in the joined original text
        replaced = {}
        colored = {}  # run -> [start, end, color] of values whose format colors them, in the run's new text
        for match in reversed(list(PLACEHOLDER.finditer(joined))):
            new = resolve(match.group(1))
            if new is None:
                continue
            start, end = match.span()
            pos = 0
            for i, size in enumerate(sizes):
                run_start, run_end = pos, pos + size
                pos = run_end
                if run_end <= start or run_start >= end:
                    continue
                text = texts[i]
                cut_from, cut_to = max(start - run_start, 0), min(end - run_start, size)  # later matches were replaced first, so this part of the run is still as it was
                middle = new if run_start <= start < run_end else ""
                texts[i] = text[:cut_from] + middle + text[cut_to:]
                shift = len(middle) - (cut_to - cut_from)
                for span in colored.get(i, []):
                    if span[0] >= cut_to:
                        span[0] += shift
                        span[1] += shift
                if middle and getattr(new, "color", None):
                    colored.setdefault(i, []).append([cut_from, cut_from + len(middle), new.color])
                if middle:
                    for lo, hi, color in getattr(new, "spans", ()):
                        if 0 <= lo < hi <= len(middle):
                            colored.setdefault(i, []).append([cut_from + lo, cut_from + hi, color])
                replaced[i] = True
        for i in replaced:
            if not texts[i]:
                paragraph.remove(runs[i])
            elif i in colored:
                _write_colored_run(runs[i], texts[i], sorted(colored[i]))
            else:
                _write_run(runs[i], texts[i])


def _write_colored_run(run, text, spans):
    """Split a run so each colored value gets its own run with that color, and the words around it keep theirs."""
    pieces, at = [], 0
    for start, end, color in spans:
        pieces += [(text[at:start], None), (text[start:end], color)]
        at = end
    pieces.append((text[at:], None))
    for piece, color in [p for p in pieces if p[0]]:
        made = copy.deepcopy(run)
        if color:
            set_run_color(made, color)
        run.addprevious(made)
        _write_run(made, piece)  # line breaks it adds go right after it, so still before the original run
    run.getparent().remove(run)


FILLS = {A + name for name in ("noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill")}


def set_run_color(run, color):
    """Give a run a solid text color, in the place the schema keeps it: after a:ln, before effects and fonts."""
    rpr = run.find(A_RPR)
    if rpr is None:
        rpr = etree.Element(A_RPR)
        run.insert(0, rpr)
    for fill in [child for child in rpr if child.tag in FILLS]:
        rpr.remove(fill)
    solid = etree.Element(A + "solidFill")
    etree.SubElement(solid, A + "srgbClr", val=color)
    line = rpr.find(A + "ln")
    if line is not None:
        line.addnext(solid)
    else:
        rpr.insert(0, solid)


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
    if getattr(text, "color", None):
        set_run_color(run, text.color)
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


def _chart_number_format(rows,column=None):
    """Number format for chart values. The XML writer in python-pptx does not escape it, so unsafe ones are dropped."""
    for line in rows[1:]:
        for raw, _text, fmt in (line[column:column+1] if column is not None else line[1:]):
            if isinstance(raw, (int, float,Fraction)) and not isinstance(raw, bool):
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
        if len(rows)<2 or len(rows[0])<3 or len(rows[0])%2!=1:
            return [Issue("chart_bubble_columns", {})]
        if any(figure_number(line[c][0]) is not None and figure_number(line[c][0])<0 for line in rows[1:] for c in range(2,len(rows[0]),2)):
            return [Issue("chart_bubble_size", {})]
    if len(rows) < 2 or len(rows[0]) < 2:
        return [Issue("chart_too_small", {})]
    issues = []
    lost = _number_in_first_row(rows)
    if lost is not None:
        issues.append(Issue("chart_header_number", {"value": lost}))

    def number(cell):
        raw, text, _fmt = cell
        value = figure_number(raw)
        try:finite=value is not None and math.isfinite(value)
        except OverflowError:finite=False
        if not finite:
            if raw is not None:
                issues.append(Issue("chart_not_number", {"value": text}))
            return None
        if isinstance(raw, str):
            converted.append(text)
        return float(value) if isinstance(value,Fraction) else value

    converted = []
    number_format = _chart_number_format(rows)
    if chart.chart_type in BUBBLE_CHARTS:
        data = BubbleChartData(number_format=number_format)
        for c in range(1,len(rows[0]),2):
            series=data.add_series(rows[0][c][1] or "",number_format=_chart_number_format(rows,c))
            points=0
            for line in rows[1:]:
                x,y,size=number(line[0]),number(line[c]),number(line[c+1])
                if x is not None and y is not None and size is not None:
                    series.add_data_point(x,y,size);points+=1
            if not points:return [Issue("chart_bubble_empty", {})]
    elif chart.chart_type in XY_CHARTS:
        data = XyChartData(number_format=number_format)
        for c in range(1, len(rows[0])):
            series = data.add_series(rows[0][c][1] or "",number_format=_chart_number_format(rows,c))
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
            data.add_series(rows[0][c][1] or "", [number(line[c]) for line in rows[1:]],number_format=_chart_number_format(rows,c))
    chart.replace_data(data)
    if converted:
        issues.append(Issue("chart_number_from_text", {"count": len(converted), "text": converted[0]}))
    return issues


# ---------------------------------------------------------------- the run

class _Run:
    def __init__(self, source, prs, insights=None):
        self.source, self.prs = source, prs
        self.guard = automatic.Session(source)
        self.insights = insights or {}
        self.requests = {}
        self.template_text = None
        self.items = {}
        self.done_parts = set()
        self.filled_paragraphs = set()  # table cells we just wrote: their data must not be scanned for placeholders
        self.text_evidence_cache = {}
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
            if text is not None:self.capture_text_input(key,item,text)
            return text
        return resolve

    def capture_text_input(self,key,item,text):
        if not hasattr(self.source,'text_evidence'):return
        if key not in self.text_evidence_cache:self.text_evidence_cache[key]=self.source.text_evidence(key)
        for evidence in self.text_evidence_cache[key]:
            self.guard.capture_text_input(evidence['source'],key,item,
                                          {k:v for k,v in evidence.items() if k!='source'},str(text))

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
            if not isinstance(text, str) or unbacked_numbers(text, (n for fact in request.facts for n in fact.numbers)):
                text = request.fallback
                issues = [Issue("insight_numbers", {})] + issues
        else:
            text = request.fallback
            issues = [Issue("insight_stale" if chosen else "insight_not_chosen", {})] + issues
        item.value = text
        self.add_issues(item, issues)
        self.capture_text_input(key,item,text)
        return text

    def fill_part_text(self, part, where):
        self.done_parts.add(part)
        resolve = self.resolve_text(where)
        for index,paragraph in enumerate(part._element.iter(A_P)):
            if paragraph not in self.filled_paragraphs:
                before=etree.tostring(paragraph)
                replace_in_paragraph(paragraph, resolve)
                if etree.tostring(paragraph)!=before:
                    self.guard.capture_paragraph(part,index,automatic.paragraph_text(paragraph))

    def fill_slide_diagrams(self, slide_part, where):
        """SmartArt keeps its words in a data part, and PowerPoint draws it from a drawn copy sized for the words
        it had: a long value would spill out of its box. So once a value goes into the data, the drawn copy is
        dropped, and PowerPoint lays the SmartArt out again from the data when it opens the report
        (tools/smartart_check.ps1). A drawn copy that cannot be matched to its data is filled the same way instead."""
        parts = _diagram_parts(slide_part)
        for data in [p for p in parts if str(p.partname).startswith("/ppt/diagrams/data")]:
            root = self.fill_diagram_text(data, where)
            if root is None:
                continue
            for ext in root.iter(f"{{{DSP}}}dataModelExt"):
                rel = slide_part.rels.get(ext.get("relId") or "")
                if rel is None or rel.is_external or rel.reltype != RT_DIAGRAM_DRAWING:
                    continue
                self.done_parts.add(rel.target_part)
                slide_part.rels.pop(ext.get("relId"))
                holder = ext.getparent()  # <a:ext uri=...>; the list it sits in goes too when it is left empty
                listing = holder.getparent()
                listing.remove(holder)
                if len(listing) == 0:
                    listing.getparent().remove(listing)
            if not isinstance(data, XmlPart):
                data._blob = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
        for drawing in [p for p in parts if str(p.partname).startswith("/ppt/diagrams/drawing") and p not in self.done_parts]:
            root = self.fill_diagram_text(drawing, where)
            if root is not None and not isinstance(drawing, XmlPart):
                drawing._blob = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)

    def fill_diagram_text(self, part, where):
        """Fill a SmartArt part's paragraphs. Returns its XML when a value went in, else None."""
        if part in self.done_parts:
            return None
        if isinstance(part, XmlPart):
            root = part._element
        else:
            raw = part.blob
            if not PLACEHOLDER.search(raw.decode("utf-8", "ignore")):
                return None
            try:
                root = etree.fromstring(raw, etree.XMLParser(resolve_entities=False, no_network=True, huge_tree=False))
            except etree.XMLSyntaxError:
                return None  # left for flag_unreachable to report
        if any(PLACEHOLDER.search(text) for text in root.xpath("//text()[not(parent::a:t)]", namespaces={"a": A[1:-1]})):
            return None  # a placeholder outside the text runs: not a place a value can go, so it is reported, not filled
        self.done_parts.add(part)
        resolve, filled = self.resolve_text(where), []

        def counted(key):
            new = resolve(key)
            if new is not None:
                filled.append(key)
            return new

        for index,paragraph in enumerate(root.iter(A_P)):
            before=etree.tostring(paragraph)
            replace_in_paragraph(paragraph, counted)
            if etree.tostring(paragraph)!=before:
                self.guard.capture_paragraph(part,index,automatic.paragraph_text(paragraph))
        return root if filled else None

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
        audited_source, audited_key = self.source, match.group(2)
        document_key=False
        if isinstance(self.source,Sources) and DOCUMENT_MARK in audited_key:
            document_key=True
            audited_source=self.source.sources[0]
        elif isinstance(self.source, Sources):
            _number, audited_source, audited_key = self.source._route(audited_key)
        try:
            recipe = self.source.document_evidence(audited_key) if document_key else audited_source.audit_evidence(audited_key, rows)
        except (IssueError, ValueError, TypeError, KeyError, IndexError, AttributeError):
            recipe = None
        self.guard.capture(shape, kind, audited_key, rows, recipe, audited_source, item,
                           mode=("bubble" if shape.chart.chart_type in BUBBLE_CHARTS else "xy" if shape.chart.chart_type in XY_CHARTS else "category") if kind == "chart" else "category",
                           date_axis=bool(shape.chart._chartSpace.xpath(".//c:dateAx")) if kind == "chart" else False)
        if skip:
            self.guard.bindings[-1].update(skip=skip, header=[cell.text for cell in shape.table.rows[0].cells])
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
        self.fill_slide_diagrams(slide.part, where)
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
        """Placeholders in parts we do not edit, such as SmartArt we could not read, are reported rather than silently left."""
        left = []
        for part in self.prs.part.package.iter_parts():
            name = str(part.partname)
            if part in self.done_parts or not name.startswith("/ppt/") or not name.endswith(".xml"):
                continue
            if isinstance(part, XmlPart) and not name.startswith(("/ppt/diagrams/", "/ppt/charts/")):
                continue
            if PLACEHOLDER.search(part.blob.decode("utf-8", "ignore")):
                left.append((part, name))
        paired = {drawing for part, name in left if name.startswith("/ppt/diagrams/data") for drawing in self._drawings_of(part)}
        for part, name in left:
            if part in paired:  # PowerPoint's drawn copy of a SmartArt already listed, not a second place to fix
                continue
            item = self.item("unsupported", name, ("part", name))
            self.add_issues(item, [Issue("unsupported_location", {"part": name, "slide": self.slide_holding(part)})])

    def _drawings_of(self, data):
        """The drawn copies of a SmartArt data part: what its dataModelExt names on the slide that holds it, or else the drawing numbered as it is."""
        found = []
        ids = re.findall(rb'dataModelExt[^>]*\brelId="([^"]+)"', data.blob)
        for slide in self.prs.slides:
            if any(not rel.is_external and rel.target_part is data for rel in slide.part.rels.values()):
                found += [slide.part.rels[i.decode()].target_part for i in ids if i.decode() in slide.part.rels]
        if not found:
            number = re.search(r"(\d+)\.xml$", str(data.partname))
            found = [part for part in self.prs.part.package.iter_parts() if number and str(part.partname) == f"/ppt/diagrams/drawing{number.group(1)}.xml"]
        return found

    def slide_holding(self, part):
        """The number of the slide that uses a part, so a person can find it."""
        for number, slide in enumerate(self.prs.slides, start=1):
            if any(not rel.is_external and rel.target_part is part for rel in slide.part.rels.values()):
                return number
        return "?"


def _diagram_parts(slide_part):
    """A slide's SmartArt data parts and their drawn copies. The drawing may hang from the slide or from the data part."""
    found, todo = [], [slide_part]
    while todo:
        part = todo.pop()
        for rel in part.rels.values():
            if rel.is_external:
                continue
            target = rel.target_part
            name = str(target.partname)
            if name.startswith(("/ppt/diagrams/data", "/ppt/diagrams/drawing")) and target not in found:
                found.append(target)
                todo.append(target)
    return found


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
        self.document_fingerprints = {path:automatic.digest(path) for path in paths if path not in workbooks}
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
                text = (f"[{document.name}] {period}에 적힌 일: {found.event.text} 같은 항목 {found.item.name}의 전 {found.window}개 기간 평균 {documents.mean_text(found.before)}, 후 {found.window}개 기간 평균 {documents.mean_text(found.after)}, {documents.change_text(found)}." if lang == "ko"
                        else f"[{document.name}] Entered for {period}: {found.event.text} {found.item.name} averaged {documents.mean_text(found.before)} over the {found.window} periods before and {documents.mean_text(found.after)} over the {found.window} after, {documents.change_text(found)}.")
                pack.append(facts.Fact(f"doc{number}_link{place}", text, tuple(facts.NUMBER.findall(f"{documents.mean_text(found.before)} {documents.mean_text(found.after)} {documents.change_text(found)}"))))
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

    def text_evidence(self,key):
        if DOCUMENT_MARK in key:
            return [{'source':row['source'],'grid':row['grid'],'recipe':row['recipe'],'expected':None}
                    for row in self.document_evidence(key,all_links=True)['rows']]
        if ALL_FILES in key:
            _word,_mark,references=key.partition(ALL_FILES)
            result=[]
            for reference in references.split('|'):
                _number,source,plain=self._route(reference)
                result.extend(source.text_evidence('reading:'+plain))
            return result
        _number,source,plain=self._route(key)
        return source.text_evidence(plain)

    def document_evidence(self,key,all_links=False):
        head,_mark,tail=key.partition(DOCUMENT_MARK)
        number,_colon,count=tail.partition(':')
        number=number.partition('#')[0]
        document,links=self.links()[int(number or 0)]
        selected=links if all_links else documents.highlights(links,int(count or DOCUMENT_LINKS))
        rows=[]
        for found in selected:
            source=self.sources[found.item.file]
            if not found.item.reference:raise ValueError('document figure has no explicit source reference')
            if len(found.item.points)>100000:raise ValueError('document source evidence exceeds budget')
            table_key='latest'+str(len(found.item.points))+':'+found.item.reference
            grid,_where,_issues=source.range_cells(table_key)
            rows.append({'source':source,'grid':grid,'recipe':source.audit_evidence(table_key,grid),
                         'event':found.event.date,'name':found.item.name,'text':found.event.text,
                         'window':found.window,'document':document.path})
        return {'kind':'document','lang':'ko' if head.startswith(('ko','읽는법','해설','통찰')) else 'en','rows':rows}

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
    return re.sub(r"^\s*(?:\d{4}년\s*(?:\d+월\s*(?:\d+일\s*(?:\(.\)\s*)?)?|\d분기\s*)?|\d+월\s*\d+일\s*(?:\(.\)\s*)?|\d{4}-\d{2}(?:-\d{2})?\s*|Q\d \d{4}\s*|Q[1-4]\s*\d{2,4}/\d{2}\s*|Q[1-4]\s*FY\s*'?\d{2,4}\s*|FY\s*'?\d{2,4}\s*Q[1-4]\s*|\d{4}\s+|[A-Z][a-z]{2} \d{4}\s*|[A-Z][a-z]{2} \d+ (?:\(\w+\) )?)", "", fact.text.partition(":")[0]).strip()


def _lead_change(request, lang):
    """(exact percent change, period, basis), from typed results rather than rendered prose.

    A points change or a legacy prose-only fact has no relative-change payload and cannot
    stand beside relative percentages. A year comparison stays preferred over the previous period.
    """
    found = {fact.id: fact for fact in request.facts}
    fact, basis = (found["year"], "year") if "year" in found else (found.get("prev"), "prev")
    change = fact.relative_change if fact is not None else None
    percent_change = change * 100 if isinstance(change, Fraction) else None
    period = next((fact.period for fact in request.facts if fact.period), "")  # as the figure's own sentence says it: yearly data has no month
    return percent_change, _period_in(period, lang), basis


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
            if fmt in ("General", "0") and isinstance(raw, (int, float,Fraction)) and not isinstance(raw, bool) and abs(raw) >= 1000:
                text = re.sub(r"^(-?)(\d+)", lambda m: f"{m.group(1)}{int(m.group(2)):,}", text)
            cells.append((raw, text, fmt))
        done.append(cells)
    return done


def run(template, excel, output=None, insights=None, source=None, thousands=False, *, collect_vba=False):
    """Fill the template. Without output this is the preview: same work, nothing written.

    insights holds sentences a person chose earlier. No model is called here, so preview and report always agree.
    source is a workbook already read from excel, so a large one need not be read twice in a row.
    thousands writes plain large numbers in tables with commas, for a report SlideFill laid out itself.
    collect_vba prepares optional native comparison inputs from verified numeric report cells.
    """
    if source is None:
        source = Sources(excel) if isinstance(excel, (list, tuple)) else ExcelSource(excel)
    source.thousands = thousands
    prs = open_template(template)
    work = _Run(source, prs, insights)
    work.walk()
    for binding in work.guard.bindings:
        if binding.get("skip"):
            def header_value(match):
                item = work.items.get(("text", match.group(1).strip().casefold()))
                return item.value if item is not None and item.status != "error" else match.group(0)
            binding["header"] = [PLACEHOLDER.sub(header_value, text) for text in binding["header"]]
    report = Report(items=list(work.items.values()), requests=work.requests, source_snapshot=source)
    blob = automatic.serialize(prs)
    if work.guard.bindings or work.guard.paragraphs or work.guard.text_inputs:
        report.verification = work.guard.verify(blob)
        for item, detail in report.verification["failures"]:
            target = item or report.items[0]
            work.add_issues(target, [Issue("verification_failed", {"detail": detail})])
        if report.incomplete and report.items:
            work.add_issues(report.items[0],[Issue('verification_incomplete',{'detail':'; '.join(report.verification['reasons'][:3])})])
    if collect_vba and not (report.blocked or report.incomplete):
        report.vba_cases = native_vba_cases.collect(work.guard.bindings, work.guard.books)
    if output:
        if report.blocked:
            raise SlideFillError("verification_blocked", detail="; ".join(reason for _item, reason in report.verification["failures"][:3]))
        if report.incomplete:
            raise SlideFillError("verification_incomplete",detail="; ".join(report.verification['reasons'][:3]))
        try: report.output = work.guard.publish(blob, output)
        except automatic.Incomplete as exc: raise SlideFillError("verification_incomplete",detail=str(exc))
        except automatic.Failure as exc: raise SlideFillError("verification_blocked", detail=str(exc))
        except OSError as exc: raise SlideFillError("cannot_save", path=output, detail=str(exc))
    return report


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
