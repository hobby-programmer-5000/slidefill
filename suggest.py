"""Reads how a sheet is laid out and offers references that work on it, each with its real result. No GUI, no network, no model.

A person should not have to learn A10:A475 or 합계(D, C=도시가스). Every reference offered here has been run through the same
code as the report, so what is listed is what will be filled, and nothing that fails is offered as if it worked.
"""

import datetime as dt
import fnmatch
import os
import re
from collections import Counter
from dataclasses import dataclass, field, replace
from itertools import islice

from openpyxl.utils.cell import get_column_letter
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.enum.shapes import MSO_CONNECTOR
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.util import Inches, Pt

import core
import documents
import facts

SAMPLE_ROWS = 3000  # enough to tell what a column holds; a sheet of a million rows need not be read twice for this
SAMPLE_COLUMNS = 60
HEAD_ROWS = 25  # periods laid out across a row are near the top
MAX_GROUP_VALUES = 200
MAX_NAME_BASES = 6  # a sheet of many text columns need not be tried in full
HEADING_LEVELS = 4  # rows of headings looked for under periods laid across
NESTED = object()  # periods laid across with more than one level of headings under them
MIN_PERIODS = 3  # two dates are more likely two headings, 2025년 and 2024년 side by side, than a series
MANY_ROWS = 0.7  # fewer distinct dates than this share of dated rows: several rows per date, not a stray repeat
ID_NAME = re.compile(r"코드|번호|순번|연번|일련|시간|요일|좌표|(?:^|[\s_(])(?:위도|경도)(?:$|[\s_)(])|\bno\b(?!\.?\s*of\b)|\b(?:time|hour|weekday)\b"  # 위도, not 위도항, a harbor; No, not No of Units
                     r"|바코드|사번|학번|[a-z]number\b"  # EmployeeNumber
                     r"|(?<![a-z0-9])(?:id|code|abn|acn|ein|postcode|pcode|zip|zipcode|latitude|longitude|isbn|issn|sku|upc|ean|doi|barcode|num|nbr|respondent)(?![a-z0-9])"  # CUSTOMER_ID, where \b would not see the underscore
                     r"|(?<![a-z0-9])(?!total|average|mean|median|net|gross|estimated)[a-z]+[ _]number(?![a-z0-9])(?!\s+of)", re.I)  # FILE NUMBER, not Total number or Number of staff
# Places in an order, and dates stored as plain numbers. They can be items to compare by, but adding them up means nothing, so they are kept apart from ID_NAME. 차수별 건수 counts something.
NOT_A_FIGURE_NAME = re.compile(r"(?:차수|회차|순위|등수|등급|일자|날짜|일시|생년월일|기준일|시작일|종료일|개시일|지정일|발행년|발행일|(?<!\d)연도|(?<!\d)년도)(?!별)(?:\s|$|\()"  # 2023년도 수량 is a figure of that year
                               r"|(?<![a-z0-9])(?:rank|ranking|grade|sequence|date|dob|year)(?![a-z0-9])(?!\s+\w)", re.I)
TOTAL_NAME = re.compile(r"총계|합계|총공급|종합|전체|\b(?:total|overall)\b", re.I)
ENDS_AS_TOTAL = re.compile(r"(?:총계|합계|소계|전체|\b(?:total|overall))\s*$", re.I)  # 전체 소계 is a total; 전체 임원 is one part under the heading 전체
AMOUNT_NAME = re.compile(r"금액|매출|수입|비용|수량|판매|사용량|공급량|인원|건수|면적|중량|\b(?:amount|sales|revenue|cost|income|quantity|qty|area|weight|number of)\b", re.I)
LANGUAGE_SAMPLE = 2000
# What find_doubts can say. Each has words as doubt_<code> in i18n.
DOUBTS = ("example_sheet", "count_only", "sum_of_levels", "first_table", "heading_below", "dates_sparse", "maybe_percent", "several_tables", "total_differs")
EXAMPLE_SHEET = re.compile(r"example|sample|template|예시|예제|양식|서식|샘플", re.I)
LEVEL_NAME = re.compile(r"평균|최대|최소|농도|온도|습도|밀도|발열량|수질|\b(?:average|mean|temperature|density|concentration)\b", re.I)
LEVEL_MIN_ROWS = 10
LEVEL_SPREAD = 0.05
HEADING_MIN_ROWS = 5
HEADING_MAX_ABOVE = 3
DATES_MISSING_SHARE = 0.2
PER_CENT_NOTE = re.compile(r"per\s?cent|percent|percentage|%|백분율|단위\s*[:：]?\s*%", re.I)
TABLE_LINE = re.compile(r"(?:table|표)\s*\d+\b", re.I)
STACKED_SHARE = 0.5
SPLITS_SHARE = 0.2  # small airports have only domestic flights, and 국내선 and 국제선 still split the rest
RUNNING_SHARE = 0.9
CODE_MIN_ROWS = 10 # fewer rows than this could pair up by chance
CODE_MIN_VALUES = 3
CODE_VALUES_BESIDE_DATES = 8  # beside dates, three regions that happen to stay flat are still a series; many that all do are not
CODE_REPEATS = 1.5  # rows per distinct number: without repeats, any list of names pairs one to one with its figures
CODE_PURITY = 0.95  # a spelling that differs here and there should not hide a code
LONG_ID_DIGITS = 9
LONG_ID_ROWS = 20
CONTENTS_TITLE = re.compile(r"목차|색인|\btable of contents\b|\b(?:table|tables)\s+index\b|\bindex of tables\b", re.I)  # not 소비자물가지수 or Consumer price index, which are figures
TITLE_ROWS = 3
TITLE_COLUMNS = 5
ALL_ROWS_WORD = re.compile(r"all\b.*|total\b.*|.*\b(?:sub-?)?total|(?:\w+[:.]\s*|\d+\s*)?(?:합계|총계|소계|전체|계)", re.I)  # 5합계 and 00:소계 carry a code in front; 가계 and 회계 only end the same way
WHOLE_TOLERANCE = 0.02  # rounded to the nearest 5 as some publishers do, 960 stands for parts that add to 975
WHOLE_SHARE = 0.9
WHOLE_MIN_GROUPS = 3  # dates or names: with fewer, a match could be chance
WHOLE_MIN_VALUES = 3  # the whole and at least two parts: 남 and 여 of much the same size are not a whole and a part
EVERY_ROW_MIN = 20  # so many rows each with a figure are items of one kind; a few are more likely different measures
MEAN_NAME = re.compile(r"평균|가격|단가|율|률|점수|\b(?:price|rate|ratio|average|mean|score)\b", re.I)
# 분기별 평균 수출금액: each row's own average over time. Across rows these add up to the category's average, and a mean of them would hide how many rows it has.
AVERAGE_OVER_TIME = re.compile(r"(?:분기|월|연|일|주|quarterly|monthly|annual|daily|weekly)\s*(?:별\s*)?(?:평균|average|mean)\s*\S*?(?:금액|수량|수입|판매|매출|액|량|건수|인원|amount|sales|revenue|exports?|imports?|quantity)", re.I)
MEASURE_COLUMN = re.compile(r"measures?\b|지표|측정항목|indicator|metric", re.I)  # 건수, 동수 and 연면적 in one column of each date
KIND_COLUMN = re.compile(r"항목|종류|구분|측정|분류|parameter|measures?\b|metric|indicator|\b(?:type|kind)\b", re.I)  # what the rows are measurements of, not who or where
PERIOD_IN_TITLE = re.compile(r"(?:19|20)\d\d\s*년(?:\s*\d{1,2}\s*월)?(?:\s*\d\s*분기)?|\b(?:jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\.?\s+(?:19|20)\d\d\b", re.I)
PERIOD_TITLE_ROWS = 6
PERIOD_TITLE_COLUMNS = 5
PERIOD_TITLE_LENGTH = 60
NEVER_ADDED_NAME = re.compile(r"나이|연령(?!별)|근속|\b(?:age|tenure)\b", re.I)  # whatever else the sheet holds, adding people's ages says nothing
UNLIKE_MIN_LARGEST = 100000
UNLIKE_SHARE = 0.1  # so many values below one, beside values in the hundreds of thousands, are not the same kind of thing
NOT_DATA_SHEET = re.compile(r"표지|목차|설명|안내|주석|index|contents|cover|notes?\b|copyright", re.I)
YEAR_IN_NAME = re.compile(r"(?:19|20)\d\d")
PLAIN_SHEET_NAME = re.compile(r"(?:data|sheet\s*\d*|시트\s*\d*|table\s*\d*|\d+)", re.I)  # says nothing, so it is not worth a heading
UNIT_NOTE = re.compile(r"단위\s*[:：]|[(（\[]\s*단위|\bunits?\s*[:：]|[(（\[]\s*units?\b", re.I)  # not 유통단계별단위명, which is only a column's name
UNIT_HEADING = re.compile(r"단위|units?", re.I)
UNIT_VALUE_LENGTH = 10
UNIT_ROWS = 6
UNIT_COLUMNS = 12
UNIT_LENGTH = 40
MAX_FIGURES = 4  # beyond this they are more likely thirty stations than the two sides of one total, and a slide cannot hold them
READ_ALL_FIGURES = 2  # so many figures can each have their reading on one slide
MAX_SECTIONS = 3
SMALL_SHEET = 5000  # a further sheet is only worth waiting for if it reads quickly
CHART_RANGE = 10  # figures further apart than this do not share a chart's axis
LINE_HEIGHT = 1.25
LEFT, FULL_WIDTH = 0.5, 12.3  # inches
TEXT_WIDTH = 6.1
RIGHT, RIGHT_WIDTH = 6.9, 5.9
CONTENT_TOP, SLIDE_BOTTOM = 1.45, 7.1
SENTENCE_GAP = 0.2
TABLE_ROW = 0.3
TABLE_POINTS = 13
SLIDE_CHARACTERS = 1100  # characters × points that fit across the full width, for mixed Korean and digits
AXIS_FORMAT = "[>=1000]#,##0;General"  # 7,000,000 can be read, and a small figure such as 0.2 keeps its decimal
QUIET_CODES = ("reading", "change", "by_item", "insight", "doc_reading", "doc_quote", "doubt_note")  # what explains the main sentence, set smaller beneath it
RULE_AT = 1.28
WIDE_ROW = 0.42
WIDE_POINTS = 11
WIDE_COLUMNS = (1.2, 1.35, 6.05, 1.25, 1.25, 0.9)  # inches: when, figure, what the document says, before, after, change
MAX_DOCUMENT_ITEMS = 200
LABEL_HEIGHT = 0.27
MAX_TABLE_FIGURES = 6  # columns a table on half a slide can hold, and the most a person can tick on one sheet
MUTED = RGBColor(0x5F, 0x63, 0x68)
GRID = RGBColor(0xD9, 0xD9, 0xD9)
TITLE_COLOR = RGBColor(0x1F, 0x38, 0x64)
BODY_COLOR = RGBColor(0x40, 0x40, 0x40)
WORDS = {
    "latest": ("최신", "latest"), "note": ("해설", "note"), "reading": ("읽는법", "reading"), "prev": ("직전대비", "prev"), "week": ("전주대비", "week"),
    "avg": ("평균대비", "avg"), "peak": ("최고일", "peak"), "insight": ("통찰", "insight"), "top": ("순위", "top"), "topvalue": ("순위값", "topvalue"),
    "share": ("구성비", "share"), "table": ("표", "table"), "chart": ("차트", "chart"), "sum": ("합계", "sum"), "count": ("건수", "count"),
    "mean": ("평균", "mean"), "rate": ("비율", "rate"),
}


def word(name, lang):
    return WORDS[name][0 if lang == "ko" else 1]


@dataclass
class Value:
    """Something that can be read as the figure."""
    spec: str  # C, or 3:122 for every row of a table across
    name: str
    every_row: bool = False  # the sum of every line, as when each row is one harbor
    counted: bool = False  # nothing to add up, so rows are counted
    reaches_newest: bool = True  # False when the sheet's newest period has no figure for it: a programme that ended years ago
    size: float = 0  # how large its figures are together, for telling a grand total from a part of it


@dataclass
class Group:
    """A column of a few repeating names, which can be a condition or the thing compared."""
    spec: str
    name: str
    values: list = field(default_factory=list)


@dataclass
class Base:
    """What the figures hang on: dates, or the names of items."""
    kind: str  # date or name
    spec: str  # A, A+B, 5, A10:A475, or A:C for names
    name: str
    values: list = field(default_factory=list)
    groups: list = field(default_factory=list)
    many_rows: bool = False  # several rows per date or per name, so they must be added up, averaged or counted
    first: str = ""
    last: str = ""
    count: int = 0
    notes: list = field(default_factory=list)  # (message code, params)
    newest_word: str = ""  # how to ask for the newest period in a condition, such as 2025-12
    date_spec: str = ""  # the date line, for that condition


@dataclass
class Shape:
    sheet: str
    bases: list = field(default_factory=list)
    glob: str = ""  # several sheets that continue one another


@dataclass
class Suggestion:
    code: str  # what it is, an i18n key after suggest_
    key: str
    place: str  # text, or alt for the alt text of a table or chart
    result: str = ""
    status: str = "ok"
    issues: list = field(default_factory=list)
    columns: int = 0
    rows: int = 0  # of data, for a table or chart
    figure: str = ""  # which figure it is about, when a report holds several
    label: str = ""  # a small line above it on the slide, such as the file it came from

    @property
    def typed(self):
        return self.result if self.place == "plain" else "{{" + self.key + "}}"  # plain: words for the slide, with nothing to look up


@dataclass
class Section:
    """One sheet's part of a report."""
    heading: str
    sentences: list
    table: Suggestion = None
    chart: Suggestion = None
    figures: list = field(default_factory=list)
    title: str = ""  # its own title, when the report's one title does not say what this slide is
    sheet: str = ""
    file: int = 0
    cross: str = ""  # the key of a sentence about several files together, once someone chooses one
    wide: bool = False  # a table across the whole slide, under the sentences
    period: str = ""  # the time its figures are for, when a sheet says so only in its title

    @property
    def items(self):
        return self.sentences + [item for item in (self.table, self.chart) if item]


@dataclass
class Plan:
    """A report made from the workbook alone."""
    sections: list
    what: str
    figures_left: int = 0
    sheets_left: int = 0
    files_left: list = field(default_factory=list)  # files nothing could be read from
    doubts: list = field(default_factory=list)  # (code, params): why a person should look before this becomes a report

    @property
    def items(self):
        return [item for section in self.sections for item in section.items]


def data_sheets(source):
    """The sheets worth offering, best first: the newest by the year in its name, or the one the workbook opens on."""
    found = []
    for ws in source.wb.worksheets:
        if ws.max_row < 2 or NOT_DATA_SHEET.search(ws.title) or _titled_as_contents(ws):
            continue
        cells = [ws.cell(r, c).value for r in range(1, min(ws.max_row, 60) + 1) for c in range(1, min(ws.max_column, SAMPLE_COLUMNS) + 1)]
        if sum(core.as_number(v) is not None for v in cells) >= 3:
            found.append(ws)
    active = source.wb.active
    years = {ws.title: max((int(y) for y in YEAR_IN_NAME.findall(ws.title)), default=0) for ws in found}
    return sorted(found, key=lambda ws: (-years[ws.title], ws is not active, source.wb.worksheets.index(ws)))


def _titled_as_contents(ws):
    """A sheet named Companies whose title says it only lists where things are: its numbers are places in other tables."""
    for row in range(1, min(ws.max_row, TITLE_ROWS) + 1):
        for column in range(1, min(ws.max_column, TITLE_COLUMNS) + 1):
            raw = ws.cell(row, column).value
            if isinstance(raw, str) and CONTENTS_TITLE.search(raw):
                return True
    return False


def _one_to_one(pairs):
    """Each number goes with one name and each name with one number, over many repeated rows."""
    if len(pairs) < CODE_MIN_ROWS:
        return False
    by_number, by_name = {}, {}
    for number, name in pairs:
        by_number.setdefault(number, Counter())[name] += 1
        by_name.setdefault(name, Counter())[number] += 1
    if len(by_number) < CODE_MIN_VALUES or len(pairs) < len(by_number) * CODE_REPEATS:
        return False
    return all(sum(counts.most_common(1)[0][1] for counts in side.values()) >= len(pairs) * CODE_PURITY for side in (by_number, by_name))


def _code_columns(ws):
    """Columns of numbers that only stand for a name beside them, with how many different numbers each holds: an 8-digit species code with its species, or a region's area repeated on each of its rows. Adding those up, or calling them a figure, means nothing."""
    limit = min(ws.max_row, SAMPLE_ROWS)
    numbers, names = {}, {}
    for column in range(1, min(ws.max_column, SAMPLE_COLUMNS) + 1):
        cells = [ws.cell(row, column).value for row in range(1, limit + 1)]
        read = {row: core.as_number(raw) for row, raw in enumerate(cells) if raw not in (None, "")}
        counted = [number for number in read.values() if number is not None]
        if len(counted) >= max(len(read) * 0.5, 2):
            numbers[column] = {row: number for row, number in read.items() if number is not None}
        elif read:
            names[column] = {row: str(cells[row]).strip() for row, number in read.items() if number is None}
    return {column: len(set(found.values())) for column, found in numbers.items() if any(_one_to_one([(found[row], words[row]) for row in found.keys() & words.keys()]) for words in names.values())}


def _long_identifiers(numbers):
    """Whole numbers all different and all of nine or more digits, like business numbers. A real figure across this many named things is not so even, whatever its column is called."""
    found = [number for number in numbers if number is not None]
    if len(found) < LONG_ID_ROWS or len(set(found)) < len(found) or any(number != int(number) or number < 0 for number in found):
        return False
    return len({len(str(int(number))) for number in found}) == 1 and found[0] >= 10 ** (LONG_ID_DIGITS - 1)


def _names_a_figure(name):
    """Not a code or a place in an order by its name, and a name at all: a heading of - says nothing about what the numbers are."""
    return not ID_NAME.search(name) and not NOT_A_FIGURE_NAME.search(name) and re.search(r"\w", name) is not None


def _unlike_things(numbers):
    """An axis of 6,378,137 m, a flattening of 0.0034 and a scale of 0.9996 in one column are constants, not one figure across items."""
    found = [abs(number) for number in numbers if number]
    return bool(found) and max(found) >= UNLIKE_MIN_LARGEST and sum(number < 1 for number in found) >= len(found) * UNLIKE_SHARE


def _profile(ws, line, by_row, limit):
    """What a column, or a row, holds: counts over the first limit cells."""
    counts = Counter()
    distinct = Counter()
    for place in range(1, limit + 1):
        cell = ws.cell(line, place) if by_row else ws.cell(place, line)
        raw = cell.value
        if raw in (None, ""):
            continue
        counts["filled"] += 1
        if core.as_date(raw) is not None or (isinstance(raw, str) and core.YEAR_SPAN.fullmatch(raw.strip())):
            counts["date"] += 1
        elif core.as_year(raw) is not None:
            counts["year"] += 1
        elif core.as_month(raw) is not None:
            counts["month"] += 1
        elif core.as_number(raw) is not None:
            counts["number"] += 1
        else:
            counts["text"] += 1
            if len(distinct) <= MAX_GROUP_VALUES + 1:
                distinct[str(raw).strip()] += 1
        part = core.as_part_of_date(raw)
        if part is not None and 1 <= part <= 12:
            counts["monthish"] += 1
    return counts, distinct


def _name(ref_line, by_row):
    return str(ref_line) if by_row else get_column_letter(ref_line)


def _date_candidates(ws):
    """Date specs worth trying, likeliest first, as (spec, by_row)."""
    limit = min(ws.max_row, SAMPLE_ROWS)
    columns = min(ws.max_column, SAMPLE_COLUMNS)
    found = []
    profiles = {c: _profile(ws, c, False, limit)[0] for c in range(1, columns + 1)}
    for c, p in profiles.items():
        filled = p["filled"]
        if not filled:
            continue
        if p["year"] >= 3 and p["year"] >= filled * 0.8:
            for other in range(c + 1, min(c + 3, columns) + 1):
                q = profiles[other]
                if q["filled"] and q["monthish"] >= q["filled"] * 0.8 and q["number"] + q["month"] + q["text"] + q["year"] >= 2:
                    found.append((p["year"] * 2, f"{get_column_letter(c)}+{get_column_letter(other)}", False))
                    break
        size = p["date"] + p["year"] + p["month"]
        if size >= 3 and size >= filled * 0.4:
            found.append((size, get_column_letter(c), False))
    for r in range(1, min(ws.max_row, HEAD_ROWS) + 1):
        p = _profile(ws, r, True, columns)[0]
        size = p["date"] + p["year"] + p["month"]
        if size >= 3 and size >= p["filled"] * 0.4:
            found.append((size, str(r), True))
    found.sort(key=lambda item: -item[0])
    return [(spec, by_row) for _size, spec, by_row in found[:6]]


def _blocks(positions):
    """Runs of positions with no more than a one-line gap between them."""
    runs = [[positions[0]]]
    for position in positions[1:]:
        if position - runs[-1][-1] > 2:
            runs.append([])
        runs[-1].append(position)
    return runs


def _last_form(ws, ref, positions):
    """(first, last) of the final run of dates written one way, when an earlier run is written another way: years and then quarters read as one series would mislead."""
    forms = []
    for position in positions:
        raw = ws.cell(*ref.at(position, ref.date)).value
        form = re.sub(r"[A-Za-z가-힣]+", "a", re.sub(r"\d+", "#", raw.strip())) if isinstance(raw, str) else type(raw).__name__
        if forms and forms[-1][0] == form:
            forms[-1][1].append(position)
        else:
            forms.append((form, [position]))
    if not 2 <= len(forms) <= 3 or len(forms[-1][1]) < 3 or not all(len(run) >= 3 for _form, run in forms):
        return None
    return forms[-1][1][0], forms[-1][1][-1]


def _bounded(spec, by_row, start, end):
    letter = spec.split("+")[0]
    if by_row:
        return f"{get_column_letter(start)}{letter}:{get_column_letter(end)}{letter}"
    return f"{letter}{start}:{letter}{end}"


def _read_date_base(source, ws, spec, by_row, codes):
    """A Base for a date line, with stacked tables and stray repeats dealt with, or None when the line does not really hold dates."""
    try:
        ref = source._parse_ref(f"{ws.title}!{spec}", spec)
    except core.IssueError:
        return None
    dated = source._dated_rows(ws, ref)
    if len({date for date, _position in dated}) < MIN_PERIODS:
        return None
    dates = [date for date, _position in dated]
    positions = [position for _date, position in dated]
    notes, used = [], spec
    span = None
    base_name = source._date_header(ws, ref, positions[0])
    repeated = [date for date, times in Counter(dates).items() if times > 1]
    if repeated:
        runs = _blocks(positions)
        by_position = dict((position, date) for date, position in dated)
        unique_runs = len(runs) >= 2 and all(len({by_position[p] for p in run}) == len(run) for run in runs)
        if unique_runs and len(runs[0]) >= 3:  # two tables one above the other, such as an index and its revisions
            span = (runs[0][0], runs[0][-1])
            notes.append(("note_first_table", {"count": len(runs)}))
        elif len(set(dates)) >= len(dates) * MANY_ROWS:  # a stray repeat, perhaps a typing slip: read only what comes after it
            second = max(p for p in positions if dates.count(by_position[p]) > 1)
            after = [p for p in positions if p > second]
            if len(after) >= 3:
                span = (after[0], after[-1])
                notes.append(("note_repeat", {"date": max(repeated).date().isoformat(), "count": len(after)}))
    if "+" not in spec and not span:
        last_form = _last_form(ws, ref, positions)
        if last_form:
            span = last_form
            notes.append(("note_last_form", {}))
    if span:
        used = _bounded(spec, by_row, *span)
        ref = source._parse_ref(f"{ws.title}!{used}", used)
        dated = source._dated_rows(ws, ref)
        dates = [date for date, _position in dated]
    positions = [position for _date, position in dated]
    many = len(set(dates)) < len(dates) * MANY_ROWS
    lines = set(ref.date_parts or (ref.date,))
    base = Base("date", used, base_name, first=min(dates).date().isoformat(), last=max(dates).date().isoformat(), count=len(set(dates)), notes=notes)
    other_count = ws.max_row if by_row else min(ws.max_column, SAMPLE_COLUMNS)
    newest_position = max(dated)[1]
    stepping = []  # 1, 2, 3 is most likely a running number, but it may be all the sheet has
    sample = positions[:SAMPLE_ROWS] if not many else positions[:SAMPLE_ROWS]
    # Rows under a date with no date of their own belong to it, as with merged cells.
    under = [p + 1 for p in sample if p + 1 not in set(positions) and ws.cell(*ref.at(p + 1, ref.date)).value in (None, "")]
    for line in range(ref.date + 1 if by_row else 1, other_count + 1):
        if line in lines:
            continue
        cells = [ws.cell(*ref.at(p, line)).value for p in sample + under[:SAMPLE_ROWS]]
        filled = [v for v in cells if v not in (None, "")]
        if not filled:
            if by_row and base.values:
                break  # the table under the dates has ended
            continue
        numbers = [core.as_number(v) for v in filled]
        count = sum(n is not None for n in numbers)
        name = source._header(ws, ref, positions[0], line)
        years_only = all(core.as_year(v) is not None for v in filled)  # 통계년도 beside the date: another way of writing it, not a figure
        dates_too = sum(core.as_date(v) is not None for v in filled) >= len(filled) * 0.8  # 20080625 is a day, and adding days means nothing
        code = not by_row and (codes.get(line, 0) >= CODE_VALUES_BESIDE_DATES or (many and _long_identifiers(numbers)))  # a book's number under each year; a series of large values has one per date
        if count >= max(len(filled) * 0.5, 2) and _names_a_figure(name) and not years_only and not dates_too and not code:
            reaches = many or core.as_number(ws.cell(*ref.at(newest_position, line)).value) is not None
            (stepping if _steps_by_one(numbers) else base.values).append(Value(_name(line, by_row), name, reaches_newest=reaches, size=sum(abs(number) for number in numbers if number)))
        elif code:
            continue
        elif not by_row:
            words = Counter(str(v).strip() for v in filled if core.as_number(v) is None)
            if 2 <= len(words) <= MAX_GROUP_VALUES and sum(words.values()) > len(words):
                base.groups.append(Group(_name(line, by_row), name, list(words)))
    groups_under = bool(under) and len(under) >= len(sample) * 0.5
    base.many_rows = many or groups_under
    if by_row and base.many_rows:
        levels = _levels_under(ws, ref, positions)
        if len(levels) > 1:
            return NESTED
        base.groups = levels
    if by_row and len(base.values) > 3:
        lines_of = [int(v.spec) for v in base.values]
        total = next((v for v in base.values if TOTAL_NAME.search(v.name)), None)
        # Each row is one item, so their sum is what a report wants first, unless they are rates or indexes that cannot be added.
        if not total and not facts.is_rate(ws.title) and len(lines_of) >= EVERY_ROW_MIN and lines_of == list(range(min(lines_of), max(lines_of) + 1)):
            base.values.insert(0, Value(f"{min(lines_of)}:{max(lines_of)}", "", True))
    if not many and not base.values:
        base.values = stepping
    if many and not base.values:  # a log of events, with nothing to add: what it can report is how many per date
        base.values.append(Value(get_column_letter(ref.date), "", counted=True))
    newest = max(dated)
    raw = source._date_cell(ref, 0, newest[1])[0]
    base.newest_word = raw.strip() if isinstance(raw, str) else newest[0].strftime("%Y-%m") if isinstance(raw, dt.datetime) else str(raw)
    base.date_spec = spec.split("+")[0] if "+" not in spec else ""
    return base


def _levels_under(ws, ref, positions):
    """The rows of headings between periods laid across and the figures below, as when each year spans 사고건수, 사망자수 and more. Those cannot be added together."""
    step = positions[-1] - positions[-2] if len(positions) >= 2 else 1
    levels = []
    for row in range(ref.date + 1, min(ref.date + HEADING_LEVELS, ws.max_row) + 1):
        cells = [core._merged_value(ws, row, column) for column in range(positions[0], positions[-1] + step)]
        if any(core.as_number(cell) is not None for cell in cells):
            break
        words = list(dict.fromkeys(str(cell).strip() for cell in cells if cell not in (None, "")))
        if len(words) >= 2:
            levels.append(Group(str(row), " / ".join(words[:2]), words))
    return levels


def _steps_by_one(numbers):
    """A running number, not a figure, even one that starts again for each kind of thing listed."""
    found = [n for n in numbers if n is not None]
    return len(found) >= 5 and sum(later - earlier == 1 for earlier, later in zip(found, found[1:])) >= (len(found) - 1) * RUNNING_SHARE


def _read_name_bases(source, ws, codes):
    """Bases for a table with no dates: which columns name the items, and which hold figures."""
    limit = min(ws.max_row, SAMPLE_ROWS)
    columns = min(ws.max_column, SAMPLE_COLUMNS)
    profiles = {c: _profile(ws, c, False, limit) for c in range(1, columns + 1)}
    numeric = [c for c, (p, _d) in profiles.items() if p["number"] + p["year"] >= max(p["filled"] * 0.5, 2) and p["number"] >= 2]
    words = [c for c, (p, d) in profiles.items() if p["text"] + p["month"] >= 2 and c not in numeric and any(core.as_date(ws.cell(r, c).value) is None for r in range(1, 3))]
    if not numeric or not words:
        return []
    data_rows = [r for r in range(1, limit + 1) if core.as_number(ws.cell(r, numeric[0]).value) is not None]
    bases = []
    for c in words[:MAX_NAME_BASES]:
        p, d = profiles[c]
        if c > numeric[-1] and len(words) > 1 and not d:
            continue
        # Words only from rows with figures, so a heading is not offered as a value to choose.
        d = Counter(str(core._merged_value(ws, r, c)).strip() for r in data_rows if core._merged_value(ws, r, c) not in (None, "") and core.as_number(core._merged_value(ws, r, c)) is None)
        ref = source._parse_ref(f"{ws.title}!{get_column_letter(c)}", "x")
        first_filled = next((r for r in range(1, limit + 1) if core.as_number(ws.cell(r, numeric[0]).value) is not None), 2)
        name = source._header(ws, ref, first_filled, c)
        stacked = _first_stacked_table(ws, c, limit, numeric)
        if stacked:
            d = Counter(str(core._merged_value(ws, r, c)).strip() for r in data_rows if stacked[0] <= r <= stacked[1] and core._merged_value(ws, r, c) not in (None, "") and core.as_number(core._merged_value(ws, r, c)) is None)
        repeats = sum(d.values()) > len(d) * 1.5
        spec = f"{get_column_letter(c)}{stacked[0]}:{get_column_letter(c)}{stacked[1]}" if stacked else get_column_letter(c)
        base = Base("name", spec, name, many_rows=repeats, count=len(d))
        bases.append((c, base, p, d))
    if not bases:
        return []
    out = []
    # Labels spread across a few columns on the left, as in a statement with indented lines.
    left = [b for b in bases if b[0] < numeric[0]]
    if len(left) >= 2:
        rows = [r for r in range(1, limit + 1) if core.as_number(ws.cell(r, numeric[0]).value) is not None]
        best = max(sum(ws.cell(r, b[0]).value not in (None, "") for r in rows) for b in left) if rows else 0
        if rows and best < len(rows) * 0.8:
            lo, hi = left[0][0], left[-1][0]
            filled = [sum(ws.cell(r, c).value not in (None, "") for c in range(lo, hi + 1)) for r in rows]
            combined = sum(count > 0 for count in filled)
            # An indented statement has one label on each line. 지역, 시설, 유형 and 구분 on one line are levels, and the last of them does not name the row.
            if combined >= best and sum(count == 1 for count in filled) >= combined * 0.8:
                out.append((-1, Base("name", f"{get_column_letter(lo)}:{get_column_letter(hi)}", left[0][1].name, count=len(rows))))
    for c, base, p, d in bases:
        filled_numeric = 0
        sample = range(1, limit + 1)
        filled_numeric = sum(ws.cell(r, c).value not in (None, "") and core.as_number(ws.cell(r, numeric[0]).value) is not None for r in sample)
        # 월계 and 누계 beside thirty regions are kinds of figure, not the things a report compares.
        kinds = base.many_rows and base.count <= 6 and base.count < max(other.count for _c, other, _p, _d in bases)
        order = (0 if not base.many_rows else 1, int(kinds), -filled_numeric)
        out.append((order, base))
    ordered = [base for _o, base in sorted(out, key=lambda item: (item[0] if isinstance(item[0], tuple) else (-1, 0)))]
    for base in ordered:
        for c in numeric:
            ref = source._parse_ref(f"{ws.title}!{_name_line(base.spec)}", "x")
            first_filled = next((r for r in range(1, limit + 1) if core.as_number(ws.cell(r, c).value) is not None), 2)
            name = source._header(ws, ref, first_filled, c)
            numbers = [core.as_number(ws.cell(r, c).value) for r in range(1, limit + 1)]
            years = [ws.cell(r, c).value for r in data_rows if ws.cell(r, c).value not in (None, "")]
            if c in codes or not _names_a_figure(name) or _steps_by_one(numbers) or _long_identifiers(numbers) or _unlike_things(numbers):
                continue
            if years and all(core.as_year(raw) is not None for raw in years):  # 회계연도 2025 on every line of a budget
                continue
            base.values.append(Value(get_column_letter(c), name))
        for c, other, p, d in bases:
            if other.spec != base.spec and d and 2 <= len(d) <= MAX_GROUP_VALUES and sum(d.values()) > len(d):
                base.groups.append(Group(_name_line(other.spec), other.name, list(d)))
        _newest_period_first(ws, base)
    return ordered


def _newest_period_first(ws, base):
    """Columns headed 2019 to 2025: a report is about the newest, not the first."""
    periods = {}
    for value in base.values:
        if value.spec.isalpha():
            period = _period_of(ws, core.column_index_from_string(value.spec), value.name)
            if period:
                periods[value.spec] = period
    if len(periods) >= 2 and len(periods) == len(base.values):
        base.values.sort(key=lambda value: periods[value.spec], reverse=True)


def read_shape(source, title):
    """How a sheet is laid out: its dates and figures, or its named items and figures."""
    ws = source._sheet(title)
    shape = Shape(ws.title)
    codes = _code_columns(ws)
    for spec, by_row in _date_candidates(ws):
        # B on its own is not another reading of dates that A+B has already used.
        if any(spec in re.split(r"[+:]", b.spec) for b in shape.bases if "+" in b.spec):
            continue
        base = _read_date_base(source, ws, spec, by_row, codes)
        if base is NESTED:  # 남 and 여 over 사망자수 and 부상자수 over each year: any one column would be a guess, and every other reading of this sheet takes one too
            return Shape(ws.title)
        if base and base.values and not any(b.spec.split("+")[0] == base.spec.split("+")[0] for b in shape.bases):
            shape.bases.append(base)
    shape.bases.extend(b for b in _read_name_bases(source, ws, codes) if b.values and not any(o.spec.split("+")[0] == b.spec for o in shape.bases if o.kind == "date"))
    shape.glob = _joined_sheets(source, ws, shape)
    return shape


def _joined_sheets(source, ws, shape):
    """A pattern for sheets that continue this one, when they read as one series: one sheet per year, say."""
    dated = next((b for b in shape.bases if b.kind == "date"), None)
    if not dated or not dated.values or ":" in dated.spec or dated.many_rows and not dated.groups:
        return ""
    others = [other for other in source.wb.worksheets if other is not ws and not NOT_DATA_SHEET.search(other.title)]
    best = ""
    for size in range(len(ws.title), 1, -1):
        prefix = ws.title[:size].casefold()
        if any(other.title.casefold().startswith(prefix) for other in others):
            best = ws.title[:size]
            break
    if not best or not re.search(r"\d|\D\D", best):
        return ""
    pattern = best + "*"
    value = next((v for v in dated.values if not v.every_row), None)
    if value is None:
        return ""
    try:
        joined = source._series(source._parse_ref(f"{pattern}!{dated.spec},{value.spec}", pattern))
        alone = source._series(source._parse_ref(f"{ws.title}!{dated.spec},{value.spec}", pattern))
    except core.IssueError:
        if not dated.many_rows:
            return ""
        return _joined_by_sum(source, pattern, ws, dated, value)
    return pattern if len(joined.points) > len(alone.points) else ""


def _joined_by_sum(source, pattern, ws, dated, value):
    spec = f"{word('sum', 'en')}({value.spec})"
    try:
        joined = source._series(source._parse_ref(f"{pattern}!{dated.spec},{spec}", pattern))
        alone = source._series(source._parse_ref(f"{ws.title}!{dated.spec},{spec}", pattern))
    except core.IssueError:
        return ""
    return pattern if len(joined.points) > len(alone.points) else ""


def default_value(base):
    """The figure a report most likely wants: a total if one is named, else the first. One that stops before the sheet's newest period gives way to one that does not: a programme that ended in 2021-22 says little about a file that runs to 2024-25."""
    plain = [v for v in base.values]
    if not plain:
        return None
    every_row = next((v for v in plain if v.every_row), None)
    if every_row:
        return every_row
    named_total = [v for v in plain if TOTAL_NAME.search(v.name)]
    ends_as_total = [v for v in named_total if ENDS_AS_TOTAL.search(v.name)]
    chosen = next(iter(ends_as_total), None) or next(iter(named_total), None) or plain[0]
    if chosen.reaches_newest:
        return chosen
    for totals in (ends_as_total, named_total):
        running = [v for v in totals if v.reaches_newest]
        if running:
            return max(running, key=lambda v: v.size)  # a grand total is at least as large as any part of it
    return next((v for v in plain if v.reaches_newest), chosen)


def _is_level(name):
    """A heating value, a concentration or a temperature: something each row has, not an amount the rows hold between them. Adding them up says only how many rows there are."""
    return bool(name) and LEVEL_NAME.search(re.sub(r"\s+", "", name)) is not None and not AVERAGE_OVER_TIME.search(name)


def default_how(base, value):
    """How rows join: as they are when each date or name has one, else a mean for prices and rates, and a sum for the rest."""
    if value is not None and value.every_row:
        return "sum"
    if value is not None and value.counted:
        return "count"
    if not base.many_rows:
        return "plain"
    if value is not None and AVERAGE_OVER_TIME.search(value.name):
        return "sum"
    if value is not None and MEAN_NAME.search(value.name):
        return "mean"
    if value is not None and NEVER_ADDED_NAME.search(value.name):
        return "count" if base.kind == "name" else "mean"
    if value is not None and _is_level(value.name):
        return "mean"
    # Rows of named things, such as staff by department: the column could be anything, an age or a salary, so it is only added if it says it is an amount.
    if base.kind == "name" and not (value is not None and AMOUNT_NAME.search(value.name)) and not _splits_by_group(base):
        return "count"
    return "sum"


def _splits_by_group(base):
    """A few repeating kinds beside the names, such as 월계 and 누계, or a row marked 소계 among many: each name's rows differ by kind, so one kind makes each name one figure."""
    return any(len(group.values) <= 6 or any(_marks_a_whole(word) for word in group.values) for group in base.groups)


def _marks_a_whole(word):
    """A word for all the rows beside it. A comma would end the condition early when it is written out."""
    return ALL_ROWS_WORD.fullmatch(word.strip()) is not None and "," not in word


def _whole_by_its_figures(ws, ref, buckets, group, value):
    """The word whose rows equal all the others, date after date: 모두 beside 남, 여 and 미상. Adding it to them would count everything twice, whatever it is called."""
    line, figure = core.column_index_from_string(group.spec), core.column_index_from_string(value.spec)
    sums = {}
    for bucket, position in buckets:
        word = ws.cell(*ref.at(position, line)).value
        number = core.as_number(ws.cell(*ref.at(position, figure)).value)
        if word not in (None, "") and number is not None:
            by_word = sums.setdefault(bucket, {})
            word = str(word).strip()
            by_word[word] = by_word.get(word, 0) + number
    if len(group.values) < WHOLE_MIN_VALUES or len(sums) < WHOLE_MIN_GROUPS:
        return None
    for word in group.values[:MAX_GROUP_VALUES]:
        seen = [by_word for by_word in sums.values() if by_word.get(word) and len(by_word) >= WHOLE_MIN_VALUES]
        matches = sum(abs(by_word[word] - (sum(by_word.values()) - by_word[word])) <= abs(by_word[word]) * WHOLE_TOLERANCE for by_word in seen)
        if "," not in word and len(seen) >= WHOLE_MIN_GROUPS and matches >= len(seen) * WHOLE_SHARE:
            return word
    return None


def _whole_rows(source, shape, base, value):
    """(column, word) for each column that has a row for all the others, by what it is called or else by what it adds up to."""
    named = {group.spec: next((word for word in group.values[:MAX_GROUP_VALUES] if _marks_a_whole(word)), None) for group in base.groups}
    found = {spec: word for spec, word in named.items() if word}
    unnamed = [group for group in base.groups if group.spec not in found and group.spec.isalpha()]
    if unnamed and value is not None and value.spec.isalpha():
        ws = source._sheet(shape.sheet)
        if base.kind == "date":
            ref = source._parse_ref(f"{shape.sheet}!{base.spec}", base.spec)
            buckets = source._dated_rows(ws, ref)[:SAMPLE_ROWS]
        else:
            ref = source._parse_ref(f"{shape.sheet}!{_name_line(base.spec)}", base.spec)
            buckets = _named_rows(ws, base)
        for group in unnamed:
            word = _whole_by_its_figures(ws, ref, buckets, group, value)
            if word:
                found[group.spec] = word
    return [(group.spec, found[group.spec]) for group in base.groups if group.spec in found]


def _leaves_items(ws, base, spec, word):
    """총계 merged across 지역 and 국가 is only the sheet's total line: choosing it would leave no countries to compare."""
    if not spec.isalpha():
        return True
    line = core.column_index_from_string(spec)
    names = {name for name, row in _named_rows(ws, base) if str(core._merged_value(ws, row, line)).strip() == word and name != word}
    return len(names) >= 2


def _splits_names(ws, base, spec):
    """Most names have more than one value in this column, so their rows differ by it."""
    if not spec.isalpha():
        return False
    line = core.column_index_from_string(spec)
    found = {}
    for name, row in _named_rows(ws, base):
        word = core._merged_value(ws, row, line)
        if word not in (None, ""):
            found.setdefault(name, set()).add(str(word).strip())
    return bool(found) and sum(len(words) >= 2 for words in found.values()) >= len(found) * SPLITS_SHARE


def _name_line(spec):
    """The column of a name base's spec: A from A, A:C or A6:A12."""
    return re.match(r"[A-Za-z]+", spec).group(0)


def _named_rows(ws, base):
    bounded = core.BOUNDED_LINE.fullmatch(base.spec)
    first, last = (int(bounded.group(2)), int(bounded.group(4))) if bounded else (1, min(ws.max_row, SAMPLE_ROWS))
    column = core.column_index_from_string(_name_line(base.spec))
    rows = ((core._merged_value(ws, row, column), row) for row in range(first, last + 1))
    return [(str(name).strip(), row) for name, row in rows if name not in (None, "")]


def _first_stacked_table(ws, column, limit, numeric):
    """(first row, last row) of the first table when one heading comes back further down with the same names under it. Read together, 동구지회 in three tables is three rows of one name."""
    cells = [(row, str(raw).strip()) for row in range(1, limit + 1) if isinstance(raw := ws.cell(row, column).value, str) and raw.strip()]
    where = {}
    for row, word in cells:
        where.setdefault(word, []).append(row)
    for row, word in cells:
        if len(where[word]) < 2 or any(core.as_number(ws.cell(row, line).value) is not None for line in numeric):
            continue
        next_heading = where[word][1]
        following = where[word][2] if len(where[word]) > 2 else limit + 1
        first = {other for at, other in cells if row < at < next_heading}
        second = {other for at, other in cells if next_heading < at < following}
        if len(first) >= 2 and len(first & second) >= len(first) * STACKED_SHARE:
            filled = [at for at, _other in cells if row < at < next_heading]
            return row + 1, filled[-1]
    return None


def default_conditions(source, shape, base, value, how):
    """A mean across different things means nothing, so one thing is chosen, one column at a time, until each date has one row."""
    if how in ("sum", "mean") and _across(base):  # each year spans 사고건수, 사망자수 and more: the first is what the rest depend on
        return [(group.spec, group.values[0]) for group in base.groups]
    if how == "sum" and base.kind == "date":  # All Health Authorities beside each authority: add both and everything counts twice
        wholes = _whole_rows(source, shape, base, value)
        taken = {spec for spec, _word in wholes}
        return (wholes + [(group.spec, group.values[0]) for group in base.groups if group.spec not in taken and MEASURE_COLUMN.search(group.name)])[:3]
    if base.kind == "name" and base.many_rows and how == "sum":
        ws = source._sheet(shape.sheet)
        wholes = [(spec, word) for spec, word in _whole_rows(source, shape, base, value) if _leaves_items(ws, base, spec, word)]
        taken = {spec for spec, _word in wholes}
        # 월계 and 누계, or 건수 and 연면적, under one name must not be added together. 대륙 goes with each country and splits nothing.
        kinds = [(group.spec, group.values[0]) for group in base.groups if group.spec not in taken and len(group.values) <= 6 and _splits_names(ws, base, group.spec) and _leaves_items(ws, base, group.spec, group.values[0])]
        measures = [kind for kind, group in zip(kinds, [g for g in base.groups if g.spec in {spec for spec, _word in kinds}]) if MEASURE_COLUMN.search(group.name)]
        chosen = kinds[:1] + [kind for kind in measures if kind not in kinds[:1]]  # one kind is enough to make a name one figure: 정기 and 부정기 can still add up, unless they are different measures
        return (wholes + chosen)[:3]
    if how != "mean" or not base.groups or base.kind != "date":
        return []
    ws = source._sheet(shape.sheet)
    ref = source._parse_ref(f"{shape.sheet}!{base.spec}", base.spec)
    rows = [(date, position) for date, position in source._dated_rows(ws, ref)][:SAMPLE_ROWS]
    chosen = []
    # Each company's heating value can be averaged, so 회사 is left alone. BOD beside COD cannot, so a column of kinds still narrows it to one.
    level_only = value is not None and _is_level(value.name) and not MEAN_NAME.search(value.name)
    for group in base.groups:
        if len({date for date, _p in rows}) == len(rows):
            break
        if level_only and not KIND_COLUMN.search(group.name):
            continue
        line = core.column_index_from_string(group.spec)
        narrowed = [(d, p) for d, p in rows if str(ws.cell(*ref.at(p, line)).value).strip() == group.values[0]]
        if narrowed:
            rows = narrowed
            chosen.append((group.spec, group.values[0]))
    return chosen


def _across(base):
    """Periods laid across, with headings under them: its groups are rows, not columns."""
    return bool(base.groups) and base.groups[0].spec.isdigit()


def value_spec(value, how, conditions, lang):
    """C, or 합계(C, B=민원)."""
    extra = [f"{group}={text}" for group, text in conditions if group and text]
    if how == "plain" and not extra:
        return value.spec
    if how == "rate":
        first, *rest = extra
        return f"{word('rate', lang)}({', '.join([first] + rest)})"
    name = "sum" if how == "plain" else how
    return f"{word(name, lang)}({', '.join([value.spec] + extra)})"


def body(shape, base, value, how="plain", conditions=(), lang="ko", joined=True):
    sheet = shape.glob if joined and shape.glob and base.kind == "date" else shape.sheet
    return f"{sheet}!{base.spec},{value_spec(value, how, conditions, lang)}"


def _period_of(ws, line, name):
    """(year, month) a column belongs to, from its heading, or from a heading above it such as a merged 2024년도."""
    words = [name]
    for row in range(1, 6):
        for column in range(line, max(line - 3, 0), -1):
            raw = ws.cell(row, column).value
            if raw not in (None, ""):
                words.append(str(raw))
                break
    for text in words:
        match = re.search(r"((?:19|20)\d\d)(?:[-./년 ]+(\d{1,2})(?!\d))?", text)
        if match:
            return int(match.group(1)), int(match.group(2) or 0)
    return None


def two_periods(source, shape, base):
    """The (older, newer) value columns of a table that sets two periods side by side, when its headings say so."""
    if base.kind != "name":
        return None
    ws = source._sheet(shape.sheet)
    dated = []
    for value in base.values:
        if value.spec.isalpha():
            period = _period_of(ws, core.column_index_from_string(value.spec), value.name)
            bare = re.sub(r"(?:19|20)\d\d[-./년 \d]*", "", value.name).strip()
            if period:
                dated.append((period, bare, value.spec))
    if len(dated) < 2:
        return None
    # Only columns of the same kind: 수량 against 수량, not 수량 against 비율. The first such pair, as the main figure comes first.
    for _period, kind, spec in dated:
        same = sorted((d for d in dated if d[1] == kind), key=lambda d: d[0])
        if len({d[0] for d in same}) >= 2:
            return same[-2][2], same[-1][2]
    return None


def _plan(base, lang, reference):
    """What can be asked of a reference, in the order a report wants it."""
    if base.kind == "date":
        return (("latest", f"{word('latest', lang)}:{reference}", "text"), ("note", f"{word('note', lang)}:{reference}", "text"), ("reading", f"{word('reading', lang)}:{reference}", "text"),
                ("prev", f"{word('prev', lang)}:{reference}", "text"), ("week", f"{word('week', lang)}:{reference}", "text"), ("avg", f"{word('avg', lang)}7:{reference}", "text"),
                ("peak", f"{word('peak', lang)}30:{reference}", "text"), ("table", f"{word('table', lang)}:{word('latest', lang)}6:{reference}", "alt"),
                ("chart", f"{word('chart', lang)}:{word('latest', lang)}12:{reference}", "alt"))
    return (("reading", f"{word('reading', lang)}:{reference}", "text"), ("note", f"{word('note', lang)}:{reference}", "text"), ("top", f"{word('top', lang)}1:{reference}", "text"),
            ("topvalue", f"{word('topvalue', lang)}1:{reference}", "text"), ("share", f"{word('share', lang)}1:{reference}", "text"),
            ("table_share", f"{word('table', lang)}:{word('share', lang)}5:{reference}", "alt"), ("table_top", f"{word('table', lang)}:{word('top', lang)}5:{reference}", "alt"),
            ("chart_top", f"{word('chart', lang)}:{word('top', lang)}5:{reference}", "alt"))


def _try(source, code, key, place):
    """Run a key through the report's own code. A suggestion is only as good as its real result."""
    found = Suggestion(code, key, place)
    try:
        if place == "alt":
            marker = core.MARKER.fullmatch("{{" + key + "}}")
            rows, _where, issues = source.range_cells(marker.group(2))
            found.columns = len(rows[0])
            found.rows = len(rows) - 1
            found.result = f"{len(rows) - 1} × {len(rows[0])}"
            issues = [issue for issue in issues if issue.code != "latest_fewer"]  # the size just shown says so already
        else:
            text, _where, issues = source.text_value(key)
            found.result = text or ""
        found.issues = list(issues)
    except core.IssueError as exc:
        found.issues = [exc.issue]
    if any(issue.is_error for issue in found.issues):
        found.status = "error"
    elif found.issues:
        found.status = "warn"
    return found


def suggestions(source, shape, base, value=None, how=None, conditions=(), lang="ko", joined=True):
    """Everything that works for this choice, with its real result. What fails is left out, except when all of it fails, so the reason can be seen."""
    value = value or default_value(base)
    if value is None:
        return []
    how = how or default_how(base, value)
    conditions = list(conditions) or default_conditions(source, shape, base, value, how)
    reference = body(shape, base, value, how, conditions, lang, joined)
    made = [_try(source, code, key, place) for code, key, place in _plan(base, lang, reference)]
    change = two_periods(source, shape, base)
    if change and not conditions:
        made.insert(0, _try(source, "change", f"{word('reading', lang)}:{shape.sheet}!{base.spec},{change[0]},{change[1]}", "text"))
    elif change and change[1] == value.spec and how == "sum":  # 월계 beside 누계: the same condition on both years
        older = value_spec(Value(change[0], ""), how, conditions, lang)
        made.insert(0, _try(source, "change", f"{word('reading', lang)}:{shape.sheet}!{base.spec},{older},{value_spec(value, how, conditions, lang)}", "text"))
    # By item, in the newest period: one reference a long table of dates and names usually wants.
    group = next((g for g in base.groups if not ID_NAME.search(g.name)), None) if not _across(base) else None  # the days of the week are no one's idea of items
    if base.kind == "date" and group and base.date_spec and base.newest_word and how != "rate":
        extra = list(conditions) + [(base.date_spec, base.newest_word)]
        by_item = f"{shape.sheet}!{group.spec},{value_spec(value, how if how != 'plain' else 'sum', extra, lang)}"
        made.append(_try(source, "by_item", f"{word('reading', lang)}:{by_item}", "text"))
        made.append(_try(source, "table_by_item", f"{word('table', lang)}:{word('share', lang)}5:{by_item}", "alt"))
    good = [item for item in made if item.status != "error"]
    return good or made[:1]


def first_choice(source, title=None, lang="ko"):
    """(shape, base, list of suggestions) for the sheet a person most likely means, trying each way of reading it until one works."""
    sheets = [source._sheet(title)] if title else data_sheets(source)
    for ws in sheets:
        shape = read_shape(source, ws.title)
        for base in shape.bases:
            found = suggestions(source, shape, base, lang=lang)
            if any(item.status != "error" for item in found):
                return shape, base, found
    return None, None, []


def find_doubts(source, shape, base, value, found):
    """Why this choice may not be what the workbook means. Rules cannot foresee every table, so where the signs are poor a person is asked, instead of a wrong figure going into a report unsaid."""
    ws = source._sheet(shape.sheet)
    how = default_how(base, value)
    name = value.name or ""
    doubts = []

    def add(code):
        if code not in [known for known, _params in doubts]:
            doubts.append((code, {"figure": name}))

    if EXAMPLE_SHEET.search(shape.sheet):
        add("example_sheet")
    if how == "count" or value.counted:
        add("count_only")
    numbers = _figure_rows(ws, value)
    if how == "sum" and not AVERAGE_OVER_TIME.search(name) and base.many_rows and _nearly_the_same(list(numbers.values())):  # a level by its name is averaged already
        add("sum_of_levels")
    if core.BOUNDED_LINE.fullmatch(base.spec) or any(code == "note_first_table" for code, _params in base.notes):
        add("first_table")
    if value.spec.isalpha() and _heading_below_first_figure(ws, value, numbers):
        add("heading_below")
    if base.kind == "date" and value.spec.isalpha() and _figures_without_dates(source, ws, shape, base, numbers):
        add("dates_sparse")
    if how not in ("mean", "count") and not facts.is_rate(name) and not core.in_per_cent(ws) and _says_per_cent(ws):  # a note that says so outright is simply believed
        add("maybe_percent")
    last_row, last_column = min(ws.max_row, SAMPLE_ROWS), min(ws.max_column, TITLE_COLUMNS)  # each asks every cell on the sheet, so once only
    if sum(isinstance(ws.cell(row, column).value, str) and TABLE_LINE.match(ws.cell(row, column).value.strip()) is not None for row in range(1, last_row + 1) for column in range(1, last_column + 1)) >= 2:
        add("several_tables")
    codes = {issue.code for item in found for issue in item.issues}
    if "cross_total_differs" in codes:
        add("total_differs")
    return doubts


def _figure_rows(ws, value):
    if not value.spec.isalpha():
        return {}
    column = core.column_index_from_string(value.spec)
    found = {row: core.as_number(ws.cell(row, column).value) for row in range(1, min(ws.max_row, SAMPLE_ROWS) + 1)}
    return {row: number for row, number in found.items() if number is not None}


def _nearly_the_same(numbers):
    """A heating value of 42.3 to 42.6 from every company: a level each of them has, and adding it up says only how many there are."""
    if len(numbers) < LEVEL_MIN_ROWS:
        return False
    middle = sum(numbers) / len(numbers)
    return bool(middle) and (sum((number - middle) ** 2 for number in numbers) / len(numbers)) ** 0.5 <= abs(middle) * LEVEL_SPREAD


def _heading_below_first_figure(ws, value, numbers):
    """A summary line above the real table: the name read for this column came from over the wrong table."""
    if len(numbers) < HEADING_MIN_ROWS:
        return False
    column = core.column_index_from_string(value.spec)
    rows = sorted(numbers)
    for row in range(rows[0] + 1, rows[-1]):
        raw = ws.cell(row, column).value
        if isinstance(raw, str) and core.as_number(raw) is None and re.search(r"[^\W\d]", raw):
            before = sum(place < row for place in rows)
            return before <= HEADING_MAX_ABOVE and len(rows) - before >= before * 2
    return False


def _figures_without_dates(source, ws, shape, base, numbers):
    ref = source._parse_ref(f"{shape.sheet}!{base.spec}", base.spec)
    if ref.by_row or len(numbers) < LEVEL_MIN_ROWS:
        return False
    dated = {position for _date, position in source._dated_rows(ws, ref)}
    figure_rows = [row for row in numbers if row > min(dated, default=0)]
    missing = sum(row not in dated and core._merged_value(ws, row, ref.date) in (None, "") for row in figure_rows)
    return bool(figure_rows) and missing >= len(figure_rows) * DATES_MISSING_SHARE


def _says_per_cent(ws):
    return any(isinstance(ws.cell(row, column).value, str) and PER_CENT_NOTE.fullmatch(ws.cell(row, column).value.strip()) is not None for row in range(1, min(ws.max_row, HEAD_ROWS) + 1) for column in range(1, min(ws.max_column, TITLE_COLUMNS) + 1))


def _added_or_shown(base, figures):
    """Rows of people counted under a column named Age: what was taken is a count, and the window must not say it was the age."""
    return [figure for figure in figures if default_how(base, figure) != "count"]


def describe(shape, *values):
    """Which figures were taken, in words a person can check against the workbook."""
    names = ", ".join(facts.readable_name(value.name) for value in values if value and not (value.every_row or value.counted))
    return " · ".join(part for part in (shape.sheet, names) if part)


def labels(found):
    """(what it is, which figure) for each suggestion, by the key the report lists it under, so the list can be read without the notation."""
    named = {}
    for item in found:
        marker = core.MARKER.fullmatch(item.typed) if item.place == "alt" else None
        named[(marker.group(2) if marker else item.key).casefold()] = (item.code, item.figure)
    return named


def heading(source, shape, title=""):
    """The sheet's name and any unit written above its table: without them 192 says nothing about what was counted."""
    ws = source._sheet(shape.sheet)
    parts = []
    if not PLAIN_SHEET_NAME.fullmatch(shape.sheet.strip()) and shape.sheet.casefold() not in title.casefold():
        parts.append(shape.sheet)
    if not any(UNIT_NOTE.search(part) for part in parts):
        parts += _unit_words(ws)
    return " · ".join(parts)


def _period_words(ws, value):
    """The time a table with no dates is for, when only its title says: 2026년 6월 건축 통계표. Two dates in a title are a span, and a figure named for its year says so itself."""
    if value is not None and YEAR_IN_NAME.search(value.name or ""):
        return ""
    for row in range(1, min(ws.max_row, PERIOD_TITLE_ROWS) + 1):
        for column in range(1, min(ws.max_column, PERIOD_TITLE_COLUMNS) + 1):
            cell = ws._cells.get((row, column))
            text = cell.value if cell is not None and isinstance(cell.value, str) else ""
            found = PERIOD_IN_TITLE.findall(text) if len(text) <= PERIOD_TITLE_LENGTH else []
            if len(found) == 1:
                return " ".join(found[0].split())
    return ""


def _unit_words(ws):
    """A note such as (단위: 백만원) near the top, or a column headed 단위 that says the same thing in every row."""
    note = core.unit_note(ws, UNIT_ROWS, UNIT_COLUMNS)[0]
    if note:
        return [note]
    for row in range(1, min(ws.max_row, UNIT_ROWS) + 1):
        for column in range(1, min(ws.max_column, UNIT_COLUMNS) + 1):
            raw = ws.cell(row, column).value
            if not isinstance(raw, str):
                continue
            text = raw.strip()
            if UNIT_HEADING.fullmatch(text):
                below = {str(ws.cell(place, column).value).strip() for place in range(row + 1, min(ws.max_row, row + SAMPLE_ROWS) + 1) if ws.cell(place, column).value not in (None, "")}
                if len(below) == 1 and len(next(iter(below))) <= UNIT_VALUE_LENGTH:
                    return [f"{text}: {next(iter(below))}"]
    return []


def _figures(base, value):
    """The figures a report needs to be read fairly. 내국인 alone reads as all sales when 외국인 stands beside it; a named total already speaks for its parts."""
    plain = [v for v in base.values if not (v.every_row or v.counted)]
    if value.every_row or value.counted or base.many_rows or TOTAL_NAME.search(value.name) or not 2 <= len(plain) <= MAX_FIGURES:
        return [value]
    if all(YEAR_IN_NAME.search(v.name) for v in plain):  # 2025년 and 2024년 are one figure at two times, so the newest speaks for them
        return [value]
    return plain


def _comparable(figures, made):
    """Figures that can share a chart's axis. Put 3 beside 3,000,000 and the small one is a flat line."""
    size = [core.as_number(next((i.result for i in items if i.code == "latest"), None)) for items in made]
    if not size[0]:
        return figures[:1]
    return [figure for figure, number in zip(figures, size) if number and min(abs(number), abs(size[0])) * CHART_RANGE >= max(abs(number), abs(size[0]))]


def _section(source, shape, base, value, found, lang, figures=None):
    """One sheet's part of the report: a sentence or two for each figure, then one table and one chart holding them."""
    figures = figures or _figures(base, value)
    if value not in figures:
        value = figures[0]
        found = suggestions(source, shape, base, value, lang=lang)
    figures = [value] + [figure for figure in figures if figure is not value]
    made = [found if figure is value else suggestions(source, shape, base, figure, lang=lang) for figure in figures]
    for figure, items in zip(figures, made):
        for item in items:
            item.figure = figure.name if not (figure.every_row or figure.counted) else ""
    words = heading(source, shape)
    period = _period_words(source._sheet(shape.sheet), value) if base.kind == "name" else ""
    if period and period not in words:
        words = " · ".join(part for part in (period, words) if part)
    sentences = []
    for number, items in enumerate(made):
        chosen = _sentence_items({item.code: item for item in items if item.status != "error"})
        if number:
            chosen = [item for item in chosen if item.code == "note" or (len(figures) <= READ_ALL_FIGURES and item.code == "reading")]
        sentences += chosen
    table, chart = _table_and_chart({item.code: item for item in found if item.status != "error"})
    if any(item.code == "reading" and item.result in facts.SEPARATE_LISTS.values() for item in sentences):
        sentences = [item for item in sentences if item.code != "reading"]  # it says only that nothing can be said, which a template's author should know and a reader need not
        chart = None  # bars would rank 자산 against 부채
    if len(figures) > 1 and base.kind == "date" and not base.many_rows:
        table = _together(source, shape, base, figures[:MAX_TABLE_FIGURES], "table", word("latest", lang) + "6", lang) or table
        together = _comparable(figures, made)[:MAX_TABLE_FIGURES]
        chart = _together(source, shape, base, together, "chart", word("latest", lang) + "12", lang) if len(together) > 1 else chart
        chart = chart or _table_and_chart({item.code: item for item in found if item.status != "error"})[1]
    return Section(words, sentences, table, chart, figures, sheet=shape.sheet, period=period)


def section_for(source, shape, base, figures, found, lang="ko"):
    """A section of exactly these figures, as someone ticked them. found is copied, as it may be used again for other choices."""
    value = default_value(base)
    return _section(source, shape, base, value, [replace(item) for item in found], lang, figures)


def _together(source, shape, base, figures, kind, rows, lang):
    """A table or chart of several figures, or None when they will not go together."""
    reference = body(shape, base, figures[0], lang=lang) + "".join("," + figure.spec for figure in figures[1:])
    made = _try(source, kind, f"{word(kind, lang)}:{rows}:{reference}", "alt")
    return made if made.status != "error" else None


def _same_layout(source, ws, base, value, lang):
    """(shape, base, value, suggestions) when a sheet is built like the first, such as sales in one sheet and customers in the next. Others may be something else entirely."""
    try:
        shape = read_shape(source, ws.title)
    except core.IssueError:  # one odd sheet is no reason to give up on the report
        return None
    other = next((b for b in shape.bases if b.kind == base.kind and b.spec == base.spec), None)
    if other is None or {v.name for v in other.values} != {v.name for v in base.values}:
        return None
    if {g.name for g in other.groups} != {g.name for g in base.groups}:  # the first sheet broken down by 성별: the same figures again, not more of them
        return None
    figure = next(v for v in other.values if v.name == value.name)
    found = suggestions(source, shape, other, figure, lang=lang)
    return (shape, other, figure, found) if any(item.status != "error" for item in found) else None


def plan_report(source, lang="ko"):
    """What a report made from the workbook alone should hold, or None when nothing can be read."""
    shape, base, found = first_choice(source, lang=lang)
    if shape is None:
        return None
    value = default_value(base)
    sections = [_section(source, shape, base, value, found, lang)]
    what = [describe(shape, *_added_or_shown(base, sections[0].figures))]
    rest = [ws for ws in data_sheets(source) if ws.title != shape.sheet and not (shape.glob and fnmatch.fnmatchcase(ws.title.casefold(), shape.glob.casefold()))]
    alike = []
    for ws in rest:
        same = _same_layout(source, ws, base, value, lang) if ws.max_row <= SMALL_SHEET else None
        if same:
            alike.append(same)
        if len(alike) >= MAX_SECTIONS:  # so many are more likely months or regions than different measures, and any few of them would be arbitrary
            alike = []
            break
    for other_shape, other_base, other_value, other_found in alike:
        sections.append(_section(source, other_shape, other_base, other_value, other_found, lang))
        what.append(describe(other_shape, *_added_or_shown(other_base, sections[-1].figures)))
    plain = [v for v in base.values if not (v.every_row or v.counted)]
    return Plan(sections, " / ".join(what), max(len(plain) - len(sections[0].figures), 0), len(rest) - len(alike), doubts=find_doubts(source, shape, base, value, found))


def _marked(key, number):
    """A key that says which file it reads: 해설:@1@시트!A,B. A sheet's name cannot hold a colon, so the last one before the ! is where the name starts."""
    head, bang, tail = key.partition("!")
    before, colon, sheet = head.rpartition(":")
    return f"{before}{colon}@{number}@{sheet}{bang}{tail}"


def _for_file(section, number, title):
    section.file, section.title = number, title
    for item in section.items:
        item.key = _marked(item.key, number)
    return section


def overview(sections, names, lang="ko", source=None):
    """A first slide for several files: what can be calculated across them, then each one's main sentence under its name. A chosen sentence can take the first place."""
    sentences, references, seen = [], [], set()
    for section in sections:
        if section.file in seen:
            continue
        lead = next((item for item in section.sentences if item.code == "note"), None) or next(iter(section.sentences), None)
        if lead is not None:
            seen.add(section.file)
            sentences.append(replace(lead, label=" · ".join(part for part in (names[section.file], section.period) if part), figure=""))
            references.append(lead.key.partition(":")[2])
    together = "|".join(references)
    key = f"{word('insight', lang)}:{core.ALL_FILES}{together}"
    has_documents = bool(getattr(source, "documents", None))
    if source is not None and len(references) >= 2:
        reading = f"{word('reading', lang)}:{core.ALL_FILES}{together}"
        try:
            text = source.text_value(reading)[0]
        except core.IssueError:
            text = None
        if text:
            sentences.insert(0, Suggestion("cross", reading, "text", result=text))
    return Section("", sentences, cross=key if len(references) >= 2 or (has_documents and references) else "", file=-1)


def document_items(sources, events):
    """The figures in the workbooks that the document's entries name: only those are read, as a workbook can hold thousands."""
    text = " ".join(event.text for event in events)
    items = []
    for number, source in enumerate(sources.sources):
        for ws in data_sheets(source):
            if ws.max_row > SMALL_SHEET and not ws.max_column > SMALL_SHEET:
                continue
            try:
                shape = read_shape(source, ws.title)
            except core.IssueError:
                continue
            base = next((b for b in shape.bases if b.kind == "date"), None)
            if base is None:
                continue
            if base.many_rows:  # a long table: a column of names beside the dates
                figure = default_value(base)
                how = default_how(base, figure) if figure else "sum"
                if figure is None or figure.counted or _across(base) or how not in ("sum", "mean"):
                    continue
                named = [(f"{word(how, 'en')}({figure.spec}, {group.spec}={value})", value) for group in base.groups if not ID_NAME.search(group.name) for value in group.values if documents.mentions(text, value)]
            else:
                named = [(value.spec, value.name) for value in base.values if value.name and not value.every_row and not value.counted and not TOTAL_NAME.search(value.name) and documents.mentions(text, value.name)]
            for spec, name in named[:MAX_DOCUMENT_ITEMS]:
                reference = f"{shape.sheet}!{base.spec},{spec}"
                try:
                    points = source._series(source._parse_ref(reference, reference)).points
                except core.IssueError:
                    continue
                items.append(documents.Item(name, shape.sheet, points, number))
    return items


def document_section(sources, number, lang="ko"):
    """One slide for a document: how it lines up with the figures, with a wide table of what happened when."""
    document = sources.documents[number]
    links = sources.links()[number][1]
    note = f"{word('note', lang)}:{core.DOCUMENT_MARK}{number}"
    reading = f"{word('reading', lang)}:{core.DOCUMENT_MARK}{number}"
    sentences = []
    for code, key in (("doc_note", note), ("doc_reading", reading)):
        text = sources.text_value(key)[0]
        if text:
            sentences.append(Suggestion(code, key, "text", result=text))
    if not links:  # nothing to line up, so the document speaks for itself, above what was checked
        quoted = [f"{note}#{place}" for place in range(1, len(documents.key_paragraphs(document)) + 1)]
        sentences[:0] = [Suggestion("doc_quote", key, "text", result=sources.text_value(key)[0]) for key in quoted]
    table = None
    if links:
        rows = min(len(links), core.DOCUMENT_LINKS)
        table = Suggestion("table_doc", f"{word('table', lang)}:{'ko' if lang == 'ko' else 'en'}{core.DOCUMENT_MARK}{number}:{rows}", "alt", columns=6, rows=rows)
    return Section("", sentences, table, title=document.title, file=-2, wide=True)


def plan_files(sources, lang="ko"):
    """A report from several workbooks: an overview, then each file's slides, with what could not be read said plainly."""
    sections, what, unread, figures_left, sheets_left, doubts = [], [], [], 0, 0, []
    for number, (source, name) in enumerate(zip(sources.sources, sources.names)):
        plan = plan_report(source, lang)
        if plan is None:
            unread.append(name)
            continue
        for place, section in enumerate(plan.sections):
            _for_file(section, number, name if place == 0 else f"{name} · {section.heading}".rstrip(" ·"))
        sections += plan.sections
        what.append(f"[{name}] {plan.what}")
        figures_left += plan.figures_left
        sheets_left += plan.sheets_left
        doubts += [(code, {**params, "file": name}) for code, params in plan.doubts]
    if not sections:
        return None
    return Plan(with_documents(sections, sources, lang), " ".join(what), figures_left, sheets_left, unread, doubts)


def with_documents(sections, sources, lang="ko"):
    """The overview first, then a slide for each document, then each workbook's slides: the reader meets what happened before the detail."""
    pages = [document_section(sources, number, lang) for number in range(len(sources.documents))]
    first, placed = [], {}
    for document, page in zip(sources.documents, pages):
        workbook = _workbook_beside(document, sources)
        if workbook is None:
            first.append(page)
        else:
            placed.setdefault(workbook, []).append(page)
            if len(sources.sources) > 1:
                page.title = f"{sources.names[workbook]} · {page.title}"
    ordered, seen = [], set()
    for section in sections:
        if section.file not in seen:
            seen.add(section.file)
            ordered += placed.get(section.file, [])
        ordered.append(section)
    ordered += [page for workbook, group in placed.items() if workbook not in seen for page in group]
    return [overview(sections, sources.names, lang, sources)] + first + ordered


def _workbook_beside(document, sources):
    """The one workbook kept in the document's folder, which it most likely describes. None when files all share a folder and nothing says which."""
    folder = os.path.dirname(os.path.abspath(document.path))
    beside = [number for number, source in enumerate(sources.sources) if os.path.dirname(os.path.abspath(source.path)) == folder]
    return beside[0] if len(beside) == 1 else None


def _slide_words(title, sections):
    """The words a report's template will hold, so the language is judged as it will be in the report."""
    parts = [title]
    for section in sections:
        parts += [section.title, section.heading] + [item.label for item in section.sentences]
    return " ".join(part for part in parts if part)


def _data_words(source, sections):
    """A report made from the workbook alone is in the language of what the sheets hold, not of a sheet's name. One letter stands for each cell, so 삼성전자 does not lose to Samsung Electronics for being shorter."""
    books = getattr(source, "sources", [source])
    korean = other = 0
    for section in sections:
        if not section.sheet:
            continue
        cells = books[section.file]._sheet(section.sheet)._cells.values()  # only cells that exist: a sheet's range can reach far past them
        for cell in islice((cell for cell in cells if isinstance(cell.value, str)), LANGUAGE_SAMPLE):
            if facts.HANGUL.search(cell.value):
                korean += 1
            elif facts.LATIN.search(cell.value):
                other += 1
    return " " + "가" * korean + " " + "a" * other


def _reference(item):
    return item.key.partition(":")[2].casefold()


def ask_insights(source, sections, title, lang="ko"):
    """What a model could be asked about each figure, by the reference its sentences share: {reference: (request, figure)}. Also what it could be asked about several files together."""
    words = _slide_words(title, sections) + _data_words(source, sections)
    asks = {}
    for section in sections:
        if section.cross:
            try:
                request = source.insight_request(section.cross, words)[0]
            except core.IssueError:
                request = None
            if request:
                asks[core.ALL_FILES] = (request, "")
        for item in section.sentences:
            if item.code not in ("reading", "insight"):
                continue
            key = item.key if item.code == "insight" else f"{word('insight', lang)}:{item.key.partition(':')[2]}"
            try:
                request = source.insight_request(key, words)[0]
            except core.IssueError:
                continue
            if request:
                asks[_reference(item)] = (request, item.figure)
    return asks


def choose_insights(sections, asks, insights):
    """Sections with the sentence a person chose in place of the calculated reading, for exactly these figures. If the figures changed, the calculated reading stands."""
    chosen = []
    for section in sections:
        sentences = []
        for item in section.sentences:
            request = asks.get(_reference(item), (None,))[0] if item.code == "reading" else None
            saved = insights.get(request.key.casefold()) if request else None
            if saved and saved.get("text") and saved.get("fingerprint") == request.fingerprint:
                sentences.append(Suggestion("insight", request.key, "text", result=saved["text"], figure=item.figure))
            else:
                sentences.append(item)
        request = asks.get(core.ALL_FILES, (None,))[0] if section.cross else None
        saved = insights.get(request.key.casefold()) if request else None
        if saved and saved.get("text") and saved.get("fingerprint") == request.fingerprint:
            sentences = [item for item in sentences if item.code != "cross"]
            sentences.insert(0, Suggestion("insight", request.key, "text", result=saved["text"]))
        chosen.append(replace(section, sentences=sentences))
    return chosen


def one_section(found, words=""):
    by_code = {item.code: item for item in found if item.status != "error"}
    return Section(words, _sentence_items(by_code), *_table_and_chart(by_code))


def _sentence_items(by_code):
    """A bare figure says nothing the sentence does not, so it is only used when there is no sentence."""
    return [by_code[code] for code in ("note", "reading", "change", "by_item") if code in by_code] or [by_code[code] for code in ("latest", "top") if code in by_code]


def _table_and_chart(by_code):
    return by_code.get("table") or by_code.get("table_share") or by_code.get("table_top"), by_code.get("chart") or by_code.get("chart_top")


def in_report(found):
    """Which of these suggestions a report made from them holds."""
    by_code = {item.code: item for item in found if item.status != "error"}
    return _sentence_items(by_code) + [item for item in _table_and_chart(by_code) if item]


def starter_path(folder, excel):
    """A new file beside the workbook. Nothing is ever overwritten."""
    stem = os.path.splitext(os.path.basename(excel))[0]
    path = os.path.join(folder, f"{stem}_template.pptx")
    number = 2
    while os.path.exists(path):
        path = os.path.join(folder, f"{stem}_template_{number}.pptx")
        number += 1
    return path


def make_starter(path, title, found, words=""):
    """A plain first template from one figure's suggestions, each already marked."""
    return make_report(path, title, [one_section(found, words)])


def make_report(path, title, sections):
    """One slide for each section: its title, the sentences on the left, the chart and table on the right, so it can be read at a glance. What does not fit goes on to another slide rather than off the bottom of this one."""
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    _use_font(prs)
    for number, section in enumerate(sections):
        width = TEXT_WIDTH if (section.table or section.chart) and not section.wide else FULL_WIDTH
        pages = _pages(_fit(section.sentences, width), width)
        for place, page in enumerate(pages):
            slide = prs.slides.add_slide(prs.slide_layouts[6])
            words = section.title or (title if number == 0 else section.heading or title)
            shown = f"{words} ({place + 1}/{len(pages)})" if len(pages) > 1 else words
            _text(slide, LEFT, 0.28, FULL_WIDTH, 0.6, shown, 26, TITLE_COLOR, bold=True)
            if (number == 0 or section.title) and section.heading and section.heading not in words:
                _text(slide, LEFT, 0.86, FULL_WIDTH, 0.35, section.heading, 13, MUTED)
            line = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(LEFT + 0.1), Inches(RULE_AT), Inches(LEFT + FULL_WIDTH - 0.1), Inches(RULE_AT))
            line.line.color.rgb = TITLE_COLOR
            line.line.width = Pt(1.25)
            y = CONTENT_TOP
            for item in page:
                size = _size(item)
                quiet = item.code in QUIET_CODES
                block = _height(item, width)
                top = y
                if item.label:
                    _text(slide, LEFT, top, width, 0.3, item.label, 12, MUTED)
                    top += LABEL_HEIGHT
                _text(slide, LEFT, top, width, y + block - top + SENTENCE_GAP, item.typed, size, BODY_COLOR if quiet else TITLE_COLOR, bold=not quiet)
                y += block + SENTENCE_GAP
            if section.wide and place == len(pages) - 1:
                _add_wide_table(slide, section, y)
            elif not section.wide and place == 0:
                _add_table_and_chart(slide, section)
    prs.save(path)
    return path


def _add_wide_table(slide, section, top):
    if not section.table:
        return
    frame = slide.shapes.add_table(2, section.table.columns, Inches(LEFT + 0.1), Inches(top), Inches(FULL_WIDTH - 0.2), Inches(WIDE_ROW * 2))
    frame._element.xpath(".//p:cNvPr")[0].set("descr", section.table.typed)
    for column, width in zip(frame.table.columns, WIDE_COLUMNS):
        column.width = Inches(width)
    for row in frame.table.rows:
        row.height = Inches(WIDE_ROW)
        for cell in row.cells:
            cell.margin_top = cell.margin_bottom = Inches(0.03)
            cell.text_frame.paragraphs[0]._p.get_or_add_endParaRPr().set("sz", str(WIDE_POINTS * 100))


def _use_font(prs):
    """One font for sentences, tables and charts: every Windows has this one, and it sets Korean and digits alike."""
    for part in prs.slide_master.part.rels.values():
        if part.reltype == RT.THEME:
            theme = part.target_part
            theme._blob = re.sub(rb'(<a:latin typeface=")[^"]*(")', rb"\g<1>Malgun Gothic\g<2>", theme.blob)
            break


def title_of(excel, sections):
    """The workbook's name as words, unless it says nothing, when the first sheet's heading says more."""
    stem = re.sub(r"_+", " ", os.path.splitext(os.path.basename(excel))[0]).strip()
    if PLAIN_SHEET_NAME.fullmatch(stem) and sections and sections[0].heading:
        return sections[0].heading
    return stem


def preview_rows(source, item):
    """The rows a table or chart will show, for drawing a picture of the slide. None when they cannot be read."""
    marker = core.MARKER.fullmatch(item.typed) if item else None
    if not marker:
        return None
    try:
        return source.range_cells(marker.group(2))[0]
    except core.IssueError:
        return None


def _add_table_and_chart(slide, section):
    room = SLIDE_BOTTOM - CONTENT_TOP
    table_height = TABLE_ROW * (section.table.rows + 1) if section.table else 0
    if section.chart:
        data = CategoryChartData()
        data.categories = ["A", "B"]
        data.add_series("S", [1, 2])
        kind = XL_CHART_TYPE.LINE_MARKERS if section.chart.code == "chart" else XL_CHART_TYPE.BAR_CLUSTERED
        height = room - table_height - (0.15 if section.table else 0)
        frame = slide.shapes.add_chart(kind, Inches(RIGHT), Inches(CONTENT_TOP), Inches(RIGHT_WIDTH), Inches(height), data)
        frame._element.xpath(".//p:cNvPr")[0].set("descr", section.chart.typed)
        chart = frame.chart
        chart.font.size = Pt(11)
        chart.plots[0].vary_by_categories = False  # one figure, one color: the names are already on the axis
        chart.has_title = False  # the legend says the same, and has room for several figures
        chart.has_legend = True
        chart.legend.position = XL_LEGEND_POSITION.BOTTOM
        chart.legend.include_in_layout = False
        chart.value_axis.tick_labels.number_format = AXIS_FORMAT
        chart.value_axis.tick_labels.number_format_is_linked = False
        chart.value_axis.major_gridlines.format.line.color.rgb = GRID
    if section.table:
        top = SLIDE_BOTTOM - table_height
        frame = slide.shapes.add_table(2, section.table.columns, Inches(RIGHT), Inches(top), Inches(RIGHT_WIDTH), Inches(TABLE_ROW * 2))
        frame._element.xpath(".//p:cNvPr")[0].set("descr", section.table.typed)
        for row in frame.table.rows:
            row.height = Inches(TABLE_ROW)
            for cell in row.cells:
                cell.margin_top = cell.margin_bottom = Inches(0.03)
                cell.text_frame.paragraphs[0]._p.get_or_add_endParaRPr().set("sz", str(TABLE_POINTS * 100))  # filling a cell keeps this


def _size(item):
    return 14 if item.code in QUIET_CODES else 18


def _height(item, width):
    """Roughly how tall a sentence is, so the next one starts below it."""
    size = _size(item)
    lines = max(1, -(-len(item.result or "") * size // int(SLIDE_CHARACTERS * width / FULL_WIDTH)))
    return lines * size * LINE_HEIGHT / 72 + (LABEL_HEIGHT if item.label else 0)


def _fit(sentences, width):
    """What fits on the slide. The main sentence of every figure stays: leaving one out would make the others read as the whole."""
    kept = list(sentences)
    room = SLIDE_BOTTOM - CONTENT_TOP
    for item in reversed(sentences):
        if sum(_height(i, width) + SENTENCE_GAP for i in kept) <= room:
            break
        if item.code in ("reading", "change", "by_item"):
            kept = [i for i in kept if i is not item]
    return kept


def _pages(sentences, width):
    """The sentences for each slide. A first slide holding every file's main sentence is the one that fills up, and none of them may be lost."""
    room = SLIDE_BOTTOM - CONTENT_TOP
    pages, used = [[]], 0
    for item in sentences:
        need = _height(item, width) + SENTENCE_GAP
        if pages[-1] and used + need > room:
            pages.append([])
            used = 0
        pages[-1].append(item)
        used += need
    return pages


def with_doubt_notes(sections, doubts):
    """Sections that say on the slide what a person was warned of and chose to keep: (file name or "", words). Whoever reads the report did not see the window."""
    done = [replace(section, sentences=list(section.sentences)) for section in sections]
    for name, words in doubts:
        section = next((s for s in done if s.sheet and (not name or s.title == name)), None)
        if section is not None:
            section.sentences.append(Suggestion("doubt_note", "", "plain", result=words))
    return done


def _text(slide, left, top, width, height, words, size, color=None, bold=False):
    box = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    box.text_frame.word_wrap = True
    box.text_frame.text = words
    for run in box.text_frame.paragraphs[0].runs:
        run.font.size = Pt(size)
        run.font.bold = bold
        if color:
            run.font.color.rgb = color
