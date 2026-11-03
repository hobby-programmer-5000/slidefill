"""Reads how a sheet is laid out and offers references that work on it, each with its real result. No GUI, no network, no model.

A person should not have to learn A10:A475 or 합계(D, C=도시가스). Every reference offered here has been run through the same
code as the report, so what is listed is what will be filled, and nothing that fails is offered as if it worked.
"""

import datetime as dt
import os
import re
from collections import Counter
from dataclasses import dataclass, field

from openpyxl.utils.cell import get_column_letter
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.util import Inches, Pt

import core
import facts

SAMPLE_ROWS = 3000  # enough to tell what a column holds; a sheet of a million rows need not be read twice for this
SAMPLE_COLUMNS = 60
HEAD_ROWS = 25  # periods laid out across a row are near the top
MAX_GROUP_VALUES = 200
MAX_NAME_BASES = 6  # a sheet of many text columns need not be tried in full
MIN_PERIODS = 3  # two dates are more likely two headings, 2025년 and 2024년 side by side, than a series
MANY_ROWS = 0.7  # fewer distinct dates than this share of dated rows: several rows per date, not a stray repeat
ID_NAME = re.compile(r"코드|번호|순번|연번|일련|시간|요일|\b(?:id|code|no\.?|time|hour|weekday)\b", re.I)
TOTAL_NAME = re.compile(r"총계|합계|총공급|종합|전체|\b(?:total|overall)\b", re.I)
AMOUNT_NAME = re.compile(r"금액|매출|수입|비용|수량|판매|사용량|공급량|\b(?:amount|sales|revenue|cost|income|quantity|qty)\b", re.I)
ALL_ROWS_WORD = re.compile(r"all\b.*|total\b.*|.*\btotal|합계|총계|전체|계", re.I)
EVERY_ROW_MIN = 20  # so many rows each with a figure are items of one kind; a few are more likely different measures
MEAN_NAME = re.compile(r"평균|가격|단가|율|률|점수|\b(?:price|rate|ratio|average|mean|score)\b", re.I)
NOT_DATA_SHEET = re.compile(r"표지|목차|설명|안내|주석|index|contents|cover|notes?\b|copyright", re.I)
YEAR_IN_NAME = re.compile(r"(?:19|20)\d\d")
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

    @property
    def typed(self):
        return "{{" + self.key + "}}"


def data_sheets(source):
    """The sheets worth offering, best first: the newest by the year in its name, or the one the workbook opens on."""
    found = []
    for ws in source.wb.worksheets:
        if ws.max_row < 2 or NOT_DATA_SHEET.search(ws.title):
            continue
        cells = [ws.cell(r, c).value for r in range(1, min(ws.max_row, 60) + 1) for c in range(1, min(ws.max_column, SAMPLE_COLUMNS) + 1)]
        if sum(core.as_number(v) is not None for v in cells) >= 3:
            found.append(ws)
    active = source.wb.active
    years = {ws.title: max((int(y) for y in YEAR_IN_NAME.findall(ws.title)), default=0) for ws in found}
    return sorted(found, key=lambda ws: (-years[ws.title], ws is not active, source.wb.worksheets.index(ws)))


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


def _read_date_base(source, ws, spec, by_row):
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
        if count >= max(len(filled) * 0.5, 2) and not ID_NAME.search(name) and not years_only:
            (stepping if _steps_by_one(numbers) else base.values).append(Value(_name(line, by_row), name))
        elif not by_row:
            words = Counter(str(v).strip() for v in filled if core.as_number(v) is None)
            if 2 <= len(words) <= MAX_GROUP_VALUES and sum(words.values()) > len(words):
                base.groups.append(Group(_name(line, by_row), name, list(words)))
    groups_under = bool(under) and len(under) >= len(sample) * 0.5
    base.many_rows = many or groups_under
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


def _steps_by_one(numbers):
    """A running number, not a figure."""
    found = [n for n in numbers if n is not None]
    return len(found) >= 5 and all(later - earlier == 1 for earlier, later in zip(found, found[1:]))


def _read_name_bases(source, ws):
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
        repeats = sum(d.values()) > len(d) * 1.5
        base = Base("name", get_column_letter(c), name, many_rows=repeats, count=len(d))
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
            combined = sum(any(ws.cell(r, c).value not in (None, "") for c in range(lo, hi + 1)) for r in rows)
            if combined >= best:
                out.append((-1, Base("name", f"{get_column_letter(lo)}:{get_column_letter(hi)}", left[0][1].name, count=len(rows))))
    for c, base, p, d in bases:
        filled_numeric = 0
        sample = range(1, limit + 1)
        filled_numeric = sum(ws.cell(r, c).value not in (None, "") and core.as_number(ws.cell(r, numeric[0]).value) is not None for r in sample)
        order = (0 if not base.many_rows else 1, -filled_numeric)
        out.append((order, base))
    ordered = [base for _o, base in sorted(out, key=lambda item: (item[0] if isinstance(item[0], tuple) else (-1, 0)))]
    for base in ordered:
        for c in numeric:
            ref = source._parse_ref(f"{ws.title}!{base.spec.split(':')[0]}", "x")
            first_filled = next((r for r in range(1, limit + 1) if core.as_number(ws.cell(r, c).value) is not None), 2)
            name = source._header(ws, ref, first_filled, c)
            if not ID_NAME.search(name) and not _steps_by_one([core.as_number(ws.cell(r, c).value) for r in range(1, limit + 1)]):
                base.values.append(Value(get_column_letter(c), name))
        for c, other, p, d in bases:
            if other.spec != base.spec and d and 2 <= len(d) <= MAX_GROUP_VALUES and sum(d.values()) > len(d):
                base.groups.append(Group(other.spec, other.name, list(d)))
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
    for spec, by_row in _date_candidates(ws):
        # B on its own is not another reading of dates that A+B has already used.
        if any(spec in re.split(r"[+:]", b.spec) for b in shape.bases if "+" in b.spec):
            continue
        base = _read_date_base(source, ws, spec, by_row)
        if base and base.values and not any(b.spec.split("+")[0] == base.spec.split("+")[0] for b in shape.bases):
            shape.bases.append(base)
    shape.bases.extend(b for b in _read_name_bases(source, ws) if b.values and not any(o.spec.split("+")[0] == b.spec for o in shape.bases if o.kind == "date"))
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
    """The figure a report most likely wants: a total if one is named, else the first."""
    plain = [v for v in base.values]
    if not plain:
        return None
    return next((v for v in plain if v.every_row), None) or next((v for v in plain if TOTAL_NAME.search(v.name)), None) or plain[0]


def default_how(base, value):
    """How rows join: as they are when each date or name has one, else a mean for prices and rates, and a sum for the rest."""
    if value is not None and value.every_row:
        return "sum"
    if value is not None and value.counted:
        return "count"
    if not base.many_rows:
        return "plain"
    if value is not None and MEAN_NAME.search(value.name):
        return "mean"
    # Rows of named things, such as staff by department: the column could be anything, an age or a salary, so it is only added if it says it is an amount.
    if base.kind == "name" and not (value is not None and AMOUNT_NAME.search(value.name)) and not _splits_by_group(base):
        return "count"
    return "sum"


def _splits_by_group(base):
    """A few repeating kinds beside the names, such as 월계 and 누계: each name's rows differ by kind, so one kind makes each name one figure."""
    return any(len(group.values) <= 6 for group in base.groups)


def default_conditions(source, shape, base, value, how):
    """A mean across different things means nothing, so one thing is chosen, one column at a time, until each date has one row."""
    if how == "sum" and base.kind == "date":  # All Health Authorities beside each authority: add both and everything counts twice
        return [(group.spec, word) for group in base.groups for word in group.values[:MAX_GROUP_VALUES] if ALL_ROWS_WORD.fullmatch(word.strip())][:3]
    if base.kind == "name" and base.many_rows and how == "sum":  # 월계 and 누계 under one name must not be added together
        return [(group.spec, group.values[0]) for group in base.groups if len(group.values) <= 6][:1]
    if how != "mean" or not base.groups or base.kind != "date":
        return []
    ws = source._sheet(shape.sheet)
    ref = source._parse_ref(f"{shape.sheet}!{base.spec}", base.spec)
    rows = [(date, position) for date, position in source._dated_rows(ws, ref)][:SAMPLE_ROWS]
    chosen = []
    for group in base.groups:
        if len({date for date, _p in rows}) == len(rows):
            break
        line = core.column_index_from_string(group.spec)
        narrowed = [(d, p) for d, p in rows if str(ws.cell(*ref.at(p, line)).value).strip() == group.values[0]]
        if narrowed:
            rows = narrowed
            chosen.append((group.spec, group.values[0]))
    return chosen


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
    # By item, in the newest period: one reference a long table of dates and names usually wants.
    if base.kind == "date" and base.groups and base.date_spec and base.newest_word and how != "rate":
        group = base.groups[0]
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


def starter_path(folder, excel):
    """A new file beside the workbook. Nothing is ever overwritten."""
    stem = os.path.splitext(os.path.basename(excel))[0]
    path = os.path.join(folder, f"{stem}_template.pptx")
    number = 2
    while os.path.exists(path):
        path = os.path.join(folder, f"{stem}_template_{number}.pptx")
        number += 1
    return path


def make_starter(path, title, found):
    """A plain first template from the suggestions: the title, the sentences, a table and a chart, each already marked."""
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    by_code = {item.code: item for item in found if item.status != "error"}
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _text(slide, 0.5, 0.35, 12.3, 0.8, title, 28)
    y = 1.4
    for code in ("latest", "top", "note", "reading", "by_item"):
        if code in by_code:
            _text(slide, 0.5, y, 12.3, 1.2, by_code[code].typed, 24 if code in ("latest", "top") else 16)
            y += 1.0 if code in ("latest", "top") else 1.35
    table = by_code.get("table") or by_code.get("table_share") or by_code.get("table_top")
    chart = by_code.get("chart") or by_code.get("chart_top")
    if table or chart:
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        if table:
            frame = slide.shapes.add_table(2, table.columns, Inches(0.5), Inches(0.5), Inches(5.6), Inches(0.9))
            frame._element.xpath(".//p:cNvPr")[0].set("descr", table.typed)
        if chart:
            data = CategoryChartData()
            data.categories = ["A", "B"]
            data.add_series("S", [1, 2])
            kind = XL_CHART_TYPE.LINE_MARKERS if chart.code == "chart" else XL_CHART_TYPE.BAR_CLUSTERED
            frame = slide.shapes.add_chart(kind, Inches(6.4), Inches(0.5), Inches(6.4), Inches(6.4), data)
            frame._element.xpath(".//p:cNvPr")[0].set("descr", chart.typed)
    prs.save(path)
    return path


def _text(slide, left, top, width, height, words, size):
    box = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    box.text_frame.word_wrap = True
    box.text_frame.text = words
    for run in box.text_frame.paragraphs[0].runs:
        run.font.size = Pt(size)
