"""Reads Word documents and lines what they say up against the figures. All calculation, no model and no GUI.

A document can say when something happened and to what. The figures say what the numbers did around then. This puts the two side by side.
That they happened at the same time is all it can know: why is not in the data, and nothing here says so.
"""

import datetime as dt
import re
import statistics
import zipfile
from dataclasses import dataclass, field

from lxml import etree

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
MONTHS = ("january", "february", "march", "april", "may", "june", "july", "august", "september", "october", "november", "december")
MONTH_WORD = "|".join(MONTHS + tuple(month[:3] for month in MONTHS))
DATE_AT_START = re.compile(
    rf"^\s*(?:(?P<day1>\d{{1,2}})\s+)?(?P<month1>{MONTH_WORD})\.?,?\s+(?:(?P<day2>\d{{1,2}}),\s*)?(?P<year1>(?:19|20)\d\d)"
    r"|^\s*(?P<year2>(?:19|20)\d\d)년\s*(?P<month2>\d{1,2})월(?:\s*(?P<day3>\d{1,2})일)?"
    r"|^\s*(?P<year3>(?:19|20)\d\d)[.\-/]\s*(?P<month3>\d{1,2})(?:[.\-/]\s*(?P<day4>\d{1,2}))?\.?(?!\d)", re.I)
AFTER_DATE = re.compile(r"^[\s:：,.)\-–—]+")
NUMBER_IN_TEXT = re.compile(r"(?<![\d.,])(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d+)?(?![\d])")
HEADING_STYLE = re.compile(r"heading|title|제목", re.I)
SKIPPED_STYLE = re.compile(r"^toc|^contents", re.I)
LATIN_NAME = re.compile(r"[A-Za-z]")
MIN_NAME = 2
LONG_PARAGRAPH = 80
MAX_CONTINUATION = 200
SMALLEST_SHARE = 0.02  # below this share of the largest, a sharp percentage is a handful of people
SMALLEST_FIGURE = 1000  # 133 and 113 turn up in any workbook by chance, so only figures large enough to mean something when found are looked for
LITTLE_OVERLAP = 0.3  # finding fewer than this says the document and the workbooks are about different things
KEY_PARAGRAPHS = 3
NOTE_START = re.compile(r"^\W*(?:note|notes|참고|주|비고)\b", re.I)
POINTS_ELSEWHERE = re.compile(r"\b(?:table|figure|chart|appendix)\s*\d+|(?:표|그림|부록)\s*\d+|(?:below|above|아래|위)\s*(?:table|figure|표|그림)", re.I)
GLOSSARY_ENTRY = re.compile(r"^[^.:：]{2,40}[:：]\s")  # Australian Fishing Zone: Waters adjacent to Australia...
CAUTION = re.compile(r"care should be taken|caution|caveat|limitation|유의|주의|한계|해석할\s*때", re.I)
STATES_A_FIGURE = re.compile(r"\d{1,3}(?:,\d{3})+|\d+(?:\.\d+)?\s*(?:%|퍼센트|명|건|원|톤|kg|억|만)")


class NotAWordFile(Exception):
    pass


@dataclass
class Event:
    date: dt.datetime
    text: str


@dataclass
class Document:
    path: str
    name: str
    title: str = ""
    paragraphs: list = field(default_factory=list)
    tables: list = field(default_factory=list)
    events: list = field(default_factory=list)


@dataclass
class Item:
    """One figure in a workbook, with its points in time: objects with .date and .value."""
    name: str
    sheet: str
    points: list
    file: int = 0


@dataclass
class Link:
    event: Event
    item: Item
    before: float
    after: float
    window: int

    @property
    def change(self):
        return (self.after / self.before - 1) * 100 if self.before > 0 else None

    @property
    def moved(self):
        return abs(self.after - self.before)


def read(path, name):
    """The words of a .docx, by zip and XML only: nothing in the file can run, and no outside entity is fetched."""
    try:
        with zipfile.ZipFile(path) as package:
            raw = package.read("word/document.xml")
    except (zipfile.BadZipFile, KeyError) as exc:
        raise NotAWordFile(str(exc)) from exc
    root = etree.fromstring(raw, etree.XMLParser(resolve_entities=False, no_network=True, huge_tree=False))
    body = root.find(W + "body")
    document = Document(path, name)
    lines = []
    for element in body if body is not None else []:
        if element.tag == W + "p":
            style = element.find(f"{W}pPr/{W}pStyle")
            style = style.get(W + "val") if style is not None else ""
            text = _paragraph_text(element)
            if text and not SKIPPED_STYLE.search(style or ""):
                lines.append((text, bool(HEADING_STYLE.search(style or ""))))
                if not document.title and style and HEADING_STYLE.search(style):
                    document.title = text
        elif element.tag == W + "tbl":
            document.tables.append([[" ".join(_paragraph_text(p) for p in cell.iter(W + "p")).strip() for cell in row.findall(W + "tc")] for row in element.findall(W + "tr")])
    document.paragraphs = [text for text, _heading in lines]
    document.title = document.title or name
    document.events = _events(lines)
    return document


def _paragraph_text(paragraph):
    parts = []
    for node in paragraph.iter():
        if node.tag == W + "t":
            parts.append(node.text or "")
        elif node.tag in (W + "tab", W + "br"):
            parts.append(" ")
    return " ".join("".join(parts).split())


def _date_of(match):
    groups = match.groupdict()
    if groups["year1"]:
        word = groups["month1"].lower()
        month = next(place for place, name in enumerate(MONTHS, start=1) if name.startswith(word[:3]))
        year, day = int(groups["year1"]), groups["day1"] or groups["day2"]
    elif groups["year2"]:
        year, month, day = int(groups["year2"]), int(groups["month2"]), groups["day3"]
    else:
        year, month, day = int(groups["year3"]), int(groups["month3"]), groups["day4"]
    try:
        return dt.datetime(year, month, int(day or 1))
    except ValueError:
        return None


def _events(lines):
    """Paragraphs that start with a date. A short line right after one, with no date of its own, belongs to the same date if another dated line follows it."""
    found = []
    for place, (text, heading) in enumerate(lines):
        match = None if heading else DATE_AT_START.match(text)
        date = _date_of(match) if match else None
        if date:
            words = AFTER_DATE.sub("", text[match.end():]).strip()
            if words:
                found.append(Event(date, words))
            continue
        ended = heading or not found or place == 0 or not DATE_AT_START.match(lines[place - 1][0])
        following = place + 1 < len(lines) and DATE_AT_START.match(lines[place + 1][0])
        if not ended and following and len(text) <= MAX_CONTINUATION:
            found.append(Event(found[-1].date, text))
    return found


def mentions(text, name):
    """A name as a whole word in Latin letters, or anywhere in Korean, which has no spaces to go by before its endings."""
    name = name.strip()
    if len(name) < MIN_NAME:
        return False
    if LATIN_NAME.search(name):
        return re.search(rf"(?<![A-Za-z]){re.escape(name)}(?![A-Za-z])", text, re.I) is not None
    return name in text


def _window(points):
    """How many periods make a fair comparison: a year's worth where seasons repeat, so the same seasons fall on both sides."""
    gaps = [(later.date - earlier.date).days for earlier, later in zip(points, points[1:])]
    gap = statistics.median(gaps) if gaps else 0
    if gap <= 1:
        return 7, gap
    if gap <= 8:
        return 4, gap
    if gap <= 35:
        return 12, gap
    if gap <= 100:
        return 4, gap
    return 3, gap


def link(event, item):
    """The figure's mean over the periods before and after the event. The period it falls in is left out, as it holds some of both. None if there is too little to compare."""
    points = item.points
    if len(points) < 3:
        return None
    window, gap = _window(points)
    yearly = gap > 100
    at = next((place for place, point in enumerate(points) if (point.date.year >= event.date.year if yearly else point.date >= event.date)), None)
    if at is None or at < window or at + window >= len(points):
        return None
    before, after = points[at - window:at], points[at + 1:at + 1 + window]
    for side in (before, after):  # a gap in the data would stretch the comparison over more time than it says
        if (side[-1].date - side[0].date).days > (window - 1) * gap * 1.5 + 1:
            return None
    mean_before, mean_after = sum(p.value for p in before) / window, sum(p.value for p in after) / window
    if not mean_before and not mean_after:
        return None
    return Link(event, item, mean_before, mean_after, window)


def link_all(events, items):
    """Every event against every figure it names, largest move first, keeping one line for a figure and event found on two sheets."""
    best = {}
    for event in events:
        for item in items:
            if mentions(event.text, item.name):
                found = link(event, item)
                if found:
                    key = (event.date, event.text, item.name.casefold(), item.file)
                    if key not in best or found.moved > best[key].moved:
                        best[key] = found
    return sorted(best.values(), key=lambda found: -found.moved)


def highlights(links, count):
    """What a slide should show: the largest moves in people, and the sharpest in percent. Sorted by people alone, a route that lost nearly everyone never appears, because it was small."""
    if len(links) <= count:
        return list(links)
    half = count // 2
    chosen = list(links[:count - half])
    largest = max(found.before for found in links) or 1
    sharp = sorted((found for found in links if found not in chosen and found.change is not None and found.before >= largest * SMALLEST_SHARE), key=lambda found: -abs(found.change))
    chosen += sharp[:half]
    chosen += [found for found in links if found not in chosen][:count - len(chosen)]
    return sorted(chosen, key=lambda found: found.event.date)


def period_text(date, lang):
    return f"{date.year}년 {date.month}월" if lang == "ko" else f"{MONTHS[date.month - 1].capitalize()[:3]} {date.year}"


def change_text(found):
    return f"{found.change:+.1f}%" if found.change is not None else f"0 → {found.after:,.0f}"


def key_paragraphs(document):
    """What a document with nothing to line up says for itself, in its own words and its own order: a caution about reading the figures if it gives one, then the first long paragraphs, those stating a figure before those that do not. A pointer to a table not on the slide or an entry in a glossary means nothing there."""
    usable = [text for text in document.paragraphs if len(text) >= LONG_PARAGRAPH and not NOTE_START.match(text) and not POINTS_ELSEWHERE.search(text) and not GLOSSARY_ENTRY.match(text)]
    cautions = [text for text in usable if CAUTION.search(text)][:1]
    rest = [text for text in usable if text not in cautions]
    chosen = cautions + [text for text in rest if STATES_A_FIGURE.search(text)] + [text for text in rest if not STATES_A_FIGURE.search(text)]
    chosen = chosen[:KEY_PARAGRAPHS]
    return [text for text in usable if text in chosen]


def headline(document, links, lang):
    """The first thing to say: how much of the document could be lined up, and the largest move. None when it has no dated entries, and its own words will do."""
    total = len(document.events)
    if not links:
        if total:
            return (f"문서의 날짜 기록 {total}건 중 수치의 항목 이름과 맞는 것이 없습니다." if lang == "ko"
                    else f"None of the document's {total} dated entries name a figure in the workbooks.")
        return None
    top = links[0]
    events = len({(found.event.date, found.event.text) for found in links})
    if lang == "ko":
        return (f"문서의 날짜 기록 {total}건 중 {events}건이 수치의 항목과 맞습니다. 전후 변화가 가장 큰 곳은 {top.item.name}, "
                f"{period_text(top.event.date, lang)} 전 {top.window}개 기간 평균 {top.before:,.0f}에서 후 {top.window}개 기간 평균 {top.after:,.0f}로 {change_text(top)}입니다.")
    return (f"{events} of the document's {total} dated entries name a figure in the workbooks. The largest move is {top.item.name} around {period_text(top.event.date, lang)}: "
            f"a mean of {top.before:,.0f} over the {top.window} periods before, {top.after:,.0f} over the {top.window} after, {change_text(top)}.")


def caution(lang):
    return ("문서에 적힌 일과 수치의 변화가 같은 시기라는 뜻이며, 그 일이 원인인지는 이 자료만으로 알 수 없습니다." if lang == "ko"
            else "This says only that the entry and the change came at the same time. Whether one caused the other is not in these files.")


def table_rows(links, count, lang):
    """Rows of (raw, text, format), as a range from a workbook would be."""
    ko = lang == "ko"
    window = links[0].window if links else 4
    head = ("시기", "항목", "문서 기록", f"전 {window}개 평균", f"후 {window}개 평균", "변화") if ko else ("When", "Figure", "In the document", f"Mean of {window} before", f"Mean of {window} after", "Change")
    rows = [[(word, word, "General") for word in head]]
    for found in highlights(links, count):
        cells = (period_text(found.event.date, lang), found.item.name, found.event.text, f"{found.before:,.0f}", f"{found.after:,.0f}", change_text(found))
        rows.append([(word, word, "General") for word in cells])
    return rows


def figures_in(document):
    """Numbers the document states, as (text, value). Years and small numbers are left out: they would match anything."""
    found = {}
    for text in document.paragraphs + [cell for table in document.tables for row in table for cell in row]:
        for match in NUMBER_IN_TEXT.finditer(text):
            digits = match.group(0).replace(",", "")
            try:
                value = float(digits)
            except ValueError:
                continue
            is_year = "," not in match.group(0) and value.is_integer() and 1900 <= value <= 2100
            if value >= SMALLEST_FIGURE and not is_year:
                found.setdefault(match.group(0), value)
    return found


def check_figures(document, workbook_values):
    """Which of the document's numbers appear in the workbooks. Not finding one does not make it wrong: it may be a sum, or from elsewhere."""
    found, missing = [], []
    for text, value in figures_in(document).items():
        (found if value in workbook_values or round(value / 100, 6) in workbook_values else missing).append(text)
    return found, missing


def figures_sentence(found, missing, lang):
    total = len(found) + len(missing)
    if not total:
        return ""
    little = len(found) < total * LITTLE_OVERLAP
    if lang == "ko":
        words = f"문서에 적힌 큰 수치 {total}개 중 {len(found)}개를 엑셀에서 찾았습니다."
        if little:
            return words + " 문서와 엑셀이 서로 다른 범위를 다루는 것으로 보이니, 함께 해석할 때 주의하세요."
        return words + (f" 찾지 못한 수치: {', '.join(missing[:5])}{' 등' if len(missing) > 5 else ''}." if missing else "")
    words = f"{len(found)} of the {total} large figures stated in the document were found in the workbooks."
    if little:
        return words + " The document and the workbooks seem to cover different ground, so take care reading them together."
    return words + (f" Not found: {', '.join(missing[:5])}{' and more' if len(missing) > 5 else ''}." if missing else "")
