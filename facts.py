"""Figures for a dated series, worked out by code so that each one can be checked. No GUI, no network, no model.

A series is a list of Point sorted by date. Functions return None when a figure has no honest answer,
such as a comparison with a day that is not in the data, and the caller decides how to report that.
"""

import datetime as dt
import hashlib
import re
from dataclasses import dataclass
from statistics import mean, median

from numfmt import format_value

NUMBER = re.compile(r"\d[\d,]*(?:\.\d+)?")
HANGUL = re.compile(r"[가-힣]")
LATIN = re.compile(r"[A-Za-z]")
PERCENT = "+0.0%;-0.0%;0.0%"
# When a reading rule speaks. Set from a year of real daily data so that a rule fires rarely enough to be worth reading.
UNUSUAL_SPREAD = 3.0  # usual spreads from the same weekday's median
UNUSUAL_CHANGE = 0.07  # and at least this far from it, so a very steady series does not flag small moves
USUAL_CHANGE = 0.05  # within this of the median, a day may be called usual. Between the two, the rules say nothing
ODD_SPREAD = 3.0  # the same two tests for a day in the recent past
ODD_CHANGE = 0.15
ODD_MOVES_MEAN = 0.02  # a day is only worth naming if leaving it out moves the comparison this much
WEEKEND_GAP = 0.10  # weekends must differ from weekdays this much for the weekend to explain anything
CONFLICT_CHANGE = 0.03  # both comparisons must really move
WEEKS_IN_A_ROW = 3
LARGE_CHANGE = 0.15  # a move this big is never called nothing to note, however widely the series usually swings
YEAR_RUN = 3  # months or quarters in a row above or below the year before
FAR_FROM_PEAK = 0.7  # below this share of the same month's best, the long view is worth a sentence
MIN_TO_READ = 3
OTHER_ITEM = re.compile(r"기타|그\s*외|others?|misc\.?|etc\.?", re.I)
AMOUNT = 0  # the scale of a series whose changes are plain amounts rather than points
PART_MARK = re.compile(r"[-–·•]\s")
NESTED_TOLERANCE = 0.001  # a parent's figure is the exact sum of its parts, give or take rounding
MACHINE_NAME = re.compile(r"[A-Za-z]+(?:_[A-Za-z0-9-]+)+|[a-z]+[A-Z][A-Za-z0-9]*")  # tourStatConfirmManageBiz, RETAINED_CATCH_KG: exported names, in no one's language
SNAKE_NAME = re.compile(r"(?<![A-Za-z0-9_])[A-Za-z][A-Za-z0-9]*(?:_[A-Za-z0-9]+)+(?![A-Za-z0-9_])")
NAME_UNITS = {"kg": "kg", "g": "g", "mj": "MJ", "kj": "kJ", "kcal": "kcal", "km": "km", "m2": "m²", "m3": "m³", "kwh": "kWh", "mwh": "MWh", "usd": "USD", "aud": "AUD", "krw": "KRW", "pct": "%", "percent": "%"}
LEGEND_MARK = re.compile(r"(?<=[^\s(])\([A-Z]\)(?=\s)")  # 제주(A) 2025년 points to a formula in a table's footnote. 시나리오(A) ends with it, and it may be all that tells it from 시나리오(B)
COUNT_MARK = re.compile(r"^\s*#\s+")
KOREAN_UNIT_MARK = re.compile(r"(?<=[가-힣])_(원|천원|만원|백만원|억원|명|건|톤|개|회)$")  # 수수료_원
PRICE_NAME = re.compile(r"^(?!.*(?:총|합계|누계|\b(?:total|sum)\b))(?=.*(?:수수료|요금|가격|단가|이용료|사용료|입장료|참가비|회비|\b(?:price|fee|tariff)s?\b)\s*[_( ]?\s*(?:원|천원|만원|달러|krw|usd|aud|\$)?\s*\)?\s*$)", re.I)
MIN_NESTED_PARENT = 100  # among hospitals with 1, 2 or 3 doctors, 2 = 1 + 1 says nothing: it took 38 doctors out of a list of 127
MAX_NESTED_ITEMS = 100  # among thousands of companies, one is bound to equal the sum of the next few by chance, and they are not its parts
CLOSE_SECOND = 1.1  # the first is less than this many times the second: a multiple like 1.0 says nothing
OUTLINE = (("roman", re.compile(r"[ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ]+\s*[.)]|[IVX]+\.\s")), ("dash", re.compile(r"\d+-\d+\.")), ("number", re.compile(r"\d+\.(?!\d)")),
           ("paren", re.compile(r"\(\d+\)")), ("hangul", re.compile(r"[가-하]\.")))
SMALL_ITEM = 0.01  # an item under this share of the largest is not named for moving most
MAX_LISTED = 6
SEPARATE_LISTS = {"ko": "번호가 겹치는 항목이 함께 있어(자산과 부채 등) 크기를 비교하지 않았습니다.", "en": "Numbers repeat among these lines, as with assets and liabilities, so their sizes were not compared."}
SEPARATE_TOTALS = {"ko": "합계 줄이 여러 개 있어(수입과 지출 등) 줄끼리 크기를 비교하지 않았습니다.", "en": "There are several total lines, as with income and expenses, so the lines were not ranked against one another."}
RATE_SYLLABLE = r"(?<![법규자])[율률](?![가-힣])"  # 달성률 and 환율, not 법률상담건수 or 자율주행
RATE_WORD = re.compile(RATE_SYLLABLE + r"|%|퍼센트|지수|평균|단가|점수|만족도|온도|증감|순위|지표|배율|\b(?:rates?|ratios?|percent(?:age)?|index|average|mean|scores?|rank)\b", re.I)
WEEKDAYS = {"ko": ["월", "화", "수", "목", "금", "토", "일"], "en": ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]}
WEEKDAY_PLURALS = ["Mondays", "Tuesdays", "Wednesdays", "Thursdays", "Fridays", "Saturdays", "Sundays"]
MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
WORDS = {
    "ko": {"up": "증가", "down": "감소", "flat": "변동 없음", "prev_day": "전일보다", "prev": "직전({date})보다",
           "week": "전주 같은 요일보다", "avg": "직전 {n}개 평균보다", "year": "전년 같은 {unit}보다", "per_month": "월평균 기준",
           "달": "달", "분기": "분기"},
    "en": {"up": "up", "down": "down", "flat": "unchanged", "prev_day": "from the previous day", "prev": "from {date}",
           "week": "from the same day last week", "avg": "from the average of the previous {n}", "year": "from the same {unit} last year", "per_month": "per month",
           "달": "month", "분기": "quarter"},
}


@dataclass(frozen=True)
class Point:
    date: dt.datetime
    value: float
    row: int = 0
    sheet: int = 0
    label: str = ""  # how the period is written when a date cannot say it: 2025-26 is a financial year, not the year 2026


@dataclass(frozen=True)
class Fact:
    """One calculated statement. numbers holds the only figures a sentence citing this fact may use."""
    id: str
    text: str
    numbers: tuple
    figure: str = ""  # on a subject fact: which figure all of a file's facts are about
    period: str = ""  # and which period is "now"; any other period is only something to compare with


def readable_name(name):
    """A heading as a report should say it: Retained catch (kg) for RETAINED_CATCH_KG, 제주 for 제주(A), EAF Items for # EAF Items. Only the marks of a machine or of a table's own footnotes go; the words stay the sheet's."""
    text = LEGEND_MARK.sub("", COUNT_MARK.sub("", name or ""))
    return KOREAN_UNIT_MARK.sub(r"(\1)", SNAKE_NAME.sub(_words_of_snake, text))


def _words_of_snake(match):
    words = match.group(0).split("_")
    if not all(word.isalpha() and len(word) >= 2 for word in words):  # E_0000_WH is a code, and words made from it would be no clearer
        return match.group(0)
    unit = NAME_UNITS.get(words[-1].casefold()) if len(words) > 1 else None
    if unit:
        words = words[:-1]
    text = " ".join(words).lower()
    text = text[:1].upper() + text[1:]
    return f"{text} ({unit})" if unit else text


def detect_language(texts):
    """The router: which language the report is mostly in, by counting letters. Needs no model and no memory."""
    joined = MACHINE_NAME.sub(" ", " ".join(texts))
    return "ko" if len(HANGUL.findall(joined)) >= len(LATIN.findall(joined)) else "en"


def fingerprint(pack, lang):
    """Changes whenever the calculated results change, so a sentence chosen for old data is not used on new data."""
    joined = "\n".join(f"{fact.id}: {fact.text}" for fact in pack)
    return hashlib.sha256(f"{lang}\n{joined}".encode("utf-8")).hexdigest()


def ratio(now, base):
    """Relative change, or None when the base is zero: a percentage of nothing has no meaning."""
    return None if not base else (now - base) / abs(base)


def percent(value):
    return "0.0%" if format_value(abs(value), "0.0%")[0] == "0.0%" else format_value(value, PERCENT)[0]


def points_text(difference, lang="ko", scale=1):
    """A change in a rate or a score, as a difference. A percentage of a percentage reads wrong."""
    if scale == AMOUNT:  # a column of ups and downs in a count: its change is an amount, in no unit of its own
        return ("+" if difference > 0 else "") + plain_number(difference)
    size = round(difference * scale, 1) + 0.0
    return f"{size:+.1f}포인트" if lang == "ko" else f"{size:+.1f} points"


def plain_number(value):
    """A figure for a sentence when its cell has no format: 4,930,916 rather than 4930915.845, and 115.9 rather than 115.884985."""
    size = abs(value)
    pattern = "#,##0" if size >= 1000 or float(value).is_integer() else "0.0" if size >= 10 else "0.00" if size >= 1 else "0.000"
    text = format_value(value, pattern)[0]
    return text.rstrip("0").rstrip(".") if "." in text else text


def previous(points):
    return points[-2] if len(points) >= 2 else None


def week_ago(points):
    """The point exactly seven days before the newest. A nearby day would compare different weekdays."""
    wanted = points[-1].date.date() - dt.timedelta(days=7)
    return next((p for p in reversed(points) if p.date.date() == wanted), None)


def mean_before(points, count):
    """Mean of the count points before the newest, or None when there are fewer. The newest is left out of its own base."""
    earlier = points[-count - 1:-1]
    return mean(p.value for p in earlier) if count >= 1 and len(earlier) == count else None


def extreme(points, count, highest=True):
    """The highest or lowest of the newest count points. On a tie, the most recent."""
    recent = points[-count:]
    pick = max if highest else min
    target = pick(p.value for p in recent)
    return next(p for p in reversed(recent) if p.value == target)


def streak(points):
    """How many steps in a row the series has moved the same way: +3 is three rises, -2 is two falls, 0 is no run."""
    count, direction = 0, 0
    for later, earlier in zip(reversed(points), reversed(points[:-1])):
        step = (later.value > earlier.value) - (later.value < earlier.value)
        if not step or (direction and step != direction):
            break
        direction, count = step, count + 1
    return direction * count


def same_weekday(points, weeks=8):
    """(median, usual spread, count) of earlier points on the newest point's weekday, or None with fewer than three.

    A median and a median-based spread, so that one holiday among the weeks does not move what counts as usual.
    The spread has a floor of 1% of the level, so a perfectly steady series can still show a real change.
    """
    day = points[-1].date.weekday()
    earlier = [p.value for p in points[:-1] if p.date.weekday() == day][-weeks:]
    if len(earlier) < 3:
        return None
    level = median(earlier)
    spread = max(1.4826 * median(abs(value - level) for value in earlier), abs(level) * 0.01)
    return (level, spread, len(earlier)) if spread else None  # every Sunday 0, closed: there is no usual range to stand outside


def rank(points, count=30):
    """(position, size): where the newest value stands among the newest count points. 1 is the highest."""
    recent = points[-count:]
    return sum(1 for p in recent if p.value > points[-1].value) + 1, len(recent)


def week_total_change(points):
    """Change of the newest 7 points' total against the 7 before them, or None with fewer than 14."""
    if len(points) < 14:
        return None
    return ratio(sum(p.value for p in points[-7:]), sum(p.value for p in points[-14:-7]))


def _usual_gap(points):
    recent = points[-31:]
    gaps = sorted((later.date.date() - earlier.date.date()).days for earlier, later in zip(recent, recent[1:]))
    return gaps[len(gaps) // 2] if gaps else None


def is_daily(points):
    """True when points are normally one day apart. Weekly and weekday figures mean nothing otherwise."""
    return _usual_gap(points) == 1


def is_monthly(points):
    gap = _usual_gap(points)
    return gap is not None and 28 <= gap <= 31


def is_quarterly(points):
    gap = _usual_gap(points)
    return gap is not None and 89 <= gap <= 92


def is_yearly(points):
    return _usual_gap(points) in (365, 366) or partial_year(points) is not None


def partial_year(points):
    """How many months the newest point covers, when it is written as part of a year among whole years, such as 2022.10. after 2021. Else None."""
    earlier = [p.date for p in points[-31:-1]]
    if len(earlier) < 2:
        return None
    place = (earlier[0].month, earlier[0].day)
    whole_years = place in ((1, 1), (12, 31)) and all((d.month, d.day) == place for d in earlier)
    in_a_row = all(later.year - before.year == 1 for before, later in zip(earlier, earlier[1:]))
    newest = points[-1].date
    if whole_years and in_a_row and newest.year == earlier[-1].year + 1 and (newest.month, newest.day) != place and newest.month > 1:
        return None if newest.month == 12 else newest.month
    return None


def point_text(date, points, lang):
    """A date as the data's own rhythm calls for: a monthly series is about months, not the first of each."""
    ko = lang == "ko"
    if points[-1].label:
        named = next((point.label for point in points if point.date == date), "")
        if named:
            return named
    if partial_year(points) and date == points[-1].date:
        return f"{date.year}년 {date.month}월까지" if ko else f"{MONTHS[date.month - 1]} {date.year} year to date"
    if is_yearly(points):
        return f"{date.year}년" if ko else str(date.year)
    if is_quarterly(points):
        quarter = (date.month - 1) // 3 + 1
        return f"{date.year}년 {quarter}분기" if ko else f"Q{quarter} {date.year}"
    if is_monthly(points):
        return f"{date.year}년 {date.month}월" if ko else f"{MONTHS[date.month - 1]} {date.year}"
    return day_text(date, lang)


def year_ago(points):
    """The point in the same month a year before the newest, for monthly or quarterly data."""
    newest = points[-1].date
    return next((p for p in reversed(points) if (p.date.year, p.date.month) == (newest.year - 1, newest.month)), None)


def years_in_a_row(points):
    """How many points in a row, newest back, have been above (+) or below (-) the same period a year before."""
    level = {(p.date.year, p.date.month): p.value for p in points}
    count, direction = 0, 0
    for point in reversed(points):
        base = level.get((point.date.year - 1, point.date.month))
        step = 0 if base is None else (point.value > base) - (point.value < base)
        if not step or (direction and step != direction):
            break
        direction, count = step, count + 1
    return direction * count


def year_to_date(points):
    """True for a running total that starts again each January: it only rises within a year, then each January falls to under half of December."""
    if not is_monthly(points):
        return False
    restarts = 0
    for earlier, later in zip(points, points[1:]):
        if later.date.year == earlier.date.year:
            if later.value < earlier.value:
                return False
        elif later.value * 2 > earlier.value:
            return False
        else:
            restarts += 1
    return restarts >= 1 and len(points) >= 13


def period_end(points):
    """The last day the newest point covers, by the series' rhythm."""
    date = points[-1].date.date()
    if is_yearly(points) and not partial_year(points):
        return dt.date(date.year, 12, 31)
    if is_quarterly(points):
        last = ((date.month - 1) // 3 + 1) * 3
    elif is_monthly(points) or partial_year(points):
        last = date.month
    else:
        return date
    return dt.date(date.year + last // 12, last % 12 + 1, 1) - dt.timedelta(days=1)


def unfinished(points, as_of):
    """True when the newest point's period had not ended on as_of. A daily value entered that same day only counts if it is also far below usual."""
    end = period_end(points)
    if end > as_of:
        return True
    if end == as_of and is_daily(points):
        usual = same_weekday(points)
        change = ratio(points[-1].value, usual[0]) if usual else None
        return change is not None and change <= -LARGE_CHANGE
    return False


def same_kind_before(points, count=7):
    """Mean of the previous count points that share the newest point's kind of day: weekday or weekend. (mean, how many)."""
    weekend = points[-1].date.weekday() >= 5
    values = [p.value for p in points[-count - 1:-1] if (p.date.weekday() >= 5) == weekend]
    return (mean(values), len(values)) if len(values) >= 2 else None


def stands_out(points):
    """(usual level, change) when the newest point is far from its weekday's usual level by both size and spread, else None."""
    usual = same_weekday(points)
    change = ratio(points[-1].value, usual[0]) if usual else None
    if change is None or abs(change) < ODD_CHANGE or abs(points[-1].value - usual[0]) / usual[1] < ODD_SPREAD:
        return None
    return usual[0], change


def odd_day(points, count=7):
    """Among the newest count points, the one furthest from its own weekday's usual level. (point, usual, change) or None."""
    best = None
    for i in range(max(len(points) - count, 0), len(points)):
        found = stands_out(points[:i + 1])
        if found and (best is None or abs(found[1]) > abs(best[2])):
            best = (points[i], found[0], found[1])
    return best


def day_text(date, lang):
    weekday = WEEKDAYS[lang][date.weekday()]
    if lang == "ko":
        return f"{date.month}월 {date.day}일({weekday})"
    return f"{MONTHS[date.month - 1]} {date.day} ({weekday})"


def _movement(change, lead, lang):
    words = WORDS[lang]
    size = format_value(abs(change), "0.0%")[0]
    if lang == "ko":
        if round(abs(change), 3) == 0:
            return f"{lead} {words['flat']}"
        return f"{lead} {size} {words['up'] if change > 0 else words['down']}"
    if round(abs(change), 3) == 0:
        return f"{words['flat']} {lead}"
    return f"{words['up'] if change > 0 else words['down']} {size} {lead}"


def weekend_gap(points, count=28):
    """How far the weekend level sits from the weekday level over the newest count points, or None."""
    recent = points[-count:]
    weekdays = [p.value for p in recent if p.date.weekday() < 5]
    weekends = [p.value for p in recent if p.date.weekday() >= 5]
    return ratio(mean(weekends), mean(weekdays)) if weekdays and len(weekends) >= 2 else None


def weeks_in_a_row(points):
    """How many weeks in a row the 7-point total has moved the same way, newest back: +3, -2, or 0."""
    totals = [sum(p.value for p in points[end - 7:end]) for end in range(len(points), 6, -7)]
    count, direction = 0, 0
    for later, earlier in zip(totals, totals[1:]):
        step = (later > earlier) - (later < earlier)
        if not step or (direction and step != direction):
            break
        direction, count = step, count + 1
    return direction * count


def _size(change):
    return format_value(abs(change), "0.0%")[0]


def readings(points, show, lang="ko", limit=3):
    """How to read the newest figure, decided by fixed rules and worded by code. The same data always reads the same.

    Each rule states when a comparison misleads or a value stands out. Rules about the newest value come first.
    Nothing here claims a cause: a rule may say that a day should be checked, never why it differed.
    """
    ko = lang == "ko"
    found = []
    if len(points) < MIN_TO_READ:
        return found
    newest = points[-1]
    day = lambda p: point_text(p.date, points, lang)  # noqa: E731
    low_high = lambda change: ("낮" if change < 0 else "높") if ko else ("below" if change < 0 else "above")  # noqa: E731
    before = previous(points)
    prev_change = ratio(newest.value, before.value)
    # A big fall in the newest value is the usual sign of a period still being counted, so say to check it. This names no cause.
    check_end = lambda change: (" 마지막 기간의 집계가 끝났는지 확인하세요." if ko else " Check that the last period is fully counted.") if change <= -LARGE_CHANGE else ""  # noqa: E731

    months = partial_year(points)
    if months:
        per_month = ratio(newest.value / months, before.value / 12)
        if prev_change is None or per_month is None:  # a year of 0 before it: nothing to take a percentage of
            found.append(("read_partial_year", f"{day(newest)} 값은 {months}개월치로 보입니다. 직전 {day(before)} 값이 {show(before.value)}이라 변화율을 낼 수 없습니다." if ko else
                          f"{day(newest)} looks like {months} months. {day(before)} was {show(before.value)}, so no percentage can be given."))
            return [Fact(name, text, tuple(NUMBER.findall(text))) for name, text in found]
        found.append(("read_partial_year", (
            f"{day(newest)} 값은 {months}개월치로 보입니다. 12개월치인 {day(before)} {show(before.value)} 대비 {percent(prev_change)}는 기간이 달라 그대로 읽을 수 없고, 월평균으로 맞추면 {percent(per_month)}입니다." if ko else
            f"{day(newest)} looks like {months} months, so {percent(prev_change)} against the 12 months of {day(before)}, {show(before.value)}, cannot be read as it stands. Per month it is {percent(per_month)}.")))
        return [Fact(name, text, tuple(NUMBER.findall(text))) for name, text in found]

    if year_to_date(points):
        base = year_ago(points)
        change = ratio(newest.value, base.value) if base else None
        same = f" 전년 같은 달 {day(base)} {show(base.value)} 대비 {percent(change)}입니다." if ko and change is not None else f" {percent(change)} against the same month last year, {day(base)}, {show(base.value)}." if change is not None else ""
        found.append(("read_year_to_date", f"이 값은 매년 1월에 다시 시작하는 누계로 보여, 직전 달 대비로는 읽을 수 없습니다.{same}" if ko else
                      f"This looks like a running total that restarts each January, so it cannot be read against the month before.{same}"))
        return [Fact(name, text, tuple(NUMBER.findall(text))) for name, text in found]

    if is_yearly(points):
        if prev_change is not None:
            found.append(("read_prev_year", f"직전 해 {day(before)} {show(before.value)} 대비 {percent(prev_change)}입니다." if ko else
                          f"{percent(prev_change)} against the year before, {day(before)}, {show(before.value)}."))
        run = streak(points)
        if abs(run) >= 2:
            found.append(("read_year_run", f"{abs(run)}년 연속 {'증가' if run > 0 else '감소'}했습니다." if ko else f"It has {'risen' if run > 0 else 'fallen'} {abs(run)} years in a row."))

    if is_monthly(points) or is_quarterly(points):
        unit = "달" if is_monthly(points) else "분기"
        word = WORDS[lang][unit]
        base = year_ago(points)
        year_change = ratio(newest.value, base.value) if base else None
        recent = points[-13:-1]
        # Only a fall from the period before can look like one still being counted. A low year on a rising month cannot.
        year_check = check_end(year_change) if year_change is not None and prev_change is not None and prev_change <= -LARGE_CHANGE else ""
        if year_change is None:
            position, size = rank(points)
            if position == size:
                where = f" {size}개 중 가장 낮습니다." if ko else f" It is the lowest of {size}."
            elif position == 1:
                where = f" {size}개 중 가장 높습니다." if ko else f" It is the highest of {size}."
            else:
                where = f" {size}개 중 {position}번째로 높습니다." if ko else f" It is number {position} highest of {size}."
            if prev_change is not None:
                found.append(("read_no_year", f"전년 같은 {word} 값이 없어 계절 영향을 가릴 수 없습니다. 직전 {word} 대비 {percent(prev_change)}입니다.{where}{check_end(prev_change)}" if ko else
                              f"There is no value for the same {word} last year, so the season cannot be separated out. {percent(prev_change)} against the {word} before.{where}{check_end(prev_change)}"))
        elif prev_change is not None and prev_change * year_change < 0 and min(abs(prev_change), abs(year_change)) >= CONFLICT_CHANGE:
            found.append(("read_year_conflict", f"직전 {word} 대비 {percent(prev_change)}와 전년 같은 {word} {day(base)} {show(base.value)} 대비 {percent(year_change)}의 방향이 다릅니다. 계절 영향이 있을 수 있으니 전년 같은 {word} 기준으로 읽으세요.{year_check}" if ko else
                          f"{percent(prev_change)} against the {word} before and {percent(year_change)} against the same {word} last year, {day(base)}, {show(base.value)}, point different ways. The season may play a part, so read it by the same {word} last year.{year_check}"))
        else:
            found.append(("read_year", f"전년 같은 {word} {day(base)} {show(base.value)} 대비 {percent(year_change)}입니다.{year_check}" if ko else
                          f"{percent(year_change)} against the same {word} last year, {day(base)}, {show(base.value)}.{year_check}"))
        if prev_change is not None and percent(prev_change) != "0.0%" and len(recent) >= 6:
            lowest, highest = before.value == min(p.value for p in recent), before.value == max(p.value for p in recent)
            if lowest or highest:
                # A low base pushes the change up and a high base pulls it down, whichever way the value moved.
                found.append(("read_prev_extreme", f"직전 {word}({day(before)}) 값이 최근 {len(recent)}개 중 가장 {'낮아' if lowest else '높아'}, 직전 {word} 대비 {percent(prev_change)}는 그만큼 {'높게' if lowest else '낮게'} 나옵니다." if ko else
                              f"The {word} before, {day(before)}, was the {'lowest' if lowest else 'highest'} of the latest {len(recent)}, which pulls {percent(prev_change)} against it {'up' if lowest else 'down'}."))
        run = years_in_a_row(points)
        if abs(run) >= YEAR_RUN:
            found.append(("read_year_run", f"전년 같은 {word}보다 {abs(run)}번 연속 {'높습' if run > 0 else '낮습'}니다." if ko else
                          f"{abs(run)} {word}s in a row have been {'above' if run > 0 else 'below'} the year before."))
        same = [p for p in points[:-1] if p.date.month == newest.date.month]
        if len(same) >= 3:
            top = max(same, key=lambda p: p.value)
            if newest.value > top.value or newest.value < min(p.value for p in same):
                found.append(("read_same_month", f"같은 {word} 기준으로 {len(same) + 1}개 해 중 가장 {'높' if newest.value > top.value else '낮'}습니다." if ko else
                              f"For this {word}, it is the {'highest' if newest.value > top.value else 'lowest'} of {len(same) + 1} years."))
            elif top.value > 0 and newest.value < top.value * FAR_FROM_PEAK:
                found.append(("read_same_month", f"같은 {word} 기준 최고는 {day(top)} {show(top.value)}이고, 이번 값은 그 {format_value(newest.value / top.value, '0.0%')[0]} 수준입니다." if ko else
                              f"The best for this {word} was {day(top)}, {show(top.value)}. This value is {format_value(newest.value / top.value, '0.0%')[0]} of it."))

    if is_daily(points):
        usual = same_weekday(points)
        week = week_ago(points)
        week_change = ratio(newest.value, week.value) if week else None
        base = mean_before(points, 7)
        avg_change = ratio(newest.value, base) if base is not None else None
        gap = weekend_gap(points)
        name = WEEKDAYS[lang][newest.date.weekday()]
        plural = WEEKDAY_PLURALS[newest.date.weekday()]

        if usual:
            change = ratio(newest.value, usual[0])
            spread = (newest.value - usual[0]) / usual[1]
            if change is not None and abs(spread) >= UNUSUAL_SPREAD and abs(change) >= UNUSUAL_CHANGE:
                found.append(("read_unusual", f"이번 값은 직전 {usual[2]}번 {name}요일의 중앙값 {show(usual[0])}보다 {_size(change)} {low_high(change)}아, 평소 범위를 벗어난 수준입니다.{check_end(change)}" if ko else
                              f"This value is {_size(change)} {low_high(change)} the median of the previous {usual[2]} {plural}, {show(usual[0])}, and outside their usual range.{check_end(change)}"))
            elif change is not None and abs(change) >= LARGE_CHANGE:
                # On a series that swings widely a large move is common, so only the lowest of its weekday is worth a second look.
                compared = [p.value for p in points[:-1] if p.date.weekday() == newest.date.weekday()][-usual[2]:]
                tail = check_end(change) if newest.value < min(compared) else ""
                found.append(("read_large", f"직전 {usual[2]}번 {name}요일의 중앙값 {show(usual[0])} 대비 {percent(change)}입니다. 변동이 큰 요일이라 평소 범위 안이지만 폭은 큽니다.{tail}" if ko else
                              f"{percent(change)} against the median of the previous {usual[2]} {plural}, {show(usual[0])}. This weekday swings widely, so it is within the usual range, but the move is large.{tail}"))

        # Last week's same day is only a fair base if that day was itself ordinary.
        week_was_odd = stands_out(points[:points.index(week) + 1]) if week else None
        trusted = week_change
        if week_was_odd and usual:
            trusted = ratio(newest.value, usual[0])
            if trusted is not None and week_change is not None:  # a holiday's 0 last week gives no change to set this against
                found.append(("read_week_odd", f"전주 같은 요일 {day(week)} {show(week.value)}가 그 요일 평소보다 {_size(week_was_odd[1])} {low_high(week_was_odd[1])}았으니, 전주 대비 {percent(week_change)} 대신 평소 {name}요일 중앙값 대비 {percent(trusted)}로 읽으세요." if ko else
                              f"Last week's {day(week)}, {show(week.value)}, was itself {_size(week_was_odd[1])} {low_high(week_was_odd[1])} usual for that weekday. Read {percent(trusted)} against the usual {name} rather than {percent(week_change)} against last week."))
        by_same_day = "같은 요일 기준" if not week_was_odd else "평소 같은 요일 기준"
        by_same_day_en = "the same weekday" if not week_was_odd else "the usual level of this weekday"

        kind = same_kind_before(points)
        mixed = sum(1 for p in points[-8:-1] if p.date.weekday() >= 5)
        # When last week's day was odd, the rule above has already said which comparison to read.
        conflict = (not week_was_odd and trusted is not None and avg_change is not None and kind and trusted * avg_change < 0 and 0 < mixed < 7
                    and gap is not None and abs(gap) >= WEEKEND_GAP and min(abs(trusted), abs(avg_change)) >= CONFLICT_CHANGE)
        if conflict:
            reads = kind[0] > base  # the same kind of day sits above the mixed mean, so the mixed comparison reads high
            found.append(("read_conflict", f"직전 7개 평균 대비 {percent(avg_change)}는 주말 {mixed}개가 섞인 기준이라 {'높게' if reads else '낮게'} 나올 수 있으니, {by_same_day} {percent(trusted)}로 읽으세요." if ko else
                          f"{percent(avg_change)} against the mean of the previous 7 can read {'high' if reads else 'low'}, because that mean includes {mixed} weekend values. Read it by {by_same_day_en}: {percent(trusted)}."))

        crossed = (before.date.weekday() >= 5) != (newest.date.weekday() >= 5)
        if (crossed and not conflict and prev_change is not None and trusted is not None and gap is not None and abs(gap) >= WEEKEND_GAP
                and (newest.date.date() - before.date.date()).days == 1):
            found.append(("read_prev_weekend", f"직전 값 대비 {percent(prev_change)}는 평일과 주말을 비교한 수치라 요일 효과가 큽니다. {by_same_day} {percent(trusted)}를 보세요." if ko else
                          f"{percent(prev_change)} against the previous value compares a weekday with a weekend day. See {by_same_day_en}: {percent(trusted)}."))

        odd = odd_day(points[:-1])
        if odd and base is not None:
            others = [p.value for p in points[-8:-1] if p is not odd[0]]
            without = ratio(newest.value, mean(others)) if others and odd[0] in points[-8:-1] else None
            if without is not None and avg_change is not None and abs(without - avg_change) >= ODD_MOVES_MEAN:
                moved = ("끌어내렸을" if odd[2] < 0 else "끌어올렸을") if ko else ("pulled down" if odd[2] < 0 else "pushed up")
                found.append(("read_odd_day", f"{day(odd[0])} {show(odd[0].value)}는 그 요일 평소 {show(odd[1])}보다 {_size(odd[2])} {low_high(odd[2])}아 직전 평균을 {moved} 수 있습니다. 그날을 빼면 직전 평균 대비는 {percent(avg_change)}에서 {percent(without)}가 됩니다. 그날 사정을 확인하세요." if ko else
                              f"{day(odd[0])}, {show(odd[0].value)}, was {_size(odd[2])} {low_high(odd[2])} its usual {show(odd[1])} and may have {moved} the recent mean. Without it, {percent(avg_change)} against that mean becomes {percent(without)}. Check what happened that day."))

        run = weeks_in_a_row(points)
        if abs(run) >= WEEKS_IN_A_ROW:
            found.append(("read_weeks", f"최근 7개 합계가 {abs(run)}주 연속 {'증가' if run > 0 else '감소'}했습니다." if ko else
                          f"The 7-value total has {'risen' if run > 0 else 'fallen'} {abs(run)} weeks in a row."))

        if not found and usual:
            change = ratio(newest.value, usual[0])
            if change is not None and abs(change) < USUAL_CHANGE:
                found.append(("read_usual", f"같은 요일 평소 수준입니다. 직전 {usual[2]}번 {name}요일의 중앙값 {show(usual[0])} 대비 {percent(change)}입니다." if ko else
                              f"This is the usual level for this weekday: {percent(change)} against the median of the previous {usual[2]} {plural}, {show(usual[0])}."))
            elif change is not None:
                found.append(("read_nothing", f"크게 짚을 점은 없습니다. 직전 {usual[2]}번 {name}요일의 중앙값 {show(usual[0])} 대비 {percent(change)}입니다." if ko else
                              f"Nothing stands out: {percent(change)} against the median of the previous {usual[2]} {plural}, {show(usual[0])}."))

    position, size = rank(points)
    if size >= (5 if is_yearly(points) else 14) and (position == 1 or sum(1 for p in points[-30:] if p.value < newest.value) == 0):
        top = position == 1
        found.append(("read_extreme", f"최근 {size}개 중 가장 {'높은' if top else '낮은'} 값입니다." if ko else f"This is the {'highest' if top else 'lowest'} of the latest {size} values."))
    if not found and prev_change is not None:
        if abs(prev_change) >= LARGE_CHANGE:
            found.append(("read_large", f"직전 값 대비 {percent(prev_change)}로 폭이 큽니다.{check_end(prev_change)}" if ko else f"{percent(prev_change)} against the previous value is a large move.{check_end(prev_change)}"))
        elif len(points) >= 14:  # enough data and no finding is an answer, not a gap
            found.append(("read_nothing", f"크게 짚을 점은 없습니다. 직전 값 대비 {percent(prev_change)}입니다." if ko else f"Nothing stands out: {percent(prev_change)} against the previous value."))
        else:
            found.append(("read_short", f"값이 {len(points)}개뿐이라 평소 수준을 판단할 수 없습니다. 직전 값 대비 {percent(prev_change)}입니다." if ko else
                          f"With only {len(points)} values, the usual level cannot be judged. {percent(prev_change)} against the previous value."))
    return [Fact(name, text, tuple(NUMBER.findall(text))) for name, text in found[:limit]]


def fact_pack(points, header, show, lang="ko"):
    """Everything calculated about the newest point, as numbered statements a later step can cite. Nothing is estimated."""
    ko = lang == "ko"
    newest = points[-1]
    daily_data = is_daily(points)
    day = lambda p: point_text(p.date, points, lang)  # noqa: E731
    pack = [("value", f"{day(newest)} {header}: {show(newest.value)}")]
    months = partial_year(points)
    if months:  # Rank, streak and means against whole years would all mislead, so only the fair comparison goes in.
        change = ratio(newest.value / months, points[-2].value / 12)
        if change is not None:
            pack.append(("per_month", f"월평균 {show(newest.value / months)}, {day(points[-2])} 월평균 {show(points[-2].value / 12)} 대비 {percent(change)}" if ko else
                         f"per month {show(newest.value / months)}, {percent(change)} against {show(points[-2].value / 12)} per month in {day(points[-2])}"))
        return readings(points, show, lang, limit=9) + [Fact(name, text, tuple(NUMBER.findall(text))) for name, text in pack]
    if year_to_date(points):  # Steps, means and rank mean nothing on a running total.
        return readings(points, show, lang, limit=9) + [Fact(name, text, tuple(NUMBER.findall(text))) for name, text in pack]
    found = week_ago(points)
    week_change = ratio(newest.value, found.value) if found else None
    if week_change is not None:
        pack.append(("week", f"전주 같은 요일 {day(found)} {show(found.value)} 대비 {percent(week_change)}" if ko else
                     f"{percent(week_change)} against the same weekday last week, {day(found)}, {show(found.value)}"))
    if is_monthly(points) or is_quarterly(points):
        found = year_ago(points)
        change = ratio(newest.value, found.value) if found else None
        word = WORDS[lang]["달" if is_monthly(points) else "분기"]
        if change is not None:
            pack.append(("year", f"전년 같은 {word} {day(found)} {show(found.value)} 대비 {percent(change)}" if ko else
                         f"{percent(change)} against the same {word} last year, {day(found)}, {show(found.value)}"))
    found = previous(points)
    change = ratio(newest.value, found.value) if found else None
    if change is not None:
        pack.append(("prev", f"직전 값 {day(found)} {show(found.value)} 대비 {percent(change)}" if ko else
                     f"{percent(change)} against the previous value, {day(found)}, {show(found.value)}"))
    base = mean_before(points, 7)
    avg_change = ratio(newest.value, base) if base is not None else None
    if avg_change is not None:
        pack.append(("avg7", f"직전 7개 평균 {show(base)} 대비 {percent(avg_change)}" if ko else
                     f"{percent(avg_change)} against the mean of the previous 7, {show(base)}"))
    usual = same_weekday(points) if daily_data else None
    spread = None
    if usual:
        change = ratio(newest.value, usual[0])
        spread = (newest.value - usual[0]) / usual[1]
        name = WEEKDAYS[lang][newest.date.weekday()]
        plural = WEEKDAY_PLURALS[newest.date.weekday()]
        pack.append(("weekday", f"직전 {usual[2]}번 {name}요일의 중앙값 {show(usual[0])} 대비 {percent(change) if change is not None else '비교 불가'}, 평소 변동폭의 {spread:+.1f}배" if ko else
                     f"{percent(change) if change is not None else 'not comparable'} against the median of the previous {usual[2]} {plural}, {show(usual[0])}, {spread:+.1f} times the usual spread"))
    if daily_data:
        kind = same_kind_before(points)
        if kind:
            change = ratio(newest.value, kind[0])
            if change is not None:
                label = ("주말" if newest.date.weekday() >= 5 else "평일") if ko else ("weekend" if newest.date.weekday() >= 5 else "weekday")
                pack.append(("signal_same_kind", f"직전 7개 중 {label} {kind[1]}개만의 평균 {show(kind[0])} 대비 {percent(change)}" if ko else
                             f"{percent(change)} against the mean of only the {kind[1]} {label} values among the previous 7, {show(kind[0])}"))
        if week_change is not None and avg_change is not None and week_change * avg_change < 0:
            mixed = sum(1 for p in points[-8:-1] if p.date.weekday() >= 5)
            pack.append(("signal_conflict", f"비교 기준의 방향이 엇갈림: 전주 같은 요일 기준 {percent(week_change)}, 직전 7개 평균 기준 {percent(avg_change)}. 직전 7개에는 주말 {mixed}개가 섞여 있음" if ko else
                         f"the comparisons point different ways: {percent(week_change)} by the same weekday last week, {percent(avg_change)} by the mean of the previous 7, which includes {mixed} weekend values"))
        if spread is not None and abs(spread) >= UNUSUAL_SPREAD and change is not None and abs(change) >= UNUSUAL_CHANGE:
            pack.append(("signal_unusual", f"같은 요일 기준으로 평소 범위를 벗어남: 평소 변동폭의 {spread:+.1f}배" if ko else
                         f"outside the usual range for this weekday: {spread:+.1f} times the usual spread"))
        odd = odd_day(points)
        if odd:
            pack.append(("signal_odd_day", f"최근 7개 중 평소 같은 요일과 가장 다른 날: {day(odd[0])} {show(odd[0].value)}, 평소 {show(odd[1])} 대비 {percent(odd[2])}" if ko else
                         f"of the latest 7, the day furthest from its usual weekday level: {day(odd[0])} {show(odd[0].value)}, {percent(odd[2])} against a usual {show(odd[1])}"))
    position, size = rank(points)
    pack.append(("rank", f"최근 {size}개 값 중 {position}번째로 높음" if ko else f"number {position} highest of the latest {size} values"))
    top, bottom = extreme(points, 30), extreme(points, 30, highest=False)
    pack.append(("high", f"최근 {min(len(points), 30)}개 중 최고 {show(top.value)} ({day(top)})" if ko else
                 f"highest of the latest {min(len(points), 30)}: {show(top.value)} ({day(top)})"))
    pack.append(("low", f"최근 {min(len(points), 30)}개 중 최저 {show(bottom.value)} ({day(bottom)})" if ko else
                 f"lowest of the latest {min(len(points), 30)}: {show(bottom.value)} ({day(bottom)})"))
    run = streak(points)
    if abs(run) >= 2:
        pack.append(("streak", f"{abs(run)}회 연속 {'증가' if run > 0 else '감소'}" if ko else f"{abs(run)} {'rises' if run > 0 else 'falls'} in a row"))
    change = week_total_change(points) if daily_data else None
    if change is not None:
        pack.append(("week_total", f"최근 7개 합계 {show(sum(p.value for p in points[-7:]))}, 그 전 7개 합계 {show(sum(p.value for p in points[-14:-7]))} 대비 {percent(change)}" if ko else
                     f"latest 7 total {show(sum(p.value for p in points[-7:]))}, {percent(change)} against the 7 before, {show(sum(p.value for p in points[-14:-7]))}"))
    recent = points[-28:]
    weekdays = [p.value for p in recent if p.date.weekday() < 5]
    weekends = [p.value for p in recent if p.date.weekday() >= 5]
    if daily_data and weekdays and len(weekends) >= 2:
        change = ratio(mean(weekends), mean(weekdays))
        if change is not None:
            pack.append(("weekend", f"최근 {len(recent)}개에서 평일 평균 {show(mean(weekdays))}, 주말 평균 {show(mean(weekends))}, 주말이 평일 대비 {percent(change)}" if ko else
                         f"in the latest {len(recent)}: weekday mean {show(mean(weekdays))}, weekend mean {show(mean(weekends))}, weekend {percent(change)} against weekdays"))
    latest =", ".join(f"{day(p)} {show(p.value)}" for p in points[-7:])
    pack.append(("recent", f"최근 값: {latest}" if ko else f"latest values: {latest}"))
    return readings(points, show, lang, limit=9) + [Fact(name, text, tuple(NUMBER.findall(text))) for name, text in pack]


@dataclass(frozen=True)
class Item:
    """One named row of a table with no dates: a value, and its earlier value when two periods are compared."""
    label: str
    value: float
    before: float = None


def ranked_items(items):
    """Largest first, with a catch-all such as 기타 last: it is not a rival to the named items, however large."""
    return sorted(items, key=lambda item: (OTHER_ITEM.fullmatch(item.label.strip()) is not None, -item.value))


def is_rate(header, fmt=""):
    """True for a rate, a score, an index or a mean, judged from its heading and format: a change in one is not a share of anything."""
    return RATE_WORD.search(header or "") is not None or "%" in (fmt or "")


def is_price(header):
    """The price of one thing per row: a fee list's 수수료_원, not 수수료 수입 or 총수수료, which are money taken and do add up."""
    return PRICE_NAME.search(header or "") is not None


def is_additive(header, fmt, values):
    """True when items can be shares of a whole. A rate, an index or a mean cannot, and neither can negative figures."""
    return not is_rate(header, fmt) and all(value >= 0 for value in values) and sum(values) > 0


def outline_kind(label):
    """Which level marker a name starts with: Ⅰ. 1. 1-1. (1) 가. or None. Account tables number their levels this way."""
    return next((kind for kind, pattern in OUTLINE if pattern.match(label.strip())), None)


def top_level(items):
    """The top level of a numbered table, whose lower levels are parts of it and must not be added or ranked alongside it. None when there are no levels."""
    kinds = [kind for kind in (outline_kind(item.label) for item in items) if kind]
    if len(set(kinds)) < 2:
        return None
    kept = [item for item in items if outline_kind(item.label) == kinds[0]]
    return kept if len(kept) >= 2 else None


def cross_file_readings(entries, lang="ko"):
    """What can be said of several files together by calculation alone. entries: (name, change in percent or None, period text, basis). Which way each moved, or why they cannot be compared. How they relate is not in the data, so it is never said."""
    ko = lang == "ko"
    known = [entry for entry in entries if entry[1] is not None]
    if len(known) < 2:
        return []
    periods = list(dict.fromkeys(entry[2] for entry in known))
    if len(periods) > 1:  # each file's slide says what it is about, so this only has to say why they cannot stand side by side
        listed = ", ".join(periods)
        return [Fact("cross_periods", f"파일마다 최신 시점이 달라({listed}) 변화 방향을 나란히 비교할 수 없습니다." if ko else f"The files end at different times ({listed}), so which way they moved cannot be compared side by side.", ())]
    bases = {entry[3] for entry in known}
    if len(bases) > 1:
        return [Fact("cross_basis", "파일마다 비교 기준이 달라 변화 방향을 나란히 비교할 수 없습니다." if ko else "The files compare against different periods, so which way they moved cannot be compared side by side.", ())]
    basis = known[0][3]
    word = {"year": "전년 같은 기간보다" if ko else "against a year earlier", "prev": "직전보다" if ko else "against the period before"}[basis]
    moved = []
    for name, change, _period, _basis in known:
        if ko:
            moved.append(f"{name} {abs(change):.1f}% {'증가' if change > 0 else '감소'}" if change else f"{name} 변동 없음")
        else:
            moved.append(f"{name} {'up' if change > 0 else 'down'} {abs(change):.1f}%" if change else f"{name} unchanged")
    signs = {(change > 0) - (change < 0) for _name, change, _period, _basis in known}
    if signs == {1}:
        ending = "모두 늘었습니다" if ko else "all rose"
    elif signs == {-1}:
        ending = "모두 줄었습니다" if ko else "all fell"
    elif signs == {1, -1} or len(signs) > 1:
        ending = "방향이 서로 다릅니다" if ko else "they moved in different directions"
    else:
        ending = "변동이 없습니다" if ko else "none changed"
    text = f"{periods[0]} {word} {', '.join(moved)}로 {ending}." if ko else f"{periods[0]} {word}: {', '.join(moved)}; {ending}."
    return [Fact("cross_direction", text, tuple(NUMBER.findall(", ".join(moved))))]


def separate_lists(items):
    """True when numbers repeat among the top level, as with I. 유동자산 and I. 유동부채: separate lists on one sheet, whose lines are not rivals."""
    marks = []
    for item in items:
        label = item.label.strip()
        match = next((pattern.match(label) for _kind, pattern in OUTLINE if pattern.match(label)), None)
        if match:
            marks.append(match.group(0).strip())
    return len(marks) >= 2 and len(set(marks)) < len(marks)


def _share(value, total):
    return format_value(value / total, "0.0%")[0]


def without_parts(items):
    """Items with the rows that are parts of another taken out: a bullet in front of the name, or rows that add up exactly to the row above them. None when nothing is nested.

    In a long list one row can equal the next few by chance. There it counts only when most of the list turns out to be parts, level under level, as in a pivot table of terms, regions and schools.
    """
    if len(items) <= MAX_NESTED_ITEMS:
        return _one_level_of_parts(items, True)
    kept = _every_level_of_parts(items)
    if len(kept) * 2 <= len(items) and len(kept) >= 2:
        return kept
    return _one_level_of_parts(items, False)


def _every_level_of_parts(items):
    """From the bottom up, so a school is taken into its region before the region is taken into its term."""
    below = []  # the rows still standing, last row first
    for item in reversed(items):
        added = 0
        for count, later in enumerate(reversed(below), start=1):
            added += later.value
            if count >= 2 and item.value and abs(added - item.value) <= abs(item.value) * NESTED_TOLERANCE:
                del below[-count:]
                break
            if item.value >= 0 and added > item.value * (1 + NESTED_TOLERANCE):
                break
        below.append(item)
    return below[::-1]


def _one_level_of_parts(items, by_sums):
    kept, skip, changed = [], 0, False
    for place, item in enumerate(items):
        if skip:
            skip -= 1
            changed = True
            continue
        if PART_MARK.match(item.label):
            changed = True
            continue
        added = 0
        for count, later in enumerate(items[place + 1:] if by_sums else (), start=1):
            if PART_MARK.match(later.label):
                continue
            added += later.value
            if count >= 2 and abs(item.value) >= MIN_NESTED_PARENT and abs(added - item.value) <= abs(item.value) * NESTED_TOLERANCE:
                skip = count
                break
            if item.value >= 0 and added > item.value * (1 + NESTED_TOLERANCE):
                break
        kept.append(item)
    return kept if changed and len(kept) >= 2 else None


def _against_budget(item, show, lang):
    """A line set against its budget: 실적 33,037, 예산 28,120 대비 +17.5%. A budget of 0 or less gives no percentage."""
    ko = lang == "ko"
    change = f", {percent(ratio(item.value, item.before))}" if item.before > 0 else ""
    if ko:
        return f"{item.label} 실적 {show(item.value)}, 예산 {show(item.before)}" + (f" 대비 {change[2:]}" if change else "") + "."
    return f"{item.label} came to {show(item.value)} against a budget of {show(item.before)}{change}."


def item_readings(items, show, lang="ko", total=None, additive=True, before_total=None, limit=3, stated=True, rate=False, scale=1, basis="", wholes=()):
    """What a table of named items says, by fixed rules. Largest and its share, or what moved most when there are two periods.

    stated says the whole came from the sheet. Otherwise the items were added up here, and the wording does not call it the whole.
    basis is budget when the older column is what was planned for the same period: the words then say actual against budget, not a rise or a fall.
    wholes are the total lines of a table that totals unlike things (Total income, Total expenses). Its lines are not ranked against one another, and only the wholes are set against their budget.
    """
    ko = lang == "ko"
    budget = basis == "budget"
    whole = ("전체" if stated else "항목 합계") if ko else ("the total" if stated else "the items' sum")
    if len(items) < 2:
        return []
    found = []
    if len(wholes) >= 2:
        paired = [item for item in wholes if item.before is not None]
        if paired and budget:
            found = [("read_whole_budget", _against_budget(item, show, lang)) for item in paired]
        elif paired:
            found = [("read_whole_change", f"{item.label}는 {show(item.before)} → {show(item.value)}" + (f", {percent(ratio(item.value, item.before))}입니다." if item.before > 0 else "입니다.") if ko else
                      f"{item.label} went from {show(item.before)} to {show(item.value)}" + (f", {percent(ratio(item.value, item.before))}." if item.before > 0 else ".")) for item in paired]
        else:
            found = [("read_separate", SEPARATE_TOTALS[lang])]
        return [Fact(name, text, tuple(NUMBER.findall(text))) for name, text in found[:limit]]
    if any(item.before is not None for item in items):
        pairs = [item for item in items if item.before is not None]
        if budget:
            pair = lambda item, change: f"예산 {show(item.before)}, 실적 {show(item.value)}, {change}" if ko else f"{show(item.value)} against a budget of {show(item.before)}, {change}"  # noqa: E731
        else:
            pair = lambda item, change: f"{show(item.before)} → {show(item.value)}, {change}" if ko else f"from {show(item.before)} to {show(item.value)}, {change}"  # noqa: E731
        if not rate:
            if additive and total is not None and before_total:
                if budget:
                    found.append(("read_total_change", f"{whole}는 예산 {show(before_total)} 대비 실적 {show(total)}, {percent(ratio(total, before_total))}입니다." if ko else
                                  f"{whole.capitalize()} came to {show(total)} against a budget of {show(before_total)}, {percent(ratio(total, before_total))}."))
                else:
                    found.append(("read_total_change", f"{whole}는 {show(before_total)} → {show(total)}, {percent(ratio(total, before_total))}입니다." if ko else
                                  f"{whole.capitalize()} went from {show(before_total)} to {show(total)}, {percent(ratio(total, before_total))}."))
            # A tiny item can double and mean nothing, so only items of some size are named.
            size = max(max(item.value, item.before) for item in pairs)
            moves = [(ratio(item.value, item.before), item) for item in pairs if item.before > 0 and max(item.value, item.before) >= size * SMALL_ITEM]
            word = lambda change: percent(change)  # noqa: E731
        else:
            moves = [(item.value - item.before, item) for item in pairs]
            word = lambda change: points_text(change, lang, scale)  # noqa: E731
        ups, downs = [m for m in moves if m[0] > 0], [m for m in moves if m[0] < 0]
        if ups:
            change, item = max(ups, key=lambda m: m[0])
            if budget:
                found.append(("read_most_up", f"예산보다 가장 크게 웃돈 항목은 {item.label}({pair(item, word(change))})입니다." if ko else f"{item.label} ran furthest over budget, {pair(item, word(change))}."))
            else:
                found.append(("read_most_up", f"가장 크게 늘어난 항목은 {item.label}({pair(item, word(change))})입니다." if ko else f"{item.label} rose most, {pair(item, word(change))}."))
            if additive and total is not None:  # A small item can lead by percent while another adds far more. With overlapping rows the sum says nothing.
                biggest = max(ups, key=lambda m: m[1].value - m[1].before)[1]
                if biggest is not item:
                    grew = percent(ratio(biggest.value, biggest.before))
                    if budget:
                        found.append(("read_most_added", f"예산 초과액이 가장 큰 항목은 {biggest.label}({pair(biggest, grew)})입니다." if ko else f"{biggest.label} went over budget by the most, {pair(biggest, grew)}."))
                    else:
                        found.append(("read_most_added", f"증가량이 가장 큰 항목은 {biggest.label}({pair(biggest, grew)})입니다." if ko else f"{biggest.label} added the most, {pair(biggest, grew)}."))
        if downs:
            change, item = min(downs, key=lambda m: m[0])
            if budget:
                found.append(("read_most_down", f"예산에 가장 크게 못 미친 항목은 {item.label}({pair(item, word(change))})입니다." if ko else f"{item.label} fell furthest short of budget, {pair(item, word(change))}."))
            else:
                found.append(("read_most_down", f"가장 크게 줄어든 항목은 {item.label}({pair(item, word(change))})입니다." if ko else f"{item.label} fell most, {pair(item, word(change))}."))
        # A percentage of a negative base means nothing, but a loss that moved is often the most important line of the table.
        negative = sorted((item for item in pairs if item.before < 0 and not rate), key=lambda item: -abs(item.value - item.before))[:2]
        if negative:
            named = ", ".join(f"{item.label}({show(item.before)} → {show(item.value)})" for item in negative)
            found.append(("read_negative_base", f"기준값이 음수라 비율을 낼 수 없는 항목: {named}." if ko else f"A negative base gives no percentage for: {named}."))
        if not moves and not negative:
            if budget:
                found.append(("read_no_change", "실적과 예산이 같습니다." if ko else "The actuals match the budget."))
            else:
                found.append(("read_no_change", "두 기간의 값이 같습니다." if ko else "The two periods have the same values."))
        return [Fact(name, text, tuple(NUMBER.findall(text))) for name, text in found[:limit]]
    if separate_lists(items):
        return [Fact("read_separate", SEPARATE_LISTS[lang], ())]
    ordered = ranked_items(items)
    named = [item for item in ordered if not OTHER_ITEM.fullmatch(item.label.strip())] or ordered
    top, second, bottom = ordered[0], ordered[1], named[-1]
    if additive and total:
        found.append(("read_top", f"가장 큰 항목은 {top.label}({show(top.value)})이고, {whole} {show(total)}의 {_share(top.value, total)}입니다." if ko else
                      f"The largest is {top.label} ({show(top.value)}), {_share(top.value, total)} of {whole}, {show(total)}."))
        if second.value > 0 and top.value / second.value < CLOSE_SECOND:
            found.append(("read_gap", f"2위 {second.label}({show(second.value)}, {_share(second.value, total)})와 차이가 크지 않습니다." if ko else
                          f"The second, {second.label} ({show(second.value)}, {_share(second.value, total)}), is close behind."))
        elif second.value > 0 and top.value / second.value < 100:  # beyond this a multiple tells a reader nothing the share does not
            found.append(("read_gap", f"2위 {second.label}({show(second.value)}, {_share(second.value, total)})의 {top.value / second.value:.1f}배입니다." if ko else
                          f"That is {top.value / second.value:.1f} times the second, {second.label} ({show(second.value)}, {_share(second.value, total)})."))
        if len(ordered) >= 5:
            found.append(("read_top_three", f"항목 {len(ordered)}개 중 상위 3개가 {whole}의 {_share(sum(item.value for item in ordered[:3]), total)}입니다." if ko else
                          f"The top 3 of {len(ordered)} items are {_share(sum(item.value for item in ordered[:3]), total)} of {whole}."))
    else:
        found.append(("read_range", f"항목 {len(ordered)}개 중 가장 높은 항목은 {top.label}({show(top.value)}), 가장 낮은 항목은 {bottom.label}({show(bottom.value)})입니다." if ko else
                      f"Of {len(ordered)} items, the highest is {top.label} ({show(top.value)}) and the lowest is {bottom.label} ({show(bottom.value)})."))
    return [Fact(name, text, tuple(NUMBER.findall(text))) for name, text in found[:limit]]


def item_sentence(items, header, show, lang="ko", total=None, additive=True, before_total=None, stated=True, rate=False, scale=1, basis="", wholes=()):
    """A plain report sentence for a table of named items. It only states figures."""
    ko = lang == "ko"
    if any(item.before is not None for item in items):
        first = item_readings(items, show, lang, total, additive, before_total, limit=1, stated=stated, rate=rate, scale=scale, basis=basis, wholes=wholes)
        return f"{header}: {first[0].text}" if first else f"{header}."
    if len(wholes) >= 2:  # income, expenses and results are not rivals: their totals are what the table says
        return f"{header}: " + ", ".join(f"{item.label} {show(item.value)}" for item in wholes[:MAX_LISTED]) + "."
    if separate_lists(items):
        return f"{header}: " + ", ".join(f"{item.label} {show(item.value)}" for item in items[:MAX_LISTED]) + "."
    ordered = ranked_items(items)
    if additive and total:
        tops = ", ".join(f"{item.label} {_share(item.value, total)}" for item in ordered[:3])
        whole = ("전체" if stated else "항목 합계") if ko else ("total" if stated else "sum of items")
        name = f"{header} " if whole not in header else f"{header.replace(whole, '').strip()} "
        return f"{name}{whole} {show(total)}. {tops}."
    tops = ", ".join(f"{item.label} {show(item.value)}" for item in ordered[:3])
    return f"{header}: {tops}."


def item_pack(items, show, lang="ko", total=None, additive=True, before_total=None, stated=True, rate=False, scale=1, basis="", wholes=()):
    """The calculated statements about a table of named items, numbered so that a later step can cite them.
    With several total lines the lines are given in the sheet's order, not as places in a ranking."""
    ko = lang == "ko"
    readings = item_readings(items, show, lang, total, additive, before_total, limit=9, stated=stated, rate=rate, scale=scale, basis=basis, wholes=wholes)
    pack = []
    if len(wholes) >= 2:
        for number, item in enumerate(wholes[:10], start=1):
            planned = (f", 예산 {show(item.before)}" if ko else f", budget {show(item.before)}") if basis == "budget" and item.before is not None else (f", 이전 {show(item.before)}" if ko else f", from {show(item.before)}") if item.before is not None else ""
            pack.append((f"whole_{number}", f"합계 줄 {item.label} {show(item.value)}{planned}" if ko else f"total line {item.label} {show(item.value)}{planned}"))
        return readings + [Fact(name, text, tuple(NUMBER.findall(text))) for name, text in pack]
    for place, item in enumerate(ranked_items(items)[:10], start=1):
        share = f", {_share(item.value, total)}" if additive and total else ""
        if item.before is None:
            earlier = ""
        elif basis == "budget":
            earlier = f", 예산 {show(item.before)}" if ko else f", budget {show(item.before)}"
        else:
            earlier = f", 이전 {show(item.before)}" if ko else f", from {show(item.before)}"
        pack.append((f"item_{place}", f"{place}위 {item.label} {show(item.value)}{share}{earlier}" if ko else f"number {place} {item.label} {show(item.value)}{share}{earlier}"))
    return readings + [Fact(name, text, tuple(NUMBER.findall(text))) for name, text in pack]


def rate_found(points, show, lang="ko", scale=1):
    """What a series of rates or scores says, in points. (id, text) pairs: its own rules, because a change of a percentage is not a percentage."""
    ko = lang == "ko"
    found = []
    if len(points) < 2:
        return found
    newest, before = points[-1], points[-2]
    day = lambda p: point_text(p.date, points, lang)  # noqa: E731
    if is_monthly(points) or is_quarterly(points):
        base = year_ago(points)
        word = WORDS[lang]["달" if is_monthly(points) else "분기"]
        if base:
            found.append(("read_year_points", f"전년 같은 {word} {day(base)} {show(base.value)} 대비 {points_text(newest.value - base.value, lang, scale)}입니다." if ko else
                          f"{points_text(newest.value - base.value, lang, scale)} against the same {word} last year, {day(base)}, {show(base.value)}."))
    found.append(("read_prev_points", f"직전 {day(before)} {show(before.value)} 대비 {points_text(newest.value - before.value, lang, scale)}입니다." if ko else
                  f"{points_text(newest.value - before.value, lang, scale)} against {day(before)}, {show(before.value)}."))
    run = streak(points)
    if abs(run) >= WEEKS_IN_A_ROW:
        found.append(("read_points_run", f"{abs(run)}번 연속 {'상승' if run > 0 else '하락'}했습니다." if ko else f"It has {'risen' if run > 0 else 'fallen'} {abs(run)} times in a row."))
    position, size = rank(points)
    if size >= 5 and position == 1:
        found.append(("read_extreme", f"최근 {size}개 중 가장 높은 값입니다." if ko else f"This is the highest of the latest {size} values."))
    elif size >= 5 and not any(p.value < newest.value for p in points[-30:]):
        found.append(("read_extreme", f"최근 {size}개 중 가장 낮은 값입니다." if ko else f"This is the lowest of the latest {size} values."))
    return found


def rate_readings(points, show, lang="ko", scale=1):
    return [Fact(name, text, tuple(NUMBER.findall(text))) for name, text in rate_found(points, show, lang, scale)]


def rate_sentence(points, header, show, lang="ko", scale=1):
    found = [text for name, text in rate_found(points, show, lang, scale) if name in ("read_year_points", "read_prev_points")]
    head = f"{point_text(points[-1].date, points, lang)} {header} {show(points[-1].value)}."
    return " ".join([head] + found)


def rate_pack(points, header, show, lang="ko", scale=1):
    value = (f"{point_text(points[-1].date, points, lang)} {header}: {show(points[-1].value)}")
    recent = ", ".join(f"{point_text(p.date, points, lang)} {show(p.value)}" for p in points[-7:])
    return rate_readings(points, show, lang, scale) + [Fact(name, text, tuple(NUMBER.findall(text))) for name, text in (("value", value), ("recent", f"최근 값: {recent}" if lang == "ko" else f"latest values: {recent}"))]


def cut_short_text(points, found, usual, show, lang="ko"):
    """The sentence for a newest period with fewer rows than usual: its total cannot be set against a whole one, so the comparison is per row."""
    newest, before = points[-1], points[-2] if len(points) > 1 else None
    day = lambda p: point_text(p.date, points, lang)  # noqa: E731
    if lang == "ko":
        text = f"{day(newest)} 값은 행이 {found}개로 평소 {usual}개보다 적어 덜 쌓인 값으로 보입니다."
    else:
        text = f"{day(newest)} has {found} rows where {usual} is usual, so it looks incomplete."
    if before:
        whole, each = ratio(newest.value, before.value), ratio(newest.value / found, before.value / usual)
        if whole is not None and each is not None:
            text += (f" {day(before)} {show(before.value)} 대비 {percent(whole)}는 그대로 읽을 수 없고, 행당 평균으로 맞추면 {percent(each)}입니다. 빠진 행의 계절 차이는 반영되지 않습니다." if lang == "ko" else
                     f" {percent(whole)} against {day(before)}, {show(before.value)}, cannot be read as it stands. Per row it is {percent(each)}, which does not allow for the season of the missing rows.")
    return text


def sentence(points, header, show, lang="ko", count=7):
    """A plain report sentence. It only states figures, so it needs no model and cannot invent anything."""
    words = WORDS[lang]
    newest = points[-1]
    parts = []
    base = week_ago(points)
    change = ratio(newest.value, base.value) if base else None
    if change is not None:
        parts.append(_movement(change, words["week"], lang))
    against_year = None
    if is_monthly(points) or is_quarterly(points):
        base = year_ago(points)
        against_year = ratio(newest.value, base.value) if base else None
        if against_year is not None:
            parts.append(_movement(against_year, words["year"].format(unit=words["달" if is_monthly(points) else "분기"]), lang))
    base = previous(points)
    change = ratio(newest.value, base.value) if base and not year_to_date(points) else None
    if change is not None:
        one_day = (newest.date.date() - base.date.date()).days == 1
        lead = words["prev_day"] if one_day else words["prev"].format(date=point_text(base.date, points, lang))
        parts.append(_movement(change, lead, lang))
    months = partial_year(points)
    if months:
        change = ratio(newest.value / months, base.value / 12)
        if change is not None:
            parts.append(_movement(change, words["per_month"], lang))
    # A mean of the months before says little once the year before is known: it mostly shows the season.
    base = mean_before(points, count) if not months and against_year is None else None
    change = ratio(newest.value, base) if base is not None else None
    if change is not None:
        parts.append(_movement(change, words["avg"].format(n=count), lang))
    head = f"{point_text(newest.date, points, lang)} {header} {show(newest.value)}"
    return head + (". " + ", ".join(parts) + "." if parts else ".")
