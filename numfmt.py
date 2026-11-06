"""Renders a cell value the way Excel displays it, from its number format string. No GUI imports.

Covers the formats people use in reports: thousands, decimals, percent, scaling, currency, scientific,
dates and times, sections and literals. Anything else raises Unsupported internally, and format_value
reports it by returning exact=False so the caller can warn instead of showing a wrong value as if it were right.
"""

import ctypes
import datetime as dt
import operator
import re
from dataclasses import dataclass
from decimal import ROUND_HALF_UP, Decimal, localcontext

from openpyxl.utils.datetime import WINDOWS_EPOCH, from_excel

MONTHS = ["January", "February", "March", "April", "May", "June", "July",
          "August", "September", "October", "November", "December"]
DAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
DAYS_KO = ["월", "화", "수", "목", "금", "토", "일"]

# Built-in formats whose meaning follows the PC's locale. openpyxl reports the US meaning for all of them.
# The right-hand sides are what Excel on Korean Windows reports (see tools/compare_with_excel.py).
KOREAN_BUILTINS = {
    '"$"#,##0_);("$"#,##0)': "₩#,##0;-₩#,##0",
    '"$"#,##0_);[Red]("$"#,##0)': "₩#,##0;[Red]-₩#,##0",
    '"$"#,##0.00_);("$"#,##0.00)': "₩#,##0.00;-₩#,##0.00",
    '"$"#,##0.00_);[Red]("$"#,##0.00)': "₩#,##0.00;[Red]-₩#,##0.00",
    "#,##0_);(#,##0)": "#,##0;-#,##0",
    "#,##0_);[Red](#,##0)": "#,##0;[Red]-#,##0",
    "#,##0.00_);(#,##0.00)": "#,##0.00;-#,##0.00",
    "#,##0.00_);[Red](#,##0.00)": "#,##0.00;[Red]-#,##0.00",
    '_(* #,##0_);_(* \\(#,##0\\);_(* "-"_);_(@_)': '_-* #,##0_-;-* #,##0_-;_-* "-"_-;_-@_-',
    '_("$"* #,##0_);_("$"* \\(#,##0\\);_("$"* "-"_);_(@_)': '_-₩* #,##0_-;-₩* #,##0_-;_-₩* "-"_-;_-@_-',
    '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)': '_-* #,##0.00_-;-* #,##0.00_-;_-* "-"??_-;_-@_-',
    '_("$"* #,##0.00_)_("$"* \\(#,##0.00\\)_("$"* "-"??_)_(@_)': '_-₩* #,##0.00_-;-₩* #,##0.00_-;_-₩* "-"??_-;_-@_-',
}


class Unsupported(Exception):
    pass


@dataclass(frozen=True)
class Locale:
    date_separator: str = "/"
    short_date: str = "m/d/yyyy"
    korean: bool = False


def system_locale():
    """Read the Windows regional settings. Excel uses these for / in dates and for its locale-dependent built-ins."""
    try:
        kernel32 = ctypes.windll.kernel32

        def info(code):
            buffer = ctypes.create_unicode_buffer(80)
            kernel32.GetLocaleInfoW(0x0400, code, buffer, 80)  # 0x0400 = the current user's settings
            return buffer.value

        short_date = info(0x1F).lower().replace("'", '"')  # LOCALE_SSHORTDATE
        return Locale(info(0x1D) or "/", short_date or "m/d/yyyy",  # LOCALE_SDATE
                      kernel32.GetUserDefaultLCID() & 0x3FF == 0x12)
    except (AttributeError, OSError):
        return Locale()


LOCALE = system_locale()


def _localize(fmt):
    if fmt == "mm-dd-yy":
        return LOCALE.short_date
    if fmt == "m/d/yy h:mm":
        return LOCALE.short_date + " h:mm"
    return KOREAN_BUILTINS.get(fmt, fmt) if LOCALE.korean else fmt


def format_value(value, number_format="General", epoch=WINDOWS_EPOCH):
    """Return (text, exact). exact is False when the format was not understood and General was used instead."""
    try:
        return _format(value, number_format or "General", epoch).strip(), True
    except Unsupported:
        return _general(value), False


# [Red] and the other names a format may color a section with, as Excel draws them. Korean Excel shows the
# same eight in Korean words, so they are read too. Black and White are left out on purpose: they mean
# "plain text" in the sheet, and on a slide the template's own text color says that better (white on a
# white slide would hide the number).
FORMAT_COLORS = {"blue": "0000FF", "cyan": "00FFFF", "green": "00FF00", "magenta": "FF00FF", "red": "FF0000",
                 "yellow": "FFFF00", "파랑": "0000FF", "녹청": "00FFFF", "녹색": "00FF00", "자홍": "FF00FF",
                 "빨강": "FF0000", "노랑": "FFFF00"}
PLAIN_COLORS = {"black", "white", "검정", "흰색"}
# [Color1]..[Color56]: Excel's standard palette (a workbook may change it; this file's own palette is not read).
PALETTE = ("000000 FFFFFF FF0000 00FF00 0000FF FFFF00 FF00FF 00FFFF 800000 008000 000080 808000 800080 008080 "
           "C0C0C0 808080 9999FF 993366 FFFFCC CCFFFF 660066 FF8080 0066CC CCCCFF 000080 FF00FF FFFF00 00FFFF "
           "800080 800000 008080 0000FF 00CCFF CCFFFF CCFFCC FFFF99 99CCFF FF99CC CC99FF FFCC99 3366FF 33CCCC "
           "99CC00 FFCC00 FF9900 FF6600 666699 969696 003366 339966 003300 333300 993300 993366 333399 333333").split()


def format_color(value, number_format="General"):
    """The color a number's format section asks for, as RRGGBB, or None for the text's own color.

    #,##0;[Red]-#,##0 gives FF0000 for -5 and None for 5. Only numbers: a format colors text in its fourth
    section, which reports do not use.
    """
    if not isinstance(value, (int, float)) or isinstance(value, bool):
        return None
    try:
        section, _negate = _pick_section(value, _split_sections(_localize(number_format or "General")))
    except Unsupported:
        return None
    bare = re.sub(r'"[^"]*"?|\\.|_.|\*.', "", section)
    for name in re.findall(r"\[([^\]]*)\]", bare):
        low = name.strip().casefold()
        if low in FORMAT_COLORS:
            return FORMAT_COLORS[low]
        number = re.fullmatch(r"(?:color|색)\s*(\d{1,2})", low)
        if number and 1 <= int(number.group(1)) <= len(PALETTE):
            found = PALETTE[int(number.group(1)) - 1]
            return None if found in ("000000", "FFFFFF") else found
        if low in PLAIN_COLORS:
            return None
    return None


def _format(value, fmt, epoch):
    if value is None:
        return ""
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    fmt = _localize(fmt)
    sections = _split_sections(fmt)
    if isinstance(value, str):
        return _format_text(value, sections)
    if isinstance(value, (dt.datetime, dt.date, dt.time, dt.timedelta)):
        tokens = _tokenize(sections[0])
        return _render_date(value, tokens) if _is_date(tokens) else _general(value)
    section, negate = _pick_section(value, sections)
    tokens = _tokenize(section)
    if _is_date(tokens):
        if value < 0:
            raise Unsupported(fmt)
        elapsed = any(t[0] == "elapsed" for t in tokens)
        return _render_date(from_excel(value, epoch, timedelta=elapsed), tokens)
    if ("sym", "@") in tokens and not any(t[0] == "digit" for t in tokens):
        return _general(value)  # Excel shows a number in a text format as General and drops the literals
    if any(t[0] == "general" for t in tokens):
        return ("-" if negate else "") + _literals(tokens, _general(abs(value)))
    if any(t[0] == "digit" for t in tokens):
        return _render_number(abs(value), tokens, negate)
    return _literals(tokens, "")


CONDITION = re.compile(r"\[(>=|<=|<>|>|<|=)\s*(-?\d+(?:\.\d+)?)\]")
_COMPARE = {">=": operator.ge, "<=": operator.le, "<>": operator.ne, ">": operator.gt, "<": operator.lt, "=": operator.eq}


def _pick_by_condition(value, sections, conditions):
    """[>=5]#,##0;[<=-5]\\-#,##0;\\<5, as some publishers hide small counts: the first section whose test passes, else the last with no test of its own."""
    for place, (section, condition) in enumerate(zip(sections, conditions)):
        if condition and _COMPARE[condition.group(1)](value, float(condition.group(2))):
            return section.replace(condition.group(0), "", 1), place == 0 and value < 0
    rest = next((section for section, condition in zip(sections, conditions) if not condition), None)
    if rest is None:
        raise Unsupported("no section for this value")
    return rest, value < 0  # Excel keeps the minus here: [>=100]0;0 shows -5 as -5


def _split_sections(fmt):
    sections, buf, i = [], "", 0
    while i < len(fmt):
        c = fmt[i]
        if c == '"':
            end = fmt.find('"', i + 1)
            end = len(fmt) - 1 if end < 0 else end
            buf += fmt[i:end + 1]
            i = end + 1
        elif c in "\\_*" and i + 1 < len(fmt):
            buf += fmt[i:i + 2]
            i += 2
        elif c == "[":
            end = fmt.find("]", i)
            end = len(fmt) - 1 if end < 0 else end
            buf += fmt[i:end + 1]
            i = end + 1
        elif c == ";":
            sections.append(buf)
            buf = ""
            i += 1
        else:
            buf += c
            i += 1
    sections.append(buf)
    return sections


def _pick_section(value, sections):
    """Return (section, negate). A negative value only gets an automatic minus when no section is meant for it."""
    conditions = [CONDITION.search(section) for section in sections]
    if any(conditions):
        return _pick_by_condition(value, sections, conditions)
    if len(sections) == 1:
        return sections[0], value < 0
    if value < 0:
        return sections[1], False
    if value == 0 and len(sections) >= 3:
        return sections[2], False
    return sections[0], False


def _tokenize(section):
    tokens, i, n = [], 0, len(section)
    while i < n:
        c, low = section[i], section[i:].lower()
        if c == '"':
            end = section.find('"', i + 1)
            end = n if end < 0 else end
            tokens.append(("lit", section[i + 1:end]))
            i = end + 1
        elif c == "\\":
            tokens.append(("lit", section[i + 1:i + 2]))
            i += 2
        elif c == "_":
            tokens.append(("lit", " "))
            i += 2
        elif c == "*":
            i += 2
        elif c == "[":
            end = section.find("]", i)
            end = n if end < 0 else end
            tokens.extend(_bracket(section[i + 1:end]))
            i = end + 1
        elif low.startswith("general"):
            tokens.append(("general", ""))
            i += 7
        elif low.startswith("am/pm") or low.startswith("a/p"):
            size = 5 if low.startswith("am/pm") else 3
            tokens.append(("ampm", low[:size]))
            i += size
        elif low.startswith("aaa"):
            size = 4 if low.startswith("aaaa") else 3
            tokens.append(("date", "a" * size))
            i += size
        elif c in "yYmMdDhHsS":
            size = 1
            while i + size < n and section[i + size].lower() == c.lower():
                size += 1
            run = c.lower() * size
            i += size
            if c in "sS":
                frac = re.match(r"\.0+", section[i:])
                if frac:
                    run += frac.group()
                    i += len(frac.group())
            tokens.append(("date", run))
        elif low[:2] in ("e+", "e-"):
            tokens.append(("sci", low[1]))
            i += 2
        elif c in "gGeEbBaA":
            raise Unsupported(section)
        elif c in "0#?":
            tokens.append(("digit", c))
            i += 1
        elif c in ".,%@":
            tokens.append(("sym", c))
            i += 1
        elif c == "/":
            tokens.append(("slash", c))
            i += 1
        else:
            tokens.append(("lit", c))
            i += 1
    return tokens


def _bracket(body):
    low = body.lower()
    if re.fullmatch(r"h+|m+|s+", low):
        return [("elapsed", low)]
    if body.startswith("$"):
        symbol, _, locale = body[1:].partition("-")
        locale = locale.lower()
        if locale.startswith(("f800", "f400")):
            raise Unsupported(body)  # "system long date": depends on the PC that shows it
        out = [("lit", symbol)] if symbol else []
        if locale.endswith("412") or locale.startswith("ko"):
            out.append(("korean", ""))
        return out
    if re.match(r"(?i)dbnum|<|>|=", body):
        raise Unsupported(body)
    return []  # a color such as [Red]


def _is_date(tokens):
    return any(t[0] in ("date", "elapsed", "ampm") for t in tokens)


def _literals(tokens, value_text):
    out = ""
    for kind, text in tokens:
        if kind in ("lit", "slash"):
            out += text
        elif kind == "general":
            out += value_text
        elif kind == "sym":
            out += value_text if text == "@" else text
    return out


def _format_text(value, sections):
    for section in reversed(sections):
        if "@" in re.sub(r'"[^"]*"', "", section):
            return _literals(_tokenize(section), value)
    return value


def _render_number(value, tokens, negate):
    positions = [i for i, t in enumerate(tokens) if t[0] in ("digit", "sci")]
    first, last = positions[0], positions[-1]
    body = tokens[first:last + 1]
    if any(t[0] in ("lit", "slash") or t[0] == "sym" and t[1] in "%@" for t in body):
        raise Unsupported("literal inside digits, e.g. a fraction or a phone mask")
    after = last + 1
    while after < len(tokens) and tokens[after] == ("sym", ","):
        after += 1
    scale = after - last - 1
    percent = sum(1 for t in tokens if t == ("sym", "%"))
    prefix = _literals(tokens[:first], "")
    suffix = _literals(tokens[after:], "")

    sci = next((i for i, t in enumerate(body) if t[0] == "sci"), None)
    mantissa = body if sci is None else body[:sci]
    point = next((i for i, t in enumerate(mantissa) if t == ("sym", ".")), None)
    int_part = mantissa if point is None else mantissa[:point]
    frac_part = [] if point is None else mantissa[point + 1:]
    while int_part and int_part[-1] == ("sym", ","):
        int_part = int_part[:-1]
        scale += 1
    grouping = ("sym", ",") in int_part
    int_digits = [t[1] for t in int_part if t[0] == "digit"]
    frac_digits = [t[1] for t in frac_part if t[0] == "digit"]

    with localcontext() as ctx:
        ctx.prec = 80
        number = Decimal(repr(value)) if isinstance(value, float) else Decimal(value)
        number = number * Decimal(100) ** percent / Decimal(1000) ** scale
        if sci is None:
            digits = _digits(number, int_digits, frac_digits, grouping)
        else:
            if len(int_digits) != 1:
                raise Unsupported("engineering notation")
            exp_width = len(body) - sci - 1
            digits = _scientific(number, int_digits, frac_digits, body[sci][1], exp_width)
    # Real Excel drops an automatic minus when the shown figure rounds to zero: -0.00038 as 0.0% is 0.0%.
    return ("-" if negate and any(c in "123456789" for c in digits) else "") + prefix + digits + suffix


def _digits(number, int_digits, frac_digits, grouping):
    """Digits fill placeholders outward from the point. A leftover 0 prints a zero; # and ? print nothing."""
    rounded = number.quantize(Decimal(1).scaleb(-len(frac_digits)), rounding=ROUND_HALF_UP)
    whole, _, frac = f"{rounded:f}".partition(".")
    frac = frac.rstrip("0")
    frac += "0" * frac_digits[len(frac):].count("0")
    whole = whole.lstrip("0")
    whole = "0" * int_digits[:max(len(int_digits) - len(whole), 0)].count("0") + whole
    if grouping:
        whole = re.sub(r"(?<=\d)(?=(\d{3})+$)", ",", whole)
    return whole + ("." + frac if frac else "")


def _scientific(number, int_digits, frac_digits, sign, exp_width):
    exponent = number.adjusted() if number else 0
    mantissa = number.scaleb(-exponent)
    text = _digits(mantissa, int_digits, frac_digits, False)
    if text.startswith("10"):  # rounding carried over, e.g. 9.99 -> 10.0
        exponent += 1
        text = _digits(mantissa.scaleb(-1), int_digits, frac_digits, False)
    mark = "-" if exponent < 0 else ("+" if sign == "+" else "")
    return f"{text}E{mark}{abs(exponent):0{exp_width}d}"


def _render_date(value, tokens):
    if any(t[0] == "elapsed" for t in tokens) or isinstance(value, dt.timedelta):
        return _render_elapsed(value, tokens)
    if isinstance(value, dt.time):
        value = dt.datetime.combine(dt.date(1899, 12, 31), value)
    elif not isinstance(value, dt.datetime):
        value = dt.datetime.combine(value, dt.time())

    dates = [t for t in tokens if t[0] == "date"]
    fraction = next((len(t[1].split(".")[1]) for t in dates if "." in t[1]), 0)
    if not fraction:
        value = (value + dt.timedelta(microseconds=500000)).replace(microsecond=0)
    korean = ("korean", "") in tokens
    twelve = any(t[0] == "ampm" for t in tokens)
    out = ""
    for i, (kind, text) in enumerate(tokens):
        if kind == "date":
            out += _date_part(value, text, _is_minute(dates, tokens[i]) if text[0] == "m" else False, twelve, korean)
        elif kind == "ampm":
            out += _ampm(value, text, korean)
        elif kind == "slash":
            out += LOCALE.date_separator
        elif kind in ("lit", "sym"):
            out += text
        elif kind == "digit":
            raise Unsupported("digits inside a date format")
    return out


def _is_minute(dates, token):
    """Excel's rule: m is minutes right after an hour or right before seconds, otherwise month."""
    i = next(k for k, t in enumerate(dates) if t is token)
    return (i > 0 and dates[i - 1][1][0] == "h") or (i + 1 < len(dates) and dates[i + 1][1][0] == "s")


def _date_part(value, code, is_minute, twelve, korean):
    head, size = code[0], len(code)
    if head == "y":
        return f"{value.year % 100:02d}" if size <= 2 else str(value.year)
    if head == "m":
        if is_minute:
            return f"{value.minute:02d}" if size >= 2 else str(value.minute)
        if size == 1:
            return str(value.month)
        if size == 2:
            return f"{value.month:02d}"
        if korean:
            return f"{value.month}월"
        name = MONTHS[value.month - 1]
        return name[:3] if size == 3 else (name if size == 4 else name[0])
    if head == "d":
        if size == 1:
            return str(value.day)
        if size == 2:
            return f"{value.day:02d}"
        if korean:
            return DAYS_KO[value.weekday()] + ("요일" if size >= 4 else "")
        name = DAYS[value.weekday()]
        return name[:3] if size == 3 else name
    if head == "a":
        return DAYS_KO[value.weekday()] + ("요일" if size == 4 else "")
    if head == "h":
        hour = (value.hour % 12 or 12) if twelve else value.hour
        return f"{hour:02d}" if size >= 2 else str(hour)
    seconds, _, frac = code.partition(".")
    text = f"{value.second:02d}" if len(seconds) >= 2 else str(value.second)
    if frac:
        text += "." + f"{value.microsecond:06d}"[:len(frac)].ljust(len(frac), "0")
    return text


def _ampm(value, code, korean):
    pm = value.hour >= 12
    if korean:
        return "오후" if pm else "오전"
    if code == "am/pm":
        return "PM" if pm else "AM"
    return "P" if pm else "A"


def _render_elapsed(value, tokens):
    if isinstance(value, dt.timedelta):
        total = value
    elif isinstance(value, dt.time):
        total = dt.timedelta(hours=value.hour, minutes=value.minute, seconds=value.second)
    else:
        raise Unsupported("elapsed time on a date")
    seconds = int(round(total.total_seconds()))
    if seconds < 0:
        raise Unsupported("negative elapsed time")
    out, top = "", True
    for kind, text in tokens:
        if kind == "elapsed":
            unit = {"h": 3600, "m": 60, "s": 1}[text[0]]
            out += str(seconds // unit).rjust(len(text), "0")
            seconds %= unit
            top = False
        elif kind == "date" and not top:
            unit = {"m": 60, "s": 1}.get(text[0])
            if unit is None:
                raise Unsupported("mixed elapsed format")
            out += str(seconds // unit).rjust(len(text), "0")
            seconds %= unit
        elif kind == "date":
            raise Unsupported("date before elapsed part")
        elif kind in ("lit", "sym", "slash"):
            out += text
    return out


def _general(value):
    """Excel's General: no grouping, about eleven characters of precision, integers kept whole."""
    if value is None:
        return ""
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, str):
        return value
    if isinstance(value, dt.datetime):
        if value.time() == dt.time():
            return value.date().isoformat()
        return value.isoformat(sep=" ", timespec="seconds")
    if isinstance(value, (dt.date, dt.time)):
        return value.isoformat()
    if isinstance(value, dt.timedelta):
        return str(value)
    if isinstance(value, int) or float(value).is_integer():
        if abs(value) < 1e15:
            return str(int(value))
    size = abs(value)
    if size >= 1e15 or (size and size < 1e-4):
        mantissa, _, exponent = f"{value:.5E}".partition("E")
        if "." in mantissa:
            mantissa = mantissa.rstrip("0").rstrip(".")
        return f"{mantissa}E{exponent}"
    with localcontext() as ctx:
        ctx.prec = 9 if size < 1 else 10
        text = f"{+Decimal(repr(float(value))):f}"
    return text.rstrip("0").rstrip(".") if "." in text else text
