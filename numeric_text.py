"""Bounded, exact identities for numbers quoted in prose. No GUI or network.

These checks compare numeric spelling with supplied facts; they do not establish
that a number belongs to the right subject, unit, sign word, or business meaning.
"""

import re


# A broad token deliberately keeps malformed grouping/decimal/exponent syntax
# together so it is rejected, rather than quietly checked as smaller numbers.
NUMBER = re.compile(r"[+-]?(?:\d[\d,.]*|\.\d+)(?:[eE][+-]?\d*)?")
NOT_A_FIGURE = re.compile(r"\bQ\d\b|(?<![\d.,])\d\s*분기|(?<![\d.,])\d+\s*(?:월(?!\s*연속)|일|년)|(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\.? \d+|(?<![\d.,])\d{4}(?![\d,.])")
_LITERAL = re.compile(r"([+-]?)(?:(\d+|\d{1,3}(?:,\d{3})+)(?:\.(\d*))?|\.(\d+))(?:[eE]([+-]?\d+))?", re.ASCII)
MAX_TOKEN_LENGTH = 4096
MAX_EXPONENT = 4096


def number_identity(text, *, magnitude=False):
    """Return a normalized decimal tuple, or None for unsupported input.

    No binary float, rounding, power expansion, or context precision is involved.
    The coefficient and exponent stay as bounded strings/integers. A final comma
    may be prose punctuation, as in ``11, while``; internal commas must group
    thousands. A final dot is the ordinary spelling ``11.`` or sentence end.
    """
    if not isinstance(text, str) or not text or len(text) > MAX_TOKEN_LENGTH:
        return None
    if text.endswith(","):
        text = text[:-1]
    if text.endswith(".") and len(text) > 1 and text[-2].isdigit():
        text = text[:-1]
    match = _LITERAL.fullmatch(text)
    if not match:
        return None
    sign, integer, fraction, leading_fraction, exponent = match.groups()
    if exponent:
        # Check before int conversion, even for an all-zero oversized exponent.
        digits = exponent.lstrip("+-")
        if len(digits) > 4 or abs(int(exponent)) > MAX_EXPONENT:
            return None
    power = int(exponent or "0")
    fraction = fraction if integer is not None else leading_fraction
    fraction = fraction or ""
    coefficient = ((integer or "0").replace(",", "") + fraction).lstrip("0")
    if not coefficient:
        return (1, "0", 0)
    trimmed = coefficient.rstrip("0")
    power += len(coefficient) - len(trimmed) - len(fraction)
    return (-1 if sign == "-" and not magnitude else 1, trimmed, power)


def figure_tokens(text):
    """Return whole numeric claims except tokens wholly covered by a date.

    Span containment avoids deleting a year-shaped piece of a larger literal
    (such as the exponent in ``1e2025``) before checking that literal.
    """
    dates = tuple(found.span() for found in NOT_A_FIGURE.finditer(text))
    return tuple(found.group(0) for found in NUMBER.finditer(text)
                 if not any(start <= found.start() and found.end() <= end for start, end in dates))


def unbacked_numbers(text, allowed_numbers, *, magnitude=True):
    """Return unsupported/unbacked numeric claims, allowing exact equivalence.

    Magnitude matching preserves the wording check's existing separation of
    numeric identity from sign/direction semantics. Callers needing signed
    comparison must explicitly pass magnitude=False and signed allowed tokens.
    """
    allowed = {identity for raw in allowed_numbers
               if (identity := number_identity(raw, magnitude=magnitude)) is not None}
    return tuple(raw for raw in figure_tokens(text)
                 if (identity := number_identity(raw, magnitude=magnitude)) is None or identity not in allowed)
