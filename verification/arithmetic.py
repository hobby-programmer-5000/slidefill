"""Exact reference arithmetic; imports no application calculation functions."""
from decimal import Decimal, ROUND_HALF_UP, localcontext
from fractions import Fraction
import re


def exact(value):
    if isinstance(value, bool) or value is None:
        raise ValueError("missing/nonnumeric value")
    text = str(value).strip()
    if not re.fullmatch(r"[+-]?(?:\d+|\d{1,3}(?:,\d{3})+)(?:\.\d+)?", text):
        raise ValueError("unsupported number")
    return Fraction(text.replace(",", ""))


def growth(now, base):
    a, b = exact(now), exact(base)
    if not b:
        raise ValueError("zero denominator")
    return (a - b) / abs(b)  # current application's policy, including negative bases


def rounded(value, places=0):
    if not 0 <= places <= 12:
        raise ValueError("unsupported precision")
    q = Fraction(value)
    with localcontext() as ctx:
        ctx.prec = max(50, len(str(abs(q.numerator))) + len(str(q.denominator)) + places + 5)
        return (Decimal(q.numerator) / Decimal(q.denominator)).quantize(Decimal(1).scaleb(-places), rounding=ROUND_HALF_UP)


def lookup(bands, population):
    """Closed integer intervals; None is UNKNOWN, never inferred infinity."""
    if type(population) is not int or population < 0:
        raise ValueError("invalid population")
    intervals = []
    for low, high, value in bands:
        if type(low) is not int or type(high) is not int or low < 0 or low > high:
            raise ValueError("unknown/invalid interval")
        intervals.append((low, high, value))
    ordered = sorted(intervals)
    if any(b[0] <= a[1] for a, b in zip(ordered, ordered[1:])):
        raise ValueError("overlapping intervals")
    found = [value for low, high, value in intervals if low <= population <= high]
    if len(found) != 1:
        raise ValueError("population not covered uniquely")
    return found[0]


def partition_sum(groups, scope):
    """Explicit atom sets and same measure/unit/period, never numeric coincidence."""
    seen, total, signature = set(), Fraction(), None
    for group in groups:
        kind = group['kind']
        if kind not in ('count', 'amount'):
            raise ValueError("nonadditive measure")
        tag = (kind, group['measure'], group['unit'], group['period'])
        if signature is not None and tag != signature:
            raise ValueError("incompatible measure/unit/period")
        signature = tag
        atoms = set(group['atoms'])
        if len(atoms) != len(group['atoms']) or seen & atoms:
            raise ValueError("duplicate/overlapping scope")
        seen |= atoms
        total += exact(group['value'])
    if seen != set(scope):
        raise ValueError("incomplete/extra scope")
    return total
