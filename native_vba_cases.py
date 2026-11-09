"""Select numeric report cells that share the bundled VBA engine's operand rules."""

from dataclasses import dataclass
from fractions import Fraction
import math
from pathlib import Path
import sys
from typing import TypedDict

from openpyxl.styles.numbers import is_date_format

from verification import formulas, recipes


MAX_CASES = 100
MAX_CASE_OPERANDS = 1_000
MAX_TOTAL_OPERANDS = 10_000
OPERATIONS = frozenset({'sum', 'mean', 'count'})
EXCEL_SUFFIXES = frozenset({'.xlsx', '.xlsm', '.xltx', '.xltm'})


class VbaCase(TypedDict):
    id: str
    operation: str
    operands: list[Fraction | None]
    expected: Fraction
    label: str


@dataclass
class VbaCases:
    cases: list[VbaCase]
    skipped: int


def _operand(book, cell):
    """Return a numeric source literal or a true blank from the independent reader."""
    sheet = book[cell['sheet']]
    ref = cell['cell']
    if ref in (sheet.formulas or {}) or any(formulas.inside(ref, span) for span in sheet.formula_regions or ()):
        raise ValueError('formula result')
    if (sheet.types or {}).get(ref, 'n') != 'n':
        raise ValueError('non-numeric cell type')
    raw = sheet.get(ref)
    if raw is None:
        return None
    fmt = sheet.formats.get(ref, 'General')
    if fmt == 'UNSUPPORTED' or is_date_format(fmt):
        raise ValueError('non-numeric cell format')
    value = Fraction(raw)
    binary = float(value)
    if not math.isfinite(binary) or (value and abs(binary) < sys.float_info.min):
        raise ValueError('operand outside finite Excel number range')
    if value.denominator == 1 and value != binary:
        raise ValueError('integer operand loses precision at the Double boundary')
    return value


def _numeric_case(book, node, actual, remaining):
    rows = node['rows']
    if node['kind'] not in OPERATIONS or node['product'] or len(rows) > min(MAX_CASE_OPERANDS, remaining):
        return None
    if any(len(row['cells']) != 1 or row['filters'] or row['numerator'] is not None for row in rows):
        return None
    if isinstance(actual, bool) or not isinstance(actual, (int, float, Fraction)):
        return None
    try:
        operands = [_operand(book, row['cells'][0]) for row in rows]
        expected = recipes.evaluate(book, node, decoded=True)
        if expected is None:
            return None
        binary = float(expected)
        if not math.isfinite(binary) or (expected and abs(binary) < sys.float_info.min):
            return None
        if actual != (binary if isinstance(actual, float) else expected):
            return None
    except (KeyError, ValueError, OverflowError, ZeroDivisionError):
        return None
    return operands, expected


def collect(bindings, books):
    """Use already verified table/chart grids and explicit source-cell recipes.

    Count omitted top-level aggregate cells separately. Other report work (including
    text, formulas, ranking decisions and share ratios) is outside this comparison.
    """
    cases = []
    skipped = 0
    remaining = MAX_TOTAL_OPERANDS
    for binding in bindings:
        recipe = binding['recipe']
        if binding['kind'] not in ('table', 'chart') or not isinstance(recipe, list):
            continue
        source = binding['source']
        book = books.get(source.path) if source is not None and Path(source.path).suffix.lower() in EXCEL_SUFFIXES else None
        for row, (expressions, cells) in enumerate(zip(recipe, binding['grid']), 1):
            if row == 1 and binding.get('skip'):
                continue
            for column, (expr, cell) in enumerate(zip(expressions, cells), 1):
                if not isinstance(expr, dict) or expr.get('op') != 'aggregate':
                    continue
                result = None if book is None or len(cases) == MAX_CASES else _numeric_case(book, expr, cell[0], remaining)
                if result is None:
                    skipped += 1
                    continue
                operands, expected = result
                remaining -= len(operands)
                cases.append({
                    'id': str(len(cases) + 1),
                    'operation': expr['kind'],
                    'operands': operands,
                    'expected': expected,
                    'label': f"{binding['key']} [{binding['part'].rsplit('/', 1)[-1]} #{binding['shape']}, {row},{column}]",
                })
    return VbaCases(cases, skipped)
