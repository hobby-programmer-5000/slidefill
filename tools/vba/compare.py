"""Developer-only native VBA comparison; does not change SlideFill's engine."""
from pathlib import Path
import argparse
from fractions import Fraction
import hashlib
import io
import json
import math
import re
import struct
import sys
from zipfile import ZipFile

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT))
import core
from openpyxl import Workbook
from verification.kernel import round_rational


CASES = (
    ("decimal_sum", "sum", [0.1, 0.2], 2, "0.0000000000000000"),
    ("decimal_variant", "decimal_sum", [0.1, 0.2], 2, "0.0000000000000000"),
    ("cancellation", "sum", [10**16, 1, -(10**16)], 0, "0"),
    ("large_literal", "sum", [9007199254740993, 0], 0, "0"),
    ("negative_mean", "mean", [-1, -1, 0], 3, "0.000"),
    ("positive_mean", "mean", [1, 2, 2], 3, "0.000"),
    ("count_blanks", "count", [1, None, 3], 0, "0"),
    ("empty_sum", "sum", [], 0, "0"),
    ("empty_mean", "mean", [], 2, "0.00"),
    ("tiny_growth", "growth", [100.00000001, 100], 12, "0.000000000000%"),
    ("negative_base", "growth", [-5, -10], 2, "0.00%"),
    ("zero_base", "growth", [1, 0], 2, "0.00%"),
    ("point_difference", "points", [10.25, 10], 2, "0.00"),
    ("worksheet_tie", "round", [2.5], 0, "0"),
    ("worksheet_negative_tie", "round", [-2.5], 0, "0"),
    ("banker_tie_probe", "vba_round_probe", [2.5], 0, "0"),
    ("decimal_round", "round", [2.675], 2, "0.00"),
    ("negative_precision", "round", [125], -1, "0"),
    ("bool_refused", "sum", [True, 1], 0, "0"),
    ("text_refused", "sum", ["2", 1], 0, "0"),
    ("unknown_refused", "unknown", [1], 0, "0"),
)


def token(value):
    if value is None:
        return {"type": "blank"}
    if isinstance(value, bool):
        return {"type": "bool", "value": value}
    if isinstance(value, (int, float)):
        if isinstance(value, float) and not math.isfinite(value):
            raise ValueError("nonfinite input")
        return {"type": "number", "decimal": str(value)}
    if isinstance(value, str):
        return {"type": "text", "value": value}
    raise ValueError("unsupported input type")


def exact_result(operation, operands, digits=2):
    """Independent mathematical policy, including half-away for the banker probe."""
    if operation not in {"sum", "decimal_sum", "mean", "count", "growth", "points", "round", "vba_round_probe"}:
        return {"kind": "error", "error_code": 2015}
    if any(item["type"] not in {"blank", "number"} for item in operands):
        return {"kind": "error", "error_code": 2015}
    values = [Fraction(item["decimal"]) for item in operands if item["type"] == "number"]
    if operation in {"sum", "decimal_sum"}:
        value = sum(values, Fraction(0))
    elif operation == "count":
        value = Fraction(len(values))
    elif operation == "mean":
        if not values:
            return {"kind": "error", "error_code": 2007}
        value = sum(values, Fraction(0)) / len(values)
    elif operation in {"growth", "points"}:
        if len(values) != 2:
            return {"kind": "error", "error_code": 2015}
        if operation == "growth" and not values[1]:
            return {"kind": "error", "error_code": 2007}
        value = values[0] - values[1]
        if operation == "growth":
            value /= abs(values[1])
    else:
        if len(values) != 1 or type(digits) is not int or not -15 <= digits <= 15 or (operation == "vba_round_probe" and digits < 0):
            return {"kind": "error", "error_code": 2015}
        factor = 10 ** max(0, -digits)
        value = Fraction(round_rational(values[0] / factor, max(0, digits))) * factor
    return {"kind": "number", "rational": str(value), "nearest_double": float(value).hex()}


def load_cases(path):
    """Use the application's real workbook reader and explicit source range."""
    source = core.ExcelSource(str(path))
    rows, where, issues = source.range_cells(f"Cases!A1:G{len(CASES) + 1}")
    if issues:
        raise ValueError("fixture read issues: " + ",".join(i.code for i in issues))
    if [cell[0] for cell in rows[0]] != ["id", "operation", "v1", "v2", "v3", "digits", "format"]:
        raise ValueError("comparison fixture header differs")
    cases = []
    for row in rows[1:]:
        raw = [cell[0] for cell in row]
        if raw[0] is None:
            continue
        identifier, operation, *rest = raw
        if not isinstance(identifier, str) or any(c["id"] == identifier for c in cases):
            raise ValueError("non-string or duplicate case id")
        operands = [token(value) for value in rest[:3]]
        digits, fmt = rest[3:]
        if type(digits) is not int or not isinstance(operation, str) or not isinstance(fmt, str):
            raise ValueError("invalid comparison metadata")
        cases.append({"id": identifier, "operation": operation, "operands": operands,
                      "digits": digits, "format": fmt,
                      "python": exact_result(operation, operands, digits)})
    return {"schema": 1, "source": str(Path(path).resolve()), "source_range": where,
            "source_sha256": hashlib.sha256(Path(path).read_bytes()).hexdigest(),
            "module_source_sha256": hashlib.sha256(Path(__file__).with_name("SlideFillOracle.bas").read_bytes()).hexdigest(),
            "cases": cases}


def prepare(folder):
    folder.mkdir(parents=True, exist_ok=False)
    path = folder / "inputs.xlsx"
    book = Workbook()
    sheet = book.active
    sheet.title = "Cases"
    sheet.append(["id", "operation", "v1", "v2", "v3", "digits", "format"])
    for identifier, operation, values, digits, fmt in CASES:
        sheet.append([identifier, operation, *(values + [None] * (3 - len(values))), digits, fmt])
    for column, width in {"A": 28, "B": 24, "C": 26, "D": 26, "E": 26, "F": 10, "G": 28}.items():
        sheet.column_dimensions[column].width = width
    sheet.freeze_panes = "C2"
    book.save(path)
    book.close()
    # openpyxl's numeric writer can round a 16-digit integer before either
    # engine sees it. Preserve the authored integer XML sentinels in this new
    # synthetic fixture so Excel decoding and app decoding can be compared.
    with ZipFile(path) as archive:
        members = [(info, archive.read(info)) for info in archive.infolist()]
    rewritten = io.BytesIO()
    with ZipFile(rewritten, "w") as archive:
        for info, data in members:
            if info.filename == "xl/worksheets/sheet1.xml":
                sheet_xml = data.decode("utf-8")
                for row, (_, _, values, _, _) in enumerate(CASES, 2):
                    for column, value in zip("CDE", values):
                        if type(value) is int:
                            pattern = rf'(<c\b[^>]*\br="{column}{row}"[^>]*>\s*<v>)[^<]*(</v>)'
                            sheet_xml, count = re.subn(pattern, lambda m: m[1] + str(value) + m[2], sheet_xml)
                            if count != 1:
                                raise ValueError("integer sentinel cell missing from generated fixture")
                data = sheet_xml.encode("utf-8")
            archive.writestr(info, data)
    path.write_bytes(rewritten.getvalue())
    payload = load_cases(path)
    (folder / "app-input.json").write_text(json.dumps(payload, indent=2) + "\n", encoding="utf8")
    print(f"Prepared {len(payload['cases'])} actual-reader cases in {folder}")


def validate_result(result):
    if result.get("kind") == "error":
        if type(result.get("error_code")) is not int:
            raise ValueError("missing native error identity")
    elif result.get("kind") == "number":
        decimal, bits = result.get("decimal"), result.get("double_hex")
        if not isinstance(decimal, str) or len(decimal) > 128 or not isinstance(bits, str):
            raise ValueError("missing native numeric identity")
        number = float(decimal)
        if not math.isfinite(number) or struct.pack("<d", number).hex().lower() != bits.lower():
            raise ValueError("native decimal/Double bit identity differs")
    else:
        raise ValueError("unsupported or missing native result")


def summarize(app_path, native_path, output):
    app_input = json.loads(app_path.read_text(encoding="utf-8-sig"))
    native = json.loads(native_path.read_text(encoding="utf-8-sig"))
    if native.get("status") not in {"executed", "ok"}:
        raise ValueError("VBA execution has not completed: " + str(native.get("status")))
    rows = native["cases"]
    expected = {case["id"]: case for case in app_input["cases"]}
    if not expected or len(expected) != len(app_input["cases"]) or len(rows) != len(expected) or {row["id"] for row in rows} != set(expected):
        raise ValueError("native comparison case coverage differs")
    result = []
    for row in rows:
        python = expected[row["id"]]["python"]
        direct, called = row["direct"], row["app"]
        validate_result(direct)
        validate_result(called)
        bridge = direct["kind"] == called["kind"] and (
            direct.get("error_code") == called.get("error_code") if direct["kind"] == "error"
            else direct["double_hex"].lower() == called["double_hex"].lower())
        matches_math = direct["kind"] == python["kind"] and (
            direct.get("error_code") == python.get("error_code") if direct["kind"] == "error"
            else float(direct["decimal"]).hex() == python["nearest_double"])
        exact_difference = None
        if direct["kind"] == python["kind"] == "number":
            exact_difference = str(Fraction.from_float(float(direct["decimal"])) - Fraction(python["rational"]))
        result.append({"id": row["id"], "bridge_equal": bridge,
                       "matches_python_at_double_boundary": matches_math,
                       "native_binary_minus_python_exact": exact_difference,
                       "python": python, "direct": direct, "app": called})
    summary = {"status": "compared", "cases": result,
               "bridge_matches": sum(row["bridge_equal"] for row in result),
               "math_matches_at_double_boundary": sum(row["matches_python_at_double_boundary"] for row in result),
               "total": len(result), "limits": "Same VBA route agreement checks transport. Double equality does not certify exact rational arithmetic or business meaning."}
    with output.open("x", encoding="utf8") as stream:
        json.dump(summary, stream, indent=2)
        stream.write("\n")
    print(json.dumps({key: value for key, value in summary.items() if key != "cases"}, indent=2))
    return 0 if summary["bridge_matches"] == len(result) else 1


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    actions = parser.add_subparsers(dest="command", required=True)
    action = actions.add_parser("prepare"); action.add_argument("output", type=Path)
    action = actions.add_parser("summarize")
    action.add_argument("app_input", type=Path); action.add_argument("native_results", type=Path); action.add_argument("output", type=Path)
    args = parser.parse_args()
    try:
        if args.command == "prepare":
            prepare(args.output)
            return 0
        return summarize(args.app_input, args.native_results, args.output)
    except (ValueError, OSError, core.SlideFillError, core.IssueError) as exc:
        print(f"Cannot compare: {exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
