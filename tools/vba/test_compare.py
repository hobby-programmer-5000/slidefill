"""Fixed expected arithmetic and receipt-failure checks for the developer spike."""
from fractions import Fraction
import importlib.util
import json
from pathlib import Path
import tempfile
import struct
import unittest

spec = importlib.util.spec_from_file_location("vba_compare", Path(__file__).with_name("compare.py"))
compare = importlib.util.module_from_spec(spec)
spec.loader.exec_module(compare)


class ExactPolicy(unittest.TestCase):
    def result(self, operation, values, digits=2):
        return compare.exact_result(operation, [compare.token(v) for v in values], digits)

    def test_fixed_cancellation_and_long_integer_expectations(self):
        self.assertEqual(self.result("sum", [10**16, 1, -(10**16)])["rational"], "1")
        self.assertEqual(self.result("sum", [9007199254740993, 0])["rational"], "9007199254740993")

    def test_fixed_decimal_and_fraction_expectations(self):
        self.assertEqual(self.result("sum", [0.1, 0.2])["rational"], "3/10")
        self.assertEqual(self.result("mean", [-1, -1, 0])["rational"], "-2/3")
        self.assertEqual(self.result("points", [10.25, 10])["rational"], "1/4")

    def test_tiny_change_and_negative_base_keep_meaning(self):
        self.assertEqual(self.result("growth", [100.00000001, 100])["rational"], "1/10000000000")
        self.assertEqual(self.result("growth", [-5, -10])["rational"], "1/2")

    def test_half_away_reference_does_not_silently_adopt_vba_bankers(self):
        self.assertEqual(self.result("round", [2.5], 0)["rational"], "3")
        self.assertEqual(self.result("round", [-2.5], 0)["rational"], "-3")
        self.assertEqual(self.result("round", [125], -1)["rational"], "130")
        self.assertEqual(self.result("vba_round_probe", [2.5], 0)["rational"], "3")

    def test_blank_count_and_refusal_contract(self):
        self.assertEqual(self.result("count", [1, None, 3])["rational"], "2")
        for operation, values, code in (("sum", [True, 1], 2015), ("mean", [], 2007),
                                        ("sum", ["2"], 2015), ("growth", [1, 0], 2007),
                                        ("unknown", [1], 2015)):
            with self.subTest(operation=operation, values=values):
                self.assertEqual(self.result(operation, values), {"kind": "error", "error_code": code})
        with self.assertRaises(ValueError):
            compare.token(float("inf"))

    def test_positive_scale_preserves_relative_change(self):
        for scale in (1, 2, 7, 1000):
            self.assertEqual(self.result("growth", [12 * scale, 10 * scale])["rational"], "1/5")


class ComparisonEvidence(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.folder = Path(self.temp.name)
        self.app = self.folder / "app.json"
        self.native = self.folder / "native.json"
        self.output = self.folder / "comparison.json"

    def write(self, app_cases, native):
        self.app.write_text(json.dumps({"cases": app_cases}), encoding="utf8")
        self.native.write_text(json.dumps(native), encoding="utf8")

    def test_blocked_execution_never_becomes_pass(self):
        self.write([], {"status": "execution_blocked", "cases": []})
        with self.assertRaisesRegex(ValueError, "has not completed"):
            compare.summarize(self.app, self.native, self.output)
        self.assertFalse(self.output.exists())

    def test_missing_case_coverage_is_not_success(self):
        self.write([{"id": "one", "python": {}}], {"status": "executed", "cases": []})
        with self.assertRaisesRegex(ValueError, "coverage"):
            compare.summarize(self.app, self.native, self.output)

    def test_empty_native_run_is_not_success(self):
        self.write([], {"status": "executed", "cases": []})
        with self.assertRaisesRegex(ValueError, "coverage"):
            compare.summarize(self.app, self.native, self.output)

    def test_missing_or_inconsistent_numeric_identity_is_not_agreement(self):
        expected = {"kind": "number", "rational": "1", "nearest_double": float(1).hex()}
        for invalid in ({"kind": "number"}, {"kind": "number", "decimal": "1", "double_hex": "0000000000000000"},
                        {"kind": "number", "decimal": "nan", "double_hex": "000000000000F87F"}):
            with self.subTest(invalid=invalid):
                self.write([{"id": "one", "python": expected}], {"status": "executed", "cases": [
                    {"id": "one", "direct": invalid, "app": invalid}]})
                with self.assertRaises(ValueError):
                    compare.summarize(self.app, self.native, self.output)

    def test_same_vba_wrong_arithmetic_is_separate_from_bridge_match(self):
        expected = {"kind": "number", "rational": "1", "nearest_double": float(1).hex()}
        native_zero = {"kind": "number", "value": 0, "decimal": "0", "double_hex": "0000000000000000"}
        self.write([{"id": "cancel", "python": expected}],
                   {"status": "executed", "cases": [{"id": "cancel", "direct": native_zero, "app": native_zero}]})
        self.assertEqual(compare.summarize(self.app, self.native, self.output), 0)
        result = json.loads(self.output.read_text())
        self.assertEqual(result["bridge_matches"], 1)
        self.assertEqual(result["math_matches_at_double_boundary"], 0)
        self.assertEqual(result["cases"][0]["native_binary_minus_python_exact"], "-1")

    def test_allowed_double_boundary_does_not_hide_integer_information_loss(self):
        integer = 9007199254740993
        converted = float(integer)
        expected = {"kind": "number", "rational": str(integer), "nearest_double": converted.hex()}
        native = {"kind": "number", "decimal": repr(converted), "double_hex": struct.pack("<d", converted).hex()}
        self.write([{"id": "large", "python": expected}], {"status": "executed", "cases": [
            {"id": "large", "direct": native, "app": native}]})
        self.assertEqual(compare.summarize(self.app, self.native, self.output), 0)
        result = json.loads(self.output.read_text())
        self.assertEqual(result["math_matches_at_double_boundary"], 1)
        self.assertEqual(result["cases"][0]["native_binary_minus_python_exact"], "-1")

    def test_real_core_reader_preparation_and_no_overwrite(self):
        directory = self.folder / "prepared"
        compare.prepare(directory)
        payload = compare.load_cases(directory / "inputs.xlsx")
        self.assertEqual(len(payload["cases"]), 21)
        by_id = {case["id"]: case for case in payload["cases"]}
        self.assertEqual(by_id["large_literal"]["python"]["rational"], "9007199254740993")
        self.assertEqual(by_id["text_refused"]["operands"][0]["type"], "text")
        self.assertEqual(by_id["bool_refused"]["operands"][0]["type"], "bool")
        with self.assertRaises(FileExistsError):
            compare.prepare(directory)


if __name__ == "__main__":
    unittest.main()
