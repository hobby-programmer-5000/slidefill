import hashlib
import json
import unittest
from decimal import Decimal
from fractions import Fraction
from types import SimpleNamespace
from unittest.mock import Mock, patch

import native_vba


class NativeVbaTest(unittest.TestCase):
    def test_bundled_engine_matches_the_reviewed_digest(self):
        data = native_vba.bundled_engine_path().read_bytes()
        self.assertEqual(hashlib.sha256(data).hexdigest(), native_vba.ENGINE_SHA256)

    def test_only_typed_operands_reach_the_host(self):
        cases = [
            {"id": "private label", "operation": "sum", "operands": [Fraction(1, 10), None, 0],
             "label": "private sheet", "path": "private.xlsx"},
            {"id": "empty", "operation": "count", "operands": []},
        ]
        ids, payload = native_vba._encode_cases(cases)
        self.assertEqual(ids, ["private label", "empty"])
        self.assertEqual(payload, [
            {"operation": "sum", "digits": 2, "operands": ["9a9999999999b93f", None, "0000000000000000"]},
            {"operation": "count", "digits": 2, "operands": [None]},
        ])

    def test_invalid_inputs_stop_before_starting_excel(self):
        base = {"id": "case", "operation": "sum", "operands": [1]}
        invalid_operands = [True, "1", float("nan"), Decimal("Infinity"), 2**53 + 1,
                            Decimal(2**53 + 1), Fraction(2**53 + 1), Fraction(1, 10**400)]
        with patch.object(native_vba.subprocess, "run") as start:
            for value in invalid_operands:
                with self.subTest(value=value):
                    self.assert_invalid([{**base, "operands": [value]}])
            self.assert_invalid([{**base, "operation": "untrusted.xlsm!Macro"}])
            self.assert_invalid([{**base, "digits": True}])
            self.assert_invalid([base, base])
            self.assert_invalid([base], digest="invalid")
            self.assert_invalid([base], path="user-workbook.xlsx")
            with patch.object(native_vba, "MAX_BATCH_OPERANDS", 1):
                self.assert_invalid([base, {**base, "id": "second"}])
            start.assert_not_called()

    def assert_invalid(self, cases, digest=native_vba.ENGINE_SHA256, path="engine.xlsm"):
        self.assertEqual(native_vba.run(path, digest, cases)["status"], "invalid")

    def test_typed_response_retains_raw_double_and_excel_error(self):
        response = {"status": "executed", "cases": [
            {"kind": "number", "double_hex": "343333333333d33f"},
            {"kind": "error", "error_code": 2007},
        ]}
        result = native_vba._response(response_file(response), ["sum", "mean"], 0)
        self.assertEqual(result, {"status": "executed", "cases": [
            {"id": "sum", "kind": "number", "double_hex": "343333333333D33F",
             "value": 0.30000000000000004, "decimal": "0.30000000000000004"},
            {"id": "mean", "kind": "error", "error_code": 2007},
        ]})

    def test_malformed_response_cannot_report_a_successful_comparison(self):
        invalid_responses = [
            ({"status": "executed", "cases": []}, 0),
            ({"status": "executed", "cases": [{"kind": "number", "double_hex": "000000000000F87F"}]}, 0),
            ({"status": "executed", "cases": [{"kind": "error", "error_code": 9999}]}, 0),
            ({"status": "executed", "cases": [{"kind": "number", "double_hex": "000000000000F03F"}]}, 1),
            ({"status": "blocked", "detail": "Office policy"}, 0),
        ]
        for response, code in invalid_responses:
            with self.subTest(response=response, code=code), self.assertRaises(ValueError):
                native_vba._response(response_file(response), ["case"], code)

    def test_oversized_host_response_is_rejected_before_reading(self):
        path = response_file({}, size=native_vba.MAX_RESPONSE_BYTES + 1)
        with self.assertRaises(ValueError):
            native_vba._response(path, ["case"], 0)
        path.read_text.assert_not_called()


def response_file(data, *, size=None):
    body = json.dumps(data)
    path = Mock()
    path.stat.return_value = SimpleNamespace(st_size=len(body) if size is None else size)
    path.read_text.return_value = body
    return path


if __name__ == "__main__":
    unittest.main()
