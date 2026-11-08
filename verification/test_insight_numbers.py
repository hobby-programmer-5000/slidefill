"""Adversarial numeric wording checks with exact, independent decimal operands."""

from fractions import Fraction
import json
import unittest
from unittest import mock

import facts
import insight
from numeric_text import MAX_TOKEN_LENGTH, figure_tokens, number_identity, unbacked_numbers


class ExactInsightNumbers(unittest.TestCase):
    @staticmethod
    def check(stated, allowed):
        pack = (facts.Fact("value", "Sales " + allowed, (allowed,)),)
        return insight.check({"statement": "Sales " + stated + ".", "fact_ids": ["value"]}, pack, "en")

    def test_adjacent_large_integers_and_long_decimals_do_not_collide(self):
        for allowed, wrong in (("1000001", "1000002"),
                               ("9007199254740993", "9007199254740994"),
                               ("123456789012345678901234567890", "123456789012345678901234567891"),
                               ("1.000001", "1.000002"),
                               ("0.123456789012345678901", "0.123456789012345678902")):
            with self.subTest(allowed=allowed):
                self.assertEqual(self.check(allowed, allowed), [])
                self.assertIn("number", self.check(wrong, allowed))

    def test_finite_scientific_literals_do_not_overflow_or_underflow(self):
        for allowed, equivalent, wrong in (("1e400", "10e399", "2e400"),
                                           ("1e-400", "10e-401", "2e-400")):
            with self.subTest(allowed=allowed):
                self.assertEqual(self.check(equivalent, allowed), [])
                self.assertIn("number", self.check(wrong, allowed))

    def test_grouping_trailing_zeroes_and_scientific_scale_are_equivalent(self):
        for allowed, stated in (("11.0", "11"), ("1,234.500", "1234.5"),
                                ("12345.678", "12345678e-3"), ("0.000", "0")):
            with self.subTest(allowed=allowed, stated=stated):
                self.assertEqual(self.check(stated, allowed), [])

    def test_decimal_rescaling_relations_and_one_unit_mutations(self):
        # Separate Fraction arithmetic establishes each relation. No formatter or
        # decoder from the wording guard computes these expected values.
        for coefficient in (1, 7, 199, 1234567, 9007199254740993, 10**90 + 37):
            for places in (1, 2, 9, 30):
                whole, remainder = divmod(coefficient, 10**places)
                decimal = f"{whole}.{remainder:0{places}d}"
                scaled = f"{coefficient}e-{places}"
                mutated = f"{coefficient + 1}e-{places}"
                self.assertEqual(Fraction(decimal), Fraction(scaled))
                self.assertNotEqual(Fraction(decimal), Fraction(mutated))
                self.assertEqual(self.check(scaled, decimal), [], (scaled, decimal))
                self.assertIn("number", self.check(mutated, decimal), (mutated, decimal))

    def test_malformed_or_unbounded_literals_fail_even_if_a_fact_repeats_them(self):
        for raw in ("1,23", "1,,000", "1.2.3", "1e", "1e+", "1e999999999", "1e4097",
                    "1" * (MAX_TOKEN_LENGTH + 1)):
            with self.subTest(raw=raw[:40]):
                self.assertIn("number", self.check(raw, raw))

    def test_missing_citation_repair_uses_exact_identity(self):
        pack = (facts.Fact("a", "Sales 1000001", ("1000001",)),
                facts.Fact("b", "Sales 1000002", ("1000002",)))
        self.assertEqual([fact.id for fact in insight._with_missing_citations([], pack, "Sales 1000002.")], ["b"])
        self.assertEqual(insight._with_missing_citations([], pack, "Sales 1000003."), [])

    def test_model_reply_with_a_changed_low_digit_is_rejected(self):
        pack = (facts.Fact("value", "Sales 1000001", ("1000001",)),)
        reply = {"response": json.dumps({"insights": [
            {"statement": "Sales 1000002.", "fact_ids": ["value"]},
            {"statement": "Sales 1000001.", "fact_ids": ["value"]}]})}
        with mock.patch.object(insight, "post", return_value=reply):
            result = insight.make_candidates(pack, "en", runs=1)
        self.assertEqual(result.rejected, 1)
        self.assertEqual([candidate.statement for candidate in result.candidates], ["Sales 1000001."])

    def test_dates_do_not_erase_part_of_a_numeric_literal(self):
        self.assertEqual(figure_tokens("Sales 1e2025."), ("1e2025",))
        self.assertEqual(figure_tokens("Sales in Dec 2025 was 11."), ("11.",))
        self.assertEqual(figure_tokens("2025년 12월 매출 11입니다."), ("11",))
        self.assertEqual(unbacked_numbers("Sales 1e2025.", ("1e2025",)), ())

    def test_signed_identity_is_available_without_changing_wording_semantics(self):
        self.assertNotEqual(number_identity("-11"), number_identity("11"))
        self.assertEqual(number_identity("-0"), number_identity("0"))
        self.assertEqual(unbacked_numbers("Sales -11.", ("11",)), ())
        self.assertEqual(unbacked_numbers("Sales -11.", ("11",), magnitude=False), ("-11.",))


if __name__ == "__main__":
    unittest.main()
