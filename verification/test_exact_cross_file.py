"""Cross-file direction contracts with exact operands and transformed series."""
import datetime as dt
from dataclasses import replace
from fractions import Fraction
from pathlib import Path
from tempfile import TemporaryDirectory
from types import SimpleNamespace
import unittest

import core
import facts


class ExactCrossFile(unittest.TestCase):
    @staticmethod
    def monthly(last=Fraction(100), scale=Fraction(1)):
        return [facts.Point(dt.datetime(2025 + m // 12, m % 12 + 1, 1),
                            (last if m == 13 else Fraction(100)) * scale)
                for m in range(14)]

    @staticmethod
    def request(points, lang="en", rate=False):
        pack = (facts.rate_pack if rate else facts.fact_pack)(points, "Sales", facts.plain_number, lang)
        period = facts.point_text(points[-1].date, points, lang)
        return SimpleNamespace(facts=(facts.Fact("subject", "", (), "Sales", period), *pack))

    def test_exact_year_and_previous_changes_survive_display_rounding(self):
        # Both rounded facts say 0.0%; the exact positive ratio is 1 / 10**12.
        request = self.request(self.monthly(Fraction(100) + Fraction(1, 10**10)))
        found = {fact.id: fact for fact in request.facts}
        self.assertIn("0.0%", found["year"].text)
        self.assertIn("0.0%", found["prev"].text)
        self.assertEqual(found["year"].relative_change, Fraction(1, 10**12))
        self.assertEqual(found["prev"].relative_change, Fraction(1, 10**12))
        self.assertEqual(core._lead_change(request, "en"), (Fraction(1, 10**10), "Feb 2026", "year"))

    def test_display_text_cannot_change_calculated_cross_file_direction(self):
        request = self.request(self.monthly(Fraction(125)))
        changed = SimpleNamespace(facts=tuple(replace(fact, text="invented -99.9%")
                                             if fact.id == "year" else fact for fact in request.facts))
        self.assertEqual(core._lead_change(changed, "en"), (Fraction(25), "Feb 2026", "year"))

    def test_legacy_prose_and_rate_points_do_not_supply_relative_changes(self):
        legacy = SimpleNamespace(facts=(facts.Fact("subject", "", (), "Sales", "Feb 2026"),
                                        facts.Fact("year", "+25.0%", ("25.0",))))
        self.assertEqual(core._lead_change(legacy, "en"), (None, "Feb 2026", "year"))
        rates = self.request(self.monthly(Fraction(101)), rate=True)
        self.assertTrue(any(fact.id == "read_prev_points" for fact in rates.facts))
        self.assertIsNone(core._lead_change(rates, "en")[0])

    def test_year_change_stays_preferred_and_previous_is_explicit_fallback(self):
        points = self.monthly(Fraction(125))
        points[-2] = replace(points[-2], value=Fraction(250))
        request = self.request(points)
        self.assertEqual(core._lead_change(request, "en")[0], Fraction(25))
        without_year = SimpleNamespace(facts=tuple(fact for fact in request.facts if fact.id != "year"))
        self.assertEqual(core._lead_change(without_year, "en"), (Fraction(-50), "Feb 2026", "prev"))

    def test_zero_base_withholds_but_true_zero_change_remains_available(self):
        request = self.request(self.monthly())
        self.assertEqual(core._lead_change(request, "en")[0], Fraction(0))
        points = [facts.Point(dt.datetime(2026, 1, day), value)
                  for day, value in ((1, 0), (2, 0), (3, 10))]
        self.assertIsNone(core._lead_change(self.request(points), "en")[0])

    def test_positive_scaling_preserves_the_exact_ratio_and_direction(self):
        # Preconditions: positive scale, nonzero base, dates and selection fixed.
        for delta in (Fraction(1, 10**20), Fraction(-1, 10**20), Fraction(0), Fraction(1, 3)):
            expected = delta  # base is 100, and _lead_change returns percentage units.
            for scale in (Fraction(1, 7), Fraction(37), Fraction(10**40)):
                request = self.request(self.monthly(Fraction(100) + delta, scale))
                self.assertEqual(core._lead_change(request, "en")[0], expected)

    def test_nonzero_subdisplay_changes_have_truthful_bilingual_directions(self):
        entries = [("A", Fraction(1, 10**12), "Feb 2026", "year"),
                   ("B", Fraction(-1, 10**12), "Feb 2026", "year")]
        english = facts.cross_file_readings(entries, "en")[0].text
        self.assertIn("A up less than 0.05%", english)
        self.assertIn("B down less than 0.05%", english)
        self.assertIn("different directions", english)
        korean = facts.cross_file_readings(entries, "ko")[0].text
        self.assertIn("A 0.05% 미만 증가", korean)
        self.assertIn("B 0.05% 미만 감소", korean)
        self.assertNotIn("0.0%", english + korean)

    def test_primary_sentence_does_not_call_nonzero_subdisplay_changes_unchanged(self):
        for delta, english, korean in ((Fraction(1, 10**10), "up less than 0.05%", "0.05% 미만 증가"),
                                       (Fraction(-1, 10**10), "down less than 0.05%", "0.05% 미만 감소")):
            points = self.monthly(Fraction(100) + delta)
            self.assertIn(english, facts.sentence(points, "Sales", facts.plain_number, "en"))
            self.assertIn(korean, facts.sentence(points, "Sales", facts.plain_number, "ko"))

    def test_cross_file_rounds_half_away_once_at_the_declared_precision(self):
        entries = [("A", Fraction(1, 4), "Feb 2026", "prev"),
                   ("B", Fraction(-1, 20), "Feb 2026", "prev")]
        text = facts.cross_file_readings(entries, "en")[0].text
        self.assertIn("A up 0.3%", text)
        self.assertIn("B down 0.1%", text)
        self.assertNotIn("less than", text)

    def test_incompatible_periods_or_bases_still_withhold_direction(self):
        for period, basis, expected in (("Jan 2026", "year", "cross_periods"),
                                        ("Feb 2026", "prev", "cross_basis")):
            found = facts.cross_file_readings([("A", Fraction(1), "Feb 2026", "year"),
                                              ("B", Fraction(-1), period, basis)], "en")
            self.assertEqual(found[0].id, expected)
            self.assertNotIn("A up", found[0].text)

    def test_fingerprint_changes_when_exact_change_changes_below_display_precision(self):
        one = facts.fact_pack(self.monthly(Fraction(100) + Fraction(1, 10**10)), "Sales", facts.plain_number, "en")
        two = facts.fact_pack(self.monthly(Fraction(100) + Fraction(2, 10**10)), "Sales", facts.plain_number, "en")
        self.assertEqual([fact.text for fact in one], [fact.text for fact in two])
        self.assertNotEqual(facts.fingerprint(one, "en"), facts.fingerprint(two, "en"))

    def test_fresh_csv_sources_keep_tiny_opposite_directions_end_to_end(self):
        # Freshly read transformed inputs, not a changed-input hash rejection.
        with TemporaryDirectory() as directory:
            paths = []
            for name, last in (("up", "100.00000001"), ("down", "99.99999999")):
                path = Path(directory) / f"{name}.csv"
                body = "Date,Sales\n" + "".join(
                    f"{2025 + m // 12}-{m % 12 + 1:02d}-01,{last if m == 13 else '100'}\n"
                    for m in range(14))
                path.write_text(body, encoding="utf-8")
                paths.append(str(path))
            source = core.Sources(paths)
            text, _where, _issues = source.text_value("reading:@all@@0@up!A,B|@1@down!A,B")
            self.assertIn("up less than 0.05%", text)
            self.assertIn("down less than 0.05%", text)
            self.assertIn("different directions", text)


if __name__ == "__main__":
    unittest.main()
