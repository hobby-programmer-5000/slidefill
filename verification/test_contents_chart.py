"""Canvas behavior checks with synthetic rows, without creating a Tk window."""
from fractions import Fraction
import math
import unittest

from contents import draw_chart, SERIES_COLORS


class Canvas:
    def __init__(self):
        self.items = []

    def record(self, kind, coordinates, options):
        self.items.append((kind, coordinates, options))
        return len(self.items)

    def create_line(self, *coordinates, **options):
        return self.record("line", coordinates, options)

    def create_rectangle(self, *coordinates, **options):
        return self.record("rectangle", coordinates, options)

    def create_oval(self, *coordinates, **options):
        return self.record("oval", coordinates, options)

    def create_text(self, *coordinates, **options):
        return self.record("text", coordinates, options)

    def bbox(self, item):
        return (0, 0, 100, 20)


def chart(values, bars=False):
    rows = [[(None, ""), (None, "Sales")]]
    rows += [[(None, str(i)), (value, str(value))] for i, value in enumerate(values)]
    canvas = Canvas()
    draw_chart(canvas, 0, 0, 500, 300, rows, bars, 10, "black")
    return canvas


def plotted(canvas, kind):
    return [coords for name, coords, options in canvas.items
            if name == kind and options.get("fill") == SERIES_COLORS[0]
            and options.get("width") == 2] if kind == "line" else [
        coords for name, coords, options in canvas.items
        if name == kind and options.get("fill") == SERIES_COLORS[0]
        and (kind != "rectangle" or coords[1] < 270)]


class ContentsChartTests(unittest.TestCase):
    def test_fraction_series_is_visible(self):
        lines = plotted(chart([Fraction(1, 3), Fraction(2, 3)]), "line")
        self.assertEqual(len(lines), 1)
        self.assertEqual(lines[0][1::2], (135.0, 4.0))

    def test_missing_slot_breaks_line(self):
        self.assertEqual(plotted(chart([1, None, 2]), "line"), [])

    def test_contiguous_runs_stay_separate(self):
        self.assertEqual(len(plotted(chart([1, 2, None, 3, 4]), "line")), 2)

    def test_isolated_valid_points_are_visible(self):
        self.assertEqual(len(plotted(chart([1, None, 2]), "oval")), 2)

    def test_bars_start_at_zero_for_both_signs(self):
        rectangles = plotted(chart([-1, 1], bars=True), "rectangle")
        self.assertEqual(len(rectangles), 2)
        self.assertEqual(rectangles[0][1::2], (135.0, 266.0))
        self.assertEqual(rectangles[1][1::2], (4.0, 135.0))

    def test_bool_and_nonfinite_are_gaps(self):
        for bad in (True, False, float("inf"), float("-inf"), float("nan")):
            with self.subTest(bad=bad):
                canvas = chart([1, bad, 2])
                self.assertEqual(plotted(canvas, "line"), [])
                self.assertEqual(len(plotted(canvas, "oval")), 2)

    def test_huge_values_have_finite_bounded_geometry(self):
        values = [Fraction(-(10**400)), Fraction(10**400)]
        for bars in (False, True):
            canvas = chart(values, bars)
            for _kind, coords, _options in canvas.items:
                self.assertTrue(all(math.isfinite(value) for value in coords))
                self.assertTrue(all(-10 <= value <= 500 for value in coords))

    def test_positive_scaling_preserves_geometry(self):
        values = [Fraction(-2, 7), Fraction(1, 3), Fraction(8, 9)]
        for bars, kind in ((False, "line"), (True, "rectangle")):
            self.assertEqual(plotted(chart(values, bars), kind),
                             plotted(chart([v * 10**400 for v in values], bars), kind))

    def test_zero_and_constant_series_are_finite(self):
        for value in (0, Fraction(1, 3), Fraction(-1, 3)):
            canvas = chart([value, value])
            self.assertTrue(plotted(canvas, "line"))
            self.assertTrue(all(math.isfinite(v) for _, coords, _ in canvas.items for v in coords))


if __name__ == "__main__":
    unittest.main()
