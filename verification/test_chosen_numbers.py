"""The saved-choice and editor boundaries enforce the same narrow number check."""
import datetime as dt
import hashlib
from pathlib import Path
import tempfile
from types import SimpleNamespace
import unittest
from unittest.mock import Mock, patch

from openpyxl import Workbook
from pptx import Presentation
from pptx.util import Inches

import app
import core
import facts


class ChosenNumberTests(unittest.TestCase):
    def request(self):
        return core.InsightRequest("insight:Sales!A,B", "en",
            (facts.Fact("value", "Sales 1000001", ("1000001",)),),
            "synthetic-fingerprint", "Calculated Sales 1000001.")

    def resolve(self, text):
        request = self.request()
        run = object.__new__(core._Run)
        run.source = SimpleNamespace(insight_request=lambda *_: (request, "synthetic", []))
        run.words_of_template = lambda: ""
        run.requests = {}
        run.insights = {request.key.casefold(): {"fingerprint": request.fingerprint, "text": text}}
        run.capture_text_input = Mock()
        item = core.Item("text", request.key)
        return run.resolve_insight(request.key, item), item, request

    def test_tampered_saved_number_uses_fallback_and_warning(self):
        for text in ("Sales 1000002.", "Sales 1,00,001.", "Sales 1e99999.", 1000002):
            with self.subTest(text=text):
                result, item, request = self.resolve(text)
                self.assertEqual(result, request.fallback)
                self.assertEqual([issue.code for issue in item.issues], ["insight_numbers"])

    def test_equivalent_numeric_spelling_is_preserved(self):
        for text in ("Sales 1,000,001.", "Sales 1000001.0.", "Sales 1.000001e6.", "A sentence without a figure."):
            with self.subTest(text=text):
                result, item, _ = self.resolve(text)
                self.assertEqual(result, text)
                self.assertEqual(item.issues, [])

    def test_editor_rejection_preserves_previous_choice_and_does_not_save(self):
        request = self.request()
        previous = {request.key.casefold(): {"fingerprint": request.fingerprint, "text": "Sales 1000001."}}
        fake = SimpleNamespace(insights=previous.copy(), insight_window=None, root=None,
                               t=lambda key: key, preview=Mock())
        with patch.object(app.messagebox, "showwarning") as warning, patch.object(core, "save_settings") as save:
            self.assertFalse(app.App.apply_insight(fake, request, "Sales 1000002."))
        self.assertEqual(fake.insights, previous)
        save.assert_not_called()
        fake.preview.assert_not_called()
        warning.assert_called_once()

    def test_editor_stays_open_until_text_is_accepted(self):
        fake = SimpleNamespace(app=SimpleNamespace(apply_insight=Mock(return_value=False)),
            request=self.request(), chosen=SimpleNamespace(get=lambda *_: "Sales 1000002."),
            window=SimpleNamespace(destroy=Mock()))
        app.InsightWindow.apply(fake)
        fake.window.destroy.assert_not_called()
        fake.app.apply_insight.return_value = True
        app.InsightWindow.apply(fake)
        fake.window.destroy.assert_called_once()

    def test_valid_choice_and_clear_still_save_and_preview(self):
        request = self.request()
        fake = SimpleNamespace(insights={}, insights_file="unused.json", auto=None,
            template_path=lambda: "", preview=Mock())
        with patch.object(core, "save_settings") as save:
            self.assertTrue(app.App.apply_insight(fake, request, "Sales 1000001."))
            self.assertIn(request.key.casefold(), fake.insights)
            self.assertTrue(app.App.apply_insight(fake, request, ""))
        self.assertEqual(fake.insights, {})
        self.assertEqual(save.call_count, 2)
        self.assertEqual(fake.preview.call_count, 2)


class ChosenNumberReports(unittest.TestCase):
    """Real source parsing, request fingerprints, verification and PPTX writing."""

    KEY = "insight:Daily!A,B"
    PREFIX = "Sales report: "

    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.folder = Path(self.temp.name)
        self.book = self.folder / "sales.xlsx"
        self.template = self.folder / "template.pptx"
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Daily"
        sheet.append(["Date", "Sales"])
        for day in range(5, 13):
            sheet.append([dt.datetime(2026, 1, day), 1000001])
            sheet.cell(sheet.max_row, 1).number_format = "yyyy-mm-dd"
            sheet.cell(sheet.max_row, 2).number_format = "0"
        workbook.save(self.book)
        workbook.close()
        presentation = Presentation()
        slide = presentation.slides.add_slide(presentation.slide_layouts[6])
        slide.shapes.add_textbox(Inches(1), Inches(1), Inches(8), Inches(4)).text_frame.text = self.PREFIX + "{{" + self.KEY + "}}"
        presentation.save(self.template)
        self.original_hashes = self.input_hashes()
        baseline = core.run(self.template, self.book)
        self.request = baseline.requests[self.KEY]
        self.assertEqual(self.request.lang, "en")
        self.assertIn("1,000,001", next(fact for fact in self.request.facts if fact.id == "value").numbers)

    def input_hashes(self):
        return tuple(hashlib.sha256(path.read_bytes()).hexdigest() for path in (self.book, self.template))

    def preview_and_save(self, text, name):
        choices = {self.KEY.casefold(): {"fingerprint": self.request.fingerprint, "text": text}}
        preview = core.run(self.template, self.book, insights=choices)
        output = self.folder / name
        saved = core.run(self.template, self.book, output, insights=choices)
        visible = Presentation(output).slides[0].shapes[0].text_frame.text
        self.assertFalse(preview.blocked)
        self.assertFalse(saved.blocked)
        self.assertEqual(self.input_hashes(), self.original_hashes)
        return (next(item for item in report.items if item.key == self.KEY) for report in (preview, saved)), visible

    def test_changed_low_digit_falls_back_in_preview_and_saved_report(self):
        items, visible = self.preview_and_save("Sales 1000002.", "rejected.pptx")
        for item in items:
            self.assertEqual(item.value, self.request.fallback)
            self.assertEqual(item.status, "warn")
            self.assertIn("insight_numbers", [issue.code for issue in item.issues])
        self.assertEqual(visible, self.PREFIX + self.request.fallback)
        self.assertNotIn("1000002", visible)

    def test_exact_equivalent_spelling_survives_preview_and_saved_report(self):
        for index, spelling in enumerate(("1,000,001.000", "1.000001e6")):
            with self.subTest(spelling=spelling):
                text = "Sales " + spelling + "."
                items, visible = self.preview_and_save(text, f"accepted-{index}.pptx")
                for item in items:
                    self.assertEqual(item.value, text)
                    self.assertEqual(item.status, "ok")
                    self.assertNotIn("insight_numbers", [issue.code for issue in item.issues])
                self.assertEqual(visible, self.PREFIX + text)


if __name__ == "__main__":
    unittest.main()
