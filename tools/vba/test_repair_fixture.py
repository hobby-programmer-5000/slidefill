"""Package/type-preservation checks; sentinel bytes are not executable VBA."""
from pathlib import Path
import hashlib
import importlib.util
import tempfile
import unittest
from zipfile import ZipFile

from openpyxl import Workbook, load_workbook

spec = importlib.util.spec_from_file_location("repair_fixture", Path(__file__).with_name("repair_fixture.py"))
repair_fixture = importlib.util.module_from_spec(spec)
spec.loader.exec_module(repair_fixture)


class RepairFixture(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        folder = Path(self.temp.name)
        self.source, self.engine, self.output = folder / "source.xlsx", folder / "engine.xlsm", folder / "copy.xlsm"
        book = Workbook(); book.active.title = "Cases"
        book.active.append(["id", "operation", "v1", "v2", "v3", "digits", "format"])
        book.active.append(["text_case", "sum", "2", 1, None, 2, "0.000"])
        book.save(self.source)
        book.active["C2"] = 2; book.active["G2"] = 0
        book.create_sheet("_SlideFillMeta")["B2"] = hashlib.sha256(self.source.read_bytes()).hexdigest()
        book.save(self.engine); book.close()

    def test_requires_imported_project_before_copy(self):
        with self.assertRaisesRegex(ValueError, "imported VBA project"):
            repair_fixture.repair(self.source, self.engine, self.output)
        self.assertFalse(self.output.exists())

    def test_new_copy_preserves_project_bytes_and_literal_types(self):
        sentinel = b"package-test-only; not compiled or executable VBA"
        with ZipFile(self.engine, "a") as archive:
            archive.writestr("xl/vbaProject.bin", sentinel)
        original = self.engine.read_bytes()
        result = repair_fixture.repair(self.source, self.engine, self.output)
        self.assertTrue(result["vba_binary_unchanged"])
        self.assertEqual(self.engine.read_bytes(), original)
        with ZipFile(self.output) as archive:
            self.assertEqual(archive.read("xl/vbaProject.bin"), sentinel)
        copied = load_workbook(self.output, keep_vba=True)
        self.assertEqual(copied["Cases"]["C2"].value, "2")
        self.assertEqual(copied["Cases"]["G2"].value, "0.000")
        copied.vba_archive.close()
        copied.close()
        with self.assertRaises(FileExistsError):
            repair_fixture.repair(self.source, self.engine, self.output)


if __name__ == "__main__":
    unittest.main()
