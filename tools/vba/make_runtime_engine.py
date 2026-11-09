"""Package the already tested VBA project in a fresh, deterministic workbook.

Usage: python tools/vba/make_runtime_engine.py TESTED.xlsm NEW.xlsm
Only the pinned VBA Project part is read from TESTED.xlsm. No workbook cells,
document properties, or relationships are copied from that source.
"""

import argparse
import hashlib
from pathlib import Path
from zipfile import ZIP_STORED, ZipFile, ZipInfo


VBA_PART = "xl/vbaProject.bin"
VBA_SHA256 = "be5fe6fad262880671e02098402098586b44e12ce344117ce748f885aa5d94c4"
SPREADSHEET = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
DOCUMENT_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PACKAGE_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
VBA_REL = "http://schemas.microsoft.com/office/2006/relationships/vbaProject"
CONTENT_TYPES = "http://schemas.openxmlformats.org/package/2006/content-types"


def build(source: Path, output: Path) -> None:
    """Write a new workbook containing the tested project and two empty sheets."""
    with ZipFile(source) as original:
        projects = [entry for entry in original.infolist() if entry.filename == VBA_PART]
        if len(projects) != 1 or projects[0].file_size > 65536:
            raise ValueError("Expected one bounded VBA project in the tested workbook.")
        project = original.read(projects[0])
    if hashlib.sha256(project).hexdigest() != VBA_SHA256:
        raise ValueError("The supplied VBA project differs from the tested project.")

    parts = {
        "[Content_Types].xml": f"""<Types xmlns="{CONTENT_TYPES}">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.ms-excel.sheet.macroEnabled.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>""",
        "_rels/.rels": f"""<Relationships xmlns="{PACKAGE_REL}">
<Relationship Id="rId1" Type="{DOCUMENT_REL}/officeDocument" Target="xl/workbook.xml"/>
</Relationships>""",
        "xl/workbook.xml": f"""<workbook xmlns="{SPREADSHEET}" xmlns:r="{DOCUMENT_REL}">
<workbookPr codeName="ThisWorkbook"/>
<sheets>
<sheet name="Engine" sheetId="1" r:id="rId1"/>
<sheet name="Reserved" sheetId="2" state="hidden" r:id="rId2"/>
</sheets>
</workbook>""",
        "xl/_rels/workbook.xml.rels": f"""<Relationships xmlns="{PACKAGE_REL}">
<Relationship Id="rId1" Type="{DOCUMENT_REL}/worksheet" Target="worksheets/sheet1.xml"/>
<Relationship Id="rId2" Type="{DOCUMENT_REL}/worksheet" Target="worksheets/sheet2.xml"/>
<Relationship Id="rId3" Type="{VBA_REL}" Target="vbaProject.bin"/>
</Relationships>""",
    }
    # Retain both sheet CodeNames because the VBA project has document
    # modules for them. Both sheets are empty, and the host uses no cells.
    for index in (1, 2):
        parts[f"xl/worksheets/sheet{index}.xml"] = f"""<worksheet xmlns="{SPREADSHEET}">
<sheetPr codeName="Sheet{index}"/>
<sheetData/>
</worksheet>"""

    with ZipFile(output, "x", compression=ZIP_STORED) as workbook:
        for name, xml in parts.items():
            workbook.writestr(ZipInfo(name), xml.encode("utf-8"))
        workbook.writestr(ZipInfo(VBA_PART), project)


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("source", type=Path)
    parser.add_argument("output", type=Path)
    args = parser.parse_args()
    build(args.source, args.output)


if __name__ == "__main__":
    main()
