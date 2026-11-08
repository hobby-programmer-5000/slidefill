"""Make a new typed-data test copy, preserving an already imported VBA binary."""
from pathlib import Path
import argparse
import hashlib
import io
import json
from zipfile import ZipFile

from openpyxl import load_workbook


def repair(source_path, engine_path, output_path):
    source_path, engine_path, output_path = map(Path, (source_path, engine_path, output_path))
    if output_path.exists():
        raise FileExistsError("Preserve existing output")
    if source_path.suffix.lower() != ".xlsx" or engine_path.suffix.lower() != ".xlsm" or output_path.suffix.lower() != ".xlsm":
        raise ValueError("Expected XLSX source and XLSM engine/copy")
    source_bytes, engine_bytes = source_path.read_bytes(), engine_path.read_bytes()
    source_hash = hashlib.sha256(source_bytes).hexdigest()
    with ZipFile(io.BytesIO(engine_bytes)) as archive:
        if "xl/vbaProject.bin" not in archive.namelist():
            raise ValueError("The saved engine does not contain an imported VBA project yet")
        vba = archive.read("xl/vbaProject.bin")
    source = load_workbook(io.BytesIO(source_bytes), data_only=False)
    engine = load_workbook(io.BytesIO(engine_bytes), keep_vba=True, data_only=False)
    try:
        if engine["_SlideFillMeta"]["B2"].value != source_hash:
            raise ValueError("The engine was prepared for a different source")
        original, target = source["Cases"], engine["Cases"]
        if original.max_column != 7 or original.max_row != target.max_row:
            raise ValueError("Case table coverage differs")
        changed = []
        for row in original:
            for cell in row:
                if cell.data_type == "f":
                    raise ValueError("The source fixture must contain literals only")
                destination = target.cell(cell.row, cell.column)
                if type(destination.value) is not type(cell.value) or destination.value != cell.value:
                    changed.append({"cell": cell.coordinate, "before_type": type(destination.value).__name__,
                                    "before": destination.value, "after_type": type(cell.value).__name__, "after": cell.value})
                destination.value = cell.value
                destination.number_format = "@" if isinstance(cell.value, str) else cell.number_format
        buffer = io.BytesIO()
        engine.save(buffer)
        repaired = buffer.getvalue()
        with ZipFile(io.BytesIO(repaired)) as archive:
            if archive.read("xl/vbaProject.bin") != vba:
                raise ValueError("VBA binary identity changed")
        with output_path.open("xb") as stream:
            stream.write(repaired)
    finally:
        source.close()
        if engine.vba_archive is not None:
            engine.vba_archive.close()
        engine.close()
    if source_path.read_bytes() != source_bytes or engine_path.read_bytes() != engine_bytes:
        raise ValueError("An input changed while preparing the copy")
    return {"status": "prepared_copy", "source_sha256": source_hash,
            "original_engine_sha256": hashlib.sha256(engine_bytes).hexdigest(),
            "copy_sha256": hashlib.sha256(repaired).hexdigest(),
            "vba_binary_sha256": hashlib.sha256(vba).hexdigest(), "vba_binary_unchanged": True,
            "changed_data_cells": changed,
            "scope": "Data-only correction on a new copy; binary preservation does not prove VBA source identity or execution."}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("source", type=Path)
    parser.add_argument("engine", type=Path)
    parser.add_argument("output", type=Path)
    args = parser.parse_args()
    try:
        print(json.dumps(repair(args.source, args.engine, args.output), indent=2))
        return 0
    except (ValueError, OSError, KeyError) as exc:
        print(f"Cannot prepare copy: {exc}")
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
