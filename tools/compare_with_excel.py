"""Proves the "as Excel displays it" claim against the real Excel on this PC.

Excel writes a workbook with many number formats and reports each cell's displayed text. SlideFill then reads
the saved file, and every cell must render to the same text. Exit code 1 lists any difference.

Run: .venv\\Scripts\\python.exe -X utf8 tools\\compare_with_excel.py
"""

import json
import os
import subprocess
import sys
import tempfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

import core  # noqa: E402

# (value or formula, number format). Dates are given as serial numbers: 46329 is 2026-11-03.
CASES = [
    (1234567, "General"), (1234567.891, "General"), ("=0.1+0.2", "General"), ("=1/3", "General"), (-2.5, "General"),
    (0, "General"), (123456789, "General"), (0.000123, "General"),
    (1234567, "#,##0"), (1234567.5, "#,##0"), (1234.5, "#,##0.00"), (-1234.5, "#,##0"), (2.675, "0.00"),
    (1.005, "0.00"), (0.5, "0"), (1.5, "0"), (2.5, "0"), (0, "#,##0"), (1.5, "0.0##"), (1.23456, "0.0##"),
    (0.853, "0.0%"), (0.853, "0%"), (1.2, "0%"), (-0.125, "0.0%"), (0.042, "+0.0%;-0.0%"), (-0.042, "+0.0%;-0.0%"),
    (1234567, "#,##0,"), (1234567, '#,##0,,"백만"'), (1234567, "0.0,"),
    (1500, '#,##0"원"'), (1500, "[$₩-412]#,##0"), (1500, "$#,##0.00"), (-1500, "$#,##0"), (5, '0"개"'),
    (-5, "#,##0;(#,##0)"), (-5, "#,##0;[Red]-#,##0"), (0, '#,##0;-#,##0;"-"'),
    (1234, '_-* #,##0_-;-* #,##0_-;_-* "-"_-;_-@_-'), (0, '_-* #,##0_-;-* #,##0_-;_-* "-"_-;_-@_-'),
    (-1234, '_-* #,##0_-;-* #,##0_-;_-* "-"_-;_-@_-'),
    (12345, "0.00E+00"), (0.00012, "0.0E+00"), (9.99, "0.0E+00"),
    (46329, "yyyy-mm-dd"), (46329, 'yyyy"년" m"월" d"일"'), (46329, 'yyyy"년" mm"월" dd"일"'), (46329, "yy.m.d"),
    (46329, "aaa"), (46329, "aaaa"), (46329, "yyyy-mm-dd (aaa)"), (46329, "[$-412]dddd"),
    (46329, "ddd, mmm d"), (46329, "mmm-yy"), (46329, "d-mmm"), (46329, "mmmm"),
    (46329.5875, "yyyy-mm-dd hh:mm"), (46329.5875, "h:mm"), (46329.5875, "h:mm AM/PM"), (46329.5875, "hh:mm:ss"),
    (46329.5875, "mm:ss"), (46329.5875, "[$-412]AM/PM h:mm"), (1.25, "[h]:mm"), (0.5, "[h]:mm:ss"),
    ("글자", "General"), ("글자", "@"), ("글자", '@" 님"'), (True, "General"),
    (46329, "yyyy/mm/dd"),
    # Found in real public files: a number in a text-format cell, and a required digit after optional ones.
    (1, "@"), (1.5, "@"), (-2, "@"), (0, "@"), (1234567.891, "@"), (1234567, '@" 님"'), (-2, '@" 님"'),
    (4.13, "0.##0"), (4, "0.##0"), (4.5, "0.#0#"), (4.123, "0.#0#"), (0, "0.##0"), (0.5, "#.##0"),
    (4.13, r"\+###,###,###,##0.##0"), (668, r"\+###,###,###,##0.0"),
    (5, "0#"), (5, "0##"), (5, "#0#"), (0, "0#"), (1234, "0#"), (5, "0,000"), (12345, "0,000"),
]
# Formats Excel stores as locale-dependent built-in ids, where openpyxl reports the US meaning.
for _fmt in ("#,##0;-#,##0", "#,##0;[Red]-#,##0", "#,##0.00;-#,##0.00", "#,##0.00;[Red]-#,##0.00",
             "₩#,##0;-₩#,##0", "₩#,##0;[Red]-₩#,##0", "₩#,##0.00;-₩#,##0.00", "₩#,##0.00;[Red]-₩#,##0.00",
             '_-* #,##0_-;-* #,##0_-;_-* "-"_-;_-@_-', '_-₩* #,##0_-;-₩* #,##0_-;_-₩* "-"_-;_-@_-',
             '_-* #,##0.00_-;-* #,##0.00_-;_-* "-"??_-;_-@_-', '_-₩* #,##0.00_-;-₩* #,##0.00_-;_-₩* "-"??_-;_-@_-'):
    CASES += [(-1234.5, _fmt), (1234.5, _fmt), (0, _fmt)]

SCRIPT = r"""
param([string]$CasesFile, [string]$Book, [string]$Result)
$ErrorActionPreference = "Stop"
$cases = Get-Content -Raw -Encoding UTF8 $CasesFile | ConvertFrom-Json
$xl = New-Object -ComObject Excel.Application
$hadBooks = $xl.Workbooks.Count -gt 0
$taskOriginalAlerts=$xl.DisplayAlerts
$xl.DisplayAlerts = $false
try {
    $wb = $xl.Workbooks.Add()
    $ws = $wb.Worksheets.Item(1)
    $ws.Name = "s"
    $row = 1
    foreach ($case in $cases) {
        $cell = $ws.Cells.Item($row, 1)
        try { if ($case.format -ne "General") { $cell.NumberFormat = $case.format } }
        catch {
            try { $cell.NumberFormatLocal = $case.format }  # some codes, such as aaa, only exist in the local dialect
            catch { $cell.NumberFormat = "@"; $cell.Value2 = "REJECTED"; $row++; continue }
        }
        $v = $case.value
        if ($v -is [bool]) { $cell.Value2 = [bool]$v }
        elseif ($v -isnot [string]) { $cell.Value2 = [double]$v }
        elseif ($v.StartsWith("=")) { $cell.Formula = [string]$v }
        else { $cell.Value2 = [string]$v }
        $row++
    }
    $ws.Columns.Item(1).ColumnWidth = 60
    $shown = @()
    for ($r = 1; $r -lt $row; $r++) { $shown += ,$ws.Cells.Item($r, 1).Text }

    # What people really make in Excel: a defined name on a formula cell, and an Excel table on a sheet with a space.
    $other = $wb.Worksheets.Add()
    $other.Name = "My Data"
    $other.Range("A1").Value2 = "Store"; $other.Range("B1").Value2 = "Sales"
    $other.Range("A2").Value2 = "Seoul"; $other.Range("B2").Value2 = [double]1500.5
    $other.Range("A3").Value2 = "Busan"; $other.Range("B3").Formula = "=B2*2"
    $other.Range("B2:B3").NumberFormat = "#,##0.0"
    $other.ListObjects.Add(1, $other.Range("A1:B3"), $null, 1).Name = "SalesTable"
    $other.Range("D1").Formula = "=SUM(B2:B3)"
    $other.Range("D1").NumberFormat = "#,##0"
    $wb.Names.Add("TotalSales", "='My Data'!`$D`$1") | Out-Null
    $wb.Names.Add("RollingRange", "=OFFSET('My Data'!`$A`$1,0,0,2,2)") | Out-Null
    $shown += ,$other.Range("D1").Text
    $shown += ,$other.Range("B3").Text
    $wb.SaveAs($Book, 51)
    $wb.Close($false)
    $shown | ConvertTo-Json | Out-File -Encoding utf8 $Result
}
finally {
    $xl.DisplayAlerts=$taskOriginalAlerts
    if (-not $hadBooks -and $xl.Workbooks.Count -eq 0) { $xl.Quit() }
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
}
"""


def main():
    folder = tempfile.mkdtemp(prefix="slidefill_compare_")
    cases, book, result, script = (os.path.join(folder, n) for n in ("cases.json", "book.xlsx", "shown.json", "run.ps1"))
    with open(cases, "w", encoding="utf-8") as fh:
        json.dump([{"value": v, "format": f} for v, f in CASES], fh, ensure_ascii=False)
    with open(script, "w", encoding="utf-8-sig") as fh:
        fh.write(SCRIPT)
    subprocess.run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", script,
                    "-CasesFile", cases, "-Book", book, "-Result", result], check=True)
    with open(result, encoding="utf-8-sig") as fh:
        shown = json.load(fh)

    if len(shown) != len(CASES) + 2:
        print(f"Excel reported {len(shown)} cells for {len(CASES) + 2} cases")
        return 1

    source = core.ExcelSource(book)
    total, table_cell = shown[-2], shown[-1]
    named = source.text_value("TotalSales")
    rows, _where, _issues = source.range_cells("SalesTable")
    try:
        source.text_value("RollingRange")
        rolling = "no issue"
    except core.IssueError as exc:
        rolling = exc.issue.code
    dynamic, _dynamic_where, dynamic_issues = source.range_cells("RollingRange")
    checks = [("defined name on a formula cell", named[0], total),
              ("Excel table size", f"{len(rows)}x{len(rows[0])}", "3x2"),
              ("Excel table formula cell", rows[2][1][1], table_cell),
              ("dynamic range is explained in text", rolling, "range_for_text"),
              ("dynamic range resolves to source cells", (len(dynamic),len(dynamic[0]),dynamic[1][0][0],dynamic[1][1][0],bool(dynamic_issues)), (2,2,"Seoul",1500.5,False))]
    failed = [c for c in checks if c[1] != c[2]]
    for label, got, wanted in failed:
        print(f"DIFF  {label}: excel={wanted!r} slidefill={got!r}")
    print(f"{len(checks)} workbook checks, {len(failed)} differences")
    if failed:
        return 1

    ws = source.wb["s"]
    differences = warned = 0
    for row, ((value, fmt), excel) in enumerate(zip(CASES, shown), start=1):
        if excel == "REJECTED":
            differences += 1
            print(f"DIFF  row {row}: Excel rejected the format {fmt!r}; fix the case list")
            continue
        _raw, text, stored, issues = source.read(ws, row, 1)
        if "-v" in sys.argv:
            print(f"{fmt!r:45} excel={excel!r:30} slidefill={text!r}")
        if issues:
            warned += 1
            print(f"WARN  row {row}: {fmt!r} -> {[i.code for i in issues]}")
        # A * in a format stretches the cell with spaces; a slide has no cell width to stretch, so spaces don't count.
        same = "".join(text.split()) == "".join(excel.split()) if "*" in stored else text == excel.strip()
        if not same:
            differences += 1
            print(f"DIFF  row {row}: value={value!r} format={fmt!r} stored={stored!r} excel={excel!r} slidefill={text!r}")
    print(f"{len(CASES)} cases, {differences} differences, {warned} warnings")
    return 1 if differences else 0


if __name__ == "__main__":
    sys.exit(main())
