# Opens a generated report in the real PowerPoint and exports each slide to PNG.
# A damaged file makes Open throw, so exit 0 means PowerPoint accepted it without needing repair.
# Also activates each chart's embedded workbook, which is what "Edit Data" opens.
#
# Usage: powershell -ExecutionPolicy Bypass -File tools\verify_powerpoint.ps1 <report.pptx> [<png folder>]
param(
    [Parameter(Mandatory = $true)][string]$Path,
    [string]$OutDir = ""
)
$ErrorActionPreference = "Stop"
$full = (Resolve-Path $Path).Path
$ppt = New-Object -ComObject PowerPoint.Application
$hadOpen = $ppt.Presentations.Count -gt 0
$exit = 0
try {
    # ReadOnly, not Untitled, no window
    $deck = $ppt.Presentations.Open($full, -1, 0, 0)
    "Opened: {0} slides" -f $deck.Slides.Count
    foreach ($slide in $deck.Slides) {
        foreach ($shape in $slide.Shapes) {
            if ($shape.HasChart) {
                $chart = $shape.Chart
                $chart.ChartData.Activate()
                $book = $chart.ChartData.Workbook
                $sheet = $book.Worksheets.Item(1)
                $used = $sheet.UsedRange
                "Slide {0} chart: embedded data {1} rows x {2} cols, A2={3}, B2={4}" -f `
                    $slide.SlideIndex, $used.Rows.Count, $used.Columns.Count, `
                    $sheet.Cells.Item(2, 1).Text, $sheet.Cells.Item(2, 2).Text
                $book.Close()
            }
            if ($shape.HasTable) {
                "Slide {0} table: {1} rows x {2} cols" -f $slide.SlideIndex, $shape.Table.Rows.Count, $shape.Table.Columns.Count
            }
        }
    }
    if ($OutDir) {
        New-Item -ItemType Directory -Force $OutDir | Out-Null
        $deck.Export((Resolve-Path $OutDir).Path, "PNG", 1280, 720)
        "Exported PNGs to $OutDir"
    }
    $deck.Close()
    "OK"
}
catch {
    "FAILED: $($_.Exception.Message)"
    $exit = 1
}
finally {
    if (-not $hadOpen -and $ppt.Presentations.Count -eq 0) { $ppt.Quit() }
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null
}
exit $exit
