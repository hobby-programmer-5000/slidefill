# Saves samples\formats\monthly.xlsx again as .xls, .xlsb and .ods with real Excel, so the readers are tested on what Excel really writes.
# The source is made by this project and holds no one else's data. Run again only if the source changes.
param([string]$Folder = (Join-Path $PSScriptRoot "..\samples\formats"))
$Folder = (Resolve-Path $Folder).Path
$excel = New-Object -ComObject Excel.Application
$excel.DisplayAlerts = $false
try {
    $book = $excel.Workbooks.Open((Join-Path $Folder "monthly.xlsx"), 0, $true)
    $book.SaveAs((Join-Path $Folder "monthly.xls"), 56)
    $book.SaveAs((Join-Path $Folder "monthly.xlsb"), 50)
    $book.SaveAs((Join-Path $Folder "monthly.ods"), 60)
    $book.Close($false)
    "saved"
} finally {
    $excel.Quit()
}
