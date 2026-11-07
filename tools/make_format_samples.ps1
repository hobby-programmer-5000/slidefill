# Saves samples\formats\monthly.xlsx again as .xls, .xlsb and .ods with real Excel, so the readers are tested on what Excel really writes.
# The source is made by this project and holds no one else's data. Run again only if the source changes.
# Excel writes the author's name and the folder it saved to into these files. They are saved with personal information removed,
# and the script fails if the Windows user name or Excel's user name is still in any of them.
# -CheckOnly looks in the files already there and saves nothing.
param([string]$Folder = (Join-Path $PSScriptRoot "..\samples\formats"), [switch]$CheckOnly)
$Folder = (Resolve-Path $Folder).Path
$excel = New-Object -ComObject Excel.Application
$excel.DisplayAlerts = $false
# Each name, each word of it, and a word without its first letter: Office may write a Korean name given name first (홍길동 as 길동 홍).
$names = @($env:USERNAME, $excel.UserName) | Where-Object { $_ } | ForEach-Object { $_; ($_ -split '\s+') } | ForEach-Object { $_; if ($_.Length -ge 3) { $_.Substring(1) } } | Where-Object { $_.Length -ge 2 } | Select-Object -Unique
try {
    if (-not $CheckOnly) {
        $book = $excel.Workbooks.Open((Join-Path $Folder "monthly.xlsx"), 0, $true)
        $book.RemovePersonalInformation = $true
        $book.SaveAs((Join-Path $Folder "monthly.xls"), 56)
        $book.SaveAs((Join-Path $Folder "monthly.xlsb"), 50)
        $book.SaveAs((Join-Path $Folder "monthly.ods"), 60)
        $book.Close($false)
    }
} finally {
    $excel.Quit()
}

# Look in the files and, for the zipped ones, in every part inside them.
Add-Type -AssemblyName System.IO.Compression.FileSystem
$found = @()
foreach ($kind in "xls", "xlsb", "ods") {
    $path = Join-Path $Folder "monthly.$kind"
    $blobs = @(, [System.IO.File]::ReadAllBytes($path))
    if ($kind -ne "xls") {
        $zip = [System.IO.Compression.ZipFile]::OpenRead($path)
        foreach ($entry in $zip.Entries) {
            $stream = $entry.Open(); $memory = New-Object System.IO.MemoryStream; $stream.CopyTo($memory); $stream.Close()
            $blobs += , $memory.ToArray()
        }
        $zip.Dispose()
    }
    foreach ($blob in $blobs) {
        $hex = [System.BitConverter]::ToString($blob)  # a byte search: decoding the whole file as UTF-16 misses text that starts on an odd byte
        foreach ($name in $names) {
            foreach ($encoding in [System.Text.Encoding]::UTF8, [System.Text.Encoding]::Unicode, [System.Text.Encoding]::GetEncoding(949)) {
                if ($hex.Contains([System.BitConverter]::ToString($encoding.GetBytes($name)))) { $found += "monthly.$kind holds '$name'" }
            }
        }
    }
}
if ($found) {
    $found | Select-Object -Unique
    "FAILED: personal information is still in the samples. Do not commit them."
    exit 1
}
"saved, no personal information found"
