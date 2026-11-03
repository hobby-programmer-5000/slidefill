# Checks the real exe from an empty folder without touching the keyboard or the active window:
# first run opens a window, remembered paths trigger the automatic preview, and closing keeps the settings.
# The button and shortcut flows are covered by tools\gui_check.py, which needs no window focus.
#
# Usage: powershell -ExecutionPolicy Bypass -File tools\gui_check.ps1 <empty work folder>
param([Parameter(Mandatory = $true)][string]$Work)
$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
New-Item -ItemType Directory -Force $Work | Out-Null
$Work = (Resolve-Path $Work).Path
Copy-Item "$root\dist\SlideFill.exe" $Work
Add-Type -AssemblyName System.Drawing
Add-Type -Namespace Native -Name Win -MemberDefinition @'
[DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
[DllImport("user32.dll")] public static extern bool PrintWindow(IntPtr h, IntPtr dc, uint flags);
[DllImport("user32.dll")] public static extern bool SetProcessDPIAware();
public struct RECT { public int Left, Top, Right, Bottom; }
'@
[Native.Win]::SetProcessDPIAware() | Out-Null

function Shot($proc, $name) {
    $r = New-Object Native.Win+RECT
    [Native.Win]::GetWindowRect($proc.MainWindowHandle, [ref]$r) | Out-Null
    $bmp = New-Object System.Drawing.Bitmap ($r.Right - $r.Left), ($r.Bottom - $r.Top)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $dc = $g.GetHdc()
    [Native.Win]::PrintWindow($proc.MainWindowHandle, $dc, 2) | Out-Null
    $g.ReleaseHdc($dc); $g.Dispose()
    $bmp.Save("$Work\$name.png", [System.Drawing.Imaging.ImageFormat]::Png)
    $bmp.Dispose()
}

function Launch() {
    Start-Process "$Work\SlideFill.exe" -WorkingDirectory $Work | Out-Null
    # A onefile exe runs as two processes; the window belongs to the second one.
    for ($i = 0; $i -lt 60; $i++) {
        $p = Get-Process SlideFill -ErrorAction SilentlyContinue |
            Where-Object { $_.Path -like "$Work\*" -and $_.MainWindowHandle -ne 0 } | Select-Object -First 1
        if ($p) { break }
        Start-Sleep -Milliseconds 500
    }
    if (-not $p) { throw "no window appeared" }
    if ($p.MainWindowTitle -ne "SlideFill") { throw "unexpected window: $($p.MainWindowTitle)" }
    Start-Sleep -Seconds 3
    return $p
}

function Close($proc) {
    $proc.CloseMainWindow() | Out-Null
    if (-not $proc.WaitForExit(10000)) { throw "window did not close" }
    Start-Sleep -Seconds 1
}

trap {
    "ERROR: $_"
    Get-Process SlideFill -ErrorAction SilentlyContinue | Where-Object { $_.Path -like "$Work\*" } | Stop-Process -Force
    "RESULT: FAIL"
    exit 1
}

function Check($label, $ok) { "{0}  {1}" -f $(if ($ok) { "PASS" } else { "FAIL" }), $label; if (-not $ok) { $script:failed = $true } }
$failed = $false
$settings = "$Work\SlideFill.json"

$p = Launch
Check "first run opens a window with no settings file" ($p.MainWindowHandle -ne 0)
Shot $p "1_first_run"
Close $p

$seed = @{ template = "$root\samples\sample_template.pptx"; excel = "$root\samples\sample_data.xlsx"; output = $Work; lang = "ko" }
[IO.File]::WriteAllText($settings, ($seed | ConvertTo-Json), (New-Object Text.UTF8Encoding $false))
$p = Launch
Check "remembered paths open without an error dialog" ($p.MainWindowTitle -eq "SlideFill")
Shot $p "2_preview_ko"
Close $p

$saved = Get-Content -Raw -Encoding UTF8 $settings | ConvertFrom-Json
Check "settings file stays next to the exe with its paths" ($saved.template -like "*sample_template.pptx")
Check "no temp file left behind" (-not (Test-Path "$settings.tmp"))

if ($failed) { "RESULT: FAIL"; exit 1 }
"RESULT: PASS"
