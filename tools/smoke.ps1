# Runs dist\SlideFill.exe on the samples from an empty temp folder. Exit 0 only if a report was written with no problems.
# A windowed exe that fails on startup shows a dialog and waits forever, so a time limit turns that into a failure.
$root = Split-Path -Parent $PSScriptRoot
$tempRoot = (Get-Item -LiteralPath $env:TEMP -ErrorAction Stop).FullName
$workName = "slidefill_smoke_" + [guid]::NewGuid().ToString('N')
$work = Join-Path $tempRoot $workName
New-Item -ItemType Directory -Force $work | Out-Null
Copy-Item "$root\dist\SlideFill.exe" $work
$out = "$work\out.pptx"
$arguments = @("--smoke", "`"$root\samples\sample_template.pptx`"", "`"$root\samples\sample_data.xlsx`"", "`"$out`"")
$p = Start-Process "$work\SlideFill.exe" -WorkingDirectory $work -ArgumentList $arguments -WindowStyle Hidden -PassThru
$code = 1
if ($p.WaitForExit(60000)) {
    $code = $p.ExitCode
} else {
    "Smoke test FAILED: no exit after 60 seconds. The exe is probably showing a startup error."
    taskkill /T /F /PID $p.Id | Out-Null
}
if ($code -eq 0 -and -not (Test-Path $out)) {
    "Smoke test FAILED: no report was written."
    $code = 1
} elseif ($code -ne 0) {
    "Smoke test FAILED: exit code $code"
    if (Test-Path "$out.log") { Get-Content "$out.log" -Encoding UTF8 }
} else {
    "Smoke test passed"
}
Start-Sleep -Seconds 1  # the onefile bootloader needs a moment to release the exe
# Delete only this run's exact directory immediately under the resolved TEMP root.
$cleanupTarget = (Get-Item -LiteralPath $work -ErrorAction Stop).FullName
if ([IO.Path]::GetDirectoryName($cleanupTarget) -ne $tempRoot -or
    [IO.Path]::GetFileName($cleanupTarget) -ne $workName -or
    $workName -notmatch '^slidefill_smoke_[0-9a-f]{32}$') {
    throw "Refusing cleanup outside the expected smoke directory."
}
Remove-Item -LiteralPath $cleanupTarget -Recurse -Force -ErrorAction Stop
exit $code
