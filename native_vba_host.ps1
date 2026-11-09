# SlideFill's optional Excel comparison. Existing MIT license.
param(
    [Parameter(Mandatory=$true)][string]$EnginePath,
    [Parameter(Mandatory=$true)][string]$ExpectedSha256,
    [Parameter(Mandatory=$true)][string]$RequestPath,
    [Parameter(Mandatory=$true)][string]$ResponsePath
)

$ErrorActionPreference = 'Stop'
$operations = @('sum', 'mean', 'count', 'growth', 'points', 'round')
$errorCodes = @(2007, 2015, 2036)
$excel = $null
$book = $null
$engineStream = $null
$ownsExcel = $false
$statusOnFailure = 'invalid'
$exitCode = 1

function Convert-Case($Case) {
    if ($Case.operation -isnot [string] -or $Case.operation -cnotin $operations -or
        ($Case.digits -isnot [int] -and $Case.digits -isnot [long]) -or
        $Case.digits -lt -15 -or $Case.digits -gt 15 -or
        $Case.operands -isnot [array] -or $Case.operands.Count -lt 1 -or
        $Case.operands.Count -gt 4096) {
        throw 'Invalid VBA comparison case.'
    }
    $matrix = New-Object 'object[,]' 1,$Case.operands.Count
    for ($index = 0; $index -lt $Case.operands.Count; $index++) {
        $bits = $Case.operands[$index]
        if ($null -eq $bits) { continue }
        if ($bits -isnot [string] -or $bits -notmatch '^[0-9a-f]{16}$') {
            throw 'Invalid binary64 operand.'
        }
        $bytes = New-Object 'byte[]' 8
        for ($offset = 0; $offset -lt 8; $offset++) {
            $bytes[$offset] = [Convert]::ToByte($bits.Substring(2 * $offset, 2), 16)
        }
        $number = [BitConverter]::ToDouble($bytes, 0)
        if ([double]::IsNaN($number) -or [double]::IsInfinity($number)) {
            throw 'The VBA operand is not finite.'
        }
        $matrix[0,$index] = $number
    }
    return @{ operation = $Case.operation; digits = [int]$Case.digits; matrix = $matrix }
}

function Convert-Result($Value) {
    if ($Value -is [double] -and -not [double]::IsNaN($Value) -and -not [double]::IsInfinity($Value)) {
        $bits = [BitConverter]::ToString([BitConverter]::GetBytes($Value)).Replace('-', '')
        return @{kind = 'number'; double_hex = $bits}
    }
    if ($Value -is [Runtime.InteropServices.ErrorWrapper]) {
        $errorNumber = $Value.ErrorCode -band 65535
    } elseif ($Value -is [int] -and ($Value -band -65536) -eq -2146828288) {
        # The VARIANT error can arrive as a signed Excel HRESULT.
        $errorNumber = $Value -band 65535
    } else {
        throw 'The VBA function did not return a finite Double or a calculation error.'
    }
    if ($errorNumber -notin $errorCodes) {
        throw 'Excel returned an unexpected VBA error.'
    }
    return @{kind = 'error'; error_code = $errorNumber}
}

try {
    if (-not [BitConverter]::IsLittleEndian) { throw 'Unsupported host byte order.' }
    if ((Get-Item -LiteralPath $RequestPath).Length -gt 512KB) { throw 'VBA comparison request is too large.' }
    $request = [IO.File]::ReadAllText($RequestPath) | ConvertFrom-Json
    if ($request.version -ne 1 -or $request.cases -isnot [array] -or
        $request.cases.Count -lt 1 -or $request.cases.Count -gt 128) {
        throw 'Invalid VBA comparison batch.'
    }
    $operandCount = 0
    $cases = @(foreach ($case in $request.cases) {
        $converted = Convert-Case $case
        $operandCount += $converted.matrix.Length
        if ($operandCount -gt 16384) { throw 'Too many operands in the VBA comparison batch.' }
        $converted
    })

    $path = [IO.Path]::GetFullPath($EnginePath)
    if ($path -notmatch '^[A-Za-z]:\\' -or $path.Substring(2).Contains(':') -or
        [IO.Path]::GetExtension($path) -ine '.xlsm' -or
        $ExpectedSha256 -notmatch '^[0-9a-f]{64}$') {
        throw 'Select a local, pinned SlideFill VBA engine workbook.'
    }
    $drive = New-Object IO.DriveInfo([IO.Path]::GetPathRoot($path))
    if ($drive.DriveType -notin @('Fixed', 'Removable')) {
        throw 'The SlideFill VBA engine must be on a local drive.'
    }
    $current = $path
    while ($current) {
        if (([IO.File]::GetAttributes($current) -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
            throw 'The SlideFill VBA engine path must not use a file or directory link.'
        }
        $current = [IO.Path]::GetDirectoryName($current)
    }
    # Hold the read lock across both digest verification and Excel's read-only open.
    $engineStream = [IO.File]::Open($path, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::Read)
    if ($engineStream.Length -lt 1 -or $engineStream.Length -gt 1MB) {
        throw 'Invalid SlideFill VBA engine file size.'
    }
    $sha = [Security.Cryptography.SHA256]::Create()
    try { $actual = [BitConverter]::ToString($sha.ComputeHash($engineStream)).Replace('-', '') }
    finally { $sha.Dispose() }
    if ($actual -ine $ExpectedSha256) {
        throw 'The selected workbook differs from the reviewed SlideFill VBA engine.'
    }

    $statusOnFailure = 'unavailable'
    Add-Type -Namespace SlideFill -Name ExcelWindow -MemberDefinition @'
[DllImport("user32.dll")]
public static extern uint GetWindowThreadProcessId(IntPtr window, out uint process);
'@
    $existingExcel = @((Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Id)
    $excel = New-Object -ComObject Excel.Application
    [uint32]$excelProcess = 0
    [void][SlideFill.ExcelWindow]::GetWindowThreadProcessId([IntPtr]([int64]$excel.Hwnd), [ref]$excelProcess)
    if ($excelProcess -eq 0 -or $excelProcess -in $existingExcel -or $excel.Workbooks.Count -ne 0 -or
        $excel.Visible -or $excel.UserControl) {
        throw 'Could not verify a separate, unused Excel automation instance.'
    }
    $ownsExcel = $true
    $excel.EnableEvents = $false
    $excel.DisplayAlerts = $false
    $excel.AskToUpdateLinks = $false
    $excel.AutomationSecurity = 2 # msoAutomationSecurityByUI: honor the existing Office policy.

    $statusOnFailure = 'blocked'
    $missing = [Type]::Missing
    $book = $excel.Workbooks.Open($path, 0, $true, $missing, $missing, $missing, $true,
                                 $missing, $missing, $false, $false, $missing, $false)
    if (-not $book.ReadOnly -or $book.FullName -ine $path) {
        $statusOnFailure = 'failed'
        throw 'Excel did not open the verified engine path read-only.'
    }
    $macro = "'" + $book.Name.Replace("'", "''") + "'!SlideFillOracle.SFCalc"
    $results = @(foreach ($case in $cases) {
        $statusOnFailure = 'blocked'
        $raw = $excel.Run($macro, $case.operation, $case.matrix, $case.digits)
        $statusOnFailure = 'failed'
        Convert-Result $raw
    })
    $response = @{status = 'executed'; cases = $results}
    $exitCode = 0
} catch {
    $response = @{status = $statusOnFailure; cases = @(); detail = $_.Exception.Message}
} finally {
    if ($book) {
        try { $book.Close($false) } catch {}
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($book)
    }
    if ($excel) {
        try {
            if ($ownsExcel -and -not $excel.Visible -and $excel.Workbooks.Count -eq 0) { $excel.Quit() }
        } catch {}
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    }
    if ($engineStream) { $engineStream.Dispose() }
}
$stream = [IO.File]::Open($ResponsePath, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write)
$writer = New-Object IO.StreamWriter($stream, (New-Object Text.UTF8Encoding($false)))
try { $writer.Write(($response | ConvertTo-Json -Depth 6 -Compress)) }
finally { $writer.Dispose(); $stream.Dispose() }
exit $exitCode
