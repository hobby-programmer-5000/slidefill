# Experimental native Excel/VBA host. Original project code; existing MIT license.
# No Trust Center, registry, trusted-location, or macro-policy changes are made.
param(
    [Parameter(Mandatory=$true)][ValidateSet('Build','Run')][string]$Mode,
    [Parameter(Mandatory=$true)][string]$EngineWorkbook,
    [Parameter(Mandatory=$true)][string]$OutputJson,
    [string]$InputWorkbook,
    [string]$ModulePath,
    [string]$AppInputJson
)
$ErrorActionPreference = 'Stop'
$invariant = [Globalization.CultureInfo]::InvariantCulture
$excel = $null
$sourceBook = $null
$engineBook = $null
$ownedConfirmed = $false
$exitCode = 1
$result = $null

function New-OutputPath([string]$Path) {
    $full = [IO.Path]::GetFullPath($Path)
    if (Test-Path -LiteralPath $full) { throw "Preserve existing output: $full" }
    if (-not (Test-Path -LiteralPath ([IO.Path]::GetDirectoryName($full)) -PathType Container)) {
        throw 'The output parent directory must already exist.'
    }
    return $full
}

function Write-NewJson([string]$Path, [object]$Data) {
    $stream = [IO.File]::Open($Path, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write)
    $writer = New-Object IO.StreamWriter($stream, (New-Object Text.UTF8Encoding($false)))
    try { $writer.WriteLine(($Data | ConvertTo-Json -Depth 16)) }
    finally { $writer.Dispose(); $stream.Dispose() }
}

function Get-Hash([string]$Path) {
    return (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash.ToLowerInvariant()
}

function Get-NativeResult([object]$Value, [string]$DisplayText) {
    if ($null -eq $Value) { return @{kind='empty'; display_text=$DisplayText} }
    $errorNumber = $null
    if ($Value -is [Runtime.InteropServices.ErrorWrapper]) {
        $errorNumber = $Value.ErrorCode -band 65535
    } elseif ($Value -is [int] -and $Value -lt 0) {
        # Excel CVErr is commonly marshalled as the signed 0x800Axxxx HRESULT.
        $candidate = $Value -band 65535
        if (($Value -band -65536) -eq -2146828288 -and
            $candidate -in @(2000,2007,2015,2023,2029,2036,2042,2043)) {
            $errorNumber = $candidate
        }
    }
    if ($null -ne $errorNumber) {
        return @{kind='error'; error_code=$errorNumber; display_text=$DisplayText;
                 clr_type=$Value.GetType().FullName}
    }
    if ($Value -is [bool] -or $Value -is [string]) {
        return @{kind='unsupported_result'; clr_type=$Value.GetType().FullName;
                 display_text=$DisplayText}
    }
    try { $number = [double]$Value }
    catch { return @{kind='unsupported_result'; clr_type=$Value.GetType().FullName} }
    if ([double]::IsNaN($number) -or [double]::IsInfinity($number)) {
        return @{kind='nonfinite_result'; clr_type=$Value.GetType().FullName}
    }
    return @{kind='number'; value=$number; decimal=$number.ToString('R',$invariant);
             double_hex=([BitConverter]::ToString([BitConverter]::GetBytes($number))).Replace('-','');
             clr_type=$Value.GetType().FullName; display_text=$DisplayText}
}

function New-OperandMatrix([object[]]$Operands) {
    if ($Operands.Count -ne 3) { throw 'Each app case needs exactly three explicit operand slots.' }
    $matrix = New-Object 'object[,]' 1,3
    for ($index=0; $index -lt 3; $index++) {
        $operand = $Operands[$index]
        switch ($operand.type) {
            'blank' { $matrix[0,$index] = $null }
            'number' {
                $number = [double]::Parse([string]$operand.decimal,
                    [Globalization.NumberStyles]::Float,$invariant)
                if ([double]::IsNaN($number) -or [double]::IsInfinity($number)) {
                    throw 'Only finite numeric bridge operands are supported.'
                }
                $matrix[0,$index] = $number
            }
            'bool' {
                if ($operand.value -isnot [bool]) { throw 'A Boolean bridge operand must contain a JSON Boolean.' }
                $matrix[0,$index] = $operand.value
            }
            'text' {
                if ($operand.value -isnot [string]) { throw 'A text bridge operand must contain a JSON string.' }
                $matrix[0,$index] = $operand.value
            }
            default { throw 'Unsupported typed bridge operand.' }
        }
    }
    # Do not enumerate the rectangular SAFEARRAY into scalar pipeline items.
    return ,$matrix
}

try {
    $outputFull = New-OutputPath $OutputJson
    if ($Mode -eq 'Build') {
        if (-not $InputWorkbook -or -not $ModulePath) { throw 'Build needs InputWorkbook and ModulePath.' }
        $inputFull = (Resolve-Path -LiteralPath $InputWorkbook).Path
        $moduleFull = (Resolve-Path -LiteralPath $ModulePath).Path
        if ([IO.Path]::GetExtension($inputFull) -ne '.xlsx') { throw 'Build accepts a literal XLSX fixture only.' }
        if ([IO.Path]::GetFileName($moduleFull) -ne 'SlideFillOracle.bas') { throw 'Use the reviewed SlideFillOracle.bas module.' }
        $engineFull = New-OutputPath $EngineWorkbook
        if ([IO.Path]::GetExtension($engineFull) -ne '.xlsm') { throw 'The engine must be a new XLSM file.' }
        $sourceBefore = Get-Hash $inputFull
        $moduleHash = Get-Hash $moduleFull
    } else {
        if (-not $AppInputJson) { throw 'Run needs AppInputJson.' }
        $engineFull = (Resolve-Path -LiteralPath $EngineWorkbook).Path
        if ([IO.Path]::GetExtension($engineFull) -ne '.xlsm') { throw 'Run accepts a reviewed XLSM engine only.' }
        $appInputFull = (Resolve-Path -LiteralPath $AppInputJson).Path
        $appInput = Get-Content -LiteralPath $appInputFull -Raw | ConvertFrom-Json
        $cases = @($appInput.cases)
        if ($cases.Count -lt 1 -or $cases.Count -gt 1000) { throw 'Expected 1..1000 bridge cases.' }
        $sourceHash = [string]$appInput.source_sha256
        $moduleHash = [string]$appInput.module_source_sha256
        if ($sourceHash -notmatch '^[0-9a-f]{64}$' -or $moduleHash -notmatch '^[0-9a-f]{64}$') {
            throw 'The app manifest needs source_sha256 and module_source_sha256.'
        }
        $sourcePath = (Resolve-Path -LiteralPath ([string]$appInput.source)).Path
        if ((Get-Hash $sourcePath) -ne $sourceHash) { throw 'SourceHashMismatch: the app source fixture changed.' }
        $engineBefore = Get-Hash $engineFull
    }

    $excel = New-Object -ComObject Excel.Application
    if ($excel.Workbooks.Count -ne 0) { throw 'Expected a new empty owned Excel instance; user workbooks are not touched.' }
    $ownedConfirmed = $true
    $excel.Visible = $false
    $excel.EnableEvents = $false
    $excel.AskToUpdateLinks = $false
    $excel.AutomationSecurity = 3 # ForceDisable for all source data.
    $environment = @{excel_version=[string]$excel.Version; excel_build=[string]$excel.Build}

    if ($Mode -eq 'Build') {
        $sourceBook = $excel.Workbooks.Open($inputFull,0,$true)
        $sourceSheet = $sourceBook.Worksheets.Item('Cases')
        $used = $sourceSheet.UsedRange
        $rows = [int]$used.Rows.Count
        if ($rows -lt 2 -or $rows -gt 1001 -or $used.Columns.Count -ne 7 -or
            $used.Row -ne 1 -or $used.Column -ne 1) { throw 'Expected bounded Cases A1:G1001 fixture data.' }
        if ($used.HasFormula -ne $false) { throw 'Source fixtures must contain literals, not formulas.' }
        $headers = @('id','operation','v1','v2','v3','digits','format')
        for ($column=1; $column -le 7; $column++) {
            if ([string]$sourceSheet.Cells.Item(1,$column).Value2 -ne $headers[$column-1]) {
                throw 'Unexpected fixture header.'
            }
        }
        $engineBook = $excel.Workbooks.Add()
        $engineSheet = $engineBook.Worksheets.Item(1)
        $engineSheet.Name = 'Cases'
        $engineSheet.Range("A1:G$rows").Value2 = $used.Value2
        $engineSheet.Cells.Item(1,8).Value2 = 'direct_vba'
        $engineSheet.Cells.Item(1,9).Value2 = 'app_vba'
        # Visible metadata binds the declared source/version, not an audit of VBA bytes.
        # Plain cells avoid Office's inconsistent CustomDocumentProperties COM marshaling.
        $metadata = $engineBook.Worksheets.Add()
        $metadata.Name = '_SlideFillMeta'
        $metadata.Cells.Item(1,1).Value2 = 'field'
        $metadata.Cells.Item(1,2).Value2 = 'value'
        $metadata.Cells.Item(2,1).Value2 = 'source_sha256'
        $metadata.Cells.Item(2,2).Value2 = $sourceBefore
        $metadata.Cells.Item(3,1).Value2 = 'module_source_sha256'
        $metadata.Cells.Item(3,2).Value2 = $moduleHash
        $metadata.Range('A:A').ColumnWidth = 28
        $metadata.Range('B:B').ColumnWidth = 72
        for ($row=2; $row -le $rows; $row++) {
            $engineSheet.Cells.Item($row,8).Formula = "=SFCalc(B$row,C${row}:E$row,F$row)"
            $engineSheet.Cells.Item($row,8).NumberFormat = [string]$sourceSheet.Cells.Item($row,7).Value2
        }
        $engineSheet.Range('A:I').ColumnWidth = 26
        $imported = $false
        $importFailure = $null
        try {
            $project = $engineBook.VBProject
            if ($null -eq $project) {
                # Reflection exposes the actual COM refusal when the PowerShell
                # property adapter silently returns null for this Office interface.
                $project = $engineBook.GetType().InvokeMember('VBProject',
                    [Reflection.BindingFlags]::GetProperty,$null,$engineBook,$null)
            }
            if ($null -eq $project) { throw 'Excel returned no VBA project object.' }
            $components = $project.VBComponents
            if ($null -eq $components) {
                $components = $project.GetType().InvokeMember('VBComponents',
                    [Reflection.BindingFlags]::GetProperty,$null,$project,$null)
            }
            if ($null -eq $components -or $components.Count -lt 1) { throw 'No valid VBA component collection was returned.' }
            $component = $components.Import($moduleFull)
            if ($null -eq $component -or $component.Name -ne 'SlideFillOracle') { throw 'The reviewed module was not imported.' }
            $imported = $true
        } catch {
            $cause = $_.Exception
            while ($null -ne $cause.InnerException) { $cause = $cause.InnerException }
            $importFailure = $cause.Message.Trim()
        }
        $excel.DisplayAlerts = $false # New task-owned output only, never an overwrite.
        $engineBook.SaveAs($engineFull,52)
        $engineBook.Close($false)
        $engineBook = $null
        if ((Get-Hash $inputFull) -ne $sourceBefore) { throw 'Source fixture changed during the build.' }
        $result = @{status=$(if($imported){'built'}else{'manual_import_required'});
            environment=$environment; engine_sha256=(Get-Hash $engineFull);
            module_source_sha256=$moduleHash; source_sha256=$sourceBefore;
            module_imported=$imported; import_failure=$importFailure; cases=$rows-1;
            host_settings_changed=$false; macros_executed=$false;
            next_step=$(if($imported){'Run the three-way comparison; this build is not a VBA execution check.'}
                        else{'Import SlideFillOracle.bas into this starter XLSM with the normal VBA editor, save it, then run comparison. No Trust Center change is required by this tool.'})}
        $exitCode = $(if($imported){0}else{3})
    } else {
        $excel.AutomationSecurity = 2 # ByUI: honor the user's existing macro policy.
        $engineBook = $excel.Workbooks.Open($engineFull,0,$true)
        $metadata = $engineBook.Worksheets.Item('_SlideFillMeta')
        if ([string]$metadata.Cells.Item(2,1).Value2 -ne 'source_sha256' -or
            [string]$metadata.Cells.Item(2,2).Value2 -ne $sourceHash) {
            throw 'SourceHashMismatch: native and app fixtures are not the same declared source.'
        }
        if ([string]$metadata.Cells.Item(3,1).Value2 -ne 'module_source_sha256' -or
            [string]$metadata.Cells.Item(3,2).Value2 -ne $moduleHash) {
            throw 'ModuleHashMismatch: native and app module versions differ.'
        }
        $engineSheet = $engineBook.Worksheets.Item('Cases')
        $macro = "'" + $engineBook.Name.Replace("'","''") + "'!SlideFillOracle.SFCalc"
        $probe = New-Object 'object[,]' 1,1
        $probe[0,0] = [double]1
        try {
            $probeResult = $excel.Run($macro,'count',$probe,2)
            if ([double]$probeResult -ne 1) { throw 'The fixed VBA execution probe returned an unexpected value.' }
        } catch {
            $result = @{status='execution_blocked'; environment=$environment;
                reason=$_.Exception.Message; engine_sha256=$engineBefore;
                host_settings_changed=$false; macros_executed=$false; cases=@()}
            $exitCode = 3
        }
        if ($null -eq $result) {
            $excel.CalculateFullRebuild()
            $seen = @{}
            $nativeRows = @{}
            $lastRow = [int]$engineSheet.UsedRange.Rows.Count
            for ($row=2; $row -le $lastRow; $row++) {
                $id = [string]$engineSheet.Cells.Item($row,1).Value2
                if (-not $id -or $nativeRows.ContainsKey($id)) { throw 'Empty or duplicate native case identifier.' }
                $nativeRows[$id] = $row
            }
            $caseResults = @()
            foreach ($case in $cases) {
                $id = [string]$case.id
                if (-not $id -or $seen.ContainsKey($id) -or -not $nativeRows.ContainsKey($id)) {
                    throw 'App case identifiers must be unique and present in the engine.'
                }
                $seen[$id] = $true
                $row = $nativeRows[$id]
                if ([string]$engineSheet.Cells.Item($row,2).Value2 -ne [string]$case.operation -or
                    [int]$engineSheet.Cells.Item($row,6).Value2 -ne [int]$case.digits -or
                    [string]$engineSheet.Cells.Item($row,7).Value2 -ne [string]$case.format) {
                    throw 'The native and app operation contracts differ.'
                }
                $directCell = $engineSheet.Cells.Item($row,8)
                $direct = Get-NativeResult $directCell.Value2 ([string]$directCell.Text)
                $matrix = New-OperandMatrix @($case.operands)
                $raw = $excel.Run($macro,[string]$case.operation,$matrix,[int]$case.digits)
                $appResult = Get-NativeResult $raw ''
                if ($appResult.kind -eq 'number') {
                    $appCell = $engineSheet.Cells.Item($row,9)
                    $appCell.NumberFormat = [string]$case.format
                    $appCell.Value2 = [double]$raw
                    $appResult.display_text = [string]$appCell.Text
                }
                $equal = $direct.kind -eq $appResult.kind -and
                    (($direct.kind -eq 'number' -and $direct.double_hex -eq $appResult.double_hex) -or
                     ($direct.kind -eq 'error' -and $direct.error_code -eq $appResult.error_code))
                $caseResults += @{id=$id; source_cells="Cases!C${row}:E$row";
                                  direct=$direct; app=$appResult; bridge_equal=[bool]$equal}
            }
            if ($seen.Count -ne $nativeRows.Count) { throw 'The app case manifest omits native cases.' }
            $result = @{status='executed'; environment=$environment; engine_sha256=$engineBefore;
                app_input_sha256=(Get-Hash $appInputFull); host_settings_changed=$false;
                macros_executed=$true; cases=$caseResults;
                scope='Worksheet UDF vs Application.Run bridge consistency; Python mathematics is a separate comparison.'}
            $exitCode = 0
        }
        $engineBook.Close($false)
        $engineBook = $null
        if ((Get-Hash $engineFull) -ne $engineBefore) { throw 'The read-only engine artifact changed.' }
    }
    Write-NewJson $outputFull $result
    Write-Output ($result.status + '; evidence: ' + $outputFull)
} catch {
    Write-Error -ErrorRecord $_ -ErrorAction Continue
    if ($_.ScriptStackTrace) { Write-Output $_.ScriptStackTrace }
    $exitCode = 1
} finally {
    if ($sourceBook) { try { $sourceBook.Close($false) } catch { Write-Warning 'The owned source workbook did not close normally.' } }
    if ($engineBook) { try { $engineBook.Close($false) } catch { Write-Warning 'The owned engine workbook did not close normally.' } }
    if ($excel -and $ownedConfirmed) { try { $excel.Quit() } catch { Write-Warning 'The owned Excel instance did not quit normally.' } }
    foreach ($comObject in @($sourceBook,$engineBook,$excel)) {
        if ($comObject) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($comObject) }
    }
}
exit $exitCode
