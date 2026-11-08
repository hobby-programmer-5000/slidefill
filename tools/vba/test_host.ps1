# Headless checks for the native result/operand protocol; no Excel instance is opened.
$ErrorActionPreference = 'Stop'
$invariant = [Globalization.CultureInfo]::InvariantCulture
$errors = $null
$tokens = $null
$ast = [Management.Automation.Language.Parser]::ParseFile(
    (Join-Path $PSScriptRoot 'host.ps1'),[ref]$tokens,[ref]$errors)
if ($errors.Count) { throw 'host.ps1 has PowerShell syntax errors.' }
$functions = @($ast.FindAll({param($node)
    $node -is [Management.Automation.Language.FunctionDefinitionAst] -and
    $node.Name -in @('Get-NativeResult','New-OperandMatrix')
},$true))
if ($functions.Count -ne 2) { throw 'The expected protocol functions were not found.' }
foreach ($function in $functions) {
    . ([ScriptBlock]::Create($function.Extent.Text))
}
$checked = 0
foreach ($pair in @(@(-2146826281,2007),@(-2146826273,2015),@(-2146826252,2036))) {
    $result = Get-NativeResult ([int]$pair[0]) '#error'
    if ($result.kind -ne 'error' -or $result.error_code -ne $pair[1]) { throw 'Excel CVErr HRESULT was not retained.' }
    $checked++
}
$result = Get-NativeResult ([double]-2146826281) ''
if ($result.kind -ne 'number') { throw 'An ordinary numeric Double was confused with a CVErr.' }
$checked++
$result = Get-NativeResult ([double]0.30000000000000004) ''
if ($result.decimal -ne '0.30000000000000004') { throw 'The native result lost its round-trip low digit.' }
$checked++
$result = Get-NativeResult (New-Object Runtime.InteropServices.ErrorWrapper(2015)) '#VALUE!'
if ($result.kind -ne 'error' -or $result.error_code -ne 2015) { throw 'An explicit COM error wrapper was lost.' }
$checked++
$matrix = New-OperandMatrix @(@{type='number';decimal='0.1'},@{type='blank'},@{type='bool';value=$false})
if ($matrix.Rank -ne 2 -or $matrix.GetLength(0) -ne 1 -or $matrix.GetLength(1) -ne 3 -or
    $matrix[0,0] -ne 0.1 -or $null -ne $matrix[0,1] -or
    $matrix[0,2] -isnot [bool] -or $matrix[0,2]) { throw 'The SAFEARRAY operand shape/types changed.' }
$checked++
foreach ($operands in @(
    @(@{type='bool';value='false'},@{type='blank'},@{type='blank'}),
    @(@{type='number';decimal='Infinity'},@{type='blank'},@{type='blank'}),
    @(@{type='blank'},@{type='blank'})
)) {
    $refused = $false
    try { New-OperandMatrix $operands | Out-Null } catch { $refused = $true }
    if (-not $refused) { throw 'An invalid typed operand payload was silently accepted.' }
    $checked++
}
Write-Output "PASS: $checked host protocol checks; no Excel/VBA execution is implied."
