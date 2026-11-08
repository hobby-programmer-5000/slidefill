Attribute VB_Name = "SlideFillOracle"
Option Explicit

' Original SlideFill comparison spike. No file/network/automatic-event behavior.
' Numeric literals/Range.Value2 only; exact Python arithmetic stays independent.
Private Const MAX_CELLS As Long = 100000

Public Function SFCalc(ByVal operation As String, ByVal values As Variant, _
                       Optional ByVal digits As Variant = 2) As Variant
    Dim op As String, raw As Variant, matrix As Variant, source As Range
    Dim row As Long, column As Long, count As Long
    Dim number As Double, total As Double, first As Double, second As Double
    Dim decimalTotal As Variant, places As Long, digitValue As Double
    On Error GoTo InvalidInput
    op = LCase$(Trim$(operation))
    Select Case op
        Case "sum", "mean", "count", "growth", "points", "round", _
             "vba_round_probe", "decimal_sum"
        Case Else: GoTo InvalidInput
    End Select
    If IsObject(values) Then
        If Not TypeOf values Is Excel.Range Then GoTo InvalidInput
        Set source = values
        If source.Areas.count <> 1 Then GoTo InvalidInput
        If source.Cells.CountLarge > MAX_CELLS Then GoTo InvalidInput
        raw = source.Value2
    Else
        raw = values
    End If
    If IsArray(raw) Then
        If SFArrayRank(raw) <> 2 Then GoTo InvalidInput
        If CDbl(UBound(raw, 1) - LBound(raw, 1) + 1) * _
           CDbl(UBound(raw, 2) - LBound(raw, 2) + 1) > MAX_CELLS Then GoTo InvalidInput
        matrix = raw
    Else
        ReDim matrix(1 To 1, 1 To 1)
        matrix(1, 1) = raw
    End If
    If op = "round" Or op = "vba_round_probe" Then
        If Not SFNumericType(digits) Then GoTo InvalidInput
        digitValue = CDbl(digits)
        If Not SFFinite(digitValue) Then GoTo InvalidInput
        If digitValue <> Fix(digitValue) Or Abs(digitValue) > 15# Then GoTo InvalidInput
        If op = "vba_round_probe" And digitValue < 0# Then GoTo InvalidInput
        places = CLng(digitValue)
    End If
    decimalTotal = CDec(0)
    For row = LBound(matrix, 1) To UBound(matrix, 1)
        For column = LBound(matrix, 2) To UBound(matrix, 2)
            If Not IsEmpty(matrix(row, column)) Then
                If Not SFNumericType(matrix(row, column)) Then GoTo InvalidInput
                On Error GoTo NumericFailure
                number = CDbl(matrix(row, column))
                If Not SFFinite(number) Then GoTo NumericFailure
                count = count + 1
                If count = 1 Then first = number
                If count = 2 Then second = number
                If op = "sum" Or op = "mean" Then total = total + number
                If op = "decimal_sum" Then decimalTotal = decimalTotal + CDec(matrix(row, column))
                On Error GoTo InvalidInput
            End If
        Next column
    Next row
    On Error GoTo NumericFailure
    Select Case op
        Case "sum": SFCalc = total
        Case "count": SFCalc = CDbl(count)
        Case "mean"
            If count = 0 Then GoTo ZeroBase
            SFCalc = total / CDbl(count)
        Case "growth", "points"
            If count <> 2 Then GoTo InvalidInput
            If op = "growth" Then
                If second = 0# Then GoTo ZeroBase
                SFCalc = (first - second) / Abs(second)
            Else
                SFCalc = first - second
            End If
        Case "round", "vba_round_probe"
            If count <> 1 Then GoTo InvalidInput
            If op = "round" Then
                SFCalc = Application.WorksheetFunction.Round(first, places)
            Else
                SFCalc = VBA.Round(first, places)
            End If
        Case "decimal_sum": SFCalc = CDbl(decimalTotal)
    End Select
    If Not SFFinite(CDbl(SFCalc)) Then GoTo NumericFailure
    Exit Function
InvalidInput:
    SFCalc = CVErr(xlErrValue)
    Exit Function
ZeroBase:
    SFCalc = CVErr(xlErrDiv0)
    Exit Function
NumericFailure:
    SFCalc = CVErr(xlErrNum)
End Function

Private Function SFNumericType(ByVal value As Variant) As Boolean
    Select Case VarType(value)
        Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal, 20
            SFNumericType = True
        Case Else
            SFNumericType = False
    End Select
End Function

Private Function SFFinite(ByVal value As Double) As Boolean
    On Error GoTo NotFinite
    If value <> value Then Exit Function
    If value <> 0# Then
        If value / 2# = value Then Exit Function
    End If
    SFFinite = True
NotFinite:
End Function

Private Function SFArrayRank(ByRef value As Variant) As Long
    Dim dimension As Long, bound As Long
    On Error GoTo Finished
    For dimension = 1 To 3
        bound = LBound(value, dimension)
        SFArrayRank = dimension
    Next dimension
Finished:
End Function

