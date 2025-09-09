Option Explicit

Sub CALx()
'
' CALx Macro
'
Sheets("Parsed").Visible = True
Sheets("Parsed").Select
Range("A12").FormulaR1C1 = _
        "=IF(R[-2]C=""PR"",DATE(YEAR(R[-10]C)-1,MONTH(R[-10]C),DAY(R[-10]C)),'DATA ENTRY CHECK'!R[39]C[8])"
Range("A40").FormulaR1C1 = "=CONCATENATE(R1C1,"" "",R1C2)"

If Range("C40") = Range("A40") Then
    Range("A41:A64").Value = Range("C41:C64").Value
ElseIf Range("D40") = Range("A40") Then
    Range("A41:A64").Value = Range("D41:D64").Value
ElseIf Range("E40") = Range("A40") Then
    Range("A41:A64").Value = Range("E41:E64").Value
ElseIf Range("F40") = Range("A40") Then
    Range("A41:A64").Value = Range("F41:F64").Value
ElseIf Range("G40") = Range("A40") Then
    Range("A41:A64").Value = Range("G41:G64").Value
ElseIf Range("H40") = Range("A40") Then
    Range("A41:A64").Value = Range("H41:H64").Value
ElseIf Range("I40") = Range("A40") Then
    Range("A41:A64").Value = Range("I41:I64").Value
ElseIf Range("J40") = Range("A40") Then
    Range("A41:A64").Value = Range("J41:J64").Value
ElseIf Range("K40") = Range("A40") Then
    Range("A41:A64").Value = Range("K41:K64").Value
ElseIf Range("L40") = Range("A40") Then
    Range("A41:A64").Value = Range("L41:L64").Value
End If

Range("A39").FormulaR1C1 = _
        "=IF(R[2]C="""","""",IF(AND(R[2]C>=R[-27]C,R[2]C<=R[-27]C[1]),""CURRENT"",""NOT CURRENT""))"
Range("A39:A64").Value = Range("A39:A64").Value
Range("C40:L64").Clear

If Range("A41") <> "" Then

Application.DisplayAlerts = False
Range("A43:A64").TextToColumns Destination:=Range("A43"), DataType:=xlDelimited, _
        TextQualifier:=xlNone, ConsecutiveDelimiter:=True, Tab:=False, Semicolon _
        :=False, Comma:=False, Space:=True, Other:=False, OtherChar:="E", _
        FieldInfo:=Array(Array(1, 2), Array(2, 2)), TrailingMinusNumbers:=True
Application.DisplayAlerts = True

Workbooks("Ab.xlsm").Sheets("PRS").Range("A39:B64").Value = Workbooks("A.xlsm").Sheets("Parsed").Range("A39:B64").Value
End If

End Sub