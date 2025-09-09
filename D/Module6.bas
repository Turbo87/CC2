' VBA Module: Application Control and User Interface
' Purpose: Manages application termination, form displays, and user interface controls.
' Handles calibration forms, cursor movement, and application exit procedures.

Option Explicit
Sub Buhby()

    ActiveWorkbook.Saved = True
    MsgBox "Thanks for using Claim Check! Click OK to exit"
    Application.Quit

End Sub
Sub RoundedRectangle1_Click()
Sheets("Calibration").Select
    frmCAL.Show
End Sub

Sub RoundedRectangle12_Click()
Sheets("Calibration").Select
    frmUNIT.Show
End Sub

Sub MoveC()
'
' MoveCursor after drop-down 2 or 30
'
    If Range("E8") < 3 Then
        Range("F12").Select
    ElseIf Range("E8") >= 3 Then
        Range("F14").Select
    End If
End Sub
Sub Cal()
Attribute Cal.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Cal Macro retains key defaults
'
    Application.ScreenUpdating = False
    Sheets("Calibration").Select
    Sheets("Calibration").Unprotect Password:="spike"
    Range("B8:C8").FormulaR1C1 = _
        "=IF(R[46]C[-1]<>R[45]C[-1],CONCATENATE(R[45]C[-1],"" Calibrations saved""),IF(R[46]C[-1]=1,CONCATENATE(R[46]C[-1],"" Calibration saved""),CONCATENATE(R[46]C[-1],"" Calibrations saved"")))"
    Range("E8:F8").FormulaR1C1 = "1"
    Range("E8:F8").Locked = False
    Range("E8:F8").FormulaHidden = True
    ActiveSheet.Shapes.Range(Array("Drop Down 30")).Visible = msoFalse
    ActiveSheet.Shapes.Range(Array("Drop Down 2")).Visible = msoTrue
    Range("E10:F10").FormulaR1C1 = "1"
    ActiveSheet.Rows("11:12").Hidden = False
    Range("D12").ClearContents
    Range("E12").Value = "FR Serial #:"
    Range("F12").ClearContents
    Range("F12").Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateTextLength, AlertStyle:=xlValidAlertStop, _
        Operator:=xlEqual, Formula1:="3"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = "Three characters are required"
        .ShowInput = False
        .ShowError = True
    End With
    Range("F12").Locked = False
    Range("F12").FormulaHidden = True
    Range("F14").ClearContents
    Range("F14").Locked = False
    Range("F14").FormulaHidden = True
    Range("F16").FormulaR1C1 = "1"
    Range("F16").Locked = False
    Range("F16").FormulaHidden = True
    Range("B10").FormulaR1C1 = _
        "=IF(R10C5=2,R[2]C[9],IF(R10C5=3,R[3]C[9],IF(R10C5=4,R[4]C[9],IF(R10C5=5,R[5]C[9],IF(R10C5=6,R16C11,"""")))))"
    Range("B11").FormulaR1C1 = _
        "=IF(R10C5=7,R[6]C[9],IF(R10C5=8,R[7]C[9],IF(R10C5=9,R[8]C[9],IF(R10C5=10,R[9]C[9],IF(R10C5=11,R[10]C[9],"""")))))"
    Range("B12").FormulaR1C1 = _
        "=IF(R10C5=12,R[10]C[9],IF(R10C5=13,R[11]C[9],IF(R10C5=14,R[12]C[9],IF(R10C5=15,R[13]C[9],IF(R10C5=16,R[14]C[9],"""")))))"
    Range("B13").FormulaR1C1 = _
        "=IF(R10C5=17,R[14]C[9],IF(R10C5=18,R[15]C[9],IF(R10C5=19,R[16]C[9],IF(R10C5=20,R[17]C[9],""""))))"
    Range("B14").FormulaR1C1 = _
        "=IF(R10C5=21,R[17]C[9],IF(R10C5=22,R[18]C[9],IF(R10C5=23,R[19]C[9],IF(R10C5=24,R[20]C[9],""""))))"

    Range("B16").FormulaR1C1 = _
        "=IF(R[-6]C<>"""",CONCATENATE(R[-6]C,"" "",R[-4]C[4]),IF(R[-5]C<>"""",CONCATENATE(R[-5]C,"" "",R[-4]C[4]),IF(R[-4]C<>"""",CONCATENATE(R[-4]C,"" "",R[-4]C[4]),IF(R[-3]C<>"""",CONCATENATE(R[-3]C,"" "",R[-4]C[4]),""""))))"
    Range("B17").FormulaR1C1 = "=CONCATENATE(""                       "",R[-1]C)"
    Range("C16").FormulaR1C1 = "=IF(OR(COUNTIF(R15C13:R24C13,R[1]C[-1]),AND(R8C5=3,R53C1>0)),""dupe"","""")"
    Range("D16").FormulaR1C1 = _
        "=IF(AND(SUM(R[-8]C[1],R[-6]C[1])>2,R[-6]C[1]<25,R[-4]C[2]<>"""",R[-2]C[2]<>"""",RC[-1]=""""),""OK"","""")"
    Range("C10").FormulaR1C1 = _
        "=IF(AND(R8C5=3,R10C5=2),RIGHT(R[5]C[10],7),IF(AND(R8C5=3,R10C5=3),RIGHT(R[6]C[10],7),IF(AND(R8C5=3,R10C5=4),RIGHT(R[7]C[10],7),IF(AND(R8C5=3,R10C5=5),RIGHT(R[8]C[10],7),IF(AND(R8C5=3,R10C5=6),RIGHT(R[9]C[10],7),"""")))))"
    Range("C12").FormulaR1C1 = _
        "=IF(AND(R8C5=3,R10C5=7),RIGHT(R[8]C[10],7),IF(AND(R8C5=3,R10C5=8),RIGHT(R[9]C[10],7),IF(AND(R8C5=3,R10C5=9),RIGHT(R[10]C[10],7),IF(AND(R8C5=3,R10C5=10),RIGHT(R[11]C[10],7),IF(AND(R8C5=3,R10C5=11),RIGHT(R[12]C[10],7),"""")))))"
    Range("C14").FormulaR1C1 = "=IF(R8C5=1,"""",MATCH(R[1]C,R[41]C2:R55C11,0)+1)"
    Range("C15").FormulaR1C1 = _
        "=IF(AND(R[-5]C="""",R[-3]C=""""),R[1]C[-1],IF(R[-5]C<>"""",R[-5]C,R[-3]C))"
    Range("A53").FormulaR1C1 = "=IF(AND(R[1]C=1,R[4]C[1]=""""),0,R[1]C)"
    Range("A54").FormulaR1C1 = "=IF(R[3]C[1]="""",0,SUM(RC[1]:RC[10]))"
    Range("B54:K54").FormulaR1C1 = "=IF(R[4]C<>"""",1,"""")"
    Range("A55").FormulaR1C1 = "=10-R[-1]C"
    Range("E8:F8").Select
    Sheets("Calibration").Protect Password:="spike"
    Application.ScreenUpdating = True
End Sub
Sub Caltyp()
Attribute Caltyp.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Caltyp Macro on selection at second drop-down (#2) w/no cal saved
'
    Application.ScreenUpdating = False
    Sheets("Calibration").Unprotect Password:="spike"

    If Range("E8") >= 3 And Range("L7") <> "" Then
        ActiveSheet.Shapes.Range(Array("Rounded Rectangle 1")).Visible = msoFalse
        ActiveSheet.Shapes.Range(Array("Drop Down 2")).Visible = msoFalse
        ActiveSheet.Shapes.Range(Array("Drop Down 30")).Visible = msoTrue
        ActiveSheet.Rows("11:12").Hidden = True
        If Range("E8") = 4 Then
            ActiveSheet.Shapes.Range(Array("Rounded Rectangle 3")).Visible = msoTrue
        End If

    ElseIf Range("E8") < 3 Or Range("L7") = "" Then
    ActiveSheet.Shapes.Range(Array("Drop Down 30")).Visible = msoFalse
    ActiveSheet.Shapes.Range(Array("Drop Down 2")).Visible = msoTrue
    ActiveSheet.Shapes.Range(Array("Rounded Rectangle 1")).Visible = msoTrue
    ActiveSheet.Shapes.Range(Array("Rounded Rectangle 3")).Visible = msoFalse
    ActiveSheet.Rows("11:12").Hidden = False

    End If
    Range("E10").Select
    Sheets("Calibration").Protect Password:="spike"
    Application.ScreenUpdating = True
End Sub

Sub AddAmendCal()

' Copies headers to 55:57 when Units selected, pass dupe to ViewCal

Application.ScreenUpdating = False
Sheets("Calibration").Unprotect Password:="spike"

Range("F12") = UCase(Left(Range("F12"), 7))

If Range("D16") = "OK" And Range("E8") < 4 Then
    If Range("A54") = 0 Then
        Range("B55").Value = Range("B16").Value
        Range("B56").Value = Range("F14").Value
    ElseIf Range("A54") = 1 Then
        Range("C55").Value = Range("B16").Value
        Range("C56").Value = Range("F14").Value
    ElseIf Range("A54") = 2 Then
        Range("D55").Value = Range("B16").Value
        Range("D56").Value = Range("F14").Value
    ElseIf Range("A54") = 3 Then
        Range("E55").Value = Range("B16").Value
        Range("E56").Value = Range("F14").Value
    ElseIf Range("A54") = 4 Then
        Range("F55").Value = Range("B16").Value
        Range("F56").Value = Range("F14").Value
    ElseIf Range("A54") = 5 Then
        Range("G55").Value = Range("B16").Value
        Range("G56").Value = Range("F14").Value
    ElseIf Range("A54") = 6 Then
        Range("H55").Value = Range("B16").Value
        Range("H56").Value = Range("F14").Value
    ElseIf Range("A54") = 7 Then
        Range("I55").Value = Range("B16").Value
        Range("I56").Value = Range("F14").Value
    ElseIf Range("A54") = 8 Then
        Range("J55").Value = Range("B16").Value
        Range("J56").Value = Range("F14").Value
    ElseIf Range("A54") = 9 Then
        Range("K55").Value = Range("B16").Value
        Range("K56").Value = Range("F14").Value
    End If
    Range("E22").Select

ElseIf Range("E8") < 4 Then Application.Run "D.xlsm!ViewCal"
End If
Sheets("Calibration").Protect Password:="spike"
Application.ScreenUpdating = True
End Sub

Sub CalDONE()
' SAVE on glider click
'
 Application.ScreenUpdating = False
 Sheets("Calibration").Unprotect Password:="spike"

If Range("D16") = "OK" Then
    If Range("B16") = Range("B55") Then
        Range("B57").FormulaR1C1 = "=IF(R16C6=2,""Metres"",""Feet"")"
        Range("B58:B78").FormulaR1C1 = "=CONCATENATE(R[-36]C[3],"" "",R[-36]C[4])"
        Range("B55:B78").Value = Range("B55:B78").Value
    ElseIf Range("B16") = Range("C55") Then
        Range("C57").FormulaR1C1 = "=IF(R16C6=2,""Metres"",""Feet"")"
        Range("C58:C78").FormulaR1C1 = "=CONCATENATE(R[-36]C[2],"" "",R[-36]C[3])"
        Range("C55:C78").Value = Range("C55:C78").Value
    ElseIf Range("B16") = Range("D55") Then
        Range("D57").FormulaR1C1 = "=IF(R16C6=2,""Metres"",""Feet"")"
        Range("D58:D78").FormulaR1C1 = "=CONCATENATE(R[-36]C[1],"" "",R[-36]C[2])"
        Range("D55:D78").Value = Range("D55:D78").Value
    ElseIf Range("B16") = Range("E55") Then
        Range("E57").FormulaR1C1 = "=IF(R16C6=2,""Metres"",""Feet"")"
        Range("E58:E78").FormulaR1C1 = "=CONCATENATE(R[-36]C,"" "",R[-36]C[1])"
        Range("E55:E78").Value = Range("E55:E78").Value
    ElseIf Range("B16") = Range("F55") Then
        Range("F57").FormulaR1C1 = "=IF(R16C6=2,""Metres"",""Feet"")"
        Range("F58:F78").FormulaR1C1 = "=CONCATENATE(R[-36]C[-1],"" "",R[-36]C)"
        Range("F55:F78").Value = Range("F55:F78").Value
    ElseIf Range("B16") = Range("G55") Then
        Range("G57").FormulaR1C1 = "=IF(R16C6=2,""Metres"",""Feet"")"
        Range("G58:G78").FormulaR1C1 = "=CONCATENATE(R[-36]C[-2],"" "",R[-36]C[-1])"
        Range("G55:G78").Value = Range("G55:G78").Value
    ElseIf Range("B16") = Range("H55") Then
        Range("H57").FormulaR1C1 = "=IF(R16C6=2,""Metres"",""Feet"")"
        Range("H58:H78").FormulaR1C1 = "=CONCATENATE(R[-36]C[-3],"" "",R[-36]C[-2])"
        Range("H55:H78").Value = Range("H55:H78").Value
    ElseIf Range("B16") = Range("I55") Then
        Range("I57").FormulaR1C1 = "=IF(R16C6=2,""Metres"",""Feet"")"
        Range("I58:I78").FormulaR1C1 = "=CONCATENATE(R[-36]C[-4],"" "",R[-36]C[-3])"
        Range("I55:I78").Value = Range("I55:I78").Value
    ElseIf Range("B16") = Range("J55") Then
        Range("J57").FormulaR1C1 = "=IF(R16C6=2,""Metres"",""Feet"")"
        Range("J58:J78").FormulaR1C1 = "=CONCATENATE(R[-36]C[-5],"" "",R[-36]C[-4])"
        Range("J55:J78").Value = Range("J55:J78").Value
    ElseIf Range("B16") = Range("K55") Then
        Range("K57").FormulaR1C1 = "=IF(R16C6=2,""Metres"",""Feet"")"
        Range("K58:K78").FormulaR1C1 = "=CONCATENATE(R[-36]C[-6],"" "",R[-36]C[-5])"
        Range("K55:K78").Value = Range("K55:K78").Value
    End If

 ElseIf Range("C16") = "dupe" Then

    If Range("C17") = Range("M15") Or Range("C15") = Range("B55") Then
        Range("B56").Value = Range("F14").Value
        Range("B57").FormulaR1C1 = "=IF(R16C6=2,""Metres"",""Feet"")"
        Range("B58:B78").FormulaR1C1 = "=CONCATENATE(R[-36]C5,"" "",R[-36]C6)"
        Range("B55:B78").Value = Range("B55:B78").Value
    ElseIf Range("C17") = Range("M16") Or Range("C15") = Range("C55") Then
        Range("C56").Value = Range("F14").Value
        Range("C57").FormulaR1C1 = "=IF(R16C6=2,""Metres"",""Feet"")"
        Range("C58:C78").FormulaR1C1 = "=CONCATENATE(R[-36]C5,"" "",R[-36]C6)"
        Range("C55:C78").Value = Range("C55:C78").Value
    ElseIf Range("C17") = Range("M17") Or Range("C15") = Range("D55") Then
        Range("D56").Value = Range("F14").Value
        Range("D57").FormulaR1C1 = "=IF(R16C6=2,""Metres"",""Feet"")"
        Range("D58:D78").FormulaR1C1 = "=CONCATENATE(R[-36]C5,"" "",R[-36]C6)"
        Range("D55:D78").Value = Range("D55:D78").Value
    ElseIf Range("C17") = Range("M18") Or Range("C15") = Range("E55") Then
        Range("E56").Value = Range("F14").Value
        Range("E57").FormulaR1C1 = "=IF(R16C6=2,""Metres"",""Feet"")"
        Range("E58:E78").FormulaR1C1 = "=CONCATENATE(R[-36]C5,"" "",R[-36]C6)"
        Range("E55:E78").Value = Range("E55:E78").Value
    ElseIf Range("C17") = Range("M19") Or Range("C15") = Range("F55") Then
        Range("F56").Value = Range("F14").Value
        Range("F57").FormulaR1C1 = "=IF(R16C6=2,""Metres"",""Feet"")"
        Range("F58:F78").FormulaR1C1 = "=CONCATENATE(R[-36]C5,"" "",R[-36]C6)"
        Range("F55:F78").Value = Range("F55:F78").Value
    ElseIf Range("C17") = Range("M20") Or Range("C15") = Range("G55") Then
        Range("G56").Value = Range("F14").Value
        Range("G57").FormulaR1C1 = "=IF(R16C6=2,""Metres"",""Feet"")"
        Range("G58:G78").FormulaR1C1 = "=CONCATENATE(R[-36]C5,"" "",R[-36]C6)"
        Range("G55:G78").Value = Range("G55:G78").Value
    ElseIf Range("C17") = Range("M21") Or Range("C15") = Range("H55") Then
        Range("H56").Value = Range("F14").Value
        Range("H57").FormulaR1C1 = "=IF(R16C6=2,""Metres"",""Feet"")"
        Range("H58:H78").FormulaR1C1 = "=CONCATENATE(R[-36]C5,"" "",R[-36]C6)"
        Range("H55:H78").Value = Range("H55:H78").Value
    ElseIf Range("C17") = Range("M22") Or Range("C15") = Range("I55") Then
        Range("I56").Value = Range("F14").Value
        Range("I57").FormulaR1C1 = "=IF(R16C6=2,""Metres"",""Feet"")"
        Range("I58:I78").FormulaR1C1 = "=CONCATENATE(R[-36]C5,"" "",R[-36]C6)"
        Range("I55:I78").Value = Range("I55:I78").Value
    ElseIf Range("C17") = Range("M23") Or Range("C15") = Range("J55") Then
        Range("J56").Value = Range("F14").Value
        Range("J57").FormulaR1C1 = "=IF(R16C6=2,""Metres"",""Feet"")"
        Range("J58:J78").FormulaR1C1 = "=CONCATENATE(R[-36]C5,"" "",R[-36]C6)"
        Range("J55:J78").Value = Range("J55:J78").Value
    ElseIf Range("C17") = Range("M24") Or Range("C15") = Range("K55") Then
        Range("K56").Value = Range("F14").Value
        Range("K57").FormulaR1C1 = "=IF(R16C6=2,""Metres"",""Feet"")"
        Range("K58:K78").FormulaR1C1 = "=CONCATENATE(R[-36]C5,"" "",R[-36]C6)"
        Range("K55:K78").Value = Range("K55:K78").Value
    End If
 End If

 Application.Run "D.xlsm!Clear"
 Range("E8").Select
 Sheets("Calibration").Protect Password:="spike"
 ActiveWorkbook.Save
 Application.ScreenUpdating = True
End Sub

Sub ViewCal()
'
' ViewCal Macro On selection second drop-down (#30) w/saved Calibration
'
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Sheets("Calibration").Unprotect Password:="spike"

    ActiveSheet.Shapes.Range(Array("Drop Down 2")).Visible = msoFalse
    ActiveSheet.Shapes.Range(Array("Drop Down 30")).Visible = msoTrue
    ActiveSheet.Rows("11:12").Hidden = True
    Range("E8").Select
        If Range("E8") = 2 Then
        Range("E8").Value = 3
        Range("E10").Select
        Range("E10").Value = Range("C14").Value
        End If

    If Range("E10") = 2 Or Range("B17") = Range("M15") Then
        Range("F14").Value = Range("B56").Value
        If Range("B57") = "Metres" Then
            Range("F16").Value = 2
            Else: Range("F16").Value = 3
        End If
        Range("B58:B78").TextToColumns Destination:=Range("E22"), DataType:=xlDelimited, _
            TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=True, _
            Semicolon:=False, Comma:=False, Space:=True, Other:=False, FieldInfo _
            :=Array(Array(1, 1), Array(2, 1)), TrailingMinusNumbers:=True
    ElseIf Range("E10") = 3 Or Range("B17") = Range("M16") Then
        Range("F14").Value = Range("C56").Value
        If Range("C57") = "Metres" Then
            Range("F16").Value = 2
            Else: Range("F16").Value = 3
        End If
        Range("C58:C78").TextToColumns Destination:=Range("E22"), DataType:=xlDelimited, _
            TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=True, _
            Semicolon:=False, Comma:=False, Space:=True, Other:=False, FieldInfo _
            :=Array(Array(1, 1), Array(2, 1)), TrailingMinusNumbers:=True
    ElseIf Range("E10") = 4 Or Range("B17") = Range("M17") Then
        Range("F14").Value = Range("D56").Value
        If Range("D57") = "Metres" Then
            Range("F16").Value = 2
            Else: Range("F16").Value = 3
        End If
        Range("D58:D78").TextToColumns Destination:=Range("E22"), DataType:=xlDelimited, _
            TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=True, _
            Semicolon:=False, Comma:=False, Space:=True, Other:=False, FieldInfo _
            :=Array(Array(1, 1), Array(2, 1)), TrailingMinusNumbers:=True
    ElseIf Range("E10") = 5 Or Range("B17") = Range("M18") Then
        Range("F14").Value = Range("E56").Value
        If Range("E57") = "Metres" Then
            Range("F16").Value = 2
            Else: Range("F16").Value = 3
        End If
        Range("E58:E78").TextToColumns Destination:=Range("E22"), DataType:=xlDelimited, _
            TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=True, _
            Semicolon:=False, Comma:=False, Space:=True, Other:=False, FieldInfo _
            :=Array(Array(1, 1), Array(2, 1)), TrailingMinusNumbers:=True
    ElseIf Range("E10") = 6 Or Range("B17") = Range("M19") Then
        Range("F14").Value = Range("F56").Value
        If Range("F57") = "Metres" Then
            Range("F16").Value = 2
            Else: Range("F16").Value = 3
        End If
        Range("F58:F78").TextToColumns Destination:=Range("E22"), DataType:=xlDelimited, _
            TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=True, _
            Semicolon:=False, Comma:=False, Space:=True, Other:=False, FieldInfo _
            :=Array(Array(1, 1), Array(2, 1)), TrailingMinusNumbers:=True
    ElseIf Range("E10") = 7 Or Range("B17") = Range("M20") Then
        Range("F14").Value = Range("G56").Value
        If Range("G57") = "Metres" Then
            Range("F16").Value = 2
            Else: Range("F16").Value = 3
        End If
        Range("G58:G78").TextToColumns Destination:=Range("E22"), DataType:=xlDelimited, _
            TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=True, _
            Semicolon:=False, Comma:=False, Space:=True, Other:=False, FieldInfo _
            :=Array(Array(1, 1), Array(2, 1)), TrailingMinusNumbers:=True
    ElseIf Range("E10") = 8 Or Range("B17") = Range("M21") Then
        Range("F14").Value = Range("H56").Value
        If Range("H57") = "Metres" Then
            Range("F16").Value = 2
            Else: Range("F16").Value = 3
        End If
        Range("H58:H78").TextToColumns Destination:=Range("E22"), DataType:=xlDelimited, _
            TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=True, _
            Semicolon:=False, Comma:=False, Space:=True, Other:=False, FieldInfo _
            :=Array(Array(1, 1), Array(2, 1)), TrailingMinusNumbers:=True
    ElseIf Range("E10") = 9 Or Range("B17") = Range("M22") Then
        Range("F14").Value = Range("I56").Value
        If Range("I57") = "Metres" Then
            Range("F16").Value = 2
            Else: Range("F16").Value = 3
        End If
        Range("I58:I78").TextToColumns Destination:=Range("E22"), DataType:=xlDelimited, _
            TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=True, _
            Semicolon:=False, Comma:=False, Space:=True, Other:=False, FieldInfo _
            :=Array(Array(1, 1), Array(2, 1)), TrailingMinusNumbers:=True
    ElseIf Range("E10") = 10 Or Range("B17") = Range("M23") Then
        Range("F14").Value = Range("J56").Value
        If Range("J57") = "Metres" Then
            Range("F16").Value = 2
            Else: Range("F16").Value = 3
        End If
        Range("J58:J78").TextToColumns Destination:=Range("E22"), DataType:=xlDelimited, _
            TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=True, _
            Semicolon:=False, Comma:=False, Space:=True, Other:=False, FieldInfo _
            :=Array(Array(1, 1), Array(2, 1)), TrailingMinusNumbers:=True
    ElseIf Range("E10") = 11 Or Range("B17") = Range("M24") Then
        Range("F14").Value = Range("K56").Value
        If Range("K57") = "Metres" Then
            Range("F16").Value = 2
            Else: Range("F16").Value = 3
        End If
        Range("K58:K78").TextToColumns Destination:=Range("E22"), DataType:=xlDelimited, _
            TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=True, _
            Semicolon:=False, Comma:=False, Space:=True, Other:=False, FieldInfo _
            :=Array(Array(1, 1), Array(2, 1)), TrailingMinusNumbers:=True
    End If

  Range("F14").Select
  Sheets("Calibration").Protect Password:="spike"
  Application.DisplayAlerts = True
  Application.ScreenUpdating = True
End Sub

Sub DeleteCal()
'
' Macro12 Macro need button!
'
Application.ScreenUpdating = False
Sheets("Calibration").Unprotect Password:="spike"
If Range("E8") = 4 Then
    If Range("E10") = 2 Then
        Range("B55:B78").Delete Shift:=xlToLeft
    ElseIf Range("E10") = 3 Then
        Range("C55:C78").Delete Shift:=xlToLeft
    ElseIf Range("E10") = 4 Then
        Range("D55:D78").Delete Shift:=xlToLeft
    ElseIf Range("E10") = 5 Then
        Range("E55:E78").Delete Shift:=xlToLeft
    ElseIf Range("E10") = 6 Then
        Range("F55:F78").Delete Shift:=xlToLeft
    ElseIf Range("E10") = 7 Then
        Range("G55:G78").Delete Shift:=xlToLeft
    ElseIf Range("E10") = 8 Then
        Range("H55:H78").Delete Shift:=xlToLeft
    ElseIf Range("E10") = 9 Then
        Range("I55:I78").Delete Shift:=xlToLeft
    ElseIf Range("E10") = 10 Then
        Range("J55:J78").Delete Shift:=xlToLeft
    ElseIf Range("E10") = 11 Then
        Range("K55:K78").Delete Shift:=xlToLeft
    End If

Application.Run "D.xlsm!Clear"
End If
Range("E8").Select
Sheets("Calibration").Protect Password:="spike"
Application.ScreenUpdating = True
End Sub

Sub Clear()

Application.ScreenUpdating = False
Sheets("Calibration").Unprotect Password:="spike"
ActiveSheet.Rows("11:12").Hidden = False
ActiveSheet.Shapes.Range(Array("Drop Down 30")).Visible = msoFalse
ActiveSheet.Shapes.Range(Array("Drop Down 2")).Visible = msoTrue
Range("F12,F14,E22:F42").ClearContents
Range("E8,E10,F16").Value = 1
Range("C14").FormulaR1C1 = "=IF(R8C5=1,"""",MATCH(R[1]C,R[41]C2:R55C11,0)+1)"
Range("C16").FormulaR1C1 = "=IF(OR(COUNTIF(R15C13:R24C13,R[1]C[-1]),AND(R8C5=3,R53C1>0)),""dupe"","""")"
Range("A53").FormulaR1C1 = "=IF(AND(R[1]C=1,R[4]C[1]=""""),0,R[1]C)"
Range("A54").FormulaR1C1 = "=IF(R[1]C[1]="""",0,SUM(RC[1]:RC[10]))"
Range("B54:K54").FormulaR1C1 = "=IF(R[4]C<>"""",1,""X"")"
If Range("B54") = "X" Then Range("B55:B78").ClearContents
If Range("C54") = "X" Then Range("C55:C78").ClearContents
If Range("D54") = "X" Then Range("D55:D78").ClearContents
If Range("E54") = "X" Then Range("E55:E78").ClearContents
If Range("F54") = "X" Then Range("F55:F78").ClearContents
If Range("G54") = "X" Then Range("G55:G78").ClearContents
If Range("H54") = "X" Then Range("H55:H78").ClearContents
If Range("I54") = "X" Then Range("I55:I78").ClearContents
If Range("G54") = "X" Then Range("J55:J78").ClearContents
If Range("H54") = "X" Then Range("K55:K78").ClearContents
Range("B55:K78").Select
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
Range("M14").FormulaR1C1 = "=IF(R8C5=4,""                   FR TO DELETE"",""               FR TO VIEW / UPDATE"")"
Range("M15").FormulaR1C1 = "=IF(R58C2<>"""",CONCATENATE(""                       "",R55C2),"""")"
Range("M16").FormulaR1C1 = "=IF(R58C3<>"""",CONCATENATE(""                       "",R55C3),"""")"
Range("M17").FormulaR1C1 = "=IF(R58C4<>"""",CONCATENATE(""                       "",R55C4),"""")"
Range("M18").FormulaR1C1 = "=IF(R58C5<>"""",CONCATENATE(""                       "",R55C5),"""")"
Range("M19").FormulaR1C1 = "=IF(R58C6<>"""",CONCATENATE(""                       "",R55C6),"""")"
Range("M20").FormulaR1C1 = "=IF(R58C7<>"""",CONCATENATE(""                       "",R55C7),"""")"
Range("M21").FormulaR1C1 = "=IF(R58C8<>"""",CONCATENATE(""                       "",R55C8),"""")"
Range("M22").FormulaR1C1 = "=IF(R58C9<>"""",CONCATENATE(""                       "",R55C9),"""")"
Range("M23").FormulaR1C1 = "=IF(R58C10<>"""",CONCATENATE(""                       "",R55C10),"""")"
Range("M24").FormulaR1C1 = "=IF(R58C11<>"""",CONCATENATE(""                       "",R55C11),"""")"
ActiveSheet.Shapes.Range(Array("Rounded Rectangle 3")).Visible = msoFalse
ActiveSheet.Shapes.Range(Array("Rounded Rectangle 1")).Visible = msoTrue
Range("A1").Select
Sheets("Calibration").Protect Password:="spike"
ActiveWorkbook.Save
Application.ScreenUpdating = True

End Sub
