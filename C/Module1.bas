' VBA Module: Calibration and Task Validation
' Purpose: Manages calibration visibility controls and task validation calculations.
' Handles dropdown change events, calibration sheet management, and complex flight
' task calculations including trigonometric computations for soaring competitions.

Option Explicit
Sub DropDown1_Change()
Application.ScreenUpdating = False
Workbooks("C.xlsm").Unprotect Password:="spike"
If Range("H1") = "" And Range("G2") <> 3 Then
        Sheets("Calibration").Visible = True
        Application.ScreenUpdating = True
ElseIf Range("H1") <> "" Or Range("G2") = 3 Then
        Sheets("Calibration").Visible = False
        Application.ScreenUpdating = True
ElseIf Range("H1") <> "" And Range("G2") = 1 Then
        Exit Sub
End If
Workbooks("C.xlsm").Protect Password:="spike"
Sheets("Verify Task").Activate
Application.Calculate

End Sub
Sub CalcC()
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
 Sheets("YDWK2").Range("E1505").FormulaR1C1 = _
        "=R[-34]C[-2]+(1-R[-2]C)*R[-35]C[-2]*R[-14]C*(ACOS(R[-19]C)+R[-2]C*SIN(ACOS(R[-19]C))*(R[-12]C+R[-2]C*R[-19]C*(-1+2*R[-12]C*R[-12]C)))"
 Sheets("YDWK2").Calculate
 Sheets("TPOrder").Calculate
 Sheets("Summary").Calculate
 Sheets("Verify Task").Calculate
'Sheets("Summary").Calculate
'Sheets("YDWK2").Calculate
'Sheets("TPOrder").Calculate
'Dim newHour As Variant
'Dim newMinute As Variant
'Dim newSecond As Variant
'Dim waitTime As Variant
'newHour = Hour(Now())
'newMinute = Minute(Now())
'newSecond = Second(Now()) + 2
'waitTime = TimeSerial(newHour, newMinute, newSecond)
'Application.Wait waitTime
'Application.Calculation = xlCalculationAutomatic
Sheets("Verify Task").Unprotect Password:="spike"
Range("C20").FormulaR1C1 = "=IF(AND(SUMMARY!R33C11=""User"",R12C6<>"""",OR(AND(R14C5<>""N/A"",R14C6<>""""),AND(R16C5<>""N/A"",R16C6<>""""))),""This Turn Point order cannot be confirmed."","""")"
Range("B25").FormulaR1C1 = "=IF(OR(R12C6="""",AND(R2C7=1,R2C8<>"""",R2C8>42278),AND(R14C5<>""N/A"",R14C6=""""),AND(R16C5<>""N/A"",R16C6=""""),R20C3<>""""),"""",""            Click on the glider to continue"")"
If Range("C20") <> "" Then
    ActiveSheet.Shapes("Rounded Rectangle 7").Visible = True
    ActiveSheet.Shapes("Rounded Rectangle 3").Visible = False
ElseIf Range("C20") = "" Then
    ActiveSheet.Shapes("Rounded Rectangle 7").Visible = False
    ActiveSheet.Shapes("Rounded Rectangle 3").Visible = True
End If
Sheets("Verify Task").Protect Password:="spike"
Application.ScreenUpdating = True
'Application.Calculation = xlCalculationManual
End Sub
Sub ReTry()
'
' JLR 7/9/2015 Provide for re-try via Rounded Rectangle 7
'
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Range("F12,F14,F16").ClearContents
    Sheets("YDWK2").Range("E1505").FormulaR1C1 = _
        "=R[-34]C[-2]+(1-R[-2]C)*R[-35]C[-2]*R[-14]C*(ACOS(R[-19]C)+R[-2]C*SIN(ACOS(R[-19]C))*(R[-12]C+R[-2]C*R[-19]C*(-1+2*R[-12]C*R[-12]C)))"
    Sheets("YDWK2").Calculate
    Sheets("TPOrder").Calculate
    Sheets("SUMMARY").Calculate
    Sheets("Verify Task").Calculate
    ActiveSheet.Shapes("Rounded Rectangle 7").Visible = False
    ActiveSheet.Shapes("Rounded Rectangle 3").Visible = True
    Range("F12").Select
    Application.ScreenUpdating = True
End Sub
Sub GotoCC()
Application.ScreenUpdating = True

If Range("B25") = "            Click on the glider to continue" Then
    Workbooks("C.xlsm").Activate
    Workbooks("C.xlsm").Unprotect Password:="spike"
    Application.DisplayFullScreen = True
    ActiveWindow.DisplayWorkbookTabs = False
    If Range("F18") > 0 Then
        Sheets("Verify Task").Unprotect Password:="spike"
        Range("B27:D27").Value = "           Final Calculations in Progress "
        Range("F27:H27").Value = "5 to 10 seconds typical"
    End If
    Sheets("Verify Task").Protect Password:="spike"
    Application.ScreenUpdating = False

    Application.Run "C.xlsm!YDWK2"
    Workbooks("C.xlsm").Unprotect Password:="spike"
    Sheets("Worksheet").Activate
    Sheets("Worksheet").Unprotect Password:="spike"
        'Next 6 lines commented out 9/17/15 w/ withdrawal of penalties for <=100 dist w/excess LoH
        'If Range("N3") >= 42278 Then
            'Range("J44").Value = ">= 10/1/2015"
            'Range("I46:I61").FormulaR1C1 = _
                "=IF(OR(RC[-3]=0,AND(R1C17="""",RC[-3]>100,RC[-1]<=1000),AND(R1C17=""PR"",RC[-3]>100,RC[-1]<=900),AND(R1C17="""",RC[-3]<=100,RC[-1]<=RC[-3]*10),AND(R1C17=""PR"",RC[-3]<=100,RC[-1]<=(RC[-3]*10)-100)),0,IF(AND(R1C17="""",RC[-3]>100,RC[-1]>1000),RC[-1]-1000,IF(AND(R1C17=""PR"",RC[-3]>100,RC[-1]>900),RC[-1]-900,IF(AND(R1C17="""",RC[-3]<=100,RC[-1]>RC[-3]*10),RC[-1]-(RC[-3]*10),IF(AND(R1C17=""PR"",RC[-3]<=100,RC[-1]>(RC[-3]*10)-100),RC[-1]-((RC[-3]*10)-100))))))"
            'Range("J46:J61").FormulaR1C1 = "=IF(OR(RC[-4]=0,RC[-1]<=0),0,IF(RC[-1]>=10*RC[-4],RC[-4],RC[-1]/10))"
        'ElseIf Range("N3") >= 41183 And Range("N3") < 42278 Then
        If Range("N3") >= 41183 Then
            Range("J44").Value = ">= 10/1/2012"
            Range("I46:I61").FormulaR1C1 = _
                "=IF(OR(RC[-3]=0,AND(R1C17<>""PR"",RC[-3]<=100,RC[-1]<=10*RC[-3]),AND(R1C17<>""PR"",RC[-3]>100,RC[-1]<=1000),AND(R1C17=""PR"",RC[-3]>100,RC[-1]<=900),AND(R1C17=""PR"",RC[-3]<=100,RC[-1]<=(10*RC[-3])-100)),0,IF(AND(R1C17<>""PR"",RC[-3]<=100),RC[-1]-0.1*RC[-3]*100,IF(AND(R1C17<>""PR"",RC[-3]>100),RC[-1]-1000,IF(AND(R1C17=""PR"",RC[-3]>100,RC[-1]>900),RC[-1]-900,RC[-1]-((RC[-3]*10)-100)))))"
            Range("J46:J61").FormulaR1C1 = _
            "=IF(OR(AND(R1C17<>""PR"",RC[-4]<=100,RC[-2]>0.01*RC[-4]*1000),AND(R1C17=""PR"",RC[-4]<=100,RC[-2]>(RC[-4]*10)-100)),RC[-4],IF(RC[-1]>0,-1*(RC[-1]*-100/1000),0))"
        ElseIf Range("N3") < 41183 Then
            Range("J44").Value = "< 10/1/2012"
            Range("I46:I61").FormulaR1C1 = "=IF(OR(RC[-3]=0,AND(RC[-3]<=100,RC[-1]<=10*RC[-3]),AND(RC[-3]>100,RC[-1]<=1000)),0,IF(RC[-3]<=100,RC[-1]-0.1*RC[-3]*100,-1*(1000-RC[-1])))"
            Range("J46:J61").FormulaR1C1 = "=IF(AND(RC[-4]<=100,RC[-2]>10*RC[-4]),RC[-4],IF(RC[-1]>0,-1*(RC[-1]*-100/1000),0))"
        End If
    Range("A1").Select
    Sheets("Worksheet").Calculate
    Sheets("Worksheet").Protect Password:="spike"
    Sheets("Worksheet").Visible = False

    Application.DisplayAlerts = False
    If Range("E10") <> "No Turn Points Declared" And Range("F18") > 0 Then
        Application.Run "C.xlsm!POfixes"
    'End If
    Sheets("Sheet11").Visible = False
    Sheets("Verify Task").Unprotect Password:="spike"
    Sheets("Verify Task").Range("C20:E20,B25:D25,B27:D27,F27:H27").ClearContents
    Sheets("Verify Task").Shapes("Rounded Rectangle 3").Visible = True
    Sheets("Verify Task").Shapes("Rounded Rectangle 7").Visible = False
    ElseIf Range("E10") = "No Turn Points Declared" And Range("G2") > 1 Then
        Sheets("Verify Task").Shapes("Rounded Rectangle 3").Visible = False
    End If
    Sheets("Verify Task").Protect Password:="spike"
    Sheets("Claim Check").Visible = True
    Sheets("CLAIM CHECK").Select
    Sheets("Claim Check").Unprotect Password:="spike"
    ActiveWindow.DisplayVerticalScrollBar = True
    ActiveWindow.DisplayWorkbookTabs = True
    Sheets("Claim Check").Calculate

    Range("A1:J31").Select
    ActiveWindow.Zoom = True
    Range("B2").Select
    Sheets("Claim Check").Protect Password:="spike"
 End If

    Workbooks("C.xlsm").Protect Password:="spike"
    Application.ScreenUpdating = True

End Sub
Sub back2CC()

Application.ScreenUpdating = False
Workbooks("C.xlsm").Unprotect Password:="spike"
Sheets("Worksheet").Unprotect Password:="spike"
Range("B1").Select
Sheets("Worksheet").Protect Password:="spike"
Sheets("Claim Check").Select
Sheets("Claim Check").Protect Password:="spike"
Sheets("Worksheet").Visible = False
ActiveWindow.DisplayVerticalScrollBar = True
Workbooks("C.xlsm").Protect Password:="spike"
Application.ScreenUpdating = True

End Sub
Sub GO2WKS()

Application.ScreenUpdating = False
Workbooks("C.xlsm").Unprotect Password:="spike"
Sheets("Worksheet").Visible = True
Sheets("Worksheet").Select
ActiveSheet.Unprotect Password:="spike"
Range("A1:P30").Select
ActiveWindow.Zoom = True
ActiveWindow.DisplayVerticalScrollBar = True
ActiveSheet.Protect Password:="spike"
Workbooks("C.xlsm").Protect Password:="spike"
Application.ScreenUpdating = True

End Sub
Sub CalData()
If Range("E5") = 1 Then
Range("E5").Select
Else: Range("G6").Select
End If
End Sub
Sub OpenForm()
    Application.ScreenUpdating = False
    Workbooks("C.xlsm").Unprotect Password:="spike"
    Sheets("PRINT THIS!").Visible = True
    Sheets("PRINT THIS!").Select
    Sheets("PRINT THIS!").Unprotect Password:="spike"
    ActiveWindow.DisplayVerticalScrollBar = True
    Sheets("PRINT THIS!").Protect Password:="spike"
    Workbooks("C.xlsm").Protect Password:="spike"
    Application.ScreenUpdating = True
End Sub

Sub CloseAll()
On Error Resume Next
    Workbooks("A.xlsm").Activate
  If Err = 0 Then
    Application.DisplayAlerts = False
    Workbooks("A.xlsm").Close
  End If
    Workbooks("C.xlsm").Saved = True
    MsgBox "Thanks for using Claim Check! Click OK to exit"
    Application.Quit

End Sub
Sub ClearC()

    Application.ScreenUpdating = False
    Workbooks("C.xlsm").Unprotect Password:="spike"
    Sheets("Print This!").Visible = True
    Sheets("Print This!").Select
    Range("G16").FormulaR1C1 = "1"
    Range("A1").Select
    Sheets("Print This!").Protect Password:="spike"
    Sheets("Print This!").Visible = False
    Sheets("Worksheet").Visible = False
    Sheets("Calibration").Visible = True
    Sheets("Calibration").Select
    ActiveSheet.Unprotect Password:="spike"
    Range("E5").FormulaR1C1 = "1"
    Range("A1:I23").Select
    ActiveWindow.Zoom = True
    Range("G7:H22").ClearContents
    Range("E5").Select
    ActiveSheet.Protect Password:="spike"
    Sheets("Calibration").Visible = False
    Sheets("Claim Check").Visible = False
    Sheets("TPOrder").Visible = True
    Sheets("TPOrder").Select
    Sheets("TPOrder").Unprotect Password:="spike"
    Range("A11:GC10025").Clear
    Range("A3:A7,C3:C5,L3:L4,O4:S4,P3:P4,R2,T3:T4,V2,Y3,Y4,O6:S6,O8:S8,O10:S10,W4,W6,W8,W10,Z3,Z4,Z9,AD2,AE10,AU5,AU8,AX7,AU10,AV3,AX3,BS5").ClearContents
    Range("AV9:BS9,CG9,CW9,ES9:FQ9,AY10,BU4:BX4,CF4:CH4,CK4:CN4,CV4:CX4,DA4:DD4,DL4:DN4,DR3:DT3,DV4,EB2,EB4,EC10,EO4:EQ4,EY10").ClearContents
    Range("X5:AC7,AD4:AD6,AI5:AI8,AO4:AO6,AQ5:AQ6,AR4:AT4,AR6:AT6,AV7:AW7").ClearContents
    Range("AZ5:AZ6,BD4:BD5,BJ5:BJ6,BW5:BW9,CC4:CC6,CE5:CE6,CH6:CI7,CI1:CI3,CM5:CM9,CS4:CS6,CU5:CU6,CX6:CY7,CY1:CY3,DC5:DC9,DI4:DI6,DK5:DK6").ClearContents
    Range("DN6:DO7,DO1:DO3,DT1:DY2,DQ8:DZ8,EB4:ED6,EF5:EF8,EM4:EM6,EO5:EO6,ER6:ES7,EU3:EW4,EZ5:EZ6,FD4:FD5,FJ5:FJ6,FR9,FS15:FW83").ClearContents
    Sheets("TPOrder").Protect Password:="spike"
    Sheets("TPOrder").Visible = False
    Sheets("YDWK1").Visible = True
    Sheets("YDWK1").Activate
    Sheets("YDWK1").Unprotect Password:="spike"
    Range("J1:L1").ClearContents
    Sheets("YDWK1").Protect Password:="spike"
    Sheets("YDWK1").Visible = False
    Sheets("B").Visible = True
    Sheets("B").Cells.Clear
    Sheets("B").Visible = False
    Sheets("Sheet11").Visible = True
    Sheets("Sheet11").Cells.Clear
    Sheets("Sheet11").Visible = False
    Sheets("Free Me").Visible = True
    Sheets("Free Me").Activate
    Sheets("Free Me").Unprotect Password:="spike"
    Range("C10:M11,C14:M17,C20:M24,G26,C27:M33").ClearContents
    Sheets("Free Me").Protect Password:="spike"
    Sheets("Verify Task").Activate
    Sheets("Verify Task").Unprotect Password:="spike"
    Range("G2").Value = "1"
    Range("F12,F14,F16,C20:E20,B25:D25,B27:D27,F27:H27").ClearContents
    Range("F12,F14,F16").Locked = True
    Range("F12,F14,F16").FormulaHidden = True
    ActiveSheet.Shapes("TextBox 2").Visible = False
    ActiveSheet.Shapes("Rounded Rectangle 3").Visible = False
    ActiveSheet.Shapes("Rounded Rectangle 7").Visible = False
    Range("A1").Select
    Sheets("Verify Task").Protect Password:="spike"
    Application.Run ("C.xlsm!YDWK2")
    Workbooks("C.xlsm").Protect Password:="spike"
    Application.ScreenUpdating = True
    'Workbooks("C.xlsm").Save
End Sub

Sub YDWK2()
'
' JL Ruprecht 3/15/15
'
    Workbooks("C.xlsm").Unprotect Password:="spike"
    Sheets("YDWK2").Visible = True
    Sheets("YDWK2").Select

    Application.Calculation = xlCalculationManual
    Range("E55,E127,E198,E270,E342,E414,E486,E558,E630").FormulaR1C1 = _
        "=R[-34]C[-2]+(1-R[-2]C)*R[-35]C[-2]*R[-14]C*(ACOS(R[-19]C)+R[-2]C*SIN(ACOS(R[-19]C))*(R[-12]C+R[-2]C*R[-19]C*(-1+2*R[-12]C*R[-12]C)))"

    Range("E702,E774,E848,E922,O922,Y922,E995,O996,Y996").FormulaR1C1 = _
        "=R[-34]C[-2]+(1-R[-2]C)*R[-35]C[-2]*R[-14]C*(ACOS(R[-19]C)+R[-2]C*SIN(ACOS(R[-19]C))*(R[-12]C+R[-2]C*R[-19]C*(-1+2*R[-12]C*R[-12]C)))"

    Range("E1069,E1142,E1214,E1287,E1359,E1432,E1505,E1578,O1578,E1651").FormulaR1C1 = _
        "=R[-34]C[-2]+(1-R[-2]C)*R[-35]C[-2]*R[-14]C*(ACOS(R[-19]C)+R[-2]C*SIN(ACOS(R[-19]C))*(R[-12]C+R[-2]C*R[-19]C*(-1+2*R[-12]C*R[-12]C)))"
    Application.Calculation = xlCalculationAutomatic

    Range("A1").Select
    Sheets("YDWK2").Visible = False
    Workbooks("C.xlsm").Protect Password:="spike"

End Sub
