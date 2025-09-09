' VBA Module: Event Processing and Validation
' Purpose: Manages event processing workflows and validation procedures.
' Handles user interface transitions and data entry validation checks.

Option Explicit
Dim linkArray As Variant
Dim newHour As Variant
Dim newMinute As Variant
Dim newSecond As Variant
Dim waitTime As Variant
Sub NewEV()
Attribute NewEV.VB_ProcData.VB_Invoke_Func = " \n14"
'
' JLR 4/8/2014 TEST
'
'ActiveWorkbook.Unprotect Password:="spike"

If Range("C24") = " Click on the glider to continue" Then
    Application.ScreenUpdating = True
    Workbooks("A.xlsm").Unprotect Password:="spike"
    Sheets("Logo").Visible = True
    Sheets("Logo").Select
    Application.DisplayFullScreen = True
    ActiveWindow.DisplayWorkbookTabs = False
    newHour = Hour(Now())
    newMinute = Minute(Now())
    newSecond = Second(Now()) + 2
    waitTime = TimeSerial(newHour, newMinute, newSecond)
    Application.Wait waitTime
    Sheets("Data Entry Check").Visible = False
    Application.ScreenUpdating = False
    Application.Cursor = xlWait

    On Error Resume Next
    Workbooks("D.xlsm").Activate
  If Err = 0 Then
    Application.DisplayAlerts = False
    Workbooks("D.xlsm").Close
  Else: Application.DisplayAlerts = False
  End If

    Application.Run ("Ab.xlsm!NEWHIlo")
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    Workbooks("A.xlsm").Activate
    'linkArray = ActiveWorkbook.LinkSources(xlExcelLinks)
    'ActiveWorkbook.OpenLinks linkArray(2)
    'Added 12/5/14 for XL2013 revised 9/25/15
    Workbooks.Open Filename:=ThisWorkbook.Path & "\C.xlsm"
    Windows("C.xlsm").Activate
    ActiveWindow.WindowState = xlMinimized

    Workbooks("C.xlsm").Unprotect Password:="spike"
    Workbooks("C.xlsm").Sheets("B").Range("A1:S30").Value = Workbooks("Ab.xlsm").Sheets("PRS").Range("A1:S30").Value
    Workbooks("C.xlsm").Sheets("B").Range("A39:B64").Value = Workbooks("Ab.xlsm").Sheets("PRS").Range("A39:B64").Value
    Workbooks("C.xlsm").Sheets("TPOrder").Unprotect Password:="spike"
    Workbooks("C.xlsm").Sheets("TPOrder").Range("A11:G10010").Value = Workbooks("Ab.xlsm").Sheets("Sheet2").Range("A1:G10000").Value
    Workbooks("C.xlsm").Sheets("TPOrder").Protect Password:="spike"
    Workbooks("C.xlsm").Sheets("Sheet11").Range("A9:E60009").Value = Workbooks("Ab.xlsm").Sheets("Sheet2").Range("I9:M60009").Value

    Workbooks("Ab.xlsm").Activate
    Application.DisplayAlerts = False
    Workbooks("Ab.xlsm").Close
    Workbooks("C.xlsm").Activate
    Sheets("Free Me").Visible = False
    Sheets("Worksheet").Visible = True
    Sheets("Worksheet").Activate
    Sheets("Worksheet").Unprotect Password:="spike"
    Range("A11:A71").EntireRow.Hidden = False
    Range("A72:A102").EntireRow.Hidden = True
    ActiveSheet.Shapes("Rectangle 2").Visible = True
    ActiveSheet.Shapes("Rectangle 4").Visible = True
    Range("A1").Select
    Sheets("Worksheet").Protect Password:="spike"
    Sheets("Worksheet").Visible = False

    Sheets("Claim Check").Visible = True
    Sheets("Claim Check").Activate
    Sheets("Claim Check").Unprotect Password:="spike"
        Range("A17:A19").EntireRow.Hidden = False
        Range("B22:C22").Merge
        Range("B22:C22").FormulaR1C1 = "Straight Distance, "
        Range("D22").FormulaR1C1 = "=Worksheet!R63C12"
        Range("B23:D23").FormulaR1C1 = "=IF(OR(SUMMARY!R14C1=""Free"",SUMMARY!R30C8>0,Worksheet!R59C6=0),""Straight Distance to a Goal N/A   "",""Straight Distance to a Goal,  Start to Finish   "")"
        Range("B24:C24").Merge
        Range("B24:C24").FormulaR1C1 = "=IF(SUMMARY!R[-10]C[-1]=""Free""=FALSE,""Distance Via Up to 3 TP,"",""Free 3-Turn Pt Distance,"")"
        Range("D24").FormulaR1C1 = "=IF(AND(R6C5<42278,SUMMARY!R[-10]C[1]<10),""Invalid   "",IF(B!R[-10]C[-2]=""None"",""No Turn Pts   "",Worksheet!R64C12))"
        Range("B25:D25").FormulaR1C1 = _
        "=IF(Worksheet!R60C3=""No Closed Course"",""No Closed Course   "",IF(AND(SUMMARY!R14C1<>""Free"",R[-6]C[2]>0,R[-6]C[4]=""""),""Out & Return, Start to Finish   "",IF(AND(SUMMARY!R14C1=""Free"",R[-6]C[2]>0,R[-6]C[4]=""""),""Free Out &Return, Start to Finish   "",IF(AND(SUMMARY!R14C1=""Free"",R[-6]C[4]>0,R[-6]C[6]=""""),""Free 2-Turn Point Triangle, Start to Finish   "",IF(AND(SUMMARY!R14C1<>""Free"",R[-6]C[4]>0,R[-6]C[6]=""""),""2-Turn Point Triangle, Start to Finish   "",IF(AND(SUMMARY!R14C1<>""Free"",R[-6]C[6]>0,R[-6]C[4]>0,Worksheet!R60C5>0),""3-Turn Point Triangle, Start to Finish   "",IF(AND(Worksheet!R60C5>0,R[-6]C[2]>0,R[-6]C[4]>0,R[-6]C[6]>0),""Free 3-Turn Point Triangle, Start to Finish   "",""No Closed Course   "")))))))"
    Sheets("Claim Check").Protect Password:="spike"
    Sheets("Claim Check").Visible = False
    Application.Run ("C.xlsm!Mega")
    Sheets("Sheet11").Visible = False
    Application.DisplayFullScreen = True
    Workbooks("C.xlsm").Protect Password:="spike"

ElseIf Range("C24") = "Click on the glider to continue" Then
    Application.ScreenUpdating = True
    Workbooks("A.xlsm").Unprotect Password:="spike"
    Sheets("Logo").Visible = True
    Sheets("Logo").Select
    Application.DisplayFullScreen = True
    ActiveWindow.DisplayWorkbookTabs = False
    newHour = Hour(Now())
    newMinute = Minute(Now())
    newSecond = Second(Now()) + 2
    waitTime = TimeSerial(newHour, newMinute, newSecond)
    Application.Wait waitTime
    Sheets("Data Entry Check").Visible = False
    Application.ScreenUpdating = False
    Application.Cursor = xlWait

    On Error Resume Next
    Workbooks("D.xlsm").Activate
  If Err = 0 Then
    Application.DisplayAlerts = False
    Workbooks("D.xlsm").Close
  Else: Application.DisplayAlerts = False
  End If

    Application.Run ("Ab.xlsm!NEWHIlo")
    Application.Calculation = xlCalculationManual
    Workbooks("A.xlsm").Activate
    Application.DisplayAlerts = False
    'linkArray = ActiveWorkbook.LinkSources(xlExcelLinks)
    'ActiveWorkbook.OpenLinks linkArray(3)
    Workbooks.Open Filename:=ThisWorkbook.Path & "\F.xlsm"
    Workbooks("F.xlsm").Activate
    Workbooks("F.xlsm").Unprotect Password:="spike"
    Workbooks("F.xlsm").Sheets("B").Range("A1:L30").Value = Workbooks("Ab.xlsm").Sheets("PRS").Range("A1:L30").Value
    Workbooks("F.xlsm").Sheets("B").Range("A39:B64").Value = Workbooks("Ab.xlsm").Sheets("PRS").Range("A39:B64").Value
    Workbooks("F.xlsm").Sheets("Sheet2").Unprotect Password:="spike"
    Workbooks("F.xlsm").Sheets("Sheet2").Range("A9:E60008").Value = Workbooks("Ab.xlsm").Sheets("Sheet2").Range("I9:M60008").Value

    Workbooks("Ab.xlsm").Activate
    Application.DisplayAlerts = False
    Workbooks("Ab.xlsm").Close
    Workbooks("F.xlsm").Activate
    Sheets("Sheet2").Select
    Application.Run ("F.xlsm!ASelect")

    Workbooks("A.xlsm").Activate
    Workbooks("A.xlsm").Sheets("Parsed").Range("O1:Y26").Value = Workbooks("F.xlsm").Sheets("TASKS").Range("C8:M33").Value
    Workbooks("A.xlsm").Sheets("Parsed").Range("A1:M64").Value = Workbooks("F.xlsm").Sheets("B").Range("A1:M64").Value
    Workbooks("F.xlsm").Close False

    Workbooks("A.xlsm").Activate
    Application.DisplayAlerts = False
    'linkArray = ActiveWorkbook.LinkSources(xlExcelLinks)
    'ActiveWorkbook.OpenLinks linkArray(2)
    Workbooks.Open Filename:=ThisWorkbook.Path & "\C.xlsm"
    Workbooks("C.xlsm").Activate
    Workbooks("C.xlsm").Unprotect Password:="spike"
    Workbooks("C.xlsm").Sheets("B").Range("A1:Y75").Value = Workbooks("A.xlsm").Sheets("Parsed").Range("A1:Y75").Value
    Sheets("Free Me").Visible = True
    Sheets("Free Me").Activate
    Sheets("Free Me").Unprotect Password:="spike"
    Workbooks("C.xlsm").Sheets("Free Me").Range("C8:M33").Value = Workbooks("A.xlsm").Sheets("Parsed").Range("O1:Y26").Value
    Range("A16").EntireRow.Hidden = True
    If Range("G26") = "NONE IDENTIFIED" Then
        Range("A28:A33").EntireRow.Hidden = True
        Range("G26").Select
        With Selection.Font
            .Color = -16776961
            .TintAndShade = 0
        End With
    ElseIf Range("G26") <> "NONE IDENTIFIED" Then
        Range("A28:A33").EntireRow.Hidden = False
        Range("G26").Select
        With Selection.Font
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
        End With
    End If

    Sheets("Free Me").Protect Password:="spike"
    Sheets("Claim Check").Visible = True
    Sheets("Claim Check").Activate
    Sheets("Claim Check").Unprotect Password:="spike"
        Range("A17:A19").EntireRow.Hidden = True
        Range("B22:C22").ClearContents
        Range("B22:C22").UnMerge
        Range("D22").FormulaR1C1 = "Free Straight Distance    "
        Range("B23:D23").FormulaR1C1 = "Free Out & Return Distance    "
        Range("B24:C24").ClearContents
        Range("B24:C24").UnMerge
        Range("D24").FormulaR1C1 = "Free Distance via Up to 3 TP  "
        Range("B25:D25").FormulaR1C1 = "Free FAI Triangle Distance    "
    Sheets("Claim Check").Protect Password:="spike"
    Sheets("Worksheet").Visible = True
    Sheets("Worksheet").Activate
    Sheets("Worksheet").Unprotect Password:="spike"
    ActiveSheet.Shapes("Rectangle 2").Visible = False
    Range("A11:A71").EntireRow.Hidden = True
    Range("A72:A102").EntireRow.Hidden = False
    Range("A1").Select
    Sheets("Worksheet").Protect Password:="spike"
    Sheets("Worksheet").Visible = False
    Sheets("Free Me").Activate
    Sheets("Claim Check").Visible = False
    Sheets("Verify Task").Visible = False
    Workbooks("C.xlsm").Protect Password:="spike"
    Application.Cursor = xlDefault
    Application.DisplayFullScreen = True
    Application.Calculation = xlCalculationAutomatic
    Workbooks("A.xlsm").Close False
End If
End Sub
