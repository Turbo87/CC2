Option Explicit
Dim linkArray As Variant
Dim newHour As Variant
Dim newMinute As Variant
Dim newSecond As Variant
Dim waitTime As Variant
Sub OpenAb()
Attribute OpenAb.VB_ProcData.VB_Invoke_Func = " \n14"
'
' JLR 6/1/2011
'
Application.ScreenUpdating = False
If Range("C23") = "      Click on the glider to continue" Then
    Dim myFile As String
    myFile = Application.GetOpenFilename("IGC Files,*.igc")
    
    If myFile <> "False" Then
    Application.ScreenUpdating = True
    Workbooks("A.xlsm").Unprotect Password:="spike"
    Sheets("Logo").Visible = True
    Sheets("Logo").Select
    ActiveWindow.DisplayWorkbookTabs = False
        newHour = Hour(Now())
        newMinute = Minute(Now())
        newSecond = Second(Now()) + 2
        waitTime = TimeSerial(newHour, newMinute, newSecond)
        Application.Wait waitTime
    
    Sheets("All Claims").Visible = False
    Application.ScreenUpdating = False
    
    Application.Cursor = xlWait
    Workbooks.OpenText Filename:=myFile, Origin:=437, StartRow:=1 _
        , DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, _
        ConsecutiveDelimiter:=False, Tab:=False, Semicolon:=False, Comma:=False _
        , Space:=False, Other:=False, FieldInfo:=Array(1, 1), _
        TrailingMinusNumbers:=True
    Application.ScreenUpdating = False
    Columns("A:A").Copy
    Application.DisplayAlerts = False
    ActiveWorkbook.Close
    Application.DisplayAlerts = False

    'linkArray = ActiveWorkbook.LinkSources(xlExcelLinks)
    'ActiveWorkbook.OpenLinks linkArray(4)
    Workbooks.Open Filename:=ThisWorkbook.Path & "\Ab.xlsm"
    Workbooks("Ab.xlsm").Activate
    ActiveWorkbook.Unprotect Password:="spike"
    Sheets("BR").Select
    Range("A1").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    ActiveWindow.WindowState = xlMinimized
       Windows("A.xlsm").Activate
       ActiveWindow.WindowState = xlMaximized
       Application.ScreenUpdating = True
    Application.Run "Ab.xlsm!NewENLA"
    Application.Run "A.xlsm!CALx"
    Sheets("Parsed").Visible = False
    Sheets("E-Dec").Visible = True
    Sheets("E-Dec").Activate
        ActiveSheet.Unprotect Password:="spike"
    Range("A1:J29").Select
    ActiveWindow.Zoom = True
        ActiveSheet.Protect Password:="spike"
    Sheets("Logo").Visible = False
    Sheets("All Claims").Visible = True
    Sheets("Data Entry Check").Visible = True
    Sheets("Data Entry Check").Select
        ActiveSheet.Unprotect Password:="spike"
    Range("A1:K30").Select
    ActiveWindow.Zoom = True
        ActiveSheet.Protect Password:="spike"
    Range("G12:H12").Select
    ActiveWindow.ScrollWorkbookTabs Sheets:=-3
    Application.Calculate
    ActiveWindow.DisplayWorkbookTabs = True
    
    Application.Cursor = xlDefault
    Application.DisplayFullScreen = True
    Application.ScreenUpdating = True
    Workbooks("A.xlsm").Protect Password:="spike"
    Else: Exit Sub
    End If
End If
End Sub
Sub STclear()
Attribute STclear.VB_ProcData.VB_Invoke_Func = " \n14"
'
' JLR 6/1/2011//updates 7/11/2015 // Changed E-Dec 5/11/2018
'
    Application.Calculation = xlCalculationAutomatic
    Application.Iteration = True
    Application.MaxIterations = 100
    Application.MaxChange = 0.001
    Application.DisplayFullScreen = True
 Application.ScreenUpdating = False
 On Error Resume Next
    Workbooks("Ab.xlsm").Activate
  If Err = 0 Then
    Application.DisplayAlerts = False
    Workbooks("Ab.xlsm").Close
  End If
   On Error Resume Next
    Workbooks("D.xlsm").Activate
  If Err = 0 Then
    Application.DisplayAlerts = False
    Workbooks("D.xlsm").Close
  End If
    Workbooks("A.xlsm").Activate
    ActiveWorkbook.Unprotect Password:="spike"
    Application.ScreenUpdating = False
    Application.MoveAfterReturn = True
    Application.MoveAfterReturnDirection = xlToRight
    Application.DisplayFormulaBar = False
    Application.DisplayStatusBar = False
    ActiveWindow.DisplayWorkbookTabs = True
    Sheets("Parsed").Visible = False
    With ActiveWindow
        .DisplayHorizontalScrollBar = False
        .DisplayVerticalScrollBar = False
    End With
    'Sheets("E-DEC").Visible = True
    Sheets("E-DEC").Select
        Sheets("E-Dec").Unprotect Password:="spike"
    Range("D6:E6").FormulaR1C1 = "=CONCATENATE('[Ab.xlsm]PRS'!R[-5]C[-3],"" , "",'[Ab.xlsm]PRS'!R[-5]C[-2])"
    Range("H6").FormulaR1C1 = "='[Ab.xlsm]PRS'!R[-4]C[-7]"
    Range("C8:E8").FormulaR1C1 = "=IF('[Ab.xlsm]PRS'!R4C2=0,'[Ab.xlsm]PRS'!R4C1,CONCATENATE('[Ab.xlsm]PRS'!R4C1,"" & "",'[Ab.xlsm]PRS'!R4C2))"
    Range("G8:H8").FormulaR1C1 = "=CONCATENATE('[Ab.xlsm]PRS'!R5C1,"" , "",'[Ab.xlsm]PRS'!R5C2)"
    Range("E10:F10").FormulaR1C1 = "=IF('ALL CLAIMS'!R18C6=3,[Ab.xlsm]IMP!R2C25,MAX([Ab.xlsm]IMP!R1C25,[Ab.xlsm]IMP!R1021C25))"
    'Range("I10").FormulaR1C1 = "=IF('ALL CLAIMS'!R18C6=3,[Ab.xlsm]IMP!R2C26,[Ab.xlsm]PRS!R14C2)"
    'Range("C13").FormulaR1C1 = "=IF('ALL CLAIMS'!R18C6=3,[Ab.xlsm]IMP!R1C41,[Ab.xlsm]PRS!R20C7)"
    'Range("E13").FormulaR1C1 = "=IF('ALL CLAIMS'!R18C6=3,[Ab.xlsm]IMP!R2C35,[Ab.xlsm]PRS!R20C1)"
    'Range("F13").FormulaR1C1 = "=IF('ALL CLAIMS'!R18C6=3,[Ab.xlsm]IMP!R2C36,[Ab.xlsm]PRS!R20C2)"
    'Range("H13").FormulaR1C1 = "=IF('ALL CLAIMS'!R18C6=3,[Ab.xlsm]IMP!R2C38,[Ab.xlsm]PRS!R20C4)"
    'Range("I13").FormulaR1C1 = "=IF('ALL CLAIMS'!R18C6=3,[Ab.xlsm]IMP!R2C39,[Ab.xlsm]PRS!R20C5)"
    'Range("C15").FormulaR1C1 = "=IF(AND(R[-5]C[6]>=1,'ALL CLAIMS'!R18C6=3),[Ab.xlsm]IMP!R1C48,IF(R[-5]C[6]>=1,[Ab.xlsm]PRS!R22C7,0))"
    'Range("E15").FormulaR1C1 = "=IF(RC[-2]=""NONE"","""",IF('ALL CLAIMS'!R18C6=3,[Ab.xlsm]IMP!R2C42,[Ab.xlsm]PRS!R22C1))"
    'Range("F15").FormulaR1C1 = "=IF(RC[-1]="""","""",IF('ALL CLAIMS'!R18C6=3,[Ab.xlsm]IMP!R2C43,[Ab.xlsm]PRS!R22C2))"
    'Range("H15").FormulaR1C1 = "=IF(RC[-2]="""","""",IF('ALL CLAIMS'!R18C6=3,[Ab.xlsm]IMP!R2C45,[Ab.xlsm]PRS!R22C4))"
    'Range("I15").FormulaR1C1 = "=IF(RC[-1]="""","""",IF('ALL CLAIMS'!R18C6=3,[Ab.xlsm]IMP!R2C46,[Ab.xlsm]PRS!R22C5))"
    'Range("C17").FormulaR1C1 = "=IF(AND(R[-7]C[6]>=2,'ALL CLAIMS'!R18C6=3),[Ab.xlsm]IMP!R1C55,IF(R[-7]C[6]>=2,[Ab.xlsm]PRS!R24C7,0))"
    'Range("E17").FormulaR1C1 = "=IF(RC[-2]=""NONE"","""",IF('ALL CLAIMS'!R18C6=3,[Ab.xlsm]IMP!R2C49,[Ab.xlsm]PRS!R24C1))"
    'Range("F17").FormulaR1C1 = "=IF(RC[-1]="""","""",IF('ALL CLAIMS'!R18C6=3,[Ab.xlsm]IMP!R2C50,[Ab.xlsm]PRS!R24C2))"
    'Range("H17").FormulaR1C1 = "=IF(RC[-2]="""","""",IF('ALL CLAIMS'!R18C6=3,[Ab.xlsm]IMP!R2C52,[Ab.xlsm]PRS!R24C4))"
    'Range("I17").FormulaR1C1 = "=IF(RC[-1]="""","""",IF('ALL CLAIMS'!R18C6=3,[Ab.xlsm]IMP!R2C53,[Ab.xlsm]PRS!R24C5))"
    'Range("C19").FormulaR1C1 = "=IF(AND(R[-9]C[6]=3,'ALL CLAIMS'!R18C6=3),[Ab.xlsm]IMP!R1C62,IF(R[-9]C[6]=3,[Ab.xlsm]PRS!R26C7,0))"
    'Range("E19").FormulaR1C1 = "=IF(AND(R[-9]C[4]=3,'ALL CLAIMS'!R18C6=3),[Ab.xlsm]IMP!R2C56,IF(R[-9]C[4]=3,[Ab.xlsm]PRS!R26C1,0))"
    'Range("F19").FormulaR1C1 = "=IF(AND(R[-9]C[3]=3,'ALL CLAIMS'!R18C6=3),[Ab.xlsm]IMP!R2C57,IF(R[-9]C[3]=3,[Ab.xlsm]PRS!R26C2,0))"
    'Range("H19").FormulaR1C1 = "=IF(AND(R[-9]C[1]=3,'ALL CLAIMS'!R18C6=3),[Ab.xlsm]IMP!R2C59,IF(R[-9]C[1]=3,[Ab.xlsm]PRS!R26C4,0))"
    'Range("I19").FormulaR1C1 = "=IF(AND(R[-9]C=3,'ALL CLAIMS'!R18C6=3),[Ab.xlsm]IMP!R2C60,IF(R[-9]C=3,[Ab.xlsm]PRS!R26C5,0))"
    'Range("C21").FormulaR1C1 = _
        "=IF(AND(R[-11]C[6]=0,'ALL CLAIMS'!R18C6=3),[Ab.xlsm]IMP!R1C48,IF(AND(R[-11]C[6]=1,'ALL CLAIMS'!R18C6=3),[Ab.xlsm]IMP!R1C55,IF(AND(R[-11]C[6]=2,'ALL CLAIMS'!R18C6=3),[Ab.xlsm]IMP!R1C62,IF(AND(R[-11]C[6]=3,'ALL CLAIMS'!R18C6=3),[Ab.xlsm]IMP!R1C69,[Ab.xlsm]PRS!R28C7))))"
    'Range("E21").FormulaR1C1 = _
        "=IF(AND(R[-11]C[4]=0,'ALL CLAIMS'!R18C6=3),[Ab.xlsm]IMP!R2C42,IF(AND(R[-11]C[4]=1,'ALL CLAIMS'!R18C6=3),[Ab.xlsm]IMP!R2C49,IF(AND(R[-11]C[4]=2,'ALL CLAIMS'!R18C6=3),[Ab.xlsm]IMP!R2C56,IF(AND(R[-11]C[4]=3,'ALL CLAIMS'!R18C6=3),[Ab.xlsm]IMP!R2C63,[Ab.xlsm]PRS!R28C1))))"
    'Range("F21").FormulaR1C1 = _
        "=IF(AND(R[-11]C[3]=0,'ALL CLAIMS'!R18C6=3),[Ab.xlsm]IMP!R2C43,IF(AND(R[-11]C[3]=1,'ALL CLAIMS'!R18C6=3),[Ab.xlsm]IMP!R2C50,IF(AND(R[-11]C[3]=2,'ALL CLAIMS'!R18C6=3),[Ab.xlsm]IMP!R2C57,IF(AND(R[-11]C[3]=3,'ALL CLAIMS'!R18C6=3),[Ab.xlsm]IMP!R2C64,[Ab.xlsm]PRS!R28C2))))"
    'Range("H21").FormulaR1C1 = _
        "=IF(AND(R[-11]C[1]=0,'ALL CLAIMS'!R18C6=3),[Ab.xlsm]IMP!R2C45,IF(AND(R[-11]C[1]=1,'ALL CLAIMS'!R18C6=3),[Ab.xlsm]IMP!R2C52,IF(AND(R[-11]C[1]=2,'ALL CLAIMS'!R18C6=3),[Ab.xlsm]IMP!R2C59,IF(AND(R[-11]C[1]=3,'ALL CLAIMS'!R18C6=3),[Ab.xlsm]IMP!R2C66,[Ab.xlsm]PRS!R28C4))))"
    'Range("I21").FormulaR1C1 = _
        "=IF(AND(R[-11]C=0,'ALL CLAIMS'!R18C6=3),[Ab.xlsm]IMP!R2C46,IF(AND(R[-11]C=1,'ALL CLAIMS'!R18C6=3),[Ab.xlsm]IMP!R2C53,IF(AND(R[-11]C=2,'ALL CLAIMS'!R18C6=3),[Ab.xlsm]IMP!R2C60,IF(AND(R[-11]C=3,'ALL CLAIMS'!R18C6=3),[Ab.xlsm]IMP!R2C67,[Ab.xlsm]PRS!R28C5))))"
    Range("I10").FormulaR1C1 = "=[Ab.xlsm]IMP!R2C26"
    Range("C13").FormulaR1C1 = "=[Ab.xlsm]IMP!R1C41"
    Range("E13").FormulaR1C1 = "=[Ab.xlsm]IMP!R2C35"
    Range("F13").FormulaR1C1 = "=[Ab.xlsm]IMP!R2C36"
    Range("H13").FormulaR1C1 = "=[Ab.xlsm]IMP!R2C38"
    Range("I13").FormulaR1C1 = "=[Ab.xlsm]IMP!R2C39"
    Range("C15").FormulaR1C1 = "=IF(R[-5]C[6]>=1,[Ab.xlsm]IMP!R1C48,""None"")"
    Range("E15").FormulaR1C1 = "=IF(RC[-2]=""NONE"","""",[Ab.xlsm]IMP!R2C42)"
    Range("F15").FormulaR1C1 = "=IF(RC[-1]="""","""",[Ab.xlsm]IMP!R2C43)"
    Range("H15").FormulaR1C1 = "=IF(RC[-2]="""","""",[Ab.xlsm]IMP!R2C45)"
    Range("I15").FormulaR1C1 = "=IF(RC[-1]="""","""",[Ab.xlsm]IMP!R2C46)"
    Range("C17").FormulaR1C1 = "=IF(R[-7]C[6]>=2,[Ab.xlsm]IMP!R1C55,""None"")"
    Range("E17").FormulaR1C1 = "=IF(RC[-2]=""NONE"","""",[Ab.xlsm]IMP!R2C49)"
    Range("F17").FormulaR1C1 = "=IF(RC[-1]="""","""",[Ab.xlsm]IMP!R2C50)"
    Range("H17").FormulaR1C1 = "=IF(RC[-2]="""","""",[Ab.xlsm]IMP!R2C52)"
    Range("I17").FormulaR1C1 = "=IF(RC[-1]="""","""",[Ab.xlsm]IMP!R2C53)"
    Range("C19").FormulaR1C1 = "=IF(R[-9]C[6]>=3,[Ab.xlsm]IMP!R1C62,""None"")"
    Range("E19").FormulaR1C1 = "=IF(R[-9]C[4]>=3,[Ab.xlsm]IMP!R2C56,"""")"
    Range("F19").FormulaR1C1 = "=IF(R[-9]C[3]>=3,[Ab.xlsm]IMP!R2C57,"""")"
    Range("H19").FormulaR1C1 = "=IF(R[-9]C[1]>=3,[Ab.xlsm]IMP!R2C59,"""")"
    Range("I19").FormulaR1C1 = "=IF(R[-9]C>=3,[Ab.xlsm]IMP!R2C60,"""")"
    Range("C21").FormulaR1C1 = "=IF(R[-11]C[6]=0,[Ab.xlsm]IMP!R1C48,IF(R[-11]C[6]=1,[Ab.xlsm]IMP!R1C55,IF(R[-11]C[6]=2,[Ab.xlsm]IMP!R1C62,IF(OR(R[-11]C[6]=""Too Many"",R[-11]C[6]>=3),[Ab.xlsm]IMP!R1C69))))"
    Range("E21").FormulaR1C1 = "=IF(R[-11]C[4]=0,[Ab.xlsm]IMP!R2C42,IF(R[-11]C[4]=1,[Ab.xlsm]IMP!R2C49,IF(R[-11]C[4]=2,[Ab.xlsm]IMP!R2C56,IF(R[-11]C[4]>=3,[Ab.xlsm]IMP!R2C63,""""))))"
    Range("F21").FormulaR1C1 = "=IF(R[-11]C[3]=0,[Ab.xlsm]IMP!R2C43,IF(R[-11]C[3]=1,[Ab.xlsm]IMP!R2C50,IF(R[-11]C[3]=2,[Ab.xlsm]IMP!R2C57,IF(R[-11]C[3]>=3,[Ab.xlsm]IMP!R2C64,""""))))"
    Range("H21").FormulaR1C1 = "=IF(10=0,[Ab.xlsm]IMP!R2C45,IF(R[-11]C[1]=1,[Ab.xlsm]IMP!R2C52,IF(R[-11]C[1]=2,[Ab.xlsm]IMP!R2C59,IF(R[-11]C[1]>=3,[Ab.xlsm]IMP!R2C66,""""))))"
    Range("I21").FormulaR1C1 = "=IF(R[-11]C=0,[Ab.xlsm]IMP!R2C46,IF(R[-11]C=1,[Ab.xlsm]IMP!R2C53,IF(R[-11]C=2,[Ab.xlsm]IMP!R2C60,IF(R[-11]C>=3,[Ab.xlsm]IMP!R2C67,""""))))"
    Range("B21").FormulaR1C1 = "=IF(R[-11]C[7]<=3,""Finish Point"",""Invalid TP"")"
    Range("A1").Select
        Sheets("E-DEC").Protect Password:="spike"
    'Sheets("OTHER").Visible = True
    Sheets("OTHER").Select
        Sheets("OTHER").Unprotect Password:="spike"
    Sheets("OTHER").Shapes("Drop Down 38").Visible = True
    Sheets("OTHER").Shapes("Drop Down 38").Locked = False
    Sheets("OTHER").Shapes("Drop Down 45").Visible = False
    Range("K15:L15").FormulaR1C1 = "=1"
    Range("K15:L15").Locked = True
    Range("A12:A28").EntireRow.Hidden = False
    Range("D6:E6").ClearContents
    Range("D6:E6").Locked = False
    Range("D6:E6").FormulaHidden = True
    Range("J6").ClearContents
    Range("J6").Locked = False
    Range("J6").FormulaHidden = True
    Range("K6").FormulaR1C1 = "=IF(OR(RC[-1]="""",'All Claims'!R65C16=0),"""",RC[-1]-'All Claims'!R65C16/24)"
    Range("M6").FormulaR1C1 = "=IF(AND(RC[-3]="""",SUM(R[14]C[-8]:R[22]C[-8])=0),"""",IF(OR(AND(OR(R4C4=2,R4C4=3),R4C10>0,SUM(RC[-9]+RC[-3])>0,OR(RC[-9]+RC[-3]-('[Ab.xlsm]PRS'!R2C2/24)>='[Ab.xlsm]PRS'!R1C7,AND(Parsed!R[9]C[-12]<Parsed!R[-5]C[-6],RC[-9]+RC[-3]-('[Ab.xlsm]PRS'!R2C2/24)<Parsed!R[9]C[-12])))),""INVALID"",""""))"
    Range("D8:F8").ClearContents
    Range("D8:F8").Locked = False
    Range("D8:F8").FormulaHidden = True
    Range("J8:L8").ClearContents
    Range("J8:L8").Locked = True
    Range("J8:L8").FormulaHidden = True
    Range("D10:G10").ClearContents
    Range("D10:G10").Locked = False
    Range("D10:G10").FormulaHidden = True
    Range("K10:L10").ClearContents
    Range("K10:L10").Locked = False
    Range("K10:L10").FormulaHidden = True
    Sheets("OTHER").Unprotect Password:="spike"
    Range("D4").FormulaR1C1 = "=1"
    Sheets("OTHER").Unprotect Password:="spike"
    Range("D4:E4").Locked = False
    Selection.FormulaHidden = True
    Sheets("OTHER").Unprotect Password:="spike"
    ActiveSheet.Shapes("Drop Down 5").Locked = False
    Sheets("OTHER").Unprotect Password:="spike"
    Range("D15").FormulaR1C1 = "=1"
    Range("D15:E15").Locked = False
    Sheets("OTHER").Shapes("Oval 14").Visible = False
    Sheets("OTHER").Shapes("Oval 16").Visible = False
    Sheets("OTHER").Shapes("Oval 18").Visible = False
    Sheets("OTHER").Shapes("Oval 20").Visible = False
    Sheets("OTHER").Shapes("Oval 22").Visible = False
    Sheets("OTHER").Shapes("Rectangle 1").Visible = False
    Range("C20:M28").ClearContents
    Range("C20,E20:H20,J20:M20").Locked = False
    Range("C22,E22:H22,J22:M22").Locked = False
    Range("C24,E24:H24,J24:M24").Locked = False
    Range("C26,E26:H26,J26:M26").Locked = False
    Range("C28,E28:H28,J28:M28").Locked = False
    Range("C69:M93").Clear
    Range("A1").RowHeight = 12
    Columns("A:A").ColumnWidth = 21
    Range("A1:O36").Select
    ActiveWindow.Zoom = True
    Range("A1").Select
        Sheets("OTHER").Protect Password:="spike"
    Sheets("OTHER").Visible = False
    Sheets("Parsed").Visible = True
    Sheets("Parsed").Select
    Range("A1:I30").Clear
    Range("A39:B64").Clear
    Range("A1").Select
    Sheets("Parsed").Visible = False
    'Sheets("Data Entry Check").Visible = True
    Sheets("Data Entry Check").Select
        Sheets("Data Entry Check").Unprotect Password:="spike"
    Range("E2:F2").FormulaR1C1 = _
        "=IF(AND('ALL CLAIMS'!R18C6=3,OTHER!R31C2=""Data Entry Required Above""),""Declaration Incomplete"",IF('ALL CLAIMS'!R18C6=4,""Custom Electronic Declaration"",IF(OR(AND('ALL CLAIMS'!R18C6=5,'E-DEC'!R2C7<>""Downloaded""),R[1]C=""Free-W"",R[1]C=""Free-E""),""Free Task"",IF(OR(R[1]C=""Electronic"",R[1]C=""Written""),CONCATENATE(R[1]C,"" Pre-flight Declaration""),IF(R[1]C=""Free-W inv"",""Free Written Dec Invalid"",IF(R[1]C=""W inv"",""Written declaration invalid"",IF(OR(R[1]C=""Free-E inv"",R[1]C=""E INV""),""Electronic declaration invalid"")))))))"
    Range("G2:I2").FormulaR1C1 = "=IF(AND('ALL CLAIMS'!R18C6=4,OR('[Ab.xlsm]PRS'!R4C1<>'[Ab.xlsm]PRS'!R8C1,'[Ab.xlsm]PRS'!R5C1<>'[Ab.xlsm]PRS'!R9C1,'[Ab.xlsm]PRS'!R5C2<>'[Ab.xlsm]PRS'!R9C2)),""User Changes"","""")"
    Range("E3").FormulaR1C1 = "=[Ab.xlsm]PRS!R11C4"
    Range("F3").FormulaR1C1 = "=[Ab.xlsm]PRS!R3C3"
    Range("J8").ClearContents
     Range("I6:J6").FormulaR1C1 = _
        "=IF(OR('ALL CLAIMS'!R18C6=5,AND('ALL CLAIMS'!R18C6>=2,'ALL CLAIMS'!R18C6<4,R[-3]C[-4]<>""Written"")),'E-DEC'!R[2]C[-2],IF(OTHER!R[4]C[-5]="""",CONCATENATE(""TYPE? ,"",OTHER!R[4]C[2]),CONCATENATE(OTHER!R[4]C[-5],"", "",OTHER!R[4]C[2])))"
    Range("D8:E8").FormulaR1C1 = _
        "=IF(OR('ALL CLAIMS'!R18C6=5,AND('ALL CLAIMS'!R18C6>=2,'ALL CLAIMS'!R18C6<4,R3C5<>""Written"")),[Ab.xlsm]PRS!R4C1,IF(OTHER!RC="""",""UNSPECIFIED"",OTHER!RC))"
    Range("F8").FormulaR1C1 = _
        "=IF(R[-5]C[-1]=""Written"",""W"",IF('ALL CLAIMS'!R18C6=4,""OO"",""""))"
    Range("G8").FormulaR1C1 = "=IF(RC[2]="""","""",""Flight Crew  "")"
    Range("C10").FormulaR1C1 = "=IF([Ab.xlsm]PRS!R1C4=0,""First Data Pt"",IF([Ab.xlsm]PRS!R2C2=0,""Take Off, UTC"",""Take Off, LCL""))"
    Range("D10").FormulaR1C1 = "=IF(AND(RC[-1]=""First Data Pt"",[Ab.xlsm]PRS!R2C2=0),[Ab.xlsm]PRS!R1C7,IF(AND(RC[-1]=""First Data Pt"",[Ab.xlsm]PRS!R2C2<>0),[Ab.xlsm]PRS!R1C7+[Ab.xlsm]PRS!R2C2/24,IF([Ab.xlsm]PRS!R2C2=0,[Ab.xlsm]PRS!R18C7,[Ab.xlsm]PRS!R18C7+[Ab.xlsm]PRS!R2C2/24)))"
    Range("E10:F10").FormulaR1C1 = "=IF(AND(RC[-2]=""First Data Pt, LCL  "",RC[-1]=RC[2]),""First Data Pt, LCL  "",IF(AND([Ab.xlsm]PRS!R2C2<>0,[Ab.xlsm]PRS!R2C7=[Ab.xlsm]PRS!R3C7),""MoP Stop, LCL  "",IF(AND([Ab.xlsm]PRS!R2C2=0,[AB.xlsm]PRS!R2C7=[AB.xlsm]PRS!R3C7),""MoP Stop, UTC  "",IF([Ab.xlsm]PRS!R2C2=0,""Default Release Time, UTC  "",""Default Release Time, LCL ""))))"
    Range("G10").FormulaR1C1 = "=IF(AND(OR('ALL CLAIMS'!R[-2]C[-1]=3,'ALL CLAIMS'!R[-2]C[-1]=5),[Ab.xlsm]PRS!R[-8]C>0),[Ab.xlsm]PRS!R[-8]C+[Ab.xlsm]PRS!R2C2/24,[Ab.xlsm]PRS!R[-7]C+[Ab.xlsm]PRS!R2C2/24)"
    Range("I10").FormulaR1C1 = "=IF(AND([Ab.xlsm]PRS!R2C2=0,[Ab.xlsm]PRS!R6C7>0,[Ab.xlsm]PRS!R6C7<[Ab.xlsm]PRS!R10C7),""MG Fini, UTC"",IF(AND([Ab.xlsm]PRS!R2C2=0,[Ab.xlsm]PRS!R6C7=0),""Landing, UTC"",IF(AND([Ab.xlsm]PRS!R6C7>0,[Ab.xlsm]PRS!R6C7<[Ab.xlsm]PRS!R10C7),""MG Fini, LCL"",""Landing, LCL"")))"
    Range("J10").FormulaR1C1 = "=IF(RC[-1]=""Landing, UTC"",[Ab.xlsm]PRS!R10C7,IF(AND([Ab.xlsm]PRS!R2C2=0,RC[-1]=""MG Fini, UTC""),[Ab.xlsm]PRS!R6C7,IF(RC[-1]=""Landing, LCL"",[Ab.xlsm]PRS!RC[-3]+[Ab.xlsm]PRS!R2C2/24,[Ab.xlsm]PRS!R[-4]C[-3]+[Ab.xlsm]PRS!R2C2/24)))"
    Range("G12:H12").ClearContents
    Range("A1").Select
        ActiveSheet.Protect Password:="spike"
    Sheets("Data Entry Check").Visible = False
    Sheets("E-Dec").Visible = False
    Sheets("Calculating").Visible = False
    Sheets("Logo").Visible = False
    ActiveWorkbook.UpdateLink Name:=ActiveWorkbook.LinkSources
    Sheets("All Claims").Visible = True
    Sheets("All Claims").Select
    Sheets("All Claims").Unprotect Password:="spike"
    Sheets("All Claims").Shapes.Range(Array("Drop Down 141")).Select
    With Selection
        .ListFillRange = "$C$39:$D$43"
        .LinkedCell = "$F$18"
        .DropDownLines = 5
        .Display3DShading = False
    End With
    Range("F6").FormulaR1C1 = "1"
    Range("F8").FormulaR1C1 = "1"
    Range("F10").FormulaR1C1 = "1"
    Range("D12").FormulaR1C1 = "1"
    Range("F14,F16").ClearContents
    Range("F18").FormulaR1C1 = "1"
    Range("B22").FormulaR1C1 = "='[C.xlsm]Verify Task'!R14C1"
    Range("A1:H28").Select
    ActiveWindow.Zoom = True
    Range("A1").Select
        Sheets("All Claims").Protect Password:="spike"
    ActiveWorkbook.Protect Password:="spike"
    Application.Cursor = xlDefault
    Application.ScreenUpdating = True
End Sub