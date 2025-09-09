Option Explicit
#If VBA7 And Win64 Then
    Dim LastRow As LongPtr
#Else
    Dim LastRow As Long
#End If

Sub Aselect()
'
' JLR 4/30/14 Revised 4/9/2015 Hp correction already in Ab
'
Application.ScreenUpdating = False
Workbooks("A.xlsm").Activate
ActiveWindow.WindowState = xlMaximized
Application.ScreenUpdating = True
ActiveWindow.DisplayWorkbookTabs = False
Application.ScreenUpdating = False
Workbooks("F.xlsm").Activate
'HpA correction copy to M2,M5,M8
Sheets("B").Range("M2").Value = Sheets("B").Range("F1").Value
Sheets("B").Range("M5").Value = Sheets("B").Range("F2").Value
Sheets("B").Range("M8").Value = Sheets("B").Range("F3").Value
Sheets("Sheet2").Activate

Range("E9").FormulaR1C1 = "=IF(OR(B!R[8]C[-4]<>"""",B!R[-6]C[-4]=""X""),""PR"","""")"
Range("D9").FormulaR1C1 = "=IF(RC[1]="""",1000,900)"
Range("D9:E9").Value = Range("D9:E9").Value

If Range("A10009") <> "" Then
    Sheets("Sheet3").Activate
    Sheets("Sheet3").Range("A9").Value = "REF"
    Sheets("Sheet3").Range("D9:E9").Value = Sheets("Sheet2").Range("D9:E9").Value
    Sheets("Sheet3").Range("A10:E60009").Value = Sheets("Sheet2").Range("A10:E60009").Value
    
    Sheets("Sheet2").Activate
    If Range("A10009") <> "" And Range("A20009") = "" Then
        Range("I12:I15").FormulaR1C1 = "=IF(R[-1]C="""",1,"""")"
    ElseIf Range("A20009") <> "" And Range("A30009") = "" Then
        Range("I12:I15").FormulaR1C1 = "=IF(SUM(R[-2]C:R[-1]C)=0,1,"""")"
    ElseIf Range("A30009") <> "" And Range("A40009") = "" Then
        Range("I13:I15").FormulaR1C1 = "=IF(SUM(R[-3]C:R[-1]C)=0,1,"""")"
    ElseIf Range("A40009") <> "" Then
        Range("I15").FormulaR1C1 = "=IF(SUM(R[-5]C:R[-1]C)=0,1,"""")"
    End If

    Application.Calculation = xlCalculationManual
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("I15:I15").AutoFill Destination:=.Range("I15:I" & LastRow), Type:=xlFillDefault
    End With
ActiveSheet.Calculate

    Range("I11:I60009").Value = Range("I11:I60009").Value

    Range("I10").Value = 1
    Range("J10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-9],"""")"
    Range("K10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-9],"""")"
    Range("L10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-9],"""")"
    Range("M10:N10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-9],"""")"
    Application.Calculation = xlCalculationManual
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("J10:N10").AutoFill Destination:=.Range("J10:N" & LastRow), Type:=xlFillDefault
    End With
ActiveSheet.Calculate

    Range("J10:N60009").Value = Range("J10:N60009").Value
    Range("J10:N60009").Sort Key1:=Range("J10"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    Range("A10:E60009").Clear
    Range("A10:E10009").Value = Range("J10:N10009").Value
    Columns("F:N").Clear
End If

    Application.Run "F.xlsm!Straight"

End Sub

Sub Straight()
'
' Revised 4/28/15 consider LoH > < 100, PR vs Fr; amended 9/29/15 for STD reverse
' Re-written 7/16/2017 to consolidate Straight & Large


Application.ScreenUpdating = False
 
    Range("A9").FormulaR1C1 = "=IF(Sheet3!R[1]C<>"""",""REF"","""")"
    ActiveSheet.Calculate
    Range("A9").Value = Range("A9").Value
    
    If Range("A9") = "REF" Then
             Sheets("Sheet3").Activate
   
    ElseIf Range("A9") <> "REF" Then
             Sheets("Sheet2").Activate
    End If

    Range("B5:C5").FormulaR1C1 = "=MIN(R[5]C:R[60004]C)"
    Range("B6:C6").FormulaR1C1 = "=MAX(R[4]C:R[60003]C)"
    Range("B5:C6").Value = Range("B5:C6").Value
    
    Range("G10").FormulaR1C1 = _
        "=IF(OR(RC[-5]=R5C2,RC[-5]=R6C2,RC[-4]=R5C3,RC[-4]=R6C3),RC[-6],"""")"
    Range("H10:K10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-6],"""")"
    
    Application.Calculation = xlCalculationManual
    With ActiveSheet
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:K10").AutoFill Destination:=.Range("G10:K" & LastRow), Type:=xlFillDefault
    End With
ActiveSheet.Calculate

    Range("G10:K60009").Value = Range("G10:K60009").Value
    Range("G10:K60009").Sort Key1:=Range("G10"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    
    Range("M10").FormulaR1C1 = _
        "=IF(R[1]C[-6]="""","""",6371*ACOS(SIN(RC[-5])*SIN(R[1]C[-5])+COS(RC[-5])*COS(R[1]C[-5])*COS(R[1]C[-4]-RC[-4])))"
    Range("N10").FormulaR1C1 = _
        "=IF(R[2]C[-7]="""","""",6371*ACOS(SIN(RC[-6])*SIN(R[2]C[-6])+COS(RC[-6])*COS(R[2]C[-6])*COS(R[2]C[-5]-RC[-5])))"
    Range("O10").FormulaR1C1 = _
        "=IF(R[3]C[-8]="""","""",6371*ACOS(SIN(RC[-7])*SIN(R[3]C[-7])+COS(RC[-7])*COS(R[3]C[-7])*COS(R[3]C[-6]-RC[-6])))"
    Range("P10").FormulaR1C1 = _
        "=IF(R[4]C[-9]="""","""",6371*ACOS(SIN(RC[-8])*SIN(R[4]C[-8])+COS(RC[-8])*COS(R[4]C[-8])*COS(R[4]C[-7]-RC[-7])))"
    Range("Q10").FormulaR1C1 = _
        "=IF(R[5]C[-10]="""","""",6371*ACOS(SIN(RC[-9])*SIN(R[5]C[-9])+COS(RC[-9])*COS(R[5]C[-9])*COS(R[5]C[-8]-RC[-8])))"
    Range("R10").FormulaR1C1 = _
        "=IF(R[6]C[-11]="""","""",6371*ACOS(SIN(RC[-10])*SIN(R[6]C[-10])+COS(RC[-10])*COS(R[6]C[-10])*COS(R[6]C[-9]-RC[-9])))"
    Range("S10").FormulaR1C1 = _
        "=IF(R[7]C[-12]="""","""",6371*ACOS(SIN(RC[-11])*SIN(R[7]C[-11])+COS(RC[-11])*COS(R[7]C[-11])*COS(R[7]C[-10]-RC[-10])))"
    Range("T10").FormulaR1C1 = _
        "=IF(R[8]C[-13]="""","""",6371*ACOS(SIN(RC[-12])*SIN(R[8]C[-12])+COS(RC[-12])*COS(R[8]C[-12])*COS(R[8]C[-11]-RC[-11])))"
    Range("U10").FormulaR1C1 = _
        "=IF(R[9]C[-14]="""","""",6371*ACOS(SIN(RC[-13])*SIN(R[9]C[-13])+COS(RC[-13])*COS(R[9]C[-13])*COS(R[9]C[-12]-RC[-12])))"
    Range("V10").FormulaR1C1 = _
        "=IF(R[10]C[-15]="""","""",6371*ACOS(SIN(RC[-14])*SIN(R[10]C[-14])+COS(RC[-14])*COS(R[10]C[-14])*COS(R[10]C[-13]-RC[-13])))"
    Range("W10").FormulaR1C1 = _
        "=IF(R[11]C[-16]="""","""",6371*ACOS(SIN(RC[-15])*SIN(R[11]C[-15])+COS(RC[-15])*COS(R[11]C[-15])*COS(R[11]C[-14]-RC[-14])))"
    
    Application.Calculation = xlCalculationManual
    With ActiveSheet
    LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
.Range("M10:W10").AutoFill Destination:=.Range("M10:W" & LastRow), Type:=xlFillDefault
    End With
ActiveSheet.Calculate

    Range("M10:W21").Value = Range("M10:W21").Value

  If Range("E9") <> "PR" Then

    Range("M23:M34").FormulaR1C1 = _
        "=IF(R[-13]C="""","""",IF(OR(AND(R[-13]C>100,R[-13]C[-3]-R[-12]C[-3]<=R9C4),AND(R[-13]C<=100,R[-13]C[-3]-R[-12]C[-3]<=10*R[-13]C)),R[-13]C,IF(AND(R[-13]C>100,R[-13]C[-3]-R[-12]C[-3]>R9C4),R[-13]C-((R[-13]C[-3]-R[-12]C[-3]-R9C4)*0.1),0)))"
    Range("N23:N34").FormulaR1C1 = _
        "=IF(R[-13]C="""","""",IF(OR(AND(R[-13]C>100,R[-13]C[-4]-R[-11]C[-4]<=R9C4),AND(R[-13]C<=100,R[-13]C[-4]-R[-11]C[-4]<=10*R[-13]C)),R[-13]C,IF(AND(R[-13]C>100,R[-13]C[-4]-R[-11]C[-4]>R9C4),R[-13]C-((R[-13]C[-4]-R[-11]C[-4]-R9C4)*0.1),0)))"
    Range("O23:O34").FormulaR1C1 = _
        "=IF(R[-13]C="""","""",IF(OR(AND(R[-13]C>100,R[-13]C[-5]-R[-10]C[-5]<=R9C4),AND(R[-13]C<=100,R[-13]C[-5]-R[-10]C[-5]<=10*R[-13]C)),R[-13]C,IF(AND(R[-13]C>100,R[-13]C[-5]-R[-10]C[-5]>R9C4),R[-13]C-((R[-13]C[-5]-R[-10]C[-5]-R9C4)*0.1),0)))"
    Range("P23:P34").FormulaR1C1 = _
        "=IF(R[-13]C="""","""",IF(OR(AND(R[-13]C>100,R[-13]C[-6]-R[-9]C[-6]<=R9C4),AND(R[-13]C<=100,R[-13]C[-6]-R[-9]C[-6]<=10*R[-13]C)),R[-13]C,IF(AND(R[-13]C>100,R[-13]C[-6]-R[-9]C[-6]>R9C4),R[-13]C-((R[-13]C[-6]-R[-9]C[-6]-R9C4)*0.1),0)))"
    Range("Q23:Q34").FormulaR1C1 = _
        "=IF(R[-13]C="""","""",IF(OR(AND(R[-13]C>100,R[-13]C[-7]-R[-8]C[-7]<=R9C4),AND(R[-13]C<=100,R[-13]C[-7]-R[-8]C[-7]<=10*R[-13]C)),R[-13]C,IF(AND(R[-13]C>100,R[-13]C[-7]-R[-8]C[-7]>R9C4),R[-13]C-((R[-13]C[-7]-R[-8]C[-7]-R9C4)*0.1),0)))"
    Range("R23:R34").FormulaR1C1 = _
        "=IF(R[-13]C="""","""",IF(OR(AND(R[-13]C>100,R[-13]C[-8]-R[-7]C[-8]<=R9C4),AND(R[-13]C<=100,R[-13]C[-8]-R[-7]C[-8]<=10*R[-13]C)),R[-13]C,IF(AND(R[-13]C>100,R[-13]C[-8]-R[-7]C[-8]>R9C4),R[-13]C-((R[-13]C[-8]-R[-7]C[-8]-R9C4)*0.1),0)))"
    Range("S23:S34").FormulaR1C1 = _
        "=IF(R[-13]C="""","""",IF(OR(AND(R[-13]C>100,R[-13]C[-9]-R[-6]C[-9]<=R9C4),AND(R[-13]C<=100,R[-13]C[-9]-R[-6]C[-9]<=10*R[-13]C)),R[-13]C,IF(AND(R[-13]C>100,R[-13]C[-9]-R[-6]C[-9]>R9C4),R[-13]C-((R[-13]C[-9]-R[-6]C[-9]-R9C4)*0.1),0)))"
    Range("T23:T34").FormulaR1C1 = _
        "=IF(R[-13]C="""","""",IF(OR(AND(R[-13]C>100,R[-13]C[-10]-R[-5]C[-10]<=R9C4),AND(R[-13]C<=100,R[-13]C[-10]-R[-5]C[-10]<=10*R[-13]C)),R[-13]C,IF(AND(R[-13]C>100,R[-13]C[-10]-R[-5]C[-10]>R9C4),R[-13]C-((R[-13]C[-10]-R[-5]C[-10]-R9C4)*0.1),0)))"
    Range("U23:U34").FormulaR1C1 = _
        "=IF(R[-13]C="""","""",IF(OR(AND(R[-13]C>100,R[-13]C[-11]-R[-4]C[-11]<=R9C4),AND(R[-13]C<=100,R[-13]C[-11]-R[-4]C[-11]<=10*R[-13]C)),R[-13]C,IF(AND(R[-13]C>100,R[-13]C[-11]-R[-4]C[-11]>R9C4),R[-13]C-((R[-13]C[-11]-R[-4]C[-11]-R9C4)*0.1),0)))"
    Range("V23:V34").FormulaR1C1 = _
        "=IF(R[-13]C="""","""",IF(OR(AND(R[-13]C>100,R[-13]C[-12]-R[-3]C[-12]<=R9C4),AND(R[-13]C<=100,R[-13]C[-12]-R[-3]C[-12]<=10*R[-13]C)),R[-13]C,IF(AND(R[-13]C>100,R[-13]C[-12]-R[-3]C[-12]>R9C4),R[-13]C-((R[-13]C[-12]-R[-3]C[-12]-R9C4)*0.1),0)))"
    Range("W23:W34").FormulaR1C1 = _
        "=IF(R[-13]C="""","""",IF(OR(AND(R[-13]C>100,R[-13]C[-13]-R[-2]C[-13]<=R9C4),AND(R[-13]C<=100,R[-13]C[-13]-R[-2]C[-13]<=10*R[-13]C)),R[-13]C,IF(AND(R[-13]C>100,R[-13]C[-13]-R[-2]C[-13]>R9C4),R[-13]C-((R[-13]C[-13]-R[-2]C[-13]-R9C4))*0.1,0)))"
 
  ElseIf Range("E9") = "PR" Then
    
    Range("M23:M34").FormulaR1C1 = _
        "=IF(R[-13]C="""","""",IF(OR(AND(R[-13]C>100,R[-13]C[-3]-R[-12]C[-3]<=R9C4),AND(R[-13]C<=100,R[-13]C[-3]-R[-12]C[-3]<=(10*R[-13]C)-100)),R[-13]C,IF(AND(R[-13]C>100,R[-13]C[-3]-R[-12]C[-3]>R9C4),R[-13]C-((R[-13]C[-3]-R[-12]C[-3]-R9C4)*0.1),0)))"
    Range("N23:N34").FormulaR1C1 = _
        "=IF(R[-13]C="""","""",IF(OR(AND(R[-13]C>100,R[-13]C[-4]-R[-11]C[-4]<=R9C4),AND(R[-13]C<=100,R[-13]C[-4]-R[-11]C[-4]<=(10*R[-13]C)-100)),R[-13]C,IF(AND(R[-13]C>100,R[-13]C[-4]-R[-11]C[-4]>R9C4),R[-13]C-((R[-13]C[-4]-R[-11]C[-4]-R9C4)*0.1),0)))"
    Range("O23:O34").FormulaR1C1 = _
        "=IF(R[-13]C="""","""",IF(OR(AND(R[-13]C>100,R[-13]C[-5]-R[-10]C[-5]<=R9C4),AND(R[-13]C<=100,R[-13]C[-5]-R[-10]C[-5]<=(10*R[-13]C)-100)),R[-13]C,IF(AND(R[-13]C>100,R[-13]C[-5]-R[-10]C[-5]>R9C4),R[-13]C-((R[-13]C[-5]-R[-10]C[-5]-R9C4)*0.1),0)))"
    Range("P23:P34").FormulaR1C1 = _
        "=IF(R[-13]C="""","""",IF(OR(AND(R[-13]C>100,R[-13]C[-6]-R[-9]C[-6]<=R9C4),AND(R[-13]C<=100,R[-13]C[-6]-R[-9]C[-6]<=(10*R[-13]C)-100)),R[-13]C,IF(AND(R[-13]C>100,R[-13]C[-6]-R[-9]C[-6]>R9C4),R[-13]C-((R[-13]C[-6]-R[-9]C[-6]-R9C4)*0.1),0)))"
    Range("Q23:Q34").FormulaR1C1 = _
        "=IF(R[-13]C="""","""",IF(OR(AND(R[-13]C>100,R[-13]C[-7]-R[-8]C[-7]<=R9C4),AND(R[-13]C<=100,R[-13]C[-7]-R[-8]C[-7]<=(10*R[-13]C)-100)),R[-13]C,IF(AND(R[-13]C>100,R[-13]C[-7]-R[-8]C[-7]>R9C4),R[-13]C-((R[-13]C[-7]-R[-8]C[-7]-R9C4)*0.1),0)))"
    Range("R23:R34").FormulaR1C1 = _
        "=IF(R[-13]C="""","""",IF(OR(AND(R[-13]C>100,R[-13]C[-8]-R[-7]C[-8]<=R9C4),AND(R[-13]C<=100,R[-13]C[-8]-R[-7]C[-8]<=(10*R[-13]C)-100)),R[-13]C,IF(AND(R[-13]C>100,R[-13]C[-8]-R[-7]C[-8]>R9C4),R[-13]C-((R[-13]C[-8]-R[-7]C[-8]-R9C4)*0.1),0)))"
    Range("S23:S34").FormulaR1C1 = _
        "=IF(R[-13]C="""","""",IF(OR(AND(R[-13]C>100,R[-13]C[-9]-R[-6]C[-9]<=R9C4),AND(R[-13]C<=100,R[-13]C[-9]-R[-6]C[-9]<=(10*R[-13]C)-100)),R[-13]C,IF(AND(R[-13]C>100,R[-13]C[-9]-R[-6]C[-9]>R9C4),R[-13]C-((R[-13]C[-9]-R[-6]C[-9]-R9C4)*0.1),0)))"
    Range("T23:T34").FormulaR1C1 = _
        "=IF(R[-13]C="""","""",IF(OR(AND(R[-13]C>100,R[-13]C[-10]-R[-5]C[-10]<=R9C4),AND(R[-13]C<=100,R[-13]C[-10]-R[-5]C[-10]<=(10*R[-13]C)-100)),R[-13]C,IF(AND(R[-13]C>100,R[-13]C[-10]-R[-5]C[-10]>R9C4),R[-13]C-((R[-13]C[-10]-R[-5]C[-10]-R9C4)*0.1),0)))"
    Range("U23:U34").FormulaR1C1 = _
        "=IF(R[-13]C="""","""",IF(OR(AND(R[-13]C>100,R[-13]C[-11]-R[-4]C[-11]<=R9C4),AND(R[-13]C<=100,R[-13]C[-11]-R[-4]C[-11]<=(10*R[-13]C)-100)),R[-13]C,IF(AND(R[-13]C>100,R[-13]C[-11]-R[-4]C[-11]>R9C4),R[-13]C-((R[-13]C[-11]-R[-4]C[-11]-R9C4)*0.1),0)))"
    Range("V23:V34").FormulaR1C1 = _
        "=IF(R[-13]C="""","""",IF(OR(AND(R[-13]C>100,R[-13]C[-12]-R[-3]C[-12]<=R9C4),AND(R[-13]C<=100,R[-13]C[-12]-R[-3]C[-12]<=(10*R[-13]C)-100)),R[-13]C,IF(AND(R[-13]C>100,R[-13]C[-12]-R[-3]C[-12]>R9C4),R[-13]C-((R[-13]C[-12]-R[-3]C[-12]-R9C4)*0.1),0)))"
    Range("W23:W34").FormulaR1C1 = _
        "=IF(R[-13]C="""","""",IF(OR(AND(R[-13]C>100,R[-13]C[-13]-R[-2]C[-13]<=R9C4),AND(R[-13]C<=100,R[-13]C[-13]-R[-2]C[-13]<=(10*R[-13]C)-100)),R[-13]C,IF(AND(R[-13]C>100,R[-13]C[-13]-R[-2]C[-13]>R9C4),R[-13]C-((R[-13]C[-13]-R[-2]C[-13]-R9C4))*0.1,0)))"
    End If

    ActiveSheet.Calculate
    Range("M10:W21").Value = Range("M23:W34").Value
    Range("M23:W34").Clear
    
    Range("I5").FormulaR1C1 = "=MAX(R10C13:R21C23)"
    ActiveSheet.Calculate

    Range("I5").Value = Range("I5").Value
    
    Range("M7").FormulaR1C1 = _
        "=IF(MAX(R[3]C:R[13]C)<>R5C9,"""",IF(MAX(R[3]C:R[13]C)=R[3]C,R[4]C7,IF(MAX(R[3]C:R[13]C)=R[4]C,R[5]C7,IF(MAX(R[3]C:R[13]C)=R[5]C,R[6]C7,IF(MAX(R[3]C:R[13]C)=R[6]C,R[7]C7,IF(MAX(R[3]C:R[13]C)=R[7]C,R[8]C7,IF(MAX(R[3]C:R[13]C)=R[8]C,R[9]C7,"""")))))))"
    Range("M8").FormulaR1C1 = _
        "=IF(MAX(R[2]C:R[12]C)<>R5C9,"""",IF(MAX(R[2]C:R[12]C)=R[8]C,R[9]C7,IF(MAX(R[2]C:R[12]C)=R[9]C,R[10]C7,IF(MAX(R[2]C:R[12]C)=R[10]C,R[11]C7,IF(MAX(R[2]C:R[12]C)=R[11]C,R[12]C7,IF(MAX(R[2]C:R[12]C)=R[12]C,R[13]C7,""""))))))"
    Range("N7").FormulaR1C1 = _
        "=IF(MAX(R[3]C:R[12]C)<>R5C9,"""",IF(MAX(R[3]C:R[12]C)=R[3]C,R[5]C7,IF(MAX(R[3]C:R[12]C)=R[4]C,R[6]C7,IF(MAX(R[3]C:R[12]C)=R[5]C,R[7]C7,IF(MAX(R[3]C:R[12]C)=R[6]C,R[8]C7,IF(MAX(R[3]C:R[12]C)=R[7]C,R[9]C7,IF(MAX(R[3]C:R[12]C)=R[8]C,R[10]C7,"""")))))))"
    Range("N8").FormulaR1C1 = _
        "=IF(MAX(R[2]C:R[11]C)<>R5C9,"""",IF(MAX(R[2]C:R[11]C)=R[8]C,R[10]C7,IF(MAX(R[2]C:R[11]C)=R[9]C,R[11]C7,IF(MAX(R[2]C:R[11]C)=R[10]C,R[12]C7,IF(MAX(R[2]C:R[11]C)=R[11]C,R[13]C7,"""")))))"
    Range("O7").FormulaR1C1 = _
        "=IF(MAX(R[3]C:R[11]C)<>R5C9,"""",IF(MAX(R[3]C:R[11]C)=R[3]C,R[6]C7,IF(MAX(R[3]C:R[11]C)=R[4]C,R[7]C7,IF(MAX(R[3]C:R[11]C)=R[5]C,R[8]C7,IF(MAX(R[3]C:R[11]C)=R[6]C,R[9]C7,IF(MAX(R[3]C:R[11]C)=R[7]C,R[10]C7,IF(MAX(R[3]C:R[11]C)=R[8]C,R[11]C7,"""")))))))"
    Range("O8").FormulaR1C1 = _
        "=IF(MAX(R[2]C:R[10]C)<>R5C9,"""",IF(MAX(R[2]C:R[10]C)=R[8]C,R[11]C7,IF(MAX(R[2]C:R[10]C)=R[9]C,R[12]C7,IF(MAX(R[2]C:R[10]C)=R[10]C,R[13]C7,""""))))"
    Range("P7").FormulaR1C1 = _
        "=IF(MAX(R[3]C:R[10]C)<>R5C9,"""",IF(MAX(R[3]C:R[10]C)=R[3]C,R[7]C7,IF(MAX(R[3]C:R[10]C)=R[4]C,R[8]C7,IF(MAX(R[3]C:R[10]C)=R[5]C,R[9]C7,IF(MAX(R[3]C:R[10]C)=R[6]C,R[10]C7,IF(MAX(R[3]C:R[10]C)=R[7]C,R[11]C7,IF(MAX(R[3]C:R[10]C)=R[8]C,R[12]C7,"""")))))))"
    Range("P8").FormulaR1C1 = _
        "=IF(MAX(R[2]C:R[9]C)<>R5C9,"""",IF(MAX(R[2]C:R[9]C)=R[8]C,R[12]C7,IF(MAX(R[2]C:R[9]C)=R[9]C,R[13]C7,"""")))"
    Range("Q7").FormulaR1C1 = _
        "=IF(MAX(R[3]C:R[9]C)<>R5C9,"""",IF(MAX(R[3]C:R[9]C)=R[3]C,R[8]C7,IF(MAX(R[3]C:R[9]C)=R[4]C,R[9]C7,IF(MAX(R[3]C:R[9]C)=R[5]C,R[10]C7,IF(MAX(R[3]C:R[9]C)=R[6]C,R[11]C7,IF(MAX(R[3]C:R[9]C)=R[7]C,R[12]C7,IF(MAX(R[3]C:R[9]C)=R[8]C,R[13]C7,"""")))))))"
    Range("Q8").FormulaR1C1 = _
        "=IF(MAX(R[2]C:R[8]C)<>R5C9,"""",IF(MAX(R[2]C:R[8]C)=R[8]C,R[13]C7,""""))"
    Range("R7").FormulaR1C1 = _
        "=IF(MAX(R[3]C:R[8]C)<>R5C9,"""",IF(MAX(R[3]C:R[8]C)=R[3]C,R[9]C7,IF(MAX(R[3]C:R[8]C)=R[4]C,R[10]C7,IF(MAX(R[3]C:R[8]C)=R[5]C,R[11]C7,IF(MAX(R[3]C:R[8]C)=R[6]C,R[12]C7,IF(MAX(R[3]C:R[8]C)=R[7]C,R[13]C7,IF(MAX(R[3]C:R[8]C)=R[8]C,R[14]C7)))))))"
    Range("S7").FormulaR1C1 = _
        "=IF(MAX(R[3]C:R[7]C)<>R5C9,"""",IF(MAX(R[3]C:R[7]C)=R[3]C,R[10]C7,IF(MAX(R[3]C:R[7]C)=R[4]C,R[11]C7,IF(MAX(R[3]C:R[7]C)=R[5]C,R[12]C7,IF(MAX(R[3]C:R[7]C)=R[6]C,R[13]C7,IF(MAX(R[3]C:R[7]C)=R[7]C,R[14]C7))))))"
    Range("T7").FormulaR1C1 = _
        "=IF(MAX(R[3]C:R[6]C)<>R5C9,"""",IF(MAX(R[3]C:R[6]C)=R[3]C,R[11]C7,IF(MAX(R[3]C:R[6]C)=R[4]C,R[12]C7,IF(MAX(R[3]C:R[6]C)=R[5]C,R[13]C7,IF(MAX(R[3]C:R[6]C)=R[6]C,R[14]C7)))))"
    Range("U7").FormulaR1C1 = _
        "=IF(MAX(R[3]C:R[5]C)<>R5C9,"""",IF(MAX(R[3]C:R[5]C)=R[3]C,R[12]C7,IF(MAX(R[3]C:R[5]C)=R[4]C,R[13]C7,IF(MAX(R[3]C:R[5]C)=R[5]C,R[14]C7))))"
    Range("V7").FormulaR1C1 = _
        "=IF(MAX(R[3]C:R[4]C)<>R5C9,"""",IF(MAX(R[3]C:R[4]C)=R[3]C,R[13]C7,R[14]C7))"
    Range("W7").FormulaR1C1 = "=IF(R[3]C<>R5C9,"""",R[14]C7)"
    ActiveSheet.Calculate

    Range("M9:W9").FormulaR1C1 = "=IF(MAX(R[1]C:R[11]C)=R5C9,MAX(R[-2]C:R[-1]C),"""")"

    Range("L10:L20").FormulaR1C1 = "=IF(MAX(RC[1]:RC[11])=R5C9,RC[-5],"""")"
    ActiveSheet.Calculate

    Range("M7:W9").Value = Range("M7:W9").Value
    Range("L10:L20").Value = Range("L10:L20").Value
    
    Range("A1").FormulaR1C1 = "=MAX(R10C12:R21C12)"
    Range("A2").FormulaR1C1 = "=MAX(R9C13:R9C23)"
    ActiveSheet.Calculate

    Range("A1:A2").Value = Range("A1:A2").Value
    
    Range("G10:L60009,L7:W21").Clear
    
    Range("G10").FormulaR1C1 = "=IF(OR(RC[-6]=R1C1,RC[-6]=R2C1),RC[-6],"""")"
    Range("H10:K10").FormulaR1C1 = "=IF(RC[-1]<"""",RC[-6],"""")"
    
    Application.Calculation = xlCalculationManual
    With ActiveSheet
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:K10").AutoFill Destination:=.Range("G10:K" & LastRow), Type:=xlFillDefault
    End With
ActiveSheet.Calculate

    Range("G10:K60009").Value = Range("G10:K60009").Value
    Range("G10:K60009").Sort Key1:=Range("G10"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    
    Range("A1:E2").Value = Range("G10:K11").Value
    
    Range("B5:I6,G10:K12").Clear
    Range("F2").FormulaR1C1 = _
        "=6371*ACOS(SIN(R[-1]C[-4])*SIN(RC[-4])+COS(R[-1]C[-4])*COS(RC[-4])*COS(RC[-3]-R[-1]C[-3]))"
    ActiveSheet.Calculate

    Range("F2").Value = Range("F2").Value
    
   If Range("E9") <> "PR" Then
    Range("G2").FormulaR1C1 = _
        "=IF(OR(AND(RC[-1]>100,R[-1]C[-3]-RC[-3]<=R9C4),AND(RC[-1]<=100,R[-1]C[-3]-RC[-3]<=10*RC[-1])),RC[-1],IF(AND(RC[-1]>100,R[-1]C[-3]-RC[-3]>R9C4),RC[-1]-((R[-1]C[-3]-RC[-3]-R9C4)*0.1),0))"
   ElseIf Range("E9") = "PR" Then
    Range("G2").FormulaR1C1 = _
        "=IF(OR(AND(RC[-1]>100,R[-1]C[-3]-RC[-3]<=R9C4),AND(RC[-1]<=100,R[-1]C[-3]-RC[-3]<=(10*RC[-1])-100)),RC[-1],IF(AND(RC[-1]>100,R[-1]C[-3]-RC[-3]>R9C4),RC[-1]-((R[-1]C[-3]-RC[-3]-R9C4)*0.1),0))"
   End If
    ActiveSheet.Calculate

    Range("G2").Value = Range("G2").Value
    Range("F2").Value = Range("G2").Value
    
    'PaynePain added 5/19/15
    Range("F6:K6").FormulaR1C1 = "=MAX(R[4]C:R[60003]C)"
    Range("F7:K7").FormulaR1C1 = "=MAX(R[3]C[7]:R[60002]C[7])"
    
    Range("F10").FormulaR1C1 = _
        "=IF(RC[-5]>=R2C1,"""",IF(RC[-2]-R2C4<R9C4,6371*ACOS(SIN(RC[-4])*SIN(R2C2)+COS(RC[-4])*COS(R2C2)*COS(R2C3-RC[-3])),6371*ACOS(SIN(RC[-4])*SIN(R2C2)+COS(RC[-4])*COS(R2C2)*COS(R2C3-RC[-3]))-((RC[-2]-R2C4-R9C4)*0.1)))"
    Range("G10").FormulaR1C1 = "=IF(RC[-1]="""","""",IF(RC[-1]=R6C6,RC[-6],""""))"
    Range("H10:K10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-6],"""")"
    'Copy Ref A
    Application.Calculation = xlCalculationManual
    With ActiveSheet
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("F10:K10").AutoFill Destination:=.Range("F10:K" & LastRow), Type:=xlFillDefault
    End With
 ActiveSheet.Calculate
   
    Range("F6:K6").Value = Range("F6:K6").Value
    Range("F6").Clear
    
    Range("M10").FormulaR1C1 = _
        "=IF(RC[-12]<R6C7,"""",IF(R6C10-RC[-9]<R9C4,6371*ACOS(SIN(RC[-11])*SIN(R6C8)+COS(RC[-11])*COS(R6C8)*COS(R6C9-RC[-10])),6371*ACOS(SIN(RC[-11])*SIN(R6C8)+COS(RC[-11])*COS(R6C8)*COS(R6C9-RC[-10]))-((R6C10-RC[-9]-R9C4)*0.1)))"
    Range("N10").FormulaR1C1 = "=IF(RC[-1]=R7C6,RC[-13],"""")"
    Range("O10:R10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-13],"""")"
   'Copy Ref A
    Application.Calculation = xlCalculationManual
    With ActiveSheet
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("M10:R10").AutoFill Destination:=.Range("M10:R" & LastRow), Type:=xlFillDefault
    End With
 ActiveSheet.Calculate
   
    Range("F7:K7").Value = Range("F7:K7").Value
    
    If Range("F7") > Range("F2") Then
        Range("A1:E1").Value = Range("G6:K6").Value
        Range("A2:E2").Value = Range("G7:K7").Value
        Range("F2").Value = Range("F7").Value
        
        If Range("E9") <> "PR" Then
            Range("G2").FormulaR1C1 = _
                "=IF(OR(AND(RC[-1]>100,R[-1]C[-3]-RC[-3]<=R9C4),AND(RC[-1]<=100,R[-1]C[-3]-RC[-3]<=10*RC[-1])),RC[-1],IF(AND(RC[-1]>100,R[-1]C[-3]-RC[-3]>R9C4),RC[-1]-((R[-1]C[-3]-RC[-3]-R9C4)*0.1),0))"
        ElseIf Range("E9") = "PR" Then
            Range("G2").FormulaR1C1 = _
            "=IF(OR(AND(RC[-1]>100,R[-1]C[-3]-RC[-3]<=R9C4),AND(RC[-1]<=100,R[-1]C[-3]-RC[-3]<=(10*RC[-1])-100)),RC[-1],IF(AND(RC[-1]>100,R[-1]C[-3]-RC[-3]>R9C4),RC[-1]-((R[-1]C[-3]-RC[-3]-R9C4)*0.1),0))"
        End If
    End If
    ActiveSheet.Calculate

    Range("F6:R60009").Clear

    'Re-calc Start
    Range("A5").Value = 1 / 24
    
    Range("I8:N8").FormulaR1C1 = "=MAX(R[2]C:R[60001]C)"
    
    Range("G10").FormulaR1C1 = _
        "=IF(AND(RC[-6]>=R1C1-R5C1,RC[-6]<=R1C1+R5C1),6371*ACOS(SIN(RC[-5])*SIN(R2C2)+COS(RC[-5])*COS(R2C2)*COS(R2C3-RC[-4])),"""")"
    Range("H10").FormulaR1C1 = _
        "=IF(RC[-1]="""","""",IF(RC[-1]>100,R9C4,IF(AND(RC[-1]<=100,R9C5<>""PR""),10*RC[-1],IF(AND(RC[-1]<=100,R9C5=""PR""),10*RC[-1]-100))))"
    Range("I10").FormulaR1C1 = "=IF(RC[-5]-R2C4<=RC[-1],RC[-2],RC[-2]-((RC[-5]-R2C4-RC[-1])*0.1))"
    Range("J10").FormulaR1C1 = "=IF(RC[-1]=R8C9,RC[-9],"""")"
    Range("K10:N10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-9],"""")"
    
     Application.Calculation = xlCalculationManual
    With ActiveSheet
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:N10").AutoFill Destination:=.Range("G10:N" & LastRow), Type:=xlFillDefault
    End With
ActiveSheet.Calculate
    
  If Range("I8") > Range("F2") Then
    Range("A1:E1").Value = Range("J8:N8").Value
    Range("F2").FormulaR1C1 = _
        "=6371*ACOS(SIN(R[-1]C[-4])*SIN(RC[-4])+COS(R[-1]C[-4])*COS(RC[-4])*COS(RC[-3]-R[-1]C[-3]))"
    ActiveSheet.Calculate
    Range("F2").Value = Range("F2").Value
    
    If Range("E9") <> "PR" Then
    Range("G2").FormulaR1C1 = _
        "=IF(OR(AND(RC[-1]>100,R[-1]C[-3]-RC[-3]<=R9C4),AND(RC[-1]<=100,R[-1]C[-3]-RC[-3]<=10*RC[-1])),RC[-1],IF(AND(RC[-1]>100,R[-1]C[-3]-RC[-3]>R9C4),RC[-1]-((R[-1]C[-3]-RC[-3]-R9C4)*0.1),0))"
    ElseIf Range("E9") = "PR" Then
    Range("G2").FormulaR1C1 = _
        "=IF(OR(AND(RC[-1]>100,R[-1]C[-3]-RC[-3]<=R9C4),AND(RC[-1]<=100,R[-1]C[-3]-RC[-3]<=(10*RC[-1])-100)),RC[-1],IF(AND(RC[-1]>100,R[-1]C[-3]-RC[-3]>R9C4),RC[-1]-((R[-1]C[-3]-RC[-3]-R9C4)*0.1),0))"
    End If
    ActiveSheet.Calculate

    Range("G2").Value = Range("G2").Value
    Range("F2").Value = Range("G2").Value
    Range("G2").Clear
   End If

    'recalc fini
    Range("G10").FormulaR1C1 = _
        "=IF(AND(RC[-6]>=R2C1-R5C1,RC[-6]<=R2C1+R5C1),6371*ACOS(SIN(RC[-5])*SIN(R1C2)+COS(RC[-5])*COS(R1C2)*COS(R1C3-RC[-4])),"""")"
    Range("H10").FormulaR1C1 = _
        "=IF(RC[-1]="""","""",IF(RC[-1]>100,R9C4,IF(AND(RC[-1]<=100,R9C5<>""PR""),10*RC[-1],IF(AND(RC[-1]<=100,R9C5=""PR""),10*RC[-1]-100))))"
    Range("I10").FormulaR1C1 = "=IF(R1C4-RC[-5]<=RC[-1],RC[-2],RC[-2]-((R1C4-RC[-5]-RC[-1])*0.1))"
    Application.Calculation = xlCalculationManual
    With ActiveSheet
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:I10").AutoFill Destination:=.Range("G10:I" & LastRow), Type:=xlFillDefault
    End With
ActiveSheet.Calculate

If Range("I8") > Range("F2") Then
    Range("A2:E2").Value = Range("J8:N8").Value
    Range("F2").FormulaR1C1 = _
        "=6371*ACOS(SIN(R[-1]C[-4])*SIN(RC[-4])+COS(R[-1]C[-4])*COS(RC[-4])*COS(RC[-3]-R[-1]C[-3]))"
    Range("F2").Value = Range("F2").Value
 If Range("E9") <> "PR" Then
    Range("G2").FormulaR1C1 = _
        "=IF(OR(AND(RC[-1]>100,R[-1]C[-3]-RC[-3]<=R9C4),AND(RC[-1]<=100,R[-1]C[-3]-RC[-3]<=10*RC[-1])),RC[-1],IF(AND(RC[-1]>100,R[-1]C[-3]-RC[-3]>R9C4),RC[-1]-((R[-1]C[-3]-RC[-3]-R9C4)*0.1),0))"
 ElseIf Range("E9") = "PR" Then
    Range("G2").FormulaR1C1 = _
        "=IF(OR(AND(RC[-1]>100,R[-1]C[-3]-RC[-3]<=R9C4),AND(RC[-1]<=100,R[-1]C[-3]-RC[-3]<=(10*RC[-1])-100)),RC[-1],IF(AND(RC[-1]>100,R[-1]C[-3]-RC[-3]>R9C4),RC[-1]-((R[-1]C[-3]-RC[-3]-R9C4)*0.1),0))"
 End If
    ActiveSheet.Calculate
End If
 
  'CK reverse
    Range("G10").FormulaR1C1 = "=IF(RC[-6]>=R2C1,6371*ACOS(SIN(RC[-5])*SIN(R2C2)+COS(RC[-5])*COS(R2C2)*COS(R2C3-RC[-4])),"""")"
    Range("H10").FormulaR1C1 = "=IF(RC[-1]="""","""",IF(RC[-1]>100,R9C4,IF(AND(RC[-1]<=100,R9C5<>""PR""),10*RC[-1],IF(AND(RC[-1]<=100,R9C5=""PR""),10*RC[-1]-100))))"
    Range("I10").FormulaR1C1 = "=IF(R2C4-RC[-5]<=RC[-1],RC[-2],RC[-2]-((R2C4-RC[-5]-RC[-1])*0.1))"
    'Copy Ref a
     With ActiveSheet
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:I10").AutoFill Destination:=.Range("G10:I" & LastRow), Type:=xlFillDefault
    End With
ActiveSheet.Calculate
If Range("I8") > Range("F2") Then
    Range("A1:E1").Value = Range("J8:N8").Value
    Range("A1:E2").Sort Key1:=Range("A1"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    
    Range("F2").FormulaR1C1 = _
        "=6371*ACOS(SIN(R[-1]C[-4])*SIN(RC[-4])+COS(R[-1]C[-4])*COS(RC[-4])*COS(RC[-3]-R[-1]C[-3]))"
    Range("F2").Value = Range("F2").Value
    If Range("E9") <> "PR" Then
        Range("G2").FormulaR1C1 = _
            "=IF(OR(AND(RC[-1]>100,R[-1]C[-3]-RC[-3]<=R9C4),AND(RC[-1]<=100,R[-1]C[-3]-RC[-3]<=10*RC[-1])),RC[-1],IF(AND(RC[-1]>100,R[-1]C[-3]-RC[-3]>R9C4),RC[-1]-((R[-1]C[-3]-RC[-3]-R9C4)*0.1),0))"
    ElseIf Range("E9") = "PR" Then
        Range("G2").FormulaR1C1 = _
            "=IF(OR(AND(RC[-1]>100,R[-1]C[-3]-RC[-3]<=R9C4),AND(RC[-1]<=100,R[-1]C[-3]-RC[-3]<=(10*RC[-1])-100)),RC[-1],IF(AND(RC[-1]>100,R[-1]C[-3]-RC[-3]>R9C4),RC[-1]-((R[-1]C[-3]-RC[-3]-R9C4)*0.1),0))"
    End If
 End If
 ActiveSheet.Calculate
 
    Range("A5,G8:N60009").Clear

  'Ck w/in 1 minute
    Range("A7").Value = 6.94444444444444E-04
    
    'FINIS
    Range("G10").FormulaR1C1 = "=IF(AND(RC[-6]>=R2C1-R7C1,RC[-6]<=R2C1+R7C1),RC[-6],"""")"
    Range("H10:K10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-6],"""")"
    'Copy Ref A Value Sort COPY to N1"
    Application.Calculation = xlCalculationManual
    With ActiveSheet
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:K10").AutoFill Destination:=.Range("G10:K" & LastRow), Type:=xlFillDefault
    End With
ActiveSheet.Calculate
    
    Range("G10:K60009").Value = Range("G10:K60009").Value
    Range("G10:K60009").Sort Key1:=Range("G10"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal

    Range("G10:K133").Copy
    Range("N1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    Application.CutCopyMode = False
    Range("G10:K133").Clear
    
    'STARTS
    Range("G10").FormulaR1C1 = "=IF(AND(RC[-6]>=R1C1-R7C1,RC[-6]<=R1C1+R7C1),RC[-6],"""")"
    Range("H10:K10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-6],"""")"
    'Copy Ref A Value Sort
    Application.Calculation = xlCalculationManual
    With ActiveSheet
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:K10").AutoFill Destination:=.Range("G10:K" & LastRow), Type:=xlFillDefault
    End With
ActiveSheet.Calculate
    
    Range("G10:K60009").Value = Range("G10:K60009").Value
    Range("G10:K60009").Sort Key1:=Range("G10"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal

    'MATRIX REVISED THIS SECTION!! N6,L10
    If Range("F2") > 100 Then
        Range("N10:EE132").FormulaR1C1 = _
            "=IF(OR(R1C="""",RC11="""",RC9=R3C),"""",IF(AND(6371*ACOS(SIN(RC8)*SIN(R2C)+COS(RC8)*COS(R2C)*COS(R3C-RC9))>100,RC10-R4C<=R9C4),6371*ACOS(SIN(RC8)*SIN(R2C)+COS(RC8)*COS(R2C)*COS(R3C-RC9)),6371*ACOS(SIN(RC8)*SIN(R2C)+COS(RC8)*COS(R2C)*COS(R3C-RC9))-((RC10-R4C-R9C4)*0.1)))"
    ElseIf Range("F2") <= 100 And Range("E9") <> "PR" Then
        Range("N10:EE132").FormulaR1C1 = _
        "=IF(OR(RC11="""",R1C="""",RC9=R3C),"""",IF(RC10-R4C<=63710*ACOS(SIN(RC8)*SIN(R2C)+COS(RC8)*COS(R2C)*COS(R3C-RC9)),6371*ACOS(SIN(RC8)*SIN(R2C)+COS(RC8)*COS(R2C)*COS(R3C-RC9)),0))"
    ElseIf Range("F2") <= 100 And Range("E9") = "PR" Then
        Range("N10:EE132").FormulaR1C1 = _
        "=IF(OR(RC11="""",R1C="""",RC9=R3C),"""",IF(RC10-R4C<=63710*ACOS(SIN(RC8)*SIN(R2C)+COS(RC8)*COS(R2C)*COS(R3C-RC9))-100,6371*ACOS(SIN(RC8)*SIN(R2C)+COS(RC8)*COS(R2C)*COS(R3C-RC9)),0))"
    End If

    Range("I5").FormulaR1C1 = "=MAX(R10C14:R132C135)"
    ActiveSheet.Calculate
    Range("I5").Value = Range("I5").Value
    Range("N6:EE6").FormulaR1C1 = "=IF(AND(R[-5]C<>"""",MAX(R[4]C:R[124]C)=R5C9),R[-5]C,"""")"
    ActiveSheet.Calculate
    Range("N6:EE6").Value = Range("N6:EE6").Value
    Range("L10:L132").FormulaR1C1 = "=IF(MAX(RC[2]:RC[123])=R5C9,RC[-5],"""")"
    ActiveSheet.Calculate
    Range("L10:L132").Value = Range("L10:L132").Value
    
    Range("N10:EE132").Clear
    
    Range("N7:EE10").FormulaR1C1 = "=IF(R[-1]C<>"""",R[-5]C,"""")"
    
    Range("H2").FormulaR1C1 = "=MAX(R[4]C[6]:R[4]C[128])"
    Range("I2").FormulaR1C1 = "=MAX(R[5]C[5]:R[5]C[127])"
    Range("J2").FormulaR1C1 = "=MAX(R[6]C[4]:R[6]C[126])"
    Range("K2").FormulaR1C1 = "=MAX(R[7]C[3]:R[7]C[125])"
    Range("L2").FormulaR1C1 = "=MAX(R[8]C[2]:R[8]C[124])"
    ActiveSheet.Calculate
    Range("H2:L2").Value = Range("H2:L2").Value
    
    Range("M10:Q122").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-5],"""")"

    Range("H1:L1").FormulaR1C1 = "=MAX(R[9]C[4]:R[131]C[4])"
    ActiveSheet.Calculate
    Range("H1:L1").Value = Range("H1:L1").Value
    Columns("N:EE").Clear
    Range("G5:M300").Clear
    
    Range("M2").FormulaR1C1 = "=6371*ACOS(SIN(R[-1]C[-4])*SIN(RC[-4])+COS(R[-1]C[-4])*COS(RC[-4])*COS(RC[-3]-R[-1]C[-3]))"
    ActiveSheet.Calculate
    Range("M2").Value = Range("M2").Value
    
   If Range("E9") <> "PR" Then
    Range("N2").FormulaR1C1 = _
        "=IF(OR(AND(RC[-1]>100,R[-1]C[-3]-RC[-3]<=R9C4),AND(RC[-1]<=100,R[-1]C[-3]-RC[-3]<=10*RC[-1])),RC[-1],IF(AND(RC[-1]>100,R[-1]C[-3]-RC[-3]>R9C4),RC[-1]-((R[-1]C[-3]-RC[-3]-R9C4)*0.1),0))"
   ElseIf Range("E9") = "PR" Then
    Range("N2").FormulaR1C1 = _
        "=IF(OR(AND(RC[-1]>100,R[-1]C[-3]-RC[-3]<=R9C4),AND(RC[-1]<=100,R[-1]C[-3]-RC[-3]<=(10*RC[-1])-100)),RC[-1],IF(AND(RC[-1]>100,R[-1]C[-3]-RC[-3]>R9C4),RC[-1]-((R[-1]C[-3]-RC[-3]-R9C4)*0.1),0))"
   End If
    ActiveSheet.Calculate
    Range("M2").Value = Range("N2").Value
    Range("N2").Clear
       
    If Range("M2") > Range("F2") Then
        Range("A1:F2").Value = Range("H1:M2").Value
        Range("A7,H1:M2").Clear
    End If
    
    Sheets("Sheet2").Activate
    Range("H1:M2").Value = Sheets("Sheet3").Range("A1:F2").Value
    
    If Range("H1") <> "" Then
        Range("A1:F2").Value = Range("H1:M2").Value
        Range("H1:M2").Clear
        Sheets("Sheet3").Range("A1:F2").Clear

    End If

    Sheets("Tasks").Activate
    Sheets("Tasks").Range("C10:E11").Value = Sheets("Sheet2").Range("A1:C2").Value
    Sheets("Tasks").Activate
    Range("F10:G11").FormulaR1C1 = "=DEGREES(RC[-2])"
    ActiveSheet.Calculate
    Range("F10:G11").Value = Range("F10:G11").Value
    Sheets("Tasks").Range("H10:J11").Value = Sheets("Sheet2").Range("D1:F2").Value

Sheets("Sheet2").Activate
Application.Run "F.xlsm!NewOR1"
End Sub