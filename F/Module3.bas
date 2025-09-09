Option Explicit
#If VBA7 And Win64 Then
    Dim LastRow As LongPtr
#Else
    Dim LastRow As Long
#End If

Sub ORDS()
'
' Works for Ramy 06
'
  Application.ScreenUpdating = False
    Range("A8:C8").FormulaR1C1 = "=MAX(R[2]C:R[10001]C)"
    Range("B7:C7").FormulaR1C1 = "=MIN(R[3]C:R[10002]C)"
   
    Range("G10").FormulaR1C1 = _
        "=IF(OR(RC[-6]=R8C1,RC[-5]=R7C2,RC[-5]=R8C2,RC[-4]=R7C3,RC[-4]=R8C3),RC[-6],"""")"
    Range("H10:K10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-6],"""")"
    
 With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:K10").AutoFill Destination:=.Range("G10:K" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    
    Range("G10:K10009").Value = Range("G10:K10009").Value
    Range("G10:K10009").Sort Key1:=Range("G10"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    'Remove ords within 2 minutes & 1km of each other
    Range("A7").Value = 1.38888888888889E-03
    Range("L11:L21").FormulaR1C1 = "=IF(RC[-5]<>"""",6371*ACOS(SIN(R[-1]C[-4])*SIN(RC[-4])+COS(R[-1]C[-4])*COS(RC[-4])*COS(RC[-3]-R[-1]C[-3])),"""")"
    Range("M10:Q10").Value = Range("G10:K10").Value
    Range("F11:F21").FormulaR1C1 = _
        "=IF(AND(RC[1]<>"""",RC[1]-R[-1]C[1]<R7C[-5],RC[6]<=1),""X"","""")"
    Range("M11:M21").FormulaR1C1 = "=IF(AND(RC[-7]<>""X"",RC[-6]<>""""),RC[-6],"""")"
    Range("N11:Q21").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-6],"""")"
    ActiveSheet.Calculate
    Range("M10:Q21").Value = Range("M10:Q21").Value
    Range("M10:Q21").Sort Key1:=Range("M10"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    Range("R11:R21").FormulaR1C1 = _
        "=IF(RC[-1]<>"""",6371*ACOS(SIN(R[-1]C[-4])*SIN(RC[-4])+COS(R[-1]C[-4])*COS(RC[-4])*COS(RC[-3]-R[-1]C[-3])),"""")"
    ActiveSheet.Calculate
    Range("A1:F5").Value = Range("M10:R14").Value
    Range("F6").FormulaR1C1 = "=SUM(R[-5]C:R[-1]C)"
    Range("F7").FormulaR1C1 = _
        "=IF(AND(R[-2]C[-2]<>"""",R[-6]C[-2]-R[-2]C[-2]>R9C4),R[-1]C-((R[-6]C[-2]-R[-2]C[-2]-R9C4)*0.1),IF(AND(R[-2]C[-2]="""",R[-6]C[-2]-R[-3]C[-2]>R9C4),R[-1]C-((R[-6]C[-2]-R[-3]C[-2]-R9C4)*0.1),IF(AND(R[-2]C[-2]="""",R[-3]C[-2]="""",R[-6]C[-2]-R[-4]C[-2]>R9C4),R[-1]C-((R[-6]C[-2]-R[-4]C[-2]-R9C4)*0.1),R[-1]C)))"
    ActiveSheet.Calculate
    Range("F6:F7").Value = Range("F6:F7").Value
    Range("A7:C8,F10:R21").Clear
    Sheets("Tasks").Range("A40:E43").Value = Sheets("Sheet2").Range("A1:E4").Value
    Range("G8:L8").FormulaR1C1 = "=MAX(R[2]C:R[10001]C)"
    
    Range("G10").FormulaR1C1 = _
        "=IF(RC[-6]<R1C[-6],6371*ACOS(SIN(RC[-5])*SIN(R1C2)+COS(RC[-5])*COS(R1C2)*COS(R1C3-RC[-4])),"""")"
    Range("H10").FormulaR1C1 = "=IF(RC[-1]=R8C[-1],RC[-7],"""")"
    Range("I10:L10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-7],"""")"
    
 With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:L10").AutoFill Destination:=.Range("G10:L" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    
    Range("G8:L8").Value = Range("G8:L8").Value
    
  If Range("G8") > Range("F5") Then
    Range("A6:E6").Value = Range("H8:L8").Value
    Range("A5:E5").Clear
    Range("A1:E6").Sort Key1:=Range("A1"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    Range("F2:F5").FormulaR1C1 = _
        "=IF(RC[-5]<>"""",6371*ACOS(SIN(RC[-4])*SIN(R[-1]C[-4])+COS(RC[-4])*COS(R[-1]C[-4])*COS(R[-1]C[-3]-RC[-3])),"""")"
    Range("F6").FormulaR1C1 = "=SUM(R[-4]C:R[-1]C)"
    Range("F7").FormulaR1C1 = _
        "=IF(R[-6]C[-2]-R[-2]C[-2]>R9C4,R[-1]C-((R[-6]C[-2]-R[-2]C[-2]-R9C4)*0.1),R[-1]C)"
    ActiveSheet.Calculate
    Range("F2:F7").Value = Range("F2:F7").Value
  End If
    Columns("G:L").Clear

    Sheets("Sheet2").Range("A8").Value = Sheets("Tasks").Range("J11").Value
    Range("A7").FormulaR1C1 = "=MAX(R2C6:R5C6)"
      
 If Range("A7") < Range("A8") Then
    Application.Run "F.xlsm!Ordsa"

 ElseIf Range("A7") >= Range("A8") Then
 
    Range("G2").FormulaR1C1 = "=IF(RC[-1]=MIN(R2C6:R5C6),""A1"","""")"
    Range("G3").FormulaR1C1 = "=IF(RC[-1]=MIN(R2C6:R5C6),""A2"","""")"
    Range("G4").FormulaR1C1 = "=IF(RC[-1]=MIN(R2C6:R5C6),""A3"","""")"
    Range("G5").FormulaR1C1 = "=IF(RC[-1]=MIN(R2C6:R5C6),""A4"","""")"
    
    If Range("G2") = "A1" Then
        Range("H2:L2").Value = Range("A1:E1").Value
        Range("P1:T1").Value = Range("A1:E1").Value
        Range("P3:T4").Value = Range("A3:E4").Value
        Range("G10").FormulaR1C1 = _
        "=IF(AND(RC[-6]>R1C1,RC[-6]<R3C1),6371*ACOS(SIN(RC[-5])*SIN(R1C2)+COS(RC[-5])*COS(R1C2)*COS(RC[-4]-R1C3)),"""")"
    ElseIf Range("G3") = "A2" Then
        Range("H2:L2").Value = Range("A2:E2").Value
        Range("P1:T2").Value = Range("A1:E2").Value
        Range("P4:T4").Value = Range("A4:E4").Value
        Range("G10").FormulaR1C1 = _
        "=IF(AND(RC[-6]>R2C1,RC[-6]<R4C1),6371*ACOS(SIN(RC[-5])*SIN(R2C2)+COS(RC[-5])*COS(R2C2)*COS(RC[-4]-R2C3)),"""")"
    ElseIf Range("G4") = "A3" Then
        Range("H2:L2").Value = Range("A3:E3").Value
        Range("P1:T3").Value = Range("A1:E3").Value
        Range("G10").FormulaR1C1 = _
        "=IF(AND(RC[-6]>R3C1,RC[-6]<R5C1),6371*ACOS(SIN(RC[-5])*SIN(R3C2)+COS(RC[-5])*COS(R3C2)*COS(RC[-4]-R3C3)),"""")"
    End If
    ActiveSheet.Calculate
    If Range("G5") = "A4" Then
        Range("A1:G7").Clear
        Application.Run "F.xlsm!TP3a"
    ElseIf Range("G5") <> "A4" Then
    Range("H10").FormulaR1C1 = "=IF(RC[-1]=R8C7,RC[-7],"""")"
    Range("I10:L10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-7],"""")"
    
    Range("G8:L8").FormulaR1C1 = "=MAX(R[2]C:R[10001]C)"
    
    Application.Calculation = xlCalculationManual
 With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:L10").AutoFill Destination:=.Range("G10:L" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    
    Range("H3:L3").Value = Range("H8:L8").Value
    
    Range("G8:L10009").Clear
    
    Range("I4").FormulaR1C1 = _
        "=ACOS(SIN(R2C)*SIN(R3C)+COS(R2C)*COS(R3C)*COS(R3C[1]-R2C[1]))"
    Range("I5").FormulaR1C1 = _
        "=ACOS((SIN(R3C)-SIN(R2C)*COS(R[-1]C))/(SIN(R[-1]C)*COS(R2C)))"
    'MAXIMUM off course after Selected ORD
    Range("G10").FormulaR1C1 = _
        "=IF(RC[-4]=R2C[3],"""",IF(AND(RC1>R2C[1],RC1<R3C[1]),ACOS(SIN(R2C[2])*SIN(RC2)+COS(R2C[2])*COS(RC2)*COS(RC3-R2C[3])),""""))"
    Range("H10").FormulaR1C1 = _
        "=IF(RC[-1]<>"""",ACOS((SIN(RC2)-SIN(R2C[1])*COS(RC[-1]))/(SIN(RC[-1])*COS(R2C[1]))),"""")"
    Range("I10").FormulaR1C1 = _
        "=IF(RC[-1]<>"""",ABS((ASIN(SIN(RC[-2])*SIN(RC[-1]-R5C))*180*60*1.852/PI())),"""")"
    Range("J10").FormulaR1C1 = "=IF(RC[-1]=R4C[1],RC1,"""")"
    Range("K10:N10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-9],"""")"
    
    Range("K4").FormulaR1C1 = "=MAX(R[6]C[-2]:R[10005]C[-2])"
    Range("L4").FormulaR1C1 = "=MAX(R[6]C[-2]:R[10005]C[-2])"
    Range("J8:N8").FormulaR1C1 = "=MAX(R[2]C:R[10001]C)"
    
    Application.Calculation = xlCalculationManual
 With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:N10").AutoFill Destination:=.Range("G10:N" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    
    If Range("G2") = "A1" Then
        Range("P2:T2").Value = Range("J8:N8").Value
    ElseIf Range("G3") = "A2" Then
        Range("P3:T3").Value = Range("J8:N8").Value
    ElseIf Range("G4") = "A3" Then
        Range("P4:T4").Value = Range("J8:N8").Value
    End If
    
    Range("G10:N10009").Clear
    Range("U2:U4").FormulaR1C1 = "=IF(SUM(RC[-5]:RC[-1])>0,6371*ACOS(SIN(R[-1]C[-4])*SIN(RC[-4])+COS(R[-1]C[-4])*COS(RC[-4])*COS(RC[-3]-R[-1]C[-3])),0)"
    Range("G8:L8").FormulaR1C1 = "=MAX(R[2]C:R[10001]C)"
    Range("G10").FormulaR1C1 = _
         "=IF(RC[-6]>R4C16,6371*ACOS(SIN(RC[-5])*SIN(R4C17)+COS(RC[-5])*COS(R4C17)*COS(R4C18-RC[-4])),"""")"
    Range("H10").FormulaR1C1 = "=IF(RC[-1]=R8C7,RC[-7],"""")"
    Range("I10:L10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-7],"""")"
    
    Application.Calculation = xlCalculationManual
 With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:L10").AutoFill Destination:=.Range("G10:L" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    
    Range("P5:T5").Value = Range("H8:L8").Value
   
    Range("U5").FormulaR1C1 = _
        "=IF(AND(SUM(RC[-5]:RC[-1])>0,SUM(R[-1]C[-5]:R[-1]C[-1])>0),6371*ACOS(SIN(R[-1]C[-4])*SIN(RC[-4])+COS(R[-1]C[-4])*COS(RC[-4])*COS(RC[-3]-R[-1]C[-3])),0)"
    Range("U6").FormulaR1C1 = "=SUM(R[-4]C:R[-1]C)"
    Range("U7").FormulaR1C1 = _
        "=IF(R[-6]C[-2]-R[-2]C[-2]>R9C4,R[-1]C-((R[-6]C[-2]-R[-2]C[-2]-R9C4)*0.1),R[-1]C)"
    ActiveSheet.Calculate
    Range("U2:U7").Value = Range("U2:U7").Value
    Range("G8:L10009").Clear
    Range("O2").FormulaR1C1 = "=IF(OR(RC[1]<R[-1]C[1],R[1]C[1]<RC[1],R[2]C[1]<R[1]C[1]),""NO"","""")"
    ActiveSheet.Calculate
    If Range("O2") = "NO" Then
        Range("O1:U7").Clear
    End If
 End If
    
    Application.Run "F.xlsm!ORDS1"
    Application.Run "F.xlsm!ORDS3"
    Application.Run "F.xlsm!ORDS2"
  
  If Range("U7") > Range("F7") Then
    Range("A1:E5").Value = Range("P1:T5").Value
    Range("F2:F7").Value = Range("U2:U7").Value
  End If
 End If
    
    Range("G2:L5,A7:E8,O1:U7").Clear
  
  Application.Run "F.xlsm!DH"
  Application.Run "F.xlsm!XCA"
  Application.Run "F.xlsm!ORDS4"
  Application.Run "F.xlsm!AllOrds"
  Application.Run "F.xlsm!TP3a"
  Application.Run "F.xlsm!TriZords"
  Application.Run "F.xlsm!Serk"
  Application.Run "F.xlsm!Essex"
  Application.Run "F.xlsm!Siba"
  
 If Range("F7") < Range("F6") Then
   Application.Run "F.xlsm!ORDS5"
 End If

    Application.Run "F.xlsm!CK3TP"
    Application.Run "F.xlsm!ORDS5"
  
  If Range("A9") = "REF" Then
    Sheets("Sheet3").Activate
    Sheets("Sheet3").Range("A1:F7").Value = Sheets("Sheet2").Range("A1:F7").Value
    Application.Run "F.xlsm!CK3TP"
    Application.Run "F.xlsm!ORDS5a"
    Sheets("Sheet2").Range("H7").Value = Sheets("Sheet3").Range("F7").Value
    Sheets("Sheet2").Activate
    If Range("H7") > Range("F7") Then
        Sheets("Sheet2").Range("A1:F7").Value = Sheets("Sheet3").Range("A1:F7").Value
    End If
    Sheets("Sheet2").Range("H7").Clear
  End If
  
   Sheets("Tasks").Activate
   Sheets("Tasks").Range("C20:E24").Value = Sheets("Sheet2").Range("A1:C5").Value
   Sheets("Tasks").Range("H20:I24").Value = Sheets("Sheet2").Range("D1:E5").Value
   Sheets("Tasks").Range("J24").Value = Sheets("Sheet2").Range("F7").Value
   Range("F20:G24").FormulaR1C1 = "=DEGREES(RC[-2])"
   ActiveSheet.Calculate
   Range("F20:G24").Value = Range("F20:G24").Value
   
   Application.Run "F.xlsm!Vince"
   Sheets("Sheet3").Range("A1:F7").Clear
   Sheets("Sheet3").Columns("G:L").Clear
   Sheets("Sheet2").Activate
   Range("A1:F8").Clear
   Range("A1").Activate
 
  Application.Run "F.xlsm!Triangle1"
End Sub

Sub ORDS1()
'
' Compares ords vs St Dist LEG and O&R - Fini Works for Doogie Amended 9/29/15 for better OCStd
' Amended G10 in FIRST Max OffCourse to avoid longitude conflict
'
Application.ScreenUpdating = False
    Sheets("Sheet2").Range("P9:R9").Value = Sheets("TASKS").Range("C10:E10").Value
    Sheets("Sheet2").Range("S9:T9").Value = Sheets("TASKS").Range("H10:I10").Value
    Sheets("Sheet2").Range("P10:R10").Value = Sheets("TASKS").Range("C11:E11").Value
    Sheets("Sheet2").Range("S10:T10").Value = Sheets("TASKS").Range("H11:I11").Value
    Sheets("Sheet2").Range("P12:R12").Value = Sheets("TASKS").Range("C16:E16").Value
    Sheets("Sheet2").Range("S12:T12").Value = Sheets("Tasks").Range("H16:I16").Value
    
    Range("H2:L2").Value = Range("P10:T10").Value
    Range("H3:L3").Value = Range("P12:T12").Value
    
    'MAX Off course from St dist FINI to OR FINI
    Range("G10").FormulaR1C1 = "=IF(AND(RC1>R2C[1],RC[-6]<R3C8,RC[-4]<>R2C10),ACOS(SIN(R2C[2])*SIN(RC2)+COS(R2C[2])*COS(RC2)*COS(RC3-R2C[3])),"""")"
    Range("H10").FormulaR1C1 = _
        "=IF(RC[-1]<>"""",ACOS((SIN(RC2)-SIN(R2C[1])*COS(RC[-1]))/(SIN(RC[-1])*COS(R2C[1]))),"""")"
    Range("I10").FormulaR1C1 = _
        "=IF(RC[-1]<>"""",ABS((ASIN(SIN(RC[-2])*SIN(RC[-1]-R5C))*180*60*1.852/PI())),"""")"
    Range("J10").FormulaR1C1 = "=IF(RC[-1]=R4C[1],RC1,"""")"
    Range("K10:N10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-9],"""")"
    
    Range("K4").FormulaR1C1 = "=MAX(R[6]C[-2]:R[10005]C[-2])"
    Range("L4").FormulaR1C1 = "=MAX(R[6]C[-2]:R[10005]C[-2])"
    
 With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:N10").AutoFill Destination:=.Range("G10:N" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    
    Range("J8:N8").FormulaR1C1 = "=MAX(R[2]C:R[10001]C)"
    Range("P11:T11").Value = Range("J8:N8").Value
    Range("G8:N10009").Clear
    
    Range("G8:L8").FormulaR1C1 = "=MAX(R[2]C:R[10001]C)"
    Range("G10").FormulaR1C1 = _
        "=IF(RC[-6]<R9C16,6371*ACOS(SIN(RC[-5])*SIN(R9C17)+COS(RC[-5])*COS(R9C17)*COS(R9C18-RC[-4])),"""")"
    Range("H10").FormulaR1C1 = "=IF(RC[-1]=R8C7,RC[-7],"""")"
    Range("I10:L10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-7],"""")"
    
 With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:L10").AutoFill Destination:=.Range("G10:L" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
       
    Range("P8:T8").Value = Range("H8:L8").Value
    Range("H8:L10009").Clear
    
    Range("U9:U12").FormulaR1C1 = _
        "=IF(AND(SUM(RC[-5]:RC[-1])>0,SUM(R[-1]C[-5]:R[-1]C[-1])>0),6371*ACOS(SIN(R[-1]C[-4])*SIN(RC[-4])+COS(R[-1]C[-4])*COS(RC[-4])*COS(RC[-3]-R[-1]C[-3])),0)"
    Range("U13").FormulaR1C1 = "=SUM(R[-4]C:R[-1]C)"
    Range("U14").FormulaR1C1 = "=IF(R[-6]C[-2]-R[-2]C[-2]>R9C4,R[-1]C-((R[-6]C[-2]-R[-2]C[-2]-R9C4)*0.1),R[-1]C)"
    ActiveSheet.Calculate
    Range("U9:U14").Value = Range("U9:U14").Value
    
If Range("U14") = Range("U13") Then
    Sheets("Sheet2").Range("H2:J3").Value = Sheets("TASKS").Range("C10:E11").Value
    Sheets("Sheet2").Range("K2:L3").Value = Sheets("TASKS").Range("H10:I11").Value
    
    Range("I4").FormulaR1C1 = "=ACOS(SIN(R2C)*SIN(R3C)+COS(R2C)*COS(R3C)*COS(R3C[1]-R2C[1]))"
    Range("I5").FormulaR1C1 = "=ACOS((SIN(R3C)-SIN(R2C)*COS(R[-1]C))/(SIN(R[-1]C)*COS(R2C)))"
    Range("K4").FormulaR1C1 = "=MAX(R[6]C[-2]:R[10005]C[-2])"
    Range("L4").FormulaR1C1 = "=MIN(R[6]C[-3]:R[10005]C[-3])"
    
    Range("G10").FormulaR1C1 = "=IF(AND(RC[-6]>=R1C1,RC[-6]<=R5C[-6],RC[-4]<>R2C[3]),ACOS(SIN(R2C[2])*SIN(RC2)+COS(R2C[2])*COS(RC2)*COS(RC3-R2C[3])),"""")"
    Range("H10").FormulaR1C1 = "=IF(RC[-1]<>"""",ACOS((SIN(RC2)-SIN(R2C[1])*COS(RC[-1]))/(SIN(RC[-1])*COS(R2C[1]))),"""")"
    Range("I10").FormulaR1C1 = "=IF(RC[-1]<>"""",ASIN(SIN(RC[-2])*SIN(RC[-1]-R5C))*180*60*1.852/PI(),"""")"
    Range("J10").FormulaR1C1 = "=IF(OR(RC[-1]=R4C[1],RC[-1]=R4C[2]),RC1,"""")"
    Range("K10:N10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-9],"""")"
    'Copy Ref A
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:N10").AutoFill Destination:=.Range("G10:N" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("G10:N10009").Value = Range("G10:N10009").Value
    Range("J10:N10009").Sort Key1:=Range("J10"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    
    Range("P15:T15").Value = Range("A1:E1").Value
    Range("P16:T17").Value = Range("J10:N11").Value
    Range("P18:T19").Value = Range("H2:L3").Value
    
    Range("P15:T19").Sort Key1:=Range("P15"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    
    Range("U16:U19").FormulaR1C1 = "=6371*ACOS(SIN(R[-1]C[-4])*SIN(RC[-4])+COS(R[-1]C[-4])*COS(RC[-4])*COS(RC[-3]-R[-1]C[-3]))"
    Range("U20").FormulaR1C1 = "=SUM(R[-4]C:R[-1]C)"
    Range("U21").FormulaR1C1 = "=IF(OR(AND(R[-1]C>100,R[-6]C[-2]-R[-2]C[-2]<=R[-12]C[-17]),AND(R[-1]C<=100,R[-6]C[-2]-R[-2]C[-2]<=10*R[-1]C)),R[-1]C,0)"
    
ElseIf Range("U14") < Range("U13") Then
    Range("A7").Value = 2.08333333333333E-02
    Range("G8:L8").FormulaR1C1 = "=MAX(R[2]C:R[10001]C)"
    
    Range("G10").FormulaR1C1 = "=IF(AND(RC[-6]>=R8C16-R7C1,RC[-6]<=R8C16+R7C1,RC[-3]-R12C19<=R9C4,RC[-4]<>R9C18),6371*ACOS(SIN(RC[-5])*SIN(R9C17)+COS(RC[-5])*COS(R9C17)*COS(R9C18-RC[-4])),"""")"
    Range("H10").FormulaR1C1 = "=IF(RC[-1]=R8C7,RC[-7],"""")"
    Range("I10:L10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-7],"""")"
    
 With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:L10").AutoFill Destination:=.Range("G10:L" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    
    Range("G10:L10009").Value = Range("G10:L10009").Value
    
    Range("E7").FormulaR1C1 = "=R9C21-R8C7"
    Range("E8").FormulaR1C1 = "=IF(R[-1]C>0,R13C21-R14C21,0)"
    ActiveSheet.Calculate
    Range("E7:E8").Value = Range("E7:E8").Value
    
  If Range("E7") < Range("E8") Then

    Range("P15:T15").Value = Range("H8:L8").Value
    Range("P16:T19").Value = Range("P9:T12").Value
    
    Range("U16:U19").FormulaR1C1 = _
        "=6371*ACOS(SIN(R[-1]C[-4])*SIN(RC[-4])+COS(R[-1]C[-4])*COS(RC[-4])*COS(RC[-3]-R[-1]C[-3]))"
    Range("U20").FormulaR1C1 = "=SUM(R[-4]C:R[-1]C)"
    Range("U21").FormulaR1C1 = _
        "=IF(R[-6]C[-2]-R[-2]C[-2]>R9C4,R[-1]C-((R[-6]C[-2]-R[-2]C[-2]-R9C4)*0.1),R[-1]C)"
    ActiveSheet.Calculate
    Range("U16:U21").Value = Range("U16:U21").Value
    End If
  End If

    Range("A7:E7,E8,G8:L10009").Clear

    Range("O1").FormulaR1C1 = "=MAX(R[6]C[6],R[13]C[6],R[20]C[6])"
    ActiveSheet.Calculate
    Range("O1").Value = Range("O1").Value
    
    If Range("O1") = Range("U7") Then
        Range("P1:U7").Value = Range("P1:U7").Value
    ElseIf Range("O1") = Range("U14") Then
        Range("P1:U7").Value = Range("P8:U14").Value
    ElseIf Range("O1") = Range("U21") Then
        Range("P1:U7").Value = Range("P15:U21").Value
    End If
    
    Range("P8:U28").Clear
    Columns("G:L").Clear
End Sub

Sub ORDS3()
'
' Works for Payne2013  /ROWS +7
'
Application.ScreenUpdating = False
    Range("P9:T9").Value = Range("A1:E1").Value
    Sheets("Sheet2").Range("P10:R10").Value = Sheets("TAsks").Range("C15:E15").Value
    Sheets("Sheet2").Range("S10:T10").Value = Sheets("TAsks").Range("H15:I15").Value
  
    Range("G8:L8").FormulaR1C1 = "=MAX(R[2]C:R[10001]C)"
    
    Range("G10").FormulaR1C1 = "=IF(RC[-6]<R1C1,6371*ACOS(SIN(RC[-5])*SIN(R1C2)+COS(RC[-5])*COS(R1C2)*COS(R1C3-RC[-4])),"""")"
    Range("H10").FormulaR1C1 = "=IF(RC[-1]=R8C7,RC[-7],"""")"
    Range("I10:L10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-7],"""")"
   
 With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:L10").AutoFill Destination:=.Range("G10:L" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
   
    Range("P8:T8").Value = Range("H8:L8").Value
    Range("H8:L8").Clear
    
    Range("H2:L2").Value = Range("P10:T10").Value
    Range("H3:L3").Value = Range("A3:E3").Value
    
    'MAXIMUM off course after Selected ORD
    Range("G10").FormulaR1C1 = "=IF(AND(RC1>R2C[1],RC1<R3C[1],RC[-4]<>R2C10),ACOS(SIN(R2C[2])*SIN(RC2)+COS(R2C[2])*COS(RC2)*COS(RC3-R2C[3])),"""")"
    Range("H10").FormulaR1C1 = _
        "=IF(RC[-1]<>"""",ACOS((SIN(RC2)-SIN(R2C[1])*COS(RC[-1]))/(SIN(RC[-1])*COS(R2C[1]))),"""")"
    Range("I10").FormulaR1C1 = _
        "=IF(RC[-1]<>"""",ABS((ASIN(SIN(RC[-2])*SIN(RC[-1]-R5C))*180*60*1.852/PI())),"""")"
    Range("J10").FormulaR1C1 = "=IF(RC[-1]=R4C[1],RC1,"""")"
    Range("K10:N10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-9],"""")"
    
    Range("K4").FormulaR1C1 = "=MAX(R[6]C[-2]:R[10005]C[-2])"
    Range("L4").FormulaR1C1 = "=MAX(R[6]C[-2]:R[10005]C[-2])"
    
 With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:N10").AutoFill Destination:=.Range("G10:N" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    
    Range("J8:N8").FormulaR1C1 = "=MAX(R[2]C:R[10001]C)"
    ActiveSheet.Calculate
    Range("P11:T11").Value = Range("J8:N8").Value
    Range("G8:N10009").Clear
    Range("G8:L8").FormulaR1C1 = "=MAX(R[2]C:R[10001]C)"

    Range("G10").FormulaR1C1 = _
        "=IF(RC1>R11C[9],6371*ACOS(SIN(R11C[10])*SIN(RC2)+COS(R11C[10])*COS(RC2)*COS(RC3-R11C[11])),"""")"
    Range("H10").FormulaR1C1 = "=IF(RC[-1]=R8C[-1],RC[-7],"""")"
    Range("I10:L10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-7],"""")"
    
 With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:L10").AutoFill Destination:=.Range("G10:L" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    
    Range("P12:T12").Value = Range("H8:L8").Value
    Range("U9:U12").FormulaR1C1 = _
        "=IF(OR(SUM(R[-1]C[-5]:R[-1]C[-1])=0,SUM(RC[-5]:RC[-1])=0),0,6371*ACOS(SIN(R[-1]C[-4])*SIN(RC[-4])+COS(R[-1]C[-4])*COS(RC[-4])*COS(R[-1]C[-3]-RC[-3])))"
    Range("U13").FormulaR1C1 = "=SUM(R[-4]C:R[-1]C)"
    Range("U14").FormulaR1C1 = _
        "=IF(R[-6]C[-2]-R[-2]C[-2]>R9C4,R[-1]C-((R[-6]C[-2]-R[-2]C[-2]-R9C4)*0.1),R[-1]C)"
    ActiveSheet.Calculate
    Range("U9:U14").Value = Range("U9:U14").Value

If Range("U14") < Range("U13") Then
    
    Range("G8:L8").FormulaR1C1 = "=MAX(R[2]C:R[10001]C)"
    Range("G10").FormulaR1C1 = _
        "=IF(AND(RC[-6]<=R9C16,RC[-3]-R12C[12]<R9C4),6371*ACOS(SIN(RC[-5])*SIN(R9C17)+COS(RC[-5])*COS(R9C17)*COS(R9C18-RC[-4])),"""")"
    Range("H10").FormulaR1C1 = "=IF(RC[-1]=R8C[-1],RC[-7],"""")"
    Range("I10:L10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-7],"""")"
    
 With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:L10").AutoFill Destination:=.Range("G10:L" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    
    Range("E7").FormulaR1C1 = "=R9C21-R8C7"
    Range("E8").FormulaR1C1 = "=IF(R[-1]C>0,R13C[16]-R14C21,0)"
    ActiveSheet.Calculate
    Range("E7:E8").Value = Range("E7:E8").Value
    
  If Range("E7") < Range("E8") Then
    
    Range("P15:T15").Value = Range("H8:L8").Value
    Range("P16:T19").Value = Range("P9:T12").Value
    Range("U16:U19").FormulaR1C1 = _
        "=6371*ACOS(SIN(R[-1]C[-4])*SIN(RC[-4])+COS(R[-1]C[-4])*COS(RC[-4])*COS(RC[-3]-R[-1]C[-3]))"
    Range("U20").FormulaR1C1 = "=SUM(R[-4]C:R[-1]C)"
    Range("U21").FormulaR1C1 = _
        "=IF(R[-6]C[-2]-R[-2]C[-2]>R9C4,R[-1]C-((R[-6]C[-2]-R[-2]C[-2]-R9C4)*0.1),R[-1]C)"
    ActiveSheet.Calculate
  End If
 End If
    Range("O1").FormulaR1C1 = "=MAX(R[6]C[6],R[13]C[6],R[20]C[6])"
    ActiveSheet.Calculate
    Range("O1").Value = Range("O1").Value
    
    If Range("O1") = Range("U7") Then
        Range("P1:U7").Value = Range("P1:U7").Value
    ElseIf Range("O1") = Range("U14") Then
        Range("P1:U7").Value = Range("P8:U14").Value
    ElseIf Range("O1") = Range("U21") Then
        Range("P1:U7").Value = Range("P15:U21").Value
    End If

    Range("P8:U28").Clear
    Columns("G:L").Clear
   
End Sub
Sub ORDS2()
'
' ORDS2 Macro First ORD > 1 hr after release
'
Application.ScreenUpdating = False
    Range("H2:L2").Value = Range("A1:E1").Value
    Range("G8:L8").FormulaR1C1 = "=MAX(R[2]C:R[10001]C)"
    
    Range("G10").FormulaR1C1 = _
        "=IF(AND(RC[-6]<R2C8,RC[-4]<>R2C10),6371*ACOS(SIN(RC[-5])*SIN(R2C9)+COS(RC[-5])*COS(R2C9)*COS(R2C10-RC[-4])),"""")"
    Range("H10").FormulaR1C1 = "=IF(RC[-1]=R8C[-1],RC[-7],"""")"
    Range("I10:L10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-7],"""")"
   
 With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:L10").AutoFill Destination:=.Range("G10:L" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    
    Range("P9:T9").Value = Range("H8:L8").Value
    
    Range("G10").FormulaR1C1 = _
        "=IF(AND(RC[-6]<R9C16,RC[-4]<>R9C18),6371*ACOS(SIN(RC[-5])*SIN(R9C17)+COS(RC[-5])*COS(R9C17)*COS(R9C18-RC[-4])),"""")"
    
 With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:G10").AutoFill Destination:=.Range("G10:G" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    
    Range("P8:T8").Value = Range("H8:L8").Value
    Range("P10:T10").Value = Range("A1:E1").Value
    Range("G10").FormulaR1C1 = _
        "=IF(AND(RC[-6]>R10C16,RC[-4]<>R10C18),6371*ACOS(SIN(RC[-5])*SIN(R10C17)+COS(RC[-5])*COS(R10C17)*COS(R10C18-RC[-4])),"""")"
    
 With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:G10").AutoFill Destination:=.Range("G10:G" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    
    Range("P11:T11").Value = Range("H8:L8").Value
    
    Range("G10").FormulaR1C1 = _
        "=IF(AND(RC[-6]>R11C16,RC[-4]<>R11C18),6371*ACOS(SIN(RC[-5])*SIN(R11C17)+COS(RC[-5])*COS(R11C17)*COS(R11C18-RC[-4])),"""")"
    
 With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:G10").AutoFill Destination:=.Range("G10:G" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    
    Range("P12:T12").Value = Range("H8:L8").Value
    Range("G8:L10009").Clear
    Range("U9:U12").FormulaR1C1 = _
        "=IF(OR(SUM(R[-1]C[-5]:R[-1]C[-5])=0,SUM(RC[-5]:RC[-1])=0),0,6371*ACOS(SIN(R[-1]C[-4])*SIN(RC[-4])+COS(R[-1]C[-4])*COS(RC[-4])*COS(RC[-3]-R[-1]C[-3])))"
    Range("U13").FormulaR1C1 = "=SUM(R[-4]C:R[-1]C)"
    Range("U14").FormulaR1C1 = _
        "=IF(R[-6]C[-2]-R[-2]C[-2]>R9C4,R[-1]C-((R[-6]C[-2]-R[-2]C[-2]-R9C4)*0.1),R[-1]C)"
    ActiveSheet.Calculate
    Range("U9:U14").Value = Range("U9:U14").Value
    
    Range("O1").FormulaR1C1 = "=MAX(R[6]C[6],R[13]C[6],R[20]C[6])"
    ActiveSheet.Calculate
    Range("O1").Value = Range("O1").Value
        
    If Range("O1") = Range("U7") Then
        Range("P1:U7").Value = Range("P1:U7").Value
    ElseIf Range("O1") = Range("U14") Then
        Range("P1:U7").Value = Range("P8:U14").Value
    ElseIf Range("O1") = Range("U21") Then
        Range("P1:U7").Value = Range("P15:U21").Value
    End If
    
    Range("P8:U28").Clear
    Columns("G:L").Clear
    
End Sub
Sub DH()
'
' DH Macro in ORDS BEFORE XCA
'
    Range("P1:U7").Value = Range("A1:F7").Value
    Range("G8:L8").FormulaR1C1 = "=MAX(R[2]C:R[10001]C)"
    Range("G10").FormulaR1C1 = "=IF(AND(RC[-6]>R4C16,6371*ACOS(SIN(RC[-5])*SIN(R4C17)+COS(RC[-5])*COS(R4C17)*COS(R4C18-RC[-4]))>=R5C21),6371*ACOS(SIN(RC[-5])*SIN(R4C17)+COS(RC[-5])*COS(R4C17)*COS(R4C18-RC[-4])),"""")"
    Range("H10").FormulaR1C1 = "=IF(RC[-1]=R8C7,RC[-7],"""")"
    Range("I10:L10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-7],"""")"
    'Copy Ref A
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:L10").AutoFill Destination:=.Range("G10:L" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("G8:L8").Value = Range("G8:L8").Value
    Range("P5:T5").Value = Range("H8:L8").Value
    
    Range("U5").FormulaR1C1 = "=6371*ACOS(SIN(RC[-4])*SIN(R[-1]C[-4])+COS(RC[-4])*COS(R[-1]C[-4])*COS(R[-1]C[-3]-RC[-3]))"
    Range("U6").FormulaR1C1 = "=SUM(R[-4]C:R[-1]C)"
    Range("U7").FormulaR1C1 = "=IF(R[-6]C[-2]-R[-2]C[-2]<=R9C4,R[-1]C,R[-1]C-((R[-6]C[-2]-R[-2]C[-2]-R9C4)*0.1))"
    ActiveSheet.Calculate
   
    Range("G8:L8").FormulaR1C1 = "=MAX(R[2]C:R[10001]C)"
    'First Try <= D9
    Range("G10").FormulaR1C1 = "=IF(AND(RC[-6]<R2C16,RC[-3]-R5C19<=R9C4),6371*ACOS(SIN(RC[-5])*SIN(R2C17)+COS(RC[-5])*COS(R2C17)*COS(R2C18-RC[-4])),"""")"
    'Copy Ref A
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:L10").AutoFill Destination:=.Range("G10:L" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    If Range("G8") = 0 Then
        Range("G10").FormulaR1C1 = "=IF(AND(RC[-6]<R2C16,RC[-3]-R5C19<R1C19-R5C19),6371*ACOS(SIN(RC[-5])*SIN(R2C17)+COS(RC[-5])*COS(R2C17)*COS(R2C18-RC[-4])),"""")"
        'Copy Ref A
        With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:L10").AutoFill Destination:=.Range("G10:L" & LastRow), Type:=xlFillDefault
    End With
    End If
    ActiveSheet.Calculate
    If Range("G8") <> 0 Then
    Range("G8:L8").Value = Range("G8:L8").Value
    Range("P1:T1").Value = Range("H8:L8").Value
    Range("U2").FormulaR1C1 = "=6371*ACOS(SIN(R[-1]C[-4])*SIN(RC[-4])+COS(R[-1]C[-4])*COS(RC[-4])*COS(RC[-3]-R[-1]C[-3]))"
    Range("U6").FormulaR1C1 = "=SUM(R[-4]C:R[-1]C)"
    Range("U7").FormulaR1C1 = "=IF(R[-6]C[-2]-R[-2]C[-2]<=R9C4,R[-1]C,R[-1]C-((R[-6]C[-2]-R[-2]C[-2]-R[2]C[-2])*0.1))"
    ActiveSheet.Calculate
End If

If Range("U7") > Range("F7") Then
    Range("A1:F7").Value = Range("P1:U7").Value
    Columns("G:U").Clear
End If
End Sub
Sub XCA()
'
' Cks min leg @ A3 PRIOR to ORDS4, move to A1:F7 10/12/14
'
Application.ScreenUpdating = False
    Range("P1:U5").Value = Range("A1:F5").Value

If Range("U3") < Range("U2") And Range("U3") < Range("U4") And Range("U3") < Range("U5") Then
    Range("H2:L3").Value = Range("P4:T5").Value
    
    Range("I4").FormulaR1C1 = "=ACOS(SIN(R2C)*SIN(R3C)+COS(R2C)*COS(R3C)*COS(R3C[1]-R2C[1]))"
    Range("I5").FormulaR1C1 = "=ACOS((SIN(R3C)-SIN(R2C)*COS(R[-1]C))/(SIN(R[-1]C)*COS(R2C)))"
    Range("K4").FormulaR1C1 = "=MAX(R[6]C[-2]:R[10005]C[-2])"
    'Range("L4").FormulaR1C1 = "=MIN(R[6]C[-2]:R[10005]C[-2])"
    ActiveSheet.Calculate
    Range("J8:N8").FormulaR1C1 = "=MAX(R[2]C:R[10001]C)"
    
    Range("G10").FormulaR1C1 = _
        "=IF(AND(RC1>R2C[1],RC1<R3C[1],RC3<>R2C[3]),ACOS(SIN(R2C[2])*SIN(RC2)+COS(R2C[2])*COS(RC2)*COS(RC3-R2C[3])),"""")"
    Range("H10").FormulaR1C1 = _
        "=IF(RC[-1]<>"""",ACOS((SIN(RC2)-SIN(R2C[1])*COS(RC[-1]))/(SIN(RC[-1])*COS(R2C[1]))),"""")"
    Range("I10").FormulaR1C1 = _
        "=IF(RC[-1]<>"""",ABS((ASIN(SIN(RC[-2])*SIN(RC[-1]-R5C))*180*60*1.852/PI())),"""")"
    Range("J10").FormulaR1C1 = "=IF(RC[-1]=R4C[1],RC1,"""")"
    Range("K10:N10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-9],"""")"
    'Copy Ref A
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:N10").AutoFill Destination:=.Range("G10:N" & LastRow), Type:=xlFillDefault
    End With
ActiveSheet.Calculate
    
    Range("P3:T3").Value = Range("J8:N8").Value
    Range("P1:T5").Select
    Range("P1:T5").Sort Key1:=Range("P1"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    
    Range("U2:U5").FormulaR1C1 = "=6371*ACOS(SIN(R[-1]C[-4])*SIN(RC[-4])+COS(R[-1]C[-4])*COS(RC[-4])*COS(RC[-3]-R[-1]C[-3]))"
    Range("U6").FormulaR1C1 = "=SUM(R[-4]C:R[-1]C)"
    Range("U7").FormulaR1C1 = "=IF(R[-6]C[-2]-R[-2]C[-2]>R9C4,R[-1]C-((R[-6]C[-2]-R[-2]C[-2]-R9C4)*0.1),R[-1]C)"
    ActiveSheet.Calculate
    Columns("G:N").Clear
    
    If Range("U7") > Range("F7") Then
        Range("A1:F7").Value = Range("P1:U7").Value
        Range("P1:U7").Clear
    End If
 End If

End Sub

Sub ORDS4()
'
' FOR KEENE - needs LoH check - STR dist start = first ORD which is also O&R TP and 2nd 3-TP TP
'
 Application.ScreenUpdating = False
    Range("H1:U7").Clear
    
    Sheets("Sheet2").Range("H2:J3").Value = Sheets("TASKS").Range("C10:E11").Value
    Sheets("Sheet2").Range("K2:L3").Value = Sheets("TASKS").Range("H10:I11").Value
    
    Range("I4").FormulaR1C1 = _
        "=ACOS(SIN(R2C)*SIN(R3C)+COS(R2C)*COS(R3C)*COS(R3C[1]-R2C[1]))"
    Range("I5").FormulaR1C1 = _
        "=ACOS((SIN(R3C)-SIN(R2C)*COS(R[-1]C))/(SIN(R[-1]C)*COS(R2C)))"
    'K4/K5 = Right & left Off Course
    Range("K4").FormulaR1C1 = "=MAX(R[6]C[-2]:R[10005]C[-2])"
    Range("K5").FormulaR1C1 = "=MIN(R[5]C[-2]:R[10004]C[-2])"
    Range("L4").FormulaR1C1 = "=MIN(R[6]C[-2]:R[10005]C[-2])"
    Range("L5").FormulaR1C1 = "=MAX(R[5]C[-2]:R[10004]C[-2])"
    
    'Varies with the claim!
    Range("G10").FormulaR1C1 = _
        "=IF(AND(RC1<R2C[1],RC3<>R2C[3]),ACOS(SIN(R2C[2])*SIN(RC2)+COS(R2C[2])*COS(RC2)*COS(RC3-R2C[3])),"""")"
    Range("H10").FormulaR1C1 = _
        "=IF(RC[-1]<>"""",ACOS((SIN(RC2)-SIN(R2C[1])*COS(RC[-1]))/(SIN(RC[-1])*COS(R2C[1]))),"""")"
    Range("I10").FormulaR1C1 = _
        "=IF(RC[-1]<>"""",(ASIN(SIN(RC[-2])*SIN(RC[-1]-R5C))*180*60*1.852/PI()),"""")"
    Range("J10").FormulaR1C1 = "=IF(OR(RC[-1]=R4C[1],RC[-1]=R5C[1]),RC1,"""")"
    Range("K10:N10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-9],"""")"
    
 With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:N10").AutoFill Destination:=.Range("G10:N" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    
  If Range("L4") < Range("H2") And Range("L5") < Range("H2") Then
    Range("J10:N10009").Value = Range("J10:N10009").Value
    Range("J10:N10009").Sort Key1:=Range("J10"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal

    Range("P2:T3").Value = Range("J10:N11").Value
    Range("P4:T5").Value = Range("H2:L3").Value
    
    Range("U3:U6").FormulaR1C1 = _
        "=IF(OR(R[-1]C[-1]=0,RC[-1]=0,R[-1]C[-1]="""",RC[-1]=""""),0,6371*ACOS(SIN(R[-1]C[-4])*SIN(RC[-4])+COS(R[-1]C[-4])*COS(RC[-4])*COS(RC[-3]-R[-1]C[-3])))"
    ActiveSheet.Calculate
    'Calculate Finish Point
    Range("P8").FormulaR1C1 = "=MAX(R[2]C:R[10001]C)"
    Range("P6:T6").FormulaR1C1 = "=MAX(R[4]C[1]:R[10003]C[1])"
    
    Range("P10").FormulaR1C1 = _
        "=IF(RC[-15]>R5C,6371*ACOS(SIN(RC[-14])*SIN(R5C[1])+COS(RC[-14])*COS(R5C[1])*COS(R5C[2]-RC[-13])),"""")"
    Range("Q10").FormulaR1C1 = "=IF(RC[-1]=R8C[-1],RC[-16],"""")"
    Range("R10:U10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-16],"""")"
    
 With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("P10:U10").AutoFill Destination:=.Range("P10:U" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    
    Range("P6:T6").Value = Range("P6:T6").Value
    
    'Recalculate Start
    Range("P8:U8").FormulaR1C1 = "=MAX(R[2]C:R[10001]C)"
    Range("P10").FormulaR1C1 = "=IF(RC[-15]<R2C,6371*ACOS(SIN(RC[-14])*SIN(R2C17)+COS(RC[-14])*COS(R2C17)*COS(R2C18-RC[-13])),"""")"
    
 With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("P10:P10").AutoFill Destination:=.Range("P10:P" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    
    Range("P8:U8").Value = Range("P8:U8").Value
    
    Range("O8").FormulaR1C1 = "=SUM(RC[1]:RC[5])"
    ActiveSheet.Calculate
    Range("O8").Value = Range("O8").Value
    
  If Range("O8") <> 0 Then
    
    Range("V8").FormulaR1C1 = _
        "=6371*ACOS(SIN(RC[-4])*SIN(R[-6]C[-5])+COS(RC[-4])*COS(R[-6]C[-5])*COS(R[-6]C[-4]-RC[-3]))"
    ActiveSheet.Calculate
    Range("V8").Value = Range("V8").Value
    Range("W8").FormulaR1C1 = "=SUM(RC[-1],R[-5]C[-2]:R[-3]C[-2])"
    ActiveSheet.Calculate
    Range("W8").Value = Range("W8").Value
    Range("V6").FormulaR1C1 = "=SUM(R[-3]C[-1]:RC[-1])"
    ActiveSheet.Calculate
    Range("V6").Value = Range("V6").Value
    
    If Range("W8") <= Range("V6") Then
        Range("P1:T1").Value = Range("P2:T2").Value
        Range("P2:T2").Value = Range("P3:T3").Value
        Range("P3:T3").Value = Range("P4:T4").Value
        Range("P4:T4").Value = Range("P5:T5").Value
        Range("P5:T5").Value = Range("P6:T6").Value
        Range("P6:U6").Clear
        Range("U2").FormulaR1C1 = _
        "=6371*ACOS(SIN(R[-1]C[-4])*SIN(RC[-4])+COS(R[-1]C[-4])*COS(RC[-4])*COS(RC[-3]-R[-1]C[-3]))"
    
    ElseIf Range("W8") > Range("V6") Then
        Range("P1:T1").Value = Range("Q8:U8").Value
        Range("P6:V6").Clear
        Range("U2").FormulaR1C1 = _
        "=6371*ACOS(SIN(R[-1]C[-4])*SIN(RC[-4])+COS(R[-1]C[-4])*COS(RC[-4])*COS(RC[-3]-R[-1]C[-3]))"
        Range("P10").FormulaR1C1 = _
        "=IF(RC[-15]>R4C,6371*ACOS(SIN(RC[-14])*SIN(R4C17)+COS(RC[-14])*COS(R4C17)*COS(R4C18-RC[-13])),"""")"

    Application.Calculation = xlCalculationManual
 With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("P10:P10").AutoFill Destination:=.Range("P10:P" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    
        Range("P8:U8").FormulaR1C1 = "=MAX(R[2]C:R[100008]C)"
    End If
        
        If Range("P8") > Range("U5") Then
        Range("P5:T5").Value = Range("Q8:U8").Value
        End If
    End If
    
    Range("U6").FormulaR1C1 = "=SUM(R[-4]C:R[-1]C)"
    Range("U7").FormulaR1C1 = _
        "=IF(R[-6]C[-2]-R[-2]C[-2]>R9C4,R[-1]C-((R[-6]C[-2]-R[-2]C[-2]-R9C4)*0.1),R[-1]C)"
    ActiveSheet.Calculate
    Range("P8:W10009").Clear
    Columns("G:N").Clear
    
    Application.Run "F.xlsm!ORDS4a"
    
    Range("O1").FormulaR1C1 = "=MAX(R7C[6],R14C[6])"
    ActiveSheet.Calculate
    Range("O1").Value = Range("O1").Value
    
    If Range("O1") = Range("U7") And Range("U7") > Range("F7") Then
        Range("A1:F7").Value = Range("P1:U7").Value
    ElseIf Range("O1") = Range("U14") And Range("U14") > Range("F7") Then
        Range("A1:F7").Value = Range("P8:U14").Value
    End If
  End If
    Columns("G:U").Clear
   
End Sub
Sub ORDS4a()
'
' Works for Paul, data at P8:T14
'
Application.ScreenUpdating = False
    Sheets("Sheet2").Range("M1").Value = Sheets("TASKS").Range("C10").Value
    Sheets("Sheet2").Range("M2").Value = Sheets("TASKS").Range("C15").Value
    
    If Range("M1") = Range("M2") Then
        Sheets("Sheet2").Range("H2:J4").Value = Sheets("TASKS").Range("C14:E16").Value
        Sheets("Sheet2").Range("K2:L4").Value = Sheets("TASKS").Range("H14:I16").Value
    Range("M1:M2").Clear
    
    Range("G5").FormulaR1C1 = "=(R[-1]C[1]-R[-2]C[1])/3"
    Range("G7").FormulaR1C1 = "=R[-4]C[1]+R[-2]C"
    Range("H7").FormulaR1C1 = "=RC[-1]+R[-2]C[-1]"
    Range("I7").FormulaR1C1 = "=RC[-1]+R[-2]C[-2]"
    
    Range("G8:L8").FormulaR1C1 = "=MAX(R[2]C:R[10001]C)"
   
    Range("G10").FormulaR1C1 = _
        "=IF(AND(RC[-6]>R3C[1],RC[-6]<=R7C),6371*ACOS(SIN(RC[-5])*SIN(R3C[2])+COS(RC[-5])*COS(R3C[2])*COS(R3C[3]-RC[-4])),"""")"
    Range("H10").FormulaR1C1 = "=IF(RC[-1]=R8C[-1],RC[-7],"""")"
    Range("I10:L10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-7],"""")"

 With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:L10").AutoFill Destination:=.Range("G10:L" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    
    Range("P8:T9").Value = Range("H2:L3").Value
    Range("P10:T10").Value = Range("H8:L8").Value
    
    Range("G10").FormulaR1C1 = _
        "=IF(AND(RC[-6]>R7C,RC[-6]<=R7C[1]),6371*ACOS(SIN(RC[-5])*SIN(R10C[10])+COS(RC[-5])*COS(R10C[10])*COS(R10C[11]-RC[-4])),"""")"
    
 With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:G10").AutoFill Destination:=.Range("G10:G" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    
    Range("P11:T11").Value = Range("H8:L8").Value
    
    Range("G10").FormulaR1C1 = _
        "=IF(RC[-6]>R7C[1],6371*ACOS(SIN(RC[-5])*SIN(R11C[10])+COS(RC[-5])*COS(R11C[10])*COS(R11C[11]-RC[-4])),"""")"
    
 With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:G10").AutoFill Destination:=.Range("G10:G" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    
    Range("P12:T12").Value = Range("H8:L8").Value
  
    Range("U9:U12").FormulaR1C1 = _
        "=6371*ACOS(SIN(R[-1]C[-4])*SIN(RC[-4])+COS(R[-1]C[-4])*COS(RC[-4])*COS(RC[-3]-R[-1]C[-3]))"
   
    Range("U13").FormulaR1C1 = "=SUM(R[-4]C:R[-1]C)"
    Range("U14").FormulaR1C1 = _
        "=IF(R[-6]C[-2]-R[-2]C[-2]>R9C4,R[-1]C-((R[-6]C[-2]-R[-2]C[-2]-R9C4)*0.1),R[-1]C)"
    ActiveSheet.Calculate
    Range("U9:U14").Value = Range("U9:U14").Value
  End If

End Sub
Sub ORDS5()
'
' LoHmatrix Macro Amended 9/29/2015 for <10K fixes @ 1 second interval
'
Application.ScreenUpdating = False
    Range("A7").Value = 1 / 48
    Range("B8").FormulaR1C1 = "=R[3]C[-1]-R[2]C[-1]"
    Range("C8").FormulaR1C1 = "=(1/24)/RC[-1]"
    ActiveSheet.Calculate
    If Range("C8") > 900 Then
        Range("A7").Value = 0.005208333
    Else: Range("A7").Value = 1 / 96
    End If
    ActiveSheet.Calculate
    Range("B8:C8").Clear
  
    Range("G10").FormulaR1C1 = _
        "=IF(AND(RC[-6]>R4C1,RC[-6]>=R5C1-R7C[-6],RC[-6]<=R5C1+R7C[-6]),RC[-6],"""")"
    Range("H10:K10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-6],"""")"
    
 With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:K10").AutoFill Destination:=.Range("G10:K" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    
    Range("G10:K60009").Value = Range("G10:K60009").Value
    Range("G10:K60009").Sort Key1:=Range("G10"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    
    Range("G10:K509").Copy
    Range("N1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=True
    Columns("G:K").Clear
    
    Range("G10").FormulaR1C1 = "=IF(RC[-6]<R2C1,RC[-6],"""")"
    Range("H10:K10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-6],"""")"
    Range("L10").FormulaR1C1 = "=IF(RC[-1]<>"""",6371*ACOS(SIN(RC8)*SIN(R2C2)+COS(RC8)*COS(R2C2)*COS(R2C3-RC9)),"""")"
    
 With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:L10").AutoFill Destination:=.Range("G10:L" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    
    Range("G10:L10009").Value = Range("G10:L10009").Value
    
    Range("N6:SS6").FormulaR1C1 = _
        "=IF(OR(R1C="""",R3C=R[-2]C3),"""",6371*ACOS(SIN(R2C)*SIN(R[-2]C2)+COS(R2C)*COS(R[-2]C2)*COS(R[-2]C3-R3C)))"
    ActiveSheet.Calculate
    Range("N6:SS6").Value = Range("N6:SS6").Value
        
    'Matrix
    Range("N10:SS10").FormulaR1C1 = _
        "=IF(R1C="""","""",IF(RC10-R4C>R9C4,(RC12+R6C)-((RC10-R4C-R9C4)*0.1),RC12+R6C))"
    
 With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
.Range("N10:SS10").AutoFill Destination:=.Range("N10:SS" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate

    Range("I5").FormulaR1C1 = "=MAX(R[5]C[5]:R[10004]C[504])"
    ActiveSheet.Calculate
    Range("I5").Value = Range("I5").Value
    Range("N10:SS10009").Value = Range("N10:SS10009").Value
    Range("N7:SS7").FormulaR1C1 = "=IF(MAX(R[3]C:R[10002]C)=R5C9,1,"""")"
    ActiveSheet.Calculate
    Range("N7:SS7").Value = Range("N7:SS7").Value
    
    Range("M10").FormulaR1C1 = "=IF(MAX(RC[1]:RC[500])=R5C9,1,"""")"
    
 With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
.Range("M10:M10").AutoFill Destination:=.Range("M10:M" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("M10:M10009").Value = Range("M10:M10009").Value
    
    Range("N10:SS10009").Clear
    
    Range("N10").FormulaR1C1 = "=IF(RC[-1]=1,RC[-13],"""")"
    Range("O10:R10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-13],"""")"

 With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
.Range("N10:R10").AutoFill Destination:=.Range("N10:R" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    
    Range("H1:L1").FormulaR1C1 = "=MAX(R[9]C[6]:R[10008]C[6])"
    ActiveSheet.Calculate
    Range("A1:E1").Value = Range("H1:L1").Value
    Range("N10:R10009").Clear
    
    Range("N10:SS10").FormulaR1C1 = "=IF(R[-3]C=1,R[-9]C,"""")"
    Range("N11:SS14").FormulaR1C1 = "=IF(R[-1]C<>"""",R[-9]C,"""")"
    
    Range("H5").FormulaR1C1 = "=MAX(R[5]C[6]:R[5]C[505])"
    Range("I5").FormulaR1C1 = "=MAX(R[6]C[5]:R[6]C[504])"
    Range("J5").FormulaR1C1 = "=MAX(R[7]C[4]:R[7]C[503])"
    Range("K5").FormulaR1C1 = "=MAX(R[8]C[3]:R[8]C[502])"
    Range("L5").FormulaR1C1 = "=MAX(R[9]C[2]:R[9]C[501])"
    ActiveSheet.Calculate
    Range("A5:E5").Value = Range("H5:L5").Value
    
    Columns("G:SS").Clear
    Range("F2:F5").FormulaR1C1 = _
        "=6371*ACOS(SIN(RC[-4])*SIN(R[-1]C[-4])+COS(RC[-4])*COS(R[-1]C[-4])*COS(R[-1]C[-3]-RC[-3]))"
    Range("F6").FormulaR1C1 = "=SUM(R[-4]C:R[-1]C)"
    Range("F7").FormulaR1C1 = _
        "=IF(R[-6]C[-2]-R[-2]C[-2]>R9C4,R[-1]C-((R[-6]C[-2]-R[-2]C[-2]-R9C4)*0.1),R[-1]C)"
    ActiveSheet.Calculate
    Range("F2:F7").Value = Range("F2:F7").Value
End Sub
Sub Ordsa()
'
' For semi, Mueller & Arcturas Amended 5/11/2018 for Howard!
'
Application.ScreenUpdating = False
    Sheets("Sheet2").Range("P2:R3").Value = Sheets("Tasks").Range("C10:E11").Value
    Sheets("Sheet2").Range("S2:T3").Value = Sheets("Tasks").Range("H10:I11").Value
    
    Range("G8:L8").FormulaR1C1 = "=MAX(R[2]C:R[10001]C)"
    
    Range("G10").FormulaR1C1 = _
        "=IF(RC[-6]>R3C[9],6371*ACOS(SIN(RC[-5])*SIN(R3C[10])+COS(RC[-5])*COS(R3C[10])*COS(R3C[11]-RC[-4])),"""")"
    Range("H10").FormulaR1C1 = "=IF(RC[-1]=R8C7,RC[-7],"""")"
    Range("I10:L10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-7],"""")"
   
 With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:L10").AutoFill Destination:=.Range("G10:L" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
   
    Range("P4:T4").Value = Range("H8:L8").Value
    Range("G10").FormulaR1C1 = _
        "=IF(RC[-6]>R4C[9],6371*ACOS(SIN(RC[-5])*SIN(R4C[10])+COS(RC[-5])*COS(R4C[10])*COS(R4C[11]-RC[-4])),"""")"
    
 With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:G10").AutoFill Destination:=.Range("G10:G" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    
  If Range("G8") <> 0 Then
    
    Range("P5:T5").Value = Range("H8:L8").Value
    Range("G10").FormulaR1C1 = _
        "=IF(RC[-6]<R2C[9],6371*ACOS(SIN(RC[-5])*SIN(R2C[10])+COS(RC[-5])*COS(R2C[10])*COS(R2C[11]-RC[-4])),"""")"
    
 With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:G10").AutoFill Destination:=.Range("G10:G" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    
   If Range("G8") <> 0 Then
    
    Range("P1:T1").Value = Range("H8:L8").Value
    
    Range("U2:U5").FormulaR1C1 = _
        "=IF(AND(SUM(R[-1]C[-5],R[-1]C[-1])<>0,SUM(RC[-5],RC[-1])<>0),6371*ACOS(SIN(R[-1]C[-4])*SIN(RC[-4])+COS(R[-1]C[-4])*COS(RC[-4])*COS(RC[-3]-R[-1]C[-3])),0)"
    Range("U6").FormulaR1C1 = "=SUM(R[-4]C:R[-1]C)"
    Range("U7").FormulaR1C1 = _
        "=IF(R[-6]C[-2]-R[-2]C[-2]>R9C4,R[-1]C-((R[-6]C[-2]-R[-2]C[-2]-R9C4)*0.1),R[-1]C)"
    ActiveSheet.Calculate
    Range("U2:U7").Value = Range("U2:U7").Value
    
    Range("B8").FormulaR1C1 = "=R[-7]C[14]-R[2]C[-1]"
    ActiveSheet.Calculate
    If Range("B8") >= 1 / 48 Then
    
    Range("G10").FormulaR1C1 = _
        "=IF(RC[-6]<R1C[9],6371*ACOS(SIN(RC[-5])*SIN(R1C[10])+COS(RC[-5])*COS(R1C[10])*COS(R1C[11]-RC[-4])),"""")"

 With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:G10").AutoFill Destination:=.Range("G10:G" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    
    Range("P8:T8").Value = Range("H8:L8").Value
    Range("P9:T12").Value = Range("P1:T4").Value
    Range("U9:U12").FormulaR1C1 = _
        "=IF(AND(SUM(R[-1]C[-5],R[-1]C[-1])<>0,SUM(RC[-5],RC[-1])<>0),6371*ACOS(SIN(R[-1]C[-4])*SIN(RC[-4])+COS(R[-1]C[-4])*COS(RC[-4])*COS(RC[-3]-R[-1]C[-3])),0)"
    Range("U13").FormulaR1C1 = "=SUM(R[-4]C:R[-1]C)"
    Range("U14").FormulaR1C1 = "=IF(R[-6]C[-2]-R[-2]C[-2]>R9C4,((R[-6]C[-2]-R[-2]C[-2]-R9C4)*.1),R[-1]C)"
    ActiveSheet.Calculate
    If Range("U14") > Range("U7") Then
        Range("P1:U7").Value = Range("P8:U14").Value
        Range("P8:U14").Clear
    End If
   End If
   
   ElseIf Range("G8") = 0 Then
   
    Range("P2:T5").Cut Destination:=Range("P1:T4")
    Range("U2:U4").FormulaR1C1 = _
        "=IF(AND(SUM(R[-1]C[-5],R[-1]C[-1])<>0,SUM(RC[-5],RC[-1])<>0),6371*ACOS(SIN(R[-1]C[-4])*SIN(RC[-4])+COS(R[-1]C[-4])*COS(RC[-4])*COS(RC[-3]-R[-1]C[-3])),0)"
    Range("G10:G10009").FormulaR1C1 = "=IF(RC[-6]>R4C[9],6371*ACOS(SIN(RC[-5])*SIN(R4C[10])+COS(RC[-5])*COS(R4C[10])*COS(R4C[11]-RC[-4])),"""")"
    ActiveSheet.Calculate
    Range("P5:T5").Value = Range("H8:L8").Value
    Range("U5").FormulaR1C1 = _
        "=IF(AND(SUM(R[-1]C[-5],R[-1]C[-1])<>0,SUM(RC[-5],RC[-1])<>0),6371*ACOS(SIN(R[-1]C[-4])*SIN(RC[-4])+COS(R[-1]C[-4])*COS(RC[-4])*COS(RC[-3]-R[-1]C[-3])),0)"
    Range("U6").FormulaR1C1 = "=SUM(R[-4]C:R[-1]C)"
    Range("U7").FormulaR1C1 = _
        "=IF(R[-6]C[-2]-R[-2]C[-2]>R9C4,R[-1]C-((R[-6]C[-2]-R[-2]C[-2]-R9C4)*0.1),R[-1]C)"
    ActiveSheet.Calculate
    Range("U2:U7").Value = Range("U2:U7").Value
   End If
   
    'Ck TP1 for Howard!
    Range("I7").FormulaR1C1 = "=SUM(R2C21:R3C21)"
    Range("I8:N8").FormulaR1C1 = "=MAX(R[2]C:R[10001]C)"
    Range("G10").FormulaR1C1 = "=IF(AND(RC[-6]>R1C16,RC[-6]<R3C16),6371*ACOS(SIN(RC[-5])*SIN(R1C17)+COS(RC[-5])*COS(R1C17)*COS(R1C18-RC[-4])),"""")"
    Range("H10").FormulaR1C1 = "=IF(RC[-1]<>"""",6371*ACOS(SIN(RC[-6])*SIN(R3C17)+COS(RC[-6])*COS(R3C17)*COS(R3C18-RC[-5])),"""")"
    Range("I10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-2]+RC[-1],"""")"
    Range("J10").FormulaR1C1 = "=IF(RC[-1]=R8C9,RC[-9],"""")"
    Range("K10:N10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-9],"""")"
    'Copy Ref A
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:N10").AutoFill Destination:=.Range("G10:N" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    
    If Range("I8") > Range("I7") Then
    
        Range("P2:T2").Value = Range("J8:N8").Value
        Range("U2:U3").FormulaR1C1 = "=6316*ACOS(SIN(R[-1]C[-4])*SIN(RC[-4])+COS(R[-1]C[-4])*COS(RC[-4])*COS(RC[-3]-R[-1]C[-3]))"
        Range("U6").FormulaR1C1 = "=SUM(R[-4]C:R[-1]C)"
        Range("U7").FormulaR1C1 = "=IF(R[-6]C[-2]-R[-2]C[-2]>R9C4,((R[-6]C[-2]-R[-2]C[-2]-R9C4)*.1),R[-1]C)"
        Range("G7:N10009").Clear
    End If

 ElseIf Range("G8") = 0 Then
 
    Range("G10").FormulaR1C1 = _
        "=IF(RC[-6]<R2C[9],6371*ACOS(SIN(RC[-5])*SIN(R2C[10])+COS(RC[-5])*COS(R2C[10])*COS(R2C[11]-RC[-4])),"""")"

 With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:G10").AutoFill Destination:=.Range("G10:G" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    
    Range("P1:T1").Value = Range("H8:L8").Value
    
    Sheets("Sheet2").Range("H2:J3").Value = Sheets("TASKS").Range("C10:E11").Value
    Sheets("Sheet2").Range("K2:L3").Value = Sheets("TASKS").Range("H10:I11").Value
    
    Range("I4").FormulaR1C1 = _
        "=ACOS(SIN(R2C)*SIN(R3C)+COS(R2C)*COS(R3C)*COS(R3C[1]-R2C[1]))"
    Range("I5").FormulaR1C1 = _
        "=ACOS((SIN(R3C)-SIN(R2C)*COS(R[-1]C))/(SIN(R[-1]C)*COS(R2C)))"
    'K4/K5 = Right & left Off Course
    Range("K4").FormulaR1C1 = "=MAX(R[6]C[-2]:R[10005]C[-2])"
    Range("K5").FormulaR1C1 = "=MIN(R[5]C[-2]:R[10004]C[-2])"
    Range("L4").FormulaR1C1 = "=MIN(R[6]C[-2]:R[10005]C[-2])"
    Range("L5").FormulaR1C1 = "=MAX(R[5]C[-2]:R[10004]C[-2])"
    
    'Off ST Dist Course between St Dist Start & FIN
    Range("G10").FormulaR1C1 = _
        "=IF(AND(RC1>R2C[1],RC1<R3C[1],RC3<>R2C[3]),ACOS(SIN(R2C[2])*SIN(RC2)+COS(R2C[2])*COS(RC2)*COS(RC3-R2C[3])),"""")"
    Range("H10").FormulaR1C1 = _
        "=IF(RC[-1]<>"""",ACOS((SIN(RC2)-SIN(R2C[1])*COS(RC[-1]))/(SIN(RC[-1])*COS(R2C[1]))),"""")"
    Range("I10").FormulaR1C1 = _
        "=IF(RC[-1]<>"""",(ASIN(SIN(RC[-2])*SIN(RC[-1]-R5C))*180*60*1.852/PI()),"""")"
    Range("J10").FormulaR1C1 = "=IF(OR(RC[-1]=R4C[1],RC[-1]=R5C[1]),RC1,"""")"
    Range("K10:N10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-9],"""")"
    
 With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:N10").AutoFill Destination:=.Range("G10:N" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    
  'If Range("L4") < Range("H2") And Range("L5") < Range("H2") Then
    Range("J10:N10009").Value = Range("J10:N10009").Value
    Range("J10:N10009").Sort Key1:=Range("J10"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal

    Range("P5:T6").Value = Range("J10:N11").Value
    
    Range("P1:T6").Sort Key1:=Range("P1"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    
    Range("U2:U6").FormulaR1C1 = _
        "=IF(AND(SUM(R[-1]C[-5],R[-1]C[-1])<>0,SUM(RC[-5],RC[-1])<>0),6371*ACOS(SIN(R[-1]C[-4])*SIN(RC[-4])+COS(R[-1]C[-4])*COS(RC[-4])*COS(RC[-3]-R[-1]C[-3])),0)"
    Range("T16").FormulaR1C1 = "=LARGE(R[-15]C[1]:R[-1]C[1],4)"
    Range("V1").FormulaR1C1 = "=IF(R[1]C[-1]>=R16C[-2],RC[-6],"""")"
    Range("V2:V6").FormulaR1C1 = "=IF(RC[-1]>=R16C[-2],RC[-6],"""")"
    ActiveSheet.Calculate
    Range("V1:V6").Value = Range("V1:V6").Value
    
    Range("P1:V6").Sort Key1:=Range("V1"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    
    Range("P6:V16").Clear
    Range("V1:V6").Clear
    
    Range("U6").FormulaR1C1 = "=SUM(R[-4]C:R[-1]C)"
    Range("U7").FormulaR1C1 = _
        "=IF(R[-6]C[-2]-R[-2]C[-2]>R9C4,R[-1]C-((R[-6]C[-2]-R[-2]C[-2]-R9C4)*0.1),R[-1]C)"
    ActiveSheet.Calculate
    Range("U2:U7").Value = Range("U2:U7").Value
  End If
    
    Range("G8:L8").FormulaR1C1 = "=MAX(R[2]C:R[10001]C)"
   
    Range("G10").FormulaR1C1 = _
        "=IF(RC[-6]<R1C[9],6371*ACOS(SIN(RC[-5])*SIN(R1C[10])+COS(RC[-5])*COS(R1C[10])*COS(R1C[11]-RC[-4])),"""")"
    Range("H10").FormulaR1C1 = "=IF(RC[-1]=R8C[-1],RC[-7],"""")"
    Range("I10:L10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-7],"""")"
   
 With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:L10").AutoFill Destination:=.Range("G10:L" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
   
    Range("O1").FormulaR1C1 = "=MIN(R1C21:R5C21)"
    ActiveSheet.Calculate
    
  If Range("G8") > Range("O1") Then
    Range("P8:T8").Value = Range("H8:L8").Value
    Range("P9:T13").Value = Range("P1:T5").Value
    Range("U9:U14").FormulaR1C1 = _
        "=IF(AND(SUM(R[-1]C[-5],R[-1]C[-1])<>0,SUM(RC[-5],RC[-1])<>0),6371*ACOS(SIN(R[-1]C[-4])*SIN(RC[-4])+COS(R[-1]C[-4])*COS(RC[-4])*COS(RC[-3]-R[-1]C[-3])),0)"
    Range("T16").FormulaR1C1 = "=LARGE(R[-7]C[1]:R[-3]C[1],4)"
    Range("V8").FormulaR1C1 = "=IF(R[1]C[-1]>=R16C[-2],RC[-6],"""")"
    Range("V9:V13").FormulaR1C1 = "=IF(RC[-1]>=R16C[-2],RC[-6],"""")"
    ActiveSheet.Calculate
    Range("V8:V13").Value = Range("V8:V13").Value
    
    Range("P8:V14").Sort Key1:=Range("V8"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal

    Range("P13:V19,V8:V13").Clear
    
    Range("U9:U12").FormulaR1C1 = _
        "=IF(AND(SUM(R[-1]C[-5],R[-1]C[-1])<>0,SUM(RC[-5],RC[-1])<>0),6371*ACOS(SIN(R[-1]C[-4])*SIN(RC[-4])+COS(R[-1]C[-4])*COS(RC[-4])*COS(RC[-3]-R[-1]C[-3])),0)"
    Range("U13").FormulaR1C1 = "=SUM(R[-4]C:R[-1]C)"
    Range("U14").FormulaR1C1 = _
        "=IF(R[-6]C[-2]-R[-2]C[-2]>R9C4,R[-1]C-((R[-6]C[-2]-R[-2]C[-2]-R9C4)*0.1),R[-1]C)"
    ActiveSheet.Calculate
  End If
   
    Range("O1").FormulaR1C1 = "=MAX(R[6]C[6],R[13]C[6])"
    ActiveSheet.Calculate
    If Range("O1") = Range("U7") And Range("U7") > Range("F7") Then
        Range("A1:F7").Value = Range("P1:U7").Value
    ElseIf Range("O1") = Range("U14") And Range("U14") > Range("F7") Then
        Range("A1:F7").Value = Range("P8:U14").Value
    End If

    Columns("G:U").Clear
    Application.Run "F.xlsm!NoReturn2"

End Sub
Sub NoReturn2()
'
' For Franke Straight OCDist each of 2 halves of the flight
'
Application.ScreenUpdating = False
    Sheets("Sheet2").Range("H2:J3").Value = Sheets("TASKS").Range("C10:E11").Value
    Sheets("Sheet2").Range("K2:L3").Value = Sheets("TASKS").Range("H10:I11").Value
    
    Range("A7").FormulaR1C1 = "=(R3C8-R2C8)/2"
    
    Range("I4").FormulaR1C1 = _
        "=ACOS(SIN(R2C)*SIN(R3C)+COS(R2C)*COS(R3C)*COS(R3C[1]-R2C[1]))"
    Range("K4").FormulaR1C1 = "=MIN(R[6]C[-2]:R[10005]C[-2])"
    Range("K5").FormulaR1C1 = "=MAX(R[6]C[-2]:R[10005]C[-2])"
    Range("I5").FormulaR1C1 = _
        "=ACOS((SIN(R3C)-SIN(R2C)*COS(R[-1]C))/(SIN(R[-1]C)*COS(R2C)))"
    
    Range("G10").FormulaR1C1 = _
        "=IF(AND(RC1>R2C[1],RC1<R3C[1],RC[-4]<>R2C[3]),ACOS(SIN(R2C[2])*SIN(RC2)+COS(R2C[2])*COS(RC2)*COS(RC3-R2C[3])),"""")"
    Range("H10").FormulaR1C1 = _
        "=IF(RC[-1]<>"""",ACOS((SIN(RC2)-SIN(R2C[1])*COS(RC[-1]))/(SIN(RC[-1])*COS(R2C[1]))),"""")"
    Range("I10").FormulaR1C1 = _
        "=IF(RC[-1]<>"""",(ASIN(SIN(RC[-2])*SIN(RC[-1]-R5C))*180*60*1.852/PI()),"""")"
    Range("J10").FormulaR1C1 = "=IF(OR(RC[-1]=R4C[1],RC[-1]=R5C[1]),RC1,"""")"
    Range("K10:N10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-9],"""")"
    
 With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:N10").AutoFill Destination:=.Range("G10:N" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    
    Range("P1:T1").Value = Range("H2:L2").Value
    
    Range("J10:N10009").Value = Range("J10:N10009").Value
    Range("J10:N10009").Sort Key1:=Range("J10"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    Range("P2:T3").Value = Range("J10:N11").Value
    
    Range("U2").FormulaR1C1 = _
        "=6371*ACOS(SIN(R[-1]C[-4])*SIN(RC[-4])+COS(R[-1]C[-4])*COS(RC[-4])*COS(RC[-3]-R[-1]C[-3]))"
    Range("U3").FormulaR1C1 = _
        "=6371*ACOS(SIN(R[-2]C[-4])*SIN(RC[-4])+COS(R[-2]C[-4])*COS(RC[-4])*COS(RC[-3]-R[-2]C[-3]))"
    ActiveSheet.Calculate
    
    If Range("U3") > Range("U2") Then
        Range("P2:U2").Delete Shift:=xlUp
    ElseIf Range("U2") > Range("U3") Then
        Range("P3:T3").Clear
    End If
    
    Range("G10").FormulaR1C1 = _
        "=IF(AND(RC1>R2C16,RC1<R2C16+R7C1,RC[-4]<>R2C[3]),ACOS(SIN(R2C[2])*SIN(RC2)+COS(R2C[2])*COS(RC2)*COS(RC3-R2C[3])),"""")"
    Range("J10").FormulaR1C1 = "=IF(OR(RC[-1]=R4C[1],RC[-1]=R5C[1]),RC1,"""")"
    Range("K10:N10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-9],"""")"
    
 With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:N10").AutoFill Destination:=.Range("G10:N" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    
    Range("J10:N10009").Value = Range("J10:N10009").Value
    Range("J10:N10009").Sort Key1:=Range("J10"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    Range("P3:T4").Value = Range("J10:N11").Value
    
    Range("U3").FormulaR1C1 = _
        "=6371*ACOS(SIN(R[-1]C[-4])*SIN(RC[-4])+COS(R[-1]C[-4])*COS(RC[-4])*COS(RC[-3]-R[-1]C[-3]))"
    Range("U4").FormulaR1C1 = _
        "=6371*ACOS(SIN(R[-2]C[-4])*SIN(RC[-4])+COS(R[-2]C[-4])*COS(RC[-4])*COS(RC[-3]-R[-2]C[-3]))"
    ActiveSheet.Calculate
    
    If Range("U4") > Range("U3") Then
        Range("P3:U3").Delete Shift:=xlUp
    ElseIf Range("U3") > Range("U4") Then
        Range("P4:T4").Clear
    End If
    
    Range("P4:T4").Value = Range("H3:L3").Value
    Range("H3:L3").Value = Range("P2:T2").Value
    
    Range("G10").FormulaR1C1 = _
        "=IF(AND(RC1>R1C16,RC1<R2C16,RC[-4]<>R2C[3]),ACOS(SIN(R2C[2])*SIN(RC2)+COS(R2C[2])*COS(RC2)*COS(RC3-R2C[3])),"""")"
    Range("J10").FormulaR1C1 = "=IF(OR(RC[-1]=R4C[1],RC[-1]=R5C[1]),RC[-9],"""")"
    Range("K10:N10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-9],"""")"
    
 With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:N10").AutoFill Destination:=.Range("G10:N" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    
    Range("J10:N10009").Value = Range("J10:N10009").Value
    Range("J10:N10009").Sort Key1:=Range("J10"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    
    Range("P5:T6").Value = Range("J10:N11").Value
    
    Range("P1:T6").Sort Key1:=Range("P1"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
     
    Range("U2:U5").FormulaR1C1 = _
        "=6371*ACOS(SIN(R[-1]C[-4])*SIN(RC[-4])+COS(R[-1]C[-4])*COS(RC[-4])*COS(RC[-3]-R[-1]C[-3]))"
    Range("U6").FormulaR1C1 = _
        "=6371*ACOS(SIN(R[-2]C[-4])*SIN(RC[-4])+COS(R[-2]C[-4])*COS(RC[-4])*COS(RC[-3]-R[-2]C[-3]))"
    ActiveSheet.Calculate
    
    If Range("U6") > Range("U5") Then
        Range("P5:U5").Delete Shift:=xlUp
    ElseIf Range("U5") > Range("U6") Then
        Range("P6:U6").Clear
    End If
    
    Range("O6:T6").FormulaR1C1 = "=MAX(R[4]C:R[10003]C)"
    Range("O10").FormulaR1C1 = _
        "=IF(RC[-14]>R5C[1],6371*ACOS(SIN(R5C[2])*SIN(RC[-13])+COS(R5C[2])*COS(RC[-13])*COS(RC[-12]-R5C[3])),"""")"
    Range("P10").FormulaR1C1 = "=IF(RC[-1]=R6C15,RC[-15],"""")"
    Range("Q10:T10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-15],"""")"
    
 With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("O10:T10").AutoFill Destination:=.Range("O10:T" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    
    Range("P6:T6").Value = Range("P6:T6").Value
    Range("U6").Value = Range("O6").Value
    Range("O6").Clear
    
    Columns("G:O").Clear
    Range("P10:T10009").Clear

    Range("T16").FormulaR1C1 = "=LARGE(R[-14]C[1]:R[-10]C[1],4)"
    Range("V1").FormulaR1C1 = "=IF(R[1]C[-1]>=R16C[-2],RC[-6],"""")"
    Range("V2:V6").FormulaR1C1 = "=IF(RC[-1]>=R16C[-2],RC[-6],"""")"
    ActiveSheet.Calculate
    Range("V1:V6").Value = Range("V1:V6").Value
    Range("P1:V6").Sort Key1:=Range("V1"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    
    Range("P5:V16").Clear
    
    Range("P1:T4").Cut Destination:=Range("P2:T5")
    Range("U1:V5").Clear
    
    Range("O1:T1").FormulaR1C1 = "=MAX(R[9]C:R[10009]C)"
    Range("O10").FormulaR1C1 = _
        "=IF(RC[-14]<R2C[1],6371*ACOS(SIN(RC[-13])*SIN(R2C[2])+COS(RC[-13])*COS(R2C[2])*COS(R2C[3]-RC[-12])),"""")"
    Range("P10").FormulaR1C1 = "=IF(RC[-1]=R1C[-1],RC[-15],"""")"
    Range("Q10:T10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-15],"""")"
   
 With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("O10:T10").AutoFill Destination:=.Range("O10:T" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    
    Range("P1:T1").Value = Range("P1:T1").Value
    
    Columns("O:O").Clear
    Range("A7,P10:T10009").Clear
    
    Range("U2:U5").FormulaR1C1 = _
        "=6371*ACOS(SIN(R[-1]C[-4])*SIN(RC[-4])+COS(R[-1]C[-4])*COS(RC[-4])*COS(RC[-3]-R[-1]C[-3]))"
    Range("U6").FormulaR1C1 = "=SUM(R[-4]C:R[-1]C)"
    Range("U7").FormulaR1C1 = _
        "=IF(R[-6]C[-2]-R[-2]C[-2]>R9C4,R[-1]C-((R[-6]C[-2]-R[-2]C[-2]-R9C4)*0.1),R[-1]C)"
    ActiveSheet.Calculate
    Range("U2:U7").Value = Range("U2:U7").Value
    
    'LoH
  If Range("U7") < Range("U6") Then
    Sheets("Sheet2").Range("A7").Value = Sheets("TASKS").Range("C10").Value
    Range("O8:T8").FormulaR1C1 = "=MAX(R[2]C:R[10001]C)"
    
    Range("O10").FormulaR1C1 = _
        "=IF(AND(RC[-14]<R1C[1],RC[-11]-R5C[4]<R1C[4]-R5C[4],RC[-12]<>R2C[3]),6371*ACOS(SIN(RC[-13])*SIN(R2C[2])+COS(RC[-13])*COS(R2C[2])*COS(R2C[3]-RC[-12])),"""")"
    Range("P10").FormulaR1C1 = "=IF(RC[-1]=R8C15,RC[-15],"""")"
    Range("Q10:T10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-15],"""")"
    
 With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("O10:T10").AutoFill Destination:=.Range("O10:T" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    
    Range("P8:T8").Value = Range("P8:T8").Value
    Range("O8:O10009,P10:T10009").Clear
    Range("P9:T12").Value = Range("P2:T5").Value
    
    Range("U9:U12").FormulaR1C1 = _
        "=6371*ACOS(SIN(R[-1]C[-4])*SIN(RC[-4])+COS(R[-1]C[-4])*COS(RC[-4])*COS(RC[-3]-R[-1]C[-3]))"
    Range("U13").FormulaR1C1 = "=SUM(R[-4]C:R[-1]C)"
    Range("U14").FormulaR1C1 = _
        "=IF(R[-6]C[-2]-R[-2]C[-2]>R9C4,R[-1]C-((R[-6]C[-2]-R[-2]C[-2]-R9C4)*0.1),R[-1]C)"
    ActiveSheet.Calculate
    
    If Range("U14") < Range("U13") And Range("P8") <> Range("A7") Then
        Sheets("Sheet2").Range("P15:R15").Value = Sheets("TASKS").Range("C10:E10").Value
        Sheets("Sheet2").Range("S15:T15").Value = Sheets("TASKS").Range("H10:I10").Value
        Range("P16:T19").Value = Range("P2:T5").Value
        Range("U16:U19").FormulaR1C1 = _
            "=6371*ACOS(SIN(R[-1]C[-4])*SIN(RC[-4])+COS(R[-1]C[-4])*COS(RC[-4])*COS(RC[-3]-R[-1]C[-3]))"
        Range("U20").FormulaR1C1 = "=SUM(R[-4]C:R[-1]C)"
        Range("U21").FormulaR1C1 = _
            "=IF(R[-6]C[-2]-R[-2]C[-2]>R9C4,R[-1]C-((R[-6]C[-2]-R[-2]C[-2]-R9C4)*0.1),R[-1]C)"
        ActiveSheet.Calculate
    End If
    
    Range("O1").Formula = "=IF(MAX(R7C21,R14C21,R21C21)=R7C21,""U7"",IF(MAX(R7C21,R14C21,R21C21)=R14C21,""U14"",IF(MAX(R7C21,R14C21,R21C21)=R21C21,""U21"")))"
    Range("A7").Clear
    
    If Range("O1") = "U7" And Range("U7") > Range("F7") Then
        Range("A1:F7").Value = Range("P1:U7").Value
    ElseIf Range("O1") = "U14" And Range("U14") > Range("F7") Then
        Range("A1:F7").Value = Range("P8:U14").Value
    ElseIf Range("O1") = "U21" And Range("U21") > Range("F7") Then
        Range("A1:F7").Value = Range("P15:U21").Value
    End If
  End If
    Columns("O:U").Clear
    Application.Run "F.xlsm!Steve"
  
End Sub
Sub Steve()
'
' Macro1 Macro For Stevenson (PO: Start=Release, TP1=StDist Fin, TP2 OFFC StDist, wander!
'
Application.ScreenUpdating = False
    Sheets("Sheet2").Range("H2:J3").Value = Sheets("TASKS").Range("C10:E11").Value
    Sheets("Sheet2").Range("K2:L3").Value = Sheets("TASKS").Range("H10:I11").Value
    
    Sheets("Sheet2").Range("P2:T3").Value = Sheets("Sheet2").Range("H2:L3").Value
    Range("U3").FormulaR1C1 = "=6371*ACOS(SIN(RC[-4])*SIN(R[-1]C[-4])+COS(RC[-4])*COS(R[-1]C[-4])*COS(R[-1]C[-3]-RC[-3]))"
    
    Range("I4").FormulaR1C1 = _
        "=ACOS(SIN(R2C)*SIN(R3C)+COS(R2C)*COS(R3C)*COS(R3C[1]-R2C[1]))"
    Range("K4").FormulaR1C1 = "=MIN(R[6]C[-2]:R[10005]C[-2])"
    Range("K5").FormulaR1C1 = "=MAX(R[6]C[-2]:R[10005]C[-2])"
    Range("I5").FormulaR1C1 = _
        "=ACOS((SIN(R3C)-SIN(R2C)*COS(R[-1]C))/(SIN(R[-1]C)*COS(R2C)))"
    'OFFC ST Dist, after St Dist FINI
    Range("G10").FormulaR1C1 = "=IF(AND(RC1>R3C[1],RC[-4]<>R2C[3]),ACOS(SIN(R2C[2])*SIN(RC2)+COS(R2C[2])*COS(RC2)*COS(RC3-R2C[3])),"""")"
    Range("H10").FormulaR1C1 = "=IF(RC[-1]<>"""",ACOS((SIN(RC2)-SIN(R2C[1])*COS(RC[-1]))/(SIN(RC[-1])*COS(R2C[1]))),"""")"
    Range("I10").FormulaR1C1 = "=IF(RC[-1]<>"""",(ASIN(SIN(RC[-2])*SIN(RC[-1]-R5C))*180*60*1.852/PI()),"""")"
    Range("J10").FormulaR1C1 = "=IF(OR(RC[-1]=R4C[1],RC[-1]=R5C[1]),RC1,"""")"
    Range("K10:N10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-9],"""")"
    'Copy Ref A
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:N10").AutoFill Destination:=.Range("G10:N" & LastRow), Type:=xlFillDefault
    End With
ActiveSheet.Calculate
    
    Range("J10:N10009").Value = Range("J10:N10009").Value
    Range("J10:N10009").Sort Key1:=Range("J10"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    
    Range("P10:T11").Value = Range("J10:N11").Value
 
    'OFFC St Dist FINI to earlier TP1 Alternative
    Range("H2:L2").Value = Range("P3:T3").Value
    Range("H3:L3").Value = Range("P10:T10").Value
    
    Range("G10").FormulaR1C1 = _
        "=IF(AND(RC1>R2C[1],RC1<R3C[1],RC[-4]<>R2C[3]),ACOS(SIN(R2C[2])*SIN(RC2)+COS(R2C[2])*COS(RC2)*COS(RC3-R2C[3])),"""")"
    Range("J10").FormulaR1C1 = "=IF(OR(RC[-1]=R4C11,RC[-1]=R5C11),RC[-9],"""")"
    Range("K10:N10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-9],"""")"
    'Copy Ref A Value Sort
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:N10").AutoFill Destination:=.Range("G10:N" & LastRow), Type:=xlFillDefault
    End With
ActiveSheet.Calculate
    
    Range("J10:N10009").Value = Range("J10:N10009").Value
    Range("J10:N10009").Sort Key1:=Range("J10"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    
    Range("P13:T14").Value = Range("J10:N11").Value
    
    Range("J10:N10009").Clear
    
    'OFFC St Dist FINI to later TP2 Alternative
    Range("H3:L3").Value = Range("P11:T11").Value
    
    Range("J10").FormulaR1C1 = "=IF(OR(RC[-1]=R4C11,RC[-1]=R5C11),RC[-9],"""")"
    Range("K10:N10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-9],"""")"
    'Copy Ref A Value Sort
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:N10").AutoFill Destination:=.Range("G10:N" & LastRow), Type:=xlFillDefault
    End With
ActiveSheet.Calculate
    
    Range("J10:N10009").Value = Range("J10:N10009").Value
    Range("J10:N10009").Sort Key1:=Range("J10"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    
    Range("P16:T17").Value = Range("J10:N11").Value
    
    Range("U10,U11,U13,U14,U16,U17").FormulaR1C1 = "=6371*ACOS(SIN(RC[-4])*SIN(R3C17)+COS(RC[-4])*COS(R3C17)*COS(R3C18-RC[-3]))"
    ActiveSheet.Calculate
    Range("W10").FormulaR1C1 = "=IF(RC[-2]=MAX(RC[-2],R[1]C[-2]),RC[-7],R[1]C[-7])"
    Range("X10").FormulaR1C1 = "=IF(RC[-1]=RC[-8],RC[-7],R[1]C[-7])"
    Range("Y10").FormulaR1C1 = "=IF(RC[-1]=RC[-8],RC[-7],R[1]C[-7])"
    Range("Z10").FormulaR1C1 = "=IF(RC[-1]=RC[-8],RC[-7],R[1]C[-7])"
    Range("AA10").FormulaR1C1 = "=IF(RC[-1]=RC[-8],RC[-7],R[1]C[-7])"
    ActiveSheet.Calculate
    Range("W13").FormulaR1C1 = "=IF(RC[-2]=MAX(RC[-2],R[1]C[-2]),RC[-7],R[1]C[-7])"
    Range("X13").FormulaR1C1 = "=IF(RC[-1]=RC[-8],RC[-7],R[1]C[-7])"
    Range("Y13").FormulaR1C1 = "=IF(RC[-1]=RC[-8],RC[-7],R[1]C[-7])"
    Range("Z13").FormulaR1C1 = "=IF(RC[-1]=RC[-8],RC[-7],R[1]C[-7])"
    Range("AA13").FormulaR1C1 = "=IF(RC[-1]=RC[-8],RC[-7],R[1]C[-7])"
    ActiveSheet.Calculate
    Range("W16").FormulaR1C1 = "=IF(RC[-2]=MAX(RC[-2],R[1]C[-2]),RC[-7],R[1]C[-7])"
    Range("X16").FormulaR1C1 = "=IF(RC[-1]=RC[-8],RC[-7],R[1]C[-7])"
    Range("Y16").FormulaR1C1 = "=IF(RC[-1]=RC[-8],RC[-7],R[1]C[-7])"
    Range("Z16").FormulaR1C1 = "=IF(RC[-1]=RC[-8],RC[-7],R[1]C[-7])"
    Range("AA16").FormulaR1C1 = "=IF(RC[-1]=RC[-8],RC[-7],R[1]C[-7])"
    ActiveSheet.Calculate
    Range("AB10").FormulaR1C1 = "=IF(RC[-5]=MIN(R10C[-5],R13C[-5],R16C[-5]),6371*ACOS(SIN(RC[-4])*SIN(R3C[-11])+COS(RC[-4])*COS(R3C[-11])*COS(R3C[-10]-RC[-3])),0)"
    Range("AB13").FormulaR1C1 = "=IF(RC[-5]=MIN(R10C[-5],R13C[-5],R16C[-5]),6371*ACOS(SIN(RC[-4])*SIN(R3C[-11])+COS(RC[-4])*COS(R3C[-11])*COS(R3C[-10]-RC[-3])),0)"
    Range("AB16").FormulaR1C1 = "=IF(RC[-5]=MIN(R10C[-5],R13C[-5],R16C[-5]),6371*ACOS(SIN(RC[-4])*SIN(R3C[-11])+COS(RC[-4])*COS(R3C[-11])*COS(R3C[-10]-RC[-3])),0)"
    ActiveSheet.Calculate
    Range("W7").FormulaR1C1 = "=IF(R[3]C[5]=MAX(R[3]C[5],R[6]C[5],R[9]C[5]),R[3]C,IF(R[6]C[5]=MAX(R[3]C[5],R[6]C[5],R[9]C[5]),R[6]C,R[9]C))"
    Range("X7").FormulaR1C1 = "=IF(RC[-1]=R[3]C[-1],R[3]C,IF(RC[-1]=R[6]C[-1],R[6]C,R[9]C))"
    Range("Y7").FormulaR1C1 = "=IF(RC[-2]=R[3]C[-2],R[3]C,IF(RC[-2]=R[6]C[-2],R[6]C,R[9]C))"
    Range("Z7").FormulaR1C1 = "=IF(RC[-3]=R[3]C[-3],R[3]C,IF(RC[-3]=R[6]C[-3],R[6]C,R[9]C))"
    Range("AA7").FormulaR1C1 = "=IF(RC[-4]=R[3]C[-4],R[3]C,IF(RC[-4]=R[6]C[-4],R[6]C,R[9]C))"
    ActiveSheet.Calculate
    Range("P4:T4").Value = Range("W7:AA7").Value
    Range("U4").FormulaR1C1 = "=6371*ACOS(SIN(RC[-4])*SIN(R[-1]C[-4])+COS(RC[-4])*COS(R[-1]C[-4])*COS(R[-1]C[-3]-RC[-3]))"
    ActiveSheet.Calculate
    Range("AB10").FormulaR1C1 = "=IF(AND(RC[-5]<>MIN(R10C[-5],R13C[-5],R16C[-5]),RC[-5]<>MAX(RC[-5],R[3]C[-5],R[6]C[-5])),6371*ACOS(SIN(RC[-4])*SIN(R4C[-11])+COS(RC[-4])*COS(R4C[-11])*COS(R4C[-10]-RC[-3])),0)"
    Range("AB13").FormulaR1C1 = "=IF(AND(RC[-5]<>MIN(R10C[-5],R13C[-5],R16C[-5]),RC[-5]<>MAX(R[-3]C[-5],RC[-5],R[3]C[-5])),6371*ACOS(SIN(RC[-4])*SIN(R4C[-11])+COS(RC[-4])*COS(R4C[-11])*COS(R4C[-10]-RC[-3])),0)"
    Range("AB16").FormulaR1C1 = "=IF(AND(RC[-5]<>MIN(R10C[-5],R13C[-5],R16C[-5]),RC[-5]<>MAX(R[-6]C[-5],R[-3]C[-5],RC[-5])),6371*ACOS(SIN(RC[-4])*SIN(R4C[-11])+COS(RC[-4])*COS(R4C[-11])*COS(R4C[-10]-RC[-3])),0)"
    Range("P5:T5").Value = Range("W7:AA7").Value
    Range("U5").FormulaR1C1 = "=6371*ACOS(SIN(RC[-4])*SIN(R[-1]C[-4])+COS(RC[-4])*COS(R[-1]C[-4])*COS(R[-1]C[-3]-RC[-3]))"
    ActiveSheet.Calculate
    Range("AB10").FormulaR1C1 = "=IF(RC[-5]=MAX(RC[-5],R[3]C[-5],R[6]C[-5]),6371*ACOS(SIN(RC[-4])*SIN(R5C[-11])+COS(RC[-4])*COS(R5C[-11])*COS(R5C[-10]-RC[-3])),0)"
    Range("AB13").FormulaR1C1 = "=IF(RC[-5]=MAX(R[-3]C[-5],RC[-5],R[3]C[-5]),6371*ACOS(SIN(RC[-4])*SIN(R5C[-11])+COS(RC[-4])*COS(R5C[-11])*COS(R5C[-10]-RC[-3])),0)"
    Range("AB16").FormulaR1C1 = "=IF(RC[-5]=MAX(R[-6]C[-5],R[-3]C[-5],RC[-5]),6371*ACOS(SIN(RC[-4])*SIN(R5C[-11])+COS(RC[-4])*COS(R5C[-11])*COS(R5C[-10]-RC[-3])),0)"
    Range("P6:T6").Value = Range("W7:AA7").Value
    Range("U6").FormulaR1C1 = "=6371*ACOS(SIN(RC[-4])*SIN(R[-1]C[-4])+COS(RC[-4])*COS(R[-1]C[-4])*COS(R[-1]C[-3]-RC[-3]))"
    ActiveSheet.Calculate
    Range("V6").FormulaR1C1 = "=SUM(R[-3]C[-1]:RC[-1])"
    Range("V7").FormulaR1C1 = "=IF(R[-5]C[-3]-R[-1]C[-3]<=R9C4,R[-1]C,R[-1]C-((R[-5]C[-3]-R[-1]C[-3]-R9C4)*0.1))"
    ActiveSheet.Calculate
    Range("U3:V7").Value = Range("U3:V7").Value
    Range("W7:AB16").Clear
    Columns("G:N").Clear
    Range("P10:U17").Clear
    'Reverse Ck
    Range("G8:L8").FormulaR1C1 = "=MAX(R[2]C:R[10001]C)"
    
    Range("G10").FormulaR1C1 = "=IF(AND(RC[-6]>R3C[9],RC[-6]<=R4C[9]),6371*ACOS(SIN(RC[-5])*SIN(R3C[10])+COS(RC[-5])*COS(R3C[10])*COS(R3C[11]-RC[-4])),"""")"
    Range("H10").FormulaR1C1 = "=IF(RC[-1]=R8C7,RC[-7],"""")"
    Range("I10:L10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-7],"""")"
    'Copy Ref A
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:L10").AutoFill Destination:=.Range("G10:L" & LastRow), Type:=xlFillDefault
    End With
ActiveSheet.Calculate
    
    Range("P4:T4").Value = Range("H8:L8").Value
    
    Range("G10").FormulaR1C1 = "=IF(AND(RC[-6]>R4C[9],RC[-6]<=R5C[9]),6371*ACOS(SIN(RC[-5])*SIN(R4C[10])+COS(RC[-5])*COS(R4C[10])*COS(R4C[11]-RC[-4])),"""")"
    'Copy Ref A
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:G10").AutoFill Destination:=.Range("G10:G" & LastRow), Type:=xlFillDefault
    End With
ActiveSheet.Calculate

    Range("P5:T5").Value = Range("H8:L8").Value
    
    Range("G10").FormulaR1C1 = "=IF(RC[-6]>R5C[9],6371*ACOS(SIN(RC[-5])*SIN(R5C[10])+COS(RC[-5])*COS(R5C[10])*COS(R5C[11]-RC[-4])),"""")"
    'Copy Ref A
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:G10").AutoFill Destination:=.Range("G10:G" & LastRow), Type:=xlFillDefault
    End With
ActiveSheet.Calculate

    Range("P6:T6").Value = Range("H8:L8").Value
    
    Range("G10").FormulaR1C1 = "=IF(AND(RC[-6]<R6C[9],RC[-6]>R5C[9]),6371*ACOS(SIN(RC[-5])*SIN(R6C[10])+COS(RC[-5])*COS(R6C[10])*COS(R6C[11]-RC[-4])),"""")"
    'Copy Ref A
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:G10").AutoFill Destination:=.Range("G10:G" & LastRow), Type:=xlFillDefault
    End With
ActiveSheet.Calculate
    
    Range("P5:T5").Value = Range("H8:L8").Value
    Range("P2:T6").Sort Key1:=Range("P2"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    
    Range("U3:U6").FormulaR1C1 = "=IF(AND(SUM(R[-1]C[-5]:R[-1]C[-1])<>0,SUM(RC[-5]:RC[-1])<>0),6371*ACOS(SIN(R[-1]C[-4])*SIN(RC[-4])+COS(R[-1]C[-4])*COS(RC[-4])*COS(RC[-3]-R[-1]C[-3])),0)"
    Range("V6").FormulaR1C1 = "=SUM(R[-3]C[-1]:RC[-1])"
    Range("V7").FormulaR1C1 = "=IF(R[-6]C[-2]-R[-2]C[-2]<=R9C4,R[-1]C,R[-1]C-((R[-6]C[-2]-R[-2]C[-2]-R9C4)*0.1))"
    ActiveSheet.Calculate
    
    If Range("V7") > Range("F7") Then
    
        Range("A1:E5").Value = Range("P2:T6").Value
        Range("F2:F5").Value = Range("U3:U6").Value
        Range("F6").FormulaR1C1 = "=SUM(R[-4]C:R[-1]C)"
        Range("F7").FormulaR1C1 = "=IF(R[-6]C[-2]-R[-2]C[-2]<=R9C4,R[-1]C,R[-1]C-((R[-6]C[-2]-R[-2]C[-2]-R9C4)*0.1))"
        ActiveSheet.Calculate
        Range("F6:F7").Value = Range("F6:F7").Value
    End If
    Columns("G:W").Clear

End Sub


Sub AllOrds2()
'
' Works for Doogie
'
Application.ScreenUpdating = False

    Range("G10").FormulaR1C1 = _
        "=IF(RC[-6]>R5C[9],6371*ACOS(SIN(RC[-5])*SIN(R5C[10])+COS(RC[-5])*COS(R5C[10])*COS(R5C[11]-RC[-4])),"""")"
    
 With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:G10").AutoFill Destination:=.Range("G10:G" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    
    Range("B8").FormulaR1C1 = "=MIN(R[-6]C[19]:R[-3]C[19])"
    
  If Range("G8") > Range("B8") Then
    
        Range("P6:T6").Value = Range("H8:L8").Value
        Range("U6").Value = Range("G8").Value
        Range("T11").FormulaR1C1 = "=LARGE(R[-9]C[1]:R[-5]C[1],4)"
        Range("V1").FormulaR1C1 = "=IF(R[1]C[-1]>=R11C[-2],RC[-6],"""")"
        Range("V2:V6").FormulaR1C1 = "=IF(RC[-1]>=R11C[-2],RC[-6],"""")"
        ActiveSheet.Calculate
        Range("V1:V6").Value = Range("V1:V6").Value
    
        Range("P1:V6").Sort Key1:=Range("V1"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
        Range("U1:V5,P5:V11").Clear
    
        Range("G10").FormulaR1C1 = _
            "=IF(RC[-6]<R1C[9],6371*ACOS(SIN(RC[-5])*SIN(R1C[10])+COS(RC[-5])*COS(R1C[10])*COS(R1C[11]-RC[-4])),"""")"
    
 With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:G10").AutoFill Destination:=.Range("G10:G" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    
    Range("P5:T5").Value = Range("H8:L8").Value
    Range("P1:T5").Sort Key1:=Range("P1"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    
    Range("U2:U5").FormulaR1C1 = _
        "=6371*ACOS(SIN(R[-1]C[-4])*SIN(RC[-4])+COS(R[-1]C[-4])*COS(RC[-4])*COS(RC[-3]-R[-1]C[-3]))"
    Range("U6").FormulaR1C1 = "=SUM(R[-4]C:R[-1]C)"
    Range("U7").FormulaR1C1 = _
        "=IF(R[-6]C[-2]-R[-2]C[-2]>R9C4,R[-1]C-((R[-6]C[-2]-R[-2]C[-2]-R9C4)*0.1),R[-1]C)"
    ActiveSheet.Calculate
    Range("U2:U7").Value = Range("U2:U7").Value
    
   End If
 'End If
End Sub

Sub ORbase()
'
' ORbase Macro works for Payne 2012 based on O&R task
'
Application.ScreenUpdating = False
    Sheets("Sheet2").Range("P2:R4").Value = Sheets("TASKS").Range("C14:E16").Value
    Sheets("Sheet2").Range("S2:T4").Value = Sheets("TASKS").Range("H14:I16").Value
     
    Range("A8,G8:L8").FormulaR1C1 = "=MAX(R[2]C:R[10001]C)"
    
    Range("G10").FormulaR1C1 = _
        "=If(RC[-6]>R3C[9],6371*ACOS(SIN(RC[-5])*SIN(R3C[10])+COS(RC[-5])*COS(R3C[10])*COS(R3C[11]-RC[-4])),"""")"
    Range("H10").FormulaR1C1 = "=IF(RC[-1]=R8C[-1],RC[-7],"""")"
    Range("I10:L10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-7],"""")"

 With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:L10").AutoFill Destination:=.Range("G10:L" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Sheets("Sheet2").Cells.NumberFormat = "General"
    Range("P5:T5").Value = Range("H8:L8").Value
    
    Range("G10").FormulaR1C1 = _
            "=IF(AND(RC[-6]>R4C[9],RC[-6]<R5C[9]),6371*ACOS(SIN(RC[-5])*SIN(R4C[10])+COS(RC[-5])*COS(R4C[10])*COS(R4C[11]-RC[-4])),"""")"

 With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:G10").AutoFill Destination:=.Range("G10:G" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
       
 If Range("G8") <> 0 Then
    Range("P8:T8").Value = Range("H8:L8").Value
    
    Range("G10").FormulaR1C1 = _
        "=IF(AND(RC[-6]<R8C[9],RC[-6]>R3C[9]),6371*ACOS(SIN(RC[-5])*SIN(R8C[10])+COS(RC[-5])*COS(R8C[10])*COS(R8C[11]-RC[-4])),"""")"
    
 With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:G10").AutoFill Destination:=.Range("G10:G" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    
    Range("P9:T9").Value = Range("H8:L8").Value
    
    Range("G10").FormulaR1C1 = _
        "=IF(RC[-6]>R8C[9],6371*ACOS(SIN(RC[-5])*SIN(R8C[10])+COS(RC[-5])*COS(R8C[10])*COS(R8C[11]-RC[-4])),"""")"
    
 With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:G10").AutoFill Destination:=.Range("G10:G" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    
    Range("P10:T10").Value = Range("H8:L8").Value
    
    Range("G10").FormulaR1C1 = _
        "=IF(RC[-6]<R9C[9],6371*ACOS(SIN(RC[-5])*SIN(R9C[10])+COS(RC[-5])*COS(R9C[10])*COS(R9C[11]-RC[-4])),"""")"
    
 With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:G10").AutoFill Destination:=.Range("G10:G" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    
    Range("P11:T11").Value = Range("H8:L8").Value
    
    Range("G10").FormulaR1C1 = _
        "=IF(RC[-6]<R11C[9],6371*ACOS(SIN(RC[-5])*SIN(R11C[10])+COS(RC[-5])*COS(R11C[10])*COS(R11C[11]-RC[-4])),"""")"
    
 With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:G10").AutoFill Destination:=.Range("G10:G" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    
    Range("P12:T12").Value = Range("H8:L8").Value
    
    Range("G10").FormulaR1C1 = _
        "=IF(RC[-6]<R2C[9],6371*ACOS(SIN(RC[-5])*SIN(R11C[10])+COS(RC[-5])*COS(R11C[10])*COS(R11C[11]-RC[-4])),"""")"
    
 With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:G10").AutoFill Destination:=.Range("G10:G" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    
    Range("P8:T12").Sort Key1:=Range("P8"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
  
    Range("U9:U12").FormulaR1C1 = _
        "=IF(OR(R[-1]C[-4] = 0, RC[-4]=0),"""",6371*ACOS(SIN(R[-1]C[-4])*SIN(RC[-4])+COS(R[-1]C[-4])*COS(RC[-4])*COS(RC[-3]-R[-1]C[-3])))"
    Range("U13").FormulaR1C1 = "=SUM(R[-4]C:R[-1]C)"
    Range("U14").FormulaR1C1 = _
        "=IF(R[-6]C[-2]-R[-2]C[-2]>R9C4,R[-1]C-((R[-6]C[-2]-R[-2]C[-2]-R9C4)*0.1),R[-1]C)"
    ActiveSheet.Calculate
    Range("P1:U7").Value = Range("P8:U14").Value
    
    Range("P8:U14").Clear
    Columns("G:L").Clear
    
 ElseIf Range("G8") = 0 Then
    
    Range("P4:T4").Clear
    Range("G10").FormulaR1C1 = _
        "=IF(RC[-6]>R5C[9],6371*ACOS(SIN(RC[-5])*SIN(R5C[10])+COS(RC[-5])*COS(R5C[10])*COS(R5C[11]-RC[-4])),"""")"
 
 With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:G10").AutoFill Destination:=.Range("G10:G" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    
    Range("P8:T8").Value = Range("H8:L8").Value
    
    Range("G10").FormulaR1C1 = _
        "=IF(RC[-6]<R2C[9],6371*ACOS(SIN(RC[-5])*SIN(R2C[10])+COS(RC[-5])*COS(R2C[10])*COS(R2C[11]-RC[-4])),"""")"
    
 With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:G10").AutoFill Destination:=.Range("G10:G" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    
    Range("P9:T9").Value = Range("H8:L8").Value
    
    Range("G10").FormulaR1C1 = _
        "=IF(AND(RC[-6]>R9C[9],RC[-6]<R3C[9]),6371*ACOS(SIN(RC[-5])*SIN(R3C[10])+COS(RC[-5])*COS(R3C[10])*COS(R3C[11]-RC[-4])),"""")"
    
 With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:G10").AutoFill Destination:=.Range("G10:G" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    
    Range("P10:T10").Value = Range("H8:L8").Value
    
    Range("P2:T2").Clear
    
    Range("P1:T10").Sort Key1:=Range("P1"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
  
    Range("U2:U7").FormulaR1C1 = _
        "=IF(OR(R[-1]C[-4] = 0, RC[-4]=0),"""",6371*ACOS(SIN(R[-1]C[-4])*SIN(RC[-4])+COS(R[-1]C[-4])*COS(RC[-4])*COS(RC[-3]-R[-1]C[-3])))"
    ActiveSheet.Calculate
    Range("G10").FormulaR1C1 = "=IF(RC[-6]<R2C[9],6371*ACOS(SIN(RC[-5])*SIN(R2C[10])+COS(RC[-5])*COS(R2C[10])*COS(R2C[11]-RC[-4])),"""")"
  
 With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:G10").AutoFill Destination:=.Range("G10:G" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
  
    If Range("G8") > Range("U2") Then
        Range("P2:T2").Value = Range("H8:L8").Value
    End If
  
    Range("U6").FormulaR1C1 = "=SUM(R[-4]C:R[-1]C)"
    Range("U7").FormulaR1C1 = _
        "=IF(R[-6]C[-2]-R[-2]C[-2]>R9C4,R[-1]C-((R[-6]C[-2]-R[-2]C[-2]-R9C4)*0.1),R[-1]C)"
    ActiveSheet.Calculate
    Range("P1:U7").Value = Range("P1:U7").Value
  
    Columns("G:L").Clear
       
 End If
    
    If Range("U7") > Range("F7") Then
        Range("A1:F7").Value = Range("P1:U7").Value
    End If
        
    Range("P1:U7").Clear
    
End Sub


Sub Fixxer()
' 7/27/17 JLR Consolidated w/Fixxer2 for brevity
'
Application.ScreenUpdating = False
    
    Range("I8,J8,N8,O8").FormulaR1C1 = "=MAX(R[2]C:R[60001]C)"
    
    Range("G10").FormulaR1C1 = "=IF(AND(RC[-6]>R3C1,RC[-6]<R5C1),6371*ACOS(SIN(RC[-5])*SIN(R5C2)+COS(RC[-5])*COS(R5C2)*COS(R5C3-RC[-4])),"""")"
    Range("H10").FormulaR1C1 = "=IF(RC[-1]<>"""",6371*ACOS(SIN(RC[-6])*SIN(R3C2)+COS(RC[-6])*COS(R3C2)*COS(R3C3-RC[-5])),"""")"
    Range("I10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-2]+RC[-1],"""")"
    Range("J10").FormulaR1C1 = "=IF(RC[-1]=R8C9,RC[-9],"""")"
    Range("L10").FormulaR1C1 = "=IF(AND(RC[-11]>R1C1,RC[-11]<R3C1),6371*ACOS(SIN(RC[-10])*SIN(R3C2)+COS(RC[-10])*COS(R3C2)*COS(R3C3-RC[-9])),"""")"
    Range("M10").FormulaR1C1 = "=IF(RC[-1]<>"""",6371*ACOS(SIN(RC[-11])*SIN(R1C2)+COS(RC[-11])*COS(R1C2)*COS(R1C3-RC[-10])),"""")"
    Range("N10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-2]+RC[-1],"""")"
    Range("O10").FormulaR1C1 = "=IF(RC[-1]=R8C14,RC[-14],"""")"
    'Copy Ref A
 With ActiveSheet
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:O10").AutoFill Destination:=.Range("G10:O" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    
    Range("J8").Value = Range("J8").Value
    Range("O8").Value = Range("O8").Value
  
    Range("G10:O60009").Clear
    
    Range("G10").FormulaR1C1 = "=IF(OR(RC[-6]=R8C15,RC[-6]=R3C1,RC[-6]=R8C10),RC[-6],"""")"
    Range("H10:K10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-6],"""")"
    'Copy Ref A Value Sort
 With ActiveSheet
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:K10").AutoFill Destination:=.Range("G10:K" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    
    Range("G10:K60009").Value = Range("G10:K60009").Value
    Range("G10:K60009").Sort Key1:=Range("G10"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    
    Range("L10").FormulaR1C1 = _
        "=IF(RC[-1]="""",0,6371*ACOS(SIN(R[-9]C[-10])*SIN(RC[-4])+COS(R[-9]C[-10])*COS(RC[-4])*COS(RC[-3]-R[-9]C[-9])))"
    Range("L11:L12").FormulaR1C1 = _
        "=IF(RC[-1]="""",0,6371*ACOS(SIN(R[-1]C[-4])*SIN(RC[-4])+COS(R[-1]C[-4])*COS(RC[-4])*COS(RC[-3]-R[-1]C[-3])))"
    Range("L13").FormulaR1C1 = _
        "=IF(RC[-1]="""",0,6371*ACOS(SIN(R[-1]C[-4])*SIN(R[-8]C[-10])+COS(R[-1]C[-4])*COS(R[-8]C[-10])*COS(R[-8]C[-9]-R[-1]C[-3])))"
    Range("L14").FormulaR1C1 = "=SUM(R[-4]C:R[-1]C)"
    Range("L15").FormulaR1C1 = _
        "=IF(R[-14]C[-8]-R[-10]C[-8]<=R9C4,R[-1]C,R[-1]C-((R[-14]C[-8]-R[-10]C[-8]-R9C4)*0.1))"
    ActiveSheet.Calculate
    
    If Range("L15") > Range("F7") Then
        Range("A2:E4").Value = Range("G10:K12").Value
        Range("F2:F7").Value = Range("L10:L15").Value
    End If
    
    Range("I6:N6").FormulaR1C1 = "=MAX(R[4]C:R[60003]C)"
    Range("J6").FormulaR1C1 = "=MAX(R[4]C:R[60003]C)"
    
    Range("G10").FormulaR1C1 = "=IF(AND(RC[-6]>R2C1,RC[-6]<R4C1),6371*ACOS(SIN(RC[-5])*SIN(R2C2)+COS(RC[-5])*COS(R2C2)*COS(R2C3-RC[-4])),"""")"
    Range("H10").FormulaR1C1 = "=IF(RC[-1]<>"""",6371*ACOS(SIN(RC[-6])*SIN(R4C2)+COS(RC[-6])*COS(R4C2)*COS(R4C3-RC[-5])),"""")"
    Range("I10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-2]+RC[-1],"""")"
    Range("J10").FormulaR1C1 = "=IF(RC[-1]=R6C9,RC[-9],"""")"
    'Copy Ref A
    With ActiveSheet
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:J10").AutoFill Destination:=.Range("G10:J" & LastRow), Type:=xlFillDefault
    End With
ActiveSheet.Calculate
    
    Range("I6:J6").Value = Range("I6:J6").Value

    Range("K10").FormulaR1C1 = "=IF(RC[-1]=R6C10,RC[-9],"""")"
    Range("L10:N10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-9],"""")"
    'Copy Ref A
    With ActiveSheet
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("K10:N10").AutoFill Destination:=.Range("K10:N" & LastRow), Type:=xlFillDefault
    End With
ActiveSheet.Calculate

    Range("I5").FormulaR1C1 = "=R[-2]C[-3]+R[-1]C[-3]"
    ActiveSheet.Calculate
    If Range("I6") > Range("I5") Then
        Range("A3:E3").Value = Range("J6:N6").Value
        Range("F3:F4").FormulaR1C1 = "=6371*ACOS(SIN(R[-1]C[-4])*SIN(RC[-4])+COS(R[-1]C[-4])*COS(RC[-4])*COS(RC[-3]-R[-1]C[-3]))"
        Range("F6").FormulaR1C1 = "=SUM(R[-4]C:R[-1]C)"
        Range("F7").FormulaR1C1 = "=IF(R[-6]C[-2]-R[-2]C[-2]<=R9C4,R[-1]C,R[-1]C-((R[-6]C[-2]-R[-2]C[-2]-R9C4)*0.1))"
        ActiveSheet.Calculate
        Range("F2:F7").Value = Range("F2:F7").Value
    End If
    Columns("G:O").Clear
End Sub

Sub TP3a()
'
' OFF ST Dist Course
'
Application.ScreenUpdating = False
    Sheets("Sheet2").Range("H2:I3").Value = Sheets("Tasks").Range("C10:D11").Value
    Sheets("Sheet2").Range("J2:J3").Value = Sheets("Tasks").Range("E10:E11").Value
    Sheets("Sheet2").Range("K2:L3").Value = Sheets("Tasks").Range("H10:I11").Value
    
    Range("I4").FormulaR1C1 = _
        "=ACOS(SIN(R2C)*SIN(R3C)+COS(R2C)*COS(R3C)*COS(R3C[1]-R2C[1]))"
    Range("I5").FormulaR1C1 = _
        "=ACOS((SIN(R3C)-SIN(R2C)*COS(R[-1]C))/(SIN(R[-1]C)*COS(R2C)))"
    Range("J8:N8").FormulaR1C1 = "=MAX(R[2]C:R[10001]C)"
    'L&R off course after ST DIST Fini
    Range("G10").FormulaR1C1 = _
        "=IF(RC[-4]=R2C10,"""",IF(RC1>R3C[1],ACOS(SIN(R2C[2])*SIN(RC2)+COS(R2C[2])*COS(RC2)*COS(RC3-R2C[3])),""""))"
    Range("H10").FormulaR1C1 = _
        "=IF(RC[-1]<>"""",ACOS((SIN(RC2)-SIN(R2C[1])*COS(RC[-1]))/(SIN(RC[-1])*COS(R2C[1]))),"""")"
    Range("I10").FormulaR1C1 = _
        "=IF(RC[-1]<>"""",(ASIN(SIN(RC[-2])*SIN(RC[-1]-R5C))*180*60*1.852/PI()),"""")"
    Range("J10").FormulaR1C1 = "=IF(OR(RC[-1]=R4C[1],RC[-1]=R5C[1]),RC1,"""")"
    Range("K10:N10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-9],"""")"
    
    Range("K4").FormulaR1C1 = "=MAX(R[6]C[-2]:R[10005]C[-2])"
    Range("K5").FormulaR1C1 = "=MIN(R[5]C[-2]:R[10004]C[-2])"
    Range("L4").FormulaR1C1 = "=MAX(R[6]C[-2]:R[10005]C[-2])"
    
 With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:N10").AutoFill Destination:=.Range("G10:N" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    
    Range("J10:N10009").Value = Range("J10:N10009").Value
    Range("G10:I10009").Clear
    Range("J10:N10009").Sort Key1:=Range("J10"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    Range("K4:L4").Clear
    
    Range("P4:T5").Value = Range("J10:N11").Value
    Range("J10:N11").Clear
    
    Range("G8:L8").FormulaR1C1 = "=MAX(R[2]C:R[10001]C)"
    
    Range("G10").FormulaR1C1 = _
        "=IF(RC[-6]<R2C8,6371*ACOS(SIN(RC[-5])*SIN(R2C9)+COS(RC[-5])*COS(R2C9)*COS(R2C10-RC[-4])),"""")"
    Range("H10").FormulaR1C1 = "=IF(RC[-1]=R8C7,RC[-7],"""")"
    Range("I10:L10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-7],"""")"
    
 With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:L10").AutoFill Destination:=.Range("G10:L" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    
    Range("P1:T1").Value = Range("H8:L8").Value
    Range("P2:T3").Value = Range("H2:L3").Value
    
    Range("O2:O5").FormulaR1C1 = _
        "=IF(OR(R[-1]C[1]=0,RC[1]=0),0,6371*ACOS(SIN(R[-1]C[2])*SIN(RC[2])+COS(R[-1]C[2])*COS(RC[2])*COS(RC[3]-R[-1]C[3])))"
    Range("O6").FormulaR1C1 = "=SUM(R[-4]C:R[-1]C)"
    ActiveSheet.Calculate
    Range("O1:T6").Value = Range("O1:T6").Value
    
    Range("G8:G10009").Clear
    
    'CK LoH, adjust Start
    Range("P6").FormulaR1C1 = _
        "=IF(R[-5]C[3]-R[-1]C[3]>R9C4,RC[-1]-((R[-5]C[3]-R[-1]C[3]-R9C4)*0.1),RC[-1])"
    ActiveSheet.Calculate
    Range("P6").Value = Range("P6").Value
 
 If Range("P6") <> Range("O6") Then
    Range("A7").Value = 2.08333333333333E-02
    Range("G8:L8").FormulaR1C1 = "=MAX(R[2]C:R[10001]C)"
    ActiveSheet.Calculate
    Range("G10").FormulaR1C1 = _
        "=IF(AND(RC[-6]>=R1C16-R7C1,RC[-6]<=R1C16+R7C1,RC[-3]-R5C19<=R9C4),6371*ACOS(SIN(RC[-5])*SIN(R2C17)+COS(RC[-5])*COS(R2C17)*COS(R2C18-RC[-4])),"""")"
     Range("H10").FormulaR1C1 = "=IF(RC[-1]=R8C7,RC[-7],"""")"
    Range("I10:L10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-7],"""")"
    
 With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:G10").AutoFill Destination:=.Range("G10:G" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    
    Range("E7").FormulaR1C1 = "=R2C15-R8C7"
    Range("E8").FormulaR1C1 = "=IF(R[-1]C>0,R6C15-R6C16,0)"
    ActiveSheet.Calculate
    Range("E7:E8").Value = Range("E7:E8").Value
    
    If Range("E7") < Range("E8") And Range("G8") <> 0 Then
    
    Range("P8:T8").Value = Range("H8:L8").Value
    Range("P9:T12").Value = Range("P2:T5").Value
    
    Range("O9:O12").FormulaR1C1 = _
        "=IF(OR(AND(R[-1]C[1]=0,R[-1]C[2]=0),AND(RC[1]=0,RC[2]=0)),0,6371*ACOS(SIN(R[-1]C[2])*SIN(RC[2])+COS(R[-1]C[2])*COS(RC[2])*COS(RC[3]-R[-1]C[3])))"
    Range("O13").FormulaR1C1 = "=SUM(R[-4]C:R[-1]C)"
    ActiveSheet.Calculate
    Range("O9:O13").Value = Range("O9:O13").Value
    
    Range("P13").FormulaR1C1 = _
        "=IF(R[-5]C[3]-R[-1]C[3]>R9C4,RC[1]-((R[-5]C[3]-R[-1]C[3]-R9C4)*0.1),RC[-1])"
    ActiveSheet.Calculate
    Range("P13").Value = Range("P13").Value
 
    End If
 End If
 
    Range("E7,E8,G8:L10009").Clear
    
    Range("E7").FormulaR1C1 = "=MIN(R2C[10]:R5C[10],R9C[10]:R12C[10])"
    ActiveSheet.Calculate
    Range("E7").Value = Range("E7").Value
    
    'If for near-duplicate TPs, ElseIf when no TP3, relese hour(s) before start
    If Range("E7") < 5 And Range("E7") > 0 Then
        Application.Run "F.xlsm!TP3b"
    ElseIf Range("E7") > 5 Then
        Application.Run "F.xlsm!TP3c"
    End If
    
    Range("N1").FormulaR1C1 = "=MAX(R7C6,R6C16,R13C16,R20C16,R27C16,R34C16,R41C16)"
    ActiveSheet.Calculate
    Range("N1").Value = Range("N1").Value
    
    If Range("N1") = Range("F7") Then
        Range("O1:T41").Clear
    ElseIf Range("N1") = Range("P6") Then
        Range("A1:E5").Value = Range("P1:T5").Value
        Range("F2:F6").Value = Range("O2:O6").Value
        Range("F7").Value = Range("P6").Value
    ElseIf Range("N1") = Range("P13") Then
        Range("A1:E5").Value = Range("P8:T12").Value
        Range("F2:F6").Value = Range("O9:O13").Value
        Range("F7").Value = Range("P13").Value
    ElseIf Range("N1") = Range("P20") Then
        Range("A1:E5").Value = Range("P15:T19").Value
        Range("F2:F6").Value = Range("O16:O20").Value
        Range("F7").Value = Range("P20").Value
    ElseIf Range("N1") = Range("P27") Then
        Range("A1:E5").Value = Range("P22:T26").Value
        Range("F2:F6").Value = Range("O23:O27").Value
        Range("F7").Value = Range("P27").Value
    ElseIf Range("N1") = Range("P34") Then
        Range("A1:E5").Value = Range("P29:T33").Value
        Range("F2:F6").Value = Range("O30:O34").Value
        Range("F7").Value = Range("P34").Value
    ElseIf Range("N1") = Range("P41") Then
        Range("A1:E5").Value = Range("P36:T40").Value
        Range("F2:F6").Value = Range("O37:O41").Value
        Range("F7").Value = Range("P41").Value
    End If
    
    Range("A7,E7:E8").Clear
    Columns("G:T").Clear
 
End Sub

Sub TP3c()
'
' JLR 12/4/13 works for Scutter Based on O&R
'
Application.ScreenUpdating = False
    Range("P17:T19").Value = Range("P8:T10").Value
    Range("O18:O19").Value = Range("O9:O10").Value
    Sheets("Sheet2").Range("H2:J2").Value = Sheets("Tasks").Range("C14:E14").Value
    Sheets("Sheet2").Range("K2:L2").Value = Sheets("Tasks").Range("H14:I14").Value
    Range("H3:L3").Value = Range("P17:T17").Value
    Range("P15:T15").Value = Range("H2:L2").Value
    
    Range("G10").FormulaR1C1 = _
        "=IF(AND(RC1<R3C[1],RC3<>R2C[3]),ACOS(SIN(R2C[2])*SIN(RC2)+COS(R2C[2])*COS(RC2)*COS(RC3-R2C[3])),"""")"
    Range("H10").FormulaR1C1 = _
        "=IF(RC[-1]<>"""",ACOS((SIN(RC2)-SIN(R2C[1])*COS(RC[-1]))/(SIN(RC[-1])*COS(R2C[1]))),"""")"
    Range("I10").FormulaR1C1 = _
        "=IF(RC[-1]<>"""",ABS((ASIN(SIN(RC[-2])*SIN(RC[-1]-R5C))*180*60*1.852/PI())),"""")"
    Range("J10").FormulaR1C1 = "=IF(RC[-1]=R4C[1],RC1,"""")"
    Range("K10:N10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-9],"""")"
    
    Range("K4").FormulaR1C1 = "=MAX(R[6]C[-2]:R[10005]C[-2])"
    Range("L4").FormulaR1C1 = "=MAX(R[6]C[-2]:R[10005]C[-2])"
    
 With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:N10").AutoFill Destination:=.Range("G10:N" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    
    Range("J10:N10009").Value = Range("J10:N10009").Value
    Range("G10:I10009").Clear
    Range("J10:N10009").Sort Key1:=Range("J10"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    
    Range("P16:T16").Value = Range("J10:N10").Value
    Range("J10:N10").Clear
    'need to sort
    Range("P15:T19").Sort Key1:=Range("P15"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    
    Range("O16:O19").FormulaR1C1 = _
        "=IF(AND(RC[1]<>"""",R[-1]C[1]<>0,RC[1]<>0),6371*ACOS(SIN(R[-1]C[2])*SIN(RC[2])+COS(R[-1]C[2])*COS(RC[2])*COS(R[-1]C[3]-RC[3])),0)"
    Range("O20").FormulaR1C1 = "=SUM(R[-4]C:R[-1]C)"
    Range("P20").FormulaR1C1 = _
        "=IF(R[-5]C[3]-R[-1]C[3]>R9C4,RC[-1]-((R[-5]C[3]-R[-1]C[3]-R9C4)*0.1),RC[-1])"
    ActiveSheet.Calculate
    Range("O16:O20").Value = Range("O16:O20").Value
    Range("P20").Value = Range("P20").Value
    
    If Range("P20") < Range("O20") Then
    Range("A7").Value = 2.08333333333333E-02
    Range("G8:L8").FormulaR1C1 = "=MAX(R[2]C:R[10001]C)"
    
    Range("G10").FormulaR1C1 = _
        "=IF(AND(RC[-6]>=R15C16-R7C1,RC[-6]<=R15C16+R7C1,RC[-3]-R19C19<=R9C4),6371*ACOS(SIN(RC[-5])*SIN(R16C17)+COS(RC[-5])*COS(R16C17)*COS(R16C18-RC[-4])),"""")"
    Range("H10").FormulaR1C1 = "=IF(RC[-1]=R8C7,RC[-7],"""")"
    Range("I10:L10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-7],"""")"
    
 With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:L10").AutoFill Destination:=.Range("G10:L" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    
    Range("E7").FormulaR1C1 = "=R16C15-R8C7"
    Range("E8").FormulaR1C1 = "=IF(R[-1]C>0,R20C15-R20C16,0)"
    ActiveSheet.Calculate
    Range("E7:E8").Value = Range("E7:E8").Value
    
    If Range("E7") < Range("E8") And Range("G8") <> 0 Then

    Range("P22:T22").Value = Range("H8:L8").Value
    Range("P23:T26").Value = Range("P16:T19").Value
    
    Range("O23:O26").FormulaR1C1 = _
        "=IF(AND(RC[1]<>"""",R[-1]C[1]<>0,RC[1]<>0),6371*ACOS(SIN(R[-1]C[2])*SIN(RC[2])+COS(R[-1]C[2])*COS(RC[2])*COS(RC[3]-R[-1]C[3])),0)"
    Range("O27").FormulaR1C1 = "=SUM(R[-4]C:R[-1]C)"
    ActiveSheet.Calculate
    Range("O23:O27").Value = Range("O23:O27").Value
    
    Range("P27").FormulaR1C1 = _
        "=IF(R[-5]C[3]-R[-1]C[3]>R9C4,RC[1]-((R[-5]C[3]-R[-1]C[3]-R9C4)*0.1),RC[-1])"
    ActiveSheet.Calculate
    Range("P27").Value = Range("P27").Value
    End If
  End If
      
End Sub
Sub Serk()
'
' Works for Serkowski, Ords 1,2,3 then wing it! (Funky)
'
 Application.ScreenUpdating = False
    'Copy ORDS from TASKS to P1:T4
    Sheets("Sheet2").Range("P1:T4").Value = Sheets("TASKS").Range("A40:E43").Value
    
    Range("A7").Value = 1 / 24
    Range("O2").FormulaR1C1 = "=RC[1]-R[-1]C[1]"
    Range("O3").FormulaR1C1 = "=RC[1]-R[-1]C[1]"
    Range("O4").FormulaR1C1 = "=RC[1]-R[-1]C[1]"
    Range("N4").FormulaR1C1 = "=R[-1]C[2]+(RC[1]/2)+R[3]C[-13]"
    
    Range("G8:L8").FormulaR1C1 = "=MAX(R[2]C:R[10001]C)"
    Range("G10").FormulaR1C1 = _
        "=IF(AND(RC[-6]>R3C16,RC[-6]<=R4C14,RC[-4]<>R3C18),6371*ACOS(SIN(RC[-5])*SIN(R3C17)+COS(RC[-5])*COS(R3C17)*COS(R3C18-RC[-4])),"""")"
    Range("H10").FormulaR1C1 = "=IF(RC[-1]=R8C7,RC[-7],"""")"
    Range("I10:L10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-7],"""")"
    'Copy Ref A
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:L10").AutoFill Destination:=.Range("G10:L" & LastRow), Type:=xlFillDefault
    End With
ActiveSheet.Calculate
    
    Range("P4:T4").Value = Range("H8:L8").Value
    
    Range("U2:U4").FormulaR1C1 = "=IF(AND(RC[-5]<>0,R[-1]C[-5]<>0),6371*ACOS(SIN(R[-1]C[-4])*SIN(RC[-4])+COS(R[-1]C[-4])*COS(RC[-4])*COS(RC[-3]-R[-1]C[-3])),"""")"
    ActiveSheet.Calculate
    Range("G10:G10009").Clear
    
    Range("G10").FormulaR1C1 = _
        "=IF(RC[-4]=R4C18,"""",IF(RC[-6]>R4C16,6371*ACOS(SIN(RC[-5])*SIN(R4C17)+COS(RC[-5])*COS(R4C17)*COS(R4C18-RC[-4])),""""))"
    'Copy Ref A
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:G10").AutoFill Destination:=.Range("G10:G" & LastRow), Type:=xlFillDefault
    End With
ActiveSheet.Calculate
    
    Range("P5:T5").Value = Range("H8:L8").Value
    
    Range("U5").FormulaR1C1 = "=IF(AND(RC[-5]<>0,R[-1]C[-5]<>0),6371*ACOS(SIN(R[-1]C[-4])*SIN(RC[-4])+COS(R[-1]C[-4])*COS(RC[-4])*COS(RC[-3]-R[-1]C[-3])),"""")"
    
    Range("U6").FormulaR1C1 = "=SUM(R[-4]C:R[-1]C)"
    Range("U7").FormulaR1C1 = "=IF(R[-6]C[-2]-R[-2]C[-2]<R9C4,R[-1]C,R[-1]C-((R[-6]C[-2]-R[-2]C[-2])*0.1))"
    ActiveSheet.Calculate
    Range("U2:U7").Value = Range("U2:U7").Value
    
    Range("G1:O10009").Clear
    
    If Range("U7") > Range("F7") Then
        Range("A1:F7").Value = Range("P1:U7").Value
    End If
    Range("A7:A8,P1:U7").Clear
    
End Sub
Sub Essex()
'
' Chk Ords, amending S&F only
'
Application.ScreenUpdating = False
    
    Sheets("Sheet2").Range("P2:T4").Value = Sheets("Tasks").Range("A41:E43").Value
    
    Range("G8:L8").FormulaR1C1 = "=MAX(R[2]C:R[10001]C)"
    Range("G10").FormulaR1C1 = "=IF(RC[-6]<R2C16,6371*ACOS(SIN(RC[-5])*SIN(R2C17)+COS(RC[-5])*COS(R2C17)*COS(R2C18-RC[-4])),"""")"
    Range("H10").FormulaR1C1 = "=IF(RC[-1]=R8C7,RC[-7],"""")"
    Range("I10:L10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-7],"""")"
    'Copy Ref A
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:L10").AutoFill Destination:=.Range("G10:L" & LastRow), Type:=xlFillDefault
    End With
ActiveSheet.Calculate
    
    Range("P1:T1").Value = Range("H8:L8").Value
    Range("G8:L10009").Clear
    
    Range("N8:S8").FormulaR1C1 = "=MAX(R[2]C:R[10001]C)"
    Range("N10").FormulaR1C1 = "=IF(RC[-13]>R4C16,6371*ACOS(SIN(RC[-12])*SIN(R4C17)+COS(RC[-12])*COS(R4C17)*COS(R4C18-RC[-11])),"""")"
    Range("O10").FormulaR1C1 = "=IF(RC[-1]=R8C14,RC[-14],"""")"
    Range("P10:S10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-14],"""")"
    'Copy Ref A
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("N10:S10").AutoFill Destination:=.Range("N10:S" & LastRow), Type:=xlFillDefault
    End With
ActiveSheet.Calculate
    
    Range("P5:T5").Value = Range("O8:S8").Value
    Range("N8:S10009").Clear
    
    Range("U2:U5").FormulaR1C1 = "=IF(AND(RC[-5]<>0,R[-1]C[-5]<>0),6371*ACOS(SIN(R[-1]C[-4])*SIN(RC[-4])+COS(R[-1]C[-4])*COS(RC[-4])*COS(RC[-3]-R[-1]C[-3])),"""")"
    
    Range("U6").FormulaR1C1 = "=SUM(R[-4]C:R[-1]C)"
    Range("U7").FormulaR1C1 = _
        "=IF(AND(R[-2]C[-2]<>"""",R[-6]C[-2]-R[-2]C[-2]>R9C4),R[-1]C-((R[-6]C[-2]-R[-2]C[-2]-R9C4)*0.1),IF(AND(R[-2]C[-2]="""",R[-6]C[-2]-R[-3]C[-2]>R9C4),R[-1]C-((R[-6]C[-2]-R[-3]C[-2]-R9C4)*0.1),IF(AND(R[-2]C[-2]="""",R[-3]C[-2]="""",R[-6]C[-2]-R[-4]C[-2]>R9C4),R[-1]C-((R[-6]C[-2]-R[-4]C[-2]-R9C4)*0.1),R[-1]C)))"
    ActiveSheet.Calculate
    
        If Range("U7") > Range("F7") Then
            Range("A1:E5").Value = Range("P1:T5").Value
            Range("F2:F7").Value = Range("U2:U7").Value
        End If
    Range("P1:U8").Clear
End Sub

Sub Siba()
'
' Siba Macro
'
'Longest ORDS Leg = TP2 & TP3; JLR added 3/20/18

    Range("G2").FormulaR1C1 = "=MIN(RC[-1]:R[4]C[-1])"
    Range("H2").FormulaR1C1 = "=R7C6*.1"
    ActiveSheet.Calculate
    If Range("F2") = Range("G2") And Range("F2") < Range("H2") Then
        Range("P1:U2").Value = Range("A2:F3").Value
        Range("U1").Clear
    
    'CK TP3 & FIN within 10 mins of ORDS 2 & 3
    Range("H5").FormulaR1C1 = "=ABS(R[-1]C[-7]-TASKS!R[36]C[-7])"
    Range("H6").FormulaR1C1 = "=ABS(R[-1]C[-7]-TASKS!R[36]C[-7])"
    Range("J6").Value = 6.94444444444444E-03
    ActiveSheet.Calculate
    
    If Range("H5") < Range("J6") And Range("H6") < Range("J6") Then
        Sheets("Sheet2").Range("P4:T5").Value = Sheets("Tasks").Range("A41:E42").Value
        Range("U5").FormulaR1C1 = "=6371*ACOS(SIN(R[-1]C[-4])*SIN(RC[-4])+COS(R[-1]C[-4])*COS(RC[-4])*COS(RC[-3]-R[-1]C[-3]))"
    
        
    Range("G8:L8").FormulaR1C1 = "=MAX(R[2]C:R[10001]C)"
    Range("G10").FormulaR1C1 = _
        "=IF(AND(RC[-6]>R2C16,RC[-6]<R4C16),6371*ACOS(SIN(RC[-5])*SIN(R2C17)+COS(RC[-5])*COS(R2C17)*COS(R2C18-RC[-4]))+6371*ACOS(SIN(RC[-5])*SIN(R4C17)+COS(RC[-5])*COS(R4C17)*COS(R4C18-RC[-4])),"""")"
    Range("H10").FormulaR1C1 = "=IF(RC[-1]=R8C7,RC[-7],"""")"
    Range("I10:L10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-7],"""")"
    'Copy Ref A
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:L10").AutoFill Destination:=.Range("G10:L" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    
    Range("P3:T3").Value = Range("H8:L8").Value

Range("U2:U5").FormulaR1C1 = "=6371*ACOS(SIN(R[-1]C[-4])*SIN(RC[-4])+COS(R[-1]C[-4])*COS(RC[-4])*COS(RC[-3]-R[-1]C[-3]))"
    
    Range("U6").FormulaR1C1 = "=SUM(R[-4]C:R[-1]C)"
    Range("U7").FormulaR1C1 = "=IF(R[-6]C[-2]-R[-2]C[-2]<=R[2]C[-17],R[-1]C,R[-1]C-((R[-6]C[-2]-R[-2]C[-2]-R[2]C[-17])*0.1))"
    Columns("G:L").Clear
    
    If Range("U7") > Range("F7") Then
    Range("A1:F7").Value = Range("P1: U7").Value
    End If
  End If
 End If
    
    Range("P1:U7").Clear
    
End Sub

Sub AllOrds()
'
' Works for Boettger - no ref to StrDist or O&R
'
Application.ScreenUpdating = False
    Range("B7:C7").FormulaR1C1 = "=MIN(R[3]C:R[10002]C)"
    Range("A8:C8").FormulaR1C1 = "=MAX(R[2]C:R[10001]C)"
   
    Range("G10").FormulaR1C1 = _
        "=IF(OR(RC[-5]=R7C[-5],RC[-5]=R8C[-5],RC[-4]=R7C[-4],RC[-4]=R8C[-4],RC[-6]=R8C[-6]),RC[-6],"""")"
    Range("H10:K10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-6],"""")"
    
 With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:K10").AutoFill Destination:=.Range("G10:K" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    
    Range("G10:K10009").Value = Range("G10:K10009").Value
    Range("G10:K10009").Sort Key1:=Range("G10"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    
    Range("P1:T12").Value = Range("G10:L21").Value
    Range("U2:U5").FormulaR1C1 = _
        "=IF(SUM(RC[-5]:RC[-1])<>0,6371*ACOS(SIN(R[-1]C[-4])*SIN(RC[-4])+COS(R[-1]C[-4])*COS(RC[-4])*COS(RC[-3]-R[-1]C[-3])),0)"
    ActiveSheet.Calculate
    Range("A7:C8,G10:K22").Clear

    'Ck dist before Ord 1
  If Range("P1") > Range("A10") Then
    Range("G8:L8").FormulaR1C1 = "=MAX(R[2]C:R[10001]C)"
    
    Range("G10").FormulaR1C1 = _
        "=IF(RC[-6]<R1C[9],6371*ACOS(SIN(RC[-5])*SIN(R1C[10])+COS(RC[-5])*COS(R1C[10])*COS(R1C[11]-RC[-4])),"""")"
    Range("H10").FormulaR1C1 = "=IF(RC[-1]=R8C[-1],RC[-7],"""")"
    Range("I10:L10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-7],"""")"
    
 With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:L10").AutoFill Destination:=.Range("G10:L" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    
    If Range("G8") > Range("U2") Then
        Range("P13:T13").Value = Range("H8:L8").Value
        
    'Ck dist before above
    Range("A7").FormulaR1C1 = "=R[1]C[7]-R[3]C"
    ActiveSheet.Calculate
    Range("A7").Value = Range("A7").Value
    If Range("A7") >= 1 / 48 Then
        Range("G10").FormulaR1C1 = _
            "=IF(RC[-6]<R13C[9],6371*ACOS(SIN(RC[-5])*SIN(R13C[10])+COS(RC[-5])*COS(R13C[10])*COS(R13C[11]-RC[-4])),"""")"
    
 With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:G10").AutoFill Destination:=.Range("G10:G" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    
        Range("P14:T14").Value = Range("H8:L8").Value
    End If
   End If
  End If
    'Ck dist before landing, if last ord >= 1/48 earlier
    Range("A7").FormulaR1C1 = "=MAX(R[3]C:R[10002]C)-R[-2]C[15]"
    ActiveSheet.Calculate
    Range("A7").Value = Range("A7").Value
    
    If Range("A7") >= 1 / 48 Then
        Range("G10").FormulaR1C1 = _
            "=IF(RC[-6]>R5C[9],6371*ACOS(SIN(RC[-5])*SIN(R5C[10])+COS(RC[-5])*COS(R5C[10])*COS(R5C[11]-RC[-4])),"""")"

 With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:G10").AutoFill Destination:=.Range("G10:G" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    
    Range("P15:T15").Value = Range("H8:L8").Value
    End If
    
    Range("P1:T15").Sort Key1:=Range("P1"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    
    Range("U2").FormulaR1C1 = _
        "=6371*ACOS(SIN(R[-1]C[-4])*SIN(RC[-4])+COS(R[-1]C[-4])*COS(RC[-4])*COS(RC[-3]-R[-1]C[-3]))"
    'Copy ref P
    Application.Calculation = xlCalculationManual
 With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "P").End(xlUp).Row
.Range("U2:U2").AutoFill Destination:=.Range("U2:U" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    
    Range("U2:U15").Value = Range("U2:U15").Value
    
    Range("T16").FormulaR1C1 = "=LARGE(R[-15]C[1]:R[-1]C[1],4)"
    Range("V1").FormulaR1C1 = "=IF(R[1]C[-1]>=R16C[-2],RC[-6],"""")"
    Range("V2:V15").FormulaR1C1 = "=IF(RC[-1]>=R16C[-2],RC[-6],"""")"
    ActiveSheet.Calculate
    Range("V1:V15").Value = Range("V1:V15").Value
    
    Range("P1:V15").Sort Key1:=Range("V1"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    
    Range("P6:V16").Clear
    Range("V1:V6").Clear
    
    Range("P1:T5").Sort Key1:=Range("P1"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    
    Range("U2:U5").FormulaR1C1 = _
        "=6371*ACOS(SIN(R[-1]C[-4])*SIN(RC[-4])+COS(R[-1]C[-4])*COS(RC[-4])*COS(RC[-3]-R[-1]C[-3]))"
    ActiveSheet.Calculate
    'Test for later Fini
    Range("G10").FormulaR1C1 = _
        "=IF(RC[-6]>R4C[9],6371*ACOS(SIN(RC[-5])*SIN(R4C[10])+COS(RC[-5])*COS(R4C[10])*COS(R4C[11]-RC[-4])),"""")"

 With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:G10").AutoFill Destination:=.Range("G10:G" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    
    If Range("G8") > Range("U5") Then
        Range("P5:T5").Value = Range("H8:L8").Value
    End If
    
    'Range("A1:F5").Value = Range("P1:U5").Value   16DEC
    Range("U6").FormulaR1C1 = "=SUM(R[-4]C:R[-1]C)"
    Range("U7").FormulaR1C1 = _
        "=IF(R[-6]C[-2]-R[-2]C[-2]>R9C4,R[-1]C-((R[-6]C[-2]-R[-2]C[-2]-R9C4)*0.1),R[-1]C)"
    ActiveSheet.Calculate
    Range("U6:U7").Value = Range("U6:U7").Value
    
    If Range("U7") > Range("F7") Then
        Range("A1:F7").Value = Range("P1:U7").Value
    
    Range("A8").FormulaR1C1 = "=MAX(R[2]C:R[10001]C)-R[-3]C[15]"
    ActiveSheet.Calculate
    If Range("A8") > 1 / 48 Then
        Application.Run "F.xlsm!AllOrds2"
        If Range("U7") > Range("F7") Then
        Range("A1:F7").Value = Range("P1:U7").Value
        End If
    End If
   End If
    Range("A7:C8").Clear
    Columns("G:U").Clear
    
    Application.Run "F.xlsm!ORBase"

End Sub

Sub TriZords()
'
' Pure ORDS assumes release & landing are NOT ORDs
'
Application.ScreenUpdating = False
    Range("A8").FormulaR1C1 = "=MAX(R[2]C:R[10001]C)"
    Range("A8").Value = Range("A8").Value
    Sheets("Sheet2").Range("P10:T13").Value = Sheets("TASKS").Range("A40:E43").Value
    
    Range("I4").FormulaR1C1 = "=ACOS(SIN(R2C)*SIN(R3C)+COS(R2C)*COS(R3C)*COS(R3C[1]-R2C[1]))"
    Range("K4").FormulaR1C1 = "=MAX(R[6]C[-2]:R[10005]C[-2])"
    Range("L4").FormulaR1C1 = "=MAX(R[6]C[-2]:R[10005]C[-2])"
    Range("I5").FormulaR1C1 = "=ACOS((SIN(R3C)-SIN(R2C)*COS(R[-1]C))/(SIN(R[-1]C)*COS(R2C)))"
    Range("J8:N8").FormulaR1C1 = "=MAX(R[2]C:R[10001]C)"
    
    Range("G10").FormulaR1C1 = "=IF(RC[-4]=R2C[3],"""",IF(AND(RC1>R2C[1],RC1<R3C[1]),ACOS(SIN(R2C[2])*SIN(RC2)+COS(R2C[2])*COS(RC2)*COS(RC3-R2C[3])),""""))"
    Range("H10").FormulaR1C1 = "=IF(RC[-1]<>"""",ACOS((SIN(RC2)-SIN(R2C[1])*COS(RC[-1]))/(SIN(RC[-1])*COS(R2C[1]))),"""")"
    Range("I10").FormulaR1C1 = "=IF(RC[-1]<>"""",ABS((ASIN(SIN(RC[-2])*SIN(RC[-1]-R5C))*180*60*1.852/PI())),"""")"
    Range("J10").FormulaR1C1 = "=IF(RC[-1]=R4C[1],RC1,"""")"
    Range("K10:N10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-9],"""")"
    'Copy Ref A
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:N10").AutoFill Destination:=.Range("G10:N" & LastRow), Type:=xlFillDefault
    End With
ActiveSheet.Calculate
    
    Range("H2:L3").Value = Range("P10:T11").Value
    Range("P14:T14").Value = Range("J8:N8").Value
    
    Range("H2:L3").Value = Range("P11:T12").Value
    Range("P15:T15").Value = Range("J8:N8").Value
    
    Range("H2:L3").Value = Range("P12:T13").Value
    Range("P16:T16").Value = Range("J8:N8").Value
    
    Columns("G:N").Clear
    
    Range("P10:T16").Sort Key1:=Range("P10"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    
    If Range("P1") <> Range("A10") And Range("P16") <> Range("A8") Then
    
    Range("P10:T16").Copy
    Range("H1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    
    Range("A10:E10").Copy
    Range("G1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    Application.CutCopyMode = False
    
    Range("G10:N16").FormulaR1C1 = "=IF(R3C=RC18,0,6371*ACOS(SIN(R2C)*SIN(RC17)+COS(R2C)*COS(RC17)*COS(RC18-R3C)))"
    ActiveSheet.Calculate
    Range("G10:N16").Value = Range("G10:N16").Value
    Range("U19").FormulaR1C1 = "=R[-8]C[-13]+R[-7]C[-12]"
    Range("V19").Value = "123"
    Range("W19").FormulaR1C1 = "=RC[-2]+R[-9]C[-16]"
    Range("U20").FormulaR1C1 = "=R[-8]C[-12]+R[-7]C[-11]"
    Range("V20").Value = "234"
    Range("W20").FormulaR1C1 = "=RC[-2]+R[-9]C[-16]"
    Range("U21").FormulaR1C1 = "=R[-8]C[-11]+R[-7]C[-10]"
    Range("V21").Value = "345"
    Range("W21").FormulaR1C1 = "=RC[-2]+R[-9]C[-16]"
    Range("U22").FormulaR1C1 = "=R[-8]C[-10]+R[-7]C[-9]"
    Range("V22").Value = "456"
    Range("W22").FormulaR1C1 = "=RC[-2]+R[-9]C[-16]"
    Range("U23").FormulaR1C1 = "=R[-8]C[-9]+R[-7]C[-8]"
    Range("V23").Value = "567"
    Range("W23").FormulaR1C1 = "=RC[-2]+R[-9]C[-16]"
    Range("U24").FormulaR1C1 = "=R[-12]C[-13]+R[-11]C[-11]"
    Range("V24").Value = "134"
    Range("W24").FormulaR1C1 = "=RC[-2]+R[-14]C[-16]"
    Range("U25").FormulaR1C1 = "=R[-13]C[-13]+R[-11]C[-11]"
    Range("V25").Value = "135"
    Range("W25").FormulaR1C1 = "=RC[-2]+R[-15]C[-16]"
    Range("U26").FormulaR1C1 = "=R[-14]C[-13]+R[-11]C[-11]"
    Range("V26").Value = "136"
    Range("W26").FormulaR1C1 = "=RC[-2]+R[-16]C[-16]"
    Range("U27").FormulaR1C1 = "=R[-15]C[-13]+R[-11]C[-11]"
    Range("V27").Value = "137"
    Range("W27").FormulaR1C1 = "=RC[-2]+R[-17]C[-16]"
    Range("U28").FormulaR1C1 = "=R[-17]C[-13]+R[-15]C[-12]"
    Range("V28").Value = "124"
    Range("W28").FormulaR1C1 = "=RC[-2]+R[-18]C[-16]"
    Range("U29").FormulaR1C1 = "=R[-18]C[-13]+R[-15]C[-12]"
    Range("V29").Value = "125"
    Range("W29").FormulaR1C1 = "=RC[-2]+R[-19]C[-16]"
    Range("U30").FormulaR1C1 = "=R[-19]C[-13]+R[-15]C[-12]"
    Range("V30").Value = "126"
    Range("W30").FormulaR1C1 = "=RC[-2]+R[-20]C[-16]"
    Range("U31").FormulaR1C1 = "=R[-20]C[-13]+R[-15]C[-12]"
    Range("V31").Value = "127"
    Range("W31").FormulaR1C1 = "=RC[-2]+R[-21]C[-16]"
    Range("U32").FormulaR1C1 = "=R[-20]C[-12]+R[-18]C[-11]"
    Range("V32").Value = "235"
    Range("W32").FormulaR1C1 = "=RC[-2]+R[-21]C[-16]"
    Range("U33").FormulaR1C1 = "=R[-21]C[-12]+R[-18]C[-11]"
    Range("V33").Value = "236"
    Range("W33").FormulaR1C1 = "=RC[-2]+R[-22]C[-16]"
    Range("U34").FormulaR1C1 = "=R[-22]C[-12]+R[-18]C[-11]"
    Range("V34").Value = "237"
    Range("W34").FormulaR1C1 = "=RC[-2]+R[-23]C[-16]"
    Range("U35").FormulaR1C1 = "=R[-22]C[-12]+R[-21]C[-10]"
    Range("V35").Value = "245"
    Range("W35").FormulaR1C1 = "=RC[-2]+R[-24]C[-16]"
    Range("U36").FormulaR1C1 = "=R[-23]C[-12]+R[-21]C[-10]"
    Range("V36").Value = "246"
    Range("W36").FormulaR1C1 = "=RC[-2]+R[-25]C[-16]"
    Range("U37").FormulaR1C1 = "=R[-24]C[-12]+R[-21]C[-10]"
    Range("V37").Value = "247"
    Range("W37").FormulaR1C1 = "=RC[-2]+R[-26]C[-16]"
    Range("U38").FormulaR1C1 = "=R[-24]C[-12]+R[-23]C[-9]"
    Range("V38").Value = "256"
    Range("W38").FormulaR1C1 = "=RC[-2]+R[-27]C[-16]"
    Range("U39").FormulaR1C1 = "=R[-25]C[-12]+R[-23]C[-9]"
    Range("V39").Value = "257"
    Range("W39").FormulaR1C1 = "=RC[-2]+R[-28]C[-16]"
    Range("U40").FormulaR1C1 = "=R[-27]C[-11]+R[-25]C[-10]"
    Range("V40").Value = "346"
    Range("W40").FormulaR1C1 = "=RC[-2]+R[-28]C[-16]"
    Range("U41").FormulaR1C1 = "=R[-28]C[-11]+R[-25]C[-10]"
    Range("V41").Value = "347"
    Range("W41").FormulaR1C1 = "=RC[-2]+R[-29]C[-16]"
    Range("U42").FormulaR1C1 = "=R[-29]C[-13]+R[-28]C[-10]"
    Range("V42").Value = "145"
    Range("W42").FormulaR1C1 = "=RC[-2]+R[-32]C[-16]"
    Range("U43").FormulaR1C1 = "=R[-30]C[-13]+R[-28]C[-10]"
    Range("V43").Value = "146"
    Range("W43").FormulaR1C1 = "=RC[-2]+R[-33]C[-16]"
    Range("U44").FormulaR1C1 = "=R[-31]C[-13]+R[-28]C[-10]"
    Range("V44").Value = "147"
    Range("W44").FormulaR1C1 = "=RC[-2]+R[-34]C[-16]"
    Range("U45").FormulaR1C1 = "=R[-31]C[-13]+R[-30]C[-9]"
    Range("V45").Value = "156"
    Range("W45").FormulaR1C1 = "=RC[-2]+R[-35]C[-16]"
    Range("U46").FormulaR1C1 = "=R[-32]C[-13]+R[-30]C[-9]"
    Range("V46").Value = "157"
    Range("W46").FormulaR1C1 = "=RC[-2]+R[-36]C[-16]"
    Range("U47").FormulaR1C1 = "=R[-32]C[-13]+R[-31]C[-8]"
    Range("V47").Value = "167"
    Range("W47").FormulaR1C1 = "=RC[-2]+R[-37]C[-16]"
    ActiveSheet.Calculate
    Range("U19:W47").Value = Range("U19:W47").Value
    Range("U19:W47").Sort Key1:=Range("W19"), Order1:=xlDescending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    Range("V1").Value = Range("V19").Value
    Range("U19:W47").Clear
    
    Range("V1").TextToColumns Destination:=Range("V1"), DataType:=xlFixedWidth, FieldInfo:=Array(Array(0, 1), Array(1, 1), Array(2, 1)), TrailingMinusNumbers:=True
    
    Range("G1:G5").Copy
    Range("P2").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    
    Range("P3").FormulaR1C1 = "=IF(R1C[6]=1,R[-2]C[-8],IF(R1C[6]=2,R[-2]C[-7],IF(R1C[6]=3,R[-2]C[-6],IF(R1C[6]=4,R[-2]C[-5],IF(R1C[6]=5,R[-2]C[-4],IF(R1C[6]=6,R[-2]C[-3],R[-2]C[-2]))))))"
    Range("Q3").FormulaR1C1 = "=IF(R1C22=1,R[-1]C8,IF(R1C22=2,R[-1]C9,IF(R1C22=3,R[-1]C10,IF(R1C22=4,R[-1]C11,IF(R1C22=5,R[-1]C12,IF(R1C22=6,R[-1]C13,R[-1]C14))))))"
    Range("R3").FormulaR1C1 = "=IF(R1C22=1,RC8,IF(R1C22=2,RC9,IF(R1C22=3,RC10,IF(R1C22=4,RC11,IF(R1C22=5,RC12,IF(R1C22=6,RC13,RC14))))))"
    Range("S3").FormulaR1C1 = "=IF(R1C22=1,R[1]C8,IF(R1C22=2,R[1]C9,IF(R1C22=3,R[1]C10,IF(R1C22=4,R[1]C11,IF(R1C22=5,R[1]C12,IF(R1C22=6,R[1]C13,R[1]C14))))))"
    Range("T3").FormulaR1C1 = "=IF(R1C22=1,R[2]C8,IF(R1C22=2,R[2]C9,IF(R1C22=3,R[2]C10,IF(R1C22=4,R[2]C11,IF(R1C22=5,R[2]C12,IF(R1C22=6,R[2]C13,R[2]C14))))))"
    Range("P4").FormulaR1C1 = "=IF(R1C[7]=2,R[-3]C[-7],IF(R1C[7]=3,R[-3]C[-6],IF(R1C[7]=4,R[-3]C[-5],IF(R1C[7]=5,R[-3]C[-4],IF(R1C[7]=6,R[-3]C[-3],R[-3]C[-2])))))"
    Range("Q4").FormulaR1C1 = "=IF(R1C[6]=2,R[-2]C[-8],IF(R1C[6]=3,R[-2]C[-7],IF(R1C[6]=4,R[-2]C[-6],IF(R1C[6]=5,R[-2]C[-5],IF(R1C[6]=6,R[-2]C[-4],R[-2]C[-3])))))"
    Range("R4").FormulaR1C1 = "=IF(R1C[5]=2,R[-1]C[-9],IF(R1C[5]=3,R[-1]C[-8],IF(R1C[5]=4,R[-1]C[-7],IF(R1C[5]=5,R[-1]C[-6],IF(R1C[5]=6,R[-1]C[-5],R[-1]C[-4])))))"
    Range("S4").FormulaR1C1 = "=IF(R1C[4]=2,RC[-10],IF(R1C[4]=3,RC[-9],IF(R1C[4]=4,RC[-8],IF(R1C[4]=5,RC[-7],IF(R1C[4]=6,RC[-6],RC[-5])))))"
    Range("T4").FormulaR1C1 = "=IF(R1C[3]=2,R[1]C[-11],IF(R1C[3]=3,R[1]C[-10],IF(R1C[3]=4,R[1]C[-9],IF(R1C[3]=5,R[1]C[-8],IF(R1C[3]=6,R[1]C[-7],R[1]C[-6])))))"
    Range("P5").FormulaR1C1 = "=IF(R1C24=3,R[-4]C10,IF(R1C24=4,R[-4]C11,IF(R1C24=5,R[-4]C12,IF(R1C24=6,R[-4]C13,R[-4]C14))))"
    Range("Q5").FormulaR1C1 = "=IF(R1C24=3,R[-3]C10,IF(R1C24=4,R[-3]C11,IF(R1C24=5,R[-3]C12,IF(R1C24=6,R[-3]C13,R[-3]C14))))"
    Range("R5").FormulaR1C1 = "=IF(R1C24=3,R[-2]C10,IF(R1C24=4,R[-2]C11,IF(R1C24=5,R[-2]C12,IF(R1C24=6,R[-2]C13,R[-2]C14))))"
    Range("S5").FormulaR1C1 = "=IF(R1C24=3,R[-1]C10,IF(R1C24=4,R[-1]C11,IF(R1C24=5,R[-1]C12,IF(R1C24=6,R[-1]C13,R[-1]C14))))"
    Range("T5").FormulaR1C1 = "=IF(R1C24=3,RC10,IF(R1C24=4,RC11,IF(R1C24=5,RC12,IF(R1C24=6,RC13,RC14))))"
    ActiveSheet.Calculate
    Range("P2:T5").Value = Range("P2:T5").Value
    
    Columns("G:N").Clear
    
    Range("G8:L8").FormulaR1C1 = "=MAX(R[2]C:R[10001]C)"
    Range("G10").FormulaR1C1 = "=IF(RC[-6]>R5C16,6371*ACOS(SIN(RC[-5])*SIN(R5C17)+COS(RC[-5])*COS(R5C17)*COS(R5C18-RC[-4])),"""")"
    Range("H10").FormulaR1C1 = "=IF(RC[-1]=R8C7,RC[-7],"""")"
    Range("I10:L10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-7],"""")"
    'Copy Ref A
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:L10").AutoFill Destination:=.Range("G10:L" & LastRow), Type:=xlFillDefault
    End With
ActiveSheet.Calculate
    
    Range("P6:T6").Value = Range("H8:L8").Value
    
    Range("U3:U6").FormulaR1C1 = "=IF(AND(R[-1]C[-5]<>0, RC[-5]<>0),6371*ACOS(SIN(R[-1]C[-4])*SIN(RC[-4])+COS(R[-1]C[-4])*COS(RC[-4])*COS(RC[-3]-R[-1]C[-3])),0)"
    Range("U7").FormulaR1C1 = "=SUM(R[-4]C:R[-1]C)"
    Range("U8").FormulaR1C1 = "=IF(R[-6]C[-2]-R[-2]C[-2]<=R9C4,R[-1]C,R[-1]C-((R[-6]C[-2]-R[-2]C[-2]-R9C4)*0.1))"
    ActiveSheet.Calculate
    Range("U3:U8").Value = Range("U3:U8").Value
    Range("V1:X1,P10:T16").Clear
    Columns("G:L").Clear
    
    If Range("U8") > Range("F7") Then
        Range("A1:F7").Value = Range("P2:U8").Value
        Range("P2:U8").Clear
    End If
  End If
  
End Sub
Sub TP3b()
'
' For Jonker (TP3 = max off course TP2 / St Dist Fini) Revised 7/17/2018 remove conflicting longitudes
'
Application.ScreenUpdating = False
    Range("P15:T16").Value = Range("P1:T2").Value
  
   'MAX off course before ST DIST Fini
    Range("G10").FormulaR1C1 = _
        "=IF(AND(RC1<R3C[1],RC3<>R2C[3]),ACOS(SIN(R2C[2])*SIN(RC2)+COS(R2C[2])*COS(RC2)*COS(RC3-R2C[3])),"""")"
    Range("H10").FormulaR1C1 = _
        "=IF(RC[-1]<>"""",ACOS((SIN(RC2)-SIN(R2C[1])*COS(RC[-1]))/(SIN(RC[-1])*COS(R2C[1]))),"""")"
    Range("I10").FormulaR1C1 = _
        "=IF(RC[-2]<>"""",ABS((ASIN(SIN(RC[-2])*SIN(RC[-1]-R5C))*180*60*1.852/PI())),"""")"
    Range("J10").FormulaR1C1 = "=IF(RC[-1]=R4C[1],RC1,"""")"
    Range("K10:N10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-9],"""")"
    
    Range("K4").FormulaR1C1 = "=MAX(R[6]C[-2]:R[10005]C[-2])"
    Range("L4").FormulaR1C1 = "=MAX(R[6]C[-2]:R[10005]C[-2])"
    
 With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:N10").AutoFill Destination:=.Range("G10:N" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    
    Range("J10:N10009").Value = Range("J10:N10009").Value
    Range("G10:I10009").Clear
    Range("J10:N10009").Sort Key1:=Range("J10"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    
    Range("P17:T17").Value = Range("J10:N10").Value
    Range("J10:N10").Clear
    'Off course from previous to ST DIST FINI
    Range("H2:L2").Value = Range("P17:T17").Value
    Range("G10").FormulaR1C1 = _
        "=IF(AND(RC1>R2C[1],RC[-6]<R3C8,RC3<>R2C10),ACOS(SIN(R2C[2])*SIN(RC2)+COS(R2C[2])*COS(RC2)*COS(RC3-R2C[3])),"""")"
    Range("H10").FormulaR1C1 = _
        "=IF(RC[-1]<>"""",ACOS((SIN(RC2)-SIN(R2C[1])*COS(RC[-1]))/(SIN(RC[-1])*COS(R2C[1]))),"""")"
    Range("I10").FormulaR1C1 = _
        "=IF(RC[-2]<>"""",ABS((ASIN(SIN(RC[-2])*SIN(RC[-1]-R5C))*180*60*1.852/PI())),"""")"
    Range("J10").FormulaR1C1 = "=IF(RC[-1]=R4C[1],RC1,"""")"
    Range("K10:N10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-9],"""")"
    
 With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:N10").AutoFill Destination:=.Range("G10:N" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    
    Range("J8:N8").FormulaR1C1 = "=MAX(R[2]C:R[10001]C)"
    
    Range("P18:T18").Value = Range("J8:N8").Value
    Range("P19:T19").Value = Range("H3:L3").Value
    
    Range("P15:T19").Sort Key1:=Range("P15"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    
    Range("O16:O19").FormulaR1C1 = _
        "=IF(OR(R[-1]C[1]=0,RC[1]=0),0,6371*ACOS(SIN(R[-1]C[2])*SIN(RC[2])+COS(R[-1]C[2])*COS(RC[2])*COS(R[-1]C[3]-RC[3])))"
    Range("O20").FormulaR1C1 = "=SUM(R[-4]C:R[-1]C)"
    Range("P20").FormulaR1C1 = _
        "=IF(R[-5]C[3]-R[-1]C[3]>R9C4,RC[-1]-((R[-5]C[3]-R[-1]C[3]-R9C4)*0.1),RC[-1])"
    ActiveSheet.Calculate
    Range("P18:T20").Value = Range("P18:T20").Value
    Range("M8:N10009").Clear
 
 If Range("P20") < Range("O20") Then
    Range("A7").Value = 2.08333333333333E-02
    Range("G8:L8").FormulaR1C1 = "=MAX(R[2]C:R[10001]C)"
    
    Range("G10").FormulaR1C1 = _
        "=IF(AND(RC[-6]>=R15C16-R7C1,RC[-6]<=R15C16+R7C1,RC[-3]-R19C19<=R9C4),6371*ACOS(SIN(RC[-5])*SIN(R16C17)+COS(RC[-5])*COS(R16C17)*COS(R16C18-RC[-4])),"""")"
    Range("H10").FormulaR1C1 = "=IF(RC[-1]=R8C7,RC[-7],"""")"
    Range("I10:L10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-7],"""")"
    
 With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:L10").AutoFill Destination:=.Range("G10:L" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    
    Range("E7").FormulaR1C1 = "=R16C15-R8C7"
    Range("E8").FormulaR1C1 = "=IF(R[-1]C>0,R20C15-R20C16,0)"
    ActiveSheet.Calculate
    Range("E7:E8").Value = Range("E7:E8").Value
    
    If Range("E7") < Range("E8") And Range("G8") <> 0 Then

    Range("P22:T22").Value = Range("H8:L8").Value
    Range("P23:T26").Value = Range("P16:T19").Value
    
    Range("O23:O26").FormulaR1C1 = _
        "=IF(OR(R[-1]C[1]=0,RC[1]=0),0,6371*ACOS(SIN(R[-1]C[2])*SIN(RC[2])+COS(R[-1]C[2])*COS(RC[2])*COS(RC[3]-R[-1]C[3])))"
    Range("O27").FormulaR1C1 = "=SUM(R[-4]C:R[-1]C)"
    ActiveSheet.Calculate
    Range("O23:O27").Value = Range("O23:O27").Value
    
    Range("P27").FormulaR1C1 = _
        "=IF(R[-5]C[3]-R[-1]C[3]>R9C4,RC[1]-((R[-5]C[3]-R[-1]C[3]-R9C4)*0.1),RC[-1])"
    ActiveSheet.Calculate
    Range("P27").Value = Range("P27").Value
    End If
  End If
   
'For Vihlen eg OFC Ord 2/3
    Range("H2:L3").Value = Range("P2:T3").Value
    Range("G10").FormulaR1C1 = _
        "=IF(RC3=R2C[3],"""",IF(AND(RC1>R2C8,RC1<R3C8),ACOS(SIN(R2C[2])*SIN(RC2)+COS(R2C[2])*COS(RC2)*COS(RC3-R2C[3])),""""))"
    Range("H10").FormulaR1C1 = "=IF(RC[-1]<>"""",ACOS((SIN(RC2)-SIN(R2C[1])*COS(RC[-1]))/(SIN(RC[-1])*COS(R2C[1]))),"""")"
    Range("I10").FormulaR1C1 = "=IF(RC[-1]<>"""",ABS((ASIN(SIN(RC[-2])*SIN(RC[-1]-R5C))*180*60*1.852/PI())),"""")"
    Range("J10").FormulaR1C1 = "=IF(RC[-1]=R4C[1],RC1,"""")"
    Range("K10:N10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-9],"""")"
    
    Range("K4").FormulaR1C1 = "=MAX(R[6]C[-2]:R[10005]C[-2])"
    Range("L4").FormulaR1C1 = "=MAX(R[6]C[-2]:R[10005]C[-2])"
    
 With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:N10").AutoFill Destination:=.Range("G10:N" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    
    Range("J8:N8").FormulaR1C1 = "=MAX(R[2]C:R[10001]C)"
    
    Range("P29:T30").Value = Range("P1:T2").Value
    Range("P31:T31").Value = Range("J8:N8").Value
    Range("P32:T32").Value = Range("H3:L3").Value
    
    Range("G10:N10009").Clear
    Range("G8:L8").FormulaR1C1 = "=MAX(R[2]C:R[10001]C)"
    
    Range("G10").FormulaR1C1 = "=IF(RC[-6]>R32C[9],6371*ACOS(SIN(RC[-5])*SIN(R32C17)+COS(RC[-5])*COS(R32C17)*COS(R32C18-RC[-4])),"""")"
    Range("H10").FormulaR1C1 = "=IF(RC[-1]=R8C7,RC[-7],"""")"
    Range("I10:L10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-7],"""")"
    'Copy Ref A
 With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:L10").AutoFill Destination:=.Range("G10:L" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    
    Range("P33:T33").Value = Range("H8:L8").Value
    
    Range("O30:O33").FormulaR1C1 = "=IF(AND(RC[1]<>"""",R[-1]C[1]<>0,RC[1]<>0),6371*ACOS(SIN(R[-1]C[2])*SIN(RC[2])+COS(R[-1]C[2])*COS(RC[2])*COS(RC[3]-R[-1]C[3])),0)"
    Range("O34").FormulaR1C1 = "=SUM(R[-4]C:R[-1]C)"
    Range("P34").FormulaR1C1 = "=IF(R[-5]C[3]-R[-1]C[3]<=R9C4,RC[-1],RC[-1]-((R[-5]C[3]-R[-1]C[3]-R9C4)*0.1))"
    ActiveSheet.Calculate
    Range("O30:P34").Value = Range("O30:P34").Value
    
    Range("G8:N10009").Clear
    
    If Range("P34") < Range("O34") Then
    Range("A7").Value = 2.08333333333333E-02
    Range("G8:L8").FormulaR1C1 = "=MAX(R[2]C:R[10001]C)"
    
    Range("G10").FormulaR1C1 = _
        "=IF(AND(RC[-6]>=R29C16-R7C1,RC[-6]<=R29C16+R7C1,RC[-3]-R33C19<=R9C4),6371*ACOS(SIN(RC[-5])*SIN(R30C17)+COS(RC[-5])*COS(R30C17)*COS(R30C18-RC[-4])),"""")"
    Range("H10").FormulaR1C1 = "=IF(RC[-1]=R8C7,RC[-7],"""")"
    Range("I10:L10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-7],"""")"
    
 With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:L10").AutoFill Destination:=.Range("G10:L" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    
    Range("E7").FormulaR1C1 = "=R30C15-R8C7"
    Range("E8").FormulaR1C1 = "=IF(R[-1]C>0,R34C15-R34C16,0)"
    ActiveSheet.Calculate
    Range("E7:E8").Value = Range("E7:E8").Value
    
    If Range("E7") < Range("E8") And Range("G8") <> 0 Then

    Range("P36:T36").Value = Range("H8:L8").Value
    Range("P37:T40").Value = Range("P30:T34").Value
    
    Range("O37:O40").FormulaR1C1 = _
        "=IF(AND(RC[1]<>"""",R[-1]C[1]<>0,RC[1]<>0),6371*ACOS(SIN(R[-1]C[2])*SIN(RC[2])+COS(R[-1]C[2])*COS(RC[2])*COS(RC[3]-R[-1]C[3])),0)"
    Range("O41").FormulaR1C1 = "=SUM(R[-4]C:R[-1]C)"
    ActiveSheet.Calculate
    Range("O37:O41").Value = Range("O37:O41").Value
    
    Range("P41").FormulaR1C1 = "=IF(R[-5]C[3]-R[-1]C[3]>R9C4,RC[1]-((R[-5]C[3]-R[-1]C[3]-R9C4)*0.1),RC[-1])"
    ActiveSheet.Calculate
    Range("P41").Value = Range("P41").Value
    End If
  End If
  
End Sub

Sub CK3TP()
'
' 7/27/17 JLR Consolidated w/ Refine3TP for brevity; 8/24/17 corrected "Ref" to "REF" for proper checking
'
Application.ScreenUpdating = False
    Range("H12:L16").Value = Range("A1:E5").Value
    Range("G10").FormulaR1C1 = "=SUM(R13C:R16C)"
    Range("G11").FormulaR1C1 = "=IF(R12C[4]-R16C[4]>R9C4,R10C-((R12C[4]-R16C[4]-R9C4)*.1),R10C)"
    Range("G13:G16").FormulaR1C1 = _
        "=6371*ACOS(SIN(RC[2])*SIN(R[-1]C[2])+COS(RC[2])*COS(R[-1]C[2])*COS(R[-1]C[3]-RC[3]))"
    ActiveSheet.Calculate
    'Check TP1
    Range("G19:L19").FormulaR1C1 = "=MAX(R[1]C:R[60000]C)"
    Range("G20").FormulaR1C1 = _
        "=IF(AND(R[-10]C1>R12C[1],R[-10]C1<R14C[1]),6371*ACOS(SIN(R[-10]C2)*SIN(R12C[2])+COS(R[-10]C2)*COS(R12C[2])*COS(R12C[3]-R[-10]C3)),"""")"
    Range("H20").FormulaR1C1 = "=IF(RC[-1]=R19C[-1],R[-10]C1,"""")"
    Range("I20:L20").FormulaR1C1 = "=IF(RC[-1]<>"""",R[-10]C[-7],"""")"
    
 With ActiveSheet
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G20:L20").AutoFill Destination:=.Range("G20:L" & LastRow), Type:=xlFillDefault
    End With
 ActiveSheet.Calculate

    Range("G19:L19").Value = Range("G19:L19").Value
    Range("G20:L60019").Clear
    
    Range("G20").FormulaR1C1 = _
        "=6371*ACOS(SIN(R[-1]C[2])*SIN(R[-6]C[2])+COS(R[-1]C[2])*COS(R[-6]C[2])*COS(R[-6]C[3]-R[-1]C[3]))"
    Range("G22").FormulaR1C1 = "=R[-3]C+R[-2]C+R[-7]C+R[-6]C"
    ActiveSheet.Calculate
    
    If Range("G22") > Range("G10") Then
        Range("H13:L13").Value = Range("H19:L19").Value
    End If
    
    'Check TP2
    Range("G19:L19").FormulaR1C1 = "=MAX(R[1]C:R[60000]C)"
    
    Range("G20").FormulaR1C1 = _
        "=IF(AND(R[-10]C1>R13C[1],R[-10]C1<R15C[1]),6371*ACOS(SIN(R[-10]C2)*SIN(R13C[2])+COS(R[-10]C2)*COS(R13C[2])*COS(R13C[3]-R[-10]C3)),"""")"
    Range("H20").FormulaR1C1 = "=IF(RC[-1]=R19C[-1],R[-10]C1,"""")"
    Range("I20:L20").FormulaR1C1 = "=IF(RC[-1]<>"""",R[-10]C[-7],"""")"
    
 With ActiveSheet
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G20:L20").AutoFill Destination:=.Range("G20:L" & LastRow), Type:=xlFillDefault
    End With
 ActiveSheet.Calculate
 
    Range("G19:L19").Value = Range("G19:L19").Value
    Range("G20:L60019").Clear
    Range("G20").FormulaR1C1 = _
        "=6371*ACOS(SIN(R[-1]C[2])*SIN(R[-5]C[2])+COS(R[-1]C[2])*COS(R[-5]C[2])*COS(R[-5]C[3]-R[-1]C[3]))"
    Range("G22").FormulaR1C1 = "=R[-3]C+R[-2]C+R[-9]C+R[-6]C"
    ActiveSheet.Calculate
    
    If Range("G22") > Range("G10") Then
        Range("H14:L14").Value = Range("H19:L19").Value
    End If
    
    'CK TP3
    Range("G19:L19").FormulaR1C1 = "=MAX(R[1]C:R[60000]C)"
    Range("G20").FormulaR1C1 = _
        "=IF(AND(R[-10]C1>R14C[1],R[-10]C1<R16C[1]),6371*ACOS(SIN(R[-10]C2)*SIN(R14C[2])+COS(R[-10]C2)*COS(R14C[2])*COS(R14C[3]-R[-10]C3)),"""")"
    Range("H20").FormulaR1C1 = "=IF(RC[-1]=R19C[-1],R[-10]C1,"""")"
    Range("I20:L20").FormulaR1C1 = "=IF(RC[-1]<>"""",R[-10]C[-7],"""")"
    
 With ActiveSheet
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G20:L20").AutoFill Destination:=.Range("G20:L" & LastRow), Type:=xlFillDefault
    End With
 ActiveSheet.Calculate
    
    Range("G19:L19").Value = Range("G19:L19").Value
    Range("G20:L60019").Clear
    Range("G20").FormulaR1C1 = _
        "=6371*ACOS(SIN(R[-6]C[2])*SIN(R[-1]C[2])+COS(R[-6]C[2])*COS(R[-1]C[2])*COS(R[-1]C[3]-R[-6]C[3]))"
    'next 2 lines changed 1/7/13 to consider effect on leg 4 (Jonker)
    Range("G21").FormulaR1C1 = _
        "=6371*ACOS(SIN(R[-5]C[2])*SIN(R[-2]C[2])+COS(R[-5]C[2])*COS(R[-2]C[2])*COS(R[-2]C[3]-R[-5]C[3]))"
    Range("G22").FormulaR1C1 = "=R[-9]C+R[-8]C+R[-2]C+R[-1]C"
    ActiveSheet.Calculate
    
    If Range("G22") > Range("G10") Then
        Range("H15:L15").Value = Range("H19:L19").Value
    End If
   
    'CK Start
    Range("G19:L19").FormulaR1C1 = "=MAX(R[1]C:R[60000]C)"
    Range("G20").FormulaR1C1 = _
        "=IF(R[-10]C1<R13C[1],6371*ACOS(SIN(R[-10]C2)*SIN(R13C[2])+COS(R[-10]C2)*COS(R13C[2])*COS(R13C[3]-R[-10]C3)),"""")"
    Range("H20").FormulaR1C1 = "=IF(RC[-1]=R19C[-1],R[-10]C1,"""")"
    Range("I20:L20").FormulaR1C1 = "=IF(RC[-1]<>"""",R[-10]C[-7],"""")"
    
 With ActiveSheet
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G20:L20").AutoFill Destination:=.Range("G20:L" & LastRow), Type:=xlFillDefault
    End With
 ActiveSheet.Calculate
 
    Range("G19:L19").Value = Range("G19:L19").Value
    Range("G20:L60019").Clear
    
    Range("G22").FormulaR1C1 = "=R[-8]C+R[-7]C+R[-6]C+R[-3]C"
    Range("G23").FormulaR1C1 = "=IF(R[-4]C[4]-R[-7]C[4]>R9C4,R[-1]C-(R[-4]C[4]-R[-7]C[4]-R9C4)*.1,R[-1]C)"
    ActiveSheet.Calculate
    
    If Range("G23") > Range("G11") Then
        Range("H12:L12").Value = Range("H19:L19").Value
    End If
    
    'Reverse CK
    Range("G19:L19").FormulaR1C1 = "=MAX(R[1]C:R[60000]C)"
    Range("M19").FormulaR1C1 = "=SUM(RC[-5]:RC[-1])"
    Range("G20").FormulaR1C1 = _
        "=IF(AND(R[-10]C[-6]>R15C[1],R[-10]C[-6]<R16C[1]),6371*ACOS(SIN(R[-10]C[-5])*SIN(R16C[2])+COS(R[-10]C[-5])*COS(R16C[2])*COS(R16C[3]-R[-10]C[-4])),"""")"
    Range("H20").FormulaR1C1 = "=IF(RC[-1]=R19C[-1],R[-10]C[-7],"""")"
    Range("I20:L20").FormulaR1C1 = "=IF(RC[-1]<>"""",R[-10]C[-7],"""")"
    
 With ActiveSheet
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G20:L20").AutoFill Destination:=.Range("G20:L" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    
  If Range("M19") <> 0 Then
    Range("G19:L19").Value = Range("G19:L19").Value
    Range("G21").FormulaR1C1 = _
        "=6371*ACOS(SIN(R[-7]C[2])*SIN(R[-2]C[2])+COS(R[-7]C[2])*COS(R[-2]C[2])*COS(R[-2]C[3]-R[-7]C[3]))"
    Range("G22").FormulaR1C1 = "=R[-1]C+R[-3]C+R[-8]C+R[-9]C"
    ActiveSheet.Calculate
    
    If Range("G22") > Range("G10") Then
        Range("H15:L15").Value = Range("H19:L19").Value
    End If
   End If
    Range("G20:L60019").Clear
    Range("G19:L19").FormulaR1C1 = "=MAX(R[1]C:R[60000]C)"
    
    Range("G20").FormulaR1C1 = _
        "=IF(AND(R[-10]C[-6]>R13C[1],R[-10]C[-6]<R14C[1]),6371*ACOS(SIN(R[-10]C[-5])*SIN(R14C[2])+COS(R[-10]C[-5])*COS(R14C[2])*COS(R14C[3]-R[-10]C[-4])),"""")"
    Range("H20").FormulaR1C1 = "=IF(RC[-1]=R19C[-1],R[-10]C[-7],"""")"
    Range("I20:L20").FormulaR1C1 = "=IF(RC[-1]<>"""",R[-10]C[-7],"""")"
    
 With ActiveSheet
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G20:L20").AutoFill Destination:=.Range("G20:L" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    
  If Range("M19") <> 0 Then
    Range("G19:L19").Value = Range("G19:L19").Value
    Range("G21").FormulaR1C1 = _
        "=6371*ACOS(SIN(R[-9]C[2])*SIN(R[-2]C[2])+COS(R[-9]C[2])*COS(R[-2]C[2])*COS(R[-2]C[3]-R[-9]C[3]))"
    Range("G22").FormulaR1C1 = "=R[-1]C+R[-3]C+R[-6]C+R[-7]C"
    ActiveSheet.Calculate
    
    If Range("G22") > Range("G10") Then
        Range("H13:L13").Value = Range("H19:L19").Value
    End If
   End If
    Range("G20:L60019").Clear
    'CK FINI
    Range("G19:L19").FormulaR1C1 = "=MAX(R[1]C:R[60000]C)"
    Range("G20").FormulaR1C1 = _
        "=IF(R[-10]C1>R15C[1],6371*ACOS(SIN(R[-10]C2)*SIN(R15C[2])+COS(R[-10]C2)*COS(R15C[2])*COS(R15C[3]-R[-10]C3)),"""")"
    Range("H20").FormulaR1C1 = "=IF(RC[-1]=R19C[-1],R[-10]C1,"""")"
    Range("I20:L20").FormulaR1C1 = "=IF(RC[-1]<>"""",R[-10]C[-7],"""")"
    
 With ActiveSheet
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G20:L20").AutoFill Destination:=.Range("G20:L" & LastRow), Type:=xlFillDefault
    End With
 ActiveSheet.Calculate
 
    Range("G19:L19").Value = Range("G19:L19").Value
    Range("G20:L60019").Clear
    
    Range("G22").FormulaR1C1 = "=R[-9]C+R[-8]C+R[-7]C+R[-3]C"
    ActiveSheet.Calculate
    
    If Range("G22") > Range("G10") Then
        'Range("H16:L16").Value = Range("H19:L19").Value
        Range("G16:L16").Value = Range("G19:L19").Value
    End If
    
    If Range("G11") > Range("F7") Then
        Range("A1:E5").Value = Range("H12:L16").Value
        Range("F2:F5").Value = Range("G13:G16").Value
        Range("F6:F7").Value = Range("G10:G11").Value
    End If
     Range("G20:L60019").Clear
    'CK later Fini
    Range("G19:L19").FormulaR1C1 = "=MAX(R[1]C:R[60000]C)"
    Range("G20").FormulaR1C1 = _
        "=IF(R[-10]C1>R16C[1],6371*ACOS(SIN(R[-10]C2)*SIN(R16C[2])+COS(R[-10]C2)*COS(R16C[2])*COS(R16C[3]-R[-10]C3)),"""")"
    Range("H20").FormulaR1C1 = "=IF(RC[-1]=R19C[-1],R[-10]C1,"""")"
    Range("I20:L20").FormulaR1C1 = "=IF(RC[-1]<>"""",R[-10]C[-7],"""")"
    
 With ActiveSheet
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G20:L20").AutoFill Destination:=.Range("G20:L" & LastRow), Type:=xlFillDefault
    End With
 ActiveSheet.Calculate
 
  If Range("M19") <> 0 Then
    Range("G19:L19").Value = Range("G19:L19").Value
    Range("G20:L60019").Clear
    
    Range("G22").FormulaR1C1 = "=R[-8]C+R[-7]C+R[-6]C+R[-3]C"
    Range("G23").FormulaR1C1 = "=IF(R[-10]C[4]-R[-4]C[4]>R9C4,R[-1]C-(R[-10]C[4]-R[-4]C[4]-R9C4)*.1,R[-1]C)"
    ActiveSheet.Calculate
    
    If Range("G23") > Range("G11") Then
        Range("G12:L12").Value = Range("G13:L13").Value
        Range("G13:L13").Value = Range("G14:L14").Value
        Range("G14:L14").Value = Range("G15:L15").Value
        Range("G15:L15").Value = Range("G16:L16").Value
        Range("G16:L16").Value = Range("G19:L19").Value
    End If
    
    Range("G19").FormulaR1C1 = "=6371*ACOS(SIN(RC[2])*SIN(R[-4]C[2])+COS(RC[2])*COS(R[-4]C[2])*COS(R[-4]C[3]-RC[3]))"
    ActiveSheet.Calculate
    
    If Range("G19") > Range("G16") Then
        Range("G16:L16").Value = Range("G19:L19").Value
    End If
   End If
    If Range("G11") > Range("F7") Then
        Range("A1:E5").Value = Range("H12:L16").Value
        Range("F2:F5").Value = Range("G13:G16").Value
        Range("F6:F7").Value = Range("G10:G11").Value
    End If
     
    Range("P1:T3").Value = Range("H12:L14").Value
    Range("U2:U3").Value = Range("G13:G14").Value
    Columns("G:L").Clear
  
    Range("H2:L2").Value = Range("P3:T3").Value
    Range("H3:L3").Value = Range("A5:E5").Value
    'OffCourse TP2 to Fini
    Range("I4").FormulaR1C1 = _
        "=ACOS(SIN(R2C)*SIN(R3C)+COS(R2C)*COS(R3C)*COS(R3C[1]-R2C[1]))"
    Range("I5").FormulaR1C1 = _
        "=ACOS((SIN(R3C)-SIN(R2C)*COS(R[-1]C))/(SIN(R[-1]C)*COS(R2C)))"
    'MAX Off Course
     Range("G10").FormulaR1C1 = _
        "=IF(RC[-4]=R2C[3],"""",IF(AND(RC1>R2C[1],RC1<R3C[1]),ACOS(SIN(R2C[2])*SIN(RC2)+COS(R2C[2])*COS(RC2)*COS(RC3-R2C[3])),""""))"
    Range("H10").FormulaR1C1 = _
        "=IF(RC[-1]<>"""",ACOS((SIN(RC2)-SIN(R2C[1])*COS(RC[-1]))/(SIN(RC[-1])*COS(R2C[1]))),"""")"
    Range("I10").FormulaR1C1 = _
        "=IF(RC[-1]<>"""",ABS((ASIN(SIN(RC[-2])*SIN(RC[-1]-R5C))*180*60*1.852/PI())),"""")"
    Range("J10").FormulaR1C1 = "=IF(RC[-1]=R4C[1],RC1,"""")"
    Range("K10:N10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-9],"""")"
    
    Range("K4").FormulaR1C1 = "=MAX(R[6]C[-2]:R[10005]C[-2])"
    Range("L4").FormulaR1C1 = "=MAX(R[6]C[-2]:R[10005]C[-2])"
    
 With ActiveSheet
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:N10").AutoFill Destination:=.Range("G10:N" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    
    Range("J10:N60009").Value = Range("J10:N60009").Value
    Range("G10:I60009").Clear
    Range("J10:N60009").Sort Key1:=Range("J10"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    Range("P4:T4").Value = Range("J10:N10").Value
    Range("P5:T5").Value = Range("A5:E5").Value
    Range("U4").FormulaR1C1 = _
        "=IF(RC[-4]<>"""",6371*ACOS(SIN(R[-1]C[-4])*SIN(RC[-4])+COS(R[-1]C[-4])*COS(RC[-4])*COS(RC[-3]-R[-1]C[-3])),"""")"
    Range("U5").FormulaR1C1 = _
        "=IF(OR(RC[-4]="""",R[-1]C[-4]=""""),"""",6371*ACOS(SIN(R[-1]C[-4])*SIN(RC[-4])+COS(R[-1]C[-4])*COS(RC[-4])*COS(RC[-3]-R[-1]C[-3])))"
    Range("U6").FormulaR1C1 = "=SUM(R[-4]C:R[-1]C)"
    Range("U7").FormulaR1C1 = _
        "=IF(R[-6]C[-2]-R[-2]C[-2]>R9C4,R[-1]C-((R[-6]C[-2]-R[-2]C[-2]-R9C4)*0.1),R[-1]C)"
    ActiveSheet.Calculate
    Range("U4:U7").Value = Range("U4:U7").Value
  
    If Range("U7") > Range("F7") Then
        Range("A1:E5").Value = Range("P1:T5").Value
        Range("F2:F7").Value = Range("U2:U7").Value
    End If
  
    Columns("G:U").Clear
    Application.Run "F.xlsm!Fixxer"

End Sub

Sub ORDS5a()
'
' LoHmatrix after ORDS5,testing +/-18 seconds
'
Application.ScreenUpdating = False
    Sheets("Sheet3").Activate
    Range("A7").Value = 2.08333333333333E-04
    
    Range("G10").FormulaR1C1 = "=IF(AND(RC[-6]>R4C1,RC[-6]>=R5C1-R7C1,RC[-6]<=R5C1+R7C1),RC[-6],"""")"
    Range("H10:K10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-6],"""")"
    
 With Worksheets("Sheet3")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:K10").AutoFill Destination:=.Range("G10:K" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    
    Range("G10:K60009").Value = Range("G10:K60009").Value
    Range("G10:K60009").Sort Key1:=Range("G10"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    
    Range("G10:K50").Copy
    Range("N1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=True
    Columns("G:K").Clear
    
    Range("G10").FormulaR1C1 = "=IF(AND(RC[-6]>=R1C1-R7C1,RC[-6]<=R1C1+R7C1),RC[-6],"""")"
    Range("H10:K10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-6],"""")"
    Range("L10").FormulaR1C1 = "=IF(RC[-1]<>"""",6371*ACOS(SIN(RC8)*SIN(R2C2)+COS(RC8)*COS(R2C2)*COS(R2C3-RC9)),"""")"
    
 With Worksheets("Sheet3")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:L10").AutoFill Destination:=.Range("G10:L" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    
    Range("G10:L60009").Value = Range("G10:L60009").Value
    Range("G10:L60009").Sort Key1:=Range("G10"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    
    Range("N6:AY6").FormulaR1C1 = _
        "=IF(OR(R1C="""",R3C=R[-2]C3),"""",6371*ACOS(SIN(R2C)*SIN(R[-2]C2)+COS(R2C)*COS(R[-2]C2)*COS(R[-2]C3-R3C)))"
    ActiveSheet.Calculate
    Range("N6:AY6").Value = Range("N6:AY6").Value
        
    'Matrix
    Range("N10:AY50").FormulaR1C1 = _
        "=IF(OR(R1C="""",RC8=""""),"""",IF(RC10-R4C>R9C4,(RC12+R6C)-((RC10-R4C-R9C4)*0.1),RC12+R6C))"
    Range("I5").FormulaR1C1 = "=MAX(R[5]C[5]:R[55]C[42])"
    ActiveSheet.Calculate
    Range("I5").Value = Range("I5").Value
    Range("N10:AY60").Value = Range("N10:AY60").Value
    Range("N7:AY7").FormulaR1C1 = "=IF(MAX(R[3]C:R[10002]C)=R5C9,1,"""")"
    ActiveSheet.Calculate
    Range("N7:AY7").Value = Range("N7:AY7").Value
    
    Range("M10").FormulaR1C1 = "=IF(MAX(RC[1]:RC[500])=R5C9,1,"""")"
    
 With Worksheets("Sheet3")
    LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
.Range("M10:M10").AutoFill Destination:=.Range("M10:M" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("M10:M10009").Value = Range("M10:M10009").Value
    
    Range("N10:AY60").Clear
    
    Range("N10").FormulaR1C1 = "=IF(RC[-1]=1,RC[-7],"""")"
    Range("O10:R10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-7],"""")"

 With Worksheets("Sheet3")
    LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
.Range("N10:R10").AutoFill Destination:=.Range("N10:R" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    
    Range("H1:L1").FormulaR1C1 = "=MAX(R[9]C[6]:R[10008]C[6])"
    ActiveSheet.Calculate
    Range("A1:E1").Value = Range("H1:L1").Value
    Range("N10:R10009").Clear
    
    Range("N10:AY10").FormulaR1C1 = "=IF(R[-3]C=1,R[-9]C,"""")"
    Range("N11:AY14").FormulaR1C1 = "=IF(R[-1]C<>"""",R[-9]C,"""")"
    
    Range("H5").FormulaR1C1 = "=MAX(R[5]C[6]:R[5]C[44])"
    Range("I5").FormulaR1C1 = "=MAX(R[6]C[5]:R[6]C[43])"
    Range("J5").FormulaR1C1 = "=MAX(R[7]C[4]:R[7]C[42])"
    Range("K5").FormulaR1C1 = "=MAX(R[8]C[3]:R[8]C[41])"
    Range("L5").FormulaR1C1 = "=MAX(R[9]C[2]:R[9]C[40])"
    ActiveSheet.Calculate
    Range("A5:E5").Value = Range("H5:L5").Value
    
    Columns("G:AY").Clear
    Range("F2:F5").FormulaR1C1 = _
        "=6371*ACOS(SIN(RC[-4])*SIN(R[-1]C[-4])+COS(RC[-4])*COS(R[-1]C[-4])*COS(R[-1]C[-3]-RC[-3]))"
    Range("F6").FormulaR1C1 = "=SUM(R[-4]C:R[-1]C)"
    Range("F7").FormulaR1C1 = _
        "=IF(R[-6]C[-2]-R[-2]C[-2]>R9C4,R[-1]C-((R[-6]C[-2]-R[-2]C[-2]-R9C4)*0.1),R[-1]C)"
    ActiveSheet.Calculate
    Range("F2:F7").Value = Range("F2:F7").Value

End Sub

Sub Vince()
'
' Geodetic Official distances
'
Application.ScreenUpdating = False
Range("A1").FormulaR1C1 = "=Sheet2!R[8]C[3]"
Range("A2").FormulaR1C1 = "=Sheet2!R[7]C[4]"
Range("A1:A2").Value = Range("A1:A2").Value

    'Straight
    Sheets("YDWK3").Range("E39").Value = Sheets("Tasks").Range("D10").Value
    Sheets("YDWK3").Range("E40").Value = Sheets("Tasks").Range("E10").Value
    Sheets("YDWK3").Range("E41").Value = Sheets("Tasks").Range("D11").Value
    Sheets("YDWK3").Range("E42").Value = Sheets("Tasks").Range("E11").Value
    Sheets("YDWK3").Calculate
    Sheets("Tasks").Range("L11:M11").Value = Sheets("YDWK3").Range("F37").Value
    
    'O&R
    Sheets("YDWK3").Range("E39").Value = Sheets("Tasks").Range("D14").Value
    Sheets("YDWK3").Range("E40").Value = Sheets("Tasks").Range("E14").Value
    Sheets("YDWK3").Range("E41").Value = Sheets("Tasks").Range("D15").Value
    Sheets("YDWK3").Range("E42").Value = Sheets("Tasks").Range("E15").Value
    Sheets("YDWK3").Calculate
    Sheets("Tasks").Range("L15").Value = Sheets("YDWK3").Range("F37").Value
    Sheets("Tasks").Range("M15").FormulaR1C1 = "=2*RC[-1]"
    ActiveSheet.Calculate
    Sheets("Tasks").Range("M15").Value = Sheets("Tasks").Range("M15").Value
    
    'PilotOption
    Sheets("YDWK3").Range("E39").Value = Sheets("Tasks").Range("D20").Value
    Sheets("YDWK3").Range("E40").Value = Sheets("Tasks").Range("E20").Value
    Sheets("YDWK3").Range("E41").Value = Sheets("Tasks").Range("D21").Value
    Sheets("YDWK3").Range("E42").Value = Sheets("Tasks").Range("E21").Value
    Sheets("YDWK3").Calculate
    Sheets("Tasks").Range("L21").Value = Sheets("YDWK3").Range("F37").Value
    
    Sheets("YDWK3").Range("E39").Value = Sheets("Tasks").Range("D21").Value
    Sheets("YDWK3").Range("E40").Value = Sheets("Tasks").Range("E21").Value
    Sheets("YDWK3").Range("E41").Value = Sheets("Tasks").Range("D22").Value
    Sheets("YDWK3").Range("E42").Value = Sheets("Tasks").Range("E22").Value
    Sheets("YDWK3").Calculate
    Sheets("Tasks").Range("L22").Value = Sheets("YDWK3").Range("F37").Value
    
    Sheets("YDWK3").Range("E39").Value = Sheets("Tasks").Range("D22").Value
    Sheets("YDWK3").Range("E40").Value = Sheets("Tasks").Range("E22").Value
    Sheets("YDWK3").Range("E41").Value = Sheets("Tasks").Range("D23").Value
    Sheets("YDWK3").Range("E42").Value = Sheets("Tasks").Range("E23").Value
    Sheets("YDWK3").Calculate
    Sheets("Tasks").Range("L23").Value = Sheets("YDWK3").Range("F37").Value
    
    Sheets("YDWK3").Range("E39").Value = Sheets("Tasks").Range("D23").Value
    Sheets("YDWK3").Range("E40").Value = Sheets("Tasks").Range("E23").Value
    Sheets("YDWK3").Range("E41").Value = Sheets("Tasks").Range("D24").Value
    Sheets("YDWK3").Range("E42").Value = Sheets("Tasks").Range("E24").Value
    Sheets("YDWK3").Calculate
    Sheets("Tasks").Range("L24").Value = Sheets("YDWK3").Range("F37").Value
    Sheets("Tasks").Range("M24").FormulaR1C1 = "=SUM(R[-3]C[-1]:RC[-1])"
    Sheets("Tasks").Range("M24").Value = Sheets("Tasks").Range("M24").Value
    
    Application.Run "F.xlsm!YDWK3clear"
    
    Sheets("TASKS").Activate
    Range("N11").FormulaR1C1 = _
        "=IF(OR(AND(RC[-1]>100,ROUND(R[-1]C[-6]-RC[-6],0)<=R1C1),AND(RC[-1]<=100,R2C1<>""PR"",ROUND(R[-1]C[-6]-RC[-6],0)<=10*RC[-1]),AND(RC[-1]<=100,R2C1=""PR"",ROUND(R[-1]C[-6]-RC[-6],0)<10*RC[-1]-100)),RC[-1],IF(AND(RC[-1]>100,ROUND(R[-1]C[-6]-RC[-6],0)>R1C1),RC[-1]-((ROUND(R[-1]C[-6]-RC[-6],0)-R1C1)*0.1),0))"
    Range("M12").FormulaR1C1 = "=IF(R[-1]C[1]<R[-1]C,""LoH"","""")"
    ActiveSheet.Calculate
    Range("M12").Value = Range("M12").Value
    Range("M11").Value = Range("N11").Value
    
    Range("N15").FormulaR1C1 = _
        "=IF(OR(AND(RC[-1]>100,ROUND(R[-1]C[-6]-R[2]C[-6],0)<=R1C1),AND(RC[-1]<=100,R2C1<>""PR"",ROUND(R[-1]C[-6]-R[2]C[-6],0)<=10*RC[-1]),AND(RC[-1]<=100,R2C1=""PR"",ROUND(R[-1]C[-6]-R[2]C[-6],0)<10*RC[-1]-100)),RC[-1],IF(AND(RC[-1]>100,ROUND(R[-1]C[-6]-R[2]C[-6],0)>R1C1),RC[-1]-((ROUND(R[-1]C[-6]-R[2]C[-6],0)-R1C1)*0.1),0))"
    Range("M17").FormulaR1C1 = "=IF(R[-2]C[1]<R[-2]C,""LoH"","""")"
    ActiveSheet.Calculate
    Range("M17").Value = Range("M17").Value
    Range("M15").Value = Range("N15").Value
    
    Range("N24").FormulaR1C1 = _
        "=IF(OR(AND(RC[-1]>100,ROUND(R[-4]C[-6]-RC[-6],0)<=R1C1),AND(RC[-1]<=100,R2C1<>""PR"",ROUND(R[-4]C[-6]-RC[-6],0)<=10*RC[-1]),AND(RC[-1]<=100,R2C1=""PR"",ROUND(R[-4]C[-6]-RC[-6],0)<10*RC[-1]-100)),RC[-1],IF(AND(RC[-1]>100,ROUND(R[-4]C[-6]-RC[-6],0)>R1C1),RC[-1]-((ROUND(R[-4]C[-6]-RC[-6],0)-R1C1)*0.1),0))"
    Range("M25").FormulaR1C1 = "=IF(R[-1]C[1]<R[-1]C,""LoH"","""")"
    ActiveSheet.Calculate
    Range("M25").Value = Range("M25").Value
    Range("M24").Value = Range("N24").Value
    
    Range("N11,N15,N24").Clear
    
    If Range("F17") = "   FINISH SECTOR" And Range("C10") >= 42278 Then
            Range("C18:M18").Value = "Note: Per SC3 rules on this flight date, Out & Return courses require Finish Line crossing"
    End If
End Sub