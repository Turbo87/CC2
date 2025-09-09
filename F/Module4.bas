' VBA Module: Triangle Task Analysis
' Purpose: Manages triangle task calculations and turn point optimization.
' Handles 2-point and 3-point triangle analysis, start/finish calculations,
' and spherical trigonometry for competitive soaring task validation.

Option Explicit
#If VBA7 And Win64 Then
    Dim LastRow As LongPtr
#Else
    Dim LastRow As Long
#End If

Sub Triangle1()
'
' 2/8/14 JLR finds 3 TP triangle, based on StDist OffCourse & O&R Start/Fini Macro CloseSF added 10/9/14; W6 modified for 2-TP Triangles via O&R S/F or NearestS/F
'
 Application.ScreenUpdating = False
    Application.Run "F.xlsm!CloseSF"
    Sheets("Sheet2").Range("A7").Value = Sheets("TASKS").Range("G40").Value
    Sheets("Sheet2").Range("B7").Value = Sheets("TASKS").Range("G41").Value
    Sheets("Sheet2").Range("A8").Value = Sheets("TASKS").Range("C14").Value
    Sheets("Sheet2").Range("B8").Value = Sheets("TASKS").Range("C16").Value
    Sheets("Sheet2").Range("H2:J3").Value = Sheets("Tasks").Range("C10:E11").Value
    Sheets("Sheet2").Range("K2:L3").Value = Sheets("Tasks").Range("H10:I11").Value
    
    Range("I4").FormulaR1C1 = _
        "=ACOS(SIN(R2C)*SIN(R3C)+COS(R2C)*COS(R3C)*COS(R3C[1]-R2C[1]))"
    Range("K4").FormulaR1C1 = "=MAX(R[6]C[-2]:R[10005]C[-2])"
    Range("L4").FormulaR1C1 = "=MAX(R[6]C[-2]:R[10005]C[-2])"
    Range("I5").FormulaR1C1 = _
        "=ACOS((SIN(R3C)-SIN(R2C)*COS(R[-1]C))/(SIN(R[-1]C)*COS(R2C)))"
        
    Range("J8:N8").FormulaR1C1 = "=MAX(R[2]C:R[10001]C)"
    
    Range("G10").FormulaR1C1 = _
        "=IF(RC[-4]=R2C[3],"""",IF(AND(RC1>R8C1,RC1<R8C2),ACOS(SIN(R2C[2])*SIN(RC2)+COS(R2C[2])*COS(RC2)*COS(RC3-R2C[3])),""""))"
    Range("H10").FormulaR1C1 = _
        "=IF(RC[-1]<>"""",ACOS((SIN(RC2)-SIN(R2C[1])*COS(RC[-1]))/(SIN(RC[-1])*COS(R2C[1]))),"""")"
    Range("I10").FormulaR1C1 = _
        "=IF(RC[-1]<>"""",ABS((ASIN(SIN(RC[-2])*SIN(RC[-1]-R5C))*180*60*1.852/PI())),"""")"
    Range("J10").FormulaR1C1 = "=IF(RC[-1]=R4C[1],RC1,"""")"
    Range("K10:N10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-9],"""")"

    Application.Calculation = xlCalculationManual
 With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:N10").AutoFill Destination:=.Range("G10:N" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("P2:T2").Value = Range("J8:N8").Value
    
    Sheets("Sheet2").Range("P3:R4").Value = Sheets("TASKS").Range("C10:E11").Value
    Sheets("Sheet2").Range("S3:T4").Value = Sheets("TASKS").Range("H10:I11").Value
    Range("P2:T4").Sort Key1:=Range("P2"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    
    Range("U2:U3").FormulaR1C1 = "=6371*ACOS(SIN(RC[-4])*SIN(R[1]C[-4])+COS(RC[-4])*COS(R[1]C[-4])*COS(R[1]C[-3]-RC[-3]))"
    Range("U4").FormulaR1C1 = "=6371*ACOS(SIN(RC[-4])*SIN(R[-2]C[-4])+COS(RC[-4])*COS(R[-2]C[-4])*COS(R[-2]C[-3]-RC[-3]))"
    Range("V5").FormulaR1C1 = "=SUM(R[-3]C[-1]:R[-1]C[-1])"
    Range("V6").FormulaR1C1 = "=IF(TASKS!R[8]C[-14]-MAX(TASKS!R[10]C[-14],TASKS!R[11]C[-14])>R9C4,R[-1]C-((TASKS!R[8]C[-14]-MAX(TASKS!R[10]C[-14],TASKS!R[11]C[-14])-R9C4)*0.1),R[-1]C)"
    Range("F9").Value = 6.94444444444444E-04
    'Range("W6").FormulaR1C1 = "=IF(OR(AND(RC[-1]<750,MIN(R[-4]C[-2]:R[-2]C[-2])>=0.28*RC[-1]),AND(RC[-1]>=750,MIN(R[-4]C[-2]:R[-2]C[-2])>=0.25*RC[-1],MAX(R[-4]C[-2]:R[-2]C[-2])<=0.45*RC[-1])),""OK"","""")"
    Range("W6").FormulaR1C1 = _
        "=IF(OR(ABS(R[-2]C[-7]-R[2]C[-21])<R[3]C[-17],ABS(R[-4]C[-7]-R[2]C[-22])<R[3]C[-17],ABS(R[-4]C[-7]-R[1]C[-22])<5*R[3]C[-17]),""2TP"",IF(OR(AND(RC[-1]<750,MIN(R[-4]C[-2]:R[-2]C[-2])>=0.28*RC[-1]),AND(RC[-1]>=750,MIN(R[-4]C[-2]:R[-2]C[-2])>=0.25*RC[-1],MAX(R[-4]C[-2]:R[-2]C[-2])<=0.45*RC[-1])),""OK"",""""))"
    Range("P2:W6").Value = Range("P2:W6").Value
  '''''''''''''''''
  If Range("W6") = "2TP" Then
    Application.Run ("F.xlsm!SibTri")
  End If
  
  If Range("W6") = "" Then
    Sheets("Sheet2").Range("A1:E4").Value = Sheets("Tasks").Range("A40:E44").Value
    Range("A6").Value = 2.08333333333333E-03
    Range("P6").FormulaR1C1 = _
        "=IF(AND(R[-4]C>=R[-4]C[-15]-RC[-15],R[-4]C<=R[-4]C[-15]+RC[-15],R[-2]C>=R[-2]C[-15]-RC[-15],R[-2]C<=R[-2]C[-15]+RC[-15]),""2 X 4"",IF(AND(R[-4]C>=R[-5]C[-15]-RC[-15],R[-4]C<=R[-5]C[-15]+RC[-15],R[-3]C>=R[-4]C[-15]-RC[-15],R[-3]C<R[-4]C[-15]+RC[-15],R[-2]C>=R[-3]C[-15]-RC[-15],R[-2]C<=R[-3]C[-15]+RC[-15]),""123"",""""))"
    Range("P6").Value = Range("P6").Value

If Range("P6") = 123 Then
    Application.Run "F.xlsm!TriAlmost"
    
ElseIf Range("P6") = "2 X 4" Then
    Range("P3:U3").Clear
    Range("U2:V6").Clear
    Range("U4").FormulaR1C1 = "=6371*ACOS(SIN(RC[-4])*SIN(R[-2]C[-4])+COS(RC[-4])*COS(R[-2]C[-4])*COS(R[-2]C[-3]-RC[-3]))"
    
    Range("G10").FormulaR1C1 = "=IF(AND(RC[-6]>R2C16,RC[-6]<R4C16),RC[-6],"""")"
    Range("H10:K10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-6],"""")"
    'Copy Ref A Value Sort
    Application.Calculation = xlCalculationManual
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:K10").AutoFill Destination:=.Range("G10:K" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("G10:K10009").Value = Range("G10:K10009").Value
    Range("G10:K10009").Sort Key1:=Range("G10"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal

    Range("P9:U9").FormulaR1C1 = "=MAX(R[1]C:R[10000]C)"
    
    Range("L10").FormulaR1C1 = "=6371*ACOS(SIN(RC[-4])*SIN(R2C17)+COS(RC[-4])*COS(R2C17)*COS(R2C18-RC[-3]))"
    Range("M10").FormulaR1C1 = "=6371*ACOS(SIN(RC[-5])*SIN(R4C17)+COS(RC[-5])*COS(R4C17)*COS(R4C18-RC[-4]))"
    Range("N10").FormulaR1C1 = "=R4C21"
    Range("O10").FormulaR1C1 = "=SUM(RC[-3]:RC[-1])"
    Range("P10").FormulaR1C1 = _
        "=IF(OR(AND(RC[-1]<750,MIN(RC[-4]:RC[-2])>=0.28*RC[-1]),AND(RC[-1]>=750,MIN(RC[-4]:RC[-2])>=0.25*RC[-1],MAX(RC[-4]:RC[-2])<=0.45*RC[-1])),RC[-1],"""")"
    Range("Q10").FormulaR1C1 = "=IF(RC[-1]=R9C16,RC[-10],"""")"
    Range("R10:U10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-10],"""")"
    'Copy Ref G
    Application.Calculation = xlCalculationManual
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
.Range("L10:U10").AutoFill Destination:=.Range("L10:U" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    If Range("Q9") <> 0 Then
        Range("P3:T3").Value = Range("Q9:U9").Value
        Range("U2:U3").FormulaR1C1 = "=6371*ACOS(SIN(RC[-4])*SIN(R[1]C[-4])+COS(RC[-4])*COS(R[1]C[-4])*COS(R[1]C[-3]-RC[-3]))"
        Range("V5").FormulaR1C1 = "=SUM(R[-3]C[-1]:R[-1]C[-1])"
        Range("W6").FormulaR1C1 = "=IF(OR(AND(R[-1]C[-1]<750,MIN(R[-4]C[-2]:R[-2]C[-2])>=0.28*R[-1]C[-1]),AND(R[-1]C[-1]>=750,MIN(R[-4]C[-2]:R[-2]C[-2])>=0.25*R[-1]C[-1],MAX(R[-4]C[-2]:R[-2]C[-2])<=0.45*R[-1]C[-1])),""OK"","""")"
        Range("U2:W6").Value = Range("U2:W6").Value
    End If
 
    Range("A6").FormulaR1C1 = "=R[-4]C-R[-5]C"
      If Range("A6") < 1 / 48 Then
        Range("F5").FormulaR1C1 = "=6371*ACOS(SIN(R[-4]C[-4])*SIN(R[-2]C[-4])+COS(R[-4]C[-4])*COS(R[-2]C[-4])*COS(R[-2]C[-3]-R[-4]C[-3]))"
        Range("F6").FormulaR1C1 = "=6371*ACOS(SIN(R[-3]C[-4])*SIN(R[-2]C[-4])+COS(R[-3]C[-4])*COS(R[-2]C[-4])*COS(R[-2]C[-3]-R[-3]C[-3]))"
        Range("F7").FormulaR1C1 = "=6371*ACOS(SIN(R[-3]C[-4])*SIN(R[-6]C[-4])+COS(R[-3]C[-4])*COS(R[-6]C[-4])*COS(R[-6]C[-3]-R[-3]C[-3]))"
        Range("F8").FormulaR1C1 = "=SUM(R[-3]C:R[-1]C)"
        Range("G8").FormulaR1C1 = "=IF(OR(AND(RC[-1]<750,MIN(R[-3]C[-1]:R[-1]C[-1])>=0.28*RC[-1]),AND(RC[-1]>=750,MIN(R[-3]C[-1]:R[-1]C[-1])>=0.25*RC[-1],MAX(R[-3]C[-1]:R[-1]C[-1])<=0.45*RC[-1])),""OK"",""NOPE"")"
        Range("F5:G8").Value = Range("F5:G8").Value
    
        If Range("W6") = "OK" And Range("G8") = "OK" And Range("V5") < Range("F8") Then
        Range("F5:G8,W6").Clear
        End If
      End If
   End If
End If
 
    Range("A6,P6,G7:U10009").Clear
  
  If Range("W6") = "OK" And Range("A1") <> "This is a 2-Turn Point Triangle" Then
    Application.Run "F.xlsm!ReCk"
    Columns("G:N").Clear
    Range("A2:F4").Value = Range("P2:U4").Value
    Range("F6:F7").Value = Range("V5:V6").Value
    Range("E6").Value = Range("X6").Value
    Range("A8:B8,P2:X7").Clear
  ElseIf Range("W6") <> "OK" And Range("A1") <> "This is a 2-Turn Point Triangle" Then
    Columns("G:N").Clear
    Application.Run "F.xlsm!TriOrds"
  End If

If Range("A1") <> "This is a 2-Turn Point Triangle" And Range("A1") <> "NONE IDENTIFIED" Then
    Sheets("Sheet2").Range("A1:C1").Value = Sheets("TASKS").Range("C14:E14").Value
    Sheets("Sheet2").Range("E1:D1").Value = Sheets("TASKS").Range("H14:I14").Value
    Sheets("Sheet2").Range("A5:C5").Value = Sheets("TASKS").Range("C17:E17").Value
    Sheets("Sheet2").Range("D5:E5").Value = Sheets("TASKS").Range("H17:I17").Value
    
    If Range("A4") > Range("A5") Or Range("A2") < Range("A1") Then
        Application.Run "F.xlsm!TriSFredux"
    End If

    Sheets("YDWK3").Range("E39").Value = Sheets("Sheet2").Range("B2").Value
    Sheets("YDWK3").Range("E40").Value = Sheets("Sheet2").Range("C2").Value
    Sheets("YDWK3").Range("E41").Value = Sheets("Sheet2").Range("B3").Value
    Sheets("YDWK3").Range("E42").Value = Sheets("Sheet2").Range("C3").Value
    Sheets("YDWK3").Calculate
    Sheets("Sheet2").Range("G2").Value = Sheets("YDWK3").Range("F37").Value
    
    Sheets("YDWK3").Range("E41").Value = Sheets("Sheet2").Range("B4").Value
    Sheets("YDWK3").Range("E42").Value = Sheets("Sheet2").Range("C4").Value
    Sheets("YDWK3").Calculate
    Sheets("Sheet2").Range("G4").Value = Sheets("YDWK3").Range("F37").Value
    
    Sheets("YDWK3").Range("E39").Value = Sheets("Sheet2").Range("B3").Value
    Sheets("YDWK3").Range("E40").Value = Sheets("Sheet2").Range("C3").Value
    Sheets("YDWK3").Calculate
    Sheets("Sheet2").Range("G3").Value = Sheets("YDWK3").Range("F37").Value
    
ElseIf Range("A1") = "This is a 2-Turn Point Triangle" Then
    Sheets("YDWK3").Range("E39").Value = Sheets("Sheet2").Range("B2").Value
    Sheets("YDWK3").Range("E40").Value = Sheets("Sheet2").Range("C2").Value
    Sheets("YDWK3").Range("E41").Value = Sheets("Sheet2").Range("B3").Value
    Sheets("YDWK3").Range("E42").Value = Sheets("Sheet2").Range("C3").Value
    Sheets("YDWK3").Calculate
    Sheets("Sheet2").Range("G2").Value = Sheets("YDWK3").Range("F37").Value
    
    Sheets("YDWK3").Range("E39").Value = Sheets("Sheet2").Range("B4").Value
    Sheets("YDWK3").Range("E40").Value = Sheets("Sheet2").Range("C4").Value
    Sheets("YDWK3").Calculate
    Sheets("Sheet2").Range("G3").Value = Sheets("YDWK3").Range("F37").Value
    
    Sheets("YDWK3").Range("E41").Value = Sheets("Sheet2").Range("B2").Value
    Sheets("YDWK3").Range("E42").Value = Sheets("Sheet2").Range("C2").Value
    Sheets("YDWK3").Calculate
    Sheets("Sheet2").Range("G4").Value = Sheets("YDWK3").Range("F37").Value
 End If
 
Application.Run "F.xlsm!YDWK3clear"
Range("G6").FormulaR1C1 = "=SUM(R[-4]C:R[-2]C)"
Range("G7").FormulaR1C1 = "=IF(OR(AND(R[-1]C>=750,MIN(R[-5]C:R[-3]C)>=0.25*R[-1]C,MAX(R[-5]C:R[-3]C)<=0.45*R[-1]C),AND(R[-1]C<750,MIN(R[-5]C:R[-3]C)>=0.28*R[-1]C)),""OK"","""")"
Range("E7").FormulaR1C1 = "=IF(OR(R1C1=""NONE IDENTIFIED"",RC[1]=""""),""NONE IDENTIFIED"",IF(OR(AND(RC[1]>=750,MIN(R[-5]C[1]:R[-3]C[1])>=0.25*RC[1],MAX(R[-5]C[1]:R[-3]C[1])<=0.45*RC[1]),AND(RC[1]<750,MIN(R[-5]C[1]:R[-3]C[1])>=0.28*RC[1])),""OK"",""""))"
Range("E8").FormulaR1C1 = "=IF(AND(R[-1]C=""OK"",R[-1]C[2]=""""),""No record-eligible Triangle found. Great Circle distances are shown in red above"","""")"
Range("G6:G7").Value = Range("G6:G7").Value
Range("E7:E8").Value = Range("E7:E8").Value

If Range("E7") = "NONE IDENTIFIED" Then
    Sheets("TASKS").Range("G26").Value = Sheets("Sheet2").Range("E7").Value
    Sheets("Sheet2").Range("A1:X8").Clear
    Sheets("TASKS").Activate
    'Restore raw pressure altitudes to K
    Range("K10,K11,K14:K17,K20:K24").FormulaR1C1 = "=IF(RC[-8]<B!R5C13,RC[-3]-B!R2C13,RC[-3]-B!R8C13)"
    Range("K27:K29").FormulaR1C1 = "=IF(R[-1]C[-4]=""NONE IDENTIFIED"","""",IF(RC[-8]<B!R5C13,RC[-3]-B!R2C13,RC[-3]-B!R8C13))"
    Range("K30").FormulaR1C1 = "=IF(OR(R26C7=""NONE IDENTIFIED"",RC[-8]=""""),"""",IF(RC[-8]<B!R5C13,RC[-3]-B!R2C13,RC[-3]-B!R8C13))"
    Range("K31").FormulaR1C1 = "=IF(OR(R26C7=""NONE IDENTIFIED"",RC[-8]=""""),"""",IF(RC[-8]<B!R5C13,RC[-3]-B!R2C13,RC[-3]-B!R8C13))"
    Range("K10:K31").Value = Range("K10:K31").Value
    If Range("A2") = "PR" Then
        Range("H10:H11,H14:H17,H20:H24").Value = "N/A"
        Range("K10,K11,K14:K17,K20:K24").FormulaR1C1 = "=RC[-2]"
    ElseIf Range("A2") <> "PR" Then
        Range("K10,K11,K14:K17,K20:K24").FormulaR1C1 = "=IF(RC[-8]<B!R5C13,RC[-3]-B!R2C13,RC[-3]-B!R8C13)"
    End If
    Range("H10:H24").Value = Range("K10:K24").Value
    Range("A27:A32").EntireRow.Hidden = True
    Exit Sub
End If

If Range("E8") = "" Then
    Range("F2:F6").Value = Range("G2:G6").Value
    Range("F7").FormulaR1C1 = "=IF(R[-6]C[-2]-R[-2]C[-2]>R9C4,R[-1]C-((R[-6]C[-2]-R[-2]C[-2]-R9C4)*0.1),R[-1]C)"
    Range("F7").Value = Range("F7").Value
    Range("F8").Value = Range("G7").Value
    Range("G2:G10009").Clear
End If
Range("A7:B8,F9").Clear
Columns("H:X").Clear
Range("A1").Select
Application.Run "F.xlsm!TriFline1"
End Sub
Sub CloseSF()
'
' Find closest S/F for Triangle claims Run early in Triangle1 for Antares et al
'
    Sheets("Sheet2").Range("A6").Value = Sheets("TASKS").Range("C14").Value
    Sheets("Sheet2").Range("B6").Value = Sheets("TASKS").Range("C17").Value
    
    Range("A7").Value = 1.38888888888889E-03
    Range("G10").FormulaR1C1 = "=IF(OR(AND(RC[-6]>=R6C1-R7C1,RC[-6]<=R6C1+R7C1),AND(RC[-6]>=R6C2-R7C1,RC[-6]<=R6C2+R7C1)),RC[-6],"""")"
    Range("H10:K10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-6],"""")"
    'Copy Ref A Value Sort
    Application.Calculation = xlCalculationManual
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:K10").AutoFill Destination:=.Range("G10:K" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("G10:K10009").Value = Range("G10:L10009").Value
    Range("G10:K10009").Sort Key1:=Range("G10"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
   
    Range("M10").FormulaR1C1 = "=IF(RC[-6]>=R6C2-R7C1,RC[-6],"""")"
    Range("N10:Q10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-6],"""")"
    'Copy Ref G Value Sort
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
.Range("M10:Q10").AutoFill Destination:=.Range("M10:Q" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("M10:Q509").Value = Range("M10:Q509").Value
    Range("M10:Q509").Sort Key1:=Range("M10"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
   
    Range("M10:Q250").Copy
    Range("N1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    
    Range("M10:Q500").Clear
    'MATRIX
    Range("N10:IV10").FormulaR1C1 = "=IF(OR(RC7>=R1C14,RC9=R3C,R1C=""""),"""",6371*ACOS(SIN(RC8)*SIN(R2C)+COS(RC8)*COS(R2C)*COS(R3C-RC9)))"
    'Copy Ref G
     Application.Calculation = xlCalculationManual
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
.Range("N10:IV10").AutoFill Destination:=.Range("N10:IV" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    
    Range("I5").FormulaR1C1 = "=MIN(R[5]C[5]:R[255]C[247])"
    Range("I5").Value = Range("I5").Value
    Range("N6:IV6").FormulaR1C1 = "=IF(MIN(R[4]C:R[314]C)=R5C9,R[-5]C,"""")"
    Range("N6:IV6").Value = Range("N6:IV6").Value
    
    Range("L10:L250").FormulaR1C1 = "=IF(MIN(RC[2]:RC[244])=R5C9,RC[-5],"""")"
    Range("L10:L250").Value = Range("L10:L250").Value
    Range("N10:IV509").Clear
    
    Range("N7:IV10").FormulaR1C1 = "=IF(R[-1]C<>"""",R[-5]C,"""")"
    
    Range("G7").FormulaR1C1 = "=MAX(R[-1]C[7]:R[-1]C[249])"
    Range("H7").FormulaR1C1 = "=MAX(RC[6]:RC[248])"
    Range("I7").FormulaR1C1 = "=MAX(R[1]C[5]:R[1]C[247])"
    Range("J7").FormulaR1C1 = "=MAX(R[2]C[4]:R[2]C[246])"
    Range("K7").FormulaR1C1 = "=MAX(R[3]C[3]:R[3]C[245])"
    ActiveSheet.Calculate
    Range("G7:K7").Value = Range("G7:K7").Value
    
    Range("M10:P250").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-5],"""")"
    
    Range("G8:K8").FormulaR1C1 = "=MAX(R[2]C[5]:R[252]C[5])"
    Range("G8:K8").Value = Range("G8:K8").Value
    
    Range("G7:K8").Sort Key1:=Range("G7"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    
    Sheets("TASKS").Range("G40:K41").Value = Sheets("Sheet2").Range("G7:K8").Value
    Sheets("Sheet2").Range("A6:B7,G5:M10009,N1:IV1000").Clear

End Sub

Sub TriAlmost()
'
' 4/28/14 JLR  when 1st TRI1 is almost ORDS 123 fo AussieNoName winch
'
Application.ScreenUpdating = False
    Range("T6").FormulaR1C1 = "=IF(RC[2]<750,0.28*RC[2],0.25*RC[2])"
    Range("T6").Value = Range("T6").Value
    
    Range("G10").FormulaR1C1 = _
        "=IF(AND(R[-4]C[-6]<R3C16,6371*ACOS(SIN(RC[-5])*SIN(R3C17)+COS(RC[-5])*COS(R3C17)*COS(R3C18-RC[-4]))>=R6C20),RC[-6],"""")"
    Range("H10:K10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-6],"""")"
    'Copy Ref A Value Sort
    Application.Calculation = xlCalculationManual
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:K10").AutoFill Destination:=.Range("G10:K" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("G10:K10009").Value = Range("G10:K10009").Value
    Range("G10:K10009").Sort Key1:=Range("G10"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
  
  If Range("G10") <> "" Then
    
    Range("P8:U8").FormulaR1C1 = "=MAX(R[2]C:R[10001]C)"
    Range("L10").FormulaR1C1 = "=6371*ACOS(SIN(RC[-4])*SIN(R3C17)+COS(RC[-4])*COS(R3C17)*COS(R3C18-RC[-3]))"
    Range("M10").FormulaR1C1 = "=R3C21"
    Range("N10").FormulaR1C1 = "=6371*ACOS(SIN(RC[-6])*SIN(R4C17)+COS(RC[-6])*COS(R4C17)*COS(R4C18-RC[-5]))"
    Range("O10").FormulaR1C1 = "=SUM(RC[-3]:RC[-1])"
    Range("P10").FormulaR1C1 = "=IF(OR(AND(RC[-1]<750,MIN(RC[-4]:RC[-2])>=0.28*RC[-1]),AND(RC[-1]>=750,MIN(RC[-4]:RC[-2])>=0.25*RC[-1],MAX(RC[-4]:RC[-2])<=0.45*RC[-1])),RC[-1],"""")"
    Range("Q10").FormulaR1C1 = "=IF(RC[-1]=R8C16,RC[-10],"""")"
    Range("R10:U10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-10],"""")"
    'Copy Ref G
     Application.Calculation = xlCalculationManual
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
.Range("L10:U10").AutoFill Destination:=.Range("L10:U" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    If Range("P8") <> 0 Then
    Range("P2:T2").Value = Range("Q8:U8").Value

    Range("U2:U3").FormulaR1C1 = "=6371*ACOS(SIN(RC[-4])*SIN(R[1]C[-4])+COS(RC[-4])*COS(R[1]C[-4])*COS(R[1]C[-3]-RC[-3]))"
    Range("U4").FormulaR1C1 = "=6371*ACOS(SIN(RC[-4])*SIN(R[-2]C[-4])+COS(RC[-4])*COS(R[-2]C[-4])*COS(R[-2]C[-3]-RC[-3]))"
    Range("V5").FormulaR1C1 = "=SUM(R[-3]C[-1]:R[-1]C[-1])"
    Range("V6").FormulaR1C1 = "=IF(TASKS!R[8]C[-14]-TASKS!R[10]C[-14]<R9C4,R[-1]C,R[-1]C-((TASKS!R[8]C[-14]-TASKS!R[10]C[-14]-R9C4)*0.1))"
    Range("W6").FormulaR1C1 = "=IF(OR(AND(RC[-1]<750,MIN(R[-4]C[-2]:R[-2]C[-2])>=0.28*RC[-1]),AND(RC[-1]>=750,MIN(R[-4]C[-2]:R[-2]C[-2])>=0.25*RC[-1],MAX(R[-4]C[-2]:R[-2]C[-2])<=0.45*RC[-1])),""OK"","""")"
    Range("U2:W6").Value = Range("U2:W6").Value
    End If
    Range("G10:O10009").Clear
    Range("T6,P8:U10009").Clear
 End If
End Sub
Sub SibTri()
'
'
' Triangle based on St Dist s/f as TPs, w/ O&R Start & Finish or NearestSF Revised 4/18/2018 to address either/or
'
    Application.ScreenUpdating = False
    Range("P8").FormulaR1C1 = "=IF(ABS(R[-6]C-RC[-15])<R[1]C[-10],R[-6]C,RC[-15])"
    Range("Q8").FormulaR1C1 = "=IF(RC[-1]=RC[-16],TASKS!R[6]C[-13],R[-6]C)"
    Range("R8").FormulaR1C1 = "=IF(RC[-2]=RC[-17],TASKS!R[6]C[-13],R[-6]C)"
    Range("S8").FormulaR1C1 = "=IF(RC[-3]=RC[-18],TASKS!R[6]C[-11],R[-6]C)"
    Range("T8").FormulaR1C1 = "=IF(RC[-4]=RC[-19],TASKS!R[6]C[-11],R[-6]C)"
    Range("U8").FormulaR1C1 = "=6371*ACOS(SIN(RC[-4])*SIN(R[1]C[-4])+COS(RC[-4])*COS(R[1]C[-4])*COS(R[1]C[-3]-RC[-3]))"
    Range("P9:T10").Value = Range("P2:T3").Value
    Range("U9:U10").FormulaR1C1 = "=6371*ACOS(SIN(RC[-4])*SIN(R[1]C[-4])+COS(RC[-4])*COS(R[1]C[-4])*COS(R[1]C[-3]-RC[-3]))"
    ActiveSheet.Calculate
    Range("P11").FormulaR1C1 = "=IF(ABS(R[-7]C-R[-3]C[-14])<R[-2]C[-10],R[-3]C[-14],R[-7]C)"
    Range("Q11").FormulaR1C1 = "=IF(RC[-1]=R[-3]C[-15],TASKS!R[5]C[-13],R[-7]C)"
    Range("R11").FormulaR1C1 = "=IF(RC[-2]=R[-3]C[-16],TASKS!R[5]C[-13],R[-7]C)"
    Range("S11").FormulaR1C1 = "=IF(RC[-3]=R[-3]C[-17],TASKS!R[5]C[-11],Sheet2!R[-7]C)"
    Range("T11").FormulaR1C1 = "=IF(RC[-4]=R[-3]C[-19],TASKS!R[5]C[-11],R[-7]C)"
    Range("V11").FormulaR1C1 = "=SUM(R[-3]C[-1]:R[-1]C[-1])"
    Range("V12").FormulaR1C1 = "=IF(R[-4]C[-3]-R[-1]C[-3]<=R[-3]C[-18],R[-1]C,R[-1]C-((R[-4]C[-3]-R[-1]C[-3]-1000)*0.1))"
    Range("W12").FormulaR1C1 = _
        "=IF(OR(AND(RC[-1]>=750,MIN(R[-4]C[-2]:R[-2]C[-2])>=0.25*RC[-1],MAX(R[-4]C[-2]:R[-2]C[-2])<=0.45*RC[-1]),RC[-1]<750,MIN(R[-4]C[-2]:R[-2]C[-2])>=0.28*RC[-1]),""OK"","""")"
    ActiveSheet.Calculate
    Range("P8:W12").Value = Range("P8:W12").Value
    
    Sheets("Sheet2").Range("P14:T14").Value = Sheets("Tasks").Range("G40:K40").Value
    Range("P15:T16").Value = Range("P3:T4").Value
    Sheets("Sheet2").Range("P17:T17").Value = Sheets("Tasks").Range("G41:K41").Value
     
    Range("U14:U16").FormulaR1C1 = "=6371*ACOS(SIN(RC[-4])*SIN(R[1]C[-4])+COS(RC[-4])*COS(R[1]C[-4])*COS(R[1]C[-3]-RC[-3]))"
    Range("V17").FormulaR1C1 = "=SUM(R[-3]C[-1]:R[-1]C[-1])"
    Range("V18").FormulaR1C1 = "=IF(R[-4]C[-3]-R[-1]C[-3]<=R[-9]C[-18],R[-1]C,R[-1]C-((R[-4]C[-3]-R[-1]C[-3]-1000)*0.1))"
    Range("W18").FormulaR1C1 = _
        "=IF(OR(AND(RC[-1]>=750,MIN(R[-4]C[-2]:R[-2]C[-2])>=0.25*RC[-1],MAX(R[-4]C[-2]:R[-2]C[-2])<=0.45*RC[-1]),RC[-1]<750,MIN(R[-4]C[-2]:R[-2]C[-2])>=0.28*RC[-1]),""OK"","""")"
    ActiveSheet.Calculate
    Range("P14:W18").Value = Range("P14:W18").Value
    
    If Range("V12") > Range("V18") And Range("W12") = "OK" Then
      Range("A1").Value = "This is a 2-Turn Point Triangle"
      Range("A2:F5").Value = Range("P8:U11").Value
      Range("F6:F7").Value = Range("V11:V12").Value
    
    ElseIf Range("V18") > Range("V12") And Range("W18") = "OK" Then
      Range("A1").Value = "This is a 2-Turn Point Triangle"
      Range("A2:F5").Value = Range("P14:U17").Value
      Range("F6:F7").Value = Range("V17:V18").Value
    End If

    Range("W6").FormulaR1C1 = _
        "=IF(OR(AND(RC[-1]<750,MIN(R[-4]C[-2]:R[-2]C[-2])>=0.28*RC[-1]),AND(RC[-1]>=750,MIN(R[-4]C[-2]:R[-2]C[-2])>=0.25*RC[-1],MAX(R[-4]C[-2]:R[-2]C[-2])<=0.45*RC[-1])),""OK"","""")"
    Columns("G:O").Clear
    Range("P8:W18").Clear
    
    End Sub

Sub TriSFredux()
'
' Addresses TRI TP3 after O&R fini, but within 10 km
'
Application.ScreenUpdating = False
    Range("G8").FormulaR1C1 = "=MAX(R[2]C:R[10001]C)"
    Range("G10").FormulaR1C1 = "=IF(RC[-4]=R1C3,"""",IF(AND(RC[-6]>R4C1,6371*ACOS(SIN(RC[-5])*SIN(R1C2)+COS(RC[-5])*COS(R1C2)*COS(R1C3-RC[-4]))<10),RC[-6],""""))"
    Range("H10:K10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-6],"""")"
    'Copy Ref A Value Sort
    Application.Calculation = xlCalculationManual
 With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:K10").AutoFill Destination:=.Range("G10:K" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    If Range("G8") = 0 Then
        Range("G10").FormulaR1C1 = "=If(RC[-4]=R1C3,"""",IF(AND(RC[-6]>R4C1,6371*ACOS(SIN(RC[-5])*SIN(R1C2)+COS(RC[-5])*COS(R1C2)*COS(R1C3-RC[-4]))<20),RC[-6],""""))"
        'Copy Ref A Value Sort
        Application.Calculation = xlCalculationManual
    With Worksheets("Sheet2")
        LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
    .Range("G10:K10").AutoFill Destination:=.Range("G10:K" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    End If
    
    Range("G10:K10009").Value = Range("G10:K10009").Value
    Range("G10:K10009").Sort Key1:=Range("G10"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    Range("G10:K509").Copy
    Range("N1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    
    Range("G10:U10009").Clear
    
    Range("N6:SS6").FormulaR1C1 = "=IF(OR(R1C="""",R3C=R[-2]C3),"""",6371*ACOS(SIN(R2C)*SIN(R4C2)+COS(R2C)*COS(R4C2)*COS(R4C3-R3C)))"
    Range("N6:SS6").Value = Range("N6:SS6").Value
    Range("L6").FormulaR1C1 = "=MIN(RC[2]:RC[501])"
    Range("M6").FormulaR1C1 = "=MAX(RC[1]:RC[500])"
    Range("L6:M6").Value = Range("L6:M6").Value
    
    Range("G10").FormulaR1C1 = "=IF(RC[-4]=R4C3,"""",IF(AND(RC[-6]<R2C1,6371*ACOS(SIN(RC[-5])*SIN(R4C2)+COS(RC[-5])*COS(R4C2)*COS(R4C3-RC[-4]))>=R6C12,6371*ACOS(SIN(RC[-5])*SIN(R4C2)+COS(RC[-5])*COS(R4C2)*COS(R4C3-RC[-4]))<=R6C13),RC[-6],""""))"
    Range("H10:K10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-6],"""")"
    Range("L10").FormulaR1C1 = "=IF(RC[-1]<>"""",6371*ACOS(SIN(RC[-4])*SIN(R4C2)+COS(RC[-4])*COS(R4C2)*COS(R4C3-RC[-3])),"""")"
    'Copy Ref A Value Sort
    Application.Calculation = xlCalculationManual
 With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:L10").AutoFill Destination:=.Range("G10:L" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("G10:L10009").Value = Range("G10:L10009").Value
    Range("G10:L10009").Sort Key1:=Range("G10"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    
If Range("G10") = "" Then
    Range("A1").Value = "NONE IDENTIFIED"
ElseIf Range("G10") <> "" Then
    
    'MATRIX
    Range("N10:SS10").FormulaR1C1 = _
        "=IF(OR(R1C="""",R3C=RC9),"""",IF(OR(R6C<RC12,6371*ACOS(SIN(R2C)*SIN(RC8)+COS(R2C)*COS(RC8)*COS(RC9-R3C))>0.5,R6C<SQRT(RC12^2+0.25),R6C>R6C[1]),"""",IF(RC10-R4C<R9C4,R6C+RC12,R6C+RC12-((RC10-R4C-R9C4)*0.1))))"
    'COPY REF G
    Application.Calculation = xlCalculationManual
 With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
.Range("N10:SS10").AutoFill Destination:=.Range("N10:SS" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    
    Range("I5").FormulaR1C1 = "=MAX(R[5]C[5]:R[3004]C[504])"
    Range("I5").Value = Range("I5").Value
    
    Range("N10:SS10009").Value = Range("N10:SS10009").Value
    Range("N7:SS7").FormulaR1C1 = "=IF(MAX(R[3]C:R[10002]C)=R5C9,R[-6]C,"""")"
    Range("N7:SS7").Value = Range("N7:SS7").Value
    
    Range("M10").FormulaR1C1 = "=IF(MAX(RC[1]:RC[500])=R5C9,RC[-6],"""")"
    'Copy Ref G
     Application.Calculation = xlCalculationManual
 With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
.Range("M10:M10").AutoFill Destination:=.Range("M10:M" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("M10:M10009").Value = Range("M10:M10009").Value
    
    Range("N10:SS10009").Clear
    
    Range("N8:SS8").FormulaR1C1 = "=IF(R1C="""","""",IF(R[-1]C=R1C,R[-6]C,""""))"
    Range("N9:SS11").FormulaR1C1 = "=IF(R[-1]C<>"""",R[-6]C,"""")"
    
    Range("A5").FormulaR1C1 = "=MAX(R[2]C[13]:R[2]C[512])"
    Range("B5").FormulaR1C1 = "=MAX(R[3]C[12]:R[3]C[511])"
    Range("C5").FormulaR1C1 = "=MAX(R[4]C[11]:R[4]C[510])"
    Range("D5").FormulaR1C1 = "=MAX(R[5]C[10]:R[5]C[509])"
    Range("E5").FormulaR1C1 = "=MAX(R[6]C[9]:R[6]C[508])"
    Range("A5:E5").Value = Range("A5:E5").Value
    
    Range("N10:Q10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-6],"""")"
    'Copy REF G
    Application.Calculation = xlCalculationManual
 With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
.Range("N10:Q10").AutoFill Destination:=.Range("N10:Q" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("A1:E1").FormulaR1C1 = "=MAX(R[9]C[12]:R[10008]C[12])"
    Range("A1:E1").Value = Range("A1:E1").Value
End If
    Columns("G:SS").Clear
End Sub
Sub ReCK()
'
' ReCK Macro for 3 ORDS only!
'
    'ReCK TP2
    Columns("G:N").Clear
    Range("G10").FormulaR1C1 = "=IF(AND(RC[-6]>R2C16,RC[-6]<R4C16),RC[-6],"""")"
    Range("H10:K10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-6],"""")"
    
    Application.Calculation = xlCalculationManual
 With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:K10").AutoFill Destination:=.Range("G10:K" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("G10:K10009").Value = Range("G10:K10009").Value
    Range("G10:K10009").Sort Key1:=Range("G10"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    
    Range("L8:Q8").FormulaR1C1 = "=MAX(R[2]C:R[10001]C)"
    Range("L10").FormulaR1C1 = "=6371*ACOS(SIN(RC[-4])*SIN(R2C17)+COS(RC[-4])*COS(R2C17)*COS(R2C18-RC[-3]))+6371*ACOS(SIN(RC[-4])*SIN(R4C17)+COS(RC[-4])*COS(R4C17)*COS(R4C18-RC[-3]))"
    Range("M10").FormulaR1C1 = "=IF(RC[-1]=R8C12,RC[-6],"""")"
    Range("N10:Q10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-6],"""")"

 With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
.Range("L10:Q10").AutoFill Destination:=.Range("L10:Q" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("R8").FormulaR1C1 = "=6371*ACOS(SIN(RC[-4])*SIN(R[-6]C[-1])+COS(RC[-4])*COS(R[-6]C[-1])*COS(R[-6]C-RC[-3]))"
    Range("S8").FormulaR1C1 = "=6371*ACOS(SIN(RC[-5])*SIN(R[-4]C[-2])+COS(RC[-5])*COS(R[-4]C[-2])*COS(R[-4]C[-1]-RC[-4]))"
    Range("T8").FormulaR1C1 = "=SUM(RC[-8],R[-4]C[1])"
    Range("U8").FormulaR1C1 = "=IF(RC[-1]<750,0.28*RC[-1],0.25*RC[-1])"
    Range("V8").FormulaR1C1 = "=IF(RC[-2]>=750,0.25*RC[-2],RC[-2])"
    Range("W8").FormulaR1C1 = "=IF(AND(RC[-11]>=RC[-2],RC[-11]<=RC[-1]),""OK"","""")"
    Range("X8").FormulaR1C1 = "=IF(MIN(RC[-6],RC[-5],R[-4]C[-3])>=RC[-3],""OK"",""NOPE"")"
    ActiveSheet.Calculate
  If Range("X8") = "OK" Then
    Range("P3:T3").Value = Range("M8:Q8").Value

    Range("U2:U3").FormulaR1C1 = _
        "=6371*ACOS(SIN(R[1]C[-4])*SIN(RC[-4])+COS(R[1]C[-4])*COS(RC[-4])*COS(RC[-3]-R[1]C[-3]))"
    Range("U4").FormulaR1C1 = _
        "=6371*ACOS(SIN(RC[-4])*SIN(R[-2]C[-4])+COS(RC[-4])*COS(R[-2]C[-4])*COS(R[-2]C[-3]-RC[-3]))"
    Range("V5").FormulaR1C1 = "=SUM(R[-3]C[-1]:R[-1]C[-1])"
    Range("V6").FormulaR1C1 = "=IF(TASKS!R[8]C[-14]-MAX(TASKS!R[10]C[-14],TASKS!R[11]C[-14])>R9C4,R[-1]C-((TASKS!R[8]C[-14]-MAX(TASKS!R[10]C[-14],TASKS!R[11]C[-14])-R9C4)*0.1),R[-1]C)"
  End If
  
    'Re-CK TP3
    Range("A7").Value = 1.38888888888889E-03
    Range("G10:Q10009").Clear
    Range("G10").FormulaR1C1 = _
        "=IF(AND(RC[-6]>=R4C16-R7C1,RC[-6]<=R4C16+R7C1),RC[-6],"""")"
    Range("H10:K10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-6],"""")"
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:K10").AutoFill Destination:=.Range("G10:K" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("G10:K10009").Value = Range("G10:K10009").Value
    Range("G10:K10009").Sort Key1:=Range("G10"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    
    Range("L10").FormulaR1C1 = _
        "=6371*ACOS(SIN(RC[-4])*SIN(R2C17)+COS(RC[-4])*COS(R2C17)*COS(R2C18-RC[-3]))+6371*ACOS(SIN(RC[-4])*SIN(R3C17)+COS(RC[-4])*COS(R3C17)*COS(R3C18-RC[-3]))"
    Range("M10").FormulaR1C1 = "=IF(RC[-1]=R8C12,RC[-6],"""")"
    Range("N10:Q10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-6],"""")"
    
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
.Range("L10:Q10").AutoFill Destination:=.Range("L10:Q" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("R8").FormulaR1C1 = _
        "=6371*ACOS(SIN(RC[-4])*SIN(R[-6]C[-1])+COS(RC[-4])*COS(R[-6]C[-1])*COS(R[-6]C-RC[-3]))"
    Range("S8").FormulaR1C1 = _
        "=6371*ACOS(SIN(RC[-5])*SIN(R[-5]C[-2])+COS(RC[-5])*COS(R[-5]C[-2])*COS(R[-5]C[-1]-RC[-4]))"
    Range("T8").FormulaR1C1 = "=SUM(RC[-8],R[-6]C[1])"
    Range("U8").FormulaR1C1 = "=IF(RC[-1]<750,0.28*RC[-1],0.25*RC[-1])"
    Range("V8").FormulaR1C1 = "=IF(RC[-2]>=750,0.25*RC[-2],RC[-2])"
    Range("W8").FormulaR1C1 = _
        "=IF(AND(RC[-11]>=RC[-2],RC[-11]<=RC[-1]),""OK"","""")"
    Range("X8").FormulaR1C1 = _
        "=IF(MIN(RC[-6],RC[-5],R[-6]C[-3])>=RC[-3],""OK"",""NOPE"")"

  If Range("X8") = "OK" Then
    Range("P4:T4").Value = Range("M8:Q8").Value
    Range("U2:U3").FormulaR1C1 = _
        "=6371*ACOS(SIN(R[1]C[-4])*SIN(RC[-4])+COS(R[1]C[-4])*COS(RC[-4])*COS(RC[-3]-R[1]C[-3]))"
    Range("U4").FormulaR1C1 = _
        "=6371*ACOS(SIN(RC[-4])*SIN(R[-2]C[-4])+COS(RC[-4])*COS(R[-2]C[-4])*COS(R[-2]C[-3]-RC[-3]))"
    Range("V5").FormulaR1C1 = "=SUM(R[-3]C[-1]:R[-1]C[-1])"
    Range("V6").FormulaR1C1 = "=IF(TASKS!R[8]C[-14]-MAX(TASKS!R[10]C[-14],TASKS!R[11]C[-14])>R9C4,R[-1]C-((TASKS!R[8]C[-14]-MAX(TASKS!R[10]C[-14],TASKS!R[11]C[-14])-R9C4)*0.1),R[-1]C)"
  End If
  
  'Re-CK TP1
  Range("G10:Q10009").Clear
    Range("G10").FormulaR1C1 = _
        "=IF(AND(RC[-6]>=R2C16-R7C1,RC[-6]<=R2C16+R7C1),RC[-6],"""")"
    Range("H10:K10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-6],"""")"
    
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:K10").AutoFill Destination:=.Range("G10:K" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("G10:K10009").Value = Range("G10:K10009").Value
    Range("G10:K10009").Sort Key1:=Range("G10"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    
    Range("L10").FormulaR1C1 = _
        "=6371*ACOS(SIN(RC[-4])*SIN(R4C17)+COS(RC[-4])*COS(R4C17)*COS(R4C18-RC[-3]))+6371*ACOS(SIN(RC[-4])*SIN(R3C17)+COS(RC[-4])*COS(R3C17)*COS(R3C18-RC[-3]))"
    Range("M10").FormulaR1C1 = "=IF(RC[-1]=R8C12,RC[-6],"""")"
    Range("N10:Q10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-6],"""")"
    
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
.Range("L10:Q10").AutoFill Destination:=.Range("L10:Q" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("R8").FormulaR1C1 = "=6371*ACOS(SIN(RC[-4])*SIN(R[-5]C[-1])+COS(RC[-4])*COS(R[-5]C[-1])*COS(R[-5]C-RC[-3]))"
    Range("S8").FormulaR1C1 = "=6371*ACOS(SIN(RC[-5])*SIN(R[-4]C[-2])+COS(RC[-5])*COS(R[-4]C[-2])*COS(R[-4]C[-1]-RC[-4]))"
    Range("T8").FormulaR1C1 = "=SUM(RC[-8],R[-5]C[1])"
    Range("U8").FormulaR1C1 = "=IF(RC[-1]<750,0.28*RC[-1],0.25*RC[-1])"
    Range("V8").FormulaR1C1 = "=IF(RC[-2]>=750,0.25*RC[-2],RC[-2])"
    Range("W8").FormulaR1C1 = "=IF(AND(RC[-11]>=RC[-2],RC[-11]<=RC[-1]),""OK"","""")"
    Range("X8").FormulaR1C1 = "=IF(MIN(RC[-6],RC[-5],R[-5]C[-3])>=RC[-3],""OK"",""NOPE"")"
    
  If Range("X8") = "OK" Then
    Range("P2:T2").Value = Range("M8:Q8").Value
    Range("U2:U3").FormulaR1C1 = _
        "=6371*ACOS(SIN(R[1]C[-4])*SIN(RC[-4])+COS(R[1]C[-4])*COS(RC[-4])*COS(RC[-3]-R[1]C[-3]))"
    Range("U4").FormulaR1C1 = _
        "=6371*ACOS(SIN(RC[-4])*SIN(R[-2]C[-4])+COS(RC[-4])*COS(R[-2]C[-4])*COS(R[-2]C[-3]-RC[-3]))"
    Range("V5").FormulaR1C1 = "=SUM(R[-3]C[-1]:R[-1]C[-1])"
    Range("V6").FormulaR1C1 = "=IF(TASKS!R[8]C[-14]-MAX(TASKS!R[10]C[-14],TASKS!R[11]C[-14])>R9C4,R[-1]C-((TASKS!R[8]C[-14]-MAX(TASKS!R[10]C[-14],TASKS!R[11]C[-14])-R9C4)*0.1),R[-1]C)"
    Range("S8:X8,G8:R10009").Clear
    Range("U2:X6").Value = Range("U2:X6").Value
  End If
  
End Sub
Sub TriOrds()
'
' Works for Leonard,Mueller, Ramy06 w/macro noted
'
Application.ScreenUpdating = False
    Sheets("Sheet2").Range("A1:E4").Value = Sheets("Tasks").Range("A40:E44").Value
    Range("F1").FormulaR1C1 = "=6371*ACOS(SIN(RC[-4])*SIN(R[1]C[-4])+COS(RC[-4])*COS(R[1]C[-4])*COS(R[1]C[-3]-RC[-3]))"
    Range("G2").FormulaR1C1 = "=6371*ACOS(SIN(RC[-5])*SIN(R[1]C[-5])+COS(RC[-5])*COS(R[1]C[-5])*COS(R[1]C[-4]-RC[-4]))"
    Range("H3").FormulaR1C1 = "=6371*ACOS(SIN(R[-2]C[-6])*SIN(RC[-6])+COS(R[-2]C[-6])*COS(RC[-6])*COS(RC[-5]-R[-2]C[-5]))"
    Range("H4").FormulaR1C1 = "=6371*ACOS(SIN(R[-2]C[-6])*SIN(RC[-6])+COS(R[-2]C[-6])*COS(RC[-6])*COS(RC[-5]-R[-2]C[-5]))"
    Range("H5").FormulaR1C1 = "=6371*ACOS(SIN(R[-2]C[-6])*SIN(R[-1]C[-6])+COS(R[-2]C[-6])*COS(R[-1]C[-6])*COS(R[-1]C[-5]-R[-2]C[-5]))"
    Range("J4").FormulaR1C1 = "=6371*ACOS(SIN(RC[-8])*SIN(R[-3]C[-8])+COS(RC[-8])*COS(R[-3]C[-8])*COS(R[-3]C[-7]-RC[-7]))"
    Range("L1").FormulaR1C1 = "=RC[-6]+R[1]C[-5]+R[2]C[-4]"
    Range("M1").FormulaR1C1 = "=IF(OR(AND(RC[-1]>=750,MIN(RC[-7],R[1]C[-6],R[2]C[-5])>=0.25*RC[-1],MAX(RC[-7],R[1]C[-6],R[2]C[-5])<=0.45*RC[-1]),AND(RC[-1]<750,MIN(RC[-7],R[1]C[-6],R[2]C[-5])>=0.28*RC[-1])),""OK"",""NOPE"")"
    Range("L2").FormulaR1C1 = "=R[-1]C[-6]+R[2]C[-4]+R[2]C[-2]"
    Range("M2").FormulaR1C1 = "=IF(OR(AND(RC[-1]>=750,MIN(R[-1]C[-7],R[2]C[-5],R[2]C[-3])>=0.25*RC[-1],MAX(R[-1]C[-7],R[2]C[-5],R[2]C[-3])<=0.45*RC[-1]),AND(RC[-1]<750,MIN(R[-1]C[-7],R[2]C[-5],R[2]C[-3])>=0.28*RC[-1])),""OK"",""NOPE"")"
    Range("L3").FormulaR1C1 = "=RC[-4]+R[2]C[-4]+R[1]C[-2]"
    Range("M3").FormulaR1C1 = "=IF(OR(AND(RC[-1]>=750,MIN(RC[-5],R[2]C[-5],R[1]C[-3])>=0.25*RC[-1],MAX(RC[-5],R[2]C[-5],R[1]C[-3])<=0.45*RC[-1]),AND(RC[-1]<750,MIN(RC[-5],R[2]C[-5],R[1]C[-3])>=0.28*RC[-1])),""OK"",""NOPE"")"
    Range("L4").FormulaR1C1 = "=R[-2]C[-5]+R[1]C[-4]+RC[-4]"
    Range("M4").FormulaR1C1 = "=IF(OR(AND(RC[-1]>=750,MIN(R[-2]C[-6],R[1]C[-5],RC[-3])>=0.25*RC[-1],MAX(R[-2]C[-6],R[1]C[-5],RC[-3])<=0.45*RC[-1]),AND(RC[-1]<750,MIN(R[-2]C[-6],R[1]C[-5],RC[-3])>=0.28*RC[-1])),""OK"",""NOPE"")"
    Range("N1:N4").FormulaR1C1 = "=IF(RC[-1]=""OK"",RC[-2],"""")"
    Range("N5").FormulaR1C1 = "=IF(SUM(R[-4]C:R[-1]C)=0,""NOPE"",IF(R[-4]C=MAX(R[-4]C:R[-1]C),123,IF(R[-3]C=MAX(R[-4]C:R[-1]C),124,IF(R[-2]C=MAX(R[-4]C:R[-1]C),134,IF(R[-1]C=MAX(R[-4]C:R[-1]C),234)))))"
    ActiveSheet.Calculate
    If Range("N5") = 123 Then
        Range("P2:T4").Value = Range("A1:E3").Value
    ElseIf Range("N5") = 124 Then
        Range("P2:T3").Value = Range("A1:E2").Value
        Range("P4:T4").Value = Range("A4:E4").Value
    ElseIf Range("N5") = 134 Then
        Range("P2:T2").Value = Range("A1:E1").Value
        Range("P3:T4").Value = Range("A3:E4").Value
    ElseIf Range("N5") = 234 Then
        Range("P2:T4").Value = Range("A2:E4").Value
    ElseIf Range("N5") = "NOPE" Then
    'Find longest possible triangle based on shortest leg/applicable FAI minimum
        '123 tri
        Range("O1").FormulaR1C1 = "=IF(RC[-3]<750,MIN(RC[-9],R[1]C[-8],R[2]C[-7])/0.28,MIN(RC[-9],R[1]C[-8],R[2]C[-7])/0.25)"
        '124 tri
        Range("O2").FormulaR1C1 = "=IF(RC[-3]<750,MIN(R[-1]C[-9],R[2]C[-7],R[2]C[-5])/0.28,MIN(R[-1]C[-9],R[2]C[-7],R[2]C[-5])/0.25)"
        '134 tri
        Range("O3").FormulaR1C1 = "=IF(RC[-3]<750,MIN(RC[-7],R[2]C[-7],R[1]C[-5])/0.28,MIN(RC[-7],R[2]C[-7],R[1]C[-5])/0.25)"
        '234 tri
        Range("O4").FormulaR1C1 = "=IF(RC[-3]<750,MIN(R[-2]C[-8],RC[-7],R[1]C[-7])/0.28,MIN(R[-2]C[-8],RC[-7],R[1]C[-7])/0.25)"
        Range("O1:O4").Value = Range("O1:O4").Value
        'Branch here for Tri2 vs TriOFC2, followed if needed by TriONE; amended for 300k limit end 10/1/14
      If Range("O1") = Range("O2") And Range("O3") = Range("O4") Then
        Application.Run "F.xlsm!Tri2Ords"
      ElseIf Range("O1") = Range("O2") And Range("O1") > Range("O3") And Range("O1") > Range("O4") And Range("O1") < 300 Then
        Application.Run "F.xlsm!Tri2TP"
      ElseIf Range("O1") = Range("O4") And Range("O2") = Range("O3") Then
        Application.Run "F.xlsm!Tri2Ordsb"
      ElseIf Range("O1") <> Range("O2") Or Range("O3") <> Range("O4") Then
        'HERE?
        Application.Run "F.xlsm!TriONEord"
      End If
    End If
    
    If Range("N5") <> "NOPE" And Range("N5") <> "" Then
        Range("U2:U3").FormulaR1C1 = "=6371*ACOS(SIN(RC[-4])*SIN(R[1]C[-4])+COS(RC[-4])*COS(R[1]C[-4])*COS(R[1]C[-3]-RC[-3]))"
        Range("U4").FormulaR1C1 = "=6371*ACOS(SIN(RC[-4])*SIN(R[-2]C[-4])+COS(RC[-4])*COS(R[-2]C[-4])*COS(R[-2]C[-3]-RC[-3]))"
        Range("V5").FormulaR1C1 = "=SUM(R[-3]C[-1]:R[-1]C[-1])"
        Range("V6").FormulaR1C1 = "=IF(TASKS!R[8]C[-14]-MAX(TASKS!R[10]C[-14],TASKS!R[11]C[-14])>R9C4,R[-1]C-((TASKS!R[8]C[-14]-MAX(TASKS!R[10]C[-14],TASKS!R[11]C[-14])-R9C4)*0.1),R[-1]C)"
        Range("W6").FormulaR1C1 = "=IF(OR(AND(RC[-1]<750,MIN(R[-4]C[-2]:R[-2]C[-2])>=0.28*RC[-1]),AND(RC[-1]>=750,MIN(R[-4]C[-2]:R[-2]C[-2])>=0.25*RC[-1],MAX(R[-4]C[-2]:R[-2]C[-2])<=0.45*RC[-1])),""OK"","""")"
        Application.Run "F.xlsm!ReCK"
    End If
    
    If Range("W6") = "OK" And Range("P1") <> "This is a 2-Turn Point Triangle" And Range("P2") < Range("A8") Then
        Application.Run "F.xlsm!Ramy06"
    End If
    
    If Range("W6") = "OK" And Range("V6") > Range("F7") Then
        Range("A1:E5").Value = Range("P1:T5").Value
        Range("F2:F4").Value = Range("U2:U4").Value
        Range("F6:F7").Value = Range("V5:V6").Value
    End If
    Range("F1,G1:W6").Clear
End Sub
Sub Tri2tp()
'
' Tri2tp Macro Antares 2-TP triangle using longest leg Ords 1/2; shortest leg Ords 2/3
'
    'If Range("O1") = Range("O2") And Range("O1") > Range("O3") And Range("O1") > Range("O4") And Range("O1") < 300 Then
    Range("P1").Value = "This is a 2-Turn Point Triangle"
    Range("P2:U3").Cut Destination:=Range("P3:U4")
    Sheets("Sheet2").Range("P2:R2").Value = Sheets("TASKS").Range("C14:E14").Value
    Sheets("Sheet2").Range("S2:T2").Value = Sheets("TASKS").Range("H14:I14").Value
    Sheets("Sheet2").Range("P5:R5").Value = Sheets("TASKS").Range("C16:E16").Value
    Sheets("Sheet2").Range("S5:T5").Value = Sheets("TASKS").Range("H16:I16").Value
    
    Range("T6").Clear
    Range("U2:U4").FormulaR1C1 = "=6371*ACOS(SIN(RC[-4])*SIN(R[1]C[-4])+COS(RC[-4])*COS(R[1]C[-4])*COS(R[1]C[-3]-RC[-3]))"
   
    Range("V5").FormulaR1C1 = "=SUM(R[-3]C[-1]:R[-1]C[-1])"
    Range("V6").FormulaR1C1 = _
        "=IF(TASKS!R[8]C[-14]-MAX(TASKS!R[10]C[-14],TASKS!R[11]C[-14])>R9C4,Sheet2!R[-1]C-((TASKS!R[8]C[-14]-MAX(TASKS!R[10]C[-14],TASKS!R[11]C[-14])-R9C4)*0.1),Sheet2!R[-1]C)"
    Range("W6").FormulaR1C1 = "=IF(AND(RC[-1]<750,MIN(R[-4]C[-2]:R[-2]C[-2])>0.28*RC[-1]),""OK"","""")"
    
    Range("G10").FormulaR1C1 = "=IF(AND(RC[-6]<R5C16,RC[-6]>R4C16),RC[-6],"""")"
    Range("H10:K10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-6],"""")"
    'Copy Ref A Value Sort
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:K10").AutoFill Destination:=.Range("G10:K" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("G10:K10009").Value = Range("G10:K10009").Value
    Range("G10:K10009").Sort Key1:=Range("G10"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
     
    Range("L10").FormulaR1C1 = "=R2C21"
    Range("M10").FormulaR1C1 = "=IF(RC[-4]<>R3C18,6371*ACOS(SIN(RC[-5])*SIN(R3C17)+COS(RC[-5])*COS(R3C17)*COS(R3C18-RC[-4])),"""")"
    Range("N10").FormulaR1C1 = "=IF(RC[-5]<>R5C18,6371*ACOS(SIN(RC[-6])*SIN(R5C17)+COS(RC[-6])*COS(R5C17)*COS(R5C18-RC[-5])),"""")"
    Range("O10").FormulaR1C1 = "=SUM(RC[-3]:RC[-1])"
    Range("P10").FormulaR1C1 = "=IF(MIN(RC[-4]:RC[-2])>=0.28*RC[-1],RC[-1],"""")"
    Range("Q10").FormulaR1C1 = "=IF(RC[-1]=R8C16,RC[-10],"""")"
    Range("R10:U10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-10],"""")"
    Range("P8:U8").FormulaR1C1 = "=MAX(R[2]C:R[10001]C)"
    'Copy Ref G
    Application.Calculation = xlCalculationManual
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
.Range("L10:U10").AutoFill Destination:=.Range("L10:U" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    If Range("Q8") <> 0 Then
        Range("P4:T4").Value = Range("Q8:U8").Value
    End If
    
    If Range("Q8") = 0 Or Range("W6") <> "OK" Then
        Sheets("Sheet2").Range("P2:T2").Value = Sheets("TASKS").Range("G40:K40").Value
        Sheets("Sheet2").Range("P5:T5").Value = Sheets("TASKS").Range("G41:K41").Value
        If Range("Q8") <> 0 Then
             Range("P4:T4").Value = Range("Q8:U8").Value
        End If
    End If

 Range("G8:U10009").Clear
End Sub

Sub Tri2Ords()
'
' Tri2Ords Macro works for Payne 09 NO reCK
'
    Range("A7").Value = 1.38888888888889E-03
    
    'IF cks ORDS 3 & 4, ElseIf cks ORDS 1 & 2
    
    If Range("O4") = Range("O3") And Range("O3") > Range("O2") Then
        Range("P2:T3").Value = Range("A3:E4").Value
    ElseIf Range("O1") = Range("O2") And Range("O2") > Range("O3") Then
        Range("P2:T3").Value = Range("A1:E2").Value
    End If
    
    Range("U2").FormulaR1C1 = "=6371*ACOS(SIN(R[1]C[-4])*SIN(RC[-4])+COS(R[1]C[-4])*COS(RC[-4])*COS(RC[-3]-R[1]C[-3]))"
    Range("L8:Q8").FormulaR1C1 = "=MAX(R[2]C:R[10001]C)"
    Range("G10").FormulaR1C1 = "=IF(AND(RC[-6]>=R2C16-R7C1,RC[-6]<=R2C16+R7C1),RC[-6],"""")"
    Range("H10:K10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-6],"""")"
    Range("L10").FormulaR1C1 = "=IF(RC[-1]<>"""",6371*ACOS(SIN(RC[-10])*SIN(R3C17)+COS(RC[-10])*COS(R3C17)*COS(R3C18-RC[-9])),"""")"
    Range("M10").FormulaR1C1 = "=IF(RC[-1]=R8C12,RC[-6],"""")"
    Range("N10:Q10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-6],"""")"
    'Copy Ref A
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:Q10").AutoFill Destination:=.Range("G10:Q" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
  If Range("L8") > Range("U2") Then
    Range("P2:T2").Value = Range("M8:Q8").Value
  End If
    
    Range("G10").FormulaR1C1 = "=IF(AND(RC[-6]>=R3C16-R7C1,RC[-6]<=R3C16+R7C1),RC[-6],"""")"
    Range("L10").FormulaR1C1 = "=IF(RC[-1]<>"""",6371*ACOS(SIN(RC[-10])*SIN(R2C17)+COS(RC[-10])*COS(R2C17)*COS(R2C18-RC[-9])),"""")"
    'Copy Ref A
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
.Range("G10:L10").AutoFill Destination:=.Range("G10:L" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
  If Range("L8") > Range("U2") Then
    Range("P3:T3").Value = Range("M8:Q8").Value
  End If
    Range("G8:Q10009").Clear
 
    Range("G10").FormulaR1C1 = "=IF(RC[-6]>R3C16,RC[-6],"""")"
    Range("H10:K10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-6],"""")"
    'Copy Ref A Value Sort
     With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:K10").AutoFill Destination:=.Range("G10:K" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("G10:K10009").Value = Range("G10:K10009").Value
    Range("G10:K10009").Sort Key1:=Range("G10"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
  
If Range("G10") = "" Then
    Exit Sub
ElseIf Range("G10") <> "" Then
    Range("O8:T8").FormulaR1C1 = "=MAX(R[2]C:R[10001]C)"
    
    Range("L10").FormulaR1C1 = "=6371*ACOS(SIN(RC[-4])*SIN(R3C17)+COS(RC[-4])*COS(R3C17)*COS(R3C18-RC[-3]))"
    Range("M10").FormulaR1C1 = "=6371*ACOS(SIN(RC[-5])*SIN(R2C17)+COS(RC[-5])*COS(R2C17)*COS(R2C18-RC[-4]))"
    Range("N10").FormulaR1C1 = "=SUM(RC[-2],RC[-1],R2C21)"
    Range("O10").FormulaR1C1 = _
        "=IF(OR(AND(RC[-1]<750,MIN(RC[-3],RC[-2],R2C21)>=0.28*RC[-1]),AND(RC[-1]>=750,MIN(RC[-3],RC[-2],R2C21)>=0.25*RC[-1],MAX(RC[-3],RC[-2],R2C21)<=0.45*RC[-1])),RC[-1],"""")"
    Range("P10").FormulaR1C1 = "=IF(RC[-1]=R8C15,RC[-9],"""")"
    Range("Q10:T10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-9],"""")"
    'Copy Ref G
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
.Range("L10:T10").AutoFill Destination:=.Range("L10:T" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("P4:T4").Value = Range("P8:T8").Value
        
    Range("U2:U3").FormulaR1C1 = _
        "=6371*ACOS(SIN(R[1]C[-4])*SIN(RC[-4])+COS(R[1]C[-4])*COS(RC[-4])*COS(RC[-3]-R[1]C[-3]))"
    Range("U4").FormulaR1C1 = "=6371*ACOS(SIN(RC[-4])*SIN(R[-2]C[-4])+COS(RC[-4])*COS(R[-2]C[-4])*COS(R[-2]C[-3]-RC[-3]))"
    Range("V5").FormulaR1C1 = "=SUM(R[-3]C[-1]:R[-1]C[-1])"
    Range("V6").FormulaR1C1 = _
        "=IF(TASKS!R[8]C[-14]-MAX(TASKS!R[10]C[-14],TASKS!R[11]C[-14])>R9C4,Sheet2!R[-1]C-((TASKS!R[8]C[-14]-MAX(TASKS!R[10]C[-14],TASKS!R[11]C[-14])-R9C4)*0.1),Sheet2!R[-1]C)"
    Range("W6").FormulaR1C1 = _
        "=IF(OR(AND(RC[-1]>=750,MIN(R[-4]C[-2]:R[-2]C[-2])>=0.25*RC[-1],MAX(R[-4]C[-2]:R[-2]C[-2])<=0.45*RC[-1]),AND(RC[-1]<750,MIN(R[-4]C[-2]:R[-2]C[-2])>=0.28*RC[-1])),""OK"","""")"

 If Range("W6") <> "" Then
        Application.Run "F.xlsm!TriGEO"
 End If

    If Range("W6") = "OK" Then
        Range("L1:O5").Value = Range("L1:O5").Value
        Range("G10:N10009").Clear
        Range("A2:F4").Value = Range("P2:U4").Value
        Range("F6:F7").Value = Range("V5:V6").Value
        Range("E6").Value = Range("X6").Value
        Range("F1,A7:B8").Clear

        Sheets("Sheet2").Range("A1:C1").Value = Sheets("TASKS").Range("C14:E14").Value
        Sheets("Sheet2").Range("E1:D1").Value = Sheets("TASKS").Range("H14:I14").Value
        Sheets("Sheet2").Range("A5:C5").Value = Sheets("TASKS").Range("C16:E16").Value
        Sheets("Sheet2").Range("D5:E5").Value = Sheets("TASKS").Range("H16:I16").Value
    End If
    
    If Range("O3") = Range("O4") And Range("O3") > Range("O2") Then
        Application.Run "F.xlsm!Tri2Ordsa"
    ElseIf Range("O1") > Range("O2") And Range("O2") > Range("O3") And Range("O3") <= Range("O4") Then
        Application.Run "F.xlsm!TriONEord"
    End If
    
        
End If
End Sub
Sub TriGEO()
'
' SOMETHING Macro
'
Application.ScreenUpdating = False
    Sheets("YDWK3").Range("E39").Value = Sheets("Sheet2").Range("Q2").Value
    Sheets("YDWK3").Range("E40").Value = Sheets("Sheet2").Range("R2").Value
    Sheets("YDWK3").Range("E41").Value = Sheets("Sheet2").Range("Q3").Value
    Sheets("YDWK3").Range("E42").Value = Sheets("Sheet2").Range("R3").Value
    Sheets("YDWK3").Calculate
    Sheets("Sheet2").Range("W2").Value = Sheets("YDWK3").Range("F37").Value
    
    Sheets("YDWK3").Range("E39").Value = Sheets("Sheet2").Range("Q4").Value
    Sheets("YDWK3").Range("E40").Value = Sheets("Sheet2").Range("R4").Value
    Sheets("YDWK3").Calculate
    Sheets("Sheet2").Range("W3").Value = Sheets("YDWK3").Range("F37").Value
    
    Sheets("YDWK3").Range("E41").Value = Sheets("Sheet2").Range("Q2").Value
    Sheets("YDWK3").Range("E42").Value = Sheets("Sheet2").Range("R2").Value
    Sheets("YDWK3").Calculate
    Sheets("Sheet2").Range("W4").Value = Sheets("YDWK3").Range("F37").Value
    
    Range("X5").FormulaR1C1 = "=SUM(R[-3]C[-1]:R[-1]C[-1])"
    Range("X6").FormulaR1C1 = _
        "=IF(TASKS!R[8]C[-16]-MAX(TASKS!R[10]C[-16],TASKS!R[11]C[-16])>R9C4,Sheet2!R[-1]C-((TASKS!R[8]C[-16]-MAX(TASKS!R[10]C[-16],TASKS!R[11]C[-16])-R9C4)*0.1),Sheet2!R[-1]C)"
    Range("Y6").FormulaR1C1 = _
        "=IF(OR(AND(RC[-1]>=750,MIN(R[-4]C[-2]:R[-2]C[-2])>=0.25*RC[-1],MAX(R[-4]C[-2]:R[-2]C[-2])<=0.45*RC[-1]),AND(RC[-1]<750,MIN(R[-4]C[-2]:R[-2]C[-2])>=0.28*RC[-1])),""OK"","""")"
    
  If Range("Y6") = "" Then
    
    Range("H6").FormulaR1C1 = _
        "=IF(AND(RC[17]="""",OR(AND(RC[16]<750,MIN(R[-4]C[15]:R[-2]C[15])<0.28*RC[16]),AND(RC[16]>=750,MIN(R[-4]C[15]:R[-2]C[15])<0.25*RC[16]))),MIN(R[-4]C[13]:R[-2]C[13])-MIN(R[-4]C[15]:R[-2]C[15]),"""")"
    Range("H7").FormulaR1C1 = _
        "=IF(AND(R[-1]C<>"""",R[-5]C[15]=MAX(R[-5]C[15]:R[-3]C[15])),R[-5]C[8],IF(AND(R[-1]C<>"""",R[-4]C[15]=MAX(R[-5]C[15]:R[-3]C[15])),R[-4]C[8],IF(AND(R[-1]C<>"""",R[-3]C[15]=MAX(R[-5]C[15]:R[-3]C[15])),R[-3]C[8],"""")))"
    Range("H8").FormulaR1C1 = "=IF(R[-2]C<>"""",R[-2]C[16]/R[-2]C[14],"""")"
    
    Range("H6:H8").Value = Range("H6:H8").Value
    
    Range("A7").Value = 2.08333333333333E-03
    Range("G10").FormulaR1C1 = _
        "=IF(OR(AND(R7C8=R4C16,RC[-6]>=R4C16-R7C1,RC[-6]<=R4C16+R7C1),AND(R7C8=R3C16,RC[-6]>=R3C16-R7C1,RC[-6]<=R3C16+R7C1),AND(R7C8=R2C16,RC[-6]>=R2C16-R7C1,RC[-6]<=R2C16+R7C1)),RC[-6],"""")"
    Range("H10:K10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-6],"""")"
    'Copy Ref A Value Sort
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:K10").AutoFill Destination:=.Range("G10:K" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("G10:K10009").Value = Range("G10:K10009").Value
    Range("G10:K10009").Sort Key1:=Range("G10"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
   
    If Range("H7") = Range("P4") Then
        Range("L10").FormulaR1C1 = "=R2C21"
        Range("M10").FormulaR1C1 = "=6371*ACOS(SIN(RC[-5])*SIN(R3C17)+COS(RC[-5])*COS(R3C17)*COS(R3C18-RC[-4]))"
        Range("N10").FormulaR1C1 = "=6371*ACOS(SIN(RC[-6])*SIN(R2C17)+COS(RC[-6])*COS(R2C17)*COS(R2C18-RC[-5]))"
        Range("O10").FormulaR1C1 = "=SUM(RC[-3]:RC[-1])"
        Range("P10").FormulaR1C1 = "=IF(OR(AND(RC[-1]<750,MIN(RC[-4]:RC[-2])>=0.28*R8C8*RC[-1]),AND(RC[-1]>=750,MIN(RC[-4]:RC[-2])>=0.25*R8C8*RC[-1],MAX(RC[-4]:RC[-2])<=0.45*R8C8*RC[-1])),R8C8*RC[-1],"""")"
    'ElseIf(s) for H7 = P3, H7=P2
    End If
    
    Range("P8:U8").FormulaR1C1 = "=MAX(R[2]C:R[10001]C)"
    Range("Q10").FormulaR1C1 = "=IF(RC[-1]=R8C16,RC[-10],"""")"
    Range("R10:U10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-10],"""")"
    'Copy Ref G
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
.Range("L10:U10").AutoFill Destination:=.Range("L10:U" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    If Range("H7") = Range("P4") Then
        Range("P4:T4").Value = Range("Q8:U8").Value
    ElseIf Range("H7") = Range("P3") Then
        Range("P3:T3").Value = Range("Q8:U8").Value
    ElseIf Range("H7") = Range("P2") Then
        Range("P2:T2").Value = Range("Q8:U8").Value
    End If
  
    Sheets("YDWK3").Range("E39").Value = Sheets("Sheet2").Range("Q2").Value
    Sheets("YDWK3").Range("E40").Value = Sheets("Sheet2").Range("R2").Value
    Sheets("YDWK3").Range("E41").Value = Sheets("Sheet2").Range("Q3").Value
    Sheets("YDWK3").Range("E42").Value = Sheets("Sheet2").Range("R3").Value
    Sheets("YDWK3").Calculate
    Sheets("Sheet2").Range("W2").Value = Sheets("YDWK3").Range("F37").Value
    
    Sheets("YDWK3").Range("E39").Value = Sheets("Sheet2").Range("Q4").Value
    Sheets("YDWK3").Range("E40").Value = Sheets("Sheet2").Range("R4").Value
    Sheets("YDWK3").Calculate
    Sheets("Sheet2").Range("W3").Value = Sheets("YDWK3").Range("F37").Value
    
    Sheets("YDWK3").Range("E41").Value = Sheets("Sheet2").Range("Q2").Value
    Sheets("YDWK3").Range("E42").Value = Sheets("Sheet2").Range("R2").Value
    Sheets("YDWK3").Calculate
    Sheets("Sheet2").Range("W4").Value = Sheets("YDWK3").Range("F37").Value
    
    Range("X5").FormulaR1C1 = "=SUM(R[-3]C[-1]:R[-1]C[-1])"
    Range("X6").FormulaR1C1 = _
        "=IF(TASKS!R[8]C[-16]-MAX(TASKS!R[10]C[-16],TASKS!R[11]C[-16])>R9C4,Sheet2!R[-1]C-((TASKS!R[8]C[-16]-MAX(TASKS!R[10]C[-16],TASKS!R[11]C[-16])-R9C4)*0.1),Sheet2!R[-1]C)"
    Range("Y6").FormulaR1C1 = _
        "=IF(OR(AND(RC[-1]>=750,MIN(R[-4]C[-2]:R[-2]C[-2])>=0.25*RC[-1],MAX(R[-4]C[-2]:R[-2]C[-2])<=0.45*RC[-1]),AND(RC[-1]<750,MIN(R[-4]C[-2]:R[-2]C[-2])>=0.28*RC[-1])),""OK"","""")"
 End If
 
 If Range("Y6") = "" Then
    Range("G6:U10009").Clear
    
    Range("A7").FormulaR1C1 = "=MAX(TASKS!R[9]C[2],TASKS!R[10]C[2])"
    Range("A7").Value = Range("A7").Value
    
    Range("G10").FormulaR1C1 = "=IF(AND(RC[-6]>R3C16,RC[-6]<R7C1),RC[-6],"""")"
    Range("H10:K10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-6],"""")"
    'Copy Ref A Value Sort
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:K10").AutoFill Destination:=.Range("G10:K" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("G10:K10009").Value = Range("G10:K10009").Value
    Range("G10:K10009").Sort Key1:=Range("G10"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
  
If Range("G10") = "" Then
    Exit Sub
ElseIf Range("G10") <> "" Then
    Range("O8:T8").FormulaR1C1 = "=MAX(R[2]C:R[10001]C)"
    
    Range("L10").FormulaR1C1 = "=6371*ACOS(SIN(RC[-4])*SIN(R3C17)+COS(RC[-4])*COS(R3C17)*COS(R3C18-RC[-3]))"
    Range("M10").FormulaR1C1 = "=6371*ACOS(SIN(RC[-5])*SIN(R2C17)+COS(RC[-5])*COS(R2C17)*COS(R2C18-RC[-4]))"
    Range("N10").FormulaR1C1 = "=SUM(RC[-2],RC[-1],R2C21)"
    Range("O10").FormulaR1C1 = _
        "=IF(OR(AND(RC[-1]<750,MIN(RC[-3],RC[-2],R2C21)>=0.28*RC[-1]),AND(RC[-1]>=750,MIN(RC[-3],RC[-2],R2C21)>=0.25*RC[-1],MAX(RC[-3],RC[-2],R2C21)<=0.45*RC[-1])),RC[-1],"""")"
    Range("P10").FormulaR1C1 = "=IF(RC[-1]=R8C15,RC[-9],"""")"
    Range("Q10:T10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-9],"""")"
    'Copy Ref G
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
.Range("L10:T10").AutoFill Destination:=.Range("L10:T" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("P4:T4").Value = Range("P8:T8").Value
        
    Range("U2:U3").FormulaR1C1 = _
        "=6371*ACOS(SIN(R[1]C[-4])*SIN(RC[-4])+COS(R[1]C[-4])*COS(RC[-4])*COS(RC[-3]-R[1]C[-3]))"
    Range("U4").FormulaR1C1 = "=6371*ACOS(SIN(RC[-4])*SIN(R[-2]C[-4])+COS(RC[-4])*COS(R[-2]C[-4])*COS(R[-2]C[-3]-RC[-3]))"
    Range("V5").FormulaR1C1 = "=SUM(R[-3]C[-1]:R[-1]C[-1])"
    Range("V6").FormulaR1C1 = _
        "=IF(TASKS!R[8]C[-14]-MAX(TASKS!R[10]C[-14],TASKS!R[11]C[-14])>R9C4,Sheet2!R[-1]C-((TASKS!R[8]C[-14]-MAX(TASKS!R[10]C[-14],TASKS!R[11]C[-14])-R9C4)*0.1),Sheet2!R[-1]C)"
    Range("W6").FormulaR1C1 = _
        "=IF(OR(AND(RC[-1]>=750,MIN(R[-4]C[-2]:R[-2]C[-2])>=0.25*RC[-1],MAX(R[-4]C[-2]:R[-2]C[-2])<=0.45*RC[-1]),AND(RC[-1]<750,MIN(R[-4]C[-2]:R[-2]C[-2])>=0.28*RC[-1])),""OK"","""")"

    If Range("W6") = "OK" Then
        Range("L1:O5").Value = Range("L1:O5").Value
        Range("G10:N10009").Clear
        Range("A2:F4").Value = Range("P2:U4").Value
        Range("F6:F7").Value = Range("V5:V6").Value
        Range("E6").Value = Range("W6").Value
        Range("F1,A7:B8,P2:X7").Clear

        Sheets("Sheet2").Range("A1:C1").Value = Sheets("TASKS").Range("C14:E14").Value
        Sheets("Sheet2").Range("E1:D1").Value = Sheets("TASKS").Range("H14:I14").Value
        Sheets("Sheet2").Range("A5:C5").Value = Sheets("TASKS").Range("C16:E16").Value
        Sheets("Sheet2").Range("D5:E5").Value = Sheets("TASKS").Range("H16:I16").Value
    End If
  End If
 End If
 
End Sub
Sub Tri2Ordsa()
'
' double-checks ORDS 3/4 after Tri2Ords
'
    Sheets("Sheet2").Range("A8").Value = Sheets("TASKS").Range("C14").Value
    Sheets("Sheet2").Range("B8").Value = Sheets("TASKS").Range("C16").Value
    Range("P3:T4").Value = Range("A2:E3").Value
    Range("U3").FormulaR1C1 = "=6371*ACOS(SIN(RC[-4])*SIN(R[1]C[-4])+COS(RC[-4])*COS(R[1]C[-4])*COS(R[1]C[-3]-RC[-3]))"
    
    Range("G10").FormulaR1C1 = "=IF(AND(RC[-6]<R3C16,RC[-6]>=R8C1),RC[-6],"""")"
    Range("H10:K10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-6],"""")"
    'Copy Ref A Value Sort
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:K10").AutoFill Destination:=.Range("G10:K" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("G10:K10009").Value = Range("G10:K10009").Value
    Range("G10:K10009").Sort Key1:=Range("G10"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    
  If Range("G10") <> "" Then
    
    Range("N8:S8").FormulaR1C1 = "=MAX(R[2]C:R[10001]C)"
    Range("L10").FormulaR1C1 = "=6371*ACOS(SIN(RC[-4])*SIN(R3C17)+COS(RC[-4])*COS(R3C17)*COS(R3C18-RC[-3]))"
    Range("M10").FormulaR1C1 = "=6371*ACOS(SIN(RC[-5])*SIN(R4C17)+COS(RC[-5])*COS(R4C17)*COS(R4C18-RC[-4]))"
    Range("N10").FormulaR1C1 = _
        "=IF(OR(AND(SUM(RC[-2],RC[-1],R3C21)<750,MIN(RC[-2],RC[-1],R3C21)>=0.28*SUM(RC[-2],RC[-1],R3C21)),AND(SUM(RC[-2],RC[-1],R3C21)>=750,MIN(RC[-2],RC[-1],R3C21)>=0.25*SUM(RC[-2],RC[-1],R3C21),MAX(RC[-2],RC[-1],R3C21)<=0.45*SUM(RC[-2],RC[-1],R3C21))),SUM(RC[-2],RC[-1],R3C21),"""")"
    Range("O10").FormulaR1C1 = "=IF(RC[-1]=R8C14,RC[-8],"""")"
    Range("P10:S10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-8],"""")"
    'Copy Ref G
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
.Range("L10:S10").AutoFill Destination:=.Range("L10:S" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("P2:T2").Value = Range("O8:S8").Value
    
    Range("U2:U3").FormulaR1C1 = "=6371*ACOS(SIN(RC[-4])*SIN(R[1]C[-4])+COS(RC[-4])*COS(R[1]C[-4])*COS(R[1]C[-3]-RC[-3]))"
    Range("U4").FormulaR1C1 = "=6371*ACOS(SIN(RC[-4])*SIN(R[-2]C[-4])+COS(RC[-4])*COS(R[-2]C[-4])*COS(R[-2]C[-3]-RC[-3]))"
    Range("V5").FormulaR1C1 = "=SUM(R[-3]C[-1]:R[-1]C[-1])"
    Range("V6").FormulaR1C1 = _
        "=IF(TASKS!R[8]C[-14]-MAX(TASKS!R[10]C[-14],TASKS!R[11]C[-14])>R9C4,Sheet2!R[-1]C-((TASKS!R[8]C[-14]-MAX(TASKS!R[10]C[-14],TASKS!R[11]C[-14])-R9C4)*0.1),Sheet2!R[-1]C)"
    Range("W6").FormulaR1C1 = _
        "=IF(OR(AND(RC[-1]<750,MIN(R[-4]C[-2]:R[-2]C[-2])>=0.28*RC[-1]),AND(RC[-1]>=750,MIN(R[-4]C[-2]:R[-2]C[-2])>=0.25*RC[-1],MAX(R[-4]C[-2]:R[-2]C[-2])<=0.45*RC[-1])),""OK"","""")"
    
    Range("G10").FormulaR1C1 = "=IF(AND(RC[-6]>R2C16,RC[-6]<R4C16),RC[-6],"""")"
    Range("H10:K10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-6],"""")"
    'Copy Ref A Value Sort
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:K10").AutoFill Destination:=.Range("G10:K" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("G10:K10009").Value = Range("G10:K10009").Value
    Range("G10:K10009").Sort Key1:=Range("G10"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    
    Range("L10").FormulaR1C1 = "=6371*ACOS(SIN(RC[-4])*SIN(R2C17)+COS(RC[-4])*COS(R2C17)*COS(R2C18-RC[-3]))"
    Range("M10").FormulaR1C1 = "=6371*ACOS(SIN(RC[-5])*SIN(R4C17)+COS(RC[-5])*COS(R4C17)*COS(R4C18-RC[-4]))"
    Range("N10").FormulaR1C1 = "=IF(OR(AND(SUM(RC[-2],RC[-1],R4C21)<750,MIN(RC[-2],RC[-1],R4C21)>=0.28*SUM(RC[-2],RC[-1],R4C21)),AND(SUM(RC[-2],RC[-1],R4C21)>=750,MIN(RC[-2],RC[-1],R4C21)>=0.25*SUM(RC[-2],RC[-1],R4C21),MAX(RC[-2],RC[-1],R4C21)<=0.45*SUM(RC[-2],RC[-1],R4C21))),SUM(RC[-2],RC[-1],R4C21),"""")"
    'Copy Ref G
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
.Range("L10:S10").AutoFill Destination:=.Range("L10:S" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
  If Range("N8") > Range("V6") Then
    Range("P3:T3").Value = Range("O8:S8").Value
  End If
    
    Range("G10").FormulaR1C1 = "=IF(RC[-6]>R4C16,RC[-6],"""")"
    Range("H10:K10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-6],"""")"
    'Copy Ref A Value Sort
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:K10").AutoFill Destination:=.Range("G10:K" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("G10:K10009").Value = Range("G10:K10009").Value
    Range("G10:K10009").Sort Key1:=Range("G10"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    
    Range("L10").FormulaR1C1 = "=6371*ACOS(SIN(RC[-4])*SIN(R2C17)+COS(RC[-4])*COS(R2C17)*COS(R2C18-RC[-3]))"
    Range("M10").FormulaR1C1 = "=6371*ACOS(SIN(RC[-5])*SIN(R3C17)+COS(RC[-5])*COS(R3C17)*COS(R3C18-RC[-4]))"
    Range("N10").FormulaR1C1 = _
        "=IF(OR(AND(SUM(RC[-2],RC[-1],R2C21)<750,MIN(RC[-2],RC[-1],R2C21)>=0.28*SUM(RC[-2],RC[-1],R2C21)),AND(SUM(RC[-2],RC[-1],R2C21)>=750,MIN(RC[-2],RC[-1],R2C21)>=0.25*SUM(RC[-2],RC[-1],R2C21),MAX(RC[-2],RC[-1],R2C21)<=0.45*SUM(RC[-2],RC[-1],R2C21))),SUM(RC[-2],RC[-1],R2C21),"""")"
    'Copy Ref G
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
.Range("L10:S10").AutoFill Destination:=.Range("L10:S" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
  If Range("N8") > Range("V6") Then
    Range("P4:T4").Value = Range("O8:S8").Value
  End If
    
  If Range("V6") > Range("F7") Then
    Range("A2:F4").Value = Range("P2:U4").Value
    Range("F6:F7").Value = Range("V5:V6").Value
  End If
  End If
Columns("G:W").Clear
End Sub
Sub Tri2Ordsb()
'
' Cks for Ords 2&3 LEWIS, future expansion for 1&4
'
Application.ScreenUpdating = False
'Check 1&4 as two Ords (If W4<V4, 1&4 not possible)
    Range("V4").FormulaR1C1 = "=R[-2]C[-7]*0.28"
    Range("W4").FormulaR1C1 = "=R[-2]C[-8]-(RC[-2]+RC[-1])"
 If Range("W4") < Range("V4") Then
 'Check 2&3 as two Ords
    Range("P3:T3").Value = Range("A3:E3").Value
    Range("U2").FormulaR1C1 = "=6371*ACOS(SIN(RC[-4])*SIN(R[1]C[-4])+COS(RC[-4])*COS(R[1]C[-4])*COS(R[1]C[-3]-RC[-3]))"
    Range("V2").FormulaR1C1 = "=RC[-1]/0.28"
 'ElseIf (future?)for Ords 1 & 4
 End If
    'Determine 2 vs 3 TP
If Range("V2") < 300 Then
    Range("A7").Value = 1.04166666666667E-02
    'START CANDIDATES for a 2-TP Triangle
    Range("G10").FormulaR1C1 = "=IF(RC[-6]<R2C16,RC[-6],"""")"
    Range("H10:K10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-6],"""")"
    'CopyRef A Value SORT
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:K10").AutoFill Destination:=.Range("G10:K" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("G10:K10009").Value = Range("G10:K10009").Value
    Range("G10:K10009").Sort Key1:=Range("G10"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
  
  If Range("G10") = "" Then
    Application.Run "F.xlsm!TriONEord"
  ElseIf Range("G10") <> "" Then
  
    Range("L10").FormulaR1C1 = "=6371*ACOS(SIN(RC[-4])*SIN(R2C17)+COS(RC[-4])*COS(R2C17)*COS(R2C18-RC[-3]))"
    Range("M10").FormulaR1C1 = "=R2C21"
    Range("N10").FormulaR1C1 = "=6371*ACOS(SIN(RC[-6])*SIN(R3C17)+COS(RC[-6])*COS(R3C17)*COS(R3C18-RC[-5]))"
    Range("O10").FormulaR1C1 = "=SUM(RC[-3]:RC[-1])"
    Range("P10").FormulaR1C1 = "=IF(AND(RC[-9]>=R8C1-R7C1,RC[-9]<=R8C1+R7C1,OR(AND(RC[-1]<=750,MIN(RC[-4]:RC[-2])>=0.28*RC[-1]),AND(RC[-1]>750,MIN(RC[-4]:RC[-2])>=0.25*RC[-1],MAX(RC[-4]:RC[-2])<=0.45*RC[-1]))),RC[-1],"""")"
    Range("Q10:U10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-10],"""")"
    'Copy Ref K Value
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "K").End(xlUp).Row
.Range("L10:U10").AutoFill Destination:=.Range("L10:U" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("P10:U10009").Value = Range("P10:U10009").Value
    Range("P10:U10009").Sort Key1:=Range("Q10"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    If Range("P10") = "" Then
        Range("A1").Value = "NONE IDENTIFIED"
        Exit Sub
    End If
    Range("G10:O10009").Clear
    
    Range("V10").FormulaR1C1 = "=6371*ACOS(SIN(RC[-4])*SIN(R3C17)+COS(RC[-4])*COS(R3C17)*COS(R3C18-RC[-3]))"
    Range("W10").FormulaR1C1 = "=IF(RC[-7]>100,R9C4,IF(R9C5<>""PR"",10*RC[-7],10*RC[-7]-100))"
    'Copy Ref U VALUE
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "U").End(xlUp).Row
.Range("V10:W10").AutoFill Destination:=.Range("V10:W" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("V10:W10009").Value = Range("V10:W10009").Value
    Range("V7").FormulaR1C1 = "=MIN(R[3]C:R[10002]C)"
    Range("V8").FormulaR1C1 = "=MAX(R[2]C:R[10001]C)"
    'FINISH CANDIDATES
    Range("Y10").FormulaR1C1 = _
        "=IF(AND(RC[-24]>R3C16,6371*ACOS(SIN(RC[-23])*SIN(R3C17)+COS(RC[-23])*COS(R3C17)*COS(R3C18-RC[-22]))>R7C22,6371*ACOS(SIN(RC[-23])*SIN(R3C17)+COS(RC[-23])*COS(R3C17)*COS(R3C18-RC[-22]))<R8C22),RC[-24],"""")"
    Range("Z10:AC10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-24],"""")"
    Range("AD10").FormulaR1C1 = "=IF(RC[-1]<>"""",6371*ACOS(SIN(RC[-4])*SIN(R3C17)+COS(RC[-4])*COS(R3C17)*COS(R3C18-RC[-3])),"""")"
    'Copy Ref A Value Sort
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("Y10:AD10").AutoFill Destination:=.Range("Y10:AD" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("Y10:AD10009").Value = Range("Y10:AD10009").Value
    Range("Y10:AD10009").Sort Key1:=Range("Y10"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
     
    If Range("Y510") <> "" Then
        Range("G10:L509").Value = Range("Y510:AD1009").Value
        Range("Y510:AD1009").Clear
    End If
    
    Range("Y10:AD509").Copy
    Range("Y1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    Range("Y10:AD10009").Clear
    
    'MATRIX
    Range("Y10:TD10").FormulaR1C1 = _
        "=IF(OR(R1C="""",RC19=R3C),"""",IF(OR(R6C<RC22,6371*ACOS(SIN(RC18)*SIN(R2C)+COS(RC18)*COS(R2C)*COS(R3C-RC19))>0.5,R6C>SQRT(RC22^2+0.25)),"""",IF(RC20-R4C<=RC23,RC16+R6C,IF(AND(RC16>100,RC20-R4C>RC23),RC16+R6C-((RC20-R4C-RC23)*0.1),0))))"
    'Copy Ref P
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "P").End(xlUp).Row
.Range("Y10:TD10").AutoFill Destination:=.Range("Y10:TD" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("Y10:TD3000").Value = Range("Y10:TD3000").Value
    Range("T5").FormulaR1C1 = "=MAX(R[5]C[5]:R[2995]C[504])"
    Range("T5").Value = Range("T5").Value

    If Range("T5") = 0 Then
        Columns("Y:TD").Clear
        Range("G10:L509").Copy
        Range("Y1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
        'MATRIX
        Range("Y10:TD10").FormulaR1C1 = _
            "=IF(OR(R1C="""",RC19=R3C),"""",IF(OR(R6C<RC22,6371*ACOS(SIN(RC18)*SIN(R2C)+COS(RC18)*COS(R2C)*COS(R3C-RC19))>0.5,R6C>SQRT(RC22^2+0.25)),"""",IF(RC20-R4C<=RC23,RC16+R6C,IF(AND(RC16>100,RC20-R4C>RC23),RC16+R6C-((RC20-R4C-RC23)*0.1),0))))"
         'Copy Ref P
        With Worksheets("Sheet2")
        LastRow = .Cells(Rows.Count, "P").End(xlUp).Row
    .Range("Y10:TD10").AutoFill Destination:=.Range("Y10:TD" & LastRow), Type:=xlFillDefault
        End With
        ActiveSheet.Calculate
        'If T5 0 again, clear Columns N:TD, go None Identified
    End If
    
    Range("Y10:TD3000").Value = Range("Y10:TD3000").Value
    Range("T5").FormulaR1C1 = "=MAX(R[5]C[5]:R[2995]C[504])"
    Range("T5").Value = Range("T5").Value
    
    Range("Y7:TD7").FormulaR1C1 = "=IF(MAX(R[3]C:R[2993]C)=R5C20,R[-6]C,"""")"
    Range("Y7:TD7").Value = Range("Y7:TD7").Value

    Range("X10").FormulaR1C1 = "=IF(MAX(RC[1]:RC[500])=R5C20,RC[-7],"""")"
    'Copy Ref W
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "W").End(xlUp).Row
.Range("X10:X10").AutoFill Destination:=.Range("X10:X" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("X10:X3000").Value = Range("X10:X3000").Value
    
    Range("P4:T4").Value = Range("P3:T3").Value
    Range("P3:T3").Value = Range("P2:T2").Value
    Range("P1").Value = "This is a 2-Turn Point Triangle"
    
    Range("Y10:AD10009").Clear
    Range("Y10:AB10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-7],"""")"
    'Copy Ref X
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "X").End(xlUp).Row
.Range("Y10:AB10").AutoFill Destination:=.Range("Y10:AB" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("P2:T2").FormulaR1C1 = "=MAX(R[8]C[8]:R[10007]C[8])"
    Range("P2:T2").Value = Range("P2:T2").Value
    
    Range("Y8:TD8").FormulaR1C1 = "=IF(R[-1]C<>"""",R[-6]C,"""")"
    Range("Y9:TD9").FormulaR1C1 = "=IF(R[-1]C<>"""",R[-6]C,"""")"
    Range("Y10:TD10").FormulaR1C1 = "=IF(R[-1]C<>"""",R[-6]C,"""")"
    Range("Y11:TD11").FormulaR1C1 = "=IF(R[-1]C<>"""",R[-6]C,"""")"
    
    Range("P5").FormulaR1C1 = "=MAX(R[2]C[9]:R[2]C[508])"
    Range("Q5").FormulaR1C1 = "=MAX(R[3]C[8]:R[3]C[507])"
    Range("R5").FormulaR1C1 = "=MAX(R[4]C[7]:R[4]C[506])"
    Range("S5").FormulaR1C1 = "=MAX(R[5]C[6]:R[5]C[505])"
    Range("T5").FormulaR1C1 = "=MAX(R[6]C[5]:R[6]C[504])"
    
    Range("P5:T5").Value = Range("P5:T5").Value
    
    Range("U2:U4").FormulaR1C1 = "=6371*ACOS(SIN(RC[-4])*SIN(R[1]C[-4])+COS(RC[-4])*COS(R[1]C[-4])*COS(R[1]C[-3]-RC[-3]))"
    Range("V5").FormulaR1C1 = "=SUM(R[-3]C[-1]:R[-1]C[-1])"
    Range("V6").FormulaR1C1 = _
        "=IF(OR(AND(R[-1]C>100,R[-4]C[-3]-R[-1]C[-3]<=R9C4),AND(R[-1]C<=100,R9C5<>""PR"",R[-4]C[-3]-R[-1]C[-3]<=10*R[-1]C),AND(R[-1]C<=100,R[-4]C[-3]-R[-1]C[-3]<=10*R[-1]C-100)),R[-1]C,IF(AND(R[-1]C>100,R[-4]C[-3]-R[-1]C[-3]>R9C4),R[-1]C-((R[-4]C[-3]-R[-1]C[-3]-R9C4)*0.1),0))"
    Range("W6").FormulaR1C1 = _
        "=IF(OR(AND(RC[-1]<750,MIN(R[-4]C[-2]:R[-2]C[-2])>=0.28*RC[-1]),AND(RC[-1]>=750,MIN(R[-4]C[-2]:R[-2]C[-2])>=0.25*RC[-1],MAX(R[-4]C[-2]:R[-2]C[-2])<=0.45*RC[-1])),""OK"",""NOPE"")"
    Columns("Y:TD").Clear
  End If
ElseIf Range("V2") >= 300 Then
    Range("G10").FormulaR1C1 = "=IF(RC[-6]>R3C16,RC[-6],"""")"
    Range("H10:K10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-6],"""")"
    'Copy Ref A Value Sort
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:K10").AutoFill Destination:=.Range("G10:K" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("G10:K10009").Value = Range("G10:K10009").Value
    Range("G10:K10009").Sort Key1:=Range("G10"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
  
  If Range("G10") = "" Then
    Application.Run "F.xlsm!TriONEord"
  ElseIf Range("G10") <> "" Then
  
    Range("L10").FormulaR1C1 = "=6371*ACOS(SIN(RC[-4])*SIN(R3C17)+COS(RC[-4])*COS(R3C17)*COS(R3C18-RC[-3]))"
    Range("M10").FormulaR1C1 = "=6371*ACOS(SIN(RC[-5])*SIN(R2C17)+COS(RC[-5])*COS(R2C17)*COS(R2C18-RC[-4]))"
    Range("N10").FormulaR1C1 = "=R2C21"
    Range("O10").FormulaR1C1 = "=SUM(RC[-3]:RC[-1])"
    'Copy Ref G Value
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
.Range("L10:O10").AutoFill Destination:=.Range("L10:O" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("L10:O10009").Value = Range("L10:O10009").Value
    
    Range("P8:U8").FormulaR1C1 = "=MAX(R[2]C:R[10001]C)"
    Range("P10").FormulaR1C1 = _
        "=IF(OR(AND(RC[-1]<750,MIN(RC[-4]:RC[-2])>=0.28*RC[-1]),AND(RC[-1]>=750,MIN(RC[-4]:RC[-2])>=0.25*RC[-1],MAX(RC[-4]:RC[-2])<=0.45*RC[-1])),RC[-1],"""")"
    Range("Q10").FormulaR1C1 = "=IF(RC[-1]=R8C16,RC[-10],"""")"
    Range("R10:U10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-10],"""")"
    'Copy Ref G
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
.Range("P10:U10").AutoFill Destination:=.Range("P10:U" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("P4:T4").Value = Range("Q8:U8").Value
    
    Range("U2:U3").FormulaR1C1 = "=6371*ACOS(SIN(RC[-4])*SIN(R[1]C[-4])+COS(RC[-4])*COS(R[1]C[-4])*COS(R[1]C[-3]-RC[-3]))"
    Range("U4").FormulaR1C1 = "=6371*ACOS(SIN(RC[-4])*SIN(R[-2]C[-4])+COS(RC[-4])*COS(R[-2]C[-4])*COS(R[-2]C[-3]-RC[-3]))"
    Range("V5").FormulaR1C1 = "=SUM(R[-3]C[-1]:R[-1]C[-1])"
    Range("V6").FormulaR1C1 = _
        "=IF(TASKS!R[8]C[-14]-MAX(TASKS!R[10]C[-14]:R[11]C[-14])<R9C4,R[-1]C,R[-1]C-((TASKS!R[8]C[-14]-MAX(TASKS!R[10]C[-14]:R[11]C[-14])-R9C4)*0.1))"
    Range("W6").FormulaR1C1 = _
        "=IF(OR(AND(RC[-1]<750,MIN(R[-4]C[-2]:R[-2]C[-2])>=0.28*RC[-1]),AND(RC[-1]>=750,MIN(R[-4]C[-2]:R[-2]C[-2])>=0.25*RC[-1],MAX(R[-4]C[-2]:R[-2]C[-2])<=0.45*RC[-1])),""OK"","""")"
   End If
End If
 Range("G7:W10009").Clear

End Sub
Sub TriONEord()
'
'Revised 3/21/14 First Part uses OFC as basis for TP 2
'
Application.ScreenUpdating = False
    Range("G10").FormulaR1C1 = "=IF(AND(RC[-6]>R4C16,RC[-6]<R8C2),RC[-6],"""")"
    Range("H10:K10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-6],"""")"
    'Copy Ref A Value Sort
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:K10").AutoFill Destination:=.Range("G10:K" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("G10:K10009").Value = Range("G10:K10009").Value
    Range("G10:K10009").Sort Key1:=Range("G10"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
  
If Range("G10") <> "" Then
    Range("P8:U8").FormulaR1C1 = "=MAX(R[2]C:R[10001]C)"
    Range("L10").FormulaR1C1 = "=6371*ACOS(SIN(RC[-4])*SIN(R4C17)+COS(RC[-4])*COS(R4C17)*COS(R4C18-RC[-3]))"
    Range("M10").FormulaR1C1 = "=6371*ACOS(SIN(RC[-5])*SIN(R3C17)+COS(RC[-5])*COS(R3C17)*COS(R3C18-RC[-4]))"
    Range("N10").FormulaR1C1 = "=R3C[7]"
    Range("O10").FormulaR1C1 = "=SUM(RC[-3],RC[-2],RC[-1])"
    Range("P10").FormulaR1C1 = _
        "=IF(OR(AND(RC[-1]<750,MIN(RC[-4]:RC[-2])>=0.28*RC[-1]),AND(RC[-1]>=750,MIN(RC[-4]:RC[-2])>=0.25*RC[-1],MAX(RC[-4]:RC[-2])<=0.45*RC[-1])),RC[-1],"""")"
    Range("Q10").FormulaR1C1 = "=IF(RC[-1]=R8C16,RC[-10],"""")"
    Range("R10:U10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-10],"""")"
    'Copy Ref G
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
.Range("L10:U10").AutoFill Destination:=.Range("L10:U" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
  If Range("P8") = 0 Then
    Range("G8:U10009").Clear
    Application.Run "F.xlsm!Triangle2"
    
  ElseIf Range("P8") <> 0 Then
    Range("P8:U8").Value = Range("P8:U8").Value
    Range("G10:U10009").Clear
    Range("V7").FormulaR1C1 = "=6371*ACOS(SIN(R[-4]C[-5])*SIN(R[-3]C[-5])+COS(R[-4]C[-5])*COS(R[-3]C[-5])*COS(R[-3]C[-4]-R[-4]C[-4]))"
    Range("V8").FormulaR1C1 = "=6371*ACOS(SIN(RC[-4])*SIN(R4C17)+COS(RC[-4])*COS(R4C17)*COS(R4C18-RC[-3]))"
    Range("V9").FormulaR1C1 = "=6371*ACOS(SIN(R[-1]C[-4])*SIN(R[-6]C[-5])+COS(R[-1]C[-4])*COS(R[-6]C[-5])*COS(R[-6]C[-4]-R[-1]C[-3]))"
    Range("V10").FormulaR1C1 = "=SUM(R[-3]C:R[-1]C)"
    Range("V11").FormulaR1C1 = "=IF(Tasks!R[3]C[-14]-Tasks!R[6]C[-14]<R9C4,R[-1]C,R[-1]C-((Tasks!R[3]C[-14]-Tasks!R[6]C[-14]-R9C4)*0.1))"
    Range("W11").FormulaR1C1 = _
        "=IF(OR(AND(RC[-1]<750,MIN(R[-4]C[-1]:R[-2]C[-1])>=0.28*RC[-1]),AND(RC[-1]>=750,MIN(R[-4]C[-1]:R[-2]C[-1])>=0.25*R[-4]C[-1],MAX(R[-4]C[-1]:R[-2]C[-1])<=0.45*R[-4]C[-1])),""OK"","""")"
    ActiveSheet.Calculate
    Range("V7:W11").Value = Range("V7:W11").Value
    
    If Range("W11") = "OK" Then
        Range("U2:W6").Clear
        Range("P2:T2").Value = Range("P3:T3").Value
        Range("P3:T3").Value = Range("P4:T4").Value
        Range("P4:T4").Value = Range("Q8:U8").Value
        Range("U2:U4").Value = Range("V7:V9").Value
        Range("V5:W6").Value = Range("V10:W11").Value
    End If
    Range("G7:W10009").Clear
  End If
 
 If Range("W6") <> "OK" Then
    'Check base Triangle1 TP1&3 +/-15 Mins, with TP2 at ORD2 (works for Keene)EG: O1= MAX(O1:O4)
    Range("A7").Value = 1.04166666666667E-02
    Range("G10").FormulaR1C1 = "=IF(AND(RC[-6]>=R2C16-R7C1,RC[-6]<=R2C16+R7C1),RC[-6],"""")"
    Range("H10:K10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-6],"""")"
    Range("L10").FormulaR1C1 = "=IF(RC[-1]<>"""",6371*ACOS(SIN(RC[-4])*SIN(R3C17)+COS(RC[-4])*COS(R3C17)*COS(R3C18-RC[-3])),"""")"
    'Copy Ref A Value Sort
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:L10").AutoFill Destination:=.Range("G10:L" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("G10:L10009").Value = Range("G10:L10009").Value
    Range("G10:L10009").Sort Key1:=Range("G10"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    
    Range("N10").FormulaR1C1 = "=IF(AND(RC[-13]>=R4C16-R7C1,RC[-13]<=R4C16+R7C1),RC[-13],"""")"
    Range("O10:R10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-13],"""")"
    Range("S10").FormulaR1C1 = "=IF(RC[-1]<>"""",6371*ACOS(SIN(RC[-4])*SIN(R3C17)+COS(RC[-4])*COS(R3C17)*COS(R3C18-RC[-3])),"""")"
    'Copy Ref A Value Sort
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("N10:S10").AutoFill Destination:=.Range("N10:S" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("N10:S10009").Value = Range("N10:S10009").Value
    Range("N10:S10009").Sort Key1:=Range("N10"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    
    Range("N10:S500").Copy
    Range("Z1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    Range("N10:S10009").Clear
   'Matrix
    Range("Z10:TE10").FormulaR1C1 = _
        "=IF(OR(AND(RC12+R6C+6371*ACOS(SIN(R2C)*SIN(RC8)+COS(R2C)*COS(RC8)*COS(RC9-R3C))<750,MIN(RC12,R6C,6371*ACOS(SIN(R2C)*SIN(RC8)+COS(R2C)*COS(RC8)*COS(RC9-R3C)))>=0.28*(RC12+R6C+6371*ACOS(SIN(R2C)*SIN(RC8)+COS(R2C)*COS(RC8)*COS(RC9-R3C)))),AND(RC12+R6C+6371*ACOS(SIN(R2C)*SIN(RC8)+COS(R2C)*COS(RC8)*COS(RC9-R3C))>=750,MIN(RC12,R6C,6371*ACOS(SIN(R2C)*SIN(RC8)+COS(R2C)*COS(RC8)*COS(RC9-R3C)))>=0.25*(RC12+R6C+6371*ACOS(SIN(R2C)*SIN(RC8)+COS(R2C)*COS(RC8)*COS(RC9-R3C))),MAX(RC12,R6C,6371*ACOS(SIN(R2C)*SIN(RC8)+COS(R2C)*COS(RC8)*COS(RC9-R3C)))<=0.45*(RC12+R6C+6371*ACOS(SIN(R2C)*SIN(RC8)+COS(R2C)*COS(RC8)*COS(RC9-R3C))))),SUM(RC12,R6C,6371*ACOS(SIN(R2C)*SIN(RC8)+COS(R2C)*COS(RC8)*COS(RC9-R3C))),"""")"
    'COPY ref G Value
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
.Range("Z10:TE10").AutoFill Destination:=.Range("Z10:TE" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("Z10:TE10009").Value = Range("Z10:TE10009").Value
    
    Range("X5").FormulaR1C1 = "=MAX(R[5]C[2]:R[495]C[501])"
    Range("X5").Value = Range("X5").Value

  If Range("X5") <> 0 Then
    Range("P10:P509").FormulaR1C1 = "=IF(MAX(RC[10]:RC[509])=R5C24,RC[-9],"""")"
    Range("Z7:TE7").FormulaR1C1 = "=IF(MAX(R[3]C:R[504]C)=R5C24,R[-6]C,"""")"

    Range("P7").FormulaR1C1 = "=MAX(R[3]C:R[502]C)"
    Range("P8").FormulaR1C1 = "=MAX(R[-1]C[10]:R[-1]C[509])"
    Range("P7:P8").Value = Range("P7:P8").Value
    Range("P10:TE509").Clear

    Range("P10").FormulaR1C1 = "=IF(OR(RC[-15]=R7C16,RC[-15]=R8C16),RC[-15],"""")"
    Range("Q10:T10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-15],"""")"
    'Copy Ref A Value Sort
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("P10:T10").AutoFill Destination:=.Range("P10:T" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("P10:T10009").Value = Range("P10:T10009").Value
    Range("P10:T10009").Sort Key1:=Range("P10"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    
    Range("P2:T2").Value = Range("P10:T10").Value
    Range("P4:T4").Value = Range("P11:T11").Value
    
    Range("U2:U3").FormulaR1C1 = "=6371*ACOS(SIN(RC[-4])*SIN(R[1]C[-4])+COS(RC[-4])*COS(R[1]C[-4])*COS(R[1]C[-3]-RC[-3]))"
    Range("U4").FormulaR1C1 = "=6371*ACOS(SIN(RC[-4])*SIN(R[-2]C[-4])+COS(RC[-4])*COS(R[-2]C[-4])*COS(R[-2]C[-3]-RC[-3]))"
    Range("V5").FormulaR1C1 = "=SUM(R[-3]C[-1]:R[-1]C[-1])"
    Range("V6").FormulaR1C1 = _
        "=IF(TASKS!R[8]C[-14]-MAX(TASKS!R[10]C[-14],TASKS!R[11]C[-14])>R9C4,R[-1]C-((TASKS!R[8]C[-14]-MAX(TASKS!R[10]C[-14],TASKS!R[11]C[-14])-R9C4)*0.1),R[-1]C)"
    Range("W6").FormulaR1C1 = _
        "=IF(OR(AND(RC[-1]<750,MIN(R[-4]C[-2]:R[-2]C[-2])>=0.28*RC[-1]),AND(RC[-1]>=750,MIN(R[-4]C[-2]:R[-2]C[-2])>=0.25*RC[-1],MAX(R[-4]C[-2]:R[-2]C[-2])<=0.45*RC[-1])),""OK"","""")"
   End If
 End If
    If Range("W6") <> "OK" Then
    Range("A1:W8").Clear
    Range("A1").Value = "NONE IDENTIFIED"
    End If
    Columns("Z:TE").Clear
    Range("G7:T10009").Clear
    
ElseIf Range("G10") = "" Then
    Range("A1:F7").Clear
    Range("A1").Value = "NONE IDENTIFIED"
End If
    
End Sub
Sub Triangle2()
'
' Triangle2 IN TRIORDS? When OneOrd fails, and O3 = O4 and O3 = MAX Os (Essex)
'
Application.ScreenUpdating = False
If Range("O3") = Range("O4") And Range("O3") > Range("O1") And Range("O3") > Range("O2") Then
    'amended 7/15/14 R4C1 instead of R4C16
    Range("G10").FormulaR1C1 = "=IF(AND(RC[-6]<R4C1,6371*ACOS(SIN(RC[-5])*SIN(R3C2)+COS(RC[-5])*COS(R3C2)*COS(R3C3-RC[-4]))<=10),RC[-6],"""")"
    Range("H10:K10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-6],"""")"
    Range("L10").FormulaR1C1 = "=IF(RC[-1]<>"""",6371*ACOS(SIN(RC[-10])*SIN(R4C2)+COS(RC[-10])*COS(R4C2)*COS(R4C3-RC[-9])),"""")"
    'Copy Ref A
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:L10").AutoFill Destination:=.Range("G10:L" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("G10:L10009").Value = Range("G10:L10009").Value
    Range("G10:L10009").Sort Key1:=Range("G10"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    
    'Determine whether 134 or 234 is closer
    Range("D6").Value = "134"
    Range("E6").FormulaR1C1 = "=6371*ACOS(SIN(R[-3]C[-3])*SIN(R[-5]C[-3])+COS(R[-3]C[-3])*COS(R[-5]C[-3])*COS(R[-5]C[-2]-R[-3]C[-2]))"
    Range("E7").FormulaR1C1 = "=6371*ACOS(SIN(R[-3]C[-3])*SIN(R[-4]C[-3])+COS(R[-3]C[-3])*COS(R[-4]C[-3])*COS(R[-4]C[-2]-R[-3]C[-2]))"
    Range("E8").FormulaR1C1 = "=6371*ACOS(SIN(R[-4]C[-3])*SIN(R[-7]C[-3])+COS(R[-4]C[-3])*COS(R[-7]C[-3])*COS(R[-7]C[-2]-R[-4]C[-2]))"
    Range("F7").FormulaR1C1 = "=SUM(R[-1]C[-1]:R[1]C[-1])"
    Range("G7").FormulaR1C1 = "=IF(RC[-1]>=750,0.25*RC[-1],0.28*RC[-1])"
    Range("H7").FormulaR1C1 = "=IF(RC[-2]>=750,0.45*RC[-2],"""")"
    Range("I6").Value = "234"
    Range("J6").FormulaR1C1 = "=6371*ACOS(SIN(R[-3]C[-8])*SIN(R[-4]C[-8])+COS(R[-3]C[-8])*COS(R[-4]C[-8])*COS(R[-4]C[-7]-R[-3]C[-7]))"
    Range("J7").FormulaR1C1 = "=6371*ACOS(SIN(R[-3]C[-8])*SIN(R[-4]C[-8])+COS(R[-3]C[-8])*COS(R[-4]C[-8])*COS(R[-4]C[-7]-R[-3]C[-7]))"
    Range("J8").FormulaR1C1 = "=6371*ACOS(SIN(R[-4]C[-8])*SIN(R[-6]C[-8])+COS(R[-4]C[-8])*COS(R[-6]C[-8])*COS(R[-6]C[-7]-R[-4]C[-7]))"
    Range("K7").FormulaR1C1 = "=SUM(R[-1]C[-1]:R[1]C[-1])"
    Range("L7").FormulaR1C1 = "=IF(RC[-1]>=750,0.25*RC[-1],0.28*RC[-1])"
    Range("M7").FormulaR1C1 = "=IF(RC[-2]>=750,0.45*RC[-2],"""")"
    Range("L8").FormulaR1C1 = "=MIN(R[2]C:R[10001]C)"
    Range("L9").FormulaR1C1 = "=MAX(R[1]C:R[10000]C)"
    Range("N7").FormulaR1C1 = "=RC[-7]-R[1]C[-2]"
    Range("O7").FormulaR1C1 = "=RC[-3]-R[1]C[-3]"
    Range("P7").FormulaR1C1 = "=SUM(RC[-2]:RC[-1])"
    Range("Q7").FormulaR1C1 = "=SUM(R[1]C[-3]:R[1]C[-2])"
    Range("N8").FormulaR1C1 = "=IF(R[-1]C[-6]<>"""",R[-1]C[-6]-R[1]C[-2],"""")"
    Range("O8").FormulaR1C1 = "=IF(R[-1]C[-2]<>"""",R[-1]C[-2]-R[1]C[-3],"""")"
    Range("P7").FormulaR1C1 = "=SUM(RC[-2]:RC[-1])"
    Range("Q7").FormulaR1C1 = "=SUM(R[1]C[-3]:R[1]C[-2])"
   
    If Range("P7") < Range("Q7") Then
        Range("P2:T2").Value = Range("A1:E1").Value
    ElseIf Range("Q7") < Range("P7") Then
        Range("P2:T2").Value = Range("A2:E2").Value
    End If
    
    Range("P3:T4").Value = Range("A3:E4").Value
    Range("U2:U3").FormulaR1C1 = "=6371*ACOS(SIN(RC[-4])*SIN(R[1]C[-4])+COS(RC[-4])*COS(R[1]C[-4])*COS(R[1]C[-3]-RC[-3]))"
    Range("U4").FormulaR1C1 = "=6371*ACOS(SIN(RC[-4])*SIN(R[-2]C[-4])+COS(RC[-4])*COS(R[-2]C[-4])*COS(R[-2]C[-3]-RC[-3]))"
    Range("U5").FormulaR1C1 = "=SUM(R[-3]C:R[-1]C)"
    Range("L9,D6:Q8").Clear
   
    Range("G10").FormulaR1C1 = "=IF(AND(RC[-6]<R4C16,6371*ACOS(SIN(RC[-5])*SIN(R3C2)+COS(RC[-5])*COS(R3C2)*COS(R3C3-RC[-4]))<=10),RC[-6],"""")"
    Range("H10:K10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-6],"""")"
    Range("L10").FormulaR1C1 = "=IF(RC[-1]<>"""",6371*ACOS(SIN(RC[-4])*SIN(R4C17)+COS(RC[-4])*COS(R4C17)*COS(R4C18-RC[-3])),"""")"
    Range("M10").FormulaR1C1 = "=IF(RC[-1]="""","""",IF(OR(AND(R5C21>=750,RC[-1]>0.25*R5C21,RC[-1]<=0.45*R5C21),AND(R5C21<750,RC[-1]>=0.28*R5C21)),1,""""))"
    'Copy Ref A value Sort
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:M10").AutoFill Destination:=.Range("G10:M" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("G10:M10009").Value = Range("G10:M10009").Value
    Range("G10:M10009").Sort Key1:=Range("G10"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    
    Range("M9").FormulaR1C1 = "=SUM(R[1]C:R[10000]C)"
    
    Range("A6").Value = 6.94444444444444E-03
    Range("O10").FormulaR1C1 = "=IF(AND(RC[-14]<R10C7,6371*ACOS(SIN(RC[-13])*SIN(R2C17)+COS(RC[-13])*COS(R2C17)*COS(R2C18-RC[-12]))<=15,RC[-14]>=R2C16-R6C1,RC[-14]<=R2C16+R6C1),RC[-14],"""")"
    Range("P10:S10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-14],"""")"
    Range("T10").FormulaR1C1 = "=IF(RC[-1]<>"""",6371*ACOS(SIN(RC[-18])*SIN(R4C17)+COS(RC[-18])*COS(R4C17)*COS(R4C18-RC[-17])),"""")"
    Range("U10").FormulaR1C1 = "=IF(RC[-1]<>"""",1,"""")"
    'Copy Ref A Value Sort
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("O10:U10").AutoFill Destination:=.Range("O10:U" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("O10:U10009").Value = Range("O10:U10009").Value
    Range("O10:U10009").Sort Key1:=Range("O10"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    
    Range("U9").FormulaR1C1 = "=SUM(R[1]C:R[10000]C)"
    
    If Range("M9") < Range("U9") And Range("M9") <= 500 Then
    
    Range("W10").FormulaR1C1 = "=IF(RC[-10]=1,RC[-16],"""")"
    Range("X10:AA10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-16],"""")"
    Range("AB10").FormulaR1C1 = "=IF(RC[-1]<>"""",6371*ACOS(SIN(RC[-20])*SIN(R4C17)+COS(RC[-20])*COS(R4C17)*COS(R4C18-RC[-19])),"""")"
    'Copy Ref G Value Sort Copy/Transpose to W1"
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
.Range("W10:AB10").AutoFill Destination:=.Range("W10:AB" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("W10:AB509").Value = Range("W10:AB509").Value
    Range("W10:AB509").Sort Key1:=Range("W10"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    
    Range("W10:AB509").Copy
    Range("W1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    Range("W10:AB509").Clear
    
    'MATRIX
    Range("W10:TB10").FormulaR1C1 = _
        "=IF(R1C="""","""",IF(AND(RC20+R6C+6371*ACOS(SIN(RC16)*SIN(R2C)+COS(RC16)*COS(R2C)*COS(R3C-RC17))>=R5C21,MIN(RC20,R6C,6371*ACOS(SIN(RC16)*SIN(R2C)+COS(RC16)*COS(R2C)*COS(R3C-RC17)))>=0.25*(SUM(RC20,R6C,6371*ACOS(SIN(RC16)*SIN(R2C)+COS(RC16)*COS(R2C)*COS(R3C-RC17)))),MAX(RC20,R6C,6371*ACOS(SIN(RC16)*SIN(R2C)+COS(RC16)*COS(R2C)*COS(R3C-RC17)))<=0.45*SUM(RC20,R6C,6371*ACOS(SIN(RC16)*SIN(R2C)+COS(RC16)*COS(R2C)*COS(R3C-RC17)))),SUM(RC20,R6C,6371*ACOS(SIN(RC16)*SIN(R2C)+COS(RC16)*COS(R2C)*COS(R3C-RC17))),""""))"
    'Copy Ref O
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "O").End(xlUp).Row
.Range("W10:TB10").AutoFill Destination:=.Range("W10:TB" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("W10:TB3000").Value = Range("W10:TB3000").Value
   
    Range("R6").FormulaR1C1 = "=MAX(R[4]C[5]:R[2994]C[500])"
    Range("R6").Value = Range("R6").Value
    
    Range("W7:TB7").FormulaR1C1 = "=IF(MAX(R[3]C:R[10002]C)=R6C18,R[-6]C,"""")"
    Range("W7:TB7").Value = Range("W7:TB7").Value
    Range("V10:V3000").FormulaR1C1 = "=IF(MAX(RC[1]:RC[500])=R6C18,RC[-7],"""")"
    ActiveSheet.Calculate
    Range("V10:V3000").Value = Range("V10:V3000").Value
    Range("W10:TB3000").Clear
    
    Range("W8:TB11").FormulaR1C1 = "=IF(R[-1]C<>"""",R[-6]C,"""")"
    Range("P2").FormulaR1C1 = "=MAX(R[5]C[7]:R[5]C[506])"
    Range("Q2").FormulaR1C1 = "=MAX(R[6]C[6]:R[6]C[505])"
    Range("R2").FormulaR1C1 = "=MAX(R[7]C[5]:R[7]C[504])"
    Range("S2").FormulaR1C1 = "=MAX(R[8]C[4]:R[8]C[503])"
    Range("T2").FormulaR1C1 = "=MAX(R[9]C[3]:R[9]C[502])"
    Range("P2:T2").Value = Range("P2:T2").Value
    Range("W8:TB11").Clear
   
    Range("W10:Z3000").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-7],"""")"
    Range("P3:T3").FormulaR1C1 = "=MAX(R[7]C[6]:R[2997]C[6])"
    
    Range("P3:T3").Value = Range("P3:T3").Value
   
    Range("P2:T4").Select
    Range("P2:T4").Sort Key1:=Range("P2"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    
    Range("U2:U3").FormulaR1C1 = "=6371*ACOS(SIN(RC[-4])*SIN(R[1]C[-4])+COS(RC[-4])*COS(R[1]C[-4])*COS(R[1]C[-3]-RC[-3]))"
    Range("U4").FormulaR1C1 = "=6371*ACOS(SIN(RC[-4])*SIN(R[-2]C[-4])+COS(RC[-4])*COS(R[-2]C[-4])*COS(R[-2]C[-3]-RC[-3]))"
    Range("U5").FormulaR1C1 = "=SUM(R[-3]C:R[-1]C)"
    Range("U2:U5").Value = Range("U2:U5").Value
    End If
    
    Range("R6,G9:TB3000,W1:TB7").Clear
  
  ElseIf Range("O1") > Range("O2") And Range("O1") > Range("O2") And Range("O1") > Range("O3") And Range("O1") > Range("O4") Then
    Application.Run "F.xlsm!TriORDS123"
  End If
    Range("V5").FormulaR1C1 = "=SUM(R[-3]C[-1]:R[-1]C[-1])"
    Range("V6").FormulaR1C1 = "=IF(TASKS!R[8]C[-14]-MAX(TASKS!R[10]C[-14],TASKS!R[11]C[-14])>R9C3,R[-1]C-((TASKS!R[8]C[-14]-MAX(TASKS!R[10]C[-14],TASKS!R[11]C[-14])-R9C4)*0.1),R[-1]C)"
    Range("W6").FormulaR1C1 = "=IF(OR(AND(R[-1]C[-1]<750,MIN(R[-4]C[-2]:R[-2]C[-2])>=0.28*R[-1]C[-1]),AND(R[-1]C[-1]>=750,MIN(R[-4]C[-2]:R[-2]C[-2])>=0.25*R[-1]C[-1],MAX(R[-4]C[-2]:R[-2]C[-2])<=0.45*R[-1]C[-1])),""OK"","""")"
    Range("U2:W6").Value = Range("U2:W6").Value
End Sub
Sub TriORDS123()
'
' If Range O1 = MAX O1:O4 And G2 = MIN F1,G2,H3
'
Application.ScreenUpdating = False
    'Range("G9:U10009,P2:U5").Clear
    Range("K1").FormulaR1C1 = "=IF(AND(RC[4]=MAX(RC[4]:R[3]C[4]),R[1]C[-4]=MIN(RC[-5],R[1]C[-4],R[2]C[-3])),123,"""")"
    Range("K1").Value = Range("K1").Value
    
 If Range("K1") = 123 Then
    Range("A7").Value = 3.47222222222222E-03
    Range("U5").FormulaR1C1 = "=IF(R[-4]C[-6]<750,R[-4]C[-6]*0.28,R[-4]C[-6]*0.25)"
    Range("U6").FormulaR1C1 = "=IF(R[-5]C[-6]>=750,0.45*R[-5]C[-6],0.44*R[-5]C[-6])"
    Range("U5:U6").Value = Range("U5:U6").Value
    
    Range("G10").FormulaR1C1 = "=IF(AND(RC[-6]>=R1C1-R7C1,RC[-6]<=R1C1+R7C1),RC[-6],"""")"
    Range("H10:K10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-6],"""")"
    'Copy Ref A value sort
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:K10").AutoFill Destination:=.Range("G10:K" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("G10:K10009").Value = Range("G10:K10009").Value
    Range("G10:K10009").Sort Key1:=Range("G10"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
     
    Range("M10").FormulaR1C1 = "=IF(AND(RC[-12]>=R2C1-R7C1,RC[-12]<=R2C1+R7C1),RC[-12],"""")"
    Range("N10:Q10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-12],"""")"
    'Copy Ref A value sort
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("M10:Q10").AutoFill Destination:=.Range("M10:Q" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("M10:Q10009").Value = Range("M10:Q10009").Value
    Range("M10:Q10009").Sort Key1:=Range("M10"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
     
    Range("S10").FormulaR1C1 = "=IF(AND(RC[-18]>=R3C1-R7C1,RC[-18]<=R3C1+R7C1),RC[-18],"""")"
    Range("T10:W10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-18],"""")"
    'Copy Ref A value sort
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("S10:W10").AutoFill Destination:=.Range("S10:W" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("S10:W10009").Value = Range("S10:W10009").Value
    Range("S10:W10009").Sort Key1:=Range("S10"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
     
    Range("M10:Q500").Copy
    Range("Z1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    
    'MATRIX
    Range("Z10:TE10").FormulaR1C1 = _
        "=IF(OR(R1C="""",RC21=R3C),"""",IF(AND(6371*ACOS(SIN(RC20)*SIN(R2C)+COS(RC20)*COS(R2C)*COS(R3C-RC21))>=R5C21,6371*ACOS(SIN(RC20)*SIN(R2C)+COS(RC20)*COS(R2C)*COS(R3C-RC21))<=R6C21),6371*ACOS(SIN(RC20)*SIN(R2C)+COS(RC20)*COS(R2C)*COS(R3C-RC21)),""""))"
    'Copy Ref S
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "S").End(xlUp).Row
.Range("Z10:TE10").AutoFill Destination:=.Range("Z10:TE" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("Z10:TE3000").Value = Range("Z10:TE3000").Value
    Range("V5").FormulaR1C1 = "=MAX(R[5]C[4]:R[995]C[153])"
    Range("V5").Value = Range("V5").Value

    Range("Z6:TE6").FormulaR1C1 = "=IF(R1C="""","""",IF(MAX(R[4]C:R[2994]C)=R5C22,R[-5]C,""""))"
    Range("Z6:TE6").Value = Range("Z6:TE6").Value
    Range("Y10:Y3000").FormulaR1C1 = "=IF(MAX(RC[1]:RC[500])=R5C22,RC[-6],"""")"
    Range("Y10:Y3000").Value = Range("Y10:Y3000").Value
    
    Range("Z10:TE3000").Clear
    
    Range("Z7:TE10").FormulaR1C1 = "=IF(R[-1]C<>"""",R[-5]C,"""")"
    
    Range("P3").FormulaR1C1 = "=MAX(R[3]C[10]:R[3]C[509])"
    Range("Q3").FormulaR1C1 = "=MAX(R[4]C[9]:R[4]C[508])"
    Range("R3").FormulaR1C1 = "=MAX(R[5]C[8]:R[5]C[507])"
    Range("S3").FormulaR1C1 = "=MAX(R[6]C[7]:R[6]C[506])"
    Range("T3").FormulaR1C1 = "=MAX(R[7]C[6]:R[7]C[505])"
    Range("P3:T3").Value = Range("P3:T3").Value
    
    Range("Z7:TE10").Clear
    
    Range("Z10:AC10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-6],"""")"
    'Copy Ref S
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "S").End(xlUp).Row
.Range("Z10:AC10").AutoFill Destination:=.Range("Z10:AC" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("P4:T4").FormulaR1C1 = "=MAX(R[6]C[9]:R[2996]C[9])"
    Range("P4:T4").Value = Range("P4:T4").Value
    
    Range("M10:W3000").Clear
    Columns("Y:TE").Clear

    Range("P8:U8").FormulaR1C1 = "=MAX(R[2]C:R[2992]C)"
    
    Range("L10").FormulaR1C1 = "=6371*ACOS(SIN(RC[-4])*SIN(R3C17)+COS(RC[-4])*COS(R3C17)*COS(R3C18-RC[-3]))"
    Range("M10").FormulaR1C1 = "=R3C21"
    Range("N10").FormulaR1C1 = "=6371*ACOS(SIN(RC[-6])*SIN(R4C17)+COS(RC[-6])*COS(R4C17)*COS(R4C18-RC[-5]))"
    Range("O10").FormulaR1C1 = "=SUM(RC[-3]:RC[-1])"
    Range("P10").FormulaR1C1 = _
        "=IF(OR(AND(RC[-1]<750,MIN(RC[-4]:RC[-2])>=0.28*RC[-1]),AND(RC[-1]>=750,MIN(RC[-4]:RC[-2])>=0.25*RC[-1],MAX(RC[-4]:RC[-2])<=0.45*RC[-1])),RC[-1],"""")"
   
    Range("Q10").FormulaR1C1 = "=IF(RC[-1]=R8C16,RC[-10],"""")"
    Range("R10:U10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-10],"""")"
    'Copy Ref G
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
.Range("L10:U10").AutoFill Destination:=.Range("L10:U" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("P2:T2").Value = Range("Q8:U8").Value
    
    Range("U2:U3").FormulaR1C1 = "=6371*ACOS(SIN(RC[-4])*SIN(R[1]C[-4])+COS(RC[-4])*COS(R[1]C[-4])*COS(R[1]C[-3]-RC[-3]))"
    Range("U4").FormulaR1C1 = "=6371*ACOS(SIN(RC[-4])*SIN(R[-2]C[-4])+COS(RC[-4])*COS(R[-2]C[-4])*COS(R[-2]C[-3]-RC[-3]))"
    'Range("V5").FormulaR1C1 = "=SUM(R[-3]C[-1]:R[-1]C[-1])"
    'Range("V6").FormulaR1C1 = "=IF(TASKS!R[8]C[-14]-MAX(TASKS!R[10]C[-14],TASKS!R[11]C[-14])>R9C4,R[-1]C-((TASKS!R[8]C[-14]-MAX(TASKS!R[10]C[-14],TASKS!R[11]C[-14])-R9C4)*0.1),R[-1]C)"
    'Range("W6").FormulaR1C1 = "=IF(OR(AND(R[-1]C[-1]<750,MIN(R[-4]C[-2]:R[-2]C[-2])>=0.28*R[-1]C[-1]),AND(R[-1]C[-1]>=750,MIN(R[-4]C[-2]:R[-2]C[-2])>=0.25*R[-1]C[-1],MAX(R[-4]C[-2]:R[-2]C[-2])<=0.45*R[-1]C[-1])),""OK"","""")"
    'Range("U2:W6").Value = Range("U2:W6").Value
    
    Range("G8:X10009").Clear
  End If
End Sub

Sub Ramy06()
'
' Checks for 2-TP Triangle if TP1 earlier than O&R Start
'
 If Range("W6") = "OK" And Range("P2") < Range("A8") Then
    'Save Tri2 results
    Range("G6:K8").Value = Range("P2:T4").Value
    Sheets("Sheet2").Range("C8:D8").Value = Sheets("TASKS").Range("D16:E16").Value
    Range("G10").FormulaR1C1 = "=IF(AND(RC[-6]>R2C16,RC[-6]<=R8C1),RC[-6],"""")"
    Range("H10:K10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-6],"""")"
    Range("L10").FormulaR1C1 = "=IF(RC[-1]<>"""",6371*ACOS(SIN(RC[-10])*SIN(R8C3)+COS(RC[-10])*COS(R8C3)*COS(R8C4-RC[-9])),"""")"
    'Copy Ref A Value Sort
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:L10").AutoFill Destination:=.Range("G10:L" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("G10:L10009").Value = Range("G10:L10009").Value
    Range("G10:L10009").Sort Key1:=Range("G10"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    
    Range("N10").FormulaR1C1 = "=6371*ACOS(SIN(RC[-6])*SIN(R3C17)+COS(RC[-6])*COS(R3C17)*COS(R3C18-RC[-5]))"
    Range("O10").FormulaR1C1 = "=6371*ACOS(SIN(RC[-7])*SIN(R4C17)+COS(RC[-7])*COS(R4C17)*COS(R4C18-RC[-6]))"
    Range("Q10").FormulaR1C1 = "=SUM(RC[-3],RC[-2],R3C21)"
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
.Range("N10:Q10").AutoFill Destination:=.Range("N10:Q" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("N10:Q10009").Value = Range("N10:Q10009").Value
   
    Range("E8").Value = 1.38888888888889E-03
    Range("S10").FormulaR1C1 = "=IF(AND(RC[-18]>=R8C2-R8C5,RC[-18]<=R8C2+R8C5),RC[-18],"""")"
    Range("T10:W10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-18],"""")"
    Range("X10").FormulaR1C1 = "=IF(RC[-1]<>"""",6371*ACOS(SIN(R4C17)*SIN(RC[-22])+COS(R4C17)*COS(RC[-22])*COS(RC[-21]-R4C18)),"""")"
    'Copy Ref A Value Sort
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("S10:X10").AutoFill Destination:=.Range("S10:X" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("S10:X10009").Value = Range("S10:X10009").Value
    Range("S10:X10009").Sort Key1:=Range("S10"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    
    Range("S10:X500").Copy
    Range("Z1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    Range("S10:X500").Clear
    'Matrix
    Range("Z10:TE10").FormulaR1C1 = _
        "=IF(OR(R1C="""",RC9=R3C),"""",IF(OR(R6C<RC15,6371*ACOS(SIN(RC8)*SIN(R2C)+COS(RC8)*COS(R2C)*COS(R3C-RC9))>0.7),"""",IF(RC10-R4C>R9C4,RC17-((RC10-R4C-R9C4)*0.1),RC17)))"
    'Copy Ref G
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
.Range("Z10:TE10").AutoFill Destination:=.Range("Z10:TE" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("X5").FormulaR1C1 = "=MAX(R[5]C[2]:R[3004]C[501])"
    Range("X5").Value = Range("X5").Value
    Range("Z10:TE3000").Value = Range("Z10:TE3000").Value
    
    Range("Z7:TE7").FormulaR1C1 = "=IF(MAX(R[3]C:R[10002]C)=R5C24,R[-6]C,"""")"
    Range("Z7:TE7").Value = Range("Z7:TE7").Value
    
    Range("R10:R3000").FormulaR1C1 = "=IF(MAX(RC[8]:RC[507])=R5C24,RC[-11],"""")"
    ActiveSheet.Calculate
    Range("R10:R3000").Value = Range("R10:R3000").Value
    
    Range("Z10:TE3000").Clear
    
    Range("S10:V10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-11],"""")"
    'Copy Ref G
     With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
.Range("S10:V10").AutoFill Destination:=.Range("S10:V" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("P2:T2").FormulaR1C1 = "=MAX(R[8]C[2]:R[10007]C[2])"
    Range("P2:T2").Value = Range("P2:T2").Value
    
    Range("Z8:TE8").FormulaR1C1 = "=IF(R[-1]C=MIN(R[-1]C26:R[-1]C525),R[-6]C,"""")"
    Range("Z9:TE11").FormulaR1C1 = "=IF(R[-1]C<>"""",R[-6]C,"""")"
    Range("Z8:TE11").Value = Range("Z8:TE11").Value
 
    Range("P5").FormulaR1C1 = "=MIN(R[2]C[10]:R[2]C[509])"
    Range("Q5").FormulaR1C1 = "=MAX(R[3]C[9]:R[3]C[508])"
    Range("R5").FormulaR1C1 = "=MAX(R[4]C[8]:R[4]C[507])"
    Range("S5").FormulaR1C1 = "=MAX(R[5]C[7]:R[5]C[506])"
    Range("T5").FormulaR1C1 = "=MAX(R[6]C[6]:R[6]C[505])"
    Range("P5:T5").Value = Range("P5:T5").Value
    
    Range("P1").Value = "This is a 2-Turn Point Triangle"
    
    Range("U2").FormulaR1C1 = "=6371*ACOS(SIN(RC[-4])*SIN(R[1]C[-4])+COS(RC[-4])*COS(R[1]C[-4])*COS(R[1]C[-3]-RC[-3]))"
    Range("U3").FormulaR1C1 = "=6371*ACOS(SIN(RC[-4])*SIN(R[1]C[-4])+COS(RC[-4])*COS(R[1]C[-4])*COS(R[1]C[-3]-RC[-3]))"
    Range("U4").FormulaR1C1 = "=6371*ACOS(SIN(RC[-4])*SIN(R[1]C[-4])+COS(RC[-4])*COS(R[1]C[-4])*COS(R[1]C[-3]-RC[-3]))"
    Range("V5").FormulaR1C1 = "=SUM(R[-3]C[-1]:R[-1]C[-1])"
    Range("V6").FormulaR1C1 = _
        "=IF(TASKS!R[8]C[-14]-MAX(TASKS!R[10]C[-14],TASKS!R[11]C[-14])>R9C4,R[-1]C-((TASKS!R[8]C[-14]-MAX(TASKS!R[10]C[-14],TASKS!R[11]C[-14])-R9C4)*0.1),R[-1]C)"

    Range("W6").FormulaR1C1 = _
        "=IF(OR(AND(RC[-1]<750,MIN(R[-4]C[-2]:R[-2]C[-2])>=0.28*RC[-1]),AND(RC[-1]>=750,MIN(R[-4]C[-2]:R[-2]C[-2])>=0.25*RC[-1],MAX(R[-4]C[-2]:R[-2]C[-2])<=0.45*RC[-1])),""OK"","""")"
    Range("C8:E8,G10:W10009,X1:TE11").Clear
 End If
    'Ck Start, given FinLine length
    Range("N5").FormulaR1C1 = "=6371*ACOS(SIN(RC[3])*SIN(R[-3]C[3])+COS(RC[3])*COS(R[-3]C[3])*COS(R[-3]C[4]-RC[4]))"
   
    If Range("N5") <= 0.5 Then
        Exit Sub
    End If
    
    'Ck for Finish candidates based on FinLine length
    If Range("N5") > 0.5 Then
        Range("G10").FormulaR1C1 = _
        "=IF(AND(RC[-6]<R3C16,6371*ACOS(SIN(RC[-5])*SIN(R5C17)+COS(RC[-5])*COS(R5C17)*COS(R5C18-RC[-4]))<=0.5),6371*ACOS(SIN(RC[-5])*SIN(R3C17)+COS(RC[-5])*COS(R3C17)*COS(R3C18-RC[-4])),"""")"
        'Copy Ref A
        With Worksheets("Sheet2")
        LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
    .Range("G10:G10").AutoFill Destination:=.Range("G10:G" & LastRow), Type:=xlFillDefault
        End With
     ActiveSheet.Calculate
        Range("L8").FormulaR1C1 = "=MAX(R[2]C[-5]:R[10001]C[-5])"
        Range("L8").Value = Range("L8").Value
        Range("G10:G10009").Clear
        'If G8 =0 ck for 3 TP triangle
        If Range("L8") = 0 Then
            Range("X4:X5").FormulaR1C1 = "=6371*ACOS(SIN(R[-1]C[-7])*SIN(RC[-7])+COS(R[-1]C[-7])*COS(RC[-7])*COS(RC[-6]-R[-1]C[-6]))"
            Range("X6").FormulaR1C1 = "=6371*ACOS(SIN(R[-1]C[-7])*SIN(R[-3]C[-7])+COS(R[-1]C[-7])*COS(R[-3]C[-7])*COS(R[-3]C[-6]-R[-1]C[-6]))"
            Range("Y5").FormulaR1C1 = "=SUM(R[-1]C[-1]:R[1]C[-1])"
            Range("Y6").FormulaR1C1 = _
                "=IF(TASKS!R[8]C[-17]-MAX(TASKS!R[10]C[-17],TASKS!R[11]C[-17])>R9C4,R[-1]C-((TASKS!R[8]C[-17]-MAX(TASKS!R[10]C[-17],TASKS!R[11]C[-17])-R9C4)*0.1),R[-1]C)"
            Range("Z6").FormulaR1C1 = _
                "=IF(OR(AND(RC[-1]<750,MIN(R[-2]C[-2]:RC[-2])>=0.28*RC[-1]),AND(RC[-1]>=750,MIN(R[-2]C[-2]:RC[-2])>=0.25*RC[-1],MAX(R[-2]C[-2]:RC[-2])<=0.45*RC[-1])),""OK"","""")"
            
            If Range("Z6") <> "OK" Then
                Range("G10").FormulaR1C1 = "=IF(RC[-6]>R4C16,RC[-6],"""")"
                Range("H10:K10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-6],"""")"
                'Copy Ref A Value Sort
                With Worksheets("Sheet2")
                LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
            .Range("G10:K10").AutoFill Destination:=.Range("G10:K" & LastRow), Type:=xlFillDefault
                End With
                ActiveSheet.Calculate
                Range("G10:K10009").Value = Range("G10:K10009").Value
                Range("G10:K10009").Sort Key1:=Range("G10"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    
                Range("L10").FormulaR1C1 = "=R4C24"
                Range("M10").FormulaR1C1 = "=6371*ACOS(SIN(RC[-5])*SIN(R4C17)+COS(RC[-5])*COS(R4C17)*COS(R4C18-RC[-4]))"
                Range("N10").FormulaR1C1 = "=6371*ACOS(SIN(RC[-6])*SIN(R3C17)+COS(RC[-6])*COS(R3C17)*COS(R3C18-RC[-5]))"
                Range("O10").FormulaR1C1 = "=SUM(RC[-3]:RC[-1])"
                Range("P8:U8").FormulaR1C1 = "=MAX(R[2]C:R[10001]C)"
                Range("P10").FormulaR1C1 = _
                    "=IF(OR(AND(RC[-1]<750,MIN(RC[-4]:RC[-2])>=0.28*RC[-1]),AND(RC[-1]>=750,MIN(RC[-4]:RC[-2])>=0.25*RC[-1],MAX(RC[-4]:RC[-2])<=0.45*RC[-1])),RC[-1],"""")"
                Range("Q10").FormulaR1C1 = "=IF(RC[-1]=R8C16,RC[-10],"""")"
                Range("R10:U10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-10],"""")"
                'Copy Ref G
                With Worksheets("Sheet2")
                LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
            .Range("L10:U10").AutoFill Destination:=.Range("L10:U" & LastRow), Type:=xlFillDefault
                End With
                ActiveSheet.Calculate
                Range("P8:U8").Value = Range("P8:U8").Value
    
                Range("G10:U10009").Clear
    
                Range("G10").FormulaR1C1 = "=IF(AND(RC[-6]<R3C16,6371*ACOS(SIN(RC[-5])*SIN(R8C18)+COS(RC[-5])*COS(R8C18)*COS(R8C19-RC[-4]))<=0.5),RC[-6],"""")"
                'Copy Ref A
                With Worksheets("Sheet2")
                LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
            .Range("G10:G10").AutoFill Destination:=.Range("G10:G" & LastRow), Type:=xlFillDefault
                End With
                ActiveSheet.Calculate
                    If Range("L8") = 0 Then
                        Range("P8:U8").Clear
                    ElseIf Range("G8") <> 0 Then
                        Range("P5:T5").Value = Range("Q8:U8").Value
                    End If
          End If
        End If
      End If
    Range("X4:Z6").Value = Range("X4:Z6").Value
    Range("P1").Clear
    Range("P2:T2").Value = Range("P3:T3").Value
    Range("P3:T3").Value = Range("P4:T4").Value
    Range("P4:T4").Value = Range("P5:T5").Value
    Range("P5:T5").Clear
    Range("U2:U4").Value = Range("X4:X6").Value
    Range("L8,N5,X4:Z6").Clear
    If Range("P1") = "" And Range("G6") > Range("A8") Then
        Range("P2:T4").Value = Range("G6:K8").Value
        Range("U2").FormulaR1C1 = "=6371*ACOS(SIN(RC[-4])*SIN(R[1]C[-4])+COS(RC[-4])*COS(R[1]C[-4])*COS(R[1]C[-3]-RC[-3]))"
        Range("U3").FormulaR1C1 = "=6371*ACOS(SIN(RC[-4])*SIN(R[1]C[-4])+COS(RC[-4])*COS(R[1]C[-4])*COS(R[1]C[-3]-RC[-3]))"
        Range("U4").FormulaR1C1 = "=6371*ACOS(SIN(RC[-4])*SIN(R[-2]C[-4])+COS(RC[-4])*COS(R[-2]C[-4])*COS(R[-2]C[-3]-RC[-3]))"
        Range("U2:U4").Value = Range("U2:U4").Value
    End If
    Range("G6:K8").Clear
        
End Sub

Sub TriFline1()
'
' Adapted from Azz2
'
'Select Tri FINI + 2 fixes before, 2 fixes after
Application.ScreenUpdating = False
 If Range("A9") <> "REF" Then
    Range("G10").FormulaR1C1 = "=IF(OR(RC[-6]=R5C1,AND(R[-1]C[-6]<R5C1,RC[-6]>R5C1)),1,"""")"
    Range("H10").FormulaR1C1 = "=IF(R[2]C[-1]=1,1,"""")"
    Range("I10").FormulaR1C1 = "=IF(R[1]C[-2]=1,1,"""")"
    Range("J10").FormulaR1C1 = "=IF(R[-1]C[-3]=1,1,"""")"
    Range("K10").FormulaR1C1 = "=IF(R[-2]C[-4]=1,1,"""")"
    Range("L10").FormulaR1C1 = _
        "=IF(OR(RC[-5]=1,RC[-4]=1,RC[-3]=1,RC[-2]=1,RC[-1]=1),RC[-11],"""")"
    Range("M10:P10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-11],"""")"
    
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:P10").AutoFill Destination:=.Range("G10:P" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("L10:P10009").Value = Range("L10:P10009").Value
    Range("G10:K10009").Clear
    
    Range("L10:P10009").Sort Key1:=Range("L10"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal

 ElseIf Range("A9") = "REF" Then
    Sheets("Sheet3").Activate
    Sheets("Sheet3").Range("A5").Value = Sheets("Sheet2").Range("A5").Value
    Range("G10").FormulaR1C1 = "=IF(OR(RC[-6]=R5C1,AND(R[-1]C[-6]<R5C1,RC[-6]>R5C1)),1,"""")"
    Range("H10").FormulaR1C1 = "=IF(R[2]C[-1]=1,1,"""")"
    Range("I10").FormulaR1C1 = "=IF(R[1]C[-2]=1,1,"""")"
    Range("J10").FormulaR1C1 = "=IF(R[-1]C[-3]=1,1,"""")"
    Range("K10").FormulaR1C1 = "=IF(R[-2]C[-4]=1,1,"""")"
    Range("L10").FormulaR1C1 = _
        "=IF(OR(RC[-5]=1,RC[-4]=1,RC[-3]=1,RC[-2]=1,RC[-1]=1),RC[-11],"""")"
    Range("M10:P10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-11],"""")"
    
    With Worksheets("Sheet3")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:P10").AutoFill Destination:=.Range("G10:P" & LastRow), Type:=xlFillDefault
    End With
 ActiveSheet.Calculate
    Range("L10:P60009").Value = Range("L10:P60009").Value
    Range("G10:K60009").Clear
    
    Range("L10:P60009").Sort Key1:=Range("L10"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    Sheets("Sheet2").Range("L10:P14").Value = Sheets("Sheet3").Range("L10:P14").Value
    Sheets("Sheet3").Range("A5,L10:P14").Clear
    Sheets("Sheet2").Activate
 End If
 
    Sheets("YDWK3").Range("M22:Q22").Value = Sheets("Sheet2").Range("A4:E4").Value
    If Range("A1") <> "This is a 2-Turn Point Triangle" Then
        Sheets("YDWK3").Range("M23:Q23").Value = Sheets("Sheet2").Range("A1:E1").Value
    ElseIf Range("A1") = "This is a 2-Turn Point Triangle" Then
        Sheets("YDWK3").Range("M23:Q23").Value = Sheets("Sheet2").Range("A2:E2").Value
    End If
    Sheets("YDWK3").Range("M25:Q29").Value = Sheets("Sheet2").Range("L10:P14").Value
    Sheets("YDWK3").Activate
Application.Run "F.xlsm!FINlines"

    Sheets("Sheet2").Activate
    Sheets("Sheet2").Range("A6").Value = Sheets("YDWK3").Range("M17").Value
    Sheets("Sheet2").Range("B6").FormulaR1C1 = "=IF(RC[-1]<>""NO FIN LINE"",""        FINISH LINE"","""")"
    Range("B6").Value = Range("B6").Value
    Sheets("Sheet2").Range("D6:E6").Value = Sheets("YDWK3").Range("N17:O17").Value
    Range("L10:P14").Clear

Application.Run "F.xlsm!YDWK3clear"

If Range("A6") = "NO FIN LINE" Then
    Application.Run "F.xlsm!TriFline1b"
End If

If Range("A9") = "REF" Then
    Application.Run "F.xlsm!TriFline1a"
End If

Application.Run "F.xlsm!TriOZSector"
End Sub
Sub TriFline1a()
'
' Finds Finish Line candidate at OR Start, within 5 mins of OR Finish
'
Application.ScreenUpdating = False
    Sheets("Sheet3").Activate
    Sheets("Sheet3").Range("A1:E5").Value = Sheets("Sheet2").Range("A1:E5").Value
    Sheets("Sheet3").Range("M5").Value = Sheets("Sheet2").Range("F7").Value
    Sheets("Sheet3").Range("A7").Value = 3.47222222222222E-03
    If Range("A1") <> "This is a 2-Turn Point Triangle" Then
        Range("G10:K10").Value = Range("A1:E1").Value
    Else:
        Range("G10:K10").Value = Range("A2:E2").Value
    End If
    Range("L10").FormulaR1C1 = "=6371*ACOS(SIN(RC[-4])*SIN(R[-6]C[-10])+COS(RC[-4])*COS(R[-6]C[-10])*COS(R[-6]C[-9]-RC[-3]))"
    
    Range("M10").FormulaR1C1 = "=IF(AND(RC[-12]>=R5C1-R7C1,RC[-12]<=R5C1+R7C1),RC[-12],"""")"
    Range("N10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-12],"""")"
    Range("O10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-12],"""")"
    Range("P10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-12],"""")"
    Range("Q10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-12],"""")"
    'Copy Ref A Value Sort to N1
 With Worksheets("Sheet3")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("M10:Q10").AutoFill Destination:=.Range("M10:Q" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("M10:Q60009").Value = Range("M10:Q60009").Value
    Range("M10:Q60009").Sort Key1:=Range("M10"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    Range("M10:Q509").Copy
    Range("N1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    Range("M10:Q509").Clear
    
    Range("N6:SS6").FormulaR1C1 = "=IF(R1C<>"""",6371*ACOS(SIN(R2C)*SIN(R4C2)+COS(R2C)*COS(R4C2)*COS(R4C3-R3C)),"""")"
    
    'MATRIX
    Range("N10:SS10").FormulaR1C1 = _
        "=IF(OR(R1C="""",R3C=RC9,R6C[1]<R6C),"""",IF(OR(R6C<RC12,6371*ACOS(SIN(RC8)*SIN(R2C)+COS(RC8)*COS(R2C)*COS(R3C-RC9))>0.5,SQRT(RC12^2+0.25)<=R2C),"""",IF(RC10-R4C>R9C4,R5C13-((RC10-R4C-R9C4)*0.1),R5C13)))"
    ActiveSheet.Calculate
    Range("M6").FormulaR1C1 = "=MAX(R[4]C[1]:R[4]C[500])"
    Range("M6").Value = Range("M6").Value
    Range("C7").FormulaR1C1 = "=MIN(RC[11]:RC[510])"
    Range("N7:SS7").FormulaR1C1 = "=IF(R[3]C=R6C13,R[-6]C,"""")"
    Range("C7").Value = Range("C7").Value
    
  If Range("C7") = 0 Then
    Range("A6").Value = "NO FIN LINE"
    Range("C7,G1:SS7,G10:L10").Clear
  ElseIf Range("C7") <> 0 Then
    
    Range("G10:L10").Clear
    'Find 5 fixes BASED ON C7
    Range("G10").FormulaR1C1 = "=IF(RC[-6]=R7C3,1,"""")"
    Range("H10").FormulaR1C1 = "=IF(R[2]C[-1]=1,1,"""")"
    Range("I10").FormulaR1C1 = "=IF(R[1]C[-2]=1,1,"""")"
    Range("J10").FormulaR1C1 = "=IF(R[-1]C[-3]=1,1,"""")"
    Range("K10").FormulaR1C1 = "=IF(R[-2]C[-4]=1,1,"""")"
    Range("L10").FormulaR1C1 = _
        "=IF(OR(RC[-5]=1,RC[-4]=1,RC[-3]=1,RC[-2]=1,RC[-1]=1),RC[-11],"""")"
    Range("M10:P10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-11],"""")"
    
    With Worksheets("Sheet3")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:P10").AutoFill Destination:=.Range("G10:P" & LastRow), Type:=xlFillDefault
    End With
ActiveSheet.Calculate
    Range("L10:P60009").Value = Range("L10:P60009").Value
    Range("G10:K60009").Clear
    
    Range("L10:P60009").Sort Key1:=Range("L10"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    Sheets("Sheet2").Range("L10:P14").Value = Sheets("Sheet3").Range("L10:P14").Value
    Sheets("Sheet3").Range("A1:SS8,L10:P14,Q10:SS10").Clear
    Sheets("Sheet2").Activate
 
    Sheets("YDWK3").Range("M22:Q22").Value = Sheets("Sheet2").Range("A4:E4").Value
    If Range("A1") <> "This is a 2-Turn Point Triangle" Then
        Sheets("YDWK3").Range("M23:Q23").Value = Sheets("Sheet2").Range("A1:E1").Value
    ElseIf Range("A1") = "This is a 2-Turn Point Triangle" Then
        Sheets("YDWK3").Range("M23:Q23").Value = Sheets("Sheet2").Range("A2:E2").Value
    End If
    Sheets("YDWK3").Range("M25:Q29").Value = Sheets("Sheet2").Range("L10:P14").Value
    Sheets("YDWK3").Activate
Application.Run "F.xlsm!FINlines"

    Sheets("Sheet2").Activate
    Sheets("Sheet2").Range("A6").Value = Sheets("YDWK3").Range("M17").Value
    Sheets("Sheet2").Range("B6").FormulaR1C1 = "=IF(RC[-1]<>""NO FIN LINE"",""        FINISH LINE"","""")"
    Range("B6").Value = Range("B6").Value
    Sheets("Sheet2").Range("D6:E6").Value = Sheets("YDWK3").Range("N17:O17").Value
    Range("L10:P14").Clear
    Range("F7").FormulaR1C1 = "=IF(R[-6]C[-2]-R[-1]C[-2]>R9C4,R[-1]C-((R[-6]C[-2]-R[-1]C[-2]-R9C4)*0.1),R[-1]C)"
    Range("F8").FormulaR1C1 = _
        "=IF(OR(AND(R[-1]C<750,MIN(R[-6]C:R[-4]C)>=0.28*R[-1]C),AND(R[-1]C>=750,MIN(R[-6]C:R[-4]C)>=0.25*R[-1]C,MAX(R[-6]C:R[-4]C)<=0.45*R[-1]C)),""OK"",""NOPE"")"
    Range("F7:F8").Value = Range("F7:F8").Value
Application.Run "F.xlsm!YDWK3clear"
 End If
End Sub
Sub TriFline1b()
'
' Cks within 5 mins OR Fini, 30 mins OR Start
'
Application.ScreenUpdating = False
    Range("A7").Value = 6.94444444444444E-03
    Range("A8").Value = 2.08333333333333E-02
    'Finis to TOP
    Range("G10").FormulaR1C1 = _
        "=IF(AND(RC[-6]>=R5C1-R7C1,RC[-6]<=R5C1+R7C1),RC[-6],"""")"
    Range("H10:K10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-6],"""")"
    Range("L10").FormulaR1C1 = "=IF(RC[-1]<>"""",6371*ACOS(SIN(RC[-10])*SIN(R4C2)+COS(RC[-10])*COS(R4C2)*COS(R4C3-RC[-9])),"""")"
    'Copy Ref A Value Sort
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:L10").AutoFill Destination:=.Range("G10:L" & LastRow), Type:=xlFillDefault
    End With
ActiveSheet.Calculate
    Range("G10:L10009").Value = Range("G10:L10009").Value
    Range("G10:L10009").Sort Key1:=Range("G10"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal

    Range("G10:L509").Copy
    Range("N1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    Range("L6").FormulaR1C1 = "=MIN(RC[2]:RC[501])"
    Range("M6").FormulaR1C1 = "=MAX(RC[1]:RC[500])"
    Range("L6:M6").Value = Range("L6:M6").Value
    
    Range("G10:L10009").Clear
    'STARTS
    If Range("A1") <> "This is a 2-Turn Point Triangle" Then
        Range("G10").FormulaR1C1 = _
            "=IF(RC[-4]=R4C3,"""",IF(AND(RC[-6]>=R1C1-R8C1,RC[-6]<=R1C1+R8C1,6371*ACOS(SIN(RC[-5])*SIN(R4C2)+COS(RC[-5])*COS(R4C2)*COS(R4C3-RC[-4]))>=R6C12),RC[-6],""""))"
    Else:
        Range("G10").FormulaR1C1 = _
            "=IF(RC[-4]=R4C3,"""",IF(AND(RC[-6]>=R2C1-R8C1,RC[-6]<=R2C1+R8C1,6371*ACOS(SIN(RC[-5])*SIN(R4C2)+COS(RC[-5])*COS(R4C2)*COS(R4C3-RC[-4]))>=R6C12),RC[-6],""""))"
    End If
    
    Range("H10:K10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-6],"""")"
    Range("L10").FormulaR1C1 = _
        "=IF(RC[-1]<>"""",6371*ACOS(SIN(RC[-10])*SIN(R4C2)+COS(RC[-10])*COS(R4C2)*COS(R4C3-RC[-9])),"""")"
    'Copy Ref A Value Sort
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:L10").AutoFill Destination:=.Range("G10:L" & LastRow), Type:=xlFillDefault
    End With
ActiveSheet.Calculate
    Range("G10:L10009").Value = Range("G10:L10009").Value
    Range("G10:L10009").Sort Key1:=Range("G10"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal

If Range("G10") <> "" Then
    'MATRIX Uses MIN!!
    Range("N10:SS10").FormulaR1C1 = _
        "=IF(OR(R1C="""",R3C=RC9),"""",IF(OR(R6C<RC12,6371*ACOS(SIN(RC8)*SIN(R2C)+COS(RC8)*COS(R2C)*COS(R3C-RC9))>0.5,SQRT(RC12^2+0.25)<=R2C),"""",R6C+RC12))"
    
    'Copy Ref G
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
.Range("N10:SS10").AutoFill Destination:=.Range("N10:SS" & LastRow), Type:=xlFillDefault
    End With
ActiveSheet.Calculate
    Range("N10:SS3000").Value = Range("N10:SS3000").Value
    
    Range("J6").FormulaR1C1 = "=MIN(R[4]C[4]:R[3003]C[503])"
    Range("J6").Value = Range("J6").Value
    Range("N7:SS7").FormulaR1C1 = "=IF(MIN(R[3]C:R[10002]C)=R6C10,R[-6]C,"""")"
    Range("N7:SS7").Value = Range("N7:SS7").Value
    
    Range("M10").FormulaR1C1 = "=IF(MIN(RC[1]:RC[500])=R6C10,RC[-6],"""")"
    'Copy Ref G
 With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
.Range("M10:M10").AutoFill Destination:=.Range("M10:M" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("M10:M3009").Value = Range("M10:M3009").Value
    
    Range("N10:SS3009").Clear
    Range("N8:SS11").FormulaR1C1 = "=IF(R[-1]C<>"""",R[-6]C,"""")"
    
    Range("H7").FormulaR1C1 = "=MAX(RC[6]:RC[505])"
    Range("I7").FormulaR1C1 = "=MAX(R[1]C[5]:R[1]C[504])"
    Range("J7").FormulaR1C1 = "=MAX(R[2]C[4]:R[2]C[503])"
    Range("K7").FormulaR1C1 = "=MAX(R[3]C[3]:R[3]C[502])"
    Range("L7").FormulaR1C1 = "=MAX(R[4]C[2]:R[4]C[501])"
    Range("H7:L7").Value = Range("H7:L7").Value
    
    Range("N10:Q10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-6],"""")"
    'Copy Ref G
 With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
.Range("N10:Q10").AutoFill Destination:=.Range("N10:Q" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("H6:L6").FormulaR1C1 = "=MAX(R[4]C[5]:R[3003]C[5])"
    Range("H6:L6").Value = Range("H6:L6").Value
    
    Columns("N:SS").Clear
    Range("M6,G8:M10009").Clear

    'Select Tri FINI + 2 fixes before, 2 fixes after
    Range("G10").FormulaR1C1 = "=IF(RC[-6]=R7C8,1,"""")"
    Range("H10").FormulaR1C1 = "=IF(R[2]C[-1]=1,1,"""")"
    Range("I10").FormulaR1C1 = "=IF(R[1]C[-2]=1,1,"""")"
    Range("J10").FormulaR1C1 = "=IF(R[-1]C[-3]=1,1,"""")"
    Range("K10").FormulaR1C1 = "=IF(R[-2]C[-4]=1,1,"""")"
    Range("L10").FormulaR1C1 = _
        "=IF(OR(RC[-5]=1,RC[-4]=1,RC[-3]=1,RC[-2]=1,RC[-1]=1),RC[-11],"""")"
    Range("M10:P10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-11],"""")"
    
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:P10").AutoFill Destination:=.Range("G10:P" & LastRow), Type:=xlFillDefault
    End With
ActiveSheet.Calculate
    Range("L10:P10009").Value = Range("L10:P10009").Value
    Range("G10:K10009").Clear
    
    Range("L10:P10009").Sort Key1:=Range("L10"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal

    Sheets("YDWK3").Activate
    Sheets("YDWK3").Range("M22:Q22").Value = Sheets("Sheet2").Range("A4:E4").Value
    Sheets("YDWK3").Range("M23:Q23").Value = Sheets("Sheet2").Range("H6:L6").Value
    Sheets("YDWK3").Range("M25:Q29").Value = Sheets("Sheet2").Range("L10:P14").Value

Application.Run "F.xlsm!FINlines"

    Sheets("Sheet2").Activate
    Sheets("Sheet2").Range("H8").Value = Sheets("YDWK3").Range("M17").Value
    Sheets("Sheet2").Range("I8").FormulaR1C1 = "=IF(RC[-1]<>""NO FIN LINE"",""        FINISH LINE"","""")"
    Range("I8").Value = Range("I8").Value
    Sheets("Sheet2").Range("K8:L8").Value = Sheets("YDWK3").Range("N17:O17").Value
    Range("L10:P14").Clear
    
    Range("M8").FormulaR1C1 = _
        "=IF(RC[-5]=""NO FIN LINE"","""",IF(R[-2]C[-2]-RC[-2]>R9C4,R[-2]C[-7]-((R[-2]C[-2]-RC[-2]-R9C4)*0.1),R[-2]C[-7]))"
    Range("M8").Value = Range("M8").Value

    Application.Run "F.xlsm!YDWK3clear"
    
    If Range("H8") <> "NO FIN LINE" And Range("A1") <> "This is a 2-Turn Point Triangle" Then
        Range("A1:E1").Value = Range("H6:L6").Value
        Range("A5:E6").Value = Range("H7:L8").Value
        Range("F7").Value = Range("M8").Value
    ElseIf Range("H8") <> "NO FIN LINE" And Range("A1") = "This is a 2-Turn Point Triangle" Then
        Range("A2:E2").Value = Range("H6:L6").Value
        Range("A5:E6").Value = Range("H7:L8").Value
        Range("F7").Value = Range("M8").Value
    End If
End If
    Range("A7:A8,H6:M8").Clear
End Sub
Sub TriOZSector()
'
' Adapted from Azz2 Tests for OZ after Triangle1; Tests for/finds OZ Sector; P1 modified 4/18/2018 so sector OK irrespective of date
'
Application.ScreenUpdating = False
    
    If Range("A1") <> "This is a 2-Turn Point Triangle" Then
        Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("B1:C1").Value
    ElseIf Range("A1") = "This is a 2-Turn Point Triangle" Then
        Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("B2:C2").Value
    End If
    Sheets("YDWK3").Range("C6:D6").Value = Sheets("Sheet2").Range("B4:C4").Value
    Sheets("YDWK3").Calculate
    Sheets("Sheet2").Range("C7:C8").Value = Sheets("YDWK3").Range("E2:E3").Value
    Sheets("YDWK3").Range("C5:D6").Value = 0
    Sheets("YDWK3").Activate
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    Sheets("YDWK3").Calculate
    Sheets("Sheet2").Activate
    'CK OZ for Triangle1 result
    If Range("A1") <> "This is a 2-Turn Point Triangle" Then
        Range("H1:L1").Value = Range("A1:E1").Value
    ElseIf Range("A1") = "This is a 2-Turn Point Triangle" Then
        Range("H1:L1").Value = Range("A2:E2").Value
    End If
    Range("H2:L3").Value = Range("A5:E6").Value
   'Find 7 Fixes - 3 before & 3 after raw Fin
    
  If Range("A9") <> "REF" Then
    Range("G10").FormulaR1C1 = "=IF(OR(RC[-6]=R5C1,AND(R[-1]C[-6]<R5C1,RC[-6]>R5C1)),1,"""")"
    Range("H10").FormulaR1C1 = "=IF(R[3]C[-1]=1,1,"""")"
    Range("I10").FormulaR1C1 = "=IF(R[2]C[-2]=1,1,"""")"
    Range("J10").FormulaR1C1 = "=IF(R[1]C[-3]=1,1,"""")"
    Range("K10").FormulaR1C1 = "=IF(R[-1]C[-4]=1,1,"""")"
    Range("L10").FormulaR1C1 = "=IF(R[-2]C[-5]=1,1,"""")"
    Range("M10").FormulaR1C1 = "=IF(R[-3]C[-6]=1,1,"""")"
    
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:M10").AutoFill Destination:=.Range("G10:M" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("G10:M10009").Value = Range("G10:M10009").Value
    Range("N10").FormulaR1C1 = _
        "=IF(OR(RC[-7]=1,RC[-6]=1,RC[-5]=1,RC[-4]=1,RC[-3]=1,RC[-2]=1,RC[-1]=1),RC[-13],"""")"
    Range("O10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-13],"""")"
    Range("P10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-13],"""")"
    Range("Q10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-13],"""")"
    Range("R10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-13],"""")"
      
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("N10:R10").AutoFill Destination:=.Range("N10:R" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("N10:R10009").Value = Range("N10:R10009").Value
    Range("G10:M10009").Clear
    Range("N10:R10009").Sort Key1:=Range("N10"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
  
  ElseIf Range("A9") = "REF" Then
    Sheets("Sheet3").Activate
    Sheets("Sheet3").Range("A5").Value = Sheets("Sheet2").Range("A5").Value
    
    Range("G10").FormulaR1C1 = "=IF(OR(RC[-6]=R5C1,AND(R[-1]C[-6]<R5C1,RC[-6]>R5C1)),1,"""")"
    Range("H10").FormulaR1C1 = "=IF(R[3]C[-1]=1,1,"""")"
    Range("I10").FormulaR1C1 = "=IF(R[2]C[-2]=1,1,"""")"
    Range("J10").FormulaR1C1 = "=IF(R[1]C[-3]=1,1,"""")"
    Range("K10").FormulaR1C1 = "=IF(R[-1]C[-4]=1,1,"""")"
    Range("L10").FormulaR1C1 = "=IF(R[-2]C[-5]=1,1,"""")"
    Range("M10").FormulaR1C1 = "=IF(R[-3]C[-6]=1,1,"""")"
    
    Application.Calculation = xlCalculationManual
    With Worksheets("Sheet3")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:M10").AutoFill Destination:=.Range("G10:M" & LastRow), Type:=xlFillDefault
    End With
ActiveSheet.Calculate
    Range("G10:M60009").Value = Range("G10:M60009").Value
    Range("N10").FormulaR1C1 = _
        "=IF(OR(RC[-7]=1,RC[-6]=1,RC[-5]=1,RC[-4]=1,RC[-3]=1,RC[-2]=1,RC[-1]=1),RC[-13],"""")"
    Range("O10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-13],"""")"
    Range("P10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-13],"""")"
    Range("Q10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-13],"""")"
    Range("R10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-13],"""")"
      
    Application.Calculation = xlCalculationManual
    With Worksheets("Sheet3")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("N10:R10").AutoFill Destination:=.Range("N10:R" & LastRow), Type:=xlFillDefault
    End With
ActiveSheet.Calculate
    Range("N10:R60009").Value = Range("N10:R60009").Value
    Range("G10:M60009").Clear
    Range("N10:R60009").Sort Key1:=Range("N10"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
  
    Sheets("Sheet2").Range("N10:R16").Value = Sheets("Sheet3").Range("N10:R16").Value
    Sheets("Sheet3").Range("A5,N10:R16").Clear
    Sheets("Sheet2").Activate
  End If
   
    'Test 3 earlier, 3 later than A5
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("O10:P10").Value
    Sheets("YDWK3").Range("C6:D6").Value = Sheets("Sheet2").Range("I1:J1").Value
    Sheets("YDWK3").Calculate
    Sheets("Sheet2").Range("M10").Value = Sheets("YDWK3").Range("C2").Value
    Sheets("YDWK3").Range("C5:D6").Value = 0
    Sheets("YDWK3").Activate
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    Sheets("YDWK3").Calculate
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("O11:P11").Value
    Sheets("YDWK3").Range("C6:D6").Value = Sheets("Sheet2").Range("I1:J1").Value
    Sheets("YDWK3").Calculate
    Sheets("Sheet2").Range("M11").Value = Sheets("YDWK3").Range("C2").Value
    Sheets("YDWK3").Range("C5:D6").Value = 0
    Sheets("YDWK3").Activate
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    Sheets("YDWK3").Calculate
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("O12:P12").Value
    Sheets("YDWK3").Range("C6:D6").Value = Sheets("Sheet2").Range("I1:J1").Value
    Sheets("YDWK3").Calculate
    Sheets("Sheet2").Range("M12").Value = Sheets("YDWK3").Range("C2").Value
    Sheets("YDWK3").Range("C5:D6").Value = 0
    Sheets("YDWK3").Activate
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    Sheets("YDWK3").Calculate
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("O13:P13").Value
    Sheets("YDWK3").Range("C6:D6").Value = Sheets("Sheet2").Range("I1:J1").Value
    Sheets("YDWK3").Calculate
    Sheets("Sheet2").Range("M13").Value = Sheets("YDWK3").Range("C2").Value
    Sheets("YDWK3").Range("C5:D6").Value = 0
    Sheets("YDWK3").Activate
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    Sheets("YDWK3").Calculate
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("O14:P14").Value
    Sheets("YDWK3").Range("C6:D6").Value = Sheets("Sheet2").Range("I1:J1").Value
    Sheets("YDWK3").Calculate
    Sheets("Sheet2").Range("M14").Value = Sheets("YDWK3").Range("C2").Value
    Sheets("YDWK3").Range("C5:D6").Value = 0
    Sheets("YDWK3").Activate
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    Sheets("YDWK3").Calculate
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("O15:P15").Value
    Sheets("YDWK3").Range("C6:D6").Value = Sheets("Sheet2").Range("I1:J1").Value
    Sheets("YDWK3").Calculate
    Sheets("Sheet2").Range("M15").Value = Sheets("YDWK3").Range("C2").Value
    Sheets("YDWK3").Range("C5:D6").Value = 0
    Sheets("YDWK3").Activate
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    Sheets("YDWK3").Calculate
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("O16:P16").Value
    Sheets("YDWK3").Range("C6:D6").Value = Sheets("Sheet2").Range("I1:J1").Value
    Sheets("YDWK3").Calculate
    Sheets("Sheet2").Range("M16").Value = Sheets("YDWK3").Range("C2").Value
    Sheets("YDWK3").Range("C5:D6").Value = 0
    Sheets("YDWK3").Activate
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    Sheets("YDWK3").Calculate
    Application.Run "F.xlsm!YDWK3clear"
    
    'Added 9/1/14 G10:G15 to find thru sector G16= former G10:G16/deleted 11/3/14 ignores longer 2nd leg, finds ourside of sector!
    'ADDED NEW G10:G15 11/3/14
    'Range("G10:G15").FormulaR1C1 = _
        "=IF(OR(6371*ACOS(SIN(RC[8])*SIN(R1C9)+COS(RC[8])*COS(R1C9)*COS(R1C10-RC[9]))>1,RC[6]=""""),"""",IF(OR(AND(R7C3>R8C3,RC[6]>=R8C3,RC[6]<=R7C3),AND(R7C3<R8C3,OR(RC[6]<=R7C3,RC[6]>=R8C3))),RC[7],IF(ABS(RC[6]-R[1]C[6])>90,RC[7],"""")))"
    Range("G10:G15").FormulaR1C1 = _
        "=IF(OR(6371*ACOS(SIN(RC[8])*SIN(R1C9)+COS(RC[8])*COS(R1C9)*COS(R1C10-RC[9]))>1,RC[6]=""""),"""",IF(OR(AND(R7C3>R8C3,RC[6]>=R8C3,RC[6]<=R7C3),AND(R7C3<R8C3,OR(RC[6]<=R7C3,RC[6]>=R8C3))),RC[7],IF(AND(ABS(RC[6]-R[1]C[6])>90,6371*ACOS(SIN(R4C2)*SIN(R1C9)+COS(R4C2)*COS(R1C9)*COS(R1C10-R4C3))<6371*ACOS(SIN(RC[8])*SIN(R1C9)+COS(RC[8])*COS(R1C9)*COS(R1C10-RC[9]))),RC[7],"""")))"
    Range("G16").FormulaR1C1 = _
        "=IF(OR(6371*ACOS(SIN(RC[8])*SIN(R1C9)+COS(RC[8])*COS(R1C9)*COS(R1C10-RC[9]))>1,RC[6]=""""),"""",IF(OR(AND(R7C3>R8C3,RC[6]>=R8C3,RC[6]<=R7C3),AND(R7C3<R8C3,OR(RC[6]<=R7C3,RC[6]>=R8C3))),RC[7],""""))"
    Range("H10:K16").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[7],"""")"
    ActiveSheet.Calculate
    Range("G8").FormulaR1C1 = "=MAX(R[2]C[3]:R[8]C[3])"
    Range("G8").Value = Range("G8").Value
    Range("H7").FormulaR1C1 = _
        "=IF(R[1]C[-1]=R[3]C[2],R[3]C[-1],IF(R[1]C[-1]=R[4]C[2],R[4]C[-1],IF(R[1]C[-1]=R[5]C[2],R[5]C[-1],IF(R[1]C[-1]=R[6]C[2],R[6]C[-1],""""))))"
    Range("H8").FormulaR1C1 = _
        "=IF(RC[-1]=R[6]C[2],R[6]C[-1],IF(RC[-1]=R[7]C[2],R[7]C[-1],IF(RC[-1]=R[8]C[2],R[8]C[-1],"""")))"
    Range("H7:H8").Value = Range("H7:H8").Value
    Range("G8").FormulaR1C1 = "=MAX(R[-1]C[1]:RC[1])"
    Range("G8").Value = Range("G8").Value
    Range("H10:H16").FormulaR1C1 = "=IF(R8C7=0,"""",IF(RC[-1]=R8C7,RC[7],""""))"
    Range("H8:K8").FormulaR1C1 = "=MAX(R[2]C:R[8]C)"
    Range("G8:K8").Value = Range("G8:K8").Value
    
    If Range("G8") > 0 Then
        Range("D7").Value = "OK"
        Range("H4:L4").Value = Range("G8:K8").Value
        Range("M4").Value = "SECTOR"
        Range("G8:R16").Clear
        Range("M5").FormulaR1C1 = "=IF(R[-4]C[-2]-R[-1]C[-2]>R9C4,R7C6-((R[-4]C[-2]-R[-1]C[-2]-R9C4)*0.1),R7C6)"
        Range("M5").Value = Range("M5").Value
        
        If Range("I3") <> "        FINISH LINE" Then
            Application.Run "F.xlsm!TriFLine2b"
        End If
     End If
    
    If Range("G8") = 0 And Range("I3") = "        FINISH LINE" Then
        Range("G8:R10009").Clear
    ElseIf Range("G8") = 0 And Range("I3") <> "        Finish Line" Then
        Application.Run "F.xlsm!TriFLine2b"
    End If
     
     Range("P1").FormulaR1C1 = _
        "=IF(AND(R[6]C[-10]=MAX(R[6]C[-10],R[4]C[-3],R[7]C[-3]),R[5]C[-15]<>""NO FIN LINE""),""A1"",IF(OR(AND(R[4]C[-3]=MAX(R[6]C[-10],R[4]C[-3],R[7]C[-3]),R[3]C[-3]=""Sector""),AND(RC[-8]<42278,R[5]C[-15]=""NO FIN LINE"",R[4]C[-3]=MAX(R[4]C[-3],R[7]C[-3]))),""H1"",""H6""))"
     Range("P1").Value = Range("P1").Value
     Range("A7:D8,L9:R16").Clear
     Range("A1").Select
    
     Application.Run "F.xlsm!TriH10"
            
    If Range("E8") <> "" Then
        Sheets("TASKS").Range("C32").Value = Sheets("Sheet2").Range("E8").Value
        Sheets("TASKS").Range("A32").Value = "X"
    End If
    
If Range("H16") = "" Then
    If Range("A1") = "This is a 2-Turn Point Triangle" Then
        Sheets("TASKS").Range("G26").Value = "This is a 2-Turn Point Triangle"
        Sheets("TASKS").Range("C27:E29").Value = Sheets("Sheet2").Range("H10:J12").Value
        Sheets("TASKS").Range("C30").Value = Sheets("Sheet2").Range("H13").Value
        Sheets("TASKS").Range("F30").Value = Sheets("Sheet2").Range("I13").Value
        Sheets("TASKS").Range("H27:I30").Value = Sheets("Sheet2").Range("K10:L13").Value
        Sheets("TASKS").Range("L28:L30").Value = Sheets("Sheet2").Range("M10:M12").Value
        Sheets("TASKS").Activate
        Range("F27:G29").FormulaR1C1 = "=DEGREES(RC[-2])"
        Range("F27:G29").Value = Range("F27:G29").Value
        Range("M30").FormulaR1C1 = "=SUM(R[-2]C[-1]:RC[-1])"
        Range("M30").Value = Range("M30").Value
        Range("N30").FormulaR1C1 = _
        "=IF(OR(AND(R30C13>100,ROUND(R27C8-MAX(R30C8:R31C8),0)<=R1C1),AND(R30C13<=100,R2C1=""PR"",ROUND(R27C8-MAX(R30C8:R31C8),0)<=10*R30C13-100),AND(R30C13<=100,R2C1=0,ROUND(R27C8-MAX(R30C8:R31C8),0)<=10*R30C13)),R30C13,IF(AND(R30C13>100,ROUND(R27C8-MAX(R30C8:R31C8),0)>R1C1),R30C13-((ROUND(R27C8-MAX(R30C8:R31C8),0)-R1C1))*0.1,0))"
        Range("M31").FormulaR1C1 = _
        "=IF(OR(AND(R30C13>100,ROUND(R27C8-MAX(R30C8:R31C8),0)>R1C1),AND(R30C13<=100,R2C1=""PR"",ROUND(R27C8-MAX(R30C8:R31C8),0)>10*R30C13-100),AND(R30C13<=100,R2C1=0,ROUND(R27C8-MAX(R30C8:R31C8),0)>10*R30C13)),""LoH"","""")"
        Range("M31").Value = Range("M31").Value
        Range("M30").Value = Range("N30").Value
        Range("N30").Clear
        If Range("A2") = "PR" Then
            Range("H10:H11,H14:H17,H20:H24,H27:H30").Value = "N/A"
        End If
    ElseIf Range("A1") <> "This is a 2-Turn Point Triangle" Then
        Sheets("TASKS").Range("G26").Value = "This is a 3-Turn Point Triangle"
        Sheets("TASKS").Range("H27:I31").Value = Sheets("Sheet2").Range("K10:L14").Value
        Sheets("TASKS").Range("C27:E31").Value = Sheets("Sheet2").Range("H10:J14").Value
        Sheets("TASKS").Range("L28:L30").Value = Sheets("Sheet2").Range("M11:M13").Value
        Sheets("TASKS").Activate
        Range("F27:G30").FormulaR1C1 = "=DEGREES(RC[-2])"
        Range("F31").FormulaR1C1 = "=RC[-2]"
        Range("F27:G31").Value = Range("F27:G31").Value
        Range("M30").FormulaR1C1 = "=SUM(R[-2]C[-1]:RC[-1])"
        Range("M30").Value = Range("M30").Value
        Range("N30").FormulaR1C1 = _
        "=IF(OR(AND(R30C13>100,ROUND(R27C8-R31C8,0)<=R1C1),AND(R30C13<=100,R2C1=""PR"",ROUND(R27C8-R31C8,0)<=10*R30C13-100),AND(R30C13<=100,R2C1=0,ROUND(R27C8-R31C8,0)<=10*R30C13)),R30C13,IF(AND(R30C13>100,ROUND(R27C8-R31C8,0)>R1C1),R30C13-((ROUND(R27C8-R31C8,0)-R1C1))*0.1,0))"
        Range("M31").FormulaR1C1 = _
        "=IF(OR(AND(R30C13>100,ROUND(R27C8-R31C8,0)>R1C1),AND(R30C13<=100,R2C1=""PR"",ROUND(R27C8-R31C8,0)>10*R30C13-100),AND(R30C13<=100,R2C1=0,ROUND(R27C8-R31C8,0)>10*R30C13)),""LoH"","""")"
        Range("M31").Value = Range("M31").Value
        Range("M30").Value = Range("N30").Value
        Range("N30").Clear
        Range("M30:M31").Value = Range("M30:M31").Value
        If Range("A2") = "PR" Then
            Range("H10:H11,H14:H17,H20:H24,H27:H31").Value = "N/A"
        End If
    End If

ElseIf Range("H16") <> "" Then
    If Range("A1") = "This is a 2-Turn Point Triangle" Then
        Sheets("TASKS").Range("G26").Value = "This is a 2-Turn Point Triangle"
        Sheets("TASKS").Range("C27:E29").Value = Sheets("Sheet2").Range("H16:J18").Value
        Sheets("TASKS").Range("C30").Value = Sheets("Sheet2").Range("H19").Value
        Sheets("TASKS").Range("F30").Value = Sheets("Sheet2").Range("I19").Value
        Sheets("TASKS").Range("H27:I30").Value = Sheets("Sheet2").Range("K16:L19").Value
        Sheets("TASKS").Range("L28:L30").Value = Sheets("Sheet2").Range("M16:M18").Value
        Sheets("TASKS").Activate
        Range("F27:G29").FormulaR1C1 = "=DEGREES(RC[-2])"
        Range("F27:G29").Value = Range("F27:G29").Value
        Range("M30").FormulaR1C1 = "=SUM(R[-2]C[-1]:RC[-1])"
        Range("M30").Value = Range("M30").Value
        Range("N30").FormulaR1C1 = _
        "=IF(OR(AND(R30C13>100,ROUND(R27C8-MAX(R30C8:R31C8),0)<=R1C1),AND(R30C13<=100,R2C1=""PR"",ROUND(R27C8-MAX(R30C8:R31C8),0)<=10*R30C13-100),AND(R30C13<=100,R2C1=0,ROUND(R27C8-MAX(R30C8:R31C8),0)<=10*R30C13)),R30C13,IF(AND(R30C13>100,ROUND(R27C8-MAX(R30C8:R31C8),0)>R1C1),R30C13-((ROUND(R27C8-MAX(R30C8:R31C8),0)-R1C1))*0.1,0))"
        Range("M31").FormulaR1C1 = _
        "=IF(OR(AND(R30C13>100,ROUND(R27C8-MAX(R30C8:R31C8),0)>R1C1),AND(R30C13<=100,R2C1=""PR"",ROUND(R27C8-MAX(R30C8:R31C8),0)>10*R30C13-100),AND(R30C13<=100,R2C1=0,ROUND(R27C8-MAX(R30C8:R31C8),0)>10*R30C13)),""LoH"","""")"
        Range("M31").Value = Range("M31").Value
        Range("M30").Value = Range("N30").Value
        Range("N30").Clear
        If Range("A2") = "PR" Then
            Range("H10:H11,H14:H17,H20:H24,H27:H30").Value = "N/A"
        End If
    ElseIf Range("A1") <> "This is a 2-Turn Point Triangle" Then
        Sheets("TASKS").Range("G26").Value = "This is a 3-Turn Point Triangle"
        Sheets("TASKS").Range("H27:I31").Value = Sheets("Sheet2").Range("K16:L20").Value
        Sheets("TASKS").Range("C27:E31").Value = Sheets("Sheet2").Range("H16:J20").Value
        Sheets("TASKS").Range("L28:L30").Value = Sheets("Sheet2").Range("M17:M19").Value
        Sheets("TASKS").Activate
        Range("F27:G30").FormulaR1C1 = "=DEGREES(RC[-2])"
        Range("F31").FormulaR1C1 = "=IF(RC[-3]=""NO FIN LINE"",""    NO FINISH LINE"",RC[-2])"
        Range("F27:G31").Value = Range("F27:G31").Value
        Range("M30").FormulaR1C1 = "=SUM(R[-2]C[-1]:RC[-1])"
        Range("M30").Value = Range("M30").Value
        Range("N30").FormulaR1C1 = _
        "=IF(OR(AND(R30C13>100,ROUND(R27C8-R31C8,0)<=R1C1),AND(R30C13<=100,R2C1=""PR"",ROUND(R27C8-R31C8,0)<=10*R30C13-100),AND(R30C13<=100,R2C1=0,ROUND(R27C8-R31C8,0)<=10*R30C13)),R30C13,IF(AND(R30C13>100,ROUND(R27C8-R31C8,0)>R1C1),R30C13-((ROUND(R27C8-R31C8,0)-R1C1))*0.1,0))"
        Range("M31").FormulaR1C1 = _
        "=IF(OR(AND(R30C13>100,ROUND(R27C8-R31C8,0)>R1C1),AND(R30C13<=100,R2C1=""PR"",ROUND(R27C8-R31C8,0)>10*R30C13-100),AND(R30C13<=100,R2C1=0,ROUND(R27C8-R31C8,0)>10*R30C13)),""LoH"","""")"
        Range("M31").Value = Range("M31").Value
        Range("M30").Value = Range("N30").Value
        Range("N30").Clear
        If Range("A2") = "PR" Then
            Range("H10:H11,H14:H17,H20:H24,H27:H31").Value = "N/A"
        End If
    End If
 End If
 
 'Date-sensitive rules & Restore raw pressure altitudes to K
    Sheets("TASKS").Activate
    If Range("C10") < 41913 And Range("G26") = "This is a 3-Turn Point Triangle" And Range("M30") < 300 Then
        Range("C33:M33").Value = "Note: Per SC3 rules on this flight date, a 3-Turn Point Triangle required an Official Distance of at least 300km"
    ElseIf Range("C10") > 42278 Then
        If Range("F30") = "    FINISH SECTOR" Then
            Range("C33:M33").Value = "Note: Per SC3 rules on this flight date, Triangle Courses require Finish Line crossing"
        End If
    End If
    
    If Range("A2") <> "PR" Then
        Range("K10,K11,K14:K17,K20:K24").FormulaR1C1 = "=IF(RC[-8]<B!R5C13,RC[-3]-B!R2C13,RC[-3]-B!R8C13)"
        Range("K10:K24").Value = Range("K10:K24").Value
        Range("K27:K29").FormulaR1C1 = "=IF(R[-1]C[-4]=""NONE IDENTIFIED"","""",IF(RC[-8]<B!R5C13,RC[-3]-B!R2C13,RC[-3]-B!R8C13))"
        Range("K30").FormulaR1C1 = "=IF(OR(R26C7=""NONE IDENTIFIED"",RC[-8]=""""),"""",IF(RC[-8]<B!R5C13,RC[-3]-B!R2C13,RC[-3]-B!R8C13))"
        Range("K31").FormulaR1C1 = "=IF(OR(R26C7=""NONE IDENTIFIED"",RC[-8]=""""),"""",IF(RC[-8]<B!R5C13,RC[-3]-B!R2C13,RC[-3]-B!R8C13))"
        Range("H10:H31").Value = Range("K10:K31").Value
    End If
End Sub
Sub TriFLine2b()
'
' Finds Finish Line after Sector - cks w/in 30 mins after OR Start; after last TP & among last 500 fixes recorded
'
Application.ScreenUpdating = False
    Range("A7").FormulaR1C1 = "=LARGE(R[3]C:R[10002]C,500)"
    Range("A7").Value = Range("A7").Value
    Range("A8").Value = 2.08333333333333E-02
    
    'Finishes to Top
    Range("M10").FormulaR1C1 = "=IF(AND(RC[-12]>R4C1,RC[-12]>=R7C1),RC[-12],"""")"
    Range("N10:Q10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-12],"""")"
    'Copy Ref A Value Sort
    Application.Calculation = xlCalculationManual
 With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("M10:Q10").AutoFill Destination:=.Range("M10:Q" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("M10:Q10009").Value = Range("M10:Q10009").Value
    Range("M10:Q10009").Sort Key1:=Range("M10"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    
    Range("M10:Q509").Copy
    Range("N1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    Range("M10:Q3009").Clear
    
    Range("N6:SS6").FormulaR1C1 = "=IF(R[-5]C<>"""",6371*ACOS(SIN(R4C2)*SIN(R2C)+COS(R4C2)*COS(R2C)*COS(R3C-R4C3)),"""")"
    Range("N6:SS6").Value = Range("N6:SS6").Value
    Range("M6").FormulaR1C1 = "=MAX(RC[1]:RC[500])"
    Range("M6").Value = Range("M6").Value
    
    'Starts
    If Range("A1") <> "This is a 2-Turn Point Triangle" Then
        Range("G10").FormulaR1C1 = _
            "=IF(AND(RC[-6]<R2C1,RC[-6]>=R1C1-R8C1,6371*ACOS(SIN(RC[-5])*SIN(R4C2)+COS(RC[-5])*COS(R4C2)*COS(R4C3-RC[-4]))<=R6C13),RC[-6],"""")"
    Else:
        Range("G10").FormulaR1C1 = _
            "=IF(AND(RC[-6]<R3C1,RC[-6]>=R2C1-R8C1,6371*ACOS(SIN(RC[-5])*SIN(R4C2)+COS(RC[-5])*COS(R4C2)*COS(R4C3-RC[-4]))<=R6C13),RC[-6],"""")"
    End If

    Range("H10:K10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-6],"""")"
    Range("L10").FormulaR1C1 = "=IF(RC[-1]<>"""",6371*ACOS(SIN(RC[-4])*SIN(R4C2)+COS(RC[-4])*COS(R4C2)*COS(R4C3-RC[-3])),"""")"
    'CopyRef A Value Sort
 With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:L10").AutoFill Destination:=.Range("G10:L" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("G10:L10009").Value = Range("G10:L10009").Value
    Range("G10:L10009").Sort Key1:=Range("G10"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    
    'Matrix
    Range("N10:SS10").FormulaR1C1 = _
        "=IF(OR(R1C="""",R3C=RC9,R6C[1]<R6C),"""",IF(OR(R6C<RC12,6371*ACOS(SIN(RC8)*SIN(R2C)+COS(RC8)*COS(R2C)*COS(R3C-RC9))>0.5,SQRT(RC12^2+0.25)<=R2C),"""",R6C+RC12))"
    
    'Copy Ref G Value
 With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
.Range("N10:SS10").AutoFill Destination:=.Range("N10:SS" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("N10:SS3009").Value = Range("N10:SS3009").Value
    
    Range("J6").FormulaR1C1 = "=MAX(R[4]C[4]:R[3003]C[503])"
    Range("J6").Value = Range("J6").Value
    Range("N7:SS7").FormulaR1C1 = "=IF(MAX(R[3]C:R[10002]C)=R6C10,R[-6]C,"""")"
    Range("N7:SS7").Value = Range("N7:SS7").Value
    
    Range("M10").FormulaR1C1 = "=IF(MAX(RC[1]:RC[500])=R6C10,RC[-6],"""")"
    'Copy Ref G
 With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
.Range("M10:M10").AutoFill Destination:=.Range("M10:M" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("M10:M3009").Value = Range("M10:M3009").Value
    
    Range("N8:SS11").FormulaR1C1 = "=IF(R[-1]C<>"""",R[-6]C,"""")"
    'Range("N10:SS12").FormulaR1C1 = "=IF(R[-1]C<>"""",R[-7]C,"""")"
    
    Range("H7").FormulaR1C1 = "=MAX(RC[6]:RC[505])"
    Range("I7").FormulaR1C1 = "=MAX(R[1]C[5]:R[1]C[504])"
    Range("J7").FormulaR1C1 = "=MAX(R[2]C[4]:R[2]C[503])"
    Range("K7").FormulaR1C1 = "=MAX(R[3]C[3]:R[3]C[502])"
    Range("L7").FormulaR1C1 = "=MAX(R[4]C[2]:R[4]C[501])"
    Range("H7:L7").Value = Range("H7:L7").Value
    
    Range("N10:Q10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-6],"""")"
    'Copy Ref G
 With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
.Range("N10:Q10").AutoFill Destination:=.Range("N10:Q" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("H6:L6").FormulaR1C1 = "=MAX(R[4]C[5]:R[3003]C[5])"
    Range("H6:L6").Value = Range("H6:L6").Value
    
    Columns("N:SS").Clear
    Range("M6,G8:M10009").Clear

    'Select Tri FINI + 2 fixes before, 2 fixes after
    Range("G10").FormulaR1C1 = "=IF(RC[-6]=R7C8,1,"""")"
    Range("H10").FormulaR1C1 = "=IF(R[2]C[-1]=1,1,"""")"
    Range("I10").FormulaR1C1 = "=IF(R[1]C[-2]=1,1,"""")"
    Range("J10").FormulaR1C1 = "=IF(R[-1]C[-3]=1,1,"""")"
    Range("K10").FormulaR1C1 = "=IF(R[-2]C[-4]=1,1,"""")"
    Range("L10").FormulaR1C1 = _
        "=IF(OR(RC[-5]=1,RC[-4]=1,RC[-3]=1,RC[-2]=1,RC[-1]=1),RC[-11],"""")"
    Range("M10:P10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-11],"""")"
    
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:P10").AutoFill Destination:=.Range("G10:P" & LastRow), Type:=xlFillDefault
    End With
ActiveSheet.Calculate
    Range("L10:P10009").Value = Range("L10:P10009").Value
    Range("G10:K10009").Clear
    
    Range("L10:P10009").Sort Key1:=Range("L10"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal

    Sheets("YDWK3").Activate
    Sheets("YDWK3").Range("M22:Q22").Value = Sheets("Sheet2").Range("A4:E4").Value
    Sheets("YDWK3").Range("M23:Q23").Value = Sheets("Sheet2").Range("H6:L6").Value
    Sheets("YDWK3").Range("M25:Q29").Value = Sheets("Sheet2").Range("L10:P14").Value

Application.Run "F.xlsm!FINlines"

    Sheets("Sheet2").Activate
    Sheets("Sheet2").Range("H8").Value = Sheets("YDWK3").Range("M17").Value
    Sheets("Sheet2").Range("I8").FormulaR1C1 = "=IF(RC[-1]<>""NO FIN LINE"",""        FINISH LINE"","""")"
    Range("I8").Value = Range("I8").Value
    Sheets("Sheet2").Range("K8:L8").Value = Sheets("YDWK3").Range("N17:O17").Value
    Range("L10:P14").Clear
    
    Range("M8").FormulaR1C1 = _
        "=IF(RC[-5]=""NO FIN LINE"","""",IF(R[-2]C[-2]-RC[-2]>R9C4,R[-2]C[-7]-((R[-2]C[-2]-RC[-2]-R9C4)*0.1),R[-2]C[-7]))"
    Range("M8").Value = Range("M8").Value

    Application.Run "F.xlsm!YDWK3clear"

End Sub
Sub TriFLine2a()
'
' amended 3/21/14; amended 11/3/14 for 2-TP Triangles
'
Application.ScreenUpdating = False
    Range("G10:R10009").Clear
    Range("A8").FormulaR1C1 = "=MAX(R[2]C:R[10001]C)"
    Range("B8").FormulaR1C1 = "=IF(AND(R8C1-R5C1<0.0208333333333333,SECOND(R[3]C[-1]-R[2]C[-1])>=4),0.0208333333333333,0.010416666666666)"
    Range("A8:B8").Value = Range("A8:B8").Value
    
    Range("G8").FormulaR1C1 = "=SMALL(R[2]C[-1]:R[10001]C[-1],500)"
    
    Range("F10").FormulaR1C1 = "=IF(AND(RC[-5]>=R5C1-R8C2,RC[-5]<=R5C1+R8C2),ABS(R5C1-RC[-5]),"""")"
    Range("G10").FormulaR1C1 = "=IF(RC[-1]<R8C7,RC[-6],"""")"
    Range("H10:K10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-6],"""")"
    'Copy Ref A Value CLEAR F10:F10009 Sort
    
    Application.Calculation = xlCalculationManual
 With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("F10:K10").AutoFill Destination:=.Range("F10:K" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    On Error Resume Next
        Range("G8").FormulaR1C1 = "=MAX(R[2]C[-1]:R[10001]C[-1])"
    
    Range("F10:K10009").Value = Range("F10:K10009").Value
    Range("G8,F10:F10009").Clear
    Range("G10:K10009").Sort Key1:=Range("G10"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    Range("G10:K509").Copy
    Range("N1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    Range("G10:K10009").Clear
    
    Range("N6:SS6").FormulaR1C1 = "=IF(R[-5]C<>"""",6371*ACOS(SIN(R[-4]C)*SIN(R4C2)+COS(R[-4]C)*COS(R4C2)*COS(R4C3-R[-3]C)),"""")"

    Range("B8").Value = 2.08333333333333E-02
   
    If Range("A1") <> "This is a 2-Turn Point Triangle" Then
        Range("G10").FormulaR1C1 = "=IF(AND(RC[-6]>=R1C1-R8C2,RC[-6]<=R1C1+R8C2),RC[-6],"""")"
    ElseIf Range("A1") = "This is a 2-Turn Point Triangle" Then
        Range("G10").FormulaR1C1 = "=IF(AND(RC[-6]>=R2C1-R8C2,RC[-6]<=R2C1+R8C2),RC[-6],"""")"
    End If
    Range("H10:K10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-6],"""")"
    Range("L10").FormulaR1C1 = _
        "=IF(RC[-1]<>"""",6371*ACOS(SIN(RC[-4])*SIN(R4C2)+COS(RC[-4])*COS(R4C2)*COS(R4C3-RC[-3])),"""")"
    'Copy Ref A Value Sort
 With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:L10").AutoFill Destination:=.Range("G10:L" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("G10:L10009").Value = Range("G10:L10009").Value
    Range("G10:L10009").Sort Key1:=Range("G10"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    
    'MATRIX
    Range("N10:SS10").FormulaR1C1 = _
        "=IF(OR(R1C="""",R3C=RC9,R6C[1]<R6C),"""",IF(OR(R6C<RC12,6371*ACOS(SIN(RC8)*SIN(R2C)+COS(RC8)*COS(R2C)*COS(R3C-RC9))>0.5,SQRT(RC12^2+0.25)<=R2C),"""",IF(RC10-R4C>R9C4,(R6C+RC12)-((RC10-R4C-R9C4)*0.1),R6C+RC12)))"
    'Copy Ref G
 With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
.Range("N10:SS10").AutoFill Destination:=.Range("N10:SS" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("J6").FormulaR1C1 = "=MAX(R[4]C[4]:R[10003]C[503])"
    Range("J6").Value = Range("J6").Value
    
    Range("N7:SS7").FormulaR1C1 = "=IF(MAX(R[3]C:R[10002]C)=R6C10,R[-6]C,"""")"
    Range("N7:SS7").Value = Range("N7:SS7").Value
    
    Range("M10").FormulaR1C1 = "=IF(MAX(RC[1]:RC[500])=R6C10,RC[-6],"""")"
    'Copy Ref G
 With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
.Range("M10:M10").AutoFill Destination:=.Range("M10:M" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("M10:M3009").Value = Range("M10:M3009").Value
    Range("N10:SS10009").Clear
    Range("N8:SS11").FormulaR1C1 = "=IF(R[-1]C<>"""",R[-6]C,"""")"
    
    Range("H7").FormulaR1C1 = "=MAX(R7C[6]:R7C[505])"
    Range("I7").FormulaR1C1 = "=MAX(R8C[5]:R8C[504])"
    Range("J7").FormulaR1C1 = "=MAX(R9C[4]:R9C[503])"
    Range("K7").FormulaR1C1 = "=MAX(R10C[3]:R10C[502])"
    Range("L7").FormulaR1C1 = "=MAX(R11C[2]:R11C[501])"
    Range("H7:L7").Value = Range("H7:L7").Value
    
    Range("N10:Q10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-6],"""")"
    'Copy Ref G
 With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
.Range("N10:Q10").AutoFill Destination:=.Range("N10:Q" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("H6:L6").FormulaR1C1 = "=MAX(R10C[5]:R3009C[5])"
    Range("H6:L6").Value = Range("H6:L6").Value
    
    Columns("N:SS").Clear
    Range("A8,B8,G8:M10009").Clear
  
    'Select Tri FINI + 2 fixes before, 2 fixes after
    Range("G10").FormulaR1C1 = "=IF(RC[-6]=R7C8,1,"""")"
    Range("H10").FormulaR1C1 = "=IF(R[2]C[-1]=1,1,"""")"
    Range("I10").FormulaR1C1 = "=IF(R[1]C[-2]=1,1,"""")"
    Range("J10").FormulaR1C1 = "=IF(R[-1]C[-3]=1,1,"""")"
    Range("K10").FormulaR1C1 = "=IF(R[-2]C[-4]=1,1,"""")"
    Range("L10").FormulaR1C1 = _
        "=IF(OR(RC[-5]=1,RC[-4]=1,RC[-3]=1,RC[-2]=1,RC[-1]=1),RC[-11],"""")"
    Range("M10:P10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-11],"""")"
    
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:P10").AutoFill Destination:=.Range("G10:P" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("L10:P10009").Value = Range("L10:P10009").Value
    Range("G10:K10009").Clear
    
    Range("L10:P10009").Sort Key1:=Range("L10"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal

    Sheets("YDWK3").Activate
    Sheets("YDWK3").Range("M22:Q22").Value = Sheets("Sheet2").Range("A4:E4").Value
    Sheets("YDWK3").Range("M23:Q23").Value = Sheets("Sheet2").Range("H6:L6").Value
    Sheets("YDWK3").Range("M25:Q29").Value = Sheets("Sheet2").Range("L10:P14").Value

Application.Run "F.xlsm!FINlines"

    Sheets("Sheet2").Activate
    Sheets("Sheet2").Range("H8").Value = Sheets("YDWK3").Range("M17").Value
    Sheets("Sheet2").Range("I8").FormulaR1C1 = "=IF(RC[-1]<>""NO FIN LINE"",""        FINISH LINE"","""")"
    Range("I8").Value = Range("I8").Value
    Sheets("Sheet2").Range("K8:L8").Value = Sheets("YDWK3").Range("N17:O17").Value
    Range("L10:P14").Clear
    
    Application.Run "F.xlsm!YDWK3clear"
    'Amended 11/3:
    If Range("A1") <> "This is a 2-Turn Point Triangle" Then
        Range("M8").FormulaR1C1 = _
            "=IF(RC[-5]=""NO FIN LINE"","""",IF(R[-2]C[-2]-RC[-2]>R9C4,R[-2]C[-7]-((R[-2]C[-2]-RC[-2]-R9C4)*0.1),R[-2]C[-7]))"
        Range("M8").Value = Range("M8").Value
    
    ElseIf Range("A1") = "This is a 2-Turn Point Triangle" Then
        Range("N6").FormulaR1C1 = "=6371*ACOS(SIN(RC[-5])*SIN(R[-3]C[-12])+COS(RC[-5])*COS(R[-3]C[-12])*COS(R[-3]C[-11]-RC[-4]))"
        Range("N7").FormulaR1C1 = "=R[-4]C[-8]"
        Range("N8").FormulaR1C1 = "=6371*ACOS(SIN(R[-4]C[-12])*SIN(R[-1]C[-5])+COS(R[-4]C[-12])*COS(R[-1]C[-5])*COS(R[-1]C[-4]-R[-4]C[-11]))"
        Range("O8").FormulaR1C1 = "=SUM(R[-2]C[-1]:RC[-1])"
        Range("O9").FormulaR1C1 = "=IF(R[-3]C[-4]-R[-1]C[-4]<=RC[-11],R[-1]C,R[-1]C-((R[-3]C[-4]-R[-1]C[-4]-RC[-11])*0.1))"
        Range("P8").FormulaR1C1 = "=IF(MIN(R[-2]C[-2]:RC[-2])>=0.28*R[1]C[-1],""OK"",""NOPE"")"
        
       If Range("P8") = "OK" Then
            Range("M8").Value = Range("O9").Value
            Range("N6:P9").Clear
        ElseIf Range("P8") = "NOPE" Then
            'Range("E7").Value = "NONE IDENTIFIED"  or M8=""??
            Range("N6:P9").Clear
        End If
    End If
End Sub

Sub TriH10()
'
' Moves data to H10
'
Application.ScreenUpdating = False
If Range("P1") = "A1" And Range("E8") = "" Then
   If Range("A1") <> "This is a 2-Turn Point Triangle" Then
       Range("H10:L13").Value = Range("A1:E4").Value
       Range("M11:M13").Value = Range("F2:F4").Value
       Range("H14:L14").Value = Range("A6:E6").Value
       Range("M14").Value = Range("F7").Value
   ElseIf Range("A1") = "This is a 2-Turn Point Triangle" Then
       Range("H10:L12").Value = Range("A2:E4").Value
       Range("H13:L13").Value = Range("A6:E6").Value
       Range("M10:M12").Value = Range("F2:F4").Value
       Range("M13").Value = Range("F7").Value
    End If

ElseIf Range("P1") = "A1" And Range("E8") <> "" And Range("A1") <> "This is a 2-Turn Point Triangle" Then
       Range("H10:L13").Value = Range("A1:E4").Value
       Range("H14:L14").Value = Range("A6:E6").Value
       Range("M11:M13").Value = Range("F2:F4").Value
       Range("M14").Value = Range("F7").Value
       Range("H15").Value = "No record-eligible Triangle found. Great Circle distances are shown in red above"
        
ElseIf Range("P1") = "H1" And Range("E8") = "" Then
   If Range("A1") <> "This is a 2-Turn Point Triangle" Then
       Range("H10:L10").Value = Range("H1:L1").Value
       Range("H11:L13").Value = Range("A2:E4").Value
       Range("M11:M13").Value = Range("F2:F4").Value
       Range("M14").Value = Range("M5").Value

       If Range("M4") = "SECTOR" And Range("H3") = "NO FIN LINE" Then
            Range("H14").Value = Range("H4").Value
            Range("I14").Value = "    FINISH SECTOR"
            Range("K14:L14").Value = Range("K4:L4").Value
       ElseIf Range("I4") = "        FINISH LINE" Then
            Range("H14").Value = Range("H3").Value
            Range("I14").Value = Range("I4").Value
            Range("K14:L14").Value = Range("K3:L3").Value
       End If
    
    ElseIf Range("A1") = "This is a 2-Turn Point Triangle" Then
        Range("H10:L10").Value = Range("H1:L1").Value
        Range("H11:L12").Value = Range("A3:E4").Value
        Range("M10:M12").Value = Range("F2:F4").Value
        Range("M13").Value = Range("F7").Value
        
        If Range("M4") = "SECTOR" And Range("H3") = "NO FIN LINE" Then
            Range("H13").Value = Range("H4").Value
            Range("I13").Value = "    FINISH SECTOR"
            Range("K13:L13").Value = Range("K4:L4").Value
        ElseIf Range("I3") = "        FINISH LINE" Then
            Range("H13").Value = Range("H3").Value
            Range("I13").Value = "        FINISH LINE"
            Range("K13:L13").Value = Range("K3:L3").Value
        End If
    End If

ElseIf Range("P1") = "H1" And Range("M4") = "SECTOR" And Range("E8") <> "" And Range("A1") <> "This is a 2-Turn Point Triangle" Then
    Range("H10:L10").Value = Range("H1:L1").Value
    Range("H11:L13").Value = Range("A2:E4").Value
    Range("M11:M13").Value = Range("F2:F4").Value
    Range("M14").Value = Range("F7").Value
    Range("H14").Value = Range("H4").Value
    Range("I14").Value = "    FINISH SECTOR"
    Range("K14:L14").Value = Range("K4:L4").Value
    Range("H15").Value = "No record-eligible Triangle found. Great Circle distances are shown in red above"
    
ElseIf Range("P1") = "H6" And Range("E8") = "" Then
  If Range("A1") <> "This is a 2-Turn Point Triangle" Then
    Range("H10:L10").Value = Range("H6:L6").Value
    Range("H11:L13").Value = Range("A2:E4").Value
    Range("M11:M13").Value = Range("F2:F4").Value
    Range("H14:L14").Value = Range("H8:L8").Value
    Range("M14").Value = Range("M8").Value
  ElseIf Range("A1") = "This is a 2-Turn Point Triangle" Then
    Range("H10:L10").Value = Range("H6:L6").Value
    Range("H11:L12").Value = Range("A3:E4").Value
    Range("H13:L14").Value = Range("H7:L8").Value
    Range("M13").Value = Range("M8").Value
  End If
End If

If Range("A9") <> "REF" Then
    Sheets("Sheet3").Range("A9:E10009").Value = Sheets("Sheet2").Range("A9:E10009").Value
    Application.Run "F.xlsm!TriREF2"
    Sheets("Sheet3").Columns("A:E").Clear
ElseIf Range("A9") = "REF" Then
    Application.Run "F.xlsm!TriREF2"
End If

End Sub
Sub TriREF2()
'
' Cks triangle @ sheet 3 +/- 3 minutes Amended 9/3/14 for 2-TP triangles
'
Application.ScreenUpdating = False
    Sheets("Sheet3").Range("A1:G5").Value = Sheets("Sheet2").Range("H10:N14").Value
    Sheets("Sheet3").Activate
    Range("A8").Value = 2.08333333333333E-03
    
  If Range("A5") <> "" Then
    Range("G2:G3").FormulaR1C1 = "=6371*ACOS(SIN(RC[-5])*SIN(R[1]C[-5])+COS(RC[-5])*COS(R[1]C[-5])*COS(R[1]C[-4]-RC[-4]))"
    Range("G4").FormulaR1C1 = "=6371*ACOS(SIN(RC[-5])*SIN(R[-2]C[-5])+COS(RC[-5])*COS(R[-2]C[-5])*COS(R[-2]C[-4]-RC[-4]))"
    Range("H4").FormulaR1C1 = "=SUM(R[-2]C[-1]:RC[-1])"
    Range("G2:H4").Value = Range("G2:H4").Value
    Range("G10").FormulaR1C1 = _
        "=IF(OR(AND(RC[-6]>=R2C1-R8C1,RC[-6]<=R2C1+R8C1),AND(RC[-6]>=R3C1-R8C1,RC[-6]<=R3C1+R8C1),AND(RC[-6]>=R4C1-R8C1,RC[-6]<=R4C1+R8C1)),RC[-6],"""")"
  ElseIf Range("A5") = "" Then
    Range("G2").FormulaR1C1 = "=6371*ACOS(SIN(R[-1]C[-5])*SIN(RC[-5])+COS(R[-1]C[-5])*COS(RC[-5])*COS(RC[-4]-R[-1]C[-4]))"
    Range("G3").FormulaR1C1 = "=6371*ACOS(SIN(R[-1]C[-5])*SIN(RC[-5])+COS(R[-1]C[-5])*COS(RC[-5])*COS(RC[-4]-R[-1]C[-4]))"
    Range("G4").FormulaR1C1 = "=6371*ACOS(SIN(R[-1]C[-5])*SIN(R[-3]C[-5])+COS(R[-1]C[-5])*COS(R[-3]C[-5])*COS(R[-3]C[-4]-R[-1]C[-4]))"
    Range("G2:H4").Value = Range("G2:H4").Value
    Range("G10").FormulaR1C1 = "=IF(OR(AND(RC[-6]>=R2C1-R8C1,RC[-6]<=R2C1+R8C1),AND(RC[-6]>=R3C1-R8C1,RC[-6]<=R3C1+R8C1)),RC[-6],"""")"
  End If
    Range("H10:K10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-6],"""")"
    'Copy Ref A Value Sort
    Application.Calculation = xlCalculationManual
    With Worksheets("Sheet3")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:K10").AutoFill Destination:=.Range("G10:K" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("G10:K60009").Value = Range("G10:K60009").Value
    Range("G10:K60009").Sort Key1:=Range("G10"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    
    Range("L7").FormulaR1C1 = "=IF(R5C1<>"""",R4C1-R8C1,R3C1-R8C1)"
    Range("L8").FormulaR1C1 = "=IF(R5C1<>"""",R3C1+R8C1,R2C1+R8C1)"
    Range("L9").FormulaR1C1 = "=R2C1+R8C1"
    
    'Range("L10").FormulaR1C1 = _
        "=IF(AND(R5C1<>"""",RC[-5]<R9C12),6371*ACOS(SIN(RC[-4])*SIN(R3C2)+COS(RC[-4])*COS(R3C2)*COS(R3C3-RC[-3])),IF(AND(R5C1="""",RC[-5]<R9C12),6371*ACOS(SIN(RC[-4])*SIN(R1C2)+COS(RC[-4])*COS(R1C2)*COS(R1C3-RC[-3])),""""))"
    'Range("M10").FormulaR1C1 = _
        "=IF(AND(RC[-1]<>"""",R5C1<>""""),6371*ACOS(SIN(RC[-5])*SIN(R4C2)+COS(RC[-5])*COS(R4C2)*COS(R4C3-RC[-4])),IF(AND(RC[-1]<>"""",R5C1=""""),6371*ACOS(SIN(RC[-5])*SIN(R3C2)+COS(RC[-5])*COS(R3C2)*COS(R3C3-RC[-4])),""""))"
    'Range("N10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-2]+RC[-1]+R3C7,"""")"
    'Copy Ref G
    'Application.Calculation = xlCalculationManual
    'With Worksheets("Sheet3")
    'LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
'.Range("L10:N10").AutoFill Destination:=.Range("L10:N" & LastRow), Type:=xlFillDefault
    'End With
'Application.Calculation = xlCalculationAutomatic
    
    Range("P9:U9").FormulaR1C1 = "=MAX(R[1]C:R[10000]C)"
    
    'TP1
    Range("L10").FormulaR1C1 = _
        "=IF(AND(R5C1<>"""",RC[-5]<R9C12),6371*ACOS(SIN(RC[-4])*SIN(R3C2)+COS(RC[-4])*COS(R3C2)*COS(R3C3-RC[-3])),IF(AND(R5C1="""",RC[-5]<R9C12),6371*ACOS(SIN(RC[-4])*SIN(R1C2)+COS(RC[-4])*COS(R1C2)*COS(R1C3-RC[-3])),""""))"
    Range("M10").FormulaR1C1 = _
        "=IF(AND(RC[-1]<>"""",R5C1<>""""),6371*ACOS(SIN(RC[-5])*SIN(R4C2)+COS(RC[-5])*COS(R4C2)*COS(R4C3-RC[-4])),IF(AND(RC[-1]<>"""",R5C1=""""),6371*ACOS(SIN(RC[-5])*SIN(R3C2)+COS(RC[-5])*COS(R3C2)*COS(R3C3-RC[-4])),""""))"
    Range("N10").FormulaR1C1 = "=IF(AND(R5C1<>"""",RC[-1]<>""""),R3C7,IF(AND(R5C1="""",RC[-1]<>""""),R4C7,""""))"
    Range("O10").FormulaR1C1 = "=IF(RC[-1]<>"""",SUM(RC[-3]:RC[-1]),"""")"
    Range("P10").FormulaR1C1 = _
        "=IF(RC[-1]="""","""",IF(OR(AND(RC[-1]<750,MIN(RC[-4]:RC[-2])>=0.28*RC[-1]),AND(RC[-1]>=750,MIN(RC[-4]:RC[-2])>=0.25*RC[-1],MAX(RC[-4]:RC[-2])<=0.45*RC[-1])),RC[-1],""""))"
    ''Range("L10").FormulaR1C1 = _
        "=IF(AND(R5C1<>"""",RC[-5]<R9C12),6371*ACOS(SIN(RC[-4])*SIN(R3C2)+COS(RC[-4])*COS(R3C2)*COS(R3C3-RC[-3])),IF(AND(R5C1="""",RC[-5]<R9C12),6371*ACOS(SIN(RC[-4])*SIN(R1C2)+COS(RC[-4])*COS(R1C2)*COS(R1C3-RC[-3])),""""))"
    ''Range("M10").FormulaR1C1 = _
        "=IF(AND(RC[-1]<>"""",R5C1<>""""),6371*ACOS(SIN(RC[-5])*SIN(R4C2)+COS(RC[-5])*COS(R4C2)*COS(R4C3-RC[-4])),IF(AND(RC[-1]<>"""",R5C1=""""),6371*ACOS(SIN(RC[-5])*SIN(R3C2)+COS(RC[-5])*COS(R3C2)*COS(R3C3-RC[-4])),""""))"
    ''Range("N10").FormulaR1C1 = "=IF(AND(R5C1<>"""",RC[-1]<>""""),RC[-2]+RC[-1]+R3C7,IF(AND(R5C1="""",RC[-1]<>""""),RC[-2]+RC[-1]+R4C7,""""))"
    ''Range("O10").FormulaR1C1 = "=IF(AND(R5C1<>"""",RC[-1]<>""""),SUM(RC[-3]:RC[-1]),IF(R5C1="""",RC[-1],""""))"
    'Range("L10").FormulaR1C1 = "=IF(RC[-5]<=R9C12,6371*ACOS(SIN(RC[-4])*SIN(R3C2)+COS(RC[-4])*COS(R3C2)*COS(R3C3-RC[-3])),"""")"
    'Range("M10").FormulaR1C1 = "=IF(RC[-1]<>"""",6371*ACOS(SIN(RC[-5])*SIN(R4C2)+COS(RC[-5])*COS(R4C2)*COS(R4C3-RC[-4])),"""")"
    'Range("N10").FormulaR1C1 = "=IF(RC[-1]<>"""",R3C7,"""")"
    'Range("O10").FormulaR1C1 = "=IF(RC[-1]<>"""",SUM(RC[-3]:RC[-1]),"""")"
    'Range("P10").FormulaR1C1 = _
        "=IF(RC[-1]="""","""",IF(OR(AND(RC[-1]<750,MIN(RC[-4]:RC[-2])>=0.28*RC[-1]),AND(RC[-1]>=750,MIN(RC[-4]:RC[-2])>=0.25*RC[-1],MAX(RC[-4]:RC[-2])<=0.45*RC[-1])),RC[-1],""""))"
    Range("Q10").FormulaR1C1 = "=IF(RC[-1]=R9C16,RC[-10],"""")"
    Range("R10:U10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-10],"""")"
    'Copy Ref G
    With Worksheets("Sheet3")
    LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
.Range("L10:U10").AutoFill Destination:=.Range("L10:U" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("P2:T2").Value = Range("Q9:U9").Value
  
  If Range("A5") = "" Then
         Range("U2").FormulaR1C1 = "=6371*ACOS(SIN(RC[-4])*SIN(R1C2)+COS(RC[-4])*COS(R1C2)*COS(R1C3-RC[-3]))"
  ElseIf Range("A5") <> "" Then
    'TP2 of 3
    Range("L10").FormulaR1C1 = "=IF(AND(RC[-5]>R2C16,RC[-5]<=R8C12),6371*ACOS(SIN(RC[-4])*SIN(R2C17)+COS(RC[-4])*COS(R2C17)*COS(R2C18-RC[-3])),"""")"
    Range("M10").FormulaR1C1 = "=IF(RC[-1]<>"""",6371*ACOS(SIN(RC[-5])*SIN(R4C2)+COS(RC[-5])*COS(R4C2)*COS(R4C3-RC[-4])),"""")"
    Range("N10").FormulaR1C1 = "=IF(RC[-1]<>"""",6371*ACOS(SIN(R4C2)*SIN(R2C17)+COS(R4C2)*COS(R2C17)*COS(R2C18-R4C3)),"""")"
    'Copy Ref G
    With Worksheets("Sheet3")
    LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
.Range("L10:N10").AutoFill Destination:=.Range("L10:N" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("P3:T3").Value = Range("Q9:U9").Value
    Range("U2").FormulaR1C1 = _
        "=6371*ACOS(SIN(RC[-4])*SIN(R[1]C[-4])+COS(RC[-4])*COS(R[1]C[-4])*COS(R[1]C[-3]-RC[-3]))"
    Range("U2").Value = Range("U2").Value
  End If
  
    'Last TP
    Range("L10").FormulaR1C1 = _
        "=IF(RC[-5]>=R7C12,6371*ACOS(SIN(RC[-4])*SIN(R2C17)+COS(RC[-4])*COS(R2C17)*COS(R2C18-RC[-3])),"""")"
    Range("M10").FormulaR1C1 = _
        "=IF(AND(R5C1<>"""",RC[-1]<>""""),6371*ACOS(SIN(RC[-5])*SIN(R3C17)+COS(RC[-5])*COS(R3C17)*COS(R3C18-RC[-4])),IF(AND(R5C1="""",RC[-1]<>""""),6371*ACOS(SIN(RC[-5])*SIN(R1C2)+COS(RC[-5])*COS(R1C2)*COS(R1C3-RC[-4])),""""))"
    Range("N10").FormulaR1C1 = "=IF(RC[-1]<>"""",R2C21,"""")"
    Range("O10").FormulaR1C1 = "=IF(RC[-1]<>"""",SUM(RC[-3]:RC[-1]),"""")"
    'Range("L10").FormulaR1C1 = "=IF(RC[-5]>=R7C12,6371*ACOS(SIN(RC[-4])*SIN(R2C17)+COS(RC[-4])*COS(R2C17)*COS(R2C18-RC[-3])),"""")"
    'Range("M10").FormulaR1C1 = "=IF(RC[-1]<>"""",6371*ACOS(SIN(RC[-5])*SIN(R3C17)+COS(RC[-5])*COS(R3C17)*COS(R3C18-RC[-4])),"""")"
    'Range("N10").FormulaR1C1 = "=IF(RC[-1]<>"""",R2C21,"""")"
    'Copy Ref G
    With Worksheets("Sheet3")
    LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
.Range("L10:O10").AutoFill Destination:=.Range("L10:O" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("P4:T4").Value = Range("Q9:U9").Value
    
  If Range("A5") <> "" Then
    Range("U3").FormulaR1C1 = "=6371*ACOS(SIN(RC[-4])*SIN(R[1]C[-4])+COS(RC[-4])*COS(R[1]C[-4])*COS(R[1]C[-3]-RC[-3]))"
    
  ElseIf Range("A5") = "" Then
    Range("P1").Value = "This is a 2-Turn Point Triangle"
    Range("P3:T3").Value = Range("P2:T2").Value
    Range("P2:T2").Value = Range("A1:E1").Value
    Range("P5:T5").Value = Range("A4:E4").Value
    Range("U2:U3").FormulaR1C1 = "=6371*ACOS(SIN(RC[-4])*SIN(R[1]C[-4])+COS(RC[-4])*COS(R[1]C[-4])*COS(R[1]C[-3]-RC[-3]))"
  End If
    
    Range("U4").FormulaR1C1 = "=6371*ACOS(SIN(RC[-4])*SIN(R[-2]C[-4])+COS(RC[-4])*COS(R[-2]C[-4])*COS(R[-2]C[-3]-RC[-3]))"
    Range("V4").FormulaR1C1 = "=SUM(R[-2]C[-1]:RC[-1])"
    Range("W5").FormulaR1C1 = _
        "=IF(OR(AND(R[-1]C[-1]<750,MIN(R[-3]C[-2]:R[-1]C[-2])>=0.28*R[-1]C[-1]),AND(R[-1]C[-1]>=750,MIN(R[-3]C[-2]:R[-1]C[-2])>=0.25*R[-1]C[-1],MAX(R[-3]C[-2]:R[-1]C[-2])<=0.45*R[-1]C[-1])),""OK"",""NOPE"")"
    'Range("W5").Value = Range("W5").Value
    Range("U2:W5").Value = Range("U2:W5").Value
    
  If Range("W5") = "NOPE" Then
    Range("G5").Value = Range("F5").Value
    Range("P2:W5,G7:U10009").Clear
  
  ElseIf Range("W5") = "OK" Then
    Range("H2:L4").Value = Range("P2:T4").Value

    Range("G7:U10009").Clear
    
    Sheets("YDWK3").Range("E39").Value = Sheets("Sheet3").Range("I2").Value
    Sheets("YDWK3").Range("E40").Value = Sheets("Sheet3").Range("J2").Value
    Sheets("YDWK3").Range("E41").Value = Sheets("Sheet3").Range("I3").Value
    Sheets("YDWK3").Range("E42").Value = Sheets("Sheet3").Range("J3").Value
    Sheets("YDWK3").Calculate
    Sheets("Sheet3").Range("M2").Value = Sheets("YDWK3").Range("F37").Value

    Sheets("YDWK3").Range("E39").Value = Sheets("Sheet3").Range("I4").Value
    Sheets("YDWK3").Range("E40").Value = Sheets("Sheet3").Range("J4").Value
    Sheets("YDWK3").Calculate
    Sheets("Sheet3").Range("M3").Value = Sheets("YDWK3").Range("F37").Value
    
    Sheets("YDWK3").Range("E41").Value = Sheets("Sheet3").Range("I2").Value
    Sheets("YDWK3").Range("E42").Value = Sheets("Sheet3").Range("J2").Value
    Sheets("YDWK3").Calculate
    Sheets("Sheet3").Range("M4").Value = Sheets("YDWK3").Range("F37").Value
    
    Range("N5").FormulaR1C1 = _
        "=IF(OR(AND(RC[-10]<>"""",R[-4]C[-10]-RC[-10]<=R9C4),AND(RC[-10]="""",R[-4]C[-10]-R[-1]C[-10]<=R9C4)),SUM(R[-3]C[-1]:R[-1]C[-1]),IF(RC[-10]<>"""",SUM(R[-3]C[-1]:R[-1]C[-1])-(R[-4]C[-10]-RC[-10]-R9C4)*0.1,SUM(R[-3]C[-1]:R[-1]C[-1])-(R[-4]C[-10]-R[-1]C[-10]-R[4]C[-10])*0.1))"
    'Range("N5").FormulaR1C1 = _
        "=IF(R[-4]C[-10]-RC[-10]<=R9C4,SUM(R[-3]C[-1]:R[-1]C[-1]),SUM(R[-3]C[-1]:R[-1]C[-1])-((R[-4]C[-10]-RC[-10]-R9C4)*0.1))"
    Range("N5").Value = Range("N5").Value

    Range("G2:G4").Clear
    
    If Range("N5") > Range("F5") Then
      If Range("A5") <> "" Then
        Range("A2:F4").Value = Range("H2:M4").Value
      Else:
        Range("A5:E5").Value = Range("A4:E4").Value
        Range("A2:F4").Value = Range("H2:M4").Value
        Range("A1:F1").Clear
        Range("A1").Value = "This is a 2-Turn Point Triangle"
      End If
      Range("F5").Value = Range("N5").Value
    End If
 End If
    Range("G6").FormulaR1C1 = _
        "=IF(R[-1]C[-1]<750,MIN(R[-4]C[-1]:R[-2]C[-1])-(0.28*R[-1]C[-1]),MIN(R[-4]C[-1]:R[-2]C[-1])-(0.25*R[-1]C[-1]))"
    Range("G6").Value = Range("G6").Value
        
If Range("G6") < 0.6 And Range("A1") <> "This is a 2-Turn Point Triangle" Then
    'Expamd to 2 up, 2 dn
    Range("G10").FormulaR1C1 = "=IF(OR(R[-2]C[-6]=R2C1,R[-1]C[-6]=R2C1,R[1]C[-6]=R2C1,R[2]C[-6]=R2C1),RC[-6],"""")"
    Range("H10:K10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-6],"""")"
    'Expand to two up, two down
    Range("M10").FormulaR1C1 = "=IF(OR(R[-2]C[-12]=R3C1,R[-1]C[-12]=R3C1,R[1]C[-12]=R3C1,R[2]C[-12]=R3C1),RC[-12],"""")"
    Range("N10:Q10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-12],"""")"
    'Expand to 2 up, 2 dn
    Range("S10").FormulaR1C1 = "=IF(OR(R[-2]C[-18]=R4C1,R[-1]C[-18]=R4C1,R[1]C[-18]=R4C1,R[2]C[-18]=R4C1),RC[-18],"""")"
    Range("T10:W10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-18],"""")"
    'Copy Ref A
    With Worksheets("Sheet3")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:W10").AutoFill Destination:=.Range("G10:W" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    
    Range("G10:W60009").Value = Range("G10:W60009").Value
    Range("G10:K60009").Sort Key1:=Range("G10"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    Range("M10:Q60009").Sort Key1:=Range("M10"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    Range("S10:W60009").Sort Key1:=Range("S10"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    
    Range("G15:G16").Value = Range("G10:G11").Value
    Range("G18:G19").Value = Range("G12:G13").Value
    Range("M15:M16").Value = Range("M10:M11").Value
    Range("M18:M19").Value = Range("M12:M13").Value
    Range("S15:S16").Value = Range("S10:S11").Value
    Range("S18:S19").Value = Range("S12:S13").Value
    
    Sheets("YDWK3").Range("E39").Value = Sheets("Sheet3").Range("I3").Value
    Sheets("YDWK3").Range("E40").Value = Sheets("Sheet3").Range("J3").Value
    Sheets("YDWK3").Range("E41").Value = Sheets("Sheet3").Range("H10").Value
    Sheets("YDWK3").Range("E42").Value = Sheets("Sheet3").Range("I10").Value
    Sheets("YDWK3").Calculate
    Sheets("Sheet3").Range("H15").Value = Sheets("YDWK3").Range("F37").Value
    
    Sheets("YDWK3").Range("E41").Value = Sheets("Sheet3").Range("H11").Value
    Sheets("YDWK3").Range("E42").Value = Sheets("Sheet3").Range("I11").Value
    Sheets("YDWK3").Calculate
    Sheets("Sheet3").Range("H16").Value = Sheets("YDWK3").Range("F37").Value
    
    Sheets("YDWK3").Range("E41").Value = Sheets("Sheet3").Range("H12").Value
    Sheets("YDWK3").Range("E42").Value = Sheets("Sheet3").Range("I12").Value
    Sheets("YDWK3").Calculate
    Sheets("Sheet3").Range("H18").Value = Sheets("YDWK3").Range("F37").Value
    
    Sheets("YDWK3").Range("E41").Value = Sheets("Sheet3").Range("H13").Value
    Sheets("YDWK3").Range("E42").Value = Sheets("Sheet3").Range("I13").Value
    Sheets("YDWK3").Calculate
    Sheets("Sheet3").Range("H19").Value = Sheets("YDWK3").Range("F37").Value
    
    Sheets("YDWK3").Activate
    Range("E39:E42").Value = 0
    Range("E81").FormulaR1C1 = "=R47C[-2]+(1-R79C)*R46C[-2]*R67C*(ACOS(R62C)+R79C*SIN(ACOS(R62C))*(R69C+R79C*R62C*(-1+2*R69C*R69C)))"
    Sheets("YDWK3").Calculate
    Sheets("YDWK3").Range("E39").Value = Sheets("Sheet3").Range("I3").Value
    Sheets("YDWK3").Range("E40").Value = Sheets("Sheet3").Range("J3").Value
    Sheets("YDWK3").Range("E41").Value = Sheets("Sheet3").Range("T10").Value
    Sheets("YDWK3").Range("E42").Value = Sheets("Sheet3").Range("U10").Value
    Sheets("YDWK3").Calculate
    Sheets("Sheet3").Range("T15").Value = Sheets("YDWK3").Range("F37").Value
    
    Sheets("YDWK3").Range("E41").Value = Sheets("Sheet3").Range("T11").Value
    Sheets("YDWK3").Range("E42").Value = Sheets("Sheet3").Range("U11").Value
    Sheets("YDWK3").Calculate
    Sheets("Sheet3").Range("T16").Value = Sheets("YDWK3").Range("F37").Value
    
    Sheets("YDWK3").Range("E41").Value = Sheets("Sheet3").Range("T12").Value
    Sheets("YDWK3").Range("E42").Value = Sheets("Sheet3").Range("U12").Value
    Sheets("YDWK3").Calculate
    Sheets("Sheet3").Range("T18").Value = Sheets("YDWK3").Range("F37").Value
    
    Sheets("YDWK3").Range("E41").Value = Sheets("Sheet3").Range("T13").Value
    Sheets("YDWK3").Range("E42").Value = Sheets("Sheet3").Range("U13").Value
    Sheets("YDWK3").Calculate
    Sheets("Sheet3").Range("T19").Value = Sheets("YDWK3").Range("F37").Value
    
    Sheets("YDWK3").Activate
    Range("E39:E42").Value = 0
    Range("E81").FormulaR1C1 = "=R47C[-2]+(1-R79C)*R46C[-2]*R67C*(ACOS(R62C)+R79C*SIN(ACOS(R62C))*(R69C+R79C*R62C*(-1+2*R69C*R69C)))"
    Sheets("YDWK3").Calculate
    Sheets("YDWK3").Range("E39").Value = Sheets("Sheet3").Range("I2").Value
    Sheets("YDWK3").Range("E40").Value = Sheets("Sheet3").Range("J2").Value
    Sheets("YDWK3").Range("E41").Value = Sheets("Sheet3").Range("N10").Value
    Sheets("YDWK3").Range("E42").Value = Sheets("Sheet3").Range("O10").Value
    Sheets("YDWK3").Calculate
    Sheets("Sheet3").Range("N15").Value = Sheets("YDWK3").Range("F37").Value
    
    Sheets("YDWK3").Range("E41").Value = Sheets("Sheet3").Range("N11").Value
    Sheets("YDWK3").Range("E42").Value = Sheets("Sheet3").Range("O11").Value
    Sheets("YDWK3").Calculate
    Sheets("Sheet3").Range("N16").Value = Sheets("YDWK3").Range("F37").Value
    
    Sheets("YDWK3").Range("E41").Value = Sheets("Sheet3").Range("N12").Value
    Sheets("YDWK3").Range("E42").Value = Sheets("Sheet3").Range("O12").Value
    Sheets("YDWK3").Calculate
    Sheets("Sheet3").Range("N18").Value = Sheets("YDWK3").Range("F37").Value
    
    Sheets("YDWK3").Range("E41").Value = Sheets("Sheet3").Range("N13").Value
    Sheets("YDWK3").Range("E42").Value = Sheets("Sheet3").Range("O13").Value
    Sheets("YDWK3").Calculate
    Sheets("Sheet3").Range("N19").Value = Sheets("YDWK3").Range("F37").Value
    
    Sheets("YDWK3").Activate
    Range("E39:E42").Value = 0
    Range("E81").FormulaR1C1 = "=R47C[-2]+(1-R79C)*R46C[-2]*R67C*(ACOS(R62C)+R79C*SIN(ACOS(R62C))*(R69C+R79C*R62C*(-1+2*R69C*R69C)))"
    Sheets("YDWK3").Calculate
    Sheets("YDWK3").Range("E39").Value = Sheets("Sheet3").Range("I4").Value
    Sheets("YDWK3").Range("E40").Value = Sheets("Sheet3").Range("J4").Value
    Sheets("YDWK3").Range("E41").Value = Sheets("Sheet3").Range("H10").Value
    Sheets("YDWK3").Range("E42").Value = Sheets("Sheet3").Range("I10").Value
    Sheets("YDWK3").Calculate
    Sheets("Sheet3").Range("I15").Value = Sheets("YDWK3").Range("F37").Value
    
    Sheets("YDWK3").Range("E41").Value = Sheets("Sheet3").Range("H11").Value
    Sheets("YDWK3").Range("E42").Value = Sheets("Sheet3").Range("I11").Value
    Sheets("YDWK3").Calculate
    Sheets("Sheet3").Range("I16").Value = Sheets("YDWK3").Range("F37").Value
    
    Sheets("YDWK3").Range("E41").Value = Sheets("Sheet3").Range("H12").Value
    Sheets("YDWK3").Range("E42").Value = Sheets("Sheet3").Range("I12").Value
    Sheets("YDWK3").Calculate
    Sheets("Sheet3").Range("I18").Value = Sheets("YDWK3").Range("F37").Value
    
    Sheets("YDWK3").Range("E41").Value = Sheets("Sheet3").Range("H13").Value
    Sheets("YDWK3").Range("E42").Value = Sheets("Sheet3").Range("I13").Value
    Sheets("YDWK3").Calculate
    Sheets("Sheet3").Range("I19").Value = Sheets("YDWK3").Range("F37").Value
    
    Sheets("YDWK3").Activate
    Range("E39:E42").Value = 0
    Range("E81").FormulaR1C1 = "=R47C[-2]+(1-R79C)*R46C[-2]*R67C*(ACOS(R62C)+R79C*SIN(ACOS(R62C))*(R69C+R79C*R62C*(-1+2*R69C*R69C)))"
    Sheets("YDWK3").Calculate
    Sheets("YDWK3").Range("E39").Value = Sheets("Sheet3").Range("I4").Value
    Sheets("YDWK3").Range("E40").Value = Sheets("Sheet3").Range("J4").Value
    Sheets("YDWK3").Range("E41").Value = Sheets("Sheet3").Range("N10").Value
    Sheets("YDWK3").Range("E42").Value = Sheets("Sheet3").Range("O10").Value
    Sheets("YDWK3").Calculate
    Sheets("Sheet3").Range("O15").Value = Sheets("YDWK3").Range("F37").Value
    
    Sheets("YDWK3").Range("E41").Value = Sheets("Sheet3").Range("N11").Value
    Sheets("YDWK3").Range("E42").Value = Sheets("Sheet3").Range("O11").Value
    Sheets("YDWK3").Calculate
    Sheets("Sheet3").Range("O16").Value = Sheets("YDWK3").Range("F37").Value
    
    Sheets("YDWK3").Range("E41").Value = Sheets("Sheet3").Range("N12").Value
    Sheets("YDWK3").Range("E42").Value = Sheets("Sheet3").Range("O12").Value
    Sheets("YDWK3").Calculate
    Sheets("Sheet3").Range("O18").Value = Sheets("YDWK3").Range("F37").Value
    
    Sheets("YDWK3").Range("E41").Value = Sheets("Sheet3").Range("N13").Value
    Sheets("YDWK3").Range("E42").Value = Sheets("Sheet3").Range("O13").Value
    Sheets("YDWK3").Calculate
    Sheets("Sheet3").Range("O19").Value = Sheets("YDWK3").Range("F37").Value
    
    Sheets("YDWK3").Activate
    Range("E39:E42").Value = 0
    Range("E81").FormulaR1C1 = "=R47C[-2]+(1-R79C)*R46C[-2]*R67C*(ACOS(R62C)+R79C*SIN(ACOS(R62C))*(R69C+R79C*R62C*(-1+2*R69C*R69C)))"
    Sheets("YDWK3").Calculate
    Sheets("YDWK3").Range("E39").Value = Sheets("Sheet3").Range("I2").Value
    Sheets("YDWK3").Range("E40").Value = Sheets("Sheet3").Range("J2").Value
    Sheets("YDWK3").Range("E41").Value = Sheets("Sheet3").Range("T10").Value
    Sheets("YDWK3").Range("E42").Value = Sheets("Sheet3").Range("U10").Value
    Sheets("YDWK3").Calculate
    Sheets("Sheet3").Range("U15").Value = Sheets("YDWK3").Range("F37").Value
    
    Sheets("YDWK3").Range("E41").Value = Sheets("Sheet3").Range("T11").Value
    Sheets("YDWK3").Range("E42").Value = Sheets("Sheet3").Range("U11").Value
    Sheets("YDWK3").Calculate
    Sheets("Sheet3").Range("U16").Value = Sheets("YDWK3").Range("F37").Value
    
    Sheets("YDWK3").Range("E41").Value = Sheets("Sheet3").Range("T12").Value
    Sheets("YDWK3").Range("E42").Value = Sheets("Sheet3").Range("U12").Value
    Sheets("YDWK3").Calculate
    Sheets("Sheet3").Range("U18").Value = Sheets("YDWK3").Range("F37").Value
    
    Sheets("YDWK3").Range("E41").Value = Sheets("Sheet3").Range("T13").Value
    Sheets("YDWK3").Range("E42").Value = Sheets("Sheet3").Range("U13").Value
    Sheets("YDWK3").Calculate
    Sheets("Sheet3").Range("U19").Value = Sheets("YDWK3").Range("F37").Value
    
    Sheets("YDWK3").Activate
    Range("E39:E42").Value = 0
    Range("E81").FormulaR1C1 = "=R47C[-2]+(1-R79C)*R46C[-2]*R67C*(ACOS(R62C)+R79C*SIN(ACOS(R62C))*(R69C+R79C*R62C*(-1+2*R69C*R69C)))"
    Sheets("YDWK3").Calculate
    Sheets("Sheet3").Activate
    Range("J15:J16").FormulaR1C1 = "=SUM(RC[-2],RC[-1],R3C6)"
    Range("J18:J19").FormulaR1C1 = "=SUM(RC[-2],RC[-1],R3C6)"
    Range("P15:P16").FormulaR1C1 = "=SUM(RC[-2],RC[-1],R4C6)"
    Range("P18:P19").FormulaR1C1 = "=SUM(RC[-2],RC[-1],R4C6)"
    Range("V15:V16").FormulaR1C1 = "=SUM(RC[-2],RC[-1],R2C6)"
    Range("V18:V19").FormulaR1C1 = "=SUM(RC[-2],RC[-1],R2C6)"
    
    Range("J15:J19").Value = Range("J15:J19").Value
    Range("P15:P19").Value = Range("P15:P19").Value
    Range("V15:V19").Value = Range("V15:V19").Value
    
    Range("K15:K16").FormulaR1C1 = _
        "=IF(OR(AND(RC[-1]<750,RC[-1]>R5C6,MIN(RC[-3],RC[-2],R3C6)>=0.28*RC[-1]),AND(RC[-1]>=750,RC[-1]>R5C6,MIN(RC[-3],RC[-2],R3C6)>=0.25*RC[-1],MAX(RC[-3],RC[-2],R3C6)<=0.45*RC[-1])),""OK"",""NO"")"
    Range("K18:K19").FormulaR1C1 = _
        "=IF(OR(AND(RC[-1]<750,RC[-1]>R5C6,MIN(RC[-3],RC[-2],R3C6)>=0.28*RC[-1]),AND(RC[-1]>=750,RC[-1]>R5C6,MIN(RC[-3],RC[-2],R3C6)>=0.25*RC[-1],MAX(RC[-3],RC[-2],R3C6)<=0.45*RC[-1])),""OK"",""NO"")"
    Range("Q15:Q16").FormulaR1C1 = _
        "=IF(OR(AND(RC[-1]<750,RC[-1]>R5C6,MIN(RC[-3],RC[-2],R4C6)>=0.28*RC[-1]),AND(RC[-1]>=750,RC[-1]>R5C6,MIN(RC[-3],RC[-2],R4C6)>=0.25*RC[-1],MAX(RC[-3],RC[-2],R4C6)<=0.45*RC[-1])),""OK"",""NO"")"
    Range("Q18:Q19").FormulaR1C1 = _
        "=IF(OR(AND(RC[-1]<750,RC[-1]>R5C6,MIN(RC[-3],RC[-2],R4C6)>=0.28*RC[-1]),AND(RC[-1]>=750,RC[-1]>R5C6,MIN(RC[-3],RC[-2],R4C6)>=0.25*RC[-1],MAX(RC[-3],RC[-2],R4C6)<=0.45*RC[-1])),""OK"",""NO"")"
    Range("W15:W16").FormulaR1C1 = _
        "=IF(OR(AND(RC[-1]<750,RC[-1]>R5C6,MIN(RC[-3],RC[-2],R2C6)>=0.28*RC[-1]),AND(RC[-1]>=750,RC[-1]>R5C6,MIN(RC[-3],RC[-2],R2C6)>=0.25*RC[-1],MAX(RC[-3],RC[-2],R2C6)<=0.45*RC[-1])),""OK"",""NO"")"
    Range("W18:W19").FormulaR1C1 = _
        "=IF(OR(AND(RC[-1]<750,RC[-1]>R5C6,MIN(RC[-3],RC[-2],R2C6)>=0.28*RC[-1]),AND(RC[-1]>=750,RC[-1]>R5C6,MIN(RC[-3],RC[-2],R2C6)>=0.25*RC[-1],MAX(RC[-3],RC[-2],R2C6)<=0.45*RC[-1])),""OK"",""NO"")"
    
    Range("G15:K19").Value = Range("G15:K19").Value
    Range("G20:K24").Value = Range("M15:Q19").Value
    'Range("M14:Q16").Clear
    Range("G25:K29").Value = Range("S15:W19").Value
    'Range("S14:W16").Clear
    
    Range("G15:K29").Sort Key1:=Range("G15"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    
    Range("M15:W19").Clear
    
    Range("L15:L26").FormulaR1C1 = "=IF(RC[-1]=""OK"",RC[-5],"""")"
    Range("M15:P18").FormulaR1C1 = "=IF(RC[-1]<>"""",R[-5]C[-5],"""")"
    Range("M19:P22").FormulaR1C1 = "=IF(RC[-1]<>"""",R[-9]C[1],"""")"
    Range("M23:P26").FormulaR1C1 = "=IF(RC[-1]<>"""",R[-13]C[7],"""")"
    
    Range("Q15:Q26").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-7],"""")"
    Range("R15:R26").FormulaR1C1 = "=IF(RC[-1]=MAX(R15C[-1]:R26C[-1]),1,"""")"
    Range("L15:R26").Value = Range("L15:R26").Value
    
    Range("L15:R26").Sort Key1:=Range("R15"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    Range("N2").FormulaR1C1 = "=IF(OR(R[13]C[-2]=R[8]C[-7],R[13]C[-2]=R[9]C[-7],R[13]C[-2]=R[10]C[-7],R[13]C[-2]=R[11]C[-7]),R[13]C[-2],RC[-6])"
    Range("O2").FormulaR1C1 = "=IF(RC14=R15C12,R15C[-2],RC[-6])"
    Range("P2").FormulaR1C1 = "=IF(RC14=R15C12,R15C[-2],RC[-6])"
    Range("Q2").FormulaR1C1 = "=IF(RC14=R15C12,R15C[-2],RC[-6])"
    Range("R2").FormulaR1C1 = "=IF(RC14=R15C12,R15C[-2],RC[-6])"
    
    Range("N3").FormulaR1C1 = "=IF(OR(R[12]C[-2]=R[7]C[-1],R[12]C[-2]=R[8]C[-1],R[12]C[-2]=R[9]C[-1],R[12]C[-2]=R[10]C[-1]),R[12]C[-2],RC[-6])"
    Range("O3").FormulaR1C1 = "=IF(RC14=R15C12,R15C[-2],RC[-6])"
    Range("P3").FormulaR1C1 = "=IF(RC14=R15C12,R15C[-2],RC[-6])"
    Range("Q3").FormulaR1C1 = "=IF(RC14=R15C12,R15C[-2],RC[-6])"
    Range("R3").FormulaR1C1 = "=IF(RC14=R15C12,R15C[-2],RC[-6])"
    
    Range("N4").FormulaR1C1 = "=IF(OR(R[11]C[-2]=R[6]C[5],R[11]C[-2]=R[7]C[5],R[11]C[-2]=R[8]C[5],R[11]C[-2]=R[9]C[5]),R[11]C[-2],RC[-6])"
    Range("O4").FormulaR1C1 = "=IF(RC14=R15C12,R15C[-2],RC[-6])"
    Range("P4").FormulaR1C1 = "=IF(RC14=R15C12,R15C[-2],RC[-6])"
    Range("Q4").FormulaR1C1 = "=IF(RC14=R15C12,R15C[-2],RC[-6])"
    Range("R4").FormulaR1C1 = "=IF(RC14=R15C12,R15C[-2],RC[-6])"
    
    Range("N2:R4").Value = Range("N2:R4").Value
    
    Sheets("YDWK3").Range("E39").Value = Sheets("Sheet3").Range("O2").Value
    Sheets("YDWK3").Range("E40").Value = Sheets("Sheet3").Range("P2").Value
    Sheets("YDWK3").Range("E41").Value = Sheets("Sheet3").Range("O3").Value
    Sheets("YDWK3").Range("E42").Value = Sheets("Sheet3").Range("P3").Value
    Sheets("YDWK3").Calculate
    Sheets("Sheet3").Range("S2").Value = Sheets("YDWK3").Range("F37").Value
    
    Sheets("YDWK3").Range("E39").Value = Sheets("Sheet3").Range("O4").Value
    Sheets("YDWK3").Range("E40").Value = Sheets("Sheet3").Range("P4").Value
    Sheets("YDWK3").Calculate
    Sheets("Sheet3").Range("S3").Value = Sheets("YDWK3").Range("F37").Value
    
    Sheets("YDWK3").Range("E41").Value = Sheets("Sheet3").Range("O2").Value
    Sheets("YDWK3").Range("E42").Value = Sheets("Sheet3").Range("P2").Value
    Sheets("YDWK3").Calculate
    Sheets("Sheet3").Range("S4").Value = Sheets("YDWK3").Range("F37").Value
    
    Application.Run "F.xlsm!YDWK3clear"
    Sheets("Sheet3").Activate
    
    Range("S5").FormulaR1C1 = "=SUM(R[-3]C:R[-1]C)"
    Range("S6").FormulaR1C1 = "=IF(R[-5]C[-15]-R[-1]C[-15]<=R9C4,R[-1]C,R[-1]C-((R[-5]C[-15]-R[-1]C[-15]-R9C4)*0.1))"
    Range("S5:S6").Value = Range("S5:S6").Value
    If Range("S6") > Range("F5") Then
        Range("A2:F4").Value = Range("N2:S4").Value
        Range("F5").Value = Range("S6").Value
    End If
    
End If
 
 If Range("A1") <> "This is a 2-Turn Point Triangle" Then
    Sheets("Sheet2").Range("H16:M20").Value = Sheets("Sheet3").Range("A1:F5").Value
 ElseIf Range("A1") = "This is a 2-Turn Point Triangle" Then
    Sheets("Sheet2").Range("H16:M19").Value = Sheets("Sheet3").Range("A2:F5").Value
 End If
    Columns("G:W").Clear
    Range("A1:F8").Clear
    Sheets("Sheet2").Activate
End Sub