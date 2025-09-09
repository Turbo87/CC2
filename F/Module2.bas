' VBA Module: Turn Point and Distance Calculations
' Purpose: Manages turn point analysis and distance calculations for flight path optimization.
' Handles start turn point accuracy, leg distance calculations, and geographic coordinate
' processing using spherical trigonometry for flight route analysis.

Option Explicit
#If VBA7 And Win64 Then
    Dim LastRow As LongPtr
#Else
    Dim LastRow As Long
#End If

Sub NewOR1()
'
' Ck STDstartTP (all options) AMENDED 1/19/14 for start accuracy Schmelzer tested w/JP 2013 OK Amended 7/24/2017 for brevity See '''
' JLR 4/5/20218 Amended for SibLap, reversed Leg ck order; Matrix includes LoH; 4/17/2018 Amended Last Leg calc N10 relative to max Leg 1
' JLR 5/15/2018 Matric revised""; 5/16/2018 Matrix revised
'
    Application.ScreenUpdating = False
If Range("A10") = Range("A1") Then
    Application.Run "F.xlsm!NewOR2"
    Exit Sub
ElseIf Range("A10") < Range("A1") Then
    
    'First Leg longest 1000 if >1000, longest 500 if 500 - 999, > medidan if < 500
    Range("F10:F10009").FormulaR1C1 = "=IF(AND(RC[-5]<>"""",RC[-5]<R1C1),6371*ACOS(SIN(RC[-4])*SIN(R1C2)+COS(RC[-4])*COS(R1C2)*COS(R1C3-RC[-3])),"""")"
    
    Range("F7").FormulaR1C1 = "=MAX(R[3]C:R[10002]C)"
    Range("F8").FormulaR1C1 = "=IF(R1009C6<>"""",LARGE(R[2]C:R[10001]C,1000),IF(R509C6<>"""",LARGE(R[2]C:R[501]C,500),MEDIAN(R[2]C:R[501]C)))"
    ActiveSheet.Calculate
    Range("F7:F10009").Value = Range("F7:F10009").Value
    Range("G10").FormulaR1C1 = "=IF(AND(RC[-1]<>"""",RC[-1]>=R8C6),RC[-6],"""")"
    Range("H10:L10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-6],"""")"
    'Range("L10").FormulaR1C1 = "=IF(RC[-1]<>"""",6371*ACOS(SIN(RC[-10])*SIN(R1C2)+COS(RC[-10])*COS(R1C2)*COS(R1C3-RC[-9])),"""")"
    'Copy Ref A Value
    Application.Calculation = xlCalculationManual
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:L10").AutoFill Destination:=.Range("G10:L" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("G10:L10009").Value = Range("G10:L10009").Value
    Range("G10:L10009").Sort Key1:=Range("G10"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    Range("F10:F10009").Clear
    
    'LAST Legs >= nth First Leg
    Range("M10:M10009").FormulaR1C1 = "=IF(AND(RC[-12]<>"""",RC[-12]>R1C1),6371*ACOS(SIN(RC[-11])*SIN(R1C2)+COS(RC[-11])*COS(R1C2)*COS(R1C3-RC[-10])),"""")"
    Range("M8").Value = Range("F8").Value
    ActiveSheet.Calculate
    
    Range("N10").FormulaR1C1 = "=IF(AND(RC[-13]>R1C1,RC[-1]<>"""",RC[-1]>=R8C13,RC[-1]<=R7C6),RC[-13],"""")"
    Range("O10:R10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-13],"""")"
    Range("S10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-6],"""")"
    'Copy Ref A Value Sort
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("N10:S10").AutoFill Destination:=.Range("N10:S" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    
    Range("N10:S10009").Value = Range("N10:S10009").Value
    Range("N10:S10009").Sort Key1:=Range("N10"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    Range("M10:M10009").Clear
    
    ''''NEXT FOUR LINES 7/24/2017; amended w/conditional 8/27/17; amended 4/15/2018
    If Range("G500") <> "" Then
        Range("L8").FormulaR1C1 = "=2*(RC[1]-5)"
        ActiveSheet.Calculate
        Range("L8").Value = Range("L8").Value
    ElseIf Range("G500") = "" Then
        Range("L8").FormulaR1C1 = "=2*(MAX(R[2]C:R[10001]C))"
        ActiveSheet.Calculate
        Range("L8").Value = Range("L8").Value
        If Range("L8") < 50 Then
           Range("G8:L500").Clear
           Application.Run "F.xlsm!NewOR2"
           Exit Sub
        End If
    End If
    
    Range("G10:L509").Copy
    Range("U1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    
    'MATRIX
    'Range("U10:SZ10").FormulaR1C1 = "=IF(OR(R1C="""",RC16=R3C),"""",IF(OR(R6C>RC19,6371*ACOS(SIN(RC15)*SIN(R2C)+COS(RC15)*COS(R2C)*COS(R3C-RC16))>0.5),"""",IF(R4C-RC17<=R9C4,2*R6C,2*R6C-((R4C-RC17-R9C4)*0.1))))"
    ''Range("U10:SZ10").FormulaR1C1 = "=IF(OR(R1C="""",RC16=R3C),"""",IF(OR(R6C>RC19,R[1]C19<RC19,6371*ACOS(SIN(RC15)*SIN(R2C)+COS(RC15)*COS(R2C)*COS(R3C-RC16))>0.5),"""",IF(R4C-RC17<=R9C4,2*R6C,2*R6C-((R4C-RC17-R9C4)*0.1))))"
    'Range("U10:SZ10").FormulaR1C1 = _
        "=IF(OR(R1C="""",RC16=R3C),"""",IF(OR(R6C>RC19,6371*ACOS(SIN(RC15)*SIN(R2C)+COS(RC15)*COS(R2C)*COS(R3C-RC16))>0.5),"""",2*R6C))"
    'Range("U10:SZ10").FormulaR1C1 = "=IF(OR(R1C="""",RC16=R3C),"""",IF(OR(R6C>RC19, AND(R[-1]C19<RC19,R[1]C19<RC19),6371*ACOS(SIN(RC15)*SIN(R2C)+COS(RC15)*COS(R2C)*COS(R3C-RC16))>0.5),"""",IF(R4C-RC17<=R9C4,2*R6C,2*R6C-((R4C-RC17-R9C4)*0.1))))"
    ''Range("U10:SZ10").FormulaR1C1 = "=IF(OR(R1C="""",RC16=R3C),"""",IF(OR(R6C>RC19, AND(R[-1]C19<RC19,R[1]C19<RC19),6371*ACOS(SIN(RC15)*SIN(R2C)+COS(RC15)*COS(R2C)*COS(R3C-RC16))>0.5),"""",IF(R4C-RC17<=R9C4,R6C+RC19,(R6C+RC19)-((R4C-RC17-R9C4)*0.1))))"
    Range("U10:SZ10").FormulaR1C1 = "=IF(OR(R1C="""",RC16=R3C),"""",IF(OR(R6C>RC19,R[1]C19<RC19,6371*ACOS(SIN(RC15)*SIN(R2C)+COS(RC15)*COS(R2C)*COS(R3C-RC16))>0.5),"""",IF(R4C-RC17<=R9C4,R6C+RC19,(R6C+RC19)-((R4C-RC17-R9C4)*0.1))))"
    'Copy Ref N
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "N").End(xlUp).Row
.Range("U10:SZ10").AutoFill Destination:=.Range("U10:SZ" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("N1").FormulaR1C1 = "=MAX(R10C21:R10009C520)"
    Range("O1").FormulaR1C1 = "=MAX(R7C21:R7C520)"
    Range("P1").FormulaR1C1 = "=MIN(R10C20:R10009C20)"

    Range("U7:SZ7").FormulaR1C1 = "=IF(AND(R1C14<>0,MAX(R[3]C:R[10002]C)=R1C14),R[-6]C,"""")"
    Range("T10:T10009").FormulaR1C1 = "=IF(AND(R1C14<>0,MAX(RC[1]:RC[500])=R1C14),RC[-6],"""")"
    ActiveSheet.Calculate
    Range("N1:P1").Value = Range("N1:P1").Value
    
    If Range("G510") <> "" Then
        Range("G510:L1009").Copy
        Range("U1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
        Range("U7:SZ7").FormulaR1C1 = "=IF(AND(R2C14<>0,MAX(R[3]C:R[10002]C)=R2C14),R[-6]C,"""")"
        Range("T10:T10009").FormulaR1C1 = "=IF(AND(R2C14<>0,MAX(RC[1]:RC[500])=R2C14),RC[-6],"""")"
        Range("N2").FormulaR1C1 = "=MAX(R10C21:R10009C520)"
        Range("O2").FormulaR1C1 = "=MAX(R7C21:R7C520)"
        Range("P2").FormulaR1C1 = "=MIN(R10C20:R10009C20)"
        ActiveSheet.Calculate
        Range("N2:P2").Value = Range("N2:P2").Value
    Else: Range("N2:P2").Clear
    End If
    
    If Range("G1010") <> "" Then
        Range("G1010:L1509").Copy
        Range("U1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
        Range("U7:SZ7").FormulaR1C1 = "=IF(AND(R3C14<>0,MAX(R[3]C:R[10002]C)=R3C14),R[-6]C,"""")"
        Range("T10:T10009").FormulaR1C1 = "=IF(AND(R3C14<>0,MAX(RC[1]:RC[500])=R3C14),RC[-6],"""")"
        Range("N3").FormulaR1C1 = "=MAX(R10C21:R10009C520)"
        Range("O3").FormulaR1C1 = "=MAX(R7C21:R7C520)"
        Range("P3").FormulaR1C1 = "=MIN(R10C20:R10009C20)"
        ActiveSheet.Calculate
        Range("N3:P3").Value = Range("N3:P3").Value
    Else: Range("N3:P3").Clear
    End If
    
    If Range("G1510") <> "" Then
        Range("G1510:L2009").Copy
        Range("U1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
        Range("U7:SZ7").FormulaR1C1 = "=IF(AND(R4C14<>0,MAX(R[3]C:R[10002]C)=R4C14),R[-6]C,"""")"
        Range("T10:T10009").FormulaR1C1 = "=IF(AND(R4C14<>0,MAX(RC[1]:RC[500])=R4C14),RC[-6],"""")"
        Range("N4").FormulaR1C1 = "=MAX(R10C21:R10009C520)"
        Range("O4").FormulaR1C1 = "=MAX(R7C21:R7C520)"
        Range("P4").FormulaR1C1 = "=MIN(R10C20:R10009C20)"
        ActiveSheet.Calculate
        Range("N4:P4").Value = Range("N4:P4").Value
    Else: Range("N4:P4").Clear
    End If
    
    If Range("G2010") <> "" Then
        Range("G2010:L2509").Copy
        Range("U1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
        Range("U7:SZ7").FormulaR1C1 = "=IF(AND(R5C14<>0,MAX(R[3]C:R[10002]C)=R5C14),R[-6]C,"""")"
        Range("T10:T10009").FormulaR1C1 = "=IF(AND(R5C14<>0,MAX(RC[1]:RC[500])=R5C14),RC[-6],"""")"
        Range("N5").FormulaR1C1 = "=MAX(R10C21:R10009C520)"
        Range("O5").FormulaR1C1 = "=MAX(R7C21:R7C520)"
        Range("P5").FormulaR1C1 = "=MIN(R10C20:R10009C20)"
        ActiveSheet.Calculate
        Range("N5:P5").Value = Range("N5:P5").Value
    Else: Range("N5:P5").Clear
    End If
    
    If Range("G2510") <> "" Then
        Range("G2510:L3009").Copy
        Range("U1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
        Range("U7:SZ7").FormulaR1C1 = "=IF(AND(R1C17<>0,MAX(R[3]C:R[10002]C)=R1C17),R[-6]C,"""")"
        Range("T10:T10009").FormulaR1C1 = "=IF(AND(R1C17<>0,MAX(RC[1]:RC[500])=R1C17),RC[-6],"""")"
        Range("Q1").FormulaR1C1 = "=MAX(R10C21:R10009C520)"
        Range("R1").FormulaR1C1 = "=MAX(R7C21:R7C520)"
        Range("S1").FormulaR1C1 = "=MIN(R10C20:R10009C20)"
        ActiveSheet.Calculate
        Range("Q1:S1").Value = Range("Q1:S1").Value
    Else: Range("Q1:S1").Clear
    End If
    
    If Range("G3010") <> "" Then
        Range("G3010:L3509").Copy
        Range("U1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
        Range("U7:SZ7").FormulaR1C1 = "=IF(AND(R2C17<>0,MAX(R[3]C:R[10002]C)=R2C17),R[-6]C,"""")"
        Range("T10:T0009").FormulaR1C1 = "=IF(AND(R2C17<>0,MAX(RC[1]:RC[500])=R2C17),RC[-6],"""")"
        Range("Q2").FormulaR1C1 = "=MAX(R10C21:R10009C520)"
        Range("R2").FormulaR1C1 = "=MAX(R7C21:R7C520)"
        Range("S2").FormulaR1C1 = "=MIN(R10C20:R10009C20)"
        ActiveSheet.Calculate
        Range("Q2:S2").Value = Range("Q2:S2").Value
    Else: Range("Q2:S2").Clear
    End If
    
    If Range("G3510") <> "" Then
        Range("G3510:L4009").Copy
        Range("U1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
        Range("U7:SZ7").FormulaR1C1 = "=IF(AND(R3C17<>0,MAX(R[3]C:R[10002]C)=R3C17),R[-6]C,"""")"
        Range("T10:T10009").FormulaR1C1 = "=IF(AND(R3C17<>0,MAX(RC[1]:RC[500])=R3C17),RC[-6],"""")"
        Range("Q3").FormulaR1C1 = "=MAX(R10C21:R10009C520)"
        Range("R3").FormulaR1C1 = "=MAX(R7C21:R7C520)"
        Range("S3").FormulaR1C1 = "=MIN(R10C20:R10009C20)"
        ActiveSheet.Calculate
        Range("Q3:S3").Value = Range("Q3:S3").Value
    Else: Range("Q3:S3").Clear
    End If
    
    If Range("G4010") <> "" Then
        Range("G4010:L4509").Copy
        Range("U1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
        Range("U7:SZ7").FormulaR1C1 = "=IF(AND(R4C17<>0,MAX(R[3]C:R[10002]C)=R4C17),R[-6]C,"""")"
        Range("T10:T10009").FormulaR1C1 = "=IF(AND(R4C17<>0,MAX(RC[1]:RC[500])=R4C17),RC[-6],"""")"
        Range("Q4").FormulaR1C1 = "=MAX(R10C21:R10009C520)"
        Range("R4").FormulaR1C1 = "=MAX(R7C21:R7C520)"
        Range("S4").FormulaR1C1 = "=MIN(R10C20:R10009C20)"
        ActiveSheet.Calculate
        Range("Q4:S4").Value = Range("Q4:S4").Value
    Else: Range("Q4:S4").Clear
    End If
    
    If Range("G4510") <> "" Then
        Range("G4510:L5009").Copy
        Range("U1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
        Range("U7:SZ7").FormulaR1C1 = "=IF(AND(R5C17<>0,MAX(R[3]C:R[10002]C)=R5C17),R[-6]C,"""")"
        Range("T10:T10009").FormulaR1C1 = "=IF(AND(R5C17<>0,MAX(RC[1]:RC[500])=R5C17),RC[-6],"""")"
        Range("Q5").FormulaR1C1 = "=MAX(R10C21:R10009C520)"
        Range("R5").FormulaR1C1 = "=MAX(R7C21:R7C520)"
        Range("S5").FormulaR1C1 = "=MIN(R10C20:R10009C20)"
        ActiveSheet.Calculate
        Range("Q5:S5").Value = Range("Q5:S5").Value
    Else: Range("Q5:S5").Clear
    End If
    
    Range("L8,L9,U1:SZ10009").Clear
    Range("G10:T10009").Clear
    
    Range("T1:T5").FormulaR1C1 = "=IF(MAX(R1C14:R5C14,R1C17:R5C17)=RC[-6],RC[-5],IF(MAX(R1C14:R5C14,R1C17:R5C17)=RC[-3],RC[-2],""""))"
    Range("U1:U5").FormulaR1C1 = "=IF(RC[-1]=RC[-6],RC[-5],IF(RC[-1]=RC[-4],RC[-2],""""))"
    
    Range("H1").FormulaR1C1 = "=MAX(RC[12]:R[4]C[12])"
    ActiveSheet.Calculate
    Range("H1").Value = Range("H1").Value
    Range("H3").FormulaR1C1 = "=MAX(R[-2]C[13]:R[2]C[13])"
    ActiveSheet.Calculate
    Range("H3").Value = Range("H3").Value
    
    Range("H2:L2").Value = Range("A1:E1").Value
    
    Range("N1:U5").Clear
    
    Range("G7:J7").FormulaR1C1 = "=MAX(R[3]C:R[10002]C)"
    Range("G10").FormulaR1C1 = "=IF(RC[-6]=R1C8,RC[-5],"""")"
    Range("H10:J10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-5],"""")"
    'Copy Ref A
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:J10").AutoFill Destination:=.Range("G10:J" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("I1:L1").Value = Range("G7:J7").Value
    
    Range("G10:G10009").FormulaR1C1 = "=IF(RC[-6]=R3C8,RC[-5],"""")"
    ActiveSheet.Calculate
    Range("I3:L3").Value = Range("G7:J7").Value
    
    Range("F8,M8,G7:J10009").Clear
    
    Range("M2").FormulaR1C1 = _
        "=IF(AND(R[-1]C[-5]<>0,R[1]C[-5]<>0),2*6371*ACOS(SIN(R[-1]C[-4])*SIN(RC[-4])+COS(R[-1]C[-4])*COS(RC[-4])*COS(RC[-3]-R[-1]C[-3])),0)"
    Range("M2").Value = Range("M2").Value
    
 End If

Application.Run "F.xlsm!AORc"

End Sub

Sub AORc()
'
' 10/20/13 needed for JP 2013 REVISED 2/1/15 (lingering data, amended matrix)
' 5/15/18 Revised Matrix""; 5/16/2018 Revised Matrix
'
Application.ScreenUpdating = False
    Range("A5").Value = 0.041666667
    'Save prior TP
    Range("A7:E7").Value = Range("H2:L2").Value
    Range("G7").FormulaR1C1 = "=MIN(R[3]C:R[10012]C)"
    Range("H7:L7").FormulaR1C1 = "=MAX(R[3]C:R[10012]C)"
    Range("G10").FormulaR1C1 = _
        "=IF(RC[-4]=R2C3,"""",IF(RC1<R2C1-R5C1,6371*ACOS(SIN(RC[-5])*SIN(R2C2)+COS(RC[-5])*COS(R2C2)*COS(R2C3-RC[-4])),""""))"
    Range("H10").FormulaR1C1 = "=IF(RC[-1]=R[-3]C[-1],RC[-7],"""")"
    Range("H10").FormulaR1C1 = "=IF(RC[-1]=R7C7,RC[-7],"""")"
    Range("I10:L10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-7],"""")"
    
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:L10").AutoFill Destination:=.Range("G10:L" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
  If Range("G7") = 0 Or Range("H7") = Range("H1") Then
      Range("H2:L2").Value = Range("A7:E7").Value
      Range("A5:E7,G7:L10009").Clear
      
  ElseIf Range("G7") <> 0 Then
    Range("G2:L2").Value = Range("G7:L7").Value

    'Fini
    Range("G10").FormulaR1C1 = _
        "=IF(RC[-4]=R1C3,"""",IF(AND(RC[-6]>R2C8,RC[-6]<R2C1),6371*ACOS(SIN(RC[-5])*SIN(R1C2)+COS(RC[-5])*COS(R1C2)*COS(R1C3-RC[-4])),""""))"
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:G10").AutoFill Destination:=.Range("G10:G" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("G7:L7").Value = Range("G7:L7").Value
    Range("G10:H10019").Clear
    Range("A4").Value = 0.003472222
      
    Range("G10").FormulaR1C1 = _
        "=IF(RC[-4]=R2C10,"""",IF(AND(RC[-6]>=R7C8-R4C1,RC[-6]<=R7C8+R4C1),6371*ACOS(SIN(RC[-5])*SIN(R2C9)+COS(RC[-5])*COS(R2C9)*COS(R2C10-RC[-4])),""""))"
    Range("H10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-7],"""")"
    
    Application.Calculation = xlCalculationManual
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:H10").AutoFill Destination:=.Range("G10:H" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("G10:L10009").Value = Range("G10:L10009").Value
    Range("G10:L10009").Sort Key1:=Range("H10"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal

    Range("H10:L509").Copy
    Range("N1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    Application.CutCopyMode = False
    
    Range("G10:G509").Copy
    Range("N6").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    Application.CutCopyMode = False
    
    Range("G10:L10009").Clear

'Starts
    Range("L6").FormulaR1C1 = "=MAX(RC[2]:RC[501])"
    Range("M6").FormulaR1C1 = "=MIN(RC[1]:RC[500])"
    Range("L6:M6").Value = Range("L6:M6").Value
    
    Range("G10").FormulaR1C1 = _
        "=IF(RC[-4]=R2C10,"""",IF(AND(RC[-6]<R2C8,6371*ACOS(SIN(RC[-5])*SIN(R2C9)+COS(RC[-5])*COS(R2C9)*COS(R2C10-RC[-4]))>R6C13,6371*ACOS(SIN(RC[-5])*SIN(R2C9)+COS(RC[-5])*COS(R2C9)*COS(R2C10-RC[-4]))<=R6C12+1),RC[-6],""""))"
    Range("H10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-6],"""")"
    Range("I10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-6],"""")"
    Range("J10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-6],"""")"
    Range("K10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-6],"""")"
    Range("L10").FormulaR1C1 = _
        "=IF(RC[-1]<>"""",6371*ACOS(SIN(RC[-10])*SIN(R2C9)+COS(RC[-10])*COS(R2C9)*COS(R2C10-RC[-9])),"""")"
    
     Application.Calculation = xlCalculationManual
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:L10").AutoFill Destination:=.Range("G10:L" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("G10:L10009").Value = Range("G10:L10009").Value
    Range("G10:L10009").Sort Key1:=Range("G10"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    
'Revert to Aorb if no Start found
If Range("G10") = "" Then
    Range("H2:L2").Value = Range("A7:E7").Value
    Range("A7:E7,G2,G5:M10009").Clear
    Columns("N:SS").Clear
    
ElseIf Range("G10") <> "" Then
'Matrix Amended 2/1/15 to add R1C<R2C8)
    'Range("N10:SS10").FormulaR1C1 = _
        "=IF(OR(R1C="""",R3C=RC9,R1C<R2C8),"""",IF(OR(6371*ACOS(SIN(R2C)*SIN(RC8)+COS(R2C)*COS(RC8)*COS(RC9-R3C))>0.5,AND(R6C>RC12,R6C<SQRT(RC12^2+0.25)),RC12>R6C),"""",IF(RC10-R4C<=R9C4,RC12,"""")))"
    ''Range("N10:SS10").FormulaR1C1 = "=IF(OR(R1C="""",R3C=RC9,R1C<R2C8),"""",IF(OR(6371*ACOS(SIN(R2C)*SIN(RC8)+COS(R2C)*COS(RC8)*COS(RC9-R3C))>0.5,AND(R6C>RC12,R6C<SQRT(RC12^2+0.25)),RC12>R6C,R6C[1]<R6C),"""",IF(RC10-R4C<=R9C4,RC12,"""")))"
    ''Range("N10:SS10").FormulaR1C1 = "=IF(OR(R1C="""",R3C=RC9,R1C<R2C8),"""",IF(OR(6371*ACOS(SIN(R2C)*SIN(RC8)+COS(R2C)*COS(RC8)*COS(RC9-R3C))>0.5,AND(R6C>RC12,R6C<SQRT(RC12^2+0.25)),RC12>R6C,AND(R6C[-1]<R6C,R6C[1]<R6C)),"""",IF(RC10-R4C<=R9C4,RC12,"""")))"
    Range("N10:SS10").FormulaR1C1 = "=IF(OR(R1C="""",R3C=RC9,R1C<R2C8),"""",IF(OR(6371*ACOS(SIN(R2C)*SIN(RC8)+COS(R2C)*COS(RC8)*COS(RC9-R3C))>0.5,AND(R6C>RC12,R6C<SQRT(RC12^2+0.25)),RC12>R6C,R6C[1]<R6C),"""",IF(RC10-R4C<=R9C4,RC12,"""")))"
    'Copy Ref G
    Application.Calculation = xlCalculationManual
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
.Range("N10:SS10").AutoFill Destination:=.Range("N10:SS" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    
    Range("I5").FormulaR1C1 = "=MAX(R[5]C[5]:R[10014]C[504])"
    Range("H5").FormulaR1C1 = "=2*RC[1]"
    Range("H5:I5").Value = Range("H5:I5").Value
    
  'Revert to Aorb if no Finish or shorter distance found
  If Range("H5") = 0 Or Range("H5") <= Range("M2") Then
    Range("H2:L2").Value = Range("A7:E7").Value
    Range("G2,G5:M10019").Clear
    Columns("N:SS").Clear
  
  ElseIf Range("H5") > Range("M2") Then
  
    Range("N10:SS3000").Value = Range("N10:SS3000").Value
    
    Range("N8:SS8").FormulaR1C1 = _
        "=IF(R[-7]C="""","""",IF(MAX(R[2]C:R[2992]C)=R5C9,R1C,""""))"
    Range("N8:SS8").Value = Range("N8:SS8").Value
    Range("M10").FormulaR1C1 = "=IF(MAX(RC[1]:RC[500])=R5C9,RC[-6],"""")"
    
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
.Range("M10:M10").AutoFill Destination:=.Range("M10:M" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("M10:M3009").Value = Range("M10:M3009").Value
    
    Range("H1").FormulaR1C1 = "=MAX(R[9]C[5]:R[2999]C[5])"
    Range("H3").FormulaR1C1 = "=MIN(R[5]C[6]:R[5]C[505])"
    Range("H1:H3").Value = Range("H1:H3").Value
    Range("N10:SS10009").Clear
                
    Range("N9:SS9").FormulaR1C1 = "=IF(R1C="""","""",IF(R[-1]C=R3C8,R2C,""""))"
    Range("N10:SS10").FormulaR1C1 = "=IF(R9C<>"""",R3C,"""")"
    Range("N11:SS11").FormulaR1C1 = "=IF(R10C<>"""",R4C,"""")"
    Range("N12:SS12").FormulaR1C1 = "=IF(R11C<>"""",R5C,"""")"
    
    Range("I3").FormulaR1C1 = "=MAX(R[6]C[5]:R[6]C[504])"
    Range("J3").FormulaR1C1 = "=MAX(R[7]C[4]:R[7]C[503])"
    Range("K3").FormulaR1C1 = "=MAX(R[8]C[3]:R[8]C[502])"
    Range("L3").FormulaR1C1 = "=MAX(R[9]C[2]:R[9]C[501])"
    Range("I3:L3").Value = Range("I3:L3").Value
    
    Range("N9:SS12").Clear
    
    Range("N10:Q10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-6],"""")"
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
.Range("N10:Q10").AutoFill Destination:=.Range("N10:Q" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("N10:Q3009").Value = Range("N10:Q3009").Value
    
    Range("I1:L1").FormulaR1C1 = "=MAX(R[9]C[5]:R[2999]C[5])"
    Range("I1:L1").Value = Range("I1:L1").Value
    
    Range("A4:E7,G2,G5:M3000,N1:SS3000").Clear
    
    Range("M2").FormulaR1C1 = _
        "=IF(OR(SUM(R1C8:R1C12)=0,SUM(R2C8:R2C12)=0),0,2*6371*ACOS(SIN(R[-1]C[-4])*SIN(RC[-4])+COS(R[-1]C[-4])*COS(RC[-4])*COS(RC[-3]-R[-1]C[-3])))"
    Range("M2").Value = Range("M2").Value
    End If
  End If
End If

Application.Run "F.xlsm!NewOR2"

End Sub

Sub NewOR2()
'
' OR2redux Macro
' JLR 4/5/2018 For speed?
' JLR 5/15/18 Revised Matrix""; 5/16/2018 Revised Matrix
'
    Range("M10").FormulaR1C1 = "=IF(AND(RC[-12]<>"""",RC[-12]>R2C1),6371*ACOS(SIN(RC[-11])*SIN(R2C2)+COS(RC[-11])*COS(R2C2)*COS(R2C3-RC[-10])),"""")"
    Range("L10").FormulaR1C1 = "=IF(RC[1]<>"""",1,"""")"
    Range("L8").FormulaR1C1 = "=SUM(R[2]C:R[10001]C)"
    'Copy Ref A
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("L10:M10").AutoFill Destination:=.Range("L10:M" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    
    Range("M8").FormulaR1C1 = "=(2*(MAX(R[1]C:R[10000]C)))"
    Range("M7").FormulaR1C1 = "=IF(R[1]C[-1]>1000,LARGE(R[3]C:R[10002]C,1000),IF(AND(R[1]C[-1]<1000,R[1]C[-1]>500),LARGE(R[3]C:R[10002]C,500),1))"

    If Range("M8") > Range("M2") Then
    
    Range("N10").FormulaR1C1 = "=IF(AND(RC[-1]<>"""",RC[-1]>=R7C13,RC[-13]<>""""),RC[-13],"""")"
    Range("O10:R10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-13],"""")"
    Range("S10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-6],"""")"
    'Copy Ref A, Value Sort
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("N10:S10").AutoFill Destination:=.Range("N10:S" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("M10:S10009").Value = Range("M10:S10009").Value
    Range("N10:S10009").Sort Key1:=Range("N10"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    
    Range("M7:M8").Value = Range("M7:M8").Value
    Range("L10:M10009").Clear
    Range("S7").FormulaR1C1 = "=Median(R[3]C:R[1002]C)"
    Range("S8").FormulaR1C1 = "=MIN(R[2]C:R[10001]C)"
    Range("S9").FormulaR1C1 = "=MAX(R[1]C:R[10000]C)"
    ActiveSheet.Calculate
    
    Range("F10:F10009").FormulaR1C1 = "=IF(AND(RC[-5]<>"""",RC[-5]<R2C1),6371*ACOS(SIN(RC[-4])*SIN(R2C2)+COS(RC[-4])*COS(R2C2)*COS(RC[-3]-R2C3)),"""")"
    ActiveSheet.Calculate
    
    Range("G10").FormulaR1C1 = "=IF(AND(RC[-6]<R2C1,RC[-1]>=R8C19,RC[-1]<=R9C19),RC[-6],"""")"
    Range("H10:L10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-6],"""")"
    'Copy Ref A
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:L10").AutoFill Destination:=.Range("G10:L" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("G10:L10009").Value = Range("G10:L10009").Value
    Range("G10:L10009").Sort Key1:=Range("G10"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    
    Range("F10:F10009").Clear
    
    Range("G10:L509").Copy
    Range("U1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    
    'MATRIX
    'Range("U10:SZ10").FormulaR1C1 = _
        "=IF(OR(R1C="""",RC16=R3C),"""",IF(OR(R6C>RC19,6371*ACOS(SIN(RC15)*SIN(R2C)+COS(RC15)*COS(R2C)*COS(R3C-RC16))>0.5),"""",IF(R4C-RC17<=R9C4,2*R6C,2*R6C-((R4C-RC17-R9C4)*0.1))))"
    ''Range("U10:SZ10").FormulaR1C1 = "=IF(OR(R1C="""",RC16=R3C),"""",IF(OR(R6C>RC19,R[1]C19<RC19,6371*ACOS(SIN(RC15)*SIN(R2C)+COS(RC15)*COS(R2C)*COS(R3C-RC16))>0.5),"""",IF(R4C-RC17<=R9C4,2*R6C,2*R6C-((R4C-RC17-R9C4)*0.1))))"
    ''' Range("U10:SZ10").FormulaR1C1 = "=IF(OR(R1C="""",RC16=R3C),"""",IF(OR(R6C>RC19,AND(R[-1]C19<RC19,R[1]C19<RC19),6371*ACOS(SIN(RC15)*SIN(R2C)+COS(RC15)*COS(R2C)*COS(R3C-RC16))>0.5),"""",IF(R4C-RC17<=R9C4,2*R6C,2*R6C-((R4C-RC17-R9C4)*0.1))))"
    ''Range("U10:SZ10").FormulaR1C1 = "=IF(OR(R1C="""",RC16=R3C),"""",IF(OR(R6C>RC19,AND(R[-1]C19<RC19,R[1]C19<RC19),6371*ACOS(SIN(RC15)*SIN(R2C)+COS(RC15)*COS(R2C)*COS(R3C-RC16))>0.5),"""",IF(R4C-RC17<=R9C4,R6C+RC19,(R6C+RC19)-((R4C-RC17-R9C4)*0.1))))"
    Range("U10:SZ10").FormulaR1C1 = "=IF(OR(R1C="""",RC16=R3C),"""",IF(OR(R6C>RC19,R[1]C19<RC19,6371*ACOS(SIN(RC15)*SIN(R2C)+COS(RC15)*COS(R2C)*COS(R3C-RC16))>0.5),"""",IF(R4C-RC17<=R9C4,R6C+RC19,(R6C+RC19)-((R4C-RC17-R9C4)*0.1))))"
     'Copy Ref N
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "N").End(xlUp).Row
.Range("U10:SZ10").AutoFill Destination:=.Range("U10:SZ" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    
    Range("N1:N5,Q1:Q5").FormulaR1C1 = "=MAX(R10C21:R10009C520)"
    Range("O1:O5,R1:R5").FormulaR1C1 = "=MAX(R7C21:R7C520)"
    Range("P1:P5,S1:S5").FormulaR1C1 = "=MIN(R10C20:R10009C20)"
    
    Range("U7:SZ7").FormulaR1C1 = "=IF(AND(R1C14<>0,MAX(R[3]C:R[10002]C)=R1C14),R[-6]C,"""")"
    Range("T10:T10009").FormulaR1C1 = "=IF(AND(R1C14<>0,MAX(RC[1]:RC[500])=R1C14),RC[-6],"""")"
    ActiveSheet.Calculate
    Range("N1:P1").Value = Range("N1:P1").Value
    
    Range("N1:N5,Q1:Q5").FormulaR1C1 = "=MAX(R10C21:R10009C520)"
    Range("O1:O5,R1:R5").FormulaR1C1 = "=MAX(R7C21:R7C520)"
    Range("P1:P5,S1:S5").FormulaR1C1 = "=MIN(R10C20:R10009C20)"
    
    Range("U7:SZ7").FormulaR1C1 = "=IF(AND(R1C14<>0,MAX(R[3]C:R[10002]C)=R1C14),R[-6]C,"""")"
    Range("T10:T10009").FormulaR1C1 = "=IF(AND(R1C14<>0,MAX(RC[1]:RC[500])=R1C14),RC[-6],"""")"
    ActiveSheet.Calculate
    Range("N1:P1").Value = Range("N1:P1").Value
    
    If Range("G510") <> "" Then
        Range("G510:L1009").Copy
        Range("U1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
        Range("U7:SZ7").FormulaR1C1 = "=IF(AND(R2C14<>0,MAX(R[3]C:R[10002]C)=R2C14),R[-6]C,"""")"
        Range("T10:T10009").FormulaR1C1 = "=IF(AND(R2C14<>0,MAX(RC[1]:RC[500])=R2C14),RC[-6],"""")"
        ActiveSheet.Calculate
        Range("N2:P2").Value = Range("N2:P2").Value
    Else: Range("N2:P2").Clear
    End If
    
    If Range("G1010") <> "" Then
        Range("G1010:L1509").Copy
        Range("U1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
        Range("U7:SZ7").FormulaR1C1 = "=IF(AND(R3C14<>0,MAX(R[3]C:R[10002]C)=R3C14),R[-6]C,"""")"
        Range("T10:T10009").FormulaR1C1 = "=IF(AND(R3C14<>0,MAX(RC[1]:RC[500])=R3C14),RC[-6],"""")"
        ActiveSheet.Calculate
        Range("N3:P3").Value = Range("N3:P3").Value
    Else: Range("N3:P3").Clear
    End If
    
    If Range("G1510") <> "" Then
        Range("G1510:L2009").Copy
        Range("U1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
        Range("U7:SZ7").FormulaR1C1 = "=IF(AND(R4C14<>0,MAX(R[3]C:R[10002]C)=R4C14),R[-6]C,"""")"
        Range("T10:T10009").FormulaR1C1 = "=IF(AND(R4C14<>0,MAX(RC[1]:RC[500])=R4C14),RC[-6],"""")"
        ActiveSheet.Calculate
        Range("N4:P4").Value = Range("N4:P4").Value
    Else: Range("N4:P4").Clear
    End If
    
    If Range("G2010") <> "" Then
        Range("G2010:L2509").Copy
        Range("U1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
        Range("U7:SZ7").FormulaR1C1 = "=IF(AND(R5C14<>0,MAX(R[3]C:R[10002]C)=R5C14),R[-6]C,"""")"
        Range("T10:T10009").FormulaR1C1 = "=IF(AND(R5C14<>0,MAX(RC[1]:RC[500])=R5C14),RC[-6],"""")"
        ActiveSheet.Calculate
        Range("N5:P5").Value = Range("N5:P5").Value
    Else: Range("N5:P5").Clear
    End If
    
    If Range("G2510") <> "" Then
        Range("G2510:L3009").Copy
        Range("U1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
        Range("U7:SZ7").FormulaR1C1 = "=IF(AND(R1C17<>0,MAX(R[3]C:R[10002]C)=R1C17),R[-6]C,"""")"
        Range("T10:T10009").FormulaR1C1 = "=IF(AND(R1C17<>0,MAX(RC[1]:RC[500])=R1C17),RC[-6],"""")"
        ActiveSheet.Calculate
        Range("Q1:S1").Value = Range("Q1:S1").Value
    Else: Range("Q1:S1").Clear
    End If
    
    If Range("G3010") <> "" Then
        Range("G3010:L3509").Copy
        Range("U1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
        Range("U7:SZ7").FormulaR1C1 = "=IF(AND(R2C17<>0,MAX(R[3]C:R[10002]C)=R2C17),R[-6]C,"""")"
        Range("T10:T0009").FormulaR1C1 = "=IF(AND(R2C17<>0,MAX(RC[1]:RC[500])=R2C17),RC[-6],"""")"
        ActiveSheet.Calculate
        Range("Q2:S2").Value = Range("Q2:S2").Value
    Else: Range("Q2:S2").Clear
    End If
    
    If Range("G3510") <> "" Then
        Range("G3510:L4009").Copy
        Range("U1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
        Range("U7:SZ7").FormulaR1C1 = "=IF(AND(R3C17<>0,MAX(R[3]C:R[10002]C)=R3C17),R[-6]C,"""")"
        Range("T10:T10009").FormulaR1C1 = "=IF(AND(R3C17<>0,MAX(RC[1]:RC[500])=R3C17),RC[-6],"""")"
        ActiveSheet.Calculate
        Range("Q3:S3").Value = Range("Q3:S3").Value
    Else: Range("Q3:S3").Clear
    End If
    
    If Range("G4010") <> "" Then
        Range("G4010:L4509").Copy
        Range("U1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
        Range("U7:SZ7").FormulaR1C1 = "=IF(AND(R4C17<>0,MAX(R[3]C:R[10002]C)=R4C17),R[-6]C,"""")"
        Range("T10:T10009").FormulaR1C1 = "=IF(AND(R4C17<>0,MAX(RC[1]:RC[500])=R4C17),RC[-6],"""")"
        ActiveSheet.Calculate
        Range("Q4:S4").Value = Range("Q4:S4").Value
    Else: Range("Q4:S4").Clear
    End If
    
    If Range("G4510") <> "" Then
        Range("G4510:L5009").Copy
        Range("U1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
        Range("U7:SZ7").FormulaR1C1 = "=IF(AND(R5C17<>0,MAX(R[3]C:R[10002]C)=R5C17),R[-6]C,"""")"
        Range("T10:T10009").FormulaR1C1 = "=IF(AND(R5C17<>0,MAX(RC[1]:RC[500])=R5C17),RC[-6],"""")"
        ActiveSheet.Calculate
        Range("Q5:S5").Value = Range("Q5:S5").Value
    Else: Range("Q5:S5").Clear
    End If
        
    'Ck thru 7509 Starts(!)
    Range("T6").FormulaR1C1 = "=SUM(R1C[-6]:R5C[-1])"
    ActiveSheet.Calculate
    Range("T6").Value = Range("T6").Value
    
    If Range("T6") = 0 Then
        Range("N1:N5,Q1:Q5").FormulaR1C1 = "=MAX(R10C21:R10009C520)"
        Range("O1:O5,R1:R5").FormulaR1C1 = "=MAX(R7C21:R7C520)"
        Range("P1:P5,S1:S5").FormulaR1C1 = "=MIN(R10C20:R10009C20)"
    
        If Range("G5010") <> "" Then
            Range("G5010:L5509").Copy
            Range("U1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
            Range("U7:SZ7").FormulaR1C1 = "=IF(AND(R1C14<>0,MAX(R[3]C:R[10002]C)=R1C14),R[-6]C,"""")"
            Range("T10:T10009").FormulaR1C1 = "=IF(AND(R1C14<>0,MAX(RC[1]:RC[500])=R1C14),RC[-6],"""")"
            ActiveSheet.Calculate
            Range("N1:P1").Value = Range("N1:P1").Value
        Else: Range("N1:P1").Clear
        End If
        
        If Range("G5510") <> "" Then
            Range("G5510:L6009").Copy
            Range("U1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
            Range("U7:SZ7").FormulaR1C1 = "=IF(AND(R2C14<>0,MAX(R[3]C:R[10002]C)=R2C14),R[-6]C,"""")"
            Range("T10:T10009").FormulaR1C1 = "=IF(AND(R2C14<>0,MAX(RC[1]:RC[500])=R2C14),RC[-6],"""")"
            ActiveSheet.Calculate
            Range("N2:P2").Value = Range("N2:P2").Value
        Else: Range("N2:P2").Clear
        End If
        
        If Range("G6010") <> "" Then
            Range("G6010:L6509").Copy
            Range("U1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
            Range("U7:SZ7").FormulaR1C1 = "=IF(AND(R3C14<>0,MAX(R[3]C:R[10002]C)=R3C14),R[-6]C,"""")"
            Range("T10:T10009").FormulaR1C1 = "=IF(AND(R3C14<>0,MAX(RC[1]:RC[500])=R3C14),RC[-6],"""")"
            ActiveSheet.Calculate
            Range("N3:P3").Value = Range("N3:P3").Value
        Else: Range("N3:P3").Clear
        End If
        
        If Range("G6510") <> "" Then
            Range("G6510:L7009").Copy
            Range("U1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
            Range("U7:SZ7").FormulaR1C1 = "=IF(AND(R4C14<>0,MAX(R[3]C:R[10002]C)=R4C14),R[-6]C,"""")"
            Range("T10:T10009").FormulaR1C1 = "=IF(AND(R4C14<>0,MAX(RC[1]:RC[500])=R4C14),RC[-6],"""")"
            ActiveSheet.Calculate
            Range("N4:P4").Value = Range("N4:P4").Value
        Else: Range("N4:P4").Clear
        End If
        
        If Range("G7010") <> "" Then
            Range("G7010:L7509").Copy
            Range("U1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
            Range("U7:SZ7").FormulaR1C1 = "=IF(AND(R5C14<>0,MAX(R[3]C:R[10002]C)=R5C14),R[-6]C,"""")"
            Range("T10:T10009").FormulaR1C1 = "=IF(AND(R5C14<>0,MAX(RC[1]:RC[500])=R5C14),RC[-6],"""")"
            ActiveSheet.Calculate
            Range("N5:P5").Value = Range("N5:P5").Value
        Else: Range("N5:P5").Clear
        End If
    End If
    
    Range("U1:SZ5009").Clear
    Range("G10:T5009").Clear
    
    Range("T1:T5").FormulaR1C1 = "=IF(MAX(R1C14:R5C14,R1C17:R5C17)=RC[-6],RC[-5],IF(MAX(R1C14:R5C14,R1C17:R5C17)=RC[-3],RC[-2],""""))"
    Range("U1:U5").FormulaR1C1 = "=IF(RC[-1]=RC[-6],RC[-5],IF(RC[-1]=RC[-4],RC[-2],""""))"
    
    Range("H5").FormulaR1C1 = "=MAX(R[-4]C[12]:RC[12])"
    Range("H5").Value = Range("H5").Value
    Range("H7").FormulaR1C1 = "=MAX(R[-6]C[13]:R[-2]C[13])"
    Range("H7").Value = Range("H7").Value
    
    Range("H6:L6").Value = Range("A2:E2").Value
    
    Range("N1:U5").Clear
    
    Range("G8:J8").FormulaR1C1 = "=MAX(R[2]C:R[10001]C)"
    Range("G10").FormulaR1C1 = "=IF(RC[-6]=R5C8,RC[-5],"""")"
    Range("H10:J10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-5],"""")"
    'Copy Ref A
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:J10").AutoFill Destination:=.Range("G10:J" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("I5:L5").Value = Range("G8:J8").Value
    
    Range("G10:G10009").FormulaR1C1 = "=IF(RC[-6]=R7C8,RC[-5],"""")"
    ActiveSheet.Calculate
    Range("I7:L7").Value = Range("G8:J8").Value
    
    Range("G8:J10009").Clear
    
    Range("M6").FormulaR1C1 = _
        "=IF(AND(R[-1]C[-5]<>0,R[1]C[-5]<>0),2*6371*ACOS(SIN(R[-1]C[-4])*SIN(RC[-4])+COS(R[-1]C[-4])*COS(RC[-4])*COS(RC[-3]-R[-1]C[-3])),0)"
    ActiveSheet.Calculate
    Range("M6").Value = Range("M6").Value
    
    If Range("M6") > Range("M2") Then
        Range("H1:M3").Value = Range("H5:M7").Value
    End If
    
    Range("T6,H5:M8,S7:S9").Clear
    'Application.Run "F.xlsm!AORc"
    
  End If
    Application.Run "F.xlsm!ORredeemer"

End Sub
Sub ORredeemer()
'
' Android - OR TP NOT at St Start, Fini or via AORc
' JLR 5/15/18 Revised Matrix""; 5/16/2018 Revised Matrix; 8/6/2018 Returned Matrix formula to prior
'
Application.ScreenUpdating = False
    'Ck TP
    Range("A7").Value = Range("H2").Value
    Range("G8:L8").FormulaR1C1 = "=MAX(R[2]C:R[10001]C)"
    Range("G10").FormulaR1C1 = _
        "=IF(OR(RC[-6]<=R1C8,RC[-6]>=R3C8),"""",IF(R1C11-R3C11<=R9C4,2*6371*ACOS(SIN(RC[-5])*SIN(R1C9)+COS(RC[-5])*COS(R1C9)*COS(R1C10-RC[-4])),2*6371*ACOS(SIN(RC[-5])*SIN(R1C9)+COS(RC[-5])*COS(R1C9)*COS(R1C10-RC[-4]))-((R1C11-R3C11-R9C4)*.1)))"
    Range("H10").FormulaR1C1 = "=IF(RC[-1]=R8C[-1],RC[-7],"""")"
    Range("I10:L10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-7],"""")"

    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:L10").AutoFill Destination:=.Range("G10:L" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("L10:P10009").Value = Range("L10:P10009").Value
    Range("G8:L8").Value = Range("G8:L8").Value
    Range("G10:L10009").Clear
    
  If Range("G8") > Range("M2") Then
    Range("H2:L2").Value = Range("H8:L8").Value
    Range("M2").FormulaR1C1 = _
        "=IF(OR(SUM(R1C8:R1C12)=0,SUM(R2C8:R2C12)=0),0,IF(R[-1]C[-2]-R[1]C[-2]<=R9C4,2*6371*ACOS(SIN(R[-1]C[-4])*SIN(RC[-4])+COS(R[-1]C[-4])*COS(RC[-4])*COS(RC[-3]-R[-1]C[-3])),2*6371*ACOS(SIN(R[-1]C[-4])*SIN(RC[-4])+COS(R[-1]C[-4])*COS(RC[-4])*COS(RC[-3]-R[-1]C[-3]))-((R[-1]C[-2]-R[1]C[-2]-R9C4)*0.1)))"
    Range("M2").Value = Range("M2").Value
  End If
    Range("G8:L8").Clear

    Range("A6").Value = 3.47222222222222E-03
    'Starts
    Range("G10").FormulaR1C1 = "=IF(RC[-4]=R2C10,"""",IF(AND(RC[-6]>=R1C8-R6C1,RC[-6]<=R1C8+R6C1),RC[-6],""""))"
    Range("H10:K10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-6],"""")"
    Range("L10").FormulaR1C1 = _
        "=IF(RC[-1]<>"""",6371*ACOS(SIN(RC[-10])*SIN(R2C9)+COS(RC[-10])*COS(R2C9)*COS(R2C10-RC[-9])),"""")"
    'Copy Ref A Value Sort
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:L10").AutoFill Destination:=.Range("G10:L" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("G10:L10009").Value = Range("G10:L10009").Value
    Range("G10:L10009").Sort Key1:=Range("G10"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
   
    Range("L6").FormulaR1C1 = "=MIN(R[4]C:R[10003]C)"
    Range("L6").Value = Range("L6").Value
    
    'FINIS TO TOP
    Range("N10").FormulaR1C1 = _
        "=IF(RC[-11]=R2C10,"""",IF(AND(RC[-13]>=R3C8-R6C1,RC[-13]<=R3C8+R6C1,6371*ACOS(SIN(RC[-12])*SIN(R2C9)+COS(RC[-12])*COS(R2C9)*COS(R2C10-RC[-11]))>R6C12),RC[-13],""""))"
    Range("O10:R10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-13],"""")"
    Range("S10").FormulaR1C1 = _
        "=IF(RC[-1]<>"""",6371*ACOS(SIN(RC[-17])*SIN(R2C9)+COS(RC[-17])*COS(R2C9)*COS(R2C10-RC[-16])),"""")"
    'Copy Ref A Value Sort
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("N10:S10").AutoFill Destination:=.Range("N10:S" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("N10:S10009").Value = Range("N10:S10009").Value
    Range("N10:S10009").Sort Key1:=Range("N10"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    Range("N10:S509").Copy
    Range("N1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    Application.CutCopyMode = False
    
    Range("N10:S10009").Clear
    
    'MATRIX Amended 7/29/17 to restore .5km radius OZSector/Line length
    'Range("N10:SS10").FormulaR1C1 = _
        "=IF(OR(R1C="""",RC9=R3C),"""",IF(OR(R6C<RC12,6371*ACOS(SIN(RC8)*SIN(R2C)+COS(RC8)*COS(R2C)*COS(R3C-RC9))>0.5,R6C<SQRT(RC12^2+0.25)),"""",IF(RC10-R4C<=R9C4,2*RC12,"""")))"
    ''Range("N10:SS10").FormulaR1C1 = "=IF(OR(R1C="""",RC9=R3C),"""",IF(OR(R6C<RC12,R6C[1]<R6C,6371*ACOS(SIN(RC8)*SIN(R2C)+COS(RC8)*COS(R2C)*COS(R3C-RC9))>0.5,R6C<SQRT(RC12^2+0.25)),"""",IF(RC10-R4C<=R9C4,2*RC12,"""")))"
    'Range("N10:SS10").FormulaR1C1 = "=IF(OR(R1C="""",RC9=R3C),"""",IF(OR(R6C<RC12,AND(R6C[1]<R6C,R6C[-1]<R6C),6371*ACOS(SIN(RC8)*SIN(R2C)+COS(RC8)*COS(R2C)*COS(R3C-RC9))>0.5,R6C<SQRT(RC12^2+0.25)),"""",IF(RC10-R4C<=R9C4,2*RC12,"""")))"
    ''Range("N10:SS10").FormulaR1C1 = "=IF(OR(R1C="""",RC9=R3C),"""",IF(OR(R6C<RC12,AND(R6C[1]<R6C,R6C[-1]<R6C),6371*ACOS(SIN(RC8)*SIN(R2C)+COS(RC8)*COS(R2C)*COS(R3C-RC9))>0.5,R6C<SQRT(RC12^2+0.25)),"""",IF(RC10-R4C<=R9C4,RC12+R6C,"""")))"
    '''Range("N10:SS10").FormulaR1C1 = "=IF(OR(R1C="""",RC9=R3C),"""",IF(OR(R6C<RC12,R6C[1]<R6C,6371*ACOS(SIN(RC8)*SIN(R2C)+COS(RC8)*COS(R2C)*COS(R3C-RC9))>0.5,R6C<SQRT(RC12^2+0.25)),"""",IF(RC10-R4C<=R9C4,RC12+R6C,"""")))"
    Range("N10:SS10").FormulaR1C1 = "=IF(OR(R1C="""",RC9=R3C),"""",IF(OR(R6C<RC12,AND(R6C[1]<R6C,R6C[-1]<R6C),6371*ACOS(SIN(RC8)*SIN(R2C)+COS(RC8)*COS(R2C)*COS(R3C-RC9))>0.5,R6C<SQRT(RC12^2+0.25)),"""",IF(RC10-R4C<=R9C4,2*RC12,"""")))"
    'Copy Ref G
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
.Range("N10:SS10").AutoFill Destination:=.Range("N10:SS" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    
    Range("I5").FormulaR1C1 = "=MAX(R[5]C[5]:R[3004]C[504])"
    Range("I5").Value = Range("I5").Value
    
    Range("N10:SS3009").Value = Range("N10:SS3009").Value
    Range("N7:SS7").FormulaR1C1 = "=IF(AND(R[-6]C<>"""",MAX(R[3]C:R[3002]C)=R5C9,RC[-1]=""""),R[-6]C,"""")"
    Range("N7:SS7").Value = Range("N7:SS7").Value
    
    Range("M10:M3009").FormulaR1C1 = "=IF(MAX(RC[1]:RC[500])=R5C9,RC[-6],"""")"
    ActiveSheet.Calculate
    Range("M10:M3009").Value = Range("M10:M3009").Value
    
    Range("H7").FormulaR1C1 = "=MIN(RC[6]:RC[505])"
    Range("N8:SS11").FormulaR1C1 = "=IF(R7C=R7C8,R[-6]C,"""")"
    Range("I7").FormulaR1C1 = "=MAX(R[1]C[5]:R[1]C[504])"
    Range("J7").FormulaR1C1 = "=MAX(R[2]C[4]:R[2]C[503])"
    Range("K7").FormulaR1C1 = "=MAX(R[3]C[3]:R[3]C[502])"
    Range("L7").FormulaR1C1 = "=MAX(R[4]C[2]:R[4]C[501])"
    Range("H7:L7").Value = Range("H7:L7").Value
    Range("H6:L6").Value = Range("H2:L2").Value
   
    Range("N10:Q10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-6],"""")"
    'Copy Ref G
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
.Range("N10:Q10").AutoFill Destination:=.Range("N10:Q" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    
    Range("H5:L5").FormulaR1C1 = "=MAX(R[5]C[5]:R[3004]C[5])"
    Range("H5:L5").Value = Range("H5:L5").Value
    Range("M6").FormulaR1C1 = "=2*6371*ACOS(SIN(R[-1]C[-4])*SIN(RC[-4])+COS(R[-1]C[-4])*COS(RC[-4])*COS(RC[-3]-R[-1]C[-3]))"
    Range("M6").Value = Range("M6").Value
    
    If Range("M6") > Range("M2") Then
        Range("H1:M3").Value = Range("H5:M7").Value
    End If
    
    Range("A6,A7,H5:M8,G10:M10009").Clear
    Columns("N:SS").Clear
    Application.Run "F.xlsm!FLine1"
End Sub

Sub Fline1()
'
'Setup to ck FINISH LINE @ AORb Start
'
Application.ScreenUpdating = False
    
'Ck TP
    Range("G8:L8").FormulaR1C1 = "=MAX(R[2]C:R[10001]C)"
    Range("G10").FormulaR1C1 = _
        "=IF(OR(RC[-6]<=R1C8,RC[-6]>=R3C8),"""",IF(R1C11-R3C11<=R9C4,2*6371*ACOS(SIN(RC[-5])*SIN(R1C9)+COS(RC[-5])*COS(R1C9)*COS(R1C10-RC[-4])),2*6371*ACOS(SIN(RC[-5])*SIN(R1C9)+COS(RC[-5])*COS(R1C9)*COS(R1C10-RC[-4]))-((R1C11-R3C11-R9C4)*.1)))"
    Range("H10").FormulaR1C1 = "=IF(RC[-1]=R8C[-1],RC[-7],"""")"
    Range("I10:L10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-7],"""")"

    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:L10").AutoFill Destination:=.Range("G10:L" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("L10:P10009").Value = Range("L10:P10009").Value
    Range("G8:L8").Value = Range("G8:L8").Value
    Range("G10:L10009").Clear
    
  If Range("G8") > Range("M2") Then
    Range("H2:L2").Value = Range("H8:L8").Value
    Range("M2").FormulaR1C1 = _
        "=IF(OR(SUM(R1C8:R1C12)=0,SUM(R2C8:R2C12)=0),0,IF(R[-1]C[-2]-R[1]C[-2]<=R9C4,2*6371*ACOS(SIN(R[-1]C[-4])*SIN(RC[-4])+COS(R[-1]C[-4])*COS(RC[-4])*COS(RC[-3]-R[-1]C[-3])),2*6371*ACOS(SIN(R[-1]C[-4])*SIN(RC[-4])+COS(R[-1]C[-4])*COS(RC[-4])*COS(RC[-3]-R[-1]C[-3]))-((R[-1]C[-2]-R[1]C[-2]-R9C4)*0.1)))"
    Range("M2").Value = Range("M2").Value
  End If
    Range("G8:L8").Clear
 'Resume as before
    Range("A1:F2").Clear
    Range("A1:F3").Value = Range("H1:M3").Value
    
'Select AORb FINI + 2 fixes before, 2 fixes after
    Range("G10").FormulaR1C1 = "=IF(RC[-6]=R3C1,1,"""")"
    Range("H10").FormulaR1C1 = "=IF(R[2]C[-1]=1,1,"""")"
    Range("I10").FormulaR1C1 = "=IF(R[1]C[-2]=1,1,"""")"
    Range("J10").FormulaR1C1 = "=IF(R[-1]C[-3]=1,1,"""")"
    Range("K10").FormulaR1C1 = "=IF(R[-2]C[-4]=1,1,"""")"
    Range("L10").FormulaR1C1 = _
        "=IF(OR(RC[-5]=1,RC[-4]=1,RC[-3]=1,RC[-2]=1,RC[-1]=1),RC[-11],"""")"
    Range("M10:P10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-11],"""")"
    
    'Copy Ref A
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:P10").AutoFill Destination:=.Range("G10:P" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("L10:P10009").Value = Range("L10:P10009").Value
    Range("G10:K10009").Clear
    Range("L10:P10009").Sort Key1:=Range("L10"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal

    Sheets("YDWK3").Activate
    Sheets("YDWK3").Range("M22:Q22").Value = Sheets("Sheet2").Range("H2:L2").Value
    Sheets("YDWK3").Range("M23:Q23").Value = Sheets("Sheet2").Range("H1:L1").Value
    Sheets("YDWK3").Range("M25:Q29").Value = Sheets("Sheet2").Range("L10:P14").Value

    Application.Run "F.xlsm!FINlines"

    Sheets("Sheet2").Activate
    Sheets("Sheet2").Range("A4").Value = Sheets("YDWK3").Range("M17").Value
    Sheets("Sheet2").Range("B4").FormulaR1C1 = "=IF(RC[-1]<>""NO FIN LINE"",""        FINISH LINE"","""")"
    Range("B4").Value = Range("B4").Value
    Sheets("Sheet2").Range("D4:E4").Value = Sheets("YDWK3").Range("N17:O17").Value
    Range("L10:P14").Clear

Application.Run "F.xlsm!YDWK3clear"

If Range("B4") = "        FINISH LINE" Then
    Application.Run ("F.xlsm!OZSectorB")
ElseIf Range("B4") <> "        FINISH LINE" Then
    Application.Run "F.xlsm!OZSector"
End If

End Sub
Sub FINlines()
'
Application.ScreenUpdating = False
' DGfinfix1to2
    Range("N23:O23").Copy
    Range("E39").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    
    Range("N25:O25").Copy
    Range("E41").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    
    Range("E81").FormulaR1C1 = "=R47C[-2]+(1-R79C)*R46C[-2]*R67C*(ACOS(R62C)+R79C*SIN(ACOS(R62C))*(R69C+R79C*R62C*(-1+2*R69C*R69C)))"
    ActiveSheet.Calculate
    Range("M35").Value = Range("F37").Value

    Range("N26:O26").Copy
    Range("E41").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    
    Range("E81").FormulaR1C1 = "=R47C[-2]+(1-R79C)*R46C[-2]*R67C*(ACOS(R62C)+R79C*SIN(ACOS(R62C))*(R69C+R79C*R62C*(-1+2*R69C*R69C)))"
    ActiveSheet.Calculate
    Range("N35").Value = Range("F37").Value
    
    Range("N25:O25").Copy
    Range("E39").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    
    Range("E81").FormulaR1C1 = "=R47C[-2]+(1-R79C)*R46C[-2]*R67C*(ACOS(R62C)+R79C*SIN(ACOS(R62C))*(R69C+R79C*R62C*(-1+2*R69C*R69C)))"
    ActiveSheet.Calculate
    Range("O35").Value = Range("F37").Value
    
    Range("N22:O22").Copy
    Range("E39").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    
    Range("N23:O23").Copy
    Range("E41").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    
    Range("E81").FormulaR1C1 = "=R47C[-2]+(1-R79C)*R46C[-2]*R67C*(ACOS(R62C)+R79C*SIN(ACOS(R62C))*(R69C+R79C*R62C*(-1+2*R69C*R69C)))"
    ActiveSheet.Calculate
    Range("Q35").Value = Range("C36").Value
    
    Range("N23:O23").Copy
    Range("E39").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
   
    Range("N25:O25").Copy
    Range("E41").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    
    Range("E81").FormulaR1C1 = "=R47C[-2]+(1-R79C)*R46C[-2]*R67C*(ACOS(R62C)+R79C*SIN(ACOS(R62C))*(R69C+R79C*R62C*(-1+2*R69C*R69C)))"
    ActiveSheet.Calculate
    Range("P44").Value = Range("C36").Value
    
    Range("N26:O26").Copy
    Range("E41").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    
    Range("E81").FormulaR1C1 = "=R47C[-2]+(1-R79C)*R46C[-2]*R67C*(ACOS(R62C)+R79C*SIN(ACOS(R62C))*(R69C+R79C*R62C*(-1+2*R69C*R69C)))"
    ActiveSheet.Calculate
    Range("P46").Value = Range("C36").Value

' DGfinfix2to3
    Range("N23:O23").Copy
    Range("E39").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    
    Range("N26:O26").Copy
    Range("E41").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    
    Range("E81").FormulaR1C1 = "=R47C[-2]+(1-R79C)*R46C[-2]*R67C*(ACOS(R62C)+R79C*SIN(ACOS(R62C))*(R69C+R79C*R62C*(-1+2*R69C*R69C)))"
    ActiveSheet.Calculate
    Range("M50").Value = Range("F37").Value

    Range("N27:O27").Copy
    Range("E41").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    
    Range("E81").FormulaR1C1 = "=R47C[-2]+(1-R79C)*R46C[-2]*R67C*(ACOS(R62C)+R79C*SIN(ACOS(R62C))*(R69C+R79C*R62C*(-1+2*R69C*R69C)))"
    ActiveSheet.Calculate
    Range("N50").Value = Range("F37").Value
    
    Range("N26:O26").Copy
    Range("E39").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    
    Range("E81").FormulaR1C1 = "=R47C[-2]+(1-R79C)*R46C[-2]*R67C*(ACOS(R62C)+R79C*SIN(ACOS(R62C))*(R69C+R79C*R62C*(-1+2*R69C*R69C)))"
    ActiveSheet.Calculate
    Range("O50").Value = Range("F37").Value
    
    Range("N22:O22").Copy
    Range("E39").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    
    Range("N23:O23").Copy
    Range("E41").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    
    Range("E81").FormulaR1C1 = "=R47C[-2]+(1-R79C)*R46C[-2]*R67C*(ACOS(R62C)+R79C*SIN(ACOS(R62C))*(R69C+R79C*R62C*(-1+2*R69C*R69C)))"
    ActiveSheet.Calculate
    Range("Q50").Value = Range("C36").Value
    
    Range("N23:O23").Copy
    Range("E39").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
   
    Range("N26:O26").Copy
    Range("E41").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    
    Range("E81").FormulaR1C1 = "=R47C[-2]+(1-R79C)*R46C[-2]*R67C*(ACOS(R62C)+R79C*SIN(ACOS(R62C))*(R69C+R79C*R62C*(-1+2*R69C*R69C)))"
    ActiveSheet.Calculate
    Range("P59").Value = Range("C36").Value
    
    Range("N27:O27").Copy
    Range("E41").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    
    Range("E81").FormulaR1C1 = "=R47C[-2]+(1-R79C)*R46C[-2]*R67C*(ACOS(R62C)+R79C*SIN(ACOS(R62C))*(R69C+R79C*R62C*(-1+2*R69C*R69C)))"
    ActiveSheet.Calculate
    Range("P61").Value = Range("C36").Value
    
' DGfinfix3to4
    Range("N23:O23").Copy
    Range("E39").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    
    Range("N27:O27").Copy
    Range("E41").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    
    Range("E81").FormulaR1C1 = "=R47C[-2]+(1-R79C)*R46C[-2]*R67C*(ACOS(R62C)+R79C*SIN(ACOS(R62C))*(R69C+R79C*R62C*(-1+2*R69C*R69C)))"
    ActiveSheet.Calculate
    Range("M65").Value = Range("F37").Value

    Range("N28:O28").Copy
    Range("E41").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    
    Range("E81").FormulaR1C1 = "=R47C[-2]+(1-R79C)*R46C[-2]*R67C*(ACOS(R62C)+R79C*SIN(ACOS(R62C))*(R69C+R79C*R62C*(-1+2*R69C*R69C)))"
    ActiveSheet.Calculate
    Range("N65").Value = Range("F37").Value
    
    Range("N27:O27").Copy
    Range("E39").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    
    Range("E81").FormulaR1C1 = "=R47C[-2]+(1-R79C)*R46C[-2]*R67C*(ACOS(R62C)+R79C*SIN(ACOS(R62C))*(R69C+R79C*R62C*(-1+2*R69C*R69C)))"
    ActiveSheet.Calculate
    Range("O65").Value = Range("F37").Value
    
    Range("N22:O22").Copy
    Range("E39").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    
    Range("N23:O23").Copy
    Range("E41").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    
    Range("E81").FormulaR1C1 = "=R47C[-2]+(1-R79C)*R46C[-2]*R67C*(ACOS(R62C)+R79C*SIN(ACOS(R62C))*(R69C+R79C*R62C*(-1+2*R69C*R69C)))"
    ActiveSheet.Calculate
    Range("Q65").Value = Range("C36").Value
    
    Range("N23:O23").Copy
    Range("E39").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
   
    Range("N27:O27").Copy
    Range("E41").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    
    Range("E81").FormulaR1C1 = "=R47C[-2]+(1-R79C)*R46C[-2]*R67C*(ACOS(R62C)+R79C*SIN(ACOS(R62C))*(R69C+R79C*R62C*(-1+2*R69C*R69C)))"
    ActiveSheet.Calculate
    Range("P74").Value = Range("C36").Value
    
    Range("N28:O28").Copy
    Range("E41").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    
    Range("E81").FormulaR1C1 = "=R47C[-2]+(1-R79C)*R46C[-2]*R67C*(ACOS(R62C)+R79C*SIN(ACOS(R62C))*(R69C+R79C*R62C*(-1+2*R69C*R69C)))"
    ActiveSheet.Calculate
    Range("P76").Value = Range("C36").Value
    
' DGfinfix4to5
    Range("N23:O23").Copy
    Range("E39").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    
    Range("N28:O28").Copy
    Range("E41").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    
    Range("E81").FormulaR1C1 = "=R47C[-2]+(1-R79C)*R46C[-2]*R67C*(ACOS(R62C)+R79C*SIN(ACOS(R62C))*(R69C+R79C*R62C*(-1+2*R69C*R69C)))"
    ActiveSheet.Calculate
    Range("M80").Value = Range("F37").Value

    Range("N29:O29").Copy
    Range("E41").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    
    Range("E81").FormulaR1C1 = "=R47C[-2]+(1-R79C)*R46C[-2]*R67C*(ACOS(R62C)+R79C*SIN(ACOS(R62C))*(R69C+R79C*R62C*(-1+2*R69C*R69C)))"
    ActiveSheet.Calculate
    Range("N80").Value = Range("F37").Value
    
    Range("N28:O28").Copy
    Range("E39").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    
    Range("E81").FormulaR1C1 = "=R47C[-2]+(1-R79C)*R46C[-2]*R67C*(ACOS(R62C)+R79C*SIN(ACOS(R62C))*(R69C+R79C*R62C*(-1+2*R69C*R69C)))"
    ActiveSheet.Calculate
    Range("O80").Value = Range("F37").Value
    
    Range("N22:O22").Copy
    Range("E39").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    
    Range("N23:O23").Copy
    Range("E41").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    
    Range("E81").FormulaR1C1 = "=R47C[-2]+(1-R79C)*R46C[-2]*R67C*(ACOS(R62C)+R79C*SIN(ACOS(R62C))*(R69C+R79C*R62C*(-1+2*R69C*R69C)))"
    ActiveSheet.Calculate
    Range("Q80").Value = Range("C36").Value
    
    Range("N23:O23").Copy
    Range("E39").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
   
    Range("N28:O28").Copy
    Range("E41").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    
    Range("E81").FormulaR1C1 = "=R47C[-2]+(1-R79C)*R46C[-2]*R67C*(ACOS(R62C)+R79C*SIN(ACOS(R62C))*(R69C+R79C*R62C*(-1+2*R69C*R69C)))"
    ActiveSheet.Calculate
    Range("P89").Value = Range("C36").Value
    
    Range("N29:O29").Copy
    Range("E41").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    
    Range("E81").FormulaR1C1 = "=R47C[-2]+(1-R79C)*R46C[-2]*R67C*(ACOS(R62C)+R79C*SIN(ACOS(R62C))*(R69C+R79C*R62C*(-1+2*R69C*R69C)))"
    ActiveSheet.Calculate
    Range("P91").Value = Range("C36").Value
    
    Range("M17").FormulaR1C1 = _
        "=IF(AND(R[20]C[3]=""GOOD FINISH"",R[21]C[3]=""GOOD FINISH""),R[28]C[1],IF(AND(R[35]C[3]=""GOOD FINISH"",R[36]C[3]=""GOOD FINISH""),R[43]C[1],IF(AND(R[50]C[3]=""GOOD FINISH"",R[51]C[3]=""GOOD FINISH""),R[58]C[1],IF(AND(R[65]C[3]=""GOOD FINISH"",R[66]C[3]=""GOOD FINISH""),R[73]C[1],""NO FIN LINE""))))"
    Range("N17").FormulaR1C1 = _
        "=IF(AND(R[20]C[2]=""GOOD FINISH"",R[21]C[2]=""GOOD FINISH""),R[28]C[1],IF(AND(R[35]C[2]=""GOOD FINISH"",R[36]C[2]=""GOOD FINISH""),R[43]C[1],IF(AND(R[50]C[2]=""GOOD FINISH"",R[51]C[2]=""GOOD FINISH""),R[58]C[1],IF(AND(R[65]C[2]=""GOOD FINISH"",R[66]C[2]=""GOOD FINISH""),R[73]C[1],""""))))"
    Range("O17").FormulaR1C1 = _
        "=IF(RC[-1]="""","""",IF(AND(R[20]C[1]=""GOOD FINISH"",R[21]C[1]=""GOOD FINISH""),R[28]C[1],IF(AND(R[35]C[1]=""GOOD FINISH"",R[36]C[1]=""GOOD FINISH""),R[43]C[1],IF(AND(R[50]C[1]=""GOOD FINISH"",R[51]C[1]=""GOOD FINISH""),R[58]C[1],IF(AND(R[65]C[1]=""GOOD FINISH"",R[66]C[1]=""GOOD FINISH""),R[73]C[1])))))"
    ActiveSheet.Calculate
    Range("M17:O17").Value = Range("M17:O17").Value
    Range("A1").Activate
''Application.Calculation = xlCalculationManual
End Sub
Sub YDWK3clear()
'
' Clears YDWK3
'

Application.ScreenUpdating = False
    Sheets("YDWK3").Activate
    Range("C5:D6").Value = 0
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Range("M17:O17,M22:Q29").Clear
    Range("E39:E42").Value = 0
    Range("M35:O35,Q35").Clear
    Range("P44,P46").Clear
    Range("M50:O50,Q50").Clear
    Range("P59,P61").Clear
    Range("M65:O65,Q65").Clear
    Range("P74,P76").Clear
    Range("M80:O80,Q80").Clear
    Range("P89,P91").Clear
    Range("E81").FormulaR1C1 = _
        "=R47C[-2]+(1-R79C)*R46C[-2]*R67C*(ACOS(R62C)+R79C*SIN(ACOS(R62C))*(R69C+R79C*R62C*(-1+2*R69C*R69C)))"
    ActiveSheet.Calculate
    Range("A1").Activate
    Sheets("Sheet2").Activate
    Range("A1").Activate
Application.Calculation = xlCalculationManual
End Sub

Sub Fline2()
'
' Sets up to ck Fin Line after OZ Sector
'
Application.ScreenUpdating = False
    
If Range("M3") = "Sector" Then
 
    Range("G10").FormulaR1C1 = "=IF(RC[-6]=R3C8,1,"""")"
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
    Sheets("YDWK3").Range("M22:Q22").Value = Sheets("Sheet2").Range("H2:L2").Value
    Sheets("YDWK3").Range("M23:Q23").Value = Sheets("Sheet2").Range("H1:L1").Value
    Sheets("YDWK3").Range("M25:Q29").Value = Sheets("Sheet2").Range("L10:P14").Value

    Application.Run "F.xlsm!FINLINES"

    Sheets("Sheet2").Activate
    Sheets("Sheet2").Range("H4").Value = Sheets("YDWK3").Range("M17").Value
    Sheets("Sheet2").Range("I4").FormulaR1C1 = "=IF(RC[-1]<>""NO FIN LINE"",""        FINISH LINE"","""")"
    Range("I4").Value = Range("I4").Value
    Sheets("Sheet2").Range("K4:L4").Value = Sheets("YDWK3").Range("N17:O17").Value
    Range("L10:P14").Clear
    Range("M4").FormulaR1C1 = _
        "=IF(RC[-5]=""NO FIN LINE"","""",IF(R[-3]C[-2]-RC[-2]<=R9C4,2*6371*ACOS(SIN(R[-3]C[-4])*SIN(R[-2]C[-4])+COS(R[-3]C[-4])*COS(R[-2]C[-4])*COS(R[-2]C[-3]-R[-3]C[-3])),2*6371*ACOS(SIN(R[-3]C[-4])*SIN(R[-2]C[-4])+COS(R[-3]C[-4])*COS(R[-2]C[-4])*COS(R[-2]C[-3]-R[-3]C[-3]))-((R[-3]C[-2]-RC[-2]-R9C4)*0.1)))"
    Range("M4").Value = Range("M4").Value

    Application.Run "F.xlsm!YDWK3clear"

    If Range("H4") = "NO FIN LINE" Then
    
    Range("A6").Value = 4.16666666666667E-02
     Range("H1:L1").Copy
    Range("N1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    
    Range("N7").FormulaR1C1 = "=6371*ACOS(SIN(R[-5]C)*SIN(R2C9)+COS(R[-5]C)*COS(R2C9)*COS(R2C10-R[-4]C))"
    Range("N8").FormulaR1C1 = "=SQRT(R[-1]C^2+0.25)"
    ActiveSheet.Calculate
    Range("N7:N8").Value = Range("N7:N8").Value

    Range("G10").FormulaR1C1 = _
        "=IF(RC[-4]=R3C14,"""",IF(AND(RC[-6]>=R3C8-R6C1/2,RC[-6]<=R3C8+R6C1/2,6371*ACOS(SIN(RC[-5])*SIN(R2C14)+COS(RC[-5])*COS(R2C14)*COS(R3C14-RC[-4]))<=0.5),RC[-6],""""))"
    Range("H10:K10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-6],"""")"
    
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:K10").AutoFill Destination:=.Range("G10:K" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("G10:K10009").Value = Range("G10:K10009").Value
    Range("G10:K10009").Sort Key1:=Range("G10"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    
 If Range("G10") = "" Or Range("G11") = "" And Range("G12") = "" Then
    Range("N1:N8").Clear
 
 ElseIf Range("G10") <> "" Then
    Range("L10").FormulaR1C1 = _
        "=IF(OR(AND(R[-1]C[-4]<>"""",6371*ACOS(SIN(RC[-4])*SIN(R2C9)+COS(RC[-4])*COS(R2C9)*COS(R2C10-RC[-3]))>=R7C14,6371*ACOS(SIN(R[-1]C[-4])*SIN(R2C9)+COS(R[-1]C[-4])*COS(R2C9)*COS(R2C10-R[-1]C[-3]))<=R8C14),AND(R[1]C[-4]<>"""",6371*ACOS(SIN(RC[-4])*SIN(R2C9)+COS(RC[-4])*COS(R2C9)*COS(R2C10-RC[-3]))<=R8C14,6371*ACOS(SIN(R[1]C[-4])*SIN(R2C9)+COS(R[1]C[-4])*COS(R2C9)*COS(R2C10-RC[-3]))>=R7C14)),1,0)"
    Range("M10").FormulaR1C1 = "=IF(RC[-1]=1,RC[-6],"""")"
    
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
.Range("L10:M10").AutoFill Destination:=.Range("L10:M" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("L10:M10009").Value = Range("L10:M10009").Value
    
    Range("F9").FormulaR1C1 = "=MIN(R[1]C13:R[1000]C13)"
    ActiveSheet.Calculate
    Range("F9").Value = Range("F9").Value
    Range("G10:M10009").Clear
    
  If Range("F9") > 0 Then
  
  'Select new FINI + 2 fixes before, 2 fixes after
    Range("G10").FormulaR1C1 = "=IF(RC[-6]=R9C6,1,"""")"
    Range("H10").FormulaR1C1 = "=IF(R[2]C[-1]=1,1,"""")"
    Range("I10").FormulaR1C1 = "=IF(R[1]C[-2]=1,1,"""")"
    Range("J10").FormulaR1C1 = "=IF(R[-1]C[-3]=1,1,"""")"
    Range("K10").FormulaR1C1 = "=IF(R[-2]C[-4]=1,1,"""")"
    Range("L10").FormulaR1C1 = _
        "=IF(OR(RC[-5]=1,RC[-4]=1,RC[-3]=1,RC[-2]=1,RC[-1]=1),RC[-11],"""")"
    Range("M10:P10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-11],"""")"
    'Sibylle fails with this IF
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:P10").AutoFill Destination:=.Range("G10:P" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("L10:P10009").Value = Range("L10:P10009").Value
    Range("F6:F9,G10:K10009,N1:N8").Clear
    
    Range("L10:P10009").Sort Key1:=Range("L10"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal

    Sheets("YDWK3").Activate
    Sheets("YDWK3").Range("M22:Q22").Value = Sheets("Sheet2").Range("H2:L2").Value
    Sheets("YDWK3").Range("M23:Q23").Value = Sheets("Sheet2").Range("H1:L1").Value
    Sheets("YDWK3").Range("M25:Q29").Value = Sheets("Sheet2").Range("L10:P14").Value

Application.Run "F.xlsm!FINlines"
    
    Sheets("Sheet2").Activate
    Sheets("Sheet2").Range("H4").Value = Sheets("YDWK3").Range("M17").Value
    Sheets("Sheet2").Range("I4").FormulaR1C1 = "=IF(RC[-1]<>""NO FIN LINE"",""        FINISH LINE"","""")"
    Range("I4").Value = Range("I4").Value
    Sheets("Sheet2").Range("K4:L4").Value = Sheets("YDWK3").Range("N17:O17").Value
    Range("L10:P14").Clear

    Application.Run "F.xlsm!YDWK3clear"
  
ElseIf Range("G10") = "" Or Range("F9") = 0 Then
    
    Range("N1:N8,F9:M10009").Clear
    Range("G10").FormulaR1C1 = _
        "=IF(AND(RC[-6]>=R1C8-R6C1,RC[-6]<=R1C8+R6C1,6371*ACOS(SIN(RC[-5])*SIN(R2C9)+COS(RC[-5])*COS(R2C9)*COS(R2C10-RC[-4]))<R2C13/2),6371*ACOS(SIN(RC[-5])*SIN(R2C9)+COS(RC[-5])*COS(R2C9)*COS(R2C10-RC[-4])),"""")"
    
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:G10").AutoFill Destination:=.Range("G10:G" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("F10:F10009").FormulaR1C1 = "=IF(RC[1]<>"""",1,"""")"
    ActiveSheet.Calculate
    Range("F8").FormulaR1C1 = "=SUM(R[2]C:R[10001]C)"
    Range("F8").Value = Range("F8").Value
    Range("F10:F10009").Clear
    
    Range("G8").FormulaR1C1 = "=IF(R8C[-1]>500,MEDIAN(R[2]C:R[10001]C),MIN(R[2]C:R[10001]C))"
    ActiveSheet.Calculate
    Range("G8").Value = Range("G8").Value
    Range("H10").FormulaR1C1 = _
        "=IF(AND(RC[-1]<>"""",RC[-1]>=R8C7,RC[-1]<=R2C13/2),RC[-7],"""")"
    Range("I10:L10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-7],"""")"
    
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:L10").AutoFill Destination:=.Range("G10:L" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("G10:L10009").Value = Range("G10:L10009").Value
    Range("H10:L10009").Sort Key1:=Range("H10"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    Range("H10:L509").Copy
    Range("N1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=True
    Application.CutCopyMode = False
    
    Range("F8:L10009").Clear
    Range("N6:SS6").FormulaR1C1 = _
        "=IF(R[-5]C<>"""",6371*ACOS(SIN(R[-4]C)*SIN(R2C9)+COS(R[-4]C)*COS(R2C9)*COS(R2C10-R[-3]C)),"""")"
    ActiveSheet.Calculate
    Range("N6:SS6").Value = Range("N6:SS6").Value
    Range("M6").FormulaR1C1 = "=MIN(RC[1]:RC[500])"
    Range("M6").Value = Range("M6").Value
    
    'FINISHES within an hour of sector Fin, >= M6 AFTER H2
    Range("G10").FormulaR1C1 = _
        "=IF(AND(RC[-6]>R2C8,RC[-6]>=R3C8-R6C1,RC[-6]<=R3C8+R6C1,6371*ACOS(SIN(RC[-5])*SIN(R2C9)+COS(RC[-5])*COS(R2C9)*COS(R2C10-RC[-4]))>R6C13,6371*ACOS(SIN(RC[-5])*SIN(R2C9)+COS(RC[-5])*COS(R2C9)*COS(R2C10-RC[-4]))<=R2C13/2),RC[-6],"""")"
    Range("H10:K10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-6],"""")"

    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:K10").AutoFill Destination:=.Range("G10:K" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("G10:K10009").Value = Range("G10:K10009").Value
    Range("G10:K10009").Sort Key1:=Range("G10"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    
    'MATRIX
    Range("N10:SS10").FormulaR1C1 = _
        "=IF(OR(R1C="""",R[1]C7="""",RC9=R3C,R[1]C9=R3C),"""",IF(OR(AND(6371*ACOS(SIN(RC8)*SIN(R2C)+COS(RC8)*COS(R2C)*COS(R3C-RC9))<=0.5,6371*ACOS(SIN(R[1]C8)*SIN(R2C)+COS(R[1]C8)*COS(R2C)*COS(R3C-R[1]C9))<=0.7),AND(6371*ACOS(SIN(RC8)*SIN(R2C)+COS(RC8)*COS(R2C)*COS(R3C-RC9))<=0.7,6371*ACOS(SIN(R[1]C8)*SIN(R2C)+COS(R[1]C8)*COS(R2C)*COS(R3C-R[1]C9))<=0.5)),1,""""))"
   
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
.Range("N10:SS10").AutoFill Destination:=.Range("N10:SS" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("N10:SS600").Value = Range("N10:SS600").Value
    
    Range("N7:SS7").FormulaR1C1 = "=IF(SUM(R[3]C:R[593]C)>=2,R[-6]C,"""")"
    Range("N7:SS7").Copy
    Range("N1000").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    Application.CutCopyMode = False
    
    Range("N7:SS7").FormulaR1C1 = "=IF(SUM(R[3]C:R[593]C)>=2,R[-5]C,"""")"
    Range("N7:SS7").Copy
    Range("O1000").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    Application.CutCopyMode = False
    
    Range("N7:SS7").FormulaR1C1 = "=IF(SUM(R[3]C:R[593]C)>=2,R[-4]C,"""")"
    Range("N7:SS7").Copy
    Range("P1000").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    Application.CutCopyMode = False
    
    Range("N7:SS7").FormulaR1C1 = "=IF(SUM(R[3]C:R[593]C)>=2,R[-3]C,"""")"
    Range("N7:SS7").Copy
    Range("Q1000").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    Application.CutCopyMode = False
    
    Range("N7:SS7").FormulaR1C1 = "=IF(SUM(R[3]C:R[593]C)>=2,R[-2]C,"""")"
    Range("N7:SS7").Copy
    Range("R1000").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    Application.CutCopyMode = False
    
    Range("N7:SS7").FormulaR1C1 = "=IF(SUM(R[3]C:R[593]C)>=2,R[-1]C,"""")"
    Range("N7:SS7").Copy
    Range("S1000").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    Application.CutCopyMode = False
    
    Range("N1000:S1599").Sort Key1:=Range("N1000"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    
    Range("N1:SS700").Clear
   
    Range("N1000:S1499").Copy
    Range("N1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    
    Range("N1000:S1499").Clear
    
If Range("N1") <> "" Then
    
    Range("N7:SS7").FormulaR1C1 = "=IF(R1C="""","""",SQRT(R[-1]C^2+0.25))"
    Range("N7:SS7").Value = Range("N7:SS7").Value
    
    Range("L10").FormulaR1C1 = "=6371*ACOS(SIN(RC[-4])*SIN(R2C9)+COS(RC[-4])*COS(R2C9)*COS(R2C10-RC[-3]))"
    
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
.Range("L10:L10").AutoFill Destination:=.Range("L10:L" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("L10:L10009").Value = Range("L10:L10009").Value
    Range("N10:SS10").FormulaR1C1 = _
        "=IF(OR(R1C="""",RC9=R3C),"""",IF(AND(RC12>R6C,6371*ACOS(SIN(RC8)*SIN(R2C)+COS(RC8)*COS(R2C)*COS(R3C-RC9))<0.7,R[1]C12>RC12),1,""""))"
    
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
.Range("N10:SS10").AutoFill Destination:=.Range("N10:SS" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("N10:SS600").Value = Range("N10:SS600").Value
    
    Range("N8:SS8").FormulaR1C1 = "=IF(R1C="""","""",IF(SUM(R[2]C:R[592]C)>1,R[-2]C,""""))"
    Range("M8").FormulaR1C1 = "=MAX(RC[1]:RC[500])"
    Range("M8:SS8").Value = Range("M8:SS8").Value
    
    Range("L10").FormulaR1C1 = "=IF(MAX(RC[2]:RC[501])=1,6371*ACOS(SIN(RC[-4])*SIN(R2C9)+COS(RC[-4])*COS(R2C9)*COS(R2C10-RC[-3])),0)"
    
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
.Range("L10:L10").AutoFill Destination:=.Range("L10:L" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("L10:L10009").Value = Range("L10:L1009").Value
    Range("N10:SS600").Clear
    
    Range("N9:SS9").FormulaR1C1 = "=IF(R[-1]C=R8C13,R1C,"""")"
    Range("N10:SS10").FormulaR1C1 = "=IF(R[-1]C<>"""",R2C,"""")"
    Range("N11:SS11").FormulaR1C1 = "=IF(R[-1]C<>"""",R3C,"""")"
    Range("N12:SS12").FormulaR1C1 = "=IF(R[-1]C<>"""",R4C,"""")"
    Range("N13:SS13").FormulaR1C1 = "=IF(R[-1]C<>"""",R5C,"""")"
    
    Range("H6").FormulaR1C1 = "=MAX(R9C[6]:R9C[505])"
    Range("I6").FormulaR1C1 = "=MAX(R10C[5]:R10C[504])"
    Range("J6").FormulaR1C1 = "=MAX(R11C[4]:R11C[503])"
    Range("K6").FormulaR1C1 = "=MAX(R12C[3]:R12C[502])"
    Range("L6").FormulaR1C1 = "=MAX(R13C[2]:R13C[501])"
    Range("H6:L6").Value = Range("H6:L6").Value
    Columns("N:SS").Clear
    Range("M6").Clear
    
    Range("H7:L7").Value = Range("H2:L2").Value
    
    Range("M10").FormulaR1C1 = "=IF(RC[-1]>=R8C13,RC[-6],"""")"
    
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
.Range("M10:M10").AutoFill Destination:=.Range("M10:M" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("G9").FormulaR1C1 = "=MIN(R[1]C[6]:R[1000]C[6])"
    ActiveSheet.Calculate
    Range("G9").Value = Range("G9").Value
    Range("G10:M10009").Clear
    Range("F10").FormulaR1C1 = "=IF(RC1=R9C7,1,"""")"
    Range("G10").FormulaR1C1 = "=IF(R[2]C[-1]=1,1,"""")"
    Range("H10").FormulaR1C1 = "=IF(R[-1]C[-1]=1,1,"""")"
    Range("I10").FormulaR1C1 = "=IF(R[-2]C[-1]=1,1,"""")"
    Range("J10").FormulaR1C1 = "=IF(R[-1]C[-1]=1,1,"""")"
    
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("F10:J10").AutoFill Destination:=.Range("F10:J" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("F10:J10009").Value = Range("F10:J10009").Value
    
    Range("L10").FormulaR1C1 = "=IF(OR(RC[-6]=1,RC[-5]=1,RC[-4]=1,RC[-3]=1,RC[-2]=1),RC[-11],"""")"
    Range("M10:P10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-11],"""")"
    
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("L10:P10").AutoFill Destination:=.Range("L10:P" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("L10:P10009").Value = Range("L10:P10009").Value
    Range("L10:P10009").Sort Key1:=Range("L10"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    Range("F9:J10009").Clear
    
    Range("H8:L8").Value = Range("L12:P12").Value
    
    Sheets("YDWK3").Activate
    Sheets("YDWK3").Range("M22:Q22").Value = Sheets("Sheet2").Range("H7:L7").Value
    Sheets("YDWK3").Range("M23:Q23").Value = Sheets("Sheet2").Range("H6:L6").Value
    Sheets("YDWK3").Range("M25:Q29").Value = Sheets("Sheet2").Range("L10:P14").Value
    
    Application.Run "F.xlsm!FINlines"

    Sheets("Sheet2").Activate
    Sheets("Sheet2").Range("H9").Value = Sheets("YDWK3").Range("M17").Value
    Sheets("Sheet2").Range("I9").FormulaR1C1 = "=IF(RC[-1]<>""NO FIN LINE"",""        FINISH LINE"","""")"
    Range("I9").Value = Range("I9").Value
    Sheets("Sheet2").Range("K9:L9").Value = Sheets("YDWK3").Range("N17:O17").Value

    Range("M7").FormulaR1C1 = _
        "=IF(R[2]C[-5]=""NO FIN LINE"","""",IF(R[-1]C[-2]-R[2]C[-2]<=R9C4,2*6371*ACOS(SIN(R[-1]C[-4])*SIN(RC[-4])+COS(R[-1]C[-4])*COS(RC[-4])*COS(RC[-3]-R[-1]C[-3])),2*6371*ACOS(SIN(R[-1]C[-4])*SIN(RC[-4])+COS(R[-1]C[-4])*COS(RC[-4])*COS(RC[-3]-R[-1]C[-3]))-((R[-1]C[-2]-R[2]C[-2]-R9C4)*0.1)))"
    Range("M7").Value = Range("M7").Value

    Range("M8,L10:P14").Clear

Application.Run "F.xlsm!YDWK3clear"
        End If
    End If
    
''''''Sib: N1 Nothing
ElseIf Range("N1") = "" Then
    Range("A5:E7,G6:M100009").Clear
            
    If Range("M3") <> "SECTOR" And ("H4") = "NO FIN LINE" Then
        Application.Run ("F.xlsm!NearestSF")
    
    ElseIf Range("A1") >= 42278 And Range("A4") = "NO FIN LINE" And Range("H4") = "NO FIN LINE" Then
        Application.Run ("F.xlsm!FLineLast")
    End If

    End If
  End If
End If

 Application.Run "F.xlsm!OZSectorB"
End Sub

Sub OZSector()
'
' Tests for OZ Sector On Sheet2 after AORb; Tests for/finds OZ Sector; revised 10/1/2015 for FlineLAST
'
Application.ScreenUpdating = False
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("B1:C1").Value
    Sheets("YDWK3").Range("C6:D6").Value = Sheets("Sheet2").Range("B2:C2").Value
    Sheets("YDWK3").Calculate
    Sheets("Sheet2").Range("C6:C7").Value = Sheets("YDWK3").Range("E2:E3").Value
    Sheets("YDWK3").Range("C5:D6").Value = 0
    Sheets("YDWK3").Activate
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Activate
    'CK OZ for AORb result
    Range("H1:L4").Value = Range("A1:E4").Value
   
   'Find 7 Fixes - 3 before & 3 after raw Fin
    Range("G10").FormulaR1C1 = "=IF(RC[-6]=R3C1,1,"""")"
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
    
    'Test 3 earlier, 3 later than A3
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("O10:P10").Value
    Sheets("YDWK3").Range("C6:D6").Value = Sheets("Sheet2").Range("B1:C1").Value
    Sheets("YDWK3").Calculate
    Sheets("Sheet2").Range("M10").Value = Sheets("YDWK3").Range("C2").Value
    Sheets("YDWK3").Range("C5:D6").Value = 0
    Sheets("YDWK3").Activate
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    Sheets("YDWK3").Calculate
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("O11:P11").Value
    Sheets("YDWK3").Range("C6:D6").Value = Sheets("Sheet2").Range("B1:C1").Value
    Sheets("YDWK3").Calculate
    Sheets("Sheet2").Range("M11").Value = Sheets("YDWK3").Range("C2").Value
    Sheets("YDWK3").Range("C5:D6").Value = 0
    Sheets("YDWK3").Activate
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    Sheets("YDWK3").Calculate
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("O12:P12").Value
    Sheets("YDWK3").Range("C6:D6").Value = Sheets("Sheet2").Range("B1:C1").Value
    Sheets("YDWK3").Calculate
    Sheets("Sheet2").Range("M12").Value = Sheets("YDWK3").Range("C2").Value
    Sheets("YDWK3").Range("C5:D6").Value = 0
    Sheets("YDWK3").Activate
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    Sheets("YDWK3").Calculate
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("O13:P13").Value
    Sheets("YDWK3").Range("C6:D6").Value = Sheets("Sheet2").Range("B1:C1").Value
    Sheets("YDWK3").Calculate
    Sheets("Sheet2").Range("M13").Value = Sheets("YDWK3").Range("C2").Value
    Sheets("YDWK3").Range("C5:D6").Value = 0
    Sheets("YDWK3").Activate
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    Sheets("YDWK3").Calculate
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("O14:P14").Value
    Sheets("YDWK3").Range("C6:D6").Value = Sheets("Sheet2").Range("B1:C1").Value
    Sheets("YDWK3").Calculate
    Sheets("Sheet2").Range("M14").Value = Sheets("YDWK3").Range("C2").Value
    Sheets("YDWK3").Range("C5:D6").Value = 0
    Sheets("YDWK3").Activate
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    Sheets("YDWK3").Calculate
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("O15:P15").Value
    Sheets("YDWK3").Range("C6:D6").Value = Sheets("Sheet2").Range("B1:C1").Value
    Sheets("YDWK3").Calculate
    Sheets("Sheet2").Range("M15").Value = Sheets("YDWK3").Range("C2").Value
    Sheets("YDWK3").Range("C5:D6").Value = 0
    Sheets("YDWK3").Activate
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    Sheets("YDWK3").Calculate
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("O16:P16").Value
    Sheets("YDWK3").Range("C6:D6").Value = Sheets("Sheet2").Range("B1:C1").Value
    Sheets("YDWK3").Calculate
    Sheets("Sheet2").Range("M16").Value = Sheets("YDWK3").Range("C2").Value
    Sheets("YDWK3").Range("C5:D6").Value = 0
    Sheets("YDWK3").Activate
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    Sheets("YDWK3").Calculate
    Application.Run "F.xlsm!YDWK3clear"
    
    Range("G10:G16").FormulaR1C1 = _
        "=IF(OR(6371*ACOS(SIN(RC[8])*SIN(R1C9)+COS(RC[8])*COS(R1C9)*COS(R1C10-RC[9]))>1,RC[6]=""""),"""",IF(OR(AND(R6C3>R7C3,RC[6]>=R7C3,RC[6]<=R6C3),AND(R6C3<R7C3,OR(RC[6]<=R6C3,RC[6]>=R7C3))),RC[7],""""))"
    Range("H10:K16").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[7],"""")"
    Range("G6").FormulaR1C1 = "=MAX(R[4]C[3]:R[10]C[3])"
    Range("G6").Value = Range("G6").Value
    Range("H7").FormulaR1C1 = _
        "=IF(R[3]C[2]=R6C7,R[3]C[-1],IF(R[4]C[2]=R6C7,R[4]C[-1],IF(R[5]C[2]=R6C7,R[5]C[-1],IF(R[6]C[2]=R6C7,R[6]C[-1],""""))))"
    Range("H8").FormulaR1C1 = _
        "=IF(R[6]C[2]=R6C7,R[6]C[-1],IF(R[7]C[2]=R6C7,R[7]C[-1],IF(R[8]C[2]=R6C7,R[8]C[-1],"""")))"
    Range("H7:H8").Value = Range("H7:H8").Value
    Range("G6").FormulaR1C1 = "=MAX(R[1]C[1],R[2]C[1])"
    Range("G6").Value = Range("G6").Value
    Range("H10:H16").FormulaR1C1 = "=IF(R6C7=0,"""",IF(RC[-1]=R6C7,RC[7],""""))"
    Range("H6:K6").FormulaR1C1 = "=MAX(R[4]C:R[10]C)"
    Range("G6:K6").Value = Range("G6:K6").Value
    
    If Range("G6") > 0 Then
        Range("D6").Value = "OK"
        Range("H3:L3").Value = Range("G6:K6").Value
        Range("M2").FormulaR1C1 = _
            "=IF(R[-1]C[-2]-R[1]C[-2]<=R9C4,2*6371*ACOS(SIN(R[-1]C[-4])*SIN(RC[-4])+COS(R[-1]C[-4])*COS(RC[-4])*COS(RC[-3]-R[-1]C[-3])),2*6371*ACOS(SIN(R[-1]C[-4])*SIN(RC[-4])+COS(R[-1]C[-4])*COS(RC[-4])*COS(RC[-3]-R[-1]C[-3]))-((R[-1]C[-2]-R[1]C[-2]-R9C4)*0.1))"
        Range("M2").Value = Range("M2").Value
        Range("M3").Value = "SECTOR"
        Range("G6:R16").Clear
           
    ElseIf Range("G6") = 0 Then
      Range("G6:K6,G10:K16").Clear
      Range("G10:K16").Value = Range("N10:R16").Value
    Range("M10:R16").Clear
    
    'Find Revised Start candidates
    Range("A5").Value = 0.00347222222222
    'find distance to TP @ each fix within 5 mins & 1.5km of Aorb Start
    Range("M10").FormulaR1C1 = _
        "=IF(RC[-10]=R1C3,"""",IF(AND(RC[-12]>=R1C1-R5C1,RC[-12]<=R1C1+R5C1,6371*ACOS(SIN(RC[-11])*SIN(R1C2)+COS(RC[-11])*COS(R1C2)*COS(R1C3-RC[-10]))<1.5),6371*ACOS(SIN(RC[-11])*SIN(R2C2)+COS(RC[-11])*COS(R2C2)*COS(R2C3-RC[-10])),""""))"
    
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("M10:M10").AutoFill Destination:=.Range("M10:M" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("M10:M10019").Value = Range("M10:M10019").Value
    
    Range("L10:L10019").FormulaR1C1 = "=IF(RC[1]<>"""",1,"""")"
    ActiveSheet.Calculate
    Range("L9").FormulaR1C1 = "=SUM(R[1]C:R[10010]C)"
    Range("L9").Value = Range("L9").Value
    Range("L10:L10019").Clear
    
    'find at most 50 longest first legs
    Range("L8").FormulaR1C1 = _
        "=IF(R[1]C<=50,MIN(R[2]C[1]:R[10011]C[1]),LARGE(R[2]C[1]:R[10011]C[1],50))"
    ActiveSheet.Calculate
    Range("N10").FormulaR1C1 = "=IF(AND(RC[-1]<>"""",RC[-1]>=R8C12),RC[-13],"""")"
    Range("O10:R10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-13],"""")"
    
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("N10:R10").AutoFill Destination:=.Range("N10:R" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("N10:R10009").Value = Range("N10:R10009").Value
    Range("N10:R10009").Sort Key1:=Range("N10"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    
    Range("N10:R59").Copy
    Range("N1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    Application.CutCopyMode = False
    
    Range("F9,L8:R10019").Clear

    Range("N8:BK8").FormulaR1C1 = "=IF(R1C<>"""",0.0033528107,"""")"
    Range("N9:BK9").FormulaR1C1 = "=IF(R1C<>"""",R2C3-R[-6]C,"""")"
    Range("N10:BK10").FormulaR1C1 = "=IF(R1C<>"""",(1-R[-2]C)*TAN(R2C),"""")"
    Range("N11:BK11").FormulaR1C1 = "=IF(R1C<>"""",ATAN(R[-1]C),"""")"
    Range("N12:BK12").FormulaR1C1 = "=IF(R1C<>"""",SIN(R[-1]C),"""")"
    Range("N13:BK13").FormulaR1C1 = "=IF(R1C<>"""",COS(R[-2]C),"""")"
    Range("N14:BK14").FormulaR1C1 = "=IF(R1C<>"""",(1-R[-6]C)*TAN(R2C2),"""")"
    Range("N15:BK15").FormulaR1C1 = "=IF(R1C<>"""",ATAN(R[-1]C),"""")"
    Range("N16:BK16").FormulaR1C1 = "=IF(R1C<>"""",SIN(R[-1]C),"""")"
    Range("N17:BK17").FormulaR1C1 = "=IF(R1C<>"""",COS(R[-2]C),"""")"
    Range("N18:BK18").FormulaR1C1 = "=R[6]C"
    Range("N19:BK19").FormulaR1C1 = _
        "=IF(R1C<>"""",(R[-2]C*SIN(R[-1]C)*R[-2]C*SIN(R[-1]C))+(R[-6]C*R[-3]C-R[-7]C*R[-2]C*COS(R[-1]C))*(R[-6]C*R[-3]C-R[-7]C*R[-2]C*COS(R[-1]C)),"""")"
    Range("N20:BK20").FormulaR1C1 = "=IF(R1C<>"""",(R[-8]C*R[-4]C)+(R[-7]C*R[-3]C*COS(R[-2]C)),"""")"
    Range("N21:BK21").FormulaR1C1 = _
        "=IF(R1C="""","""",IF(R[-2]C=0,0,R[-8]C*R[-4]C*SIN(R[-3]C)/SQRT(R[-2]C)))"
    Range("N22:BK22").FormulaR1C1 = _
        "=IF(R1C<>"""",R[-2]C-2*R[-10]C*R[-6]C/(COS(ASIN(R[-1]C))*COS(ASIN(R[-1]C))),"""")"
    Range("N23:BK23").FormulaR1C1 = _
        "=IF(R1C<>"""",R[-15]C/16*COS(ASIN(R[-2]C))*COS(ASIN(R[-2]C))*(4+R[-15]C*(4-3*COS(ASIN(R[-2]C))*COS(ASIN(R[-2]C)))),"""")"
    Range("N24:BK24").FormulaR1C1 = _
        "=IF(R1C<>"""",R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C))),"""")"
    Range("N25:BK25").FormulaR1C1 = _
        "=IF(R1C<>"""",ATAN2((R[-12]C*R[-9]C-R[-13]C*R[-8]C*COS(R[-1]C)),R[-8]C*SIN(R[-1]C)),"""")"
    Range("N26:BK26").FormulaR1C1 = _
        "=IF(R1C<>"""",ATAN2((-R[-14]C*R[-9]C+R[-13]C*R[-10]C*COS(R[-2]C)),R[-13]C*SIN(R[-2]C)),"""")"
    Range("N27:BK27").FormulaR1C1 = "=IF(R1C="""","""",IF(AND(R2C=R2C2,R3C=R2C3),""samepoint"",""N.A. "" ))"
    Range("N28:BK28").FormulaR1C1 = "=IF(R1C="""","""",IF(AND(R3C=R2C3,R2C2>R2C),""northsouth"",""N.A. ""))"
    Range("N29:BK29").FormulaR1C1 = "=IF(R1C="""","""",IF(AND(R3C=R2C3,R2C>R2C2),""southnorth"",""N.A. ""))"
    Range("N30:BK30").FormulaR1C1 = _
        "=IF(R1C="""","""",IF(R[-3]C=""samepoint"",0,IF(R[-2]C=""northsouth"",0,IF(R[-1]C=""southnorth"",180,IF(R[-5]C<0,R[-5]C*180/PI()+360,R[-5]C*180/PI())))))"
    Range("N31:BK31").FormulaR1C1 = _
        "=IF(R1C="""","""",IF(R[-4]C=""samepoint"",0,IF(R[-3]C=""northsouth"",180,IF(R[-2]C=""southnorth"",0,R[-5]C*180/PI()+180))))"
    Range("N32:BK32").FormulaR1C1 = "=IF(R1C="""","""",IF(R[-2]C+45>360,R[-2]C+45-360,R[-2]C+45))"
    Range("N33:BK33").FormulaR1C1 = "=IF(R1C="""","""",IF(R[-3]C-45<0,R[-3]C-45+360,R[-3]C-45))"
    ActiveSheet.Calculate
    Range("N8:BK33").Value = Range("N8:BK33").Value
    Range("N6:BK7").Value = Range("N32:BK33").Value
    
    Range("N8:BK10008").Clear
    
    Application.Run "F.xlsm!Sect1"
    Application.Run "F.xlsm!Last"

    If Range("A1") < 42278 And Range("M3") <> "" Then
        Application.Run "F.xlsm!OZsectorB"
        Exit Sub
    
    ElseIf Range("A4") = "NO FIN LINE" And Range("H4") = "NO FIN LINE" And Range("H9") = "NO FIN LINE" And Range("M3") = "" Then
        Application.Run "F.xlsm!NearestSF"
        Exit Sub
    
    ElseIf Range("A1") >= 42278 And Range("A4") = "NO FIN LINE" And Range("H4") = "NO FIN LINE" Then
        Application.Run "F.xlsm!FlineLast"
        Exit Sub
    End If
        
End If

Application.Run "F.xlsm!Fline2"
'Application.Run "F.xlsm!OZSectorB"
End Sub

Sub Sect1()
'
' CK for FIN OZ BY COLUMN  10/7/13 Cks 7 Finish Alternatives only 2 flaws corrected 11/1/13
'
'Column N
Application.ScreenUpdating = False
If Range("N1") <> "" Then
    Sheets("YDWK3").Activate
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H10:I10").Value
    Sheets("YDWK3").Range("C6").Value = Sheets("Sheet2").Range("N2").Value
    Sheets("YDWK3").Range("D6").Value = Sheets("Sheet2").Range("N3").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("N10").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H11:I11").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("N11").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H12:I12").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("N12").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H13:I13").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("N13").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H14:I14").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("N14").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H15:I15").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("N15").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H16:I16").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("N16").Value = Sheets("YDWK3").Range("C2").Value

' Test Column(Starts) for FIN in OZ Sector
    Sheets("Sheet2").Activate
    Range("N8").FormulaR1C1 = "=MAX(R[2]C12:R[493]C12)"
    Range("L10").FormulaR1C1 = _
        "=IF(OR(AND(R6C[2]>R7C[2],RC[2]>=R7C[2],RC[2]<=R6C[2]),AND(R6C[2]<R7C[2],OR(RC[2]>=R7C[2],RC[2]<=R6C[2]))),RC[2],-1)"

    Application.Calculation = xlCalculationManual
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
.Range("L10:L10").AutoFill Destination:=.Range("L10:L" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("N8").Value = Range("N8").Value

If Range("N8") < 0 Then
    Range("N10:N500").ClearContents
End If
End If

'Column O
If Range("O1") <> "" Then
    Sheets("YDWK3").Activate
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H10:I10").Value
    Sheets("YDWK3").Range("C6").Value = Sheets("Sheet2").Range("O2").Value
    Sheets("YDWK3").Range("D6").Value = Sheets("Sheet2").Range("O3").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("O10").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H11:I11").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("O11").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H12:I12").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("O12").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H13:I13").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("O13").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H14:I14").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("O14").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H15:I15").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("O15").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H16:I16").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("O16").Value = Sheets("YDWK3").Range("C2").Value

' Test Column(Starts) for FIN in OZ Sector
    Sheets("Sheet2").Activate
    Range("O8").FormulaR1C1 = "=MAX(R[2]C12:R[493]C12)"
    Range("L10").FormulaR1C1 = _
        "=IF(OR(AND(R6C[3]>R7C[3],RC[3]>=R7C[3],RC[3]<=R6C[3]),AND(R6C[3]<R7C[3],OR(RC[3]>=R7C[3],RC[3]<=R6C[3]))),RC[3],-1)"

    Application.Calculation = xlCalculationManual
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
.Range("L10:L10").AutoFill Destination:=.Range("L10:L" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("O8").Value = Range("O8").Value

If Range("O8") < 0 Then
    Range("O10:O500").ClearContents
End If
End If

'Column P
If Range("P1") <> "" Then
Sheets("YDWK3").Activate
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H10:I10").Value
    Sheets("YDWK3").Range("C6").Value = Sheets("Sheet2").Range("P2").Value
    Sheets("YDWK3").Range("D6").Value = Sheets("Sheet2").Range("P3").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("P10").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H11:I11").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("P11").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H12:I12").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("P12").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H13:I13").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("P13").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H14:I14").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("P14").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H15:I15").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("P15").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H16:I16").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("P16").Value = Sheets("YDWK3").Range("C2").Value

' Test Column(Starts) for FIN in OZ Sector
    Sheets("Sheet2").Activate
    Range("P8").FormulaR1C1 = "=MAX(R[2]C12:R[493]C12)"
    Range("L10").FormulaR1C1 = _
        "=IF(OR(AND(R6C[4]>R7C[4],RC[4]>=R7C[4],RC[4]<=R6C[4]),AND(R6C[4]<R7C[4],OR(RC[4]>=R7C[4],RC[4]<=R6C[4]))),RC[4],-1)"
    
    Application.Calculation = xlCalculationManual
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
.Range("L10:L10").AutoFill Destination:=.Range("L10:L" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("P8").Value = Range("P8").Value

If Range("P8") < 0 Then
    Range("P10:P500").ClearContents
End If
End If

'Column Q
If Range("Q1") <> "" Then
Sheets("YDWK3").Activate
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H10:I10").Value
    Sheets("YDWK3").Range("C6").Value = Sheets("Sheet2").Range("Q2").Value
    Sheets("YDWK3").Range("D6").Value = Sheets("Sheet2").Range("Q3").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("Q10").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H11:I11").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("Q11").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H12:I12").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("Q12").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H13:I13").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("Q13").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H14:I14").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("Q14").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H15:I15").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("Q15").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H16:I16").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("Q16").Value = Sheets("YDWK3").Range("C2").Value

' Test Column(Starts) for FIN in OZ Sector
    Sheets("Sheet2").Activate
    Range("Q8").FormulaR1C1 = "=MAX(R[2]C12:R[493]C12)"
    Range("L10").FormulaR1C1 = _
        "=IF(OR(AND(R6C[5]>R7C[5],RC[5]>=R7C[5],RC[5]<=R6C[5]),AND(R6C[5]<R7C[5],OR(RC[5]>=R7C[5],RC[5]<=R6C[5]))),RC[5],-1)"
    
    Application.Calculation = xlCalculationManual
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
.Range("L10:L10").AutoFill Destination:=.Range("L10:L" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("Q8").Value = Range("Q8").Value

If Range("Q8") < 0 Then
    Range("Q10:Q500").ClearContents
End If
End If

'Column R
If Range("R1") <> "" Then
Sheets("YDWK3").Activate
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H10:I10").Value
    Sheets("YDWK3").Range("C6").Value = Sheets("Sheet2").Range("R2").Value
    Sheets("YDWK3").Range("D6").Value = Sheets("Sheet2").Range("R3").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("R10").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H11:I11").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("R11").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H12:I12").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("R12").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H13:I13").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("R13").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H14:I14").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("R14").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H15:I15").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("R15").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H16:I16").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("R16").Value = Sheets("YDWK3").Range("C2").Value

' Test Column(Starts) for FIN in OZ Sector
    Sheets("Sheet2").Activate
    Range("R8").FormulaR1C1 = "=MAX(R[2]C12:R[493]C12)"
    Range("L10").FormulaR1C1 = _
        "=IF(OR(AND(R6C[6]>R7C[6],RC[6]>=R7C[6],RC[6]<=R6C[6]),AND(R6C[6]<R7C[6],OR(RC[6]>=R7C[6],RC[6]<=R6C[6]))),RC[6],-1)"
    
    Application.Calculation = xlCalculationManual
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
.Range("L10:L10").AutoFill Destination:=.Range("L10:L" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("R8").Value = Range("R8").Value

If Range("R8") < 0 Then
    Range("R10:R500").ClearContents
End If
End If

'Column S
If Range("S1") <> "" Then
Sheets("YDWK3").Activate
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H10:I10").Value
    Sheets("YDWK3").Range("C6").Value = Sheets("Sheet2").Range("S2").Value
    Sheets("YDWK3").Range("D6").Value = Sheets("Sheet2").Range("S3").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("S10").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H11:I11").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("S11").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H12:I12").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("S12").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H13:I13").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("S13").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H14:I14").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("S14").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H15:I15").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("S15").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H16:I16").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("S16").Value = Sheets("YDWK3").Range("C2").Value

' Test Column(Starts) for FIN in OZ Sector
    Sheets("Sheet2").Activate
    Range("S8").FormulaR1C1 = "=MAX(R[2]C12:R[493]C12)"
    Range("L10").FormulaR1C1 = _
        "=IF(OR(AND(R6C[7]>R7C[7],RC[7]>=R7C[7],RC[7]<=R6C[7]),AND(R6C[7]<R7C[7],OR(RC[7]>=R7C[7],RC[7]<=R6C[7]))),RC[7],-1)"
    
    Application.Calculation = xlCalculationManual
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
.Range("L10:L10").AutoFill Destination:=.Range("L10:L" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("S8").Value = Range("S8").Value

If Range("S8") < 0 Then
    Range("S10:S500").ClearContents
End If
End If

'Column T
If Range("T1") <> "" Then
Sheets("YDWK3").Activate
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H10:I10").Value
    Sheets("YDWK3").Range("C6").Value = Sheets("Sheet2").Range("T2").Value
    Sheets("YDWK3").Range("D6").Value = Sheets("Sheet2").Range("T3").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("T10").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H11:I11").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("T11").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H12:I12").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("T12").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H13:I13").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("T13").Value = Sheets("YDWK3").Range("C2").Value

     Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H14:I14").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("T14").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H15:I15").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("T15").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H16:I16").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("T16").Value = Sheets("YDWK3").Range("C2").Value

' Test Column(Starts) for FIN in OZ Sector
    Sheets("Sheet2").Activate
    Range("T8").FormulaR1C1 = "=MAX(R[2]C12:R[493]C12)"
    Range("L10").FormulaR1C1 = _
        "=IF(OR(AND(R6C[8]>R7C[8],RC[8]>=R7C[8],RC[8]<=R6C[8]),AND(R6C[8]<R7C[8],OR(RC[8]>=R7C[8],RC[8]<=R6C[8]))),RC[8],-1)"
    
    Application.Calculation = xlCalculationManual
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
.Range("L10:L10").AutoFill Destination:=.Range("L10:L" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("T8").Value = Range("T8").Value

If Range("T8") < 0 Then
    Range("T10:T500").ClearContents
End If
End If

'Column U
If Range("U1") <> "" Then
Sheets("YDWK3").Activate
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H10:I10").Value
    Sheets("YDWK3").Range("C6").Value = Sheets("Sheet2").Range("U2").Value
    Sheets("YDWK3").Range("D6").Value = Sheets("Sheet2").Range("U3").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("U10").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H11:I11").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("U11").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H12:I12").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("U12").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H13:I13").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("U13").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H14:I14").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("U14").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H15:I15").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("U15").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H16:I16").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("U16").Value = Sheets("YDWK3").Range("C2").Value

' Test Column(Starts) for FIN in OZ Sector
    Sheets("Sheet2").Activate
    Range("U8").FormulaR1C1 = "=MAX(R[2]C12:R[493]C12)"
    Range("L10").FormulaR1C1 = _
        "=IF(OR(AND(R6C[9]>R7C[9],RC[9]>=R7C[9],RC[9]<=R6C[9]),AND(R6C[9]<R7C[9],OR(RC[9]>=R7C[9],RC[9]<=R6C[9]))),RC[9],-1)"
    
    Application.Calculation = xlCalculationManual
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
.Range("L10:L10").AutoFill Destination:=.Range("L10:L" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("U8").Value = Range("U8").Value

If Range("U8") < 0 Then
    Range("U10:U500").ClearContents
End If
End If

'Column V
If Range("V1") <> "" Then
Sheets("YDWK3").Activate
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H10:I10").Value
    Sheets("YDWK3").Range("C6").Value = Sheets("Sheet2").Range("V2").Value
    Sheets("YDWK3").Range("D6").Value = Sheets("Sheet2").Range("V3").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("V10").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H11:I11").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("V11").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H12:I12").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("V12").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H13:I13").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("V13").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H14:I14").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("V14").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H15:I15").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("V15").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H16:I16").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("V16").Value = Sheets("YDWK3").Range("C2").Value

' Test Column(Starts) for FIN in OZ Sector
    Sheets("Sheet2").Activate
    Range("V8").FormulaR1C1 = "=MAX(R[2]C12:R[493]C12)"
    Range("L10").FormulaR1C1 = _
        "=IF(OR(AND(R6C[10]>R7C[10],RC[10]>=R7C[10],RC[10]<=R6C[10]),AND(R6C[10]<R7C[10],OR(RC[10]>=R7C[10],RC[10]<=R6C[10]))),RC[10],-1)"
    
    Application.Calculation = xlCalculationManual
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
.Range("L10:L10").AutoFill Destination:=.Range("L10:L" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("V8").Value = Range("V8").Value

If Range("V8") < 0 Then
    Range("V10:V500").ClearContents
End If
End If

'Column W
If Range("W1") <> "" Then
Sheets("YDWK3").Activate
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H10:I10").Value
    Sheets("YDWK3").Range("C6").Value = Sheets("Sheet2").Range("W2").Value
    Sheets("YDWK3").Range("D6").Value = Sheets("Sheet2").Range("W3").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("W10").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H11:I11").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("W11").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H12:I12").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("W12").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H13:I13").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("W13").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H14:I14").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("W14").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H15:I15").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("W15").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H16:I16").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("W16").Value = Sheets("YDWK3").Range("C2").Value

' Test Column(Starts) for FIN in OZ Sector
    Sheets("Sheet2").Activate
    Range("W8").FormulaR1C1 = "=MAX(R[2]C12:R[493]C12)"
    Range("L10").FormulaR1C1 = _
        "=IF(OR(AND(R6C[11]>R7C[11],RC[11]>=R7C[11],RC[11]<=R6C[11]),AND(R6C[11]<R7C[11],OR(RC[11]>=R7C[11],RC[11]<=R6C[11]))),RC[11],-1)"
    
    Application.Calculation = xlCalculationManual
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
.Range("L10:L10").AutoFill Destination:=.Range("L10:L" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("W8").Value = Range("W8").Value

If Range("W8") < 0 Then
    Range("W10:W500").ClearContents
End If
End If

If Range("X1") <> "" Then Application.Run "F.xlsm!Sect2"

End Sub

Sub Sect2()
'
' OZSect2 Macro
'
'Column X
Application.ScreenUpdating = False
If Range("X1") <> "" Then
    Sheets("YDWK3").Activate
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H10:I10").Value
    Sheets("YDWK3").Range("C6").Value = Sheets("Sheet2").Range("X2").Value
    Sheets("YDWK3").Range("D6").Value = Sheets("Sheet2").Range("X3").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("X10").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H11:I11").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("X11").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H12:I12").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("X12").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H13:I13").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("X13").Value = Sheets("YDWK3").Range("C2").Value

     Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H14:I14").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("X14").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H15:I15").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("X15").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H16:I16").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("X16").Value = Sheets("YDWK3").Range("C2").Value

' Test Column(Starts) for FIN in OZ Sector
    Sheets("Sheet2").Activate
    Range("X8").FormulaR1C1 = "=MAX(R[2]C12:R[493]C12)"
    Range("L10").FormulaR1C1 = _
        "=IF(OR(AND(R6C[12]>R7C[12],RC[12]>=R7C[12],RC[12]<=R6C[12]),AND(R6C[12]<R7C[12],OR(RC[12]>=R7C[12],RC[12]<=R6C[12]))),RC[12],-1)"

    Application.Calculation = xlCalculationManual
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
.Range("L10:L10").AutoFill Destination:=.Range("L10:L" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("X8").Value = Range("X8").Value

If Range("X8") < 0 Then
    Range("X10:X500").ClearContents
End If
End If

'Column Y
If Range("Y1") <> "" Then
    Sheets("YDWK3").Activate
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H10:I10").Value
    Sheets("YDWK3").Range("C6").Value = Sheets("Sheet2").Range("Y2").Value
    Sheets("YDWK3").Range("D6").Value = Sheets("Sheet2").Range("Y3").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("Y10").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H11:I11").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("Y11").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H12:I12").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("Y12").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H13:I13").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("Y13").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H14:I14").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("Y14").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H15:I15").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("Y15").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H16:I16").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("Y16").Value = Sheets("YDWK3").Range("C2").Value

' Test Column(Starts) for FIN in OZ Sector
    Sheets("Sheet2").Activate
    Range("Y8").FormulaR1C1 = "=MAX(R[2]C12:R[493]C12)"
    Range("L10").FormulaR1C1 = _
        "=IF(OR(AND(R6C[13]>R7C[13],RC[13]>=R7C[13],RC[13]<=R6C[13]),AND(R6C[13]<R7C[13],OR(RC[13]>=R7C[13],RC[13]<=R6C[13]))),RC[13],-1)"

    Application.Calculation = xlCalculationManual
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
.Range("L10:L10").AutoFill Destination:=.Range("L10:L" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("Y8").Value = Range("Y8").Value

If Range("Y8") < 0 Then
    Range("Y10:Y500").ClearContents
End If
End If

'Column Z
If Range("Z1") <> "" Then
Sheets("YDWK3").Activate
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H10:I10").Value
    Sheets("YDWK3").Range("C6").Value = Sheets("Sheet2").Range("Z2").Value
    Sheets("YDWK3").Range("D6").Value = Sheets("Sheet2").Range("Z3").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("Z10").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H11:I11").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("Z11").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H12:I12").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("Z12").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H13:I13").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("Z13").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H14:I14").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("Z14").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H15:I15").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("Z15").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H16:I16").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("Z16").Value = Sheets("YDWK3").Range("C2").Value

' Test Column(Starts) for FIN in OZ Sector
    Sheets("Sheet2").Activate
    Range("Z8").FormulaR1C1 = "=MAX(R[2]C12:R[493]C12)"
    Range("L10").FormulaR1C1 = _
        "=IF(OR(AND(R6C[14]>R7C[14],RC[14]>=R7C[14],RC[14]<=R6C[14]),AND(R6C[14]<R7C[14],OR(RC[14]>=R7C[14],RC[14]<=R6C[14]))),RC[14],-1)"
    
    Application.Calculation = xlCalculationManual
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
.Range("L10:L10").AutoFill Destination:=.Range("L10:L" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("Z8").Value = Range("Z8").Value

If Range("Z8") < 0 Then
    Range("Z10:Z500").ClearContents
End If
End If

'Column AA
If Range("AA1") <> "" Then
Sheets("YDWK3").Activate
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H10:I10").Value
    Sheets("YDWK3").Range("C6").Value = Sheets("Sheet2").Range("AA2").Value
    Sheets("YDWK3").Range("D6").Value = Sheets("Sheet2").Range("AA3").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AA10").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H11:I11").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AA11").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H12:I12").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AA12").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H13:I13").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AA13").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H14:I14").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AA14").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H15:I15").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AA15").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H16:I16").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AA16").Value = Sheets("YDWK3").Range("C2").Value

' Test Column(Starts) for FIN in OZ Sector
    Sheets("Sheet2").Activate
    Range("AA8").FormulaR1C1 = "=MAX(R[2]C12:R[493]C12)"
    Range("L10").FormulaR1C1 = _
        "=IF(OR(AND(R6C[15]>R7C[15],RC[15]>=R7C[15],RC[15]<=R6C[15]),AND(R6C[15]<R7C[15],OR(RC[15]>=R7C[15],RC[15]<=R6C[15]))),RC[15],-1)"
    
    Application.Calculation = xlCalculationManual
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
.Range("L10:L10").AutoFill Destination:=.Range("L10:L" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("AA8").Value = Range("AA8").Value

If Range("AA8") < 0 Then
    Range("AA10:AA500").ClearContents
End If
End If

'Column AB
If Range("AB1") <> "" Then
Sheets("YDWK3").Activate
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H10:I10").Value
    Sheets("YDWK3").Range("C6").Value = Sheets("Sheet2").Range("AB2").Value
    Sheets("YDWK3").Range("D6").Value = Sheets("Sheet2").Range("AB3").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AB10").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H11:I11").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AB11").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H12:I12").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AB12").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H13:I13").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AB13").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H14:I14").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AB14").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H15:I15").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AB15").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H16:I16").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AB16").Value = Sheets("YDWK3").Range("C2").Value

' Test Column(Starts) for FIN in OZ Sector
    Sheets("Sheet2").Activate
    Range("AB8").FormulaR1C1 = "=MAX(R[2]C12:R[493]C12)"
    Range("L10").FormulaR1C1 = _
        "=IF(OR(AND(R6C[16]>R7C[16],RC[16]>=R7C[16],RC[16]<=R6C[16]),AND(R6C[16]<R7C[16],OR(RC[16]>=R7C[16],RC[16]<=R6C[16]))),RC[16],-1)"
    
    Application.Calculation = xlCalculationManual
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
.Range("L10:L10").AutoFill Destination:=.Range("L10:L" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("AB8").Value = Range("AB8").Value

If Range("AB8") < 0 Then
    Range("AB10:AB500").ClearContents
End If
End If

'Column AC
If Range("AC1") <> "" Then
Sheets("YDWK3").Activate
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H10:I10").Value
    Sheets("YDWK3").Range("C6").Value = Sheets("Sheet2").Range("AC2").Value
    Sheets("YDWK3").Range("D6").Value = Sheets("Sheet2").Range("AC3").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AC10").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H11:I11").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AC11").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H12:I12").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AC12").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H13:I13").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AC13").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H14:I14").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AC14").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H15:I15").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AC15").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H16:I16").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AC16").Value = Sheets("YDWK3").Range("C2").Value

' Test Column(Starts) for FIN in OZ Sector
    Sheets("Sheet2").Activate
    Range("AC8").FormulaR1C1 = "=MAX(R[2]C12:R[493]C12)"
    Range("L10").FormulaR1C1 = _
        "=IF(OR(AND(R6C[17]>R7C[17],RC[17]>=R7C[17],RC[17]<=R6C[17]),AND(R6C[17]<R7C[17],OR(RC[17]>=R7C[17],RC[17]<=R6C[17]))),RC[17],-1)"
    
    Application.Calculation = xlCalculationManual
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
.Range("L10:L10").AutoFill Destination:=.Range("L10:L" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("AC8").Value = Range("AC8").Value

If Range("AC8") < 0 Then
    Range("AC10:AC500").ClearContents
End If
End If

'Column AD
If Range("AD1") <> "" Then
Sheets("YDWK3").Activate
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H10:I10").Value
    Sheets("YDWK3").Range("C6").Value = Sheets("Sheet2").Range("AD2").Value
    Sheets("YDWK3").Range("D6").Value = Sheets("Sheet2").Range("AD3").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AD10").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H11:I11").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AD11").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H12:I12").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AD12").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H13:I13").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AD13").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H14:I14").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AD14").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H15:I15").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AD15").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H16:I16").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AD16").Value = Sheets("YDWK3").Range("C2").Value

' Test Column(Starts) for FIN in OZ Sector
    Sheets("Sheet2").Activate
    Range("AD8").FormulaR1C1 = "=MAX(R[2]C12:R[493]C12)"
    Range("L10").FormulaR1C1 = _
        "=IF(OR(AND(R6C[18]>R7C[18],RC[18]>=R7C[18],RC[18]<=R6C[18]),AND(R6C[18]<R7C[18],OR(RC[18]>=R7C[18],RC[18]<=R6C[18]))),RC[18],-1)"
    
    Application.Calculation = xlCalculationManual
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
.Range("L10:L10").AutoFill Destination:=.Range("L10:L" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("AD8").Value = Range("AD8").Value

If Range("AD8") < 0 Then
    Range("AD10:AD500").ClearContents
End If
End If

'Column AE
If Range("AE1") <> "" Then
Sheets("YDWK3").Activate
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H10:I10").Value
    Sheets("YDWK3").Range("C6").Value = Sheets("Sheet2").Range("AE2").Value
    Sheets("YDWK3").Range("D6").Value = Sheets("Sheet2").Range("AE3").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AE10").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H11:I11").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AE11").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H12:I12").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AE12").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H13:I13").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AE13").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H14:I14").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AE14").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H15:I15").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AE15").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H16:I16").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AE16").Value = Sheets("YDWK3").Range("C2").Value

' Test Column(Starts) for FIN in OZ Sector
    Sheets("Sheet2").Activate
    Range("AE8").FormulaR1C1 = "=MAX(R[2]C12:R[493]C12)"
    Range("L10").FormulaR1C1 = _
        "=IF(OR(AND(R6C[19]>R7C[19],RC[19]>=R7C[19],RC[19]<=R6C[19]),AND(R6C[19]<R7C[19],OR(RC[19]>=R7C[19],RC[19]<=R6C[19]))),RC[19],-1)"
    
    Application.Calculation = xlCalculationManual
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
.Range("L10:L10").AutoFill Destination:=.Range("L10:L" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("AE8").Value = Range("AE8").Value

If Range("AE8") < 0 Then
    Range("AE10:AE500").ClearContents
End If
End If

'Column AF
If Range("AF1") <> "" Then
Sheets("YDWK3").Activate
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H10:I10").Value
    Sheets("YDWK3").Range("C6").Value = Sheets("Sheet2").Range("AF2").Value
    Sheets("YDWK3").Range("D6").Value = Sheets("Sheet2").Range("AF3").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AF10").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H11:I11").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AF11").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H12:I12").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AF12").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H13:I13").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AF13").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H14:I14").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AF14").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H15:I15").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AF15").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H16:I16").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AF16").Value = Sheets("YDWK3").Range("C2").Value

' Test Column(Starts) for FIN in OZ Sector
    Sheets("Sheet2").Activate
    Range("AF8").FormulaR1C1 = "=MAX(R[2]C12:R[493]C12)"
    Range("L10").FormulaR1C1 = _
        "=IF(OR(AND(R6C[20]>R7C[20],RC[20]>=R7C[20],RC[20]<=R6C[20]),AND(R6C[20]<R7C[20],OR(RC[20]>=R7C[20],RC[20]<=R6C[20]))),RC[20],-1)"
    
    Application.Calculation = xlCalculationManual
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
.Range("L10:L10").AutoFill Destination:=.Range("L10:L" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("AF8").Value = Range("AF8").Value

If Range("AF8") < 0 Then
    Range("AF10:AF500").ClearContents
End If
End If

'Column AG
If Range("AG1") <> "" Then
Sheets("YDWK3").Activate
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H10:I10").Value
    Sheets("YDWK3").Range("C6").Value = Sheets("Sheet2").Range("AG2").Value
    Sheets("YDWK3").Range("D6").Value = Sheets("Sheet2").Range("AG3").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AG10").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H11:I11").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AG11").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H12:I12").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AG12").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H13:I13").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AG13").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H14:I14").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AG14").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H15:I15").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AG15").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H16:I16").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AG16").Value = Sheets("YDWK3").Range("C2").Value

' Test Column(Starts) for FIN in OZ Sector
    Sheets("Sheet2").Activate
    Range("AG8").FormulaR1C1 = "=MAX(R[2]C12:R[493]C12)"
    Range("L10").FormulaR1C1 = _
        "=IF(OR(AND(R6C[21]>R7C[21],RC[21]>=R7C[21],RC[21]<=R6C[21]),AND(R6C[21]<R7C[21],OR(RC[21]>=R7C[21],RC[21]<=R6C[21]))),RC[21],-1)"
    
    Application.Calculation = xlCalculationManual
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
.Range("L10:L10").AutoFill Destination:=.Range("L10:L" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("AG8").Value = Range("AG8").Value

If Range("AG8") < 0 Then
    Range("AG10:AG500").ClearContents
End If
End If

If Range("AH1") <> "" Then Application.Run "F.xlsm!Sect3"

End Sub

Sub Sect3()
'
' OZSect3 Macro
'
'Column AH
Application.ScreenUpdating = False
If Range("AH1") <> "" Then
    Sheets("YDWK3").Activate
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H10:I10").Value
    Sheets("YDWK3").Range("C6").Value = Sheets("Sheet2").Range("AH2").Value
    Sheets("YDWK3").Range("D6").Value = Sheets("Sheet2").Range("AH3").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AH10").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H11:I11").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AH11").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H12:I12").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AH12").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H13:I13").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AH13").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H14:I14").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AH14").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H15:I15").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AH15").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H16:I16").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AH16").Value = Sheets("YDWK3").Range("C2").Value

' Test Column(Starts) for FIN in OZ Sector
    Sheets("Sheet2").Activate
    Range("AH8").FormulaR1C1 = "=MAX(R[2]C12:R[493]C12)"
    Range("L10").FormulaR1C1 = _
        "=IF(OR(AND(R6C[22]>R7C[22],RC[22]>=R7C[22],RC[22]<=R6C[22]),AND(R6C[22]<R7C[22],OR(RC[22]>=R7C[22],RC[22]<=R6C[22]))),RC[22],-1)"

    Application.Calculation = xlCalculationManual
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
.Range("L10:L10").AutoFill Destination:=.Range("L10:L" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("AH8").Value = Range("AH8").Value

If Range("AH8") < 0 Then
    Range("AH10:AH500").ClearContents
End If
End If

'Column AI
If Range("AI1") <> "" Then
    Sheets("YDWK3").Activate
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H10:I10").Value
    Sheets("YDWK3").Range("C6").Value = Sheets("Sheet2").Range("AI2").Value
    Sheets("YDWK3").Range("D6").Value = Sheets("Sheet2").Range("AI3").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AI10").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H11:I11").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AI11").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H12:I12").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AI12").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H13:I13").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AI13").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H14:I14").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AI14").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H15:I15").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AI15").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H16:I16").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AI16").Value = Sheets("YDWK3").Range("C2").Value

' Test Column(Starts) for FIN in OZ Sector
    Sheets("Sheet2").Activate
    Range("AI8").FormulaR1C1 = "=MAX(R[2]C12:R[493]C12)"
    Range("L10").FormulaR1C1 = _
        "=IF(OR(AND(R6C[23]>R7C[23],RC[23]>=R7C[23],RC[23]<=R6C[23]),AND(R6C[23]<R7C[23],OR(RC[23]>=R7C[23],RC[23]<=R6C[23]))),RC[23],-1)"

    Application.Calculation = xlCalculationManual
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
.Range("L10:L10").AutoFill Destination:=.Range("L10:L" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("AI8").Value = Range("AI8").Value

If Range("AI8") < 0 Then
    Range("AI10:AI500").ClearContents
End If
End If

'Column AJ
If Range("AJ1") <> "" Then
    Sheets("YDWK3").Activate
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H10:I10").Value
    Sheets("YDWK3").Range("C6").Value = Sheets("Sheet2").Range("AJ2").Value
    Sheets("YDWK3").Range("D6").Value = Sheets("Sheet2").Range("AJ3").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AJ10").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H11:I11").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AJ11").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H12:I12").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AJ12").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H13:I13").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AJ13").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H14:I14").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AJ14").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H15:I15").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AJ15").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H16:I16").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AJ16").Value = Sheets("YDWK3").Range("C2").Value

' Test Column(Starts) for FIN in OZ Sector
    Sheets("Sheet2").Activate
    Range("AJ8").FormulaR1C1 = "=MAX(R[2]C12:R[493]C12)"
    Range("L10").FormulaR1C1 = _
        "=IF(OR(AND(R6C[24]>R7C[24],RC[24]>=R7C[24],RC[24]<=R6C[24]),AND(R6C[24]<R7C[24],OR(RC[24]>=R7C[24],RC[24]<=R6C[24]))),RC[24],-1)"

    Application.Calculation = xlCalculationManual
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
.Range("L10:L10").AutoFill Destination:=.Range("L10:L" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("AJ8").Value = Range("AJ8").Value

If Range("AJ8") < 0 Then
    Range("AJ10:AJ500").ClearContents
End If
End If

'Column AK
If Range("AK1") <> "" Then
Sheets("YDWK3").Activate
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H10:I10").Value
    Sheets("YDWK3").Range("C6").Value = Sheets("Sheet2").Range("AK2").Value
    Sheets("YDWK3").Range("D6").Value = Sheets("Sheet2").Range("AK3").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AK10").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H11:I11").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AK11").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H12:I12").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AK12").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H13:I13").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AK13").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H14:I14").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AK14").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H15:I15").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AK15").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H16:I16").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AK16").Value = Sheets("YDWK3").Range("C2").Value

' Test Column(Starts) for FIN in OZ Sector
    Sheets("Sheet2").Activate
    Range("AK8").FormulaR1C1 = "=MAX(R[2]C12:R[493]C12)"
    Range("L10").FormulaR1C1 = _
        "=IF(OR(AND(R6C[25]>R7C[25],RC[25]>=R7C[25],RC[25]<=R6C[25]),AND(R6C[25]<R7C[25],OR(RC[25]>=R7C[25],RC[25]<=R6C[25]))),RC[25],-1)"
    
    Application.Calculation = xlCalculationManual
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
.Range("L10:L10").AutoFill Destination:=.Range("L10:L" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("AK8").Value = Range("AK8").Value

If Range("AK8") < 0 Then
    Range("AK10:AK500").ClearContents
End If
End If

'Column AL
If Range("AL1") <> "" Then
Sheets("YDWK3").Activate
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H10:I10").Value
    Sheets("YDWK3").Range("C6").Value = Sheets("Sheet2").Range("AL2").Value
    Sheets("YDWK3").Range("D6").Value = Sheets("Sheet2").Range("AL3").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AL10").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H11:I11").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AL11").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H12:I12").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AL12").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H13:I13").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AL13").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H14:I14").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AL14").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H15:I15").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AL15").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H16:I16").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AL16").Value = Sheets("YDWK3").Range("C2").Value

' Test Column(Starts) for FIN in OZ Sector
    Sheets("Sheet2").Activate
    Range("AL8").FormulaR1C1 = "=MAX(R[2]C12:R[493]C12)"
    Range("L10").FormulaR1C1 = _
        "=IF(OR(AND(R6C[26]>R7C[26],RC[26]>=R7C[26],RC[26]<=R6C[26]),AND(R6C[26]<R7C[26],OR(RC[26]>=R7C[26],RC[26]<=R6C[26]))),RC[26],-1)"
    
    Application.Calculation = xlCalculationManual
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
.Range("L10:L10").AutoFill Destination:=.Range("L10:L" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("AL8").Value = Range("AL8").Value

If Range("AL8") < 0 Then
    Range("AL10:AL500").ClearContents
End If
End If

'Column AM
If Range("AM1") <> "" Then
Sheets("YDWK3").Activate
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H10:I10").Value
    Sheets("YDWK3").Range("C6").Value = Sheets("Sheet2").Range("AM2").Value
    Sheets("YDWK3").Range("D6").Value = Sheets("Sheet2").Range("AM3").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AM10").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H11:I11").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AM11").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H12:I12").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AM12").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H13:I13").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AM13").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H14:I14").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AM14").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H15:I15").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AM15").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H16:I16").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AM16").Value = Sheets("YDWK3").Range("C2").Value

' Test Column(Starts) for FIN in OZ Sector
    Sheets("Sheet2").Activate
    Range("AM8").FormulaR1C1 = "=MAX(R[2]C12:R[493]C12)"
    Range("L10").FormulaR1C1 = _
        "=IF(OR(AND(R6C[27]>R7C[27],RC[27]>=R7C[27],RC[27]<=R6C[27]),AND(R6C[27]<R7C[27],OR(RC[27]>=R7C[27],RC[27]<=R6C[27]))),RC[27],-1)"
    
    Application.Calculation = xlCalculationManual
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
.Range("L10:L10").AutoFill Destination:=.Range("L10:L" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("AM8").Value = Range("AM8").Value

If Range("AM8") < 0 Then
    Range("AM10:AM500").ClearContents
End If
End If

'Column AN
If Range("AN1") <> "" Then
Sheets("YDWK3").Activate
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H10:I10").Value
    Sheets("YDWK3").Range("C6").Value = Sheets("Sheet2").Range("AN2").Value
    Sheets("YDWK3").Range("D6").Value = Sheets("Sheet2").Range("AN3").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AN10").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H11:I11").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AN11").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H12:I12").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AN12").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H13:I13").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AN13").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H14:I14").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AN14").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H15:I15").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AN15").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H16:I16").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AN16").Value = Sheets("YDWK3").Range("C2").Value

' Test Column(Starts) for FIN in OZ Sector
    Sheets("Sheet2").Activate
    Range("AN8").FormulaR1C1 = "=MAX(R[2]C12:R[493]C12)"
    Range("L10").FormulaR1C1 = _
        "=IF(OR(AND(R6C[28]>R7C[28],RC[28]>=R7C[28],RC[28]<=R6C[28]),AND(R6C[28]<R7C[28],OR(RC[28]>=R7C[28],RC[28]<=R6C[28]))),RC[28],-1)"
    
    Application.Calculation = xlCalculationManual
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
.Range("L10:L10").AutoFill Destination:=.Range("L10:L" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("AN8").Value = Range("AN8").Value

If Range("AN8") < 0 Then
    Range("AN10:AN500").ClearContents
End If
End If

'Column AO
If Range("AO1") <> "" Then
Sheets("YDWK3").Activate
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H10:I10").Value
    Sheets("YDWK3").Range("C6").Value = Sheets("Sheet2").Range("AO2").Value
    Sheets("YDWK3").Range("D6").Value = Sheets("Sheet2").Range("AO3").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AO10").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H11:I11").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AO11").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H12:I12").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AO12").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H13:I13").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AO13").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H14:I14").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AO14").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H15:I15").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AO15").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H16:I16").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AO16").Value = Sheets("YDWK3").Range("C2").Value

' Test Column(Starts) for FIN in OZ Sector
    Sheets("Sheet2").Activate
    Range("AO8").FormulaR1C1 = "=MAX(R[2]C12:R[493]C12)"
    Range("L10").FormulaR1C1 = _
        "=IF(OR(AND(R6C[29]>R7C[29],RC[29]>=R7C[29],RC[29]<=R6C[29]),AND(R6C[29]<R7C[29],OR(RC[29]>=R7C[29],RC[29]<=R6C[29]))),RC[29],-1)"
    
    Application.Calculation = xlCalculationManual
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
.Range("L10:L10").AutoFill Destination:=.Range("L10:L" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("AO8").Value = Range("AO8").Value

If Range("AO8") < 0 Then
    Range("AO10:AO500").ClearContents
End If
End If

'Column AP
If Range("AP1") <> "" Then
Sheets("YDWK3").Activate
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H10:I10").Value
    Sheets("YDWK3").Range("C6").Value = Sheets("Sheet2").Range("AP2").Value
    Sheets("YDWK3").Range("D6").Value = Sheets("Sheet2").Range("AP3").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AP10").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H11:I11").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AP11").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H12:I12").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AP12").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H13:I13").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AP13").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H14:I14").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AP14").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H15:I15").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AP15").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H16:I16").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AP16").Value = Sheets("YDWK3").Range("C2").Value

' Test Column(Starts) for FIN in OZ Sector
    Sheets("Sheet2").Activate
    Range("AP8").FormulaR1C1 = "=MAX(R[2]C12:R[493]C12)"
    Range("L10").FormulaR1C1 = _
        "=IF(OR(AND(R6C[30]>R7C[30],RC[30]>=R7C[30],RC[30]<=R6C[30]),AND(R6C[30]<R7C[30],OR(RC[30]>=R7C[30],RC[30]<=R6C[30]))),RC[30],-1)"
    
    Application.Calculation = xlCalculationManual
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
.Range("L10:L10").AutoFill Destination:=.Range("L10:L" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("AP8").Value = Range("AP8").Value

If Range("AP8") < 0 Then
    Range("AP10:AP500").ClearContents
End If
End If

'Column AQ
If Range("AQ1") <> "" Then
Sheets("YDWK3").Activate
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H10:I10").Value
    Sheets("YDWK3").Range("C6").Value = Sheets("Sheet2").Range("AQ2").Value
    Sheets("YDWK3").Range("D6").Value = Sheets("Sheet2").Range("AQ3").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AQ10").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H11:I11").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AQ11").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H12:I12").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AQ12").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H13:I13").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AQ13").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H14:I14").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AQ14").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H15:I15").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AQ15").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H16:I16").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AQ16").Value = Sheets("YDWK3").Range("C2").Value

' Test Column(Starts) for FIN in OZ Sector
    Sheets("Sheet2").Activate
    Range("AQ8").FormulaR1C1 = "=MAX(R[2]C12:R[493]C12)"
    Range("L10").FormulaR1C1 = _
        "=IF(OR(AND(R6C[31]>R7C[31],RC[31]>=R7C[31],RC[31]<=R6C[31]),AND(R6C[31]<R7C[31],OR(RC[31]>=R7C[31],RC[31]<=R6C[31]))),RC[31],-1)"
    
    Application.Calculation = xlCalculationManual
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
.Range("L10:L10").AutoFill Destination:=.Range("L10:L" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("AQ8").Value = Range("AQ8").Value

If Range("AQ8") < 0 Then
    Range("AQ10:AQ500").ClearContents
End If
End If

If Range("AR1") <> "" Then Application.Run "F.xlsm!Sect4"

End Sub

Sub Sect4()
'
' OZSect4 Macro
'
'Column AR
Application.ScreenUpdating = False
If Range("AR1") <> "" Then
Sheets("YDWK3").Activate
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H10:I10").Value
    Sheets("YDWK3").Range("C6").Value = Sheets("Sheet2").Range("AR2").Value
    Sheets("YDWK3").Range("D6").Value = Sheets("Sheet2").Range("AR3").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AR10").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H11:I11").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AR11").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H12:I12").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AR12").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H13:I13").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AR13").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H14:I14").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AR14").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H15:I15").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AR15").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H16:I16").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AR16").Value = Sheets("YDWK3").Range("C2").Value

' Test Column(Starts) for FIN in OZ Sector
    Sheets("Sheet2").Activate
    Range("AR8").FormulaR1C1 = "=MAX(R[2]C12:R[493]C12)"
    Range("L10").FormulaR1C1 = _
        "=IF(OR(AND(R6C[32]>R7C[32],RC[32]>=R7C[32],RC[32]<=R6C[32]),AND(R6C[32]<R7C[32],OR(RC[32]>=R7C[32],RC[32]<=R6C[32]))),RC[32],-1)"
    
    Application.Calculation = xlCalculationManual
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
.Range("L10:L10").AutoFill Destination:=.Range("L10:L" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("AR8").Value = Range("AR8").Value

If Range("AR8") < 0 Then
    Range("AR10:AR500").ClearContents
End If
End If

'Column AS
If Range("AS1") <> "" Then
    Sheets("YDWK3").Activate
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H10:I10").Value
    Sheets("YDWK3").Range("C6").Value = Sheets("Sheet2").Range("AS2").Value
    Sheets("YDWK3").Range("D6").Value = Sheets("Sheet2").Range("AS3").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AS10").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H11:I11").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AS11").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H12:I12").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AS12").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H13:I13").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AS13").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H14:I14").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AS14").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H15:I15").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AS15").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H16:I16").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AS16").Value = Sheets("YDWK3").Range("C2").Value

' Test Column(Starts) for FIN in OZ Sector
    Sheets("Sheet2").Activate
    Range("AS8").FormulaR1C1 = "=MAX(R[2]C12:R[493]C12)"
    Range("L10").FormulaR1C1 = _
        "=IF(OR(AND(R6C[33]>R7C[33],RC[33]>=R7C[33],RC[33]<=R6C[33]),AND(R6C[33]<R7C[33],OR(RC[33]>=R7C[33],RC[33]<=R6C[33]))),RC[33],-1)"

    Application.Calculation = xlCalculationManual
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
.Range("L10:L10").AutoFill Destination:=.Range("L10:L" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("AS8").Value = Range("AS8").Value

If Range("AS8") < 0 Then
    Range("AS10:AS500").ClearContents
End If
End If

'Column AT
If Range("AT1") <> "" Then
    Sheets("YDWK3").Activate
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H10:I10").Value
    Sheets("YDWK3").Range("C6").Value = Sheets("Sheet2").Range("AT2").Value
    Sheets("YDWK3").Range("D6").Value = Sheets("Sheet2").Range("AT3").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AT10").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H11:I11").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AT11").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H12:I12").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AT12").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H13:I13").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AT13").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H14:I14").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AT14").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H15:I15").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AT15").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H16:I16").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AT16").Value = Sheets("YDWK3").Range("C2").Value

' Test Column(Starts) for FIN in OZ Sector
    Sheets("Sheet2").Activate
    Range("AT8").FormulaR1C1 = "=MAX(R[2]C12:R[493]C12)"
    Range("L10").FormulaR1C1 = _
        "=IF(OR(AND(R6C[34]>R7C[34],RC[34]>=R7C[34],RC[34]<=R6C[34]),AND(R6C[34]<R7C[34],OR(RC[34]>=R7C[34],RC[34]<=R6C[34]))),RC[34],-1)"

    Application.Calculation = xlCalculationManual
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
.Range("L10:L10").AutoFill Destination:=.Range("L10:L" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("AT8").Value = Range("AT8").Value

If Range("AT8") < 0 Then
    Range("AT10:AT500").ClearContents
End If
End If

'Column AU
If Range("AU1") <> "" Then
Sheets("YDWK3").Activate
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H10:I10").Value
    Sheets("YDWK3").Range("C6").Value = Sheets("Sheet2").Range("AU2").Value
    Sheets("YDWK3").Range("D6").Value = Sheets("Sheet2").Range("AU3").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AU10").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H11:I11").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AU11").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H12:I12").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AU12").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H13:I13").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AU13").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H14:I14").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AU14").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H15:I15").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AU15").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H16:I16").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AU16").Value = Sheets("YDWK3").Range("C2").Value

' Test Column(Starts) for FIN in OZ Sector
    Sheets("Sheet2").Activate
    Range("AU8").FormulaR1C1 = "=MAX(R[2]C12:R[493]C12)"
    Range("L10").FormulaR1C1 = _
        "=IF(OR(AND(R6C[35]>R7C[35],RC[35]>=R7C[35],RC[35]<=R6C[35]),AND(R6C[35]<R7C[35],OR(RC[35]>=R7C[35],RC[35]<=R6C[35]))),RC[35],-1)"
    
    Application.Calculation = xlCalculationManual
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
.Range("L10:L10").AutoFill Destination:=.Range("L10:L" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("AU8").Value = Range("AU8").Value

If Range("AU8") < 0 Then
    Range("AU10:AU500").ClearContents
End If
End If

'Column AV
If Range("AV1") <> "" Then
Sheets("YDWK3").Activate
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H10:I10").Value
    Sheets("YDWK3").Range("C6").Value = Sheets("Sheet2").Range("AV2").Value
    Sheets("YDWK3").Range("D6").Value = Sheets("Sheet2").Range("AV3").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AV10").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H11:I11").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AV11").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H12:I12").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AV12").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H13:I13").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AV13").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H14:I14").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AV14").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H15:I15").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AV15").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H16:I16").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AV16").Value = Sheets("YDWK3").Range("C2").Value

' Test Column(Starts) for FIN in OZ Sector
    Sheets("Sheet2").Activate
    Range("AV8").FormulaR1C1 = "=MAX(R[2]C12:R[493]C12)"
    Range("L10").FormulaR1C1 = _
        "=IF(OR(AND(R6C[36]>R7C[36],RC[36]>=R7C[36],RC[36]<=R6C[36]),AND(R6C[36]<R7C[36],OR(RC[36]>=R7C[36],RC[36]<=R6C[36]))),RC[36],-1)"
    
    Application.Calculation = xlCalculationManual
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
.Range("L10:L10").AutoFill Destination:=.Range("L10:L" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("AV8").Value = Range("AV8").Value

If Range("AV8") < 0 Then
    Range("AV10:AV500").ClearContents
End If
End If

'Column AW
If Range("AW1") <> "" Then
    Sheets("YDWK3").Activate
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H10:I10").Value
    Sheets("YDWK3").Range("C6").Value = Sheets("Sheet2").Range("AW2").Value
    Sheets("YDWK3").Range("D6").Value = Sheets("Sheet2").Range("AW3").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AW10").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H11:I11").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AW11").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H12:I12").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AW12").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H13:I13").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AW13").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H14:I14").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AW14").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H15:I15").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AW15").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H16:I16").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AW16").Value = Sheets("YDWK3").Range("C2").Value

' Test Column(Starts) for FIN in OZ Sector
    Sheets("Sheet2").Activate
    Range("AW8").FormulaR1C1 = "=MAX(R[2]C12:R[493]C12)"
    Range("L10").FormulaR1C1 = _
        "=IF(OR(AND(R6C[37]>R7C[37],RC[37]>=R7C[37],RC[37]<=R6C[37]),AND(R6C[37]<R7C[37],OR(RC[37]>=R7C[37],RC[37]<=R6C[37]))),RC[37],-1)"
    
    Application.Calculation = xlCalculationManual
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
.Range("L10:L10").AutoFill Destination:=.Range("L10:L" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("AW8").Value = Range("AW8").Value

If Range("AW8") < 0 Then
    Range("AW10:AW500").ClearContents
End If
End If

'Column AX
If Range("AX1") <> "" Then
Sheets("YDWK3").Activate
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H10:I10").Value
    Sheets("YDWK3").Range("C6").Value = Sheets("Sheet2").Range("AX2").Value
    Sheets("YDWK3").Range("D6").Value = Sheets("Sheet2").Range("AX3").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AX10").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H11:I11").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AX11").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H12:I12").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AX12").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H13:I13").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AX13").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H14:I14").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AX14").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H15:I15").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AX15").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H16:I16").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AX16").Value = Sheets("YDWK3").Range("C2").Value

' Test Column(Starts) for FIN in OZ Sector
    Sheets("Sheet2").Activate
    Range("AX8").FormulaR1C1 = "=MAX(R[2]C12:R[493]C12)"
    Range("L10").FormulaR1C1 = _
        "=IF(OR(AND(R6C[38]>R7C[38],RC[38]>=R7C[38],RC[38]<=R6C[38]),AND(R6C[38]<R7C[38],OR(RC[38]>=R7C[38],RC[38]<=R6C[38]))),RC[38],-1)"
    
    Application.Calculation = xlCalculationManual
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
.Range("L10:L10").AutoFill Destination:=.Range("L10:L" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("AX8").Value = Range("AX8").Value

If Range("AX8") < 0 Then
    Range("AX10:AX500").ClearContents
End If
End If

'Column AY
If Range("AY1") <> "" Then
Sheets("YDWK3").Activate
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H10:I10").Value
    Sheets("YDWK3").Range("C6").Value = Sheets("Sheet2").Range("AY2").Value
    Sheets("YDWK3").Range("D6").Value = Sheets("Sheet2").Range("AY3").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AY10").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H11:I11").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AY11").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H12:I12").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AY12").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H13:I13").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AY13").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H14:I14").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AY14").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H15:I15").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AY15").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H16:I16").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AY16").Value = Sheets("YDWK3").Range("C2").Value

' Test Column(Starts) for FIN in OZ Sector
    Sheets("Sheet2").Activate
    Range("AY8").FormulaR1C1 = "=MAX(R[2]C12:R[493]C12)"
    Range("L10").FormulaR1C1 = _
        "=IF(OR(AND(R6C[39]>R7C[39],RC[39]>=R7C[39],RC[39]<=R6C[39]),AND(R6C[39]<R7C[39],OR(RC[39]>=R7C[39],RC[39]<=R6C[39]))),RC[39],-1)"
    
    Application.Calculation = xlCalculationManual
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
.Range("L10:L10").AutoFill Destination:=.Range("L10:L" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("AY8").Value = Range("AY8").Value

If Range("AY8") < 0 Then
    Range("AY10:AY500").ClearContents
End If
End If

'Column AZ
If Range("AZ1") <> "" Then
Sheets("YDWK3").Activate
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H10:I10").Value
    Sheets("YDWK3").Range("C6").Value = Sheets("Sheet2").Range("AZ2").Value
    Sheets("YDWK3").Range("D6").Value = Sheets("Sheet2").Range("AZ3").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AZ10").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H11:I11").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AZ11").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H12:I12").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AZ12").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H13:I13").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AZ13").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H14:I14").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AZ14").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H15:I15").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AZ15").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H16:I16").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("AZ16").Value = Sheets("YDWK3").Range("C2").Value
  
' Test Column(Starts) for FIN in OZ Sector
    Sheets("Sheet2").Activate
    Range("AZ8").FormulaR1C1 = "=MAX(R[2]C12:R[493]C12)"
    Range("L10").FormulaR1C1 = _
        "=IF(OR(AND(R6C[40]>R7C[40],RC[40]>=R7C[40],RC[40]<=R6C[40]),AND(R6C[40]<R7C[40],OR(RC[40]>=R7C[40],RC[40]<=R6C[40]))),RC[40],-1)"
    
    Application.Calculation = xlCalculationManual
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
.Range("L10:L10").AutoFill Destination:=.Range("L10:L" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("AZ8").Value = Range("AZ8").Value

If Range("AZ8") < 0 Then
    Range("AZ10:AZ500").ClearContents
End If
End If

'Column BA
If Range("BA1") <> "" Then
Sheets("YDWK3").Activate
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H10:I10").Value
    Sheets("YDWK3").Range("C6").Value = Sheets("Sheet2").Range("BA2").Value
    Sheets("YDWK3").Range("D6").Value = Sheets("Sheet2").Range("BA3").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("BA10").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H11:I11").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("BA11").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H12:I12").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("BA12").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H13:I13").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("BA13").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H14:I14").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("BA14").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H15:I15").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("BA15").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H16:I16").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("BA16").Value = Sheets("YDWK3").Range("C2").Value

' Test Column(Starts) for FIN in OZ Sector
    Sheets("Sheet2").Activate
    Range("BA8").FormulaR1C1 = "=MAX(R[2]C12:R[493]C12)"
    Range("L10").FormulaR1C1 = _
        "=IF(OR(AND(R6C[41]>R7C[41],RC[41]>=R7C[41],RC[41]<=R6C[41]),AND(R6C[41]<R7C[41],OR(RC[41]>=R7C[41],RC[41]<=R6C[41]))),RC[41],-1)"
    
    Application.Calculation = xlCalculationManual
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
.Range("L10:L10").AutoFill Destination:=.Range("L10:L" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("BA8").Value = Range("BA8").Value

If Range("BA8") < 0 Then
    Range("BA10:BA500").ClearContents
End If
End If
If Range("BB1") <> "" Then Application.Run "F.xlsm!Sect5"

End Sub

Sub Sect5()
'
' Sect5 Macro
'
'Column BB
Application.ScreenUpdating = False
If Range("BB1") <> "" Then
Sheets("YDWK3").Activate
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H10:I10").Value
    Sheets("YDWK3").Range("C6").Value = Sheets("Sheet2").Range("BB2").Value
    Sheets("YDWK3").Range("D6").Value = Sheets("Sheet2").Range("BB3").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("BB10").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H11:I11").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("BB11").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H12:I12").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("BB12").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H13:I13").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("BB13").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H14:I14").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("BB14").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H15:I15").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("BB15").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H16:I16").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("BB16").Value = Sheets("YDWK3").Range("C2").Value

' Test Column(Starts) for FIN in OZ Sector
    Sheets("Sheet2").Activate
    Range("BB8").FormulaR1C1 = "=MAX(R[2]C12:R[493]C12)"
    Range("L10").FormulaR1C1 = _
        "=IF(OR(AND(R6C[42]>R7C[42],RC[42]>=R7C[42],RC[42]<=R6C[42]),AND(R6C[42]<R7C[42],OR(RC[42]>=R7C[42],RC[42]<=R6C[42]))),RC[42],-1)"
    
    Application.Calculation = xlCalculationManual
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
.Range("L10:L10").AutoFill Destination:=.Range("L10:L" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("BB8").Value = Range("BB8").Value

If Range("BB8") < 0 Then
    Range("BB10:BB500").ClearContents
End If
End If

'Column BC
If Range("BC1") <> "" Then
    Sheets("YDWK3").Activate
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H10:I10").Value
    Sheets("YDWK3").Range("C6").Value = Sheets("Sheet2").Range("BC2").Value
    Sheets("YDWK3").Range("D6").Value = Sheets("Sheet2").Range("BC3").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("BC10").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H11:I11").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("BC11").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H12:I12").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("BC12").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H13:I13").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("BC13").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H14:I14").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("BC14").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H15:I15").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("BC15").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H16:I16").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("BC16").Value = Sheets("YDWK3").Range("C2").Value

' Test Column(Starts) for FIN in OZ Sector
    Sheets("Sheet2").Activate
    Range("BC8").FormulaR1C1 = "=MAX(R[2]C12:R[493]C12)"
    Range("L10").FormulaR1C1 = _
        "=IF(OR(AND(R6C[43]>R7C[43],RC[43]>=R7C[43],RC[43]<=R6C[43]),AND(R6C[43]<R7C[43],OR(RC[43]>=R7C[43],RC[43]<=R6C[43]))),RC[43],-1)"

    Application.Calculation = xlCalculationManual
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
.Range("L10:L10").AutoFill Destination:=.Range("L10:L" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("BC8").Value = Range("BC8").Value

If Range("BC8") < 0 Then
    Range("BC10:BC500").ClearContents
End If
End If

'Column BD
If Range("BD1") <> "" Then
    Sheets("YDWK3").Activate
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H10:I10").Value
    Sheets("YDWK3").Range("C6").Value = Sheets("Sheet2").Range("BD2").Value
    Sheets("YDWK3").Range("D6").Value = Sheets("Sheet2").Range("BD3").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("BD10").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H11:I11").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("BD11").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H12:I12").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("BD12").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H13:I13").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("BD13").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H14:I14").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("BD14").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H15:I15").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("BD15").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H16:I16").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("BD16").Value = Sheets("YDWK3").Range("C2").Value

' Test Column(Starts) for FIN in OZ Sector
    Sheets("Sheet2").Activate
    Range("BD8").FormulaR1C1 = "=MAX(R[2]C12:R[493]C12)"
    Range("L10").FormulaR1C1 = _
        "=IF(OR(AND(R6C[44]>R7C[44],RC[44]>=R7C[44],RC[44]<=R6C[44]),AND(R6C[44]<R7C[44],OR(RC[44]>=R7C[44],RC[44]<=R6C[44]))),RC[44],-1)"
    
    Application.Calculation = xlCalculationManual
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
.Range("L10:L10").AutoFill Destination:=.Range("L10:L" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    
    Range("BD8").Value = Range("BD8").Value

If Range("BD8") < 0 Then
    Range("BD10:BD500").ClearContents
End If
End If

'Column BE
If Range("BE1") <> "" Then
Sheets("YDWK3").Activate
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H10:I10").Value
    Sheets("YDWK3").Range("C6").Value = Sheets("Sheet2").Range("BE2").Value
    Sheets("YDWK3").Range("D6").Value = Sheets("Sheet2").Range("BE3").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("BE10").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H11:I11").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("BE11").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H12:I12").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("BE12").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H13:I13").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("BE13").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H14:I14").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("BE14").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H15:I15").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("BE15").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H16:I16").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("BE16").Value = Sheets("YDWK3").Range("C2").Value

' Test Column(Starts) for FIN in OZ Sector
    Sheets("Sheet2").Activate
    Range("BE8").FormulaR1C1 = "=MAX(R[2]C12:R[493]C12)"
    Range("L10").FormulaR1C1 = _
        "=IF(OR(AND(R6C[45]>R7C[45],RC[45]>=R7C[45],RC[45]<=R6C[45]),AND(R6C[45]<R7C[45],OR(RC[45]>=R7C[45],RC[45]<=R6C[45]))),RC[45],-1)"
    
    Application.Calculation = xlCalculationManual
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
.Range("L10:L10").AutoFill Destination:=.Range("L10:L" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("BE8").Value = Range("BE8").Value

If Range("BE8") < 0 Then
    Range("BE10:BE500").ClearContents
End If
End If

'Column BF
If Range("BF1") <> "" Then
Sheets("YDWK3").Activate
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H10:I10").Value
    Sheets("YDWK3").Range("C6").Value = Sheets("Sheet2").Range("BF2").Value
    Sheets("YDWK3").Range("D6").Value = Sheets("Sheet2").Range("BF3").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("BF10").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H11:I11").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("BF11").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H12:I12").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("BF12").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H13:I13").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("BF13").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H14:I14").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("BF14").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H15:I15").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("BF15").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H16:I16").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    Sheets("Sheet2").Range("BF16").Value = Sheets("YDWK3").Range("C2").Value

' Test Column(Starts) for FIN in OZ Sector
    Sheets("Sheet2").Activate
    Range("BF8").FormulaR1C1 = "=MAX(R[2]C12:R[493]C12)"
    Range("L10").FormulaR1C1 = _
        "=IF(OR(AND(R6C[46]>R7C[46],RC[46]>=R7C[46],RC[46]<=R6C[46]),AND(R6C[46]<R7C[46],OR(RC[46]>=R7C[46],RC[46]<=R6C[46]))),RC[46],-1)"
    
    Application.Calculation = xlCalculationManual
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
.Range("L10:L10").AutoFill Destination:=.Range("L10:L" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("BF8").Value = Range("BF8").Value

If Range("BF8") < 0 Then
    Range("BF10:BF500").ClearContents
End If
End If

'Column BG
If Range("BG1") <> "" Then
Sheets("YDWK3").Activate
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H10:I10").Value
    Sheets("YDWK3").Range("C6").Value = Sheets("Sheet2").Range("BG2").Value
    Sheets("YDWK3").Range("D6").Value = Sheets("Sheet2").Range("BG3").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("BG10").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H11:I11").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("BG11").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H12:I12").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("BG12").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H13:I13").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("BG13").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H14:I14").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("BG14").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H15:I15").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("BG15").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H16:I16").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("BG16").Value = Sheets("YDWK3").Range("C2").Value

' Test Column(Starts) for FIN in OZ Sector
    Sheets("Sheet2").Activate
    Range("BG8").FormulaR1C1 = "=MAX(R[2]C12:R[493]C12)"
    Range("L10").FormulaR1C1 = _
        "=IF(OR(AND(R6C[47]>R7C[47],RC[47]>=R7C[47],RC[47]<=R6C[47]),AND(R6C[47]<R7C[47],OR(RC[47]>=R7C[47],RC[47]<=R6C[47]))),RC[47],-1)"
    
    Application.Calculation = xlCalculationManual
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
.Range("L10:L10").AutoFill Destination:=.Range("L10:L" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("BG8").Value = Range("BG8").Value

If Range("BG8") < 0 Then
    Range("BG10:BG500").ClearContents
End If
End If

'Column BH
If Range("BH1") <> "" Then
Sheets("YDWK3").Activate
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H10:I10").Value
    Sheets("YDWK3").Range("C6").Value = Sheets("Sheet2").Range("BH2").Value
    Sheets("YDWK3").Range("D6").Value = Sheets("Sheet2").Range("BH3").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("BH10").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H11:I11").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("BH11").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H12:I12").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("BH12").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H13:I13").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("BH13").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H14:I14").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("BH14").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H15:I15").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("BH15").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H16:I16").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("BH16").Value = Sheets("YDWK3").Range("C2").Value

' Test Column(Starts) for FIN in OZ Sector
    Sheets("Sheet2").Activate
    Range("BH8").FormulaR1C1 = "=MAX(R[2]C12:R[493]C12)"
    Range("L10").FormulaR1C1 = _
        "=IF(OR(AND(R6C[48]>R7C[48],RC[48]>=R7C[48],RC[48]<=R6C[48]),AND(R6C[48]<R7C[48],OR(RC[48]>=R7C[48],RC[48]<=R6C[48]))),RC[48],-1)"
    
    Application.Calculation = xlCalculationManual
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
.Range("L10:L10").AutoFill Destination:=.Range("L10:L" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("BH8").Value = Range("BH8").Value

If Range("BH8") < 0 Then
    Range("BH10:BH500").ClearContents
End If
End If

'Column BI
If Range("BI1") <> "" Then
Sheets("YDWK3").Activate
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H10:I10").Value
    Sheets("YDWK3").Range("C6").Value = Sheets("Sheet2").Range("BI2").Value
    Sheets("YDWK3").Range("D6").Value = Sheets("Sheet2").Range("BI3").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("BI10").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H11:I11").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("BI11").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H12:I12").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("BI12").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H13:I13").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("BI13").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H14:I14").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("BI14").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H15:I15").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    Sheets("Sheet2").Range("BI15").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H16:I16").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("BI16").Value = Sheets("YDWK3").Range("C2").Value

' Test Column(Starts) for FIN in OZ Sector
    Sheets("Sheet2").Activate
    Range("BI8").FormulaR1C1 = "=MAX(R[2]C12:R[493]C12)"
    Range("L10").FormulaR1C1 = _
        "=IF(OR(AND(R6C[49]>R7C[49],RC[49]>=R7C[49],RC[49]<=R6C[49]),AND(R6C[49]<R7C[49],OR(RC[49]>=R7C[49],RC[49]<=R6C[49]))),RC[49],-1)"
    
    Application.Calculation = xlCalculationManual
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
.Range("L10:L10").AutoFill Destination:=.Range("L10:L" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("BI8").Value = Range("BI8").Value

If Range("BI8") < 0 Then
    Range("BI10:BI500").ClearContents
End If
End If

'Column BJ
If Range("BJ1") <> "" Then
Sheets("YDWK3").Activate
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H10:I10").Value
    Sheets("YDWK3").Range("C6").Value = Sheets("Sheet2").Range("BJ2").Value
    Sheets("YDWK3").Range("D6").Value = Sheets("Sheet2").Range("BJ3").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("BJ10").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H11:I11").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("BJ11").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H12:I12").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("BJ12").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H13:I13").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("BJ13").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H14:I14").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("BJ14").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H15:I15").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("BJ15").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H16:I16").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("BJ16").Value = Sheets("YDWK3").Range("C2").Value

' Test Column(Starts) for FIN in OZ Sector
    Sheets("Sheet2").Activate
    Range("BJ8").FormulaR1C1 = "=MAX(R[2]C12:R[493]C12)"
    Range("L10").FormulaR1C1 = _
        "=IF(OR(AND(R6C[50]>R7C[50],RC[50]>=R7C[50],RC[50]<=R6C[50]),AND(R6C[50]<R7C[50],OR(RC[50]>=R7C[50],RC[50]<=R6C[50]))),RC[50],-1)"
    
    Application.Calculation = xlCalculationManual
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
.Range("L10:L10").AutoFill Destination:=.Range("L10:L" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("BJ8").Value = Range("BJ8").Value

If Range("BJ8") < 0 Then
    Range("BJ10:BJ500").ClearContents
End If
End If

'Column BK
If Range("BK1") <> "" Then
Sheets("YDWK3").Activate
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H10:I10").Value
    Sheets("YDWK3").Range("C6").Value = Sheets("Sheet2").Range("BK2").Value
    Sheets("YDWK3").Range("D6").Value = Sheets("Sheet2").Range("BK3").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("BK10").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H11:I11").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("BK11").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H12:I12").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("BK12").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H13:I13").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("BK13").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H14:I14").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("BK14").Value = Sheets("YDWK3").Range("C2").Value

    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H15:I15").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("BK15").Value = Sheets("YDWK3").Range("C2").Value
    
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet2").Range("H16:I16").Value
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    ActiveSheet.Calculate
    Sheets("Sheet2").Range("BK16").Value = Sheets("YDWK3").Range("C2").Value

' Test Column(Starts) for FIN in OZ Sector
    Sheets("Sheet2").Activate
    Range("BK8").FormulaR1C1 = "=MAX(R[2]C12:R[493]C12)"
    Range("L10").FormulaR1C1 = _
        "=IF(OR(AND(R6C[51]>R7C[51],RC[51]>=R7C[51],RC[51]<=R6C[51]),AND(R6C[51]<R7C[51],OR(RC[51]>=R7C[51],RC[51]<=R6C[51]))),RC[51],-1)"
    
    Application.Calculation = xlCalculationManual
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
.Range("L10:L10").AutoFill Destination:=.Range("L10:L" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("BK8").Value = Range("BK8").Value

If Range("BK8") < 0 Then
    Range("BK10:BK500").ClearContents
End If
End If

End Sub

Sub Last()
'
' Last Macro
'
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Range("N102:BK102").FormulaR1C1 = _
        "=IF(R8C>0,6371*ACOS(SIN(R2C)*SIN(R2C2)+COS(R2C)*COS(R2C2)*COS(R2C3-R3C)),"""")"
    
    Range("N101").FormulaR1C1 = "=MAX(R[1]C:R[1]C[49])"
    Range("N101").Value = Range("N101").Value
    
  If Range("N101") = 0 Then
    Columns("N:BK").Clear
  
  ElseIf Range("N101") <> 0 Then
    Range("N103:BK103").FormulaR1C1 = "=IF(R102C=R101C14,R1C,"""")"
    Range("N104:BK104").FormulaR1C1 = "=IF(R[-1]C<>"""",R2C,"""")"
    Range("N105:BK105").FormulaR1C1 = "=IF(R[-1]C<>"""",R3C,"""")"
    Range("N106:BK106").FormulaR1C1 = "=IF(R[-1]C<>"""",R4C,"""")"
    Range("N107:BK107").FormulaR1C1 = "=IF(R[-1]C<>"""",R5C,"""")"
    
    Range("L103:L107").FormulaR1C1 = "=MAX(RC[2]:RC[51])"
    Range("L103:L107").Value = Range("L103:L107").Value
    
    Range("L103:L107").Copy
    Range("H1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    
    Range("G105:K129").Value = Range("G10:K34").Value
     
    Range("N103:BK103").FormulaR1C1 = "=IF(R102C=R101C14,R6C,"""")"
    Range("N104:BK104").FormulaR1C1 = "=IF(R[-1]C<>"""",R7C,"""")"
    Range("N105:BK105").FormulaR1C1 = "=IF(R[-1]C<>"""",R10C,"""")"
    Range("N106:BK106").FormulaR1C1 = "=IF(R[-1]C<>"""",R11C,"""")"
    Range("N107:BK107").FormulaR1C1 = "=IF(R[-1]C<>"""",R12C,"""")"
    Range("N108:BK108").FormulaR1C1 = "=IF(R[-1]C<>"""",R13C,"""")"
    Range("N109:BK109").FormulaR1C1 = "=IF(R[-1]C<>"""",R14C,"""")"
    Range("N110:BK110").FormulaR1C1 = "=IF(R[-1]C<>"""",R15C,"""")"
    Range("N111:BK111").FormulaR1C1 = "=IF(R[-1]C<>"""",R16C,"""")"
    Range("N103:BK111").Value = Range("N103:BK111").Value
    Range("L103:L109").Clear
    Range("L103:L111").FormulaR1C1 = "=MAX(RC[2]:RC[51])"
     
    Range("M105").FormulaR1C1 = _
        "=IF(OR(AND(R103C[-1]>R104C[-1],RC[-1]>=R104C[-1],RC[-1]<=R103C[-1]),AND(R103C[-1]<R104C[-1],OR(RC[-1]<=R103C[-1],RC[-1]>=R104C[-1]))),RC[-6],"""")"

    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
.Range("M105:M105").AutoFill Destination:=.Range("M105:M" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    
    Range("G101").FormulaR1C1 = "=MAX(R[4]C[6]:R[28]C[6])"
    Range("G101").Value = Range("G101").Value
    
    Range("H105").FormulaR1C1 = "=IF(RC[-1]=R101C[-1],R[-95]C,"""")"
    Range("I105").FormulaR1C1 = "=IF(RC[-1]<>"""",R[-95]C,"""")"
    Range("J105").FormulaR1C1 = "=IF(RC[-1]<>"""",R[-95]C,"""")"
    Range("K105").FormulaR1C1 = "=IF(RC[-1]<>"""",R[-95]C,"""")"
    
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
.Range("H105:K105").AutoFill Destination:=.Range("H105:K" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    
    Range("H101:K101").FormulaR1C1 = "=MAX(R[4]C:R[24]C)"
    Range("H101:K101").Value = Range("H101:K101").Value
    Range("H3:L3").Value = Range("G101:K101").Value
    
    Range("H2:L2").Value = Range("A2:E2").Value
    Range("M2").FormulaR1C1 = _
        "=IF(R[-1]C[-2]-R[1]C[-2]<=R9C4,2*6371*ACOS(SIN(R[-1]C[-4])*SIN(RC[-4])+COS(R[-1]C[-4])*COS(RC[-4])*COS(RC[-3]-R[-1]C[-3])),2*6371*ACOS(SIN(R[-1]C[-4])*SIN(RC[-4])+COS(R[-1]C[-4])*COS(RC[-4])*COS(RC[-3]-R[-1]C[-3]))-((R[-1]C[-2]-R[1]C[-2]-R9C4)*0.1))"
    Range("M2").Value = Range("M2").Value
    Range("M3").Value = "SECTOR"

    Range("A5,C6:E7").Clear
    Range("G10:M10009").Clear
    Columns("N:BK").Clear
    
  End If
End Sub
Sub NearestSF()
'
' NearestSF Macro; amended 8/27/17 for Franke et similar
'
Application.ScreenUpdating = False

    Range("H6:L9").Clear
    Range("A5").Value = "0.00138888888888889"
    
    Range("M10").FormulaR1C1 = "=IF(AND(RC[-12]>=R3C8-R5C1,RC[-12]<=R3C8+R5C1),RC[-12],"""")"
    Range("N10:Q10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-12],"""")"
    Range("R10").FormulaR1C1 = "=IF(OR(RC[-1]="""",RC[-3]=R1C10),"""",6371*ACOS(SIN(RC[-4])*SIN(R1C9)+COS(RC[-4])*COS(R1C9)*COS(R1C10-RC[-3])))"
    'Copy Ref A Value Sort
    Application.Calculation = xlCalculationManual
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("M10:R10").AutoFill Destination:=.Range("M10:R" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("M10:R10009").Value = Range("M10:R10009").Value
    Range("M10:R10009").Sort Key1:=Range("M10"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
     
    Range("R8").FormulaR1C1 = "=MIN(R[2]C:R[501]C)"
    Range("S8:W8").FormulaR1C1 = "=MAX(R[2]C:R[501]C)"
    
    Range("S10").FormulaR1C1 = "=IF(RC[-1]=R8C18,RC[-6],"""")"
    Range("T10:W10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-6],"""")"
    'Copy Ref M
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "M").End(xlUp).Row
.Range("S10:W10").AutoFill Destination:=.Range("S10:W" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("R8:W8").Value = Range("R8:W8").Value
    
    Range("R10:R509").Clear
    
    Range("R10").FormulaR1C1 = "=IF(AND(RC[-17]>=R1C8-R5C1,RC[-17]<=R1C8+R5C1),RC[-17],"""")"
    Range("S10:V10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-17],"""")"
    Range("W10").FormulaR1C1 = "=IF(OR(RC[-1]="""",RC[-2]=R8C21),"""",6371*ACOS(SIN(RC[-4])*SIN(R8C20)+COS(RC[-4])*COS(R8C20)*COS(R8C21-RC[-3])))"
    'Copy Ref A Value Sort
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("R10:W10").AutoFill Destination:=.Range("R10:W" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("R10:W10009").Value = Range("R10:W10009").Value
    Range("R10:W10009").Sort Key1:=Range("R10"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
     
    Range("R6").FormulaR1C1 = "=MIN(R[4]C[5]:R[496]C[5])"
    Range("S6:W6").FormulaR1C1 = "=MAX(R[4]C[5]:R[496]C[5])"
    
    Range("X10").FormulaR1C1 = "=IF(RC[-1]=R6C18,RC[-6],"""")"
    Range("Y10:AB10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-6],"""")"
    'Copy Ref R
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "R").End(xlUp).Row
.Range("X10:AB10").AutoFill Destination:=.Range("X10:AB" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("S6:W6").Value = Range("S6:W6").Value
    Range("S7:W7").Value = Range("H2:L2").Value
    Range("H6:L8").Value = Range("S6:W8").Value
    
    Range("S6:W8,M10:W10009").Clear
    Columns("X:AB").Clear
    Range("M7").FormulaR1C1 = "=2*6371*ACOS(SIN(R[-1]C[-4])*SIN(RC[-4])+COS(R[-1]C[-4])*COS(RC[-4])*COS(RC[-3]-R[-1]C[-3]))"
    ActiveSheet.Calculate
    Range("A5").Value = "0.000208333333333333"
    
    Range("M10").FormulaR1C1 = "=IF(AND(RC[-12]>=R8C8-R5C1,RC[-12]<=R8C8+R5C1),RC[-12],"""")"
    Range("N10:Q10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-12],"""")"
    Range("R10").FormulaR1C1 = "=IF(RC[-1]<>"""",6371*ACOS(SIN(RC[-4])*SIN(R7C9)+COS(RC[-4])*COS(R7C9)*COS(R7C10-RC[-3])),"""")"
    'Copy Ref A Value Sort
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("M10:R10").AutoFill Destination:=.Range("M10:R" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("M10:R10009").Value = Range("M10:R10009").Value
    Range("M10:R10009").Sort Key1:=Range("M10"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
     
    Range("M10:R509").Copy
    Range("N1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
        
    Range("M10:R10009").Clear
    
    Range("G10").FormulaR1C1 = "=IF(AND(RC[-6]>=R6C8-R5C1,RC[-6]<=R6C8+R5C1),RC[-6],"""")"
    Range("H10:K10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-6],"""")"
    Range("L10").FormulaR1C1 = "=IF(RC[-1]<>"""",6371*ACOS(SIN(RC[-4])*SIN(R7C9)+COS(RC[-4])*COS(R7C9)*COS(R7C10-RC[-3])),"""")"
    'Copy Ref A Value Sort
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:L10").AutoFill Destination:=.Range("G10:L" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("G10:L10009").Value = Range("G10:L10009").Value
    Range("G10:L10009").Sort Key1:=Range("G10"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
     
    'MATRIX
    Range("N10:BK10").FormulaR1C1 = _
        "=IF(OR(R1C="""",RC9=R3C),"""",IF(OR(R6C<RC12,6371*ACOS(SIN(RC8)*SIN(R2C)+COS(RC8)*COS(R2C)*COS(R3C-RC9))>0.5),"""",6371*ACOS(SIN(RC8)*SIN(R2C)+COS(RC8)*COS(R2C)*COS(R3C-RC9))))"
    'Copy Ref L
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "L").End(xlUp).Row
.Range("N10:BK10").AutoFill Destination:=.Range("N10:BK" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("I5").FormulaR1C1 = "=MIN(R[5]C[5]:R[495]C[42])"
    Range("I5").Value = Range("I5").Value
    Range("N10:BK50").Value = Range("N10:BK50").Value
    Range("N7:BK7").FormulaR1C1 = "=IF(MIN(R[3]C:R[493]C)=R5C9,R[-6]C,"""")"
    Range("N7:BK7").Value = Range("N7:BK7").Value
    Range("M10:M50").FormulaR1C1 = "=IF(MIN(RC[1]:RC[50])=R5C9,RC[-6],"""")"
    Range("M10:M50").Value = Range("M10:M50").Value
    
    Range("N10:BK50").Clear
    
    Range("N8:BK11").FormulaR1C1 = "=IF(R[-1]C<>"""",R[-6]C,"""")"
    
    Range("H8").FormulaR1C1 = "=MAX(R[-1]C[6]:R[-1]C[55])"
    Range("I8").FormulaR1C1 = "=MAX(RC[5]:RC[54])"
    Range("J8").FormulaR1C1 = "=MAX(R[1]C[4]:R[1]C[53])"
    Range("K8").FormulaR1C1 = "=MAX(R[2]C[3]:R[2]C[52])"
    Range("L8").FormulaR1C1 = "=MAX(R[3]C[2]:R[3]C[51])"
    Range("H8:L8").Value = Range("H8:L8").Value
    Range("N7:BK11").Clear
    
    Range("N10:Q50").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-6],"""")"
    
    Range("H6:L6").FormulaR1C1 = "=MAX(R[4]C[5]:R[494]C[5])"
    ActiveSheet.Calculate
    Range("H6:L6").Value = Range("H6:L6").Value
    
    Range("M7").FormulaR1C1 = _
        "=2*6371*ACOS(SIN(R[-1]C[-4])*SIN(RC[-4])+COS(R[-1]C[-4])*COS(RC[-4])*COS(RC[-3]-R[-1]C[-3]))"
    ActiveSheet.Calculate
    Range("I5,G10:Q10009,N1:BK6").Clear
    
    'Revised 8/27/17 Franke no change
    If Range("M7") > Range("M2") Then
        Range("A1:F3").Value = Range("H6:M8").Value
        Application.Run "F.xlsm!OZSector"
    ElseIf Range("M7") <= Range("M2") Then
        Range("H6:M8").Clear
        Application.Run "F.xlsm!OZSectorB"
    End If
End Sub
Sub FLineLast()
'
' Tests within 3 minutes of initial S/F pt for Fin Line ONLY
'
    Range("A5").Value = 0.002083333
    'Starts within 3 minutes of H1
    Range("G10").FormulaR1C1 = "=IF(AND(RC[-6]>=R1C8-R5C1,RC[-6]<=R1C8+R5C1),RC[-6],"""")"
    Range("H10:K10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-6],"""")"
    Range("L10").FormulaR1C1 = "=IF(RC[-1]<>"""",6371*ACOS(SIN(RC[-4])*SIN(R2C9)+COS(RC[-4])*COS(R2C9)*COS(R2C10-RC[-3])),"""")"
    'Copy Ref A Value Sort
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:L10").AutoFill Destination:=.Range("G10:L" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("G10:L10009").Value = Range("G10:L10009").Value
    Range("G10:L10009").Sort Key1:=Range("G10"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    
    'Finis
    Range("O10").FormulaR1C1 = "=IF(AND(RC[-14]>=R3C8-R5C1,RC[-14]<=R3C8+R5C1),RC[-14],"""")"
    Range("P10:S10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-14],"""")"
    Range("T10").FormulaR1C1 = "=IF(RC[-1]<>"""",6371*ACOS(SIN(RC[-4])*SIN(R2C9)+COS(RC[-4])*COS(R2C9)*COS(R2C10-RC[-3])),"""")"
    'Copy Ref A Value Sort Copy to N1
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("O10:T10").AutoFill Destination:=.Range("O10:T" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("O10:T10009").Value = Range("O10:T10009").Value
    Range("O10:T10009").Sort Key1:=Range("O10"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    Range("O10:T190").Copy
    Range("N1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    
    Range("O10:T10009").Clear
    
    'Matrix
     Range("N10:GK10").FormulaR1C1 = _
        "=IF(OR(R1C="""",R3C=RC9),"""",IF(OR(R6C<RC12,6371*ACOS(SIN(R2C)*SIN(RC8)+COS(R2C)*COS(RC8)*COS(RC9-R3C))>0.5,R6C<SQRT(RC12^2+0.25),R6C>R6C[1]),"""",IF(RC10-R4C<R9C4,R6C+RC12,R6C+RC12-((RC10-R4C-R9C4)*0.1))))"
    'Range("N10:GK10").FormulaR1C1 = "=IF(OR(R1C="""",RC9=R3C),"""",IF(OR(6371*ACOS(SIN(RC8)*SIN(R2C)+COS(RC8)*COS(R2C)*COS(R3C-RC9))>0.5,R6C[1]<R6C),0,2*RC12))"
    'Copy Ref G
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
.Range("N10:GK10").AutoFill Destination:=.Range("N10:GK" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("N10:GK190").Value = Range("N10:GK190").Value

    Range("M8").FormulaR1C1 = "=MAX(R[2]C[1]:R[352]C[180])"
    Range("M8").Value = Range("M8").Value
    
    If Range("M8") <> 0 Then
    
    Range("M10").FormulaR1C1 = "=IF(MAX(RC[1]:RC[180])=R8C13,RC[-6],"""")"
    'Copy Ref G
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("M10:M10").AutoFill Destination:=.Range("M10:M" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("M10:M10009").Value = Range("M10:M10009").Value
    
    Range("N7:GK7").FormulaR1C1 = "=IF(MAX(R[3]C:R[193]C)=R8C13,R[-6]C,"""")"
    ActiveSheet.Calculate
    Range("N7:GK7").Value = Range("N7:GK7").Value
    Range("H8").FormulaR1C1 = "=MIN(R[-1]C[6]:R[-1]C[185])"
    Range("H8").Value = Range("H8").Value
    
    Range("M8,N10:GK190").Clear
    
    Range("N10:Q190").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-6],"""")"
    ActiveSheet.Calculate
    Range("H6:L6").FormulaR1C1 = "=MAX(R[4]C[5]:R[194]C[5])"
    Range("H6:L6").Value = Range("H6:L6").Value
    
    Range("N8:GK8").FormulaR1C1 = "=IF(R[-1]C=R8C8,R[-6]C,"""")"
    Range("N9:GK11").FormulaR1C1 = "=IF(R[-1]C<>"""",R[-6]C,"""")"
    ActiveSheet.Calculate
    
    Range("I8").FormulaR1C1 = "=MAX(RC[5]:RC[184])"
    Range("J8").FormulaR1C1 = "=MAX(R[1]C[4]:R[1]C[183])"
    Range("K8").FormulaR1C1 = "=MAX(R[2]C[3]:R[2]C[182])"
    Range("L8").FormulaR1C1 = "=MAX(R[3]C[2]:R[3]C[181])"
    Range("I8:L8").Value = Range("I8:L8").Value
    
    Range("H7:L7").Value = Range("H2:L2").Value
    Range("M7").FormulaR1C1 = "=2*6371*ACOS(SIN(R[-1]C[-4])*SIN(RC[-4])+COS(R[-1]C[-4])*COS(RC[-4])*COS(RC[-3]-R[-1]C[-3]))"
    Columns("N:GK").Clear
    Range("G10:L190").Clear
    Range("A1:F3").Value = Range("H6:M8").Value
    Range("H1:M3").Value = Range("H6:M8").Value
    Range("H6:M8").Clear
    Application.Run "F.xlsm!Fline1"
  End If
     
End Sub
Sub ORRref()
'
' On Sheet 3
'
Application.ScreenUpdating = False
    Sheets("Sheet3").Activate
    Sheets("Sheet3").Range("A5").Value = 2.08333333333333E-04
    Sheets("Sheet3").Range("H1:M4").Value = Sheets("Sheet2").Range("H10:M13").Value
    If Range("H4") = "NO FIN LINE" Then
        Range("H4").Clear
    End If
    
    'SORT +/- 18 SECONDS
    Range("G10").FormulaR1C1 = _
        "=IF(OR(AND(RC[-6]>=R1C8-R5C1,RC[-6]<=R1C8+R5C1),AND(RC[-6]>=R2C8-R5C1,RC[-6]<=R2C8+R5C1),AND(RC[-6]>=R3C8-R5C1,RC[-6]<=R3C8+R5C1),AND(R4C8<>"""",RC[-6]>=R4C8-R5C1,RC[-6]<=R4C8+R5C1),AND(R5C8<>"""",RC[-6]>=R5C8-R5C1,RC[-6]<=R5C8+R5C1)),RC[-6],"""")"
    Range("H10:K10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-6],"""")"
    'Copy Ref A Value Sort
    With Worksheets("Sheet3")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("G10:K10").AutoFill Destination:=.Range("G10:K" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("G10:K60009").Value = Range("G10:K60009").Value
    Range("G10:K60009").Sort Key1:=Range("G10"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    'CHECK TP
    Range("M8:R8").FormulaR1C1 = "=MAX(R[2]C:R[192]C)"
    Range("M9").FormulaR1C1 = "=2*R8C"
    Range("M10:M200").FormulaR1C1 = _
        "=IF(RC[-6]="""","""",IF(AND(RC[-6]>R1C8,RC[-6]<R3C8),6371*ACOS(SIN(RC[-5])*SIN(R1C9)+COS(RC[-5])*COS(R1C9)*COS(R1C10-RC[-4])),""""))"
    Range("N10:N200").FormulaR1C1 = "=IF(RC[-1]=R8C13,RC[-7],"""")"
    Range("O10:R200").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-7],"""")"
    ActiveSheet.Calculate
    
    If Range("M9") > Range("M2") Then
        Range("H2:L2").Value = Range("N8:R8").Value
        Range("M2").Value = Range("M9").Value
    End If
    Range("M8:R200").Clear
    
    'CK OZ if Fini by Sector ONLY
If Range("H4") = "" And Range("M3") = "SECTOR" Then
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet3").Range("I1:J1").Value
    Sheets("YDWK3").Range("C6:D6").Value = Sheets("Sheet3").Range("I2:J2").Value
    Sheets("YDWK3").Calculate
    Sheets("Sheet3").Range("J6:J7").Value = Sheets("YDWK3").Range("E2:E3").Value
    Sheets("YDWK3").Range("C5:D6").Value = 0
    Sheets("YDWK3").Activate
    Range("C24").FormulaR1C1 = "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    Sheets("YDWK3").Calculate
    Sheets("Sheet3").Activate
   'Find 7 Fixes - 3 before & 3 after raw Fin
    Range("M10").FormulaR1C1 = "=IF(RC[-6]=R3C8,1,"""")"
    Range("N10").FormulaR1C1 = "=IF(R[3]C[-1]=1,1,"""")"
    Range("O10").FormulaR1C1 = "=IF(R[2]C[-2]=1,1,"""")"
    Range("P10").FormulaR1C1 = "=IF(R[1]C[-3]=1,1,"""")"
    Range("Q10").FormulaR1C1 = "=IF(R[-1]C[-4]=1,1,"""")"
    Range("R10").FormulaR1C1 = "=IF(R[-2]C[-5]=1,1,"""")"
    Range("S10").FormulaR1C1 = "=IF(R[-3]C[-6]=1,1,"""")"
    
    With Worksheets("Sheet3")
    LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
.Range("M10:S10").AutoFill Destination:=.Range("M10:S" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("M10:S209").Value = Range("M10:S209").Value
    Range("T10").FormulaR1C1 = "=IF(OR(RC[-7]=1,RC[-6]=1,RC[-5]=1,RC[-4]=1,RC[-3]=1,RC[-2]=1,RC[-1]=1),RC[-13],"""")"
    Range("U10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-13],"""")"
    Range("V10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-13],"""")"
    Range("W10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-13],"""")"
    Range("X10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-13],"""")"
      
    With Worksheets("Sheet3")
    LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
.Range("T10:X10").AutoFill Destination:=.Range("T10:X" & LastRow), Type:=xlFillDefault
    End With
    ActiveSheet.Calculate
    Range("T10:X209").Value = Range("T10:X209").Value
    Range("M10:S209").Clear
    Range("T10:X209").Sort Key1:=Range("T10"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    
    'Test 3 earlier, 3 later than A3
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet3").Range("U10:V10").Value
    Sheets("YDWK3").Range("C6:D6").Value = Sheets("Sheet3").Range("I1:J1").Value
    Sheets("YDWK3").Calculate
    Sheets("Sheet3").Range("S10").Value = Sheets("YDWK3").Range("C2").Value
    Sheets("YDWK3").Range("C5:D6").Value = 0
    Sheets("YDWK3").Activate
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    Sheets("YDWK3").Calculate
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet3").Range("U11:V11").Value
    Sheets("YDWK3").Range("C6:D6").Value = Sheets("Sheet3").Range("I1:J1").Value
    Sheets("YDWK3").Calculate
    Sheets("Sheet3").Range("S11").Value = Sheets("YDWK3").Range("C2").Value
    Sheets("YDWK3").Range("C5:D6").Value = 0
    Sheets("YDWK3").Activate
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    Sheets("YDWK3").Calculate
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet3").Range("U12:V12").Value
    Sheets("YDWK3").Range("C6:D6").Value = Sheets("Sheet3").Range("I1:J1").Value
    Sheets("YDWK3").Calculate
    Sheets("Sheet3").Range("S12").Value = Sheets("YDWK3").Range("C2").Value
    Sheets("YDWK3").Range("C5:D6").Value = 0
    Sheets("YDWK3").Activate
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    Sheets("YDWK3").Calculate
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet3").Range("U13:V13").Value
    Sheets("YDWK3").Range("C6:D6").Value = Sheets("Sheet3").Range("I1:J1").Value
    Sheets("YDWK3").Calculate
    Sheets("Sheet3").Range("S13").Value = Sheets("YDWK3").Range("C2").Value
    Sheets("YDWK3").Range("C5:D6").Value = 0
    Sheets("YDWK3").Activate
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    Sheets("YDWK3").Calculate
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet3").Range("U14:V14").Value
    Sheets("YDWK3").Range("C6:D6").Value = Sheets("Sheet3").Range("I1:J1").Value
    Sheets("YDWK3").Calculate
    Sheets("Sheet3").Range("S14").Value = Sheets("YDWK3").Range("C2").Value
    Sheets("YDWK3").Range("C5:D6").Value = 0
    Sheets("YDWK3").Activate
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    Sheets("YDWK3").Calculate
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet3").Range("U15:V15").Value
    Sheets("YDWK3").Range("C6:D6").Value = Sheets("Sheet3").Range("I1:J1").Value
    Sheets("YDWK3").Calculate
    Sheets("Sheet3").Range("S15").Value = Sheets("YDWK3").Range("C2").Value
    Sheets("YDWK3").Range("C5:D6").Value = 0
    Sheets("YDWK3").Activate
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    Sheets("YDWK3").Calculate
    Sheets("YDWK3").Range("C5:D5").Value = Sheets("Sheet3").Range("U16:V16").Value
    Sheets("YDWK3").Range("C6:D6").Value = Sheets("Sheet3").Range("I1:J1").Value
    Sheets("YDWK3").Calculate
    Sheets("Sheet3").Range("S16").Value = Sheets("YDWK3").Range("C2").Value
    Sheets("YDWK3").Range("C5:D6").Value = 0
    Sheets("YDWK3").Activate
    Range("C24").FormulaR1C1 = _
        "=R[-15]C+(1-R[-1]C)*R[-16]C*R[-3]C*(ACOS(R[-4]C)+R[-1]C*SIN(ACOS(R[-4]C))*(R[-2]C+R[-1]C*R[-4]C*(-1+2*R[-2]C*R[-2]C)))"
    Sheets("YDWK3").Calculate
    Application.Run "F.xlsm!YDWK3clear"
    
    Sheets("Sheet3").Activate
    Range("M6").FormulaR1C1 = "=MIN(R[4]C:R[10]C)"
    Range("N6:Q6").FormulaR1C1 = "=MAX(R[4]C:R[10]C)"
    Range("M10:M16").FormulaR1C1 = _
        "=IF(OR(6371*ACOS(SIN(RC[8])*SIN(R1C9)+COS(RC[8])*COS(R1C9)*COS(R1C10-RC[9]))>1,RC[6]=""""),"""",IF(OR(AND(R6C10>R7C10,RC[6]>=R7C10,RC[6]<=R6C10),AND(R6C10<R7C10,OR(RC[6]<=R6C10,RC[6]>=R7C10))),RC[7],""""))"
    Range("N10:N16").FormulaR1C1 = "=IF(RC[-1]=R6C13,RC[7],"""")"
    Range("O10:Q16").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[7],"""")"
    ActiveSheet.Calculate
    Range("M6:Q6").Value = Range("M6:Q6").Value
    
    If Range("M6") > 0 Then
        Range("K6").Value = "OK"
        Range("H3:L3").Value = Range("M6:Q6").Value
        Range("M2").FormulaR1C1 = "=2*6371*ACOS(SIN(R[-1]C[-4])*SIN(RC[-4])+COS(R[-1]C[-4])*COS(RC[-4])*COS(RC[-3]-R[-1]C[-3]))"
        Range("N2").FormulaR1C1 = _
        "=IF(OR(AND(RC[-1]>100,R[-1]C[-3]-R[1]C[-3]<=R9C4),AND(RC[-1]<=100,R9C5<>""PR"",R[-1]C[-3]-R[1]C[-3]<=(10*RC[-1])),AND(RC[-1]<=100,R9C5=""PR"",R[-1]C[-3]-R[1]C[-3]<=(10*RC[-1]-100))),RC[-1],IF(AND(RC[-1]>100,R[-1]C[-3]-R[1]C[-3]>R9C4),RC[-1]-((R[-1]C[-3]-R[1]C[-3]-R9C4)*0.1),0))"
        ActiveSheet.Calculate
        If Range("N2") < Range("M2") Then
            Range("M2").Value = Range("N2").Value
            Range("N3").Value = "LOH"
        End If
        Range("M2").Value = Range("M2").Value
        Range("N2").Clear
        Range("M3").Value = "SECTOR"
        Range("H4").Value = "NO FIN LINE"
        
        Sheets("Sheet2").Range("H15:M18").Value = Sheets("Sheet3").Range("H1:N4").Value
        Sheets("Sheet3").Range("A5,H1:N3").Clear
    End If
    
ElseIf Range("I4") = "        FINISH LINE" Then
    Range("M10:M209").FormulaR1C1 = "=IF(RC[-6]="""","""",IF(AND(RC[-6]>=R1C8-R5C1,RC[-6]<=R1C8+R5C1),RC[-6],""""))"
    Range("N10:Q209").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-6],"""")"
    Range("R10:R209").FormulaR1C1 = "=IF(RC[-1]<>"""",6371*ACOS(SIN(RC[-10])*SIN(R2C9)+COS(RC[-10])*COS(R2C9)*COS(R2C10-RC[-9])),"""")"
    Range("S10:S209").FormulaR1C1 = "=IF(RC[-6]="""","""",IF(RC[-1]>100,R9C4,IF(R9C5<>""PR"",20*RC[-1],20*RC[-1]-100)))"
    ActiveSheet.Calculate
    Range("M10:S209").Value = Range("M10:S209").Value
    Range("M10:S209").Sort Key1:=Range("M10"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    
    Range("T10:T209").FormulaR1C1 = "=IF(RC[-13]="""","""",IF(AND(RC[-13]>=R3C8-R5C1,RC[-13]<=R3C8+R5C1),RC[-13],""""))"
    Range("U10:X209").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-13],"""")"
    Range("Y10:Y209").FormulaR1C1 = "=IF(RC[-1]<>"""",6371*ACOS(SIN(RC[-17])*SIN(R2C9)+COS(RC[-17])*COS(R2C9)*COS(R2C10-RC[-16])),"""")"
    ActiveSheet.Calculate
    Range("T10:Y209").Value = Range("T10:Y209").Value
    Range("T10:Y209").Sort Key1:=Range("T10"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    
    Range("T10:Y209").Copy
    Range("T1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    Range("T10:Y209").Clear
 
    'MATRIX
    Range("T10:HL209").FormulaR1C1 = _
        "=IF(OR(RC13="""",R1C="""",RC15=R3C),"""",IF(OR(6371*ACOS(SIN(RC14)*SIN(R2C)+COS(RC14)*COS(R2C)*COS(R3C-RC15))>0.5,R6C<RC18,R6C[1]<RC18,R6C<SQRT(RC18^2+0.25)),"""",IF(RC16-R4C<=RC19,R6C+RC18,IF(AND(RC16-R4C>RC19,R9C5<>""PR""),R6C+RC18-((RC16-R4C-RC19)*0.1),0))))"
    ActiveSheet.Calculate
    Range("T10:HL209").Value = Range("T10:HL209").Value
    Range("O5").FormulaR1C1 = "=MAX(R[5]C[5]:R[195]C[205])"
    ActiveSheet.Calculate
    Range("O5").Value = Range("O5").Value
    Range("T7:HL7").FormulaR1C1 = "=IF(MAX(R[3]C:R[193]C)=R5C15,R1C,"""")"
    ActiveSheet.Calculate
    Range("T7:HL7").Value = Range("T7:HL7").Value
    
    Range("S10:S209").FormulaR1C1 = "=IF(MAX(RC[1]:RC[201])=R5C15,RC[-6],"""")"
    ActiveSheet.Calculate
    Range("S10:S209").Value = Range("S10:S209").Value
    
    Range("T10:HL209").Clear
    Range("T10:W209").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-6],"""")"
    
    Range("N1:R1").FormulaR1C1 = "=MAX(R[9]C[5]:R[199]C[5])"
    ActiveSheet.Calculate
    Range("N1:R1").Value = Range("N1:R1").Value
    Range("N2:R2").Value = Range("H2:L2").Value
    Range("T8:HL11").FormulaR1C1 = "=IF(R[-1]C<>"""",R[-6]C,"""")"
    ActiveSheet.Calculate
    Range("N3").FormulaR1C1 = "=MAX(R[4]C[6]:R[4]C[206])"
    Range("O3").FormulaR1C1 = "=MAX(R[5]C[5]:R[5]C[205])"
    Range("P3").FormulaR1C1 = "=MAX(R[6]C[4]:R[6]C[204])"
    Range("Q3").FormulaR1C1 = "=MAX(R[7]C[3]:R[7]C[203])"
    Range("R3").FormulaR1C1 = "=MAX(R[8]C[2]:R[8]C[202])"
    ActiveSheet.Calculate
    Range("N3:R3").Value = Range("N3:R3").Value
    
    Range("T1:HL209").Clear
    Range("S2").FormulaR1C1 = "=2*6371*ACOS(SIN(R[-1]C[-4])*SIN(RC[-4])+COS(R[-1]C[-4])*COS(RC[-4])*COS(RC[-3]-R[-1]C[-3]))"
    Range("T2").FormulaR1C1 = _
        "=IF(OR(AND(RC[-1]>100,R[-1]C[-3]-R[1]C[-3]<R9C4),AND(RC[-1]<=100,R9C5<>""PR"",R[-1]C[-3]-R[1]C[-3]<=10*RC[-1]),AND(RC[-1]<=100,R9C5=""PR"",R[-1]C[-3]-R[1]C[-3]<=10*RC[-1]-100)),RC[-1],IF(R9C5<>""PR"",RC[-1]-((R[-1]C[-3]-R[1]C[-3]-R9C4)*0.1),0))"
    Range("S3").FormulaR1C1 = "=IF(R[-1]C[1]<R[-1]C,""LOH"","""")"
    ActiveSheet.Calculate
    Range("S3").Value = Range("S3").Value
    Range("S2").Value = Range("T2").Value
    Range("T2,M5:S209").Clear
    
    'Select New FINI + 2 fixes before, 2 fixes after
    Range("M10:M209").FormulaR1C1 = "=IF(RC[-6]=R3C14,1,"""")"
    Range("N10:N209").FormulaR1C1 = "=IF(R[2]C[-1]=1,1,"""")"
    Range("O10:O209").FormulaR1C1 = "=IF(R[1]C[-2]=1,1,"""")"
    Range("P10:P209").FormulaR1C1 = "=IF(R[-1]C[-3]=1,1,"""")"
    Range("Q10:Q209").FormulaR1C1 = "=IF(R[-2]C[-4]=1,1,"""")"
    Range("R10:R209").FormulaR1C1 = "=IF(AND(RC[-11]<>"""",OR(RC[-5]=1,RC[-4]=1,RC[-3]=1,RC[-2]=1,RC[-1]=1)),RC[-11],"""")"
    Range("S10:V209").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-11],"""")"
    ActiveSheet.Calculate
    Range("R10:V209").Value = Range("R10:V209").Value
    Range("M10:Q209").Clear
    Range("R10:V209").Sort Key1:=Range("R10"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    
    Sheets("YDWK3").Activate
    Sheets("YDWK3").Range("M22:Q22").Value = Sheets("Sheet3").Range("N2:R2").Value
    Sheets("YDWK3").Range("M23:Q23").Value = Sheets("Sheet3").Range("N1:R1").Value
    Sheets("YDWK3").Range("M25:Q29").Value = Sheets("Sheet3").Range("R10:V14").Value

    Application.Run "F.xlsm!FINlines"

    Sheets("Sheet3").Activate
    Sheets("Sheet3").Range("N4").Value = Sheets("YDWK3").Range("M17").Value
    Sheets("Sheet3").Range("O4").FormulaR1C1 = "=IF(RC[-1]<>""NO FIN LINE"",""        FINISH LINE"","""")"
    ActiveSheet.Calculate
    Range("O4").Value = Range("O4").Value
    Sheets("Sheet3").Range("Q4:R4").Value = Sheets("YDWK3").Range("N17:O17").Value
    Range("A5,R10:V14").Clear

Application.Run "F.xlsm!YDWK3clear"

    Sheets("Sheet3").Activate

    If Range("N4") = "NO FIN LINE" Then
        'CK ORIGINAL Start; FINI + 2 fixes before, 2 fixes after
        Range("M10:M209").FormulaR1C1 = "=IF(RC[-6]=R3C8,1,"""")"
        Range("N10:N209").FormulaR1C1 = "=IF(R[2]C[-1]=1,1,"""")"
        Range("O10:O209").FormulaR1C1 = "=IF(R[1]C[-2]=1,1,"""")"
        Range("P10:P209").FormulaR1C1 = "=IF(R[-1]C[-3]=1,1,"""")"
        Range("Q10:Q209").FormulaR1C1 = "=IF(R[-2]C[-4]=1,1,"""")"
        Range("R10:R209").FormulaR1C1 = "=IF(OR(RC[-5]=1,RC[-4]=1,RC[-3]=1,RC[-2]=1,RC[-1]=1),RC[-11],"""")"
        Range("S10:V209").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-11],"""")"
        ActiveSheet.Calculate
        Range("R10:V209").Value = Range("R10:V209").Value
        Range("M10:Q209").Clear
        Range("R10:V209").Sort Key1:=Range("R10"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    
        Sheets("YDWK3").Activate
        Sheets("YDWK3").Range("M22:Q22").Value = Sheets("Sheet3").Range("H2:L2").Value
        Sheets("YDWK3").Range("M23:Q23").Value = Sheets("Sheet3").Range("H1:L1").Value
        Sheets("YDWK3").Range("M25:Q29").Value = Sheets("Sheet3").Range("R10:V14").Value

    Application.Run "F.xlsm!FINlines"

        Sheets("Sheet3").Activate
        Sheets("Sheet3").Range("H5").Value = Sheets("YDWK3").Range("M17").Value
        Sheets("Sheet3").Range("I5").FormulaR1C1 = "=IF(RC[-1]<>""NO FIN LINE"",""        FINISH LINE"","""")"
        Range("I5").Value = Range("I5").Value
        Sheets("Sheet3").Range("K5:L5").Value = Sheets("YDWK3").Range("N17:O17").Value
        Range("A5,N1:S4,R10:V14").Clear

    Application.Run "F.xlsm!YDWK3clear"
    End If
    
    Sheets("Sheet3").Activate
    If Range("S2") > Range("M2") And Range("I4") = Range("O4") And Range("M3") = Range("S3") Then
        Sheets("Sheet2").Range("H15:M18").Value = Sheets("Sheet3").Range("N1:S4").Value
    ElseIf Range("I5") = "        FINISH LINE" Then
        Sheets("Sheet2").Range("H15:M17").Value = Sheets("Sheet3").Range("H1:M3").Value
        Sheets("Sheet2").Range("H18:M18").Value = Sheets("Sheet3").Range("H5:M5").Value
    ElseIf Range("H5") = "NO FIN LINE" And Range("I4") = "        FINISH LINE" Then
        Sheets("Sheet2").Range("H15:M18").Value = Sheets("Sheet3").Range("H1:M4").Value
    End If
End If
Sheets("Sheet3").Columns("G:V").Clear

End Sub

Sub OZSectorB()
'
' OZSectorB Macro
'
    'P1 differentiates between Sect / Line based on flight date 10/1/15
    Range("P1").FormulaR1C1 = _
        "=IF(AND(R2C6=MAX(R2C6,R2C13,R7C13),R4C1<>""NO FIN LINE""),""A1"",IF(OR(AND(R2C13=MAX(R2C13,R7C13),R3C13=""SECTOR"",R4C8=""NO FIN LINE""),AND(R2C13=MAX(R2C13,R7C13),R4C8<>""NO FIN LINE"")),""H1"",IF(AND(R6C8<>"""",R9C8<>""NO FIN LINE""),""H6"")))"
    ActiveSheet.Calculate
    Range("P1").Value = Range("P1").Value
    
    If Range("P1") = "H1" Then
        Range("H10:M13").Value = Range("H1:M4").Value
    ElseIf Range("P1") = "A1" Then
        Range("H10:M13").Value = Range("A1:F4").Value
    ElseIf Range("P1") = "H6" Then
        Range("H10:M13").Value = Range("H6:M9").Value
    End If
    
  If Range("A9") = "REF" Or Range("B9") = "REF2" Then
    Application.Run "F.xlsm!ORRref"
    Sheets("Sheet2").Activate
  End If

If Range("H15") <> "" Then
    If Range("M17") = "SECTOR" And Range("H18") = "NO FIN LINE" Then
        Range("I17:J17").Value = Range("I15:J15").Value
        Range("I18").Value = "   FINISH SECTOR"
        Range("M17").Clear
        Range("K18:L18").Value = Range("K17:L17").Value
        Range("H18").Value = Range("H17").Value
    ElseIf Range("I18") = "        FINISH LINE" Then
        Range("H17").Value = Range("H18").Value
        Range("I17:J17").Value = Range("I15:J15").Value
        Range("K17:L17").Value = Range("K18:L18").Value
        Range("M17").Clear
    End If

    Sheets("TASKS").Activate
    Sheets("TASKS").Range("C14:E17").Value = Sheets("Sheet2").Range("H15:J18").Value
    Range("F14:G16").FormulaR1C1 = "=DEGREES(RC[-2])"
    Sheets("TASKS").Range("H14:J17").Value = Sheets("Sheet2").Range("K15:M18").Value
    Range("F17").FormulaR1C1 = "=RC[-2]"
    Range("F14:G17").Value = Range("F14:G17").Value
     
ElseIf Range("H15") = "" Then
    If Range("M12") = "SECTOR" And Range("H13") = "NO FIN LINE" Then
        Range("I12:J12").Value = Range("I10:J10").Value
        Range("I13").Value = "   FINISH SECTOR"
        Range("M12").Clear
        Range("K13:L13").Value = Range("K12:L12").Value
        Range("H13").Value = Range("H12").Value
    ElseIf Range("I13") = "        FINISH LINE" Then
        Range("H12").Value = Range("H13").Value
        Range("I12:J12").Value = Range("I10:J10").Value
        Range("K12:L12").Value = Range("K13:L13").Value
        Range("M12").Clear
    End If
    Sheets("TASKS").Activate
    Sheets("TASKS").Range("C14:E17").Value = Sheets("Sheet2").Range("H10:J13").Value
    Range("F14:G16").FormulaR1C1 = "=DEGREES(RC[-2])"
    Sheets("TASKS").Range("H14:J17").Value = Sheets("Sheet2").Range("K10:M13").Value
    Range("F17").FormulaR1C1 = "=RC[-2]"
    Range("F14:G17").Value = Range("F14:G17").Value
End If

Sheets("Sheet2").Activate
Range("A1:F7").Clear
Columns("F:P").Clear
Application.Run "F.xlsm!ORDS"

End Sub
