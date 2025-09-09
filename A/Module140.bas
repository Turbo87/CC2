Option Explicit
#If Win64 Then
    Dim LastRow As LongPtr
#Else
    Dim LastRow As Long
#End If
    
Sub PreB()
'
' JLR 2/8/2015  Amended to have edited fixes @ A:G for TPOrder; all post-Rel fixes @ I:M for iteration
' JLR 10/28/2017 Amended to get S/F Line candidates w/in .5 km of S/F points
'
Sheets("Sheet2").Select
    Range("A1").Value = Sheets("PRS").Range("D4").Value
    Range("A2").FormulaR1C1 = "=IF(PRS!R6C7<>0,MIN(PRS!R6C7,PRS!R10C7),PRS!R10C7)"
    Range("A2").Value = Range("A2").Value
    Range("A4").Value = Sheets("PRS").Range("E11").Value
    ActiveSheet.Range("B1:C60000").Value = Sheets("BR").Range("J1:K60000").Value
    Range("D1").FormulaR1C1 = "=IF(AND(RC[-2]>=R1C1,RC[-2]<=R2C1),1,"""")"
    Range("E1").FormulaR1C1 = "=IF(RC[-1]=1,RC[-3],"""")"
    Range("F1").FormulaR1C1 = "=IF(RC[-1]="""","""",RC[-3])"
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "B").End(xlUp).Row
.Range("D1:F1").AutoFill Destination:=.Range("D1:F" & LastRow), Type:=xlFillDefault
    End With
    Range("A3").FormulaR1C1 = "=SUM(C[3])"
    Range("A1:F60000").Value = Range("A1:F60000").Value
    Range("D1").FormulaR1C1 = "1"
    Range("D2").FormulaR1C1 = "=R[-1]C+1"
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "B").End(xlUp).Row
    .Range("D2").AutoFill Destination:=.Range("D2:D" & LastRow), Type:=xlFillDefault
    End With
    Range("D1:D60000").Value = Range("D1:D60000").Value
    Range("E1:F60000").Sort Key1:=Range("E1"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
   
    Columns("F:F").TextToColumns Destination:=Range("F1"), DataType:=xlFixedWidth, _
        OtherChar:="E", FieldInfo:=Array(Array(0, 9), Array(6, 1), Array(8, 1), Array(13, _
        1), Array(14, 1), Array(17, 1), Array(22, 1), Array(23, 9), Array(24, 1), Array(29, 1), _
        Array(34, 9)), TrailingMinusNumbers:=True
   'For "X" and PRs w/zeroes
    Range("A5").FormulaR1C1 = "=IF(AND(PRS!R10C1=""PR"",PRS!R3C1=""X""),0,SUM(R1C12:R5000C12))"
    Range("A5").Value = Range("A5").Value
    If Range("A5") = 0 Then
        Range("L1:L60000").Value = Range("M1:M60000").Value
    End If
    
    'For iterative stuff eg STD, ST/FIN Fixes
    Range("AB1:AB60000").Value = Range("E1:E60000").Value
    Range("AE1:AF60000").Value = Range("L1:M60000").Value
    Range("AC1").FormulaR1C1 = _
        "=IF(AND(RC[-1]<>"""",RC[-21]=""N""),RADIANS(RC[-23]+((RC[-22]/1000)/60)),IF(AND(RC[-1]<>"""",RC[-21]=""S""),-1*RADIANS(RC[-23]+((RC[-22]/1000)/60)),""""))"
    Range("AD1").FormulaR1C1 = _
        "=IF(AND(RC[-1]<>"""",RC[-19]=""E""),RADIANS(RC[-21]+((RC[-20]/1000)/60)),IF(AND(RC[-1]<>"""",RC[-19]=""W""),-1*RADIANS(RC[-21]+((RC[-20]/1000)/60)),""""))"
    'Copy Ref E
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "E").End(xlUp).Row
.Range("AC1:AD1").AutoFill Destination:=.Range("AC1:AD" & LastRow), Type:=xlFillDefault
    End With
    Range("AB1:AF60000").Value = Range("AB1:AF60000").Value
      
    Application.Calculation = xlCalculationManual
    Range("N1").FormulaR1C1 = "=IF(RC[-6]=""N"",RADIANS(RC[-8]+(RC[-7]*0.001/60)),RADIANS(-1*(RC[-8])+-1*(RC[-7]*0.001/60)))"
    Range("O1").FormulaR1C1 = "=IF(RC[-4]=""E"",RADIANS(RC[-6]+(RC[-5]*0.001/60)),RADIANS(-1*(RC[-6])+-1*(RC[-5]*0.001/60)))"
    Range("P1").FormulaR1C1 = _
        "=IF(OR(AND(R3C1<10000,SUM(PRS!R20C9:PRS!R28C9)=0),AND(RC[-2]=PRS!R20C11,RC[-1]=PRS!R20C12),AND(RC[-2]=PRS!R22C11,RC[-1]=PRS!R22C12),AND(RC[-2]=PRS!R24C11,RC[-1]=PRS!R24C12),AND(RC[-2]=PRS!R26C11,RC[-1]=PRS!R26C12),AND(RC[-2]=PRS!R28C11,RC[-1]=PRS!R28C12)),1,IF(OR(6371*(ACOS(SIN(RC[-2])*SIN(PRS!R20C11)+COS(RC[-2])*COS(PRS!R20C11)*COS(RC[-1]-PRS!R20C12)))<40,6371*(ACOS(SIN(RC[-2])*SIN(PRS!R22C11)+COS(RC[-2])*COS(PRS!R22C11)*COS(RC[-1]-PRS!R22C12)))<40,6371*(ACOS(SIN(RC[-2])*SIN(PRS!R24C11)+COS(RC[-2])*COS(PRS!R24C11)*COS(RC[-1]-PRS!R24C12)))<40,6371*(ACOS(SIN(RC[-2])*SIN(PRS!R26C11)+COS(RC[-2])*COS(PRS!R26C11)*COS(RC[-1]-PRS!R26C12)))<40,6371*(ACOS(SIN(RC[-2])*SIN(PRS!R28C11)+COS(RC[-2])*COS(PRS!R28C11)*COS(RC[-1]-PRS!R28C12)))<40),1,""""))"
    Range("Q1").FormulaR1C1 = _
    "=IF(AND(RC[-13]>=R3C1/2,OR(6371*(ACOS(SIN(RC[-3])*SIN(PRS!R22C11)+COS(RC[-3])*COS(PRS!R22C11)*COS(RC[-2]-PRS!R22C12)))>6371*(ACOS(SIN(PRS!R28C11)*SIN(PRS!R22C11)+COS(PRS!R28C11)*COS(PRS!R22C11)*COS(PRS!R28C12-PRS!R22C12))),6371*(ACOS(SIN(RC[-3])*SIN(PRS!R24C11)+COS(RC[-3])*COS(PRS!R24C11)*COS(RC[-2]-PRS!R24C12)))>6371*(ACOS(SIN(PRS!R28C11)*SIN(PRS!R24C11)+COS(PRS!R28C11)*COS(PRS!R24C11)*COS(PRS!R28C12-PRS!R24C12))),6371*(ACOS(SIN(RC[-3])*SIN(PRS!R26C11)+COS(RC[-3])*COS(PRS!R26C11)*COS(RC[-2]-PRS!R26C12)))>6371*(ACOS(SIN(PRS!R28C11)*SIN(PRS!R26C11)+COS(PRS!R28C11)*COS(PRS!R26C11)*COS(PRS!R28C12-PRS!R26C12))))),1,"""")"
    Range("S1").FormulaR1C1 = _
        "=IF(OR(6371*(ACOS(SIN(RC[-5])*SIN(PRS!R20C11)+COS(RC[-5])*COS(PRS!R20C11)*COS(RC[-4]-PRS!R20C12)))<0.5,6371*(ACOS(SIN(RC[-5])*SIN(PRS!R28C11)+COS(RC[-5])*COS(PRS!R28C11)*COS(RC[-4]-PRS!R28C12)))<0.5,AND(RC[-5]=PRS!R20C11,RC[-4]=PRS!R20C12),AND(RC[-5]=PRS!R22C11,RC[-4]=PRS!R22C12),AND(RC[-5]=PRS!R24C11,RC[-4]=PRS!R24C12),AND(RC[-5]=PRS!R26C11,RC[-4]=PRS!R26C12),AND(RC[-5]=PRS!R28C11,RC[-4]=PRS!R28C12)),1,"""")"
    Range("T1").FormulaR1C1 = _
        "=IF(AND(R3C1<10000,OR(RC[-4]=1,RC[-3]=1,RC[-2]=1,RC[-1]=1)),RC[-15],IF(AND(R3C1>=10000,OR(RC[-1]=1,AND(RC[-2]=1,OR(RC[-4]=1,RC[-3]=1)))),RC[-15],""""))"
    Range("U1").FormulaR1C1 = _
        "=IF(RC[-1]="""","""",IF(RC[-13]=""N"",RC[-15]+(RC[-14]*0.001/60),-1*RC[-15]+(-1*(RC[-14]*0.001/60))))"
    Range("V1").FormulaR1C1 = "=IF(RC[-1]="""","""",RC[-8])"
    Range("W1").FormulaR1C1 = _
        "=IF(RC[-1]="""","""",IF(RC[-12]=""E"",RC[-14]+(RC[-13]*0.001/60),-1*RC[-14]+(-1*(RC[-13]*0.001/60))))"
    Range("X1").FormulaR1C1 = "=IF(RC[-1]="""","""",RC[-9])"
    Range("Y1").FormulaR1C1 = "=IF(RC[-1]="""","""",RC[-13])"
    Range("Z1").FormulaR1C1 = "=IF(RC[-1]="""","""",RC[-13])"
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "E").End(xlUp).Row
.Range("N1:Q1").AutoFill Destination:=.Range("N1:Q" & LastRow), Type:=xlFillDefault
.Range("S1:Z1").AutoFill Destination:=Range("S1:Z" & LastRow), Type:=xlFillDefault
    End With
    Range("R1:R4").FormulaR1C1 = "1"
    Range("R6").FormulaR1C1 = _
        "=IF(OR(AND(R3C1>=40000,R[-5]C=1,SUM(R[-4]C:R[-1]C)=0),AND(R3C1>=30000,R3C1<40000,R[-4]C=1,SUM(R[-3]C:R[-1]C)=0),AND(R3C1>=20000,R3C1<30000,R[-3]C=1,SUM(R[-2]C,R[-1]C)=0),AND(R3C1>10000,R3C1<20000,R[-1]C="""")),1,"""")"
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "E").End(xlUp).Row
.Range("R6").AutoFill Destination:=.Range("R6:R" & LastRow), Type:=xlFillDefault
    End With
    
    Worksheets("Sheet2").Calculate
    Range("N1:Z60000").Value = Range("N1:Z60000").Value
    Range("T1:Z60000").Sort Key1:=Range("T1"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    
    Columns("A:S").Clear
    Range("A1:G10000").Value = Range("T1:Z10000").Value
    Columns("T:Z").Clear
    Columns("I:AA").Delete Shift:=xlToLeft
    
Application.Calculation = xlCalculationAutomatic
End Sub