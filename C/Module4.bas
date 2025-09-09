' VBA Module: Flight Performance Analysis and Optimization
' Purpose: Performs detailed flight performance analysis including GPS altitude tracking,
' start/finish line calculations, and optimization computations for competitive soaring.
' Handles complex trigonometric calculations for flight path analysis.

Option Explicit
#If VBA7 And Win64 Then
    Dim LastRow As LongPtr
#Else
    Dim LastRow As Long
#End If

Sub POfixes()
'
'1/15 - 1/30/15 initial testing// 3/30/15 return to rel/Start as only Start Options; 4/21/15 revised for C; 5/1/15 simplified;
'amended 9/6/16 to retain GPS alt; amended 9/7/15 to find GPS alt @ Start/Finish Lines
'amended 10/14/2017 to ck First TP OZ using Start at Release; correction FinFix 11/23/2017 in LoHRedux for LARGE;
'amended 11/26/2007 for MAX via St, Rel, Fini Fix
'
Application.ScreenUpdating = False
Sheets("YDWK2").Activate
Range("E1505").FormulaR1C1 = _
        "=R[-34]C[-2]+(1-R[-2]C)*R[-35]C[-2]*R[-14]C*(ACOS(R[-19]C)+R[-2]C*SIN(ACOS(R[-19]C))*(R[-12]C+R[-2]C*R[-19]C*(-1+2*R[-12]C*R[-12]C)))"
Sheets("YDWK2").Calculate
Application.Calculation = xlCalculationManual
Sheets("Sheet11").Activate

    'Get TPOrder data to Sheet2 G10:K10009
    Sheets("Sheet11").Range("G10:G10009").Value = Sheets("TPOrder").Range("A11:A10010").Value
    Sheets("Sheet11").Range("H10:H10009").Value = Sheets("TPOrder").Range("C11:C10010").Value
    Sheets("Sheet11").Range("I10:K10009").Value = Sheets("TPOrder").Range("E11:G10010").Value
   
    Range("G6:K6").Value = Range("A10:E10").Value
    'Keep A5:A7, B5 format
    Sheets("Sheet11").Range("A5").Value = Sheets("TPOrder").Range("CI1").Value
    Sheets("Sheet11").Range("A6").Value = Sheets("TPOrder").Range("CY1").Value
    Sheets("Sheet11").Range("A7").Value = Sheets("TPOrder").Range("DO1").Value
    
    'Get FIST TP - added 10/29/27 in case < 3 TP
    Range("A4").FormulaR1C1 = "=MAX(R5C1:R7C1)"
    If Range("A4") = 3 Then
        Range("B5").FormulaR1C1 = "=MIN(RC[-1]:R[2]C[-1])"
    ElseIf Range("A4") = 2 Then
        Range("B5").FormulaR1C1 = "=MIN(RC[-1]:R[1]C[-1])"
    ElseIf Range("A4") = 1 Then
        Range("B5").FormulaR1C1 = "R5C1"
    End If
    
    Range("B5").Value = Range("B5").Value
    Range("A4").Clear
    Range("B6").FormulaR1C1 = "=IF(R5C=R7C1,TPOrder!R3C119,IF(R5C=R6C1,TPOrder!R3C103,TPOrder!R3C87))"
    Range("C6").FormulaR1C1 = "=IF(AND(R5C2=R7C1,RC2=""SECTOR""),TPOrder!R4C118,IF(AND(R5C2=R7C1,RC2=""CYLINDER""),TPOrder!R6C119,IF(AND(R5C2=R6C1,RC2=""SECTOR""),TPOrder!R4C102,IF(AND(R5C2=R6C1,RC2=""CYLINDER""),TPOrder!R6C103,IF(AND(R5C2=R5C1,RC2=""SECTOR""),TPOrder!R4C86,IF(AND(R5C2=R5C1,RC2=""CYLINDER""),TPOrder!R6C87,""""))))))"
    Range("B6:C6").Value = Range("B6:C6").Value
    Sheets("Sheet11").Range("E2").Value = Sheets("Worksheet").Range("I34").Value
    Sheets("Sheet11").Range("E3").Value = Sheets("Worksheet").Range("J34").Value
    'T5 variable w/first TP
    If Range("B5") = 3 Then
        Sheets("Sheet11").Range("T5").Value = Sheets("YDWK2").Range("C513").Value
    ElseIf Range("B5") = 2 Then
        Sheets("Sheet11").Range("T5").Value = Sheets("YDWK2").Range("C441").Value
    ElseIf Range("B5") = 1 And Range("E2") > Range("E3") Then
        Sheets("Sheet11").Range("T5").Value = Sheets("YDWK2").Range("C297").Value
    ElseIf Range("B5") = 1 And Range("E3") > Range("E2") Then
        Sheets("Sheet11").Range("T5").Value = Sheets("YDWK2").Range("C369").Value
    End If
    
    Range("M10:Q10").Value = Range("G10:K10").Value
    
    Sheets("Sheet11").Range("S10").Value = Sheets("YDWK2").Range("C441").Value
    Range("T10").FormulaR1C1 = "=IF(RC[-1]-180<0,RC[-1]-180+360,RC[-1]-180)"
    Range("U10").FormulaR1C1 = _
        "=IF(R5C20+RC[-1]<180,(R5C20+RC[-1])/2,IF(AND(R5C20+RC[-1]>180,R5C20+RC[-1]<360,ABS(R5C20-RC[-1])<180),(R5C20+RC[-1])/2,IF(OR(AND(R5C20>270,RC[-1]<90),AND(RC[-1]>270,R5C20<90)),((R5C20+RC[-1])/2)+180,IF(AND(R5C20+RC[-1]>360,R5C20-RC[-1]<180),(R5C20+RC[-1])/2,IF(R5C20+RC[-1]>540,(R5C20+RC[-1])/2,(R5C20+RC[-1])/2-180)))))"
    Range("V10").FormulaR1C1 = "=IF(RC[-1]+45>360,RC[-1]+45-360,RC[-1]+45)"
    Range("W10").FormulaR1C1 = "=IF(RC[-2]-45<0,RC[-2]-45+360,RC[-2]-45)"
    Sheets("Sheet11").Calculate
    Range("T10:W10").Value = Range("T10:W10").Value
    
    Range("M10").Copy
    Range("Y1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
    Range("V10:W10").Copy
    Range("Y2").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    Application.CutCopyMode = False
    
    Range("Y7").FormulaR1C1 = "=MIN(R[3]C:R[10001]C)"
    Range("Y8").FormulaR1C1 = "=MAX(R[2]C:R[10001]C)"
    'X10 et al per TP
    If Range("A7") = Range("B5") Then
        Sheets("Sheet11").Range("X10:X10009").Value = Sheets("TPOrder").Range("DL11:DL10010").Value
    ElseIf Range("A6") = Range("B5") Then
        Sheets("Sheet11").Range("X10:X10009").Value = Sheets("TPOrder").Range("CV11:CV10010").Value
    ElseIf Range("A5") = Range("B5") Then
        Sheets("Sheet11").Range("X10:X10009").Value = Sheets("TPOrder").Range("CF11:CG10010").Value
    End If
    
    Range("Y10").FormulaR1C1 = _
        "=IF(OR(R1C="""",RC7="""",RC[-1]=""""),"""",IF(OR(AND(R2C>R3C,RC24>=R3C,RC24<=R2C),AND(R3C>R2C,OR(RC24>=R3C,RC24<=R2C))),RC7,""""))"
    'Copy Ref G
    Application.Calculation = xlCalculationManual
    With Worksheets("Sheet11")
    LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
.Range("Y10:Y10").AutoFill Destination:=.Range("Y10:Y" & LastRow), Type:=xlFillDefault
    End With
    Sheets("Sheet11").Calculate
    Range("R10").Value = Range("Y7").Value
    Sheets("TPOrder").Unprotect Password:="spike"
    Sheets("TPOrder").Range("DT1:DY1").Value = Sheets("Sheet11").Range("M10:R10").Value
    Sheets("TPOrder").Protect Password:="spike"
    Sheets("Sheet11").Activate
    Range("B5:E7,M10:Y10009").Clear

    'ID lastTP
    Range("B5").FormulaR1C1 = "=MAX(RC[-1]:R[2]C[-1])"
    Range("B5").Value = Range("B5").Value
    Range("B6").FormulaR1C1 = "=IF(R5C=R7C1,TPOrder!R3C119,IF(R5C=R6C1,TPOrder!R3C103,TPOrder!R3C87))"
    Range("C6").FormulaR1C1 = "=IF(AND(R5C2=R7C1,RC2=""SECTOR""),TPOrder!R6C118,IF(AND(R5C2=R7C1,RC2=""CYLINDER""),TPOrder!R6C119,IF(AND(R5C2=R6C1,RC2=""SECTOR""),TPOrder!R6C102,IF(AND(R5C2=R6C1,RC2=""CYLINDER""),TPOrder!R6C103,IF(AND(R5C2=R5C1,RC2=""SECTOR""),TPOrder!R6C86,IF(AND(R5C2=R5C1,RC2=""CYLINDER""),TPOrder!R6C87,""""))))))"
    Range("B6:C6").Value = Range("B6:C6").Value
    
    'TESTS OK 5/1/15!!
    'Ck (E2) Rel to TP1 & (E3) St to TP1 in case only 1 TP achieved
    Sheets("Sheet11").Range("E2").Value = Sheets("Worksheet").Range("I34").Value
    Sheets("Sheet11").Range("E3").Value = Sheets("Worksheet").Range("J34").Value
    'T5 variable w/last TP
    If Range("B5") = 3 Then
        Sheets("Sheet11").Range("T5").Value = Sheets("YDWK2").Range("C513").Value
    ElseIf Range("B5") = 2 Then
        Sheets("Sheet11").Range("T5").Value = Sheets("YDWK2").Range("C441").Value
    ElseIf Range("B5") = 1 And Range("E2") > Range("E3") Then
        Sheets("Sheet11").Range("T5").Value = Sheets("YDWK2").Range("C297").Value
    ElseIf Range("B5") = 1 And Range("E3") > Range("E2") Then
        Sheets("Sheet11").Range("T5").Value = Sheets("YDWK2").Range("C369").Value
    End If
    
    'Check TPOrder FinFix value Added 12/30/2015 for HG (Wild departure from declared task)
    Sheets("Sheet11").Range("H7:I7").Value = Sheets("TPOrder").Range("DR3:DS3").Value
    Sheets("Sheet11").Range("J7").Value = Sheets("TPOrder").Range("DQ8").Value
    
    Range("L8:Q8").FormulaR1C1 = "=MAX(R[2]C:R[59992]C)"
    Range("L10").FormulaR1C1 = "=IF(RC[-11]>R6C3,6371*ACOS(SIN(RC[-10])*SIN(R7C8)+COS(RC[-10])*COS(R7C8)*COS(R7C9-RC[-9])),"""")"
    Range("M10").FormulaR1C1 = "=IF(RC[-1]=R8C12,RC[-12],"""")"
    Range("N10:Q10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-12],"""")"
    'Copy Ref A"
    Application.Calculation = xlCalculationManual
    With Worksheets("Sheet11")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("L10:Q10").AutoFill Destination:=.Range("L10:Q" & LastRow), Type:=xlFillDefault
    End With
    Sheets("Sheet11").Calculate

    If Range("L8") > Range("J7") Then
        Sheets("TPOrder").Unprotect Password:="spike"
        Sheets("TPOrder").Range("DQ8").Value = Sheets("Sheet11").Range("L8").Value
        Sheets("TPOrder").Range("DR8:DS8").Value = Sheets("Sheet11").Range("N8:O8").Value
        Sheets("TPOrder").Range("DT8").Value = Sheets("Sheet11").Range("M8").Value
        Sheets("TPOrder").Range("DU8:DV8").Value = Sheets("Sheet11").Range("P8:Q8").Value
        Sheets("TPOrder").Protect Password:="spike"
        Sheets("TPOrder").Calculate
        Sheets("YDWK2").Calculate
    End If
    
    Sheets("Sheet11").Range("M10").Value = Sheets("TPOrder").Range("DT8").Value
    Sheets("Sheet11").Range("N10").Value = Sheets("TPOrder").Range("DR8").Value
    Sheets("Sheet11").Range("O10").Value = Sheets("TPOrder").Range("DS8").Value
    Sheets("Sheet11").Range("P10").Value = Sheets("TPOrder").Range("DU8").Value
    Sheets("Sheet11").Range("Q10").Value = Sheets("TPOrder").Range("DV8").Value
    Sheets("Sheet11").Range("R10").Value = Sheets("TPOrder").Range("DQ8").Value
    Sheets("Sheet11").Range("S10").Value = Sheets("YDWK2").Range("C1460").Value
    Range("T10").FormulaR1C1 = "=IF(RC[-1]-180<0,RC[-1]-180+360,RC[-1]-180)"
    Range("U10").FormulaR1C1 = _
        "=IF(R5C20+RC[-1]<180,(R5C20+RC[-1])/2,IF(AND(R5C20+RC[-1]>180,R5C20+RC[-1]<360,ABS(R5C20-RC[-1])<180),(R5C20+RC[-1])/2,IF(OR(AND(R5C20>270,RC[-1]<90),AND(RC[-1]>270,R5C20<90)),((R5C20+RC[-1])/2)+180,IF(AND(R5C20+RC[-1]>360,R5C20-RC[-1]<180),(R5C20+RC[-1])/2,IF(R5C20+RC[-1]>540,(R5C20+RC[-1])/2,(R5C20+RC[-1])/2-180)))))"
    Range("V10").FormulaR1C1 = "=IF(RC[-1]+45>360,RC[-1]+45-360,RC[-1]+45)"
    Range("W10").FormulaR1C1 = "=IF(RC[-2]-45<0,RC[-2]-45+360,RC[-2]-45)"
    Sheets("Sheet11").Calculate
    Range("T10:W10").Value = Range("T10:W10").Value
    
    Range("M10").Copy
    Range("Y1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
    Range("V10:W10").Copy
    Range("Y2").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    Application.CutCopyMode = False
    
    Range("Y8").FormulaR1C1 = "=MAX(R[2]C:R[10001]C)"
    'X10 et al per TP
    If Range("A7") = Range("B5") Then
        Sheets("Sheet11").Range("X10:X10009").Value = Sheets("TPOrder").Range("DL11:DL10010").Value
    ElseIf Range("A6") = Range("B5") Then
        Sheets("Sheet11").Range("X10:X10009").Value = Sheets("TPOrder").Range("CV11:CV10010").Value
    ElseIf Range("A5") = Range("B5") Then
        Sheets("Sheet11").Range("X10:X10009").Value = Sheets("TPOrder").Range("CF11:CG10010").Value
    End If
    
    Range("Y10").FormulaR1C1 = _
        "=IF(OR(R1C="""",RC7="""",RC7>R1C,RC[-1]=""""),"""",IF(OR(AND(R2C>R3C,RC24>=R3C,RC24<=R2C),AND(R3C>R2C,OR(RC24>=R3C,RC24<=R2C))),RC7,""""))"
    'Copy Ref G
    Application.Calculation = xlCalculationManual
    With Worksheets("Sheet11")
    LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
.Range("Y10:Y10").AutoFill Destination:=.Range("Y10:Y" & LastRow), Type:=xlFillDefault
    End With
    Sheets("Sheet11").Calculate
    Range("Y8").Value = Range("Y8").Value
    
    'NEW STUFF HERE
    If Range("B6") = "SECTOR" And Range("Y8") = 0 Then
        'We're in trouble!
        Application.Run "C.xlsm!POsolve"
    End If
    
    'Correct for non-std PA
    Range("P11").FormulaR1C1 = "=IF(R[-1]C[-3]<=B!R2C6,R[-1]C+B!R1C6,R[-1]C+B!R3C6)"
    Sheets("Sheet11").Calculate
    Range("P10").Value = Range("P11").Value
    Range("P11").Clear
    ' Check LOH
    Sheets("Sheet11").Range("F2").Value = Sheets("Worksheet").Range("B14").Value
    Sheets("Sheet11").Range("F3").Value = Sheets("Worksheet").Range("D14").Value
    Range("G2").FormulaR1C1 = "=Worksheet!R34C11+Worksheet!R34C12"
    Range("H2").FormulaR1C1 = "=RC[-3]+RC[-1]+R[8]C[10]"
    Range("H3").FormulaR1C1 = "=RC[-3]+R[-1]C[-1]+R[7]C[10]"
    Sheets("Sheet11").Calculate
    Sheets("Sheet11").Range("E4").Value = Sheets("Worksheet").Range("D9").Value
    Range("I2,I3").FormulaR1C1 = "=IF(R4C5=""Feet"",(RC[-3]/3.2808399)-R10C16,RC[-3]-R10C16)"
    Range("J2,J3").FormulaR1C1 = "=IF(RC[-2]= MAX(R2C[-2],R3C[-2]),""MAX"","""")"
    
    Range("K2,K3").FormulaR1C1 = "=IF(RC[-3]>100,R9C4,IF(AND(RC[-3]<100,R9C5<>""PR""),10*RC[-3],(10*RC[-3])-100))"
    Range("L2,L3").FormulaR1C1 = "=IF(RC[-3]<RC[-1],""OK"","""")"
    
    Range("M2").FormulaR1C1 = "=IF(RC[-4]>RC[-2],(RC[-8]+RC[-6]+R[8]C[5])-((RC[-4]-RC[-2])*0.1),RC[-8]+RC[-6]+R[8]C[5])"
    Range("M3").FormulaR1C1 = "=IF(RC[-4]>RC[-2],(RC[-8]+R[-1]C[-6]+R[7]C[5])-((RC[-4]-RC[-2])*0.1),RC[-8]+R[-1]C[-6]+R[7]C[5])"
    Range("M4").FormulaR1C1 = "=MAX(R[-2]C:R[-1]C)"
    Range("M2:M4").Value = Range("M2:M4").Value
    Range("L8:Q8").Value = Range("L8:Q8").Value
    Range("R8").Value = Range("Y8").Value
    Sheets("Sheet11").Calculate
    
    If Range("J2") = "MAX" And Range("I2") > Range("K2") Then
        Application.Run "C.xlsm!LoHREDUX"
    ElseIf Range("J3") = "MAX" And Range("I3") > Range("K3") Then
        Application.Run "C.xlsm!LoHREDUX"
    End If
    
     If Range("L10") < Range("M4") Or Range("L10") = "" Then
        Range("M10:R10").Value = Range("M8:R8").Value
    End If
    
    'Restore altitude as recorded
    Range("P11").FormulaR1C1 = "=IF(R[-1]C[-3]<=B!R2C6,R[-1]C-B!R1C6,R[-1]C-B!R3C6)"
    Range("P10").Value = Range("P11").Value
    'Put data somewhere!
    Sheets("TPOrder").Unprotect Password:="spike"
    Sheets("TPOrder").Range("DT2:DX2").Value = Sheets("Sheet11").Range("M10:Q10").Value
    If Range("B6") <> "CYLINDER" Then
        Sheets("TPOrder").Range("DY2").Value = Sheets("Sheet11").Range("R10").Value
    ElseIf Range("B6") = "CYLINDER" Then
        Sheets("TPOrder").Range("DY2").Value = Sheets("Sheet11").Range("C6").Value
    End If
    Sheets("TPOrder").Protect Password:="spike"
    Range("A1:F8").Clear
    Application.Run "C.xlsm!GPSsfLines"
    Columns("G:Y").Clear
    Sheets("Summary").Calculate
    Sheets("YDWK1").Calculate
    Sheets("Worksheet").Calculate
    Application.Calculation = xlCalculationAutomatic
    
End Sub
Sub POsolve()
'
' Testing on Gai, w/ TP 1 & 2 TPO FF invalidates TP2 Sector; revised 9/6/15 to retain GPS alt
' Amended 9/13/2017 in case fewer than 500 @ AP6
    
    Range("M10").FormulaR1C1 = "=IF(RC[-12]>R6C3,RC[-12],"""")"
    Range("N10:Q10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-12],"""")"
    'Copy Ref A Value Sort
    Application.Calculation = xlCalculationManual
    With Worksheets("Sheet11")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("M10:Q10").AutoFill Destination:=.Range("M10:Q" & LastRow), Type:=xlFillDefault
    End With
'Application.Calculation = xlCalculationAutomatic
Sheets("Sheet11").Calculate

    Range("M10:Q10009").Value = Range("M10:Q10009").Value
    Range("M10:Q10009").Sort Key1:=Range("M10"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    
    'Flattening @ O3
    Range("O3").Value = 3.35281066474748E-03
    'TO point lat @ AJ4, lon @ AL4 TEST CASE VALUES FROM TPO
    If Range("B5") = Range("A5") Then
        Sheets("Sheet11").Range("AJ4").Value = Sheets("TPOrder").Range("CF4").Value
        Sheets("Sheet11").Range("AL4").Value = Sheets("TPOrder").Range("CG4").Value
    ElseIf Range("B5") = Range("A6") Then
        Sheets("Sheet11").Range("AJ4").Value = Sheets("TPOrder").Range("CV4").Value
        Sheets("Sheet11").Range("AL4").Value = Sheets("TPOrder").Range("CW4").Value
    ElseIf Range("B5") = Range("A7") Then
        Sheets("Sheet11").Range("AJ4").Value = Sheets("TPOrder").Range("DL4").Value
        Sheets("Sheet11").Range("AL4").Value = Sheets("TPOrder").Range("DM4").Value
    End If
    
    'Constants
    Range("AA5").FormulaR1C1 = "=(1-R3C15)*TAN(R4C36)"
    Range("AA6").FormulaR1C1 = "=ATAN(R[-1]C)"
    Range("AA7").FormulaR1C1 = "=SIN(R[-1]C)"
    Range("AA8").FormulaR1C1 = "=COS(R[-2]C)"
    
    'Calcs for each candidate
    Range("T10").FormulaR1C1 = "=IF(RC[-3]="""","""",(1-R3C15)*TAN(RC[-6]))"
    Range("U10").FormulaR1C1 = "=IF(RC[-1]="""","""",ATAN(RC[-1]))"
    Range("V10").FormulaR1C1 = "=IF(RC[-1]="""","""",SIN(RC[-1]))"
    Range("W10").FormulaR1C1 = "=IF(RC[-1]="""","""",COS(RC[-2]))"
    'Azimuth TO selected TO POINT
    Range("Y10").FormulaR1C1 = "=R4C38-RC[-10]"
    Range("Z10").FormulaR1C1 = "=RC[-1]+(1-RC[5])*R3C15*RC[3]*(ACOS(RC[2])+RC[5]*SIN(ACOS(RC[2]))*(RC[4]+RC[5]*RC[2]*(-1+2*RC[4]*RC[4])))"
    Range("AA10").FormulaR1C1 = "=(R8C*SIN(RC[-2])*R8C*SIN(RC[-2]))+(RC[-4]*R7C-RC[-5]*R8C*COS(RC[-2]))*(RC[-4]*R7C-RC[-5]*R8C*COS(RC[-2]))"
    Range("AB10").FormulaR1C1 = "=(RC[-6]*R7C[-1])+(RC[-5]*R8C[-1]*COS(RC[-3]))"
    Range("AC10").FormulaR1C1 = "=IF(RC[-2]=0,0,RC[-6]*R8C[-2]*SIN(RC[-4])/SQRT(RC[-2]))"
    Range("AD10").FormulaR1C1 = "=RC[-2]-2*RC[-8]*R7C[-3]/(COS(ASIN(RC[-1]))*COS(ASIN(RC[-1])))"
    Range("AE10").FormulaR1C1 = "=R3C15/16*COS(ASIN(RC[-2]))*COS(ASIN(RC[-2]))*(4+R3C15*(4-3*COS(ASIN(RC[-2]))*COS(ASIN(RC[-2]))))"
    Range("AF10").FormulaR1C1 = "=IF(AND(RC[-18]=R4C[4],RC[-17]=R4C[6]),""samepoint"",""N.A."")"
    Range("AG10").FormulaR1C1 = "=IF(AND(RC[-18]=R4C[5],R4C[3]>RC[-19]),""northsouth"",""N.A."")"
    Range("AH10").FormulaR1C1 = "=IF(AND(RC[-19]=R4C[4],RC[-20]>R4C[2]),""southnorth"",""N.A."")"
    Range("AI10").FormulaR1C1 = "=ATAN2((RC[-12]*R7C[-8]-RC[-13]*R8C[-8]*COS(RC[-9])),R8C[-8]*SIN(RC[-9]))"
    Range("AJ10").FormulaR1C1 = "=IF(RC[-4]=""samepoint"",0,IF(RC[-3]=""northsouth"",0,IF(RC[-2]=""southnorth"",180,IF(RC[-1]<0,RC[-1]*180/PI()+360,RC[-1]*180/PI()))))"
    'Copy Ref M
    Application.Calculation = xlCalculationManual
    With Worksheets("Sheet11")
    LastRow = .Cells(Rows.Count, "M").End(xlUp).Row
.Range("T10:AJ10").AutoFill Destination:=.Range("T10:AJ" & LastRow), Type:=xlFillDefault
    End With
'Application.Calculation = xlCalculationAutomatic
Sheets("Sheet11").Calculate
Range("T10:AJ10009").Value = Range("T10:AJ10009").Value

    Sheets("Sheet11").Range("AK5").Value = Sheets("YDWK2").Range("C657").Value
    'Range("AK5").Value = 205.5650434952
    Range("AK10").FormulaR1C1 = "=IF(AND(RC[-1]>=R5C37-25,RC[-1]<=R5C37+25),RC[-24],"""")"
    Range("AL10:AO10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-24],"""")"
    Range("AP10").FormulaR1C1 = "=IF(RC[-1]<>"""",6371*ACOS(SIN(RC[-28])*SIN(R4C36)+COS(RC[-28])*COS(R4C36)*COS(R4C38-RC[-27])),"""")"
    Range("AQ10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-7],"""")"
    'Copy Ref M
    Application.Calculation = xlCalculationManual
    With Worksheets("Sheet11")
    LastRow = .Cells(Rows.Count, "M").End(xlUp).Row
.Range("AK10:AQ10").AutoFill Destination:=.Range("AK10:AQ" & LastRow), Type:=xlFillDefault
    End With
'Application.Calculation = xlCalculationAutomatic
Sheets("Sheet11").Calculate
Range("AK10:AK10009").Value = Range("AK10:AK10009").Value

    Range("AK7").FormulaR1C1 = "=MIN(R[3]C:R[10002]C)"
    Range("AK8").FormulaR1C1 = "=MAX(R[2]C:R[10001]C)"
    Range("AK9").FormulaR1C1 = "=R[-1]C-R[-2]C"
    Sheets("Sheet11").Calculate
    Range("AK7:AK9").Value = Range("AK7:AK9").Value
    
    If Range("AK7") = 0 And Range("AK8") = 0 Then
         Range("AK10").FormulaR1C1 = "=IF(AND(RC[-1]>=R5C37-20,RC[-1]<=R5C37+20),RC[-24],"""")"
         'Copy Ref M
        Application.Calculation = xlCalculationManual
        With Worksheets("Sheet11")
        LastRow = .Cells(Rows.Count, "M").End(xlUp).Row
    .Range("AK10:AK10").AutoFill Destination:=.Range("AK10:AK" & LastRow), Type:=xlFillDefault
        End With
    'Application.Calculation = xlCalculationAutomatic
    Sheets("Sheet11").Calculate
    
         If Range("AK7") = 0 And Range("AK8") = 0 Then
             Range("AK10").FormulaR1C1 = "=IF(AND(RC[-1]>=R5C37-15,RC[-1]<=R5C37+15),RC[-24],"""")"
            'Copy Ref M
            Application.Calculation = xlCalculationManual
            With Worksheets("Sheet11")
            LastRow = .Cells(Rows.Count, "M").End(xlUp).Row
        .Range("AK10:AK10").AutoFill Destination:=.Range("AK10:AK" & LastRow), Type:=xlFillDefault
            End With
        'Application.Calculation = xlCalculationAutomatic
        Sheets("Sheet11").Calculate
        
            If Range("AK7") = 0 And Range("AK8") = 0 Then
                Range("AK10").FormulaR1C1 = "=IF(AND(RC[-1]>=R5C37-15,RC[-1]<=R5C37+15),RC[-24],"""")"
                'Copy Ref M
                Application.Calculation = xlCalculationManual
                With Worksheets("Sheet11")
                LastRow = .Cells(Rows.Count, "M").End(xlUp).Row
            .Range("AK10:AK10").AutoFill Destination:=.Range("AK10:AK" & LastRow), Type:=xlFillDefault
                End With
            'Application.Calculation = xlCalculationAutomatic
            Sheets("Sheet11").Calculate
                If Range("AK7") = 0 And Range("AK8") = 0 Then
                    Range("AK10").FormulaR1C1 = "=IF(AND(RC[-1]>=R5C37-10,RC[-1]<=R5C37+10),RC[-24],"""")"
                    'Copy Ref M
                    Application.Calculation = xlCalculationManual
                    With Worksheets("Sheet11")
                    LastRow = .Cells(Rows.Count, "M").End(xlUp).Row
                .Range("AK10:AK10").AutoFill Destination:=.Range("AK10:AK" & LastRow), Type:=xlFillDefault
                    End With
                'Application.Calculation = xlCalculationAutomatic
                Sheets("Sheet11").Calculate
                
                    If Range("AK7") = 0 And Range("AK8") = 0 Then
                        Range("AK10").FormulaR1C1 = "=IF(AND(RC[-1]>=R5C37-5,RC[-1]<=R5C37+5),RC[-24],"""")"
                    'Copy Ref M
                    Application.Calculation = xlCalculationManual
                    With Worksheets("Sheet11")
                    LastRow = .Cells(Rows.Count, "M").End(xlUp).Row
                .Range("AK10:AK10").AutoFill Destination:=.Range("AK10:AK" & LastRow), Type:=xlFillDefault
                    End With
                'Application.Calculation = xlCalculationAutomatic
                Sheets("Sheet11").Calculate
                
                End If
            End If
        End If
    End If
    End If
     
    Range("AK10:AQ10009").Value = Range("AK10:AQ10009").Value
    Range("AK10:AQ10009").Sort Key1:=Range("AK10"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
                      
    Range("AR10").FormulaR1C1 = "=IF(RC[-1]-180<0,RC[-1]-180+360,RC[-1]-180)"
    'Copy Ref AK
    Application.Calculation = xlCalculationManual
    With Worksheets("Sheet11")
    LastRow = .Cells(Rows.Count, "AK").End(xlUp).Row
.Range("AR10:AR10").AutoFill Destination:=.Range("AR10:AR" & LastRow), Type:=xlFillDefault
    End With
    'Application.Calculation = xlCalculationAutomatic
    Sheets("Sheet11").Calculate
    
    If Range("B5") = 3 Then
        Sheets("Sheet11").Range("AR5").Value = Sheets("YDWK2").Range("C513").Value
    ElseIf Range("B5") = 2 Then
        Sheets("Sheet11").Range("AR5").Value = Sheets("YDWK2").Range("C441").Value
    ElseIf Range("B5") = 1 And Range("E2") > Range("E3") Then
        Sheets("Sheet11").Range("AR5").Value = Sheets("YDWK2").Range("C297").Value
    ElseIf Range("B5") = 1 And Range("E3") > Range("E2") Then
        Sheets("Sheet11").Range("AR5").Value = Sheets("YDWK2").Range("C369").Value
    End If
    Range("AS10").FormulaR1C1 = _
        "=IF(R5C[-1]+RC[-1]<180,(R5C[-1]+RC[-1])/2,IF(AND(R5C[-1]+RC[-1]>180,R5C[-1]+RC[-1]<360,ABS(R5C[-1]-RC[-1])<180),(R5C[-1]+RC[-1])/2,IF(OR(AND(R5C[-1]>270,RC[-1]<90),AND(RC[-1]>270,R5C[-1]<90)),((R5C[-1]+RC[-1])/2)+180,IF(AND(R5C[-1]+RC[-1]>360,R5C[-1]-RC[-1]<180),(R5C[-1]+RC[-1])/2,IF(R5C[-1]+RC[-1]>540,(R5C[-1]+RC[-1])/2,(R5C[-1]+RC[-1])/2-180)))))"
    Range("AT10").FormulaR1C1 = "=IF(RC[-1]+45>360,RC[-1]+45-360,RC[-1]+45)"
    Range("AU10").FormulaR1C1 = "=IF(RC[-2]-45<0,RC[-2]-45+360,RC[-2]-45)"
    'Copy Ref AK
    Application.Calculation = xlCalculationManual
    With Worksheets("Sheet11")
    LastRow = .Cells(Rows.Count, "AK").End(xlUp).Row
.Range("AS10:AU10").AutoFill Destination:=.Range("AS10:AU" & LastRow), Type:=xlFillDefault
    End With
    'Application.Calculation = xlCalculationAutomatic
    Sheets("Sheet11").Calculate
    
    Range("AR10:AU10009").Value = Range("AR10:AU1009").Value
    
    If Range("A7") = Range("B5") Then
        Sheets("Sheet11").Range("AV10:AV10009").Value = Sheets("TPOrder").Range("DL11:DL10010").Value
    ElseIf Range("A6") = Range("B5") Then
        Sheets("Sheet11").Range("AV10:AV10009").Value = Sheets("TPOrder").Range("CV11:CV10010").Value
    ElseIf Range("A5") = Range("B5") Then
        Sheets("Sheet11").Range("AV10:AV10009").Value = Sheets("TPOrder").Range("CF11:CG10010").Value
    End If
    
    Range("AP6").FormulaR1C1 = "=IF(R[514]C<>"""",LARGE(R[4]C:R[10003]C,500),MAX(R[4]C:R[10003]C))"
    'Range("AP6").FormulaR1C1 = "=LARGE(R[4]C:R[10003]C,500)"
    Sheets("Sheet11").Calculate
    Range("AP6").Value = Range("AP6").Value
    
    Range("AK10:AU10009").Sort Key1:=Range("AP10"), Order1:=xlDescending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    Range("AK10:AK509").Copy
    Range("AW1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    Range("AT10:AU509").Copy
    Range("AW2").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    Range("AP10:AP509").Copy
    Range("AW4").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    
    'Matrix
    Range("AW10:UB10").FormulaR1C1 = "=IF(OR(RC48="""",R1C=""""),"""",IF(OR(AND(R3C<R2C,RC48>=R3C,RC48<=R3C),AND(R2C<R3C,OR(RC48>=R3C,RC48<=R2C))),RC7,""""))"
    'Copy Ref G
    Application.Calculation = xlCalculationManual
    With Worksheets("Sheet11")
    LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
.Range("AW10:UB10").AutoFill Destination:=.Range("AW10:UB" & LastRow), Type:=xlFillDefault
    End With
    'Application.Calculation = xlCalculationAutomatic
    Sheets("Sheet11").Calculate
    
    Range("AW10:UB10009").Value = Range("AW10:UB10009").Value
    Range("AW6:UB6").FormulaR1C1 = "=MAX(R[4]C:R[10003]C)"
    Range("AW7:UB7").FormulaR1C1 = "=IF(R[-1]C<>0,R[-3]C,"""")"
    Range("AT5").FormulaR1C1 = "=MAX(R7C49:R7C548)"
    Sheets("Sheet11").Calculate
    
    Range("AT5").Value = Range("AT5").Value
    Range("AW8:UB8").FormulaR1C1 = "=IF(R[-1]C=R5C46,R[-7]C,"""")"
    Range("AW9:UB9").FormulaR1C1 = "=IF(R[-1]C<>"""",R[-3]C,"""")"
    Range("AT4").FormulaR1C1 = "=MAX(R8C[3]:R8C[502])"
    Range("AT6").FormulaR1C1 = "=MAX(R9C[3]:R9C[502])"
    Sheets("Sheet11").Calculate
    Range("AT4:AT6").Value = Range("AT4:AT6").Value
    
    If Range("AP6") > Range("AT5") Then
        'Save AT4:AT6 somewhere, Test next 500
        Range("AS4:AS6").Value = Range("AT4:AT6").Value
        Range("AW1:UB10009").Clear
        Range("AK510:AK1009").Copy
        Range("AW1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
        Range("AT510:AU1009").Copy
        Range("AW2").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
        Range("AP510:AP1009").Copy
        Range("AW4").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    
        'Matrix
        Range("AW10:UB10").FormulaR1C1 = "=IF(OR(RC48="""",R1C=""""),"""",IF(OR(AND(R3C<R2C,RC48>=R3C,RC48<=R3C),AND(R2C<R3C,OR(RC48>=R3C,RC48<=R2C))),RC7,""""))"
        'Copy Ref G
        Application.Calculation = xlCalculationManual
        With Worksheets("Sheet11")
        LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
    .Range("AW10:UB10").AutoFill Destination:=.Range("AW10:UB" & LastRow), Type:=xlFillDefault
        End With
        'Application.Calculation = xlCalculationAutomatic
        Sheets("Sheet11").Calculate
    
        Range("AW10:UB10009").Value = Range("AW10:UB10009").Value
        Range("AW6:UB6").FormulaR1C1 = "=MAX(R[4]C:R[10003]C)"
        Range("AW7:UB7").FormulaR1C1 = "=IF(R[-1]C<>0,R[-3]C,"""")"
        Range("AT5").FormulaR1C1 = "=MAX(R7C49:R7C548)"
        Sheets("Sheet11").Calculate
    
        Range("AT5").Value = Range("AT5").Value
        Range("AW8:UB8").FormulaR1C1 = "=IF(R[-1]C=R5C46,R[-7]C,"""")"
        Range("AW9:UB9").FormulaR1C1 = "=IF(R[-1]C<>"""",R[-3]C,"""")"
        Range("AT4").FormulaR1C1 = "=MAX(R8C[3]:R8C[502])"
        Range("AT6").FormulaR1C1 = "=MAX(R9C[3]:R9C[502])"
        Sheets("Sheet11").Calculate
        Range("AT4:AT6").Value = Range("AT4:AT6").Value
          
        If Range("AS4") > Range("AT4") Then
            Range("AT4:AT6").Value = Range("AS4:AS6").Value
            Range("AS4:AS6").Clear
        End If
    End If
       
    'Once all tested
    Columns("AW:UC").Clear
    Columns("M:AS").Clear
    Range("AT10:AV10009").Clear

    Range("M10:R10").FormulaR1C1 = "=MAX(R[1]C:R[59999]C)"
    Range("M11").FormulaR1C1 = "=IF(RC[-12]=R4C46,RC[-12],"""")"
    Range("N11:Q11").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-12],"""")"
    Range("R11").FormulaR1C1 = "=IF(RC[-2]<>"""",R5C46,"""")"
    'Copy Ref A
    Range("Y8").FormulaR1C1 = "=IF(R[2]C[-7]<>"""",R[-2]C[21],"""")"
    Application.Calculation = xlCalculationManual
    With Worksheets("Sheet11")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("M11:R11").AutoFill Destination:=.Range("M11:R" & LastRow), Type:=xlFillDefault
    End With
    'Application.Calculation = xlCalculationAutomatic
    Sheets("Sheet11").Calculate
          
    Range("M10:R10").Value = Range("M10:R10").Value
    Range("Y8").Value = Range("Y8").Value
    Range("M11:R60009").Clear
          
End Sub
Sub LoHREDUX()
'
' LoHREDUX Macro
'
    'Last TP lat/lon to A2:B2
    If Range("A7") = Range("B5") Then
        Sheets("Sheet11").Range("A1").Value = Sheets("TPOrder").Range("DN6").Value
        Sheets("Sheet11").Range("B1").Value = Sheets("TPOrder").Range("DO3").Value
        Sheets("Sheet11").Range("A2:B2").Value = Sheets("TPOrder").Range("DL4:DM4").Value
        If Range("B1") = "Cylinder" Then
            Sheets("Sheet11").Range("C1").Value = Sheets("TPOrder").Range("DO6").Value
        End If
    ElseIf Range("A6") = Range("B5") Then
        Sheets("Sheet11").Range("A1").Value = Sheets("TPOrder").Range("CX6").Value
        Sheets("Sheet11").Range("B1").Value = Sheets("TPOrder").Range("CY3").Value
        Sheets("Sheet11").Range("A2:B2").Value = Sheets("TPOrder").Range("CV4:CW4").Value
        If Range("B1") = "Cylinder" Then
            Sheets("Sheet11").Range("C1").Value = Sheets("TPOrder").Range("CY6").Value
        End If
    ElseIf Range("A5") = Range("B5") Then
        Sheets("Sheet11").Range("A1").Value = Sheets("TPOrder").Range("CH6").Value
        Sheets("Sheet11").Range("B1").Value = Sheets("TPOrder").Range("CI3").Value
        Sheets("Sheet11").Range("A2:B2").Value = Sheets("TPOrder").Range("CF4:CG4").Value
        If Range("B1") = "Cylinder" Then
            Sheets("Sheet11").Range("C1").Value = Sheets("TPOrder").Range("CI6").Value
        End If
    End If
    
    Range("AA10").FormulaR1C1 = "=IF(RC[-26]<MAX(R1C1,R1C3),0,6371*ACOS(SIN(R2C1)*SIN(RC[-25])+COS(R2C1)*COS(RC[-25])*COS(RC[-24]-R2C2)))"
    'Copy Ref A
    Application.Calculation = xlCalculationManual
    With Worksheets("Sheet11")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("AA10:AA10").AutoFill Destination:=.Range("AA10:AA" & LastRow), Type:=xlFillDefault
    End With
    
    Range("AA8").FormulaR1C1 = "=IF(LARGE(R[2]C:R[60001]C,500)<>0,LARGE(R[2]C:R[60001]C,500),MIN(R[2]C:R[60001]C))"
    'Range("AA8").FormulaR1C1 = "=LARGE(R[2]C:R[60001]C,500)"
    Sheets("Sheet11").Calculate
    Range("AA8").Value = Range("AA8").Value
   
    'Range("AB10").FormulaR1C1 = "=IF(RC[-1]<>0,RC[-27],"""")"
    Range("AB10").FormulaR1C1 = "=IF(RC[-1]>=R8C27,RC[-27],"""")"
    Range("AC10:AF10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-27],"""")"
    'Copy Ref A
    Application.Calculation = xlCalculationManual
    With Worksheets("Sheet11")
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("AB10:AF10").AutoFill Destination:=.Range("AB10:AF" & LastRow), Type:=xlFillDefault
    End With
    
    Sheets("Sheet11").Calculate
    
    Range("AA10:AF60009").Value = Range("AA10:AF60009").Value
    Range("AA10:AF60009").Sort Key1:=Range("AB10"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    
    If Range("J2") = "MAX" Then
        Range("AG10:AG509").FormulaR1C1 = "=IF(RC[-1]<>"""",R2C5+R2C7+RC[-6],"""")"
    ElseIf Range("J3") = "MAX" Then
        Range("AG10:AG509").FormulaR1C1 = "=IF(RC[-1]<>"""",R3C5+R2C7+RC[-6],"""")"
    End If

    If Range("J2") = "MAX" Then
        Range("AI1").FormulaR1C1 = "=IF(Worksheet!R9C4=""Feet"",R2C[-29]/3.2808399,R2C[-29])"
    ElseIf Range("J3") = "MAX" Then
        Range("AI1").FormulaR1C1 = "=IF(Worksheet!R9C4=""Feet"",R3C[-29]/3.2808399,R3C[-29])"
    End If
    
    'LoH Matrix ckd w / calculator
    If Range("E9") = "" Then
        Range("AI10:AI509").FormulaR1C1 = _
        "=IF(OR(R1C="""",RC[-2]=""""),"""",IF(OR(AND(RC33>100,R1C-RC31<=R9C4),AND(RC33<=100,R1C-RC31<=RC[-2]*10)),RC33,IF(RC33>100,RC33-(10*(R1C-RC31-R9C4)*0.1),RC33-0.1*(R1C-RC31-((10*RC[-2]))))))"
     ElseIf Range("E9") = "PR" Then
         Range("AI10").FormulaR1C1 = _
        "=IF(OR(R1C="""",RC[-2]=""""),"""",IF(OR(AND(RC33>100,R1C-RC31<R9C4),AND(RC33<=100,R1C-RC31<=0.01*(RC33*1000)-100)),RC33,IF(RC33>100,RC33-(R1C-RC31-R9C4)*0.1,RC33-(R1C-RC31-(0.01*(RC33*1000)-100))*0.1)))"
     End If
    
    Sheets("Sheet11").Calculate
    Range("AI10:AI509").Value = Range("AI10:AI509").Value
    
    Range("AI8").FormulaR1C1 = "=MAX(R10C:R509C)"
    Range("AJ8:AN8").FormulaR1C1 = "=MAX(R[2]C:R[501]C)"
    Sheets("Sheet11").Calculate
    Range("AI8").Value = Range("AI8").Value
    Range("AJ10:AJ509").FormulaR1C1 = "=IF(RC[-1]=R8C[-1],RC[-8],"""")"
    Range("AK10:AN509").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-8],"""")"
    Sheets("Sheet11").Calculate
    Range("L10:Q10").Value = Range("AI8:AN8").Value
    Columns("R:AN").Clear
 
 Sheets("YDWK2").Range("E1609").Value = 0
 Sheets("YDWK2").Range("E1610").Value = 0
 Sheets("YDWK2").Range("E1611").Value = 0
 Sheets("YDWK2").Range("E1612").Value = 0
    Sheets("YDWK2").Range("E1651").FormulaR1C1 = _
        "=R[-34]C[-2]+(1-R[-2]C)*R[-35]C[-2]*R[-14]C*(ACOS(R[-19]C)+R[-2]C*SIN(ACOS(R[-19]C))*(R[-12]C+R[-2]C*R[-19]C*(-1+2*R[-12]C*R[-12]C)))"
 Sheets("YDWK2").Calculate
 Sheets("Sheet11").Activate
 If Range("C1") = "" Then
    Application.Run ("C.xlsm!MinLOH")
 End If
     
End Sub
Sub MinLOH()
'
' re-calc last TP OZ at revised Fin Fix
'
    If Range("A7") = Range("B5") Then
        Sheets("YDWK2").Range("E1609").Value = Sheets("TPOrder").Range("DL4").Value
        Sheets("YDWK2").Range("E1610").Value = Sheets("TPOrder").Range("DM4").Value
    ElseIf Range("A6") = Range("B5") Then
        Sheets("YDWK2").Range("E1609").Value = Sheets("TPOrder").Range("CV4").Value
        Sheets("YDWK2").Range("E1610").Value = Sheets("TPOrder").Range("CW4").Value
    ElseIf Range("A5") = Range("B5") Then
        Sheets("YDWK2").Range("E1609").Value = Sheets("TPOrder").Range("CF4").Value
        Sheets("YDWK2").Range("E1610").Value = Sheets("TPOrder").Range("CG4").Value
    End If
    
    Sheets("YDWK2").Range("E1611").Value = Sheets("Sheet11").Range("N10").Value
    Sheets("YDWK2").Range("E1612").Value = Sheets("Sheet11").Range("O10").Value
    Sheets("YDWK2").Range("E1651").FormulaR1C1 = _
        "=R[-34]C[-2]+(1-R[-2]C)*R[-35]C[-2]*R[-14]C*(ACOS(R[-19]C)+R[-2]C*SIN(ACOS(R[-19]C))*(R[-12]C+R[-2]C*R[-19]C*(-1+2*R[-12]C*R[-12]C)))"
    'Application.Calculation = xlCalculationAutomatic
    Sheets("YDWK2").Calculate
    'Application.Calculation = xlCalculationManual
    
    'T5 variable w/last TP
    If Range("B5") = 3 Then
        Sheets("Sheet11").Range("T5").Value = Sheets("YDWK2").Range("C513").Value
    ElseIf Range("B5") = 2 Then
        Sheets("Sheet11").Range("T5").Value = Sheets("YDWK2").Range("C441").Value
    ElseIf Range("B5") = 1 And Range("E2") > Range("E3") Then
        Sheets("Sheet11").Range("T5").Value = Sheets("YDWK2").Range("C297").Value
    ElseIf Range("B5") = 1 And Range("E3") > Range("E2") Then
        Sheets("Sheet11").Range("T5").Value = Sheets("YDWK2").Range("C369").Value
    End If
            
    Sheets("Sheet11").Range("S10").Value = Sheets("YDWK2").Range("C1606").Value
    Range("T10").FormulaR1C1 = "=IF(RC[-1]-180<0,RC[-1]-180+360,RC[-1]-180)"
    Range("U10").FormulaR1C1 = _
        "=IF(R5C20+RC[-1]<180,(R5C20+RC[-1])/2,IF(AND(R5C20+RC[-1]>180,R5C20+RC[-1]<360,ABS(R5C20-RC[-1])<180),(R5C20+RC[-1])/2,IF(OR(AND(R5C20>270,RC[-1]<90),AND(RC[-1]>270,R5C20<90)),((R5C20+RC[-1])/2)+180,IF(AND(R5C20+RC[-1]>360,R5C20-RC[-1]<180),(R5C20+RC[-1])/2,IF(R5C20+RC[-1]>540,(R5C20+RC[-1])/2,(R5C20+RC[-1])/2-180)))))"
    Range("V10").FormulaR1C1 = "=IF(RC[-1]+45>360,RC[-1]+45-360,RC[-1]+45)"
    Range("W10").FormulaR1C1 = "=IF(RC[-2]-45<0,RC[-2]-45+360,RC[-2]-45)"
    Sheets("Sheet11").Calculate
    Range("T10:W10").Value = Range("T10:W10").Value
            
    Range("M10").Copy
    Range("Y1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
    Range("V10:W10").Copy
    Range("Y2").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    Application.CutCopyMode = False
    
    Range("Y8").FormulaR1C1 = "=MAX(R[2]C:R[10001]C)"
    'X10 et al per TP
    If Range("A7") = Range("B5") Then
        Sheets("Sheet11").Range("X10:X10009").Value = Sheets("TPOrder").Range("DL11:DL10010").Value
    ElseIf Range("A6") = Range("B5") Then
        Sheets("Sheet11").Range("X10:X10009").Value = Sheets("TPOrder").Range("CV11:CV10010").Value
    ElseIf Range("A5") = Range("B5") Then
        Sheets("Sheet11").Range("X10:X10009").Value = Sheets("TPOrder").Range("CF11:CG10010").Value
    End If
    
    Range("Y10").FormulaR1C1 = _
        "=IF(OR(R1C="""",RC7="""",RC7>R1C,RC[-1]=""""),"""",IF(OR(AND(R2C>R3C,RC24>=R3C,RC24<=R2C),AND(R3C>R2C,OR(RC24>=R3C,RC24<=R2C))),RC7,""""))"
    'Copy Ref G
    Application.Calculation = xlCalculationManual
    With Worksheets("Sheet11")
    LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
.Range("Y10:Y10").AutoFill Destination:=.Range("Y10:Y" & LastRow), Type:=xlFillDefault
    End With
    Sheets("Sheet11").Calculate
    Range("R10").Value = Range("Y8").Value
End Sub
Sub GPSsfLines()
'
' GPS Start/Fin Lines & landing added 9/7/15
'
    Sheets("Sheet11").Activate
    Sheets("Sheet11").Range("M5").Value = Sheets("TPOrder").Range("AY9").Value
    Range("M10").FormulaR1C1 = "=IF(OR(R[-1]C[-6]=R5C,RC[-6]=R5C),RC[-6],"""")"
    Range("N10").FormulaR1C1 = "=IF(RC[-1]="""","""",RC[-3])"
    'Copy Ref G Value Sort
    With Worksheets("Sheet11")
    LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
.Range("M10:N10").AutoFill Destination:=.Range("M10:N" & LastRow), Type:=xlFillDefault
    End With
    Sheets("Sheet11").Calculate
    Range("M10:N10009").Value = Range("M10:N10009").Value
    Range("M10:N10009").Sort Key1:=Range("M10"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
   
    Sheets("Sheet11").Range("P5").Value = Sheets("TPOrder").Range("EY9").Value
    Range("P10").FormulaR1C1 = "=IF(OR(RC[-9]=R5C,R[1]C[-9]="""",R[1]C[-9]=R5C),RC[-9],"""")"
    Range("Q10").FormulaR1C1 = "=IF(RC[-1]="""","""",RC[-6])"
    'Copy Ref G Value Sort
    With Worksheets("Sheet11")
    LastRow = .Cells(Rows.Count, "G").End(xlUp).Row
.Range("P10:Q10").AutoFill Destination:=.Range("P10:Q" & LastRow), Type:=xlFillDefault
    End With
    Sheets("Sheet11").Calculate
    Range("P10:Q10009").Value = Range("P10:Q10009").Value
    Range("P10:Q10009").Sort Key1:=Range("P10"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
   
    Range("M11:Q12").Cut Destination:=Range("M12:Q13")
    Sheets("Sheet11").Range("O6:O7").Value = Sheets("TPOrder").Range("FV29:FV30").Value
    Range("N11").FormulaR1C1 = "=IF(OR(R[-5]C[1]="""",R[-4]C[1]=""""),0,(R[1]C-R[-1]C)*R[-5]C[1]+R[-1]C)"
    Sheets("Sheet11").Range("R6:R7").Value = Sheets("TPOrder").Range("FV45:FV46").Value
    Range("Q11").FormulaR1C1 = "=IF(OR(R[-5]C[1]="""",R[-4]C[1]=""""),0,(R[1]C-R[-1]C)*R[-5]C[1]+R[-1]C)"
    Sheets("Sheet11").Calculate
    'Put it somewhere
    Sheets("B").Range("J20").Value = Sheets("Sheet11").Range("N11").Value
    Sheets("B").Range("J28").Value = Sheets("Sheet11").Range("Q11").Value
    Sheets("B").Range("J30").Value = Sheets("Sheet11").Range("Q13").Value
End Sub
