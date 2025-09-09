' VBA Module: Comprehensive Flight Analysis and Calculations
' Purpose: Performs complex flight analysis including ground track calculations, GPS altitude processing,
' speed calculations, and competitive soaring record validation. Handles multiple flight parameters
' including start accuracy, finish timing, and turn point validations for official flight records.

Option Explicit
#If VBA7 And Win64 Then
    Dim LastRow As LongPtr
#Else
    Dim LastRow As Long
#End If

Sub Mega()
'
' JL Ruprecht 3/17/13; Ground track AS11,CG11,CW11,DM11,EP11 revised 4/21/15 for ground track Other revisions commented out for TP1=TP3
' Revised 7/5/15 & 7/7/2015 for Longest Flt; manual calcs; Revised 7/24 & 27/2015 for Start accuracy @ AU5 needs to be at end!
'AT11,AV11,AW11,AX11,BR11; 8/4/2015 @ BU11,CK11,DA11 & TP GR TRK; 8/9/2015 @ AT11 & AU5 for no TP; 8/10/15 AV11,EQ11,EY4,EY9,EY11,for accuracy @ Start, FIni;
'8/17/15 CX11 for Longest Flt Last TP OZ time < Fin Fix;Revised 9/6/15 for GPS Alt @ Fin Fix; revised 9/16 EZ5 = YDWK2 R877C6 ("N/A" from YDWK2)
'Amended 10/2/2015 for GPS altitde alternative @ AV11,BD11,BE11,BN11,BO11,BS11,ES11,EW11,EX11,FD11,FE11,FN12,FO12,FR12; 10/13/2015 CW5 &CX11 for ST@TP; 12/25/2015 CW6 & CX11 amended
'Amended 12/30/15 @ CI1,CY1,DO1 eg: # if achieved (Previous =R[1]C ); Amended 1/31/16 @ CI1 restored to original; amended 6/24/16 @ AS9 Start Sector by ground track
'Amended 9/13/2017 @ DO1, CY1 & CI1, removed ref to Fin Fix from CX11; Amended @ EW4 2/2/18 for 100 km Speed based on calculated best Start Time
'Amended 7/3/2018 @ EY11 to eliminate wrong-way fini
'
Application.ScreenUpdating = False
Workbooks("C.xlsm").Unprotect Password:="spike"
Application.Calculation = xlCalculationAutomatic
Sheets("YDWK1").Visible = True
Sheets("YDWK1").Activate
Sheets("YDWK1").Unprotect Password:="spike"
Sheets("YDWK1").Range("J1").Value = Sheets("B").Range("A44").Value
Sheets("YDWK1").Range("K1").Value = Sheets("Worksheet").Range("N1").Value
Sheets("YDWK1").Range("L1").Value = Sheets("Worksheet").Range("Q1").Value

If Range("J1") = "" And Range("K1") = "OO Cert" And Range("L1") = "" Then
    Range("E4").FormulaR1C1 = "=Worksheet!R[1]C[8]"
    Range("F7:G14").FormulaR1C1 = "=CALIBRATION!R[-1]C[1]"
    Range("I7:J14").FormulaR1C1 = "=CALIBRATION!R[7]C[-2]"
    Sheets("YDWK1").Protect Password:="spike"
    Sheets("Calibration").Visible = True
    
ElseIf Range("J1") > 0 Or Range("K1") <> "OO Cert" Or Range("L1") <> "" Then
    Sheets("YDWK1").Range("E4").Value = Sheets("B").Range("A42").Value
    Sheets("YDWK1").Range("F7:G14").Value = Sheets("B").Range("A43:B50").Value
    Sheets("YDWK1").Range("I7:J19").Value = Sheets("B").Range("A51:B63").Value
    Sheets("YDWK1").Protect Password:="spike"
    Sheets("Calibration").Visible = False
End If
Sheets("YDWK1").Visible = False
Sheets("TPOrder").Activate
Sheets("TPOrder").Unprotect Password:="spike"
Application.Calculation = xlCalculationManual
With Worksheets("TPOrder")

    Range("A3").Value = Sheets("B").Range("I20").Value
    Range("A4").Value = Sheets("B").Range("I22").Value
    Range("A5").Value = Sheets("B").Range("I24").Value
    Range("A6").Value = Sheets("B").Range("I26").Value
    Range("A7").Value = Sheets("B").Range("I28").Value
    Range("C3").FormulaR1C1 = "0.00335281066474748"
    Range("C5").FormulaR1C1 = "=SUMMARY!R[9]C[6]"
    Range("O1,S1,W1").FormulaR1C1 = "=IF(AND(R4C>0,R10C>0),1,0)"
    
 If Range("A3") > 0 Then
    Range("H11").FormulaR1C1 = "=IF(RC[-3]="""","""",(1-R3C3)*TAN(RC[-5]))"
    Range("I11").FormulaR1C1 = "=IF(RC[-1]="""","""",ATAN(RC[-1]))"
    Range("J11").FormulaR1C1 = "=IF(RC[-1]="""","""",SIN(RC[-1]))"
    Range("K11").FormulaR1C1 = "=IF(RC[-1]="""","""",COS(RC[-2]))"
    Range("L3").FormulaR1C1 = "=RADIANS(B!R[19]C[-11]+B!R[19]C[-10]/60)"
    Range("L4").FormulaR1C1 = "=RADIANS(B!R[18]C[-8]+B!R[18]C[-7]/60)"
    Range("L11").FormulaR1C1 = _
        "=IF(AND(RC[-11]=""""=FALSE,RC[-11]<R5C3,R4C12=0=FALSE),6371*ACOS((SIN(RC[-9])*SIN(R3C12)+COS(RC[-9])*COS(R3C12)*COS(RC[-7]-R4C12))),"""")"
    Range("L10017").FormulaR1C1 = "=IF(MIN(R[-10006]C:R[-7]C)>50,0,MIN(R[-10006]C:R[-7]C))"
    Range("M11").FormulaR1C1 = "=IF(OR(RC[-12]="""",AND(R10017C[-1]<=10,RC[-1]>10),AND(R10017C[-1]>10,RC[-1]>2*R10017C[-1])),"""",RC[-12])"
    Range("M10019").FormulaR1C1 = "=MAX(R[-10008]C:R[-9]C)"
    Range("M10021").FormulaR1C1 = "=MIN(R[-10010]C:R[-11]C)"
    Range("N11").FormulaR1C1 = "=IF(AND(R10021C[-1]>0,RC[-13]>=R10021C[-1],RC[-2]>=10,RC[-13]<=R10019C[-1]),RC[-13],"""")"
    Range("N10019").FormulaR1C1 = "=IF(MIN(R[-10008]C:R[-9]C)=0,RC[-1],MIN(R[-10008]C:R[-9]C))"
    Range("O11").FormulaR1C1 = "=IF(AND(RC[-14]>R8C15,RC[-14]<R4C15),RC[-14],"""")"
    Range("O4").FormulaR1C1 = "=R10019C[-2]"
    Range("O6").FormulaR1C1 = "=IF(MAX(R[5]C:R[10004]C)>0,""YES"",""NO"")"
    Range("O8").FormulaR1C1 = "=R10019C[-1]"
    Range("O10").FormulaR1C1 = "=R10021C[-2]"
    Range("P3").FormulaR1C1 = "=RADIANS(B!R[21]C[-15]+B!R[21]C[-14]/60)"
    Range("P4").FormulaR1C1 = "=RADIANS(B!R[20]C[-12]+B!R[20]C[-11]/60)"
    Range("P11").FormulaR1C1 = "=IF(AND(RC[-15]=""""=FALSE,RC[-15]<R5C3,R4C16=0=FALSE),6371*ACOS((SIN(RC[-13])*SIN(R3C16)+COS(RC[-13])*COS(R3C16)*COS(RC[-11]-R4C16))),"""")"
    Range("P10017").FormulaR1C1 = "=IF(MIN(R[-10006]C:R[-7]C)>50,0,MIN(R[-10006]C:R[-7]C))"
    Range("Q11").FormulaR1C1 = "=IF(OR(RC[-16]="""",AND(R10017C[-1]<=10,RC[-1]>10),AND(R10017C[-1]>10,RC[-1]>2*R10017C[-1])),"""",RC[-16])"
    Range("Q10019").FormulaR1C1 = "=MAX(R[-10008]C:R[-9]C)"
    Range("Q10021").FormulaR1C1 = "=MIN(R[-10010]C:R[-11]C)"
    Range("R2").FormulaR1C1 = "=6371*ACOS(SIN(R[1]C[-2])*SIN(R[1]C[-6])+COS(R[1]C[-2])*COS(R[1]C[-6])*COS(R[2]C[-6]-R[2]C[-2]))"
    Range("R11").FormulaR1C1 = "=IF(AND(R10021C[-1]>0,RC[-17]>=R10021C[-1],RC[-2]>=10,RC[-17]<=R10019C[-1]),RC[-17],"""")"
    Range("R10019").FormulaR1C1 = "=IF(MIN(R[-10008]C:R[-9]C)=0,R[-10015]C[1],MIN(R[-10008]C:R[-9]C))"
    Range("S11").FormulaR1C1 = "=IF(AND(RC[-18]>R8C19,RC[-18]<R4C19),RC[-18],"""")"
    Range("S4").FormulaR1C1 = "=R10019C[-2]"
    Range("S6").FormulaR1C1 = "=IF(MAX(R[5]C:R[10004]C)>0,""YES"",""NO"")"
    Range("S8").FormulaR1C1 = "=R10019C[-1]"
    Range("S10").FormulaR1C1 = "=R10021C[-2]"
    Range("T3").FormulaR1C1 = "=RADIANS(B!R[23]C[-19]+B!R[23]C[-18]/60)"
    Range("T4").FormulaR1C1 = "=RADIANS(B!R[22]C[-16]+B!R[22]C[-15]/60)"
    Range("T11").FormulaR1C1 = "=IF(AND(RC[-19]=""""=FALSE,RC[-19]<R5C3,R4C20=0=FALSE),6371*ACOS((SIN(RC[-17])*SIN(R3C20)+COS(RC[-17])*COS(R3C20)*COS(RC[-15]-R4C20))),"""")"
    Range("T10017").FormulaR1C1 = "=IF(MIN(R[-10006]C:R[-7]C)>50,0,MIN(R[-10006]C:R[-7]C))"
    Range("U11").FormulaR1C1 = "=IF(OR(RC[-20]="""",AND(R10017C[-1]<=10,RC[-1]>10),AND(R10017C[-1]>10,RC[-1]>2*R10017C[-1])),"""",RC[-20])"
    Range("U10019").FormulaR1C1 = "=MAX(R11C21:R10010C21)"
    Range("U10021").FormulaR1C1 = "=MIN(R11C21:R10010C21)"
    Range("V2").FormulaR1C1 = "=6371*ACOS(SIN(R[1]C[-2])*SIN(R[1]C[-6])+COS(R[1]C[-2])*COS(R[1]C[-6])*COS(R[2]C[-6]-R[2]C[-2]))"
    Range("V11").FormulaR1C1 = "=IF(AND(R10021C[-1]>0,RC[-21]>=R10021C[-1],RC[-2]>=10,RC[-21]<=R10019C[-1]),RC[-21],"""")"
    Range("V10019").FormulaR1C1 = "=IF(MIN(R11C22:R10010C22)=0,R4C23,MIN(R11C22:R10010C22))"
    Range("W11").FormulaR1C1 = "=IF(AND(RC[-22]>R8C,RC[-22]<R4C),RC[-22],"""")"
    Range("W4").FormulaR1C1 = "=R10019C[-2]"
    Range("W6").FormulaR1C1 = "=IF(MAX(R[5]C:R[10004]C)>0,""YES"",""NO"")"
    Range("W8").FormulaR1C1 = "=R10019C[-1]"
    Range("W10").FormulaR1C1 = "=R10021C[-2]"
    Range("X5").FormulaR1C1 = "=IF(OR(AND(R5C27>0,R6C27>0,R7C27>0,OR(AND(R10C19>0,R10C23>0,R10C15=MIN(R10C15,R10C19,R10C23),R4C15=MIN(R4C15,R4C19,R4C23)),R10C15=MIN(R10C15,R10C19,R10C23))),AND(SUMMARY!R30C[-16]=2,OR(R10C19>0,R10C23>0),R4C15<MAX(R4C19,R4C23)),AND(SUMMARY!R30C8=1,R10C15=MAX(R10C15,R10C19,R10C23))),1,"""")"
    Range("Y3").FormulaR1C1 = "=MAX(ABS(R[1]C[-2]-R[7]C[-10]),ABS(R[1]C[-10]-R[7]C[-2]),ABS(R[5]C[-10]-R[5]C[-2]))"
    Range("Y4").FormulaR1C1 = "=(HOUR(R[-1]C)+MINUTE(R[-1]C)*60+SECOND(R[-1]C))"
    Range("Y5").FormulaR1C1 = "=IF(OR(AND(R5C27>0,R6C27>0,R7C27>0,OR(AND(R10C19>0,R10C23>0,R10C19=MIN(R10C15,R10C19,R10C23),R4C19=MIN(R4C15,R4C19,R4C23)),R10C19=MIN(R10C19,R10C15,R10C23))),AND(SUMMARY!R30C[-16]=2,OR(R10C15>0,R10C23>0),R4C19<MAX(R4C15,R4C23)),AND(SUMMARY!R30C8=1,R10C19=MAX(R10C15,R10C19,R10C23))),1,"""")"
    Range("Z5").FormulaR1C1 = "=IF(OR(AND(R5C27>0,R6C27>0,R7C27>0,OR(AND(R10C15>0,R10C23>0,R10C23=MIN(R10C15,R10C19,R10C23),R4C23=MIN(R4C15,R4C19,R4C23)),R10C23=MIN(R10C23,R10C19,R10C15))),AND(SUMMARY!R30C[-16]=2,OR(R10C15>0,R10C19>0),R4C23<MAX(R4C19,R4C15)),AND(SUMMARY!R30C8=1,R10C23=MAX(R10C15,R10C19,R10C23))),1,"""")"
    Range("X6").FormulaR1C1 = "=IF(OR(AND(R[-1]C[3]>0,RC[3]>0,R[1]C[3]>0,OR(AND(R5C27=MIN(R5C27:R7C27),MAX(R5C27:R7C27)-MIN(R5C27:R7C27)>1/24),AND(R10C15>MIN(R10C19,R10C23),R10C15<MAX(R10C19,R10C23),R4C15>MIN(R4C19,R4C23),R4C15<MAX(R4C19,R4C23)))),AND(SUMMARY!R30C8=2,OR(R10C15=MAX(R10C15,R10C19,R10C23),R4C15=MAX(R4C15,R4C23)))),2,"""")"
    Range("Y6").FormulaR1C1 = "=IF(OR(AND(R5C27>0,R6C27>0,R7C27>0,OR(AND(R6C27=MIN(R5C27:R7C27),R4C19=R8C19,OR(AND(R4C19>R8C15,R8C19<R4C23),AND(R4C19>R8C23,R4C19<R8C15))),AND(R10C19>MIN(R10C15,R10C23),R10C19<MAX(R10C15,R10C23),R4C19>MIN(R4C15,R4C23),R4C19<MAX(R4C15,R4C23)))),AND(SUMMARY!R30C8=2,OR(R10C19=MAX(R10C15,R10C19,R10C23),R4C19=MAX(R4C15,R4C23)))),2,"""")"
    Range("Z6").FormulaR1C1 = "=IF(OR(AND(R[-1]C[1]>0,RC[1]>0,R[1]C[1]>0,OR(AND(R7C27=MIN(R5C27:R7C27),MAX(R5C27:R7C27)-MIN(R5C27:R7C27)>1/24),AND(R10C23>MIN(R10C19,R10C15),R10C23<MAX(R10C19,R10C15),R4C23>MIN(R4C19,R4C15),R4C23<MAX(R4C19,R4C15)))),AND(SUMMARY!R30C8=2,OR(R10C23=MAX(R10C15,R10C19,R10C23),R4C23=MAX(R4C15,R4C23)))),2,"""")"
    Range("X7").FormulaR1C1 = "=IF(MAX(R[-2]C[5]:RC[5])<3,"""",IF(OR(AND(R5C25=1,R6C26=2),AND(R6C25=2,R5C26=1),AND(R10C15>0,R10C19>0,R10C23>0,R10C15=MAX(R10C15,R10C19,R10C23),R4C15=MAX(R4C15,R4C19,R4C23))),3,""""))"
    Range("Y7").FormulaR1C1 = "=IF(MAX(R[-2]C[4]:RC[4])<3,"""",IF(OR(AND(R5C24=1,R6C26=2),AND(R6C24=1,R5C26=1),AND(R10C15>0,R10C19>0,R10C23>0,R10C19=MAX(R10C15,R10C19,R10C23),R4C19=MAX(R4C15,R4C19,R4C23))),3,""""))"
    Range("Z7").FormulaR1C1 = "=IF(MAX(R[-2]C[3]:RC[3])<3,"""",IF(OR(AND(R5C24=1,R6C25=2),AND(R6C24=2,R5C25=1),AND(R10C15>0,R10C19>0,R10C23>0,R10C23=MAX(R10C15,R10C19,R10C23),R4C23=MAX(R4C15,R4C19,R4C23))),3,""""))"
    Range("Z9").FormulaR1C1 = _
        "=IF(SUM(R5C[1]:R7C[1])=0,"""",IF(AND(R[-1]C[-11]<>0,R[-1]C[-7]<>0,R[-1]C[-7]=MAX(R[-1]C[-11],R[-1]C[-7],R[-1]C[-3]),R[-1]C[-3]<>0,R[-1]C[-3]<R[-1]C[-11],R[-1]C[-3]<R[-5]C[-11],R[-5]C[-3]<R[-1]C[-11],R[-7]C[-4]/R[-5]C[-1]*3600>=400),""X"",""""))"
    Range("X8").FormulaR1C1 = "='Verify Task'!R12C4"
    Range("Y8").FormulaR1C1 = "='Verify Task'!R14C4"
    Range("Z8").FormulaR1C1 = "='Verify Task'!R16C4"
    Range("AA5").FormulaR1C1 = "=R[-1]C[-12]-R[5]C[-12]"
    Range("AA6").FormulaR1C1 = "=R[-2]C[-8]-R[4]C[-8]"
    Range("AA7").FormulaR1C1 = "=R[-3]C[-4]-R[3]C[-4]"
    Range("AB5").FormulaR1C1 = "=IF('Verify Task'!R[7]C[-22]="""","""",'Verify Task'!R[7]C[-22])"
    Range("AB6").FormulaR1C1 = "=IF('Verify Task'!R[8]C[-22]="""","""",'Verify Task'!R[8]C[-22])"
    Range("AB7").FormulaR1C1 = "=IF('Verify Task'!R[9]C[-22]="""","""",'Verify Task'!R[9]C[-22])"
    Range("AC5").FormulaR1C1 = "=IF(RC[-1]="""",SUMMARY!R[22]C[-17],RC[-1])"
    Range("AC6").FormulaR1C1 = "=IF(RC[-1]="""",SUMMARY!R[22]C[-17],RC[-1])"
    Range("AC7").FormulaR1C1 = "=IF(RC[-1]="""",SUMMARY!R[22]C[-17],RC[-1])"
    Range("Z15").FormulaR1C1 = "=IF(AND(SUM(R[-10]C[1]:R[-8]C[1])=0,YDWK2!R224C3=0),YDWK2!R[67]C[-23],IF(SUM(R[-10]C[3]:R[-8]C[3])=0,YDWK2!R224C3,YDWK2!R368C3))"
    Range("AC15").FormulaR1C1 = "=IF(OR(R[-12]C[19]=""USE RELEASE"",AND(SUMMARY!R[9]C[-28]=SUMMARY!R[-8]C[-20],SUMMARY!R[9]C[-27]=SUMMARY!R[-8]C[-19],SUMMARY!R[11]C[-28]=SUMMARY!R[-6]C[-20],SUMMARY!R[11]C[-27]=SUMMARY!R[-6]C[-19])),YDWK2!R297C3,YDWK2!R369C3)"
    Range("Z18").FormulaR1C1 = "=IF(AND(YDWK2!R296C3-180<0,OR(R3C48=""USE RELEASE"",AND(SUMMARY!R24C1=SUMMARY!R7C9,SUMMARY!R24C2 =SUMMARY!R7C10,SUMMARY!R26C1=SUMMARY! R9C9,SUMMARY!R26C2=SUMMARY!R9C10))),YDWK2!R296C3-180+360,IF(AND(YDWK2!R296C3-180>0,OR(R3C48=""USE RELEASE"",AND(SUMMARY!R24C1 =SUMMARY! R7C9,SUMMARY!R24C2=SUMMARY!R7C10,SUMMARY!R26C1=SUMMARY!R9C9,SUMMARY!R28C2=SUMMARY!R9C10))),YDWK2!R296C3-180,IF(YDWK2!R368C3-180<0,YDWK2!R368C3-180+360,YDWK2!R368C3-180)))"
    Range("Z19").FormulaR1C1 = "=YDWK2!R[421]C[-23]"
    Range("AA19").FormulaR1C1 = "=R[-1]C[-1]-RC[-1]"
    Range("Z21").FormulaR1C1 = "=IF(R[-3]C+R[-2]C<180,(R[-3]C+R[-2]C)/2,IF(AND(R[-3]C+R[-2]C>180,R[-3]C+R[-2]C<360,ABS(R[-3]C-R[-2]C)<180),(R[-3]C+R[-2]C)/2,IF(OR(AND(R[-3]C>270,R[-2]C<90),AND(R[-2]C>270,R[-3]C<90)),((R[-3]C+R[-2]C)/2)+180,IF(AND(R[-3]C+R[-2]C>360,R[-2]C[1]<180),(R[-3]C+R[-2]C)/2,IF(R[-3]C+R[-2]C>540,(R[-3]C+R[-2]C)/2,(R[-3]C+R[-2]C)/2-180)))))"
    Range("Z24").FormulaR1C1 = "=IF(YDWK2!R440C3-180<0,YDWK2!R440C3-180+360,YDWK2!R440C3-180)"
    Range("Z25").FormulaR1C1 = "=YDWK2!R512C3"
    Range("AA25").FormulaR1C1 = "=R[-1]C[-1]-RC[-1]"
    Range("Z27").FormulaR1C1 = "=IF(R[-3]C+R[-2]C<180,(R[-3]C+R[-2]C)/2,IF(AND(R[-3]C+R[-2]C>180,R[-3]C+R[-2]C<360,ABS(R[-3]C-R[-2]C)<180),(R[-3]C+R[-2]C)/2,IF(OR(AND(R[-3]C>270,R[-2]C<90),AND(R[-2]C>270,R[-3]C<90)),((R[-3]C+R[-2]C)/2)+180,IF(AND(R[-3]C+R[-2]C>360,R[-2]C[1]<180),(R[-3]C+R[-2]C)/2,IF(R[-3]C+R[-2]C>540,(R[-3]C+R[-2]C)/2,(R[-3]C+R[-2]C)/2-180)))))"
    Range("Z30").FormulaR1C1 = "=IF(SUMMARY!R[1]C[-18]=0,0,IF(AND(SUMMARY!R[1]C[-18]=3,YDWK2!R512C3-180>0),YDWK2!R512C3-180,IF(AND(SUMMARY!R[1]C[-18]=3,YDWK2!R512C3-180<0),YDWK2!R512C3-180+360,IF(AND(SUMMARY!R[1]C[-18]=2,YDWK2!R440C3-180>0),YDWK2!R440C3-180,IF(AND(SUMMARY!R[1]C[-18]=2,YDWK2!R440C3-180<0),YDWK2!R440C3-180+360,IF(AND(SUMMARY!R[1]C[-18]=1,YDWK2!R368C3-180<0),YDWK2!R368C3-180+360,YDWK2!R368C3-180))))))"
    Range("Z31").FormulaR1C1 = "=IF(OR(SUMMARY!R[-1]C[-25]=0,YDWK2!R[625]C[-23]=0),YDWK2!R[697]C[-23],YDWK2!R656C3)"
    Range("AA31").FormulaR1C1 = "=R[-1]C[-1]-RC[-1]"
    Range("Z33").FormulaR1C1 = "=IF(R[-3]C+R[-2]C<180,(R[-3]C+R[-2]C)/2,IF(AND(R[-3]C+R[-2]C>180,R[-3]C+R[-2]C<360,ABS(R[-3]C-R[-2]C)<180),(R[-3]C+R[-2]C)/2,IF(OR(AND(R[-3]C>270,R[-2]C<90),AND(R[-2]C>270,R[-3]C<90)),((R[-3]C+R[-2]C)/2)+180,IF(AND(R[-3]C+R[-2]C>360,R[-2]C[1]<180),(R[-3]C+R[-2]C)/2,IF(R[-3]C+R[-2]C>540,(R[-3]C+R[-2]C)/2,(R[-3]C+R[-2]C)/2-180)))))"
    Range("Z37").FormulaR1C1 = "=IF(SUMMARY!R[-6]C[-18]=0,YDWK2!R[188]C[-23],YDWK2!R657C3)"
    Range("AD2").FormulaR1C1 = "=B!R20C1+B!R20C2/60"
    Range("AD4").FormulaR1C1 = "=B!R20C4+B!R20C5/60"
    Range("AR4").FormulaR1C1 = "=RADIANS(R2C30)"
    Range("AT4").FormulaR1C1 = "=RADIANS(R4C30)"
    Range("AS5").FormulaR1C1 = "=SUMMARY!R4C7"
    Range("AI5").FormulaR1C1 = "=(1-R3C3)*TAN(R4C44)"
    Range("AI6").FormulaR1C1 = "=ATAN(R[-1]C)"
    Range("AI7").FormulaR1C1 = "=SIN(R[-1]C)"
    Range("AI8").FormulaR1C1 = "=COS(R[-2]C)"
    Range("AO4").FormulaR1C1 = "=R[11]C[-15]"
    Range("AO5").FormulaR1C1 = "=IF(R[-1]C="""","""",R[-1]C+45)"
    Range("AQ5").FormulaR1C1 = "=IF(R[-1]C[-2]="""","""",IF(RC[-2]>360,R[-1]C[-2]+45-360,RC[-2]))"
    Range("AO6").FormulaR1C1 = "=IF(R[-2]C="""","""",R[-2]C-45)"
    Range("AQ6").FormulaR1C1 = "=IF(R[-1]C="""","""",IF(R[-2]C[-2]-45<0,R[-2]C[-2]-45+360,RC[-2]))"
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("H11:W11").AutoFill Destination:=.Range("H11:W" & LastRow), Type:=xlFillDefault
  
    Range("AD11").FormulaR1C1 = _
        "=IF(AND(RC[-29]<>"""",RC[-28]=R2C,RC[-26]=R4C),0,IF(RC[-19]="""","""",6371*ACOS((SIN(RC[-27])*SIN(RADIANS(R2C30))+COS(RC[-27])*COS(RADIANS(R2C30))*COS(RC[-25]-RADIANS(R4C30))))))"
    Range("AD6").FormulaR1C1 = "=MIN(R[5]C:R[10004]C)"
    Range("AE11").FormulaR1C1 = "=IF(OR(RC[-1]="""",RC[-1]>20, AND(R10C[-16]>0,RC[-30]>R10C[-16])),"""",RC[-30])"
    Range("AE10").FormulaR1C1 = "=MIN(R[1]C:R[9990]C)"
    Range("AF11").FormulaR1C1 = "=IF(OR(RC[-2]="""",RC[-2]>1, AND(R10C[-17]>0,RC[-31]>R10C[-17])),"""",RC[-31])"
    Range("AG11").FormulaR1C1 = "=IF(RC[-3]="""","""",R4C46-RC[-28])"
    Range("AI11").FormulaR1C1 = "=IF(RC[-5]="""","""",(R8C*SIN(RC[-2])*R8C*SIN(RC[-2]))+(RC[-24]*R7C-RC[-25]*R8C*COS(RC[-2]))*(RC[-24]*R7C-RC[-25]*R8C*COS(RC[-2])))"
    Range("AJ11").FormulaR1C1 = "=IF(RC[-6]="""","""",(RC[-26]*R7C[-1])+(RC[-25]*R8C[-1]*COS(RC[-3])))"
    Range("AK11").FormulaR1C1 = "=IF(RC[-7]="""","""",IF(RC[-2]=0,0,RC[-26]*R8C[-2]*SIN(RC[-4])/SQRT(RC[-2])))"
    Range("AL11").FormulaR1C1 = "=IF(RC[-8]="""","""",RC[-2]-2*RC[-28]*R7C[-3]/(COS(ASIN(RC[-1]))*COS(ASIN(RC[-1]))))"
    Range("AM11").FormulaR1C1 = "=IF(RC[-9]="""","""",R3C3/16*COS(ASIN(RC[-2]))*COS(ASIN(RC[-2]))*(4+R3C3*(4-3*COS(ASIN(RC[-2]))*COS(ASIN(RC[-2])))))"
    Range("AH11").FormulaR1C1 = "=IF(RC[-4]="""","""",RC[-1]+(1-RC[5])*R3C3*RC[3]*(ACOS(RC[2])+RC[5]*SIN(ACOS(RC[2]))*(RC[4]+RC[5]*RC[2]*(-1+2*RC[4]*RC[4]))))"
    Range("AN11").FormulaR1C1 = "=IF(RC[-10]="""","""",IF(AND(RC[-38]=R2C[-10],RC[-36]=R4C[-10]),""samepoint"",""N.A.""))"
    Range("AO11").FormulaR1C1 = "=IF(RC[-11]="""","""",IF(AND(RC[-37]=R4C[-8],R4C[-11]>RC[-39]),""northsouth"",""N.A.""))"
    Range("AP11").FormulaR1C1 = "=IF(RC[-12]="""","""",IF(AND(RC[-38]=R4C[-9],RC[-40]>R4C[-12]),""southnorth"",""N.A.""))"
    Range("AQ11").FormulaR1C1 = "=IF(RC[-13]="""","""",ATAN2((RC[-32]*R7C[-8]-RC[-33]*R8C[-8]*COS(RC[-9])),R8C[-8]*SIN(RC[-9])))"
    Range("AR11").FormulaR1C1 = "=IF(RC[-14]="""","""",IF(RC[-4]=""samepoint"",0,IF(RC[-3]=""northsouth"",0,IF(RC[-2]=""southnorth"",180,IF(RC[-1]<0,RC[-1]*180/PI()+360,RC[-1]*180/PI())))))"
    'Range("AS9").FormulaR1C1 = "=IF(AND(MIN(R[2]C:R[10001]C)>0,MIN(R[2]C:R[10001]C)<R10C15),MIN(R[2]C:R[10001]C),"""")"
    Range("AS9").FormulaR1C1 = "=IF(AND(MIN(R[2]C:R[10001]C)>0,MIN(R[2]C:R[10001]C)<R10C15,MIN(R[2]C[2]:R[10001]C[2])=""""),MIN(R[2]C:R[10001]C),"""")"
    Range("AS11").FormulaR1C1 = "=IF(OR(RC[-1]="""",R[1]C[-1]="""",RC[1]<>"""",R[1]C[1]<>""""),"""",IF(AND(RC[-15]<1,ABS(R[1]C[-1]-RC[-1])>90,ABS(R[1]C[-1]-RC[-1])<180),RC[-44]+(RC[-15]/(RC[-15]+R[1]C[-15]))*(R[1]C1-RC1),""""))"
    ''Range("AS11").FormulaR1C1 = "=IF(OR(RC[-1]="""",R[1]C[-1]="""",RC[1]<>"""",R[1]C[1]<>""""),"""",IF(AND(RC[-15]<1,ABS(R[1]C[-1]-RC[-1])>90,OR(AND(R4C41>=180,R4C41<=270,OR(AND(R[1]C3>=R4C44,R[1]C5<=R4C46),AND(RC3>=R4C44,RC5<=R4C46))),AND(R4C41>=270,R4C41<=360,OR(AND(R[1]C3<=R4C44,R[1]C5<=R4C46),AND(RC3<=R4C44,RC5<=R4C46))),AND(R4C41>=0,R4C41<=90,OR(AND(R[1]C3<=R4C44,R[1]C5>=R4C46),AND(RC3<=R4C44,RC5>=R4C46))),AND(R4C41>=90,R4C41<=180,OR(AND(R[1]C3>=R4C44,R[1]C5>=R4C46),AND(RC3>=R4C44,RC5>=R4C46))))),RC[-44]+(RC[-15]/(RC[-15]+R[1]C[-15]))*(R[1]C1-RC1),""""))"
     Range("AT11").FormulaR1C1 = _
        "=IF(OR(RC[-2]="""",AND(RC[-45]>R5C[1],SUM(R5C[-17]:R7C[-17])>0),RC[-45]>R5C3,RC[-45]>R10C[87]),"""",IF(OR(AND(RC[-44]=R2C[-16],RC[-42]=R4C[-16]),AND(R5C[-3]>R6C[-3],RC[-2]>=R6C[-3],RC[-2]<=R5C[-3]),AND(R6C[-3]>R5C[-3],OR(RC[-2]>=R6C[-3],RC[-2]<=R5C[-3]))),RC[-45],""""))"
    ''Range("AT11").FormulaR1C1 = _
        "=IF(OR(RC[-2]="""",RC[-45]>R5C[1],RC[-45]>R5C3,RC[-45]>R10C[87]),"""",IF(OR(AND(RC[-44]=R2C[-16],RC[-42]=R4C[-16]),AND(R5C[-3]>R6C[-3],RC[-2]>=R6C[-3],RC[-2]<=R5C[-3]),AND(R6C[-3]>R5C[-3],OR(RC[-2]>=R6C[-3],RC[-2]<=R5C[-3]))),RC[-45],""""))"
    'Range("AT11").FormulaR1C1 = _
        "=IF(OR(RC[-2]="""",RC[-45]>MIN(MAX(R4C[40],R6C[41]),MAX(R4C[56],R6C[57]),MAX(R4C[72],R6C[73])),RC[-45]>R5C3,RC[-45]>R10C[87]),"""",IF(OR(AND(RC[-44]=R2C[-16],RC[-42]=R4C[-16]),AND(R5C[-3]>R6C[-3],RC[-2]>=R6C[-3],RC[-2]<=R5C[-3]),AND(R6C[-3]>R5C[-3],OR(RC[-2]>=R6C[-3],RC[-2]<=R5C[-3]))),RC[-45],""""))"
    ''Range("AT11").FormulaR1C1 = _
        "=IF(OR(RC[-2]="""",AND(R4C[-2]=R4C[99],R4C[101]=R4C,RC[-45]>R10C[87]),AND(R4C[-31]>0,RC[-45]>R4C[-31]),AND(R4C[-27]>0,RC[-45]>R4C[-27]),AND(R4C[-23]>0,RC[-45]>R4C[-23]),RC[-45]>R5C3,RC[-45]>R10C[87]),"""",IF(OR(AND(RC[-44]=R2C[-16],RC[-42]=R4C[-16]),AND(R5C[-3]>R6C[-3],RC[-2]>=R6C[-3],RC[-2]<=R5C[-3]),AND(R6C[-3]>R5C[-3],OR(RC[-2]>=R6C[-3],RC[-2]<=R5C[-3]))),RC[-45],""""))"
    Range("AR6").FormulaR1C1 = "=IF(R4C44=0,0,MIN(R[5]C[2]:R[10004]C[2],R[3]C[1]))"
    Range("AT6").FormulaR1C1 = "=IF(R6C44=0,0,MAX(R[5]C:R[10004]C,R[3]C[-1]))"
    Range("AV7").FormulaR1C1 = "=IF(ABS(R6C44)>0,R6C44,0)"
    Range("AV3").FormulaR1C1 = "=IF(OR(SUMMARY!R22C2=""Release"",AND(R4C15>0,R7C48>R4C15),AND(R4C19>0,R7C48>R4C19),AND(R4C23>0,R7C48>R4C23),R7C48=0),""USE RELEASE"",IF(R7C48>0,""SECTOR"",""""))"
    Range("AW7").FormulaR1C1 = "=IF(ABS(R6C44)>0,R6C46,0)"
    Range("AX7").FormulaR1C1 = "=IF(MIN(R[4]C:R[10003]C)>0,MIN(R[4]C:R[10003]C),0)"
    Range("AX9").FormulaR1C1 = "=IF(MIN(R[2]C[20]:R[10001]C[20])>0,MIN(R[2]C[20]:R[10001]C[20]),0)"
    Range("AU11").FormulaR1C1 = "=IF(RC[-1]="""","""",RC[-15])"
    Range("AU8").FormulaR1C1 = "=MIN(R[3]C:R[10002]C,R[1]C[-2])"
    Range("AU10").FormulaR1C1 = "=MAX(R[1]C:R[10000]C,R[-1]C[-2])"
    'Start before any TP
    Range("AV11").FormulaR1C1 = "=IF(AND(R5C45<>3,RC[-3]<>""""),(RC[-42]+R[1]C[-42])/2,IF(AND(R5C45=3,RC[-3]<>""""),(RC[-41]+R[1]C[-41])/2,IF(RC[-2]="""","""",IF(R5C45<>3,RC[-42],RC[-41]))))"
    'Range("AV11").FormulaR1C1 = "=IF(RC[-3]<>"""",(RC[-42]+R[1]C[-42])/2,IF(RC[-2]="""","""",RC[-42]))"
    ''Range("AV11").FormulaR1C1 = "=IF(RC[-47]>R7C[1],"""",IF(RC[-3]<>"""",(RC[-42]+R[1]C[-42])/2,IF(RC[-2]="""","""",RC[-42])))"
    '''Range("AV11").FormulaR1C1 = "=IF(RC[-47]>R7C[1],"""",IF(RC[-3]<>"""",(RC[-42]+R[1]C[-42])/2,IF(RC[-2]="""","""",RC[-42])))"
    Range("AV9").FormulaR1C1 = "=IF(AND(ABS(R6C[-4])>0,MIN(R[2]C:R[10001]C)>0),MIN(R[2]C:R[10001]C),0)"
    Range("AW11").FormulaR1C1 = "=IF(AND(RC[-4]="""",RC[-3]=""""),"""",IF(RC[-1]=R9C[-1],RC[-48],""""))"
    'Range("AW11").FormulaR1C1 = "=IF(AND(RC[-4]<>"""",RC[-4]=R7C49),RC[-4],IF(AND(RC[-3]<>"""",RC[-1]=RC[-43]),RC[-48],""""))"
    Range("AW9").FormulaR1C1 = "=MAX(R[2]C:R[10001]C)"
    Range("AX11").FormulaR1C1 = "=IF(AND(RC[-5]<>"""",RC[-20]<=1,RC[-2]=R9C[20]),(RC[-49]+R[1]C[-49])/2,IF(AND(RC[-4]<>"""",RC[-20]<=1,RC[-2]=R9C[20]),RC[-49],""""))"
    ''Range("AX11").FormulaR1C1 = "=IF(AND(RC[-5]<>"""",RC[-20]<=1),(RC[-49]+R[1]C[-49])/2,IF(AND(RC[-4]<>"""",RC[-20]<=1),RC[-49],""""))"
    Range("AX3").FormulaR1C1 = "=IF(OR(R9C51=0,R34C176=0),""No Start Line"",""Start Line OK"")"
    Range("AY9").FormulaR1C1 = "=IF(R4C46=0,0,IF(OR(RC[-6]=R[-1]C[-4],RC[-6]=R[1]C[-4]),MIN(R11C:R[1000]C),IF(OR(R15C176+90<360,R15C178>R15C179,AND(R15C176+90>360,R15C178<R15C179)),MAX(R11C51:R10009C51),MIN(R11C51:R10009C51))))"
    Range("AY10").FormulaR1C1 = "=RC[-4]"
    Range("AY11").FormulaR1C1 = "=IF(OR(AND(R[2]C[-7]=0,OR(R9C45=R7C48,R9C45=R7C49)),RC[-21]>0.66,AND(R4C86>0,RC[-50]>R4C86),AND(R4C102>0,RC[-50]>R4C102),AND(R4C118>0,RC[-50]>R4C118)),"""",IF(OR(AND(R15C179>R15C178,OR(RC[-7]>R15C179,RC[-7]<R15C178)),AND(R15C179<R15C178,RC[-7]<R15C178,RC[-7]>R15C179)),RC[-50],""""))"
    Range("AZ11").FormulaR1C1 = "=IF(RC[-1]=R9C[-1],RC[-50],"""")"
    Range("AZ9").FormulaR1C1 = "=MIN(R[2]C:R[10001]C)"
    Range("BA11").FormulaR1C1 = "=IF(AND(R10C[-2]>0,RC[-52]=R10C[-2]),RC[-51],"""")"
    Range("BA9").FormulaR1C1 = "=MIN(R[2]C:R[10001]C)"
    Range("BB11").FormulaR1C1 = "=IF(RC[-2]="""","""",RC[-50])"
    Range("BB9").FormulaR1C1 = "=MIN(R[2]C:R[10001]C)"
    Range("BC11").FormulaR1C1 = "=IF(RC[-2]="""","""",RC[-51])"
    Range("BC9").FormulaR1C1 = "=MIN(R[2]C:R[10001]C)"
    Range("BD11").FormulaR1C1 = "=IF(RC[-2]="""","""",IF(R5C45<>3,RC[-50],RC[-49]))"
    'Range("BD11").FormulaR1C1 = "=IF(RC[-2]="""","""",RC[-50])"
    Range("BD9").FormulaR1C1 = "=MIN(R[2]C:R[9991]C)"
    Range("BE11").FormulaR1C1 = "=IF(RC[-2]="""","""",IF(R5C45<>3,RC[-51],RC[-50]))"
    'Range("BE11").FormulaR1C1 = "=IF(RC[-2]="""","""",RC[-51])"
    Range("BE9").FormulaR1C1 = "=MIN(R[2]C:R[10001]C)"
    Range("BF11").FormulaR1C1 = "=IF(RC[-2]="""","""",RC[-14])"
    Range("BF9").FormulaR1C1 = "=MIN(R[2]C:R[10001]C)"
    Range("BG11").FormulaR1C1 = "=IF(RC[-2]="""","""",RC[-15])"
    Range("BG9").FormulaR1C1 = "=MIN(R[2]C:R[10001]C)"
    Range("BH11").FormulaR1C1 = "=IF(AND(R9C[-9]>0,RC[-59]=R9C[-9]),R[1]C[-59],"""")"
    Range("BH9").FormulaR1C1 = "=MAX(R[2]C:R[10001]C)"
    Range("BI11").FormulaR1C1 = "=IF(AND(R10C[-10]>0,RC[-60]=R10C[-10]),R[1]C[-60],"""")"
    Range("BI9").FormulaR1C1 = "=MIN(R[2]C:R[10001]C)"
    Range("BJ11").FormulaR1C1 = "=IF(RC[-2]=R9C[-2],R[1]C[-60],"""")"
    Range("BJ9").FormulaR1C1 = "=MIN(R[2]C:R[10001]C)"
    Range("BK11").FormulaR1C1 = "=IF(RC[-2]="""","""",R[1]C[-61])"
    Range("BK9").FormulaR1C1 = "=MIN(R[2]C:R[10001]C)"
    Range("BL11").FormulaR1C1 = "=IF(RC[-2]="""","""",R[1]C[-60])"
    Range("BL9").FormulaR1C1 = "=MIN(R[2]C:R[10001]C)"
    Range("BM11").FormulaR1C1 = "=IF(RC[-2]="""","""",R[1]C[-61])"
    Range("BM9").FormulaR1C1 = "=MIN(R[2]C:R[10001]C)"
    Range("BN11").FormulaR1C1 = "=IF(RC[-2]="""","""",IF(R5C45<>3,R[1]C[-60],R[1]C[-59]))"
    'Range("BN11").FormulaR1C1 = "=IF(RC[-2]="""","""",R[1]C[-60])"
    Range("BN9").FormulaR1C1 = "=MIN(R[2]C:R[10001]C)"
    Range("BO11").FormulaR1C1 = "=IF(RC[-2]="""","""",IF(R5C45<>3,R[1]C[-61],R[1]C[-60]))"
    'Range("BO11").FormulaR1C1 = "=IF(RC[-2]="""","""",R[1]C[-61])"
    Range("BO9").FormulaR1C1 = "=MIN(R[2]C:R[10001]C)"
    Range("BP11").FormulaR1C1 = "=IF(RC[-2]="""","""",R[1]C[-24])"
    Range("BP9").FormulaR1C1 = "=MIN(R[2]C:R[10001]C)"
    Range("BQ11").FormulaR1C1 = "=IF(RC[-2]="""","""",R[1]C[-25])"
    Range("BQ9").FormulaR1C1 = "=MIN(R[2]C:R[10001]C)"
    Range("BJ5").FormulaR1C1 = "=IF(YDWK2!R951C16=0,""N/A"",YDWK2!R951C16)"
    Range("BJ6").FormulaR1C1 = "=YDWK2!R1242C6"
    Range("BD4").FormulaR1C1 = "=IF(YDWK2!R951C26=0,""N/A"",YDWK2!R951C26)"
    Range("BD5").FormulaR1C1 = "=YDWK2!R1024C6"
    Range("AZ5").FormulaR1C1 = "=IF(OR(YDWK2!R951C6>0,AND(R[4]C[1]=R[-3]C[-22],R[-1]C[-22]=R[4]C[3])),YDWK2!R951C6,""N/A"")"
    'Range("AZ5").FormulaR1C1 = "=IF(YDWK2!R951C6=0,""N/A"",YDWK2!R951C6)"
    Range("AZ6").FormulaR1C1 = "=YDWK2!R1169C6"
    Range("AX3").FormulaR1C1 = "=IF(OR(R9C51=0,R34C176=0),""No Start Line"",""Start Line OK"")"
    Range("AV3").FormulaR1C1 = "=IF(OR(SUMMARY!R22C2=""Release"",AND(R4C15>0,R7C48>R4C15),AND(R4C19>0,R7C48>R4C19),AND(R4C23>0,R7C48>R4C23),R7C48=0),""USE RELEASE"",IF(R7C48>0,""SECTOR"",""""))"
    Range("BS5").FormulaR1C1 = "=IF(AND(MIN(R[6]C[-25]:R[10005]C[-25])>0,R11C1>=MIN(R[6]C[-25]:R[10005]C[-25])),""Release in ST OZ"",""Release not in OZ"")"
    Range("BR9").FormulaR1C1 = "=MIN(R[2]C:R[10001]C)"
    Range("BR11").FormulaR1C1 = "=IF(OR(RC[-24]="""",RC[-40]>1),"""",RC[-22])"
    ''Range("BR11").FormulaR1C1 = "=IF(RC[-20]="""","""",RC[-22])"
    Range("BS11").FormulaR1C1 = "=IF(AND(RC[-70]=R10C47,R5C45<>3),RC[-65],IF(AND(RC[-70]=R10C47,R5C45=3),RC[-64],""""))"
    'Range("BS11").FormulaR1C1 = "=IF(RC[-70]=R10C47,RC[-65],"""")"
    Range("BS6").FormulaR1C1 = "=R4C3"
    Range("BS9").FormulaR1C1 = "=MAX(R[2]C:R[10001]C)"
    Range("BU4").FormulaR1C1 = "=B!R22C1+B!R22C2/60"
    Range("BX4").FormulaR1C1 = "=B!R22C4+B!R22C5/60"
    Range("CF4").FormulaR1C1 = "=RADIANS(R4C73)"
    Range("CG4").FormulaR1C1 = "=RADIANS(R4C76)"
    Range("BW5").FormulaR1C1 = "=(1-R3C3)*TAN(R4C84)"
    Range("BW6").FormulaR1C1 = "=ATAN(R[-1]C)"
    Range("BW7").FormulaR1C1 = "=SIN(R[-1]C)"
    Range("BW8").FormulaR1C1 = "=COS(R[-2]C)"
    Range("CC4").FormulaR1C1 = _
        "=IF(AND(R5C[-52]=1,SUM(R[1]C[-52]:R[3]C[-52])>1),R21C[-55],IF(AND(R5C[-52]=2,SUM(R[1]C[-52]:R[3]C[-52])=6),R27C[-55],IF(OR(AND(R5C[-52]=1,SUM(R[1]C[-52]:R[3]C[-52])=1),AND(R5C[-52]=2,SUM(R[1]C[-52]:R[3]C[-52])=3),AND(R5C[-52]=3,SUM(R[1]C[-52]:R[3]C[-52])=6)),R33C[-55],"""")))"
    Range("CC5").FormulaR1C1 = "=IF(R[-1]C="""","""",R[-1]C+45)"
    Range("CE5").FormulaR1C1 = "=IF(RC[-2]="""","""",IF(R[-1]C[-2]+45>360,R[-1]C[-2]+45-360,RC[-2]))"
    Range("CC6").FormulaR1C1 = "=IF(R[-2]C="""","""",R[-2]C-45)"
    Range("CE6").FormulaR1C1 = "=IF(R[-1]C="""","""",IF(R[-2]C[-2]-45<0,R[-2]C[-2]-45+360,RC[-2]))"
    Range("CI1").FormulaR1C1 = "=R[1]C"
    'Range("CI1").FormulaR1C1 = "=IF(R[2]C<>"""",R[1]C,0)"
    'Range("CI1").FormulaR1C1 = "=IF(MAX(R[3]C[-1],R[5]C[-1],R[5]C)>0,R[4]C[-58],0)"
    Range("CI2").FormulaR1C1 = "=R[3]C[-58]"
 LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("AD11:BS11").AutoFill Destination:=.Range("AD11:BS" & LastRow), Type:=xlFillDefault
End If
 
 If Range("A4") > 0 Then
    Range("BU11").FormulaR1C1 = "=IF(AND(RC[-61]>40,RC[-60]="""",RC[-58]=""""),"""",R4C[12]-RC[-68])"
    'Range("BU11").FormulaR1C1 = "=IF(AND(RC[-60]="""",RC[-58]=""""),"""",R4C[12]-RC[-68])"
    Range("BV11").FormulaR1C1 = "=IF(RC[-1]="""","""",RC[-1]+(1-RC[5])*R3C3*RC[3]*(ACOS(RC[2])+RC[5]*SIN(ACOS(RC[2]))*(RC[4]+RC[5]*RC[2]*(-1+2*RC[4]*RC[4]))))"
    Range("BW11").FormulaR1C1 = "=IF(RC[-2]="""","""",(R8C*SIN(RC[-2])*R8C*SIN(RC[-2]))+(RC[-64]*R7C-RC[-65]*R8C*COS(RC[-2]))*(RC[-64]*R7C-RC[-65]*R8C*COS(RC[-2])))"
    Range("BX11").FormulaR1C1 = "=IF(RC[-3]="""","""",(RC[-66]*R7C[-1])+(RC[-65]*R8C[-1]*COS(RC[-3])))"
    Range("BY11").FormulaR1C1 = "=IF(RC[-4]="""","""",IF(RC[-2]=0,0,RC[-66]*R8C[-2]*SIN(RC[-4])/SQRT(RC[-2])))"
    Range("BZ11").FormulaR1C1 = "=IF(RC[-5]="""","""",RC[-2]-2*RC[-68]*R7C[-3]/(COS(ASIN(RC[-1]))*COS(ASIN(RC[-1]))))"
    Range("CA11").FormulaR1C1 = "=IF(RC[-6]="""","""",R3C3/16*COS(ASIN(RC[-2]))*COS(ASIN(RC[-2]))*(4+R3C3*(4-3*COS(ASIN(RC[-2]))*COS(ASIN(RC[-2])))))"
    Range("CB11").FormulaR1C1 = "=IF(RC[-7]="""","""",IF(AND(RC[-78]=R4C[-7],RC[-76]=R4C[-4]),""samepoint"",""N.A.""))"
    Range("CC11").FormulaR1C1 = "=IF(RC[-8]="""","""",IF(AND(RC[-77]=R4C[-5],R4C[-8]>RC[-79]),""northsouth"",""N.A.""))"
    Range("CD11").FormulaR1C1 = "=IF(RC[-9]="""","""",IF(AND(RC[-78]=R4C[-6],RC[-80]>R4C[-9]),""southnorth"",""N.A.""))"
    Range("CE11").FormulaR1C1 = "=IF(RC[-10]="""","""",ATAN2((RC[-72]*R7C[-8]-RC[-73]*R8C[-8]*COS(RC[-9])),R8C[-8]*SIN(RC[-9])))"
    Range("CF11").FormulaR1C1 = "=IF(RC[-11]="""","""",IF(RC[-4]=""samepoint"",0,IF(RC[-3]=""northsouth"",0,IF(RC[-2]=""southnorth"",180,IF(RC[-1]<0,RC[-1]*180/PI()+360,RC[-1]*180/PI())))))"
    Range("CG11").FormulaR1C1 = "=IF(OR(RC[-1]="""",R[1]C[-1]="""",RC[1]<>"""",R[1]C[1]<>""""),"""",IF(AND(ABS(R[1]C[-1]-RC[-1])>90,ABS(R[1]C[-1]-RC[-1])<180),RC[-84]+(RC12/(RC12+R[1]C12))*(R[1]C1-RC1),""""))"
    ''Range("CG11").FormulaR1C1 = "=IF(OR(RC[-1]="""",R[1]C[-1]="""",RC[1]<>"""",R[1]C[1]<>""""),"""",IF(AND(ABS(R[1]C[-1]-RC[-1])>90,OR(AND(R4C81>=180,R4C81<=270,OR(AND(R[1]C3>=R4C84,R[1]C5<=R4C85),AND(RC3>=R4C84,RC5<=R4C85))),AND(R4C81>=270,R4C81<=360,OR(AND(R[1]C3<=R4C84,R[1]C5<=R4C85),AND(RC3<=R4C84,RC5<=R4C85))),AND(R4C81>=0,R4C81<=90,OR(AND(R[1]C3<=R4C84,R[1]C5>=R4C85),AND(RC3<=R4C84,RC5>=R4C85))),AND(R4C81>=90,R4C81<=180,OR(AND(R[1]C3>=R4C84,R[1]C5>=R4C85),AND(RC3>=R4C84,RC5>=R4C85))))),RC[-84]+(RC12/(RC12+R[1]C12))*(R[1]C1-RC1),""""))"
    Range("CG9").FormulaR1C1 = "=MAX(R[2]C:R[10001]C)"
    Range("CH11").FormulaR1C1 = "=IF(OR(R4C[-2]=0,R4C[-1]=0,RC[-85]>R5C3,RC[-2]=""""),"""",IF(OR(AND(RC[-84]=R4C[-13],RC[-82]=R4C[-10]),AND(R5C[-3]>R6C[-3],RC[-2]>=R6C[-3],RC[-2]<=R5C[-3]),AND(R5C[-3]<R6C[-3],OR(RC[-2]>=R6C[-3],RC[-2]<=R5C[-3]))),RC[-85],""""))"
    'Range("CH11").FormulaR1C1 = "=IF(OR(RC[-85]<R10C[-71],RC[-85]>R4C[-71],R4C[-2]=0,R4C[-1]=0,RC[-85]>R5C3,RC[-2]=""""),"""",IF(OR(AND(RC[-84]=R4C[-13],RC[-82]=R4C[-10]),AND(R5C[-3]>R6C[-3],RC[-2]>=R6C[-3],RC[-2]<=R5C[-3]),AND(R5C[-3]<R6C[-3],OR(RC[-2]>=R6C[-3],RC[-2]<=R5C[-3]))),RC[-85],""""))"
    Range("CH4").FormulaR1C1 = "=IF(MIN(R[7]C:R[10006]C)=0,R[5]C[-1],MIN(R[7]C:R[10006]C))"
    Range("CH6").FormulaR1C1 = "=IF(AND(R[3]C[-1]<>0,OR(R[-2]C=R[3]C[-1],AND(MAX(R6C148:R6C149)>0,MAX(R[5]C:R[10004]C)>MAX(R6C148:R6C149)))),R[3]C[-1],MAX(R[5]C:R[10004]C))"
    Range("CI11").FormulaR1C1 = "=IF(RC[-75]="""","""",IF(AND(R2C=1,RC[-75]<=0.5),RC[-86],IF(AND(R2C=2,RC[-86]>MIN(R10C[-72],R10C[-68],R10C[-64]),RC[-75]<=0.5),RC[-86],IF(AND(R2C=3,RC[-86]>MAX(R10C[-72],R10C[-68],R10C[-64]),RC[-75]<=0.5),RC[-86],""""))))"
    Range("CI6").FormulaR1C1 = "=MIN(R[5]C:R[10004]C)"
    Range("CI7").FormulaR1C1 = "=MAX(R[4]C:R[10003]C)"
    Range("CI3").FormulaR1C1 = _
        "=IF(OR(R[-1]C=0,AND(R[1]C[-1]=0,R[3]C[-1]=0,R[3]C=0)),"""",IF(R[1]C[-1]>0,""SECTOR"",""CYLINDER""))"
    Range("CK4").FormulaR1C1 = "=B!R24C1+B!R24C2/60"
    Range("CN4").FormulaR1C1 = "=B!R24C4+B!R24C5/60"
    Range("CV4").FormulaR1C1 = "=RADIANS(RC89)"
    Range("CW4").FormulaR1C1 = "=RADIANS(R4C92)"
    Range("CM5").FormulaR1C1 = "=(1-R3C3)*TAN(R4C100)"
    Range("CM6").FormulaR1C1 = "=ATAN(R[-1]C)"
    Range("CM7").FormulaR1C1 = "=SIN(R[-1]C)"
    Range("CM8").FormulaR1C1 = "=COS(R[-2]C)"
    Range("CS4").FormulaR1C1 = "=IF(AND(R6C[-68]=1,SUM(R5C29:R7C29)>1),R21C26,IF(AND(R6C[-68]=2,SUM(R5C29:R7C29)=6),R27C26,IF(OR(AND(R6C29=1,SUM(R5C29:R7C29)=1),AND(R6C29=2,SUM(R5C29:R7C29)=3),AND(R6C29=3,SUM(R5C29:R7C29)=6)),R33C26,"""")))"
    Range("CS5").FormulaR1C1 = "=IF(R[-1]C="""","""",R[-1]C+45)"
    Range("CU5").FormulaR1C1 = "=IF(R[-1]C[-2]="""","""",IF(RC[-2]>360,R[-1]C[-2]+45-360,RC[-2]))"
    Range("CS6").FormulaR1C1 = "=IF(R[-2]C="""","""",R[-2]C-45)"
    Range("CU6").FormulaR1C1 = "=IF(R[-1]C="""","""",IF(RC[-2]<0,R[-2]C[-2]-45+360,R[-2]C[-2]-45))"
    Range("CY1").FormulaR1C1 = "=R[1]C"
    'Range("CY1").FormulaR1C1 = "=IF(R[2]C<>"""",R[1]C,0)"
    'Range("CY1").FormulaR1C1 = "=IF(MAX(R[3]C[-1],R[5]C[-1],R[5]C)>0,R[5]C[-74],0)"
    Range("CY2").FormulaR1C1 = "=R[4]C[-74]"
LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("BU11:CI11").AutoFill Destination:=.Range("BU11:CI" & LastRow), Type:=xlFillDefault

End If

If Range("A5") > 0 Then
    Range("CK11").FormulaR1C1 = "=IF(AND(RC[-73]>40,RC[-72]="""",RC[-70]=""""),"""",R4C[12]-RC[-84])"
    'Range("CK11").FormulaR1C1 = "=IF(AND(RC[-72]="""",RC[-70]=""""),"""",R4C[12]-RC[-84])"
    Range("CL11").FormulaR1C1 = "=IF(RC[-1]="""","""",RC[-1]+(1-RC[5])*R3C3*RC[3]*(ACOS(RC[2])+RC[5]*SIN(ACOS(RC[2]))*(RC[4]+RC[5]*RC[2]*(-1+2*RC[4]*RC[4]))))"
    Range("CM11").FormulaR1C1 = "=IF(RC[-2]="""","""",(R8C*SIN(RC[-2])*R8C*SIN(RC[-2]))+(RC[-80]*R7C-RC[-81]*R8C*COS(RC[-2]))*(RC[-80]*R7C-RC[-81]*R8C*COS(RC[-2])))"
    Range("CN11").FormulaR1C1 = "=IF(RC[-3]="""","""",(RC[-82]*R7C[-1])+(RC[-81]*R8C[-1]*COS(RC[-3])))"
    Range("CO11").FormulaR1C1 = "=IF(RC[-4]="""","""",IF(RC[-2]=0,0,RC[-82]*R8C[-2]*SIN(RC[-4])/SQRT(RC[-2])))"
    Range("CP11").FormulaR1C1 = "=IF(RC[-5]="""","""",RC[-2]-2*RC[-84]*R7C[-3]/(COS(ASIN(RC[-1]))*COS(ASIN(RC[-1]))))"
    Range("CQ11").FormulaR1C1 = "=IF(RC[-6]="""","""",R3C3/16*COS(ASIN(RC[-2]))*COS(ASIN(RC[-2]))*(4+R3C3*(4-3*COS(ASIN(RC[-2]))*COS(ASIN(RC[-2])))))"
    Range("CR11").FormulaR1C1 = "=IF(RC[-7]="""","""",IF(AND(RC[-94]=R4C[-7],RC[-92]=R4C[-4]),""samepoint"",""N.A.""))"
    Range("CS11").FormulaR1C1 = "=IF(RC[-8]="""","""",IF(AND(RC[-93]=R4C[-8],R4C[-5]>RC[-95]),""northsouth"",""N.A.""))"
    Range("CT11").FormulaR1C1 = "=IF(RC[-9]="""","""",IF(AND(RC[-94]=R4C[-6],RC[-96]>R4C[-9]),""southnorth"",""N.A.""))"
    Range("CU11").FormulaR1C1 = "=IF(RC[-10]="""","""",ATAN2((RC[-88]*R7C[-8]-RC[-89]*R8C[-8]*COS(RC[-9])),R8C[-8]*SIN(RC[-9])))"
    Range("CV11").FormulaR1C1 = "=IF(RC[-11]="""","""",IF(RC[-4]=""samepoint"",0,IF(RC[-3]=""northsouth"",0,IF(RC[-2]=""southnorth"",180,IF(RC[-1]<0,RC[-1]*180/PI()+360,RC[-1]*180/PI())))))"
    Range("CW11").FormulaR1C1 = "=IF(OR(RC[-1]="""",R[1]C[-1]="""",RC[1]<>"""",R[1]C[1]<>""""),"""",IF(AND(ABS(R[1]C[-1]-RC[-1])>90,ABS(R[1]C[-1]-RC[-1])<180),RC[-100]+(RC16/(RC16+R[1]C16))*(R[1]C1-RC1),""""))"
    ''Range("CW11").FormulaR1C1 = "=IF(OR(RC[-1]="""",R[1]C[-1]="""",RC[1]<>"""",R[1]C[1]<>""""),"""",IF(AND(ABS(R[1]C[-1]-RC[-1])>90,OR(AND(R4C97>=180,R4C97<=270,OR(AND(R[1]C3>=R4C100,R[1]C5<=R4C101),AND(RC3>=R4C100,RC5<=R4C101))),AND(R4C97>=270,R4C97<=360,OR(AND(R[1]C3<=R4C100,R[1]C5<=R4C101),AND(RC3<=R4C100,RC5<=R4C101))),AND(R4C97>=0,R4C97<=90,OR(AND(R[1]C3<=R4C100,R[1]C5>=R4C101),AND(RC3<=R4C100,RC5>=R4C101))),AND(R4C97>=90,R4C97<=180,OR(AND(R[1]C3>=R4C100,R[1]C5>=R4C101),AND(RC3>=R4C100,RC5>=R4C101))))),RC[-100]+(RC16/(RC16+R[1]C16))*(R[1]C1-RC1),""""))"
    Range("CW9").FormulaR1C1 = "=MAX(R[2]C:R[10001]C)"
    Range("CX11").FormulaR1C1 = _
        "=IF(OR(RC[-101]<R4C[-16],RC[-101]>R5C3,RC[-2]=""""),"""",IF(OR(AND(RC[-100]=R4C[-13],RC[-98]=R4C[-10]),AND(R5C[-3]>R6C[-3],RC[-2]>=R6C[-3],RC[-2]<=R5C[-3]),AND(R5C[-3]<R6C[-3],OR(RC[-2]>=R6C[-3],RC[-2]<=R5C[-3]))),RC[-101],""""))"
    ''Range("CX11").FormulaR1C1 = _
        "=IF(OR(RC[-101]<R4C[-16],RC[-101]>R5C3,AND(R2C129<>"""",RC[-101]>R2C129),RC[-2]=""""),"""",IF(OR(AND(RC[-100]=R4C[-13],RC[-98]=R4C[-10]),AND(R5C[-3]>R6C[-3],RC[-2]>=R6C[-3],RC[-2]<=R5C[-3]),AND(R5C[-3]<R6C[-3],OR(RC[-2]>=R6C[-3],RC[-2]<=R5C[-3]))),RC[-101],""""))"
    '''Sheets("TPOrder").Range("CW6").Value = Sheets("B").Range("E12").Value
    '''Range("CX11").FormulaR1C1 = _
        "=IF(OR(AND(R6C[-1]<>""ST@TP"",RC[-101]<R6C[-16],R4C[-2]=R4C[-58],R4C[-56]=R4C[-1]),AND(R6C[-1]=""ST@TP"",RC[-101]<R4C[-16],R4C[-2]=R4C[-58],R4C[-56]=R4C[-1]),AND(R4C[-2]=0,R4C[-1]=0),RC[-101]>R5C3,AND(R2C129<>"""",RC[-101]>R2C129),RC[-2]=""""),"""",IF(OR(AND(RC[-100]=R4C[-13],RC[-98]=R4C[-10]),AND(R5C[-3]>R6C[-3],RC[-2]>=R6C[-3],RC[-2]<=R5C[-3]),AND(R5C[-3]<R6C[-3],OR(RC[-2]>=R6C[-3],RC[-2]<=R5C[-3]))),RC[-101],""""))"
    ''Range("CX11").FormulaR1C1 = "=IF(OR(AND(RC[-101]<R6C[-16],R4C[-2]=R4C[-58],R4C[-56]=R4C[-1]),AND(R4C[-2]=0,R4C[-1]=0),RC[-101]>R5C3,AND(R2C129<>"""",RC[-101]>R2C129),RC[-2]=""""),"""",IF(OR(AND(RC[-100]=R4C[-13],RC[-98]=R4C[-10]),AND(R5C[-3]>R6C[-3],RC[-2]>=R6C[-3],RC[-2]<=R5C[-3]),AND(R5C[-3]<R6C[-3],OR(RC[-2]>=R6C[-3],RC[-2]<=R5C[-3]))),RC[-101],""""))"
    ''Range("CX11").FormulaR1C1 = "=IF(OR(AND(RC[-101]<R6C[-16],R4C[-2]=R4C[-58],R4C[-56]=R4C[-1]),AND(R4C[-2]=0,R4C[-1]=0),RC[-101]>R5C3,RC[-2]=""""),"""",IF(OR(AND(RC[-100]=R4C[-13],RC[-98]=R4C[-10]),AND(R5C[-3]>R6C[-3],RC[-2]>=R6C[-3],RC[-2]<=R5C[-3]),AND(R5C[-3]<R6C[-3],OR(RC[-2]>=R6C[-3],RC[-2]<=R5C[-3]))),RC[-101],""""))"
    ''Range("CX11").FormulaR1C1 = "=IF(OR(AND(RC[-101]<R6C[-16],R4C[-2]=R4C[-58],R4C[-56]=R4C[-1]),AND(R4C[-2]=0,R4C[-1]=0),AND(R4C19>R8C19,R9C26=""X"",RC1>R8C19),RC[-2]=""""),"""",IF(OR(AND(RC[-100]=R4C[-13],RC[-98]=R4C[-10]),AND(R5C[-3]>R6C[-3],RC[-2]>=R6C[-3],RC[-2]<=R5C[-3]),AND(R5C[-3]<R6C[-3],OR(RC[-2]>=R6C[-3],RC[-2]<=R5C[-3]))),RC[-101],""""))"
    ''Range("CX11").FormulaR1C1 = "=IF(OR(AND(RC[-101]<R6C[-16],R4C[-2]=R4C[-58],R4C[-56]=R4C[-1]),AND(R4C[-2]=0,R4C[-1]=0),RC[-101]>R5C3,RC[-2]=""""),"""",IF(OR(AND(RC[-100]=R4C[-13],RC[-98]=R4C[-10]),AND(R5C[-3]>R6C[-3],RC[-2]>=R6C[-3],RC[-2]<=R5C[-3]),AND(R5C[-3]<R6C[-3],OR(RC[-2]>=R6C[-3],RC[-2]<=R5C[-3]))),RC[-101],""""))"
    Range("CX4").FormulaR1C1 = "=IF(MIN(R[7]C:R[10006]C)=0,R[5]C[-1],MIN(R[7]C:R[10006]C))"
    Range("CX6").FormulaR1C1 = "=IF(AND(R[3]C[-1]<>0,OR(R[-2]C=R[3]C[-1],AND(MAX(R6C148:R6C149)>0,MAX(R[5]C:R[10004]C)>MAX(R6C148:R6C149)))),R[3]C[-1],MAX(R[5]C:R[10004]C))"
    Range("CY11").FormulaR1C1 = "=IF(RC[-87]="""","""",IF(AND(R2C=1,RC[-87]<=0.5),RC[-102],IF(AND(R2C=2,RC[-102]>MIN(R10C[-84],R10C[-80],R10C[-88]),RC[-87]<=0.5),RC[-102],IF(AND(R2C=3,RC[-102]>MAX(R10C[-84],R10C[-80],R10C[-88]),RC[-87]<=0.5),RC[-102],""""))))"
    Range("CY6").FormulaR1C1 = "=MIN(R[5]C:R[10004]C)"
    Range("CY7").FormulaR1C1 = "=MAX(R[4]C:R[10003]C)"
    Range("CY3").FormulaR1C1 = _
        "=IF(OR(R[-1]C=0,AND(R[1]C[-1]=0,R[3]C[-1]=0,R[3]C=0)),"""",IF(R[1]C[-1]>0,""SECTOR"",""CYLINDER""))"
    Range("DA4").FormulaR1C1 = "=B!R26C1+B!R26C2/60"
    Range("DD4").FormulaR1C1 = "=B!R26C4+B!R26C5/60"
    Range("DL4").FormulaR1C1 = "=IF(OR(SUMMARY!R[26]C[-112]="""",SUMMARY!R[28]C[-112]=""""),"""",RADIANS(R4C105))"
    Range("DM4").FormulaR1C1 = "=IF(OR(SUMMARY!R[26]C[-113]="""",SUMMARY!R[28]C[-113]=""""),"""",RADIANS(R4C108))"
    Range("DC5").FormulaR1C1 = "=(1-R3C3)*TAN(R4C116)"
    Range("DC6").FormulaR1C1 = "=ATAN(R[-1]C)"
    Range("DC7").FormulaR1C1 = "=SIN(R[-1]C)"
    Range("DC8").FormulaR1C1 = "=COS(R[-2]C)"
    Range("DI4").FormulaR1C1 = "=IF(AND(R7C[-84]=1,SUM(R5C29:R7C29)>1),R21C26,IF(AND(R7C[-84]=2,SUM(R5C29:R7C29)=6),R27C26,IF(OR(AND(R7C29=1,SUM(R5C29:R7C29)=1),AND(R7C29=2,SUM(R5C29:R7C29)=3),AND(R7C29=3,SUM(R5C29:R7C29)=6)),R33C26,"""")))"
    Range("DI5").FormulaR1C1 = "=IF(R[-1]C="""","""",R[-1]C+45)"
    Range("DK5").FormulaR1C1 = "=IF(RC[-2]="""","""",IF(R[-1]C[-2]+45>360,R[-1]C[-2]+45-360,RC[-2]))"
    Range("DI6").FormulaR1C1 = "=IF(R[-2]C="""","""",R[-2]C-45)"
    Range("DK6").FormulaR1C1 = "=IF(R[-1]C="""","""",IF(R[-2]C[-2]-45<0,R[-2]C[-2]-45+360,RC[-2]))"
    Range("DO1").FormulaR1C1 = "=R[1]C"
    'Range("DO1").FormulaR1C1 = "=IF(R[2]C<>"""",R[1]C,0)"
    'Range("DO1").FormulaR1C1 = "=IF(MAX(R[3]C[-1],R[5]C[-1],R[5]C)>0,R[6]C[-90],0)"
    Range("DO2").FormulaR1C1 = "=R[5]C[-90]"
LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("CK11:CY11").AutoFill Destination:=.Range("CK11:CY" & LastRow), Type:=xlFillDefault
End If

If Range("A6") > 0 Then
    Range("DA11").FormulaR1C1 = "=IF(AND(RC[-85]>40,RC[-84]="""",RC[-82]=""""),"""",R4C[12]-RC[-100])"
    'Range("DA11").FormulaR1C1 = "=IF(AND(RC[-84]="""",RC[-82]=""""),"""",R4C[12]-RC[-100])"
    Range("DB11").FormulaR1C1 = "=IF(RC[-1]="""","""",RC[-1]+(1-RC[5])*R3C3*RC[3]*(ACOS(RC[2])+RC[5]*SIN(ACOS(RC[2]))*(RC[4]+RC[5]*RC[2]*(-1+2*RC[4]*RC[4]))))"
    Range("DC11").FormulaR1C1 = "=IF(RC[-2]="""","""",(R8C*SIN(RC[-2])*R8C*SIN(RC[-2]))+(RC[-96]*R7C-RC[-97]*R8C*COS(RC[-2]))*(RC[-96]*R7C-RC[-97]*R8C*COS(RC[-2])))"
    Range("DD11").FormulaR1C1 = "=IF(RC[-3]="""","""",(RC[-98]*R7C[-1])+(RC[-97]*R8C[-1]*COS(RC[-3])))"
    Range("DE11").FormulaR1C1 = "=IF(RC[-4]="""","""",IF(RC[-2]=0,0,RC[-98]*R8C[-2]*SIN(RC[-4])/SQRT(RC[-2])))"
    Range("DF11").FormulaR1C1 = "=IF(RC[-5]="""","""",RC[-2]-2*RC[-100]*R7C[-3]/(COS(ASIN(RC[-1]))*COS(ASIN(RC[-1]))))"
    Range("DG11").FormulaR1C1 = "=IF(RC[-6]="""","""",R3C3/16*COS(ASIN(RC[-2]))*COS(ASIN(RC[-2]))*(4+R3C3*(4-3*COS(ASIN(RC[-2]))*COS(ASIN(RC[-2])))))"
    Range("DH11").FormulaR1C1 = "=IF(RC[-7]="""","""",IF(AND(RC[-110]=R4C[-7],RC[-108]=R4C[-4]),""samepoint"",""N.A.""))"
    Range("DI11").FormulaR1C1 = "=IF(RC[-8]="""","""",IF(AND(RC[-109]=R4C[-5],R4C[-8]>RC[-111]),""northsouth"",""N.A.""))"
    Range("DJ11").FormulaR1C1 = "=IF(RC[-9]="""","""",IF(AND(RC[-110]=R4C[-6],RC[-112]>R4C[-9]),""southnorth"",""N.A.""))"
    Range("DK11").FormulaR1C1 = "=IF(RC[-10]="""","""",ATAN2((RC[-104]*R7C[-8]-RC[-105]*R8C[-8]*COS(RC[-9])),R8C[-8]*SIN(RC[-9])))"
    Range("DL11").FormulaR1C1 = "=IF(RC[-11]="""","""",IF(RC[-4]=""samepoint"",0,IF(RC[-3]=""northsouth"",0,IF(RC[-2]=""southnorth"",180,IF(RC[-1]<0,RC[-1]*180/PI()+360,RC[-1]*180/PI())))))"
    Range("DM11").FormulaR1C1 = "=IF(OR(RC[-1]="""",R[1]C[-1]="""",RC[1]<>"""",R[1]C[1]<>""""),"""",IF(AND(ABS(R[1]C[-1]-RC[-1])>90,ABS(R[1]C[-1]-RC[-1])<180),RC[-116]+(RC20/(RC20+R[1]C20))*(R[1]C1-RC1),""""))"
    ''Range("DM11").FormulaR1C1 = "=IF(OR(RC[-1]="""",R[1]C[-1]="""",RC[1]<>"""",R[1]C[1]<>""""),"""",IF(AND(ABS(R[1]C[-1]-RC[-1])>90,OR(AND(R4C113>=180,R4C113<=270,OR(AND(R[1]C3>=R4C116,R[1]C5<=R4C117),AND(RC3>=R4C116,RC5<=R4C117))),AND(R4C113>=270,R4C113<=360,OR(AND(R[1]C3<=R4C116,R[1]C5<=R4C117),AND(RC3<=R4C116,RC5<=R4C117))),AND(R4C113>=0,R4C113<=90,OR(AND(R[1]C3<=R4C116,R[1]C5>=R4C117),AND(RC3<=R4C116,RC5>=R4C117))),AND(R4C113>=90,R4C113<=180,OR(AND(R[1]C3>=R4C116,R[1]C5>=R4C117),AND(RC3>=R4C116,RC5>=R4C117))))),RC[-116]+(RC20/(RC20+R[1]C20))*(R[1]C1-RC1),""""))"
    Range("DM9").FormulaR1C1 = "=MAX(R[2]C:R[10001]C)"
    Range("DN11").FormulaR1C1 = "=IF(OR(R4C[-2]=0,R4C[-1]=0,RC[-117]>R5C3,RC[-2]=""""),"""",IF(OR(AND(RC[-116]=R4C[-13],RC[-114]=R4C[-10]),AND(R5C[-3]>R6C[-3],RC[-2]>=R6C[-3],RC[-2]<=R5C[-3]),AND(R5C[-3]<R6C[-3],OR(RC[-2]>=R6C[-3],RC[-2]<=R5C[-3]))),RC[-117],""""))"
    'Range("DN11").FormulaR1C1 = "=IF(OR(R4C[-2]=0,R4C[-1]=0,RC[-117]<R10C[-95],RC[-117]>R4C[-95],RC[-117]>R5C3,RC[-2]=""""),"""",IF(OR(AND(RC[-116]=R4C[-13],RC[-114]=R4C[-10]),AND(R5C[-3]>R6C[-3],RC[-2]>=R6C[-3],RC[-2]<=R5C[-3]),AND(R5C[-3]<R6C[-3],OR(RC[-2]>=R6C[-3],RC[-2]<=R5C[-3]))),RC[-117],""""))"
    Range("DN4").FormulaR1C1 = "=IF(MIN(R[7]C:R[10006]C)=0,R[5]C[-1],MIN(R[7]C:R[10006]C))"
    Range("DN6").FormulaR1C1 = "=IF(AND(R[3]C[-1]<>0,OR(R[-2]C=R[3]C[-1],AND(MAX(R6C148:R6C149)>0,MAX(R[5]C:R[10004]C)>MAX(R6C148:R6C149)))),R[3]C[-1],MAX(R[5]C:R[10004]C))"
    Range("DO11").FormulaR1C1 = "=IF(OR(R2C=0,RC[-99]=""""),"""",IF(AND(R2C=1,RC[-99]<=0.5),RC[-118],IF(AND(RC[-99]<=0.5,R2C=2,RC[-118]<R5C3,RC[-118]>MIN(R10C[-104],R10C[-100],R10C[-96])),RC[-118],IF(AND(R2C=3,RC[-99]<=0.5,RC[-118]>=MAX(R10C[-96],R10C[-100],R10C[-104])),RC[-118],""""))))"
    Range("DO6").FormulaR1C1 = "=MIN(R[5]C:R[10004]C)"
    Range("DO7").FormulaR1C1 = "=MAX(R[4]C:R[10003]C)"
    Range("DO3").FormulaR1C1 = _
        "=IF(OR(R[-1]C=0,AND(R[1]C[-1]=0,R[3]C[-1]=0,R[3]C=0)),"""",IF(R[1]C[-1]>0,""SECTOR"",""CYLINDER""))"
LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("DA11:DO11").AutoFill Destination:=.Range("DA11:DO" & LastRow), Type:=xlFillDefault
End If

If Range("A4") > 0 Or Range("A5") > 0 Or Range("A6") > 0 Then
    Range("DR3").FormulaR1C1 = "=IF(MAX(R5C29:R7C29)=0,0,YDWK2!R660C5)"
    ''Range("DR3").FormulaR1C1 = "=IF(MAX(R5C29:R7C29)=0,0,IF(MAX(R[-2]C[-35],R[-2]C[-19],R[-2]C[-3])=R[-2]C[-35],R[1]C[-38],IF(MAX(R[-2]C[-119],R[-2]C[-19],R[-2]C[-3])=R[-2]C[-19],R[1]C[-22],IF(MAX(R[-2]C[-119],R[-2]C[-19],R[-2]C[-3])=R[-2]C[-3],R[1]C[-6]))))"
    Range("DS3").FormulaR1C1 = "=IF(MAX(R5C29:R7C29)=0,0,YDWK2!R661C5)"
    ''Range("DS3").FormulaR1C1 = "=IF(MAX(R5C29:R7C29)=0,0,IF(MAX(R[-2]C[-36],R[-2]C[-20],R[-2]C[-4])=R[-2]C[-36],R[1]C[-38],IF(MAX(R[-2]C[-120],R[-2]C[-20],R[-2]C[-4])=R[-2]C[-20],R[1]C[-22],IF(MAX(R[-2]C[-120],R[-2]C[-20],R[-2]C[-4])=R[-2]C[-4],R[1]C[-6]))))"
    Range("DT3").FormulaR1C1 = "=IF(AND(R3C122=R4C84,R4C85=R3C123,R4C86=MAX(R4C118,R4C102,R4C86)),MAX(R4C86,R7C87),IF(AND(R3C122=R4C84,R3C123=R4C85),MAX(R4C86,R6C86),IF(AND(R3C122=R4C100,R4C101=R3C123,R4C102=MAX(R4C118,R4C102,R4C86)),MAX(R4C102,R7C103),IF(AND(R3C122=R4C100,R4C101=R3C123),MAX(R4C102,R6C102),IF(AND(R3C122=R4C116,R4C117=R3C123,R4C118=MAX(R4C118,R4C102,R4C86)),MAX(R4C118,R7C119),IF(AND(R3C122=R4C116,R4C117=R3C123),MAX(R4C118,R6C118),0))))))"
    Range("DQ11").FormulaR1C1 = "=IF(R3C[3]=0,"""",IF(AND(RC[-120]>R3C[3],RC[-120]<=SUMMARY!R14C[-112]),6371*ACOS(SIN(R3C[1])*SIN(RC[-118])+COS(R3C[1])*COS(RC[-118])*COS(R3C[2]-RC[-116])),""""))"
    Range("DQ8").FormulaR1C1 = "=MAX(R[3]C:R[10002]C)"
    Range("DR11").FormulaR1C1 = "=IF(RC[-1]=R8C[-1],RC[-119],"""")"
    'Range("DR11").FormulaR1C1 = "=IF(RC[-1]=R8C[-1],RC[-120],"""")"
    Range("DR8").FormulaR1C1 = "=MAX(R[3]C:R[10002]C)"
    Range("DS11").FormulaR1C1 = "=IF(RC[-1]="""","""",RC[-118])"
    'Range("DS11").FormulaR1C1 = "=IF(RC[-1]="""","""",RC[-119])"
    Range("DS8").FormulaR1C1 = "=MAX(R[3]C:R[10002]C)"
    Range("DT11").FormulaR1C1 = "=IF(RC[-1]="""","""",RC[-123])"
    Range("DT8").FormulaR1C1 = "=MAX(R[3]C:R[10002]C)"
    Range("DU11").FormulaR1C1 = "=IF(RC[-1]="""","""",RC[-119])"
    Range("DU8").FormulaR1C1 = "=MAX(R[3]C:R[10002]C)"
    Range("DV11").FormulaR1C1 = "=IF(RC[-1]="""","""",RC[-119])"
    Range("DV8").FormulaR1C1 = "=MAX(R[3]C:R[10002]C)"
    'Range("DW11").FormulaR1C1 = "=IF(RC[-1]=R8C[-1],RC[-121],"""")"
    'Range("DW8").FormulaR1C1 = "=MIN(R[3]C:R[10002]C)"
    'Range("DX11").FormulaR1C1 = "=IF(RC[-7]=R8C[-2],RC[-126],"""")"
    'Range("DX8").FormulaR1C1 = "=MAX(R[3]C:R[10002]C)"
    'Range("DY11").FormulaR1C1 = "=IF(RC[-1]="""","""",RC[-125])"
    'Range("DY8").FormulaR1C1 = "=MAX(R[3]C:R[10002]C)"
    'Range("DZ11").FormulaR1C1 = "=IF(RC[-1]="""","""",RC[-129])"
    'Range("DZ8").FormulaR1C1 = "=MAX(R[3]C:R[10002]C)"
    'Range("DV4").FormulaR1C1 = "=IF(R[4]C[4]=0,R[4]C[-2],R[4]C[4])"
LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("DQ11:DV11").AutoFill Destination:=.Range("DQ11:DV" & LastRow), Type:=xlFillDefault
 End If
    
 If Range("A7") > 0 Then
    Range("EB2").FormulaR1C1 = "=B!R28C1+B!R28C2/60"
    Range("EB4").FormulaR1C1 = "=B!R28C4+B!R28C5/60"
    Range("EB6").FormulaR1C1 = "=MIN(R[5]C:R[10004]C)"
    Range("EO4").FormulaR1C1 = "=RADIANS(R2C132)"
    Range("EQ4").FormulaR1C1 = "=RADIANS(R4C132)"
    Range("EF5").FormulaR1C1 = "=(1-R3C3)*TAN(R4C145)"
    Range("EF6").FormulaR1C1 = "=ATAN(R[-1]C)"
    Range("EF7").FormulaR1C1 = "=SIN(R[-1]C)"
    Range("EF8").FormulaR1C1 = "=COS(R[-2]C)"
    Range("EM4").FormulaR1C1 = "=R[33]C[-117]"
    Range("EM5").FormulaR1C1 = "=IF(R[-1]C="""","""",R[-1]C+45)"
    Range("EO5").FormulaR1C1 = "=IF(RC[-2]="""","""",IF(R[-1]C[-2]+45>360,R[-1]C[-2]+45-360,RC[-2]))"
    Range("EM6").FormulaR1C1 = "=IF(R[-2]C="""","""",R[-2]C-45)"
    Range("EO6").FormulaR1C1 = "=IF(R[-1]C="""","""",IF(R[-2]C[-2]-45<0,R[-2]C[-2]-45+360,RC[-2]))"
    Range("EB11").FormulaR1C1 = _
        "=IF(RC[-131]="""","""",IF(OR(AND(R4C[-43]=R2C,R4C=R4C[-40],R2C119=0,RC[-131]<R10C[-113]),AND(R4C[-43]=R2C,R4C=R4C[-40],RC[-131]>=MAX(R10C[-109],R10C[-113],R4C[-117],R10C[-101])),RC[-131]>MAX(R10C15,R10C19,R10C23)),6371*ACOS((SIN(RC[-129])*SIN(RADIANS(R2C132))+COS(RC[-129])*COS(RADIANS(R2C132))*COS(RC[-127]-RADIANS(R4C132)))),IF(AND(RC[-131]>=MAX(R4C[-109],R4C[-113],R4C[-117]),R2C=0=FALSE),6371*ACOS((SIN(RC[-129])*SIN(RADIANS(R2C132))+COS(RC[-129])*COS(RADIANS(R2C132))*COS(RC[-127]-RADIANS(R4C132)))),"""")))"
    Range("EC11").FormulaR1C1 = "=IF(OR(RC[-132]="""",RC[-1]>20,RC[-132]<MAX(R10C[-110],R10C[-114],R10C[-118],R10C[-102])),"""",RC[-132])"
    Range("EC10").FormulaR1C1 = "=IF(MIN(R[1]C:R[10000]C)=0,"""",MIN(R[1]C:R[10000]C))"
    Range("ED11").FormulaR1C1 = "=IF(RC[-1]="""","""",RADIANS(R4C132)-RC[-129])"
    Range("EF11").FormulaR1C1 = "=IF(RC[-3]="""","""",(R8C*SIN(RC[-2])*R8C*SIN(RC[-2]))+(RC[-125]*R7C-RC[-126]*R8C*COS(RC[-2]))*(RC[-125]*R7C-RC[-126]*R8C*COS(RC[-2])))"
    Range("EG11").FormulaR1C1 = "=IF(RC[-4]="""","""",(RC[-127]*R7C[-1])+(RC[-126]*R8C[-1]*COS(RC[-3])))"
    Range("EH11").FormulaR1C1 = "=IF(RC[-5]="""","""",IF(RC[-2]=0,0,RC[-127]*R8C[-2]*SIN(RC[-4])/SQRT(RC[-2])))"
    Range("EI11").FormulaR1C1 = "=IF(RC[-6]="""","""",RC[-2]-2*RC[-129]*R7C[-3]/(COS(ASIN(RC[-1]))*COS(ASIN(RC[-1]))))"
    Range("EJ11").FormulaR1C1 = "=IF(RC[-7]="""","""",R3C3/16*COS(ASIN(RC[-2]))*COS(ASIN(RC[-2]))*(4+R3C3*(4-3*COS(ASIN(RC[-2]))*COS(ASIN(RC[-2])))))"
    Range("EE11").FormulaR1C1 = "=IF(RC[-2]="""","""",RC[-1]+(1-RC[5])*R3C3*RC[3]*(ACOS(RC[2])+RC[5]*SIN(ACOS(RC[2]))*(RC[4]+RC[5]*RC[2]*(-1+2*RC[4]*RC[4]))))"
    Range("EK11").FormulaR1C1 = "=IF(RC[-8]="""","""",IF(AND(RC[-139]=R2C[-9],RC[-137]=R4C[-9]),""samepoint"",""N.A.""))"
    Range("EL11").FormulaR1C1 = "=IF(RC[-9]="""","""",IF(AND(RC[-138]=R4C[-8],R4C[-10]>RC[-140]),""northsouth"",""N.A.""))"
    Range("EM11").FormulaR1C1 = "=IF(RC[-10]="""","""",IF(AND(RC[-139]=R4C[-9],RC[-141]>R4C[-11]),""southnorth"",""N.A.""))"
    Range("EN11").FormulaR1C1 = "=IF(RC[-11]="""","""",ATAN2((RC[-133]*R7C[-8]-RC[-134]*R8C[-8]*COS(RC[-9])),R8C[-8]*SIN(RC[-9])))"
    Range("EO11").FormulaR1C1 = "=IF(RC[-12]="""","""",IF(RC[-4]=""samepoint"",0,IF(RC[-3]=""northsouth"",0,IF(RC[-2]=""southnorth"",180,IF(RC[-1]<0,RC[-1]*180/PI()+360,RC[-1]*180/PI())))))"
    Range("EP11").FormulaR1C1 = "=IF(OR(RC[-1]="""",R[1]C[-1]="""",RC[1]<>"""",R[1]C[1]<>""""),"""",IF(AND(RC[-14]<1,ABS(R[1]C[-1]-RC[-1])>90,ABS(R[1]C[-1]-RC[-1])<180),RC[-145]+(RC[-14]/(RC[-14]+R[1]C[-14]))*(R[1]C[-145]-RC[-145]),""""))"
    'Range("EP11").FormulaR1C1 = _
        "=IF(OR(RC[-1]="""",R[1]C[-1]="""",RC[2]<>"""",R[1]C[2]<>""""),"""",IF(AND(RC[-14]<1,ABS(R[1]C[-1]-RC[-1])>90,OR(AND(R4C143>=180,R4C143<=270,OR(AND(R[1]C3>=R4C145,R[1]C5<=R4C147),AND(RC3>=R4C145,RC5<=R4C147))),AND(R4C143>=270,R4C143<=360,OR(AND(R[1]C3<=R4C145,R[1]C5<=R4C147),AND(RC3<=R4C145,RC5<=R4C147))),AND(R4C143>=0,R4C143<=90,OR(AND(R[1]C3<=R4C145,R[1]C5>=R4C147),AND(RC3<=R4C145,RC5>=R4C14))),AND(R4C143>=90,R4C143<=180,OR(AND(R[1]C3>=R4C145,R[1]C5>=R4C147),AND(RC3>=R4C145,RC5>=R4C147))))),RC[-145]+(RC[-14]/(RC[-14]+R[1]C[-14]))*(R[1]C[-145]-RC[-145]),""""))"
    Range("EP9").FormulaR1C1 = "=IF(MIN(R[2]C:R[10001]C)=0,"""",MIN(R[2]C:R[10001]C))"
    Range("EQ11").FormulaR1C1 = "=IF(R[1]C[-146]="""","""",IF(OR(RC[-146]>R5C3,AND(RC[-146]<MAX(R10C[-132],R10C[-128],R10C[-124]),OR(R10C[-132]>0,R10C[-128]>0,R10C[-124]>0))),"""",IF(R3C[-25]="""",6371*ACOS(SIN(R[1]C[-144])*SIN(R4C[-103])+COS(R[1]C[-144])*COS(R4C[-103])*COS(R4C[-101]-R[1]C[-142])),6371*ACOS(SIN(R[1]C[-144])*SIN(R3C[-25])+COS(R[1]C[-144])*COS(R3C[-25])*COS(R3C[-24]-R[1]C[-142])))))"
    'Range("EQ11").FormulaR1C1 = "=IF(RC[-146]="""","""",IF(OR(RC[-146]>R5C3,AND(RC[-146]<MAX(R10C[-132],R10C[-128],R10C[-124]),OR(R10C[-132]>0,R10C[-128]>0,R10C[-124]>0))),"""",RC[-146]))"
    Range("ER11").FormulaR1C1 = _
        "=IF(OR(RC[-1]="""",RC[-16]="""",RC[-3]="""",AND(SUMMARY!R31C[-140]>0,RC[-147]<MAX(R10C[-133],R10C[-129],R10C[-125]))),"""",IF(OR(AND(RC[-146]=R2C[-16],RC[-144]=R4C[-16]),AND(R5C[-3]>R6C[-3],RC[-3]>=R6C[-3],RC[-3]<=R5C[-3]),AND(R6C[-3]>R5C[-3],OR(RC[-3]>=R6C[-3],RC[-3]<=R5C[-3]))),RC[-147],""""))"
    Range("ER6").FormulaR1C1 = "=IF(R[1]C>0,R[1]C,IF(AND(R[1]C=0,ABS(R4C132)>0),0,0))"
    Range("ER7").FormulaR1C1 = _
        "=IF(AND(SUMMARY!R32C8=1,R4C143=0,R4C145=YDWK2!R660C5,YDWK2!R661C5=R4C147),0,MIN(R11C148:R10010C148,R[2]C[-2]))"
    ''Range("ES11").FormulaR1C1 = "=IF(RC[-3]<>"""",RC[-143]+(RC[-17]/(RC[-17]+R[1]C[-17]))*(R[1]C[-143]-RC[-143]),IF(OR(RC[-148]="""",RC[-1]="""",RC[-2]="""",RC[-148]<MAX(SUMMARY!R27C[-140]:R29C[-140])),"""",RC[-143]))"
    'Range("ES11").FormulaR1C1 = "=IF(RC[-3]<>"""",RC[-143]+(RC[-17]/(RC[-17]+R[1]C[-17]))*(R[1]C[-143]-RC[-143]),IF(OR(RC[-148]="""",RC[-1]="""",RC[-2]="""",RC[-148]<MAX(SUMMARY!R27C[-140]:R29C[-140]),AND(OR(SUMMARY!R10C2=3,SUMMARY!R10C2=5),SUMMARY!R20C[-141]>0,RC[-148]>SUMMARY!R20C[-141])),"""",RC[-143]))"
    Range("ES11").FormulaR1C1 = _
        "=IF(AND(RC[-3]<>"""",R5C45<>3),RC[-143]+(RC[-17]/(RC[-17]+R[1]C[-17]))*(R[1]C[-143]-RC[-143]),IF(AND(RC[-3]<>"""",R5C45=3),RC[-142]+(RC[-17]/(RC[-17]+R[1]C[-17]))*(R[1]C[-142]-RC[-142]),IF(OR(RC[-148]="""",RC[-1]="""",RC[-2]="""",RC[-148]<MAX(SUMMARY!R27C[-140]:R29C[-140]),AND(OR(SUMMARY!R10C2=3,SUMMARY!R10C2=5),SUMMARY!R20C[-141]>0,RC[-148]>SUMMARY!R20C[-141])),"""",IF(R5C45<>3,RC[-143],RC[-142]))))"
    Range("ES6").FormulaR1C1 = "=IF(R6C148=0,0,MAX(R[5]C[-1]:R[10004]C[-1],R[3]C[-3]))"
    Range("ES9").FormulaR1C1 = "=IF(R6C148=0,0,MAX(R[2]C:R[10001]C))"
    Range("ET11").FormulaR1C1 = "=IF(RC[-1]="""","""",IF(AND(RC[-4]<>"""",RC[-1]=R9C[-1]),RC[-4],IF(RC[-1]=R9C[-1],RC[-149],"""")))"
    Range("ET9").FormulaR1C1 = "=MAX(R[2]C:R[10001]C)"
    Range("EU11").FormulaR1C1 = "=IF(RC[-2]="""","""",RC[-19])"
    Range("EU9").FormulaR1C1 = "=MIN(R[2]C:R[10001]C)"
    Range("EV11").FormulaR1C1 = "=IF(RC[-6]<>"""",RC[-6],IF(OR(RC[-7]="""",RC[-20]>1,RC[-4]=""""),"""",RC[-151]))"
    Range("EV9").FormulaR1C1 = "=MIN(R[2]C:R[10001]C)"
    Range("EW11").FormulaR1C1 = "=IF(AND(RC[-7]<>"""",R5C45<>3),RC[-147]+(RC[-21]/(RC[-21]+R[1]C[-21]))*(R[1]C[-147]-RC[-147]),IF(AND(RC[-7]<>"""",R5C45=3),RC[-146]+(RC[-21]/(RC[-21]+R[1]C[-21]))*(R[1]C[-146]-RC[-146]),IF(RC[-1]="""","""",IF(R5C45<>3,RC[-147],RC[-146]))))"
    'Range("EW11").FormulaR1C1 = "=IF(RC[-7]<>"""",RC[-147]+(RC[-21]/(RC[-21]+R[1]C[-21]))*(R[1]C[-147]-RC[-147]),IF(RC[-1]="""","""",RC[-147]))"
    'Range("EW4").FormulaR1C1 = "=MAX(SUMMARY!R27C9:R29C9)"
    Range("EW4").FormulaR1C1 = "=IF(B!R[9]C[-150]<>"""",R[2]C[-5],MAX(SUMMARY!R27C9:R29C9))"
    Range("EW9").FormulaR1C1 = "=MAX(R[2]C:R[10001]C)"
    Range("EX11").FormulaR1C1 = _
        "=IF(OR(R9C152=0,RC[-22]=""""),"""",IF(AND(RC[-8]<>"""",OR(RC[-148]+(RC[-22]/(RC[-22]+R[1]C[-22]))*(R[1]C[-148]-RC[-148])=R9C[-1],RC[-147]+(RC[-22]/(RC[-22]+R[1]C[-22]))*(R[1]C[-147]-R[1]C[-148])=R9C[-1])),RC[-8],IF(RC[-1]=R9C[-1],RC[-153],"""")))"
    'Range("EX11").FormulaR1C1 = "=IF(OR(R9C152=0,RC[-22]=""""),"""",IF(AND(RC[-8]<>"""",RC[-148]+(RC[-22]/(RC[-22]+R[1]C[-22]))*(R[1]C[-148]-RC[-148])=R9C[-1]),RC[-8],IF(RC[-1]=R9C[-1],RC[-153],"""")))"
    Range("EX9").FormulaR1C1 = "=MIN(R[2]C:R[10001]C)"
    ''NEW EY11,EY4,EY9 For FinLine
    ''Range("EY11").FormulaR1C1 = "=IF(OR(RC[-23]>0.66,RC[-8]>R[1]C[-8],RC[-154]<R4C[-2],AND(OR(R4C[-140]>0,R10C[-136]>0,R10C[-132]>0),RC[-22]="""")),"""",IF(RC[-8]>=SQRT(R4C^2+RC[-23]^2),RC[-154],""""))"
    '''Range("EY11").FormulaR1C1 = "=IF(OR(RC[-23]>0.66,RC[-154]<R4C[-2],AND(OR(R4C[-140]>0,R10C[-136]>0,R10C[-132]>0),RC[-22]="""")),"""",IF(AND(R[1]C[-8]>=SQRT(R4C^2+R[1]C[-23]^2),OR(RC[-10]=R19C176,RC[-10]=R19C178,RC[-10]=R19C179,AND(OR(AND(R19C176>=0,R19C176<=90),AND(R19C176>270,R19C176<=360)),RC[-10]>R19C178,RC[-10]<R19C179),AND(RC[-10]<R19C179,RC[-10]>R19C178),AND(R19C176>90,R19C176<180,OR(RC[-10]<=R19C179,RC[-10]>=R19C178)),AND(R19C176>=180,R19C176<=270,OR(RC[-10]<R19C179,RC[-10]>R19C178)))),RC[-154],""""))"
    Range("EY11").FormulaR1C1 = "=IF(OR(RC[-23]>0.66,RC[-154]<R4C[-2],R[1]C[-8]<RC[-8],AND(OR(R4C[-140]>0,R10C[-136]>0,R10C[-132]>0),RC[-22]="""")),"""",IF(AND(R[1]C[-8]>=SQRT(R4C^2+R[1]C[-23]^2),OR(RC[-10]=R19C176,RC[-10]=R19C178,RC[-10]=R19C179,AND(OR(AND(R19C176>=0,R19C176<=90),AND(R19C176>270,R19C176<=360)),RC[-10]>R19C178,RC[-10]<R19C179),AND(RC[-10]<R19C179,RC[-10]>R19C178),AND(R19C176>90,R19C176<180,OR(RC[-10]<=R19C179,RC[-10]>=R19C178)),AND(R19C176>=180,R19C176<=270,OR(RC[-10]<R19C179,RC[-10]>R19C178)))),RC[-154],""""))"
    'Range("EY11").FormulaR1C1 = "=IF(OR(RC[-23]>0.66,RC[-154]<R4C[-2],AND(OR(R4C[-140]>0,R10C[-136]>0,R10C[-132]>0),RC[-22]="""")),"""",IF(OR(RC[-10]=R19C176,RC[-10]=R19C178,RC[-10]=R19C179,AND(OR(AND(R19C176>=0,R19C176<=90),AND(R19C176>270,R19C176<=360)),RC[-10]>R19C178,RC[-10]<R19C179),AND(RC[-10]<R19C179,RC[-10]>R19C178),AND(R19C176>90,R19C176<180,OR(RC[-10]<=R19C179,RC[-10]>=R19C178)),AND(R19C176>=180,R19C176<=270,OR(RC[-10]<R19C179,RC[-10]>R19C178))),RC[-154],""""))"
    Range("EY4").FormulaR1C1 = "=IF(MAX(R[1]C[-126]:R[3]C[-126])=0,6371*ACOS(SIN(R4C[-111])*SIN(R4C[-10])+COS(R4C[-111])*COS(R4C[-10])*COS(R4C[-8]-R4C[-109])),6371*ACOS(SIN(R4C[-10])*SIN(R3C[-33])+COS(R4C[-10])*COS(R3C[-33])*COS(R3C[-32]-R4C[-8])))"
    ''Range("EY9").FormulaR1C1 = "=IF(MIN(R[2]C:R[10001]C)=MAX(R[2]C:R[10001]C),R[1]C,SMALL((R[2]C:R[10001]C),2))"
    Range("EY9").FormulaR1C1 = "=MIN(R[2]C:R[10001]C)"
    Range("EY10").FormulaR1C1 = "=MAX(R[1]C:R[10000]C)"
    Range("EZ11").FormulaR1C1 = "=IF(RC[-1]=R9C[-1],RC[-154],"""")"
    Range("EZ9").FormulaR1C1 = "=MIN(R[2]C:R[10001]C)"
    Range("FA11").FormulaR1C1 = "=IF(AND(R9C[-5]>0,RC[-156]=R9C[-5]),RC[-155],"""")"
    Range("FA9").FormulaR1C1 = "=MIN(R[2]C:R[10001]C)"
    Range("FB11").FormulaR1C1 = "=IF(RC[-2]="""","""",RC[-154])"
    Range("FB9").FormulaR1C1 = "=MIN(R[2]C:R[10001]C)"
    Range("FC11").FormulaR1C1 = "=IF(RC[-2]="""","""",RC[-155])"
    Range("FC9").FormulaR1C1 = "=MIN(R[2]C:R[10001]C)"
    Range("FD11").FormulaR1C1 = "=IF(RC[-2]="""","""",IF(R5C45<>3,RC[-154],RC[-153]))"
    'Range("FD11").FormulaR1C1 = "=IF(RC[-2]="""","""",RC[-154])"
    Range("FD9").FormulaR1C1 = "=MIN(R[2]C:R[10001]C)"
    Range("FE11").FormulaR1C1 = "=IF(RC[-2]="""","""",IF(R5C45<>3,RC[-155],RC[-154]))"
    'Range("FE11").FormulaR1C1 = "=IF(RC[-2]="""","""",RC[-155])"
    Range("FE9").FormulaR1C1 = "=MIN(R[2]C:R[10001]C)"
    Range("FF11").FormulaR1C1 = "=IF(RC[-2]="""","""",RC[-17])"
    Range("FF9").FormulaR1C1 = "=360-MIN(R[2]C:R[10001]C)"
    Range("FG11").FormulaR1C1 = "=IF(RC[-2]="""","""",RC[-18])"
    Range("FG9").FormulaR1C1 = "=MIN(R[2]C:R[10001]C)"
    Range("FH11").FormulaR1C1 = "=IF(RC[-9]=R9C[-9],R[-1]C[-163],"""")"
    Range("FH9").FormulaR1C1 = "=MIN(R[2]C:R[9991]C)"
    Range("FI11").FormulaR1C1 = "=IF(AND(R9C[-13]>0,RC[-164]=R9C[-13]),R[-1]C[-164],"""")"
    Range("FI9").FormulaR1C1 = "=MIN(R[2]C:R[10001]C)"
    Range("FJ11").FormulaR1C1 = "=IF(RC[-2]="""","""",R[-1]C[-164])"
    Range("FJ9").FormulaR1C1 = "=MIN(R[2]C:R[10001]C)"
    Range("FK11").FormulaR1C1 = "=IF(RC[-2]="""","""",R[-1]C[-165])"
    Range("FK9").FormulaR1C1 = "=MIN(R[2]C:R[10001]C)"
    Range("FL11").FormulaR1C1 = "=IF(RC[-2]="""","""",R[-1]C[-164])"
    Range("FL9").FormulaR1C1 = "=MIN(R[2]C:R[10001]C)"
    Range("FM11").FormulaR1C1 = "=IF(RC[-2]="""","""",R[-1]C[-165])"
LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("EB11:FM11").AutoFill Destination:=.Range("EB11:FM" & LastRow), Type:=xlFillDefault

    Range("FM9").FormulaR1C1 = "=MIN(R[2]C:R[10001]C)"
    Range("FN12").FormulaR1C1 = "=IF(RC[-2]="""","""",IF(R5C45<>3,R[-1]C[-164],R[-1]C[-163]))"
    'Range("FN12").FormulaR1C1 = "=IF(RC[-2]="""","""",R[-1]C[-164])"
    Range("FN9").FormulaR1C1 = "=MIN(R[2]C:R[10001]C)"
    Range("FO12").FormulaR1C1 = "=IF(RC[-2]="""","""",IF(R5C45<>3,R[-1]C[-165],R[-1]C[-164]))"
    'Range("FO12").FormulaR1C1 = "=IF(RC[-2]="""","""",R[-1]C[-165])"
    Range("FO9").FormulaR1C1 = "=MIN(R[2]C:R[10001]C)"
    Range("FP12").FormulaR1C1 = "=IF(RC[-2]="""","""",R[-1]C[-27])"
    Range("FP9").FormulaR1C1 = "=360-MIN(R[2]C:R[10001]C)"
    Range("FQ12").FormulaR1C1 = "=IF(RC[-2]="""","""",R[-1]C[-28])"
    Range("FQ9").FormulaR1C1 = "=MIN(R[2]C:R[10001]C)"
    Range("FR12").FormulaR1C1 = _
        "=IF(AND(R5C45<>3,RC[-173]=R9C152),RC[-168],IF(AND(R5C45=3,RC[-173]=R9C152),RC[-167],IF(AND(R5C45<>3,RC[-28]=R9C152),RC[-168]+(RC[-42]/(RC[-42]+R[1]C[-42]))*(R[1]C[-168]-RC[-168]),IF(AND(R5C45=3,RC[-28]=R9C152),RC[-167]+(RC[-42]/(RC[-42]+R[1]C[-42]))*(R[1]C[-167]-RC[-167]),""""))))"
    'Range("FR12").FormulaR1C1 = "=IF(RC[-173]=R9C152,RC[-168],IF(RC[-28]=R9C152,RC[-168]+(RC[-42]/(RC[-42]+R[1]C[-42]))*(R[1]C[-168]-RC[-168]),""""))"
    Range("FR9").FormulaR1C1 = "=MAX(R[2]C:R[10001]C)"
    ''Range("EZ5").FormulaR1C1 = "=IF(OR(YDWK2!R877C6>0,AND(R[-3]C[-24]=R[4]C[1],R[-1]C[-24]=R[4]C[3])),YDWK2!R877C6,""N/A"")"
    Range("EZ5").FormulaR1C1 = "=YDWK2!R877C6"
    Range("EZ6").FormulaR1C1 = "=YDWK2!R1314C6"
    Range("FD4").FormulaR1C1 = "=IF(YDWK2!R877C26=0,""N/A"",YDWK2!R877C26)"
    Range("FD5").FormulaR1C1 = "=YDWK2!R1097C6"
    Range("FJ5").FormulaR1C1 = "=IF(YDWK2!R877C16=0,""N/A"",YDWK2!R877C16)"
    Range("FJ6").FormulaR1C1 = "=YDWK2!R1387C6"
LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("FN12:FR12").AutoFill Destination:=.Range("FN12:FR" & LastRow), Type:=xlFillDefault
End If
 
 If Range("A3") > 0 Or Range("A7") > 0 Then
    Range("FT15").FormulaR1C1 = "=RC[-150]"
    Range("FV15").FormulaR1C1 = "=IF(RC[-2]+90>360,RC[-2]+90-360,RC[-2]+90)"
    Range("FW15").FormulaR1C1 = "=IF(RC[-3]-90<0,360+RC[-3]-90,RC[-3]-90)"
    Range("FT19").FormulaR1C1 = "=IF(SUMMARY!R[12]C[-168]=0,YDWK2!R[205]C[-173],YDWK2!R[637]C[-173])"
    Range("FV19").FormulaR1C1 = "=IF(RC[-2]+90>360,RC[-2]+90-360,RC[-2]+90)"
    Range("FW19").FormulaR1C1 = "=IF(RC[-3]-90<0,RC[-3]-90+360,RC[-3]-90)"
    Range("FS22").FormulaR1C1 = "=R6C52"
    Range("FT22").FormulaR1C1 = "=R6C62"
    Range("FU22").FormulaR1C1 = "=R5C56"
    Range("FW22").FormulaR1C1 = "=IF(MAX(R5C29:R7C29)=0,YDWK2!R224C3,YDWK2!R368C3)"
    Range("FT25").FormulaR1C1 = _
        "=IF(R[-3]C[-1]=""N/A"",""N/A"",R[-3]C[-1]*(COS(RADIANS(R[-3]C[3]))*SIN(RADIANS(R[8]C[2]))-SIN(RADIANS(R[-3]C[3]))*COS(RADIANS(R[8]C[2]))))"
    Range("FU25").FormulaR1C1 = _
        "=IF(R[-3]C[-2]=""N/A"",""N/A"",R[-3]C[-2]*(SIN(RADIANS(R[-3]C[2]))*SIN(RADIANS(R[8]C[1]))+COS(RADIANS(R[-3]C[2]))*COS(RADIANS(R[8]C[1]))))"
    Range("FV25").FormulaR1C1 = _
        "=IF(OR(R[-3]C[-3]=0,R[-3]C[-2]=0,AND(RC[-1]<0,R[1]C[-1]>0)),""Good Start"",""NO START"")"
    Range("FT26").FormulaR1C1 = _
        "=IF(R[-4]C=""N/A"",""N/A"",R[-4]C*(COS(RADIANS(R[-4]C[3]))*SIN(RADIANS(R[9]C[2]))-SIN(RADIANS(R[-4]C[3]))*COS(RADIANS(R[9]C[2]))))"
    Range("FU26").FormulaR1C1 = _
        "=IF(R[-4]C[-1]=""N/A"",""N/A"",R[-4]C[-1]*(SIN(RADIANS(R[-4]C[2]))*SIN(RADIANS(R[9]C[1]))+COS(RADIANS(R[-4]C[2]))*COS(RADIANS(R[9]C[1]))))"
    Range("FV26").FormulaR1C1 = _
        "=IF(OR(RC[-1]=""N/A"",R[-1]C[-1]=""N/A""),""NO START"",IF(ABS(R[4]C[-2])<0.5,""Good Start"",""NO START""))"
    Range("FT27").FormulaR1C1 = _
        "=IF(OR(R[-2]C=""N/A"",R[-1]C=""N/A""),"""",R[-1]C-R[-2]C)"
    Range("FU27").FormulaR1C1 = "=IF(OR(R[-5]C[-2]=""N/A"",R[-5]C[-1]=""N/A""),"""",R[-1]C-R[-2]C)"
    Range("FT28").FormulaR1C1 = "=IF(OR(R[-6]C[-1]=""N/A"",R[-6]C=""N/A""),"""",R[-1]C/SQRT(R[-1]C^2+R[-1]C[1]^2))"
    Range("FU28").FormulaR1C1 = "=IF(OR(R[-6]C[-2]=""N/A"",R[-6]C[-1]=""N/A""),"""",R[-1]C/SQRT(R[-1]C[-1]^2+R[-1]C^2))"
    Range("FT29").FormulaR1C1 = "=IF(OR(R[-7]C[-1]=""N/A"",R[-7]C=""N/A""),"""",-R[-4]C[1]/R[-1]C[1])"
    Range("FT30").FormulaR1C1 = "=IF(OR(R[-8]C[-1]=""N/A"",R[-8]C=""N/A""),"""",R[-5]C+R[-1]C*R[-2]C)"
    Range("FV29").FormulaR1C1 = _
        "=IF(OR(R[-7]C[-3]=""N/A"",R[-7]C[-2]=""N/A""),"""",IF(AND(R[-4]C[-1]<=0,R[-3]C[-1]>=0),RC[-2]/R[-7]C[-1],""""))"
    Range("FV30").FormulaR1C1 = _
        "=IF(OR(R[-8]C[-3]=""N/A"",R[-8]C[-2]=""N/A""),"""",IF(AND(R[-5]C[-1]<=0,R[-4]C[-1]>=0),(R[-8]C[-1]-R[-1]C[-2])/R[-8]C[-1],""""))"
    Range("FT33").FormulaR1C1 = "=IF(AND(R[-8]C[2]=""Good Start"",R[-7]C[2]=""Good Start""),R9C51,0)"
    Range("FU33").FormulaR1C1 = "=IF(RC[-1]=0,0,R9C56)"
    Range("FV33").FormulaR1C1 = "=YDWK2!R1169C3"
    Range("FT34").FormulaR1C1 = _
        "=IF(OR(R[-5]C[2]="""",R[-4]C[2]="""",R[-9]C[2]=""No Start"",R[-8]C[2]=""No Start""),0,(R[1]C-R[-1]C)*R[-5]C[2]+R[-1]C)"
    Range("FU34").FormulaR1C1 = "=IF(OR(R[-5]C[1]="""",R[-4]C[1]=""""),0,(R[1]C-R[-1]C)*R[-5]C[1]+R[-1]C)"
    Range("FT35").FormulaR1C1 = "=IF(AND(R[-10]C[2]=""Good Start"",R[-9]C[2]=""Good Start""),R9C60,0)"
    Range("FU35").FormulaR1C1 = "=IF(RC[-1]=0,0,R9C66)"
    Range("FV35").FormulaR1C1 = "=YDWK2!R1242C3"
    Range("FS38").FormulaR1C1 = "=R[-32]C[-9]"
    Range("FT38").FormulaR1C1 = "=R[-32]C[-20]"
    Range("FU38").FormulaR1C1 = "=R5C160"
    Range("FW38").FormulaR1C1 = "=IF(MAX(R5C29:R7C29)=0,YDWK2!R224C3,YDWK2!R[618]C[-176])"
    Range("FT41").FormulaR1C1 = _
        "=IF(R[-3]C[-1]=""N/A"",""N/A"",R[-3]C[-1]*(COS(RADIANS(R[-3]C[3]))*SIN(RADIANS(R[8]C[2]))-SIN(RADIANS(R[-3]C[3]))*COS(RADIANS(R[8]C[2]))))"
    Range("FU41").FormulaR1C1 = _
        "=IF(R[-3]C[-2]=""N/A"",""N/A"",R[-3]C[-2]*(SIN(RADIANS(R[-3]C[2]))*SIN(RADIANS(R[8]C[1]))+COS(RADIANS(R[-3]C[2]))*COS(RADIANS(R[8]C[1]))))"
    Range("FV41").FormulaR1C1 = _
        "=IF(OR(R[-3]C[-2]=0,R[-3]C[-3]=0,AND(RC[-1]<0,R[1]C[-1]>0)),""Good Finish"",""NO FINISH"")"
    Range("FT42").FormulaR1C1 = _
        "=IF(R[-4]C=""N/A"",""N/A"",R[-4]C*(COS(RADIANS(R[-4]C[3]))*SIN(RADIANS(R[9]C[2]))-SIN(RADIANS(R[-4]C[3]))*COS(RADIANS(R[9]C[2]))))"
    Range("FU42").FormulaR1C1 = _
        "=IF(R[-4]C[-1]=""N/A"",""N/A"",R[-4]C[-1]*(SIN(RADIANS(R[-4]C[2]))*SIN(RADIANS(R[9]C[1]))+COS(RADIANS(R[-4]C[2]))*COS(RADIANS(R[9]C[1]))))"
    Range("FV42").FormulaR1C1 = _
        "=IF(OR(R[-1]C[-1]=""N/A"",RC[-1]=""N/A""),""NO FINISH"",IF(ABS(R[4]C[-2])<0.5,""Good Finish"",""NO FINISH""))"
    Range("FT43").FormulaR1C1 = "=IF(OR(R[-5]C[-1]=""N/A"",R[-5]C=""N/A""),"""",R[-1]C-R[-2]C)"
    Range("FU43").FormulaR1C1 = "=IF(OR(R[-5]C[-2]=""N/A"",R[-5]C[-1]=""N/A""),"""",R[-1]C-R[-2]C)"
    Range("FT44").FormulaR1C1 = "=IF(OR(R[-6]C[-1]=""N/A"",R[-6]C=""N/A""),"""",R[-1]C/SQRT(R[-1]C^2+R[-1]C[1]^2))"
    Range("FU44").FormulaR1C1 = "=IF(OR(R[-6]C[-2]=""N/A"",R[-6]C[-1]=""N/A""),"""",R[-1]C/SQRT(R[-1]C[-1]^2+R[-1]C^2))"
    Range("FT45").FormulaR1C1 = "=IF(OR(R[-7]C[-1]=""N/A"",R[-7]C=""N/A""),"""",-R[-4]C[1]/R[-1]C[1])"
    Range("FV45").FormulaR1C1 = "=IF(OR(R[-7]C[-3]=""N/A"",R[-7]C[-2]=""N/A""),"""",IF(AND(R[-3]C[-1]>=0,R[-4]C[-1]<=0),RC[-2]/R[-7]C[-1],""""))"
    Range("FT46").FormulaR1C1 = "=IF(OR(R[-8]C[-1]=""N/A"",R[-8]C=""N/A""),"""",R[-5]C+R[-1]C*R[-2]C)"
    Range("FV46").FormulaR1C1 = _
        "=IF(OR(R[-8]C[-3]=""N/A"",R[-8]C[-2]=""N/A""),"""",IF(AND(R[-4]C[-1]>=0,R[-5]C[-1]<=0),(R[-8]C[-1]-R[-1]C[-2])/R[-8]C[-1],""""))"
    Range("FT49").FormulaR1C1 = _
        "=IF(AND(R[-8]C[2]=""Good Finish"",R[-7]C[2]=""Good Finish""),R9C164,0)"
    Range("FU49").FormulaR1C1 = "=IF(RC[-1]=0,0,R9C170)"
    Range("FV49").FormulaR1C1 = "=YDWK2!R1387C3"
    Range("FT50").FormulaR1C1 = _
        "=IF(OR(R[-5]C[2]="""",R[-4]C[2]="""",R[-9]C[2]=""No Finish"",R[-8]C[2]=""No Finish""),0,(R[1]C-R[-1]C)*R[-5]C[2]+R[-1]C)"
    Range("FU50").FormulaR1C1 = _
        "=IF(RC[-1]=R[-1]C[-1],R[-1]C,IF(RC[-1]=R[1]C[-1],R[1]C,IF(OR(R[-5]C[1]="""",R[-4]C[1]="""",R[-9]C[1]=""No Finish"",R[-8]C[1]=""No Finish""),0,(R[1]C-R[-1]C)*R[-5]C[1]+R[-1]C)))"
    Range("FT51").FormulaR1C1 = _
        "=IF(AND(R[-10]C[2]=""Good Finish"",R[-9]C[2]=""Good Finish""),R9C155,0)"
    Range("FU51").FormulaR1C1 = "=IF(RC[-1]=0,0,R9C160)"
    Range("FV51").FormulaR1C1 = "=YDWK2!R1314C3"
    Range("EV3").FormulaR1C1 = "=IF(SUMMARY!R20C8=0,0,MIN(R6C148,R[1]C[1],R[47]C[24]))"
    Range("EU3").FormulaR1C1 = "=IF(SUMMARY!R20C8=0,0,MAX(R[7]C[-18],R[1]C[2],R[47]C[25]))"
    Range("EU4").FormulaR1C1 = "=IF(AND(SUMMARY!R20C8>0,OR(SUMMARY!R20C8<R[-1]C,SUMMARY!R20C8<R[-1]C[1])),""MoP Finish"","""")"
    Range("FS54").FormulaR1C1 = "=R[-49]C[-123]"
    Range("FT54").FormulaR1C1 = "=R[-49]C[-114]"
    Range("FU54").FormulaR1C1 = "=R4C56"
    Range("FW54").FormulaR1C1 = "=IF(MAX(R5C29:R7C29)=0,YDWK2!R224C3-45,YDWK2!R368C3-45)"
    Range("FT57").FormulaR1C1 = "=IF(R[-3]C[-1]=""N/A"",""N/A"",R[-3]C[-1]*(COS(RADIANS(R[-3]C[3]))*SIN(RADIANS(R[8]C[2]))-SIN(RADIANS(R[-3]C[3]))*COS(RADIANS(R[8]C[2]))))"
    Range("FU57").FormulaR1C1 = "=IF(R[-3]C[-2]=""N/A"",""N/A"",R[-3]C[-2]*(SIN(RADIANS(R[-3]C[2]))*SIN(RADIANS(R[8]C[1]))+COS(RADIANS(R[-3]C[2]))*COS(RADIANS(R[8]C[1]))))"
    Range("FV57").FormulaR1C1 = "=IF(OR(R[-3]C[-3]=0,R[-3]C[-2]=0,AND(RC[-1]<0,R[1]C[-1]>0)),""Good Start"",""NO START"")"
    Range("FT58").FormulaR1C1 = "=IF(R[-4]C=""N/A"",""N/A"",R[-4]C*(COS(RADIANS(R[-4]C[3]))*SIN(RADIANS(R[9]C[2]))-SIN(RADIANS(R[-4]C[3]))*COS(RADIANS(R[9]C[2]))))"
    Range("FU58").FormulaR1C1 = "=IF(R[-4]C[-1]=""N/A"",""N/A"",R[-4]C[-1]*(SIN(RADIANS(R[-4]C[2]))*SIN(RADIANS(R[9]C[1]))+COS(RADIANS(R[-4]C[2]))*COS(RADIANS(R[9]C[1]))))"
    Range("FV58").FormulaR1C1 = "=IF(OR(RC[-1]=""N/A"",R[-1]C[-1]=""N/A""),""NO START"",IF(ABS(R[4]C[-2])<1,""Good Start"",""NO START""))"
    Range("FT59").FormulaR1C1 = "=IF(OR(R[-2]C=""N/A"",R[-1]C=""N/A""),"""",R[-1]C-R[-2]C)"
    Range("FU59").FormulaR1C1 = "=IF(OR(R[-5]C[-2]=""N/A"",R[-5]C[-1]=""N/A""),"""",R[-1]C-R[-2]C)"
    Range("FT60").FormulaR1C1 = "=IF(OR(R[-6]C[-1]=""N/A"",R[-6]C=""N/A""),"""",R[-1]C/SQRT(R[-1]C^2+R[-1]C[1]^2))"
    Range("FU60").FormulaR1C1 = "=IF(OR(R[-6]C[-2]=""N/A"",R[-6]C[-1]=""N/A""),"""",R[-1]C/SQRT(R[-1]C[-1]^2+R[-1]C^2))"
    Range("FT61").FormulaR1C1 = "=IF(OR(R[-7]C[-1]=""N/A"",R[-7]C=""N/A""),"""",-R[-4]C[1]/R[-1]C[1])"
    Range("FV61").FormulaR1C1 = "=IF(OR(R[-7]C[-3]=""N/A"",R[-7]C[-2]=""N/A""),"""",IF(AND(R[-4]C[-1]<=0,R[-3]C[-1]>=0),RC[-2]/R[-7]C[-1],""""))"
    Range("FT62").FormulaR1C1 = "=IF(OR(R[-8]C[-1]=""N/A"",R[-8]C=""N/A""),"""",R[-5]C+R[-1]C*R[-2]C)"
    Range("FV62").FormulaR1C1 = "=IF(OR(R[-8]C[-3]=""N/A"",R[-8]C[-2]=""N/A""),"""",IF(AND(R[-5]C[-1]<=0,R[-4]C[-1]>=0),(R[-8]C[-1]-R[-1]C[-2])/R[-8]C[-1],""""))"
    Range("FT65").FormulaR1C1 = "=IF(AND(R[-8]C[2]=""Good Start"",R[-7]C[2]=""Good Start""),R[-55]C[-129],0)"
    Range("FU65").FormulaR1C1 = "=IF(RC[-1]=0,0,R[-56]C[-120])"
    Range("FV65").FormulaR1C1 = "=YDWK2!R[885]C[-175]"
    Range("FT66").FormulaR1C1 = "=IF(OR(R[-5]C[2]="""",R[-4]C[2]="""",R[-9]C[2]=""No Start"",R[-8]C[2]=""No Start""),0,(R[1]C-R[-1]C)*R[-5]C[2]+R[-1]C)"
    Range("FU66").FormulaR1C1 = "=IF(OR(R[-5]C[1]="""",R[-4]C[1]=""""),0,(R[1]C-R[-1]C)*R[-5]C[1]+R[-1]C)"
    Range("FT67").FormulaR1C1 = "=IF(AND(R[-10]C[2]=""Good Start"",R[-9]C[2]=""Good Start""),R[-58]C[-115],0)"
    Range("FU67").FormulaR1C1 = "=IF(RC[-1]=0,0,R[-58]C[-110])"
    Range("FV67").FormulaR1C1 = "=YDWK2!R[883]C[-165]"
    Range("FT68").FormulaR1C1 = "=IF(AND(R[-2]C>0,R[14]C>0),MIN(R[-2]C,R[14]C),IF(AND(R[-61]C[-128]=R[-61]C[-127],R[-61]C[-127]=R[-61]C[-126]),R[-61]C[-126],0))"
    Range("FU68").FormulaR1C1 = "=IF(RC[-1]=R[-59]C[-128],R[-59]C[-129],IF(RC[-1]=0,0,IF(RC[-1]=R[-2]C[-1],R[-2]C,R[14]C)))"
    Range("FS70").FormulaR1C1 = "=R[-16]C"
    Range("FT70").FormulaR1C1 = "=R[-16]C"
    Range("FU70").FormulaR1C1 = "=R[-16]C"
    Range("FW70").FormulaR1C1 = "=IF(MAX(R5C29:R7C29)=0,YDWK2!R224C3+45,YDWK2!R368C3+45)"
    Range("FT73").FormulaR1C1 = "=IF(R[-3]C[-1]=""N/A"",""N/A"",R[-3]C[-1]*(COS(RADIANS(R[-3]C[3]))*SIN(RADIANS(R[8]C[2]))-SIN(RADIANS(R[-3]C[3]))*COS(RADIANS(R[8]C[2]))))"
    Range("FU73").FormulaR1C1 = "=IF(R[-3]C[-2]=""N/A"",""N/A"",R[-3]C[-2]*(SIN(RADIANS(R[-3]C[2]))*SIN(RADIANS(R[8]C[1]))+COS(RADIANS(R[-3]C[2]))*COS(RADIANS(R[8]C[1]))))"
    Range("FV73").FormulaR1C1 = "=IF(OR(R[-19]C[-3]=0,R[-19]C[-2]=0,AND(RC[-1]<0,R[1]C[-1]>0)),""Good Start"",""NO START"")"
    Range("FT74").FormulaR1C1 = "=IF(R[-4]C=""N/A"",""N/A"",R[-4]C*(COS(RADIANS(R[-4]C[3]))*SIN(RADIANS(R[9]C[2]))-SIN(RADIANS(R[-4]C[3]))*COS(RADIANS(R[9]C[2]))))"
    Range("FU74").FormulaR1C1 = "=IF(R[-4]C[-1]=""N/A"",""N/A"",R[-4]C[-1]*(SIN(RADIANS(R[-4]C[2]))*SIN(RADIANS(R[9]C[1]))+COS(RADIANS(R[-4]C[2]))*COS(RADIANS(R[9]C[1]))))"
    Range("FV74").FormulaR1C1 = "=IF(OR(RC[-1]=""N/A"",R[-1]C[-1]=""N/A""),""NO START"",IF(ABS(R[4]C[-2])<1,""Good Start"",""NO START""))"
    Range("FT75").FormulaR1C1 = "=IF(OR(R[-2]C=""N/A"",R[-1]C=""N/A""),"""",R[-1]C-R[-2]C)"
    Range("FU75").FormulaR1C1 = "=IF(OR(R[-5]C[-2]=""N/A"",R[-5]C[-1]=""N/A""),"""",R[-1]C-R[-2]C)"
    Range("FT76").FormulaR1C1 = "=IF(OR(R[-6]C[-1]=""N/A"",R[-6]C=""N/A""),"""",R[-1]C/SQRT(R[-1]C^2+R[-1]C[1]^2))"
    Range("FU76").FormulaR1C1 = "=IF(OR(R[-6]C[-2]=""N/A"",R[-6]C[-1]=""N/A""),"""",R[-1]C/SQRT(R[-1]C[-1]^2+R[-1]C^2))"
    Range("FT77").FormulaR1C1 = "=IF(OR(R[-7]C[-1]=""N/A"",R[-7]C=""N/A""),"""",-R[-4]C[1]/R[-1]C[1])"
    Range("FV77").FormulaR1C1 = "=IF(OR(R[-7]C[-3]=""N/A"",R[-7]C[-2]=""N/A""),"""",IF(AND(R[-4]C[-1]<=0,R[-3]C[-1]>=0),RC[-2]/R[-7]C[-1],""""))"
    Range("FT78").FormulaR1C1 = "=IF(OR(R[-8]C[-1]=""N/A"",R[-8]C=""N/A""),"""",R[-5]C+R[-1]C*R[-2]C)"
    Range("FV78").FormulaR1C1 = "=IF(OR(R[-8]C[-3]=""N/A"",R[-8]C[-2]=""N/A""),"""",IF(AND(R[-5]C[-1]<=0,R[-4]C[-1]>=0),(R[-8]C[-1]-R[-1]C[-2])/R[-8]C[-1],""""))"
    Range("FT81").FormulaR1C1 = "=IF(AND(R[-8]C[2]=""Good Start"",R[-7]C[2]=""Good Start""),R[-71]C[-129],0)"
    Range("FU81").FormulaR1C1 = "=IF(RC[-1]=0,0,R[-72]C[-120])"
    Range("FV81").FormulaR1C1 = "=R[-16]C"
    Range("FT82").FormulaR1C1 = "=IF(OR(R[-5]C[2]="""",R[-4]C[2]="""",R[-9]C[2]=""No Start"",R[-8]C[2]=""No Start""),0,(R[1]C-R[-1]C)*R[-5]C[2]+R[-1]C)"
    Range("FU82").FormulaR1C1 = "=IF(OR(R[-5]C[1]="""",R[-4]C[1]=""""),0,(R[1]C-R[-1]C)*R[-5]C[1]+R[-1]C)"
    Range("FT83").FormulaR1C1 = "=IF(AND(R[-10]C[2]=""Good Start"",R[-9]C[2]=""Good Start""),R[-74]C[-115],0)"
    Range("FU83").FormulaR1C1 = "=IF(RC[-1]=0,0,R[-74]C[-110])"
    Range("FV83").FormulaR1C1 = "=R[-16]C"
    Range("FY54").FormulaR1C1 = "=R[-49]C[-15]"
    Range("FZ54").FormulaR1C1 = "=R[-49]C[-26]"
    Range("GA54").FormulaR1C1 = "=R4C160"
    Range("GC54").FormulaR1C1 = "=IF(MAX(R5C29:R7C29)=0,YDWK2!R224C3-45,YDWK2!R[602]C[-182]-45)"
    Range("FZ57").FormulaR1C1 = "=IF(R[-3]C[-1]=""N/A"",""N/A"",R[-3]C[-1]*(COS(RADIANS(R[-3]C[3]))*SIN(RADIANS(R[8]C[2]))-SIN(RADIANS(R[-3]C[3]))*COS(RADIANS(R[8]C[2]))))"
    Range("GA57").FormulaR1C1 = "=IF(R[-3]C[-2]=""N/A"",""N/A"",R[-3]C[-2]*(SIN(RADIANS(R[-3]C[2]))*SIN(RADIANS(R[8]C[1]))+COS(RADIANS(R[-3]C[2]))*COS(RADIANS(R[8]C[1]))))"
    Range("GB57").FormulaR1C1 = "=IF(OR(R[-3]C[-3]=0,R[-3]C[-2]=0,AND(RC[-1]<0,R[1]C[-1]>0)),""Good Finish"",""NO FINISH"")"
    Range("FZ58").FormulaR1C1 = "=IF(R[-4]C=""N/A"",""N/A"",R[-4]C*(COS(RADIANS(R[-4]C[3]))*SIN(RADIANS(R[9]C[2]))-SIN(RADIANS(R[-4]C[3]))*COS(RADIANS(R[9]C[2]))))"
    Range("GA58").FormulaR1C1 = "=IF(R[-4]C[-1]=""N/A"",""N/A"",R[-4]C[-1]*(SIN(RADIANS(R[-4]C[2]))*SIN(RADIANS(R[9]C[1]))+COS(RADIANS(R[-4]C[2]))*COS(RADIANS(R[9]C[1]))))"
    Range("GB58").FormulaR1C1 = "=IF(OR(R[-4]C[-3]=""N/A"",R[-4]C[-2]=""N/A""),""NO FINISH"",IF(ABS(R[4]C[-2])<1,""Good Finish"",""NO FINISH""))"
    Range("FZ59").FormulaR1C1 = "=IF(OR(R[-5]C[-1]=""N/A"",R[-5]C=""N/A""),"""",R[-1]C-R[-2]C)"
    Range("GA59").FormulaR1C1 = "=IF(OR(R[-5]C[-2]=""N/A"",R[-5]C[-1]=""N/A""),"""",R[-1]C-R[-2]C)"
    Range("FZ60").FormulaR1C1 = "=IF(OR(R[-6]C[-1]=""N/A"",R[-6]C=""N/A""),"""",R[-1]C/SQRT(R[-1]C^2+R[-1]C[1]^2))"
    Range("GA60").FormulaR1C1 = "=IF(OR(R[-6]C[-2]=""N/A"",R[-6]C[-1]=""N/A""),"""",R[-1]C/SQRT(R[-1]C[-1]^2+R[-1]C^2))"
    Range("FZ61").FormulaR1C1 = "=IF(OR(R[-7]C[-1]=""N/A"",R[-7]C=""N/A""),"""",-R[-4]C[1]/R[-1]C[1])"
    Range("GB61").FormulaR1C1 = "=IF(OR(R[-7]C[-3]=""N/A"",R[-7]C[-2]=""N/A""),"""",IF(AND(R[-3]C[-1]>=0,R[-4]C[-1]<=0),RC[-2]/R[-7]C[-1],""""))"
    Range("FZ62").FormulaR1C1 = "=IF(OR(R[-8]C[-1]=""N/A"",R[-8]C=""N/A""),"""",R[-5]C+R[-1]C*R[-2]C)"
    Range("GB62").FormulaR1C1 = "=IF(OR(R[-8]C[-3]=""N/A"",R[-8]C[-2]=""N/A""),"""",IF(AND(R[-4]C[-1]>=0,R[-5]C[-1]<=0),(R[-8]C[-1]-R[-1]C[-2])/R[-8]C[-1],""""))"
    Range("FZ65").FormulaR1C1 = "=IF(AND(R[-8]C[2]=""Good Finish"",R[-7]C[2]=""Good Finish""),R[-56]C[-17],0)"
    Range("GA65").FormulaR1C1 = "=IF(RC[-1]=0,0,R[-56]C[-12])"
    Range("GB65").FormulaR1C1 = "=YDWK2!R[811]C[-171]"
    Range("FZ66").FormulaR1C1 = "=IF(OR(R[-5]C[2]="""",R[-4]C[2]="""",R[-9]C[2]=""No Finish"",R[-8]C[2]=""No Finish""),0,(R[1]C-R[-1]C)*R[-5]C[2]+R[-1]C)"
    Range("GA66").FormulaR1C1 = "=IF(RC[-1]=R[-1]C[-1],R[-1]C,IF(RC[-1]=R[1]C[-1],R[1]C,IF(OR(R[-5]C[1]="""",R[-4]C[1]="""",R[-9]C[1]=""No Finish"",R[-8]C[1]=""No Finish""),0,(R[1]C-R[-1]C)*R[-5]C[1]+R[-1]C)))"
    Range("FZ67").FormulaR1C1 = "=IF(AND(R[-10]C[2]=""Good Finish"",R[-9]C[2]=""Good Finish""),R[-58]C[-30],0)"
    Range("GA67").FormulaR1C1 = "=IF(RC[-1]=0,0,R[-58]C[-22])"
    Range("GB67").FormulaR1C1 = "=YDWK2!R[809]C[-181]"
    Range("FZ68").FormulaR1C1 = "=IF(AND(R[-2]C>0,R[14]C>0),MAX(R[-2]C,R[14]C),IF(AND(R[-62]C[-33]=R6C148,R[-62]C[-33]=R[-59]C[-30]),R[-59]C[-30],0))"
    Range("GA68").FormulaR1C1 = "=IF(RC[-1]=R[-59]C[-31],R[-59]C[-30],IF(RC[-1]=0,0,IF(RC[-1]=R[-2]C[-1],R[-2]C,R[14]C)))"
    Range("FY70").FormulaR1C1 = "=R[-16]C"
    Range("FZ70").FormulaR1C1 = "=R[-16]C"
    Range("GA70").FormulaR1C1 = "=R4C160"
    Range("GC70").FormulaR1C1 = "=IF(MAX(R5C29:R7C29)=0,YDWK2!R224C3+45,YDWK2!R[586]C[-182]+45)"
    Range("FZ73").FormulaR1C1 = "=IF(R[-3]C[-1]=""N/A"",""N/A"",R[-3]C[-1]*(COS(RADIANS(R[-3]C[3]))*SIN(RADIANS(R[8]C[2]))-SIN(RADIANS(R[-3]C[3]))*COS(RADIANS(R[8]C[2]))))"
    Range("GA73").FormulaR1C1 = "=IF(R[-3]C[-2]=""N/A"",""N/A"",R[-3]C[-2]*(SIN(RADIANS(R[-3]C[2]))*SIN(RADIANS(R[8]C[1]))+COS(RADIANS(R[-3]C[2]))*COS(RADIANS(R[8]C[1]))))"
    Range("GB73").FormulaR1C1 = "=IF(OR(R[-3]C[-3]=0,R[-3]C[-2]=0,AND(RC[-1]<0,R[1]C[-1]>0)),""Good Finish"",""NO FINISH"")"
    Range("FZ74").FormulaR1C1 = "=IF(R[-4]C=""N/A"",""N/A"",R[-4]C*(COS(RADIANS(R[-4]C[3]))*SIN(RADIANS(R[9]C[2]))-SIN(RADIANS(R[-4]C[3]))*COS(RADIANS(R[9]C[2]))))"
    Range("GA74").FormulaR1C1 = "=IF(R[-4]C[-1]=""N/A"",""N/A"",R[-4]C[-1]*(SIN(RADIANS(R[-4]C[2]))*SIN(RADIANS(R[9]C[1]))+COS(RADIANS(R[-4]C[2]))*COS(RADIANS(R[9]C[1]))))"
    Range("GB74").FormulaR1C1 = "=IF(OR(R[-4]C[-3]=""N/A"",R[-4]C[-2]=""N/A""),""NO FINISH"",IF(ABS(R[4]C[-2])<1,""Good Finish"",""NO FINISH""))"
    Range("FZ75").FormulaR1C1 = "=IF(OR(R[-5]C[-1]=""N/A"",R[-5]C=""N/A""),"""",R[-1]C-R[-2]C)"
    Range("GA75").FormulaR1C1 = "=IF(OR(R[-5]C[-2]=""N/A"",R[-5]C[-1]=""N/A""),"""",R[-1]C-R[-2]C)"
    Range("FZ76").FormulaR1C1 = "=IF(OR(R[-6]C[-1]=""N/A"",R[-6]C=""N/A""),"""",R[-1]C/SQRT(R[-1]C^2+R[-1]C[1]^2))"
    Range("GA76").FormulaR1C1 = "=IF(OR(R[-6]C[-2]=""N/A"",R[-6]C[-1]=""N/A""),"""",R[-1]C/SQRT(R[-1]C[-1]^2+R[-1]C^2))"
    Range("FZ77").FormulaR1C1 = "=IF(OR(R[-7]C[-1]=""N/A"",R[-7]C=""N/A""),"""",-R[-4]C[1]/R[-1]C[1])"
    Range("GB77").FormulaR1C1 = "=IF(OR(R[-7]C[-3]=""N/A"",R[-7]C[-2]=""N/A""),"""",IF(AND(R[-3]C[-1]>=0,R[-4]C[-1]<=0),RC[-2]/R[-7]C[-1],""""))"
    Range("FZ78").FormulaR1C1 = "=IF(OR(R[-8]C[-1]=""N/A"",R[-8]C=""N/A""),"""",R[-5]C+R[-1]C*R[-2]C)"
    Range("GB78").FormulaR1C1 = "=IF(OR(R[-8]C[-3]=""N/A"",R[-8]C[-2]=""N/A""),"""",IF(AND(R[-4]C[-1]>=0,R[-5]C[-1]<=0),(R[-8]C[-1]-R[-1]C[-2])/R[-8]C[-1],""""))"
    Range("FZ81").FormulaR1C1 = "=IF(AND(R[-8]C[2]=""Good Finish"",R[-7]C[2]=""Good Finish""),R[-72]C[-17],0)"
    Range("GA81").FormulaR1C1 = "=IF(RC[-1]=0,0,R[-72]C[-12])"
    Range("GB81").FormulaR1C1 = "=R[-16]C"
    Range("FZ82").FormulaR1C1 = "=IF(OR(R[-5]C[2]="""",R[-4]C[2]="""",R[-9]C[2]=""No Finish"",R[-8]C[2]=""No Finish""),0,(R[1]C-R[-1]C)*R[-5]C[2]+R[-1]C)"
    Range("GA82").FormulaR1C1 = "=IF(RC[-1]=R[-1]C[-1],R[-1]C,IF(RC[-1]=R[1]C[-1],R[1]C,IF(OR(R[-5]C[1]="""",R[-4]C[1]="""",R[-9]C[1]=""No Finish"",R[-8]C[1]=""No Finish""),0,(R[1]C-R[-1]C)*R[-5]C[1]+R[-1]C)))"
    Range("FZ83").FormulaR1C1 = "=IF(AND(R[-10]C[2]=""Good Finish"",R[-9]C[2]=""Good Finish""),R[-74]C[-30],0)"
    Range("GA83").FormulaR1C1 = "=IF(RC[-1]=0,0,R[-74]C[-22])"
    Range("GB83").FormulaR1C1 = "=R[-16]C"
End If
End With
    'AU5 Added 7/27/2015 for START Needs to be at end!
    Range("AU5").FormulaR1C1 = _
        "=IF(MAX(R4C86,R6C87,R4C102,R6C103,R4C118,R6C119)=0,"""",IF(MIN(R4C86,R6C87,R4C102,R6C103,R4C118,R6C119)<>0,MIN(R4C86,R6C87,R4C102,R6C103,R4C118,R6C119),IF(SMALL((R4C86,R6C87,R4C102,R6C103,R4C118,R6C119),2)<>0,SMALL((R4C86,R6C87,R4C102,R6C103,R4C118,R6C119),2),IF(SMALL((R4C86,R6C87,R4C102,R6C103,R4C118,R6C119),3)<>0,SMALL((R4C86,R6C87,R4C102,R6C103,R4C118,R6C119),3),IF(SMALL((R4C86,R6C87,R4C102,R6C103,R4C118,R6C119),4)<>0,SMALL((R4C86,R6C87,R4C102,R6C103,R4C118,R6C119),4),IF(SMALL((R4C86,R6C87,R4C102,R6C103,R4C118,R6C119),5)<>0,SMALL((R4C86,R6C87,R4C102,R6C103,R4C118,R6C119),5),IF(SMALL((R4C86,R6C87,R4C102,R6C103,R4C118,R6C119),5)=0,SMALL((R4C86,R6C87,R4C102,R6C103,R4C118,R6C119),6))))))))"
    
    Application.Calculation = xlCalculationAutomatic
    Sheets("YDWK2").Range("E1505").FormulaR1C1 = _
        "=R[-34]C[-2]+(1-R[-2]C)*R[-35]C[-2]*R[-14]C*(ACOS(R[-19]C)+R[-2]C*SIN(ACOS(R[-19]C))*(R[-12]C+R[-2]C*R[-19]C*(-1+2*R[-12]C*R[-12]C)))"
    Sheets("YDWK2").Calculate
    Sheets("TPOrder").Calculate
    Sheets("Summary").Calculate
    Sheets("Verify Task").Calculate
    'Application.CalculateFull
    Sheets("TPOrder").Protect Password:="spike"
    Sheets("Verify Task").Visible = True
    Sheets("Verify Task").Select
    ActiveWindow.DisplayVerticalScrollBar = False
    ActiveWindow.WindowState = xlMaximized
    Sheets("Verify Task").Unprotect Password:="spike"
    Sheets("Verify Task").Range("H1").Value = Sheets("YDWK1").Range("J1").Value
    If Range("H2") = "" Then
        ActiveSheet.Shapes("Drop Down 1").Visible = False
        Range("G2").Locked = True
    ElseIf Range("H2") <> "" Then
        ActiveSheet.Shapes("Drop Down 1").Visible = True
        Range("G2").Locked = False
    End If
    If Range("E10") = "No Turn Points Declared" Then
        Application.Calculation = xlCalculationAutomatic
        Range("F12,F14,F16").Locked = True
        ActiveSheet.Shapes("TextBox 2").Visible = False
        ActiveSheet.Shapes("Rounded Rectangle 3").Visible = False
        Range("B25").Value = "            Click on the glider to continue"
    ElseIf Range("E10") <> "No Turn Points Declared" Then
        Application.Calculation = xlCalculationManual
        Range("F12").Locked = False
        ActiveSheet.Shapes("TextBox 2").Visible = True
        ActiveSheet.Shapes("Rounded Rectangle 3").Visible = True
        If Range("E14") <> "N/A" Then
            Range("F14").Locked = False
        End If
        If Range("E16") <> "N/A" Then
            Range("F16").Locked = False
        End If
    End If
    Application.DisplayFullScreen = True
    Range("A1:I31").Select
    ActiveWindow.Zoom = True
    Range("A1").Select
    Sheets("Verify Task").Protect Password:="spike"
    Application.Cursor = xlDefault
    Workbooks("C.xlsm").Protect Password:="spike"
    Application.DisplayAlerts = False
    Workbooks("A.xlsm").Close
    Application.ScreenUpdating = True
    'Workbooks("C.xlsm").Application.Calculation = xlCalculationManual
End Sub