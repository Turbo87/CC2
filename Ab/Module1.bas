' VBA Module: Data Logger Processing and Analysis
' Purpose: Processes logger data from flight recording devices, performs data parsing,
' time calculations, and format conversions for flight analysis. Handles data sorting,
' filtering, and preparation for import into analysis worksheets.
' Contains functions for processing large datasets (10K-60K records) with time-based calculations.

Option Explicit
Dim myFile As Variant
#If VBA7 And Win64 Then
    Dim LastRow As LongPtr
#Else
    Dim LastRow As Long
#End If

Sub NewENLA()
    ' NewENLA processes Engine Noise Level (ENL) data from IGC flight files
    ' This function analyzes motor operation periods for self-launching gliders
    ' It extracts and processes both I-records (intermittent data) and H-records (header data)
    ' The function determines Motor Operating Periods (MoP) and engine noise levels for competition verification
    '
    ' JLR 3/27/12 7/26/16 Corrected O49 on Imp (no = sign)
    ' JLR 3/2/17 Added Range 07 now used in Corrected Z11:Z10010 & AL25:AL10010; corrected AM25:AM10010
    ' JLR 6/3/17 Amended '' for interim MoP @ ZZ25, AA3, AA4, AA25, AL25 Works for Mirja
    ' JLR 7/6/17 Amended Z2,Z3,Z4,Z25:et al & AL4,AL5 & AN4 for NON-interim MoP (eg: engine run before task ONLY) for Sibylle Andresen, July 2 & 3 WORLD RECORDS!!
    ' JLR 7/7/17 Amended AA3 & AL4 to differentiate between Mirja (re-start after task) from Sibylle (No re-start after task); re-activated O11:AB10010 value
    ' JLR 11/30/2017 Amended AL4 for Schart (no re-start after second ENL); Z2,Z4,Z25,AL4,AN4,AO11 for Anja
    ' JLR 07/14/2018 AN11 amended for landing time (Howard2)

    'Workbooks("Ab.xlsm").Activate
    'ActiveWorkbook.Unprotect Password:="spike"
    'Sheets("BR").Select
    'Range("A1").Select
    'ActiveSheet.Paste
    'Application.CutCopyMode = False
    'Workbooks("A.xlsm").Activate
    'ActiveWindow.WindowState = xlMaximized
    'Application.ScreenUpdating = True
    'ActiveWindow.DisplayWorkbookTabs = False
    ' Disable screen updating for performance during bulk operations
    Application.ScreenUpdating = False
    Workbooks("Ab.xlsm").Activate
    ' Convert formulas to values in raw IGC data range
    Sheets("BR").Range("A1:A60000") = Sheets("BR").Range("A1:A60000").Value
    ' Parse IGC records by splitting first character (record type) from data
    Columns("A:A").TextToColumns Destination:=Range("B1"), DataType:=xlFixedWidth, FieldInfo:=Array(Array(0, 1), Array(1, 1)), TrailingMinusNumbers:=True
    ' Extract non-B records (everything except GPS position fixes)
    Range("D1:D1001").FormulaR1C1 = "=IF(RC[-2]<>""B"",RC[-3],"""")"
    ' Mark rows containing B-records for counting
    Range("E3:E60002").FormulaR1C1 = "=IF(RC[-3]=""B"",1,"""")"
    Range("E3:E60002") = Range("E3:E60002").Value
    ' Count total number of B-records in the flight file
    Range("E1").FormulaR1C1 = "=SUM(R[3]C:R[60000]C)"
    Range("E1").Value = Range("E1").Value
    ' Parse first B-record for time reference
    Range("E2") = Range("D2").Value
    Range("E2").TextToColumns Destination:=Range("E2"), DataType:=xlFixedWidth, _
        OtherChar:="E", FieldInfo:=Array(Array(0, 9), Array(5, 4)), _
        TrailingMinusNumbers:=True
    ' Extract B-record data where marked
    Range("F3:F60002").FormulaR1C1 = "=IF(RC[-1]=1,RC[-3],"""")"
    Range("F3:F60002") = Range("F3:F60002").Value
    ' Parse B-records into time components (hours, minutes, seconds) and position data
    Columns("F:F").TextToColumns Destination:=Range("F1"), DataType:=xlFixedWidth, _
        OtherChar:="E", FieldInfo:=Array(Array(0, 1), Array(2, 1), Array(4, 1), Array(6, 9 _
        )), TrailingMinusNumbers:=True
    ' Find maximum time value in B-record hours column
    Range("F2").FormulaR1C1 = "=MAX(R[2]C:R[49998]C)"
    Range("F2").Value = Range("F2").Value
    ' Identify B-records that follow 7+ consecutive non-B records (flight restart detection)
    Range("E8:E600").FormulaR1C1 = "=IF(AND(R[-1]C[-3]<>""B"",R[-2]C[-3]<>""B"",R[-3]C[-3]<>""B"",R[-4]C[-3]<>""B"",R[-5]C[-3]<>""B"",R[-6]C[-3]<>""B"",R[-7]C[-3]<>""B"",RC[-3]=""B""),RC[1],"""")"
    ' Find maximum time among flight restart candidates
    Range("F3").FormulaR1C1 = "=MAX(R[5]C[-1]:R[597]C[-1])"
    Range("F3").Value = Range("F3").Value
    ' Convert time components to Excel time format, handling day rollover
    ' If time is within normal range, add base time; if rolled over, add 1 day
    Range("J2:J60001").FormulaR1C1 = _
        "=IF(RC[-2]="""","""",IF(AND(RC[-4]<=R2C6,RC[-4]>=R3C6),TIME(RC[-4],RC[-3],RC[-2])+R2C5,TIME(RC[-4],RC[-3],RC[-2])+R2C5+1))"
    ' Copy record type indicator for each time entry
    Range("K2:K60001").FormulaR1C1 = "=IF(RC[-1]="""","""",RC[-8])"
    Range("J2:K60001") = Range("J2:K60001").Value
    ' Sort all data by time to create chronological sequence
    Range("J1:K60000").Sort Key1:=Range("J1"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal

    Range("L1").FormulaR1C1 = "1"
    Range("L2:L60000").FormulaR1C1 = "=IF(RC[-1]=""""=FALSE,R[-1]C+1,"""")"
    Range("L2:L60000") = Range("L2:L60000").Value

    If Range("E1") > 40000 Then
        Range("P6:P60000").FormulaR1C1 = "=IF(SUM(R[-5]C:R[-1]C)=0,1,0)"
        Range("P6:P60000") = Range("P6:P60000").Value

    ElseIf Range("E1") > 30000 Then
        Range("O4:O40000").FormulaR1C1 = "=IF(SUM(R[-3]C:R[-1]C)=0,1,0)"
        Range("O4:O40000") = Range("O4:O40000").Value

    ElseIf Range("E1") > 20000 Then
        Range("N3:N30000").FormulaR1C1 = "=IF(SUM(R[-2]C:R[-1]C)>0,0,1)"
        Range("N3:N30000") = Range("N3:N30000").Value

    ElseIf Range("E1") > 10000 Then
        Range("M2:M20000").FormulaR1C1 = "=IF(R[-1]C=0,1,0)"
        Range("M2:M20000") = Range("M2:M20000").Value
    End If

    Range("Q1").FormulaR1C1 = "1"
    Range("Q2:Q60000").FormulaR1C1 = "=IF(RC[-5]="""","""",IF(R1C5<=10000,RC[-5],IF(AND(R1C5>10000,R1C5<=20000,RC[-4]=1),RC[-5],IF(AND(R1C5>20000,R1C5<=30000,RC[-3]=1),RC[-5],IF(AND(R1C5>30000,R1C5<=40000,RC[-2]=1),RC[-5],IF(AND(R1C5>40000,R1C5<=60000,RC[-1]=1),RC[-5],""""))))))"
    Range("Q1:Q60000") = Range("Q1:Q60000").Value
    Range("R1:R60000").FormulaR1C1 = "=IF(RC[-1]=RC[-6],RC[-7],"""")"
    Range("R1:R60000") = Range("R1:R60000").Value
    Range("Q1:R60000").Sort Key1:=Range("Q1"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal

    Range("X1:X10000").Value = Range("R1:R10000").Value
    Columns("X:X").TextToColumns Destination:=Range("W1"), DataType:=xlFixedWidth, _
        OtherChar:="E", FieldInfo:=Array(Array(0, 1), Array(6, 1), Array(8, 1), Array(10, _
        1), Array(13, 1), Array(14, 1), Array(17, 1), Array(19, 1), Array(22, 1), Array(23, 1)), _
        TrailingMinusNumbers:=True
    Range("AK1:AL10000").Value = Range("Q1:R10000").Value
    Range("AK1:AL10000").Sort Key1:=Range("AK1"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal

    Sheets("IMP").Range("A1:A1000").Value = Sheets("BR").Range("D1:D1000").Value

    ' Switch to IMP sheet for H-record processing
    Sheets("IMP").Select
    ' Copy non-B record data for header analysis
    Range("B1:B1000") = Range("A1:A1000").Value
    ' Parse record type from data content
    Columns("B:B").TextToColumns Destination:=Range("B1"), DataType:=xlFixedWidth, _
        FieldInfo:=Array(Array(0, 1), Array(1, 1)), TrailingMinusNumbers:=True
    ' Extract first character and remaining data from first record
    Range("C1").TextToColumns Destination:=Range("C1"), DataType:=xlFixedWidth, _
        OtherChar:="E", FieldInfo:=Array(Array(0, 1), Array(1, 9)), _
        TrailingMinusNumbers:=True
        ' Handle different H-record formats based on first character
        If Range("C1") = "X" Then
            ' Process X-type records (manufacturer extension format)
            Range("D1").Value = Range("A1").Value
            Range("D1").TextToColumns Destination:=Range("D1"), DataType:=xlFixedWidth, _
            OtherChar:=":", FieldInfo:=Array(Array(0, 9), Array(2, 1), Array(4, 1)), _
            TrailingMinusNumbers:=True
        ElseIf Range("C1") = "X" = False Then
            ' Process standard H-record format
            Range("D1").Value = Range("A1").Value
            Range("D1").TextToColumns Destination:=Range("D1"), DataType:=xlFixedWidth, _
            OtherChar:="E", FieldInfo:=Array(Array(0, 9), Array(1, 1), Array(4, 1), Array(7, 9)), TrailingMinusNumbers:=True
        End If
    Range("D2").FormulaR1C1 = "=IF(RC[-2]=""H"",RC[-1],"""")"
    Range("D2") = Range("D2").Value
    Range("D2").TextToColumns Destination:=Range("D2"), DataType:=xlFixedWidth, _
        OtherChar:=":", FieldInfo:=Array(Array(0, 1), Array(4, 4)), _
        TrailingMinusNumbers:=True
    Range("C4:C50").Select
    Selection.TextToColumns Destination:=Range("C4"), DataType:=xlFixedWidth, _
        OtherChar:=":", FieldInfo:=Array(Array(0, 1), Array(4, 2)), _
        TrailingMinusNumbers:=True
    Range("D4:D50").Select
    Selection.TextToColumns Destination:=Range("D4"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :=":", FieldInfo:=Array(Array(1, 2), Array(2, 2)), TrailingMinusNumbers:=True

   'PUT H RECORDS IN TECH SPECS ORDER 4/16/15
    Range("R1").Value = "FPLT"
    Range("S1").Value = "FCM2"
    Range("T1").Value = "FGTY"
    Range("U1").Value = "FGID"
    Range("V1").Value = "FDTM"
    Range("W1").Value = "FRFW"
    Range("X1").Value = "FRHW"
    Range("Y1").Value = "FFTY"
    Range("Z1").Value = "FGPS"
    Range("AA1").Value = "FPRS"

    Range("L4:L50").FormulaR1C1 = "=IF(RC2=""H"",RC3,"""")"
    Range("M4:M50").FormulaR1C1 = "=IF(RC2=""H"",RC4,"""")"
    Range("N4:N50").FormulaR1C1 = "=IF(RC2=""H"",RC5,"""")"

    Range("L4:N50").Value = Range("L4:N50").Value

    Range("R4:R50").FormulaR1C1 = "=IF(OR(RC12=R1C,RC12=""OPLT"",RC12=""PPLT""),1,"""")"
    Range("S4:S50").FormulaR1C1 = "=IF(OR(RC12=R1C,RC12=""OCM2"",RC12=""PCM2""),2,"""")"
    Range("T4:T50").FormulaR1C1 = "=IF(OR(RC12=R1C,R24C12=""OGTY"",R24C12=""PGTY""),3,"""")"
    Range("U4:U50").FormulaR1C1 = "=IF(OR(RC12=R1C,RC12=""OGID"",RC12=""PGID""),4,"""")"
    Range("V4:V50").FormulaR1C1 = "=IF(RC12=R1C,5,"""")"
    Range("W4:W50").FormulaR1C1 = "=IF(RC12=R1C,6,"""")"
    Range("X4:X50").FormulaR1C1 = "=IF(RC12=R1C,7,"""")"
    Range("Y4:Y50").FormulaR1C1 = "=IF(RC12=R1C,8,"""")"
    Range("Z4:Z50").FormulaR1C1 = "=IF(RC12=R1C,9,"""")"
    Range("AA4:AA50").FormulaR1C1 = "=IF(RC12=R1C,10,"""")"

    Range("R4:AA50").Value = Range("R4:AA50").Value

    Range("O4:O50").FormulaR1C1 = "=IF(MAX(RC[3]:RC[13])=0,"""",MAX(RC[3]:RC[13]))"
    Range("O4:O50").Value = Range("O4:O50").Value

    Range("L4:O50").Sort Key1:=Range("O10"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    'Amended 4/30/15 delete former ref to FW/HW conundrum, go with order only; revised 5/5/15 for strict order; no FFTY added 5/11/15
    Range("O45").FormulaR1C1 = "=IF(R4C=1,"""",1)"
    Range("O46").FormulaR1C1 = "=IF(R5C=2,"""",2)"
    Range("O47").FormulaR1C1 = "=IF(OR(R4C=3,R5C=3,R6C=3),"""",IF(AND(R6C=4,R7C=5),"""",3))"
    Range("O48").FormulaR1C1 = "=IF(OR(R4C=4,R5C=4,R6C=4,R7C=4),"""",IF(AND(R7C=5,R8C=6),"""",4))"
    Range("O49").FormulaR1C1 = "=IF(MAX(R4C:R13C)=7,8,IF(MAX(R4C:R13C)=8,9,IF(MAX(R4C:R13C)=9,10,"""")))"
    Range("O45:O49").Value = Range("O45:O49").Value

    Range("L4:O50").Sort Key1:=Range("O10"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    Range("C4:E13").Value = Range("L4:N13").Value
    Range("R1:AB1,J9:J10,L4:AB50").Clear

    ' Extract I-records (intermittent data records containing ENL information)
    Range("G2:G50").FormulaR1C1 = "=IF(RC[-5]=""I"",RC[-6],"""")"
    Range("G2:G50").Value = Range("G2:G50").Value
    ' Count I-records to determine if ENL data is available
    Range("H2:H50").FormulaR1C1 = "=IF(RC[-6]<>""I"",0,1)"
    Range("H1").FormulaR1C1 = "=SUM(R[1]C:R[49]C)"
    Range("H1").Value = Range("H1").Value
    ' Process ENL data if I-records are found
    If Range("H1") > 0 Then
     Range("H2").FormulaR1C1 = "ENL"
     Range("G3:G50").FormulaR1C1 = "=IF(RC[-5]<>""I"","""",IF(ISERROR(FIND(R2C[1],RC[-6]))=TRUE,"""",RC[-6]))"
     Range("G3:G50").Value = Range("G3:G50").Value
     Range("H3:H50").FormulaR1C1 = "=IF(RC[-1]="""","""",FIND(R2C,RC[-1]))"
     Range("I2").FormulaR1C1 = "=MAX(R[1]C[-1]:R[48]C[-1])-4"
        If Range("I2") > 0 Then
        Range("I3:I50").FormulaR1C1 = "=MID(RC[-2],R2C,2)"
        Range("I3:I50").Value = Range("I3:I50").Value
        Range("Q1").FormulaR1C1 = "=MAX(R[2]C[-8]:R[49]C[-8])"
        Range("Q1").Value = Range("Q1").Value
        Else: Range("Q1") = 0
        End If
    Range("H2").FormulaR1C1 = "MOP"
    Range("G3:G50").FormulaR1C1 = "=IF(RC[-5]<>""I"","""",IF(ISERROR(FIND(R2C[1],RC[-6]))=TRUE,"""",RC[-6]))"
    Range("G3:G50").Value = Range("G3:G50").Value
    Range("H3:H50").FormulaR1C1 = "=IF(RC[-1]="""","""",FIND(R2C,RC[-1]))"
    Range("I2").FormulaR1C1 = "=MAX(R[1]C[-1]:R[48]C[-1])-4"
        If Range("I2") > 0 Then
        Range("I3:I50").FormulaR1C1 = "=MID(RC[-2],R2C,2)"
        Range("I3:I50").Value = Range("I3:I50").Value
        Range("Q2").FormulaR1C1 = "=MAX(R[2]C[-8]:R[49]C[-8])"
        Range("Q2").Value = Range("Q2").Value
        Else: Range("Q2") = 0
        End If
    End If
    Range("G1").FormulaR1C1 = "='[A.xlsm]ALL CLAIMS'!R18C6"

    ' Switch to main data sheet for coordinate processing
    Sheets("Sheet1").Select
    ' Determine ENL status based on I-record availability
    Range("A1").FormulaR1C1 = "=IF(IMP!R1C8=0,""NO I Record"",IF(IMP!R1C17=0,""No ENL"",""ENL""))"
    Sheets("Sheet1").Range("A11:A10010").Value = Sheets("BR").Range("AL1:AL10000").Value
    Range("A11:A10010").TextToColumns Destination:=Range("A11"), DataType:=xlFixedWidth, _
        OtherChar:="E", FieldInfo:=Array(Array(0, 1), Array(2, 1), Array(4, 1), Array(6, 1 _
        ), Array(8, 1), Array(13, 1), Array(14, 1), Array(17, 1), Array(22, 1), Array(23, 1), Array( _
        24, 1), Array(29, 1), Array(34, 2)), TrailingMinusNumbers:=True

    Dim MyCell As Range
    Application.Calculation = xlCalculationManual
        ' Convert latitude coordinates: apply negative multiplier for Southern hemisphere
        Range("D11:E10010").Select
        For Each MyCell In Selection.Cells
    If Range("F" & MyCell.Row) = "S" Then
        MyCell.Value = MyCell.Value * (-1)
    End If
    Next
        ' Convert longitude coordinates: apply negative multiplier for Western hemisphere
        Range("G11:H10010").Select
        For Each MyCell In Selection.Cells
    If Range("I" & MyCell.Row) = "W" Then
        MyCell.Value = MyCell.Value * (-1)
    End If
    Next
Application.Calculation = xlCalculationAutomatic

    If Range("A1") = "ENL" Then
    Range("O1").Value = Sheets("IMP").Range("Q1").Value
    Range("P1").Value = Sheets("IMP").Range("Q2").Value
    Range("O2:P2").FormulaR1C1 = "=R[-1]C-35"
    Range("O2:P2").Value = Range("O2:P2").Value
    Range("O11:O10010").FormulaR1C1 = "=MID(RC[-2],R2C15,3)"
    If Range("P2") > 0 Then
    Range("P11:P10010").FormulaR1C1 = "=MID(RC[-2],R2C16,3)"
    End If
    End If
    Range("O11:P10010").Value = Range("O11:P10010").Value
    Columns("O:P").Cut Destination:=Columns("M:N")
    Range("M1:N2").Select
    Selection.Clear
    Range("C4").FormulaR1C1 = "=MAX(C[14])"
    Range("K6").FormulaR1C1 = "=SUM(R[5]C:R[10004]C)"
    Sheets("PRS").Range("M2").Value = Sheets("Sheet1").Range("K6").Value
    Range("L5").FormulaR1C1 = _
        "=IF(OR(IMP!R10C5=""SeeYou Mobile"",IMP!R[-4]C[-9]=""X"",R[1]C[-1]=R[1]C,R[1]C[-1]=0),""PR"","""")"
    Range("L6").FormulaR1C1 = "=SUM(R[5]C:R[10004]C)"
    Sheets("PRS").Range("M3").Value = Sheets("Sheet1").Range("L6").Value
    Sheets("PRS").Range("A10").Value = Sheets("Sheet1").Range("L5").Value
    Range("P3").FormulaR1C1 = "=PRS!R2C1"
    Range("P4").FormulaR1C1 = "=MAX(R[7]C[-15]:R[10006]C[-15])"
    Sheets("Sheet1").Range("P5").Value = Sheets("Sheet1").Range("A11").Value
    Range("P11:P10010").FormulaR1C1 = _
        "=IF(RC[-14]="""","""",IF(AND(RC[-15]>=R5C16,RC[-15]<=R4C16),TIME(RC[-15],RC[-14],RC[-13])+R3C16,TIME(RC[-15],RC[-14],RC[-13])+R3C16+1))"
    Range("O12:O10010").FormulaR1C1 = "=IF(RC[-13]="""","""",RC[1]-R[-1]C[1])"
    Range("O6").FormulaR1C1 = "=Average(R[5]C:R[10004]C)"
    'Sets 00:00:04 as interval for 5-fix vs 14-fix calcs @ Z & AL
    Range("O7").Value = "0.0000462962962962963"
    Range("A2").FormulaR1C1 = "=SECOND(R6C15)"
    Range("Q11:Q10010").FormulaR1C1 = "=IF(RC[-15]="""","""",R[-1]C+1)"
    Range("R11:R10010").FormulaR1C1 = _
        "=IF(OR(R[1]C[-17]="""",AND(RC[-14]=R[1]C[-14],RC[-13]=R[1]C[-13],RC[-11]=R[1]C[-11],RC[-10]=R[1]C[-10])),"""",1)"
   If Range("L5") = "" Then
    Range("S11:S10010").FormulaR1C1 = "=IF(AND(RC[-1]=1,RC[-8]>R[-1]C[-8]),1,"""")"
    Range("T11:T10010").FormulaR1C1 = "=IF(AND(RC[-4]<R4C[2],RC[-2]=""""),RC[-9],"""")"
    Range("T2").FormulaR1C1 = "=IF(R[2]C=0,R11C11,IF(R2C22=R4C22,R4C20,R7C24))"
    Range("T4").FormulaR1C1 = _
        "=IF(SUM(R[7]C:R[10011]C)=0,0,AVERAGE(R[7]C:R[10011]C))"
    Range("U11:U10010").FormulaR1C1 = _
        "=IF(R[-1]C[-20]="""","""",IF(AND(SUM(RC[-3]:R[7]C[-3])=8,RC[-10]>R[-1]C[-10]+3,RC[-10]<R[-1]C[-10]+6),3,IF(AND(SUM(RC[-3]:R[7]C[-3])=8,RC[-10]>R[-1]C[-10]+6),6,0)))"
   ElseIf Range("L5") <> "" Then
    Range("S11:S10010").FormulaR1C1 = "=IF(AND(RC[-1]=1,RC[-7]>R[-1]C[-7]),1,"""")"
    Range("T11:T10010").FormulaR1C1 = "=IF(AND(RC[-4]<R4C[2],RC[-2]=""""),RC[-8],"""")"
    Range("T2").FormulaR1C1 = "=IF(R[2]C=0,0,IF(R2C22=R4C22,R4C20,R7C24))"
    Range("T4").FormulaR1C1 = _
        "=IF(SUM(R[7]C:R[10011]C)=0,0,AVERAGE(R[7]C:R[10011]C))"
    Range("U11:U10010").FormulaR1C1 = _
        "=IF(R[-1]C[-20]="""","""",IF(AND(SUM(RC[-3]:R[7]C[-3])=8,RC[-9]>R[-1]C[-9]+3,RC[-9]<R[-1]C[-9]+6),3,IF(AND(SUM(RC[-3]:R[7]C[-3])=8,RC[-9]>R[-1]C[-9]+6),6,0)))"
   End If
    'Range("V11:V10010").FormulaR1C1 = "=IF(OR(AND(RC[-4]=1,RC[-3]=1,R[1]C[-3]=1,RC[-1]=6),AND(R[1]C[-1]=3,RC[-1]=3,R[-1]C[-1]="""")),RC[-6],"""")"
    Range("V11:V10010").FormulaR1C1 = "=IF(OR(AND(RC[-4]=1,RC[-3]=1,R[1]C[-3]=1,RC[-1]=6),AND(R[1]C[-1]=3,RC[-1]=3,R[-1]C[-1]=0)),RC[-6],"""")"
    Range("V2").FormulaR1C1 = _
        "=IF(AND(R[2]C[-2]=0,R[5]C[1]=0),R[9]C[-6],IF(R[5]C[1]>R[2]C,R[5]C[1],R[2]C))"
    Range("V4").FormulaR1C1 = "=MIN(R[7]C:R[10011]C)"
   If Range("L5") = "" Then
    Range("W12:W10010").FormulaR1C1 = _
        "=IF(OR(RC[-22]="""",RC[-12]<R4C[-3]+5,RC[-12]>R4C[-3]+20,RC[-6]>0.75*R4C3),"""",IF(AND(AVERAGE(R[-10]C[-12]:R[-1]C[-12])<R4C[-3]+20,RC[-7]>R4C[-1]),RC[-7],""""))"
    Range("X11:X10010").FormulaR1C1 = "=IF(OR(RC[-1]="""",RC[-8]=R7C[-1]=FALSE),"""",AVERAGE(R[-10]C[-13]:R[-1]C[-13]))"
   ElseIf Range("L5") <> "" Then
    Range("W12:W10010").FormulaR1C1 = _
        "=IF(OR(RC[-22]="""",RC[-11]<R4C[-3]+5,RC[-11]>R4C[-3]+20,RC[-6]>0.75*R4C3),"""",IF(AND(AVERAGE(R[-10]C[-11]:R[-1]C[-11])<R4C[-3]+20,RC[-7]>R4C[-1]),RC[-7],""""))"
    Range("X11:X10010").FormulaR1C1 = "=IF(OR(RC[-1]="""",RC[-8]=R7C[-1]=FALSE),"""",AVERAGE(R[-10]C[-12]:R[-1]C[-12]))"
   End If
    Range("W7").FormulaR1C1 = "=MAX(R[4]C:R[10011]C)"
    Range("X7").FormulaR1C1 = "=MAX(R[4]C:R[10011]C)"
    Range("Y11:Y10010").FormulaR1C1 = "=IF(RC[-24]="""","""",IF(RC[-12]>R6C25,1,""""))"
    Range("Y3").FormulaR1C1 = "='[A.xlsm]ALL CLAIMS'!R8C6"
    Range("Y6").FormulaR1C1 = "=PRS!R6C1"
    'Range("Z25:Z10009").FormulaR1C1 = "=IF(AND(RC[-1]=1,RC[-9]<R4C3/2,OR(AND(R6C15>=0.00013,SUM(R[-4]C[-1]:RC[-1])>=5),AND(R6C15<0.00013,SUM(R[-14]C[-1]:RC[-1])>=14))),R[1]C[-10],"""")"
    ''Range("Z25:Z10010").FormulaR1C1 = "=IF(AND(R[-1]C[-1]=1,RC[-9]<R4C3/2,R[1]C[-1]<>1,OR(AND(R6C15<=R7C15,SUM(R[-5]C[-1]:R[-1]C[-1])>=3),AND(R6C15>R7C15,SUM(R[-15]C[-1]:R[-1]C[-1])>=8))),RC[-10],"""")"
    '''Range("Z25:Z10009").FormulaR1C1 = _
       "=IF(AND(R[-1]C[-1]=1,RC[-10]>R5C28,R[1]C[-1]<>1,OR(AND(R6C15<=R7C15,SUM(R[-5]C[-1]:R[-1]C[-1])>=3),AND(R6C15>R7C15,SUM(R[-15]C[-1]:R[-1]C[-1])>=8))),RC[-10],"""")"
    ''''Range("Z25:Z10009").FormulaR1C1 = "=IF(AND(R[-1]C[-1]=1,RC[-10]>R5C28,R[1]C[-1]<>1,SUM(R[-15]C[-1]:R[-1]C[-1])>=8),RC[-10],"""")"
    '''''Range("AA25:AA10009").FormulaR1C1 = "=IF(AND(R[-1]C[-2]="""",RC[-2]=1,RC[-11]>R3C26,SUM(RC[-2]:R[15]C[-2])>8),RC[-11],"""")"
    Range("AA25:AA10009").FormulaR1C1 = "=IF(AND(R[-1]C[-2]="""",RC[-2]=1,SUM(RC[-2]:R[15]C[-2])>8),RC[-11],"""")"
    Range("AA3").FormulaR1C1 = "=MIN(R[22]C:R[10006]C)"
    Range("AA4").FormulaR1C1 = "=MAX(R[21]C:R[10005]C)"
    Range("AA5").FormulaR1C1 = "=IF(R[-1]C<>0,(R[-1]C-R[-2]C),"""")"
    Range("AA3:AA5").Value = Range("AA3:AA5").Value
    Range("Z25:Z10009").FormulaR1C1 = "=IF(AND(R[-1]C[-1]=1,R[1]C[-1]<>1,SUM(R[-15]C[-1]:R[-1]C[-1])>=8),RC[-10],"""")"
    '''''Range("Z25:Z10009").FormulaR1C1 = "=IF(AND(R[-1]C[-1]=1,RC[-10]>R5C28,RC[-10]<R4C27,R[1]C[-1]<>1,SUM(R[-15]C[-1]:R[-1]C[-1])>=8),RC[-10],"""")"
    ''''''Range("Z25:Z10009").FormulaR1C1 = "=IF(AND(OR(R4C27=0,R4C27<RC[-10]),R[-1]C[-1]=1,R[1]C[-1]<>1,SUM(R[-15]C[-1]:R[-1]C[-1])>=8),RC[-10],"""")"
    ''''''Range("Z25:Z10009").FormulaR1C1 = "=IF(AND(R[-1]C[-1]=1,R[1]C[-1]<>1,SUM(R[-15]C[-1]:R[-1]C[-1])>=8),RC[-10],"""")"
    'Range("Z25").FormulaR1C1 = "=IF(AND(R4C27=0,R[-1]C[-1]=1,R[1]C[-1]<>1,SUM(R[-15]C[-1]:R[-1]C[-1])>=8),RC[-10],IF(AND(R4C27>0,RC[-10]>=R3C27,RC[-10]<R4C27,R[-1]C[-1]=1,R[1]C[-1]<>1,SUM(R[-15]C[-1]:R[-1]C[-1])>=8),RC[-10],""""))"
    Range("Z2").FormulaR1C1 = "=IF(R1C1=""NO ENL"",0,IF(R[6]C=0,R[2]C,R[6]C))"
    Range("Z3").FormulaR1C1 = "=MIN(R[22]C:R[10006]C)"
    Range("Z4").FormulaR1C1 = "=MAX(R[21]C:R[10005]C)"

    If Range("AA4") <> 0 And Range("AA4") < Range("Z4") And Range("AA5") >= 0.001388889 Then
         Range("Z25:Z10009").FormulaR1C1 = "=IF(AND(RC[-10]<R4C27,RC[-10]>=R3C27,R[-1]C[-1]=1,R[1]C[-1]<>1,SUM(R[-15]C[-1]:R[-1]C[-1])>=8),RC[-10],"""")"
    End If
    'Range("Z2").FormulaR1C1 = "=IF(OR(R[1]C[-1]=2,R[1]C[-1]=4,R1C1=""NO ENL""),0,MAX(R[9]C:R[10008]C))"
    ''Range("Z2").FormulaR1C1 = "=IF(OR(R[1]C[-1]=2,R[1]C[-1]=4,R1C1=""NO ENL""),0,IF(R[6]C=0,MAX(R[9]C:R[10008]C),R[6]C))"
    '''Range("Z2").FormulaR1C1 = "=IF(R1C1=""NO ENL"",0,IF(R[6]C=0,MIN(R[9]C:R[10008]C),R[6]C))"
    ''''Range("Z2").FormulaR1C1 = "=IF(R1C1=""NO ENL"",0,IF(R[6]C=0,R[1]C,R[6]C))"
    '''''Range("Z2").FormulaR1C1 = "=IF(R1C1=""NO ENL"",0,IF(R[6]C=0,R[2]C,R[6]C))"
    '''Range("Z3").FormulaR1C1 = "=MIN(R[22]C:R[10006]C)"
    '''''Range("Z3").FormulaR1C1 = "=IF(R[1]C[12]<>""NONE"",MIN(R[22]C:R[10006]C),R[1]C)"
    '''''Range("Z4").FormulaR1C1 = "=MAX(R[21]C:R[10005]C)"
    ''Range("AA3").FormulaR1C1 = "=MIN(R[22]C:R[10006]C)"
    '''Range("AA3").FormulaR1C1 = "=IF(OR(R3C25=2,R3C25=4),0,MIN(R[22]C:R[10006]C))"
    ''''Range("AA3").FormulaR1C1 = "=IF(OR(R3C25=2,R3C25=4,MIN(R[22]C:R[10006]C)-RC[-1]<=0.000694),0,MIN(R[22]C:R[10006]C))"
    ''''Range("AA4").FormulaR1C1 = "=MAX(R[21]C:R[10005]C)"
    ''''Range("AA3:AA4").Value = Range("AA3:AA4").Value

    Range("AB2").FormulaR1C1 = "=IF(R[2]C=R3C16+PRS!R2C2/24,R[1]C,R[2]C)"
    Range("AB3").FormulaR1C1 = "=IF(R[-1]C[-8]=0,R[-1]C[-6],MAX(RC[-2],R[2]C))"

   If Range("L5") = "" Then
    Range("AB11:AB10010").FormulaR1C1 = _
        "=IF(R[3]C[-12]="""","""",IF(AND(RC[-12]>R2C22,OR(AND(RC[-9]=1,R[1]C[-9]="""",RC[-7]>=3,SUM(R[1]C[-7]:R[3]C[-7])=0),(AND(RC[-9]=1,SUM(R[1]C[-9]:R[4]C[-9])=0)),(AND(RC[-7]>=3,RC[-17]-R[3]C[-17]>=30)))),RC[-12],""""))"
    Range("AC11:AC10010").FormulaR1C1 = "=IF(RC[-13]=R3C[-1],RC[-18],"""")"
   ElseIf Range("L5") <> "" Then
    Range("AB11:AB10010").FormulaR1C1 = _
        "=IF(OR(RC[-16]=0,R[3]C[-12]=""""),"""",IF(AND(RC[-12]>R2C22,OR(AND(RC[-9]=1,R[1]C[-9]="""",RC[-7]>=3,SUM(R[1]C[-7]:R[3]C[-7])=0),(AND(RC[-9]=1,SUM(R[1]C[-9]:R[4]C[-9])=0)),(AND(R[3]C[-16]<>0,RC[-7]>=3,RC[-16]-R[3]C[-16]>=30)))),RC[-12],""""))"
    Range("AC11:AC10010").FormulaR1C1 = "=IF(RC[-13]=R3C[-1],RC[-17],"""")"
   End If
    Range("O11:AB10010").Value = Range("O11:AB10010").Value
    Range("AB2").FormulaR1C1 = "=IF(R[2]C=R3C16+PRS!R2C2/24,R[1]C,R[2]C)"
    Range("AB3").FormulaR1C1 = "=IF(R[-1]C[-8]=0,R[-1]C[-6],MAX(R[-1]C[-2],R[2]C))"
    Range("AB4").FormulaR1C1 = "='[A.xlsm]DATA ENTRY CHECK'!R12C7+R3C16-PRS!R2C2/24"
    Range("AB5").FormulaR1C1 = "=MIN(R[6]C:R[9995]C)"
    Range("AC6").FormulaR1C1 = "=MAX(R[5]C:R[9994]C)"
    Range("AD11:AD10010").FormulaR1C1 = "=IF(OR(RC[-14]=R2C28,RC[2]=1),RC[-26],"""")"
    'Range("AD2").FormulaR1C1 = "=MAX(R[8]C:R[10009]C)"
    Range("AD2").FormulaR1C1 = "=IF(R[6]C=0,MAX(R[8]C:R[10009]C),R[6]C)"
    Range("AE11:AE10010").FormulaR1C1 = "=IF(RC[-1]="""","""",RC[-26])"
    'Range("AE2").FormulaR1C1 = "=MAX(R[8]C:R[10009]C)*0.001"
    Range("AE2").FormulaR1C1 = "=IF(R[6]C=0,MAX(R[8]C:R[10009]C)*0.001,R[6]C)"
    Range("AF11:AF10010").FormulaR1C1 = "=IF(AND(PRS!R14C4>0,R4C28>R3C16,R4C[-4]>R[-1]C[-16],R4C[-4]<RC[-16]),1,"""")"
    Range("AG11:AG10010").FormulaR1C1 = "=IF(RC[-2]="""","""",RC[-26])"
    'Range("AG2").FormulaR1C1 = "=MAX(R[8]C:R[10009]C)"
    Range("AG2").FormulaR1C1 = "=IF(R[6]C=0,MAX(R[8]C:R[10009]C),R[6]C)"
    Range("AH11:AH10010").FormulaR1C1 = "=IF(RC[-1]="""","""",RC[-26])"
    'Range("AH2").FormulaR1C1 = "=MAX(R[8]C:R[10009]C)*0.001"
    Range("AH2").FormulaR1C1 = "=IF(R[6]C=0,MAX(R[8]C:R[10009]C)*0.001,R[6]C)"
   If Range("L5") = "" Then
    Range("AJ11:AJ10010").FormulaR1C1 = "=IF(RC[-2]="""","""",RC[-25])"
   ElseIf Range("L5") <> "" Then
    Range("AJ11:AJ10010").FormulaR1C1 = "=IF(RC[-2]="""","""",RC[-24])"
   End If
    'Range("AJ2").FormulaR1C1 = "=IF(R4C[-8]=0,R6C29,MAX(R[9]C:R[10008]C))"
    Range("AJ2").FormulaR1C1 = "=IF(R4C[-8]=0,R6C29,IF(R[6]C=0,MAX(R[9]C:R[10008]C),R[6]C))"
    'Range("AL11:AL10010").FormulaR1C1 = "=IF(AND(RC[-22]>R2C28,OR(AND(R6C15>=0.00013,SUM(RC[-13]:R[4]C[-13])=5),AND(R6C15<0.00013,SUM(RC[-13]:R[14]C[-13])=15))),RC[-22],"""")"
    ''Range("AL11:AL10010").FormulaR1C1 = "=IF(AND(RC[-22]>R2C28,RC[-13]<>1,R[1]C[-13]=1,OR(AND(R6C15<=R7C15,SUM(R[2]C[-13]:R[6]C[-13])>=3,RC[-13]=0),AND(R6C15>R7C15,SUM(R[2]C[-13]:R[15]C[-13])>=8,R[1]C[-13]=1))),RC[-22],"""")"
    Range("AL11:AL10010").FormulaR1C1 = "=IF(RC[-11]=R3C27,RC[-22],IF(AND(RC[-22]>R2C28,RC[-13]<>1,R[1]C[-13]=1,OR(AND(R6C15<=R7C15,SUM(R[2]C[-13]:R[6]C[-13])>=3,RC[-13]=0),AND(R6C15>R7C15,SUM(R[2]C[-13]:R[15]C[-13])>=8,R[1]C[-13]=1))),RC[-22],""""))"
    Range("AL2").FormulaR1C1 = "=R4C38"
    '''Range("AL4").FormulaR1C1 = "=IF(OR(R[-1]C[-13]=2,R[3]C[1]=0,MIN(R[7]C:R[10008]C)=0),""NONE"",MIN(R[7]C:R[10008]C))"
    '''Range("AL4").FormulaR1C1 = "=IF(OR(R[-1]C[-13]=2,R[3]C[1]=0,MIN(R[7]C:R[10008]C)<RC[-12],MIN(R[7]C:R[10008]C)=0),""NONE"",MIN(R[7]C:R[10008]C))"
    '''Range("AL4").FormulaR1C1 = "=IF(OR(R[-1]C[-13]=2,R[3]C[1]=0,MIN(R[7]C:R[10008]C)=0),""NONE"",MIN(R[7]C:R[10008]C))"
    ''''Range("AL4").FormulaR1C1 = "=IF(OR(R[-1]C[-13]=2,R[-1]C[-13]=4,R[3]C[1]=0,R[-1]C[-11]=0),""NONE"",MIN(R[7]C:R[10008]C))"
    '''''Range("AL4").FormulaR1C1 = "=IF(OR(R[-1]C[-13]=2,R[-1]C[-13]=4,R[3]C[1]=0,R[-1]C[-11]=0,RC[-11]<RC[-12]),""NONE"",MIN(R[7]C:R[10008]C))"
    Range("AL4").FormulaR1C1 = "=IF(OR(R[-1]C[-13]=2,R[-1]C[-13]=4,R[3]C[1]=0,R[-1]C[-11]=0,RC[-11]<RC[-12]),""NONE"",RC[-11])"
    Range("AL5").FormulaR1C1 = "=IF(R[-1]C<>""NONE"",MAX(R[20]C:R[10004]C),0)"
    Range("AL7").FormulaR1C1 = "=IF(R[-3]C=""NONE"",""NONE"",R2C47-R[-3]C)"
   If Range("L5") = "" Then
    'Range("AM11:AM10010").FormulaR1C1 = "=IF(OR(RC[-1]="""",RC[-20]="""",R[-1]C[-28]>RC[-28]),0,1)"
    Range("AM11:AM10010").FormulaR1C1 = "=IF(AND(RC[-1]="""",RC[-20]="""",R[-1]C[-26]<R6C25),0,1)"
    '''Range("AO11:AO10010").FormulaR1C1 = "=IF(R4C38=""NONE"","""",IF(RC[-3]=R4C40,RC[-30],""""))"
    Range("AO11:AO10009").FormulaR1C1 = "=IF(R4C38=""NONE"","""",IF(RC[-25]=R4C40,RC[-30],""""))"
   ElseIf Range("L5") <> "" Then
    'Range("AM11:AM10010").FormulaR1C1 = "=IF(OR(RC[-1]="""",RC[-20]="""",R[-1]C[-27]>RC[-27]),0,1)"
    Range("AM11:AM10010").FormulaR1C1 = "=IF(AND(RC[-1]="""",RC[-20]="""",R[-1]C[-26]<R6C25),0,1)"
    ''''Range("AO11:AO10010").FormulaR1C1 = _"=IF(R4C38=""NONE"","""",IF(RC[-3]=R4C40,RC[-29],""""))"
    Range("AO11:AO10009").FormulaR1C1 = "=IF(R4C38=""NONE"","""",IF(RC[-25]=R4C40,RC[-29],""""))"
   End If
    Range("AM7").FormulaR1C1 = "=SUM(R[4]C:R[10005]C)"
    Range("AN11:AN10010").FormulaR1C1 = _
        "=IF(OR(RC[-39]="""",RC[-23]<R4C3-0.4*R4C3),"""",IF(AND(RC[-24]>R2C28,RC[-22]="""",R[1]C[-22]="""",RC[-21]=""""),RC[-24],""""))"
    ''Range("AN11:AN10010").FormulaR1C1 = "=IF(OR(RC[-39]="""",RC[-23]<R4C3-0.4*R4C3),"""",IF(AND(RC[-24]>R2C28,RC[-22]="""",RC[-21]=""""),RC[-24],""""))"
    'Range("AN2").FormulaR1C1 = "=R[2]C"
    ''''Range("AN2").FormulaR1C1 = "=IF(R[6]C=0,R[2]C,R[6]C)"
    Range("AN2").FormulaR1C1 = "=IF(R[6]C=0,R[2]C,R[6]C)"
    '''Range("AN4").FormulaR1C1 = "=IF(AND(OR(R[-1]C[-15]=3,R[-1]C[-15]=5),R[3]C[-1]>0),MIN(R[7]C[-2]:R[10008]C[-2]),0)"
    ''''Range("AN4").FormulaR1C1 = "=IF(AND(OR(R[-1]C[-15]=3,R[-1]C[-15]=5),RC[-2]<>""NONE"",R[3]C[-1]>0),MIN(R[7]C[-2]:R[10008]C[-2]),0)"
    Range("AN4").FormulaR1C1 = "=IF(AND(OR(R[-1]C[-15]=3,R[-1]C[-15]=5),RC[-2]<>""NONE"",R[3]C[-1]>0),RC[-2],0)"
    'Range("AO2").FormulaR1C1 = "=IF(R2C40="""","""",MAX(R[9]C:R[10010]C))"
    Range("AO2").FormulaR1C1 = "=IF(R2C40="""","""",IF(R[6]C=0,MAX(R[9]C:R[10010]C),R[6]C))"
    Range("AP11:AP10010").FormulaR1C1 = "=IF(R4C40="""","""",IF(RC[-26]=R4C40,RC[-38],""""))"
    'Range("AP2").FormulaR1C1 = "=MAX(R[8]C:R[10010]C)"
    Range("AP2").FormulaR1C1 = "=IF(R[6]C=0,MAX(R[8]C:R[10010]C),R[6]C)"
    Range("AQ11:AQ10010").FormulaR1C1 = "=IF(RC[-1]="""","""",RC[-38])"
    'Range("AQ2").FormulaR1C1 = "=MAX(R[8]C:R[10010]C)*0.001"
    Range("AQ2").FormulaR1C1 = "=IF(R[6]C=0,MAX(R[8]C:R[10010]C)*0.001,R[6]C)"
    Range("AR11:AR10010").FormulaR1C1 = "=IF(RC[-1]="""","""",RC[-37])"
    'Range("AR2").FormulaR1C1 = "=MAX(R[8]C:R[10010]C)"
    Range("AR2").FormulaR1C1 = "=IF(R[6]C=0,MAX(R[8]C:R[10010]C),R[6]C)"
    Range("AS11:AS10010").FormulaR1C1 = "=IF(RC[-1]="""","""",RC[-37])"
    'Range("AS2").FormulaR1C1 = "=MAX(R[8]C:R[10010]C)*0.001"
    Range("AS2").FormulaR1C1 = "=IF(R[6]C=0,MAX(R[8]C:R[10010]C)*0.001,R[6]C)"
   If Range("L5") = "" Then
    Range("AT11:AT10010").FormulaR1C1 = _
        "=IF(OR(RC[-6]="""",RC[-35]>PRS!R7C4+152),"""",RC[-35])"
    Range("AU11:AU10010").FormulaR1C1 = _
        "=IF(RC[-30]=R4C[-44],RC[-31],IF(RC[-1]="""","""",IF(AND(RC[-29]="""",RC[-28]="""",OR(RC[-36]=R9C[-1],AND(RC[-1]=R[1]C[-1],R[1]C[-1]=R[2]C[-1]),AND(RC[-36]<=R9C[-1]+5,RC[-36]>=R9C[-1]-5))),RC[-7],"""")))"
   ElseIf Range("L5") <> "" Then
    Range("AT11:AT10010").FormulaR1C1 = _
        "=IF(OR(RC[-6]="""",RC[-34]>PRS!R7C4+152,RC[-34]=0),"""",RC[-34])"
    Range("AU11:AU10010").FormulaR1C1 = _
        "=IF(RC[-30]=R4C[-44],RC[-31],IF(RC[-1]="""","""",IF(AND(RC[-29]="""",RC[-28]="""",OR(RC[-35]=R9C[-1],AND(RC[-1]=R[1]C[-1],R[1]C[-1]=R[2]C[-1]),AND(RC[-35]<=R9C[-1]+5,RC[-35]>=R9C[-1]-5))),RC[-7],"""")))"
   End If
    Range("AT9").FormulaR1C1 = "=AVERAGE(R[2]C:R[10003]C)"
    Range("AU2").FormulaR1C1 = "=R[2]C"
    Range("AU4").FormulaR1C1 = "=MIN(R[7]C:R[10008]C)"
    Range("AU5").FormulaR1C1 = "=IF(R[-1]C=MAX(R[6]C[-31]:R[9995]C[-31]),""Last DP"",""NO"")"
    Range("AV11:AV10010").FormulaR1C1 = "=IF(RC[-32]=R2C47,RC[-44],"""")"
    Range("AV2").FormulaR1C1 = "=MAX(R[8]C:R[10010]C)"
    Range("AW11:AW10010").FormulaR1C1 = "=IF(RC[-1]="""","""",RC[-44])"
    Range("AW2").FormulaR1C1 = "=MAX(R[8]C:R[10010]C)*0.001"
    Range("AY11:AY10010").FormulaR1C1 = "=IF(RC[-2]="""","""",RC[-44])"
    Range("AY2").FormulaR1C1 = "=MAX(R[8]C:R[10010]C)"
    Range("AZ11:AZ10010").FormulaR1C1 = "=IF(RC[-1]="""","""",RC[-44])"
    Range("AZ2").FormulaR1C1 = "=MAX(R[8]C:R[10010]C)*0.001"
   If Range("L5") = "" Then
    Range("BB11:BB10010").FormulaR1C1 = _
        "=IF(AND(RC[-7]=R2C47,R5C47=""Last DP""),RC[-43],IF(OR(RC[-14]="""",RC[-38]<R2C47,RC[-36]=1),"""",RC[-43]))"
   ElseIf Range("L5") <> "" Then
     Range("BB11:BB10010").FormulaR1C1 = _
        "=IF(AND(RC[-7]=R2C47,R5C47=""Last DP""),RC[-42],IF(OR(RC[-14]="""",RC[-38]<R2C47,RC[-36]=1,RC[-42]=0),"""",RC[-42]))"
   End If
    Range("BB2").FormulaR1C1 = "=R[2]C"
    Range("BB4").FormulaR1C1 = "=AVERAGE(R[7]C:R[10008]C)"
    Range("CF11:CF2510").FormulaR1C1 = "=IF(AND(RC[-64]=""""=FALSE,RC[-68]<R2C[-62]),R[2]C[-68],"""")"
    Range("CF4").FormulaR1C1 = "=MAX(R[7]C:R[10006]C)"
    Range("CF2").FormulaR1C1 = "=IF(R[2]C=0,R2C22,R[2]C)"
    Range("CG11:CG2510").FormulaR1C1 = "=IF(RC[-69]=R2C[-1],RC[-81],"""")"
    Range("CH11:CH2510").FormulaR1C1 = "=IF(RC[-1]="""","""",RC[-81]/1000)"
    Range("CJ11:CJ2510").FormulaR1C1 = "=IF(RC[-2]="""","""",RC[-81])"
    Range("CK11:CK2510").FormulaR1C1 = "=IF(RC[-1]="""","""",RC[-81]/1000)"
    Range("CG2").FormulaR1C1 = "=MAX(R[9]C:R[10008]C)"
    Range("CH2").FormulaR1C1 = "=MAX(R[9]C:R[10008]C)"
    Range("CJ2").FormulaR1C1 = "=MAX(R[9]C:R[10008]C)"
    Range("CK2").FormulaR1C1 = "=MAX(R[9]C:R[10008]C)"
    Range("A1").Select

    Sheets("Sheet1").Range("C3").Value = Sheets("PRS").Range("B10").Value
    If Range("C3") > 10000 Then
        Application.Run "Ab.xlsm!RefineLDG"
    End If

    If Range("Y3") = 3 And Range("C4") > 10000 Then
        Application.Run "Ab.xlsm!ENLrefine"
    End If
    Application.Run "Ab.xlsm!NewBRecords"
End Sub
Sub RefineLDG()
    ' RefineLDG improves landing point accuracy for flights with high fix counts (>10,000)
    ' This function re-processes the last portion of flight data to identify a more precise landing location
    ' It sorts data by time, filters for landing candidates, and updates coordinate information
    ' Used when initial processing may have missed the exact landing point due to data volume
    '
    ' Revise Landing if over 10K fixes  7/17/18

    ' Activate BR sheet containing processed flight data
    Sheets("BR").Activate
    ' Copy time and record data for landing analysis
    Range("AT1:AU60000").Value = Range("J1:K60000").Value
    ' Calculate cutoff time (max time minus 1 hour) to focus on landing phase
    Range("AV1").FormulaR1C1 = "=MAX(RC[-2]:R[60000]C[-2])-1/24"
    ' Filter for records in the last hour of flight (potential landing candidates)
    Range("AV2:AV60001").FormulaR1C1 = "=IF(RC[-2]>R1C,RC[-2],"""")"
    ' Extract corresponding record data for filtered times
    Range("AW2:AW60001").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-2],"""")"

    Range("AV2:AW60001").Value = Range("AV2:AW60001").Value
    ' Sort landing candidates chronologically for processing
    ActiveWorkbook.Worksheets("BR").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("BR").Sort.SortFields.Add Key:=Range("AV2:AV60001") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("BR").Sort
        .SetRange Range("AV2:AW60001")
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    ' Parse filtered B-record data into coordinate components for landing analysis
    ' Extracts: time, lat degrees, lat minutes, lat direction, lon degrees, lon minutes, lon direction, altitudes
    Range("AW2:AW3601").Select
    Selection.TextToColumns Destination:=Range("AW2"), DataType:=xlFixedWidth, _
        OtherChar:="E", FieldInfo:=Array(Array(0, 9), Array(6, 1), Array(8, 1), Array(13, _
        1), Array(14, 1), Array(17, 1), Array(22, 1), Array(23, 9), Array(24, 1), Array(29, 9), _
        Array(33, 9)), TrailingMinusNumbers:=True

    Range("BD1").FormulaR1C1 = "=MIN(R[1]C:R[3600]C)"
    Range("BD2:BI2").FormulaR1C1 = "=MIN(R[1]C:R[3598]C)"

    Range("BD3").FormulaR1C1 = "=IF(AND(RC[-7]=R[1]C[-7],RC[-6]=R[1]C[-6],RC[-4]=R[1]C[-4],RC[-3]=R[1]C[-3],ABS(RC[-1]-PRS!R8C4)<4),RC[-8],"""")"
    Range("BE3").FormulaR1C1 = "=IF(RC[-1]<>R1C56,"""",IF(RC[-6]=""S"",-1*RC[-8],RC[-8]))"
    Range("BF3").FormulaR1C1 = "=IF(RC[-1]="""","""",IF(RC[-7]=""S"",-1*(RC[-8]/1000),RC[-8]/1000))"
    Range("BG3").FormulaR1C1 = "=IF(RC[-1]="""","""",IF(RC[-5]=""W"",-1*RC[-7],RC[-7]))"
    Range("BH3").FormulaR1C1 = "=IF(RC[-1]="""","""",IF(RC[-6]=""W"",-1*(RC[-7]/1000),RC[-7]/1000))"
    Range("BI3").FormulaR1C1 = "=IF(RC[-1]="""","""",RC[-6])"
    'Copy Ref AV
    Application.Calculation = xlCalculationManual
 With Worksheets("BR")
    LastRow = .Cells(Rows.Count, "AV").End(xlUp).Row
.Range("BD3:BI3").AutoFill Destination:=.Range("BD3:BI" & LastRow), Type:=xlFillDefault
    End With
    Application.Calculation = xlCalculationAutomatic

    ' Switch to main data sheet to update landing coordinates
    Sheets("Sheet1").Activate
    ' Update landing data only if refined landing time differs from current value
    If Sheets("Sheet1").Range("AU4") <> Sheets("BR").Range("BD2") Then
        ' Update refined landing time
        Sheets("Sheet1").Range("AU4").Value = Sheets("BR").Range("BD2").Value
        ' Update refined landing coordinates (latitude degrees and minutes)
        Sheets("Sheet1").Range("AV2").Value = Sheets("BR").Range("BE2").Value
        Sheets("Sheet1").Range("AW2").Value = Sheets("BR").Range("BF2").Value
        ' Update refined landing coordinates (longitude degrees and minutes)
        Sheets("Sheet1").Range("AY2").Value = Sheets("BR").Range("BG2").Value
        Sheets("Sheet1").Range("AZ2").Value = Sheets("BR").Range("BH2").Value
        ' Update refined landing altitude
        Sheets("Sheet1").Range("BB2").Value = Sheets("BR").Range("BI2").Value
    End If

End Sub

Sub ENLrefine()
'
' ENLrefine Macro
' Improved accuracy for ENL on flights >10000 fixes
'
    Application.ScreenUpdating = False
    Sheets("BR").Activate
    Range("G1").Value = 0.00005787037037
    Range("I1:I60000").FormulaR1C1 = "=IF(OR(AND(RC[1]<=Sheet1!R2C26+R1C7,RC[1]>=Sheet1!R2C26-R1C7),AND(RC[1]<=Sheet1!R4C38+R1C7,RC[1]>=Sheet1!R4C38-R1C7)),RC[1],"""")"
    Range("I1:I60000").Value = Range("I1:I60000").Value

    Range("AT1:AT60000").Value = Range("I1:I60000").Value
    Range("I1:I60000").FormulaR1C1 = "=IF(RC[37]<>"""",RC[2],"""")"
    Range("AU1:AU60000").Value = Range("I1:I60000").Value

    Columns("AT:AU").Select
    ActiveWorkbook.Worksheets("BR").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("BR").Sort.SortFields.Add Key:=Columns("AT:AT"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("BR").Sort
        .SetRange Columns("AT:AU")
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("AR1").FormulaR1C1 = "=MIN(RC[12]:R[10]C[12])"
    Range("AR2").FormulaR1C1 = "=MAX(R[10]C[12]:R[20]C[12])"
    Range("AS1").FormulaR1C1 = "=Sheet1!R6C25"
    Range("AS2").FormulaR1C1 = "=IMP!R[-1]C[-28]-1"

    Range("BC1:BC22").FormulaR1C1 = "=MID(RC[-8],R2C45,3)"
    Range("BC1:BC22").Value = Range("BC1:BC22").Value
    Range("BD1:BD22").FormulaR1C1 = "=IF(RC[-1]<R1C45,RC[-10],"""")"
    Range("BE1:BE22").FormulaR1C1 = "=IF(OR(RC[-1]=R1C44,RC[-1]=R2C44),RC[-1],"""")"
    Range("BF1:BF22").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-11],"""")"
    Range("BG1:BG22").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-4],"""")"

    Range("BE1:BG22").Value = Range("BE1:BG22").Value

    ActiveWorkbook.Worksheets("BR").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("BR").Sort.SortFields.Add Key:=Range("BE1"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("BR").Sort
        .SetRange Range("BE1:BG22")
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    Range("AT1:BD22").Select
    Selection.Clear

    Range("AT1:AV2").Value = Range("BE1:BG2").Value
    Range("BC1:BC2").Value = Range("AV1:AV2").Value
    Range("AV1:AV2").Clear
    Range("AU1:AU2").Select
    Selection.TextToColumns Destination:=Range("AU1"), DataType:=xlFixedWidth, _
        OtherChar:="E", FieldInfo:=Array(Array(0, 9), Array(6, 1), Array(8, 1), Array(13, _
        1), Array(14, 1), Array(17, 1), Array(22, 1), Array(23, 9), Array(24, 1), Array(29, 1), _
        Array(34, 9)), TrailingMinusNumbers:=True

    Range("AT3:AT4").FormulaR1C1 = "=R[-2]C"
    Range("AU3:AU4").FormulaR1C1 = "=IF(R[-2]C[2]=""N"",R[-2]C,-1*R[-2]C)"
    Range("AV3:AV4").FormulaR1C1 = "=IF(R[-2]C[1]=""N"",R[-2]C/1000,-1*(R[-2]C/1000))"
    Range("AX3:AX4").FormulaR1C1 = "=IF(R[-2]C[2]=""E"",R[-2]C,-1*R[-2]C)"
    Range("AY3:AY4").FormulaR1C1 = "=IF(R[-2]C[1]=""E"",R[-2]C/1000,-1*(R[-2]C/1000))"
    Range("BA3:BC4").FormulaR1C1 = "=R[-2]C"

    Range("AT3:BC4").Value = Range("AT3:BC4").Value
    Range("AR1:BG2").Clear

    Range("AT1:BC2").Value = Range("AT3:BC4").Value
    Range("AT3:BC4").Clear

    Sheets("Sheet1").Activate
    Range("Z8").FormulaR1C1 = "=BR!R[-7]C[20]"
    Range("AD8").FormulaR1C1 = "=BR!R[-7]C[17]"
    Range("AE8").FormulaR1C1 = "=BR!R[-7]C[17]"
    Range("AG8").FormulaR1C1 = "=BR!R[-7]C[17]"
    Range("AH8").FormulaR1C1 = "=BR!R[-7]C[17]"
    'Range("AJ3").FormulaR1C1 = "=BR!R[-2]C[18]"
    Range("AJ8").FormulaR1C1 = "=BR!R[-7]C[17]"
    Range("AK8").FormulaR1C1 = "=BR!R[-7]C[18]"
    Range("AN8").FormulaR1C1 = "=BR!R[-6]C[6]"
    'Range("AO3").FormulaR1C1 = "=BR!R[-1]C[13]"
    Range("AO8").FormulaR1C1 = "=BR!R[-6]C[12]"
    Range("AP8").FormulaR1C1 = "=BR!R[-6]C[5]"
    Range("AQ8").FormulaR1C1 = "=BR!R[-6]C[5]"
    Range("AR8").FormulaR1C1 = "=BR!R[-6]C[6]"
    Range("AS8").FormulaR1C1 = "=BR!R[-6]C[6]"
    Range("AT8").FormulaR1C1 = "=BR!R[-6]C[9]"
    Range("Z8:AT8").Value = Range("Z8:AT8").Value
    'Range("AJ3:AO3").Value = Range("AJ3:AO3").Value
    Sheets("BR").Activate
    Columns("I:I").Clear
    Range("G1,AT1:BC2").Clear
    Range("A1").Select
End Sub

Sub NewBRecords()
    ' NewBRecords processes imported flight data to extract and parse B-record information
    ' B-records contain GPS position fixes with coordinates, altitude, and time data
    ' This function handles multiple coordinate records and converts them from IGC format

    ' Switch to IMP sheet containing imported flight data
    Sheets("IMP").Select
    ' Extract C-type records (B-records) from imported data
    Range("R1:R1020").FormulaR1C1 = "=IF(RC[-16]=""C"",RC[-17],"""")"
    Range("R1:R1020").Value = Range("R1:R1020").Value
    ' Identify start positions of new B-record sequences
    Range("S2:S1020").FormulaR1C1 = "=IF(AND(R[-1]C[-1]="""",RC[-17]=""C""),1,"""")"
    Range("S2:S1020").Value = Range("S2:S1020").Value
    ' Count total number of B-record sequences
    Range("S1").FormulaR1C1 = "=SUM(R[1]C:R[1020]C)"
    ' Extract actual B-record data for processing
    Range("T1:T1020").FormulaR1C1 = "=IF(RC[-1]=1,RC[-2],"""")"
    Range("T1:T1020").Value = Range("T1:T1020").Value
If Range("S1") > 0 Then
    ' Parse B-record data using fixed-width text-to-columns conversion
    ' IGC B-record format: BHHMMSSDDMMMMMNDDDMMMMMEVPPPPPGGGGGCRLF
    ' B=record type, H=hours, M=minutes, S=seconds, DD=latitude degrees, MMM=latitude minutes (thousandths),
    ' N/S=latitude hemisphere, DDD=longitude degrees, MMM=longitude minutes (thousandths), E/W=longitude hemisphere,
    ' V=validity (A=valid, V=invalid), PPPPP=pressure altitude, GGGGG=GPS altitude
    Columns("T:T").TextToColumns Destination:=Range("T1"), DataType:=xlFixedWidth, _
        OtherChar:="E", FieldInfo:=Array(Array(0, 9), Array(1, 4), Array(7, 1), Array(9, 1 _
        ), Array(11, 1), Array(13, 9), Array(19, 9), Array(23, 1), Array(25, 1)), _
        TrailingMinusNumbers:=True
    ' Convert time components to Excel time format (hours + minutes + seconds)
    Range("Y3:Y1020").FormulaR1C1 = _
        "=IF(RC[-6]="""","""",RC[-5]+TIME(RC[-4],RC[-3],RC[-2]))"
    Range("Y3:Y1020").Value = Range("Y3:Y1020").Value
    ' Find maximum time value in the dataset
    Range("Y2").FormulaR1C1 = "=MAX(R[1]C:R[998]C)"
    Range("Y1").FormulaR1C1 = "=R[1]C"
    Range("Z3:Z1020").FormulaR1C1 = "=IF(OR(RC[-1]=R2C25,RC[-1]=R1021C25),RC[-2],"""")"
    Range("Z3:Z1020").Value = Range("Z3:Z1020").Value
    Range("Z1").FormulaR1C1 = "=MAX(R[2]C:R[1021]C)"
    Range("Z2").FormulaR1C1 = "=MAX(R[1]C:R[998]C)"
    'Added 7/26/16 for LXN7007F
    Range("Z1:Z1020").Value = Range("Z1:Z1020").Value
End If

'If Range("Z1") <> "" Then changed per next 7/26/16
If Range("Z1") <> 0 Then
    ' Process coordinate records when data is available
    Range("AA3:AA1020").FormulaR1C1 = "=IF(AND(RC[-2]>0,OR(RC[-2]=R2C25,RC[-2]=R1021C25)),1,"""")"
    Range("AA2:AA1020").Value = Range("AA2:AA1020").Value
    ' Extract coordinate data for first record set
    Range("AI3:AI1020").FormulaR1C1 = "=IF(R[-2]C[-8]=1,RC[-17],"""")"
    Range("AI3:AI1020").Value = Range("AI3:AI1020").Value
    ' Parse coordinate string into components (latitude/longitude degrees, minutes, direction)
    Columns("AI:AI").TextToColumns Destination:=Range("AI1"), DataType:=xlFixedWidth, OtherChar:="E", FieldInfo:=Array(Array(0, 9), Array(1, 1), Array(3, 1), Array(8, 1), Array(9, 1), Array(12, 1), Array(17, 1), Array(18, 1)), TrailingMinusNumbers:=True
    ' Sort latitude direction indicators (N/S)
    Range("AK1:AK1000").Sort Key1:=Range("AK1"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    ' Convert N/S indicators to numeric multipliers (+1 for North, -1 for South)
    Range("AK2").FormulaR1C1 = "=IF(R[-1]C=""N"",1,IF(R[-1]C=""S"",-1,0))"
    ' Sort longitude direction indicators (E/W)
    Range("AN1:AN1000").Sort Key1:=Range("AN1"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    ' Convert E/W indicators to numeric multipliers (+1 for East, -1 for West)
    Range("AN2").FormulaR1C1 = "=IF(R[-1]C=""E"",1,IF(R[-1]C""W"",-1,0))"
    ' Sort coordinate data by timestamp
    Range("AI3:AM1000").Sort Key1:=Range("AI3"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    ' Convert coordinates to decimal degrees format
    ' Latitude: degrees + (minutes * direction multiplier)
    Range("AI2").FormulaR1C1 = "=IF(AND(R[1019]C<>"""",R1C7=5),"""",R[1]C*RC[2])"
    ' Latitude minutes with direction and conversion factor
    Range("AJ2").FormulaR1C1 = "=IF(AND(R[1019]C<>"""",R1C7=5),"""",R[1]C*RC[1]*0.001)"
    ' Longitude: degrees + (minutes * direction multiplier)
    Range("AL2").FormulaR1C1 = "=IF(AND(R[1019]C<>"""",R1C7=5),"""",R[1]C*RC[2])"
    ' Longitude minutes with direction and conversion factor
    Range("AM2").FormulaR1C1 = "=IF(AND(R[1019]C<>"""",R1C7=5),"""",R[1]C*RC[1]*0.001)"
    Range("AO1:AO1000").Sort Key1:=Range("AO1"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    Range("AP4:AP1020").FormulaR1C1 = "=IF(R[-3]C[-15]=1,RC[-24],"""")"
    Range("AP4:AP1020").Value = Range("AP4:AP1020").Value
    Columns("AP:AP").TextToColumns Destination:=Range("AP1"), DataType:=xlFixedWidth, OtherChar:="E", FieldInfo:=Array(Array(0, 9), Array(1, 1), Array(3, 1), Array(8, 1), Array(9, 1), Array(12, 1), Array(17, 1), Array(18, 1)), TrailingMinusNumbers:=True
    Range("AR1:AR1000").Sort Key1:=Range("AR1"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    Range("AR2").FormulaR1C1 = "=IF(R[-1]C=""N"",1,IF(R[-1]C=""S"",-1,0))"
    Range("AU1:AU1000").Sort Key1:=Range("AU1"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    Range("AU2").FormulaR1C1 = "=IF(R[-1]C=""E"",1,IF(R[-1]C=""W"",-1,0))"
    Range("AP3:AT1000").Sort Key1:=Range("AP3"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    Range("AP2").FormulaR1C1 = "=R[1]C*RC[2]"
    Range("AQ2").FormulaR1C1 = "=R[1]C*RC[1]*0.001"
    Range("AS2").FormulaR1C1 = "=R[1]C*RC[2]"
    Range("AT2").FormulaR1C1 = "=R[1]C*RC[1]*0.001"
    Range("AV1:AV1000").Sort Key1:=Range("AV1"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
End If

If Range("Z1") >= 1 Then
    Range("AW5:AW1020").FormulaR1C1 = "=IF(R[-4]C[-22]=1,RC[-31],"""")"
    Range("AW5:AW1020").Value = Range("AW5:AW1020").Value
    Columns("AW:AW").TextToColumns Destination:=Range("AW1"), DataType:=xlFixedWidth, OtherChar:="E", FieldInfo:=Array(Array(0, 9), Array(1, 1), Array(3, 1), Array(8, 1), Array(9, 1), Array(12, 1), Array(17, 1), Array(18, 1)), TrailingMinusNumbers:=True
    Range("AY1:AY1000").Sort Key1:=Range("AY1"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    Range("AY2").FormulaR1C1 = "=IF(R[-1]C=""N"",1,IF(R[-1]C=""S"",-1,0))"
    Range("BB1:BB1000").Sort Key1:=Range("BB1"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    Range("BB2").FormulaR1C1 = "=IF(R[-1]C=""E"",1,IF(R[-1]C=""W"",-1,0))"
    Range("AW3:BA1000").Sort Key1:=Range("AW3"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    Range("AW2").FormulaR1C1 = "=R[1]C*RC[2]"
    Range("AX2").FormulaR1C1 = "=R[1]C*RC[1]*0.001"
    Range("AZ2").FormulaR1C1 = "=R[1]C*RC[2]"
    Range("BA2").FormulaR1C1 = "=R[1]C*RC[1]*0.001"
    Range("BC1:BC1000").Sort Key1:=Range("BC1"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
End If

If Range("Z1") >= 2 Then
    Range("BD6:BD1020").FormulaR1C1 = "=IF(R[-5]C[-29]=1,RC[-38],"""")"
    Range("BD6:BD1020").Value = Range("BD6:BD1020").Value
    Columns("BD:BD").TextToColumns Destination:=Range("BD1"), DataType:=xlFixedWidth, OtherChar:="E", FieldInfo:=Array(Array(0, 9), Array(1, 1), Array(3, 1), Array(8, 1), Array(9, 1), Array(12, 1), Array(17, 1), Array(18, 1)), TrailingMinusNumbers:=True
    Range("BF1:BF1000").Sort Key1:=Range("BF1"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    Range("BF2").FormulaR1C1 = "=IF(R[-1]C=""N"",1,IF(R[-1]C=""S"",-1,0))"
    Range("BI1:BI1000").Sort Key1:=Range("BI1"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    Range("BI2").FormulaR1C1 = "=IF(R[-1]C=""E"",1,IF(R[-1]C=""W"",-1,0))"
    Range("BD3:BH1000").Sort Key1:=Range("BD3"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    Range("BD2").FormulaR1C1 = "=R[1]C*RC[2]"
    Range("BE2").FormulaR1C1 = "=R[1]C*RC[1]*0.001"
    Range("BG2").FormulaR1C1 = "=R[1]C*RC[2]"
    Range("BH2").FormulaR1C1 = "=R[1]C*RC[1]*0.001"
    Range("BJ1:BJ1000").Sort Key1:=Range("BJ1"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
End If

If Range("Z1") >= 3 Then
    Range("BK7:BK1020").FormulaR1C1 = "=IF(R[-6]C[-36]=1,RC[-45],"""")"
    Range("BK7:BK1020").Value = Range("BK7:BK1020").Value
    Columns("BK:BK").TextToColumns Destination:=Range("BK1"), DataType:=xlFixedWidth, OtherChar:="E", FieldInfo:=Array(Array(0, 9), Array(1, 1), Array(3, 1), Array(8, 1), Array(9, 1), Array(12, 1), Array(17, 1), Array(18, 1)), TrailingMinusNumbers:=True
    Range("BM1:BM1000").Sort Key1:=Range("BM1"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    Range("BM2").FormulaR1C1 = "=IF(R[-1]C=""N"",1,IF(R[-1]C=""S"",-1,0))"
    Range("BK3:BO1000").Sort Key1:=Range("BK3"), Order1:=xlDescending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    Range("BK2").FormulaR1C1 = "=R[1]C*RC[2]"
    Range("BL2").FormulaR1C1 = "=R[1]C*RC[1]*0.001"
    Range("BN2").FormulaR1C1 = "=R[1]C*RC[2]"
    Range("BO2").FormulaR1C1 = "=R[1]C*RC[1]*0.001"
    Range("BP1:BP1000").Sort Key1:=Range("BP1"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    Range("BP2").FormulaR1C1 = "=IF(R[-1]C=""E"",1,IF(R[-1]C=""W"",-1,0))"
    Range("BQ1:BQ1000").Sort Key1:=Range("BQ1"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
End If

    ' Switch to PRS sheet for final data organization
    Sheets("PRS").Select
    ' Link sunrise time calculation
    Range("G13").FormulaR1C1 = "=sunrise!R17C4"
    ' Copy coordinate references from processed data
    Range("M2").FormulaR1C1 = "=Sheet1!R6C11"
    Range("M3").FormulaR1C1 = "=Sheet1!R6C12"
    ' Transfer parsed coordinate data to main processing workbook
    Workbooks("A.xlsm").Sheets("Parsed").Range("A1:H30").Value = Workbooks("Ab.xlsm").Worksheets("PRS").Range("A1:H30").Value
    ' Activate main processing workbook for subsequent operations
    Workbooks("A.xlsm").Activate
    'Application.Run "A.xlsm!CALx"
    'Sheets("Parsed").Visible = False
    'Sheets("E-Dec").Visible = True
    'Sheets("E-Dec").Activate
    'ActiveSheet.Unprotect Password:="spike"
    'Range("A1:J29").Select
    'ActiveWindow.Zoom = True
    'ActiveSheet.Protect Password:="spike"
    'Sheets("Logo").Visible = False
    'Sheets("Data Entry Check").Visible = True
    'Sheets("Data Entry Check").Select
    'ActiveSheet.Unprotect Password:="spike"
    'Range("A1:K30").Select
    'ActiveWindow.Zoom = True
    'ActiveSheet.Protect Password:="spike"
    'ActiveWindow.ScrollWorkbookTabs Sheets:=-3
    'ActiveWorkbook.Protect Password:="spike"
    'Application.Calculate
End Sub
Sub NEWHilo()
'
' JLR 4/9/14; amended 9/5/2015 for GPS altitude
'
' The NEWHilo subroutine processes flight data to determine high and low points in a flight path.
' It is typically called after the main data import and initial processing of flight logs.
' The function reads data from the "Sheet1" worksheet, which contains the processed flight data points.
' It performs a series of calculations to identify significant altitude changes, taking into account various flight parameters.
' The results, including key high and low altitude points, are written to the "Sheet2" worksheet.
' This subroutine is a critical step in analyzing flight performance and is used in preparation for generating flight summaries and reports.
' It relies on several other subroutines and functions, such as 'Duration' and 'Detangler', to ensure the data is correctly processed and categorized.
' The logic handles different scenarios based on the flight date and specific data patterns found in the logs.
'
    ' Initialize workbook and prepare data sheets
    Workbooks("Ab.xlsm").Activate
    Sheets("Sheet2").Select
    Application.Run "A.xlsm!PreB"
    Application.Run "Ab.xlsm!Duration"
    Sheets("Sheet1").Select
    ' Convert formulas to values to improve processing speed
    Range("A1:CK10010").Value = Range("A1:CK10010").Value
    ' Disable automatic calculation for performance during bulk operations
    Application.Calculation = xlCalculationManual
      ' Check if GPS altitude data is available (L5 cell empty or not)
      ' This determines which altitude column to use for calculations
      If Range("L5") = "" Then
        ' Process using barometric altitude when GPS altitude not available
        Range("BC11:BC10010").FormulaR1C1 = _
        "=IF(R[1]C[-54]="""","""",IF(AND(RC[-39]>R2C28,R[1]C[-44]>RC[-44],R[2]C[-44]>RC[-44],R[3]C[-44]>RC[-44],R[4]C[-44]>RC[-44]),1,""""))"
        Range("BD11:BD10010").FormulaR1C1 = _
        "=IF(AND(RC[-40]>R2C28,OR(R2C40=0,RC[-40]<R2C40)),RC[-45],"""")"
        Range("BE4").FormulaR1C1 = _
        "=IF(R2C40=0,MAX(R[7]C[-46]:R[10005]C[-46]),MAX(R[7]C[-1]:R[10005]C[-1]))"
        Range("BE11:BE10010").FormulaR1C1 = "=IF(RC[-1]="""","""",IF(RC[-46]=R4C57,RC[-41],""""))"
      ElseIf Range("L5") <> "" Then
        ' Process using GPS altitude when available (amended 9/5/2015)
        Range("BC11:BC10010").FormulaR1C1 = _
        "=IF(R[1]C[-54]="""","""",IF(AND(RC[-39]>R2C28,R[1]C[-43]>RC[-43],R[2]C[-43]>RC[-43],R[3]C[-43]>RC[-43],R[4]C[-43]>RC[-43]),1,""""))"
        Range("BD11:BD10010").FormulaR1C1 = _
        "=IF(AND(RC[-40]>R2C28,OR(R2C40=0,RC[-40]<R2C40)),RC[-44],"""")"
        Range("BE4").FormulaR1C1 = _
        "=IF(R2C40=0,MAX(R[7]C[-45]:R[10005]C[-45]),MAX(R[7]C[-1]:R[10005]C[-1]))"
        Range("BE11:BE10010").FormulaR1C1 = "=IF(RC[-1]="""","""",IF(RC[-45]=R4C57,RC[-41],""""))"
      End If
    ' Calculate maximum altitude in flight data range
    Range("BD4").FormulaR1C1 = "=MAX(R[7]C[1]:R[10005]C[1])"
      ' Identify low altitude points based on GPS availability
      If Range("L5") = "" Then
        Range("BG11:BG10010").FormulaR1C1 = "=IF(AND(RC[-43]>=R2C28,RC[-43]<R4C56),RC[-48],"""")"
      ElseIf Range("L5") <> "" Then
        Range("BG11:BG10010").FormulaR1C1 = "=IF(AND(RC[-43]>=R2C28,RC[-43]<R4C56,RC[-47]<>0),RC[-47],"""")"
      End If
    ' Calculate minimum altitude values for low point identification
    Range("BH4").FormulaR1C1 = "=MIN(R[7]C[-1]:R[10005]C[-1])"
    Range("BH11:BH10010").FormulaR1C1 = "=IF(RC[-1]=R4C,RC[-44],"""")"
    Range("BG4").FormulaR1C1 = "=MIN(R[7]C[1]:R[10005]C[1])"
    ' Calculate altitude difference between high and low points
    Range("BF5").FormulaR1C1 = "=R[-1]C[-1]-R[-1]C[2]"
      If Range("L5") = "" Then
        Range("BJ11:BJ10010").FormulaR1C1 = _
        "=IF(OR(RC[-46]="""",AND(R2C40>0,RC[-46]>=R2C40)),"""",IF(AND(RC[-46]>R4C56,RC[-7]=1,RC[-51]<R[-1]C[-51],RC[-51]<R[1]C[-51]),RC[-51],""""))"
        Range("BK11:BK10010").FormulaR1C1 = _
        "=IF(AND(R2C40>0,RC[-47]>=R2C40),"""",IF(AND(RC[-47]>R8C64,RC[-52]>R7C62),RC[-52],""""))"
        Range("BL11:BL10010").FormulaR1C1 = "=IF(AND(RC[-48]>R8C64,RC[-53]=R7C63),RC[-48],"""")"
        Range("BM11:BM10010").FormulaR1C1 = "=IF(AND(RC[-49]>R4C56,RC[-54]=R7C[-1]),RC[-49],"""")"
        Range("BN11:BN10010").FormulaR1C1 = _
        "=IF(OR(RC[-50]="""",AND(R2C40>0,RC[-50]>=R2C40)),"""",IF(AND(RC[-50]>R4C56,RC[-50]>R8C63,RC[-55]<R[-1]C[-55],RC[-55]<R[1]C[-55],RC[-55]<R[2]C[-55]),RC[-55],""""))"
        Range("BO11:BO10010").FormulaR1C1 = "=IF(RC[-1]="""","""",IF(RC[-56]=R7C67,RC[-51],""""))"
        Range("BP11:BP10010").FormulaR1C1 = "=IF(RC[-52]>R8C67,RC[-57],"""")"
      ElseIf Range("L5") <> "" Then
        Range("BJ11:BJ10010").FormulaR1C1 = _
        "=IF(OR(RC[-46]="""",RC[-50]=0,AND(R2C40>0,RC[-46]>=R2C40)),"""",IF(AND(RC[-46]>R4C56,RC[-7]=1,RC[-50]<R[-1]C[-50],RC[-50]<R[1]C[-50]),RC[-50],""""))"
        Range("BK11:BK10010").FormulaR1C1 = _
        "=IF(AND(R2C40>0,RC[-47]>=R2C40),"""",IF(AND(RC[-47]>R8C64,RC[-51]>R7C62),RC[-51],""""))"
        Range("BL11:BL10010").FormulaR1C1 = "=IF(AND(RC[-48]>R8C64,RC[-52]=R7C63),RC[-48],"""")"
        Range("BM11:BM10010").FormulaR1C1 = "=IF(AND(RC[-49]>R4C56,RC[-53]=R7C[-1]),RC[-49],"""")"
        Range("BN11:BN10010").FormulaR1C1 = _
        "=IF(OR(RC[-50]="""",RC[-54]=0,AND(R2C40>0,RC[-50]>=R2C40)),"""",IF(AND(RC[-50]>R4C56,RC[-50]>R8C63,RC[-54]<R[-1]C[-54],RC[-54]<R[1]C[-54],RC[-54]<R[2]C[-54]),RC[-54],""""))"
        Range("BO11:BO10010").FormulaR1C1 = "=IF(RC[-1]="""","""",IF(RC[-55]=R7C67,RC[-51],""""))"
        Range("BP11:BP10010").FormulaR1C1 = "=IF(RC[-52]>R8C67,RC[-56],"""")"
      End If
    Range("BJ7").FormulaR1C1 = "=MIN(R[4]C:R[10993]C)"
    Range("BK5").FormulaR1C1 = "=IF(R[2]C[-1]=0,"""",R[2]C-R[2]C[1])"
    Range("BK7").FormulaR1C1 = "=MAX(R[4]C:R[10002]C)"
    Range("BK8").FormulaR1C1 = "=MAX(R[3]C[1]:R[10001]C[1])"
    Range("BL7").FormulaR1C1 = "=RC[-2]"
    Range("BL8").FormulaR1C1 = "=MIN(R[3]C[1]:R[10001]C[1])"
    Range("BO7").FormulaR1C1 = "=MIN(R[4]C[-1]:R[9993]C[-1])"
    Range("BO8").FormulaR1C1 = "=MAX(R[3]C:R[9992]C)"
    Range("BP7").FormulaR1C1 = "=MAX(R[4]C:R[9993]C)"
    Range("BQ1").FormulaR1C1 = "=PRS!R[5]C[-62]"
    Range("BQ2").FormulaR1C1 = "=PRS!R[2]C[-65]"
    Range("BQ3").FormulaR1C1 = "=PRS!R[7]C[-62]"
    Range("BQ11:BQ10010").FormulaR1C1 = "=IF(RC[-1]=R7C68,RC[-53],"""")"
    ' Calculate great circle distances using haversine formula
    ' Formula: 6371 * ACOS(SIN(lat1)*SIN(lat2) + COS(lat1)*COS(lat2)*COS(lon2-lon1))
    ' where 6371 is Earth's radius in kilometers
    If Range("L5") = "" Then
    Range("BR11:BR10010").FormulaR1C1 = _
        "=IF(OR(AND(R1C[-1]=0,RC[-54]>R2C[-1],RC[-54]<R3C[-1]),AND(RC[-54]>R2C[-1],RC[-54]<MIN(R1C[-1],R3C[-1]))),6371*ACOS(SIN(R2C)*SIN((RC[-66]+(RC[-65]*0.001)/60)*PI()/180)+COS(R2C)*COS((RC[-66]+(RC[-65]*0.001)/60)*PI()/180)*COS(R3C-((RC[-63]+(RC[-62]*0.001)/60)*PI()/180))),"""")"
    ElseIf Range("L5") <> "" Then
    Range("BR11:BR10010").FormulaR1C1 = _
        "=IF(RC[-58]=0,"""",IF(OR(AND(R1C[-1]=0,RC[-54]>R2C[-1],RC[-54]<R3C[-1]),AND(RC[-54]>R2C[-1],RC[-54]<MIN(R1C[-1],R3C[-1]))),6371*ACOS(SIN(R2C)*SIN((RC[-66]+(RC[-65]*0.001)/60)*PI()/180)+COS(R2C)*COS((RC[-66]+(RC[-65]*0.001)/60)*PI()/180)*COS(R3C-((RC[-63]+(RC[-62]*0.001)/60)*PI()/180))),""""))"
    End If
    Range("BP8").FormulaR1C1 = "=MAX(R[3]C[1]:R[9992]C[1])"
    Range("BR2").FormulaR1C1 = "=(PRS!R[2]C[-64]+PRS!R[2]C[-63]/60)*PI()/180"
    Range("BR3").FormulaR1C1 = "=(PRS!R[2]C[-64]+PRS!R[2]C[-63]/60)*PI()/180"
    Range("BR8").FormulaR1C1 = "=MAX(R[3]C:R[10002]C)"
    Range("BS11:BS10010").FormulaR1C1 = "=IF(RC[-1]=R8C[-1],RC[-55],"""")"
    Range("BS9").FormulaR1C1 = "=MAX(R[2]C:R[10001]C)"
    Range("BT11:BT10010").FormulaR1C1 = "=IF(RC[-1]="""","""",RC[-68]+((RC[-67]*0.001)/60))"
    Range("BT9").FormulaR1C1 = "=MAX(R[2]C:R[10001]C)"
    Range("BU11:BU10010").FormulaR1C1 = "=IF(RC[-1]="""","""",RC[-66]+((RC[-65]*0.001)/60))"
    Range("BU9").FormulaR1C1 = "=MAX(R[2]C:R[10001]C)"
    If Range("L5") = "" Then
    Range("BV11:BV10010").FormulaR1C1 = "=IF(RC[-1]="""","""",RC[-63])"
    ElseIf Range("L5") <> "" Then
    Range("BV11:BV10010").FormulaR1C1 = "=IF(RC[-1]="""","""",RC[-62])"
    End If
    Range("BV9").FormulaR1C1 = "=MAX(R[2]C:R[10001]C)"
    Range("BW2").FormulaR1C1 = "=(PRS!R[18]C[-74]+PRS!R[18]C[-73]/60)*PI()/180"
    Range("BW3").FormulaR1C1 = "=(PRS!R[17]C[-71]+PRS!R[17]C[-70]/60)*PI()/180"
    If Range("L5") = "" Then
    Range("BW11:BW10010").FormulaR1C1 = _
        "=IF(SUM(R2C:R3C)=0,"""",IF(AND(((RC[-71]+(RC[-70]*0.001)/60)*PI()/180)=R2C,((RC[-68]+(RC[-67]*0.001)/60)*PI()/180)=R3C),0,IF(OR(AND(R1C[-6]=0,RC[-59]>R2C[-6],RC[-59]<R3C[-6]),AND(RC[-59]>R2C[-6],RC[-59]<MIN(R1C[-6],R3C[-6]))),6371*ACOS(SIN(R2C)*SIN((RC[-71]+(RC[-70]*0.001)/60)*PI()/180)+COS(R2C)*COS((RC[-71]+(RC[-70]*0.001)/60)*PI()/180)*COS(R3C-((RC[-68]+(RC[-67]*0.001)/60)*PI()/180))),"""")))"
    ElseIf Range("L5") <> "" Then
    Range("BW11:BW10010").FormulaR1C1 = _
        "=IF(OR(SUM(R2C:R3C)=0,RC[-63]=0),"""",IF(AND(((RC[-71]+(RC[-70]*0.001)/60)*PI()/180)=R2C,((RC[-68]+(RC[-67]*0.001)/60)*PI()/180)=R3C),0,IF(OR(AND(R1C[-6]=0,RC[-59]>R2C[-6],RC[-59]<R3C[-6]),AND(RC[-59]>R2C[-6],RC[-59]<MIN(R1C[-6],R3C[-6]))),6371*ACOS(SIN(R2C)*SIN((RC[-71]+(RC[-70]*0.001)/60)*PI()/180)+COS(R2C)*COS((RC[-71]+(RC[-70]*0.001)/60)*PI()/180)*COS(R3C-((RC[-68]+(RC[-67]*0.001)/60)*PI()/180))),"""")))"
    End If
    Range("BW8").FormulaR1C1 = "=MAX(R[3]C:R[10002]C)"
    Range("BX11:BX10010").FormulaR1C1 = "=IF(RC[-1]=R8C[-1],RC[-60],"""")"
    Range("BX9").FormulaR1C1 = "=MAX(R[2]C:R[10001]C)"
    Range("BY11:BY10010").FormulaR1C1 = "=IF(RC[-1]="""","""",RC[-73]+((RC[-72]*0.001)/60))"
    Range("BY9").FormulaR1C1 = "=MAX(R[2]C:R[10001]C)"
    Range("BZ11:BZ10010").FormulaR1C1 = "=IF(RC[-1]="""","""",RC[-71]+((RC[-70]*0.001)/60))"
    Range("BZ9").FormulaR1C1 = "=MAX(R[2]C:R[10001]C)"
    If Range("L5") <> "" Then
    Range("CA11:CA10010").FormulaR1C1 = "=IF(RC[-1]="""","""",RC[-68])"
    Range("CC11:CC10010").FormulaR1C1 = _
        "=IF(SUM(RC[-80],RC[-79],RC[-78])=0,"""",IF(AND(RC[-70]>PRS!R8C[-77],RC[-65]-PRS!R4C[-77]>=5/24),RC[-70],""""))"
    ElseIf Range("L5") = "" Then
    Range("CA11:CA10010").FormulaR1C1 = "=IF(RC[-1]="""","""",RC[-67])"
    Range("CC11:CC10010").FormulaR1C1 = _
        "=IF(SUM(RC[-80],RC[-79],RC[-78])=0,"""",IF(AND(RC[-69]>PRS!R8C[-77],RC[-65]-PRS!R4C[-77]>=5/24),RC[-69],""""))"
      End If
    Range("CA9").FormulaR1C1 = "=MAX(R[2]C:R[10001]C)"
    Range("CC9").FormulaR1C1 = "=MAX(R[2]C:R[10001]C)"
    Range("CD11:CD10010").FormulaR1C1 = "=IF(RC[-1]=R9C[-1],RC[-66],"""")"
    Range("CD4").FormulaR1C1 = "=IF(R[5]C=0,0,R[5]C-PRS!RC[-78])"
    Range("CD9").FormulaR1C1 = "=MAX(R[2]C:R[10001]C)"
    Range("BD2").FormulaR1C1 = _
        "=IF(MAX(R[3]C[2],R[3]C[7],R[3]C[11])=R[3]C[2],R[2]C,IF(MAX(R[3]C[2],R[3]C[7],R[3]C[11])=R[3]C[7],R[6]C[7],R[6]C[12]))"
    Range("BE2").FormulaR1C1 = _
        "=IF(MAX(R[3]C[1],R[3]C[6],R[3]C[10])=R[3]C[1],R[2]C,IF(MAX(R[3]C[1],R[3]C[6],R[3]C[10])=R[3]C[6],R[5]C[6],R[5]C[11]))"
    Range("BG2").FormulaR1C1 = _
        "=IF(MAX(R[3]C[-1],R[3]C[4],R[3]C[8])=R[3]C[-1],R[2]C,IF(MAX(R[3]C[-1],R[3]C[4],R[3]C[8])=R[3]C[4],R[6]C[5],R[6]C[8]))"
    Range("BH2").FormulaR1C1 = _
        "=IF(MAX(R[3]C[-2],R[3]C[3],R[3]C[7])=R[3]C[-2],R[2]C,IF(MAX(R[3]C[-2],R[3]C[3],R[3]C[7])=R[3]C[3],R[5]C[4],R[5]C[7]))"
    ' Re-enable automatic calculation after bulk formula operations
    Application.Calculation = xlCalculationAutomatic
    ' Convert header formulas to values for performance
    Range("A1:CK10").Value = Range("A1:CK10").Value
    ' Clear temporary calculation columns
    Range("Q11:CK10010").Clear

    ' Final calculations to identify exact high/low points and times
    Range("BD11:BD10010").FormulaR1C1 = "=IF(RC[-40]=R4C,RC[-44],"""")"
    Range("BE11:BE10010").FormulaR1C1 = "=IF(RC[-41]=R2C[-1],RC[-45],"""")"
    Range("BG11:BG10010").FormulaR1C1 = "=IF(RC[-43]=R2C,RC[-47],"""")"

    Range("BF2").FormulaR1C1 = "=MAX(R[9]C[-1]:R[10008]C[-1])"
    Range("BF4").FormulaR1C1 = "=MAX(R[7]C[-2]:R[10006]C[-2])"
    Range("BI2").FormulaR1C1 = "=MAX(R[9]C[-2]:R[10008]C[-2])"
    Range("BF2:BI4").Value = Range("BF2:BI4").Value
    Range("A11:BG10010").Clear

    Sheets("PRS").Range("J2").Value = Sheets("Sheet1").Range("BF4").Value
    Sheets("PRS").Range("J5").Value = Sheets("Sheet1").Range("BF2").Value
    Sheets("PRS").Range("J8").Value = Sheets("Sheet1").Range("BI2").Value

    'Added 10/16/2015 for ST@TP or goofy dec
    'Sheets("PRS").Activate
    'Range("J24").FormulaR1C1 = _
        "=IF(R[-10]C[-8]<>3,"""",IF(OR(AND(R[-12]C[-5]=""ST@TP"",OR(RC[1]<>R[-4]C[1],RC[2]<>R[-4]C[2])),AND(RC[1]=R[-2]C[1],RC[2]=R[-2]C[2],RC[1]<>R[2]C[1],RC[2]<>R[2]C[2]),AND(R[2]C[1]=R[4]C[1],R[2]C[2]=R[4]C[2],RC[1]<>R[4]C[1],RC[2]<>R[4]C[2])),3,""""))"
    'Range("J26").FormulaR1C1 = "=IF(OR(R[-2]C=3,AND(RC[1]=R[2]C[1],RC[2]=R[2]C[2],RC[1]<>R[-4]C[1],RC[2]<>R[-4]C[2],R[-2]C[1]<>RC[1],R[-2]C[2]<>RC[2])),2,"""")"
    'Range("J24:J26").Value = Range("J24:J26").Value
    'If Range("J24") = 3 Or Range("J26") = 2 Then
        'Range("A26:G26").Value = Range("A24:G24").Value
        'Range("A24:G24").Value = Range("A20:G20").Value
    'End If
    ' Run Detangler subroutine to finalize data processing and cleanup
    Application.Run "Ab.xlsm!Detangler"

End Sub
Sub Detangler()
'
' Determines what to do with bungled declarations; E12 value must be calculated here!ELSEWHERE, Ck dependents on E12!
' 12/16/15: Use K22:K26 for TP Order?(w/names @ S22:S26) NEED SOMETHING QUICK WHEN B14=0 or B14=1
'
    'Step 1: Do nothing if declared order is plausible, with or without duplicates
    'Range("T20:T28").Value = Range("L20:L28").Value

    Sheets("PRS").Activate
    Range("E12").FormulaR1C1 = _
        "=IF(AND(OR(R20C11<>R28C11,R20C12<>R28C12),OR(AND(R20C11=R22C11,R20C12=R22C12),AND(R20C11=R24C11,R20C12=R24C12),AND(R20C11=R26C11,R20C12=R26C12)),OR(AND(R28C11=R26C11,R28C12=R26C12),AND(R28C11=R24C11,R28C12=R24C12),AND(R28C11=R22C11,R28C12=R22C12))),""S&F@TPS"",IF(AND(R20C11=R28C11,R20C12=R28C12,OR(AND(R20C11=R22C11,R20C12=R22C12),AND(R20C11=R24C11,R20C12=R24C12),AND(R20C11=R26C11,R20C12=R26C12))),""SF@TP"",IF(OR(AND(R20C11=R22C11,R20C12=R22C12),AND(R20C11=R24C11,R20C12=R24C12),AND(R20C11=R26C11,R20C12=R26C12)),""ST@TP"",IF(OR(AND(R28C11=R26C11,R28C12=R26C12),AND(R28C11=R24C11,R28C12=R24C12),AND(R28C11=R22C11,R28C12=R22C12)),""FI@TP"",""""))))"
    If Range("B14") > 1 Then
        Range("J20").FormulaR1C1 = _
            "=IF(AND(R12C5="""",OR(AND(R14C2=3,R22C11=R24C11,R24C11=R26C11,R22C12=R24C12,R24C12=R26C12),AND(R14C2=2,R22C11=R24C11,R22C12=R24C12))),""1 TP"",IF(R28C10=0,0,IF(AND(R14C2=2,R28C10<3),""1 TP"",IF(AND(R14C2=2,R28C=3),""2 TP"",IF(AND(R14C2=3,OR(R28C10<3,AND(R22C10=""X"",R24C10=""X""),AND(R24C10=""X"",R26C10=""X""),AND(R22C10=""X"",R26C10=""X""))),""1 TP"",IF(AND(R14C2=3,OR(R22C10=""X"",R24C10=""X"",R26C10=""X"")),""2 TP"",""3 TP""))))))"
        ActiveSheet.Calculate

        Range("J22").FormulaR1C1 = _
            "=IF(AND(OR(RC[1]<>R[-2]C[1],RC[2]<>R[-2]C[2]),OR(RC[1]<>R[2]C[1],RC[2]<>R[2]C[2])),1,""X"")"
        If Range("B14") = 3 Then
            Range("J24").FormulaR1C1 = _
                "=IF(AND(OR(RC[1]<>R[-2]C[1],RC[2]<>R[-2]C[2]),OR(RC[1]<>R[2]C[1],RC[2]<>R[2]C[2])),2,""X"")"
            Range("J26").FormulaR1C1 = _
                "=IF(AND(OR(RC[1]<>R[-2]C[1],RC[2]<>R[-2]C[2]),OR(RC[1]<>R[2]C[1],RC[2]<>R[2]C[2])),3,""X"")"
        ElseIf Range("B14") = 2 Then
            Range("J24").FormulaR1C1 = _
                "=IF(AND(OR(RC[1]<>R[-2]C[1],RC[2]<>R[-2]C[2]),OR(RC[1]<>R[4]C[1],RC[2]<>R[4]C[2])),2,""X"")"
        End If
    ElseIf Range("B14") = 1 Then
        Range("J22").FormulaR1C1 = _
            "=IF(AND(OR(RC[1]<>R[-2]C[1],RC[2]<>R[-2]C[2]),OR(RC[1]<>R[6]C[1],RC[2]<>R[6]C[2])),1,""X"")"
        If Range("J22") = "X" Then
            Range("I22").Value = 0
            Range("J20").Value = "TPO"
        End If
    End If

    Range("J28").FormulaR1C1 = "=SUM(R22C:R26C)"

    If Range("B14") = 2 And Range("J28") = 3 Then
        Range("J20:J28").Clear
        Range("I22").Value = 1
        Range("I24").Value = 2
        Range("M20:S28").Value = Range("A20:G28").Value
    ElseIf Range("B14") = 3 And Range("J28") = 6 Then
        Range("J20:J28").Clear
        Range("I22").Value = 1
        Range("I24").Value = 2
        Range("I26").Value = 3
        Range("M20:S28").Value = Range("A20:G28").Value
    Exit Sub
    End If

    'Step 2a:Bungled without ST or Fin as TP
     If Range("E12") = "" Then
     'When B14 = 2 here - has to be 1 TP
        If Range("B14") = 3 And Range("J20") = "1 TP" Then
            If Range("J22") = "X" And Range("J24") = "X" And Range("J26") = "X" Then
                Range("J22").Value = 1
                Range("J24,J26").Value = "X"
            ElseIf Range("J22") = "X" And Range("J24") = "X" Then
                Range("J22").Value = 1
                Range("J24").Value = 3
                Range("J26").Value = 2
            ElseIf Range("J24") = "X" And Range("J26") = "X" Then
                If Range("K22") <> Range("K24") Or Range("L22") <> Range("L24") Then
                    Range("J22").Value = 2
                    Range("J24").Value = 1
                    Range("J26").Value = 3
                End If
             ElseIf Range("J24") = "X" And Range("J26") = "X" Then
                Range("J22").Value = 1
                Range("J24").Value = 2
             End If
        ElseIf Range("B14") = 2 And Range("J20") = "1 TP" Then
            If Range("J28") = 1 Then
                Range("J22").Value = 1
            ElseIf Range("J28") = 2 Then
                Range("J24") = 1
            End If
        End If
     End If

    'Step 2b: Bungled TP with 2 declared OK 12/14/15
    If Range("B14") = 2 Then
        If Range("E12") = "ST@TP" Then
            If Range("J28") = 2 Then
                Range("J22").Value = 2
                Range("J24").Value = 1
            End If
        ElseIf Range("E12") = "FI@TP" Then
            If Range("J28") = 1 Then
                Range("J22").Value = 2
                Range("J24").Value = 1
            End If
        ElseIf Range("E12") = "S&F@TPS" Then
            If Range("J28") = 0 Then
                Range("J22").Value = 2
                Range("J24").Value = 1
            End If
        ElseIf Range("E12") = "SF@TP" Then
            If Range("J28") = 2 Then
                Range("J22").Value = Range("J22").Value
                Range("J24").Value = 1
            ElseIf Range("J28") = 1 Then
                Range("J22:J24").Value = Range("J22:J24").Value
            End If
        End If
    End If

    'Step 2c: Bungled TP with 3 Declared
    If Range("B14") = 3 And Range("E12") = "ST@TP" Then
        'OK 12/14/15 2&3 OK 1 invalid
        If Range("J28") = 5 Then
            Range("J22").Value = "X"
            Range("J24").Value = 1
            Range("J26").Value = 2
        'OK 12/14/15
        ElseIf Range("J28") = 3 Then
            If Range("J26") = 3 Then
                If Range("K22") <> Range("K20") Or Range("L22") <> Range("L20") Then
                    Range("J22").Value = 1
                    Range("J24").Value = 3
                    Range("J26").Value = 2
                ElseIf Range("K22") = Range("K20") And Range("L22") = Range("L20") Then
                    If Range("K22") = Range("K24") And Range("L22") = Range("L24") Then
                        Range("J22").Value = "X"
                        Range("J24").Value = 2
                        Range("J26").Value = 1
                    End If
                End If
            End If
        ElseIf Range("J28") = 1 Then
            If Range("K24") = Range("K26") And Range("L24") = Range("L26") Then
                Range("J22").Value = 1
                Range("J24").Value = 2
                Range("J26").Value = "X"
            End If
        ElseIf Range("J28") = 0 And Range("K22") = Range("K24") And Range("K24") = Range("K26") And Range("L22") = Range("L24") And Range("L24") = Range("L26") Then
            Range("J22,J24,J26").Value = "X"
        ElseIf Range("J28") = 0 Then
            Range("J22").Value = 2
            Range("J24").Value = 1
            Range("J26").Value = 3
        End If

    ElseIf Range("B14") = 3 And Range("E12") = "S&F@TPS" Then
        If Range("J28") = 5 Then
            If Range("J22") = "X" Then
                If Range("K22") = Range("K26") And Range("L22") = Range("L26") Then
                    Range("J22").Value = "X"
                    Range("J24").Value = 1
                    Range("J26").Value = 2
                ElseIf Range("K26") <> Range("K28") Or Range("L26") <> Range("L28") Then
                    Range("J22").Value = 2
                    Range("J24").Value = 1
                    Range("J26").Value = 3
                ElseIf Range("K22") <> Range("K28") Or Range("L22") <> Range("L28") Then
                    Range("J22").Value = 3
                    Range("J24").Value = 1
                    Range("J26").Value = 2
                End If
            End If
            'Trial here
        ElseIf Range("J28") = 3 And Range("J26") = 3 Then
            If Range("K24") = Range("K22") And Range("L24") = Range("L22") Then
                If Range("K22") <> Range("K20") Or Range("L22") <> Range("L20") Then
                    Range("J22").Value = 1
                    Range("J24").Value = "X"
                    Range("J26").Value = 2
                End If
            End If
        ElseIf Range("J28") = 3 And Range("J26") = "X" Then
            If Range("K24") = Range("K20") And Range("L24") = Range("L20") Then
                If Range("K22") <> Range("K28") Or Range("L22") <> Range("L28") Then
                    Range("J22").Value = 3
                    Range("J24").Value = 2
                    Range("J26").Value = 1
                End If
            ElseIf Range("J26") = 3 And Range("K22") = Range("K24") And Range("L22") = Range("L24") Then
                If Range("K22") <> Range("K20") Or Range("L22") <> Range("L20") Then
                    Range("J22").Value = 1
                    Range("J24").Value = "X"
                    Range("J26").Value = 2
                End If
            End If
        ElseIf Range("J28") = 2 And Range("J24") = 2 Then
            If Range("K22") <> Range("K28") Or Range("L22") <> Range("L28") Then
                Range("J22").Value = 3
                Range("J24").Value = 2
                Range("J26").Value = 1
            End If
        ElseIf Range("J28") = 3 And Range("J26") = "X" Then
            If Range("K26") <> Range("K20") Or Range("L26") <> Range("L20") Then
                If Range("K22") <> Range("K28") Or Range("L22") <> Range("L28") Then
                    Range("J22").Value = 3
                    Range("J24").Value = Range("J24").Value
                    Range("J26").Value = 1
                ElseIf Range("K22") = Range("K28") And Range("L22") = Range("L28") Then
                    Range("J22:J26").Value = Range("J22:J26").Value
                End If
            End If
        ElseIf Range("J28") = 1 Then
            If Range("K24") = Range("K26") And Range("L24") = Range("L26") Then
                If Range("K24") <> Range("K28") Or Range("L24") <> Range("L28") Then
                    Range("J24").Value = 2
                    Range("J26").Value = "X"
                End If
            End If
        ElseIf Range("J28") = 0 Then
            If Range("K24") = Range("K28") And Range("L24") = Range("L28") And Range("K24") = Range("K26") And Range("L24") = Range("L26") Then
                Range("J22").Value = 2
                Range("J24").Value = 1
                Range("J26").Value = "X"
            ElseIf Range("K20") = Range("K22") And Range("L20") = Range("L22") And Range("K22") = Range("K24") And Range("L22") = Range("L24") Then
                Range("J22").Value = "X"
                Range("J24").Value = 2
                Range("J26").Value = 1
            End If
        End If

    ElseIf Range("B14") = 3 And Range("E12") = "SF@TP" Then
        If Range("J28") = 5 Then
            Range("J22").Value = 2
            Range("J24").Value = 1
            Range("J26").Value = 3
        ElseIf Range("J28") = 3 Then
            If Range("J22") = 1 Then
                Range("J22").Value = Range("J22").Value
                Range("J26").Value = 2
                Range("J24").Value = 3
            ElseIf Range("J22") = 2 Then
                Range("J22").Value = Range("J22").Value
                Range("J24").Value = 1
                Range("J26").Value = 3
            ElseIf Range("J26") = 3 Then
                Range("J22,J24").Value = "X"
                Range("J26").Value = 1
            End If
        ElseIf Range("J28") = 2 Then
            If Range("K22") = Range("K26") And Range("L22") = Range("L26") Then
                Range("J22").Value = "X"
                Range("J24").Value = 1
                Range("J26").Value = "X"
            End If
        ElseIf Range("J28") = 0 Then
            If Range("K22") = Range("K24") And Range("L22") = Range("L24") And Range("K24") = Range("K26") And Range("L24") = Range("L26") Then
                Range("J22,J24,J26").Value = "X"
            ElseIf Range("K24") = Range("K26") And Range("L24") = Range("L26") Then
                If Range("K22") = Range("K20") And Range("L22") = Range("L20") Then
                    Range("J22").Value = 2
                    Range("J24").Value = 1
                    Range("J26").Value = 3
                End If
            ElseIf Range("K26") <> Range("K28") Or Range("L26") <> Range("L28") Then
                'Testing!
                If Range("K22") <> Range("K24") And Range("L22") <> Range("L24") And Range("K26") <> Range("K20") And Range("L26") <> Range("L20") Then
                    Range("J22").Value = 2
                    Range("J24").Value = 3
                    Range("J26").Value = 1
                End If
            ElseIf Range("K26") = Range("K28") And Range("L26") = Range("L28") Then
                Range("J22").Value = 1
                Range("J24").Value = 3
                Range("J26").Value = 2
            End If
        End If

    ElseIf Range("B14") = 3 And Range("E12") = "FI@TP" Then
        If Range("J28") = 3 Then
            If Range("J26") = "X" Then
                If Range("K26") <> Range("K24") Or Range("L26") <> Range("L24") Then
                    If Range("K24") <> Range("K28") Or Range("L24") <> Range("L28") Then
                        If Range("K22") <> Range("K26") Or Range("L22") <> Range("L26") Then
                            Range("J22").Value = 1
                            Range("J24").Value = 3
                            Range("J26").Value = 2
                        End If
                    End If
                End If
            ElseIf Range("J26") = 3 Then
                If Range("K24") = Range("K22") And Range("L24") = Range("L22") Then
                    Range("J22").Value = 1
                    Range("J24").Value = "X"
                    Range("J26").Value = 2
                End If
            End If
        ElseIf Range("J28") = 1 Then
            If Range("K24") = Range("K26") And Range("L24") = Range("L26") Then
                    If Range("K26") <> Range("K28") Or Range("L26") <> Range("L28") Then
                        Range("J22").Value = 2
                        Range("J24").Value = 1
                        Range("J26").Value = 3
                    ElseIf Range("K24") = Range("K26") And Range("L24") = Range("L26") And Range("K26") = Range("K28") And Range("L26") = Range("L28") Then
                        Range("J22").Value = 2
                        Range("J24").Value = 1
                        Range("J26").Value = "X"
                    ElseIf Range("K26") = Range("K28") And Range("L26") = Range("L28") Then
                        Range("J22").Value = 1
                        Range("J24").Value = "X"
                        Range("J26").Value = "X"
                    End If
            End If
        ElseIf Range("J28") = 0 Then
            If Range("K22") = Range("K24") And Range("K24") = Range("K26") And Range("L22") = Range("L24") And Range("L24") = Range("L26") Then
                Range("J22,J24,J26").Value = "X"
            ElseIf Range("K26") = Range("K28") And Range("L26") = Range("L28") Then
                Range("J26").Value = 2
                If Range("K22") <> Range("K20") Or Range("L22") <> Range("L20") Then
                    Range("J22").Value = 1
                End If
                If Range("K22") = Range("K24") And Range("L22") = Range("L24") Then
                    If Range("K24") <> Range("K28") Or Range("L24") <> Range("L28") Then
                    Range("J24").Value = 3
                    End If
                End If
            ElseIf Range("K26") <> Range("K24") Or Range("L26") <> Range("L24") Then
                If Range("K22") <> Range("K28") Or Range("L22") <> Range("L28") Then
                    Range("J22").Value = 2
                    Range("J24").Value = "X"
                    Range("J26").Value = 1
                ElseIf Range("K22") = Range("K28") And Range("L22") = Range("L28") Then
                    Range("J22").Value = 1
                    Range("J24").Value = "X"
                    Range("J26").Value = 2
                End If
            ElseIf Range("K26") = Range("K24") And Range("L26") = Range("L24") Then
                    Range("J22").Value = 1
                    Range("J24").Value = 2
                    Range("J26").Value = "X"
            End If
         End If
     End If
     'Range("I22:I26").Value = Range("J22:J26").Value
     Range("J22:J26").Value = Range("J22:J26").Value

     Range("M22:S26").Value = Range("A22:G26").Value

     If Range("B14") >= 2 And Range("J22") <> "X" And Range("J24") <> "X" And Range("J26") <> "X" Then
        If Range("J24") = 1 Then
            Range("A22:G22").Value = Range("M24:S24").Value
        ElseIf Range("J26") = 1 Then
            Range("A22:G22").Value = Range("M26:S26").Value
        End If

        If Range("J22") = 2 Then
            Range("A24:G24").Value = Range("M22:S22").Value
        ElseIf Range("J26") = 2 Then
            Range("A24:G24").Value = Range("M26:S26").Value
        End If

        If Range("J22") = 3 Then
            Range("A26:G26").Value = Range("M22:S22").Value
        ElseIf Range("J24") = 3 Then
            Range("A26:G26").Value = Range("M24:S24").Value
        End If
     End If

     If Range("J20") = "1 TP" And Range("B14") = 2 Then
        If Range("J22") = "X" Then
            Range("A22:G22").Value = Range("M24:S24").Value
            Range("A24:G24").Clear
        ElseIf Range("J24") = "X" Then
            Range("A22:G22").Value = Range("M22:S22").Value
            Range("A24:G24").Clear
        End If
     ElseIf Range("J20") = "1 TP" And Range("B14") = 3 Then
        If Range("J22") = "X" And Range("J24") = "X" Then
            Range("A22:G22").Value = Range("M26:S26").Value
            Range("A24:G26").Clear
        ElseIf Range("J22") = "X" And Range("J26") = "X" Then
            Range("A22:G22").Value = Range("M24:S24").Value
            Range("A24:G26").Clear
        ElseIf Range("J24") = "X" And Range("J26") = "X" Then
            Range("A24:G26").Clear
        End If
     ElseIf Range("J20") = "2 TP" And Range("B14") = 3 Then
        If Range("J22") = "X" And Range("J24") = 1 And Range("J26") = 2 Then
            Range("A22:G24").Value = Range("M24:S26").Value
            Range("A26:G26").Clear
        ElseIf Range("J22") = "X" And Range("J26") = 1 And Range("J24") = 2 Then
            Range("A22:G22").Value = Range("M26:S26").Value
            Range("A24:G24").Value = Range("M24:S24").Value
            Range("A26:G26").Clear
        ElseIf Range("J26") = "X" And Range("J22") = 1 And Range("J24") = 2 Then
            Range("A22:G24").Value = Range("M22:S24").Value
            Range("A26:G26").Clear
        ElseIf Range("J26") = "X" And Range("J22") = 2 And Range("J24") = 1 Then
            Range("A22:G22").Value = Range("M24:S24").Value
            Range("A24:G24").Value = Range("M22:S22").Value
            Range("A26:G26").Clear
        ElseIf Range("J24") = "X" And Range("J22") = 1 And Range("J26") = 2 Then
            Range("A22:G22").Value = Range("M22:S22").Value
            Range("A24:G24").Value = Range("M26:S26").Value
            Range("A26:G26").Clear
        ElseIf Range("J24") = "X" And Range("J26") = 1 And Range("J22") = 2 Then
            Range("A22:G22").Value = Range("M26:S26").Value
            Range("A24:G24").Value = Range("M22:S22").Value
            Range("A26:G26").Clear
        End If
    End If

End Sub

Sub ClearAb()
'
' ClearAb Macro - preserves PRS equations pre-detangler  A22:G26 Revised 5/11/2018 to remove reference to C:CC desktop
'
    Sheets("PRS").Select
    Range("A22").FormulaR1C1 = "=IF(R14C2<1,0,IF(R[-9]C[3]<>3,IMP!R[-20]C[41],[A.xlsm]OTHER!R62C5))"
    Range("B22").FormulaR1C1 = "=IF(R14C2<1,0,IF(R[-9]C[2]<>3,IMP!R[-20]C[41],[A.xlsm]OTHER!R62C6))"
    Range("C22").FormulaR1C1 = "=IF(RC[-2]>0,""N"",""S"")"
    Range("D22").FormulaR1C1 = "=IF(R14C2<1,0,IF(R[-9]C<>3,IMP!R[-20]C[41],[A.xlsm]OTHER!R62C9))"
    Range("E22").FormulaR1C1 = "=IF(R14C2<1,0,IF(R[-9]C[-1]<>3,IMP!R[-20]C[41],[A.xlsm]OTHER!R62C10))"
    Range("F22").FormulaR1C1 = "=IF(RC[-2]>0,""E"",""W"")"
    Range("G22").FormulaR1C1 = "=IF(AND(R[-8]C[-5]>0,R[-9]C[-3]<>3),IMP!R[-21]C[41],[A.xlsm]OTHER!R22C3)"

    Range("A24").FormulaR1C1 = "=IF(R14C2<2,0,IF(R[-11]C[3]<>3,IMP!R[-22]C[48],[A.xlsm]OTHER!R64C5))"
    Range("B24").FormulaR1C1 = "=IF(R14C2<2,0,IF(R[-11]C[2]<>3,IMP!R[-22]C[48],[A.xlsm]OTHER!R64C6))"
    Range("C24").FormulaR1C1 = "=IF(RC[-2]>0,""N"",""S"")"
    Range("D24").FormulaR1C1 = "=IF(R14C2<2,0,IF(R[-11]C<>3,IMP!R[-22]C[48],[A.xlsm]OTHER!R64C9))"
    Range("E24").FormulaR1C1 = "=IF(R14C2<2,0,IF(R[-11]C[-1]<>3,IMP!R[-22]C[48],[A.xlsm]OTHER!R64C10))"
    Range("F24").FormulaR1C1 = "=IF(RC[-2]>0,""E"",""W"")"
    Range("G24").FormulaR1C1 = "=IF(AND(R[-10]C[-5]>=2,R[-11]C[-3]<>3),IMP!R[-23]C[48],[A.xlsm]OTHER!R24C3)"

    Range("A26").FormulaR1C1 = "=IF(AND(R[-12]C[1]>2,R[-13]C[3]<>3),IMP!R[-24]C[55],IF(R[-12]C[1]>2,[A.xlsm]OTHER!R66C5,0))"
    Range("B26").FormulaR1C1 = "=IF(AND(R14C2>2,R[-13]C[2]<>3),IMP!R[-24]C[55],IF(R[-12]C>2,[A.xlsm]OTHER!R66C6,0))"
    Range("C26").FormulaR1C1 = "=IF(RC[-2]>0,""N"",""S"")"
    Range("D26").FormulaR1C1 = "=IF(AND(R14C2>2,R[-13]C<>3),IMP!R[-24]C[55],IF(R[-12]C[-2]>2,[A.xlsm]OTHER!R66C9,0))"
    Range("E26").FormulaR1C1 = "=IF(AND(R[-12]C[-3]>2,R[-13]C[-1]<>3),IMP!R[-24]C[55],IF(R[-12]C[-3]>2,[A.xlsm]OTHER!R66C10,0))"
    Range("F26").FormulaR1C1 = "=IF(RC[-2]>0,""E"",""W"")"
    Range("G26").FormulaR1C1 = "=IF(AND(R14C2=3,R[-13]C[-3]<>3),IMP!R[-25]C[55],[A.xlsm]OTHER!R26C3)"

End Sub
Sub Duration()
' Finds best Duration w/out LoH penalty FOR AB SHEET2 After PreB; in corporates F.xlsm pressure correction; amended 9/5/15 to cite GPS alts
' 6/4/2017 Amended to add Duration on/after 10/1/2017, NO LOH LIMIT
    Application.ScreenUpdating = False

'
    Application.ScreenUpdating = False
    Sheets("Sheet2").Activate
    Sheets("Sheet2").Range("H1").Value = Sheets("PRS").Range("A2").Value
    Sheets("Sheet2").Range("H2").Value = 43009
    ' FIRST IF: Duration as of 10/01/2017 - NO LOH LIMIT

 If Range("H1") >= Range("H2") Then
         Range("O1").FormulaR1C1 = "=RC[-6]"
         Range("P1").FormulaR1C1 = "=RC[-4]"
         Range("Q1").FormulaR1C1 = "=RC[-4]"
         Range("O2").FormulaR1C1 = "=MAX(C[-6])"
         Range("P2:Q2").FormulaR1C1 = "=MAX(R[1]C:R[59998]C)"
         Range("P3").FormulaR1C1 = "=IF(RC[-7]=R2C15,RC[-4],"""")"
         Range("Q3").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-4],"""")"

        'Copy P3,Q3 per Col I
    Application.Calculation = xlCalculationManual
 With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "I").End(xlUp).Row
.Range("P3:Q3").AutoFill Destination:=.Range("P3:Q" & LastRow), Type:=xlFillDefault
    End With
Sheets("Sheet2").Calculate
    'Application.Calculation = xlCalculationAutomatic

         Range("O1:Q2").Value = Range("O1:Q2").Value
         Range("O3:Q60000").Clear

         Sheets("PRS").Range("G14").Value = Sheets("Sheet2").Range("O1").Value
         Sheets("PRS").Range("H14").Value = Sheets("Sheet2").Range("P1").Value
         Sheets("PRS").Range("G15").Value = Sheets("Sheet2").Range("Q1").Value
         Sheets("PRS").Range("G16").Value = Sheets("Sheet2").Range("O2").Value
         Sheets("PRS").Range("H16").Value = Sheets("Sheet2").Range("P2").Value
         Sheets("PRS").Range("H15").Value = Sheets("Sheet2").Range("Q2").Value
         Range("O1:Q2").Clear
     'Need following line for STD
        Range("I1:M60000").Cut Destination:=Range("I10:M60009")

ElseIf Range("H1") < Range("H2") Then

    'Leave PR altitude data alone
    Sheets("Sheet2").Activate
    Range("I1:M60000").Cut Destination:=Range("I10:M60009")
    Range("L8").FormulaR1C1 = "=SUM(R[2]C:R[103]C)"
    Range("M8").FormulaR1C1 = "=SUM(R[2]C:R[103]C)"
    Sheets("Sheet2").Calculate
    If Range("L8") = Range("M8") Then
        Range("L9").Value = 900
        Range("M9").Value = "PR"
        Range("L8:M8").Clear
    ElseIf Range("L8") <> Range("M8") Then
        Range("L9").Value = 1000
        Range("L8:M8").Clear
'HpA correction REVISED NOW AT F1:F3 on PRS
    Sheets("Sheet2").Range("R1").Value = Sheets("PRS").Range("F1").Value
    Sheets("Sheet2").Range("S1").Value = Sheets("PRS").Range("F2").Value
    Sheets("Sheet2").Range("T1").Value = Sheets("PRS").Range("F3").Value
    Range("R10").FormulaR1C1 = "=IF(RC[-9]<R1C19,RC[-6]+R1C18,RC[-6]+R1C20)"
    'Copy Ref I
    Application.Calculation = xlCalculationManual
 With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "I").End(xlUp).Row
.Range("R10:R10").AutoFill Destination:=.Range("R10:R" & LastRow), Type:=xlFillDefault
    End With
    Sheets("Sheet2").Calculate
    'Application.Calculation = xlCalculationAutomatic
    Range("L10:L60009").Value = Range("R10:R60009").Value
    Columns("R:T").Clear
    End If

    Range("I3").FormulaR1C1 = "=MAX(R[7]C:R[60006]C)"
    Range("J3").FormulaR1C1 = "=MAX(R[7]C14:R[60006]C14)"
    Range("O5").FormulaR1C1 = "=MIN(R[5]C:R[60004]C)"
    Range("P5:V5").FormulaR1C1 = "=MAX(R[5]C:R[60004]C)"
    Range("I5").FormulaR1C1 = "=MAX(R[5]C:R[60004]C)-MIN(R[5]C:R[60004]C)"
    Range("I7").FormulaR1C1 = "=R[-2]C/2"
    Range("I3:I7").Value = Range("I3:I7").Value

    Range("N10").FormulaR1C1 = "=IF(RC[-5]=R3C9,RC[-2],"""")"
    'Copy N10 Ref I
    Application.Calculation = xlCalculationManual
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "I").End(xlUp).Row
.Range("N10:N10").AutoFill Destination:=.Range("N10:N" & LastRow), Type:=xlFillDefault
    End With
Sheets("Sheet2").Calculate
'Application.Calculation = xlCalculationAutomatic

    Range("L5").FormulaR1C1 = "=R10C-R3C10"

If Range("L5") <= Range("L9") Then
        'Rel @ G1:K1; Ldg/MoP @ G2:K2
        Range("I1:M1").Value = Range("I10:M10").Value
        Range("O10:S10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-6],"""")"
        'Copy Ref I
        Application.Calculation = xlCalculationManual
    With Worksheets("Sheet2")
        LastRow = .Cells(Rows.Count, "I").End(xlUp).Row
.Range("O10:S10").AutoFill Destination:=.Range("O10:S" & LastRow), Type:=xlFillDefault
    End With
Sheets("Sheet2").Calculate
'Application.Calculation = xlCalculationAutomatic

        Range("I2:M2").FormulaR1C1 = "=MAX(R[8]C[6]:R[60009]C[6])"
        Range("I2:M2").Value = Range("I2:M2").Value

' When Rel/ldg NOT OK, look for lowest pt in first half of flight, w/ LoH < D9
ElseIf Range("L5") > Range("L9") Then
    Range("O10").FormulaR1C1 = "=IF(RC[-6]<=R10C9+R7C9,RC[-3],"""")"
    Range("P10").FormulaR1C1 = "=IF(RC[-1]=R5C15,RC[-7],"""")"
    Range("Q10").FormulaR1C1 = "=IF(AND(RC[-8]>=R5C16,RC[-5]-R3C10<R9C12),R3C9-RC[-8],"""")"
    Range("R10").FormulaR1C1 = "=IF(RC[-1]=R5C[-1],RC[-9],"""")"
    Range("S10:V10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-9],"""")"
    'Copy Ref I
    Application.Calculation = xlCalculationManual
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "I").End(xlUp).Row
.Range("O10:V10").AutoFill Destination:=.Range("O10:V" & LastRow), Type:=xlFillDefault
    End With
Sheets("Sheet2").Calculate
'Application.Calculation = xlCalculationAutomatic

    Range("P5:V5").Value = Range("P5:V5").Value
    Range("P10:V60009").Clear

    Range("W5:AB5").FormulaR1C1 = "=MAX(R[5]C:R[60004]C)"

    Range("W10").FormulaR1C1 = "=IF(AND(RC[-14]<R5C18,RC[-11]-R5C21<R9C12),R5C18-RC[-14],"""")"
    Range("X10").FormulaR1C1 = "=IF(RC[-1]=R5C23,RC[-15],"""")"
    Range("Y10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-15],"""")"
    Range("Z10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-15],"""")"
    Range("AA10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-15],"""")"
    Range("AB10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-15],"""")"
    'Copy Ref I
    Application.Calculation = xlCalculationManual
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "I").End(xlUp).Row
.Range("W10:AB10").AutoFill Destination:=.Range("W10:AB" & LastRow), Type:=xlFillDefault
    End With
Sheets("Sheet2").Calculate
'Application.Calculation = xlCalculationAutomatic

    Range("W5:AB5").Value = Range("W5:AB5").Value

    Range("P5:V5").FormulaR1C1 = "=MAX(R[5]C:R[60004]C)"

    Range("P10").FormulaR1C1 = "=IF(AND(RC[-7]>R5C24,R5C27-RC[-4]<R9C12),RC[-7]-R5C24,"""")"
    Range("Q10").FormulaR1C1 = "=IF(RC[-1]=R5C16,RC[-8],"""")"
    Range("R10:V10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-8],"""")"
    'Copy Ref I
    Application.Calculation = xlCalculationManual
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "I").End(xlUp).Row
.Range("P10:U10").AutoFill Destination:=.Range("P10:U" & LastRow), Type:=xlFillDefault
    End With
Sheets("Sheet2").Calculate
'Application.Calculation = xlCalculationAutomatic

    Range("I1:M1").Value = Range("X5:AB5").Value
    Range("I2:M2").Value = Range("Q5:U5").Value

  End If
    Columns("N:AH").Clear
    Range("I3:M7").Clear
    'Restore Pressure Altitude as recorded
    Range("L4:L5").FormulaR1C1 = "=IF(R[-3]C[-3]<=PRS!R2C6,R[-3]C-PRS!R1C6,R[-3]C-PRS!R3C6)"
    Sheets("Sheet2").Calculate
    Range("L1:L2").Value = Range("L4:L5").Value
    Range("J4:J5").Clear
    Sheets("PRS").Range("G14").Value = Sheets("Sheet2").Range("I1").Value
    Sheets("PRS").Range("H14").Value = Sheets("Sheet2").Range("L1").Value
    Sheets("PRS").Range("G15").Value = Sheets("Sheet2").Range("M1").Value
    Sheets("PRS").Range("G16").Value = Sheets("Sheet2").Range("I2").Value
    Sheets("PRS").Range("H16").Value = Sheets("Sheet2").Range("L2").Value
    Sheets("PRS").Range("H15").Value = Sheets("Sheet2").Range("M2").Value
 End If
    Range("H1:M5").Clear
    Application.Run "Ab.xlsm!STD"
    Application.Run "Ab.xlsm!LapRes"
End Sub
Sub STD()
'
' JLR 12/17/14; amended 1/8/15 for CC; amended 2/21/15 CalcSheet; amended 3/25/15 to ck reverse; 4/3/15 to run in Ab
' Amended 6/13/15 to ck STD from declared Start; Amended 9/5/15 for GPS alt; "Free" STD deleted 4/29/2018 for SWEDES mod effective 10/1/2018
Application.ScreenUpdating = False

 'NOW, STD from Release
    Range("O8:U8").FormulaR1C1 = "=MAX(R[2]C:R[60001]C)"
    Range("O7").FormulaR1C1 = "=IF(R1509C<>"""",LARGE(R[3]C:R[60002]C,1500),MAX(R[3]C:R[60002]C))"

    Range("O10").FormulaR1C1 = _
        "=IF(RC[-4]=R10C11,"""",IF(RC[-6]>R10C9,6371*ACOS(SIN(RC[-5])*SIN(R10C10)+COS(RC[-5])*COS(R10C10)*COS(R10C11-RC[-4])),""""))"
    'Copy Ref I
    Application.Calculation = xlCalculationManual
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "I").End(xlUp).Row
.Range("O10:O10").AutoFill Destination:=.Range("O10:O" & LastRow), Type:=xlFillDefault
    End With
'Application.Calculation = xlCalculationAutomatic
    Sheets("Sheet2").Calculate

    Range("P10").FormulaR1C1 = _
        "=IF(RC[-1]="""","""",IF(OR(AND(RC[-1]>100,RC[-1]>=R7C15,R10C12-RC[-4]<=R9C12),AND(R9C13="""",RC[-1]<=100,RC[-1]>=R7C15,R10C12-RC[-4]<=10*RC[-1]),(AND(R9C13=""PR"",RC[-1]<=100,RC[-1]>R7C15,R10C4-RC[-4]<=(10*RC[-1])-100))),RC[-1],""""))"

    'OLD thru 9/30/2018; NEW as of 10/1/2018
    If Range("A1") < 43374 Then
        Range("Q10").FormulaR1C1 = "=IF(RC[-1]=R8C16,RC[-8],"""")"

    ElseIf Range("A1") >= 43374 Then
        Range("O4").FormulaR1C1 = "=RADIANS(PRS!R[14]C[-14]+(PRS!R[14]C[-13]/60))"
        Range("O5").FormulaR1C1 = "=RADIANS(PRS!R[13]C[-11]+(PRS!R[13]C[-10]/60))"
        Range("Q10").FormulaR1C1 = "=IF(AND(RC[-1]=R8C16,6371*ACOS(SIN(RC[-7])*SIN(R4C15)+COS(RC[-7])*COS(R4C15)*COS(R5C15-RC[-6]))>=50),RC[-8],"""")"
    End If

 Range("R10:U10").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-8],"""")"
    'Copy Ref I
     Application.Calculation = xlCalculationManual
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "I").End(xlUp).Row
.Range("P10:U10").AutoFill Destination:=.Range("P10:U" & LastRow), Type:=xlFillDefault
    End With
'Application.Calculation = xlCalculationAutomatic
    Sheets("Sheet2").Calculate

    'For Longest Flight(!)
 If Range("P8") = 0 Then
    Range("O7").FormulaR1C1 = "=LARGE(R[3]C:R[60002]C,3000)"
    Sheets("Sheet2").Calculate
 End If

    'For Mphillip (no Silver - max STD from Rel)
 If Range("Q8") = 0 Then
    Range("Q10").FormulaR1C1 = "=IF(RC[-1]=R8C16,RC[-8],"""")"
    With Worksheets("Sheet2")
    LastRow = .Cells(Rows.Count, "I").End(xlUp).Row
.Range("Q10:Q10").AutoFill Destination:=.Range("Q10:Q" & LastRow), Type:=xlFillDefault
    End With
    Sheets("Sheet2").Calculate
 End If

    Range("I1:M1").Value = Range("Q8:U8").Value
     'CK best Fix from St Pt
    Range("J10").FormulaR1C1 = "=RADIANS(PRS!R[10]C[-9]+(PRS!R[10]C[-8]/60))"
    Range("K10").FormulaR1C1 = "=RADIANS(PRS!R[10]C[-7]+(PRS!R[10]C[-6]/60))"
    ActiveSheet.Calculate
    Range("I2:M2").Value = Range("Q8:U8").Value
    'RESTORE REL COORDS!!
    Range("J10").FormulaR1C1 = "=RADIANS(PRS!R4C6+(PRS!R4C7/60))"
    Range("K10").FormulaR1C1 = "=RADIANS(PRS!R5C6+(PRS!R5C7/60))"
    Range("J10:K10").Value = Range("J10:K10").Value

    Range("O7:U60009").Clear

   'Restore Pressure Altitude as recorded
    Sheets("Sheet2").Range("R1").Value = Sheets("PRS").Range("F1").Value
    Sheets("Sheet2").Range("S1").Value = Sheets("PRS").Range("F2").Value
    Sheets("Sheet2").Range("T1").Value = Sheets("PRS").Range("F3").Value

    Range("P1:P2").FormulaR1C1 = "=IF(RC[-7]<=R1C19,RC[-4]-R1C18,RC[-4]-R1C20)"
    Sheets("Sheet2").Calculate
    Range("L1:L2").Value = Range("P1:P2").Value
    Range("P1:T2").Clear

   'Put it somewhere!
    Sheets("PRS").Range("J11").Value = Sheets("Sheet2").Range("I1").Value
    Sheets("PRS").Range("J12").Value = Sheets("Sheet2").Range("J1").Value
    Sheets("PRS").Range("J13").Value = Sheets("Sheet2").Range("K1").Value
    Sheets("PRS").Range("J14").Value = Sheets("Sheet2").Range("L1").Value
    Sheets("PRS").Range("J15").Value = Sheets("Sheet2").Range("M1").Value
    Sheets("PRS").Range("I11").Value = Sheets("Sheet2").Range("I2").Value
    Sheets("PRS").Range("I12").Value = Sheets("Sheet2").Range("J2").Value
    Sheets("PRS").Range("I13").Value = Sheets("Sheet2").Range("K2").Value
    Sheets("PRS").Range("I14").Value = Sheets("Sheet2").Range("L2").Value
    Sheets("PRS").Range("I15").Value = Sheets("Sheet2").Range("M2").Value
    Sheets("Sheet2").Activate
    Range("I1:O6").Clear
    Range("A1").Activate

End Sub

Sub LapRes()
'
' 2/2/18 Works - 4 seconds  TESTED IN C for use in Ab after STD; amended 2/5/18 to ref Col A for brevity now 3 secs
'

Sheets("Sheet2").Activate

    Range("O1").FormulaR1C1 = "=SUM(R[1]C[3]:R[5]C[3])"
    Range("O2").FormulaR1C1 = "=IF(AND(R[-1]C[1]=R[4]C[1],R[-1]C[2]=R[4]C[2],MIN(RC[3],R[2]C[3],R[4]C[3])>=0.28*R[-1]C),""FAI"","""")"
    Range("O3").FormulaR1C1 = "=PRS!R14C2"
    Range("P1").FormulaR1C1 = "=PRS!R20C11"
    Range("Q1").FormulaR1C1 = "=PRS!R20C12"

    Range("P2").FormulaR1C1 = "=PRS!R22C11"
    Range("Q2").FormulaR1C1 = "=PRS!R22C12"
    Range("R2").FormulaR1C1 = "=6371*ACOS(SIN(RC[-2])*SIN(R[-1]C[-2])+COS(RC[-2])*COS(R[-1]C[-2])*COS(RC[-1]-R[-1]C[-1]))"
    Range("S2").Value = "First Leg"
    Range("P4").FormulaR1C1 = "=PRS!R24C11"
    Range("Q4").FormulaR1C1 = "=PRS!R24C12"
    Range("R4").FormulaR1C1 = "=6371*ACOS(SIN(RC[-2])*SIN(R[-2]C[-2])+COS(RC[-2])*COS(R[-2]C[-2])*COS(R[-2]C[-1]-RC[-1]))"
    Range("S4").Value = "2nd leg"

    Range("P6").FormulaR1C1 = "=PRS!R28C11"
    Range("Q6").FormulaR1C1 = "=PRS!R28C12"
    Range("R6").FormulaR1C1 = "=6371*ACOS(SIN(R[-2]C[-2])*SIN(RC[-2])+COS(R[-2]C[-2])*COS(RC[-2])*COS(R[-2]C[-1]-RC[-1]))"
    Range("S6").Value = "last leg"

    Range("R9").Value = 1000

    ActiveSheet.Calculate

    'TRIGGER
    If Range("O1") > 150 Or Range("O2") <> "FAI" Then
        Range("Q1:S6").Clear
        Exit Sub
    ElseIf Range("O1") < 150 And Range("O2") = "FAI" Then

    ' First TP DO THIS BEFORE START!
    Range("AA3").FormulaR1C1 = "=MAX(R[7]C28:R[10009]C28)"
    Range("AA10").FormulaR1C1 = "=6371*ACOS(SIN(R[-9]C[-24])*SIN(R2C16)+COS(R[-9]C[-24])*COS(R2C16)*COS(R2C17-R[-9]C[-22]))"
    Range("AB10").FormulaR1C1 = "=IF(AND(RC[-1]<5,6371*ACOS(SIN(R[-9]C[-25])*SIN(R1C16)+COS(R[-9]C[-25])*COS(R1C16)*COS(R1C17-R[-9]C[-23]))>=R2C18,6371*ACOS(SIN(R[-9]C[-25])*SIN(R4C16)+COS(R[-9]C[-25])*COS(R4C16)*COS(R4C17-R[-9]C[-23]))>=R4C18),R[-9]C[-27],"""")"
    Range("AC10").FormulaR1C1 = "=IF(AND(RC[-2]<5,R[-1]C[-1]=""""),RC[-1],"""")"
    'Copy Ref A
    With Worksheets("Sheet2")
     LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("AA10:AC10").AutoFill Destination:=.Range("AA10:AC" & LastRow), Type:=xlFillDefault
    End With

    ActiveSheet.Calculate
    Range("AA3").Value = Range("AA3").Value
    Range("AA10:AC10009").Value = Range("AA10:AC10009").Value
    'SORT Range("AC1:AC60000")
     Range("AC1:AC10009").Sort Key1:=Range("AC10"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal

    'Start Line Times
    Range("U10").FormulaR1C1 = "=6371*ACOS(SIN(R[-9]C[-18])*SIN(R1C16)+COS(R[-9]C[-18])*COS(R1C16)*COS(R[-9]C[-16]-R1C17))"
    Range("V10").FormulaR1C1 = "=IF(RC[-1]<=0.5,R[-9]C[-21],"""")"
    Range("W10").FormulaR1C1 = "=IF(AND(RC[-1]<>"""",6371*ACOS(SIN(R[-9]C[-20])*SIN(R2C16)+COS(R[-9]C[-20])*COS(R2C16)*COS(R2C17-R[-9]C[-18]))>=R2C18,6371*ACOS(SIN(R[-8]C[-20])*SIN(R2C16)+COS(R[-8]C[-20])*COS(R2C16)*COS(R2C17-R[-8]C[-18]))<=R2C18,RC[-1]<R3C27),RC[-1],"""")"
    Range("X10").FormulaR1C1 = "=IF(AND(RC[-1]<>"""",R[1]C[-1]=""""),RC[-1],"""")"
    Range("Y10").FormulaR1C1 = "=IF(RC[-1]<>"""",R[-9]C[-19],"""")"
    'Copy Ref A
    With Worksheets("Sheet2")
     LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("U10:Y10").AutoFill Destination:=.Range("U10:Y" & LastRow), Type:=xlFillDefault
    End With

    ActiveSheet.Calculate

    Range("X1:Y10009").Value = Range("X1:Y10009").Value
    Range("X1:Y10009").Sort Key1:=Range("X1"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal

    Range("U10:Y10009").Clear

    'Start Line Correction
    Range("Z1:Z10").FormulaR1C1 = "=IF(RC[-1]="""","""",IF(RC[-2]<=PRS!R2C6,RC[-1] +PRS!R1C6,IF(RC[-2]>PRS!R2C6,RC[-1]+PRS!R3C6)))"
    ActiveSheet.Calculate
    Range("Y1:Y10").Value = Range("Z1:Z10").Value
    Range("Z1:Z10").Clear

    'Last TP OZ  DO THIS AFTER START!!
    Range("AE3").FormulaR1C1 = "=MAX(R[7]C:R[10009]C)"
    Range("AD10").FormulaR1C1 = "=6371*ACOS(SIN(R[-9]C[-27])*SIN(R4C16)+COS(R[-9]C[-27])*COS(R4C16)*COS(R4C17-R[-9]C[-25]))"
    Range("AE10").FormulaR1C1 = "=IF(AND(RC[-1]<5,6371*ACOS(SIN(R[-9]C[-28])*SIN(R6C16)+COS(R[-9]C[-28])*COS(R6C16)*COS(R6C17-R[-9]C[-26]))>=R6C18,6371*ACOS(SIN(R[-9]C[-28])*SIN(R2C16)+COS(R[-9]C[-28])*COS(R2C16)*COS(R2C17-R[-9]C[-26]))>=R4C4),R[-9]C[-30],"""")"
    Range("AF10").FormulaR1C1 = "=IF(R[1]C[-1]="""",RC[-1],"""")"

     With Worksheets("Sheet2")
     LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("AD10:AF10").AutoFill Destination:=.Range("AD10:AF" & LastRow), Type:=xlFillDefault
    End With

    ActiveSheet.Calculate

    Range("AE3").Value = Range("AE3").Value

    Range("AF1:AF10009").Value = Range("AF1:AF10009").Value
    'SORT
    Range("AF1:AF10009").Sort Key1:=Range("AF1"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal

    Range("AA10:AF10009").ClearContents

    'Finish Line Time
    Range("T10").FormulaR1C1 = "=6371*ACOS(SIN(R[-9]C[-17])*SIN(R6C16)+COS(R[-9]C[-17])*COS(R6C16)*COS(R[-9]C[-15]-R6C17))"
    Range("U10").FormulaR1C1 = "=IF(AND(RC[-1]<=0.5,6371*ACOS(SIN(R[-9]C[-18])*SIN(R4C16)+COS(R[-9]C[-18])*COS(R4C16)*COS(R4C17-R[-9]C[-16]))>=R6C18,R[-9]C[-20]>R3C27,R[-9]C[-20]>R3C31),R[-9]C[-20],"""")"
    Range("V10").FormulaR1C1 = "=IF(AND(RC[-1]<>"""",R[-1]C[-1]=""""),RC[-1],"""")"
    Range("W10").FormulaR1C1 = "=IF(AND(RC[-1]<>"""",R[-1]C[-1]=""""),R[-9]C[-17],"""")"
    'Copy Ref A
    With Worksheets("Sheet2")
     LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("T10:W10").AutoFill Destination:=.Range("T10:W" & LastRow), Type:=xlFillDefault
    End With

     ActiveSheet.Calculate

     Range("T10:W10009").Value = Range("T10:W10009").Value
     Range("T10:U10009").Clear
    'SORT VW
     Range("V10:W10009").Sort Key1:=Range("V10"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal

    ' FinAltCorrection
    Range("U10:U16").FormulaR1C1 = "=IF(RC[2]="""","""",IF(RC[1]<=PRS!R2C6,RC[2]+PRS!R1C6,IF(RC[1]>PRS!R2C6,RC[2]+PRS!R3C6)))"
    ActiveSheet.Calculate
    Range("W10:W16").Value = Range("U10:U16").Value
    Range("U10:U16").Clear

    'SORTHELPER
    Range("AA3,AE3").Clear
    Range("U10:U16").FormulaR1C1 = "=IF(RC[1]<>"""",""FIN"","""")"
    Range("Z1:Z10").FormulaR1C1 = "=IF(RC[-1]<>"""",""ST"","""")"
    Range("AD1:AD6").FormulaR1C1 = "=IF(RC[-1]<>"""",1,"""")"
    Range("AG1:AG6").FormulaR1C1 = "=IF(RC[-1]<>"""",2,"""")"

    ActiveSheet.Calculate

    Range("U10:U16").Value = Range("U10:U16").Value
    Range("Z1:Z10").Value = Range("Z1:Z10").Value
    Range("AD1:AD6").Value = Range("AD1:AD6").Value
    Range("AG1:AG6").Value = Range("AG1:AG6").Value

    Range("X11:X16").Value = Range("AC1:AC6").Value
    Range("Z11:Z16").Value = Range("AD1:AD6").Value
    Range("X17:X22").Value = Range("AF1:AF6").Value
    Range("Z17:Z22").Value = Range("AG1:AG6").Value
    Range("X23:X28").Value = Range("V10:V15").Value
    Range("Y23:Y28").Value = Range("W10:W15").Value
    Range("Z23:Z28").Value = Range("U10:U15").Value

    'SORT LAPS
    Range("X1:Z28").Sort Key1:=Range("X1"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    Range("X1:X28").NumberFormat = "h:mm:ss;@"

    Range("AC1:AG6").Clear

    ' BEST 3 ETs including LoH
    Range("AB1:AB28").FormulaR1C1 = "=IF(AND(RC[-2]=""ST"",R10C22>RC[-4],R10C22<>"""",RC[-3]-R10C23<R9C18),R10C22-RC[-4],"""")"
    Range("AC1:AC28").FormulaR1C1 = "=IF(AND(RC[-3]=""ST"",R11C22>RC[-5],R11C22<>"""",RC[-4]-R11C23<R9C18),R11C22-RC[-5],"""")"
    Range("AD1:AD28").FormulaR1C1 = "=IF(AND(RC[-4]=""ST"",R12C22>RC[-6],R12C22<>"""",RC[-5]-R12C23<R9C18),R12C22-RC[-6],"""")"
    Range("AE1:AE28").FormulaR1C1 = "=IF(AND(RC[-5]=""ST"",R13C22>RC[-7],R13C22<>"""",RC[-6]-R13C23<R9C18),R13C22-RC[-7],"""")"
    Range("AF1:AF28").FormulaR1C1 = "=IF(AND(RC[-6]=""ST"",R14C22>RC[-8],R14C22<>"""",RC[-7]-R14C23<R9C18),R14C22-RC[-8],"""")"
    Range("AG1:AG28").FormulaR1C1 = "=IF(AND(RC[-7]=""ST"",R15C22>RC[-9],R15C22<>"""",RC[-8]-R15C23<R9C18),R15C22-RC[-9],"""")"
    ActiveSheet.Calculate
    Range("AB1:AG28").Value = Range("AB1:AG28").Value

    Range("V1").FormulaR1C1 = "=MIN(RC[6]:R[27]C[11])"
    Range("V2").FormulaR1C1 = "=SMALL(R[-1]C[6]:R[26]C[11],2)"
    Range("V3").FormulaR1C1 = "=SMALL(R[-2]C[6]:R[25]C[11],3)"
    ActiveSheet.Calculate
    Range("V1:V3").Value = Range("V1:V3").Value

    Range("AC1:AC28").FormulaR1C1 = "=IF(RC[-1]="""","""",IF(OR(RC[-1]=R1C22,RC[-1]=R2C22,RC[-1]=R3C22),RC[-5]))"
    ActiveSheet.Calculate
    Range("AC1:AC28").Value = Range("AC1:AC28").Value
    'SORT
    Range("AB1:AC28").Sort Key1:=Range("AB1"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal

    'Put candidate Start options somewhere. BEST AT AC1

    Sheets("PRS").Range("C13").Value = Sheets("Sheet2").Range("AC1").Value
    Range("AD1").FormulaR1C1 = "=RC[-1]-((1/24)/30)"
    ActiveSheet.Calculate
    Range("O1:U1").Value = Range("A1:G1").Value
    Range("O2").FormulaR1C1 = "=IF(RC[-14]>=R1C30,RC[-14],"""")"
    Range("P2:U2").FormulaR1C1 = "=IF(RC[-1]<>"""",RC[-14],"""")"
    'Copy Ref A
    With Worksheets("Sheet2")
     LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
.Range("O2:U2").AutoFill Destination:=.Range("O2:U" & LastRow), Type:=xlFillDefault
    End With

    ActiveSheet.Calculate
    Range("A1:G10000").Value = Range("O1:U10000").Value

    Columns("O:AG").Clear
  End If

End Sub
