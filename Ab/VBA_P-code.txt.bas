' Processing file: Ab_vbaProject.bin
' ===============================================================================
' Module streams:
' VBA/ThisWorkbook - 1837 bytes
' Line #0:
' 	Option  (Explicit)
' Line #1:
' 	FuncDefn (Private Sub Workbook_Open())
' Line #2:
' Line #3:
' 	LitVarSpecial (False)
' 	Ld Application 
' 	MemSt EnableAutoComplete 
' Line #4:
' 	LitDI2 0x0032 
' 	Ld Application 
' 	MemSt MaxIterations 
' Line #5:
' 	LitR8 0xA9FC 0xD2F1 0x624D 0x3F50 
' 	Ld Application 
' 	MemSt MaxChange 
' Line #6:
' 	EndSub 
' VBA/Sheet1 - 1121 bytes
' VBA/Sheet2 - 1174 bytes
' VBA/Sheet3 - 1121 bytes
' VBA/Module1 - 201098 bytes
' Line #0:
' 	Option  (Explicit)
' Line #1:
' 	Dim 
' 	VarDefn myFile (As Variant)
' Line #2:
' 	LbMark 
' 	Ld VBA7 
' 	Ld Win64 
' 	And 
' 	LbIf 
' Line #3:
' 	Dim 
' 	VarDefn LastRow (As LongPtr)
' Line #4:
' 	LbMark 
' 	LbElse 
' Line #5:
' 	Dim 
' 	VarDefn LastRow (As Long)
' Line #6:
' 	LbMark 
' 	LbEndIf 
' Line #7:
' Line #8:
' 	FuncDefn (Sub NewENLA())
' Line #9:
' 	QuoteRem 0x0000 0x0000 ""
' Line #10:
' 	QuoteRem 0x0000 0x0035 " JLR 3/27/12 7/26/16 Corrected O49 on Imp (no = sign)"
' Line #11:
' 	QuoteRem 0x0000 0x0062 " JLR 3/2/17 Added Range 07 now used in Corrected Z11:Z10010 & AL25:AL10010; corrected AM25:AM10010"
' Line #12:
' 	QuoteRem 0x0000 0x0053 " JLR 6/3/17 Amended '' for interim MoP @ ZZ25, AA3, AA4, AA25, AL25 Works for Mirja"
' Line #13:
' 	QuoteRem 0x0000 0x009D " JLR 7/6/17 Amended Z2,Z3,Z4,Z25:et al & AL4,AL5 & AN4 for NON-interim MoP (eg: engine run before task ONLY) for Sibylle Andresen, July 2 & 3 WORLD RECORDS!!"
' Line #14:
' 	QuoteRem 0x0000 0x0098 " JLR 7/7/17 Amended AA3 & AL4 to differentiate between Mirja (re-start after task) from Sibylle (No re-start after task); re-activated O11:AB10010 value"
' Line #15:
' 	QuoteRem 0x0000 0x0066 " JLR 11/30/2017 Amended AL4 for Schart (no re-start after second ENL); Z2,Z4,Z25,AL4,AN4,AO11 for Anja"
' Line #16:
' 	QuoteRem 0x0000 0x0038 "' JLR 07/14/2018 AN11 amended for landing time (Howard2)"
' Line #17:
' 	QuoteRem 0x0004 0x001D "Workbooks("Ab.xlsm").Activate"
' Line #18:
' 	QuoteRem 0x0004 0x002A "ActiveWorkbook.Unprotect Password:="spike""
' Line #19:
' 	QuoteRem 0x0004 0x0013 "Sheets("BR").Select"
' Line #20:
' 	QuoteRem 0x0004 0x0012 "Range("A1").Select"
' Line #21:
' 	QuoteRem 0x0004 0x0011 "ActiveSheet.Paste"
' Line #22:
' 	QuoteRem 0x0004 0x001F "Application.CutCopyMode = False"
' Line #23:
' 	QuoteRem 0x0004 0x001C "Workbooks("A.xlsm").Activate"
' Line #24:
' 	QuoteRem 0x0004 0x0026 "ActiveWindow.WindowState = xlMaximized"
' Line #25:
' 	QuoteRem 0x0004 0x0021 "Application.ScreenUpdating = True"
' Line #26:
' 	QuoteRem 0x0004 0x0028 "ActiveWindow.DisplayWorkbookTabs = False"
' Line #27:
' 	LitVarSpecial (False)
' 	Ld Application 
' 	MemSt ScreenUpdating 
' Line #28:
' 	LitStr 0x0007 "Ab.xlsm"
' 	ArgsLd Workbooks 0x0001 
' 	ArgsMemCall Activate 0x0000 
' Line #29:
' 	LitStr 0x0009 "A1:A60000"
' 	LitStr 0x0002 "BR"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0009 "A1:A60000"
' 	LitStr 0x0002 "BR"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemSt Range 0x0001 
' Line #30:
' 	LitStr 0x0002 "B1"
' 	ArgsLd Range 0x0001 
' 	ParamNamed Destination 
' 	Ld xlFixedWidth 
' 	ParamNamed DataType 
' 	LitDI2 0x0000 
' 	LitDI2 0x0001 
' 	ArgsArray Array 0x0002 
' 	LitDI2 0x0001 
' 	LitDI2 0x0001 
' 	ArgsArray Array 0x0002 
' 	ArgsArray Array 0x0002 
' 	ParamNamed FieldInfo 
' 	LitVarSpecial (True)
' 	ParamNamed TrailingMinusNumbers 
' 	LitStr 0x0003 "A:A"
' 	ArgsLd Columns 0x0001 
' 	ArgsMemCall TextToColumns 0x0004 
' Line #31:
' 	LitStr 0x001A "=IF(RC[-2]<>"B",RC[-3],"")"
' 	LitStr 0x0008 "D1:D1001"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #32:
' 	LitStr 0x0014 "=IF(RC[-3]="B",1,"")"
' 	LitStr 0x0009 "E3:E60002"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #33:
' 	LitStr 0x0009 "E3:E60002"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0009 "E3:E60002"
' 	ArgsSt Range 0x0001 
' Line #34:
' 	LitStr 0x0015 "=SUM(R[3]C:R[60000]C)"
' 	LitStr 0x0002 "E1"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #35:
' 	LitStr 0x0002 "E1"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0002 "E1"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #36:
' 	LitStr 0x0002 "D2"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0002 "E2"
' 	ArgsSt Range 0x0001 
' Line #37:
' 	LineCont 0x0008 11 00 08 00 28 00 08 00
' 	LitStr 0x0002 "E2"
' 	ArgsLd Range 0x0001 
' 	ParamNamed Destination 
' 	Ld xlFixedWidth 
' 	ParamNamed DataType 
' 	LitStr 0x0001 "E"
' 	ParamNamed OtherChar 
' 	LitDI2 0x0000 
' 	LitDI2 0x0009 
' 	ArgsArray Array 0x0002 
' 	LitDI2 0x0005 
' 	LitDI2 0x0004 
' 	ArgsArray Array 0x0002 
' 	ArgsArray Array 0x0002 
' 	ParamNamed FieldInfo 
' 	LitVarSpecial (True)
' 	ParamNamed TrailingMinusNumbers 
' 	LitStr 0x0002 "E2"
' 	ArgsLd Range 0x0001 
' 	ArgsMemCall TextToColumns 0x0005 
' Line #38:
' 	LitStr 0x0017 "=IF(RC[-1]=1,RC[-3],"")"
' 	LitStr 0x0009 "F3:F60002"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #39:
' 	LitStr 0x0009 "F3:F60002"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0009 "F3:F60002"
' 	ArgsSt Range 0x0001 
' Line #40:
' 	LineCont 0x0008 11 00 08 00 33 00 08 00
' 	LitStr 0x0002 "F1"
' 	ArgsLd Range 0x0001 
' 	ParamNamed Destination 
' 	Ld xlFixedWidth 
' 	ParamNamed DataType 
' 	LitStr 0x0001 "E"
' 	ParamNamed OtherChar 
' 	LitDI2 0x0000 
' 	LitDI2 0x0001 
' 	ArgsArray Array 0x0002 
' 	LitDI2 0x0002 
' 	LitDI2 0x0001 
' 	ArgsArray Array 0x0002 
' 	LitDI2 0x0004 
' 	LitDI2 0x0001 
' 	ArgsArray Array 0x0002 
' 	LitDI2 0x0006 
' 	LitDI2 0x0009 
' 	ArgsArray Array 0x0002 
' 	ArgsArray Array 0x0004 
' 	ParamNamed FieldInfo 
' 	LitVarSpecial (True)
' 	ParamNamed TrailingMinusNumbers 
' 	LitStr 0x0003 "F:F"
' 	ArgsLd Columns 0x0001 
' 	ArgsMemCall TextToColumns 0x0005 
' Line #41:
' 	LitStr 0x0015 "=MAX(R[2]C:R[49998]C)"
' 	LitStr 0x0002 "F2"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #42:
' 	LitStr 0x0002 "F2"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0002 "F2"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #43:
' 	LitStr 0x008D "=IF(AND(R[-1]C[-3]<>"B",R[-2]C[-3]<>"B",R[-3]C[-3]<>"B",R[-4]C[-3]<>"B",R[-5]C[-3]<>"B",R[-6]C[-3]<>"B",R[-7]C[-3]<>"B",RC[-3]="B"),RC[1],"")"
' 	LitStr 0x0007 "E8:E600"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #44:
' 	LitStr 0x001B "=MAX(R[5]C[-1]:R[597]C[-1])"
' 	LitStr 0x0002 "F3"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #45:
' 	LitStr 0x0002 "F3"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0002 "F3"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #46:
' 	LineCont 0x0004 07 00 08 00
' 	LitStr 0x0076 "=IF(RC[-2]="","",IF(AND(RC[-4]<=R2C6,RC[-4]>=R3C6),TIME(RC[-4],RC[-3],RC[-2])+R2C5,TIME(RC[-4],RC[-3],RC[-2])+R2C5+1))"
' 	LitStr 0x0009 "J2:J60001"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #47:
' 	LitStr 0x0018 "=IF(RC[-1]="","",RC[-8])"
' 	LitStr 0x0009 "K2:K60001"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #48:
' 	LitStr 0x0009 "J2:K60001"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0009 "J2:K60001"
' 	ArgsSt Range 0x0001 
' Line #49:
' 	LitStr 0x0002 "J1"
' 	ArgsLd Range 0x0001 
' 	ParamNamed Key1 
' 	Ld xlAscending 
' 	ParamNamed Order1 
' 	Ld xlGuess 
' 	ParamNamed Header 
' 	LitDI2 0x0001 
' 	ParamNamed OrderCustom 
' 	LitVarSpecial (False)
' 	ParamNamed MatchCase 
' 	Ld xlTopToBottom 
' 	ParamNamed Orientation 
' 	Ld xlSortNormal 
' 	ParamNamed DataOption1 
' 	LitStr 0x0009 "J1:K60000"
' 	ArgsLd Range 0x0001 
' 	ArgsMemCall Sort 0x0007 
' Line #50:
' Line #51:
' 	LitStr 0x0001 "1"
' 	LitStr 0x0002 "L1"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #52:
' 	LitStr 0x0020 "=IF(RC[-1]=""=FALSE,R[-1]C+1,"")"
' 	LitStr 0x0009 "L2:L60000"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #53:
' 	LitStr 0x0009 "L2:L60000"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0009 "L2:L60000"
' 	ArgsSt Range 0x0001 
' Line #54:
' Line #55:
' 	LitStr 0x0002 "E1"
' 	ArgsLd Range 0x0001 
' 	LitDI4 0x9C40 0x0000 
' 	Gt 
' 	IfBlock 
' Line #56:
' 	LitStr 0x001D "=IF(SUM(R[-5]C:R[-1]C)=0,1,0)"
' 	LitStr 0x0009 "P6:P60000"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #57:
' 	LitStr 0x0009 "P6:P60000"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0009 "P6:P60000"
' 	ArgsSt Range 0x0001 
' Line #58:
' Line #59:
' 	LitStr 0x0002 "E1"
' 	ArgsLd Range 0x0001 
' 	LitDI2 0x7530 
' 	Gt 
' 	ElseIfBlock 
' Line #60:
' 	LitStr 0x001D "=IF(SUM(R[-3]C:R[-1]C)=0,1,0)"
' 	LitStr 0x0009 "O4:O40000"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #61:
' 	LitStr 0x0009 "O4:O40000"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0009 "O4:O40000"
' 	ArgsSt Range 0x0001 
' Line #62:
' Line #63:
' 	LitStr 0x0002 "E1"
' 	ArgsLd Range 0x0001 
' 	LitDI2 0x4E20 
' 	Gt 
' 	ElseIfBlock 
' Line #64:
' 	LitStr 0x001D "=IF(SUM(R[-2]C:R[-1]C)>0,0,1)"
' 	LitStr 0x0009 "N3:N30000"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #65:
' 	LitStr 0x0009 "N3:N30000"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0009 "N3:N30000"
' 	ArgsSt Range 0x0001 
' Line #66:
' Line #67:
' 	LitStr 0x0002 "E1"
' 	ArgsLd Range 0x0001 
' 	LitDI2 0x2710 
' 	Gt 
' 	ElseIfBlock 
' Line #68:
' 	LitStr 0x0011 "=IF(R[-1]C=0,1,0)"
' 	LitStr 0x0009 "M2:M20000"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #69:
' 	LitStr 0x0009 "M2:M20000"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0009 "M2:M20000"
' 	ArgsSt Range 0x0001 
' Line #70:
' 	EndIfBlock 
' Line #71:
' Line #72:
' 	LitStr 0x0001 "1"
' 	LitStr 0x0002 "Q1"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #73:
' 	LitStr 0x00EB "=IF(RC[-5]="","",IF(R1C5<=10000,RC[-5],IF(AND(R1C5>10000,R1C5<=20000,RC[-4]=1),RC[-5],IF(AND(R1C5>20000,R1C5<=30000,RC[-3]=1),RC[-5],IF(AND(R1C5>30000,R1C5<=40000,RC[-2]=1),RC[-5],IF(AND(R1C5>40000,R1C5<=60000,RC[-1]=1),RC[-5],""))))))"
' 	LitStr 0x0009 "Q2:Q60000"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #74:
' 	LitStr 0x0009 "Q1:Q60000"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0009 "Q1:Q60000"
' 	ArgsSt Range 0x0001 
' Line #75:
' 	LitStr 0x001C "=IF(RC[-1]=RC[-6],RC[-7],"")"
' 	LitStr 0x0009 "R1:R60000"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #76:
' 	LitStr 0x0009 "R1:R60000"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0009 "R1:R60000"
' 	ArgsSt Range 0x0001 
' Line #77:
' 	LitStr 0x0002 "Q1"
' 	ArgsLd Range 0x0001 
' 	ParamNamed Key1 
' 	Ld xlAscending 
' 	ParamNamed Order1 
' 	Ld xlGuess 
' 	ParamNamed Header 
' 	LitDI2 0x0001 
' 	ParamNamed OrderCustom 
' 	LitVarSpecial (False)
' 	ParamNamed MatchCase 
' 	Ld xlTopToBottom 
' 	ParamNamed Orientation 
' 	Ld xlSortNormal 
' 	ParamNamed DataOption1 
' 	LitStr 0x0009 "Q1:R60000"
' 	ArgsLd Range 0x0001 
' 	ArgsMemCall Sort 0x0007 
' Line #78:
' Line #79:
' 	LitStr 0x0009 "R1:R10000"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0009 "X1:X10000"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #80:
' 	LineCont 0x000C 11 00 08 00 32 00 08 00 60 00 08 00
' 	LitStr 0x0002 "W1"
' 	ArgsLd Range 0x0001 
' 	ParamNamed Destination 
' 	Ld xlFixedWidth 
' 	ParamNamed DataType 
' 	LitStr 0x0001 "E"
' 	ParamNamed OtherChar 
' 	LitDI2 0x0000 
' 	LitDI2 0x0001 
' 	ArgsArray Array 0x0002 
' 	LitDI2 0x0006 
' 	LitDI2 0x0001 
' 	ArgsArray Array 0x0002 
' 	LitDI2 0x0008 
' 	LitDI2 0x0001 
' 	ArgsArray Array 0x0002 
' 	LitDI2 0x000A 
' 	LitDI2 0x0001 
' 	ArgsArray Array 0x0002 
' 	LitDI2 0x000D 
' 	LitDI2 0x0001 
' 	ArgsArray Array 0x0002 
' 	LitDI2 0x000E 
' 	LitDI2 0x0001 
' 	ArgsArray Array 0x0002 
' 	LitDI2 0x0011 
' 	LitDI2 0x0001 
' 	ArgsArray Array 0x0002 
' 	LitDI2 0x0013 
' 	LitDI2 0x0001 
' 	ArgsArray Array 0x0002 
' 	LitDI2 0x0016 
' 	LitDI2 0x0001 
' 	ArgsArray Array 0x0002 
' 	LitDI2 0x0017 
' 	LitDI2 0x0001 
' 	ArgsArray Array 0x0002 
' 	ArgsArray Array 0x000A 
' 	ParamNamed FieldInfo 
' 	LitVarSpecial (True)
' 	ParamNamed TrailingMinusNumbers 
' 	LitStr 0x0003 "X:X"
' 	ArgsLd Columns 0x0001 
' 	ArgsMemCall TextToColumns 0x0005 
' Line #81:
' 	LitStr 0x0009 "Q1:R10000"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x000B "AK1:AL10000"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #82:
' 	LitStr 0x0003 "AK1"
' 	ArgsLd Range 0x0001 
' 	ParamNamed Key1 
' 	Ld xlAscending 
' 	ParamNamed Order1 
' 	Ld xlGuess 
' 	ParamNamed Header 
' 	LitDI2 0x0001 
' 	ParamNamed OrderCustom 
' 	LitVarSpecial (False)
' 	ParamNamed MatchCase 
' 	Ld xlTopToBottom 
' 	ParamNamed Orientation 
' 	Ld xlSortNormal 
' 	ParamNamed DataOption1 
' 	LitStr 0x000B "AK1:AL10000"
' 	ArgsLd Range 0x0001 
' 	ArgsMemCall Sort 0x0007 
' Line #83:
' Line #84:
' 	LitStr 0x0008 "D1:D1000"
' 	LitStr 0x0002 "BR"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0008 "A1:A1000"
' 	LitStr 0x0003 "IMP"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemSt Value 
' Line #85:
' Line #86:
' 	LitStr 0x0003 "IMP"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemCall Select 0x0000 
' Line #87:
' 	LitStr 0x0008 "A1:A1000"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0008 "B1:B1000"
' 	ArgsSt Range 0x0001 
' Line #88:
' 	LineCont 0x0004 11 00 08 00
' 	LitStr 0x0002 "B1"
' 	ArgsLd Range 0x0001 
' 	ParamNamed Destination 
' 	Ld xlFixedWidth 
' 	ParamNamed DataType 
' 	LitDI2 0x0000 
' 	LitDI2 0x0001 
' 	ArgsArray Array 0x0002 
' 	LitDI2 0x0001 
' 	LitDI2 0x0001 
' 	ArgsArray Array 0x0002 
' 	ArgsArray Array 0x0002 
' 	ParamNamed FieldInfo 
' 	LitVarSpecial (True)
' 	ParamNamed TrailingMinusNumbers 
' 	LitStr 0x0003 "B:B"
' 	ArgsLd Columns 0x0001 
' 	ArgsMemCall TextToColumns 0x0004 
' Line #89:
' 	LineCont 0x0008 11 00 08 00 28 00 08 00
' 	LitStr 0x0002 "C1"
' 	ArgsLd Range 0x0001 
' 	ParamNamed Destination 
' 	Ld xlFixedWidth 
' 	ParamNamed DataType 
' 	LitStr 0x0001 "E"
' 	ParamNamed OtherChar 
' 	LitDI2 0x0000 
' 	LitDI2 0x0001 
' 	ArgsArray Array 0x0002 
' 	LitDI2 0x0001 
' 	LitDI2 0x0009 
' 	ArgsArray Array 0x0002 
' 	ArgsArray Array 0x0002 
' 	ParamNamed FieldInfo 
' 	LitVarSpecial (True)
' 	ParamNamed TrailingMinusNumbers 
' 	LitStr 0x0002 "C1"
' 	ArgsLd Range 0x0001 
' 	ArgsMemCall TextToColumns 0x0005 
' Line #90:
' 	LitStr 0x0002 "C1"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0001 "X"
' 	Eq 
' 	IfBlock 
' Line #91:
' 	LitStr 0x0002 "A1"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0002 "D1"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #92:
' 	LineCont 0x0008 11 00 0C 00 2F 00 0C 00
' 	LitStr 0x0002 "D1"
' 	ArgsLd Range 0x0001 
' 	ParamNamed Destination 
' 	Ld xlFixedWidth 
' 	ParamNamed DataType 
' 	LitStr 0x0001 ":"
' 	ParamNamed OtherChar 
' 	LitDI2 0x0000 
' 	LitDI2 0x0009 
' 	ArgsArray Array 0x0002 
' 	LitDI2 0x0002 
' 	LitDI2 0x0001 
' 	ArgsArray Array 0x0002 
' 	LitDI2 0x0004 
' 	LitDI2 0x0001 
' 	ArgsArray Array 0x0002 
' 	ArgsArray Array 0x0003 
' 	ParamNamed FieldInfo 
' 	LitVarSpecial (True)
' 	ParamNamed TrailingMinusNumbers 
' 	LitStr 0x0002 "D1"
' 	ArgsLd Range 0x0001 
' 	ArgsMemCall TextToColumns 0x0005 
' Line #93:
' 	LitStr 0x0002 "C1"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0001 "X"
' 	Eq 
' 	LitVarSpecial (False)
' 	Eq 
' 	ElseIfBlock 
' Line #94:
' 	LitStr 0x0002 "A1"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0002 "D1"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #95:
' 	LineCont 0x0004 11 00 0C 00
' 	LitStr 0x0002 "D1"
' 	ArgsLd Range 0x0001 
' 	ParamNamed Destination 
' 	Ld xlFixedWidth 
' 	ParamNamed DataType 
' 	LitStr 0x0001 "E"
' 	ParamNamed OtherChar 
' 	LitDI2 0x0000 
' 	LitDI2 0x0009 
' 	ArgsArray Array 0x0002 
' 	LitDI2 0x0001 
' 	LitDI2 0x0001 
' 	ArgsArray Array 0x0002 
' 	LitDI2 0x0004 
' 	LitDI2 0x0001 
' 	ArgsArray Array 0x0002 
' 	LitDI2 0x0007 
' 	LitDI2 0x0009 
' 	ArgsArray Array 0x0002 
' 	ArgsArray Array 0x0004 
' 	ParamNamed FieldInfo 
' 	LitVarSpecial (True)
' 	ParamNamed TrailingMinusNumbers 
' 	LitStr 0x0002 "D1"
' 	ArgsLd Range 0x0001 
' 	ArgsMemCall TextToColumns 0x0005 
' Line #96:
' 	EndIfBlock 
' Line #97:
' 	LitStr 0x0019 "=IF(RC[-2]="H",RC[-1],"")"
' 	LitStr 0x0002 "D2"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #98:
' 	LitStr 0x0002 "D2"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0002 "D2"
' 	ArgsSt Range 0x0001 
' Line #99:
' 	LineCont 0x0008 11 00 08 00 28 00 08 00
' 	LitStr 0x0002 "D2"
' 	ArgsLd Range 0x0001 
' 	ParamNamed Destination 
' 	Ld xlFixedWidth 
' 	ParamNamed DataType 
' 	LitStr 0x0001 ":"
' 	ParamNamed OtherChar 
' 	LitDI2 0x0000 
' 	LitDI2 0x0001 
' 	ArgsArray Array 0x0002 
' 	LitDI2 0x0004 
' 	LitDI2 0x0004 
' 	ArgsArray Array 0x0002 
' 	ArgsArray Array 0x0002 
' 	ParamNamed FieldInfo 
' 	LitVarSpecial (True)
' 	ParamNamed TrailingMinusNumbers 
' 	LitStr 0x0002 "D2"
' 	ArgsLd Range 0x0001 
' 	ArgsMemCall TextToColumns 0x0005 
' Line #100:
' 	LitStr 0x0006 "C4:C50"
' 	ArgsLd Range 0x0001 
' 	ArgsMemCall Select 0x0000 
' Line #101:
' 	LineCont 0x0008 0E 00 08 00 25 00 08 00
' 	LitStr 0x0002 "C4"
' 	ArgsLd Range 0x0001 
' 	ParamNamed Destination 
' 	Ld xlFixedWidth 
' 	ParamNamed DataType 
' 	LitStr 0x0001 ":"
' 	ParamNamed OtherChar 
' 	LitDI2 0x0000 
' 	LitDI2 0x0001 
' 	ArgsArray Array 0x0002 
' 	LitDI2 0x0004 
' 	LitDI2 0x0002 
' 	ArgsArray Array 0x0002 
' 	ArgsArray Array 0x0002 
' 	ParamNamed FieldInfo 
' 	LitVarSpecial (True)
' 	ParamNamed TrailingMinusNumbers 
' 	Ld Selection 
' 	ArgsMemCall TextToColumns 0x0005 
' Line #102:
' 	LitStr 0x0006 "D4:D50"
' 	ArgsLd Range 0x0001 
' 	ArgsMemCall Select 0x0000 
' Line #103:
' 	LineCont 0x000C 0E 00 08 00 1A 00 08 00 2B 00 08 00
' 	LitStr 0x0002 "D4"
' 	ArgsLd Range 0x0001 
' 	ParamNamed Destination 
' 	Ld xlDelimited 
' 	ParamNamed DataType 
' 	Ld xlDoubleQuote 
' 	ParamNamed TextQualifier 
' 	LitVarSpecial (False)
' 	ParamNamed ConsecutiveDelimiter 
' 	LitVarSpecial (False)
' 	ParamNamed Tab 
' 	LitVarSpecial (False)
' 	ParamNamed Semicolon 
' 	LitVarSpecial (False)
' 	ParamNamed Comma 
' 	LitVarSpecial (False)
' 	ParamNamed Space 
' 	LitVarSpecial (True)
' 	ParamNamed Other 
' 	LitStr 0x0001 ":"
' 	ParamNamed OtherChar 
' 	LitDI2 0x0001 
' 	LitDI2 0x0002 
' 	ArgsArray Array 0x0002 
' 	LitDI2 0x0002 
' 	LitDI2 0x0002 
' 	ArgsArray Array 0x0002 
' 	ArgsArray Array 0x0002 
' 	ParamNamed FieldInfo 
' 	LitVarSpecial (True)
' 	ParamNamed TrailingMinusNumbers 
' 	Ld Selection 
' 	ArgsMemCall TextToColumns 0x000C 
' Line #104:
' Line #105:
' 	QuoteRem 0x0003 0x0029 "PUT H RECORDS IN TECH SPECS ORDER 4/16/15"
' Line #106:
' 	LitStr 0x0004 "FPLT"
' 	LitStr 0x0002 "R1"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #107:
' 	LitStr 0x0004 "FCM2"
' 	LitStr 0x0002 "S1"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #108:
' 	LitStr 0x0004 "FGTY"
' 	LitStr 0x0002 "T1"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #109:
' 	LitStr 0x0004 "FGID"
' 	LitStr 0x0002 "U1"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #110:
' 	LitStr 0x0004 "FDTM"
' 	LitStr 0x0002 "V1"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #111:
' 	LitStr 0x0004 "FRFW"
' 	LitStr 0x0002 "W1"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #112:
' 	LitStr 0x0004 "FRHW"
' 	LitStr 0x0002 "X1"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #113:
' 	LitStr 0x0004 "FFTY"
' 	LitStr 0x0002 "Y1"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #114:
' 	LitStr 0x0004 "FGPS"
' 	LitStr 0x0002 "Z1"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #115:
' 	LitStr 0x0004 "FPRS"
' 	LitStr 0x0003 "AA1"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #116:
' Line #117:
' 	LitStr 0x0013 "=IF(RC2="H",RC3,"")"
' 	LitStr 0x0006 "L4:L50"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #118:
' 	LitStr 0x0013 "=IF(RC2="H",RC4,"")"
' 	LitStr 0x0006 "M4:M50"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #119:
' 	LitStr 0x0013 "=IF(RC2="H",RC5,"")"
' 	LitStr 0x0006 "N4:N50"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #120:
' Line #121:
' 	LitStr 0x0006 "L4:N50"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0006 "L4:N50"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #122:
' Line #123:
' 	LitStr 0x002E "=IF(OR(RC12=R1C,RC12="OPLT",RC12="PPLT"),1,"")"
' 	LitStr 0x0006 "R4:R50"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #124:
' 	LitStr 0x002E "=IF(OR(RC12=R1C,RC12="OCM2",RC12="PCM2"),2,"")"
' 	LitStr 0x0006 "S4:S50"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #125:
' 	LitStr 0x0032 "=IF(OR(RC12=R1C,R24C12="OGTY",R24C12="PGTY"),3,"")"
' 	LitStr 0x0006 "T4:T50"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #126:
' 	LitStr 0x002E "=IF(OR(RC12=R1C,RC12="OGID",RC12="PGID"),4,"")"
' 	LitStr 0x0006 "U4:U50"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #127:
' 	LitStr 0x0012 "=IF(RC12=R1C,5,"")"
' 	LitStr 0x0006 "V4:V50"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #128:
' 	LitStr 0x0012 "=IF(RC12=R1C,6,"")"
' 	LitStr 0x0006 "W4:W50"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #129:
' 	LitStr 0x0012 "=IF(RC12=R1C,7,"")"
' 	LitStr 0x0006 "X4:X50"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #130:
' 	LitStr 0x0012 "=IF(RC12=R1C,8,"")"
' 	LitStr 0x0006 "Y4:Y50"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #131:
' 	LitStr 0x0012 "=IF(RC12=R1C,9,"")"
' 	LitStr 0x0006 "Z4:Z50"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #132:
' 	LitStr 0x0013 "=IF(RC12=R1C,10,"")"
' 	LitStr 0x0008 "AA4:AA50"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #133:
' Line #134:
' 	LitStr 0x0007 "R4:AA50"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0007 "R4:AA50"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #135:
' Line #136:
' 	LitStr 0x002D "=IF(MAX(RC[3]:RC[13])=0,"",MAX(RC[3]:RC[13]))"
' 	LitStr 0x0006 "O4:O50"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #137:
' 	LitStr 0x0006 "O4:O50"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0006 "O4:O50"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #138:
' Line #139:
' 	LitStr 0x0003 "O10"
' 	ArgsLd Range 0x0001 
' 	ParamNamed Key1 
' 	Ld xlAscending 
' 	ParamNamed Order1 
' 	Ld xlGuess 
' 	ParamNamed Header 
' 	LitDI2 0x0001 
' 	ParamNamed OrderCustom 
' 	LitVarSpecial (False)
' 	ParamNamed MatchCase 
' 	Ld xlTopToBottom 
' 	ParamNamed Orientation 
' 	Ld xlSortNormal 
' 	ParamNamed DataOption1 
' 	LitStr 0x0006 "L4:O50"
' 	ArgsLd Range 0x0001 
' 	ArgsMemCall Sort 0x0007 
' Line #140:
' 	QuoteRem 0x0004 0x0080 "Amended 4/30/15 delete former ref to FW/HW conundrum, go with order only; revised 5/5/15 for strict order; no FFTY added 5/11/15"
' Line #141:
' 	LitStr 0x000F "=IF(R4C=1,"",1)"
' 	LitStr 0x0003 "O45"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #142:
' 	LitStr 0x000F "=IF(R5C=2,"",2)"
' 	LitStr 0x0003 "O46"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #143:
' 	LitStr 0x0037 "=IF(OR(R4C=3,R5C=3,R6C=3),"",IF(AND(R6C=4,R7C=5),"",3))"
' 	LitStr 0x0003 "O47"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #144:
' 	LitStr 0x003D "=IF(OR(R4C=4,R5C=4,R6C=4,R7C=4),"",IF(AND(R7C=5,R8C=6),"",4))"
' 	LitStr 0x0003 "O48"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #145:
' 	LitStr 0x0046 "=IF(MAX(R4C:R13C)=7,8,IF(MAX(R4C:R13C)=8,9,IF(MAX(R4C:R13C)=9,10,"")))"
' 	LitStr 0x0003 "O49"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #146:
' 	LitStr 0x0007 "O45:O49"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0007 "O45:O49"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #147:
' Line #148:
' 	LitStr 0x0003 "O10"
' 	ArgsLd Range 0x0001 
' 	ParamNamed Key1 
' 	Ld xlAscending 
' 	ParamNamed Order1 
' 	Ld xlGuess 
' 	ParamNamed Header 
' 	LitDI2 0x0001 
' 	ParamNamed OrderCustom 
' 	LitVarSpecial (False)
' 	ParamNamed MatchCase 
' 	Ld xlTopToBottom 
' 	ParamNamed Orientation 
' 	Ld xlSortNormal 
' 	ParamNamed DataOption1 
' 	LitStr 0x0006 "L4:O50"
' 	ArgsLd Range 0x0001 
' 	ArgsMemCall Sort 0x0007 
' Line #149:
' 	LitStr 0x0006 "L4:N13"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0006 "C4:E13"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #150:
' 	LitStr 0x0015 "R1:AB1,J9:J10,L4:AB50"
' 	ArgsLd Range 0x0001 
' 	ArgsMemCall Clear 0x0000 
' Line #151:
' Line #152:
' 	LitStr 0x0019 "=IF(RC[-5]="I",RC[-6],"")"
' 	LitStr 0x0006 "G2:G50"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #153:
' 	LitStr 0x0006 "G2:G50"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0006 "G2:G50"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #154:
' 	LitStr 0x0014 "=IF(RC[-6]<>"I",0,1)"
' 	LitStr 0x0006 "H2:H50"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #155:
' 	LitStr 0x0012 "=SUM(R[1]C:R[49]C)"
' 	LitStr 0x0002 "H1"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #156:
' 	LitStr 0x0002 "H1"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0002 "H1"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #157:
' 	LitStr 0x0002 "H1"
' 	ArgsLd Range 0x0001 
' 	LitDI2 0x0000 
' 	Gt 
' 	IfBlock 
' Line #158:
' 	LitStr 0x0003 "ENL"
' 	LitStr 0x0002 "H2"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #159:
' 	LitStr 0x0043 "=IF(RC[-5]<>"I","",IF(ISERROR(FIND(R2C[1],RC[-6]))=TRUE,"",RC[-6]))"
' 	LitStr 0x0006 "G3:G50"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #160:
' 	LitStr 0x0006 "G3:G50"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0006 "G3:G50"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #161:
' 	LitStr 0x0022 "=IF(RC[-1]="","",FIND(R2C,RC[-1]))"
' 	LitStr 0x0006 "H3:H50"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #162:
' 	LitStr 0x001C "=MAX(R[1]C[-1]:R[48]C[-1])-4"
' 	LitStr 0x0002 "I2"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #163:
' 	LitStr 0x0002 "I2"
' 	ArgsLd Range 0x0001 
' 	LitDI2 0x0000 
' 	Gt 
' 	IfBlock 
' Line #164:
' 	LitStr 0x0012 "=MID(RC[-2],R2C,2)"
' 	LitStr 0x0006 "I3:I50"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #165:
' 	LitStr 0x0006 "I3:I50"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0006 "I3:I50"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #166:
' 	LitStr 0x001A "=MAX(R[2]C[-8]:R[49]C[-8])"
' 	LitStr 0x0002 "Q1"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #167:
' 	LitStr 0x0002 "Q1"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0002 "Q1"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #168:
' 	ElseBlock 
' 	BoS 0x0000 
' 	LitDI2 0x0000 
' 	LitStr 0x0002 "Q1"
' 	ArgsSt Range 0x0001 
' Line #169:
' 	EndIfBlock 
' Line #170:
' 	LitStr 0x0003 "MOP"
' 	LitStr 0x0002 "H2"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #171:
' 	LitStr 0x0043 "=IF(RC[-5]<>"I","",IF(ISERROR(FIND(R2C[1],RC[-6]))=TRUE,"",RC[-6]))"
' 	LitStr 0x0006 "G3:G50"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #172:
' 	LitStr 0x0006 "G3:G50"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0006 "G3:G50"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #173:
' 	LitStr 0x0022 "=IF(RC[-1]="","",FIND(R2C,RC[-1]))"
' 	LitStr 0x0006 "H3:H50"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #174:
' 	LitStr 0x001C "=MAX(R[1]C[-1]:R[48]C[-1])-4"
' 	LitStr 0x0002 "I2"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #175:
' 	LitStr 0x0002 "I2"
' 	ArgsLd Range 0x0001 
' 	LitDI2 0x0000 
' 	Gt 
' 	IfBlock 
' Line #176:
' 	LitStr 0x0012 "=MID(RC[-2],R2C,2)"
' 	LitStr 0x0006 "I3:I50"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #177:
' 	LitStr 0x0006 "I3:I50"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0006 "I3:I50"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #178:
' 	LitStr 0x001A "=MAX(R[2]C[-8]:R[49]C[-8])"
' 	LitStr 0x0002 "Q2"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #179:
' 	LitStr 0x0002 "Q2"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0002 "Q2"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #180:
' 	ElseBlock 
' 	BoS 0x0000 
' 	LitDI2 0x0000 
' 	LitStr 0x0002 "Q2"
' 	ArgsSt Range 0x0001 
' Line #181:
' 	EndIfBlock 
' Line #182:
' 	EndIfBlock 
' Line #183:
' 	LitStr 0x001B "='[A.xlsm]ALL CLAIMS'!R18C6"
' 	LitStr 0x0002 "G1"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #184:
' Line #185:
' 	LitStr 0x0006 "Sheet1"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemCall Select 0x0000 
' Line #186:
' 	LitStr 0x003C "=IF(IMP!R1C8=0,"NO I Record",IF(IMP!R1C17=0,"No ENL","ENL"))"
' 	LitStr 0x0002 "A1"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #187:
' 	LitStr 0x000B "AL1:AL10000"
' 	LitStr 0x0002 "BR"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x000A "A11:A10010"
' 	LitStr 0x0006 "Sheet1"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemSt Value 
' Line #188:
' 	LineCont 0x000C 11 00 08 00 33 00 08 00 61 00 08 00
' 	LitStr 0x0003 "A11"
' 	ArgsLd Range 0x0001 
' 	ParamNamed Destination 
' 	Ld xlFixedWidth 
' 	ParamNamed DataType 
' 	LitStr 0x0001 "E"
' 	ParamNamed OtherChar 
' 	LitDI2 0x0000 
' 	LitDI2 0x0001 
' 	ArgsArray Array 0x0002 
' 	LitDI2 0x0002 
' 	LitDI2 0x0001 
' 	ArgsArray Array 0x0002 
' 	LitDI2 0x0004 
' 	LitDI2 0x0001 
' 	ArgsArray Array 0x0002 
' 	LitDI2 0x0006 
' 	LitDI2 0x0001 
' 	ArgsArray Array 0x0002 
' 	LitDI2 0x0008 
' 	LitDI2 0x0001 
' 	ArgsArray Array 0x0002 
' 	LitDI2 0x000D 
' 	LitDI2 0x0001 
' 	ArgsArray Array 0x0002 
' 	LitDI2 0x000E 
' 	LitDI2 0x0001 
' 	ArgsArray Array 0x0002 
' 	LitDI2 0x0011 
' 	LitDI2 0x0001 
' 	ArgsArray Array 0x0002 
' 	LitDI2 0x0016 
' 	LitDI2 0x0001 
' 	ArgsArray Array 0x0002 
' 	LitDI2 0x0017 
' 	LitDI2 0x0001 
' 	ArgsArray Array 0x0002 
' 	LitDI2 0x0018 
' 	LitDI2 0x0001 
' 	ArgsArray Array 0x0002 
' 	LitDI2 0x001D 
' 	LitDI2 0x0001 
' 	ArgsArray Array 0x0002 
' 	LitDI2 0x0022 
' 	LitDI2 0x0002 
' 	ArgsArray Array 0x0002 
' 	ArgsArray Array 0x000D 
' 	ParamNamed FieldInfo 
' 	LitVarSpecial (True)
' 	ParamNamed TrailingMinusNumbers 
' 	LitStr 0x000A "A11:A10010"
' 	ArgsLd Range 0x0001 
' 	ArgsMemCall TextToColumns 0x0005 
' Line #189:
' Line #190:
' 	Dim 
' 	VarDefn MyCell
' Line #191:
' 	Ld xlCalculationManual 
' 	Ld Application 
' 	MemSt Calculation 
' Line #192:
' 	LitStr 0x000A "D11:E10010"
' 	ArgsLd Range 0x0001 
' 	ArgsMemCall Select 0x0000 
' Line #193:
' 	StartForVariable 
' 	Ld MyCell 
' 	EndForVariable 
' 	Ld Selection 
' 	MemLd Cells 
' 	ForEach 
' Line #194:
' 	LitStr 0x0001 "F"
' 	Ld MyCell 
' 	MemLd Row 
' 	Concat 
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0001 "S"
' 	Eq 
' 	IfBlock 
' Line #195:
' 	Ld MyCell 
' 	MemLd Value 
' 	LitDI2 0x0001 
' 	UMi 
' 	Paren 
' 	Mul 
' 	Ld MyCell 
' 	MemSt Value 
' Line #196:
' 	EndIfBlock 
' Line #197:
' 	StartForVariable 
' 	Next 
' Line #198:
' 	LitStr 0x000A "G11:H10010"
' 	ArgsLd Range 0x0001 
' 	ArgsMemCall Select 0x0000 
' Line #199:
' 	StartForVariable 
' 	Ld MyCell 
' 	EndForVariable 
' 	Ld Selection 
' 	MemLd Cells 
' 	ForEach 
' Line #200:
' 	LitStr 0x0001 "I"
' 	Ld MyCell 
' 	MemLd Row 
' 	Concat 
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0001 "W"
' 	Eq 
' 	IfBlock 
' Line #201:
' 	Ld MyCell 
' 	MemLd Value 
' 	LitDI2 0x0001 
' 	UMi 
' 	Paren 
' 	Mul 
' 	Ld MyCell 
' 	MemSt Value 
' Line #202:
' 	EndIfBlock 
' Line #203:
' 	StartForVariable 
' 	Next 
' Line #204:
' 	Ld xlCalculationAutomatic 
' 	Ld Application 
' 	MemSt Calculation 
' Line #205:
' Line #206:
' 	LitStr 0x0002 "A1"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0003 "ENL"
' 	Eq 
' 	IfBlock 
' Line #207:
' 	LitStr 0x0002 "Q1"
' 	LitStr 0x0003 "IMP"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0002 "O1"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #208:
' 	LitStr 0x0002 "Q2"
' 	LitStr 0x0003 "IMP"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0002 "P1"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #209:
' 	LitStr 0x000A "=R[-1]C-35"
' 	LitStr 0x0005 "O2:P2"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #210:
' 	LitStr 0x0005 "O2:P2"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0005 "O2:P2"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #211:
' 	LitStr 0x0014 "=MID(RC[-2],R2C15,3)"
' 	LitStr 0x000A "O11:O10010"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #212:
' 	LitStr 0x0002 "P2"
' 	ArgsLd Range 0x0001 
' 	LitDI2 0x0000 
' 	Gt 
' 	IfBlock 
' Line #213:
' 	LitStr 0x0014 "=MID(RC[-2],R2C16,3)"
' 	LitStr 0x000A "P11:P10010"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #214:
' 	EndIfBlock 
' Line #215:
' 	EndIfBlock 
' Line #216:
' 	LitStr 0x000A "O11:P10010"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x000A "O11:P10010"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #217:
' 	LitStr 0x0003 "M:N"
' 	ArgsLd Columns 0x0001 
' 	ParamNamed Destination 
' 	LitStr 0x0003 "O:P"
' 	ArgsLd Columns 0x0001 
' 	ArgsMemCall Cut 0x0001 
' Line #218:
' 	LitStr 0x0005 "M1:N2"
' 	ArgsLd Range 0x0001 
' 	ArgsMemCall Select 0x0000 
' Line #219:
' 	Ld Selection 
' 	ArgsMemCall Clear 0x0000 
' Line #220:
' 	LitStr 0x000B "=MAX(C[14])"
' 	LitStr 0x0002 "C4"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #221:
' 	LitStr 0x0015 "=SUM(R[5]C:R[10004]C)"
' 	LitStr 0x0002 "K6"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #222:
' 	LitStr 0x0002 "K6"
' 	LitStr 0x0006 "Sheet1"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0002 "M2"
' 	LitStr 0x0003 "PRS"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemSt Value 
' Line #223:
' 	LineCont 0x0004 07 00 08 00
' 	LitStr 0x0059 "=IF(OR(IMP!R10C5="SeeYou Mobile",IMP!R[-4]C[-9]="X",R[1]C[-1]=R[1]C,R[1]C[-1]=0),"PR","")"
' 	LitStr 0x0002 "L5"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #224:
' 	LitStr 0x0015 "=SUM(R[5]C:R[10004]C)"
' 	LitStr 0x0002 "L6"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #225:
' 	LitStr 0x0002 "L6"
' 	LitStr 0x0006 "Sheet1"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0002 "M3"
' 	LitStr 0x0003 "PRS"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemSt Value 
' Line #226:
' 	LitStr 0x0002 "L5"
' 	LitStr 0x0006 "Sheet1"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0003 "A10"
' 	LitStr 0x0003 "PRS"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemSt Value 
' Line #227:
' 	LitStr 0x0009 "=PRS!R2C1"
' 	LitStr 0x0002 "P3"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #228:
' 	LitStr 0x001F "=MAX(R[7]C[-15]:R[10006]C[-15])"
' 	LitStr 0x0002 "P4"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #229:
' 	LitStr 0x0003 "A11"
' 	LitStr 0x0006 "Sheet1"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0002 "P5"
' 	LitStr 0x0006 "Sheet1"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemSt Value 
' Line #230:
' 	LineCont 0x0004 07 00 08 00
' 	LitStr 0x0083 "=IF(RC[-14]="","",IF(AND(RC[-15]>=R5C16,RC[-15]<=R4C16),TIME(RC[-15],RC[-14],RC[-13])+R3C16,TIME(RC[-15],RC[-14],RC[-13])+R3C16+1))"
' 	LitStr 0x000A "P11:P10010"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #231:
' 	LitStr 0x0022 "=IF(RC[-13]="","",RC[1]-R[-1]C[1])"
' 	LitStr 0x000A "O12:O10010"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #232:
' 	LitStr 0x0019 "=Average(R[5]C:R[10004]C)"
' 	LitStr 0x0002 "O6"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #233:
' 	QuoteRem 0x0004 0x003C "Sets 00:00:04 as interval for 5-fix vs 14-fix calcs @ Z & AL"
' Line #234:
' 	LitStr 0x0015 "0.0000462962962962963"
' 	LitStr 0x0002 "O7"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #235:
' 	LitStr 0x000E "=SECOND(R6C15)"
' 	LitStr 0x0002 "A2"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #236:
' 	LitStr 0x001B "=IF(RC[-15]="","",R[-1]C+1)"
' 	LitStr 0x000A "Q11:Q10010"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #237:
' 	LineCont 0x0004 07 00 08 00
' 	LitStr 0x006C "=IF(OR(R[1]C[-17]="",AND(RC[-14]=R[1]C[-14],RC[-13]=R[1]C[-13],RC[-11]=R[1]C[-11],RC[-10]=R[1]C[-10])),"",1)"
' 	LitStr 0x000A "R11:R10010"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #238:
' 	LitStr 0x0002 "L5"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0000 ""
' 	Eq 
' 	IfBlock 
' Line #239:
' 	LitStr 0x0029 "=IF(AND(RC[-1]=1,RC[-8]>R[-1]C[-8]),1,"")"
' 	LitStr 0x000A "S11:S10010"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #240:
' 	LitStr 0x002B "=IF(AND(RC[-4]<R4C[2],RC[-2]=""),RC[-9],"")"
' 	LitStr 0x000A "T11:T10010"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #241:
' 	LitStr 0x002F "=IF(R[2]C=0,R11C11,IF(R2C22=R4C22,R4C20,R7C24))"
' 	LitStr 0x0002 "T2"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #242:
' 	LineCont 0x0004 07 00 08 00
' 	LitStr 0x0036 "=IF(SUM(R[7]C:R[10011]C)=0,0,AVERAGE(R[7]C:R[10011]C))"
' 	LitStr 0x0002 "T4"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #243:
' 	LineCont 0x0004 07 00 08 00
' 	LitStr 0x00A0 "=IF(R[-1]C[-20]="","",IF(AND(SUM(RC[-3]:R[7]C[-3])=8,RC[-10]>R[-1]C[-10]+3,RC[-10]<R[-1]C[-10]+6),3,IF(AND(SUM(RC[-3]:R[7]C[-3])=8,RC[-10]>R[-1]C[-10]+6),6,0)))"
' 	LitStr 0x000A "U11:U10010"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #244:
' 	LitStr 0x0002 "L5"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0000 ""
' 	Ne 
' 	ElseIfBlock 
' Line #245:
' 	LitStr 0x0029 "=IF(AND(RC[-1]=1,RC[-7]>R[-1]C[-7]),1,"")"
' 	LitStr 0x000A "S11:S10010"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #246:
' 	LitStr 0x002B "=IF(AND(RC[-4]<R4C[2],RC[-2]=""),RC[-8],"")"
' 	LitStr 0x000A "T11:T10010"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #247:
' 	LitStr 0x002A "=IF(R[2]C=0,0,IF(R2C22=R4C22,R4C20,R7C24))"
' 	LitStr 0x0002 "T2"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #248:
' 	LineCont 0x0004 07 00 08 00
' 	LitStr 0x0036 "=IF(SUM(R[7]C:R[10011]C)=0,0,AVERAGE(R[7]C:R[10011]C))"
' 	LitStr 0x0002 "T4"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #249:
' 	LineCont 0x0004 07 00 08 00
' 	LitStr 0x009A "=IF(R[-1]C[-20]="","",IF(AND(SUM(RC[-3]:R[7]C[-3])=8,RC[-9]>R[-1]C[-9]+3,RC[-9]<R[-1]C[-9]+6),3,IF(AND(SUM(RC[-3]:R[7]C[-3])=8,RC[-9]>R[-1]C[-9]+6),6,0)))"
' 	LitStr 0x000A "U11:U10010"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #250:
' 	EndIfBlock 
' Line #251:
' 	QuoteRem 0x0004 0x008E "Range("V11:V10010").FormulaR1C1 = "=IF(OR(AND(RC[-4]=1,RC[-3]=1,R[1]C[-3]=1,RC[-1]=6),AND(R[1]C[-1]=3,RC[-1]=3,R[-1]C[-1]="""")),RC[-6],"""")""
' Line #252:
' 	LitStr 0x0065 "=IF(OR(AND(RC[-4]=1,RC[-3]=1,R[1]C[-3]=1,RC[-1]=6),AND(R[1]C[-1]=3,RC[-1]=3,R[-1]C[-1]=0)),RC[-6],"")"
' 	LitStr 0x000A "V11:V10010"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #253:
' 	LineCont 0x0004 07 00 08 00
' 	LitStr 0x004C "=IF(AND(R[2]C[-2]=0,R[5]C[1]=0),R[9]C[-6],IF(R[5]C[1]>R[2]C,R[5]C[1],R[2]C))"
' 	LitStr 0x0002 "V2"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #254:
' 	LitStr 0x0015 "=MIN(R[7]C:R[10011]C)"
' 	LitStr 0x0002 "V4"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #255:
' 	LitStr 0x0002 "L5"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0000 ""
' 	Eq 
' 	IfBlock 
' Line #256:
' 	LineCont 0x0004 07 00 08 00
' 	LitStr 0x009B "=IF(OR(RC[-22]="",RC[-12]<R4C[-3]+5,RC[-12]>R4C[-3]+20,RC[-6]>0.75*R4C3),"",IF(AND(AVERAGE(R[-10]C[-12]:R[-1]C[-12])<R4C[-3]+20,RC[-7]>R4C[-1]),RC[-7],""))"
' 	LitStr 0x000A "W12:W10010"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #257:
' 	LitStr 0x004C "=IF(OR(RC[-1]="",RC[-8]=R7C[-1]=FALSE),"",AVERAGE(R[-10]C[-13]:R[-1]C[-13]))"
' 	LitStr 0x000A "X11:X10010"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #258:
' 	LitStr 0x0002 "L5"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0000 ""
' 	Ne 
' 	ElseIfBlock 
' Line #259:
' 	LineCont 0x0004 07 00 08 00
' 	LitStr 0x009B "=IF(OR(RC[-22]="",RC[-11]<R4C[-3]+5,RC[-11]>R4C[-3]+20,RC[-6]>0.75*R4C3),"",IF(AND(AVERAGE(R[-10]C[-11]:R[-1]C[-11])<R4C[-3]+20,RC[-7]>R4C[-1]),RC[-7],""))"
' 	LitStr 0x000A "W12:W10010"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #260:
' 	LitStr 0x004C "=IF(OR(RC[-1]="",RC[-8]=R7C[-1]=FALSE),"",AVERAGE(R[-10]C[-12]:R[-1]C[-12]))"
' 	LitStr 0x000A "X11:X10010"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #261:
' 	EndIfBlock 
' Line #262:
' 	LitStr 0x0015 "=MAX(R[4]C:R[10011]C)"
' 	LitStr 0x0002 "W7"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #263:
' 	LitStr 0x0015 "=MAX(R[4]C:R[10011]C)"
' 	LitStr 0x0002 "X7"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #264:
' 	LitStr 0x0029 "=IF(RC[-24]="","",IF(RC[-12]>R6C25,1,""))"
' 	LitStr 0x000A "Y11:Y10010"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #265:
' 	LitStr 0x001A "='[A.xlsm]ALL CLAIMS'!R8C6"
' 	LitStr 0x0002 "Y3"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #266:
' 	LitStr 0x0009 "=PRS!R6C1"
' 	LitStr 0x0002 "Y6"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #267:
' 	QuoteRem 0x0004 0x00B5 "Range("Z25:Z10009").FormulaR1C1 = "=IF(AND(RC[-1]=1,RC[-9]<R4C3/2,OR(AND(R6C15>=0.00013,SUM(R[-4]C[-1]:RC[-1])>=5),AND(R6C15<0.00013,SUM(R[-14]C[-1]:RC[-1])>=14))),R[1]C[-10],"""")""
' Line #268:
' 	QuoteRem 0x0004 0x00C7 "'Range("Z25:Z10010").FormulaR1C1 = "=IF(AND(R[-1]C[-1]=1,RC[-9]<R4C3/2,R[1]C[-1]<>1,OR(AND(R6C15<=R7C15,SUM(R[-5]C[-1]:R[-1]C[-1])>=3),AND(R6C15>R7C15,SUM(R[-15]C[-1]:R[-1]C[-1])>=8))),RC[-10],"""")""
' Line #269:
' 	LineCont 0x0004 01 00 DD FF
' 	QuoteRem 0x0004 0x00CE "''Range("Z25:Z10009").FormulaR1C1 =       "=IF(AND(R[-1]C[-1]=1,RC[-10]>R5C28,R[1]C[-1]<>1,OR(AND(R6C15<=R7C15,SUM(R[-5]C[-1]:R[-1]C[-1])>=3),AND(R6C15>R7C15,SUM(R[-15]C[-1]:R[-1]C[-1])>=8))),RC[-10],"""")""
' Line #270:
' 	QuoteRem 0x0004 0x0084 "'''Range("Z25:Z10009").FormulaR1C1 = "=IF(AND(R[-1]C[-1]=1,RC[-10]>R5C28,R[1]C[-1]<>1,SUM(R[-15]C[-1]:R[-1]C[-1])>=8),RC[-10],"""")""
' Line #271:
' 	QuoteRem 0x0004 0x0080 "''''Range("AA25:AA10009").FormulaR1C1 = "=IF(AND(R[-1]C[-2]="""",RC[-2]=1,RC[-11]>R3C26,SUM(RC[-2]:R[15]C[-2])>8),RC[-11],"""")""
' Line #272:
' 	LitStr 0x0044 "=IF(AND(R[-1]C[-2]="",RC[-2]=1,SUM(RC[-2]:R[15]C[-2])>8),RC[-11],"")"
' 	LitStr 0x000C "AA25:AA10009"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #273:
' 	LitStr 0x0016 "=MIN(R[22]C:R[10006]C)"
' 	LitStr 0x0003 "AA3"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #274:
' 	LitStr 0x0016 "=MAX(R[21]C:R[10005]C)"
' 	LitStr 0x0003 "AA4"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #275:
' 	LitStr 0x0021 "=IF(R[-1]C<>0,(R[-1]C-R[-2]C),"")"
' 	LitStr 0x0003 "AA5"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #276:
' 	LitStr 0x0007 "AA3:AA5"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0007 "AA3:AA5"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #277:
' 	LitStr 0x004D "=IF(AND(R[-1]C[-1]=1,R[1]C[-1]<>1,SUM(R[-15]C[-1]:R[-1]C[-1])>=8),RC[-10],"")"
' 	LitStr 0x000A "Z25:Z10009"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #278:
' 	QuoteRem 0x0004 0x0093 "''''Range("Z25:Z10009").FormulaR1C1 = "=IF(AND(R[-1]C[-1]=1,RC[-10]>R5C28,RC[-10]<R4C27,R[1]C[-1]<>1,SUM(R[-15]C[-1]:R[-1]C[-1])>=8),RC[-10],"""")""
' Line #279:
' 	QuoteRem 0x0004 0x0092 "'''''Range("Z25:Z10009").FormulaR1C1 = "=IF(AND(OR(R4C27=0,R4C27<RC[-10]),R[-1]C[-1]=1,R[1]C[-1]<>1,SUM(R[-15]C[-1]:R[-1]C[-1])>=8),RC[-10],"""")""
' Line #280:
' 	QuoteRem 0x0004 0x0078 "'''''Range("Z25:Z10009").FormulaR1C1 = "=IF(AND(R[-1]C[-1]=1,R[1]C[-1]<>1,SUM(R[-15]C[-1]:R[-1]C[-1])>=8),RC[-10],"""")""
' Line #281:
' 	QuoteRem 0x0004 0x00E3 "Range("Z25").FormulaR1C1 = "=IF(AND(R4C27=0,R[-1]C[-1]=1,R[1]C[-1]<>1,SUM(R[-15]C[-1]:R[-1]C[-1])>=8),RC[-10],IF(AND(R4C27>0,RC[-10]>=R3C27,RC[-10]<R4C27,R[-1]C[-1]=1,R[1]C[-1]<>1,SUM(R[-15]C[-1]:R[-1]C[-1])>=8),RC[-10],""""))""
' Line #282:
' 	LitStr 0x002C "=IF(R1C1="NO ENL",0,IF(R[6]C=0,R[2]C,R[6]C))"
' 	LitStr 0x0002 "Z2"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #283:
' 	LitStr 0x0016 "=MIN(R[22]C:R[10006]C)"
' 	LitStr 0x0002 "Z3"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #284:
' 	LitStr 0x0016 "=MAX(R[21]C:R[10005]C)"
' 	LitStr 0x0002 "Z4"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #285:
' Line #286:
' 	LitStr 0x0003 "AA4"
' 	ArgsLd Range 0x0001 
' 	LitDI2 0x0000 
' 	Ne 
' 	LitStr 0x0003 "AA4"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0002 "Z4"
' 	ArgsLd Range 0x0001 
' 	Lt 
' 	And 
' 	LitStr 0x0003 "AA5"
' 	ArgsLd Range 0x0001 
' 	LitR8 0x2BEC 0x354C 0xC16C 0x3F56 
' 	Ge 
' 	And 
' 	IfBlock 
' Line #287:
' 	LitStr 0x006A "=IF(AND(RC[-10]<R4C27,RC[-10]>=R3C27,R[-1]C[-1]=1,R[1]C[-1]<>1,SUM(R[-15]C[-1]:R[-1]C[-1])>=8),RC[-10],"")"
' 	LitStr 0x000A "Z25:Z10009"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #288:
' 	EndIfBlock 
' Line #289:
' 	QuoteRem 0x0004 0x0063 "Range("Z2").FormulaR1C1 = "=IF(OR(R[1]C[-1]=2,R[1]C[-1]=4,R1C1=""NO ENL""),0,MAX(R[9]C:R[10008]C))""
' Line #290:
' 	QuoteRem 0x0004 0x0076 "'Range("Z2").FormulaR1C1 = "=IF(OR(R[1]C[-1]=2,R[1]C[-1]=4,R1C1=""NO ENL""),0,IF(R[6]C=0,MAX(R[9]C:R[10008]C),R[6]C))""
' Line #291:
' 	QuoteRem 0x0004 0x005B "''Range("Z2").FormulaR1C1 = "=IF(R1C1=""NO ENL"",0,IF(R[6]C=0,MIN(R[9]C:R[10008]C),R[6]C))""
' Line #292:
' 	QuoteRem 0x0004 0x004D "'''Range("Z2").FormulaR1C1 = "=IF(R1C1=""NO ENL"",0,IF(R[6]C=0,R[1]C,R[6]C))""
' Line #293:
' 	QuoteRem 0x0004 0x004E "''''Range("Z2").FormulaR1C1 = "=IF(R1C1=""NO ENL"",0,IF(R[6]C=0,R[2]C,R[6]C))""
' Line #294:
' 	QuoteRem 0x0004 0x0034 "''Range("Z3").FormulaR1C1 = "=MIN(R[22]C:R[10006]C)""
' Line #295:
' 	QuoteRem 0x0004 0x0054 "''''Range("Z3").FormulaR1C1 = "=IF(R[1]C[12]<>""NONE"",MIN(R[22]C:R[10006]C),R[1]C)""
' Line #296:
' 	QuoteRem 0x0004 0x0036 "''''Range("Z4").FormulaR1C1 = "=MAX(R[21]C:R[10005]C)""
' Line #297:
' 	QuoteRem 0x0004 0x0034 "'Range("AA3").FormulaR1C1 = "=MIN(R[22]C:R[10006]C)""
' Line #298:
' 	QuoteRem 0x0004 0x004F "''Range("AA3").FormulaR1C1 = "=IF(OR(R3C25=2,R3C25=4),0,MIN(R[22]C:R[10006]C))""
' Line #299:
' 	QuoteRem 0x0004 0x0077 "'''Range("AA3").FormulaR1C1 = "=IF(OR(R3C25=2,R3C25=4,MIN(R[22]C:R[10006]C)-RC[-1]<=0.000694),0,MIN(R[22]C:R[10006]C))""
' Line #300:
' 	QuoteRem 0x0004 0x0036 "'''Range("AA4").FormulaR1C1 = "=MAX(R[21]C:R[10005]C)""
' Line #301:
' 	QuoteRem 0x0004 0x0032 "'''Range("AA3:AA4").Value = Range("AA3:AA4").Value"
' Line #302:
' Line #303:
' 	LitStr 0x0028 "=IF(R[2]C=R3C16+PRS!R2C2/24,R[1]C,R[2]C)"
' 	LitStr 0x0003 "AB2"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #304:
' 	LitStr 0x002E "=IF(R[-1]C[-8]=0,R[-1]C[-6],MAX(RC[-2],R[2]C))"
' 	LitStr 0x0003 "AB3"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #305:
' Line #306:
' 	LitStr 0x0002 "L5"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0000 ""
' 	Eq 
' 	IfBlock 
' Line #307:
' 	LineCont 0x0004 07 00 08 00
' 	LitStr 0x00CE "=IF(R[3]C[-12]="","",IF(AND(RC[-12]>R2C22,OR(AND(RC[-9]=1,R[1]C[-9]="",RC[-7]>=3,SUM(R[1]C[-7]:R[3]C[-7])=0),(AND(RC[-9]=1,SUM(R[1]C[-9]:R[4]C[-9])=0)),(AND(RC[-7]>=3,RC[-17]-R[3]C[-17]>=30)))),RC[-12],""))"
' 	LitStr 0x000C "AB11:AB10010"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #308:
' 	LitStr 0x001F "=IF(RC[-13]=R3C[-1],RC[-18],"")"
' 	LitStr 0x000C "AC11:AC10010"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #309:
' 	LitStr 0x0002 "L5"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0000 ""
' 	Ne 
' 	ElseIfBlock 
' Line #310:
' 	LineCont 0x0004 07 00 08 00
' 	LitStr 0x00EA "=IF(OR(RC[-16]=0,R[3]C[-12]=""),"",IF(AND(RC[-12]>R2C22,OR(AND(RC[-9]=1,R[1]C[-9]="",RC[-7]>=3,SUM(R[1]C[-7]:R[3]C[-7])=0),(AND(RC[-9]=1,SUM(R[1]C[-9]:R[4]C[-9])=0)),(AND(R[3]C[-16]<>0,RC[-7]>=3,RC[-16]-R[3]C[-16]>=30)))),RC[-12],""))"
' 	LitStr 0x000C "AB11:AB10010"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #311:
' 	LitStr 0x001F "=IF(RC[-13]=R3C[-1],RC[-17],"")"
' 	LitStr 0x000C "AC11:AC10010"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #312:
' 	EndIfBlock 
' Line #313:
' 	LitStr 0x000B "O11:AB10010"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x000B "O11:AB10010"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #314:
' 	LitStr 0x0028 "=IF(R[2]C=R3C16+PRS!R2C2/24,R[1]C,R[2]C)"
' 	LitStr 0x0003 "AB2"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #315:
' 	LitStr 0x0032 "=IF(R[-1]C[-8]=0,R[-1]C[-6],MAX(R[-1]C[-2],R[2]C))"
' 	LitStr 0x0003 "AB3"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #316:
' 	LitStr 0x0033 "='[A.xlsm]DATA ENTRY CHECK'!R12C7+R3C16-PRS!R2C2/24"
' 	LitStr 0x0003 "AB4"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #317:
' 	LitStr 0x0014 "=MIN(R[6]C:R[9995]C)"
' 	LitStr 0x0003 "AB5"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #318:
' 	LitStr 0x0014 "=MAX(R[5]C:R[9994]C)"
' 	LitStr 0x0003 "AC6"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #319:
' 	LitStr 0x0029 "=IF(OR(RC[-14]=R2C28,RC[2]=1),RC[-26],"")"
' 	LitStr 0x000C "AD11:AD10010"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #320:
' 	QuoteRem 0x0004 0x0032 "Range("AD2").FormulaR1C1 = "=MAX(R[8]C:R[10009]C)""
' Line #321:
' 	LitStr 0x0027 "=IF(R[6]C=0,MAX(R[8]C:R[10009]C),R[6]C)"
' 	LitStr 0x0003 "AD2"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #322:
' 	LitStr 0x0019 "=IF(RC[-1]="","",RC[-26])"
' 	LitStr 0x000C "AE11:AE10010"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #323:
' 	QuoteRem 0x0004 0x0038 "Range("AE2").FormulaR1C1 = "=MAX(R[8]C:R[10009]C)*0.001""
' Line #324:
' 	LitStr 0x002D "=IF(R[6]C=0,MAX(R[8]C:R[10009]C)*0.001,R[6]C)"
' 	LitStr 0x0003 "AE2"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #325:
' 	LitStr 0x004A "=IF(AND(PRS!R14C4>0,R4C28>R3C16,R4C[-4]>R[-1]C[-16],R4C[-4]<RC[-16]),1,"")"
' 	LitStr 0x000C "AF11:AF10010"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #326:
' 	LitStr 0x0019 "=IF(RC[-2]="","",RC[-26])"
' 	LitStr 0x000C "AG11:AG10010"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #327:
' 	QuoteRem 0x0004 0x0032 "Range("AG2").FormulaR1C1 = "=MAX(R[8]C:R[10009]C)""
' Line #328:
' 	LitStr 0x0027 "=IF(R[6]C=0,MAX(R[8]C:R[10009]C),R[6]C)"
' 	LitStr 0x0003 "AG2"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #329:
' 	LitStr 0x0019 "=IF(RC[-1]="","",RC[-26])"
' 	LitStr 0x000C "AH11:AH10010"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #330:
' 	QuoteRem 0x0004 0x0038 "Range("AH2").FormulaR1C1 = "=MAX(R[8]C:R[10009]C)*0.001""
' Line #331:
' 	LitStr 0x002D "=IF(R[6]C=0,MAX(R[8]C:R[10009]C)*0.001,R[6]C)"
' 	LitStr 0x0003 "AH2"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #332:
' 	LitStr 0x0002 "L5"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0000 ""
' 	Eq 
' 	IfBlock 
' Line #333:
' 	LitStr 0x0019 "=IF(RC[-2]="","",RC[-25])"
' 	LitStr 0x000C "AJ11:AJ10010"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #334:
' 	LitStr 0x0002 "L5"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0000 ""
' 	Ne 
' 	ElseIfBlock 
' Line #335:
' 	LitStr 0x0019 "=IF(RC[-2]="","",RC[-24])"
' 	LitStr 0x000C "AJ11:AJ10010"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #336:
' 	EndIfBlock 
' Line #337:
' 	QuoteRem 0x0004 0x0046 "Range("AJ2").FormulaR1C1 = "=IF(R4C[-8]=0,R6C29,MAX(R[9]C:R[10008]C))""
' Line #338:
' 	LitStr 0x003B "=IF(R4C[-8]=0,R6C29,IF(R[6]C=0,MAX(R[9]C:R[10008]C),R[6]C))"
' 	LitStr 0x0003 "AJ2"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #339:
' 	QuoteRem 0x0004 0x00AB "Range("AL11:AL10010").FormulaR1C1 = "=IF(AND(RC[-22]>R2C28,OR(AND(R6C15>=0.00013,SUM(RC[-13]:R[4]C[-13])=5),AND(R6C15<0.00013,SUM(RC[-13]:R[14]C[-13])=15))),RC[-22],"""")""
' Line #340:
' 	QuoteRem 0x0004 0x00DE "'Range("AL11:AL10010").FormulaR1C1 = "=IF(AND(RC[-22]>R2C28,RC[-13]<>1,R[1]C[-13]=1,OR(AND(R6C15<=R7C15,SUM(R[2]C[-13]:R[6]C[-13])>=3,RC[-13]=0),AND(R6C15>R7C15,SUM(R[2]C[-13]:R[15]C[-13])>=8,R[1]C[-13]=1))),RC[-22],"""")""
' Line #341:
' 	LitStr 0x00CF "=IF(RC[-11]=R3C27,RC[-22],IF(AND(RC[-22]>R2C28,RC[-13]<>1,R[1]C[-13]=1,OR(AND(R6C15<=R7C15,SUM(R[2]C[-13]:R[6]C[-13])>=3,RC[-13]=0),AND(R6C15>R7C15,SUM(R[2]C[-13]:R[15]C[-13])>=8,R[1]C[-13]=1))),RC[-22],""))"
' 	LitStr 0x000C "AL11:AL10010"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #342:
' 	LitStr 0x0006 "=R4C38"
' 	LitStr 0x0003 "AL2"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #343:
' 	QuoteRem 0x0004 0x0075 "''Range("AL4").FormulaR1C1 = "=IF(OR(R[-1]C[-13]=2,R[3]C[1]=0,MIN(R[7]C:R[10008]C)=0),""NONE"",MIN(R[7]C:R[10008]C))""
' Line #344:
' 	QuoteRem 0x0004 0x0092 "''Range("AL4").FormulaR1C1 = "=IF(OR(R[-1]C[-13]=2,R[3]C[1]=0,MIN(R[7]C:R[10008]C)<RC[-12],MIN(R[7]C:R[10008]C)=0),""NONE"",MIN(R[7]C:R[10008]C))""
' Line #345:
' 	QuoteRem 0x0004 0x0075 "''Range("AL4").FormulaR1C1 = "=IF(OR(R[-1]C[-13]=2,R[3]C[1]=0,MIN(R[7]C:R[10008]C)=0),""NONE"",MIN(R[7]C:R[10008]C))""
' Line #346:
' 	QuoteRem 0x0004 0x007B "'''Range("AL4").FormulaR1C1 = "=IF(OR(R[-1]C[-13]=2,R[-1]C[-13]=4,R[3]C[1]=0,R[-1]C[-11]=0),""NONE"",MIN(R[7]C:R[10008]C))""
' Line #347:
' 	QuoteRem 0x0004 0x008C "''''Range("AL4").FormulaR1C1 = "=IF(OR(R[-1]C[-13]=2,R[-1]C[-13]=4,R[3]C[1]=0,R[-1]C[-11]=0,RC[-11]<RC[-12]),""NONE"",MIN(R[7]C:R[10008]C))""
' Line #348:
' 	LitStr 0x005C "=IF(OR(R[-1]C[-13]=2,R[-1]C[-13]=4,R[3]C[1]=0,R[-1]C[-11]=0,RC[-11]<RC[-12]),"NONE",RC[-11])"
' 	LitStr 0x0003 "AL4"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #349:
' 	LitStr 0x002B "=IF(R[-1]C<>"NONE",MAX(R[20]C:R[10004]C),0)"
' 	LitStr 0x0003 "AL5"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #350:
' 	LitStr 0x0026 "=IF(R[-3]C="NONE","NONE",R2C47-R[-3]C)"
' 	LitStr 0x0003 "AL7"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #351:
' 	LitStr 0x0002 "L5"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0000 ""
' 	Eq 
' 	IfBlock 
' Line #352:
' 	QuoteRem 0x0004 0x005F "Range("AM11:AM10010").FormulaR1C1 = "=IF(OR(RC[-1]="""",RC[-20]="""",R[-1]C[-28]>RC[-28]),0,1)""
' Line #353:
' 	LitStr 0x0034 "=IF(AND(RC[-1]="",RC[-20]="",R[-1]C[-26]<R6C25),0,1)"
' 	LitStr 0x000C "AM11:AM10010"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #354:
' 	QuoteRem 0x0004 0x005E "''Range("AO11:AO10010").FormulaR1C1 = "=IF(R4C38=""NONE"","""",IF(RC[-3]=R4C40,RC[-30],""""))""
' Line #355:
' 	LitStr 0x0031 "=IF(R4C38="NONE","",IF(RC[-25]=R4C40,RC[-30],""))"
' 	LitStr 0x000C "AO11:AO10009"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #356:
' 	LitStr 0x0002 "L5"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0000 ""
' 	Ne 
' 	ElseIfBlock 
' Line #357:
' 	QuoteRem 0x0004 0x005F "Range("AM11:AM10010").FormulaR1C1 = "=IF(OR(RC[-1]="""",RC[-20]="""",R[-1]C[-27]>RC[-27]),0,1)""
' Line #358:
' 	LitStr 0x0034 "=IF(AND(RC[-1]="",RC[-20]="",R[-1]C[-26]<R6C25),0,1)"
' 	LitStr 0x000C "AM11:AM10010"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #359:
' 	QuoteRem 0x0004 0x0060 "'''Range("AO11:AO10010").FormulaR1C1 = _"=IF(R4C38=""NONE"","""",IF(RC[-3]=R4C40,RC[-29],""""))""
' Line #360:
' 	LitStr 0x0031 "=IF(R4C38="NONE","",IF(RC[-25]=R4C40,RC[-29],""))"
' 	LitStr 0x000C "AO11:AO10009"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #361:
' 	EndIfBlock 
' Line #362:
' 	LitStr 0x0015 "=SUM(R[4]C:R[10005]C)"
' 	LitStr 0x0003 "AM7"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #363:
' 	LineCont 0x0004 07 00 08 00
' 	LitStr 0x0072 "=IF(OR(RC[-39]="",RC[-23]<R4C3-0.4*R4C3),"",IF(AND(RC[-24]>R2C28,RC[-22]="",R[1]C[-22]="",RC[-21]=""),RC[-24],""))"
' 	LitStr 0x000C "AN11:AN10010"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #364:
' 	QuoteRem 0x0004 0x0095 "'Range("AN11:AN10010").FormulaR1C1 = "=IF(OR(RC[-39]="""",RC[-23]<R4C3-0.4*R4C3),"""",IF(AND(RC[-24]>R2C28,RC[-22]="""",RC[-21]=""""),RC[-24],""""))""
' Line #365:
' 	QuoteRem 0x0004 0x0023 "Range("AN2").FormulaR1C1 = "=R[2]C""
' Line #366:
' 	QuoteRem 0x0004 0x0038 "'''Range("AN2").FormulaR1C1 = "=IF(R[6]C=0,R[2]C,R[6]C)""
' Line #367:
' 	LitStr 0x0018 "=IF(R[6]C=0,R[2]C,R[6]C)"
' 	LitStr 0x0003 "AN2"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #368:
' 	QuoteRem 0x0004 0x0073 "''Range("AN4").FormulaR1C1 = "=IF(AND(OR(R[-1]C[-15]=3,R[-1]C[-15]=5),R[3]C[-1]>0),MIN(R[7]C[-2]:R[10008]C[-2]),0)""
' Line #369:
' 	QuoteRem 0x0004 0x0085 "'''Range("AN4").FormulaR1C1 = "=IF(AND(OR(R[-1]C[-15]=3,R[-1]C[-15]=5),RC[-2]<>""NONE"",R[3]C[-1]>0),MIN(R[7]C[-2]:R[10008]C[-2]),0)""
' Line #370:
' 	LitStr 0x004D "=IF(AND(OR(R[-1]C[-15]=3,R[-1]C[-15]=5),RC[-2]<>"NONE",R[3]C[-1]>0),RC[-2],0)"
' 	LitStr 0x0003 "AN4"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #371:
' 	QuoteRem 0x0004 0x0046 "Range("AO2").FormulaR1C1 = "=IF(R2C40="""","""",MAX(R[9]C:R[10010]C))""
' Line #372:
' 	LitStr 0x0037 "=IF(R2C40="","",IF(R[6]C=0,MAX(R[9]C:R[10010]C),R[6]C))"
' 	LitStr 0x0003 "AO2"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #373:
' 	LitStr 0x002D "=IF(R4C40="","",IF(RC[-26]=R4C40,RC[-38],""))"
' 	LitStr 0x000C "AP11:AP10010"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #374:
' 	QuoteRem 0x0004 0x0032 "Range("AP2").FormulaR1C1 = "=MAX(R[8]C:R[10010]C)""
' Line #375:
' 	LitStr 0x0027 "=IF(R[6]C=0,MAX(R[8]C:R[10010]C),R[6]C)"
' 	LitStr 0x0003 "AP2"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #376:
' 	LitStr 0x0019 "=IF(RC[-1]="","",RC[-38])"
' 	LitStr 0x000C "AQ11:AQ10010"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #377:
' 	QuoteRem 0x0004 0x0038 "Range("AQ2").FormulaR1C1 = "=MAX(R[8]C:R[10010]C)*0.001""
' Line #378:
' 	LitStr 0x002D "=IF(R[6]C=0,MAX(R[8]C:R[10010]C)*0.001,R[6]C)"
' 	LitStr 0x0003 "AQ2"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #379:
' 	LitStr 0x0019 "=IF(RC[-1]="","",RC[-37])"
' 	LitStr 0x000C "AR11:AR10010"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #380:
' 	QuoteRem 0x0004 0x0032 "Range("AR2").FormulaR1C1 = "=MAX(R[8]C:R[10010]C)""
' Line #381:
' 	LitStr 0x0027 "=IF(R[6]C=0,MAX(R[8]C:R[10010]C),R[6]C)"
' 	LitStr 0x0003 "AR2"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #382:
' 	LitStr 0x0019 "=IF(RC[-1]="","",RC[-37])"
' 	LitStr 0x000C "AS11:AS10010"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #383:
' 	QuoteRem 0x0004 0x0038 "Range("AS2").FormulaR1C1 = "=MAX(R[8]C:R[10010]C)*0.001""
' Line #384:
' 	LitStr 0x002D "=IF(R[6]C=0,MAX(R[8]C:R[10010]C)*0.001,R[6]C)"
' 	LitStr 0x0003 "AS2"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #385:
' 	LitStr 0x0002 "L5"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0000 ""
' 	Eq 
' 	IfBlock 
' Line #386:
' 	LineCont 0x0004 07 00 08 00
' 	LitStr 0x0032 "=IF(OR(RC[-6]="",RC[-35]>PRS!R7C4+152),"",RC[-35])"
' 	LitStr 0x000C "AT11:AT10010"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #387:
' 	LineCont 0x0004 07 00 08 00
' 	LitStr 0x00C0 "=IF(RC[-30]=R4C[-44],RC[-31],IF(RC[-1]="","",IF(AND(RC[-29]="",RC[-28]="",OR(RC[-36]=R9C[-1],AND(RC[-1]=R[1]C[-1],R[1]C[-1]=R[2]C[-1]),AND(RC[-36]<=R9C[-1]+5,RC[-36]>=R9C[-1]-5))),RC[-7],"")))"
' 	LitStr 0x000C "AU11:AU10010"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #388:
' 	LitStr 0x0002 "L5"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0000 ""
' 	Ne 
' 	ElseIfBlock 
' Line #389:
' 	LineCont 0x0004 07 00 08 00
' 	LitStr 0x003C "=IF(OR(RC[-6]="",RC[-34]>PRS!R7C4+152,RC[-34]=0),"",RC[-34])"
' 	LitStr 0x000C "AT11:AT10010"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #390:
' 	LineCont 0x0004 07 00 08 00
' 	LitStr 0x00C0 "=IF(RC[-30]=R4C[-44],RC[-31],IF(RC[-1]="","",IF(AND(RC[-29]="",RC[-28]="",OR(RC[-35]=R9C[-1],AND(RC[-1]=R[1]C[-1],R[1]C[-1]=R[2]C[-1]),AND(RC[-35]<=R9C[-1]+5,RC[-35]>=R9C[-1]-5))),RC[-7],"")))"
' 	LitStr 0x000C "AU11:AU10010"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #391:
' 	EndIfBlock 
' Line #392:
' 	LitStr 0x0019 "=AVERAGE(R[2]C:R[10003]C)"
' 	LitStr 0x0003 "AT9"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #393:
' 	LitStr 0x0006 "=R[2]C"
' 	LitStr 0x0003 "AU2"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #394:
' 	LitStr 0x0015 "=MIN(R[7]C:R[10008]C)"
' 	LitStr 0x0003 "AU4"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #395:
' 	LitStr 0x0038 "=IF(R[-1]C=MAX(R[6]C[-31]:R[9995]C[-31]),"Last DP","NO")"
' 	LitStr 0x0003 "AU5"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #396:
' 	LitStr 0x001D "=IF(RC[-32]=R2C47,RC[-44],"")"
' 	LitStr 0x000C "AV11:AV10010"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #397:
' 	LitStr 0x0015 "=MAX(R[8]C:R[10010]C)"
' 	LitStr 0x0003 "AV2"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #398:
' 	LitStr 0x0019 "=IF(RC[-1]="","",RC[-44])"
' 	LitStr 0x000C "AW11:AW10010"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #399:
' 	LitStr 0x001B "=MAX(R[8]C:R[10010]C)*0.001"
' 	LitStr 0x0003 "AW2"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #400:
' 	LitStr 0x0019 "=IF(RC[-2]="","",RC[-44])"
' 	LitStr 0x000C "AY11:AY10010"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #401:
' 	LitStr 0x0015 "=MAX(R[8]C:R[10010]C)"
' 	LitStr 0x0003 "AY2"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #402:
' 	LitStr 0x0019 "=IF(RC[-1]="","",RC[-44])"
' 	LitStr 0x000C "AZ11:AZ10010"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #403:
' 	LitStr 0x001B "=MAX(R[8]C:R[10010]C)*0.001"
' 	LitStr 0x0003 "AZ2"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #404:
' 	LitStr 0x0002 "L5"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0000 ""
' 	Eq 
' 	IfBlock 
' Line #405:
' 	LineCont 0x0004 07 00 08 00
' 	LitStr 0x0064 "=IF(AND(RC[-7]=R2C47,R5C47="Last DP"),RC[-43],IF(OR(RC[-14]="",RC[-38]<R2C47,RC[-36]=1),"",RC[-43]))"
' 	LitStr 0x000C "BB11:BB10010"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #406:
' 	LitStr 0x0002 "L5"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0000 ""
' 	Ne 
' 	ElseIfBlock 
' Line #407:
' 	LineCont 0x0004 07 00 08 00
' 	LitStr 0x006E "=IF(AND(RC[-7]=R2C47,R5C47="Last DP"),RC[-42],IF(OR(RC[-14]="",RC[-38]<R2C47,RC[-36]=1,RC[-42]=0),"",RC[-42]))"
' 	LitStr 0x000C "BB11:BB10010"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #408:
' 	EndIfBlock 
' Line #409:
' 	LitStr 0x0006 "=R[2]C"
' 	LitStr 0x0003 "BB2"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #410:
' 	LitStr 0x0019 "=AVERAGE(R[7]C:R[10008]C)"
' 	LitStr 0x0003 "BB4"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #411:
' 	LitStr 0x0039 "=IF(AND(RC[-64]=""=FALSE,RC[-68]<R2C[-62]),R[2]C[-68],"")"
' 	LitStr 0x000B "CF11:CF2510"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #412:
' 	LitStr 0x0015 "=MAX(R[7]C:R[10006]C)"
' 	LitStr 0x0003 "CF4"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #413:
' 	LitStr 0x0018 "=IF(R[2]C=0,R2C22,R[2]C)"
' 	LitStr 0x0003 "CF2"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #414:
' 	LitStr 0x001F "=IF(RC[-69]=R2C[-1],RC[-81],"")"
' 	LitStr 0x000B "CG11:CG2510"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #415:
' 	LitStr 0x001E "=IF(RC[-1]="","",RC[-81]/1000)"
' 	LitStr 0x000B "CH11:CH2510"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #416:
' 	LitStr 0x0019 "=IF(RC[-2]="","",RC[-81])"
' 	LitStr 0x000B "CJ11:CJ2510"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #417:
' 	LitStr 0x001E "=IF(RC[-1]="","",RC[-81]/1000)"
' 	LitStr 0x000B "CK11:CK2510"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #418:
' 	LitStr 0x0015 "=MAX(R[9]C:R[10008]C)"
' 	LitStr 0x0003 "CG2"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #419:
' 	LitStr 0x0015 "=MAX(R[9]C:R[10008]C)"
' 	LitStr 0x0003 "CH2"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #420:
' 	LitStr 0x0015 "=MAX(R[9]C:R[10008]C)"
' 	LitStr 0x0003 "CJ2"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #421:
' 	LitStr 0x0015 "=MAX(R[9]C:R[10008]C)"
' 	LitStr 0x0003 "CK2"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #422:
' 	LitStr 0x0002 "A1"
' 	ArgsLd Range 0x0001 
' 	ArgsMemCall Select 0x0000 
' Line #423:
' Line #424:
' 	LitStr 0x0003 "B10"
' 	LitStr 0x0003 "PRS"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0002 "C3"
' 	LitStr 0x0006 "Sheet1"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemSt Value 
' Line #425:
' 	LitStr 0x0002 "C3"
' 	ArgsLd Range 0x0001 
' 	LitDI2 0x2710 
' 	Gt 
' 	IfBlock 
' Line #426:
' 	LitStr 0x0011 "Ab.xlsm!RefineLDG"
' 	Ld Application 
' 	ArgsMemCall Run 0x0001 
' Line #427:
' 	EndIfBlock 
' Line #428:
' Line #429:
' 	LitStr 0x0002 "Y3"
' 	ArgsLd Range 0x0001 
' 	LitDI2 0x0003 
' 	Eq 
' 	LitStr 0x0002 "C4"
' 	ArgsLd Range 0x0001 
' 	LitDI2 0x2710 
' 	Gt 
' 	And 
' 	IfBlock 
' Line #430:
' 	LitStr 0x0011 "Ab.xlsm!ENLrefine"
' 	Ld Application 
' 	ArgsMemCall Run 0x0001 
' Line #431:
' 	EndIfBlock 
' Line #432:
' 	LitStr 0x0013 "Ab.xlsm!NewBRecords"
' 	Ld Application 
' 	ArgsMemCall Run 0x0001 
' Line #433:
' 	EndSub 
' Line #434:
' 	FuncDefn (Sub RefineLDG())
' Line #435:
' 	QuoteRem 0x0000 0x0000 ""
' Line #436:
' 	QuoteRem 0x0000 0x002A " Revise Landing if over 10K fixes  7/17/18"
' Line #437:
' 	QuoteRem 0x0000 0x0000 ""
' Line #438:
' 	LitStr 0x0002 "BR"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemCall Activate 0x0000 
' Line #439:
' 	LitStr 0x0009 "J1:K60000"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x000B "AT1:AU60000"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #440:
' 	LitStr 0x001F "=MAX(RC[-2]:R[60000]C[-2])-1/24"
' 	LitStr 0x0003 "AV1"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #441:
' 	LitStr 0x0019 "=IF(RC[-2]>R1C,RC[-2],"")"
' 	LitStr 0x000B "AV2:AV60001"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #442:
' 	LitStr 0x0019 "=IF(RC[-1]<>"",RC[-2],"")"
' 	LitStr 0x000B "AW2:AW60001"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #443:
' Line #444:
' 	LitStr 0x000B "AV2:AW60001"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x000B "AV2:AW60001"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #445:
' 	LitStr 0x0002 "BR"
' 	Ld ENLrefine 
' 	ArgsMemLd Worksheets 0x0001 
' 	MemLd Sort 
' 	MemLd ActiveWorkbook 
' 	ArgsMemCall Clear 0x0000 
' Line #446:
' 	LineCont 0x0004 12 00 08 00
' 	LitStr 0x000B "AV2:AV60001"
' 	ArgsLd Range 0x0001 
' 	ParamNamed Add 
' 	Ld SortOn 
' 	ParamNamed Key 
' 	Ld xlAscending 
' 	ParamNamed xlSortOnValues 
' 	Ld xlSortNormal 
' 	ParamNamed Order 
' 	LitStr 0x0002 "BR"
' 	Ld ENLrefine 
' 	ArgsMemLd Worksheets 0x0001 
' 	MemLd Sort 
' 	MemLd ActiveWorkbook 
' 	ArgsMemCall SortFields 0x0004 
' Line #447:
' 	StartWithExpr 
' 	LitStr 0x0002 "BR"
' 	Ld ENLrefine 
' 	ArgsMemLd Worksheets 0x0001 
' 	MemLd Sort 
' 	With 
' Line #448:
' 	LitStr 0x000B "AV2:AW60001"
' 	ArgsLd Range 0x0001 
' 	ArgsMemCallWith DataOption 0x0001 
' Line #449:
' 	Ld xlGuess 
' 	MemStWith Header 
' Line #450:
' 	LitVarSpecial (False)
' 	MemStWith MatchCase 
' Line #451:
' 	Ld xlTopToBottom 
' 	MemStWith Orientation 
' Line #452:
' 	Ld SortMethod 
' 	MemStWith SetRange 
' Line #453:
' 	ArgsMemCallWith xlPinYin 0x0000 
' Line #454:
' 	EndWith 
' Line #455:
' Line #456:
' 	LitStr 0x000A "AW2:AW3601"
' 	ArgsLd Range 0x0001 
' 	ArgsMemCall Select 0x0000 
' Line #457:
' 	LineCont 0x000C 0E 00 08 00 2F 00 08 00 5C 00 08 00
' 	LitStr 0x0003 "AW2"
' 	ArgsLd Range 0x0001 
' 	ParamNamed Destination 
' 	Ld xlFixedWidth 
' 	ParamNamed DataType 
' 	LitStr 0x0001 "E"
' 	ParamNamed OtherChar 
' 	LitDI2 0x0000 
' 	LitDI2 0x0009 
' 	ArgsArray Array 0x0002 
' 	LitDI2 0x0006 
' 	LitDI2 0x0001 
' 	ArgsArray Array 0x0002 
' 	LitDI2 0x0008 
' 	LitDI2 0x0001 
' 	ArgsArray Array 0x0002 
' 	LitDI2 0x000D 
' 	LitDI2 0x0001 
' 	ArgsArray Array 0x0002 
' 	LitDI2 0x000E 
' 	LitDI2 0x0001 
' 	ArgsArray Array 0x0002 
' 	LitDI2 0x0011 
' 	LitDI2 0x0001 
' 	ArgsArray Array 0x0002 
' 	LitDI2 0x0016 
' 	LitDI2 0x0001 
' 	ArgsArray Array 0x0002 
' 	LitDI2 0x0017 
' 	LitDI2 0x0009 
' 	ArgsArray Array 0x0002 
' 	LitDI2 0x0018 
' 	LitDI2 0x0001 
' 	ArgsArray Array 0x0002 
' 	LitDI2 0x001D 
' 	LitDI2 0x0009 
' 	ArgsArray Array 0x0002 
' 	LitDI2 0x0021 
' 	LitDI2 0x0009 
' 	ArgsArray Array 0x0002 
' 	ArgsArray Array 0x000B 
' 	ParamNamed FieldInfo 
' 	LitVarSpecial (True)
' 	ParamNamed TrailingMinusNumbers 
' 	Ld Selection 
' 	ArgsMemCall TextToColumns 0x0005 
' Line #458:
' Line #459:
' 	LitStr 0x0014 "=MIN(R[1]C:R[3600]C)"
' 	LitStr 0x0003 "BD1"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #460:
' 	LitStr 0x0014 "=MIN(R[1]C:R[3598]C)"
' 	LitStr 0x0007 "BD2:BI2"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #461:
' Line #462:
' 	LitStr 0x006E "=IF(AND(RC[-7]=R[1]C[-7],RC[-6]=R[1]C[-6],RC[-4]=R[1]C[-4],RC[-3]=R[1]C[-3],ABS(RC[-1]-PRS!R8C4)<4),RC[-8],"")"
' 	LitStr 0x0003 "BD3"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #463:
' 	LitStr 0x0035 "=IF(RC[-1]<>R1C56,"",IF(RC[-6]="S",-1*RC[-8],RC[-8]))"
' 	LitStr 0x0003 "BE3"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #464:
' 	LitStr 0x003D "=IF(RC[-1]="","",IF(RC[-7]="S",-1*(RC[-8]/1000),RC[-8]/1000))"
' 	LitStr 0x0003 "BF3"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #465:
' 	LitStr 0x0031 "=IF(RC[-1]="","",IF(RC[-5]="W",-1*RC[-7],RC[-7]))"
' 	LitStr 0x0003 "BG3"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #466:
' 	LitStr 0x003D "=IF(RC[-1]="","",IF(RC[-6]="W",-1*(RC[-7]/1000),RC[-7]/1000))"
' 	LitStr 0x0003 "BH3"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #467:
' 	LitStr 0x0018 "=IF(RC[-1]="","",RC[-6])"
' 	LitStr 0x0003 "BI3"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #468:
' 	QuoteRem 0x0004 0x000B "Copy Ref AV"
' Line #469:
' 	Ld xlCalculationManual 
' 	Ld Application 
' 	MemSt Calculation 
' Line #470:
' 	StartWithExpr 
' 	LitStr 0x0002 "BR"
' 	ArgsLd Worksheets 0x0001 
' 	With 
' Line #471:
' 	Ld xlUp 
' 	Ld Rows 
' 	MemLd Count 
' 	LitStr 0x0002 "AV"
' 	ArgsMemLdWith Cells 0x0002 
' 	ArgsMemLd End 0x0001 
' 	MemLd Row 
' 	St LastRow 
' Line #472:
' 	LitStr 0x0006 "BD3:BI"
' 	Ld LastRow 
' 	Concat 
' 	ArgsMemLdWith Range 0x0001 
' 	ParamNamed Destination 
' 	Ld xlFillDefault 
' 	ParamNamed Type 
' 	LitStr 0x0007 "BD3:BI3"
' 	ArgsMemLdWith Range 0x0001 
' 	ArgsMemCall AutoFill 0x0002 
' Line #473:
' 	EndWith 
' Line #474:
' 	Ld xlCalculationAutomatic 
' 	Ld Application 
' 	MemSt Calculation 
' Line #475:
' Line #476:
' 	LitStr 0x0006 "Sheet1"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemCall Activate 0x0000 
' Line #477:
' 	LitStr 0x0003 "AU4"
' 	LitStr 0x0006 "Sheet1"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	LitStr 0x0003 "BD2"
' 	LitStr 0x0002 "BR"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	Ne 
' 	IfBlock 
' Line #478:
' 	LitStr 0x0003 "BD2"
' 	LitStr 0x0002 "BR"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0003 "AU4"
' 	LitStr 0x0006 "Sheet1"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemSt Value 
' Line #479:
' 	LitStr 0x0003 "BE2"
' 	LitStr 0x0002 "BR"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0003 "AV2"
' 	LitStr 0x0006 "Sheet1"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemSt Value 
' Line #480:
' 	LitStr 0x0003 "BF2"
' 	LitStr 0x0002 "BR"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0003 "AW2"
' 	LitStr 0x0006 "Sheet1"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemSt Value 
' Line #481:
' 	LitStr 0x0003 "BG2"
' 	LitStr 0x0002 "BR"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0003 "AY2"
' 	LitStr 0x0006 "Sheet1"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemSt Value 
' Line #482:
' 	LitStr 0x0003 "BH2"
' 	LitStr 0x0002 "BR"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0003 "AZ2"
' 	LitStr 0x0006 "Sheet1"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemSt Value 
' Line #483:
' 	LitStr 0x0003 "BI2"
' 	LitStr 0x0002 "BR"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0003 "BB2"
' 	LitStr 0x0006 "Sheet1"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemSt Value 
' Line #484:
' 	EndIfBlock 
' Line #485:
' Line #486:
' 	EndSub 
' Line #487:
' Line #488:
' 	FuncDefn (Sub Macro1())
' Line #489:
' 	QuoteRem 0x0000 0x0000 ""
' Line #490:
' 	QuoteRem 0x0000 0x0010 " ENLrefine Macro"
' Line #491:
' 	QuoteRem 0x0000 0x0032 " Improved accuracy for ENL on flights >10000 fixes"
' Line #492:
' 	QuoteRem 0x0000 0x0000 ""
' Line #493:
' 	LitVarSpecial (False)
' 	Ld Application 
' 	MemSt ScreenUpdating 
' Line #494:
' 	LitStr 0x0002 "BR"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemCall Activate 0x0000 
' Line #495:
' 	LitR8 0x0FF3 0xC901 0x573A 0x3F0E 
' 	LitStr 0x0002 "G1"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #496:
' 	LitStr 0x007F "=IF(OR(AND(RC[1]<=Sheet1!R2C26+R1C7,RC[1]>=Sheet1!R2C26-R1C7),AND(RC[1]<=Sheet1!R4C38+R1C7,RC[1]>=Sheet1!R4C38-R1C7)),RC[1],"")"
' 	LitStr 0x0009 "I1:I60000"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #497:
' 	LitStr 0x0009 "I1:I60000"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0009 "I1:I60000"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #498:
' Line #499:
' 	LitStr 0x0009 "I1:I60000"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x000B "AT1:AT60000"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #500:
' 	LitStr 0x0018 "=IF(RC[37]<>"",RC[2],"")"
' 	LitStr 0x0009 "I1:I60000"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #501:
' 	LitStr 0x0009 "I1:I60000"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x000B "AU1:AU60000"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #502:
' Line #503:
' 	LitStr 0x0005 "AT:AU"
' 	ArgsLd Columns 0x0001 
' 	ArgsMemCall Select 0x0000 
' Line #504:
' 	LitStr 0x0002 "BR"
' 	Ld ENLrefine 
' 	ArgsMemLd Worksheets 0x0001 
' 	MemLd Sort 
' 	MemLd ActiveWorkbook 
' 	ArgsMemCall Clear 0x0000 
' Line #505:
' 	LineCont 0x0004 13 00 08 00
' 	LitStr 0x0005 "AT:AT"
' 	ArgsLd Columns 0x0001 
' 	ParamNamed Add 
' 	Ld SortOn 
' 	ParamNamed Key 
' 	Ld xlAscending 
' 	ParamNamed xlSortOnValues 
' 	Ld xlSortNormal 
' 	ParamNamed Order 
' 	LitStr 0x0002 "BR"
' 	Ld ENLrefine 
' 	ArgsMemLd Worksheets 0x0001 
' 	MemLd Sort 
' 	MemLd ActiveWorkbook 
' 	ArgsMemCall SortFields 0x0004 
' Line #506:
' 	StartWithExpr 
' 	LitStr 0x0002 "BR"
' 	Ld ENLrefine 
' 	ArgsMemLd Worksheets 0x0001 
' 	MemLd Sort 
' 	With 
' Line #507:
' 	LitStr 0x0005 "AT:AU"
' 	ArgsLd Columns 0x0001 
' 	ArgsMemCallWith DataOption 0x0001 
' Line #508:
' 	Ld xlGuess 
' 	MemStWith Header 
' Line #509:
' 	LitVarSpecial (False)
' 	MemStWith MatchCase 
' Line #510:
' 	Ld xlTopToBottom 
' 	MemStWith Orientation 
' Line #511:
' 	Ld SortMethod 
' 	MemStWith SetRange 
' Line #512:
' 	ArgsMemCallWith xlPinYin 0x0000 
' Line #513:
' 	EndWith 
' Line #514:
' 	LitStr 0x0017 "=MIN(RC[12]:R[10]C[12])"
' 	LitStr 0x0003 "AR1"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #515:
' 	LitStr 0x001B "=MAX(R[10]C[12]:R[20]C[12])"
' 	LitStr 0x0003 "AR2"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #516:
' 	LitStr 0x000D "=Sheet1!R6C25"
' 	LitStr 0x0003 "AS1"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #517:
' 	LitStr 0x0012 "=IMP!R[-1]C[-28]-1"
' 	LitStr 0x0003 "AS2"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #518:
' Line #519:
' 	LitStr 0x0014 "=MID(RC[-8],R2C45,3)"
' 	LitStr 0x0008 "BC1:BC22"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #520:
' 	LitStr 0x0008 "BC1:BC22"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0008 "BC1:BC22"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #521:
' 	LitStr 0x001C "=IF(RC[-1]<R1C45,RC[-10],"")"
' 	LitStr 0x0008 "BD1:BD22"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #522:
' 	LitStr 0x002C "=IF(OR(RC[-1]=R1C44,RC[-1]=R2C44),RC[-1],"")"
' 	LitStr 0x0008 "BE1:BE22"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #523:
' 	LitStr 0x001A "=IF(RC[-1]<>"",RC[-11],"")"
' 	LitStr 0x0008 "BF1:BF22"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #524:
' 	LitStr 0x0019 "=IF(RC[-1]<>"",RC[-4],"")"
' 	LitStr 0x0008 "BG1:BG22"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #525:
' Line #526:
' 	LitStr 0x0008 "BE1:BG22"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0008 "BE1:BG22"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #527:
' Line #528:
' 	LitStr 0x0002 "BR"
' 	Ld ENLrefine 
' 	ArgsMemLd Worksheets 0x0001 
' 	MemLd Sort 
' 	MemLd ActiveWorkbook 
' 	ArgsMemCall Clear 0x0000 
' Line #529:
' 	LineCont 0x0004 13 00 08 00
' 	LitStr 0x0003 "BE1"
' 	ArgsLd Range 0x0001 
' 	ParamNamed Add 
' 	Ld SortOn 
' 	ParamNamed Key 
' 	Ld xlAscending 
' 	ParamNamed xlSortOnValues 
' 	Ld xlSortNormal 
' 	ParamNamed Order 
' 	LitStr 0x0002 "BR"
' 	Ld ENLrefine 
' 	ArgsMemLd Worksheets 0x0001 
' 	MemLd Sort 
' 	MemLd ActiveWorkbook 
' 	ArgsMemCall SortFields 0x0004 
' Line #530:
' 	StartWithExpr 
' 	LitStr 0x0002 "BR"
' 	Ld ENLrefine 
' 	ArgsMemLd Worksheets 0x0001 
' 	MemLd Sort 
' 	With 
' Line #531:
' 	LitStr 0x0008 "BE1:BG22"
' 	ArgsLd Range 0x0001 
' 	ArgsMemCallWith DataOption 0x0001 
' Line #532:
' 	Ld xlGuess 
' 	MemStWith Header 
' Line #533:
' 	LitVarSpecial (False)
' 	MemStWith MatchCase 
' Line #534:
' 	Ld xlTopToBottom 
' 	MemStWith Orientation 
' Line #535:
' 	Ld SortMethod 
' 	MemStWith SetRange 
' Line #536:
' 	ArgsMemCallWith xlPinYin 0x0000 
' Line #537:
' 	EndWith 
' Line #538:
' Line #539:
' 	LitStr 0x0008 "AT1:BD22"
' 	ArgsLd Range 0x0001 
' 	ArgsMemCall Select 0x0000 
' Line #540:
' 	Ld Selection 
' 	ArgsMemCall Clear 0x0000 
' Line #541:
' Line #542:
' 	LitStr 0x0007 "BE1:BG2"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0007 "AT1:AV2"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #543:
' 	LitStr 0x0007 "AV1:AV2"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0007 "BC1:BC2"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #544:
' 	LitStr 0x0007 "AV1:AV2"
' 	ArgsLd Range 0x0001 
' 	ArgsMemCall Clear 0x0000 
' Line #545:
' 	LitStr 0x0007 "AU1:AU2"
' 	ArgsLd Range 0x0001 
' 	ArgsMemCall Select 0x0000 
' Line #546:
' 	LineCont 0x000C 0E 00 08 00 2F 00 08 00 5C 00 08 00
' 	LitStr 0x0003 "AU1"
' 	ArgsLd Range 0x0001 
' 	ParamNamed Destination 
' 	Ld xlFixedWidth 
' 	ParamNamed DataType 
' 	LitStr 0x0001 "E"
' 	ParamNamed OtherChar 
' 	LitDI2 0x0000 
' 	LitDI2 0x0009 
' 	ArgsArray Array 0x0002 
' 	LitDI2 0x0006 
' 	LitDI2 0x0001 
' 	ArgsArray Array 0x0002 
' 	LitDI2 0x0008 
' 	LitDI2 0x0001 
' 	ArgsArray Array 0x0002 
' 	LitDI2 0x000D 
' 	LitDI2 0x0001 
' 	ArgsArray Array 0x0002 
' 	LitDI2 0x000E 
' 	LitDI2 0x0001 
' 	ArgsArray Array 0x0002 
' 	LitDI2 0x0011 
' 	LitDI2 0x0001 
' 	ArgsArray Array 0x0002 
' 	LitDI2 0x0016 
' 	LitDI2 0x0001 
' 	ArgsArray Array 0x0002 
' 	LitDI2 0x0017 
' 	LitDI2 0x0009 
' 	ArgsArray Array 0x0002 
' 	LitDI2 0x0018 
' 	LitDI2 0x0001 
' 	ArgsArray Array 0x0002 
' 	LitDI2 0x001D 
' 	LitDI2 0x0001 
' 	ArgsArray Array 0x0002 
' 	LitDI2 0x0022 
' 	LitDI2 0x0009 
' 	ArgsArray Array 0x0002 
' 	ArgsArray Array 0x000B 
' 	ParamNamed FieldInfo 
' 	LitVarSpecial (True)
' 	ParamNamed TrailingMinusNumbers 
' 	Ld Selection 
' 	ArgsMemCall TextToColumns 0x0005 
' Line #547:
' Line #548:
' 	LitStr 0x0007 "=R[-2]C"
' 	LitStr 0x0007 "AT3:AT4"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #549:
' 	LitStr 0x0023 "=IF(R[-2]C[2]="N",R[-2]C,-1*R[-2]C)"
' 	LitStr 0x0007 "AU3:AU4"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #550:
' 	LitStr 0x002F "=IF(R[-2]C[1]="N",R[-2]C/1000,-1*(R[-2]C/1000))"
' 	LitStr 0x0007 "AV3:AV4"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #551:
' 	LitStr 0x0023 "=IF(R[-2]C[2]="E",R[-2]C,-1*R[-2]C)"
' 	LitStr 0x0007 "AX3:AX4"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #552:
' 	LitStr 0x002F "=IF(R[-2]C[1]="E",R[-2]C/1000,-1*(R[-2]C/1000))"
' 	LitStr 0x0007 "AY3:AY4"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #553:
' 	LitStr 0x0007 "=R[-2]C"
' 	LitStr 0x0007 "BA3:BC4"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #554:
' Line #555:
' 	LitStr 0x0007 "AT3:BC4"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0007 "AT3:BC4"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #556:
' 	LitStr 0x0007 "AR1:BG2"
' 	ArgsLd Range 0x0001 
' 	ArgsMemCall Clear 0x0000 
' Line #557:
' Line #558:
' 	LitStr 0x0007 "AT3:BC4"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0007 "AT1:BC2"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #559:
' 	LitStr 0x0007 "AT3:BC4"
' 	ArgsLd Range 0x0001 
' 	ArgsMemCall Clear 0x0000 
' Line #560:
' Line #561:
' 	LitStr 0x0006 "Sheet1"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemCall Activate 0x0000 
' Line #562:
' 	LitStr 0x000E "=BR!R[-7]C[20]"
' 	LitStr 0x0002 "Z8"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #563:
' 	LitStr 0x000E "=BR!R[-7]C[17]"
' 	LitStr 0x0003 "AD8"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #564:
' 	LitStr 0x000E "=BR!R[-7]C[17]"
' 	LitStr 0x0003 "AE8"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #565:
' 	LitStr 0x000E "=BR!R[-7]C[17]"
' 	LitStr 0x0003 "AG8"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #566:
' 	LitStr 0x000E "=BR!R[-7]C[17]"
' 	LitStr 0x0003 "AH8"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #567:
' 	QuoteRem 0x0004 0x002B "Range("AJ3").FormulaR1C1 = "=BR!R[-2]C[18]""
' Line #568:
' 	LitStr 0x000E "=BR!R[-7]C[17]"
' 	LitStr 0x0003 "AJ8"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #569:
' 	LitStr 0x000E "=BR!R[-7]C[18]"
' 	LitStr 0x0003 "AK8"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #570:
' 	LitStr 0x000D "=BR!R[-6]C[6]"
' 	LitStr 0x0003 "AN8"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #571:
' 	QuoteRem 0x0004 0x002B "Range("AO3").FormulaR1C1 = "=BR!R[-1]C[13]""
' Line #572:
' 	LitStr 0x000E "=BR!R[-6]C[12]"
' 	LitStr 0x0003 "AO8"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #573:
' 	LitStr 0x000D "=BR!R[-6]C[5]"
' 	LitStr 0x0003 "AP8"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #574:
' 	LitStr 0x000D "=BR!R[-6]C[5]"
' 	LitStr 0x0003 "AQ8"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #575:
' 	LitStr 0x000D "=BR!R[-6]C[6]"
' 	LitStr 0x0003 "AR8"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #576:
' 	LitStr 0x000D "=BR!R[-6]C[6]"
' 	LitStr 0x0003 "AS8"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #577:
' 	LitStr 0x000D "=BR!R[-6]C[9]"
' 	LitStr 0x0003 "AT8"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #578:
' 	LitStr 0x0006 "Z8:AT8"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0006 "Z8:AT8"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #579:
' 	QuoteRem 0x0004 0x002F "Range("AJ3:AO3").Value = Range("AJ3:AO3").Value"
' Line #580:
' 	LitStr 0x0002 "BR"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemCall Activate 0x0000 
' Line #581:
' 	LitStr 0x0003 "I:I"
' 	ArgsLd Columns 0x0001 
' 	ArgsMemCall Clear 0x0000 
' Line #582:
' 	LitStr 0x000A "G1,AT1:BC2"
' 	ArgsLd Range 0x0001 
' 	ArgsMemCall Clear 0x0000 
' Line #583:
' 	LitStr 0x0002 "A1"
' 	ArgsLd Range 0x0001 
' 	ArgsMemCall Select 0x0000 
' Line #584:
' 	EndSub 
' Line #585:
' Line #586:
' 	FuncDefn (Sub NewBRecords())
' Line #587:
' 	QuoteRem 0x0000 0x0000 ""
' Line #588:
' 	QuoteRem 0x0000 0x0024 " JLR 04/04/14 removed ref to CU Free"
' Line #589:
' 	QuoteRem 0x0000 0x0000 ""
' Line #590:
' 	LitStr 0x0003 "IMP"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemCall Select 0x0000 
' Line #591:
' 	LitStr 0x001B "=IF(RC[-16]="C",RC[-17],"")"
' 	LitStr 0x0008 "R1:R1020"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #592:
' 	LitStr 0x0008 "R1:R1020"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0008 "R1:R1020"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #593:
' 	LitStr 0x0028 "=IF(AND(R[-1]C[-1]="",RC[-17]="C"),1,"")"
' 	LitStr 0x0008 "S2:S1020"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #594:
' 	LitStr 0x0008 "S2:S1020"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0008 "S2:S1020"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #595:
' 	LitStr 0x0014 "=SUM(R[1]C:R[1020]C)"
' 	LitStr 0x0002 "S1"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #596:
' 	LitStr 0x0017 "=IF(RC[-1]=1,RC[-2],"")"
' 	LitStr 0x0008 "T1:T1020"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #597:
' 	LitStr 0x0008 "T1:T1020"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0008 "T1:T1020"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #598:
' 	LitStr 0x0002 "S1"
' 	ArgsLd Range 0x0001 
' 	LitDI2 0x0000 
' 	Gt 
' 	IfBlock 
' Line #599:
' 	LineCont 0x000C 11 00 08 00 33 00 08 00 59 00 08 00
' 	LitStr 0x0002 "T1"
' 	ArgsLd Range 0x0001 
' 	ParamNamed Destination 
' 	Ld xlFixedWidth 
' 	ParamNamed DataType 
' 	LitStr 0x0001 "E"
' 	ParamNamed OtherChar 
' 	LitDI2 0x0000 
' 	LitDI2 0x0009 
' 	ArgsArray Array 0x0002 
' 	LitDI2 0x0001 
' 	LitDI2 0x0004 
' 	ArgsArray Array 0x0002 
' 	LitDI2 0x0007 
' 	LitDI2 0x0001 
' 	ArgsArray Array 0x0002 
' 	LitDI2 0x0009 
' 	LitDI2 0x0001 
' 	ArgsArray Array 0x0002 
' 	LitDI2 0x000B 
' 	LitDI2 0x0001 
' 	ArgsArray Array 0x0002 
' 	LitDI2 0x000D 
' 	LitDI2 0x0009 
' 	ArgsArray Array 0x0002 
' 	LitDI2 0x0013 
' 	LitDI2 0x0009 
' 	ArgsArray Array 0x0002 
' 	LitDI2 0x0017 
' 	LitDI2 0x0001 
' 	ArgsArray Array 0x0002 
' 	LitDI2 0x0019 
' 	LitDI2 0x0001 
' 	ArgsArray Array 0x0002 
' 	ArgsArray Array 0x0009 
' 	ParamNamed FieldInfo 
' 	LitVarSpecial (True)
' 	ParamNamed TrailingMinusNumbers 
' 	LitStr 0x0003 "T:T"
' 	ArgsLd Columns 0x0001 
' 	ArgsMemCall TextToColumns 0x0005 
' Line #600:
' 	LineCont 0x0004 07 00 08 00
' 	LitStr 0x0033 "=IF(RC[-6]="","",RC[-5]+TIME(RC[-4],RC[-3],RC[-2]))"
' 	LitStr 0x0008 "Y3:Y1020"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #601:
' 	LitStr 0x0008 "Y3:Y1020"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0008 "Y3:Y1020"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #602:
' 	LitStr 0x0013 "=MAX(R[1]C:R[998]C)"
' 	LitStr 0x0002 "Y2"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #603:
' 	LitStr 0x0006 "=R[1]C"
' 	LitStr 0x0002 "Y1"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #604:
' 	LitStr 0x002F "=IF(OR(RC[-1]=R2C25,RC[-1]=R1021C25),RC[-2],"")"
' 	LitStr 0x0008 "Z3:Z1020"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #605:
' 	LitStr 0x0008 "Z3:Z1020"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0008 "Z3:Z1020"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #606:
' 	LitStr 0x0014 "=MAX(R[2]C:R[1021]C)"
' 	LitStr 0x0002 "Z1"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #607:
' 	LitStr 0x0013 "=MAX(R[1]C:R[998]C)"
' 	LitStr 0x0002 "Z2"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #608:
' 	QuoteRem 0x0004 0x001A "Added 7/26/16 for LXN7007F"
' Line #609:
' 	LitStr 0x0008 "Z1:Z1020"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0008 "Z1:Z1020"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #610:
' 	EndIfBlock 
' Line #611:
' Line #612:
' 	QuoteRem 0x0000 0x0032 "If Range("Z1") <> "" Then changed per next 7/26/16"
' Line #613:
' 	LitStr 0x0002 "Z1"
' 	ArgsLd Range 0x0001 
' 	LitDI2 0x0000 
' 	Ne 
' 	IfBlock 
' Line #614:
' 	LitStr 0x0038 "=IF(AND(RC[-2]>0,OR(RC[-2]=R2C25,RC[-2]=R1021C25)),1,"")"
' 	LitStr 0x000A "AA3:AA1020"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #615:
' 	LitStr 0x000A "AA2:AA1020"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x000A "AA2:AA1020"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #616:
' 	LitStr 0x001C "=IF(R[-2]C[-8]=1,RC[-17],"")"
' 	LitStr 0x000A "AI3:AI1020"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #617:
' 	LitStr 0x000A "AI3:AI1020"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x000A "AI3:AI1020"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #618:
' 	LitStr 0x0003 "AI1"
' 	ArgsLd Range 0x0001 
' 	ParamNamed Destination 
' 	Ld xlFixedWidth 
' 	ParamNamed DataType 
' 	LitStr 0x0001 "E"
' 	ParamNamed OtherChar 
' 	LitDI2 0x0000 
' 	LitDI2 0x0009 
' 	ArgsArray Array 0x0002 
' 	LitDI2 0x0001 
' 	LitDI2 0x0001 
' 	ArgsArray Array 0x0002 
' 	LitDI2 0x0003 
' 	LitDI2 0x0001 
' 	ArgsArray Array 0x0002 
' 	LitDI2 0x0008 
' 	LitDI2 0x0001 
' 	ArgsArray Array 0x0002 
' 	LitDI2 0x0009 
' 	LitDI2 0x0001 
' 	ArgsArray Array 0x0002 
' 	LitDI2 0x000C 
' 	LitDI2 0x0001 
' 	ArgsArray Array 0x0002 
' 	LitDI2 0x0011 
' 	LitDI2 0x0001 
' 	ArgsArray Array 0x0002 
' 	LitDI2 0x0012 
' 	LitDI2 0x0001 
' 	ArgsArray Array 0x0002 
' 	ArgsArray Array 0x0008 
' 	ParamNamed FieldInfo 
' 	LitVarSpecial (True)
' 	ParamNamed TrailingMinusNumbers 
' 	LitStr 0x0005 "AI:AI"
' 	ArgsLd Columns 0x0001 
' 	ArgsMemCall TextToColumns 0x0005 
' Line #619:
' 	LitStr 0x0003 "AK1"
' 	ArgsLd Range 0x0001 
' 	ParamNamed Key1 
' 	Ld xlAscending 
' 	ParamNamed Order1 
' 	Ld xlGuess 
' 	ParamNamed Header 
' 	LitDI2 0x0001 
' 	ParamNamed OrderCustom 
' 	LitVarSpecial (False)
' 	ParamNamed MatchCase 
' 	Ld xlTopToBottom 
' 	ParamNamed Orientation 
' 	Ld xlSortNormal 
' 	ParamNamed DataOption1 
' 	LitStr 0x000A "AK1:AK1000"
' 	ArgsLd Range 0x0001 
' 	ArgsMemCall Sort 0x0007 
' Line #620:
' 	LitStr 0x0025 "=IF(R[-1]C="N",1,IF(R[-1]C="S",-1,0))"
' 	LitStr 0x0003 "AK2"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #621:
' 	LitStr 0x0003 "AN1"
' 	ArgsLd Range 0x0001 
' 	ParamNamed Key1 
' 	Ld xlAscending 
' 	ParamNamed Order1 
' 	Ld xlGuess 
' 	ParamNamed Header 
' 	LitDI2 0x0001 
' 	ParamNamed OrderCustom 
' 	LitVarSpecial (False)
' 	ParamNamed MatchCase 
' 	Ld xlTopToBottom 
' 	ParamNamed Orientation 
' 	Ld xlSortNormal 
' 	ParamNamed DataOption1 
' 	LitStr 0x000A "AN1:AN1000"
' 	ArgsLd Range 0x0001 
' 	ArgsMemCall Sort 0x0007 
' Line #622:
' 	LitStr 0x0025 "=IF(R[-1]C="E",1,IF(R[-1]C="W",-1,0))"
' 	LitStr 0x0003 "AN2"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #623:
' 	LitStr 0x0003 "AI3"
' 	ArgsLd Range 0x0001 
' 	ParamNamed Key1 
' 	Ld xlAscending 
' 	ParamNamed Order1 
' 	Ld xlGuess 
' 	ParamNamed Header 
' 	LitDI2 0x0001 
' 	ParamNamed OrderCustom 
' 	LitVarSpecial (False)
' 	ParamNamed MatchCase 
' 	Ld xlTopToBottom 
' 	ParamNamed Orientation 
' 	Ld xlSortNormal 
' 	ParamNamed DataOption1 
' 	LitStr 0x000A "AI3:AM1000"
' 	ArgsLd Range 0x0001 
' 	ArgsMemCall Sort 0x0007 
' Line #624:
' 	LitStr 0x002C "=IF(AND(R[1019]C<>"",R1C7=5),"",R[1]C*RC[2])"
' 	LitStr 0x0003 "AI2"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #625:
' 	LitStr 0x0032 "=IF(AND(R[1019]C<>"",R1C7=5),"",R[1]C*RC[1]*0.001)"
' 	LitStr 0x0003 "AJ2"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #626:
' 	LitStr 0x002C "=IF(AND(R[1019]C<>"",R1C7=5),"",R[1]C*RC[2])"
' 	LitStr 0x0003 "AL2"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #627:
' 	LitStr 0x0032 "=IF(AND(R[1019]C<>"",R1C7=5),"",R[1]C*RC[1]*0.001)"
' 	LitStr 0x0003 "AM2"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #628:
' 	LitStr 0x0003 "AO1"
' 	ArgsLd Range 0x0001 
' 	ParamNamed Key1 
' 	Ld xlAscending 
' 	ParamNamed Order1 
' 	Ld xlGuess 
' 	ParamNamed Header 
' 	LitDI2 0x0001 
' 	ParamNamed OrderCustom 
' 	LitVarSpecial (False)
' 	ParamNamed MatchCase 
' 	Ld xlTopToBottom 
' 	ParamNamed Orientation 
' 	Ld xlSortNormal 
' 	ParamNamed DataOption1 
' 	LitStr 0x000A "AO1:AO1000"
' 	ArgsLd Range 0x0001 
' 	ArgsMemCall Sort 0x0007 
' Line #629:
' 	LitStr 0x001D "=IF(R[-3]C[-15]=1,RC[-24],"")"
' 	LitStr 0x000A "AP4:AP1020"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #630:
' 	LitStr 0x000A "AP4:AP1020"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x000A "AP4:AP1020"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #631:
' 	LitStr 0x0003 "AP1"
' 	ArgsLd Range 0x0001 
' 	ParamNamed Destination 
' 	Ld xlFixedWidth 
' 	ParamNamed DataType 
' 	LitStr 0x0001 "E"
' 	ParamNamed OtherChar 
' 	LitDI2 0x0000 
' 	LitDI2 0x0009 
' 	ArgsArray Array 0x0002 
' 	LitDI2 0x0001 
' 	LitDI2 0x0001 
' 	ArgsArray Array 0x0002 
' 	LitDI2 0x0003 
' 	LitDI2 0x0001 
' 	ArgsArray Array 0x0002 
' 	LitDI2 0x0008 
' 	LitDI2 0x0001 
' 	ArgsArray Array 0x0002 
' 	LitDI2 0x0009 
' 	LitDI2 0x0001 
' 	ArgsArray Array 0x0002 
' 	LitDI2 0x000C 
' 	LitDI2 0x0001 
' 	ArgsArray Array 0x0002 
' 	LitDI2 0x0011 
' 	LitDI2 0x0001 
' 	ArgsArray Array 0x0002 
' 	LitDI2 0x0012 
' 	LitDI2 0x0001 
' 	ArgsArray Array 0x0002 
' 	ArgsArray Array 0x0008 
' 	ParamNamed FieldInfo 
' 	LitVarSpecial (True)
' 	ParamNamed TrailingMinusNumbers 
' 	LitStr 0x0005 "AP:AP"
' 	ArgsLd Columns 0x0001 
' 	ArgsMemCall TextToColumns 0x0005 
' Line #632:
' 	LitStr 0x0003 "AR1"
' 	ArgsLd Range 0x0001 
' 	ParamNamed Key1 
' 	Ld xlAscending 
' 	ParamNamed Order1 
' 	Ld xlGuess 
' 	ParamNamed Header 
' 	LitDI2 0x0001 
' 	ParamNamed OrderCustom 
' 	LitVarSpecial (False)
' 	ParamNamed MatchCase 
' 	Ld xlTopToBottom 
' 	ParamNamed Orientation 
' 	Ld xlSortNormal 
' 	ParamNamed DataOption1 
' 	LitStr 0x000A "AR1:AR1000"
' 	ArgsLd Range 0x0001 
' 	ArgsMemCall Sort 0x0007 
' Line #633:
' 	LitStr 0x0025 "=IF(R[-1]C="N",1,IF(R[-1]C="S",-1,0))"
' 	LitStr 0x0003 "AR2"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #634:
' 	LitStr 0x0003 "AU1"
' 	ArgsLd Range 0x0001 
' 	ParamNamed Key1 
' 	Ld xlAscending 
' 	ParamNamed Order1 
' 	Ld xlGuess 
' 	ParamNamed Header 
' 	LitDI2 0x0001 
' 	ParamNamed OrderCustom 
' 	LitVarSpecial (False)
' 	ParamNamed MatchCase 
' 	Ld xlTopToBottom 
' 	ParamNamed Orientation 
' 	Ld xlSortNormal 
' 	ParamNamed DataOption1 
' 	LitStr 0x000A "AU1:AU1000"
' 	ArgsLd Range 0x0001 
' 	ArgsMemCall Sort 0x0007 
' Line #635:
' 	LitStr 0x0025 "=IF(R[-1]C="E",1,IF(R[-1]C="W",-1,0))"
' 	LitStr 0x0003 "AU2"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #636:
' 	LitStr 0x0003 "AP3"
' 	ArgsLd Range 0x0001 
' 	ParamNamed Key1 
' 	Ld xlAscending 
' 	ParamNamed Order1 
' 	Ld xlGuess 
' 	ParamNamed Header 
' 	LitDI2 0x0001 
' 	ParamNamed OrderCustom 
' 	LitVarSpecial (False)
' 	ParamNamed MatchCase 
' 	Ld xlTopToBottom 
' 	ParamNamed Orientation 
' 	Ld xlSortNormal 
' 	ParamNamed DataOption1 
' 	LitStr 0x000A "AP3:AT1000"
' 	ArgsLd Range 0x0001 
' 	ArgsMemCall Sort 0x0007 
' Line #637:
' 	LitStr 0x000C "=R[1]C*RC[2]"
' 	LitStr 0x0003 "AP2"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #638:
' 	LitStr 0x0012 "=R[1]C*RC[1]*0.001"
' 	LitStr 0x0003 "AQ2"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #639:
' 	LitStr 0x000C "=R[1]C*RC[2]"
' 	LitStr 0x0003 "AS2"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #640:
' 	LitStr 0x0012 "=R[1]C*RC[1]*0.001"
' 	LitStr 0x0003 "AT2"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #641:
' 	LitStr 0x0003 "AV1"
' 	ArgsLd Range 0x0001 
' 	ParamNamed Key1 
' 	Ld xlAscending 
' 	ParamNamed Order1 
' 	Ld xlGuess 
' 	ParamNamed Header 
' 	LitDI2 0x0001 
' 	ParamNamed OrderCustom 
' 	LitVarSpecial (False)
' 	ParamNamed MatchCase 
' 	Ld xlTopToBottom 
' 	ParamNamed Orientation 
' 	Ld xlSortNormal 
' 	ParamNamed DataOption1 
' 	LitStr 0x000A "AV1:AV1000"
' 	ArgsLd Range 0x0001 
' 	ArgsMemCall Sort 0x0007 
' Line #642:
' 	EndIfBlock 
' Line #643:
' Line #644:
' 	LitStr 0x0002 "Z1"
' 	ArgsLd Range 0x0001 
' 	LitDI2 0x0001 
' 	Ge 
' 	IfBlock 
' Line #645:
' 	LitStr 0x001D "=IF(R[-4]C[-22]=1,RC[-31],"")"
' 	LitStr 0x000A "AW5:AW1020"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #646:
' 	LitStr 0x000A "AW5:AW1020"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x000A "AW5:AW1020"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #647:
' 	LitStr 0x0003 "AW1"
' 	ArgsLd Range 0x0001 
' 	ParamNamed Destination 
' 	Ld xlFixedWidth 
' 	ParamNamed DataType 
' 	LitStr 0x0001 "E"
' 	ParamNamed OtherChar 
' 	LitDI2 0x0000 
' 	LitDI2 0x0009 
' 	ArgsArray Array 0x0002 
' 	LitDI2 0x0001 
' 	LitDI2 0x0001 
' 	ArgsArray Array 0x0002 
' 	LitDI2 0x0003 
' 	LitDI2 0x0001 
' 	ArgsArray Array 0x0002 
' 	LitDI2 0x0008 
' 	LitDI2 0x0001 
' 	ArgsArray Array 0x0002 
' 	LitDI2 0x0009 
' 	LitDI2 0x0001 
' 	ArgsArray Array 0x0002 
' 	LitDI2 0x000C 
' 	LitDI2 0x0001 
' 	ArgsArray Array 0x0002 
' 	LitDI2 0x0011 
' 	LitDI2 0x0001 
' 	ArgsArray Array 0x0002 
' 	LitDI2 0x0012 
' 	LitDI2 0x0001 
' 	ArgsArray Array 0x0002 
' 	ArgsArray Array 0x0008 
' 	ParamNamed FieldInfo 
' 	LitVarSpecial (True)
' 	ParamNamed TrailingMinusNumbers 
' 	LitStr 0x0005 "AW:AW"
' 	ArgsLd Columns 0x0001 
' 	ArgsMemCall TextToColumns 0x0005 
' Line #648:
' 	LitStr 0x0003 "AY1"
' 	ArgsLd Range 0x0001 
' 	ParamNamed Key1 
' 	Ld xlAscending 
' 	ParamNamed Order1 
' 	Ld xlGuess 
' 	ParamNamed Header 
' 	LitDI2 0x0001 
' 	ParamNamed OrderCustom 
' 	LitVarSpecial (False)
' 	ParamNamed MatchCase 
' 	Ld xlTopToBottom 
' 	ParamNamed Orientation 
' 	Ld xlSortNormal 
' 	ParamNamed DataOption1 
' 	LitStr 0x000A "AY1:AY1000"
' 	ArgsLd Range 0x0001 
' 	ArgsMemCall Sort 0x0007 
' Line #649:
' 	LitStr 0x0025 "=IF(R[-1]C="N",1,IF(R[-1]C="S",-1,0))"
' 	LitStr 0x0003 "AY2"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #650:
' 	LitStr 0x0003 "BB1"
' 	ArgsLd Range 0x0001 
' 	ParamNamed Key1 
' 	Ld xlAscending 
' 	ParamNamed Order1 
' 	Ld xlGuess 
' 	ParamNamed Header 
' 	LitDI2 0x0001 
' 	ParamNamed OrderCustom 
' 	LitVarSpecial (False)
' 	ParamNamed MatchCase 
' 	Ld xlTopToBottom 
' 	ParamNamed Orientation 
' 	Ld xlSortNormal 
' 	ParamNamed DataOption1 
' 	LitStr 0x000A "BB1:BB1000"
' 	ArgsLd Range 0x0001 
' 	ArgsMemCall Sort 0x0007 
' Line #651:
' 	LitStr 0x0025 "=IF(R[-1]C="E",1,IF(R[-1]C="W",-1,0))"
' 	LitStr 0x0003 "BB2"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #652:
' 	LitStr 0x0003 "AW3"
' 	ArgsLd Range 0x0001 
' 	ParamNamed Key1 
' 	Ld xlAscending 
' 	ParamNamed Order1 
' 	Ld xlGuess 
' 	ParamNamed Header 
' 	LitDI2 0x0001 
' 	ParamNamed OrderCustom 
' 	LitVarSpecial (False)
' 	ParamNamed MatchCase 
' 	Ld xlTopToBottom 
' 	ParamNamed Orientation 
' 	Ld xlSortNormal 
' 	ParamNamed DataOption1 
' 	LitStr 0x000A "AW3:BA1000"
' 	ArgsLd Range 0x0001 
' 	ArgsMemCall Sort 0x0007 
' Line #653:
' 	LitStr 0x000C "=R[1]C*RC[2]"
' 	LitStr 0x0003 "AW2"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #654:
' 	LitStr 0x0012 "=R[1]C*RC[1]*0.001"
' 	LitStr 0x0003 "AX2"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #655:
' 	LitStr 0x000C "=R[1]C*RC[2]"
' 	LitStr 0x0003 "AZ2"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #656:
' 	LitStr 0x0012 "=R[1]C*RC[1]*0.001"
' 	LitStr 0x0003 "BA2"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #657:
' 	LitStr 0x0003 "BC1"
' 	ArgsLd Range 0x0001 
' 	ParamNamed Key1 
' 	Ld xlAscending 
' 	ParamNamed Order1 
' 	Ld xlGuess 
' 	ParamNamed Header 
' 	LitDI2 0x0001 
' 	ParamNamed OrderCustom 
' 	LitVarSpecial (False)
' 	ParamNamed MatchCase 
' 	Ld xlTopToBottom 
' 	ParamNamed Orientation 
' 	Ld xlSortNormal 
' 	ParamNamed DataOption1 
' 	LitStr 0x000A "BC1:BC1000"
' 	ArgsLd Range 0x0001 
' 	ArgsMemCall Sort 0x0007 
' Line #658:
' 	EndIfBlock 
' Line #659:
' Line #660:
' 	LitStr 0x0002 "Z1"
' 	ArgsLd Range 0x0001 
' 	LitDI2 0x0002 
' 	Ge 
' 	IfBlock 
' Line #661:
' 	LitStr 0x001D "=IF(R[-5]C[-29]=1,RC[-38],"")"
' 	LitStr 0x000A "BD6:BD1020"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #662:
' 	LitStr 0x000A "BD6:BD1020"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x000A "BD6:BD1020"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #663:
' 	LitStr 0x0003 "BD1"
' 	ArgsLd Range 0x0001 
' 	ParamNamed Destination 
' 	Ld xlFixedWidth 
' 	ParamNamed DataType 
' 	LitStr 0x0001 "E"
' 	ParamNamed OtherChar 
' 	LitDI2 0x0000 
' 	LitDI2 0x0009 
' 	ArgsArray Array 0x0002 
' 	LitDI2 0x0001 
' 	LitDI2 0x0001 
' 	ArgsArray Array 0x0002 
' 	LitDI2 0x0003 
' 	LitDI2 0x0001 
' 	ArgsArray Array 0x0002 
' 	LitDI2 0x0008 
' 	LitDI2 0x0001 
' 	ArgsArray Array 0x0002 
' 	LitDI2 0x0009 
' 	LitDI2 0x0001 
' 	ArgsArray Array 0x0002 
' 	LitDI2 0x000C 
' 	LitDI2 0x0001 
' 	ArgsArray Array 0x0002 
' 	LitDI2 0x0011 
' 	LitDI2 0x0001 
' 	ArgsArray Array 0x0002 
' 	LitDI2 0x0012 
' 	LitDI2 0x0001 
' 	ArgsArray Array 0x0002 
' 	ArgsArray Array 0x0008 
' 	ParamNamed FieldInfo 
' 	LitVarSpecial (True)
' 	ParamNamed TrailingMinusNumbers 
' 	LitStr 0x0005 "BD:BD"
' 	ArgsLd Columns 0x0001 
' 	ArgsMemCall TextToColumns 0x0005 
' Line #664:
' 	LitStr 0x0003 "BF1"
' 	ArgsLd Range 0x0001 
' 	ParamNamed Key1 
' 	Ld xlAscending 
' 	ParamNamed Order1 
' 	Ld xlGuess 
' 	ParamNamed Header 
' 	LitDI2 0x0001 
' 	ParamNamed OrderCustom 
' 	LitVarSpecial (False)
' 	ParamNamed MatchCase 
' 	Ld xlTopToBottom 
' 	ParamNamed Orientation 
' 	Ld xlSortNormal 
' 	ParamNamed DataOption1 
' 	LitStr 0x000A "BF1:BF1000"
' 	ArgsLd Range 0x0001 
' 	ArgsMemCall Sort 0x0007 
' Line #665:
' 	LitStr 0x0025 "=IF(R[-1]C="N",1,IF(R[-1]C="S",-1,0))"
' 	LitStr 0x0003 "BF2"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #666:
' 	LitStr 0x0003 "BI1"
' 	ArgsLd Range 0x0001 
' 	ParamNamed Key1 
' 	Ld xlAscending 
' 	ParamNamed Order1 
' 	Ld xlGuess 
' 	ParamNamed Header 
' 	LitDI2 0x0001 
' 	ParamNamed OrderCustom 
' 	LitVarSpecial (False)
' 	ParamNamed MatchCase 
' 	Ld xlTopToBottom 
' 	ParamNamed Orientation 
' 	Ld xlSortNormal 
' 	ParamNamed DataOption1 
' 	LitStr 0x000A "BI1:BI1000"
' 	ArgsLd Range 0x0001 
' 	ArgsMemCall Sort 0x0007 
' Line #667:
' 	LitStr 0x0025 "=IF(R[-1]C="E",1,IF(R[-1]C="W",-1,0))"
' 	LitStr 0x0003 "BI2"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #668:
' 	LitStr 0x0003 "BD3"
' 	ArgsLd Range 0x0001 
' 	ParamNamed Key1 
' 	Ld xlAscending 
' 	ParamNamed Order1 
' 	Ld xlGuess 
' 	ParamNamed Header 
' 	LitDI2 0x0001 
' 	ParamNamed OrderCustom 
' 	LitVarSpecial (False)
' 	ParamNamed MatchCase 
' 	Ld xlTopToBottom 
' 	ParamNamed Orientation 
' 	Ld xlSortNormal 
' 	ParamNamed DataOption1 
' 	LitStr 0x000A "BD3:BH1000"
' 	ArgsLd Range 0x0001 
' 	ArgsMemCall Sort 0x0007 
' Line #669:
' 	LitStr 0x000C "=R[1]C*RC[2]"
' 	LitStr 0x0003 "BD2"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #670:
' 	LitStr 0x0012 "=R[1]C*RC[1]*0.001"
' 	LitStr 0x0003 "BE2"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #671:
' 	LitStr 0x000C "=R[1]C*RC[2]"
' 	LitStr 0x0003 "BG2"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #672:
' 	LitStr 0x0012 "=R[1]C*RC[1]*0.001"
' 	LitStr 0x0003 "BH2"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #673:
' 	LitStr 0x0003 "BJ1"
' 	ArgsLd Range 0x0001 
' 	ParamNamed Key1 
' 	Ld xlAscending 
' 	ParamNamed Order1 
' 	Ld xlGuess 
' 	ParamNamed Header 
' 	LitDI2 0x0001 
' 	ParamNamed OrderCustom 
' 	LitVarSpecial (False)
' 	ParamNamed MatchCase 
' 	Ld xlTopToBottom 
' 	ParamNamed Orientation 
' 	Ld xlSortNormal 
' 	ParamNamed DataOption1 
' 	LitStr 0x000A "BJ1:BJ1000"
' 	ArgsLd Range 0x0001 
' 	ArgsMemCall Sort 0x0007 
' Line #674:
' 	EndIfBlock 
' Line #675:
' Line #676:
' 	LitStr 0x0002 "Z1"
' 	ArgsLd Range 0x0001 
' 	LitDI2 0x0003 
' 	Ge 
' 	IfBlock 
' Line #677:
' 	LitStr 0x001D "=IF(R[-6]C[-36]=1,RC[-45],"")"
' 	LitStr 0x000A "BK7:BK1020"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #678:
' 	LitStr 0x000A "BK7:BK1020"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x000A "BK7:BK1020"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #679:
' 	LitStr 0x0003 "BK1"
' 	ArgsLd Range 0x0001 
' 	ParamNamed Destination 
' 	Ld xlFixedWidth 
' 	ParamNamed DataType 
' 	LitStr 0x0001 "E"
' 	ParamNamed OtherChar 
' 	LitDI2 0x0000 
' 	LitDI2 0x0009 
' 	ArgsArray Array 0x0002 
' 	LitDI2 0x0001 
' 	LitDI2 0x0001 
' 	ArgsArray Array 0x0002 
' 	LitDI2 0x0003 
' 	LitDI2 0x0001 
' 	ArgsArray Array 0x0002 
' 	LitDI2 0x0008 
' 	LitDI2 0x0001 
' 	ArgsArray Array 0x0002 
' 	LitDI2 0x0009 
' 	LitDI2 0x0001 
' 	ArgsArray Array 0x0002 
' 	LitDI2 0x000C 
' 	LitDI2 0x0001 
' 	ArgsArray Array 0x0002 
' 	LitDI2 0x0011 
' 	LitDI2 0x0001 
' 	ArgsArray Array 0x0002 
' 	LitDI2 0x0012 
' 	LitDI2 0x0001 
' 	ArgsArray Array 0x0002 
' 	ArgsArray Array 0x0008 
' 	ParamNamed FieldInfo 
' 	LitVarSpecial (True)
' 	ParamNamed TrailingMinusNumbers 
' 	LitStr 0x0005 "BK:BK"
' 	ArgsLd Columns 0x0001 
' 	ArgsMemCall TextToColumns 0x0005 
' Line #680:
' 	LitStr 0x0003 "BM1"
' 	ArgsLd Range 0x0001 
' 	ParamNamed Key1 
' 	Ld xlAscending 
' 	ParamNamed Order1 
' 	Ld xlGuess 
' 	ParamNamed Header 
' 	LitDI2 0x0001 
' 	ParamNamed OrderCustom 
' 	LitVarSpecial (False)
' 	ParamNamed MatchCase 
' 	Ld xlTopToBottom 
' 	ParamNamed Orientation 
' 	Ld xlSortNormal 
' 	ParamNamed DataOption1 
' 	LitStr 0x000A "BM1:BM1000"
' 	ArgsLd Range 0x0001 
' 	ArgsMemCall Sort 0x0007 
' Line #681:
' 	LitStr 0x0025 "=IF(R[-1]C="N",1,IF(R[-1]C="S",-1,0))"
' 	LitStr 0x0003 "BM2"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #682:
' 	LitStr 0x0003 "BK3"
' 	ArgsLd Range 0x0001 
' 	ParamNamed Key1 
' 	Ld xlDescending 
' 	ParamNamed Order1 
' 	Ld xlGuess 
' 	ParamNamed Header 
' 	LitDI2 0x0001 
' 	ParamNamed OrderCustom 
' 	LitVarSpecial (False)
' 	ParamNamed MatchCase 
' 	Ld xlTopToBottom 
' 	ParamNamed Orientation 
' 	Ld xlSortNormal 
' 	ParamNamed DataOption1 
' 	LitStr 0x000A "BK3:BO1000"
' 	ArgsLd Range 0x0001 
' 	ArgsMemCall Sort 0x0007 
' Line #683:
' 	LitStr 0x000C "=R[1]C*RC[2]"
' 	LitStr 0x0003 "BK2"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #684:
' 	LitStr 0x0012 "=R[1]C*RC[1]*0.001"
' 	LitStr 0x0003 "BL2"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #685:
' 	LitStr 0x000C "=R[1]C*RC[2]"
' 	LitStr 0x0003 "BN2"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #686:
' 	LitStr 0x0012 "=R[1]C*RC[1]*0.001"
' 	LitStr 0x0003 "BO2"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #687:
' 	LitStr 0x0003 "BP1"
' 	ArgsLd Range 0x0001 
' 	ParamNamed Key1 
' 	Ld xlAscending 
' 	ParamNamed Order1 
' 	Ld xlGuess 
' 	ParamNamed Header 
' 	LitDI2 0x0001 
' 	ParamNamed OrderCustom 
' 	LitVarSpecial (False)
' 	ParamNamed MatchCase 
' 	Ld xlTopToBottom 
' 	ParamNamed Orientation 
' 	Ld xlSortNormal 
' 	ParamNamed DataOption1 
' 	LitStr 0x000A "BP1:BP1000"
' 	ArgsLd Range 0x0001 
' 	ArgsMemCall Sort 0x0007 
' Line #688:
' 	LitStr 0x0025 "=IF(R[-1]C="E",1,IF(R[-1]C="W",-1,0))"
' 	LitStr 0x0003 "BP2"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #689:
' 	LitStr 0x0003 "BQ1"
' 	ArgsLd Range 0x0001 
' 	ParamNamed Key1 
' 	Ld xlAscending 
' 	ParamNamed Order1 
' 	Ld xlGuess 
' 	ParamNamed Header 
' 	LitDI2 0x0001 
' 	ParamNamed OrderCustom 
' 	LitVarSpecial (False)
' 	ParamNamed MatchCase 
' 	Ld xlTopToBottom 
' 	ParamNamed Orientation 
' 	Ld xlSortNormal 
' 	ParamNamed DataOption1 
' 	LitStr 0x000A "BQ1:BQ1000"
' 	ArgsLd Range 0x0001 
' 	ArgsMemCall Sort 0x0007 
' Line #690:
' 	EndIfBlock 
' Line #691:
' Line #692:
' 	LitStr 0x0003 "PRS"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemCall Select 0x0000 
' Line #693:
' 	LitStr 0x000E "=sunrise!R17C4"
' 	LitStr 0x0003 "G13"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #694:
' 	LitStr 0x000D "=Sheet1!R6C11"
' 	LitStr 0x0002 "M2"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #695:
' 	LitStr 0x000D "=Sheet1!R6C12"
' 	LitStr 0x0002 "M3"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #696:
' 	LitStr 0x0006 "A1:H30"
' 	LitStr 0x0003 "PRS"
' 	LitStr 0x0007 "Ab.xlsm"
' 	ArgsLd Workbooks 0x0001 
' 	ArgsMemLd Worksheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0006 "A1:H30"
' 	LitStr 0x0006 "Parsed"
' 	LitStr 0x0006 "A.xlsm"
' 	ArgsLd Workbooks 0x0001 
' 	ArgsMemLd Sheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemSt Value 
' Line #697:
' 	LitStr 0x0006 "A.xlsm"
' 	ArgsLd Workbooks 0x0001 
' 	ArgsMemCall Activate 0x0000 
' Line #698:
' 	QuoteRem 0x0004 0x001D "Application.Run "A.xlsm!CALx""
' Line #699:
' 	QuoteRem 0x0004 0x0020 "Sheets("Parsed").Visible = False"
' Line #700:
' 	QuoteRem 0x0004 0x001E "Sheets("E-Dec").Visible = True"
' Line #701:
' 	QuoteRem 0x0004 0x0018 "Sheets("E-Dec").Activate"
' Line #702:
' 	QuoteRem 0x0004 0x0027 "ActiveSheet.Unprotect Password:="spike""
' Line #703:
' 	QuoteRem 0x0004 0x0016 "Range("A1:J29").Select"
' Line #704:
' 	QuoteRem 0x0004 0x0018 "ActiveWindow.Zoom = True"
' Line #705:
' 	QuoteRem 0x0004 0x0025 "ActiveSheet.Protect Password:="spike""
' Line #706:
' 	QuoteRem 0x0004 0x001E "Sheets("Logo").Visible = False"
' Line #707:
' 	QuoteRem 0x0004 0x0029 "Sheets("Data Entry Check").Visible = True"
' Line #708:
' 	QuoteRem 0x0004 0x0021 "Sheets("Data Entry Check").Select"
' Line #709:
' 	QuoteRem 0x0004 0x0027 "ActiveSheet.Unprotect Password:="spike""
' Line #710:
' 	QuoteRem 0x0004 0x0016 "Range("A1:K30").Select"
' Line #711:
' 	QuoteRem 0x0004 0x0018 "ActiveWindow.Zoom = True"
' Line #712:
' 	QuoteRem 0x0004 0x0025 "ActiveSheet.Protect Password:="spike""
' Line #713:
' 	QuoteRem 0x0004 0x002A "ActiveWindow.ScrollWorkbookTabs Sheets:=-3"
' Line #714:
' 	QuoteRem 0x0004 0x0028 "ActiveWorkbook.Protect Password:="spike""
' Line #715:
' 	QuoteRem 0x0004 0x0015 "Application.Calculate"
' Line #716:
' 	EndSub 
' Line #717:
' 	FuncDefn (Sub NEWHilo())
' Line #718:
' 	QuoteRem 0x0000 0x0000 ""
' Line #719:
' 	QuoteRem 0x0000 0x002E " JLR 4/9/14; amended 9/5/2015 for GPS altitude"
' Line #720:
' 	QuoteRem 0x0000 0x0000 ""
' Line #721:
' 	LitStr 0x0007 "Ab.xlsm"
' 	ArgsLd Workbooks 0x0001 
' 	ArgsMemCall Activate 0x0000 
' Line #722:
' 	LitStr 0x0006 "Sheet2"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemCall Select 0x0000 
' Line #723:
' 	LitStr 0x000B "A.xlsm!PreB"
' 	Ld Application 
' 	ArgsMemCall Run 0x0001 
' Line #724:
' 	LitStr 0x0010 "Ab.xlsm!Duration"
' 	Ld Application 
' 	ArgsMemCall Run 0x0001 
' Line #725:
' 	LitStr 0x0006 "Sheet1"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemCall Select 0x0000 
' Line #726:
' 	LitStr 0x000A "A1:CK10010"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x000A "A1:CK10010"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #727:
' 	Ld xlCalculationManual 
' 	Ld Application 
' 	MemSt Calculation 
' Line #728:
' 	LitStr 0x0002 "L5"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0000 ""
' 	Eq 
' 	IfBlock 
' Line #729:
' 	LineCont 0x0004 07 00 08 00
' 	LitStr 0x007D "=IF(R[1]C[-54]="","",IF(AND(RC[-39]>R2C28,R[1]C[-44]>RC[-44],R[2]C[-44]>RC[-44],R[3]C[-44]>RC[-44],R[4]C[-44]>RC[-44]),1,""))"
' 	LitStr 0x000C "BC11:BC10010"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #730:
' 	LineCont 0x0004 07 00 08 00
' 	LitStr 0x003C "=IF(AND(RC[-40]>R2C28,OR(R2C40=0,RC[-40]<R2C40)),RC[-45],"")"
' 	LitStr 0x000C "BD11:BD10010"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #731:
' 	LineCont 0x0004 07 00 08 00
' 	LitStr 0x0048 "=IF(R2C40=0,MAX(R[7]C[-46]:R[10005]C[-46]),MAX(R[7]C[-1]:R[10005]C[-1]))"
' 	LitStr 0x0003 "BE4"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #732:
' 	LitStr 0x002E "=IF(RC[-1]="","",IF(RC[-46]=R4C57,RC[-41],""))"
' 	LitStr 0x000C "BE11:BE10010"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #733:
' 	LitStr 0x0002 "L5"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0000 ""
' 	Ne 
' 	ElseIfBlock 
' Line #734:
' 	LineCont 0x0004 07 00 08 00
' 	LitStr 0x007D "=IF(R[1]C[-54]="","",IF(AND(RC[-39]>R2C28,R[1]C[-43]>RC[-43],R[2]C[-43]>RC[-43],R[3]C[-43]>RC[-43],R[4]C[-43]>RC[-43]),1,""))"
' 	LitStr 0x000C "BC11:BC10010"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #735:
' 	LineCont 0x0004 07 00 08 00
' 	LitStr 0x003C "=IF(AND(RC[-40]>R2C28,OR(R2C40=0,RC[-40]<R2C40)),RC[-44],"")"
' 	LitStr 0x000C "BD11:BD10010"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #736:
' 	LineCont 0x0004 07 00 08 00
' 	LitStr 0x0048 "=IF(R2C40=0,MAX(R[7]C[-45]:R[10005]C[-45]),MAX(R[7]C[-1]:R[10005]C[-1]))"
' 	LitStr 0x0003 "BE4"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #737:
' 	LitStr 0x002E "=IF(RC[-1]="","",IF(RC[-45]=R4C57,RC[-41],""))"
' 	LitStr 0x000C "BE11:BE10010"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #738:
' 	EndIfBlock 
' Line #739:
' 	LitStr 0x001B "=MAX(R[7]C[1]:R[10005]C[1])"
' 	LitStr 0x0003 "BD4"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #740:
' 	LitStr 0x0002 "L5"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0000 ""
' 	Eq 
' 	IfBlock 
' Line #741:
' 	LitStr 0x0031 "=IF(AND(RC[-43]>=R2C28,RC[-43]<R4C56),RC[-48],"")"
' 	LitStr 0x000C "BG11:BG10010"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #742:
' 	LitStr 0x0002 "L5"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0000 ""
' 	Ne 
' 	ElseIfBlock 
' Line #743:
' 	LitStr 0x003C "=IF(AND(RC[-43]>=R2C28,RC[-43]<R4C56,RC[-47]<>0),RC[-47],"")"
' 	LitStr 0x000C "BG11:BG10010"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #744:
' 	EndIfBlock 
' Line #745:
' 	LitStr 0x001D "=MIN(R[7]C[-1]:R[10005]C[-1])"
' 	LitStr 0x0003 "BH4"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #746:
' 	LitStr 0x001A "=IF(RC[-1]=R4C,RC[-44],"")"
' 	LitStr 0x000C "BH11:BH10010"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #747:
' 	LitStr 0x001B "=MIN(R[7]C[1]:R[10005]C[1])"
' 	LitStr 0x0003 "BG4"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #748:
' 	LitStr 0x0015 "=R[-1]C[-1]-R[-1]C[2]"
' 	LitStr 0x0003 "BF5"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #749:
' 	LitStr 0x0002 "L5"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0000 ""
' 	Eq 
' 	IfBlock 
' Line #750:
' 	LineCont 0x0004 07 00 08 00
' 	LitStr 0x0084 "=IF(OR(RC[-46]="",AND(R2C40>0,RC[-46]>=R2C40)),"",IF(AND(RC[-46]>R4C56,RC[-7]=1,RC[-51]<R[-1]C[-51],RC[-51]<R[1]C[-51]),RC[-51],""))"
' 	LitStr 0x000C "BJ11:BJ10010"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #751:
' 	LineCont 0x0004 07 00 08 00
' 	LitStr 0x0053 "=IF(AND(R2C40>0,RC[-47]>=R2C40),"",IF(AND(RC[-47]>R8C64,RC[-52]>R7C62),RC[-52],""))"
' 	LitStr 0x000C "BK11:BK10010"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #752:
' 	LitStr 0x0030 "=IF(AND(RC[-48]>R8C64,RC[-53]=R7C63),RC[-48],"")"
' 	LitStr 0x000C "BL11:BL10010"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #753:
' 	LitStr 0x0032 "=IF(AND(RC[-49]>R4C56,RC[-54]=R7C[-1]),RC[-49],"")"
' 	LitStr 0x000C "BM11:BM10010"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #754:
' 	LineCont 0x0004 07 00 08 00
' 	LitStr 0x009C "=IF(OR(RC[-50]="",AND(R2C40>0,RC[-50]>=R2C40)),"",IF(AND(RC[-50]>R4C56,RC[-50]>R8C63,RC[-55]<R[-1]C[-55],RC[-55]<R[1]C[-55],RC[-55]<R[2]C[-55]),RC[-55],""))"
' 	LitStr 0x000C "BN11:BN10010"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #755:
' 	LitStr 0x002E "=IF(RC[-1]="","",IF(RC[-56]=R7C67,RC[-51],""))"
' 	LitStr 0x000C "BO11:BO10010"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #756:
' 	LitStr 0x001D "=IF(RC[-52]>R8C67,RC[-57],"")"
' 	LitStr 0x000C "BP11:BP10010"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #757:
' 	LitStr 0x0002 "L5"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0000 ""
' 	Ne 
' 	ElseIfBlock 
' Line #758:
' 	LineCont 0x0004 07 00 08 00
' 	LitStr 0x008E "=IF(OR(RC[-46]="",RC[-50]=0,AND(R2C40>0,RC[-46]>=R2C40)),"",IF(AND(RC[-46]>R4C56,RC[-7]=1,RC[-50]<R[-1]C[-50],RC[-50]<R[1]C[-50]),RC[-50],""))"
' 	LitStr 0x000C "BJ11:BJ10010"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #759:
' 	LineCont 0x0004 07 00 08 00
' 	LitStr 0x0053 "=IF(AND(R2C40>0,RC[-47]>=R2C40),"",IF(AND(RC[-47]>R8C64,RC[-51]>R7C62),RC[-51],""))"
' 	LitStr 0x000C "BK11:BK10010"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #760:
' 	LitStr 0x0030 "=IF(AND(RC[-48]>R8C64,RC[-52]=R7C63),RC[-48],"")"
' 	LitStr 0x000C "BL11:BL10010"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #761:
' 	LitStr 0x0032 "=IF(AND(RC[-49]>R4C56,RC[-53]=R7C[-1]),RC[-49],"")"
' 	LitStr 0x000C "BM11:BM10010"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #762:
' 	LineCont 0x0004 07 00 08 00
' 	LitStr 0x00A6 "=IF(OR(RC[-50]="",RC[-54]=0,AND(R2C40>0,RC[-50]>=R2C40)),"",IF(AND(RC[-50]>R4C56,RC[-50]>R8C63,RC[-54]<R[-1]C[-54],RC[-54]<R[1]C[-54],RC[-54]<R[2]C[-54]),RC[-54],""))"
' 	LitStr 0x000C "BN11:BN10010"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #763:
' 	LitStr 0x002E "=IF(RC[-1]="","",IF(RC[-55]=R7C67,RC[-51],""))"
' 	LitStr 0x000C "BO11:BO10010"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #764:
' 	LitStr 0x001D "=IF(RC[-52]>R8C67,RC[-56],"")"
' 	LitStr 0x000C "BP11:BP10010"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #765:
' 	EndIfBlock 
' Line #766:
' 	LitStr 0x0015 "=MIN(R[4]C:R[10993]C)"
' 	LitStr 0x0003 "BJ7"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #767:
' 	LitStr 0x0022 "=IF(R[2]C[-1]=0,"",R[2]C-R[2]C[1])"
' 	LitStr 0x0003 "BK5"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #768:
' 	LitStr 0x0015 "=MAX(R[4]C:R[10002]C)"
' 	LitStr 0x0003 "BK7"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #769:
' 	LitStr 0x001B "=MAX(R[3]C[1]:R[10001]C[1])"
' 	LitStr 0x0003 "BK8"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #770:
' 	LitStr 0x0007 "=RC[-2]"
' 	LitStr 0x0003 "BL7"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #771:
' 	LitStr 0x001B "=MIN(R[3]C[1]:R[10001]C[1])"
' 	LitStr 0x0003 "BL8"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #772:
' 	LitStr 0x001C "=MIN(R[4]C[-1]:R[9993]C[-1])"
' 	LitStr 0x0003 "BO7"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #773:
' 	LitStr 0x0014 "=MAX(R[3]C:R[9992]C)"
' 	LitStr 0x0003 "BO8"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #774:
' 	LitStr 0x0014 "=MAX(R[4]C:R[9993]C)"
' 	LitStr 0x0003 "BP7"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #775:
' 	LitStr 0x000F "=PRS!R[5]C[-62]"
' 	LitStr 0x0003 "BQ1"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #776:
' 	LitStr 0x000F "=PRS!R[2]C[-65]"
' 	LitStr 0x0003 "BQ2"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #777:
' 	LitStr 0x000F "=PRS!R[7]C[-62]"
' 	LitStr 0x0003 "BQ3"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #778:
' 	LitStr 0x001C "=IF(RC[-1]=R7C68,RC[-53],"")"
' 	LitStr 0x000C "BQ11:BQ10010"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #779:
' 	LitStr 0x0002 "L5"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0000 ""
' 	Eq 
' 	IfBlock 
' Line #780:
' 	LineCont 0x0004 07 00 08 00
' 	LitStr 0x0110 "=IF(OR(AND(R1C[-1]=0,RC[-54]>R2C[-1],RC[-54]<R3C[-1]),AND(RC[-54]>R2C[-1],RC[-54]<MIN(R1C[-1],R3C[-1]))),6371*ACOS(SIN(R2C)*SIN((RC[-66]+(RC[-65]*0.001)/60)*PI()/180)+COS(R2C)*COS((RC[-66]+(RC[-65]*0.001)/60)*PI()/180)*COS(R3C-((RC[-63]+(RC[-62]*0.001)/60)*PI()/180))),"")"
' 	LitStr 0x000C "BR11:BR10010"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #781:
' 	LitStr 0x0002 "L5"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0000 ""
' 	Ne 
' 	ElseIfBlock 
' Line #782:
' 	LineCont 0x0004 07 00 08 00
' 	LitStr 0x0121 "=IF(RC[-58]=0,"",IF(OR(AND(R1C[-1]=0,RC[-54]>R2C[-1],RC[-54]<R3C[-1]),AND(RC[-54]>R2C[-1],RC[-54]<MIN(R1C[-1],R3C[-1]))),6371*ACOS(SIN(R2C)*SIN((RC[-66]+(RC[-65]*0.001)/60)*PI()/180)+COS(R2C)*COS((RC[-66]+(RC[-65]*0.001)/60)*PI()/180)*COS(R3C-((RC[-63]+(RC[-62]*0.001)/60)*PI()/180))),""))"
' 	LitStr 0x000C "BR11:BR10010"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #783:
' 	EndIfBlock 
' Line #784:
' 	LitStr 0x001A "=MAX(R[3]C[1]:R[9992]C[1])"
' 	LitStr 0x0003 "BP8"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #785:
' 	LitStr 0x002C "=(PRS!R[2]C[-64]+PRS!R[2]C[-63]/60)*PI()/180"
' 	LitStr 0x0003 "BR2"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #786:
' 	LitStr 0x002C "=(PRS!R[2]C[-64]+PRS!R[2]C[-63]/60)*PI()/180"
' 	LitStr 0x0003 "BR3"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #787:
' 	LitStr 0x0015 "=MAX(R[3]C:R[10002]C)"
' 	LitStr 0x0003 "BR8"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #788:
' 	LitStr 0x001E "=IF(RC[-1]=R8C[-1],RC[-55],"")"
' 	LitStr 0x000C "BS11:BS10010"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #789:
' 	LitStr 0x0015 "=MAX(R[2]C:R[10001]C)"
' 	LitStr 0x0003 "BS9"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #790:
' 	LitStr 0x002E "=IF(RC[-1]="","",RC[-68]+((RC[-67]*0.001)/60))"
' 	LitStr 0x000C "BT11:BT10010"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #791:
' 	LitStr 0x0015 "=MAX(R[2]C:R[10001]C)"
' 	LitStr 0x0003 "BT9"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #792:
' 	LitStr 0x002E "=IF(RC[-1]="","",RC[-66]+((RC[-65]*0.001)/60))"
' 	LitStr 0x000C "BU11:BU10010"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #793:
' 	LitStr 0x0015 "=MAX(R[2]C:R[10001]C)"
' 	LitStr 0x0003 "BU9"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #794:
' 	LitStr 0x0002 "L5"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0000 ""
' 	Eq 
' 	IfBlock 
' Line #795:
' 	LitStr 0x0019 "=IF(RC[-1]="","",RC[-63])"
' 	LitStr 0x000C "BV11:BV10010"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #796:
' 	LitStr 0x0002 "L5"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0000 ""
' 	Ne 
' 	ElseIfBlock 
' Line #797:
' 	LitStr 0x0019 "=IF(RC[-1]="","",RC[-62])"
' 	LitStr 0x000C "BV11:BV10010"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #798:
' 	EndIfBlock 
' Line #799:
' 	LitStr 0x0015 "=MAX(R[2]C:R[10001]C)"
' 	LitStr 0x0003 "BV9"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #800:
' 	LitStr 0x002E "=(PRS!R[18]C[-74]+PRS!R[18]C[-73]/60)*PI()/180"
' 	LitStr 0x0003 "BW2"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #801:
' 	LitStr 0x002E "=(PRS!R[17]C[-71]+PRS!R[17]C[-70]/60)*PI()/180"
' 	LitStr 0x0003 "BW3"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #802:
' 	LitStr 0x0002 "L5"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0000 ""
' 	Eq 
' 	IfBlock 
' Line #803:
' 	LineCont 0x0004 07 00 08 00
' 	LitStr 0x0189 "=IF(SUM(R2C:R3C)=0,"",IF(AND(((RC[-71]+(RC[-70]*0.001)/60)*PI()/180)=R2C,((RC[-68]+(RC[-67]*0.001)/60)*PI()/180)=R3C),0,IF(OR(AND(R1C[-6]=0,RC[-59]>R2C[-6],RC[-59]<R3C[-6]),AND(RC[-59]>R2C[-6],RC[-59]<MIN(R1C[-6],R3C[-6]))),6371*ACOS(SIN(R2C)*SIN((RC[-71]+(RC[-70]*0.001)/60)*PI()/180)+COS(R2C)*COS((RC[-71]+(RC[-70]*0.001)/60)*PI()/180)*COS(R3C-((RC[-68]+(RC[-67]*0.001)/60)*PI()/180))),"")))"
' 	LitStr 0x000C "BW11:BW10010"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #804:
' 	LitStr 0x0002 "L5"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0000 ""
' 	Ne 
' 	ElseIfBlock 
' Line #805:
' 	LineCont 0x0004 07 00 08 00
' 	LitStr 0x0197 "=IF(OR(SUM(R2C:R3C)=0,RC[-63]=0),"",IF(AND(((RC[-71]+(RC[-70]*0.001)/60)*PI()/180)=R2C,((RC[-68]+(RC[-67]*0.001)/60)*PI()/180)=R3C),0,IF(OR(AND(R1C[-6]=0,RC[-59]>R2C[-6],RC[-59]<R3C[-6]),AND(RC[-59]>R2C[-6],RC[-59]<MIN(R1C[-6],R3C[-6]))),6371*ACOS(SIN(R2C)*SIN((RC[-71]+(RC[-70]*0.001)/60)*PI()/180)+COS(R2C)*COS((RC[-71]+(RC[-70]*0.001)/60)*PI()/180)*COS(R3C-((RC[-68]+(RC[-67]*0.001)/60)*PI()/180))),"")))"
' 	LitStr 0x000C "BW11:BW10010"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #806:
' 	EndIfBlock 
' Line #807:
' 	LitStr 0x0015 "=MAX(R[3]C:R[10002]C)"
' 	LitStr 0x0003 "BW8"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #808:
' 	LitStr 0x001E "=IF(RC[-1]=R8C[-1],RC[-60],"")"
' 	LitStr 0x000C "BX11:BX10010"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #809:
' 	LitStr 0x0015 "=MAX(R[2]C:R[10001]C)"
' 	LitStr 0x0003 "BX9"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #810:
' 	LitStr 0x002E "=IF(RC[-1]="","",RC[-73]+((RC[-72]*0.001)/60))"
' 	LitStr 0x000C "BY11:BY10010"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #811:
' 	LitStr 0x0015 "=MAX(R[2]C:R[10001]C)"
' 	LitStr 0x0003 "BY9"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #812:
' 	LitStr 0x002E "=IF(RC[-1]="","",RC[-71]+((RC[-70]*0.001)/60))"
' 	LitStr 0x000C "BZ11:BZ10010"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #813:
' 	LitStr 0x0015 "=MAX(R[2]C:R[10001]C)"
' 	LitStr 0x0003 "BZ9"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #814:
' 	LitStr 0x0002 "L5"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0000 ""
' 	Ne 
' 	IfBlock 
' Line #815:
' 	LitStr 0x0019 "=IF(RC[-1]="","",RC[-68])"
' 	LitStr 0x000C "CA11:CA10010"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #816:
' 	LineCont 0x0004 07 00 08 00
' 	LitStr 0x006A "=IF(SUM(RC[-80],RC[-79],RC[-78])=0,"",IF(AND(RC[-70]>PRS!R8C[-77],RC[-65]-PRS!R4C[-77]>=5/24),RC[-70],""))"
' 	LitStr 0x000C "CC11:CC10010"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #817:
' 	LitStr 0x0002 "L5"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0000 ""
' 	Eq 
' 	ElseIfBlock 
' Line #818:
' 	LitStr 0x0019 "=IF(RC[-1]="","",RC[-67])"
' 	LitStr 0x000C "CA11:CA10010"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #819:
' 	LineCont 0x0004 07 00 08 00
' 	LitStr 0x006A "=IF(SUM(RC[-80],RC[-79],RC[-78])=0,"",IF(AND(RC[-69]>PRS!R8C[-77],RC[-65]-PRS!R4C[-77]>=5/24),RC[-69],""))"
' 	LitStr 0x000C "CC11:CC10010"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #820:
' 	EndIfBlock 
' Line #821:
' 	LitStr 0x0015 "=MAX(R[2]C:R[10001]C)"
' 	LitStr 0x0003 "CA9"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #822:
' 	LitStr 0x0015 "=MAX(R[2]C:R[10001]C)"
' 	LitStr 0x0003 "CC9"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #823:
' 	LitStr 0x001E "=IF(RC[-1]=R9C[-1],RC[-66],"")"
' 	LitStr 0x000C "CD11:CD10010"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #824:
' 	LitStr 0x0020 "=IF(R[5]C=0,0,R[5]C-PRS!RC[-78])"
' 	LitStr 0x0003 "CD4"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #825:
' 	LitStr 0x0015 "=MAX(R[2]C:R[10001]C)"
' 	LitStr 0x0003 "CD9"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #826:
' 	LineCont 0x0004 07 00 08 00
' 	LitStr 0x0075 "=IF(MAX(R[3]C[2],R[3]C[7],R[3]C[11])=R[3]C[2],R[2]C,IF(MAX(R[3]C[2],R[3]C[7],R[3]C[11])=R[3]C[7],R[6]C[7],R[6]C[12]))"
' 	LitStr 0x0003 "BD2"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #827:
' 	LineCont 0x0004 07 00 08 00
' 	LitStr 0x0075 "=IF(MAX(R[3]C[1],R[3]C[6],R[3]C[10])=R[3]C[1],R[2]C,IF(MAX(R[3]C[1],R[3]C[6],R[3]C[10])=R[3]C[6],R[5]C[6],R[5]C[11]))"
' 	LitStr 0x0003 "BE2"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #828:
' 	LineCont 0x0004 07 00 08 00
' 	LitStr 0x0075 "=IF(MAX(R[3]C[-1],R[3]C[4],R[3]C[8])=R[3]C[-1],R[2]C,IF(MAX(R[3]C[-1],R[3]C[4],R[3]C[8])=R[3]C[4],R[6]C[5],R[6]C[8]))"
' 	LitStr 0x0003 "BG2"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #829:
' 	LineCont 0x0004 07 00 08 00
' 	LitStr 0x0075 "=IF(MAX(R[3]C[-2],R[3]C[3],R[3]C[7])=R[3]C[-2],R[2]C,IF(MAX(R[3]C[-2],R[3]C[3],R[3]C[7])=R[3]C[3],R[5]C[4],R[5]C[7]))"
' 	LitStr 0x0003 "BH2"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #830:
' 	Ld xlCalculationAutomatic 
' 	Ld Application 
' 	MemSt Calculation 
' Line #831:
' 	LitStr 0x0007 "A1:CK10"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0007 "A1:CK10"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #832:
' 	LitStr 0x000B "Q11:CK10010"
' 	ArgsLd Range 0x0001 
' 	ArgsMemCall Clear 0x0000 
' Line #833:
' Line #834:
' 	LitStr 0x001B "=IF(RC[-40]=R4C,RC[-44],"")"
' 	LitStr 0x000C "BD11:BD10010"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #835:
' 	LitStr 0x001F "=IF(RC[-41]=R2C[-1],RC[-45],"")"
' 	LitStr 0x000C "BE11:BE10010"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #836:
' 	LitStr 0x001B "=IF(RC[-43]=R2C,RC[-47],"")"
' 	LitStr 0x000C "BG11:BG10010"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #837:
' Line #838:
' 	LitStr 0x001D "=MAX(R[9]C[-1]:R[10008]C[-1])"
' 	LitStr 0x0003 "BF2"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #839:
' 	LitStr 0x001D "=MAX(R[7]C[-2]:R[10006]C[-2])"
' 	LitStr 0x0003 "BF4"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #840:
' 	LitStr 0x001D "=MAX(R[9]C[-2]:R[10008]C[-2])"
' 	LitStr 0x0003 "BI2"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #841:
' 	LitStr 0x0007 "BF2:BI4"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0007 "BF2:BI4"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #842:
' 	LitStr 0x000B "A11:BG10010"
' 	ArgsLd Range 0x0001 
' 	ArgsMemCall Clear 0x0000 
' Line #843:
' Line #844:
' 	LitStr 0x0003 "BF4"
' 	LitStr 0x0006 "Sheet1"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0002 "J2"
' 	LitStr 0x0003 "PRS"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemSt Value 
' Line #845:
' 	LitStr 0x0003 "BF2"
' 	LitStr 0x0006 "Sheet1"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0002 "J5"
' 	LitStr 0x0003 "PRS"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemSt Value 
' Line #846:
' 	LitStr 0x0003 "BI2"
' 	LitStr 0x0006 "Sheet1"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0002 "J8"
' 	LitStr 0x0003 "PRS"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemSt Value 
' Line #847:
' Line #848:
' 	QuoteRem 0x0004 0x0027 "Added 10/16/2015 for ST@TP or goofy dec"
' Line #849:
' 	QuoteRem 0x0004 0x0016 "Sheets("PRS").Activate"
' Line #850:
' 	LineCont 0x0004 01 00 E6 FF
' 	QuoteRem 0x0004 0x011A "Range("J24").FormulaR1C1 =        "=IF(R[-10]C[-8]<>3,"""",IF(OR(AND(R[-12]C[-5]=""ST@TP"",OR(RC[1]<>R[-4]C[1],RC[2]<>R[-4]C[2])),AND(RC[1]=R[-2]C[1],RC[2]=R[-2]C[2],RC[1]<>R[2]C[1],RC[2]<>R[2]C[2]),AND(R[2]C[1]=R[4]C[1],R[2]C[2]=R[4]C[2],RC[1]<>R[4]C[1],RC[2]<>R[4]C[2])),3,""""))""
' Line #851:
' 	QuoteRem 0x0004 0x009C "Range("J26").FormulaR1C1 = "=IF(OR(R[-2]C=3,AND(RC[1]=R[2]C[1],RC[2]=R[2]C[2],RC[1]<>R[-4]C[1],RC[2]<>R[-4]C[2],R[-2]C[1]<>RC[1],R[-2]C[2]<>RC[2])),2,"""")""
' Line #852:
' 	QuoteRem 0x0004 0x002F "Range("J24:J26").Value = Range("J24:J26").Value"
' Line #853:
' 	QuoteRem 0x0004 0x002C "If Range("J24") = 3 Or Range("J26") = 2 Then"
' Line #854:
' 	QuoteRem 0x0008 0x002F "Range("A26:G26").Value = Range("A24:G24").Value"
' Line #855:
' 	QuoteRem 0x0008 0x002F "Range("A24:G24").Value = Range("A20:G20").Value"
' Line #856:
' 	QuoteRem 0x0004 0x0006 "End If"
' Line #857:
' 	LitStr 0x0011 "Ab.xlsm!Detangler"
' 	Ld Application 
' 	ArgsMemCall Run 0x0001 
' Line #858:
' Line #859:
' 	EndSub 
' Line #860:
' 	FuncDefn (Sub Detangler())
' Line #861:
' 	QuoteRem 0x0000 0x0000 ""
' Line #862:
' 	QuoteRem 0x0000 0x0074 " Determines what to do with bungled declarations; E12 value must be calculated here!ELSEWHERE, Ck dependents on E12!"
' Line #863:
' 	QuoteRem 0x0000 0x0060 " 12/16/15: Use K22:K26 for TP Order?(w/names @ S22:S26) NEED SOMETHING QUICK WHEN B14=0 or B14=1"
' Line #864:
' 	QuoteRem 0x0000 0x0000 ""
' Line #865:
' 	QuoteRem 0x0004 0x004D "Step 1: Do nothing if declared order is plausible, with or without duplicates"
' Line #866:
' 	QuoteRem 0x0004 0x002F "Range("T20:T28").Value = Range("L20:L28").Value"
' Line #867:
' Line #868:
' 	LitStr 0x0003 "PRS"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemCall Activate 0x0000 
' Line #869:
' 	LineCont 0x0004 07 00 08 00
' 	LitStr 0x0280 "=IF(AND(OR(R20C11<>R28C11,R20C12<>R28C12),OR(AND(R20C11=R22C11,R20C12=R22C12),AND(R20C11=R24C11,R20C12=R24C12),AND(R20C11=R26C11,R20C12=R26C12)),OR(AND(R28C11=R26C11,R28C12=R26C12),AND(R28C11=R24C11,R28C12=R24C12),AND(R28C11=R22C11,R28C12=R22C12))),"S&F@TPS",IF(AND(R20C11=R28C11,R20C12=R28C12,OR(AND(R20C11=R22C11,R20C12=R22C12),AND(R20C11=R24C11,R20C12=R24C12),AND(R20C11=R26C11,R20C12=R26C12))),"SF@TP",IF(OR(AND(R20C11=R22C11,R20C12=R22C12),AND(R20C11=R24C11,R20C12=R24C12),AND(R20C11=R26C11,R20C12=R26C12)),"ST@TP",IF(OR(AND(R28C11=R26C11,R28C12=R26C12),AND(R28C11=R24C11,R28C12=R24C12),AND(R28C11=R22C11,R28C12=R22C12)),"FI@TP",""))))"
' 	LitStr 0x0003 "E12"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #870:
' 	LitStr 0x0003 "B14"
' 	ArgsLd Range 0x0001 
' 	LitDI2 0x0001 
' 	Gt 
' 	IfBlock 
' Line #871:
' 	LineCont 0x0004 07 00 0C 00
' 	LitStr 0x0194 "=IF(AND(R12C5="",OR(AND(R14C2=3,R22C11=R24C11,R24C11=R26C11,R22C12=R24C12,R24C12=R26C12),AND(R14C2=2,R22C11=R24C11,R22C12=R24C12))),"1 TP",IF(R28C10=0,0,IF(AND(R14C2=2,R28C10<3),"1 TP",IF(AND(R14C2=2,R28C=3),"2 TP",IF(AND(R14C2=3,OR(R28C10<3,AND(R22C10="X",R24C10="X"),AND(R24C10="X",R26C10="X"),AND(R22C10="X",R26C10="X"))),"1 TP",IF(AND(R14C2=3,OR(R22C10="X",R24C10="X",R26C10="X")),"2 TP","3 TP"))))))"
' 	LitStr 0x0003 "J20"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #872:
' 	Ld ActiveSheet 
' 	ArgsMemCall Calculate 0x0000 
' Line #873:
' Line #874:
' 	LineCont 0x0004 07 00 0C 00
' 	LitStr 0x0059 "=IF(AND(OR(RC[1]<>R[-2]C[1],RC[2]<>R[-2]C[2]),OR(RC[1]<>R[2]C[1],RC[2]<>R[2]C[2])),1,"X")"
' 	LitStr 0x0003 "J22"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #875:
' 	LitStr 0x0003 "B14"
' 	ArgsLd Range 0x0001 
' 	LitDI2 0x0003 
' 	Eq 
' 	IfBlock 
' Line #876:
' 	LineCont 0x0004 07 00 10 00
' 	LitStr 0x0059 "=IF(AND(OR(RC[1]<>R[-2]C[1],RC[2]<>R[-2]C[2]),OR(RC[1]<>R[2]C[1],RC[2]<>R[2]C[2])),2,"X")"
' 	LitStr 0x0003 "J24"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #877:
' 	LineCont 0x0004 07 00 10 00
' 	LitStr 0x0059 "=IF(AND(OR(RC[1]<>R[-2]C[1],RC[2]<>R[-2]C[2]),OR(RC[1]<>R[2]C[1],RC[2]<>R[2]C[2])),3,"X")"
' 	LitStr 0x0003 "J26"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #878:
' 	LitStr 0x0003 "B14"
' 	ArgsLd Range 0x0001 
' 	LitDI2 0x0002 
' 	Eq 
' 	ElseIfBlock 
' Line #879:
' 	LineCont 0x0004 07 00 10 00
' 	LitStr 0x0059 "=IF(AND(OR(RC[1]<>R[-2]C[1],RC[2]<>R[-2]C[2]),OR(RC[1]<>R[4]C[1],RC[2]<>R[4]C[2])),2,"X")"
' 	LitStr 0x0003 "J24"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #880:
' 	EndIfBlock 
' Line #881:
' 	LitStr 0x0003 "B14"
' 	ArgsLd Range 0x0001 
' 	LitDI2 0x0001 
' 	Eq 
' 	ElseIfBlock 
' Line #882:
' 	LineCont 0x0004 07 00 0C 00
' 	LitStr 0x0059 "=IF(AND(OR(RC[1]<>R[-2]C[1],RC[2]<>R[-2]C[2]),OR(RC[1]<>R[6]C[1],RC[2]<>R[6]C[2])),1,"X")"
' 	LitStr 0x0003 "J22"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #883:
' 	LitStr 0x0003 "J22"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0001 "X"
' 	Eq 
' 	IfBlock 
' Line #884:
' 	LitDI2 0x0000 
' 	LitStr 0x0003 "I22"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #885:
' 	LitStr 0x0003 "TPO"
' 	LitStr 0x0003 "J20"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #886:
' 	EndIfBlock 
' Line #887:
' 	EndIfBlock 
' Line #888:
' Line #889:
' 	LitStr 0x000F "=SUM(R22C:R26C)"
' 	LitStr 0x0003 "J28"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #890:
' Line #891:
' 	LitStr 0x0003 "B14"
' 	ArgsLd Range 0x0001 
' 	LitDI2 0x0002 
' 	Eq 
' 	LitStr 0x0003 "J28"
' 	ArgsLd Range 0x0001 
' 	LitDI2 0x0003 
' 	Eq 
' 	And 
' 	IfBlock 
' Line #892:
' 	LitStr 0x0007 "J20:J28"
' 	ArgsLd Range 0x0001 
' 	ArgsMemCall Clear 0x0000 
' Line #893:
' 	LitDI2 0x0001 
' 	LitStr 0x0003 "I22"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #894:
' 	LitDI2 0x0002 
' 	LitStr 0x0003 "I24"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #895:
' 	LitStr 0x0007 "A20:G28"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0007 "M20:S28"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #896:
' 	LitStr 0x0003 "B14"
' 	ArgsLd Range 0x0001 
' 	LitDI2 0x0003 
' 	Eq 
' 	LitStr 0x0003 "J28"
' 	ArgsLd Range 0x0001 
' 	LitDI2 0x0006 
' 	Eq 
' 	And 
' 	ElseIfBlock 
' Line #897:
' 	LitStr 0x0007 "J20:J28"
' 	ArgsLd Range 0x0001 
' 	ArgsMemCall Clear 0x0000 
' Line #898:
' 	LitDI2 0x0001 
' 	LitStr 0x0003 "I22"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #899:
' 	LitDI2 0x0002 
' 	LitStr 0x0003 "I24"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #900:
' 	LitDI2 0x0003 
' 	LitStr 0x0003 "I26"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #901:
' 	LitStr 0x0007 "A20:G28"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0007 "M20:S28"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #902:
' 	ExitSub 
' Line #903:
' 	EndIfBlock 
' Line #904:
' Line #905:
' 	QuoteRem 0x0004 0x0027 "Step 2a:Bungled without ST or Fin as TP"
' Line #906:
' 	LitStr 0x0003 "E12"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0000 ""
' 	Eq 
' 	IfBlock 
' Line #907:
' 	QuoteRem 0x0005 0x0022 "When B14 = 2 here - has to be 1 TP"
' Line #908:
' 	LitStr 0x0003 "B14"
' 	ArgsLd Range 0x0001 
' 	LitDI2 0x0003 
' 	Eq 
' 	LitStr 0x0003 "J20"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0004 "1 TP"
' 	Eq 
' 	And 
' 	IfBlock 
' Line #909:
' 	LitStr 0x0003 "J22"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0001 "X"
' 	Eq 
' 	LitStr 0x0003 "J24"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0001 "X"
' 	Eq 
' 	And 
' 	LitStr 0x0003 "J26"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0001 "X"
' 	Eq 
' 	And 
' 	IfBlock 
' Line #910:
' 	LitDI2 0x0001 
' 	LitStr 0x0003 "J22"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #911:
' 	LitStr 0x0001 "X"
' 	LitStr 0x0007 "J24,J26"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #912:
' 	LitStr 0x0003 "J22"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0001 "X"
' 	Eq 
' 	LitStr 0x0003 "J24"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0001 "X"
' 	Eq 
' 	And 
' 	ElseIfBlock 
' Line #913:
' 	LitDI2 0x0001 
' 	LitStr 0x0003 "J22"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #914:
' 	LitDI2 0x0003 
' 	LitStr 0x0003 "J24"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #915:
' 	LitDI2 0x0002 
' 	LitStr 0x0003 "J26"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #916:
' 	LitStr 0x0003 "J24"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0001 "X"
' 	Eq 
' 	LitStr 0x0003 "J26"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0001 "X"
' 	Eq 
' 	And 
' 	ElseIfBlock 
' Line #917:
' 	LitStr 0x0003 "K22"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0003 "K24"
' 	ArgsLd Range 0x0001 
' 	Ne 
' 	LitStr 0x0003 "L22"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0003 "L24"
' 	ArgsLd Range 0x0001 
' 	Ne 
' 	Or 
' 	IfBlock 
' Line #918:
' 	LitDI2 0x0002 
' 	LitStr 0x0003 "J22"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #919:
' 	LitDI2 0x0001 
' 	LitStr 0x0003 "J24"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #920:
' 	LitDI2 0x0003 
' 	LitStr 0x0003 "J26"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #921:
' 	EndIfBlock 
' Line #922:
' 	LitStr 0x0003 "J24"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0001 "X"
' 	Eq 
' 	LitStr 0x0003 "J26"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0001 "X"
' 	Eq 
' 	And 
' 	ElseIfBlock 
' Line #923:
' 	LitDI2 0x0001 
' 	LitStr 0x0003 "J22"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #924:
' 	LitDI2 0x0002 
' 	LitStr 0x0003 "J24"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #925:
' 	EndIfBlock 
' Line #926:
' 	LitStr 0x0003 "B14"
' 	ArgsLd Range 0x0001 
' 	LitDI2 0x0002 
' 	Eq 
' 	LitStr 0x0003 "J20"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0004 "1 TP"
' 	Eq 
' 	And 
' 	ElseIfBlock 
' Line #927:
' 	LitStr 0x0003 "J28"
' 	ArgsLd Range 0x0001 
' 	LitDI2 0x0001 
' 	Eq 
' 	IfBlock 
' Line #928:
' 	LitDI2 0x0001 
' 	LitStr 0x0003 "J22"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #929:
' 	LitStr 0x0003 "J28"
' 	ArgsLd Range 0x0001 
' 	LitDI2 0x0002 
' 	Eq 
' 	ElseIfBlock 
' Line #930:
' 	LitDI2 0x0001 
' 	LitStr 0x0003 "J24"
' 	ArgsSt Range 0x0001 
' Line #931:
' 	EndIfBlock 
' Line #932:
' 	EndIfBlock 
' Line #933:
' 	EndIfBlock 
' Line #934:
' Line #935:
' 	QuoteRem 0x0004 0x002F "Step 2b: Bungled TP with 2 declared OK 12/14/15"
' Line #936:
' 	LitStr 0x0003 "B14"
' 	ArgsLd Range 0x0001 
' 	LitDI2 0x0002 
' 	Eq 
' 	IfBlock 
' Line #937:
' 	LitStr 0x0003 "E12"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0005 "ST@TP"
' 	Eq 
' 	IfBlock 
' Line #938:
' 	LitStr 0x0003 "J28"
' 	ArgsLd Range 0x0001 
' 	LitDI2 0x0002 
' 	Eq 
' 	IfBlock 
' Line #939:
' 	LitDI2 0x0002 
' 	LitStr 0x0003 "J22"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #940:
' 	LitDI2 0x0001 
' 	LitStr 0x0003 "J24"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #941:
' 	EndIfBlock 
' Line #942:
' 	LitStr 0x0003 "E12"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0005 "FI@TP"
' 	Eq 
' 	ElseIfBlock 
' Line #943:
' 	LitStr 0x0003 "J28"
' 	ArgsLd Range 0x0001 
' 	LitDI2 0x0001 
' 	Eq 
' 	IfBlock 
' Line #944:
' 	LitDI2 0x0002 
' 	LitStr 0x0003 "J22"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #945:
' 	LitDI2 0x0001 
' 	LitStr 0x0003 "J24"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #946:
' 	EndIfBlock 
' Line #947:
' 	LitStr 0x0003 "E12"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0007 "S&F@TPS"
' 	Eq 
' 	ElseIfBlock 
' Line #948:
' 	LitStr 0x0003 "J28"
' 	ArgsLd Range 0x0001 
' 	LitDI2 0x0000 
' 	Eq 
' 	IfBlock 
' Line #949:
' 	LitDI2 0x0002 
' 	LitStr 0x0003 "J22"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #950:
' 	LitDI2 0x0001 
' 	LitStr 0x0003 "J24"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #951:
' 	EndIfBlock 
' Line #952:
' 	LitStr 0x0003 "E12"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0005 "SF@TP"
' 	Eq 
' 	ElseIfBlock 
' Line #953:
' 	LitStr 0x0003 "J28"
' 	ArgsLd Range 0x0001 
' 	LitDI2 0x0002 
' 	Eq 
' 	IfBlock 
' Line #954:
' 	LitStr 0x0003 "J22"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0003 "J22"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #955:
' 	LitDI2 0x0001 
' 	LitStr 0x0003 "J24"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #956:
' 	LitStr 0x0003 "J28"
' 	ArgsLd Range 0x0001 
' 	LitDI2 0x0001 
' 	Eq 
' 	ElseIfBlock 
' Line #957:
' 	LitStr 0x0007 "J22:J24"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0007 "J22:J24"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #958:
' 	EndIfBlock 
' Line #959:
' 	EndIfBlock 
' Line #960:
' 	EndIfBlock 
' Line #961:
' Line #962:
' 	QuoteRem 0x0004 0x0023 "Step 2c: Bungled TP with 3 Declared"
' Line #963:
' 	LitStr 0x0003 "B14"
' 	ArgsLd Range 0x0001 
' 	LitDI2 0x0003 
' 	Eq 
' 	LitStr 0x0003 "E12"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0005 "ST@TP"
' 	Eq 
' 	And 
' 	IfBlock 
' Line #964:
' 	QuoteRem 0x0008 0x001C "OK 12/14/15 2&3 OK 1 invalid"
' Line #965:
' 	LitStr 0x0003 "J28"
' 	ArgsLd Range 0x0001 
' 	LitDI2 0x0005 
' 	Eq 
' 	IfBlock 
' Line #966:
' 	LitStr 0x0001 "X"
' 	LitStr 0x0003 "J22"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #967:
' 	LitDI2 0x0001 
' 	LitStr 0x0003 "J24"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #968:
' 	LitDI2 0x0002 
' 	LitStr 0x0003 "J26"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #969:
' 	QuoteRem 0x0008 0x000B "OK 12/14/15"
' Line #970:
' 	LitStr 0x0003 "J28"
' 	ArgsLd Range 0x0001 
' 	LitDI2 0x0003 
' 	Eq 
' 	ElseIfBlock 
' Line #971:
' 	LitStr 0x0003 "J26"
' 	ArgsLd Range 0x0001 
' 	LitDI2 0x0003 
' 	Eq 
' 	IfBlock 
' Line #972:
' 	LitStr 0x0003 "K22"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0003 "K20"
' 	ArgsLd Range 0x0001 
' 	Ne 
' 	LitStr 0x0003 "L22"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0003 "L20"
' 	ArgsLd Range 0x0001 
' 	Ne 
' 	Or 
' 	IfBlock 
' Line #973:
' 	LitDI2 0x0001 
' 	LitStr 0x0003 "J22"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #974:
' 	LitDI2 0x0003 
' 	LitStr 0x0003 "J24"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #975:
' 	LitDI2 0x0002 
' 	LitStr 0x0003 "J26"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #976:
' 	LitStr 0x0003 "K22"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0003 "K20"
' 	ArgsLd Range 0x0001 
' 	Eq 
' 	LitStr 0x0003 "L22"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0003 "L20"
' 	ArgsLd Range 0x0001 
' 	Eq 
' 	And 
' 	ElseIfBlock 
' Line #977:
' 	LitStr 0x0003 "K22"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0003 "K24"
' 	ArgsLd Range 0x0001 
' 	Eq 
' 	LitStr 0x0003 "L22"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0003 "L24"
' 	ArgsLd Range 0x0001 
' 	Eq 
' 	And 
' 	IfBlock 
' Line #978:
' 	LitStr 0x0001 "X"
' 	LitStr 0x0003 "J22"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #979:
' 	LitDI2 0x0002 
' 	LitStr 0x0003 "J24"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #980:
' 	LitDI2 0x0001 
' 	LitStr 0x0003 "J26"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #981:
' 	EndIfBlock 
' Line #982:
' 	EndIfBlock 
' Line #983:
' 	EndIfBlock 
' Line #984:
' 	LitStr 0x0003 "J28"
' 	ArgsLd Range 0x0001 
' 	LitDI2 0x0001 
' 	Eq 
' 	ElseIfBlock 
' Line #985:
' 	LitStr 0x0003 "K24"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0003 "K26"
' 	ArgsLd Range 0x0001 
' 	Eq 
' 	LitStr 0x0003 "L24"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0003 "L26"
' 	ArgsLd Range 0x0001 
' 	Eq 
' 	And 
' 	IfBlock 
' Line #986:
' 	LitDI2 0x0001 
' 	LitStr 0x0003 "J22"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #987:
' 	LitDI2 0x0002 
' 	LitStr 0x0003 "J24"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #988:
' 	LitStr 0x0001 "X"
' 	LitStr 0x0003 "J26"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #989:
' 	EndIfBlock 
' Line #990:
' 	LitStr 0x0003 "J28"
' 	ArgsLd Range 0x0001 
' 	LitDI2 0x0000 
' 	Eq 
' 	LitStr 0x0003 "K22"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0003 "K24"
' 	ArgsLd Range 0x0001 
' 	Eq 
' 	And 
' 	LitStr 0x0003 "K24"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0003 "K26"
' 	ArgsLd Range 0x0001 
' 	Eq 
' 	And 
' 	LitStr 0x0003 "L22"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0003 "L24"
' 	ArgsLd Range 0x0001 
' 	Eq 
' 	And 
' 	LitStr 0x0003 "L24"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0003 "L26"
' 	ArgsLd Range 0x0001 
' 	Eq 
' 	And 
' 	ElseIfBlock 
' Line #991:
' 	LitStr 0x0001 "X"
' 	LitStr 0x000B "J22,J24,J26"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #992:
' 	LitStr 0x0003 "J28"
' 	ArgsLd Range 0x0001 
' 	LitDI2 0x0000 
' 	Eq 
' 	ElseIfBlock 
' Line #993:
' 	LitDI2 0x0002 
' 	LitStr 0x0003 "J22"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #994:
' 	LitDI2 0x0001 
' 	LitStr 0x0003 "J24"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #995:
' 	LitDI2 0x0003 
' 	LitStr 0x0003 "J26"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #996:
' 	EndIfBlock 
' Line #997:
' Line #998:
' 	LitStr 0x0003 "B14"
' 	ArgsLd Range 0x0001 
' 	LitDI2 0x0003 
' 	Eq 
' 	LitStr 0x0003 "E12"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0007 "S&F@TPS"
' 	Eq 
' 	And 
' 	ElseIfBlock 
' Line #999:
' 	LitStr 0x0003 "J28"
' 	ArgsLd Range 0x0001 
' 	LitDI2 0x0005 
' 	Eq 
' 	IfBlock 
' Line #1000:
' 	LitStr 0x0003 "J22"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0001 "X"
' 	Eq 
' 	IfBlock 
' Line #1001:
' 	LitStr 0x0003 "K22"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0003 "K26"
' 	ArgsLd Range 0x0001 
' 	Eq 
' 	LitStr 0x0003 "L22"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0003 "L26"
' 	ArgsLd Range 0x0001 
' 	Eq 
' 	And 
' 	IfBlock 
' Line #1002:
' 	LitStr 0x0001 "X"
' 	LitStr 0x0003 "J22"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1003:
' 	LitDI2 0x0001 
' 	LitStr 0x0003 "J24"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1004:
' 	LitDI2 0x0002 
' 	LitStr 0x0003 "J26"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1005:
' 	LitStr 0x0003 "K26"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0003 "K28"
' 	ArgsLd Range 0x0001 
' 	Ne 
' 	LitStr 0x0003 "L26"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0003 "L28"
' 	ArgsLd Range 0x0001 
' 	Ne 
' 	Or 
' 	ElseIfBlock 
' Line #1006:
' 	LitDI2 0x0002 
' 	LitStr 0x0003 "J22"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1007:
' 	LitDI2 0x0001 
' 	LitStr 0x0003 "J24"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1008:
' 	LitDI2 0x0003 
' 	LitStr 0x0003 "J26"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1009:
' 	LitStr 0x0003 "K22"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0003 "K28"
' 	ArgsLd Range 0x0001 
' 	Ne 
' 	LitStr 0x0003 "L22"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0003 "L28"
' 	ArgsLd Range 0x0001 
' 	Ne 
' 	Or 
' 	ElseIfBlock 
' Line #1010:
' 	LitDI2 0x0003 
' 	LitStr 0x0003 "J22"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1011:
' 	LitDI2 0x0001 
' 	LitStr 0x0003 "J24"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1012:
' 	LitDI2 0x0002 
' 	LitStr 0x0003 "J26"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1013:
' 	EndIfBlock 
' Line #1014:
' 	EndIfBlock 
' Line #1015:
' 	QuoteRem 0x000C 0x000A "Trial here"
' Line #1016:
' 	LitStr 0x0003 "J28"
' 	ArgsLd Range 0x0001 
' 	LitDI2 0x0003 
' 	Eq 
' 	LitStr 0x0003 "J26"
' 	ArgsLd Range 0x0001 
' 	LitDI2 0x0003 
' 	Eq 
' 	And 
' 	ElseIfBlock 
' Line #1017:
' 	LitStr 0x0003 "K24"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0003 "K22"
' 	ArgsLd Range 0x0001 
' 	Eq 
' 	LitStr 0x0003 "L24"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0003 "L22"
' 	ArgsLd Range 0x0001 
' 	Eq 
' 	And 
' 	IfBlock 
' Line #1018:
' 	LitStr 0x0003 "K22"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0003 "K20"
' 	ArgsLd Range 0x0001 
' 	Ne 
' 	LitStr 0x0003 "L22"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0003 "L20"
' 	ArgsLd Range 0x0001 
' 	Ne 
' 	Or 
' 	IfBlock 
' Line #1019:
' 	LitDI2 0x0001 
' 	LitStr 0x0003 "J22"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1020:
' 	LitStr 0x0001 "X"
' 	LitStr 0x0003 "J24"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1021:
' 	LitDI2 0x0002 
' 	LitStr 0x0003 "J26"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1022:
' 	EndIfBlock 
' Line #1023:
' 	EndIfBlock 
' Line #1024:
' 	LitStr 0x0003 "J28"
' 	ArgsLd Range 0x0001 
' 	LitDI2 0x0003 
' 	Eq 
' 	LitStr 0x0003 "J26"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0001 "X"
' 	Eq 
' 	And 
' 	ElseIfBlock 
' Line #1025:
' 	LitStr 0x0003 "K24"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0003 "K20"
' 	ArgsLd Range 0x0001 
' 	Eq 
' 	LitStr 0x0003 "L24"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0003 "L20"
' 	ArgsLd Range 0x0001 
' 	Eq 
' 	And 
' 	IfBlock 
' Line #1026:
' 	LitStr 0x0003 "K22"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0003 "K28"
' 	ArgsLd Range 0x0001 
' 	Ne 
' 	LitStr 0x0003 "L22"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0003 "L28"
' 	ArgsLd Range 0x0001 
' 	Ne 
' 	Or 
' 	IfBlock 
' Line #1027:
' 	LitDI2 0x0003 
' 	LitStr 0x0003 "J22"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1028:
' 	LitDI2 0x0002 
' 	LitStr 0x0003 "J24"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1029:
' 	LitDI2 0x0001 
' 	LitStr 0x0003 "J26"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1030:
' 	EndIfBlock 
' Line #1031:
' 	LitStr 0x0003 "J26"
' 	ArgsLd Range 0x0001 
' 	LitDI2 0x0003 
' 	Eq 
' 	LitStr 0x0003 "K22"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0003 "K24"
' 	ArgsLd Range 0x0001 
' 	Eq 
' 	And 
' 	LitStr 0x0003 "L22"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0003 "L24"
' 	ArgsLd Range 0x0001 
' 	Eq 
' 	And 
' 	ElseIfBlock 
' Line #1032:
' 	LitStr 0x0003 "K22"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0003 "K20"
' 	ArgsLd Range 0x0001 
' 	Ne 
' 	LitStr 0x0003 "L22"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0003 "L20"
' 	ArgsLd Range 0x0001 
' 	Ne 
' 	Or 
' 	IfBlock 
' Line #1033:
' 	LitDI2 0x0001 
' 	LitStr 0x0003 "J22"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1034:
' 	LitStr 0x0001 "X"
' 	LitStr 0x0003 "J24"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1035:
' 	LitDI2 0x0002 
' 	LitStr 0x0003 "J26"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1036:
' 	EndIfBlock 
' Line #1037:
' 	EndIfBlock 
' Line #1038:
' 	LitStr 0x0003 "J28"
' 	ArgsLd Range 0x0001 
' 	LitDI2 0x0002 
' 	Eq 
' 	LitStr 0x0003 "J24"
' 	ArgsLd Range 0x0001 
' 	LitDI2 0x0002 
' 	Eq 
' 	And 
' 	ElseIfBlock 
' Line #1039:
' 	LitStr 0x0003 "K22"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0003 "K28"
' 	ArgsLd Range 0x0001 
' 	Ne 
' 	LitStr 0x0003 "L22"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0003 "L28"
' 	ArgsLd Range 0x0001 
' 	Ne 
' 	Or 
' 	IfBlock 
' Line #1040:
' 	LitDI2 0x0003 
' 	LitStr 0x0003 "J22"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1041:
' 	LitDI2 0x0002 
' 	LitStr 0x0003 "J24"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1042:
' 	LitDI2 0x0001 
' 	LitStr 0x0003 "J26"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1043:
' 	EndIfBlock 
' Line #1044:
' 	LitStr 0x0003 "J28"
' 	ArgsLd Range 0x0001 
' 	LitDI2 0x0003 
' 	Eq 
' 	LitStr 0x0003 "J26"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0001 "X"
' 	Eq 
' 	And 
' 	ElseIfBlock 
' Line #1045:
' 	LitStr 0x0003 "K26"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0003 "K20"
' 	ArgsLd Range 0x0001 
' 	Ne 
' 	LitStr 0x0003 "L26"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0003 "L20"
' 	ArgsLd Range 0x0001 
' 	Ne 
' 	Or 
' 	IfBlock 
' Line #1046:
' 	LitStr 0x0003 "K22"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0003 "K28"
' 	ArgsLd Range 0x0001 
' 	Ne 
' 	LitStr 0x0003 "L22"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0003 "L28"
' 	ArgsLd Range 0x0001 
' 	Ne 
' 	Or 
' 	IfBlock 
' Line #1047:
' 	LitDI2 0x0003 
' 	LitStr 0x0003 "J22"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1048:
' 	LitStr 0x0003 "J24"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0003 "J24"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1049:
' 	LitDI2 0x0001 
' 	LitStr 0x0003 "J26"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1050:
' 	LitStr 0x0003 "K22"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0003 "K28"
' 	ArgsLd Range 0x0001 
' 	Eq 
' 	LitStr 0x0003 "L22"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0003 "L28"
' 	ArgsLd Range 0x0001 
' 	Eq 
' 	And 
' 	ElseIfBlock 
' Line #1051:
' 	LitStr 0x0007 "J22:J26"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0007 "J22:J26"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1052:
' 	EndIfBlock 
' Line #1053:
' 	EndIfBlock 
' Line #1054:
' 	LitStr 0x0003 "J28"
' 	ArgsLd Range 0x0001 
' 	LitDI2 0x0001 
' 	Eq 
' 	ElseIfBlock 
' Line #1055:
' 	LitStr 0x0003 "K24"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0003 "K26"
' 	ArgsLd Range 0x0001 
' 	Eq 
' 	LitStr 0x0003 "L24"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0003 "L26"
' 	ArgsLd Range 0x0001 
' 	Eq 
' 	And 
' 	IfBlock 
' Line #1056:
' 	LitStr 0x0003 "K24"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0003 "K28"
' 	ArgsLd Range 0x0001 
' 	Ne 
' 	LitStr 0x0003 "L24"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0003 "L28"
' 	ArgsLd Range 0x0001 
' 	Ne 
' 	Or 
' 	IfBlock 
' Line #1057:
' 	LitDI2 0x0002 
' 	LitStr 0x0003 "J24"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1058:
' 	LitStr 0x0001 "X"
' 	LitStr 0x0003 "J26"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1059:
' 	EndIfBlock 
' Line #1060:
' 	EndIfBlock 
' Line #1061:
' 	LitStr 0x0003 "J28"
' 	ArgsLd Range 0x0001 
' 	LitDI2 0x0000 
' 	Eq 
' 	ElseIfBlock 
' Line #1062:
' 	LitStr 0x0003 "K24"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0003 "K28"
' 	ArgsLd Range 0x0001 
' 	Eq 
' 	LitStr 0x0003 "L24"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0003 "L28"
' 	ArgsLd Range 0x0001 
' 	Eq 
' 	And 
' 	LitStr 0x0003 "K24"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0003 "K26"
' 	ArgsLd Range 0x0001 
' 	Eq 
' 	And 
' 	LitStr 0x0003 "L24"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0003 "L26"
' 	ArgsLd Range 0x0001 
' 	Eq 
' 	And 
' 	IfBlock 
' Line #1063:
' 	LitDI2 0x0002 
' 	LitStr 0x0003 "J22"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1064:
' 	LitDI2 0x0001 
' 	LitStr 0x0003 "J24"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1065:
' 	LitStr 0x0001 "X"
' 	LitStr 0x0003 "J26"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1066:
' 	LitStr 0x0003 "K20"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0003 "K22"
' 	ArgsLd Range 0x0001 
' 	Eq 
' 	LitStr 0x0003 "L20"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0003 "L22"
' 	ArgsLd Range 0x0001 
' 	Eq 
' 	And 
' 	LitStr 0x0003 "K22"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0003 "K24"
' 	ArgsLd Range 0x0001 
' 	Eq 
' 	And 
' 	LitStr 0x0003 "L22"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0003 "L24"
' 	ArgsLd Range 0x0001 
' 	Eq 
' 	And 
' 	ElseIfBlock 
' Line #1067:
' 	LitStr 0x0001 "X"
' 	LitStr 0x0003 "J22"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1068:
' 	LitDI2 0x0002 
' 	LitStr 0x0003 "J24"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1069:
' 	LitDI2 0x0001 
' 	LitStr 0x0003 "J26"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1070:
' 	EndIfBlock 
' Line #1071:
' 	EndIfBlock 
' Line #1072:
' Line #1073:
' 	LitStr 0x0003 "B14"
' 	ArgsLd Range 0x0001 
' 	LitDI2 0x0003 
' 	Eq 
' 	LitStr 0x0003 "E12"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0005 "SF@TP"
' 	Eq 
' 	And 
' 	ElseIfBlock 
' Line #1074:
' 	LitStr 0x0003 "J28"
' 	ArgsLd Range 0x0001 
' 	LitDI2 0x0005 
' 	Eq 
' 	IfBlock 
' Line #1075:
' 	LitDI2 0x0002 
' 	LitStr 0x0003 "J22"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1076:
' 	LitDI2 0x0001 
' 	LitStr 0x0003 "J24"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1077:
' 	LitDI2 0x0003 
' 	LitStr 0x0003 "J26"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1078:
' 	LitStr 0x0003 "J28"
' 	ArgsLd Range 0x0001 
' 	LitDI2 0x0003 
' 	Eq 
' 	ElseIfBlock 
' Line #1079:
' 	LitStr 0x0003 "J22"
' 	ArgsLd Range 0x0001 
' 	LitDI2 0x0001 
' 	Eq 
' 	IfBlock 
' Line #1080:
' 	LitStr 0x0003 "J22"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0003 "J22"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1081:
' 	LitDI2 0x0002 
' 	LitStr 0x0003 "J26"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1082:
' 	LitDI2 0x0003 
' 	LitStr 0x0003 "J24"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1083:
' 	LitStr 0x0003 "J22"
' 	ArgsLd Range 0x0001 
' 	LitDI2 0x0002 
' 	Eq 
' 	ElseIfBlock 
' Line #1084:
' 	LitStr 0x0003 "J22"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0003 "J22"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1085:
' 	LitDI2 0x0001 
' 	LitStr 0x0003 "J24"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1086:
' 	LitDI2 0x0003 
' 	LitStr 0x0003 "J26"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1087:
' 	LitStr 0x0003 "J26"
' 	ArgsLd Range 0x0001 
' 	LitDI2 0x0003 
' 	Eq 
' 	ElseIfBlock 
' Line #1088:
' 	LitStr 0x0001 "X"
' 	LitStr 0x0007 "J22,J24"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1089:
' 	LitDI2 0x0001 
' 	LitStr 0x0003 "J26"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1090:
' 	EndIfBlock 
' Line #1091:
' 	LitStr 0x0003 "J28"
' 	ArgsLd Range 0x0001 
' 	LitDI2 0x0002 
' 	Eq 
' 	ElseIfBlock 
' Line #1092:
' 	LitStr 0x0003 "K22"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0003 "K26"
' 	ArgsLd Range 0x0001 
' 	Eq 
' 	LitStr 0x0003 "L22"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0003 "L26"
' 	ArgsLd Range 0x0001 
' 	Eq 
' 	And 
' 	IfBlock 
' Line #1093:
' 	LitStr 0x0001 "X"
' 	LitStr 0x0003 "J22"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1094:
' 	LitDI2 0x0001 
' 	LitStr 0x0003 "J24"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1095:
' 	LitStr 0x0001 "X"
' 	LitStr 0x0003 "J26"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1096:
' 	EndIfBlock 
' Line #1097:
' 	LitStr 0x0003 "J28"
' 	ArgsLd Range 0x0001 
' 	LitDI2 0x0000 
' 	Eq 
' 	ElseIfBlock 
' Line #1098:
' 	LitStr 0x0003 "K22"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0003 "K24"
' 	ArgsLd Range 0x0001 
' 	Eq 
' 	LitStr 0x0003 "L22"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0003 "L24"
' 	ArgsLd Range 0x0001 
' 	Eq 
' 	And 
' 	LitStr 0x0003 "K24"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0003 "K26"
' 	ArgsLd Range 0x0001 
' 	Eq 
' 	And 
' 	LitStr 0x0003 "L24"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0003 "L26"
' 	ArgsLd Range 0x0001 
' 	Eq 
' 	And 
' 	IfBlock 
' Line #1099:
' 	LitStr 0x0001 "X"
' 	LitStr 0x000B "J22,J24,J26"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1100:
' 	LitStr 0x0003 "K24"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0003 "K26"
' 	ArgsLd Range 0x0001 
' 	Eq 
' 	LitStr 0x0003 "L24"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0003 "L26"
' 	ArgsLd Range 0x0001 
' 	Eq 
' 	And 
' 	ElseIfBlock 
' Line #1101:
' 	LitStr 0x0003 "K22"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0003 "K20"
' 	ArgsLd Range 0x0001 
' 	Eq 
' 	LitStr 0x0003 "L22"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0003 "L20"
' 	ArgsLd Range 0x0001 
' 	Eq 
' 	And 
' 	IfBlock 
' Line #1102:
' 	LitDI2 0x0002 
' 	LitStr 0x0003 "J22"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1103:
' 	LitDI2 0x0001 
' 	LitStr 0x0003 "J24"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1104:
' 	LitDI2 0x0003 
' 	LitStr 0x0003 "J26"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1105:
' 	EndIfBlock 
' Line #1106:
' 	LitStr 0x0003 "K26"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0003 "K28"
' 	ArgsLd Range 0x0001 
' 	Ne 
' 	LitStr 0x0003 "L26"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0003 "L28"
' 	ArgsLd Range 0x0001 
' 	Ne 
' 	Or 
' 	ElseIfBlock 
' Line #1107:
' 	QuoteRem 0x0010 0x0008 "Testing!"
' Line #1108:
' 	LitStr 0x0003 "K22"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0003 "K24"
' 	ArgsLd Range 0x0001 
' 	Ne 
' 	LitStr 0x0003 "L22"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0003 "L24"
' 	ArgsLd Range 0x0001 
' 	Ne 
' 	And 
' 	LitStr 0x0003 "K26"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0003 "K20"
' 	ArgsLd Range 0x0001 
' 	Ne 
' 	And 
' 	LitStr 0x0003 "L26"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0003 "L20"
' 	ArgsLd Range 0x0001 
' 	Ne 
' 	And 
' 	IfBlock 
' Line #1109:
' 	LitDI2 0x0002 
' 	LitStr 0x0003 "J22"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1110:
' 	LitDI2 0x0003 
' 	LitStr 0x0003 "J24"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1111:
' 	LitDI2 0x0001 
' 	LitStr 0x0003 "J26"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1112:
' 	EndIfBlock 
' Line #1113:
' 	LitStr 0x0003 "K26"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0003 "K28"
' 	ArgsLd Range 0x0001 
' 	Eq 
' 	LitStr 0x0003 "L26"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0003 "L28"
' 	ArgsLd Range 0x0001 
' 	Eq 
' 	And 
' 	ElseIfBlock 
' Line #1114:
' 	LitDI2 0x0001 
' 	LitStr 0x0003 "J22"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1115:
' 	LitDI2 0x0003 
' 	LitStr 0x0003 "J24"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1116:
' 	LitDI2 0x0002 
' 	LitStr 0x0003 "J26"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1117:
' 	EndIfBlock 
' Line #1118:
' 	EndIfBlock 
' Line #1119:
' Line #1120:
' 	LitStr 0x0003 "B14"
' 	ArgsLd Range 0x0001 
' 	LitDI2 0x0003 
' 	Eq 
' 	LitStr 0x0003 "E12"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0005 "FI@TP"
' 	Eq 
' 	And 
' 	ElseIfBlock 
' Line #1121:
' 	LitStr 0x0003 "J28"
' 	ArgsLd Range 0x0001 
' 	LitDI2 0x0003 
' 	Eq 
' 	IfBlock 
' Line #1122:
' 	LitStr 0x0003 "J26"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0001 "X"
' 	Eq 
' 	IfBlock 
' Line #1123:
' 	LitStr 0x0003 "K26"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0003 "K24"
' 	ArgsLd Range 0x0001 
' 	Ne 
' 	LitStr 0x0003 "L26"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0003 "L24"
' 	ArgsLd Range 0x0001 
' 	Ne 
' 	Or 
' 	IfBlock 
' Line #1124:
' 	LitStr 0x0003 "K24"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0003 "K28"
' 	ArgsLd Range 0x0001 
' 	Ne 
' 	LitStr 0x0003 "L24"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0003 "L28"
' 	ArgsLd Range 0x0001 
' 	Ne 
' 	Or 
' 	IfBlock 
' Line #1125:
' 	LitStr 0x0003 "K22"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0003 "K26"
' 	ArgsLd Range 0x0001 
' 	Ne 
' 	LitStr 0x0003 "L22"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0003 "L26"
' 	ArgsLd Range 0x0001 
' 	Ne 
' 	Or 
' 	IfBlock 
' Line #1126:
' 	LitDI2 0x0001 
' 	LitStr 0x0003 "J22"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1127:
' 	LitDI2 0x0003 
' 	LitStr 0x0003 "J24"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1128:
' 	LitDI2 0x0002 
' 	LitStr 0x0003 "J26"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1129:
' 	EndIfBlock 
' Line #1130:
' 	EndIfBlock 
' Line #1131:
' 	EndIfBlock 
' Line #1132:
' 	LitStr 0x0003 "J26"
' 	ArgsLd Range 0x0001 
' 	LitDI2 0x0003 
' 	Eq 
' 	ElseIfBlock 
' Line #1133:
' 	LitStr 0x0003 "K24"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0003 "K22"
' 	ArgsLd Range 0x0001 
' 	Eq 
' 	LitStr 0x0003 "L24"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0003 "L22"
' 	ArgsLd Range 0x0001 
' 	Eq 
' 	And 
' 	IfBlock 
' Line #1134:
' 	LitDI2 0x0001 
' 	LitStr 0x0003 "J22"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1135:
' 	LitStr 0x0001 "X"
' 	LitStr 0x0003 "J24"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1136:
' 	LitDI2 0x0002 
' 	LitStr 0x0003 "J26"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1137:
' 	EndIfBlock 
' Line #1138:
' 	EndIfBlock 
' Line #1139:
' 	LitStr 0x0003 "J28"
' 	ArgsLd Range 0x0001 
' 	LitDI2 0x0001 
' 	Eq 
' 	ElseIfBlock 
' Line #1140:
' 	LitStr 0x0003 "K24"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0003 "K26"
' 	ArgsLd Range 0x0001 
' 	Eq 
' 	LitStr 0x0003 "L24"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0003 "L26"
' 	ArgsLd Range 0x0001 
' 	Eq 
' 	And 
' 	IfBlock 
' Line #1141:
' 	LitStr 0x0003 "K26"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0003 "K28"
' 	ArgsLd Range 0x0001 
' 	Ne 
' 	LitStr 0x0003 "L26"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0003 "L28"
' 	ArgsLd Range 0x0001 
' 	Ne 
' 	Or 
' 	IfBlock 
' Line #1142:
' 	LitDI2 0x0002 
' 	LitStr 0x0003 "J22"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1143:
' 	LitDI2 0x0001 
' 	LitStr 0x0003 "J24"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1144:
' 	LitDI2 0x0003 
' 	LitStr 0x0003 "J26"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1145:
' 	LitStr 0x0003 "K24"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0003 "K26"
' 	ArgsLd Range 0x0001 
' 	Eq 
' 	LitStr 0x0003 "L24"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0003 "L26"
' 	ArgsLd Range 0x0001 
' 	Eq 
' 	And 
' 	LitStr 0x0003 "K26"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0003 "K28"
' 	ArgsLd Range 0x0001 
' 	Eq 
' 	And 
' 	LitStr 0x0003 "L26"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0003 "L28"
' 	ArgsLd Range 0x0001 
' 	Eq 
' 	And 
' 	ElseIfBlock 
' Line #1146:
' 	LitDI2 0x0002 
' 	LitStr 0x0003 "J22"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1147:
' 	LitDI2 0x0001 
' 	LitStr 0x0003 "J24"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1148:
' 	LitStr 0x0001 "X"
' 	LitStr 0x0003 "J26"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1149:
' 	LitStr 0x0003 "K26"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0003 "K28"
' 	ArgsLd Range 0x0001 
' 	Eq 
' 	LitStr 0x0003 "L26"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0003 "L28"
' 	ArgsLd Range 0x0001 
' 	Eq 
' 	And 
' 	ElseIfBlock 
' Line #1150:
' 	LitDI2 0x0001 
' 	LitStr 0x0003 "J22"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1151:
' 	LitStr 0x0001 "X"
' 	LitStr 0x0003 "J24"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1152:
' 	LitStr 0x0001 "X"
' 	LitStr 0x0003 "J26"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1153:
' 	EndIfBlock 
' Line #1154:
' 	EndIfBlock 
' Line #1155:
' 	LitStr 0x0003 "J28"
' 	ArgsLd Range 0x0001 
' 	LitDI2 0x0000 
' 	Eq 
' 	ElseIfBlock 
' Line #1156:
' 	LitStr 0x0003 "K22"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0003 "K24"
' 	ArgsLd Range 0x0001 
' 	Eq 
' 	LitStr 0x0003 "K24"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0003 "K26"
' 	ArgsLd Range 0x0001 
' 	Eq 
' 	And 
' 	LitStr 0x0003 "L22"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0003 "L24"
' 	ArgsLd Range 0x0001 
' 	Eq 
' 	And 
' 	LitStr 0x0003 "L24"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0003 "L26"
' 	ArgsLd Range 0x0001 
' 	Eq 
' 	And 
' 	IfBlock 
' Line #1157:
' 	LitStr 0x0001 "X"
' 	LitStr 0x000B "J22,J24,J26"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1158:
' 	LitStr 0x0003 "K26"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0003 "K28"
' 	ArgsLd Range 0x0001 
' 	Eq 
' 	LitStr 0x0003 "L26"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0003 "L28"
' 	ArgsLd Range 0x0001 
' 	Eq 
' 	And 
' 	ElseIfBlock 
' Line #1159:
' 	LitDI2 0x0002 
' 	LitStr 0x0003 "J26"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1160:
' 	LitStr 0x0003 "K22"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0003 "K20"
' 	ArgsLd Range 0x0001 
' 	Ne 
' 	LitStr 0x0003 "L22"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0003 "L20"
' 	ArgsLd Range 0x0001 
' 	Ne 
' 	Or 
' 	IfBlock 
' Line #1161:
' 	LitDI2 0x0001 
' 	LitStr 0x0003 "J22"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1162:
' 	EndIfBlock 
' Line #1163:
' 	LitStr 0x0003 "K22"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0003 "K24"
' 	ArgsLd Range 0x0001 
' 	Eq 
' 	LitStr 0x0003 "L22"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0003 "L24"
' 	ArgsLd Range 0x0001 
' 	Eq 
' 	And 
' 	IfBlock 
' Line #1164:
' 	LitStr 0x0003 "K24"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0003 "K28"
' 	ArgsLd Range 0x0001 
' 	Ne 
' 	LitStr 0x0003 "L24"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0003 "L28"
' 	ArgsLd Range 0x0001 
' 	Ne 
' 	Or 
' 	IfBlock 
' Line #1165:
' 	LitDI2 0x0003 
' 	LitStr 0x0003 "J24"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1166:
' 	EndIfBlock 
' Line #1167:
' 	EndIfBlock 
' Line #1168:
' 	LitStr 0x0003 "K26"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0003 "K24"
' 	ArgsLd Range 0x0001 
' 	Ne 
' 	LitStr 0x0003 "L26"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0003 "L24"
' 	ArgsLd Range 0x0001 
' 	Ne 
' 	Or 
' 	ElseIfBlock 
' Line #1169:
' 	LitStr 0x0003 "K22"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0003 "K28"
' 	ArgsLd Range 0x0001 
' 	Ne 
' 	LitStr 0x0003 "L22"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0003 "L28"
' 	ArgsLd Range 0x0001 
' 	Ne 
' 	Or 
' 	IfBlock 
' Line #1170:
' 	LitDI2 0x0002 
' 	LitStr 0x0003 "J22"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1171:
' 	LitStr 0x0001 "X"
' 	LitStr 0x0003 "J24"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1172:
' 	LitDI2 0x0001 
' 	LitStr 0x0003 "J26"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1173:
' 	LitStr 0x0003 "K22"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0003 "K28"
' 	ArgsLd Range 0x0001 
' 	Eq 
' 	LitStr 0x0003 "L22"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0003 "L28"
' 	ArgsLd Range 0x0001 
' 	Eq 
' 	And 
' 	ElseIfBlock 
' Line #1174:
' 	LitDI2 0x0001 
' 	LitStr 0x0003 "J22"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1175:
' 	LitStr 0x0001 "X"
' 	LitStr 0x0003 "J24"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1176:
' 	LitDI2 0x0002 
' 	LitStr 0x0003 "J26"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1177:
' 	EndIfBlock 
' Line #1178:
' 	LitStr 0x0003 "K26"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0003 "K24"
' 	ArgsLd Range 0x0001 
' 	Eq 
' 	LitStr 0x0003 "L26"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0003 "L24"
' 	ArgsLd Range 0x0001 
' 	Eq 
' 	And 
' 	ElseIfBlock 
' Line #1179:
' 	LitDI2 0x0001 
' 	LitStr 0x0003 "J22"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1180:
' 	LitDI2 0x0002 
' 	LitStr 0x0003 "J24"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1181:
' 	LitStr 0x0001 "X"
' 	LitStr 0x0003 "J26"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1182:
' 	EndIfBlock 
' Line #1183:
' 	EndIfBlock 
' Line #1184:
' 	EndIfBlock 
' Line #1185:
' 	QuoteRem 0x0005 0x002F "Range("I22:I26").Value = Range("J22:J26").Value"
' Line #1186:
' 	LitStr 0x0007 "J22:J26"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0007 "J22:J26"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1187:
' Line #1188:
' 	LitStr 0x0007 "A22:G26"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0007 "M22:S26"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1189:
' Line #1190:
' 	LitStr 0x0003 "B14"
' 	ArgsLd Range 0x0001 
' 	LitDI2 0x0002 
' 	Ge 
' 	LitStr 0x0003 "J22"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0001 "X"
' 	Ne 
' 	And 
' 	LitStr 0x0003 "J24"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0001 "X"
' 	Ne 
' 	And 
' 	LitStr 0x0003 "J26"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0001 "X"
' 	Ne 
' 	And 
' 	IfBlock 
' Line #1191:
' 	LitStr 0x0003 "J24"
' 	ArgsLd Range 0x0001 
' 	LitDI2 0x0001 
' 	Eq 
' 	IfBlock 
' Line #1192:
' 	LitStr 0x0007 "M24:S24"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0007 "A22:G22"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1193:
' 	LitStr 0x0003 "J26"
' 	ArgsLd Range 0x0001 
' 	LitDI2 0x0001 
' 	Eq 
' 	ElseIfBlock 
' Line #1194:
' 	LitStr 0x0007 "M26:S26"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0007 "A22:G22"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1195:
' 	EndIfBlock 
' Line #1196:
' Line #1197:
' 	LitStr 0x0003 "J22"
' 	ArgsLd Range 0x0001 
' 	LitDI2 0x0002 
' 	Eq 
' 	IfBlock 
' Line #1198:
' 	LitStr 0x0007 "M22:S22"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0007 "A24:G24"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1199:
' 	LitStr 0x0003 "J26"
' 	ArgsLd Range 0x0001 
' 	LitDI2 0x0002 
' 	Eq 
' 	ElseIfBlock 
' Line #1200:
' 	LitStr 0x0007 "M26:S26"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0007 "A24:G24"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1201:
' 	EndIfBlock 
' Line #1202:
' Line #1203:
' 	LitStr 0x0003 "J22"
' 	ArgsLd Range 0x0001 
' 	LitDI2 0x0003 
' 	Eq 
' 	IfBlock 
' Line #1204:
' 	LitStr 0x0007 "M22:S22"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0007 "A26:G26"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1205:
' 	LitStr 0x0003 "J24"
' 	ArgsLd Range 0x0001 
' 	LitDI2 0x0003 
' 	Eq 
' 	ElseIfBlock 
' Line #1206:
' 	LitStr 0x0007 "M24:S24"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0007 "A26:G26"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1207:
' 	EndIfBlock 
' Line #1208:
' 	EndIfBlock 
' Line #1209:
' Line #1210:
' 	LitStr 0x0003 "J20"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0004 "1 TP"
' 	Eq 
' 	LitStr 0x0003 "B14"
' 	ArgsLd Range 0x0001 
' 	LitDI2 0x0002 
' 	Eq 
' 	And 
' 	IfBlock 
' Line #1211:
' 	LitStr 0x0003 "J22"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0001 "X"
' 	Eq 
' 	IfBlock 
' Line #1212:
' 	LitStr 0x0007 "M24:S24"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0007 "A22:G22"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1213:
' 	LitStr 0x0007 "A24:G24"
' 	ArgsLd Range 0x0001 
' 	ArgsMemCall Clear 0x0000 
' Line #1214:
' 	LitStr 0x0003 "J24"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0001 "X"
' 	Eq 
' 	ElseIfBlock 
' Line #1215:
' 	LitStr 0x0007 "M22:S22"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0007 "A22:G22"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1216:
' 	LitStr 0x0007 "A24:G24"
' 	ArgsLd Range 0x0001 
' 	ArgsMemCall Clear 0x0000 
' Line #1217:
' 	EndIfBlock 
' Line #1218:
' 	LitStr 0x0003 "J20"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0004 "1 TP"
' 	Eq 
' 	LitStr 0x0003 "B14"
' 	ArgsLd Range 0x0001 
' 	LitDI2 0x0003 
' 	Eq 
' 	And 
' 	ElseIfBlock 
' Line #1219:
' 	LitStr 0x0003 "J22"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0001 "X"
' 	Eq 
' 	LitStr 0x0003 "J24"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0001 "X"
' 	Eq 
' 	And 
' 	IfBlock 
' Line #1220:
' 	LitStr 0x0007 "M26:S26"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0007 "A22:G22"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1221:
' 	LitStr 0x0007 "A24:G26"
' 	ArgsLd Range 0x0001 
' 	ArgsMemCall Clear 0x0000 
' Line #1222:
' 	LitStr 0x0003 "J22"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0001 "X"
' 	Eq 
' 	LitStr 0x0003 "J26"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0001 "X"
' 	Eq 
' 	And 
' 	ElseIfBlock 
' Line #1223:
' 	LitStr 0x0007 "M24:S24"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0007 "A22:G22"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1224:
' 	LitStr 0x0007 "A24:G26"
' 	ArgsLd Range 0x0001 
' 	ArgsMemCall Clear 0x0000 
' Line #1225:
' 	LitStr 0x0003 "J24"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0001 "X"
' 	Eq 
' 	LitStr 0x0003 "J26"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0001 "X"
' 	Eq 
' 	And 
' 	ElseIfBlock 
' Line #1226:
' 	LitStr 0x0007 "A24:G26"
' 	ArgsLd Range 0x0001 
' 	ArgsMemCall Clear 0x0000 
' Line #1227:
' 	EndIfBlock 
' Line #1228:
' 	LitStr 0x0003 "J20"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0004 "2 TP"
' 	Eq 
' 	LitStr 0x0003 "B14"
' 	ArgsLd Range 0x0001 
' 	LitDI2 0x0003 
' 	Eq 
' 	And 
' 	ElseIfBlock 
' Line #1229:
' 	LitStr 0x0003 "J22"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0001 "X"
' 	Eq 
' 	LitStr 0x0003 "J24"
' 	ArgsLd Range 0x0001 
' 	LitDI2 0x0001 
' 	Eq 
' 	And 
' 	LitStr 0x0003 "J26"
' 	ArgsLd Range 0x0001 
' 	LitDI2 0x0002 
' 	Eq 
' 	And 
' 	IfBlock 
' Line #1230:
' 	LitStr 0x0007 "M24:S26"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0007 "A22:G24"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1231:
' 	LitStr 0x0007 "A26:G26"
' 	ArgsLd Range 0x0001 
' 	ArgsMemCall Clear 0x0000 
' Line #1232:
' 	LitStr 0x0003 "J22"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0001 "X"
' 	Eq 
' 	LitStr 0x0003 "J26"
' 	ArgsLd Range 0x0001 
' 	LitDI2 0x0001 
' 	Eq 
' 	And 
' 	LitStr 0x0003 "J24"
' 	ArgsLd Range 0x0001 
' 	LitDI2 0x0002 
' 	Eq 
' 	And 
' 	ElseIfBlock 
' Line #1233:
' 	LitStr 0x0007 "M26:S26"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0007 "A22:G22"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1234:
' 	LitStr 0x0007 "M24:S24"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0007 "A24:G24"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1235:
' 	LitStr 0x0007 "A26:G26"
' 	ArgsLd Range 0x0001 
' 	ArgsMemCall Clear 0x0000 
' Line #1236:
' 	LitStr 0x0003 "J26"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0001 "X"
' 	Eq 
' 	LitStr 0x0003 "J22"
' 	ArgsLd Range 0x0001 
' 	LitDI2 0x0001 
' 	Eq 
' 	And 
' 	LitStr 0x0003 "J24"
' 	ArgsLd Range 0x0001 
' 	LitDI2 0x0002 
' 	Eq 
' 	And 
' 	ElseIfBlock 
' Line #1237:
' 	LitStr 0x0007 "M22:S24"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0007 "A22:G24"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1238:
' 	LitStr 0x0007 "A26:G26"
' 	ArgsLd Range 0x0001 
' 	ArgsMemCall Clear 0x0000 
' Line #1239:
' 	LitStr 0x0003 "J26"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0001 "X"
' 	Eq 
' 	LitStr 0x0003 "J22"
' 	ArgsLd Range 0x0001 
' 	LitDI2 0x0002 
' 	Eq 
' 	And 
' 	LitStr 0x0003 "J24"
' 	ArgsLd Range 0x0001 
' 	LitDI2 0x0001 
' 	Eq 
' 	And 
' 	ElseIfBlock 
' Line #1240:
' 	LitStr 0x0007 "M24:S24"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0007 "A22:G22"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1241:
' 	LitStr 0x0007 "M22:S22"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0007 "A24:G24"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1242:
' 	LitStr 0x0007 "A26:G26"
' 	ArgsLd Range 0x0001 
' 	ArgsMemCall Clear 0x0000 
' Line #1243:
' 	LitStr 0x0003 "J24"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0001 "X"
' 	Eq 
' 	LitStr 0x0003 "J22"
' 	ArgsLd Range 0x0001 
' 	LitDI2 0x0001 
' 	Eq 
' 	And 
' 	LitStr 0x0003 "J26"
' 	ArgsLd Range 0x0001 
' 	LitDI2 0x0002 
' 	Eq 
' 	And 
' 	ElseIfBlock 
' Line #1244:
' 	LitStr 0x0007 "M22:S22"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0007 "A22:G22"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1245:
' 	LitStr 0x0007 "M26:S26"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0007 "A24:G24"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1246:
' 	LitStr 0x0007 "A26:G26"
' 	ArgsLd Range 0x0001 
' 	ArgsMemCall Clear 0x0000 
' Line #1247:
' 	LitStr 0x0003 "J24"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0001 "X"
' 	Eq 
' 	LitStr 0x0003 "J26"
' 	ArgsLd Range 0x0001 
' 	LitDI2 0x0001 
' 	Eq 
' 	And 
' 	LitStr 0x0003 "J22"
' 	ArgsLd Range 0x0001 
' 	LitDI2 0x0002 
' 	Eq 
' 	And 
' 	ElseIfBlock 
' Line #1248:
' 	LitStr 0x0007 "M26:S26"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0007 "A22:G22"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1249:
' 	LitStr 0x0007 "M22:S22"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0007 "A24:G24"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1250:
' 	LitStr 0x0007 "A26:G26"
' 	ArgsLd Range 0x0001 
' 	ArgsMemCall Clear 0x0000 
' Line #1251:
' 	EndIfBlock 
' Line #1252:
' 	EndIfBlock 
' Line #1253:
' Line #1254:
' 	EndSub 
' Line #1255:
' Line #1256:
' 	FuncDefn (Sub ClearAb())
' Line #1257:
' 	QuoteRem 0x0000 0x0000 ""
' Line #1258:
' 	QuoteRem 0x0000 0x0075 " ClearAb Macro - preserves PRS equations pre-detangler  A22:G26 Revised 5/11/2018 to remove reference to C:CC desktop"
' Line #1259:
' 	QuoteRem 0x0000 0x0000 ""
' Line #1260:
' 	LitStr 0x0003 "PRS"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemCall Select 0x0000 
' Line #1261:
' 	LitStr 0x0043 "=IF(R14C2<1,0,IF(R[-9]C[3]<>3,IMP!R[-20]C[41],[A.xlsm]OTHER!R62C5))"
' 	LitStr 0x0003 "A22"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1262:
' 	LitStr 0x0043 "=IF(R14C2<1,0,IF(R[-9]C[2]<>3,IMP!R[-20]C[41],[A.xlsm]OTHER!R62C6))"
' 	LitStr 0x0003 "B22"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1263:
' 	LitStr 0x0015 "=IF(RC[-2]>0,"N","S")"
' 	LitStr 0x0003 "C22"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1264:
' 	LitStr 0x0040 "=IF(R14C2<1,0,IF(R[-9]C<>3,IMP!R[-20]C[41],[A.xlsm]OTHER!R62C9))"
' 	LitStr 0x0003 "D22"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1265:
' 	LitStr 0x0045 "=IF(R14C2<1,0,IF(R[-9]C[-1]<>3,IMP!R[-20]C[41],[A.xlsm]OTHER!R62C10))"
' 	LitStr 0x0003 "E22"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1266:
' 	LitStr 0x0015 "=IF(RC[-2]>0,"E","W")"
' 	LitStr 0x0003 "F22"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1267:
' 	LitStr 0x0048 "=IF(AND(R[-8]C[-5]>0,R[-9]C[-3]<>3),IMP!R[-21]C[41],[A.xlsm]OTHER!R22C3)"
' 	LitStr 0x0003 "G22"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1268:
' Line #1269:
' 	LitStr 0x0044 "=IF(R14C2<2,0,IF(R[-11]C[3]<>3,IMP!R[-22]C[48],[A.xlsm]OTHER!R64C5))"
' 	LitStr 0x0003 "A24"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1270:
' 	LitStr 0x0044 "=IF(R14C2<2,0,IF(R[-11]C[2]<>3,IMP!R[-22]C[48],[A.xlsm]OTHER!R64C6))"
' 	LitStr 0x0003 "B24"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1271:
' 	LitStr 0x0015 "=IF(RC[-2]>0,"N","S")"
' 	LitStr 0x0003 "C24"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1272:
' 	LitStr 0x0041 "=IF(R14C2<2,0,IF(R[-11]C<>3,IMP!R[-22]C[48],[A.xlsm]OTHER!R64C9))"
' 	LitStr 0x0003 "D24"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1273:
' 	LitStr 0x0046 "=IF(R14C2<2,0,IF(R[-11]C[-1]<>3,IMP!R[-22]C[48],[A.xlsm]OTHER!R64C10))"
' 	LitStr 0x0003 "E24"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1274:
' 	LitStr 0x0015 "=IF(RC[-2]>0,"E","W")"
' 	LitStr 0x0003 "F24"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1275:
' 	LitStr 0x004B "=IF(AND(R[-10]C[-5]>=2,R[-11]C[-3]<>3),IMP!R[-23]C[48],[A.xlsm]OTHER!R24C3)"
' 	LitStr 0x0003 "G24"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1276:
' Line #1277:
' 	LitStr 0x005B "=IF(AND(R[-12]C[1]>2,R[-13]C[3]<>3),IMP!R[-24]C[55],IF(R[-12]C[1]>2,[A.xlsm]OTHER!R66C5,0))"
' 	LitStr 0x0003 "A26"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1278:
' 	LitStr 0x0053 "=IF(AND(R14C2>2,R[-13]C[2]<>3),IMP!R[-24]C[55],IF(R[-12]C>2,[A.xlsm]OTHER!R66C6,0))"
' 	LitStr 0x0003 "B26"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1279:
' 	LitStr 0x0015 "=IF(RC[-2]>0,"N","S")"
' 	LitStr 0x0003 "C26"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1280:
' 	LitStr 0x0054 "=IF(AND(R14C2>2,R[-13]C<>3),IMP!R[-24]C[55],IF(R[-12]C[-2]>2,[A.xlsm]OTHER!R66C9,0))"
' 	LitStr 0x0003 "D26"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1281:
' 	LitStr 0x005F "=IF(AND(R[-12]C[-3]>2,R[-13]C[-1]<>3),IMP!R[-24]C[55],IF(R[-12]C[-3]>2,[A.xlsm]OTHER!R66C10,0))"
' 	LitStr 0x0003 "E26"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1282:
' 	LitStr 0x0015 "=IF(RC[-2]>0,"E","W")"
' 	LitStr 0x0003 "F26"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1283:
' 	LitStr 0x0044 "=IF(AND(R14C2=3,R[-13]C[-3]<>3),IMP!R[-25]C[55],[A.xlsm]OTHER!R26C3)"
' 	LitStr 0x0003 "G26"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1284:
' Line #1285:
' 	EndSub 
' Line #1286:
' 	FuncDefn (Sub Duration())
' Line #1287:
' 	QuoteRem 0x0000 0x008A " Finds best Duration w/out LoH penalty FOR AB SHEET2 After PreB; in corporates F.xlsm pressure correction; amended 9/5/15 to cite GPS alts"
' Line #1288:
' 	QuoteRem 0x0000 0x0042 " 6/4/2017 Amended to add Duration on/after 10/1/2017, NO LOH LIMIT"
' Line #1289:
' 	LitVarSpecial (False)
' 	Ld Application 
' 	MemSt ScreenUpdating 
' Line #1290:
' Line #1291:
' 	QuoteRem 0x0000 0x0000 ""
' Line #1292:
' 	LitVarSpecial (False)
' 	Ld Application 
' 	MemSt ScreenUpdating 
' Line #1293:
' 	LitStr 0x0006 "Sheet2"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemCall Activate 0x0000 
' Line #1294:
' 	LitStr 0x0002 "A2"
' 	LitStr 0x0003 "PRS"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0002 "H1"
' 	LitStr 0x0006 "Sheet2"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemSt Value 
' Line #1295:
' 	LitDI4 0xA801 0x0000 
' 	LitStr 0x0002 "H2"
' 	LitStr 0x0006 "Sheet2"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemSt Value 
' Line #1296:
' 	QuoteRem 0x0004 0x0033 " FIRST IF: Duration as of 10/01/2017 - NO LOH LIMIT"
' Line #1297:
' Line #1298:
' 	LitStr 0x0002 "H1"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0002 "H2"
' 	ArgsLd Range 0x0001 
' 	Ge 
' 	IfBlock 
' Line #1299:
' 	LitStr 0x0007 "=RC[-6]"
' 	LitStr 0x0002 "O1"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1300:
' 	LitStr 0x0007 "=RC[-4]"
' 	LitStr 0x0002 "P1"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1301:
' 	LitStr 0x0007 "=RC[-4]"
' 	LitStr 0x0002 "Q1"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1302:
' 	LitStr 0x000B "=MAX(C[-6])"
' 	LitStr 0x0002 "O2"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1303:
' 	LitStr 0x0015 "=MAX(R[1]C:R[59998]C)"
' 	LitStr 0x0005 "P2:Q2"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1304:
' 	LitStr 0x001B "=IF(RC[-7]=R2C15,RC[-4],"")"
' 	LitStr 0x0002 "P3"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1305:
' 	LitStr 0x0019 "=IF(RC[-1]<>"",RC[-4],"")"
' 	LitStr 0x0002 "Q3"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1306:
' Line #1307:
' 	QuoteRem 0x0008 0x0014 "Copy P3,Q3 per Col I"
' Line #1308:
' 	Ld xlCalculationManual 
' 	Ld Application 
' 	MemSt Calculation 
' Line #1309:
' 	StartWithExpr 
' 	LitStr 0x0006 "Sheet2"
' 	ArgsLd Worksheets 0x0001 
' 	With 
' Line #1310:
' 	Ld xlUp 
' 	Ld Rows 
' 	MemLd Count 
' 	LitStr 0x0001 "I"
' 	ArgsMemLdWith Cells 0x0002 
' 	ArgsMemLd End 0x0001 
' 	MemLd Row 
' 	St LastRow 
' Line #1311:
' 	LitStr 0x0004 "P3:Q"
' 	Ld LastRow 
' 	Concat 
' 	ArgsMemLdWith Range 0x0001 
' 	ParamNamed Destination 
' 	Ld xlFillDefault 
' 	ParamNamed Type 
' 	LitStr 0x0005 "P3:Q3"
' 	ArgsMemLdWith Range 0x0001 
' 	ArgsMemCall AutoFill 0x0002 
' Line #1312:
' 	EndWith 
' Line #1313:
' 	LitStr 0x0006 "Sheet2"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemCall Calculate 0x0000 
' Line #1314:
' 	QuoteRem 0x0004 0x0030 "Application.Calculation = xlCalculationAutomatic"
' Line #1315:
' Line #1316:
' 	LitStr 0x0005 "O1:Q2"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0005 "O1:Q2"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1317:
' 	LitStr 0x0009 "O3:Q60000"
' 	ArgsLd Range 0x0001 
' 	ArgsMemCall Clear 0x0000 
' Line #1318:
' Line #1319:
' 	LitStr 0x0002 "O1"
' 	LitStr 0x0006 "Sheet2"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0003 "G14"
' 	LitStr 0x0003 "PRS"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemSt Value 
' Line #1320:
' 	LitStr 0x0002 "P1"
' 	LitStr 0x0006 "Sheet2"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0003 "H14"
' 	LitStr 0x0003 "PRS"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemSt Value 
' Line #1321:
' 	LitStr 0x0002 "Q1"
' 	LitStr 0x0006 "Sheet2"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0003 "G15"
' 	LitStr 0x0003 "PRS"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemSt Value 
' Line #1322:
' 	LitStr 0x0002 "O2"
' 	LitStr 0x0006 "Sheet2"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0003 "G16"
' 	LitStr 0x0003 "PRS"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemSt Value 
' Line #1323:
' 	LitStr 0x0002 "P2"
' 	LitStr 0x0006 "Sheet2"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0003 "H16"
' 	LitStr 0x0003 "PRS"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemSt Value 
' Line #1324:
' 	LitStr 0x0002 "Q2"
' 	LitStr 0x0006 "Sheet2"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0003 "H15"
' 	LitStr 0x0003 "PRS"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemSt Value 
' Line #1325:
' 	LitStr 0x0005 "O1:Q2"
' 	ArgsLd Range 0x0001 
' 	ArgsMemCall Clear 0x0000 
' Line #1326:
' 	QuoteRem 0x0005 0x001B "Need following line for STD"
' Line #1327:
' 	LitStr 0x000A "I10:M60009"
' 	ArgsLd Range 0x0001 
' 	ParamNamed Destination 
' 	LitStr 0x0009 "I1:M60000"
' 	ArgsLd Range 0x0001 
' 	ArgsMemCall Cut 0x0001 
' Line #1328:
' Line #1329:
' 	LitStr 0x0002 "H1"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0002 "H2"
' 	ArgsLd Range 0x0001 
' 	Lt 
' 	ElseIfBlock 
' Line #1330:
' Line #1331:
' 	QuoteRem 0x0004 0x001C "Leave PR altitude data alone"
' Line #1332:
' 	LitStr 0x0006 "Sheet2"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemCall Activate 0x0000 
' Line #1333:
' 	LitStr 0x000A "I10:M60009"
' 	ArgsLd Range 0x0001 
' 	ParamNamed Destination 
' 	LitStr 0x0009 "I1:M60000"
' 	ArgsLd Range 0x0001 
' 	ArgsMemCall Cut 0x0001 
' Line #1334:
' 	LitStr 0x0013 "=SUM(R[2]C:R[103]C)"
' 	LitStr 0x0002 "L8"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1335:
' 	LitStr 0x0013 "=SUM(R[2]C:R[103]C)"
' 	LitStr 0x0002 "M8"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1336:
' 	LitStr 0x0006 "Sheet2"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemCall Calculate 0x0000 
' Line #1337:
' 	LitStr 0x0002 "L8"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0002 "M8"
' 	ArgsLd Range 0x0001 
' 	Eq 
' 	IfBlock 
' Line #1338:
' 	LitDI2 0x0384 
' 	LitStr 0x0002 "L9"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1339:
' 	LitStr 0x0002 "PR"
' 	LitStr 0x0002 "M9"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1340:
' 	LitStr 0x0005 "L8:M8"
' 	ArgsLd Range 0x0001 
' 	ArgsMemCall Clear 0x0000 
' Line #1341:
' 	LitStr 0x0002 "L8"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0002 "M8"
' 	ArgsLd Range 0x0001 
' 	Ne 
' 	ElseIfBlock 
' Line #1342:
' 	LitDI2 0x03E8 
' 	LitStr 0x0002 "L9"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1343:
' 	LitStr 0x0005 "L8:M8"
' 	ArgsLd Range 0x0001 
' 	ArgsMemCall Clear 0x0000 
' Line #1344:
' 	QuoteRem 0x0000 0x002A "HpA correction REVISED NOW AT F1:F3 on PRS"
' Line #1345:
' 	LitStr 0x0002 "F1"
' 	LitStr 0x0003 "PRS"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0002 "R1"
' 	LitStr 0x0006 "Sheet2"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemSt Value 
' Line #1346:
' 	LitStr 0x0002 "F2"
' 	LitStr 0x0003 "PRS"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0002 "S1"
' 	LitStr 0x0006 "Sheet2"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemSt Value 
' Line #1347:
' 	LitStr 0x0002 "F3"
' 	LitStr 0x0003 "PRS"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0002 "T1"
' 	LitStr 0x0006 "Sheet2"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemSt Value 
' Line #1348:
' 	LitStr 0x002B "=IF(RC[-9]<R1C19,RC[-6]+R1C18,RC[-6]+R1C20)"
' 	LitStr 0x0003 "R10"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1349:
' 	QuoteRem 0x0004 0x000A "Copy Ref I"
' Line #1350:
' 	Ld xlCalculationManual 
' 	Ld Application 
' 	MemSt Calculation 
' Line #1351:
' 	StartWithExpr 
' 	LitStr 0x0006 "Sheet2"
' 	ArgsLd Worksheets 0x0001 
' 	With 
' Line #1352:
' 	Ld xlUp 
' 	Ld Rows 
' 	MemLd Count 
' 	LitStr 0x0001 "I"
' 	ArgsMemLdWith Cells 0x0002 
' 	ArgsMemLd End 0x0001 
' 	MemLd Row 
' 	St LastRow 
' Line #1353:
' 	LitStr 0x0005 "R10:R"
' 	Ld LastRow 
' 	Concat 
' 	ArgsMemLdWith Range 0x0001 
' 	ParamNamed Destination 
' 	Ld xlFillDefault 
' 	ParamNamed Type 
' 	LitStr 0x0007 "R10:R10"
' 	ArgsMemLdWith Range 0x0001 
' 	ArgsMemCall AutoFill 0x0002 
' Line #1354:
' 	EndWith 
' Line #1355:
' 	LitStr 0x0006 "Sheet2"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemCall Calculate 0x0000 
' Line #1356:
' 	QuoteRem 0x0004 0x0030 "Application.Calculation = xlCalculationAutomatic"
' Line #1357:
' 	LitStr 0x000A "R10:R60009"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x000A "L10:L60009"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1358:
' 	LitStr 0x0003 "R:T"
' 	ArgsLd Columns 0x0001 
' 	ArgsMemCall Clear 0x0000 
' Line #1359:
' 	EndIfBlock 
' Line #1360:
' Line #1361:
' 	LitStr 0x0015 "=MAX(R[7]C:R[60006]C)"
' 	LitStr 0x0002 "I3"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1362:
' 	LitStr 0x0019 "=MAX(R[7]C14:R[60006]C14)"
' 	LitStr 0x0002 "J3"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1363:
' 	LitStr 0x0015 "=MIN(R[5]C:R[60004]C)"
' 	LitStr 0x0002 "O5"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1364:
' 	LitStr 0x0015 "=MAX(R[5]C:R[60004]C)"
' 	LitStr 0x0005 "P5:V5"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1365:
' 	LitStr 0x002A "=MAX(R[5]C:R[60004]C)-MIN(R[5]C:R[60004]C)"
' 	LitStr 0x0002 "I5"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1366:
' 	LitStr 0x0009 "=R[-2]C/2"
' 	LitStr 0x0002 "I7"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1367:
' 	LitStr 0x0005 "I3:I7"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0005 "I3:I7"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1368:
' Line #1369:
' 	LitStr 0x001A "=IF(RC[-5]=R3C9,RC[-2],"")"
' 	LitStr 0x0003 "N10"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1370:
' 	QuoteRem 0x0004 0x000E "Copy N10 Ref I"
' Line #1371:
' 	Ld xlCalculationManual 
' 	Ld Application 
' 	MemSt Calculation 
' Line #1372:
' 	StartWithExpr 
' 	LitStr 0x0006 "Sheet2"
' 	ArgsLd Worksheets 0x0001 
' 	With 
' Line #1373:
' 	Ld xlUp 
' 	Ld Rows 
' 	MemLd Count 
' 	LitStr 0x0001 "I"
' 	ArgsMemLdWith Cells 0x0002 
' 	ArgsMemLd End 0x0001 
' 	MemLd Row 
' 	St LastRow 
' Line #1374:
' 	LitStr 0x0005 "N10:N"
' 	Ld LastRow 
' 	Concat 
' 	ArgsMemLdWith Range 0x0001 
' 	ParamNamed Destination 
' 	Ld xlFillDefault 
' 	ParamNamed Type 
' 	LitStr 0x0007 "N10:N10"
' 	ArgsMemLdWith Range 0x0001 
' 	ArgsMemCall AutoFill 0x0002 
' Line #1375:
' 	EndWith 
' Line #1376:
' 	LitStr 0x0006 "Sheet2"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemCall Calculate 0x0000 
' Line #1377:
' 	QuoteRem 0x0000 0x0030 "Application.Calculation = xlCalculationAutomatic"
' Line #1378:
' Line #1379:
' 	LitStr 0x000B "=R10C-R3C10"
' 	LitStr 0x0002 "L5"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1380:
' Line #1381:
' 	LitStr 0x0002 "L5"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0002 "L9"
' 	ArgsLd Range 0x0001 
' 	Le 
' 	IfBlock 
' Line #1382:
' 	QuoteRem 0x0008 0x001C "Rel @ G1:K1; Ldg/MoP @ G2:K2"
' Line #1383:
' 	LitStr 0x0007 "I10:M10"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0005 "I1:M1"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1384:
' 	LitStr 0x0019 "=IF(RC[-1]<>"",RC[-6],"")"
' 	LitStr 0x0007 "O10:S10"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1385:
' 	QuoteRem 0x0008 0x000A "Copy Ref I"
' Line #1386:
' 	Ld xlCalculationManual 
' 	Ld Application 
' 	MemSt Calculation 
' Line #1387:
' 	StartWithExpr 
' 	LitStr 0x0006 "Sheet2"
' 	ArgsLd Worksheets 0x0001 
' 	With 
' Line #1388:
' 	Ld xlUp 
' 	Ld Rows 
' 	MemLd Count 
' 	LitStr 0x0001 "I"
' 	ArgsMemLdWith Cells 0x0002 
' 	ArgsMemLd End 0x0001 
' 	MemLd Row 
' 	St LastRow 
' Line #1389:
' 	LitStr 0x0005 "O10:S"
' 	Ld LastRow 
' 	Concat 
' 	ArgsMemLdWith Range 0x0001 
' 	ParamNamed Destination 
' 	Ld xlFillDefault 
' 	ParamNamed Type 
' 	LitStr 0x0007 "O10:S10"
' 	ArgsMemLdWith Range 0x0001 
' 	ArgsMemCall AutoFill 0x0002 
' Line #1390:
' 	EndWith 
' Line #1391:
' 	LitStr 0x0006 "Sheet2"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemCall Calculate 0x0000 
' Line #1392:
' 	QuoteRem 0x0000 0x0030 "Application.Calculation = xlCalculationAutomatic"
' Line #1393:
' Line #1394:
' 	LitStr 0x001B "=MAX(R[8]C[6]:R[60009]C[6])"
' 	LitStr 0x0005 "I2:M2"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1395:
' 	LitStr 0x0005 "I2:M2"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0005 "I2:M2"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1396:
' Line #1397:
' 	QuoteRem 0x0000 0x004D " When Rel/ldg NOT OK, look for lowest pt in first half of flight, w/ LoH < D9"
' Line #1398:
' 	LitStr 0x0002 "L5"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0002 "L9"
' 	ArgsLd Range 0x0001 
' 	Gt 
' 	ElseIfBlock 
' Line #1399:
' 	LitStr 0x0021 "=IF(RC[-6]<=R10C9+R7C9,RC[-3],"")"
' 	LitStr 0x0003 "O10"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1400:
' 	LitStr 0x001B "=IF(RC[-1]=R5C15,RC[-7],"")"
' 	LitStr 0x0003 "P10"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1401:
' 	LitStr 0x0039 "=IF(AND(RC[-8]>=R5C16,RC[-5]-R3C10<R9C12),R3C9-RC[-8],"")"
' 	LitStr 0x0003 "Q10"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1402:
' 	LitStr 0x001D "=IF(RC[-1]=R5C[-1],RC[-9],"")"
' 	LitStr 0x0003 "R10"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1403:
' 	LitStr 0x0019 "=IF(RC[-1]<>"",RC[-9],"")"
' 	LitStr 0x0007 "S10:V10"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1404:
' 	QuoteRem 0x0004 0x000A "Copy Ref I"
' Line #1405:
' 	Ld xlCalculationManual 
' 	Ld Application 
' 	MemSt Calculation 
' Line #1406:
' 	StartWithExpr 
' 	LitStr 0x0006 "Sheet2"
' 	ArgsLd Worksheets 0x0001 
' 	With 
' Line #1407:
' 	Ld xlUp 
' 	Ld Rows 
' 	MemLd Count 
' 	LitStr 0x0001 "I"
' 	ArgsMemLdWith Cells 0x0002 
' 	ArgsMemLd End 0x0001 
' 	MemLd Row 
' 	St LastRow 
' Line #1408:
' 	LitStr 0x0005 "O10:V"
' 	Ld LastRow 
' 	Concat 
' 	ArgsMemLdWith Range 0x0001 
' 	ParamNamed Destination 
' 	Ld xlFillDefault 
' 	ParamNamed Type 
' 	LitStr 0x0007 "O10:V10"
' 	ArgsMemLdWith Range 0x0001 
' 	ArgsMemCall AutoFill 0x0002 
' Line #1409:
' 	EndWith 
' Line #1410:
' 	LitStr 0x0006 "Sheet2"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemCall Calculate 0x0000 
' Line #1411:
' 	QuoteRem 0x0000 0x0030 "Application.Calculation = xlCalculationAutomatic"
' Line #1412:
' Line #1413:
' 	LitStr 0x0005 "P5:V5"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0005 "P5:V5"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1414:
' 	LitStr 0x000A "P10:V60009"
' 	ArgsLd Range 0x0001 
' 	ArgsMemCall Clear 0x0000 
' Line #1415:
' Line #1416:
' 	LitStr 0x0015 "=MAX(R[5]C:R[60004]C)"
' 	LitStr 0x0006 "W5:AB5"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1417:
' Line #1418:
' 	LitStr 0x003C "=IF(AND(RC[-14]<R5C18,RC[-11]-R5C21<R9C12),R5C18-RC[-14],"")"
' 	LitStr 0x0003 "W10"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1419:
' 	LitStr 0x001C "=IF(RC[-1]=R5C23,RC[-15],"")"
' 	LitStr 0x0003 "X10"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1420:
' 	LitStr 0x001A "=IF(RC[-1]<>"",RC[-15],"")"
' 	LitStr 0x0003 "Y10"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1421:
' 	LitStr 0x001A "=IF(RC[-1]<>"",RC[-15],"")"
' 	LitStr 0x0003 "Z10"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1422:
' 	LitStr 0x001A "=IF(RC[-1]<>"",RC[-15],"")"
' 	LitStr 0x0004 "AA10"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1423:
' 	LitStr 0x001A "=IF(RC[-1]<>"",RC[-15],"")"
' 	LitStr 0x0004 "AB10"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1424:
' 	QuoteRem 0x0004 0x000A "Copy Ref I"
' Line #1425:
' 	Ld xlCalculationManual 
' 	Ld Application 
' 	MemSt Calculation 
' Line #1426:
' 	StartWithExpr 
' 	LitStr 0x0006 "Sheet2"
' 	ArgsLd Worksheets 0x0001 
' 	With 
' Line #1427:
' 	Ld xlUp 
' 	Ld Rows 
' 	MemLd Count 
' 	LitStr 0x0001 "I"
' 	ArgsMemLdWith Cells 0x0002 
' 	ArgsMemLd End 0x0001 
' 	MemLd Row 
' 	St LastRow 
' Line #1428:
' 	LitStr 0x0006 "W10:AB"
' 	Ld LastRow 
' 	Concat 
' 	ArgsMemLdWith Range 0x0001 
' 	ParamNamed Destination 
' 	Ld xlFillDefault 
' 	ParamNamed Type 
' 	LitStr 0x0008 "W10:AB10"
' 	ArgsMemLdWith Range 0x0001 
' 	ArgsMemCall AutoFill 0x0002 
' Line #1429:
' 	EndWith 
' Line #1430:
' 	LitStr 0x0006 "Sheet2"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemCall Calculate 0x0000 
' Line #1431:
' 	QuoteRem 0x0000 0x0030 "Application.Calculation = xlCalculationAutomatic"
' Line #1432:
' Line #1433:
' 	LitStr 0x0006 "W5:AB5"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0006 "W5:AB5"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1434:
' Line #1435:
' 	LitStr 0x0015 "=MAX(R[5]C:R[60004]C)"
' 	LitStr 0x0005 "P5:V5"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1436:
' Line #1437:
' 	LitStr 0x0039 "=IF(AND(RC[-7]>R5C24,R5C27-RC[-4]<R9C12),RC[-7]-R5C24,"")"
' 	LitStr 0x0003 "P10"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1438:
' 	LitStr 0x001B "=IF(RC[-1]=R5C16,RC[-8],"")"
' 	LitStr 0x0003 "Q10"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1439:
' 	LitStr 0x0019 "=IF(RC[-1]<>"",RC[-8],"")"
' 	LitStr 0x0007 "R10:V10"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1440:
' 	QuoteRem 0x0004 0x000A "Copy Ref I"
' Line #1441:
' 	Ld xlCalculationManual 
' 	Ld Application 
' 	MemSt Calculation 
' Line #1442:
' 	StartWithExpr 
' 	LitStr 0x0006 "Sheet2"
' 	ArgsLd Worksheets 0x0001 
' 	With 
' Line #1443:
' 	Ld xlUp 
' 	Ld Rows 
' 	MemLd Count 
' 	LitStr 0x0001 "I"
' 	ArgsMemLdWith Cells 0x0002 
' 	ArgsMemLd End 0x0001 
' 	MemLd Row 
' 	St LastRow 
' Line #1444:
' 	LitStr 0x0005 "P10:U"
' 	Ld LastRow 
' 	Concat 
' 	ArgsMemLdWith Range 0x0001 
' 	ParamNamed Destination 
' 	Ld xlFillDefault 
' 	ParamNamed Type 
' 	LitStr 0x0007 "P10:U10"
' 	ArgsMemLdWith Range 0x0001 
' 	ArgsMemCall AutoFill 0x0002 
' Line #1445:
' 	EndWith 
' Line #1446:
' 	LitStr 0x0006 "Sheet2"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemCall Calculate 0x0000 
' Line #1447:
' 	QuoteRem 0x0000 0x0030 "Application.Calculation = xlCalculationAutomatic"
' Line #1448:
' Line #1449:
' 	LitStr 0x0006 "X5:AB5"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0005 "I1:M1"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1450:
' 	LitStr 0x0005 "Q5:U5"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0005 "I2:M2"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1451:
' Line #1452:
' 	EndIfBlock 
' Line #1453:
' 	LitStr 0x0004 "N:AH"
' 	ArgsLd Columns 0x0001 
' 	ArgsMemCall Clear 0x0000 
' Line #1454:
' 	LitStr 0x0005 "I3:M7"
' 	ArgsLd Range 0x0001 
' 	ArgsMemCall Clear 0x0000 
' Line #1455:
' 	QuoteRem 0x0004 0x0025 "Restore Pressure Altitude as recorded"
' Line #1456:
' 	LitStr 0x0039 "=IF(R[-3]C[-3]<=PRS!R2C6,R[-3]C-PRS!R1C6,R[-3]C-PRS!R3C6)"
' 	LitStr 0x0005 "L4:L5"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1457:
' 	LitStr 0x0006 "Sheet2"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemCall Calculate 0x0000 
' Line #1458:
' 	LitStr 0x0005 "L4:L5"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0005 "L1:L2"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1459:
' 	LitStr 0x0005 "J4:J5"
' 	ArgsLd Range 0x0001 
' 	ArgsMemCall Clear 0x0000 
' Line #1460:
' 	LitStr 0x0002 "I1"
' 	LitStr 0x0006 "Sheet2"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0003 "G14"
' 	LitStr 0x0003 "PRS"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemSt Value 
' Line #1461:
' 	LitStr 0x0002 "L1"
' 	LitStr 0x0006 "Sheet2"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0003 "H14"
' 	LitStr 0x0003 "PRS"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemSt Value 
' Line #1462:
' 	LitStr 0x0002 "M1"
' 	LitStr 0x0006 "Sheet2"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0003 "G15"
' 	LitStr 0x0003 "PRS"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemSt Value 
' Line #1463:
' 	LitStr 0x0002 "I2"
' 	LitStr 0x0006 "Sheet2"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0003 "G16"
' 	LitStr 0x0003 "PRS"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemSt Value 
' Line #1464:
' 	LitStr 0x0002 "L2"
' 	LitStr 0x0006 "Sheet2"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0003 "H16"
' 	LitStr 0x0003 "PRS"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemSt Value 
' Line #1465:
' 	LitStr 0x0002 "M2"
' 	LitStr 0x0006 "Sheet2"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0003 "H15"
' 	LitStr 0x0003 "PRS"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemSt Value 
' Line #1466:
' 	EndIfBlock 
' Line #1467:
' 	LitStr 0x0005 "H1:M5"
' 	ArgsLd Range 0x0001 
' 	ArgsMemCall Clear 0x0000 
' Line #1468:
' 	LitStr 0x000B "Ab.xlsm!STD"
' 	Ld Application 
' 	ArgsMemCall Run 0x0001 
' Line #1469:
' 	LitStr 0x000E "Ab.xlsm!LapRes"
' 	Ld Application 
' 	ArgsMemCall Run 0x0001 
' Line #1470:
' 	EndSub 
' Line #1471:
' 	FuncDefn (Sub STD())
' Line #1472:
' 	QuoteRem 0x0000 0x0000 ""
' Line #1473:
' 	QuoteRem 0x0000 0x0073 " JLR 12/17/14; amended 1/8/15 for CC; amended 2/21/15 CalcSheet; amended 3/25/15 to ck reverse; 4/3/15 to run in Ab"
' Line #1474:
' 	QuoteRem 0x0000 0x008B " Amended 6/13/15 to ck STD from declared Start; Amended 9/5/15 for GPS alt; "Free" STD deleted 4/29/2018 for SWEDES mod effective 10/1/2018"
' Line #1475:
' 	LitVarSpecial (False)
' 	Ld Application 
' 	MemSt ScreenUpdating 
' Line #1476:
' Line #1477:
' 	QuoteRem 0x0001 0x0015 "NOW, STD from Release"
' Line #1478:
' 	LitStr 0x0015 "=MAX(R[2]C:R[60001]C)"
' 	LitStr 0x0005 "O8:U8"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1479:
' 	LitStr 0x0040 "=IF(R1509C<>"",LARGE(R[3]C:R[60002]C,1500),MAX(R[3]C:R[60002]C))"
' 	LitStr 0x0002 "O7"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1480:
' Line #1481:
' 	LineCont 0x0004 07 00 08 00
' 	LitStr 0x0077 "=IF(RC[-4]=R10C11,"",IF(RC[-6]>R10C9,6371*ACOS(SIN(RC[-5])*SIN(R10C10)+COS(RC[-5])*COS(R10C10)*COS(R10C11-RC[-4])),""))"
' 	LitStr 0x0003 "O10"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1482:
' 	QuoteRem 0x0004 0x000A "Copy Ref I"
' Line #1483:
' 	Ld xlCalculationManual 
' 	Ld Application 
' 	MemSt Calculation 
' Line #1484:
' 	StartWithExpr 
' 	LitStr 0x0006 "Sheet2"
' 	ArgsLd Worksheets 0x0001 
' 	With 
' Line #1485:
' 	Ld xlUp 
' 	Ld Rows 
' 	MemLd Count 
' 	LitStr 0x0001 "I"
' 	ArgsMemLdWith Cells 0x0002 
' 	ArgsMemLd End 0x0001 
' 	MemLd Row 
' 	St LastRow 
' Line #1486:
' 	LitStr 0x0005 "O10:O"
' 	Ld LastRow 
' 	Concat 
' 	ArgsMemLdWith Range 0x0001 
' 	ParamNamed Destination 
' 	Ld xlFillDefault 
' 	ParamNamed Type 
' 	LitStr 0x0007 "O10:O10"
' 	ArgsMemLdWith Range 0x0001 
' 	ArgsMemCall AutoFill 0x0002 
' Line #1487:
' 	EndWith 
' Line #1488:
' 	QuoteRem 0x0000 0x0030 "Application.Calculation = xlCalculationAutomatic"
' Line #1489:
' 	LitStr 0x0006 "Sheet2"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemCall Calculate 0x0000 
' Line #1490:
' Line #1491:
' 	LineCont 0x0004 07 00 08 00
' 	LitStr 0x00E0 "=IF(RC[-1]="","",IF(OR(AND(RC[-1]>100,RC[-1]>=R7C15,R10C12-RC[-4]<=R9C12),AND(R9C13="",RC[-1]<=100,RC[-1]>=R7C15,R10C12-RC[-4]<=10*RC[-1]),(AND(R9C13="PR",RC[-1]<=100,RC[-1]>R7C15,R10C4-RC[-4]<=(10*RC[-1])-100))),RC[-1],""))"
' 	LitStr 0x0003 "P10"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1492:
' Line #1493:
' 	QuoteRem 0x0004 0x0027 "OLD thru 9/30/2018; NEW as of 10/1/2018"
' Line #1494:
' 	LitStr 0x0002 "A1"
' 	ArgsLd Range 0x0001 
' 	LitDI4 0xA96E 0x0000 
' 	Lt 
' 	IfBlock 
' Line #1495:
' 	LitStr 0x001B "=IF(RC[-1]=R8C16,RC[-8],"")"
' 	LitStr 0x0003 "Q10"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1496:
' Line #1497:
' 	LitStr 0x0002 "A1"
' 	ArgsLd Range 0x0001 
' 	LitDI4 0xA96E 0x0000 
' 	Ge 
' 	ElseIfBlock 
' Line #1498:
' 	LitStr 0x002E "=RADIANS(PRS!R[14]C[-14]+(PRS!R[14]C[-13]/60))"
' 	LitStr 0x0002 "O4"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1499:
' 	LitStr 0x002E "=RADIANS(PRS!R[13]C[-11]+(PRS!R[13]C[-10]/60))"
' 	LitStr 0x0002 "O5"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1500:
' 	LitStr 0x006F "=IF(AND(RC[-1]=R8C16,6371*ACOS(SIN(RC[-7])*SIN(R4C15)+COS(RC[-7])*COS(R4C15)*COS(R5C15-RC[-6]))>=50),RC[-8],"")"
' 	LitStr 0x0003 "Q10"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1501:
' 	EndIfBlock 
' Line #1502:
' Line #1503:
' 	LitStr 0x0019 "=IF(RC[-1]<>"",RC[-8],"")"
' 	LitStr 0x0007 "R10:U10"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1504:
' 	QuoteRem 0x0004 0x000A "Copy Ref I"
' Line #1505:
' 	Ld xlCalculationManual 
' 	Ld Application 
' 	MemSt Calculation 
' Line #1506:
' 	StartWithExpr 
' 	LitStr 0x0006 "Sheet2"
' 	ArgsLd Worksheets 0x0001 
' 	With 
' Line #1507:
' 	Ld xlUp 
' 	Ld Rows 
' 	MemLd Count 
' 	LitStr 0x0001 "I"
' 	ArgsMemLdWith Cells 0x0002 
' 	ArgsMemLd End 0x0001 
' 	MemLd Row 
' 	St LastRow 
' Line #1508:
' 	LitStr 0x0005 "P10:U"
' 	Ld LastRow 
' 	Concat 
' 	ArgsMemLdWith Range 0x0001 
' 	ParamNamed Destination 
' 	Ld xlFillDefault 
' 	ParamNamed Type 
' 	LitStr 0x0007 "P10:U10"
' 	ArgsMemLdWith Range 0x0001 
' 	ArgsMemCall AutoFill 0x0002 
' Line #1509:
' 	EndWith 
' Line #1510:
' 	QuoteRem 0x0000 0x0030 "Application.Calculation = xlCalculationAutomatic"
' Line #1511:
' 	LitStr 0x0006 "Sheet2"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemCall Calculate 0x0000 
' Line #1512:
' Line #1513:
' 	QuoteRem 0x0004 0x0015 "For Longest Flight(!)"
' Line #1514:
' 	LitStr 0x0002 "P8"
' 	ArgsLd Range 0x0001 
' 	LitDI2 0x0000 
' 	Eq 
' 	IfBlock 
' Line #1515:
' 	LitStr 0x001C "=LARGE(R[3]C:R[60002]C,3000)"
' 	LitStr 0x0002 "O7"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1516:
' 	LitStr 0x0006 "Sheet2"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemCall Calculate 0x0000 
' Line #1517:
' 	EndIfBlock 
' Line #1518:
' Line #1519:
' 	QuoteRem 0x0004 0x002B "For Mphillip (no Silver - max STD from Rel)"
' Line #1520:
' 	LitStr 0x0002 "Q8"
' 	ArgsLd Range 0x0001 
' 	LitDI2 0x0000 
' 	Eq 
' 	IfBlock 
' Line #1521:
' 	LitStr 0x001B "=IF(RC[-1]=R8C16,RC[-8],"")"
' 	LitStr 0x0003 "Q10"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1522:
' 	StartWithExpr 
' 	LitStr 0x0006 "Sheet2"
' 	ArgsLd Worksheets 0x0001 
' 	With 
' Line #1523:
' 	Ld xlUp 
' 	Ld Rows 
' 	MemLd Count 
' 	LitStr 0x0001 "I"
' 	ArgsMemLdWith Cells 0x0002 
' 	ArgsMemLd End 0x0001 
' 	MemLd Row 
' 	St LastRow 
' Line #1524:
' 	LitStr 0x0005 "Q10:Q"
' 	Ld LastRow 
' 	Concat 
' 	ArgsMemLdWith Range 0x0001 
' 	ParamNamed Destination 
' 	Ld xlFillDefault 
' 	ParamNamed Type 
' 	LitStr 0x0007 "Q10:Q10"
' 	ArgsMemLdWith Range 0x0001 
' 	ArgsMemCall AutoFill 0x0002 
' Line #1525:
' 	EndWith 
' Line #1526:
' 	LitStr 0x0006 "Sheet2"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemCall Calculate 0x0000 
' Line #1527:
' 	EndIfBlock 
' Line #1528:
' Line #1529:
' 	LitStr 0x0005 "Q8:U8"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0005 "I1:M1"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1530:
' 	QuoteRem 0x0005 0x0016 "CK best Fix from St Pt"
' Line #1531:
' 	LitStr 0x002C "=RADIANS(PRS!R[10]C[-9]+(PRS!R[10]C[-8]/60))"
' 	LitStr 0x0003 "J10"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1532:
' 	LitStr 0x002C "=RADIANS(PRS!R[10]C[-7]+(PRS!R[10]C[-6]/60))"
' 	LitStr 0x0003 "K10"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1533:
' 	Ld ActiveSheet 
' 	ArgsMemCall Calculate 0x0000 
' Line #1534:
' 	LitStr 0x0005 "Q8:U8"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0005 "I2:M2"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1535:
' 	QuoteRem 0x0004 0x0014 "RESTORE REL COORDS!!"
' Line #1536:
' 	LitStr 0x0020 "=RADIANS(PRS!R4C6+(PRS!R4C7/60))"
' 	LitStr 0x0003 "J10"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1537:
' 	LitStr 0x0020 "=RADIANS(PRS!R5C6+(PRS!R5C7/60))"
' 	LitStr 0x0003 "K10"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1538:
' 	LitStr 0x0007 "J10:K10"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0007 "J10:K10"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1539:
' Line #1540:
' 	LitStr 0x0009 "O7:U60009"
' 	ArgsLd Range 0x0001 
' 	ArgsMemCall Clear 0x0000 
' Line #1541:
' Line #1542:
' 	QuoteRem 0x0003 0x0025 "Restore Pressure Altitude as recorded"
' Line #1543:
' 	LitStr 0x0002 "F1"
' 	LitStr 0x0003 "PRS"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0002 "R1"
' 	LitStr 0x0006 "Sheet2"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemSt Value 
' Line #1544:
' 	LitStr 0x0002 "F2"
' 	LitStr 0x0003 "PRS"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0002 "S1"
' 	LitStr 0x0006 "Sheet2"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemSt Value 
' Line #1545:
' 	LitStr 0x0002 "F3"
' 	LitStr 0x0003 "PRS"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0002 "T1"
' 	LitStr 0x0006 "Sheet2"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemSt Value 
' Line #1546:
' Line #1547:
' 	LitStr 0x002C "=IF(RC[-7]<=R1C19,RC[-4]-R1C18,RC[-4]-R1C20)"
' 	LitStr 0x0005 "P1:P2"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1548:
' 	LitStr 0x0006 "Sheet2"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemCall Calculate 0x0000 
' Line #1549:
' 	LitStr 0x0005 "P1:P2"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0005 "L1:L2"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1550:
' 	LitStr 0x0005 "P1:T2"
' 	ArgsLd Range 0x0001 
' 	ArgsMemCall Clear 0x0000 
' Line #1551:
' Line #1552:
' 	QuoteRem 0x0003 0x0011 "Put it somewhere!"
' Line #1553:
' 	LitStr 0x0002 "I1"
' 	LitStr 0x0006 "Sheet2"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0003 "J11"
' 	LitStr 0x0003 "PRS"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemSt Value 
' Line #1554:
' 	LitStr 0x0002 "J1"
' 	LitStr 0x0006 "Sheet2"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0003 "J12"
' 	LitStr 0x0003 "PRS"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemSt Value 
' Line #1555:
' 	LitStr 0x0002 "K1"
' 	LitStr 0x0006 "Sheet2"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0003 "J13"
' 	LitStr 0x0003 "PRS"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemSt Value 
' Line #1556:
' 	LitStr 0x0002 "L1"
' 	LitStr 0x0006 "Sheet2"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0003 "J14"
' 	LitStr 0x0003 "PRS"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemSt Value 
' Line #1557:
' 	LitStr 0x0002 "M1"
' 	LitStr 0x0006 "Sheet2"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0003 "J15"
' 	LitStr 0x0003 "PRS"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemSt Value 
' Line #1558:
' 	LitStr 0x0002 "I2"
' 	LitStr 0x0006 "Sheet2"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0003 "I11"
' 	LitStr 0x0003 "PRS"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemSt Value 
' Line #1559:
' 	LitStr 0x0002 "J2"
' 	LitStr 0x0006 "Sheet2"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0003 "I12"
' 	LitStr 0x0003 "PRS"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemSt Value 
' Line #1560:
' 	LitStr 0x0002 "K2"
' 	LitStr 0x0006 "Sheet2"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0003 "I13"
' 	LitStr 0x0003 "PRS"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemSt Value 
' Line #1561:
' 	LitStr 0x0002 "L2"
' 	LitStr 0x0006 "Sheet2"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0003 "I14"
' 	LitStr 0x0003 "PRS"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemSt Value 
' Line #1562:
' 	LitStr 0x0002 "M2"
' 	LitStr 0x0006 "Sheet2"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0003 "I15"
' 	LitStr 0x0003 "PRS"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemSt Value 
' Line #1563:
' 	LitStr 0x0006 "Sheet2"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemCall Activate 0x0000 
' Line #1564:
' 	LitStr 0x0005 "I1:O6"
' 	ArgsLd Range 0x0001 
' 	ArgsMemCall Clear 0x0000 
' Line #1565:
' 	LitStr 0x0002 "A1"
' 	ArgsLd Range 0x0001 
' 	ArgsMemCall Activate 0x0000 
' Line #1566:
' Line #1567:
' 	EndSub 
' Line #1568:
' Line #1569:
' 	FuncDefn (Sub LapRes())
' Line #1570:
' 	QuoteRem 0x0000 0x0000 ""
' Line #1571:
' 	QuoteRem 0x0000 0x0072 " 2/2/18 Works - 4 seconds  TESTED IN C for use in Ab after STD; amended 2/5/18 to ref Col A for brevity now 3 secs"
' Line #1572:
' 	QuoteRem 0x0000 0x0000 ""
' Line #1573:
' Line #1574:
' 	LitStr 0x0006 "Sheet2"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemCall Activate 0x0000 
' Line #1575:
' Line #1576:
' 	LitStr 0x0017 "=SUM(R[1]C[3]:R[5]C[3])"
' 	LitStr 0x0002 "O1"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1577:
' 	LitStr 0x0062 "=IF(AND(R[-1]C[1]=R[4]C[1],R[-1]C[2]=R[4]C[2],MIN(RC[3],R[2]C[3],R[4]C[3])>=0.28*R[-1]C),"FAI","")"
' 	LitStr 0x0002 "O2"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1578:
' 	LitStr 0x000A "=PRS!R14C2"
' 	LitStr 0x0002 "O3"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1579:
' 	LitStr 0x000B "=PRS!R20C11"
' 	LitStr 0x0002 "P1"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1580:
' 	LitStr 0x000B "=PRS!R20C12"
' 	LitStr 0x0002 "Q1"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1581:
' Line #1582:
' 	LitStr 0x000B "=PRS!R22C11"
' 	LitStr 0x0002 "P2"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1583:
' 	LitStr 0x000B "=PRS!R22C12"
' 	LitStr 0x0002 "Q2"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1584:
' 	LitStr 0x005A "=6371*ACOS(SIN(RC[-2])*SIN(R[-1]C[-2])+COS(RC[-2])*COS(R[-1]C[-2])*COS(RC[-1]-R[-1]C[-1]))"
' 	LitStr 0x0002 "R2"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1585:
' 	LitStr 0x0009 "First Leg"
' 	LitStr 0x0002 "S2"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1586:
' 	LitStr 0x000B "=PRS!R24C11"
' 	LitStr 0x0002 "P4"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1587:
' 	LitStr 0x000B "=PRS!R24C12"
' 	LitStr 0x0002 "Q4"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1588:
' 	LitStr 0x005A "=6371*ACOS(SIN(RC[-2])*SIN(R[-2]C[-2])+COS(RC[-2])*COS(R[-2]C[-2])*COS(R[-2]C[-1]-RC[-1]))"
' 	LitStr 0x0002 "R4"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1589:
' 	LitStr 0x0007 "2nd leg"
' 	LitStr 0x0002 "S4"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1590:
' Line #1591:
' 	LitStr 0x000B "=PRS!R28C11"
' 	LitStr 0x0002 "P6"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1592:
' 	LitStr 0x000B "=PRS!R28C12"
' 	LitStr 0x0002 "Q6"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1593:
' 	LitStr 0x005A "=6371*ACOS(SIN(R[-2]C[-2])*SIN(RC[-2])+COS(R[-2]C[-2])*COS(RC[-2])*COS(R[-2]C[-1]-RC[-1]))"
' 	LitStr 0x0002 "R6"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1594:
' 	LitStr 0x0008 "last leg"
' 	LitStr 0x0002 "S6"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1595:
' Line #1596:
' 	LitDI2 0x03E8 
' 	LitStr 0x0002 "R9"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1597:
' Line #1598:
' 	Ld ActiveSheet 
' 	ArgsMemCall Calculate 0x0000 
' Line #1599:
' Line #1600:
' 	QuoteRem 0x0004 0x0007 "TRIGGER"
' Line #1601:
' 	LitStr 0x0002 "O1"
' 	ArgsLd Range 0x0001 
' 	LitDI2 0x0096 
' 	Gt 
' 	LitStr 0x0002 "O2"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0003 "FAI"
' 	Ne 
' 	Or 
' 	IfBlock 
' Line #1602:
' 	LitStr 0x0005 "Q1:S6"
' 	ArgsLd Range 0x0001 
' 	ArgsMemCall Clear 0x0000 
' Line #1603:
' 	ExitSub 
' Line #1604:
' 	LitStr 0x0002 "O1"
' 	ArgsLd Range 0x0001 
' 	LitDI2 0x0096 
' 	Lt 
' 	LitStr 0x0002 "O2"
' 	ArgsLd Range 0x0001 
' 	LitStr 0x0003 "FAI"
' 	Eq 
' 	And 
' 	ElseIfBlock 
' Line #1605:
' Line #1606:
' 	QuoteRem 0x0004 0x001F " First TP DO THIS BEFORE START!"
' Line #1607:
' 	LitStr 0x0019 "=MAX(R[7]C28:R[10009]C28)"
' 	LitStr 0x0003 "AA3"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1608:
' 	LitStr 0x005A "=6371*ACOS(SIN(R[-9]C[-24])*SIN(R2C16)+COS(R[-9]C[-24])*COS(R2C16)*COS(R2C17-R[-9]C[-22]))"
' 	LitStr 0x0004 "AA10"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1609:
' 	LitStr 0x00E3 "=IF(AND(RC[-1]<5,6371*ACOS(SIN(R[-9]C[-25])*SIN(R1C16)+COS(R[-9]C[-25])*COS(R1C16)*COS(R1C17-R[-9]C[-23]))>=R2C18,6371*ACOS(SIN(R[-9]C[-25])*SIN(R4C16)+COS(R[-9]C[-25])*COS(R4C16)*COS(R4C17-R[-9]C[-23]))>=R4C18),R[-9]C[-27],"")"
' 	LitStr 0x0004 "AB10"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1610:
' 	LitStr 0x002A "=IF(AND(RC[-2]<5,R[-1]C[-1]=""),RC[-1],"")"
' 	LitStr 0x0004 "AC10"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1611:
' 	QuoteRem 0x0004 0x000A "Copy Ref A"
' Line #1612:
' 	StartWithExpr 
' 	LitStr 0x0006 "Sheet2"
' 	ArgsLd Worksheets 0x0001 
' 	With 
' Line #1613:
' 	Ld xlUp 
' 	Ld Rows 
' 	MemLd Count 
' 	LitStr 0x0001 "A"
' 	ArgsMemLdWith Cells 0x0002 
' 	ArgsMemLd End 0x0001 
' 	MemLd Row 
' 	St LastRow 
' Line #1614:
' 	LitStr 0x0007 "AA10:AC"
' 	Ld LastRow 
' 	Concat 
' 	ArgsMemLdWith Range 0x0001 
' 	ParamNamed Destination 
' 	Ld xlFillDefault 
' 	ParamNamed Type 
' 	LitStr 0x0009 "AA10:AC10"
' 	ArgsMemLdWith Range 0x0001 
' 	ArgsMemCall AutoFill 0x0002 
' Line #1615:
' 	EndWith 
' Line #1616:
' Line #1617:
' 	Ld ActiveSheet 
' 	ArgsMemCall Calculate 0x0000 
' Line #1618:
' 	LitStr 0x0003 "AA3"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0003 "AA3"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1619:
' 	LitStr 0x000C "AA10:AC10009"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x000C "AA10:AC10009"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1620:
' 	QuoteRem 0x0004 0x0019 "SORT Range("AC1:AC60000")"
' Line #1621:
' 	LitStr 0x0004 "AC10"
' 	ArgsLd Range 0x0001 
' 	ParamNamed Key1 
' 	Ld xlAscending 
' 	ParamNamed Order1 
' 	Ld xlGuess 
' 	ParamNamed Header 
' 	LitDI2 0x0001 
' 	ParamNamed OrderCustom 
' 	LitVarSpecial (False)
' 	ParamNamed MatchCase 
' 	Ld xlTopToBottom 
' 	ParamNamed Orientation 
' 	Ld xlSortNormal 
' 	ParamNamed DataOption1 
' 	LitStr 0x000B "AC1:AC10009"
' 	ArgsLd Range 0x0001 
' 	ArgsMemCall Sort 0x0007 
' Line #1622:
' Line #1623:
' 	QuoteRem 0x0004 0x0010 "Start Line Times"
' Line #1624:
' 	LitStr 0x005A "=6371*ACOS(SIN(R[-9]C[-18])*SIN(R1C16)+COS(R[-9]C[-18])*COS(R1C16)*COS(R[-9]C[-16]-R1C17))"
' 	LitStr 0x0003 "U10"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1625:
' 	LitStr 0x001F "=IF(RC[-1]<=0.5,R[-9]C[-21],"")"
' 	LitStr 0x0003 "V10"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1626:
' 	LitStr 0x00ED "=IF(AND(RC[-1]<>"",6371*ACOS(SIN(R[-9]C[-20])*SIN(R2C16)+COS(R[-9]C[-20])*COS(R2C16)*COS(R2C17-R[-9]C[-18]))>=R2C18,6371*ACOS(SIN(R[-8]C[-20])*SIN(R2C16)+COS(R[-8]C[-20])*COS(R2C16)*COS(R2C17-R[-8]C[-18]))<=R2C18,RC[-1]<R3C27),RC[-1],"")"
' 	LitStr 0x0003 "W10"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1627:
' 	LitStr 0x002B "=IF(AND(RC[-1]<>"",R[1]C[-1]=""),RC[-1],"")"
' 	LitStr 0x0003 "X10"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1628:
' 	LitStr 0x001E "=IF(RC[-1]<>"",R[-9]C[-19],"")"
' 	LitStr 0x0003 "Y10"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1629:
' 	QuoteRem 0x0004 0x000A "Copy Ref A"
' Line #1630:
' 	StartWithExpr 
' 	LitStr 0x0006 "Sheet2"
' 	ArgsLd Worksheets 0x0001 
' 	With 
' Line #1631:
' 	Ld xlUp 
' 	Ld Rows 
' 	MemLd Count 
' 	LitStr 0x0001 "A"
' 	ArgsMemLdWith Cells 0x0002 
' 	ArgsMemLd End 0x0001 
' 	MemLd Row 
' 	St LastRow 
' Line #1632:
' 	LitStr 0x0005 "U10:Y"
' 	Ld LastRow 
' 	Concat 
' 	ArgsMemLdWith Range 0x0001 
' 	ParamNamed Destination 
' 	Ld xlFillDefault 
' 	ParamNamed Type 
' 	LitStr 0x0007 "U10:Y10"
' 	ArgsMemLdWith Range 0x0001 
' 	ArgsMemCall AutoFill 0x0002 
' Line #1633:
' 	EndWith 
' Line #1634:
' Line #1635:
' 	Ld ActiveSheet 
' 	ArgsMemCall Calculate 0x0000 
' Line #1636:
' Line #1637:
' 	LitStr 0x0009 "X1:Y10009"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0009 "X1:Y10009"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1638:
' 	LitStr 0x0002 "X1"
' 	ArgsLd Range 0x0001 
' 	ParamNamed Key1 
' 	Ld xlAscending 
' 	ParamNamed Order1 
' 	Ld xlGuess 
' 	ParamNamed Header 
' 	LitDI2 0x0001 
' 	ParamNamed OrderCustom 
' 	LitVarSpecial (False)
' 	ParamNamed MatchCase 
' 	Ld xlTopToBottom 
' 	ParamNamed Orientation 
' 	Ld xlSortNormal 
' 	ParamNamed DataOption1 
' 	LitStr 0x0009 "X1:Y10009"
' 	ArgsLd Range 0x0001 
' 	ArgsMemCall Sort 0x0007 
' Line #1639:
' Line #1640:
' 	LitStr 0x000A "U10:Y10009"
' 	ArgsLd Range 0x0001 
' 	ArgsMemCall Clear 0x0000 
' Line #1641:
' Line #1642:
' 	QuoteRem 0x0004 0x0015 "Start Line Correction"
' Line #1643:
' 	LitStr 0x005B "=IF(RC[-1]="","",IF(RC[-2]<=PRS!R2C6,RC[-1] +PRS!R1C6,IF(RC[-2]>PRS!R2C6,RC[-1]+PRS!R3C6)))"
' 	LitStr 0x0006 "Z1:Z10"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1644:
' 	Ld ActiveSheet 
' 	ArgsMemCall Calculate 0x0000 
' Line #1645:
' 	LitStr 0x0006 "Z1:Z10"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0006 "Y1:Y10"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1646:
' 	LitStr 0x0006 "Z1:Z10"
' 	ArgsLd Range 0x0001 
' 	ArgsMemCall Clear 0x0000 
' Line #1647:
' Line #1648:
' 	QuoteRem 0x0004 0x0021 "Last TP OZ  DO THIS AFTER START!!"
' Line #1649:
' 	LitStr 0x0015 "=MAX(R[7]C:R[10009]C)"
' 	LitStr 0x0003 "AE3"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1650:
' 	LitStr 0x005A "=6371*ACOS(SIN(R[-9]C[-27])*SIN(R4C16)+COS(R[-9]C[-27])*COS(R4C16)*COS(R4C17-R[-9]C[-25]))"
' 	LitStr 0x0004 "AD10"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1651:
' 	LitStr 0x00E2 "=IF(AND(RC[-1]<5,6371*ACOS(SIN(R[-9]C[-28])*SIN(R6C16)+COS(R[-9]C[-28])*COS(R6C16)*COS(R6C17-R[-9]C[-26]))>=R6C18,6371*ACOS(SIN(R[-9]C[-28])*SIN(R2C16)+COS(R[-9]C[-28])*COS(R2C16)*COS(R2C17-R[-9]C[-26]))>=R4C4),R[-9]C[-30],"")"
' 	LitStr 0x0004 "AE10"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1652:
' 	LitStr 0x001B "=IF(R[1]C[-1]="",RC[-1],"")"
' 	LitStr 0x0004 "AF10"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1653:
' Line #1654:
' 	StartWithExpr 
' 	LitStr 0x0006 "Sheet2"
' 	ArgsLd Worksheets 0x0001 
' 	With 
' Line #1655:
' 	Ld xlUp 
' 	Ld Rows 
' 	MemLd Count 
' 	LitStr 0x0001 "A"
' 	ArgsMemLdWith Cells 0x0002 
' 	ArgsMemLd End 0x0001 
' 	MemLd Row 
' 	St LastRow 
' Line #1656:
' 	LitStr 0x0007 "AD10:AF"
' 	Ld LastRow 
' 	Concat 
' 	ArgsMemLdWith Range 0x0001 
' 	ParamNamed Destination 
' 	Ld xlFillDefault 
' 	ParamNamed Type 
' 	LitStr 0x0009 "AD10:AF10"
' 	ArgsMemLdWith Range 0x0001 
' 	ArgsMemCall AutoFill 0x0002 
' Line #1657:
' 	EndWith 
' Line #1658:
' Line #1659:
' 	Ld ActiveSheet 
' 	ArgsMemCall Calculate 0x0000 
' Line #1660:
' Line #1661:
' 	LitStr 0x0003 "AE3"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0003 "AE3"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1662:
' Line #1663:
' 	LitStr 0x000B "AF1:AF10009"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x000B "AF1:AF10009"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1664:
' 	QuoteRem 0x0004 0x0004 "SORT"
' Line #1665:
' 	LitStr 0x0003 "AF1"
' 	ArgsLd Range 0x0001 
' 	ParamNamed Key1 
' 	Ld xlAscending 
' 	ParamNamed Order1 
' 	Ld xlGuess 
' 	ParamNamed Header 
' 	LitDI2 0x0001 
' 	ParamNamed OrderCustom 
' 	LitVarSpecial (False)
' 	ParamNamed MatchCase 
' 	Ld xlTopToBottom 
' 	ParamNamed Orientation 
' 	Ld xlSortNormal 
' 	ParamNamed DataOption1 
' 	LitStr 0x000B "AF1:AF10009"
' 	ArgsLd Range 0x0001 
' 	ArgsMemCall Sort 0x0007 
' Line #1666:
' Line #1667:
' 	LitStr 0x000C "AA10:AF10009"
' 	ArgsLd Range 0x0001 
' 	ArgsMemCall ClearContents 0x0000 
' Line #1668:
' Line #1669:
' 	QuoteRem 0x0004 0x0010 "Finish Line Time"
' Line #1670:
' 	LitStr 0x005A "=6371*ACOS(SIN(R[-9]C[-17])*SIN(R6C16)+COS(R[-9]C[-17])*COS(R6C16)*COS(R[-9]C[-15]-R6C17))"
' 	LitStr 0x0003 "T10"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1671:
' 	LitStr 0x00A9 "=IF(AND(RC[-1]<=0.5,6371*ACOS(SIN(R[-9]C[-18])*SIN(R4C16)+COS(R[-9]C[-18])*COS(R4C16)*COS(R4C17-R[-9]C[-16]))>=R6C18,R[-9]C[-20]>R3C27,R[-9]C[-20]>R3C31),R[-9]C[-20],"")"
' 	LitStr 0x0003 "U10"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1672:
' 	LitStr 0x002C "=IF(AND(RC[-1]<>"",R[-1]C[-1]=""),RC[-1],"")"
' 	LitStr 0x0003 "V10"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1673:
' 	LitStr 0x0031 "=IF(AND(RC[-1]<>"",R[-1]C[-1]=""),R[-9]C[-17],"")"
' 	LitStr 0x0003 "W10"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1674:
' 	QuoteRem 0x0004 0x000A "Copy Ref A"
' Line #1675:
' 	StartWithExpr 
' 	LitStr 0x0006 "Sheet2"
' 	ArgsLd Worksheets 0x0001 
' 	With 
' Line #1676:
' 	Ld xlUp 
' 	Ld Rows 
' 	MemLd Count 
' 	LitStr 0x0001 "A"
' 	ArgsMemLdWith Cells 0x0002 
' 	ArgsMemLd End 0x0001 
' 	MemLd Row 
' 	St LastRow 
' Line #1677:
' 	LitStr 0x0005 "T10:W"
' 	Ld LastRow 
' 	Concat 
' 	ArgsMemLdWith Range 0x0001 
' 	ParamNamed Destination 
' 	Ld xlFillDefault 
' 	ParamNamed Type 
' 	LitStr 0x0007 "T10:W10"
' 	ArgsMemLdWith Range 0x0001 
' 	ArgsMemCall AutoFill 0x0002 
' Line #1678:
' 	EndWith 
' Line #1679:
' Line #1680:
' 	Ld ActiveSheet 
' 	ArgsMemCall Calculate 0x0000 
' Line #1681:
' Line #1682:
' 	LitStr 0x000A "T10:W10009"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x000A "T10:W10009"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1683:
' 	LitStr 0x000A "T10:U10009"
' 	ArgsLd Range 0x0001 
' 	ArgsMemCall Clear 0x0000 
' Line #1684:
' 	QuoteRem 0x0004 0x0007 "SORT VW"
' Line #1685:
' 	LitStr 0x0003 "V10"
' 	ArgsLd Range 0x0001 
' 	ParamNamed Key1 
' 	Ld xlAscending 
' 	ParamNamed Order1 
' 	Ld xlGuess 
' 	ParamNamed Header 
' 	LitDI2 0x0001 
' 	ParamNamed OrderCustom 
' 	LitVarSpecial (False)
' 	ParamNamed MatchCase 
' 	Ld xlTopToBottom 
' 	ParamNamed Orientation 
' 	Ld xlSortNormal 
' 	ParamNamed DataOption1 
' 	LitStr 0x000A "V10:W10009"
' 	ArgsLd Range 0x0001 
' 	ArgsMemCall Sort 0x0007 
' Line #1686:
' Line #1687:
' 	QuoteRem 0x0004 0x0011 " FinAltCorrection"
' Line #1688:
' 	LitStr 0x0055 "=IF(RC[2]="","",IF(RC[1]<=PRS!R2C6,RC[2]+PRS!R1C6,IF(RC[1]>PRS!R2C6,RC[2]+PRS!R3C6)))"
' 	LitStr 0x0007 "U10:U16"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1689:
' 	Ld ActiveSheet 
' 	ArgsMemCall Calculate 0x0000 
' Line #1690:
' 	LitStr 0x0007 "U10:U16"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0007 "W10:W16"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1691:
' 	LitStr 0x0007 "U10:U16"
' 	ArgsLd Range 0x0001 
' 	ArgsMemCall Clear 0x0000 
' Line #1692:
' Line #1693:
' 	QuoteRem 0x0004 0x000A "SORTHELPER"
' Line #1694:
' 	LitStr 0x0007 "AA3,AE3"
' 	ArgsLd Range 0x0001 
' 	ArgsMemCall Clear 0x0000 
' Line #1695:
' 	LitStr 0x0017 "=IF(RC[1]<>"","FIN","")"
' 	LitStr 0x0007 "U10:U16"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1696:
' 	LitStr 0x0017 "=IF(RC[-1]<>"","ST","")"
' 	LitStr 0x0006 "Z1:Z10"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1697:
' 	LitStr 0x0014 "=IF(RC[-1]<>"",1,"")"
' 	LitStr 0x0007 "AD1:AD6"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1698:
' 	LitStr 0x0014 "=IF(RC[-1]<>"",2,"")"
' 	LitStr 0x0007 "AG1:AG6"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1699:
' Line #1700:
' 	Ld ActiveSheet 
' 	ArgsMemCall Calculate 0x0000 
' Line #1701:
' Line #1702:
' 	LitStr 0x0007 "U10:U16"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0007 "U10:U16"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1703:
' 	LitStr 0x0006 "Z1:Z10"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0006 "Z1:Z10"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1704:
' 	LitStr 0x0007 "AD1:AD6"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0007 "AD1:AD6"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1705:
' 	LitStr 0x0007 "AG1:AG6"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0007 "AG1:AG6"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1706:
' Line #1707:
' 	LitStr 0x0007 "AC1:AC6"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0007 "X11:X16"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1708:
' 	LitStr 0x0007 "AD1:AD6"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0007 "Z11:Z16"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1709:
' 	LitStr 0x0007 "AF1:AF6"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0007 "X17:X22"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1710:
' 	LitStr 0x0007 "AG1:AG6"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0007 "Z17:Z22"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1711:
' 	LitStr 0x0007 "V10:V15"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0007 "X23:X28"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1712:
' 	LitStr 0x0007 "W10:W15"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0007 "Y23:Y28"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1713:
' 	LitStr 0x0007 "U10:U15"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0007 "Z23:Z28"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1714:
' Line #1715:
' 	QuoteRem 0x0004 0x0009 "SORT LAPS"
' Line #1716:
' 	LitStr 0x0002 "X1"
' 	ArgsLd Range 0x0001 
' 	ParamNamed Key1 
' 	Ld xlAscending 
' 	ParamNamed Order1 
' 	Ld xlGuess 
' 	ParamNamed Header 
' 	LitDI2 0x0001 
' 	ParamNamed OrderCustom 
' 	LitVarSpecial (False)
' 	ParamNamed MatchCase 
' 	Ld xlTopToBottom 
' 	ParamNamed Orientation 
' 	Ld xlSortNormal 
' 	ParamNamed DataOption1 
' 	LitStr 0x0006 "X1:Z28"
' 	ArgsLd Range 0x0001 
' 	ArgsMemCall Sort 0x0007 
' Line #1717:
' 	LitStr 0x0009 "h:mm:ss;@"
' 	LitStr 0x0006 "X1:X28"
' 	ArgsLd Range 0x0001 
' 	MemSt ScrollRow 
' Line #1718:
' Line #1719:
' 	LitStr 0x0007 "AC1:AG6"
' 	ArgsLd Range 0x0001 
' 	ArgsMemCall Clear 0x0000 
' Line #1720:
' Line #1721:
' 	QuoteRem 0x0004 0x0019 " BEST 3 ETs including LoH"
' Line #1722:
' 	LitStr 0x0053 "=IF(AND(RC[-2]="ST",R10C22>RC[-4],R10C22<>"",RC[-3]-R10C23<R9C18),R10C22-RC[-4],"")"
' 	LitStr 0x0008 "AB1:AB28"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1723:
' 	LitStr 0x0053 "=IF(AND(RC[-3]="ST",R11C22>RC[-5],R11C22<>"",RC[-4]-R11C23<R9C18),R11C22-RC[-5],"")"
' 	LitStr 0x0008 "AC1:AC28"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1724:
' 	LitStr 0x0053 "=IF(AND(RC[-4]="ST",R12C22>RC[-6],R12C22<>"",RC[-5]-R12C23<R9C18),R12C22-RC[-6],"")"
' 	LitStr 0x0008 "AD1:AD28"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1725:
' 	LitStr 0x0053 "=IF(AND(RC[-5]="ST",R13C22>RC[-7],R13C22<>"",RC[-6]-R13C23<R9C18),R13C22-RC[-7],"")"
' 	LitStr 0x0008 "AE1:AE28"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1726:
' 	LitStr 0x0053 "=IF(AND(RC[-6]="ST",R14C22>RC[-8],R14C22<>"",RC[-7]-R14C23<R9C18),R14C22-RC[-8],"")"
' 	LitStr 0x0008 "AF1:AF28"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1727:
' 	LitStr 0x0053 "=IF(AND(RC[-7]="ST",R15C22>RC[-9],R15C22<>"",RC[-8]-R15C23<R9C18),R15C22-RC[-9],"")"
' 	LitStr 0x0008 "AG1:AG28"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1728:
' 	Ld ActiveSheet 
' 	ArgsMemCall Calculate 0x0000 
' Line #1729:
' 	LitStr 0x0008 "AB1:AG28"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0008 "AB1:AG28"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1730:
' Line #1731:
' 	LitStr 0x0016 "=MIN(RC[6]:R[27]C[11])"
' 	LitStr 0x0002 "V1"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1732:
' 	LitStr 0x001E "=SMALL(R[-1]C[6]:R[26]C[11],2)"
' 	LitStr 0x0002 "V2"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1733:
' 	LitStr 0x001E "=SMALL(R[-2]C[6]:R[25]C[11],3)"
' 	LitStr 0x0002 "V3"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1734:
' 	Ld ActiveSheet 
' 	ArgsMemCall Calculate 0x0000 
' Line #1735:
' 	LitStr 0x0005 "V1:V3"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0005 "V1:V3"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1736:
' Line #1737:
' 	LitStr 0x0047 "=IF(RC[-1]="","",IF(OR(RC[-1]=R1C22,RC[-1]=R2C22,RC[-1]=R3C22),RC[-5]))"
' 	LitStr 0x0008 "AC1:AC28"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1738:
' 	Ld ActiveSheet 
' 	ArgsMemCall Calculate 0x0000 
' Line #1739:
' 	LitStr 0x0008 "AC1:AC28"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0008 "AC1:AC28"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1740:
' 	QuoteRem 0x0004 0x0004 "SORT"
' Line #1741:
' 	LitStr 0x0003 "AB1"
' 	ArgsLd Range 0x0001 
' 	ParamNamed Key1 
' 	Ld xlAscending 
' 	ParamNamed Order1 
' 	Ld xlGuess 
' 	ParamNamed Header 
' 	LitDI2 0x0001 
' 	ParamNamed OrderCustom 
' 	LitVarSpecial (False)
' 	ParamNamed MatchCase 
' 	Ld xlTopToBottom 
' 	ParamNamed Orientation 
' 	Ld xlSortNormal 
' 	ParamNamed DataOption1 
' 	LitStr 0x0008 "AB1:AC28"
' 	ArgsLd Range 0x0001 
' 	ArgsMemCall Sort 0x0007 
' Line #1742:
' Line #1743:
' 	QuoteRem 0x0004 0x0032 "Put candidate Start options somewhere. BEST AT AC1"
' Line #1744:
' Line #1745:
' 	LitStr 0x0003 "AC1"
' 	LitStr 0x0006 "Sheet2"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0003 "C13"
' 	LitStr 0x0003 "PRS"
' 	ArgsLd Sheets 0x0001 
' 	ArgsMemLd Range 0x0001 
' 	MemSt Value 
' Line #1746:
' 	LitStr 0x0013 "=RC[-1]-((1/24)/30)"
' 	LitStr 0x0003 "AD1"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1747:
' 	Ld ActiveSheet 
' 	ArgsMemCall Calculate 0x0000 
' Line #1748:
' 	LitStr 0x0005 "A1:G1"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0005 "O1:U1"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1749:
' 	LitStr 0x001E "=IF(RC[-14]>=R1C30,RC[-14],"")"
' 	LitStr 0x0002 "O2"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1750:
' 	LitStr 0x001A "=IF(RC[-1]<>"",RC[-14],"")"
' 	LitStr 0x0005 "P2:U2"
' 	ArgsLd Range 0x0001 
' 	MemSt FormulaR1C1 
' Line #1751:
' 	QuoteRem 0x0004 0x000A "Copy Ref A"
' Line #1752:
' 	StartWithExpr 
' 	LitStr 0x0006 "Sheet2"
' 	ArgsLd Worksheets 0x0001 
' 	With 
' Line #1753:
' 	Ld xlUp 
' 	Ld Rows 
' 	MemLd Count 
' 	LitStr 0x0001 "A"
' 	ArgsMemLdWith Cells 0x0002 
' 	ArgsMemLd End 0x0001 
' 	MemLd Row 
' 	St LastRow 
' Line #1754:
' 	LitStr 0x0004 "O2:U"
' 	Ld LastRow 
' 	Concat 
' 	ArgsMemLdWith Range 0x0001 
' 	ParamNamed Destination 
' 	Ld xlFillDefault 
' 	ParamNamed Type 
' 	LitStr 0x0005 "O2:U2"
' 	ArgsMemLdWith Range 0x0001 
' 	ArgsMemCall AutoFill 0x0002 
' Line #1755:
' 	EndWith 
' Line #1756:
' Line #1757:
' 	Ld ActiveSheet 
' 	ArgsMemCall Calculate 0x0000 
' Line #1758:
' 	LitStr 0x0009 "O1:U10000"
' 	ArgsLd Range 0x0001 
' 	MemLd Value 
' 	LitStr 0x0009 "A1:G10000"
' 	ArgsLd Range 0x0001 
' 	MemSt Value 
' Line #1759:
' Line #1760:
' 	LitStr 0x0004 "O:AG"
' 	ArgsLd Columns 0x0001 
' 	ArgsMemCall Clear 0x0000 
' Line #1761:
' 	EndIfBlock 
' Line #1762:
' Line #1763:
' 	EndSub 
' VBA/Sheet4 - 1105 bytes
' VBA/Sheet6 - 1174 bytes
' VBA/Module2 - 115093 bytes
' Line #0:
' 	QuoteRem 0x0000 0x003E " Calculation of local times of sunrise, solar noon, and sunset"
' Line #1:
' 	QuoteRem 0x0000 0x0040 " based on the calculation procedure by NOAA in the javascript in"
' Line #2:
' 	QuoteRem 0x0000 0x003D " http://www.srrb.noaa.gov/highlights/sunrise/sunrise.html and"
' Line #3:
' 	QuoteRem 0x0000 0x0036 " http://www.srrb.noaa.gov/highlights/sunrise/azel.html"
' Line #4:
' 	QuoteRem 0x0000 0x0000 ""
' Line #5:
' 	QuoteRem 0x0000 0x003F " The calculations in the NOAA Sunrise/Sunset and Solar Position"
' Line #6:
' 	QuoteRem 0x0000 0x0041 " Calculators are based on equations from Astronomical Algorithms,"
' Line #7:
' 	QuoteRem 0x0000 0x0042 " by Jean Meeus. NOAA also included atmospheric refraction effects."
' Line #8:
' 	QuoteRem 0x0000 0x0035 " The sunrise and sunset results were reported by NOAA"
' Line #9:
' 	QuoteRem 0x0000 0x0044 " to be accurate to within +/- 1 minute for locations between +/- 72°"
' Line #10:
' 	QuoteRem 0x0000 0x003D " latitude, and within ten minutes outside of those latitudes."
' Line #11:
' 	QuoteRem 0x0000 0x0000 ""
' Line #12:
' 	QuoteRem 0x0000 0x0033 " This translation was tested for selected locations"
' Line #13:
' 	QuoteRem 0x0000 0x0038 " and found to provide results within +/- 1 minute of the"
' Line #14:
' 	QuoteRem 0x0000 0x001A " original Javascript code."
' Line #15:
' 	QuoteRem 0x0000 0x0000 ""
' Line #16:
' 	QuoteRem 0x0000 0x003F " This translation does not include calculation of prior or next"
' Line #17:
' 	QuoteRem 0x0000 0x003B " susets for locations above the Arctic Circle and below the"
' Line #18:
' 	QuoteRem 0x0000 0x003B " Antarctic Circle, when a sunrise or sunset does not occur."
' Line #19:
' 	QuoteRem 0x0000 0x0000 ""
' Line #20:
' 	QuoteRem 0x0000 0x0033 " Translated from NOAA's Javascript to Excel VBA by:"
' Line #21:
' 	QuoteRem 0x0000 0x0000 ""
' Line #22:
' 	QuoteRem 0x0000 0x000F " Greg Pelletier"
' Line #23:
' 	QuoteRem 0x0000 0x0016 " Department of Ecology"
' Line #24:
' 	QuoteRem 0x0000 0x000E " P.O.Box 47600"
' Line #25:
' 	QuoteRem 0x0000 0x0017 " Olympia, WA 98504-7600"
' Line #26:
' 	QuoteRem 0x0000 0x001B " email: gpel461@ ecy.wa.gov"
' Line #27:
' Line #28:
' 	Option  (Explicit)
' Line #29:
' Line #30:
' Line #31:
' 	FuncDefn (Function radToDeg(angleRad))
' Line #32:
' 	QuoteRem 0x0000 0x0022 "// Convert radian angle to degrees"
' Line #33:
' Line #34:
' 	LitR8 0x0000 0x0000 0x8000 0x4066 
' 	Ld angleRad 
' 	Mul 
' 	Ld Application 
' 	MemLd WorksheetFunction 
' 	ArgsMemLd Pi 0x0000 
' 	Div 
' 	Paren 
' 	St radToDeg 
' Line #35:
' Line #36:
' 	EndFunc 
' Line #37:
' Line #38:
' Line #39:
' 	FuncDefn (Function degToRad(angleDeg))
' Line #40:
' 	QuoteRem 0x0000 0x0022 "// Convert degree angle to radians"
' Line #41:
' Line #42:
' 	Ld Application 
' 	MemLd WorksheetFunction 
' 	ArgsMemLd Pi 0x0000 
' 	Ld angleDeg 
' 	Mul 
' 	LitR8 0x0000 0x0000 0x8000 0x4066 
' 	Div 
' 	Paren 
' 	St degToRad 
' Line #43:
' Line #44:
' 	EndFunc 
' Line #45:
' Line #46:
' Line #47:
' 	FuncDefn (Function calcJD(year, month, day))
' Line #48:
' Line #49:
' 	QuoteRem 0x0000 0x0048 "***********************************************************************/"
' Line #50:
' 	QuoteRem 0x0000 0x0011 "* Name:    calcJD"
' Line #51:
' 	QuoteRem 0x0000 0x0013 "* Type:    Function"
' Line #52:
' 	QuoteRem 0x0000 0x0027 "* Purpose: Julian day from calendar day"
' Line #53:
' 	QuoteRem 0x0000 0x000C "* Arguments:"
' Line #54:
' 	QuoteRem 0x0000 0x0017 "*   year : 4 digit year"
' Line #55:
' 	QuoteRem 0x0000 0x0016 "*   month: January = 1"
' Line #56:
' 	QuoteRem 0x0000 0x0011 "*   day  : 1 - 31"
' Line #57:
' 	QuoteRem 0x0000 0x000F "* Return value:"
' Line #58:
' 	QuoteRem 0x0000 0x002C "*   The Julian day corresponding to the date"
' Line #59:
' 	QuoteRem 0x0000 0x0007 "* Note:"
' Line #60:
' 	QuoteRem 0x0000 0x0043 "*   Number is returned for start of day.  Fractional days should be"
' Line #61:
' 	QuoteRem 0x0000 0x0010 "*   added later."
' Line #62:
' 	QuoteRem 0x0000 0x0048 "***********************************************************************/"
' Line #63:
' Line #64:
' 	Dim 
' 	VarDefn A (As Double)
' 	VarDefn B (As Double)
' 	VarDefn JD (As Double)
' Line #65:
' Line #66:
' 	Ld month 
' 	LitDI2 0x0002 
' 	Le 
' 	Paren 
' 	IfBlock 
' Line #67:
' 	Ld year 
' 	LitDI2 0x0001 
' 	Sub 
' 	St year 
' Line #68:
' 	Ld month 
' 	LitDI2 0x000C 
' 	Add 
' 	St month 
' Line #69:
' 	EndIfBlock 
' Line #70:
' Line #71:
' 	Ld year 
' 	LitDI2 0x0064 
' 	Div 
' 	LitDI2 0x0001 
' 	Ld Application 
' 	MemLd WorksheetFunction 
' 	ArgsMemLd Floor 0x0002 
' 	St A 
' Line #72:
' 	LitDI2 0x0002 
' 	Ld A 
' 	Sub 
' 	Ld A 
' 	LitDI2 0x0004 
' 	Div 
' 	LitDI2 0x0001 
' 	Ld Application 
' 	MemLd WorksheetFunction 
' 	ArgsMemLd Floor 0x0002 
' 	Add 
' 	St B 
' Line #73:
' Line #74:
' 	LineCont 0x0004 13 00 0D 00
' 	LitR8 0x0000 0x0000 0xD400 0x4076 
' 	Ld year 
' 	LitDI2 0x126C 
' 	Add 
' 	Paren 
' 	Mul 
' 	LitDI2 0x0001 
' 	Ld Application 
' 	MemLd WorksheetFunction 
' 	ArgsMemLd Floor 0x0002 
' 	LitR8 0x5461 0x2752 0x99A0 0x403E 
' 	Ld month 
' 	LitDI2 0x0001 
' 	Add 
' 	Paren 
' 	Mul 
' 	LitDI2 0x0001 
' 	Ld Application 
' 	MemLd WorksheetFunction 
' 	ArgsMemLd Floor 0x0002 
' 	Add 
' 	Ld day 
' 	Add 
' 	Ld B 
' 	Add 
' 	LitR8 0x0000 0x0000 0xD200 0x4097 
' 	Sub 
' 	St JD 
' Line #75:
' 	Ld JD 
' 	St calcJD 
' Line #76:
' Line #77:
' 	QuoteRem 0x0000 0x0030 "gp put the year and month back where they belong"
' Line #78:
' 	Ld month 
' 	LitDI2 0x000D 
' 	Eq 
' 	IfBlock 
' Line #79:
' 	LitDI2 0x0001 
' 	St month 
' Line #80:
' 	Ld year 
' 	LitDI2 0x0001 
' 	Add 
' 	St year 
' Line #81:
' 	EndIfBlock 
' Line #82:
' 	Ld month 
' 	LitDI2 0x000E 
' 	Eq 
' 	IfBlock 
' Line #83:
' 	LitDI2 0x0002 
' 	St month 
' Line #84:
' 	Ld year 
' 	LitDI2 0x0001 
' 	Add 
' 	St year 
' Line #85:
' 	EndIfBlock 
' Line #86:
' Line #87:
' 	EndFunc 
' Line #88:
' Line #89:
' Line #90:
' 	FuncDefn (Function calcTimeJulianCent(JD))
' Line #91:
' Line #92:
' 	QuoteRem 0x0000 0x0048 "***********************************************************************/"
' Line #93:
' 	QuoteRem 0x0000 0x001D "* Name:    calcTimeJulianCent"
' Line #94:
' 	QuoteRem 0x0000 0x0013 "* Type:    Function"
' Line #95:
' 	QuoteRem 0x0000 0x0039 "* Purpose: convert Julian Day to centuries since J2000.0."
' Line #96:
' 	QuoteRem 0x0000 0x000C "* Arguments:"
' Line #97:
' 	QuoteRem 0x0000 0x0022 "*   jd : the Julian Day to convert"
' Line #98:
' 	QuoteRem 0x0000 0x000F "* Return value:"
' Line #99:
' 	QuoteRem 0x0000 0x002F "*   the T value corresponding to the Julian Day"
' Line #100:
' 	QuoteRem 0x0000 0x0048 "***********************************************************************/"
' Line #101:
' Line #102:
' 	Dim 
' 	VarDefn t (As Double)
' Line #103:
' Line #104:
' 	Ld JD 
' 	LitR8 0x0000 0x8000 0xB42C 0x4142 
' 	Sub 
' 	Paren 
' 	LitR8 0x0000 0x0000 0xD5A0 0x40E1 
' 	Div 
' 	St t 
' Line #105:
' 	Ld t 
' 	St calcTimeJulianCent 
' Line #106:
' Line #107:
' 	EndFunc 
' Line #108:
' Line #109:
' Line #110:
' 	FuncDefn (Function calcJDFromJulianCent(t))
' Line #111:
' Line #112:
' 	QuoteRem 0x0000 0x0048 "***********************************************************************/"
' Line #113:
' 	QuoteRem 0x0000 0x001F "* Name:    calcJDFromJulianCent"
' Line #114:
' 	QuoteRem 0x0000 0x0013 "* Type:    Function"
' Line #115:
' 	QuoteRem 0x0000 0x0039 "* Purpose: convert centuries since J2000.0 to Julian Day."
' Line #116:
' 	QuoteRem 0x0000 0x000C "* Arguments:"
' Line #117:
' 	QuoteRem 0x0000 0x0030 "*   t : number of Julian centuries since J2000.0"
' Line #118:
' 	QuoteRem 0x0000 0x000F "* Return value:"
' Line #119:
' 	QuoteRem 0x0000 0x002F "*   the Julian Day corresponding to the t value"
' Line #120:
' 	QuoteRem 0x0000 0x0048 "***********************************************************************/"
' Line #121:
' Line #122:
' 	Dim 
' 	VarDefn JD (As Double)
' Line #123:
' Line #124:
' 	Ld t 
' 	LitR8 0x0000 0x0000 0xD5A0 0x40E1 
' 	Mul 
' 	LitR8 0x0000 0x8000 0xB42C 0x4142 
' 	Add 
' 	St JD 
' Line #125:
' 	Ld JD 
' 	St calcJDFromJulianCent 
' Line #126:
' Line #127:
' 	EndFunc 
' Line #128:
' Line #129:
' Line #130:
' 	FuncDefn (Function calcGeomMeanLongSun(t))
' Line #131:
' Line #132:
' 	QuoteRem 0x0000 0x0048 "***********************************************************************/"
' Line #133:
' 	QuoteRem 0x0000 0x001D "* Name:    calGeomMeanLongSun"
' Line #134:
' 	QuoteRem 0x0000 0x0013 "* Type:    Function"
' Line #135:
' 	QuoteRem 0x0000 0x003C "* Purpose: calculate the Geometric Mean Longitude of the Sun"
' Line #136:
' 	QuoteRem 0x0000 0x000C "* Arguments:"
' Line #137:
' 	QuoteRem 0x0000 0x0030 "*   t : number of Julian centuries since J2000.0"
' Line #138:
' 	QuoteRem 0x0000 0x000F "* Return value:"
' Line #139:
' 	QuoteRem 0x0000 0x0036 "*   the Geometric Mean Longitude of the Sun in degrees"
' Line #140:
' 	QuoteRem 0x0000 0x0048 "***********************************************************************/"
' Line #141:
' Line #142:
' 	Dim 
' 	VarDefn l0 (As Double)
' Line #143:
' Line #144:
' 	LitR8 0xCE46 0x9EC2 0x8776 0x4071 
' 	Ld t 
' 	LitR8 0x862F 0xA272 0x9418 0x40E1 
' 	LitR8 0xABC0 0x158A 0xDEDA 0x3F33 
' 	Ld t 
' 	Mul 
' 	Add 
' 	Paren 
' 	Mul 
' 	Add 
' 	St l0 
' Line #145:
' 	Do 
' Line #146:
' 	Ld l0 
' 	LitDI2 0x0168 
' 	Le 
' 	Paren 
' 	Ld l0 
' 	LitDI2 0x0000 
' 	Ge 
' 	Paren 
' 	And 
' 	If 
' 	BoSImplicit 
' 	ExitDo 
' 	EndIf 
' Line #147:
' 	Ld l0 
' 	LitDI2 0x0168 
' 	Gt 
' 	If 
' 	BoSImplicit 
' 	Ld l0 
' 	LitDI2 0x0168 
' 	Sub 
' 	St l0 
' 	EndIf 
' Line #148:
' 	Ld l0 
' 	LitDI2 0x0000 
' 	Lt 
' 	If 
' 	BoSImplicit 
' 	Ld l0 
' 	LitDI2 0x0168 
' 	Add 
' 	St l0 
' 	EndIf 
' Line #149:
' 	Loop 
' Line #150:
' Line #151:
' 	Ld l0 
' 	St calcGeomMeanLongSun 
' Line #152:
' Line #153:
' 	EndFunc 
' Line #154:
' Line #155:
' Line #156:
' 	FuncDefn (Function calcGeomMeanAnomalySun(t))
' Line #157:
' Line #158:
' 	QuoteRem 0x0000 0x0048 "***********************************************************************/"
' Line #159:
' 	QuoteRem 0x0000 0x001C "* Name:    calGeomAnomalySun"
' Line #160:
' 	QuoteRem 0x0000 0x0013 "* Type:    Function"
' Line #161:
' 	QuoteRem 0x0000 0x003A "* Purpose: calculate the Geometric Mean Anomaly of the Sun"
' Line #162:
' 	QuoteRem 0x0000 0x000C "* Arguments:"
' Line #163:
' 	QuoteRem 0x0000 0x0030 "*   t : number of Julian centuries since J2000.0"
' Line #164:
' 	QuoteRem 0x0000 0x000F "* Return value:"
' Line #165:
' 	QuoteRem 0x0000 0x0034 "*   the Geometric Mean Anomaly of the Sun in degrees"
' Line #166:
' 	QuoteRem 0x0000 0x0048 "***********************************************************************/"
' Line #167:
' Line #168:
' 	Dim 
' 	VarDefn m (As Double)
' Line #169:
' Line #170:
' 	LitR8 0x1FC9 0x3C0C 0x5877 0x4076 
' 	Ld t 
' 	LitR8 0xC62A 0x9BF9 0x93E1 0x40E1 
' 	LitR8 0xDB0C 0xF260 0x2550 0x3F24 
' 	Ld t 
' 	Mul 
' 	Sub 
' 	Paren 
' 	Mul 
' 	Add 
' 	St m 
' Line #171:
' 	Ld m 
' 	St calcGeomMeanAnomalySun 
' Line #172:
' Line #173:
' 	EndFunc 
' Line #174:
' Line #175:
' Line #176:
' 	FuncDefn (Function calcEccentricityEarthOrbit(t))
' Line #177:
' Line #178:
' 	QuoteRem 0x0000 0x0048 "***********************************************************************/"
' Line #179:
' 	QuoteRem 0x0000 0x0025 "* Name:    calcEccentricityEarthOrbit"
' Line #180:
' 	QuoteRem 0x0000 0x0013 "* Type:    Function"
' Line #181:
' 	QuoteRem 0x0000 0x0036 "* Purpose: calculate the eccentricity of earth's orbit"
' Line #182:
' 	QuoteRem 0x0000 0x000C "* Arguments:"
' Line #183:
' 	QuoteRem 0x0000 0x0030 "*   t : number of Julian centuries since J2000.0"
' Line #184:
' 	QuoteRem 0x0000 0x000F "* Return value:"
' Line #185:
' 	QuoteRem 0x0000 0x001D "*   the unitless eccentricity"
' Line #186:
' 	QuoteRem 0x0000 0x0048 "***********************************************************************/"
' Line #187:
' Line #188:
' 	Dim 
' 	VarDefn e (As Double)
' Line #189:
' Line #190:
' 	LitR8 0x0380 0x725D 0x1C11 0x3F91 
' 	Ld t 
' 	LitR8 0xE303 0x525F 0x0A1C 0x3F06 
' 	LitR8 0xD800 0xFC64 0x0160 0x3E81 
' 	Ld t 
' 	Mul 
' 	Add 
' 	Paren 
' 	Mul 
' 	Sub 
' 	St e 
' Line #191:
' 	Ld e 
' 	St calcEccentricityEarthOrbit 
' Line #192:
' Line #193:
' 	EndFunc 
' Line #194:
' Line #195:
' Line #196:
' 	FuncDefn (Function calcSunEqOfCenter(t))
' Line #197:
' Line #198:
' 	QuoteRem 0x0000 0x0048 "***********************************************************************/"
' Line #199:
' 	QuoteRem 0x0000 0x001C "* Name:    calcSunEqOfCenter"
' Line #200:
' 	QuoteRem 0x0000 0x0013 "* Type:    Function"
' Line #201:
' 	QuoteRem 0x0000 0x0037 "* Purpose: calculate the equation of center for the sun"
' Line #202:
' 	QuoteRem 0x0000 0x000C "* Arguments:"
' Line #203:
' 	QuoteRem 0x0000 0x0030 "*   t : number of Julian centuries since J2000.0"
' Line #204:
' 	QuoteRem 0x0000 0x000F "* Return value:"
' Line #205:
' 	QuoteRem 0x0000 0x000E "*   in degrees"
' Line #206:
' 	QuoteRem 0x0000 0x0048 "***********************************************************************/"
' Line #207:
' Line #208:
' 	Dim 
' 	VarDefn m (As Double)
' 	VarDefn mrad (As Double)
' 	VarDefn sinm (As Double)
' 	VarDefn sin2m (As Double)
' 	VarDefn sin3m (As Double)
' Line #209:
' 	Dim 
' 	VarDefn c (As Double)
' Line #210:
' Line #211:
' 	Ld t 
' 	ArgsLd calcGeomMeanAnomalySun 0x0001 
' 	St m 
' Line #212:
' Line #213:
' 	Ld m 
' 	ArgsLd degToRad 0x0001 
' 	St mrad 
' Line #214:
' 	Ld mrad 
' 	ArgsLd Sin 0x0001 
' 	St sinm 
' Line #215:
' 	Ld mrad 
' 	Ld mrad 
' 	Add 
' 	ArgsLd Sin 0x0001 
' 	St sin2m 
' Line #216:
' 	Ld mrad 
' 	Ld mrad 
' 	Add 
' 	Ld mrad 
' 	Add 
' 	ArgsLd Sin 0x0001 
' 	St sin3m 
' Line #217:
' Line #218:
' 	LineCont 0x0004 11 00 0C 00
' 	Ld sinm 
' 	LitR8 0xB2F6 0xB4ED 0xA235 0x3FFE 
' 	Ld t 
' 	LitR8 0xF3AE 0x976F 0xBAFD 0x3F73 
' 	LitR8 0x5FB7 0x593E 0x5C31 0x3EED 
' 	Ld t 
' 	Mul 
' 	Add 
' 	Paren 
' 	Mul 
' 	Sub 
' 	Paren 
' 	Mul 
' 	Ld sin2m 
' 	LitR8 0x8095 0x8498 0x790B 0x3F94 
' 	LitR8 0x1AE3 0xC99F 0x79FE 0x3F1A 
' 	Ld t 
' 	Mul 
' 	Sub 
' 	Paren 
' 	Mul 
' 	Add 
' 	Ld sin3m 
' 	LitR8 0x612C 0x8C6D 0xF09D 0x3F32 
' 	Mul 
' 	Add 
' 	St c 
' Line #219:
' Line #220:
' 	Ld c 
' 	St calcSunEqOfCenter 
' Line #221:
' Line #222:
' 	EndFunc 
' Line #223:
' Line #224:
' Line #225:
' 	FuncDefn (Function calcSunTrueLong(t))
' Line #226:
' Line #227:
' 	QuoteRem 0x0000 0x0048 "***********************************************************************/"
' Line #228:
' 	QuoteRem 0x0000 0x001A "* Name:    calcSunTrueLong"
' Line #229:
' 	QuoteRem 0x0000 0x0013 "* Type:    Function"
' Line #230:
' 	QuoteRem 0x0000 0x0032 "* Purpose: calculate the true longitude of the sun"
' Line #231:
' 	QuoteRem 0x0000 0x000C "* Arguments:"
' Line #232:
' 	QuoteRem 0x0000 0x0030 "*   t : number of Julian centuries since J2000.0"
' Line #233:
' 	QuoteRem 0x0000 0x000F "* Return value:"
' Line #234:
' 	QuoteRem 0x0000 0x0023 "*   sun's true longitude in degrees"
' Line #235:
' 	QuoteRem 0x0000 0x0048 "***********************************************************************/"
' Line #236:
' Line #237:
' 	Dim 
' 	VarDefn l0 (As Double)
' 	VarDefn c (As Double)
' 	VarDefn O (As Double)
' Line #238:
' Line #239:
' 	Ld t 
' 	ArgsLd calcGeomMeanLongSun 0x0001 
' 	St l0 
' Line #240:
' 	Ld t 
' 	ArgsLd calcSunEqOfCenter 0x0001 
' 	St c 
' Line #241:
' Line #242:
' 	Ld l0 
' 	Ld c 
' 	Add 
' 	St O 
' Line #243:
' 	Ld O 
' 	St calcSunTrueLong 
' Line #244:
' Line #245:
' 	EndFunc 
' Line #246:
' Line #247:
' Line #248:
' 	FuncDefn (Function calcSunTrueAnomaly(t))
' Line #249:
' Line #250:
' 	QuoteRem 0x0000 0x0048 "***********************************************************************/"
' Line #251:
' 	QuoteRem 0x0000 0x0046 "* Name:    calcSunTrueAnomaly (not used by sunrise, solarnoon, sunset)"
' Line #252:
' 	QuoteRem 0x0000 0x0013 "* Type:    Function"
' Line #253:
' 	QuoteRem 0x0000 0x0030 "* Purpose: calculate the true anamoly of the sun"
' Line #254:
' 	QuoteRem 0x0000 0x000C "* Arguments:"
' Line #255:
' 	QuoteRem 0x0000 0x0030 "*   t : number of Julian centuries since J2000.0"
' Line #256:
' 	QuoteRem 0x0000 0x000F "* Return value:"
' Line #257:
' 	QuoteRem 0x0000 0x0021 "*   sun's true anamoly in degrees"
' Line #258:
' 	QuoteRem 0x0000 0x0048 "***********************************************************************/"
' Line #259:
' Line #260:
' 	Dim 
' 	VarDefn m (As Double)
' 	VarDefn c (As Double)
' 	VarDefn v (As Double)
' Line #261:
' Line #262:
' 	Ld t 
' 	ArgsLd calcGeomMeanAnomalySun 0x0001 
' 	St m 
' Line #263:
' 	Ld t 
' 	ArgsLd calcSunEqOfCenter 0x0001 
' 	St c 
' Line #264:
' Line #265:
' 	Ld m 
' 	Ld c 
' 	Add 
' 	St v 
' Line #266:
' 	Ld v 
' 	St calcSunTrueAnomaly 
' Line #267:
' Line #268:
' 	EndFunc 
' Line #269:
' Line #270:
' Line #271:
' 	FuncDefn (Function calcSunRadVector(t))
' Line #272:
' Line #273:
' 	QuoteRem 0x0000 0x0048 "***********************************************************************/"
' Line #274:
' 	QuoteRem 0x0000 0x0044 "* Name:    calcSunRadVector (not used by sunrise, solarnoon, sunset)"
' Line #275:
' 	QuoteRem 0x0000 0x0013 "* Type:    Function"
' Line #276:
' 	QuoteRem 0x0000 0x0032 "* Purpose: calculate the distance to the sun in AU"
' Line #277:
' 	QuoteRem 0x0000 0x000C "* Arguments:"
' Line #278:
' 	QuoteRem 0x0000 0x0030 "*   t : number of Julian centuries since J2000.0"
' Line #279:
' 	QuoteRem 0x0000 0x000F "* Return value:"
' Line #280:
' 	QuoteRem 0x0000 0x001C "*   sun radius vector in AUs"
' Line #281:
' 	QuoteRem 0x0000 0x0048 "***********************************************************************/"
' Line #282:
' Line #283:
' 	Dim 
' 	VarDefn v (As Double)
' 	VarDefn e (As Double)
' 	VarDefn R (As Double)
' Line #284:
' Line #285:
' 	Ld t 
' 	ArgsLd calcSunTrueAnomaly 0x0001 
' 	St v 
' Line #286:
' 	Ld t 
' 	ArgsLd calcEccentricityEarthOrbit 0x0001 
' 	St e 
' Line #287:
' Line #288:
' 	LitR8 0x6D65 0x1144 0x0001 0x3FF0 
' 	LitDI2 0x0001 
' 	Ld e 
' 	Ld e 
' 	Mul 
' 	Sub 
' 	Paren 
' 	Mul 
' 	Paren 
' 	LitDI2 0x0001 
' 	Ld e 
' 	Ld v 
' 	ArgsLd degToRad 0x0001 
' 	ArgsLd Cos 0x0001 
' 	Mul 
' 	Add 
' 	Paren 
' 	Div 
' 	St R 
' Line #289:
' 	Ld R 
' 	St calcSunRadVector 
' Line #290:
' Line #291:
' 	EndFunc 
' Line #292:
' Line #293:
' Line #294:
' 	FuncDefn (Function calcSunApparentLong(t))
' Line #295:
' Line #296:
' 	QuoteRem 0x0000 0x0048 "***********************************************************************/"
' Line #297:
' 	QuoteRem 0x0000 0x0047 "* Name:    calcSunApparentLong (not used by sunrise, solarnoon, sunset)"
' Line #298:
' 	QuoteRem 0x0000 0x0013 "* Type:    Function"
' Line #299:
' 	QuoteRem 0x0000 0x0036 "* Purpose: calculate the apparent longitude of the sun"
' Line #300:
' 	QuoteRem 0x0000 0x000C "* Arguments:"
' Line #301:
' 	QuoteRem 0x0000 0x0030 "*   t : number of Julian centuries since J2000.0"
' Line #302:
' 	QuoteRem 0x0000 0x000F "* Return value:"
' Line #303:
' 	QuoteRem 0x0000 0x0027 "*   sun's apparent longitude in degrees"
' Line #304:
' 	QuoteRem 0x0000 0x0048 "***********************************************************************/"
' Line #305:
' Line #306:
' 	Dim 
' 	VarDefn O (As Double)
' 	VarDefn omega (As Double)
' 	VarDefn lambda (As Double)
' Line #307:
' Line #308:
' 	Ld t 
' 	ArgsLd calcSunTrueLong 0x0001 
' 	St O 
' Line #309:
' Line #310:
' 	LitR8 0xF5C3 0x5C28 0x428F 0x405F 
' 	LitR8 0x8106 0x4395 0x388B 0x409E 
' 	Ld t 
' 	Mul 
' 	Sub 
' 	St omega 
' Line #311:
' 	Ld O 
' 	LitR8 0xBA1F 0xBEA0 0x4E65 0x3F77 
' 	Sub 
' 	LitR8 0x4EF9 0x7ACC 0x9431 0x3F73 
' 	Ld omega 
' 	ArgsLd degToRad 0x0001 
' 	ArgsLd Sin 0x0001 
' 	Mul 
' 	Sub 
' 	St lambda 
' Line #312:
' 	Ld lambda 
' 	St calcSunApparentLong 
' Line #313:
' Line #314:
' 	EndFunc 
' Line #315:
' Line #316:
' Line #317:
' 	FuncDefn (Function calcMeanObliquityOfEcliptic(t))
' Line #318:
' Line #319:
' 	QuoteRem 0x0000 0x0048 "***********************************************************************/"
' Line #320:
' 	QuoteRem 0x0000 0x0026 "* Name:    calcMeanObliquityOfEcliptic"
' Line #321:
' 	QuoteRem 0x0000 0x0013 "* Type:    Function"
' Line #322:
' 	QuoteRem 0x0000 0x0037 "* Purpose: calculate the mean obliquity of the ecliptic"
' Line #323:
' 	QuoteRem 0x0000 0x000C "* Arguments:"
' Line #324:
' 	QuoteRem 0x0000 0x0030 "*   t : number of Julian centuries since J2000.0"
' Line #325:
' 	QuoteRem 0x0000 0x000F "* Return value:"
' Line #326:
' 	QuoteRem 0x0000 0x001D "*   mean obliquity in degrees"
' Line #327:
' 	QuoteRem 0x0000 0x0048 "***********************************************************************/"
' Line #328:
' Line #329:
' 	Dim 
' 	VarDefn seconds (As Double)
' 	VarDefn e0 (As Double)
' Line #330:
' Line #331:
' 	LitR8 0x9BA6 0x20C4 0x72B0 0x4035 
' 	Ld t 
' 	LitR8 0x1EB8 0xEB85 0x6851 0x4047 
' 	Ld t 
' 	LitR8 0xA4BE 0x5A31 0x5547 0x3F43 
' 	Ld t 
' 	LitR8 0x1AD6 0xED4A 0xB445 0x3F5D 
' 	Paren 
' 	Mul 
' 	Sub 
' 	Paren 
' 	Mul 
' 	Add 
' 	Paren 
' 	Mul 
' 	Sub 
' 	St seconds 
' Line #332:
' 	LitR8 0x0000 0x0000 0x0000 0x4037 
' 	LitR8 0x0000 0x0000 0x0000 0x403A 
' 	Ld seconds 
' 	LitR8 0x0000 0x0000 0x0000 0x404E 
' 	Div 
' 	Paren 
' 	Add 
' 	Paren 
' 	LitR8 0x0000 0x0000 0x0000 0x404E 
' 	Div 
' 	Add 
' 	St e0 
' Line #333:
' 	Ld e0 
' 	St calcMeanObliquityOfEcliptic 
' Line #334:
' Line #335:
' 	EndFunc 
' Line #336:
' Line #337:
' Line #338:
' 	FuncDefn (Function calcObliquityCorrection(t))
' Line #339:
' Line #340:
' 	QuoteRem 0x0000 0x0048 "***********************************************************************/"
' Line #341:
' 	QuoteRem 0x0000 0x0022 "* Name:    calcObliquityCorrection"
' Line #342:
' 	QuoteRem 0x0000 0x0013 "* Type:    Function"
' Line #343:
' 	QuoteRem 0x0000 0x003C "* Purpose: calculate the corrected obliquity of the ecliptic"
' Line #344:
' 	QuoteRem 0x0000 0x000C "* Arguments:"
' Line #345:
' 	QuoteRem 0x0000 0x0030 "*   t : number of Julian centuries since J2000.0"
' Line #346:
' 	QuoteRem 0x0000 0x000F "* Return value:"
' Line #347:
' 	QuoteRem 0x0000 0x0022 "*   corrected obliquity in degrees"
' Line #348:
' 	QuoteRem 0x0000 0x0048 "***********************************************************************/"
' Line #349:
' Line #350:
' 	Dim 
' 	VarDefn e0 (As Double)
' 	VarDefn omega (As Double)
' 	VarDefn e (As Double)
' Line #351:
' Line #352:
' 	Ld t 
' 	ArgsLd calcMeanObliquityOfEcliptic 0x0001 
' 	St e0 
' Line #353:
' Line #354:
' 	LitR8 0xF5C3 0x5C28 0x428F 0x405F 
' 	LitR8 0x8106 0x4395 0x388B 0x409E 
' 	Ld t 
' 	Mul 
' 	Sub 
' 	St omega 
' Line #355:
' 	Ld e0 
' 	LitR8 0x68F1 0x88E3 0xF8B5 0x3F64 
' 	Ld omega 
' 	ArgsLd degToRad 0x0001 
' 	ArgsLd Cos 0x0001 
' 	Mul 
' 	Add 
' 	St e 
' Line #356:
' 	Ld e 
' 	St calcObliquityCorrection 
' Line #357:
' Line #358:
' 	EndFunc 
' Line #359:
' Line #360:
' Line #361:
' 	FuncDefn (Function calcSunRtAscension(t))
' Line #362:
' Line #363:
' 	QuoteRem 0x0000 0x0048 "***********************************************************************/"
' Line #364:
' 	QuoteRem 0x0000 0x0046 "* Name:    calcSunRtAscension (not used by sunrise, solarnoon, sunset)"
' Line #365:
' 	QuoteRem 0x0000 0x0013 "* Type:    Function"
' Line #366:
' 	QuoteRem 0x0000 0x0033 "* Purpose: calculate the right ascension of the sun"
' Line #367:
' 	QuoteRem 0x0000 0x000C "* Arguments:"
' Line #368:
' 	QuoteRem 0x0000 0x0030 "*   t : number of Julian centuries since J2000.0"
' Line #369:
' 	QuoteRem 0x0000 0x000F "* Return value:"
' Line #370:
' 	QuoteRem 0x0000 0x0024 "*   sun's right ascension in degrees"
' Line #371:
' 	QuoteRem 0x0000 0x0048 "***********************************************************************/"
' Line #372:
' Line #373:
' 	Dim 
' 	VarDefn e (As Double)
' 	VarDefn lambda (As Double)
' 	VarDefn tananum (As Double)
' 	VarDefn tanadenom (As Double)
' Line #374:
' 	Dim 
' 	VarDefn alpha (As Double)
' Line #375:
' Line #376:
' 	Ld t 
' 	ArgsLd calcObliquityCorrection 0x0001 
' 	St e 
' Line #377:
' 	Ld t 
' 	ArgsLd calcSunApparentLong 0x0001 
' 	St lambda 
' Line #378:
' Line #379:
' 	Ld e 
' 	ArgsLd degToRad 0x0001 
' 	ArgsLd Cos 0x0001 
' 	Ld lambda 
' 	ArgsLd degToRad 0x0001 
' 	ArgsLd Sin 0x0001 
' 	Mul 
' 	Paren 
' 	St tananum 
' Line #380:
' 	Ld lambda 
' 	ArgsLd degToRad 0x0001 
' 	ArgsLd Cos 0x0001 
' 	Paren 
' 	St tanadenom 
' Line #381:
' Line #382:
' 	QuoteRem 0x0000 0x003F "original NOAA code using javascript Math.Atan2(y,x) convention:"
' Line #383:
' 	QuoteRem 0x0000 0x003D "        var alpha = radToDeg(Math.atan2(tananum, tanadenom));"
' Line #384:
' 	QuoteRem 0x0000 0x0051 "        alpha = radToDeg(Application.WorksheetFunction.Atan2(tananum, tanadenom))"
' Line #385:
' Line #386:
' 	QuoteRem 0x0000 0x004F "translated using Excel VBA Application.WorksheetFunction.Atan2(x,y) convention:"
' Line #387:
' 	Ld tanadenom 
' 	Ld tananum 
' 	Ld Application 
' 	MemLd WorksheetFunction 
' 	ArgsMemLd Atan2 0x0002 
' 	ArgsLd radToDeg 0x0001 
' 	St alpha 
' Line #388:
' Line #389:
' 	Ld alpha 
' 	St calcSunRtAscension 
' Line #390:
' Line #391:
' 	EndFunc 
' Line #392:
' Line #393:
' Line #394:
' 	FuncDefn (Function calcSunDeclination(t))
' Line #395:
' Line #396:
' 	QuoteRem 0x0000 0x0048 "***********************************************************************/"
' Line #397:
' 	QuoteRem 0x0000 0x001D "* Name:    calcSunDeclination"
' Line #398:
' 	QuoteRem 0x0000 0x0013 "* Type:    Function"
' Line #399:
' 	QuoteRem 0x0000 0x002F "* Purpose: calculate the declination of the sun"
' Line #400:
' 	QuoteRem 0x0000 0x000C "* Arguments:"
' Line #401:
' 	QuoteRem 0x0000 0x0030 "*   t : number of Julian centuries since J2000.0"
' Line #402:
' 	QuoteRem 0x0000 0x000F "* Return value:"
' Line #403:
' 	QuoteRem 0x0000 0x0020 "*   sun's declination in degrees"
' Line #404:
' 	QuoteRem 0x0000 0x0048 "***********************************************************************/"
' Line #405:
' Line #406:
' 	Dim 
' 	VarDefn e (As Double)
' 	VarDefn lambda (As Double)
' 	VarDefn sint (As Double)
' 	VarDefn theta (As Double)
' Line #407:
' Line #408:
' 	Ld t 
' 	ArgsLd calcObliquityCorrection 0x0001 
' 	St e 
' Line #409:
' 	Ld t 
' 	ArgsLd calcSunApparentLong 0x0001 
' 	St lambda 
' Line #410:
' Line #411:
' 	Ld e 
' 	ArgsLd degToRad 0x0001 
' 	ArgsLd Sin 0x0001 
' 	Ld lambda 
' 	ArgsLd degToRad 0x0001 
' 	ArgsLd Sin 0x0001 
' 	Mul 
' 	St sint 
' Line #412:
' 	Ld sint 
' 	Ld Application 
' 	MemLd WorksheetFunction 
' 	ArgsMemLd Asin 0x0001 
' 	ArgsLd radToDeg 0x0001 
' 	St theta 
' Line #413:
' 	Ld theta 
' 	St calcSunDeclination 
' Line #414:
' Line #415:
' 	EndFunc 
' Line #416:
' Line #417:
' Line #418:
' 	FuncDefn (Function calcEquationOfTime(t))
' Line #419:
' Line #420:
' 	QuoteRem 0x0000 0x0048 "***********************************************************************/"
' Line #421:
' 	QuoteRem 0x0000 0x001D "* Name:    calcEquationOfTime"
' Line #422:
' 	QuoteRem 0x0000 0x0013 "* Type:    Function"
' Line #423:
' 	QuoteRem 0x0000 0x0044 "* Purpose: calculate the difference between true solar time and mean"
' Line #424:
' 	QuoteRem 0x0000 0x0010 "*     solar time"
' Line #425:
' 	QuoteRem 0x0000 0x000C "* Arguments:"
' Line #426:
' 	QuoteRem 0x0000 0x0030 "*   t : number of Julian centuries since J2000.0"
' Line #427:
' 	QuoteRem 0x0000 0x000F "* Return value:"
' Line #428:
' 	QuoteRem 0x0000 0x0027 "*   equation of time in minutes of time"
' Line #429:
' 	QuoteRem 0x0000 0x0048 "***********************************************************************/"
' Line #430:
' Line #431:
' 	Dim 
' 	VarDefn epsilon (As Double)
' 	VarDefn l0 (As Double)
' 	VarDefn e (As Double)
' 	VarDefn m (As Double)
' Line #432:
' 	Dim 
' 	VarDefn y (As Double)
' 	VarDefn sin2l0 (As Double)
' 	VarDefn sinm (As Double)
' Line #433:
' 	Dim 
' 	VarDefn cos2l0 (As Double)
' 	VarDefn sin4l0 (As Double)
' 	VarDefn sin2m (As Double)
' 	VarDefn Etime (As Double)
' Line #434:
' Line #435:
' 	Ld t 
' 	ArgsLd calcObliquityCorrection 0x0001 
' 	St epsilon 
' Line #436:
' 	Ld t 
' 	ArgsLd calcGeomMeanLongSun 0x0001 
' 	St l0 
' Line #437:
' 	Ld t 
' 	ArgsLd calcEccentricityEarthOrbit 0x0001 
' 	St e 
' Line #438:
' 	Ld t 
' 	ArgsLd calcGeomMeanAnomalySun 0x0001 
' 	St m 
' Line #439:
' Line #440:
' 	Ld epsilon 
' 	ArgsLd degToRad 0x0001 
' 	LitR8 0x0000 0x0000 0x0000 0x4000 
' 	Div 
' 	ArgsLd Tan 0x0001 
' 	St y 
' Line #441:
' 	Ld y 
' 	LitDI2 0x0002 
' 	Pwr 
' 	St y 
' Line #442:
' Line #443:
' 	LitR8 0x0000 0x0000 0x0000 0x4000 
' 	Ld l0 
' 	ArgsLd degToRad 0x0001 
' 	Mul 
' 	ArgsLd Sin 0x0001 
' 	St sin2l0 
' Line #444:
' 	Ld m 
' 	ArgsLd degToRad 0x0001 
' 	ArgsLd Sin 0x0001 
' 	St sinm 
' Line #445:
' 	LitR8 0x0000 0x0000 0x0000 0x4000 
' 	Ld l0 
' 	ArgsLd degToRad 0x0001 
' 	Mul 
' 	ArgsLd Cos 0x0001 
' 	St cos2l0 
' Line #446:
' 	LitR8 0x0000 0x0000 0x0000 0x4010 
' 	Ld l0 
' 	ArgsLd degToRad 0x0001 
' 	Mul 
' 	ArgsLd Sin 0x0001 
' 	St sin4l0 
' Line #447:
' 	LitR8 0x0000 0x0000 0x0000 0x4000 
' 	Ld m 
' 	ArgsLd degToRad 0x0001 
' 	Mul 
' 	ArgsLd Sin 0x0001 
' 	St sin2m 
' Line #448:
' Line #449:
' 	LineCont 0x0004 15 00 10 00
' 	Ld y 
' 	Ld sin2l0 
' 	Mul 
' 	LitR8 0x0000 0x0000 0x0000 0x4000 
' 	Ld e 
' 	Mul 
' 	Ld sinm 
' 	Mul 
' 	Sub 
' 	LitR8 0x0000 0x0000 0x0000 0x4010 
' 	Ld e 
' 	Mul 
' 	Ld y 
' 	Mul 
' 	Ld sinm 
' 	Mul 
' 	Ld cos2l0 
' 	Mul 
' 	Add 
' 	LitR8 0x0000 0x0000 0x0000 0x3FE0 
' 	Ld y 
' 	Mul 
' 	Ld y 
' 	Mul 
' 	Ld sin4l0 
' 	Mul 
' 	Sub 
' 	LitR8 0x0000 0x0000 0x0000 0x3FF4 
' 	Ld e 
' 	Mul 
' 	Ld e 
' 	Mul 
' 	Ld sin2m 
' 	Mul 
' 	Sub 
' 	St Etime 
' Line #450:
' Line #451:
' 	Ld Etime 
' 	ArgsLd radToDeg 0x0001 
' 	LitR8 0x0000 0x0000 0x0000 0x4010 
' 	Mul 
' 	St calcEquationOfTime 
' Line #452:
' Line #453:
' 	EndFunc 
' Line #454:
' Line #455:
' Line #456:
' 	FuncDefn (Function calcHourAngleDawn(lat, solarDec, solardepression))
' Line #457:
' Line #458:
' 	QuoteRem 0x0000 0x0048 "***********************************************************************/"
' Line #459:
' 	QuoteRem 0x0000 0x001C "* Name:    calcHourAngleDawn"
' Line #460:
' 	QuoteRem 0x0000 0x0013 "* Type:    Function"
' Line #461:
' 	QuoteRem 0x0000 0x003E "* Purpose: calculate the hour angle of the sun at dawn for the"
' Line #462:
' 	QuoteRem 0x0000 0x0012 "*         latitude"
' Line #463:
' 	QuoteRem 0x0000 0x003A "*         for user selected solar depression below horizon"
' Line #464:
' 	QuoteRem 0x0000 0x000C "* Arguments:"
' Line #465:
' 	QuoteRem 0x0000 0x0029 "*   lat : latitude of observer in degrees"
' Line #466:
' 	QuoteRem 0x0000 0x0032 "*   solarDec : declination angle of sun in degrees"
' Line #467:
' 	QuoteRem 0x0000 0x0043 "*   solardepression: angle of the sun below the horizion in degrees"
' Line #468:
' 	QuoteRem 0x0000 0x000F "* Return value:"
' Line #469:
' 	QuoteRem 0x0000 0x0021 "*   hour angle of dawn in radians"
' Line #470:
' 	QuoteRem 0x0000 0x0048 "***********************************************************************/"
' Line #471:
' Line #472:
' 	Dim 
' 	VarDefn latRad (As Double)
' 	VarDefn sdRad (As Double)
' 	VarDefn HAarg (As Double)
' 	VarDefn HA (As Double)
' Line #473:
' Line #474:
' 	Ld lat 
' 	ArgsLd degToRad 0x0001 
' 	St latRad 
' Line #475:
' 	Ld solarDec 
' 	ArgsLd degToRad 0x0001 
' 	St sdRad 
' Line #476:
' Line #477:
' 	LitDI2 0x005A 
' 	Ld solardepression 
' 	Add 
' 	ArgsLd degToRad 0x0001 
' 	ArgsLd Cos 0x0001 
' 	Ld latRad 
' 	ArgsLd Cos 0x0001 
' 	Ld sdRad 
' 	ArgsLd Cos 0x0001 
' 	Mul 
' 	Paren 
' 	Div 
' 	Ld latRad 
' 	ArgsLd Tan 0x0001 
' 	Ld sdRad 
' 	ArgsLd Tan 0x0001 
' 	Mul 
' 	Sub 
' 	Paren 
' 	St HAarg 
' Line #478:
' Line #479:
' 	LineCont 0x0004 12 00 0E 00
' 	LitDI2 0x005A 
' 	Ld solardepression 
' 	Add 
' 	ArgsLd degToRad 0x0001 
' 	ArgsLd Cos 0x0001 
' 	Ld latRad 
' 	ArgsLd Cos 0x0001 
' 	Ld sdRad 
' 	ArgsLd Cos 0x0001 
' 	Mul 
' 	Paren 
' 	Div 
' 	Ld latRad 
' 	ArgsLd Tan 0x0001 
' 	Ld sdRad 
' 	ArgsLd Tan 0x0001 
' 	Mul 
' 	Sub 
' 	Ld Application 
' 	MemLd WorksheetFunction 
' 	ArgsMemLd Acos 0x0001 
' 	Paren 
' 	St HA 
' Line #480:
' Line #481:
' 	Ld HA 
' 	St calcHourAngleDawn 
' Line #482:
' Line #483:
' 	EndFunc 
' Line #484:
' Line #485:
' Line #486:
' 	FuncDefn (Function calcHourAngleSunrise(lat, solarDec))
' Line #487:
' Line #488:
' 	QuoteRem 0x0000 0x0048 "***********************************************************************/"
' Line #489:
' 	QuoteRem 0x0000 0x001F "* Name:    calcHourAngleSunrise"
' Line #490:
' 	QuoteRem 0x0000 0x0013 "* Type:    Function"
' Line #491:
' 	QuoteRem 0x0000 0x0041 "* Purpose: calculate the hour angle of the sun at sunrise for the"
' Line #492:
' 	QuoteRem 0x0000 0x0012 "*         latitude"
' Line #493:
' 	QuoteRem 0x0000 0x000C "* Arguments:"
' Line #494:
' 	QuoteRem 0x0000 0x0029 "*   lat : latitude of observer in degrees"
' Line #495:
' 	QuoteRem 0x0000 0x0030 "* solarDec : declination angle of sun in degrees"
' Line #496:
' 	QuoteRem 0x0000 0x000F "* Return value:"
' Line #497:
' 	QuoteRem 0x0000 0x0024 "*   hour angle of sunrise in radians"
' Line #498:
' 	QuoteRem 0x0000 0x0001 "*"
' Line #499:
' 	QuoteRem 0x0000 0x0057 "* Note: For sunrise and sunset calculations, we assume 0.833° of atmospheric refraction"
' Line #500:
' 	QuoteRem 0x0000 0x005F "* For details about refraction see http://www.srrb.noaa.gov/highlights/sunrise/calcdetails.html"
' Line #501:
' 	QuoteRem 0x0000 0x0001 "*"
' Line #502:
' 	QuoteRem 0x0000 0x0048 "***********************************************************************/"
' Line #503:
' Line #504:
' 	Dim 
' 	VarDefn latRad (As Double)
' 	VarDefn sdRad (As Double)
' 	VarDefn HAarg (As Double)
' 	VarDefn HA (As Double)
' Line #505:
' Line #506:
' 	Ld lat 
' 	ArgsLd degToRad 0x0001 
' 	St latRad 
' Line #507:
' 	Ld solarDec 
' 	ArgsLd degToRad 0x0001 
' 	St sdRad 
' Line #508:
' Line #509:
' 	LitR8 0x645A 0xDF3B 0xB54F 0x4056 
' 	ArgsLd degToRad 0x0001 
' 	ArgsLd Cos 0x0001 
' 	Ld latRad 
' 	ArgsLd Cos 0x0001 
' 	Ld sdRad 
' 	ArgsLd Cos 0x0001 
' 	Mul 
' 	Paren 
' 	Div 
' 	Ld latRad 
' 	ArgsLd Tan 0x0001 
' 	Ld sdRad 
' 	ArgsLd Tan 0x0001 
' 	Mul 
' 	Sub 
' 	Paren 
' 	St HAarg 
' Line #510:
' Line #511:
' 	LineCont 0x0004 10 00 0E 00
' 	LitR8 0x645A 0xDF3B 0xB54F 0x4056 
' 	ArgsLd degToRad 0x0001 
' 	ArgsLd Cos 0x0001 
' 	Ld latRad 
' 	ArgsLd Cos 0x0001 
' 	Ld sdRad 
' 	ArgsLd Cos 0x0001 
' 	Mul 
' 	Paren 
' 	Div 
' 	Ld latRad 
' 	ArgsLd Tan 0x0001 
' 	Ld sdRad 
' 	ArgsLd Tan 0x0001 
' 	Mul 
' 	Sub 
' 	Ld Application 
' 	MemLd WorksheetFunction 
' 	ArgsMemLd Acos 0x0001 
' 	Paren 
' 	St HA 
' Line #512:
' Line #513:
' 	Ld HA 
' 	St calcHourAngleSunrise 
' Line #514:
' Line #515:
' 	EndFunc 
' Line #516:
' Line #517:
' Line #518:
' 	FuncDefn (Function calcHourAngleSunset(lat, solarDec))
' Line #519:
' Line #520:
' 	QuoteRem 0x0000 0x0048 "***********************************************************************/"
' Line #521:
' 	QuoteRem 0x0000 0x001E "* Name:    calcHourAngleSunset"
' Line #522:
' 	QuoteRem 0x0000 0x0013 "* Type:    Function"
' Line #523:
' 	QuoteRem 0x0000 0x0040 "* Purpose: calculate the hour angle of the sun at sunset for the"
' Line #524:
' 	QuoteRem 0x0000 0x0012 "*         latitude"
' Line #525:
' 	QuoteRem 0x0000 0x000C "* Arguments:"
' Line #526:
' 	QuoteRem 0x0000 0x0029 "*   lat : latitude of observer in degrees"
' Line #527:
' 	QuoteRem 0x0000 0x0030 "* solarDec : declination angle of sun in degrees"
' Line #528:
' 	QuoteRem 0x0000 0x000F "* Return value:"
' Line #529:
' 	QuoteRem 0x0000 0x0023 "*   hour angle of sunset in radians"
' Line #530:
' 	QuoteRem 0x0000 0x0001 "*"
' Line #531:
' 	QuoteRem 0x0000 0x0057 "* Note: For sunrise and sunset calculations, we assume 0.833° of atmospheric refraction"
' Line #532:
' 	QuoteRem 0x0000 0x005F "* For details about refraction see http://www.srrb.noaa.gov/highlights/sunrise/calcdetails.html"
' Line #533:
' 	QuoteRem 0x0000 0x0001 "*"
' Line #534:
' 	QuoteRem 0x0000 0x0048 "***********************************************************************/"
' Line #535:
' Line #536:
' 	Dim 
' 	VarDefn latRad (As Double)
' 	VarDefn sdRad (As Double)
' 	VarDefn HAarg (As Double)
' 	VarDefn HA (As Double)
' Line #537:
' Line #538:
' 	Ld lat 
' 	ArgsLd degToRad 0x0001 
' 	St latRad 
' Line #539:
' 	Ld solarDec 
' 	ArgsLd degToRad 0x0001 
' 	St sdRad 
' Line #540:
' Line #541:
' 	LitR8 0x645A 0xDF3B 0xB54F 0x4056 
' 	ArgsLd degToRad 0x0001 
' 	ArgsLd Cos 0x0001 
' 	Ld latRad 
' 	ArgsLd Cos 0x0001 
' 	Ld sdRad 
' 	ArgsLd Cos 0x0001 
' 	Mul 
' 	Paren 
' 	Div 
' 	Ld latRad 
' 	ArgsLd Tan 0x0001 
' 	Ld sdRad 
' 	ArgsLd Tan 0x0001 
' 	Mul 
' 	Sub 
' 	Paren 
' 	St HAarg 
' Line #542:
' Line #543:
' 	LineCont 0x0004 10 00 0F 00
' 	LitR8 0x645A 0xDF3B 0xB54F 0x4056 
' 	ArgsLd degToRad 0x0001 
' 	ArgsLd Cos 0x0001 
' 	Ld latRad 
' 	ArgsLd Cos 0x0001 
' 	Ld sdRad 
' 	ArgsLd Cos 0x0001 
' 	Mul 
' 	Paren 
' 	Div 
' 	Ld latRad 
' 	ArgsLd Tan 0x0001 
' 	Ld sdRad 
' 	ArgsLd Tan 0x0001 
' 	Mul 
' 	Sub 
' 	Ld Application 
' 	MemLd WorksheetFunction 
' 	ArgsMemLd Acos 0x0001 
' 	Paren 
' 	St HA 
' Line #544:
' Line #545:
' 	Ld HA 
' 	UMi 
' 	St calcHourAngleSunset 
' Line #546:
' Line #547:
' 	EndFunc 
' Line #548:
' Line #549:
' Line #550:
' 	FuncDefn (Function calcHourAngleDusk(lat, solarDec, solardepression))
' Line #551:
' Line #552:
' 	QuoteRem 0x0000 0x0048 "***********************************************************************/"
' Line #553:
' 	QuoteRem 0x0000 0x001C "* Name:    calcHourAngleDusk"
' Line #554:
' 	QuoteRem 0x0000 0x0013 "* Type:    Function"
' Line #555:
' 	QuoteRem 0x0000 0x003E "* Purpose: calculate the hour angle of the sun at dusk for the"
' Line #556:
' 	QuoteRem 0x0000 0x0012 "*         latitude"
' Line #557:
' 	QuoteRem 0x0000 0x003A "*         for user selected solar depression below horizon"
' Line #558:
' 	QuoteRem 0x0000 0x000C "* Arguments:"
' Line #559:
' 	QuoteRem 0x0000 0x0029 "*   lat : latitude of observer in degrees"
' Line #560:
' 	QuoteRem 0x0000 0x0032 "*   solarDec : declination angle of sun in degrees"
' Line #561:
' 	QuoteRem 0x0000 0x003A "*   solardepression: angle of sun below horizon in degrees"
' Line #562:
' 	QuoteRem 0x0000 0x000F "* Return value:"
' Line #563:
' 	QuoteRem 0x0000 0x0021 "*   hour angle of dusk in radians"
' Line #564:
' 	QuoteRem 0x0000 0x0048 "***********************************************************************/"
' Line #565:
' Line #566:
' 	Dim 
' 	VarDefn latRad (As Double)
' 	VarDefn sdRad (As Double)
' 	VarDefn HAarg (As Double)
' 	VarDefn HA (As Double)
' Line #567:
' Line #568:
' 	Ld lat 
' 	ArgsLd degToRad 0x0001 
' 	St latRad 
' Line #569:
' 	Ld solarDec 
' 	ArgsLd degToRad 0x0001 
' 	St sdRad 
' Line #570:
' Line #571:
' 	LitDI2 0x005A 
' 	Ld solardepression 
' 	Add 
' 	ArgsLd degToRad 0x0001 
' 	ArgsLd Cos 0x0001 
' 	Ld latRad 
' 	ArgsLd Cos 0x0001 
' 	Ld sdRad 
' 	ArgsLd Cos 0x0001 
' 	Mul 
' 	Paren 
' 	Div 
' 	Ld latRad 
' 	ArgsLd Tan 0x0001 
' 	Ld sdRad 
' 	ArgsLd Tan 0x0001 
' 	Mul 
' 	Sub 
' 	Paren 
' 	St HAarg 
' Line #572:
' Line #573:
' 	LineCont 0x0004 12 00 0F 00
' 	LitDI2 0x005A 
' 	Ld solardepression 
' 	Add 
' 	ArgsLd degToRad 0x0001 
' 	ArgsLd Cos 0x0001 
' 	Ld latRad 
' 	ArgsLd Cos 0x0001 
' 	Ld sdRad 
' 	ArgsLd Cos 0x0001 
' 	Mul 
' 	Paren 
' 	Div 
' 	Ld latRad 
' 	ArgsLd Tan 0x0001 
' 	Ld sdRad 
' 	ArgsLd Tan 0x0001 
' 	Mul 
' 	Sub 
' 	Ld Application 
' 	MemLd WorksheetFunction 
' 	ArgsMemLd Acos 0x0001 
' 	Paren 
' 	St HA 
' Line #574:
' Line #575:
' 	Ld HA 
' 	UMi 
' 	St calcHourAngleDusk 
' Line #576:
' Line #577:
' 	EndFunc 
' Line #578:
' Line #579:
' Line #580:
' 	FuncDefn (Function calcDawnUTC(JD, latitude, longitude, solardepression))
' Line #581:
' Line #582:
' 	QuoteRem 0x0000 0x0048 "***********************************************************************/"
' Line #583:
' 	QuoteRem 0x0000 0x0016 "* Name:    calcDawnUTC"
' Line #584:
' 	QuoteRem 0x0000 0x0013 "* Type:    Function"
' Line #585:
' 	QuoteRem 0x0000 0x0041 "* Purpose: calculate the Universal Coordinated Time (UTC) of dawn"
' Line #586:
' 	QuoteRem 0x0000 0x003A "*         for the given day at the given location on earth"
' Line #587:
' 	QuoteRem 0x0000 0x003A "*         for user selected solar depression below horizon"
' Line #588:
' 	QuoteRem 0x0000 0x000C "* Arguments:"
' Line #589:
' 	QuoteRem 0x0000 0x0014 "*   JD  : julian day"
' Line #590:
' 	QuoteRem 0x0000 0x002E "*   latitude : latitude of observer in degrees"
' Line #591:
' 	QuoteRem 0x0000 0x0030 "*   longitude : longitude of observer in degrees"
' Line #592:
' 	QuoteRem 0x0000 0x003E "*   solardepression: angle of sun below the horizon in degrees"
' Line #593:
' 	QuoteRem 0x0000 0x000F "* Return value:"
' Line #594:
' 	QuoteRem 0x0000 0x001F "*   time in minutes from zero Z"
' Line #595:
' 	QuoteRem 0x0000 0x0048 "***********************************************************************/"
' Line #596:
' Line #597:
' 	Dim 
' 	VarDefn t (As Double)
' 	VarDefn eqtime (As Double)
' 	VarDefn solarDec (As Double)
' 	VarDefn hourangle (As Double)
' Line #598:
' 	Dim 
' 	VarDefn delta (As Double)
' 	VarDefn timeDiff (As Double)
' 	VarDefn timeUTC (As Double)
' Line #599:
' 	Dim 
' 	VarDefn newt (As Double)
' Line #600:
' Line #601:
' 	Ld JD 
' 	ArgsLd calcTimeJulianCent 0x0001 
' 	St t 
' Line #602:
' Line #603:
' 	QuoteRem 0x0000 0x0030 "        // *** First pass to approximate sunrise"
' Line #604:
' Line #605:
' 	Ld t 
' 	ArgsLd calcEquationOfTime 0x0001 
' 	St eqtime 
' Line #606:
' 	Ld t 
' 	ArgsLd calcSunDeclination 0x0001 
' 	St solarDec 
' Line #607:
' 	Ld latitude 
' 	Ld solarDec 
' 	ArgsLd calcHourAngleSunrise 0x0002 
' 	St hourangle 
' Line #608:
' Line #609:
' 	Ld longitude 
' 	Ld hourangle 
' 	ArgsLd radToDeg 0x0001 
' 	Sub 
' 	St delta 
' Line #610:
' 	LitDI2 0x0004 
' 	Ld delta 
' 	Mul 
' 	St timeDiff 
' Line #611:
' 	QuoteRem 0x0000 0x0013 " in minutes of time"
' Line #612:
' 	LitDI2 0x02D0 
' 	Ld timeDiff 
' 	Add 
' 	Ld eqtime 
' 	Sub 
' 	St timeUTC 
' Line #613:
' 	QuoteRem 0x0000 0x000B " in minutes"
' Line #614:
' Line #615:
' 	QuoteRem 0x0000 0x0037 " *** Second pass includes fractional jday in gamma calc"
' Line #616:
' Line #617:
' 	Ld t 
' 	ArgsLd calcJDFromJulianCent 0x0001 
' 	Ld timeUTC 
' 	LitR8 0x0000 0x0000 0x8000 0x4096 
' 	Div 
' 	Add 
' 	ArgsLd calcTimeJulianCent 0x0001 
' 	St newt 
' Line #618:
' 	Ld newt 
' 	ArgsLd calcEquationOfTime 0x0001 
' 	St eqtime 
' Line #619:
' 	Ld newt 
' 	ArgsLd calcSunDeclination 0x0001 
' 	St solarDec 
' Line #620:
' 	Ld latitude 
' 	Ld solarDec 
' 	Ld solardepression 
' 	ArgsLd calcHourAngleDawn 0x0003 
' 	St hourangle 
' Line #621:
' 	Ld longitude 
' 	Ld hourangle 
' 	ArgsLd radToDeg 0x0001 
' 	Sub 
' 	St delta 
' Line #622:
' 	LitDI2 0x0004 
' 	Ld delta 
' 	Mul 
' 	St timeDiff 
' Line #623:
' 	LitDI2 0x02D0 
' 	Ld timeDiff 
' 	Add 
' 	Ld eqtime 
' 	Sub 
' 	St timeUTC 
' Line #624:
' 	QuoteRem 0x0000 0x000B " in minutes"
' Line #625:
' Line #626:
' 	Ld timeUTC 
' 	St calcDawnUTC 
' Line #627:
' Line #628:
' 	EndFunc 
' Line #629:
' Line #630:
' Line #631:
' 	FuncDefn (Function calcSunriseUTC(JD, latitude, longitude))
' Line #632:
' Line #633:
' 	QuoteRem 0x0000 0x0048 "***********************************************************************/"
' Line #634:
' 	QuoteRem 0x0000 0x0019 "* Name:    calcSunriseUTC"
' Line #635:
' 	QuoteRem 0x0000 0x0013 "* Type:    Function"
' Line #636:
' 	QuoteRem 0x0000 0x0044 "* Purpose: calculate the Universal Coordinated Time (UTC) of sunrise"
' Line #637:
' 	QuoteRem 0x0000 0x003A "*         for the given day at the given location on earth"
' Line #638:
' 	QuoteRem 0x0000 0x000C "* Arguments:"
' Line #639:
' 	QuoteRem 0x0000 0x0014 "*   JD  : julian day"
' Line #640:
' 	QuoteRem 0x0000 0x002E "*   latitude : latitude of observer in degrees"
' Line #641:
' 	QuoteRem 0x0000 0x0030 "*   longitude : longitude of observer in degrees"
' Line #642:
' 	QuoteRem 0x0000 0x000F "* Return value:"
' Line #643:
' 	QuoteRem 0x0000 0x001F "*   time in minutes from zero Z"
' Line #644:
' 	QuoteRem 0x0000 0x0048 "***********************************************************************/"
' Line #645:
' Line #646:
' 	Dim 
' 	VarDefn t (As Double)
' 	VarDefn eqtime (As Double)
' 	VarDefn solarDec (As Double)
' 	VarDefn hourangle (As Double)
' Line #647:
' 	Dim 
' 	VarDefn delta (As Double)
' 	VarDefn timeDiff (As Double)
' 	VarDefn timeUTC (As Double)
' Line #648:
' 	Dim 
' 	VarDefn newt (As Double)
' Line #649:
' Line #650:
' 	Ld JD 
' 	ArgsLd calcTimeJulianCent 0x0001 
' 	St t 
' Line #651:
' Line #652:
' 	QuoteRem 0x0000 0x0030 "        // *** First pass to approximate sunrise"
' Line #653:
' Line #654:
' 	Ld t 
' 	ArgsLd calcEquationOfTime 0x0001 
' 	St eqtime 
' Line #655:
' 	Ld t 
' 	ArgsLd calcSunDeclination 0x0001 
' 	St solarDec 
' Line #656:
' 	Ld latitude 
' 	Ld solarDec 
' 	ArgsLd calcHourAngleSunrise 0x0002 
' 	St hourangle 
' Line #657:
' Line #658:
' 	Ld longitude 
' 	Ld hourangle 
' 	ArgsLd radToDeg 0x0001 
' 	Sub 
' 	St delta 
' Line #659:
' 	LitDI2 0x0004 
' 	Ld delta 
' 	Mul 
' 	St timeDiff 
' Line #660:
' 	QuoteRem 0x0000 0x0013 " in minutes of time"
' Line #661:
' 	LitDI2 0x02D0 
' 	Ld timeDiff 
' 	Add 
' 	Ld eqtime 
' 	Sub 
' 	St timeUTC 
' Line #662:
' 	QuoteRem 0x0000 0x000B " in minutes"
' Line #663:
' Line #664:
' 	QuoteRem 0x0000 0x0037 " *** Second pass includes fractional jday in gamma calc"
' Line #665:
' Line #666:
' 	Ld t 
' 	ArgsLd calcJDFromJulianCent 0x0001 
' 	Ld timeUTC 
' 	LitR8 0x0000 0x0000 0x8000 0x4096 
' 	Div 
' 	Add 
' 	ArgsLd calcTimeJulianCent 0x0001 
' 	St newt 
' Line #667:
' 	Ld newt 
' 	ArgsLd calcEquationOfTime 0x0001 
' 	St eqtime 
' Line #668:
' 	Ld newt 
' 	ArgsLd calcSunDeclination 0x0001 
' 	St solarDec 
' Line #669:
' 	Ld latitude 
' 	Ld solarDec 
' 	ArgsLd calcHourAngleSunrise 0x0002 
' 	St hourangle 
' Line #670:
' 	Ld longitude 
' 	Ld hourangle 
' 	ArgsLd radToDeg 0x0001 
' 	Sub 
' 	St delta 
' Line #671:
' 	LitDI2 0x0004 
' 	Ld delta 
' 	Mul 
' 	St timeDiff 
' Line #672:
' 	LitDI2 0x02D0 
' 	Ld timeDiff 
' 	Add 
' 	Ld eqtime 
' 	Sub 
' 	St timeUTC 
' Line #673:
' 	QuoteRem 0x0000 0x000B " in minutes"
' Line #674:
' Line #675:
' 	Ld timeUTC 
' 	St calcSunriseUTC 
' Line #676:
' Line #677:
' 	EndFunc 
' Line #678:
' Line #679:
' Line #680:
' 	FuncDefn (Function calcSolNoonUTC(t, longitude))
' Line #681:
' Line #682:
' 	QuoteRem 0x0000 0x0048 "***********************************************************************/"
' Line #683:
' 	QuoteRem 0x0000 0x0019 "* Name:    calcSolNoonUTC"
' Line #684:
' 	QuoteRem 0x0000 0x0013 "* Type:    Function"
' Line #685:
' 	QuoteRem 0x0000 0x0042 "* Purpose: calculate the Universal Coordinated Time (UTC) of solar"
' Line #686:
' 	QuoteRem 0x0000 0x003B "*     noon for the given day at the given location on earth"
' Line #687:
' 	QuoteRem 0x0000 0x000C "* Arguments:"
' Line #688:
' 	QuoteRem 0x0000 0x0030 "*   t : number of Julian centuries since J2000.0"
' Line #689:
' 	QuoteRem 0x0000 0x0030 "*   longitude : longitude of observer in degrees"
' Line #690:
' 	QuoteRem 0x0000 0x000F "* Return value:"
' Line #691:
' 	QuoteRem 0x0000 0x001F "*   time in minutes from zero Z"
' Line #692:
' 	QuoteRem 0x0000 0x0048 "***********************************************************************/"
' Line #693:
' Line #694:
' 	Dim 
' 	VarDefn newt (As Double)
' 	VarDefn eqtime (As Double)
' 	VarDefn solarNoonDec (As Double)
' 	VarDefn solNoonUTC (As Double)
' Line #695:
' Line #696:
' 	Ld t 
' 	ArgsLd calcJDFromJulianCent 0x0001 
' 	LitR8 0x0000 0x0000 0x0000 0x3FE0 
' 	Add 
' 	Ld longitude 
' 	LitR8 0x0000 0x0000 0x8000 0x4076 
' 	Div 
' 	Add 
' 	ArgsLd calcTimeJulianCent 0x0001 
' 	St newt 
' Line #697:
' Line #698:
' 	Ld newt 
' 	ArgsLd calcEquationOfTime 0x0001 
' 	St eqtime 
' Line #699:
' 	Ld newt 
' 	ArgsLd calcSunDeclination 0x0001 
' 	St solarNoonDec 
' Line #700:
' 	LitDI2 0x02D0 
' 	Ld longitude 
' 	LitDI2 0x0004 
' 	Mul 
' 	Paren 
' 	Add 
' 	Ld eqtime 
' 	Sub 
' 	St solNoonUTC 
' Line #701:
' Line #702:
' 	Ld solNoonUTC 
' 	St calcSolNoonUTC 
' Line #703:
' Line #704:
' 	EndFunc 
' Line #705:
' Line #706:
' Line #707:
' 	FuncDefn (Function calcSunsetUTC(JD, latitude, longitude))
' Line #708:
' Line #709:
' 	QuoteRem 0x0000 0x0048 "***********************************************************************/"
' Line #710:
' 	QuoteRem 0x0000 0x0018 "* Name:    calcSunsetUTC"
' Line #711:
' 	QuoteRem 0x0000 0x0013 "* Type:    Function"
' Line #712:
' 	QuoteRem 0x0000 0x0043 "* Purpose: calculate the Universal Coordinated Time (UTC) of sunset"
' Line #713:
' 	QuoteRem 0x0000 0x003A "*         for the given day at the given location on earth"
' Line #714:
' 	QuoteRem 0x0000 0x000C "* Arguments:"
' Line #715:
' 	QuoteRem 0x0000 0x0014 "*   JD  : julian day"
' Line #716:
' 	QuoteRem 0x0000 0x002E "*   latitude : latitude of observer in degrees"
' Line #717:
' 	QuoteRem 0x0000 0x0030 "*   longitude : longitude of observer in degrees"
' Line #718:
' 	QuoteRem 0x0000 0x000F "* Return value:"
' Line #719:
' 	QuoteRem 0x0000 0x001F "*   time in minutes from zero Z"
' Line #720:
' 	QuoteRem 0x0000 0x0048 "***********************************************************************/"
' Line #721:
' Line #722:
' 	Dim 
' 	VarDefn t (As Double)
' 	VarDefn eqtime (As Double)
' 	VarDefn solarDec (As Double)
' 	VarDefn hourangle (As Double)
' Line #723:
' 	Dim 
' 	VarDefn delta (As Double)
' 	VarDefn timeDiff (As Double)
' 	VarDefn timeUTC (As Double)
' Line #724:
' 	Dim 
' 	VarDefn newt (As Double)
' Line #725:
' Line #726:
' 	Ld JD 
' 	ArgsLd calcTimeJulianCent 0x0001 
' 	St t 
' Line #727:
' Line #728:
' 	QuoteRem 0x0000 0x003C "        // First calculates sunrise and approx length of day"
' Line #729:
' Line #730:
' 	Ld t 
' 	ArgsLd calcEquationOfTime 0x0001 
' 	St eqtime 
' Line #731:
' 	Ld t 
' 	ArgsLd calcSunDeclination 0x0001 
' 	St solarDec 
' Line #732:
' 	Ld latitude 
' 	Ld solarDec 
' 	ArgsLd calcHourAngleSunset 0x0002 
' 	St hourangle 
' Line #733:
' Line #734:
' 	Ld longitude 
' 	Ld hourangle 
' 	ArgsLd radToDeg 0x0001 
' 	Sub 
' 	St delta 
' Line #735:
' 	LitDI2 0x0004 
' 	Ld delta 
' 	Mul 
' 	St timeDiff 
' Line #736:
' 	LitDI2 0x02D0 
' 	Ld timeDiff 
' 	Add 
' 	Ld eqtime 
' 	Sub 
' 	St timeUTC 
' Line #737:
' Line #738:
' 	QuoteRem 0x0000 0x0042 "        // first pass used to include fractional day in gamma calc"
' Line #739:
' Line #740:
' 	Ld t 
' 	ArgsLd calcJDFromJulianCent 0x0001 
' 	Ld timeUTC 
' 	LitR8 0x0000 0x0000 0x8000 0x4096 
' 	Div 
' 	Add 
' 	ArgsLd calcTimeJulianCent 0x0001 
' 	St newt 
' Line #741:
' 	Ld newt 
' 	ArgsLd calcEquationOfTime 0x0001 
' 	St eqtime 
' Line #742:
' 	Ld newt 
' 	ArgsLd calcSunDeclination 0x0001 
' 	St solarDec 
' Line #743:
' 	Ld latitude 
' 	Ld solarDec 
' 	ArgsLd calcHourAngleSunset 0x0002 
' 	St hourangle 
' Line #744:
' Line #745:
' 	Ld longitude 
' 	Ld hourangle 
' 	ArgsLd radToDeg 0x0001 
' 	Sub 
' 	St delta 
' Line #746:
' 	LitDI2 0x0004 
' 	Ld delta 
' 	Mul 
' 	St timeDiff 
' Line #747:
' 	LitDI2 0x02D0 
' 	Ld timeDiff 
' 	Add 
' 	Ld eqtime 
' 	Sub 
' 	St timeUTC 
' Line #748:
' 	QuoteRem 0x0000 0x0015 "        // in minutes"
' Line #749:
' Line #750:
' 	Ld timeUTC 
' 	St calcSunsetUTC 
' Line #751:
' Line #752:
' 	EndFunc 
' Line #753:
' Line #754:
' Line #755:
' 	FuncDefn (Function calcDuskUTC(JD, latitude, longitude, solardepression))
' Line #756:
' Line #757:
' 	QuoteRem 0x0000 0x0048 "***********************************************************************/"
' Line #758:
' 	QuoteRem 0x0000 0x0016 "* Name:    calcDuskUTC"
' Line #759:
' 	QuoteRem 0x0000 0x0013 "* Type:    Function"
' Line #760:
' 	QuoteRem 0x0000 0x0041 "* Purpose: calculate the Universal Coordinated Time (UTC) of dusk"
' Line #761:
' 	QuoteRem 0x0000 0x003A "*         for the given day at the given location on earth"
' Line #762:
' 	QuoteRem 0x0000 0x003A "*         for user selected solar depression below horizon"
' Line #763:
' 	QuoteRem 0x0000 0x000C "* Arguments:"
' Line #764:
' 	QuoteRem 0x0000 0x0014 "*   JD  : julian day"
' Line #765:
' 	QuoteRem 0x0000 0x002E "*   latitude : latitude of observer in degrees"
' Line #766:
' 	QuoteRem 0x0000 0x0030 "*   longitude : longitude of observer in degrees"
' Line #767:
' 	QuoteRem 0x0000 0x002F "*   solardepression: angle of sun below horizon"
' Line #768:
' 	QuoteRem 0x0000 0x000F "* Return value:"
' Line #769:
' 	QuoteRem 0x0000 0x001F "*   time in minutes from zero Z"
' Line #770:
' 	QuoteRem 0x0000 0x0048 "***********************************************************************/"
' Line #771:
' Line #772:
' 	Dim 
' 	VarDefn t (As Double)
' 	VarDefn eqtime (As Double)
' 	VarDefn solarDec (As Double)
' 	VarDefn hourangle (As Double)
' Line #773:
' 	Dim 
' 	VarDefn delta (As Double)
' 	VarDefn timeDiff (As Double)
' 	VarDefn timeUTC (As Double)
' Line #774:
' 	Dim 
' 	VarDefn newt (As Double)
' Line #775:
' Line #776:
' 	Ld JD 
' 	ArgsLd calcTimeJulianCent 0x0001 
' 	St t 
' Line #777:
' Line #778:
' 	QuoteRem 0x0000 0x003C "        // First calculates sunrise and approx length of day"
' Line #779:
' Line #780:
' 	Ld t 
' 	ArgsLd calcEquationOfTime 0x0001 
' 	St eqtime 
' Line #781:
' 	Ld t 
' 	ArgsLd calcSunDeclination 0x0001 
' 	St solarDec 
' Line #782:
' 	Ld latitude 
' 	Ld solarDec 
' 	ArgsLd calcHourAngleSunset 0x0002 
' 	St hourangle 
' Line #783:
' Line #784:
' 	Ld longitude 
' 	Ld hourangle 
' 	ArgsLd radToDeg 0x0001 
' 	Sub 
' 	St delta 
' Line #785:
' 	LitDI2 0x0004 
' 	Ld delta 
' 	Mul 
' 	St timeDiff 
' Line #786:
' 	LitDI2 0x02D0 
' 	Ld timeDiff 
' 	Add 
' 	Ld eqtime 
' 	Sub 
' 	St timeUTC 
' Line #787:
' Line #788:
' 	QuoteRem 0x0000 0x0042 "        // first pass used to include fractional day in gamma calc"
' Line #789:
' Line #790:
' 	Ld t 
' 	ArgsLd calcJDFromJulianCent 0x0001 
' 	Ld timeUTC 
' 	LitR8 0x0000 0x0000 0x8000 0x4096 
' 	Div 
' 	Add 
' 	ArgsLd calcTimeJulianCent 0x0001 
' 	St newt 
' Line #791:
' 	Ld newt 
' 	ArgsLd calcEquationOfTime 0x0001 
' 	St eqtime 
' Line #792:
' 	Ld newt 
' 	ArgsLd calcSunDeclination 0x0001 
' 	St solarDec 
' Line #793:
' 	Ld latitude 
' 	Ld solarDec 
' 	Ld solardepression 
' 	ArgsLd calcHourAngleDusk 0x0003 
' 	St hourangle 
' Line #794:
' Line #795:
' 	Ld longitude 
' 	Ld hourangle 
' 	ArgsLd radToDeg 0x0001 
' 	Sub 
' 	St delta 
' Line #796:
' 	LitDI2 0x0004 
' 	Ld delta 
' 	Mul 
' 	St timeDiff 
' Line #797:
' 	LitDI2 0x02D0 
' 	Ld timeDiff 
' 	Add 
' 	Ld eqtime 
' 	Sub 
' 	St timeUTC 
' Line #798:
' 	QuoteRem 0x0000 0x0015 "        // in minutes"
' Line #799:
' Line #800:
' 	Ld timeUTC 
' 	St calcDuskUTC 
' Line #801:
' Line #802:
' 	EndFunc 
' Line #803:
' Line #804:
' Line #805:
' 	FuncDefn (Function dawn(lat, lon, year, month, day, timezone, dlstime, solardepression))
' Line #806:
' Line #807:
' 	QuoteRem 0x0000 0x0048 "***********************************************************************/"
' Line #808:
' 	QuoteRem 0x0000 0x000F "* Name:    dawn"
' Line #809:
' 	QuoteRem 0x0000 0x002E "* Type:    Main Function called by spreadsheet"
' Line #810:
' 	QuoteRem 0x0000 0x0037 "* Purpose: calculate time of dawn  for the entered date"
' Line #811:
' 	QuoteRem 0x0000 0x0013 "*     and location."
' Line #812:
' 	QuoteRem 0x0000 0x0041 "* For latitudes greater than 72 degrees N and S, calculations are"
' Line #813:
' 	QuoteRem 0x0000 0x0040 "* accurate to within 10 minutes. For latitudes less than +/- 72°"
' Line #814:
' 	QuoteRem 0x0000 0x0027 "* accuracy is approximately one minute."
' Line #815:
' 	QuoteRem 0x0000 0x000C "* Arguments:"
' Line #816:
' 	QuoteRem 0x0000 0x0028 "   latitude = latitude (decimal degrees)"
' Line #817:
' 	QuoteRem 0x0000 0x002A "   longitude = longitude (decimal degrees)"
' Line #818:
' 	QuoteRem 0x0000 0x0046 "    NOTE: longitude is negative for western hemisphere for input cells"
' Line #819:
' 	QuoteRem 0x0000 0x003D "          in the spreadsheet for calls to the functions named"
' Line #820:
' 	QuoteRem 0x0000 0x0045 "          sunrise, solarnoon, and sunset. Those functions convert the"
' Line #821:
' 	QuoteRem 0x0000 0x0047 "          longitude to positive for the western hemisphere for calls to"
' Line #822:
' 	QuoteRem 0x0000 0x003C "          other functions using the original sign convention"
' Line #823:
' 	QuoteRem 0x0000 0x0028 "          from the NOAA javascript code."
' Line #824:
' 	QuoteRem 0x0000 0x000E "   year = year"
' Line #825:
' 	QuoteRem 0x0000 0x0010 "   month = month"
' Line #826:
' 	QuoteRem 0x0000 0x000C "   day = day"
' Line #827:
' 	QuoteRem 0x0000 0x0039 "   timezone = time zone hours relative to GMT/UTC (hours)"
' Line #828:
' 	QuoteRem 0x0000 0x003C "   dlstime = daylight savings time (0 = no, 1 = yes) (hours)"
' Line #829:
' 	QuoteRem 0x0000 0x003A "   solardepression = angle of sun below horizon in degrees"
' Line #830:
' 	QuoteRem 0x0000 0x000F "* Return value:"
' Line #831:
' 	QuoteRem 0x0000 0x0022 "*   dawn time in local time (days)"
' Line #832:
' 	QuoteRem 0x0000 0x0048 "***********************************************************************/"
' Line #833:
' Line #834:
' 	Dim 
' 	VarDefn longitude (As Double)
' 	VarDefn latitude (As Double)
' 	VarDefn JD (As Double)
' Line #835:
' 	Dim 
' 	VarDefn riseTimeGMT (As Double)
' 	VarDefn riseTimeLST (As Double)
' Line #836:
' Line #837:
' 	QuoteRem 0x0000 0x0055 " change sign convention for longitude from negative to positive in western hemisphere"
' Line #838:
' 	Ld lon 
' 	LitDI2 0x0001 
' 	UMi 
' 	Mul 
' 	St longitude 
' Line #839:
' 	Ld lat 
' 	St latitude 
' Line #840:
' 	Ld latitude 
' 	LitR8 0x3333 0x3333 0x7333 0x4056 
' 	Gt 
' 	Paren 
' 	If 
' 	BoSImplicit 
' 	LitR8 0x3333 0x3333 0x7333 0x4056 
' 	St latitude 
' 	EndIf 
' Line #841:
' 	Ld latitude 
' 	LitR8 0x3333 0x3333 0x7333 0x4056 
' 	UMi 
' 	Lt 
' 	Paren 
' 	If 
' 	BoSImplicit 
' 	LitR8 0x3333 0x3333 0x7333 0x4056 
' 	UMi 
' 	St latitude 
' 	EndIf 
' Line #842:
' Line #843:
' 	Ld year 
' 	Ld month 
' 	Ld day 
' 	ArgsLd calcJD 0x0003 
' 	St JD 
' Line #844:
' Line #845:
' 	QuoteRem 0x0000 0x002E "            // Calculate sunrise for this date"
' Line #846:
' 	Ld JD 
' 	Ld latitude 
' 	Ld longitude 
' 	Ld solardepression 
' 	ArgsLd calcDawnUTC 0x0004 
' 	St riseTimeGMT 
' Line #847:
' Line #848:
' 	QuoteRem 0x0000 0x0049 "            //  adjust for time zone and daylight savings time in minutes"
' Line #849:
' 	Ld riseTimeGMT 
' 	LitDI2 0x003C 
' 	Ld timezone 
' 	Mul 
' 	Paren 
' 	Add 
' 	Ld dlstime 
' 	LitDI2 0x003C 
' 	Mul 
' 	Paren 
' 	Add 
' 	St riseTimeLST 
' Line #850:
' Line #851:
' 	QuoteRem 0x0000 0x001F "            //  convert to days"
' Line #852:
' 	Ld riseTimeLST 
' 	LitDI2 0x05A0 
' 	Div 
' 	St dawn 
' Line #853:
' Line #854:
' 	EndFunc 
' Line #855:
' Line #856:
' Line #857:
' 	FuncDefn (Function sunrise(lat, lon, year, month, day, timezone, dlstime))
' Line #858:
' Line #859:
' 	QuoteRem 0x0000 0x0048 "***********************************************************************/"
' Line #860:
' 	QuoteRem 0x0000 0x0012 "* Name:    sunrise"
' Line #861:
' 	QuoteRem 0x0000 0x002E "* Type:    Main Function called by spreadsheet"
' Line #862:
' 	QuoteRem 0x0000 0x003A "* Purpose: calculate time of sunrise  for the entered date"
' Line #863:
' 	QuoteRem 0x0000 0x0013 "*     and location."
' Line #864:
' 	QuoteRem 0x0000 0x0041 "* For latitudes greater than 72 degrees N and S, calculations are"
' Line #865:
' 	QuoteRem 0x0000 0x0040 "* accurate to within 10 minutes. For latitudes less than +/- 72°"
' Line #866:
' 	QuoteRem 0x0000 0x0027 "* accuracy is approximately one minute."
' Line #867:
' 	QuoteRem 0x0000 0x000C "* Arguments:"
' Line #868:
' 	QuoteRem 0x0000 0x0028 "   latitude = latitude (decimal degrees)"
' Line #869:
' 	QuoteRem 0x0000 0x002A "   longitude = longitude (decimal degrees)"
' Line #870:
' 	QuoteRem 0x0000 0x0046 "    NOTE: longitude is negative for western hemisphere for input cells"
' Line #871:
' 	QuoteRem 0x0000 0x003D "          in the spreadsheet for calls to the functions named"
' Line #872:
' 	QuoteRem 0x0000 0x0045 "          sunrise, solarnoon, and sunset. Those functions convert the"
' Line #873:
' 	QuoteRem 0x0000 0x0047 "          longitude to positive for the western hemisphere for calls to"
' Line #874:
' 	QuoteRem 0x0000 0x003C "          other functions using the original sign convention"
' Line #875:
' 	QuoteRem 0x0000 0x0028 "          from the NOAA javascript code."
' Line #876:
' 	QuoteRem 0x0000 0x000E "   year = year"
' Line #877:
' 	QuoteRem 0x0000 0x0010 "   month = month"
' Line #878:
' 	QuoteRem 0x0000 0x000C "   day = day"
' Line #879:
' 	QuoteRem 0x0000 0x0039 "   timezone = time zone hours relative to GMT/UTC (hours)"
' Line #880:
' 	QuoteRem 0x0000 0x003C "   dlstime = daylight savings time (0 = no, 1 = yes) (hours)"
' Line #881:
' 	QuoteRem 0x0000 0x000F "* Return value:"
' Line #882:
' 	QuoteRem 0x0000 0x0025 "*   sunrise time in local time (days)"
' Line #883:
' 	QuoteRem 0x0000 0x0048 "***********************************************************************/"
' Line #884:
' Line #885:
' 	Dim 
' 	VarDefn longitude (As Double)
' 	VarDefn latitude (As Double)
' 	VarDefn JD (As Double)
' Line #886:
' 	Dim 
' 	VarDefn riseTimeGMT (As Double)
' 	VarDefn riseTimeLST (As Double)
' Line #887:
' Line #888:
' 	QuoteRem 0x0000 0x0055 " change sign convention for longitude from negative to positive in western hemisphere"
' Line #889:
' 	Ld lon 
' 	LitDI2 0x0001 
' 	UMi 
' 	Mul 
' 	St longitude 
' Line #890:
' 	Ld lat 
' 	St latitude 
' Line #891:
' 	Ld latitude 
' 	LitR8 0x3333 0x3333 0x7333 0x4056 
' 	Gt 
' 	Paren 
' 	If 
' 	BoSImplicit 
' 	LitR8 0x3333 0x3333 0x7333 0x4056 
' 	St latitude 
' 	EndIf 
' Line #892:
' 	Ld latitude 
' 	LitR8 0x3333 0x3333 0x7333 0x4056 
' 	UMi 
' 	Lt 
' 	Paren 
' 	If 
' 	BoSImplicit 
' 	LitR8 0x3333 0x3333 0x7333 0x4056 
' 	UMi 
' 	St latitude 
' 	EndIf 
' Line #893:
' Line #894:
' 	Ld year 
' 	Ld month 
' 	Ld day 
' 	ArgsLd calcJD 0x0003 
' 	St JD 
' Line #895:
' Line #896:
' 	QuoteRem 0x0000 0x002E "            // Calculate sunrise for this date"
' Line #897:
' 	Ld JD 
' 	Ld latitude 
' 	Ld longitude 
' 	ArgsLd calcSunriseUTC 0x0003 
' 	St riseTimeGMT 
' Line #898:
' Line #899:
' 	QuoteRem 0x0000 0x0049 "            //  adjust for time zone and daylight savings time in minutes"
' Line #900:
' 	Ld riseTimeGMT 
' 	LitDI2 0x003C 
' 	Ld timezone 
' 	Mul 
' 	Paren 
' 	Add 
' 	Ld dlstime 
' 	LitDI2 0x003C 
' 	Mul 
' 	Paren 
' 	Add 
' 	St riseTimeLST 
' Line #901:
' Line #902:
' 	QuoteRem 0x0000 0x001F "            //  convert to days"
' Line #903:
' 	Ld riseTimeLST 
' 	LitDI2 0x05A0 
' 	Div 
' 	St sunrise 
' Line #904:
' Line #905:
' 	EndFunc 
' Line #906:
' Line #907:
' Line #908:
' 	FuncDefn (Function solarnoon(lat, lon, year, month, day, timezone, dlstime))
' Line #909:
' Line #910:
' 	QuoteRem 0x0000 0x0048 "***********************************************************************/"
' Line #911:
' 	QuoteRem 0x0000 0x0014 "* Name:    solarnoon"
' Line #912:
' 	QuoteRem 0x0000 0x002E "* Type:    Main Function called by spreadsheet"
' Line #913:
' 	QuoteRem 0x0000 0x0042 "* Purpose: calculate the Universal Coordinated Time (UTC) of solar"
' Line #914:
' 	QuoteRem 0x0000 0x003B "*     noon for the given day at the given location on earth"
' Line #915:
' 	QuoteRem 0x0000 0x000C "* Arguments:"
' Line #916:
' 	QuoteRem 0x0000 0x0008 "    year"
' Line #917:
' 	QuoteRem 0x0000 0x0009 "    month"
' Line #918:
' 	QuoteRem 0x0000 0x0007 "    day"
' Line #919:
' 	QuoteRem 0x0000 0x0030 "*   longitude : longitude of observer in degrees"
' Line #920:
' 	QuoteRem 0x0000 0x0046 "    NOTE: longitude is negative for western hemisphere for input cells"
' Line #921:
' 	QuoteRem 0x0000 0x003D "          in the spreadsheet for calls to the functions named"
' Line #922:
' 	QuoteRem 0x0000 0x0045 "          sunrise, solarnoon, and sunset. Those functions convert the"
' Line #923:
' 	QuoteRem 0x0000 0x0047 "          longitude to positive for the western hemisphere for calls to"
' Line #924:
' 	QuoteRem 0x0000 0x003C "          other functions using the original sign convention"
' Line #925:
' 	QuoteRem 0x0000 0x0028 "          from the NOAA javascript code."
' Line #926:
' 	QuoteRem 0x0000 0x000F "* Return value:"
' Line #927:
' 	QuoteRem 0x0000 0x0029 "*   time of solar noon in local time days"
' Line #928:
' 	QuoteRem 0x0000 0x0048 "***********************************************************************/"
' Line #929:
' Line #930:
' 	Dim 
' 	VarDefn longitude (As Double)
' 	VarDefn latitude (As Double)
' 	VarDefn JD (As Double)
' Line #931:
' 	Dim 
' 	VarDefn t (As Double)
' 	VarDefn newt (As Double)
' 	VarDefn eqtime (As Double)
' Line #932:
' 	Dim 
' 	VarDefn solarNoonDec (As Double)
' 	VarDefn solNoonUTC (As Double)
' Line #933:
' Line #934:
' 	QuoteRem 0x0000 0x0055 " change sign convention for longitude from negative to positive in western hemisphere"
' Line #935:
' 	Ld lon 
' 	LitDI2 0x0001 
' 	UMi 
' 	Mul 
' 	St longitude 
' Line #936:
' 	Ld lat 
' 	St latitude 
' Line #937:
' 	Ld latitude 
' 	LitR8 0x3333 0x3333 0x7333 0x4056 
' 	Gt 
' 	Paren 
' 	If 
' 	BoSImplicit 
' 	LitR8 0x3333 0x3333 0x7333 0x4056 
' 	St latitude 
' 	EndIf 
' Line #938:
' 	Ld latitude 
' 	LitR8 0x3333 0x3333 0x7333 0x4056 
' 	UMi 
' 	Lt 
' 	Paren 
' 	If 
' 	BoSImplicit 
' 	LitR8 0x3333 0x3333 0x7333 0x4056 
' 	UMi 
' 	St latitude 
' 	EndIf 
' Line #939:
' Line #940:
' 	Ld year 
' 	Ld month 
' 	Ld day 
' 	ArgsLd calcJD 0x0003 
' 	St JD 
' Line #941:
' 	Ld JD 
' 	ArgsLd calcTimeJulianCent 0x0001 
' 	St t 
' Line #942:
' Line #943:
' 	Ld t 
' 	ArgsLd calcJDFromJulianCent 0x0001 
' 	LitR8 0x0000 0x0000 0x0000 0x3FE0 
' 	Add 
' 	Ld longitude 
' 	LitR8 0x0000 0x0000 0x8000 0x4076 
' 	Div 
' 	Add 
' 	ArgsLd calcTimeJulianCent 0x0001 
' 	St newt 
' Line #944:
' Line #945:
' 	Ld newt 
' 	ArgsLd calcEquationOfTime 0x0001 
' 	St eqtime 
' Line #946:
' 	Ld newt 
' 	ArgsLd calcSunDeclination 0x0001 
' 	St solarNoonDec 
' Line #947:
' 	LitDI2 0x02D0 
' 	Ld longitude 
' 	LitDI2 0x0004 
' 	Mul 
' 	Paren 
' 	Add 
' 	Ld eqtime 
' 	Sub 
' 	St solNoonUTC 
' Line #948:
' Line #949:
' 	QuoteRem 0x0000 0x0049 "            //  adjust for time zone and daylight savings time in minutes"
' Line #950:
' 	Ld solNoonUTC 
' 	LitDI2 0x003C 
' 	Ld timezone 
' 	Mul 
' 	Paren 
' 	Add 
' 	Ld dlstime 
' 	LitDI2 0x003C 
' 	Mul 
' 	Paren 
' 	Add 
' 	St solarnoon 
' Line #951:
' Line #952:
' 	QuoteRem 0x0000 0x001F "            //  convert to days"
' Line #953:
' 	Ld solarnoon 
' 	LitDI2 0x05A0 
' 	Div 
' 	St solarnoon 
' Line #954:
' Line #955:
' 	EndFunc 
' Line #956:
' Line #957:
' Line #958:
' 	FuncDefn (Function sunset(lat, lon, year, month, day, timezone, dlstime))
' Line #959:
' Line #960:
' 	QuoteRem 0x0000 0x0048 "***********************************************************************/"
' Line #961:
' 	QuoteRem 0x0000 0x0011 "* Name:    sunset"
' Line #962:
' 	QuoteRem 0x0000 0x002E "* Type:    Main Function called by spreadsheet"
' Line #963:
' 	QuoteRem 0x0000 0x0044 "* Purpose: calculate time of sunrise and sunset for the entered date"
' Line #964:
' 	QuoteRem 0x0000 0x0013 "*     and location."
' Line #965:
' 	QuoteRem 0x0000 0x0041 "* For latitudes greater than 72 degrees N and S, calculations are"
' Line #966:
' 	QuoteRem 0x0000 0x0040 "* accurate to within 10 minutes. For latitudes less than +/- 72°"
' Line #967:
' 	QuoteRem 0x0000 0x0027 "* accuracy is approximately one minute."
' Line #968:
' 	QuoteRem 0x0000 0x000C "* Arguments:"
' Line #969:
' 	QuoteRem 0x0000 0x0028 "   latitude = latitude (decimal degrees)"
' Line #970:
' 	QuoteRem 0x0000 0x002A "   longitude = longitude (decimal degrees)"
' Line #971:
' 	QuoteRem 0x0000 0x0046 "    NOTE: longitude is negative for western hemisphere for input cells"
' Line #972:
' 	QuoteRem 0x0000 0x003D "          in the spreadsheet for calls to the functions named"
' Line #973:
' 	QuoteRem 0x0000 0x0045 "          sunrise, solarnoon, and sunset. Those functions convert the"
' Line #974:
' 	QuoteRem 0x0000 0x0047 "          longitude to positive for the western hemisphere for calls to"
' Line #975:
' 	QuoteRem 0x0000 0x003C "          other functions using the original sign convention"
' Line #976:
' 	QuoteRem 0x0000 0x0028 "          from the NOAA javascript code."
' Line #977:
' 	QuoteRem 0x0000 0x000E "   year = year"
' Line #978:
' 	QuoteRem 0x0000 0x0010 "   month = month"
' Line #979:
' 	QuoteRem 0x0000 0x000C "   day = day"
' Line #980:
' 	QuoteRem 0x0000 0x0039 "   timezone = time zone hours relative to GMT/UTC (hours)"
' Line #981:
' 	QuoteRem 0x0000 0x003C "   dlstime = daylight savings time (0 = no, 1 = yes) (hours)"
' Line #982:
' 	QuoteRem 0x0000 0x000F "* Return value:"
' Line #983:
' 	QuoteRem 0x0000 0x0024 "*   sunset time in local time (days)"
' Line #984:
' 	QuoteRem 0x0000 0x0048 "***********************************************************************/"
' Line #985:
' Line #986:
' 	Dim 
' 	VarDefn longitude (As Double)
' 	VarDefn latitude (As Double)
' 	VarDefn JD (As Double)
' Line #987:
' 	Dim 
' 	VarDefn setTimeGMT (As Double)
' 	VarDefn setTimeLST (As Double)
' Line #988:
' Line #989:
' 	QuoteRem 0x0000 0x0055 " change sign convention for longitude from negative to positive in western hemisphere"
' Line #990:
' 	Ld lon 
' 	LitDI2 0x0001 
' 	UMi 
' 	Mul 
' 	St longitude 
' Line #991:
' 	Ld lat 
' 	St latitude 
' Line #992:
' 	Ld latitude 
' 	LitR8 0x3333 0x3333 0x7333 0x4056 
' 	Gt 
' 	Paren 
' 	If 
' 	BoSImplicit 
' 	LitR8 0x3333 0x3333 0x7333 0x4056 
' 	St latitude 
' 	EndIf 
' Line #993:
' 	Ld latitude 
' 	LitR8 0x3333 0x3333 0x7333 0x4056 
' 	UMi 
' 	Lt 
' 	Paren 
' 	If 
' 	BoSImplicit 
' 	LitR8 0x3333 0x3333 0x7333 0x4056 
' 	UMi 
' 	St latitude 
' 	EndIf 
' Line #994:
' Line #995:
' 	Ld year 
' 	Ld month 
' 	Ld day 
' 	ArgsLd calcJD 0x0003 
' 	St JD 
' Line #996:
' Line #997:
' 	QuoteRem 0x0000 0x002C "           // Calculate sunset for this date"
' Line #998:
' 	Ld JD 
' 	Ld latitude 
' 	Ld longitude 
' 	ArgsLd calcSunsetUTC 0x0003 
' 	St setTimeGMT 
' Line #999:
' Line #1000:
' 	QuoteRem 0x0000 0x0049 "            //  adjust for time zone and daylight savings time in minutes"
' Line #1001:
' 	Ld setTimeGMT 
' 	LitDI2 0x003C 
' 	Ld timezone 
' 	Mul 
' 	Paren 
' 	Add 
' 	Ld dlstime 
' 	LitDI2 0x003C 
' 	Mul 
' 	Paren 
' 	Add 
' 	St setTimeLST 
' Line #1002:
' Line #1003:
' 	QuoteRem 0x0000 0x001F "            //  convert to days"
' Line #1004:
' 	Ld setTimeLST 
' 	LitDI2 0x05A0 
' 	Div 
' 	St sunset 
' Line #1005:
' Line #1006:
' 	EndFunc 
' Line #1007:
' Line #1008:
' Line #1009:
' 	FuncDefn (Function dusk(lat, lon, year, month, day, timezone, dlstime, solardepression))
' Line #1010:
' Line #1011:
' 	QuoteRem 0x0000 0x0048 "***********************************************************************/"
' Line #1012:
' 	QuoteRem 0x0000 0x000F "* Name:    dusk"
' Line #1013:
' 	QuoteRem 0x0000 0x002E "* Type:    Main Function called by spreadsheet"
' Line #1014:
' 	QuoteRem 0x0000 0x0044 "* Purpose: calculate time of sunrise and sunset for the entered date"
' Line #1015:
' 	QuoteRem 0x0000 0x0013 "*     and location."
' Line #1016:
' 	QuoteRem 0x0000 0x0041 "* For latitudes greater than 72 degrees N and S, calculations are"
' Line #1017:
' 	QuoteRem 0x0000 0x0040 "* accurate to within 10 minutes. For latitudes less than +/- 72°"
' Line #1018:
' 	QuoteRem 0x0000 0x0027 "* accuracy is approximately one minute."
' Line #1019:
' 	QuoteRem 0x0000 0x000C "* Arguments:"
' Line #1020:
' 	QuoteRem 0x0000 0x0028 "   latitude = latitude (decimal degrees)"
' Line #1021:
' 	QuoteRem 0x0000 0x002A "   longitude = longitude (decimal degrees)"
' Line #1022:
' 	QuoteRem 0x0000 0x0046 "    NOTE: longitude is negative for western hemisphere for input cells"
' Line #1023:
' 	QuoteRem 0x0000 0x003D "          in the spreadsheet for calls to the functions named"
' Line #1024:
' 	QuoteRem 0x0000 0x0045 "          sunrise, solarnoon, and sunset. Those functions convert the"
' Line #1025:
' 	QuoteRem 0x0000 0x0047 "          longitude to positive for the western hemisphere for calls to"
' Line #1026:
' 	QuoteRem 0x0000 0x003C "          other functions using the original sign convention"
' Line #1027:
' 	QuoteRem 0x0000 0x0028 "          from the NOAA javascript code."
' Line #1028:
' 	QuoteRem 0x0000 0x000E "   year = year"
' Line #1029:
' 	QuoteRem 0x0000 0x0010 "   month = month"
' Line #1030:
' 	QuoteRem 0x0000 0x000C "   day = day"
' Line #1031:
' 	QuoteRem 0x0000 0x0039 "   timezone = time zone hours relative to GMT/UTC (hours)"
' Line #1032:
' 	QuoteRem 0x0000 0x003C "   dlstime = daylight savings time (0 = no, 1 = yes) (hours)"
' Line #1033:
' 	QuoteRem 0x0000 0x003A "   solardepression = angle of sun below horizon in degrees"
' Line #1034:
' 	QuoteRem 0x0000 0x000F "* Return value:"
' Line #1035:
' 	QuoteRem 0x0000 0x0022 "*   dusk time in local time (days)"
' Line #1036:
' 	QuoteRem 0x0000 0x0048 "***********************************************************************/"
' Line #1037:
' Line #1038:
' 	Dim 
' 	VarDefn longitude (As Double)
' 	VarDefn latitude (As Double)
' 	VarDefn JD (As Double)
' Line #1039:
' 	Dim 
' 	VarDefn setTimeGMT (As Double)
' 	VarDefn setTimeLST (As Double)
' Line #1040:
' Line #1041:
' 	QuoteRem 0x0000 0x0055 " change sign convention for longitude from negative to positive in western hemisphere"
' Line #1042:
' 	Ld lon 
' 	LitDI2 0x0001 
' 	UMi 
' 	Mul 
' 	St longitude 
' Line #1043:
' 	Ld lat 
' 	St latitude 
' Line #1044:
' 	Ld latitude 
' 	LitR8 0x3333 0x3333 0x7333 0x4056 
' 	Gt 
' 	Paren 
' 	If 
' 	BoSImplicit 
' 	LitR8 0x3333 0x3333 0x7333 0x4056 
' 	St latitude 
' 	EndIf 
' Line #1045:
' 	Ld latitude 
' 	LitR8 0x3333 0x3333 0x7333 0x4056 
' 	UMi 
' 	Lt 
' 	Paren 
' 	If 
' 	BoSImplicit 
' 	LitR8 0x3333 0x3333 0x7333 0x4056 
' 	UMi 
' 	St latitude 
' 	EndIf 
' Line #1046:
' Line #1047:
' 	Ld year 
' 	Ld month 
' 	Ld day 
' 	ArgsLd calcJD 0x0003 
' 	St JD 
' Line #1048:
' Line #1049:
' 	QuoteRem 0x0000 0x002C "           // Calculate sunset for this date"
' Line #1050:
' 	Ld JD 
' 	Ld latitude 
' 	Ld longitude 
' 	Ld solardepression 
' 	ArgsLd calcDuskUTC 0x0004 
' 	St setTimeGMT 
' Line #1051:
' Line #1052:
' 	QuoteRem 0x0000 0x0049 "            //  adjust for time zone and daylight savings time in minutes"
' Line #1053:
' 	Ld setTimeGMT 
' 	LitDI2 0x003C 
' 	Ld timezone 
' 	Mul 
' 	Paren 
' 	Add 
' 	Ld dlstime 
' 	LitDI2 0x003C 
' 	Mul 
' 	Paren 
' 	Add 
' 	St setTimeLST 
' Line #1054:
' Line #1055:
' 	QuoteRem 0x0000 0x001F "            //  convert to days"
' Line #1056:
' 	Ld setTimeLST 
' 	LitDI2 0x05A0 
' 	Div 
' 	St dusk 
' Line #1057:
' Line #1058:
' 	EndFunc 
' Line #1059:
' Line #1060:
' Line #1061:
' 	LineCont 0x0004 0D 00 16 00
' 	FuncDefn (Function solarazimuth(lat, lon, year, month, day, hours, minutes, seconds, timezone, dlstime))
' Line #1062:
' Line #1063:
' 	QuoteRem 0x0000 0x0048 "***********************************************************************/"
' Line #1064:
' 	QuoteRem 0x0000 0x0017 "* Name:    solarazimuth"
' Line #1065:
' 	QuoteRem 0x0000 0x0018 "* Type:    Main Function"
' Line #1066:
' 	QuoteRem 0x0000 0x0043 "* Purpose: calculate solar azimuth (deg from north) for the entered"
' Line #1067:
' 	QuoteRem 0x0000 0x004B "*          date, time and location. Returns -999999 if darker than twilight"
' Line #1068:
' 	QuoteRem 0x0000 0x0001 "*"
' Line #1069:
' 	QuoteRem 0x0000 0x000C "* Arguments:"
' Line #1070:
' 	QuoteRem 0x0000 0x0040 "*   latitude, longitude, year, month, day, hour, minute, second,"
' Line #1071:
' 	QuoteRem 0x0000 0x0021 "*   timezone, daylightsavingstime"
' Line #1072:
' 	QuoteRem 0x0000 0x000F "* Return value:"
' Line #1073:
' 	QuoteRem 0x0000 0x0027 "*   solar azimuth in degrees from north"
' Line #1074:
' 	QuoteRem 0x0000 0x0001 "*"
' Line #1075:
' 	QuoteRem 0x0000 0x003F "* Note: solarelevation and solarazimuth functions are identical"
' Line #1076:
' 	QuoteRem 0x0000 0x0044 "*       and could be converted to a VBA subroutine that would return"
' Line #1077:
' 	QuoteRem 0x0000 0x0014 "*       both values."
' Line #1078:
' 	QuoteRem 0x0000 0x0001 "*"
' Line #1079:
' 	QuoteRem 0x0000 0x0048 "***********************************************************************/"
' Line #1080:
' Line #1081:
' 	Dim 
' 	VarDefn longitude (As Double)
' 	VarDefn latitude (As Double)
' Line #1082:
' 	Dim 
' 	VarDefn zone (As Double)
' 	VarDefn zone (As Double)
' Line #1083:
' 	Dim 
' 	VarDefn daySavings (As Double)
' 	VarDefn hh (As Double)
' 	VarDefn mm (As Double)
' 	VarDefn ss (As Double)
' Line #1084:
' 	Dim 
' 	VarDefn JD (As Double)
' 	VarDefn t (As Double)
' 	VarDefn R (As Double)
' Line #1085:
' 	Dim 
' 	VarDefn alpha (As Double)
' 	VarDefn theta (As Double)
' 	VarDefn Etime (As Double)
' 	VarDefn eqtime (As Double)
' Line #1086:
' 	Dim 
' 	VarDefn solarDec (As Double)
' 	VarDefn timenow (As Double)
' 	VarDefn earthRadVec (As Double)
' Line #1087:
' 	Dim 
' 	VarDefn solarTimeFix (As Double)
' 	VarDefn hourangle (As Double)
' 	VarDefn trueSolarTime (As Double)
' Line #1088:
' 	Dim 
' 	VarDefn harad (As Double)
' 	VarDefn csz (As Double)
' 	VarDefn zenith (As Double)
' 	VarDefn azDenom (As Double)
' Line #1089:
' 	Dim 
' 	VarDefn azRad (As Double)
' 	VarDefn azimuth (As Double)
' Line #1090:
' 	Dim 
' 	VarDefn exoatmElevation (As Double)
' 	VarDefn step1 (As Double)
' 	VarDefn step2 (As Double)
' Line #1091:
' 	Dim 
' 	VarDefn step3 (As Double)
' 	VarDefn refractionCorrection (As Double)
' 	VarDefn te (As Double)
' Line #1092:
' Line #1093:
' 	QuoteRem 0x0000 0x0055 " change sign convention for longitude from negative to positive in western hemisphere"
' Line #1094:
' 	Ld lon 
' 	LitDI2 0x0001 
' 	UMi 
' 	Mul 
' 	St longitude 
' Line #1095:
' 	Ld lat 
' 	St latitude 
' Line #1096:
' 	Ld latitude 
' 	LitR8 0x3333 0x3333 0x7333 0x4056 
' 	Gt 
' 	Paren 
' 	If 
' 	BoSImplicit 
' 	LitR8 0x3333 0x3333 0x7333 0x4056 
' 	St latitude 
' 	EndIf 
' Line #1097:
' 	Ld latitude 
' 	LitR8 0x3333 0x3333 0x7333 0x4056 
' 	UMi 
' 	Lt 
' 	Paren 
' 	If 
' 	BoSImplicit 
' 	LitR8 0x3333 0x3333 0x7333 0x4056 
' 	UMi 
' 	St latitude 
' 	EndIf 
' Line #1098:
' Line #1099:
' 	QuoteRem 0x0000 0x0039 "change time zone to ppositive hours in western hemisphere"
' Line #1100:
' 	Ld timezone 
' 	LitDI2 0x0001 
' 	UMi 
' 	Mul 
' 	St zone 
' Line #1101:
' 	Ld dlstime 
' 	LitDI2 0x003C 
' 	Mul 
' 	St zone 
' Line #1102:
' 	Ld hours 
' 	Ld zone 
' 	LitDI2 0x003C 
' 	Div 
' 	Paren 
' 	Sub 
' 	St daySavings 
' Line #1103:
' 	Ld minutes 
' 	St hh 
' Line #1104:
' 	Ld seconds 
' 	St mm 
' Line #1105:
' Line #1106:
' 	QuoteRem 0x0000 0x003B "//    timenow is GMT time for calculation in hours since 0Z"
' Line #1107:
' 	Ld daySavings 
' 	Ld hh 
' 	LitDI2 0x003C 
' 	Div 
' 	Add 
' 	Ld mm 
' 	LitDI2 0x0E10 
' 	Div 
' 	Add 
' 	Ld zone 
' 	Add 
' 	St ss 
' Line #1108:
' Line #1109:
' 	Ld year 
' 	Ld month 
' 	Ld day 
' 	ArgsLd calcJD 0x0003 
' 	St JD 
' Line #1110:
' 	Ld JD 
' 	Ld ss 
' 	LitR8 0x0000 0x0000 0x0000 0x4038 
' 	Div 
' 	Add 
' 	ArgsLd calcTimeJulianCent 0x0001 
' 	St t 
' Line #1111:
' 	Ld t 
' 	ArgsLd calcSunRadVector 0x0001 
' 	St R 
' Line #1112:
' 	Ld t 
' 	ArgsLd calcSunRtAscension 0x0001 
' 	St alpha 
' Line #1113:
' 	Ld t 
' 	ArgsLd calcSunDeclination 0x0001 
' 	St theta 
' Line #1114:
' 	Ld t 
' 	ArgsLd calcEquationOfTime 0x0001 
' 	St Etime 
' Line #1115:
' Line #1116:
' 	Ld Etime 
' 	St eqtime 
' Line #1117:
' 	Ld theta 
' 	St solarDec 
' 	QuoteRem 0x001D 0x0010 "//    in degrees"
' Line #1118:
' 	Ld R 
' 	St timenow 
' Line #1119:
' Line #1120:
' 	Ld eqtime 
' 	LitR8 0x0000 0x0000 0x0000 0x4010 
' 	Ld longitude 
' 	Mul 
' 	Sub 
' 	LitR8 0x0000 0x0000 0x0000 0x404E 
' 	Ld zone 
' 	Mul 
' 	Add 
' 	St earthRadVec 
' Line #1121:
' 	Ld daySavings 
' 	LitR8 0x0000 0x0000 0x0000 0x404E 
' 	Mul 
' 	Ld hh 
' 	Add 
' 	Ld mm 
' 	LitR8 0x0000 0x0000 0x0000 0x404E 
' 	Div 
' 	Add 
' 	Ld earthRadVec 
' 	Add 
' 	St solarTimeFix 
' Line #1122:
' 	QuoteRem 0x000C 0x0010 "//    in minutes"
' Line #1123:
' Line #1124:
' 	Ld solarTimeFix 
' 	LitDI2 0x05A0 
' 	Gt 
' 	Paren 
' 	DoWhile 
' Line #1125:
' 	Ld solarTimeFix 
' 	LitDI2 0x05A0 
' 	Sub 
' 	St solarTimeFix 
' Line #1126:
' 	Loop 
' Line #1127:
' Line #1128:
' 	Ld solarTimeFix 
' 	LitR8 0x0000 0x0000 0x0000 0x4010 
' 	Div 
' 	LitR8 0x0000 0x0000 0x8000 0x4066 
' 	Sub 
' 	St hourangle 
' Line #1129:
' 	QuoteRem 0x000C 0x0034 "//    Thanks to Louis Schwarzmayr for the next line:"
' Line #1130:
' 	Ld hourangle 
' 	LitDI2 0x00B4 
' 	UMi 
' 	Lt 
' 	Paren 
' 	If 
' 	BoSImplicit 
' 	Ld hourangle 
' 	LitR8 0x0000 0x0000 0x8000 0x4076 
' 	Add 
' 	St hourangle 
' 	EndIf 
' Line #1131:
' Line #1132:
' 	Ld hourangle 
' 	ArgsLd degToRad 0x0001 
' 	St trueSolarTime 
' Line #1133:
' Line #1134:
' 	LineCont 0x000C 0A 00 12 00 12 00 12 00 1A 00 12 00
' 	Ld latitude 
' 	ArgsLd degToRad 0x0001 
' 	ArgsLd Sin 0x0001 
' 	Ld solarDec 
' 	ArgsLd degToRad 0x0001 
' 	ArgsLd Sin 0x0001 
' 	Mul 
' 	Ld latitude 
' 	ArgsLd degToRad 0x0001 
' 	ArgsLd Cos 0x0001 
' 	Ld solarDec 
' 	ArgsLd degToRad 0x0001 
' 	ArgsLd Cos 0x0001 
' 	Mul 
' 	Ld trueSolarTime 
' 	ArgsLd Cos 0x0001 
' 	Mul 
' 	Add 
' 	St harad 
' Line #1135:
' Line #1136:
' 	Ld harad 
' 	LitR8 0x0000 0x0000 0x0000 0x3FF0 
' 	Gt 
' 	Paren 
' 	IfBlock 
' Line #1137:
' 	LitR8 0x0000 0x0000 0x0000 0x3FF0 
' 	St harad 
' Line #1138:
' 	Ld harad 
' 	LitR8 0x0000 0x0000 0x0000 0x3FF0 
' 	UMi 
' 	Lt 
' 	Paren 
' 	ElseIfBlock 
' Line #1139:
' 	LitR8 0x0000 0x0000 0x0000 0x3FF0 
' 	UMi 
' 	St harad 
' Line #1140:
' 	EndIfBlock 
' Line #1141:
' Line #1142:
' 	Ld harad 
' 	Ld Application 
' 	MemLd WorksheetFunction 
' 	ArgsMemLd Acos 0x0001 
' 	ArgsLd radToDeg 0x0001 
' 	St csz 
' Line #1143:
' Line #1144:
' 	Ld latitude 
' 	ArgsLd degToRad 0x0001 
' 	ArgsLd Cos 0x0001 
' 	Ld csz 
' 	ArgsLd degToRad 0x0001 
' 	ArgsLd Sin 0x0001 
' 	Mul 
' 	Paren 
' 	St zenith 
' Line #1145:
' Line #1146:
' 	Ld zenith 
' 	FnAbs 
' 	LitR8 0xA9FC 0xD2F1 0x624D 0x3F50 
' 	Gt 
' 	Paren 
' 	IfBlock 
' Line #1147:
' 	LineCont 0x0008 0C 00 14 00 15 00 14 00
' 	Ld latitude 
' 	ArgsLd degToRad 0x0001 
' 	ArgsLd Sin 0x0001 
' 	Ld csz 
' 	ArgsLd degToRad 0x0001 
' 	ArgsLd Cos 0x0001 
' 	Mul 
' 	Paren 
' 	Ld solarDec 
' 	ArgsLd degToRad 0x0001 
' 	ArgsLd Sin 0x0001 
' 	Sub 
' 	Paren 
' 	Ld zenith 
' 	Div 
' 	St azDenom 
' Line #1148:
' 	Ld azDenom 
' 	FnAbs 
' 	LitR8 0x0000 0x0000 0x0000 0x3FF0 
' 	Gt 
' 	Paren 
' 	IfBlock 
' Line #1149:
' 	Ld azDenom 
' 	LitDI2 0x0000 
' 	Lt 
' 	Paren 
' 	IfBlock 
' Line #1150:
' 	LitR8 0x0000 0x0000 0x0000 0x3FF0 
' 	UMi 
' 	St azDenom 
' Line #1151:
' 	ElseBlock 
' Line #1152:
' 	LitR8 0x0000 0x0000 0x0000 0x3FF0 
' 	St azDenom 
' Line #1153:
' 	EndIfBlock 
' Line #1154:
' 	EndIfBlock 
' Line #1155:
' Line #1156:
' 	LitR8 0x0000 0x0000 0x8000 0x4066 
' 	Ld azDenom 
' 	Ld Application 
' 	MemLd WorksheetFunction 
' 	ArgsMemLd Acos 0x0001 
' 	ArgsLd radToDeg 0x0001 
' 	Sub 
' 	St azRad 
' Line #1157:
' Line #1158:
' 	Ld hourangle 
' 	LitR8 0x0000 0x0000 0x0000 0x0000 
' 	Gt 
' 	Paren 
' 	IfBlock 
' Line #1159:
' 	Ld azRad 
' 	UMi 
' 	St azRad 
' Line #1160:
' 	EndIfBlock 
' Line #1161:
' 	ElseBlock 
' Line #1162:
' 	Ld latitude 
' 	LitR8 0x0000 0x0000 0x0000 0x0000 
' 	Gt 
' 	Paren 
' 	IfBlock 
' Line #1163:
' 	LitR8 0x0000 0x0000 0x8000 0x4066 
' 	St azRad 
' Line #1164:
' 	ElseBlock 
' Line #1165:
' 	LitR8 0x0000 0x0000 0x0000 0x0000 
' 	St azRad 
' Line #1166:
' 	EndIfBlock 
' Line #1167:
' 	EndIfBlock 
' Line #1168:
' 	Ld azRad 
' 	LitR8 0x0000 0x0000 0x0000 0x0000 
' 	Lt 
' 	Paren 
' 	IfBlock 
' Line #1169:
' 	Ld azRad 
' 	LitR8 0x0000 0x0000 0x8000 0x4076 
' 	Add 
' 	St azRad 
' Line #1170:
' 	EndIfBlock 
' Line #1171:
' Line #1172:
' 	LitR8 0x0000 0x0000 0x8000 0x4056 
' 	Ld csz 
' 	Sub 
' 	St azimuth 
' Line #1173:
' Line #1174:
' 	QuoteRem 0x0000 0x002D "beginning of complex expression commented out"
' Line #1175:
' 	QuoteRem 0x0000 0x002B "            If (exoatmElevation > 85#) Then"
' Line #1176:
' 	QuoteRem 0x0000 0x0029 "                refractionCorrection = 0#"
' Line #1177:
' 	QuoteRem 0x0000 0x0010 "            Else"
' Line #1178:
' 	QuoteRem 0x0000 0x0033 "                te = Tan(degToRad(exoatmElevation))"
' Line #1179:
' 	QuoteRem 0x0000 0x002E "                If (exoatmElevation > 5#) Then"
' Line #1180:
' 	LineCont 0x0004 01 00 B2 FF
' 	QuoteRem 0x0000 0x008A "                    refractionCorrection = 58.1 / te - 0.07 / (te * te * te) +'                        0.000086 / (te * te * te * te * te)"
' Line #1181:
' 	QuoteRem 0x0000 0x0036 "                ElseIf (exoatmElevation > -0.575) Then"
' Line #1182:
' 	LineCont 0x000C 01 00 BC FF 01 00 7F FF 01 00 4B FF
' 	QuoteRem 0x0000 0x00E8 "                    refractionCorrection = 1735# + exoatmElevation *'                        (-518.2 + exoatmElevation * (103.4 +'                        exoatmElevation * (-12.79 +'                        exoatmElevation * 0.711)))"
' Line #1183:
' 	QuoteRem 0x0000 0x0014 "                Else"
' Line #1184:
' 	QuoteRem 0x0000 0x0037 "                    refractionCorrection = -20.774 / te"
' Line #1185:
' 	QuoteRem 0x0000 0x0016 "                End If"
' Line #1186:
' 	QuoteRem 0x0000 0x0043 "                refractionCorrection = refractionCorrection / 3600#"
' Line #1187:
' 	QuoteRem 0x0000 0x0012 "            End If"
' Line #1188:
' 	QuoteRem 0x0000 0x0019 "end of complex expression"
' Line #1189:
' Line #1190:
' 	QuoteRem 0x0000 0x0022 "beginning of simplified expression"
' Line #1191:
' 	Ld azimuth 
' 	LitR8 0x0000 0x0000 0x4000 0x4055 
' 	Gt 
' 	Paren 
' 	IfBlock 
' Line #1192:
' 	LitR8 0x0000 0x0000 0x0000 0x0000 
' 	St step3 
' Line #1193:
' 	ElseBlock 
' Line #1194:
' 	Ld azimuth 
' 	ArgsLd degToRad 0x0001 
' 	ArgsLd Tan 0x0001 
' 	St refractionCorrection 
' Line #1195:
' 	Ld azimuth 
' 	LitR8 0x0000 0x0000 0x0000 0x4014 
' 	Gt 
' 	Paren 
' 	IfBlock 
' Line #1196:
' 	LineCont 0x0004 10 00 18 00
' 	LitR8 0xCCCD 0xCCCC 0x0CCC 0x404D 
' 	Ld refractionCorrection 
' 	Div 
' 	LitR8 0x51EC 0x1EB8 0xEB85 0x3FB1 
' 	Ld refractionCorrection 
' 	Ld refractionCorrection 
' 	Mul 
' 	Ld refractionCorrection 
' 	Mul 
' 	Paren 
' 	Div 
' 	Sub 
' 	LitR8 0x7736 0xBFF4 0x8B5C 0x3F16 
' 	Ld refractionCorrection 
' 	Ld refractionCorrection 
' 	Mul 
' 	Ld refractionCorrection 
' 	Mul 
' 	Ld refractionCorrection 
' 	Mul 
' 	Ld refractionCorrection 
' 	Mul 
' 	Paren 
' 	Div 
' 	Add 
' 	St step3 
' Line #1197:
' 	Ld azimuth 
' 	LitR8 0x6666 0x6666 0x6666 0x3FE2 
' 	UMi 
' 	Gt 
' 	Paren 
' 	ElseIfBlock 
' Line #1198:
' 	LitR8 0xAE14 0xE147 0x947A 0x4029 
' 	UMi 
' 	Ld azimuth 
' 	LitR8 0x978D 0x126E 0xC083 0x3FE6 
' 	Mul 
' 	Add 
' 	Paren 
' 	St exoatmElevation 
' Line #1199:
' 	LitR8 0x999A 0x9999 0xD999 0x4059 
' 	Ld azimuth 
' 	Ld exoatmElevation 
' 	Paren 
' 	Mul 
' 	Add 
' 	Paren 
' 	St step1 
' Line #1200:
' 	LitR8 0x999A 0x9999 0x3199 0x4080 
' 	UMi 
' 	Ld azimuth 
' 	Ld step1 
' 	Paren 
' 	Mul 
' 	Add 
' 	Paren 
' 	St step2 
' Line #1201:
' 	LitR8 0x0000 0x0000 0x1C00 0x409B 
' 	Ld azimuth 
' 	Ld step2 
' 	Paren 
' 	Mul 
' 	Add 
' 	St step3 
' Line #1202:
' 	ElseBlock 
' Line #1203:
' 	LitR8 0x1AA0 0xDD2F 0xC624 0x4034 
' 	UMi 
' 	Ld refractionCorrection 
' 	Div 
' 	St step3 
' Line #1204:
' 	EndIfBlock 
' Line #1205:
' 	Ld step3 
' 	LitR8 0x0000 0x0000 0x2000 0x40AC 
' 	Div 
' 	St step3 
' Line #1206:
' 	EndIfBlock 
' Line #1207:
' 	QuoteRem 0x0000 0x001C "end of simplified expression"
' Line #1208:
' Line #1209:
' 	Ld csz 
' 	Ld step3 
' 	Sub 
' 	St te 
' Line #1210:
' Line #1211:
' 	QuoteRem 0x0000 0x0025 "            If (solarZen < 108#) Then"
' Line #1212:
' 	Ld azRad 
' 	St solarazimuth 
' Line #1213:
' 	QuoteRem 0x0000 0x002D "              solarelevation = 90# - solarZen"
' Line #1214:
' 	QuoteRem 0x0000 0x0026 "              If (solarZen < 90#) Then"
' Line #1215:
' 	QuoteRem 0x0000 0x0030 "                coszen = Cos(degToRad(solarZen))"
' Line #1216:
' 	QuoteRem 0x0000 0x0012 "              Else"
' Line #1217:
' 	QuoteRem 0x0000 0x001B "                coszen = 0#"
' Line #1218:
' 	QuoteRem 0x0000 0x0014 "              End If"
' Line #1219:
' 	QuoteRem 0x0000 0x0042 "            Else    '// do not report az & el after astro twilight"
' Line #1220:
' 	QuoteRem 0x0000 0x0024 "              solarazimuth = -999999"
' Line #1221:
' 	QuoteRem 0x0000 0x0026 "              solarelevation = -999999"
' Line #1222:
' 	QuoteRem 0x0000 0x001E "              coszen = -999999"
' Line #1223:
' 	QuoteRem 0x0000 0x0012 "            End If"
' Line #1224:
' Line #1225:
' 	EndFunc 
' Line #1226:
' Line #1227:
' Line #1228:
' 	LineCont 0x0004 0D 00 16 00
' 	FuncDefn (Function solarzen(lat, lon, year, month, day, hours, minutes, seconds, timezone, dlstime))
' Line #1229:
' Line #1230:
' 	QuoteRem 0x0000 0x0048 "***********************************************************************/"
' Line #1231:
' 	QuoteRem 0x0000 0x0017 "* Name:    solarazimuth"
' Line #1232:
' 	QuoteRem 0x0000 0x0018 "* Type:    Main Function"
' Line #1233:
' 	QuoteRem 0x0000 0x0043 "* Purpose: calculate solar azimuth (deg from north) for the entered"
' Line #1234:
' 	QuoteRem 0x0000 0x004B "*          date, time and location. Returns -999999 if darker than twilight"
' Line #1235:
' 	QuoteRem 0x0000 0x0001 "*"
' Line #1236:
' 	QuoteRem 0x0000 0x000C "* Arguments:"
' Line #1237:
' 	QuoteRem 0x0000 0x0040 "*   latitude, longitude, year, month, day, hour, minute, second,"
' Line #1238:
' 	QuoteRem 0x0000 0x0021 "*   timezone, daylightsavingstime"
' Line #1239:
' 	QuoteRem 0x0000 0x000F "* Return value:"
' Line #1240:
' 	QuoteRem 0x0000 0x0027 "*   solar azimuth in degrees from north"
' Line #1241:
' 	QuoteRem 0x0000 0x0001 "*"
' Line #1242:
' 	QuoteRem 0x0000 0x003F "* Note: solarelevation and solarazimuth functions are identical"
' Line #1243:
' 	QuoteRem 0x0000 0x0041 "*       and could converted to a VBA subroutine that would return"
' Line #1244:
' 	QuoteRem 0x0000 0x0014 "*       both values."
' Line #1245:
' 	QuoteRem 0x0000 0x0001 "*"
' Line #1246:
' 	QuoteRem 0x0000 0x0048 "***********************************************************************/"
' Line #1247:
' Line #1248:
' 	Dim 
' 	VarDefn longitude (As Double)
' 	VarDefn latitude (As Double)
' Line #1249:
' 	Dim 
' 	VarDefn zone (As Double)
' 	VarDefn zone (As Double)
' Line #1250:
' 	Dim 
' 	VarDefn daySavings (As Double)
' 	VarDefn hh (As Double)
' 	VarDefn mm (As Double)
' 	VarDefn ss (As Double)
' Line #1251:
' 	Dim 
' 	VarDefn JD (As Double)
' 	VarDefn t (As Double)
' 	VarDefn R (As Double)
' Line #1252:
' 	Dim 
' 	VarDefn alpha (As Double)
' 	VarDefn theta (As Double)
' 	VarDefn Etime (As Double)
' 	VarDefn eqtime (As Double)
' Line #1253:
' 	Dim 
' 	VarDefn solarDec (As Double)
' 	VarDefn timenow (As Double)
' 	VarDefn earthRadVec (As Double)
' Line #1254:
' 	Dim 
' 	VarDefn solarTimeFix (As Double)
' 	VarDefn hourangle (As Double)
' 	VarDefn trueSolarTime (As Double)
' Line #1255:
' 	Dim 
' 	VarDefn harad (As Double)
' 	VarDefn csz (As Double)
' 	VarDefn zenith (As Double)
' 	VarDefn azDenom (As Double)
' Line #1256:
' 	Dim 
' 	VarDefn azRad (As Double)
' 	VarDefn azimuth (As Double)
' Line #1257:
' 	Dim 
' 	VarDefn exoatmElevation (As Double)
' 	VarDefn step1 (As Double)
' 	VarDefn step2 (As Double)
' Line #1258:
' 	Dim 
' 	VarDefn step3 (As Double)
' 	VarDefn refractionCorrection (As Double)
' 	VarDefn te (As Double)
' Line #1259:
' Line #1260:
' 	QuoteRem 0x0000 0x0055 " change sign convention for longitude from negative to positive in western hemisphere"
' Line #1261:
' 	Ld lon 
' 	LitDI2 0x0001 
' 	UMi 
' 	Mul 
' 	St longitude 
' Line #1262:
' 	Ld lat 
' 	St latitude 
' Line #1263:
' 	Ld latitude 
' 	LitR8 0x3333 0x3333 0x7333 0x4056 
' 	Gt 
' 	Paren 
' 	If 
' 	BoSImplicit 
' 	LitR8 0x3333 0x3333 0x7333 0x4056 
' 	St latitude 
' 	EndIf 
' Line #1264:
' 	Ld latitude 
' 	LitR8 0x3333 0x3333 0x7333 0x4056 
' 	UMi 
' 	Lt 
' 	Paren 
' 	If 
' 	BoSImplicit 
' 	LitR8 0x3333 0x3333 0x7333 0x4056 
' 	UMi 
' 	St latitude 
' 	EndIf 
' Line #1265:
' Line #1266:
' 	QuoteRem 0x0000 0x0039 "change time zone to ppositive hours in western hemisphere"
' Line #1267:
' 	Ld timezone 
' 	LitDI2 0x0001 
' 	UMi 
' 	Mul 
' 	St zone 
' Line #1268:
' 	Ld dlstime 
' 	LitDI2 0x003C 
' 	Mul 
' 	St zone 
' Line #1269:
' 	Ld hours 
' 	Ld zone 
' 	LitDI2 0x003C 
' 	Div 
' 	Paren 
' 	Sub 
' 	St daySavings 
' Line #1270:
' 	Ld minutes 
' 	St hh 
' Line #1271:
' 	Ld seconds 
' 	St mm 
' Line #1272:
' Line #1273:
' 	QuoteRem 0x0000 0x003B "//    timenow is GMT time for calculation in hours since 0Z"
' Line #1274:
' 	Ld daySavings 
' 	Ld hh 
' 	LitDI2 0x003C 
' 	Div 
' 	Add 
' 	Ld mm 
' 	LitDI2 0x0E10 
' 	Div 
' 	Add 
' 	Ld zone 
' 	Add 
' 	St ss 
' Line #1275:
' Line #1276:
' 	Ld year 
' 	Ld month 
' 	Ld day 
' 	ArgsLd calcJD 0x0003 
' 	St JD 
' Line #1277:
' 	Ld JD 
' 	Ld ss 
' 	LitR8 0x0000 0x0000 0x0000 0x4038 
' 	Div 
' 	Add 
' 	ArgsLd calcTimeJulianCent 0x0001 
' 	St t 
' Line #1278:
' 	Ld t 
' 	ArgsLd calcSunRadVector 0x0001 
' 	St R 
' Line #1279:
' 	Ld t 
' 	ArgsLd calcSunRtAscension 0x0001 
' 	St alpha 
' Line #1280:
' 	Ld t 
' 	ArgsLd calcSunDeclination 0x0001 
' 	St theta 
' Line #1281:
' 	Ld t 
' 	ArgsLd calcEquationOfTime 0x0001 
' 	St Etime 
' Line #1282:
' Line #1283:
' 	Ld Etime 
' 	St eqtime 
' Line #1284:
' 	Ld theta 
' 	St solarDec 
' 	QuoteRem 0x001D 0x0010 "//    in degrees"
' Line #1285:
' 	Ld R 
' 	St timenow 
' Line #1286:
' Line #1287:
' 	Ld eqtime 
' 	LitR8 0x0000 0x0000 0x0000 0x4010 
' 	Ld longitude 
' 	Mul 
' 	Sub 
' 	LitR8 0x0000 0x0000 0x0000 0x404E 
' 	Ld zone 
' 	Mul 
' 	Add 
' 	St earthRadVec 
' Line #1288:
' 	Ld daySavings 
' 	LitR8 0x0000 0x0000 0x0000 0x404E 
' 	Mul 
' 	Ld hh 
' 	Add 
' 	Ld mm 
' 	LitR8 0x0000 0x0000 0x0000 0x404E 
' 	Div 
' 	Add 
' 	Ld earthRadVec 
' 	Add 
' 	St solarTimeFix 
' Line #1289:
' 	QuoteRem 0x000C 0x0010 "//    in minutes"
' Line #1290:
' Line #1291:
' 	Ld solarTimeFix 
' 	LitDI2 0x05A0 
' 	Gt 
' 	Paren 
' 	DoWhile 
' Line #1292:
' 	Ld solarTimeFix 
' 	LitDI2 0x05A0 
' 	Sub 
' 	St solarTimeFix 
' Line #1293:
' 	Loop 
' Line #1294:
' Line #1295:
' 	Ld solarTimeFix 
' 	LitR8 0x0000 0x0000 0x0000 0x4010 
' 	Div 
' 	LitR8 0x0000 0x0000 0x8000 0x4066 
' 	Sub 
' 	St hourangle 
' Line #1296:
' 	QuoteRem 0x000C 0x0034 "//    Thanks to Louis Schwarzmayr for the next line:"
' Line #1297:
' 	Ld hourangle 
' 	LitDI2 0x00B4 
' 	UMi 
' 	Lt 
' 	Paren 
' 	If 
' 	BoSImplicit 
' 	Ld hourangle 
' 	LitR8 0x0000 0x0000 0x8000 0x4076 
' 	Add 
' 	St hourangle 
' 	EndIf 
' Line #1298:
' Line #1299:
' 	Ld hourangle 
' 	ArgsLd degToRad 0x0001 
' 	St trueSolarTime 
' Line #1300:
' Line #1301:
' 	LineCont 0x000C 0A 00 12 00 12 00 12 00 1A 00 12 00
' 	Ld latitude 
' 	ArgsLd degToRad 0x0001 
' 	ArgsLd Sin 0x0001 
' 	Ld solarDec 
' 	ArgsLd degToRad 0x0001 
' 	ArgsLd Sin 0x0001 
' 	Mul 
' 	Ld latitude 
' 	ArgsLd degToRad 0x0001 
' 	ArgsLd Cos 0x0001 
' 	Ld solarDec 
' 	ArgsLd degToRad 0x0001 
' 	ArgsLd Cos 0x0001 
' 	Mul 
' 	Ld trueSolarTime 
' 	ArgsLd Cos 0x0001 
' 	Mul 
' 	Add 
' 	St harad 
' Line #1302:
' Line #1303:
' 	Ld harad 
' 	LitR8 0x0000 0x0000 0x0000 0x3FF0 
' 	Gt 
' 	Paren 
' 	IfBlock 
' Line #1304:
' 	LitR8 0x0000 0x0000 0x0000 0x3FF0 
' 	St harad 
' Line #1305:
' 	Ld harad 
' 	LitR8 0x0000 0x0000 0x0000 0x3FF0 
' 	UMi 
' 	Lt 
' 	Paren 
' 	ElseIfBlock 
' Line #1306:
' 	LitR8 0x0000 0x0000 0x0000 0x3FF0 
' 	UMi 
' 	St harad 
' Line #1307:
' 	EndIfBlock 
' Line #1308:
' Line #1309:
' 	Ld harad 
' 	Ld Application 
' 	MemLd WorksheetFunction 
' 	ArgsMemLd Acos 0x0001 
' 	ArgsLd radToDeg 0x0001 
' 	St csz 
' Line #1310:
' Line #1311:
' 	Ld latitude 
' 	ArgsLd degToRad 0x0001 
' 	ArgsLd Cos 0x0001 
' 	Ld csz 
' 	ArgsLd degToRad 0x0001 
' 	ArgsLd Sin 0x0001 
' 	Mul 
' 	Paren 
' 	St zenith 
' Line #1312:
' Line #1313:
' 	Ld zenith 
' 	FnAbs 
' 	LitR8 0xA9FC 0xD2F1 0x624D 0x3F50 
' 	Gt 
' 	Paren 
' 	IfBlock 
' Line #1314:
' 	LineCont 0x0008 0C 00 14 00 15 00 14 00
' 	Ld latitude 
' 	ArgsLd degToRad 0x0001 
' 	ArgsLd Sin 0x0001 
' 	Ld csz 
' 	ArgsLd degToRad 0x0001 
' 	ArgsLd Cos 0x0001 
' 	Mul 
' 	Paren 
' 	Ld solarDec 
' 	ArgsLd degToRad 0x0001 
' 	ArgsLd Sin 0x0001 
' 	Sub 
' 	Paren 
' 	Ld zenith 
' 	Div 
' 	St azDenom 
' Line #1315:
' 	Ld azDenom 
' 	FnAbs 
' 	LitR8 0x0000 0x0000 0x0000 0x3FF0 
' 	Gt 
' 	Paren 
' 	IfBlock 
' Line #1316:
' 	Ld azDenom 
' 	LitDI2 0x0000 
' 	Lt 
' 	Paren 
' 	IfBlock 
' Line #1317:
' 	LitR8 0x0000 0x0000 0x0000 0x3FF0 
' 	UMi 
' 	St azDenom 
' Line #1318:
' 	ElseBlock 
' Line #1319:
' 	LitR8 0x0000 0x0000 0x0000 0x3FF0 
' 	St azDenom 
' Line #1320:
' 	EndIfBlock 
' Line #1321:
' 	EndIfBlock 
' Line #1322:
' Line #1323:
' 	LitR8 0x0000 0x0000 0x8000 0x4066 
' 	Ld azDenom 
' 	Ld Application 
' 	MemLd WorksheetFunction 
' 	ArgsMemLd Acos 0x0001 
' 	ArgsLd radToDeg 0x0001 
' 	Sub 
' 	St azRad 
' Line #1324:
' Line #1325:
' 	Ld hourangle 
' 	LitR8 0x0000 0x0000 0x0000 0x0000 
' 	Gt 
' 	Paren 
' 	IfBlock 
' Line #1326:
' 	Ld azRad 
' 	UMi 
' 	St azRad 
' Line #1327:
' 	EndIfBlock 
' Line #1328:
' 	ElseBlock 
' Line #1329:
' 	Ld latitude 
' 	LitR8 0x0000 0x0000 0x0000 0x0000 
' 	Gt 
' 	Paren 
' 	IfBlock 
' Line #1330:
' 	LitR8 0x0000 0x0000 0x8000 0x4066 
' 	St azRad 
' Line #1331:
' 	ElseBlock 
' Line #1332:
' 	LitR8 0x0000 0x0000 0x0000 0x0000 
' 	St azRad 
' Line #1333:
' 	EndIfBlock 
' Line #1334:
' 	EndIfBlock 
' Line #1335:
' 	Ld azRad 
' 	LitR8 0x0000 0x0000 0x0000 0x0000 
' 	Lt 
' 	Paren 
' 	IfBlock 
' Line #1336:
' 	Ld azRad 
' 	LitR8 0x0000 0x0000 0x8000 0x4076 
' 	Add 
' 	St azRad 
' Line #1337:
' 	EndIfBlock 
' Line #1338:
' Line #1339:
' 	LitR8 0x0000 0x0000 0x8000 0x4056 
' 	Ld csz 
' 	Sub 
' 	St azimuth 
' Line #1340:
' Line #1341:
' 	QuoteRem 0x0000 0x002D "beginning of complex expression commented out"
' Line #1342:
' 	QuoteRem 0x0000 0x002B "            If (exoatmElevation > 85#) Then"
' Line #1343:
' 	QuoteRem 0x0000 0x0029 "                refractionCorrection = 0#"
' Line #1344:
' 	QuoteRem 0x0000 0x0010 "            Else"
' Line #1345:
' 	QuoteRem 0x0000 0x0033 "                te = Tan(degToRad(exoatmElevation))"
' Line #1346:
' 	QuoteRem 0x0000 0x002E "                If (exoatmElevation > 5#) Then"
' Line #1347:
' 	LineCont 0x0004 01 00 B2 FF
' 	QuoteRem 0x0000 0x008A "                    refractionCorrection = 58.1 / te - 0.07 / (te * te * te) +'                        0.000086 / (te * te * te * te * te)"
' Line #1348:
' 	QuoteRem 0x0000 0x0036 "                ElseIf (exoatmElevation > -0.575) Then"
' Line #1349:
' 	LineCont 0x000C 01 00 BC FF 01 00 7F FF 01 00 4B FF
' 	QuoteRem 0x0000 0x00E8 "                    refractionCorrection = 1735# + exoatmElevation *'                        (-518.2 + exoatmElevation * (103.4 +'                        exoatmElevation * (-12.79 +'                        exoatmElevation * 0.711)))"
' Line #1350:
' 	QuoteRem 0x0000 0x0014 "                Else"
' Line #1351:
' 	QuoteRem 0x0000 0x0037 "                    refractionCorrection = -20.774 / te"
' Line #1352:
' 	QuoteRem 0x0000 0x0016 "                End If"
' Line #1353:
' 	QuoteRem 0x0000 0x0043 "                refractionCorrection = refractionCorrection / 3600#"
' Line #1354:
' 	QuoteRem 0x0000 0x0012 "            End If"
' Line #1355:
' 	QuoteRem 0x0000 0x0019 "end of complex expression"
' Line #1356:
' Line #1357:
' 	QuoteRem 0x0000 0x0022 "beginning of simplified expression"
' Line #1358:
' 	Ld azimuth 
' 	LitR8 0x0000 0x0000 0x4000 0x4055 
' 	Gt 
' 	Paren 
' 	IfBlock 
' Line #1359:
' 	LitR8 0x0000 0x0000 0x0000 0x0000 
' 	St step3 
' Line #1360:
' 	ElseBlock 
' Line #1361:
' 	Ld azimuth 
' 	ArgsLd degToRad 0x0001 
' 	ArgsLd Tan 0x0001 
' 	St refractionCorrection 
' Line #1362:
' 	Ld azimuth 
' 	LitR8 0x0000 0x0000 0x0000 0x4014 
' 	Gt 
' 	Paren 
' 	IfBlock 
' Line #1363:
' 	LineCont 0x0004 10 00 18 00
' 	LitR8 0xCCCD 0xCCCC 0x0CCC 0x404D 
' 	Ld refractionCorrection 
' 	Div 
' 	LitR8 0x51EC 0x1EB8 0xEB85 0x3FB1 
' 	Ld refractionCorrection 
' 	Ld refractionCorrection 
' 	Mul 
' 	Ld refractionCorrection 
' 	Mul 
' 	Paren 
' 	Div 
' 	Sub 
' 	LitR8 0x7736 0xBFF4 0x8B5C 0x3F16 
' 	Ld refractionCorrection 
' 	Ld refractionCorrection 
' 	Mul 
' 	Ld refractionCorrection 
' 	Mul 
' 	Ld refractionCorrection 
' 	Mul 
' 	Ld refractionCorrection 
' 	Mul 
' 	Paren 
' 	Div 
' 	Add 
' 	St step3 
' Line #1364:
' 	Ld azimuth 
' 	LitR8 0x6666 0x6666 0x6666 0x3FE2 
' 	UMi 
' 	Gt 
' 	Paren 
' 	ElseIfBlock 
' Line #1365:
' 	LitR8 0xAE14 0xE147 0x947A 0x4029 
' 	UMi 
' 	Ld azimuth 
' 	LitR8 0x978D 0x126E 0xC083 0x3FE6 
' 	Mul 
' 	Add 
' 	Paren 
' 	St exoatmElevation 
' Line #1366:
' 	LitR8 0x999A 0x9999 0xD999 0x4059 
' 	Ld azimuth 
' 	Ld exoatmElevation 
' 	Paren 
' 	Mul 
' 	Add 
' 	Paren 
' 	St step1 
' Line #1367:
' 	LitR8 0x999A 0x9999 0x3199 0x4080 
' 	UMi 
' 	Ld azimuth 
' 	Ld step1 
' 	Paren 
' 	Mul 
' 	Add 
' 	Paren 
' 	St step2 
' Line #1368:
' 	LitR8 0x0000 0x0000 0x1C00 0x409B 
' 	Ld azimuth 
' 	Ld step2 
' 	Paren 
' 	Mul 
' 	Add 
' 	St step3 
' Line #1369:
' 	ElseBlock 
' Line #1370:
' 	LitR8 0x1AA0 0xDD2F 0xC624 0x4034 
' 	UMi 
' 	Ld refractionCorrection 
' 	Div 
' 	St step3 
' Line #1371:
' 	EndIfBlock 
' Line #1372:
' 	Ld step3 
' 	LitR8 0x0000 0x0000 0x2000 0x40AC 
' 	Div 
' 	St step3 
' Line #1373:
' 	EndIfBlock 
' Line #1374:
' 	QuoteRem 0x0000 0x001C "end of simplified expression"
' Line #1375:
' Line #1376:
' 	Ld csz 
' 	Ld step3 
' 	Sub 
' 	St te 
' Line #1377:
' Line #1378:
' 	QuoteRem 0x0000 0x0025 "            If (solarZen < 108#) Then"
' Line #1379:
' 	QuoteRem 0x0000 0x0024 "              solarazimuth = azimuth"
' Line #1380:
' 	LitR8 0x0000 0x0000 0x8000 0x4056 
' 	Ld te 
' 	Sub 
' 	St solarzen 
' Line #1381:
' 	QuoteRem 0x0000 0x0026 "              If (solarZen < 90#) Then"
' Line #1382:
' 	QuoteRem 0x0000 0x0030 "                coszen = Cos(degToRad(solarZen))"
' Line #1383:
' 	QuoteRem 0x0000 0x0012 "              Else"
' Line #1384:
' 	QuoteRem 0x0000 0x001B "                coszen = 0#"
' Line #1385:
' 	QuoteRem 0x0000 0x0014 "              End If"
' Line #1386:
' 	QuoteRem 0x0000 0x0042 "            Else    '// do not report az & el after astro twilight"
' Line #1387:
' 	QuoteRem 0x0000 0x0024 "              solarazimuth = -999999"
' Line #1388:
' 	QuoteRem 0x0000 0x0026 "              solarelevation = -999999"
' Line #1389:
' 	QuoteRem 0x0000 0x001E "              coszen = -999999"
' Line #1390:
' 	QuoteRem 0x0000 0x0012 "            End If"
' Line #1391:
' Line #1392:
' 	EndFunc 
' Line #1393:
' Line #1394:
' Line #1395:
' 	LineCont 0x0004 0D 00 08 00
' 	FuncDefn (Sub solarelevation(lat, lon, year, month, day, hours, minutes, seconds, timezone, dlstime, solarazimuth, solarzen))
' Line #1396:
' Line #1397:
' 	QuoteRem 0x0000 0x0048 "***********************************************************************/"
' Line #1398:
' 	QuoteRem 0x0000 0x0017 "* Name:    solarazimuth"
' Line #1399:
' 	QuoteRem 0x0000 0x0018 "* Type:    Main Function"
' Line #1400:
' 	QuoteRem 0x0000 0x0043 "* Purpose: calculate solar azimuth (deg from north) for the entered"
' Line #1401:
' 	QuoteRem 0x0000 0x004B "*          date, time and location. Returns -999999 if darker than twilight"
' Line #1402:
' 	QuoteRem 0x0000 0x0001 "*"
' Line #1403:
' 	QuoteRem 0x0000 0x000C "* Arguments:"
' Line #1404:
' 	QuoteRem 0x0000 0x0040 "*   latitude, longitude, year, month, day, hour, minute, second,"
' Line #1405:
' 	QuoteRem 0x0000 0x0021 "*   timezone, daylightsavingstime"
' Line #1406:
' 	QuoteRem 0x0000 0x000F "* Return value:"
' Line #1407:
' 	QuoteRem 0x0000 0x0027 "*   solar azimuth in degrees from north"
' Line #1408:
' 	QuoteRem 0x0000 0x0001 "*"
' Line #1409:
' 	QuoteRem 0x0000 0x003F "* Note: solarelevation and solarazimuth functions are identical"
' Line #1410:
' 	QuoteRem 0x0000 0x0041 "*       and could converted to a VBA subroutine that would return"
' Line #1411:
' 	QuoteRem 0x0000 0x0014 "*       both values."
' Line #1412:
' 	QuoteRem 0x0000 0x0001 "*"
' Line #1413:
' 	QuoteRem 0x0000 0x0048 "***********************************************************************/"
' Line #1414:
' Line #1415:
' 	Dim 
' 	VarDefn longitude (As Double)
' 	VarDefn latitude (As Double)
' Line #1416:
' 	Dim 
' 	VarDefn zone (As Double)
' 	VarDefn zone (As Double)
' Line #1417:
' 	Dim 
' 	VarDefn daySavings (As Double)
' 	VarDefn hh (As Double)
' 	VarDefn mm (As Double)
' 	VarDefn ss (As Double)
' Line #1418:
' 	Dim 
' 	VarDefn JD (As Double)
' 	VarDefn t (As Double)
' 	VarDefn R (As Double)
' Line #1419:
' 	Dim 
' 	VarDefn alpha (As Double)
' 	VarDefn theta (As Double)
' 	VarDefn Etime (As Double)
' 	VarDefn eqtime (As Double)
' Line #1420:
' 	Dim 
' 	VarDefn solarDec (As Double)
' 	VarDefn timenow (As Double)
' 	VarDefn earthRadVec (As Double)
' Line #1421:
' 	Dim 
' 	VarDefn solarTimeFix (As Double)
' 	VarDefn hourangle (As Double)
' 	VarDefn trueSolarTime (As Double)
' Line #1422:
' 	Dim 
' 	VarDefn harad (As Double)
' 	VarDefn csz (As Double)
' 	VarDefn zenith (As Double)
' 	VarDefn azDenom (As Double)
' Line #1423:
' 	Dim 
' 	VarDefn azRad (As Double)
' 	VarDefn azimuth (As Double)
' Line #1424:
' 	Dim 
' 	VarDefn exoatmElevation (As Double)
' 	VarDefn step1 (As Double)
' 	VarDefn step2 (As Double)
' Line #1425:
' 	Dim 
' 	VarDefn step3 (As Double)
' 	VarDefn refractionCorrection (As Double)
' 	VarDefn te (As Double)
' Line #1426:
' Line #1427:
' 	QuoteRem 0x0000 0x0055 " change sign convention for longitude from negative to positive in western hemisphere"
' Line #1428:
' 	Ld lon 
' 	LitDI2 0x0001 
' 	UMi 
' 	Mul 
' 	St longitude 
' Line #1429:
' 	Ld lat 
' 	St latitude 
' Line #1430:
' 	Ld latitude 
' 	LitR8 0x3333 0x3333 0x7333 0x4056 
' 	Gt 
' 	Paren 
' 	If 
' 	BoSImplicit 
' 	LitR8 0x3333 0x3333 0x7333 0x4056 
' 	St latitude 
' 	EndIf 
' Line #1431:
' 	Ld latitude 
' 	LitR8 0x3333 0x3333 0x7333 0x4056 
' 	UMi 
' 	Lt 
' 	Paren 
' 	If 
' 	BoSImplicit 
' 	LitR8 0x3333 0x3333 0x7333 0x4056 
' 	UMi 
' 	St latitude 
' 	EndIf 
' Line #1432:
' Line #1433:
' 	QuoteRem 0x0000 0x0039 "change time zone to ppositive hours in western hemisphere"
' Line #1434:
' 	Ld timezone 
' 	LitDI2 0x0001 
' 	UMi 
' 	Mul 
' 	St zone 
' Line #1435:
' 	Ld dlstime 
' 	LitDI2 0x003C 
' 	Mul 
' 	St zone 
' Line #1436:
' 	Ld hours 
' 	Ld zone 
' 	LitDI2 0x003C 
' 	Div 
' 	Paren 
' 	Sub 
' 	St daySavings 
' Line #1437:
' 	Ld minutes 
' 	St hh 
' Line #1438:
' 	Ld seconds 
' 	St mm 
' Line #1439:
' Line #1440:
' 	QuoteRem 0x0000 0x003B "//    timenow is GMT time for calculation in hours since 0Z"
' Line #1441:
' 	Ld daySavings 
' 	Ld hh 
' 	LitDI2 0x003C 
' 	Div 
' 	Add 
' 	Ld mm 
' 	LitDI2 0x0E10 
' 	Div 
' 	Add 
' 	Ld zone 
' 	Add 
' 	St ss 
' Line #1442:
' Line #1443:
' 	Ld year 
' 	Ld month 
' 	Ld day 
' 	ArgsLd calcJD 0x0003 
' 	St JD 
' Line #1444:
' 	Ld JD 
' 	Ld ss 
' 	LitR8 0x0000 0x0000 0x0000 0x4038 
' 	Div 
' 	Add 
' 	ArgsLd calcTimeJulianCent 0x0001 
' 	St t 
' Line #1445:
' 	Ld t 
' 	ArgsLd calcSunRadVector 0x0001 
' 	St R 
' Line #1446:
' 	Ld t 
' 	ArgsLd calcSunRtAscension 0x0001 
' 	St alpha 
' Line #1447:
' 	Ld t 
' 	ArgsLd calcSunDeclination 0x0001 
' 	St theta 
' Line #1448:
' 	Ld t 
' 	ArgsLd calcEquationOfTime 0x0001 
' 	St Etime 
' Line #1449:
' Line #1450:
' 	Ld Etime 
' 	St eqtime 
' Line #1451:
' 	Ld theta 
' 	St solarDec 
' 	QuoteRem 0x001D 0x0010 "//    in degrees"
' Line #1452:
' 	Ld R 
' 	St timenow 
' Line #1453:
' Line #1454:
' 	Ld eqtime 
' 	LitR8 0x0000 0x0000 0x0000 0x4010 
' 	Ld longitude 
' 	Mul 
' 	Sub 
' 	LitR8 0x0000 0x0000 0x0000 0x404E 
' 	Ld zone 
' 	Mul 
' 	Add 
' 	St earthRadVec 
' Line #1455:
' 	Ld daySavings 
' 	LitR8 0x0000 0x0000 0x0000 0x404E 
' 	Mul 
' 	Ld hh 
' 	Add 
' 	Ld mm 
' 	LitR8 0x0000 0x0000 0x0000 0x404E 
' 	Div 
' 	Add 
' 	Ld earthRadVec 
' 	Add 
' 	St solarTimeFix 
' Line #1456:
' 	QuoteRem 0x000C 0x0010 "//    in minutes"
' Line #1457:
' Line #1458:
' 	Ld solarTimeFix 
' 	LitDI2 0x05A0 
' 	Gt 
' 	Paren 
' 	DoWhile 
' Line #1459:
' 	Ld solarTimeFix 
' 	LitDI2 0x05A0 
' 	Sub 
' 	St solarTimeFix 
' Line #1460:
' 	Loop 
' Line #1461:
' Line #1462:
' 	Ld solarTimeFix 
' 	LitR8 0x0000 0x0000 0x0000 0x4010 
' 	Div 
' 	LitR8 0x0000 0x0000 0x8000 0x4066 
' 	Sub 
' 	St hourangle 
' Line #1463:
' 	QuoteRem 0x000C 0x0034 "//    Thanks to Louis Schwarzmayr for the next line:"
' Line #1464:
' 	Ld hourangle 
' 	LitDI2 0x00B4 
' 	UMi 
' 	Lt 
' 	Paren 
' 	If 
' 	BoSImplicit 
' 	Ld hourangle 
' 	LitR8 0x0000 0x0000 0x8000 0x4076 
' 	Add 
' 	St hourangle 
' 	EndIf 
' Line #1465:
' Line #1466:
' 	Ld hourangle 
' 	ArgsLd degToRad 0x0001 
' 	St trueSolarTime 
' Line #1467:
' Line #1468:
' 	LineCont 0x000C 0A 00 12 00 12 00 12 00 1A 00 12 00
' 	Ld latitude 
' 	ArgsLd degToRad 0x0001 
' 	ArgsLd Sin 0x0001 
' 	Ld solarDec 
' 	ArgsLd degToRad 0x0001 
' 	ArgsLd Sin 0x0001 
' 	Mul 
' 	Ld latitude 
' 	ArgsLd degToRad 0x0001 
' 	ArgsLd Cos 0x0001 
' 	Ld solarDec 
' 	ArgsLd degToRad 0x0001 
' 	ArgsLd Cos 0x0001 
' 	Mul 
' 	Ld trueSolarTime 
' 	ArgsLd Cos 0x0001 
' 	Mul 
' 	Add 
' 	St harad 
' Line #1469:
' Line #1470:
' 	Ld harad 
' 	LitR8 0x0000 0x0000 0x0000 0x3FF0 
' 	Gt 
' 	Paren 
' 	IfBlock 
' Line #1471:
' 	LitR8 0x0000 0x0000 0x0000 0x3FF0 
' 	St harad 
' Line #1472:
' 	Ld harad 
' 	LitR8 0x0000 0x0000 0x0000 0x3FF0 
' 	UMi 
' 	Lt 
' 	Paren 
' 	ElseIfBlock 
' Line #1473:
' 	LitR8 0x0000 0x0000 0x0000 0x3FF0 
' 	UMi 
' 	St harad 
' Line #1474:
' 	EndIfBlock 
' Line #1475:
' Line #1476:
' 	Ld harad 
' 	Ld Application 
' 	MemLd WorksheetFunction 
' 	ArgsMemLd Acos 0x0001 
' 	ArgsLd radToDeg 0x0001 
' 	St csz 
' Line #1477:
' Line #1478:
' 	Ld latitude 
' 	ArgsLd degToRad 0x0001 
' 	ArgsLd Cos 0x0001 
' 	Ld csz 
' 	ArgsLd degToRad 0x0001 
' 	ArgsLd Sin 0x0001 
' 	Mul 
' 	Paren 
' 	St zenith 
' Line #1479:
' Line #1480:
' 	Ld zenith 
' 	FnAbs 
' 	LitR8 0xA9FC 0xD2F1 0x624D 0x3F50 
' 	Gt 
' 	Paren 
' 	IfBlock 
' Line #1481:
' 	LineCont 0x0008 0C 00 14 00 15 00 14 00
' 	Ld latitude 
' 	ArgsLd degToRad 0x0001 
' 	ArgsLd Sin 0x0001 
' 	Ld csz 
' 	ArgsLd degToRad 0x0001 
' 	ArgsLd Cos 0x0001 
' 	Mul 
' 	Paren 
' 	Ld solarDec 
' 	ArgsLd degToRad 0x0001 
' 	ArgsLd Sin 0x0001 
' 	Sub 
' 	Paren 
' 	Ld zenith 
' 	Div 
' 	St azDenom 
' Line #1482:
' 	Ld azDenom 
' 	FnAbs 
' 	LitR8 0x0000 0x0000 0x0000 0x3FF0 
' 	Gt 
' 	Paren 
' 	IfBlock 
' Line #1483:
' 	Ld azDenom 
' 	LitDI2 0x0000 
' 	Lt 
' 	Paren 
' 	IfBlock 
' Line #1484:
' 	LitR8 0x0000 0x0000 0x0000 0x3FF0 
' 	UMi 
' 	St azDenom 
' Line #1485:
' 	ElseBlock 
' Line #1486:
' 	LitR8 0x0000 0x0000 0x0000 0x3FF0 
' 	St azDenom 
' Line #1487:
' 	EndIfBlock 
' Line #1488:
' 	EndIfBlock 
' Line #1489:
' Line #1490:
' 	LitR8 0x0000 0x0000 0x8000 0x4066 
' 	Ld azDenom 
' 	Ld Application 
' 	MemLd WorksheetFunction 
' 	ArgsMemLd Acos 0x0001 
' 	ArgsLd radToDeg 0x0001 
' 	Sub 
' 	St azRad 
' Line #1491:
' Line #1492:
' 	Ld hourangle 
' 	LitR8 0x0000 0x0000 0x0000 0x0000 
' 	Gt 
' 	Paren 
' 	IfBlock 
' Line #1493:
' 	Ld azRad 
' 	UMi 
' 	St azRad 
' Line #1494:
' 	EndIfBlock 
' Line #1495:
' 	ElseBlock 
' Line #1496:
' 	Ld latitude 
' 	LitR8 0x0000 0x0000 0x0000 0x0000 
' 	Gt 
' 	Paren 
' 	IfBlock 
' Line #1497:
' 	LitR8 0x0000 0x0000 0x8000 0x4066 
' 	St azRad 
' Line #1498:
' 	ElseBlock 
' Line #1499:
' 	LitR8 0x0000 0x0000 0x0000 0x0000 
' 	St azRad 
' Line #1500:
' 	EndIfBlock 
' Line #1501:
' 	EndIfBlock 
' Line #1502:
' 	Ld azRad 
' 	LitR8 0x0000 0x0000 0x0000 0x0000 
' 	Lt 
' 	Paren 
' 	IfBlock 
' Line #1503:
' 	Ld azRad 
' 	LitR8 0x0000 0x0000 0x8000 0x4076 
' 	Add 
' 	St azRad 
' Line #1504:
' 	EndIfBlock 
' Line #1505:
' Line #1506:
' 	LitR8 0x0000 0x0000 0x8000 0x4056 
' 	Ld csz 
' 	Sub 
' 	St azimuth 
' Line #1507:
' Line #1508:
' 	QuoteRem 0x0000 0x002D "beginning of complex expression commented out"
' Line #1509:
' 	QuoteRem 0x0000 0x002B "            If (exoatmElevation > 85#) Then"
' Line #1510:
' 	QuoteRem 0x0000 0x0029 "                refractionCorrection = 0#"
' Line #1511:
' 	QuoteRem 0x0000 0x0010 "            Else"
' Line #1512:
' 	QuoteRem 0x0000 0x0033 "                te = Tan(degToRad(exoatmElevation))"
' Line #1513:
' 	QuoteRem 0x0000 0x002E "                If (exoatmElevation > 5#) Then"
' Line #1514:
' 	LineCont 0x0004 01 00 B2 FF
' 	QuoteRem 0x0000 0x008A "                    refractionCorrection = 58.1 / te - 0.07 / (te * te * te) +'                        0.000086 / (te * te * te * te * te)"
' Line #1515:
' 	QuoteRem 0x0000 0x0036 "                ElseIf (exoatmElevation > -0.575) Then"
' Line #1516:
' 	LineCont 0x000C 01 00 BC FF 01 00 7F FF 01 00 4B FF
' 	QuoteRem 0x0000 0x00E8 "                    refractionCorrection = 1735# + exoatmElevation *'                        (-518.2 + exoatmElevation * (103.4 +'                        exoatmElevation * (-12.79 +'                        exoatmElevation * 0.711)))"
' Line #1517:
' 	QuoteRem 0x0000 0x0014 "                Else"
' Line #1518:
' 	QuoteRem 0x0000 0x0037 "                    refractionCorrection = -20.774 / te"
' Line #1519:
' 	QuoteRem 0x0000 0x0016 "                End If"
' Line #1520:
' 	QuoteRem 0x0000 0x0043 "                refractionCorrection = refractionCorrection / 3600#"
' Line #1521:
' 	QuoteRem 0x0000 0x0012 "            End If"
' Line #1522:
' 	QuoteRem 0x0000 0x0019 "end of complex expression"
' Line #1523:
' Line #1524:
' Line #1525:
' 	QuoteRem 0x0000 0x0022 "beginning of simplified expression"
' Line #1526:
' 	Ld azimuth 
' 	LitR8 0x0000 0x0000 0x4000 0x4055 
' 	Gt 
' 	Paren 
' 	IfBlock 
' Line #1527:
' 	LitR8 0x0000 0x0000 0x0000 0x0000 
' 	St step3 
' Line #1528:
' 	ElseBlock 
' Line #1529:
' 	Ld azimuth 
' 	ArgsLd degToRad 0x0001 
' 	ArgsLd Tan 0x0001 
' 	St refractionCorrection 
' Line #1530:
' 	Ld azimuth 
' 	LitR8 0x0000 0x0000 0x0000 0x4014 
' 	Gt 
' 	Paren 
' 	IfBlock 
' Line #1531:
' 	LineCont 0x0004 10 00 18 00
' 	LitR8 0xCCCD 0xCCCC 0x0CCC 0x404D 
' 	Ld refractionCorrection 
' 	Div 
' 	LitR8 0x51EC 0x1EB8 0xEB85 0x3FB1 
' 	Ld refractionCorrection 
' 	Ld refractionCorrection 
' 	Mul 
' 	Ld refractionCorrection 
' 	Mul 
' 	Paren 
' 	Div 
' 	Sub 
' 	LitR8 0x7736 0xBFF4 0x8B5C 0x3F16 
' 	Ld refractionCorrection 
' 	Ld refractionCorrection 
' 	Mul 
' 	Ld refractionCorrection 
' 	Mul 
' 	Ld refractionCorrection 
' 	Mul 
' 	Ld refractionCorrection 
' 	Mul 
' 	Paren 
' 	Div 
' 	Add 
' 	St step3 
' Line #1532:
' 	Ld azimuth 
' 	LitR8 0x6666 0x6666 0x6666 0x3FE2 
' 	UMi 
' 	Gt 
' 	Paren 
' 	ElseIfBlock 
' Line #1533:
' 	LitR8 0xAE14 0xE147 0x947A 0x4029 
' 	UMi 
' 	Ld azimuth 
' 	LitR8 0x978D 0x126E 0xC083 0x3FE6 
' 	Mul 
' 	Add 
' 	Paren 
' 	St exoatmElevation 
' Line #1534:
' 	LitR8 0x999A 0x9999 0xD999 0x4059 
' 	Ld azimuth 
' 	Ld exoatmElevation 
' 	Paren 
' 	Mul 
' 	Add 
' 	Paren 
' 	St step1 
' Line #1535:
' 	LitR8 0x999A 0x9999 0x3199 0x4080 
' 	UMi 
' 	Ld azimuth 
' 	Ld step1 
' 	Paren 
' 	Mul 
' 	Add 
' 	Paren 
' 	St step2 
' Line #1536:
' 	LitR8 0x0000 0x0000 0x1C00 0x409B 
' 	Ld azimuth 
' 	Ld step2 
' 	Paren 
' 	Mul 
' 	Add 
' 	St step3 
' Line #1537:
' 	ElseBlock 
' Line #1538:
' 	LitR8 0x1AA0 0xDD2F 0xC624 0x4034 
' 	UMi 
' 	Ld refractionCorrection 
' 	Div 
' 	St step3 
' Line #1539:
' 	EndIfBlock 
' Line #1540:
' 	Ld step3 
' 	LitR8 0x0000 0x0000 0x2000 0x40AC 
' 	Div 
' 	St step3 
' Line #1541:
' 	EndIfBlock 
' Line #1542:
' 	QuoteRem 0x0000 0x001C "end of simplified expression"
' Line #1543:
' Line #1544:
' Line #1545:
' 	Ld csz 
' 	Ld step3 
' 	Sub 
' 	St te 
' Line #1546:
' Line #1547:
' 	QuoteRem 0x0000 0x0025 "            If (solarZen < 108#) Then"
' Line #1548:
' 	Ld azRad 
' 	St solarazimuth 
' Line #1549:
' 	LitR8 0x0000 0x0000 0x8000 0x4056 
' 	Ld te 
' 	Sub 
' 	St solarzen 
' Line #1550:
' 	QuoteRem 0x0000 0x0026 "              If (solarZen < 90#) Then"
' Line #1551:
' 	QuoteRem 0x0000 0x0030 "                coszen = Cos(degToRad(solarZen))"
' Line #1552:
' 	QuoteRem 0x0000 0x0012 "              Else"
' Line #1553:
' 	QuoteRem 0x0000 0x001B "                coszen = 0#"
' Line #1554:
' 	QuoteRem 0x0000 0x0014 "              End If"
' Line #1555:
' 	QuoteRem 0x0000 0x0042 "            Else    '// do not report az & el after astro twilight"
' Line #1556:
' 	QuoteRem 0x0000 0x0024 "              solarazimuth = -999999"
' Line #1557:
' 	QuoteRem 0x0000 0x0026 "              solarelevation = -999999"
' Line #1558:
' 	QuoteRem 0x0000 0x001E "              coszen = -999999"
' Line #1559:
' 	QuoteRem 0x0000 0x0012 "            End If"
' Line #1560:
' Line #1561:
' 	EndSub 
' Line #1562:
' VBA/Sheet7 - 1105 bytes
+----------+--------------------+---------------------------------------------+
|Type      |Keyword             |Description                                  |
+----------+--------------------+---------------------------------------------+
|AutoExec  |Workbook_Open       |Runs when the Excel Workbook is opened       |
|Suspicious|PUT                 |May write to a file (if combined with Open)  |
|Suspicious|run                 |May run an executable file or a system       |
|          |                    |command                                      |
|Suspicious|Hex Strings         |Hex-encoded strings were detected, may be    |
|          |                    |used to obfuscate strings (option --decode to|
|          |                    |see all)                                     |
|Suspicious|Base64 Strings      |Base64-encoded strings were detected, may be |
|          |                    |used to obfuscate strings (option --decode to|
|          |                    |see all)                                     |
|IOC       |http://www.srrb.noaa|URL                                          |
|          |.gov/highlights/sunr|                                             |
|          |ise/sunrise.html    |                                             |
|IOC       |http://www.srrb.noaa|URL                                          |
|          |.gov/highlights/sunr|                                             |
|          |ise/azel.html       |                                             |
|IOC       |http://www.srrb.noaa|URL                                          |
|          |.gov/highlights/sunr|                                             |
|          |ise/calcdetails.html|                                             |
|Base64    |8#6                 |OCM2                                         |
|String    |                    |                                             |
|Base64    |<#6                 |PCM2                                         |
|String    |                    |                                             |
|Suspicious|VBA Stomping        |VBA Stomping was detected: the VBA source    |
|          |                    |code and P-code are different, this may have |
|          |                    |been used to hide malicious code             |
+----------+--------------------+---------------------------------------------+
VBA Stomping detection is experimental: please report any false positive/negative at https://github.com/decalage2/oletools/issues