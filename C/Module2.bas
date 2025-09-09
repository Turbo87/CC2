' VBA Module: Calibration Data Management
' Purpose: Manages calibration data loading and worksheet visibility controls.
' Handles calibration value transfers between sheets and conditional sheet visibility
' based on existing calibration data availability.

Option Explicit
Sub FreeMe()
Attribute FreeMe.VB_ProcData.VB_Invoke_Func = " \n14"
'
' JLR 4/11/14 Go to CC, Calibration visible if not previously provided
'
    Application.ScreenUpdating = False
    ActiveWorkbook.Unprotect Password:="spike"

    Sheets("YDWK1").Visible = True
    Sheets("YDWK1").Activate
    Sheets("YDWK1").Unprotect Password:="spike"
    Sheets("YDWK1").Range("J1").Value = Sheets("B").Range("A44").Value
    'Sheets("YDWK1").Range("K1").Value = Sheets("Worksheet").Range("N1").Value
    Sheets("YDWK1").Range("L1").Value = Sheets("Worksheet").Range("Q1").Value

If Range("J1") = "" And Range("K1") = "" And Range("L1") = "" Then
    Range("E4").FormulaR1C1 = "=Worksheet!R[1]C[8]"
    Range("F7:G14").FormulaR1C1 = "=CALIBRATION!R[-1]C[1]"
    Range("I7:J14").FormulaR1C1 = "=CALIBRATION!R[7]C[-2]"
    Sheets("Calibration").Visible = True
     
ElseIf Range("J1") > 0 Or Range("K1") <> "" Or Range("L1") <> "" Then
    Sheets("YDWK1").Range("E4").Value = Sheets("B").Range("A42").Value
    Sheets("YDWK1").Range("F7:G14").Value = Sheets("B").Range("A43:B50").Value
    Sheets("YDWK1").Range("I7:J19").Value = Sheets("B").Range("A51:B63").Value
    Sheets("Calibration").Visible = False
End If
Sheets("YDWK1").Protect Password:="spike"
Sheets("YDWK1").Visible = False
    
    Sheets("Claim Check").Visible = True
    ActiveWorkbook.Protect Password:="spike"
    Sheets("Claim Check").Activate
    ActiveWindow.DisplayVerticalScrollBar = True
    Sheets("Claim Check").Protect Password:="spike"
    Application.ScreenUpdating = True
End Sub