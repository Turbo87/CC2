Option Explicit
Sub List()
Attribute List.VB_ProcData.VB_Invoke_Func = " \n14"
'
' JLR 9/26/11
'
Application.ScreenUpdating = False
Workbooks("D.xlsm").Activate
Sheets("Saved Way Points").Select
ActiveSheet.Unprotect Password:="spike"
Range("B2").FormulaR1C1 = "For Written Declarations:  SAVED WAY POINTS"
ActiveSheet.Shapes("Rectangle 1").Visible = False
If Range("O10") < 25 Then
ActiveSheet.Shapes("Rectangle 2").Visible = True
Else: ActiveSheet.Shapes("Rectangle 2").Visible = False
End If
ActiveSheet.Shapes("Drop Down 1").Visible = False
Range("A4:A13").Select
Selection.EntireRow.Hidden = True
Workbooks("D.xlsm").Sheets("Saved Way Points").Range("D4").Value = Workbooks("A.xlsm").Sheets("OTHER").Range("D15").Value
If Range("D4") = 2 Then
    Range("E1").EntireColumn.Hidden = False
    Range("F1").EntireColumn.Hidden = True
    Range("J1").EntireColumn.Hidden = False
    Range("K1").EntireColumn.Hidden = True
    Range("F9,K9").ClearContents
    ActiveSheet.Shapes("Picture 3").Visible = True
    ActiveSheet.Shapes("Picture 6").Visible = False
    Range("B9").Select
ElseIf Range("D4") = 3 Then
    Range("E1").EntireColumn.Hidden = True
    Range("F1").EntireColumn.Hidden = False
    Range("J1").EntireColumn.Hidden = True
    Range("K1").EntireColumn.Hidden = False
    Range("E9,J9").ClearContents
    ActiveSheet.Shapes("Picture 3").Visible = False
    ActiveSheet.Shapes("Picture 6").Visible = True
    Range("B9").Select
End If
If Range("O10") < 12 Then
    Range("A28:A40").Select
    Selection.EntireRow.Hidden = True
Else
    Range("A1:M47").Select
    ActiveWindow.Zoom = True
    Range("A1").ColumnWidth = 55
    ActiveWindow.DisplayVerticalScrollBar = False
End If
Range("B16:B40").Locked = True
ActiveWindow.ScrollRow = 1
Range("B2").Select
ActiveSheet.Protect Password:="spike"
Application.ScreenUpdating = True
'
End Sub