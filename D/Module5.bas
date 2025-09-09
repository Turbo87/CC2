' VBA Module: Window Management and View Control
' Purpose: Manages window resizing, view control, and display formatting for waypoint interface.
' Handles screen zoom, row visibility, and worksheet layout for waypoint entry mode.

Option Explicit
Sub resz()
Attribute resz.VB_ProcData.VB_Invoke_Func = " \n14"
'
' JLR 9/26/2011
'
Application.ScreenUpdating = False
Application.DisplayFullScreen = True
Workbooks("D.xlsm").Unprotect Password:="spike"
Sheets("Saved Way Points").Select
ActiveSheet.Unprotect Password:="spike"
Range("B2").FormulaR1C1 = "For Written Declarations:  ADD A SAVED WAY POINT"
Range("A4:A13").EntireRow.Hidden = False
Range("A28:A40").EntireRow.Hidden = False
ActiveSheet.Shapes("Rectangle 1").Visible = True
ActiveSheet.Shapes("Rectangle 2").Visible = False
ActiveSheet.Shapes("Drop Down 1").Visible = True
Range("B16:B40").Locked = False
Range("A1").ColumnWidth = 25
Range("A1:M22").Select
ActiveWindow.Zoom = True
Range("D4").FormulaR1C1 = "1"
ActiveWindow.ScrollRow = 1
ActiveSheet.Protect Password:="spike"
Range("D4").Select
Workbooks("D.xlsm").Protect Password:="spike"
Application.ScreenUpdating = True
'
End Sub