' VBA Class: Workbook Event Handler
' Purpose: Manages workbook initialization and setup for waypoint management application.
' Configures display settings, sheet protection, and initial view state on workbook open.

Option Explicit
Private Sub Workbook_Open()

Application.DisplayFullScreen = True
Application.EnableAutoComplete = False
Workbooks("D.xlsm").Unprotect Password:="spike"
Sheets("Calibration").Select
ActiveSheet.Unprotect Password:="spike"
Range("A1").Select
ActiveSheet.Protect Password:="spike"
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
Application.MoveAfterReturnDirection = xlToRight
Workbooks("D.xlsm").Protect Password:="spike"
    
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
Workbooks("D.xlsm").Saved = True
Application.DisplayFullScreen = False
End Sub