' VBA Class: Workbook Event Handler
' Purpose: Manages workbook closing procedures including cleanup operations,
' calculation mode restoration, and application termination for the calibration workbook.

Option Explicit

Private Sub Workbook_BeforeClose(Cancel As Boolean)
Application.DisplayFullScreen = False
Application.ScreenUpdating = False
Application.Run "C.xlsm!ClearC"
Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True
Workbooks("C.xlsm").Protect Password:="spike"
Workbooks("C.xlsm").Saved = True
Application.Quit
End Sub
