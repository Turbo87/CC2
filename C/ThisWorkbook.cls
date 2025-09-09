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
'
Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)

Dim Password As String
Dim EnteredPassword As String

Password = "Spike'sadog2"
EnteredPassword = InputBox("Enter password")
If EnteredPassword <> Password Then Cancel = True
End Sub