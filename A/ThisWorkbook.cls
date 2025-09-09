Option Explicit

'Private Declare Function InternetGetConnectedStateEx Lib "wininet.dll" (ByRef lpdwFlags As Long, ByVal lpszConnectionName As String, ByVal dwNameLen As Integer, ByVal dwReserved As Long) As Long
'Dim sConnType As String * 255

#If VBA7 And Win64 Then
    Private Declare PtrSafe Function InternetGetConnectedStateEx Lib "wininet.dll" (ByRef lpdwFlags As LongPtr, ByVal lpszConnectionName As String, ByVal dwNameLen As Integer, ByVal dwReserved As Long) As LongPtr
#Else
    Private Declare Function InternetGetConnectedStateEx Lib "wininet.dll" (ByRef lpdwFlags As Long, ByVal lpszConnectionName As String, ByVal dwNameLen As Integer, ByVal dwReserved As Long) As Long
#End If
Dim sConnType As String * 255

Private Sub Workbook_Open()

Application.ScreenUpdating = False
Workbooks("A.xlsm").Unprotect Password:="spike"
Sheets("All Claims").Unprotect Password:="spike"
Application.Cursor = xlDefault
ActiveWorkbook.UpdateLink Name:=ActiveWorkbook.LinkSources
Sheets("Parsed").Range("A39:A64").Clear
Sheets("Parsed").Range("C40:L64").FormulaR1C1 = "='[D.xlsm]Calibration'!R[15]C[-1]"
Sheets("Parsed").Range("C40:L64").Value = Sheets("Parsed").Range("C40:L64").Value
Application.Calculation = xlCalculationAutomatic
Application.Iteration = True
Application.MaxIterations = 100
Application.MaxChange = 0.001
Application.DisplayFullScreen = True
Application.MoveAfterReturn = True
Application.MoveAfterReturnDirection = xlToRight
Application.DisplayFormulaBar = False
Application.DisplayStatusBar = False
'Application.ScreenUpdating = False

#If VBA7 And Win64 Then
    Dim Ret As LongPtr
#Else
    Dim Ret As Long
#End If
Ret = InternetGetConnectedStateEx(Ret, sConnType, 254, 0)
If Ret = 1 Then
    Range("A2").QueryTable.Refresh BackgroundQuery:=False
End If

Application.ScreenUpdating = True
Application.DisplayFullScreen = True
Range("A1:H28").Select
ActiveWindow.Zoom = True
Range("A1").Select
'ActiveWorkbook.Protect Password:="spike"
Sheets("All Claims").Protect Password:="spike"
Workbooks("A.xlsm").Protect Password:="spike"
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
 On Error Resume Next
    Workbooks("C.xlsm").Activate
  If Err = 0 Then
    Application.DisplayFullScreen = True
  Else
    Application.DisplayFullScreen = False
  End If
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
Dim Password As String
Dim EnteredPassword As String

Password = "Spike'sadog2"
EnteredPassword = InputBox("Enter password")
If EnteredPassword <> Password Then Cancel = True
End Sub