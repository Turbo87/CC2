Option Explicit

Sub OpenD()
'
' JLR 9/13/2011
'
Dim linkArray As Variant
Sheets("OTHER").Activate
If Range("D15") = 1 Then
    Range("D15").Select
ElseIf Range("D15") > 1 And Range("K15") = 1 Then
    Range("K15").Select
ElseIf Range("D15") > 1 And Range("K15") <> 2 Then
    Application.ScreenUpdating = False
    ActiveWorkbook.Unprotect Password:="spike"
    Sheets("OTHER").Unprotect Password:="spike"
    Sheets("OTHER").Shapes("Oval 14").Visible = False
    Sheets("OTHER").Shapes("Oval 16").Visible = False
    Sheets("OTHER").Shapes("Oval 18").Visible = False
    Sheets("OTHER").Shapes("Oval 20").Visible = False
    Sheets("OTHER").Shapes("Oval 22").Visible = False
    Sheets("OTHER").Shapes("Rectangle 1").Visible = False
    Range("C69:O93").Clear
    Sheets("OTHER").Protect Password:="spike"
    Range("C20").Select
    ActiveWorkbook.Protect Password:="spike"
    Application.ScreenUpdating = True
ElseIf Range("D15") > 1 And Range("K15") = 2 Then
Application.ScreenUpdating = False
    On Error Resume Next
    Workbooks("D.xlsm").Activate
  If Err = 0 Then
    Workbooks("D.xlsm").WindowState = xlMinimized
    Sheets("Saved Way Points").Select
  Else
    Application.DisplayAlerts = False
    'linkArray = ActiveWorkbook.LinkSources(xlExcelLinks)
    'ActiveWorkbook.OpenLinks linkArray(1)
    Workbooks.Open Filename:=ThisWorkbook.Path & "\D.xlsm"
    Workbooks("D.xlsm").Activate
    Workbooks("D.xlsm").WindowState = xlMinimized
  End If
    Workbooks("A.xlsm").Activate
    Workbooks("A.xlsm").WindowState = xlMaximized
    Application.ScreenUpdating = True
    Application.DisplayFullScreen = True
    Application.Run "D.xlsm!WPSave"
End If
End Sub
Sub SeeList()
'
' JL Ruprecht 9/21/2011
'
Dim linkArray As Variant
Application.ScreenUpdating = False
    On Error Resume Next
    Workbooks("D.xlsm").Activate
  If Err = 0 Then
    Workbooks("D.xlsm").WindowState = xlMaximized
    Application.DisplayFullScreen = True
    Sheets("Saved Way Points").Select
    Application.Run "D.xlsm!List"
    Range("D4").Select
  Else
    Application.DisplayAlerts = False
    'linkArray = ActiveWorkbook.LinkSources(xlExcelLinks)
    'ActiveWorkbook.OpenLinks linkArray(1)
    Workbooks.Open Filename:=ThisWorkbook.Path & "\D.xlsm"
    Workbooks("D.xlsm").Activate
    ActiveWorkbook.WindowState = xlMaximized
    Application.DisplayFullScreen = True
    Application.Run "D.xlsm!List"
    Range("D4").Select
  End If
Application.ScreenUpdating = True
End Sub
