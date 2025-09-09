' VBA Module: Waypoint Data Transfer and Processing
' Purpose: Handles waypoint data saving and transfer operations between workbooks.
' Manages text parsing, data formatting, and coordinate processing for saved waypoints.

Option Explicit
Sub WPSave()
Attribute WPSave.VB_ProcData.VB_Invoke_Func = " \n14"
'
' JLR 9/13/2011
'
Workbooks("A.xlsm").Activate
Sheets("Other").Select
Application.ScreenUpdating = False
If Range("D15") > 1 And Range("K15") = 2 And Range("N18") = 1 Then
    Workbooks("D.xlsm").Activate
    ActiveWindow.WindowState = xlMaximized
    Range("D4").Select
ElseIf Range("D15") > 1 And Range("K15") = 2 And Range("N18") = 2 Then
    Application.DisplayAlerts = False
    ActiveSheet.Unprotect Password:="spike"
    Workbooks("A.xlsm").Sheets("OTHER").Range("C69:C93").Value = Workbooks("D.xlsm").Sheets("SAVED Way Points").Range("B50:B74").Value
    
    If Range("D15") = 2 Then
    Application.DisplayAlerts = False
    Range("C69:C93").TextToColumns Destination:=Range("C69"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :=":", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, _
        1), Array(6, 9), Array(7, 1), Array(8, 1), Array(9, 1), Array(10, 1), Array(11, 1), Array(12 _
        , 9), Array(13, 1)), TrailingMinusNumbers:=True
    Range("I20:I28").ClearContents
    
    ElseIf Range("D15") = 3 Then
    Application.DisplayAlerts = False
    Range("C69:C93").TextToColumns Destination:=Range("C69"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :=":", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, _
        9), Array(6, 1), Array(7, 1), Array(8, 1), Array(9, 1), Array(10, 1), Array(11, 9), Array(12 _
        , 1), Array(13, 1)), TrailingMinusNumbers:=True
    Range("I20:I28").ClearContents
    End If

    ActiveSheet.Shapes("Oval 14").Visible = True
    ActiveSheet.Shapes("Oval 16").Visible = True
    ActiveSheet.Shapes("Oval 18").Visible = True
    ActiveSheet.Shapes("Oval 20").Visible = True
    ActiveSheet.Shapes("Oval 22").Visible = True
    ActiveSheet.Shapes("Rectangle 1").Visible = True
    ActiveSheet.Protect Password:="spike"
    Range("C20").Select
    Workbooks("A.xlsm").Activate
    ActiveWindow.WindowState = xlMaximized
    Application.DisplayFullScreen = True
    Application.DisplayAlerts = True
    ActiveWorkbook.Protect Password:="spike"
    Application.ScreenUpdating = True
End If
End Sub