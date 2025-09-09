' VBA Module: Waypoint Management and Data Operations
' Purpose: Manages waypoint addition, deletion, and list operations for flight planning.
' Handles waypoint saving, data sorting, and user interface controls for waypoint management.

Option Explicit

Sub ADDwp()
'
' JLR 9/7/2011
'
 Application.ScreenUpdating = False
 Sheets("Saved Way Points").Unprotect Password:="spike"
 If Range("D4") = 1 Then
    ActiveSheet.Shapes("Picture 3").Visible = True
    ActiveSheet.Shapes("Picture 6").Visible = False
    Range("D4").Select
ElseIf Range("D4") = 2 Then
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
Sheets("Saved Way Points").Protect Password:="spike"
Application.ScreenUpdating = True
End Sub
Sub AdDeleteWP()
'
' JLR 9/7/2011
'
Application.ScreenUpdating = False

If Range("N9") = "NO" Then
    ActiveSheet.Unprotect Password:="spike"
    Range("B16:L40").SpecialCells(xlCellTypeBlanks).Select
    Range("C" & ActiveCell.Row & ":L" & ActiveCell.Row).ClearContents
        If Range("B16") = "" Then
            Range("B16:L16").Value = Range("B40:L40").Value
            Range("B40:L40").ClearContents
        End If
    Range("B16:L40").Sort Key1:=Range("B16"), Order1:=xlDescending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    Range("B16:L40").Sort Key1:=Range("B16"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    
    Range("B50:B74").FormulaR1C1 = _
        "=CONCATENATE(R[-34]C,""::"",R[-34]C[1],"":"",R[-34]C[2],"":"",R[-34]C[3],"":"",R[-34]C[4],"":"",R[-34]C[5],""::"",R[-34]C[6],"":"",R[-34]C[7],"":"",R[-34]C[8],"":"",R[-34]C[9],"":"",R[-34]C[10])"
      Range("D4").FormulaR1C1 = "1"
      Range("A1:M22").Select
      ActiveWindow.Zoom = True
      ActiveSheet.Protect Password:="spike"
      ActiveWorkbook.Save
      Application.ScreenUpdating = True
    Range("D4").Select
Else
    Sheets("Saved Way Points").Unprotect Password:="spike"
    Range("B40:L40").Value = Range("B9:L9").Value
    Range("E8:L8").EntireColumn.Hidden = False
    If Range("E40") = "" And Range("B40") <> "" Then
        Range("E40").Formula = "=ROUND(RC[1]/60,3)"
        Range("E40").Value = Range("E40").Value
    ElseIf Range("F40") = "" And Range("B40") <> "" Then
        Range("F40").FormulaR1C1 = "=ROUND(RC[-1]*60,0)"
        Range("F40").Value = Range("F40").Value
    End If
      If Range("J40") = "" And Range("B40") <> "" Then
        Range("J40").Formula = "=ROUND(RC[1]/60,3)"
        Range("J40").Value = Range("J40").Value
      ElseIf Range("K40") = "" And Range("B40") <> "" Then
        Range("K40").FormulaR1C1 = "=ROUND(RC[-1]*60,0)"
        Range("K40").Value = Range("K40").Value
      End If
        If Range("D4") <= 2 Then
            Range("E1").EntireColumn.Hidden = False
            Range("F1").EntireColumn.Hidden = True
            Range("J1").EntireColumn.Hidden = False
            Range("K1").EntireColumn.Hidden = True
        ElseIf Range("D4") = 3 Then
            Range("E1").EntireColumn.Hidden = True
            Range("F1").EntireColumn.Hidden = False
            Range("J1").EntireColumn.Hidden = True
            Range("K1").EntireColumn.Hidden = False
        End If
        
    Range("B16:L40").SpecialCells(xlCellTypeBlanks).Select
    Range("C" & ActiveCell.Row & ":L" & ActiveCell.Row).ClearContents
    
        If Range("B16") = "" Then
            Range("B16:L16").Value = Range("B40:L40").Value
            Range("B40:L40").ClearContents
            Range("B9:L9").ClearContents
        End If
    Range("B16:L40").Sort Key1:=Range("B16"), Order1:=xlDescending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    Range("B16:L40").Sort Key1:=Range("B16"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
   
      Range("B50:B74").FormulaR1C1 = _
        "=CONCATENATE(R[-34]C,""::"",R[-34]C[1],"":"",R[-34]C[2],"":"",R[-34]C[3],"":"",R[-34]C[4],"":"",R[-34]C[5],""::"",R[-34]C[6],"":"",R[-34]C[7],"":"",R[-34]C[8],"":"",R[-34]C[9],"":"",R[-34]C[10])"
      Range("B9:L9").ClearContents
      Range("D4").FormulaR1C1 = "1"
      Range("A1:M22").Select
      ActiveWindow.Zoom = True
      Range("A1").Select
      ActiveSheet.Protect Password:="spike"
      ActiveWorkbook.Save
      Application.ScreenUpdating = True
End If
End Sub
Sub Rwritt()
'
' JLR 9/8/2011
'
Dim Visible As Variant, w As Workbook
    Application.ScreenUpdating = False
    
    On Error Resume Next
    Workbooks("A.xlsm").Activate
  If Err = 0 Then
    Workbooks("D.xlsm").Activate
    Application.Run "D.xlsm!WPSave"
    Sheets("Saved Way Points").Select
    ActiveWindow.ScrollRow = 1
    Range("D4:G4").Select
    Workbooks("A.xlsm").Activate
    Application.DisplayFullScreen = True
    Sheets("OTHER").Select
    ActiveSheet.Unprotect Password:="spike"
    If ActiveSheet.Shapes("Oval 14").Visible = True Then
    Range("K15") = 2
    Else: Range("K15") = 1
    End If
    Range("A1:N36").Select
    ActiveWindow.Zoom = True
    ActiveSheet.Protect Password:="spike"
    Range("K15:L15").Select
    Application.ScreenUpdating = True
  Else
    Sheets("Saved Way Points").Select
    ActiveWindow.ScrollRow = 1
    Range("D4:G4").Select
    For Each w In Application.Workbooks
    w.Save
Next w
Application.Quit
End If
End Sub