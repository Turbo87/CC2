' VBA Module: Calendar and Date Management
' Purpose: Handles calendar date selection and dropdown list management.
' Controls date picker functionality and manages date range selections.

Option Explicit
Sub CalDate()
'
'JLR 6/15/2011
'
Application.ScreenUpdating = False
ActiveSheet.Unprotect Password:="spike"
If Range("F10") = 1 Then
ActiveSheet.Shapes.Range(Array("Drop Down 142")).Select
    With Selection
        .ListFillRange = "$P$43:$P$54"
        .LinkedCell = "$D$12"
        .DropDownLines = 12
        .Display3DShading = False
    End With
Range("D12") = 1
Range("F10").Select

ElseIf Range("F10") = 2 Then
 ActiveSheet.Shapes.Range(Array("Drop Down 142")).Select
    With Selection
        .ListFillRange = "$P$43:$P$44"
        .LinkedCell = "$D$12"
        .DropDownLines = 2
        .Display3DShading = False
    End With
Range("D12") = 2
Range("F14").Select

ElseIf Range("F10") > 2 And Range("D12") > 2 Then
ActiveSheet.Shapes.Range(Array("Drop Down 142")).Select
    With Selection
        .ListFillRange = "$P$43:$P$54"
        .LinkedCell = "$D$12"
        .DropDownLines = 12
        .Display3DShading = False
    End With
Range("F14").Select

ElseIf Range("F10") <> 2 Then
ActiveSheet.Shapes.Range(Array("Drop Down 142")).Select
    With Selection
        .ListFillRange = "$P$43:$P$54"
        .LinkedCell = "$D$12"
        .DropDownLines = 12
        .Display3DShading = False
    End With
Range("D12") = 1
Range("D12").Select
End If
ActiveSheet.Protect Password:="spike"
Application.ScreenUpdating = True
End Sub
Sub Macro3()
'
'JLR 6/1/2011
'
Application.ScreenUpdating = False
ActiveWorkbook.Unprotect Password:="spike"

If Range("F18") = 3 Then
     Sheets("OTHER").Visible = True
     Sheets("OTHER").Select
     ActiveSheet.Unprotect Password:="spike"
     ActiveSheet.Shapes("Drop Down 38").Visible = True
     ActiveSheet.Shapes("Drop Down 5").Locked = False
     Range("D4:E4").FormulaR1C1 = "1"
     Range("D4:E4").Locked = False
     Range("D4:E4").FormulaHidden = True
     Range("A12:A28").Select
     Selection.EntireRow.Hidden = False
     Range("D6:E6,J6,D8:F8,J8:L8,D10:G10,K10:L10").ClearContents
     Range("A1").RowHeight = 12
     Range("A1").ColumnWidth = 21
     Range("A1:O36").Select
     ActiveWindow.Zoom = True
     Range("D4").Select
     ActiveSheet.Calculate
     ActiveSheet.Protect Password:="spike"
ElseIf Range("F18") < 3 And Range("C23") <> "        Click a tab below to continue" Then
     Sheets("OTHER").Visible = False
ElseIf Range("F18") = 4 Then
     Application.Run "A.xlsm!Custom"
ElseIf Range("F18") = 5 Then
     Sheets("Other").Visible = True
     Sheets("Other").Select
     ActiveSheet.Unprotect Password:="spike"
     Range("D4:E4").FormulaR1C1 = "1"
     Range("D8:F8,D10:G10,K10:L10").ClearContents
     Sheets("Other").Visible = False
End If
ActiveWorkbook.Protect Password:="spike"
Application.ScreenUpdating = True
End Sub
Sub ChgSolo()
'
' JLR 7/1/2011
'
Application.ScreenUpdating = False
ActiveWorkbook.Unprotect Password:="spike"
Sheets("Other").Visible = True
Sheets("Other").Select
    
If Range("M4") = 2 Then
       ActiveSheet.Unprotect Password:="spike"
       Range("J8:L8").Locked = False
       ActiveSheet.Protect Password:="spike"
ElseIf Range("M4") = 1 Then
       ActiveSheet.Unprotect Password:="spike"
       Range("J8:L8").Locked = True
       ActiveSheet.Protect Password:="spike"
End If

Sheets("All Claims").Select
Range("F10").Select
If Range("C23") = "" Then
Sheets("Other").Visible = False
End If
ActiveWorkbook.Protect Password:="spike"
Application.ScreenUpdating = True

End Sub
Sub Custom()
'
' JLR 6/10/2011
'
    Application.ScreenUpdating = False
    ActiveWorkbook.Unprotect Password:="spike"
    Sheets("Other").Visible = True
    Sheets("Other").Select
    ActiveSheet.Unprotect Password:="spike"
    ActiveSheet.Shapes("Drop Down 38").Visible = False
    ActiveSheet.Shapes("Drop Down 45").Visible = False
    ActiveSheet.Shapes("Oval 14").Visible = False
    ActiveSheet.Shapes("Oval 16").Visible = False
    ActiveSheet.Shapes("Oval 18").Visible = False
    ActiveSheet.Shapes("Oval 20").Visible = False
    ActiveSheet.Shapes("Oval 22").Visible = False
    Range("A12:A28").EntireRow.Hidden = True
    Range("D6:E6").ClearContents
    Range("D6:E6").Locked = True
    Range("J6").ClearContents
    Range("J6").Locked = True
    ActiveSheet.Range("D4").FormulaR1C1 = "4"
    Range("D4:E4").Locked = True
    Range("D4:E4").FormulaHidden = True
    ActiveSheet.Unprotect Password:="spike"
    ActiveSheet.Shapes("Drop Down 5").Locked = True
    ActiveSheet.Protect Password:="spike"

    ActiveSheet.Unprotect Password:="spike"
    Range("D8:F8,J8:L8,D10:G10,K10:L10").Locked = False
    Range("D8:F8,J8:L8,D10:G10,K10:L10").FormulaHidden = True
    Range("A1").RowHeight = 60
    Range("A1").ColumnWidth = 10
    Range("A1:N43").Select
    ActiveWindow.Zoom = True
    ActiveSheet.Protect Password:="spike"
    Range("D8:F8").Select
    ActiveWorkbook.Protect Password:="spike"
    Application.ScreenUpdating = True
End Sub
Sub DecAltDist()
'
' JLR 6/11/2011//Revised 7/22/2015 for written Free task check using DD:MM.mmm only; revised 1/19/2016 so DD45 visible only
' once coordinate format is selected via Sub Coord
'
 Application.ScreenUpdating = False
 ActiveSheet.Unprotect Password:="spike"
 
 If Range("B2") = "Written Declaration" And Range("D4") >= 1 And Range("D4") <> 4 Then
        Range("D6:E6,J6,D8:F8,J8:L8,D10:G10,K10:L10").Locked = False
        ActiveSheet.Shapes("Drop Down 45").Visible = False
        Range("K15:L15").Locked = False
        Range("D15:E15").Locked = False
        Range("D6:E6").Select
 
 ElseIf Range("B2") = "Written Declaration" And Range("D4") = 4 Then
    Range("D6:E6,J6,D8:F8,J8:L8,D10:G10,K10:L10").ClearContents
    Range("D6:E6,J6,D8:F8,J8:L8,D10:G10,K10:L10,K15:L15").Locked = True
    Range("D15").Value = 2
    Range("D15:E15").Locked = True
    ActiveSheet.Shapes("Drop Down 45").Visible = False
    Range("K15:L15").Locked = True
    
 ElseIf Range("B2") <> "Written Declaration" And Range("D4") = 4 Then
        Range("D6:E6,J6").Locked = True
        Range("D8:F8,J8:L8,D10:G10,K10:L10").Locked = False
End If
ActiveSheet.Protect Password:="spike"
Application.ScreenUpdating = True
End Sub
Sub Coord()
'
' JL Ruprecht 9/22/2011
'
Application.ScreenUpdating = False
Sheets("OTHER").Select
ActiveSheet.Unprotect Password:="spike"
If Range("D15") = 1 And Range("C69") = "" Then
ActiveSheet.Shapes("Drop Down 45").Visible = False
Range("K15:L15").Locked = True
Range("D15").Select
Else:
ActiveSheet.Shapes("Drop Down 45").Visible = True
Range("K15:L15").Locked = False
End If
If Range("D15") > 1 And Range("C69") <> "" Then
 Workbooks("A.xlsm").Sheets("Other").Range("C69:C93").Value = Workbooks("D.xlsm").Sheets("Saved Way Points").Range("B50:B74").Value
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
Range("C20").Select
End If
ActiveSheet.Protect Password:="spike"
Application.ScreenUpdating = True
End Sub
Sub ConfirmWrit()
'
' JLR 6/15/2011
'
If Range("B31") = "Done? Click on the Glider to Confirm!" Then
    Sheets("ALL CLAIMS").Select
    Range("F6").Select
End If
End Sub