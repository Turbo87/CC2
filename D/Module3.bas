Option Explicit
Sub SVenter()
Attribute SVenter.VB_ProcData.VB_Invoke_Func = " \n14"
'
' JLR 8/2/2012
'
Application.ScreenUpdating = False
Workbooks("A.xlsm").Activate
Sheets("Other").Activate
ActiveSheet.Unprotect Password:="spike"
Range("C69:C93").NumberFormat = "@"

 If Range("C" & ActiveCell.Row) = Range("C69") And Range("C69") <> "" Then
    Range("E" & ActiveCell.Row & ":M" & ActiveCell.Row).Formula = "=R69C"
 ElseIf Range("C" & ActiveCell.Row) = Range("C70") And Range("C70") <> "" Then
    Range("E" & ActiveCell.Row & ":M" & ActiveCell.Row).Formula = "=R70C"
 ElseIf Range("C" & ActiveCell.Row) = Range("C71") And Range("C71") <> "" Then
    Range("E" & ActiveCell.Row & ":M" & ActiveCell.Row).Formula = "=R71C"
 ElseIf Range("C" & ActiveCell.Row) = Range("C72") And Range("C72") <> "" Then
    Range("E" & ActiveCell.Row & ":M" & ActiveCell.Row).Formula = "=R72C"
 ElseIf Range("C" & ActiveCell.Row) = Range("C73") And Range("C73") <> "" Then
    Range("E" & ActiveCell.Row & ":M" & ActiveCell.Row).Formula = "=R73C"
 ElseIf Range("C" & ActiveCell.Row) = Range("C74") And Range("C74") <> "" Then
    Range("E" & ActiveCell.Row & ":M" & ActiveCell.Row).Formula = "=R74C"
 ElseIf Range("C" & ActiveCell.Row) = Range("C75") And Range("C75") <> "" Then
    Range("E" & ActiveCell.Row & ":M" & ActiveCell.Row).Formula = "=R75C"
 ElseIf Range("C" & ActiveCell.Row) = Range("C76") And Range("C76") <> "" Then
    Range("E" & ActiveCell.Row & ":M" & ActiveCell.Row).Formula = "=R76C"
 ElseIf Range("C" & ActiveCell.Row) = Range("C77") And Range("C77") <> "" Then
    Range("E" & ActiveCell.Row & ":M" & ActiveCell.Row).Formula = "=R77C"
 ElseIf Range("C" & ActiveCell.Row) = Range("C78") And Range("C78") <> "" Then
    Range("E" & ActiveCell.Row & ":M" & ActiveCell.Row).Formula = "=R78C"
 ElseIf Range("C" & ActiveCell.Row) = Range("C79") And Range("C79") <> "" Then
    Range("E" & ActiveCell.Row & ":M" & ActiveCell.Row).Formula = "=R79C"
 ElseIf Range("C" & ActiveCell.Row) = Range("C80") And Range("C80") <> "" Then
    Range("E" & ActiveCell.Row & ":M" & ActiveCell.Row).Formula = "=R80C"
 ElseIf Range("C" & ActiveCell.Row) = Range("C81") And Range("C81") <> "" Then
    Range("E" & ActiveCell.Row & ":M" & ActiveCell.Row).Formula = "=R81C"
 ElseIf Range("C" & ActiveCell.Row) = Range("C82") And Range("C82") <> "" Then
    Range("E" & ActiveCell.Row & ":M" & ActiveCell.Row).Formula = "=R82C"
 ElseIf Range("C" & ActiveCell.Row) = Range("C83") And Range("C83") <> "" Then
    Range("E" & ActiveCell.Row & ":M" & ActiveCell.Row).Formula = "=R83C"
 ElseIf Range("C" & ActiveCell.Row) = Range("C84") And Range("C84") <> "" Then
    Range("E" & ActiveCell.Row & ":M" & ActiveCell.Row).Formula = "=R84C"
 ElseIf Range("C" & ActiveCell.Row) = Range("C85") And Range("C85") <> "" Then
    Range("E" & ActiveCell.Row & ":M" & ActiveCell.Row).Formula = "=R85C"
 ElseIf Range("C" & ActiveCell.Row) = Range("C86") And Range("C86") <> "" Then
    Range("E" & ActiveCell.Row & ":M" & ActiveCell.Row).Formula = "=R86C"
 ElseIf Range("C" & ActiveCell.Row) = Range("C87") And Range("C87") <> "" Then
    Range("E" & ActiveCell.Row & ":M" & ActiveCell.Row).Formula = "=R87C"
 ElseIf Range("C" & ActiveCell.Row) = Range("C88") And Range("C88") <> "" Then
    Range("E" & ActiveCell.Row & ":M" & ActiveCell.Row).Formula = "=R88C"
 ElseIf Range("C" & ActiveCell.Row) = Range("C89") And Range("C89") <> "" Then
    Range("E" & ActiveCell.Row & ":M" & ActiveCell.Row).Formula = "=R89C"
 ElseIf Range("C" & ActiveCell.Row) = Range("C90") And Range("C90") <> "" Then
    Range("E" & ActiveCell.Row & ":M" & ActiveCell.Row).Formula = "=R90C"
 ElseIf Range("C" & ActiveCell.Row) = Range("C91") And Range("C91") <> "" Then
    Range("E" & ActiveCell.Row & ":M" & ActiveCell.Row).Formula = "=R91C"
 ElseIf Range("C" & ActiveCell.Row) = Range("C92") And Range("C92") <> "" Then
    Range("E" & ActiveCell.Row & ":M" & ActiveCell.Row).Formula = "=R92C"
 ElseIf Range("C" & ActiveCell.Row) = Range("C93") And Range("C93") <> "" Then
    Range("E" & ActiveCell.Row & ":M" & ActiveCell.Row).Formula = "=R93C"
 End If
 Range("I20:I28").ClearContents
ActiveSheet.Protect Password:="spike"
Application.ScreenUpdating = True
End Sub