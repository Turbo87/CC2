' VBA Class: Sheet6 Event Handler
' Purpose: Handles worksheet activation and change events for Sheet6.
' Manages automatic calculations and cell change responses.

Private Sub Worksheet_Activate()

 'Application.ScreenUpdating=False
 Application.Calculate
 'Application.ScreenUpdating = True

 End Sub

Private Sub Worksheet_Change(ByVal Target As Range)

If Target.Address = "$G$12" Then

Application.EnableEvents = False
Application.ScreenUpdating = False
ActiveSheet.Unprotect Password:=("spike")
Range("K8").Value = Range("K10").Value
ActiveSheet.Protect Password:=("spike")
Application.ScreenUpdating = True
Application.EnableEvents = True
If Range("K8") = 1 Then
 MsgBox "Inconsistent with IGC file data. Click OK to continue"
 Range("G12:H12").ClearContents
End If
End If
End Sub
