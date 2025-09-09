' VBA Class: Workbook Event Handler
' Purpose: Handles workbook initialization for flight analysis application.
' Configures Excel settings for iterative calculations and disables AutoComplete.

Option Explicit
Private Sub Workbook_Open()

Application.EnableAutoComplete = False
Application.MaxIterations = 50
Application.MaxChange = 0.001
End Sub
