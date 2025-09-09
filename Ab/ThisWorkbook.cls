' VBA Class: Workbook Event Handler
' Purpose: Handles workbook-level events and initialization settings.
' Configures Excel application settings including AutoComplete and iteration parameters.

Option Explicit
Private Sub Workbook_Open()

Application.EnableAutoComplete = False
Application.MaxIterations = 50
Application.MaxChange = 0.001
End Sub
