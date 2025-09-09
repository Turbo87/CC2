Option Explicit
Public Function GetDesktop() As String
    GetDesktop = CreateObject("WScript.Shell").SpecialFolders("Desktop") & _
        Application.PathSeparator
End Function
Sub PrintThis()
'Amended 10/2/2017 to send 'Claim Check' page to printer
    
If Range("G16") = 2 Then
   
    Application.Run ("C.xlsm!ElecCopy")

ElseIf Range("G16") = 3 Then
    
    On Error Resume Next
        
        If Range("A16") <> "OPTIM" Then
            Sheets(Array("PRINT THIS!", "CLAIM CHECK")).PrintOut Copies:=1, Collate:=True, IgnorePrintAreas:=False
        ElseIf Range("A16") = "OPTIM" Then
           Sheets(Array("PRINT THIS!", "CLAIM CHECK", "FREE ME")).PrintOut Copies:=1, Collate:=True, IgnorePrintAreas:=False
        End If
    
    If Err = 0 Then
        Application.ScreenUpdating = False
        Sheets("VERIFY TASK").Visible = True
        Sheets("CLAIM CHECK").Visible = False
        Sheets("PRINT THIS!").Visible = False
        ActiveWorkbook.Protect Password:="spike"
        ActiveWorkbook.Saved = True
        Application.ScreenUpdating = True
        Application.Quit
    Else: MsgBox "No printer installed. Click OK, take screen shots then click end/exit"
    End If
End If
End Sub
Sub ElecCopy()
'Amended 8/8/2018 for Deutsche version ' TEST

    Dim Msg, Style, Response
    Dim FName As Variant
    Dim DTAddress As String
    ActiveWorkbook.Unprotect Password:="spike"
    Sheets("Verify Task").Visible = False
    Sheets("Calibration").Visible = False
    
    DTAddress = GetDesktop
    ChDir DTAddress
    'FName = Application.GetSaveAsFilename(FileFilter:="PDF files, *.pdf", Title:="Export to PDF")
    FName = Application.GetSaveAsFilename(FileFilter:="PDF files, *.pdf")
    
    On Error Resume Next
    If FName <> False Then
        ActiveWorkbook.ExportAsFixedFormat Type:=xlTypePDF, Filename:=FName _
            , Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
            :=False, OpenAfterPublish:=True

        If Err = 0 Then
        Application.ScreenUpdating = False
        Sheets("Verify Task").Visible = True
        Sheets("Claim Check").Visible = False
        Sheets("Print This!").Visible = False
        ActiveWorkbook.Protect Password:="spike"
        ActiveWorkbook.Saved = True
        Application.Quit
        
        ElseIf Err <> 0 Then
        
        DTAddress = GetDesktop
        ChDir DTAddress
    
    FName = Application.GetSaveAsFilename(FileFilter:="XPS files, *.xps", Title:="Export to XPS")
    
    On Error Resume Next
    If FName <> False Then
        ActiveWorkbook.ExportAsFixedFormat Type:=xlTypeXPS, Filename:=FName _
            , Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
            :=False, OpenAfterPublish:=True

        If Err = 0 Then
        Application.ScreenUpdating = False
        Sheets("Verify Task").Visible = True
        Sheets("Claim Check").Visible = False
        Sheets("Print This!").Visible = False
        ActiveWorkbook.Protect Password:="spike"
        ActiveWorkbook.Saved = True
        Application.Quit
            
        ElseIf Err <> 0 Then
    
        Msg = "ERROR! Neither a PDF nor an XPS document can be created." & vbNewLine & "Do you want to save results as a Word document?"
        Style = vbYesNo + vbCritical + vbDefaultButton1
        Response = MsgBox(Msg, Style)
            If Response = vbYes Then
                Application.Run "C.xlsm!SavDoc"
            ElseIf Response = vbNo Then
                Msg = "Click OK to exit then select 'Send to Printer' as the save method"
                Style = vbOKOnly
                Range("G16").Value = 1
             End If
           End If
         End If
       End If
     End If
End Sub

Sub SavDoc()
'
' SavDoc Macro
'
Dim wdApp As Object
Dim wdDoc As Object

Application.ScreenUpdating = False
ActiveSheet.Unprotect Password:="spike"
Range("Print_Area").Select
Selection.CopyPicture Appearance:=xlScreen, Format:=xlPicture
ActiveSheet.Protect Password:="spike"
  
   On Error Resume Next
   Set wdApp = GetObject(, "Word.Application")
   If Err <> 0 Then Set wdApp = CreateObject("Word.Application")

 Set wdDoc = wdApp.Documents.Add
 wdApp.Visible = True
 wdDoc.ActiveWindow.Selection.Paste
    
 Sheets("Claim Check").Activate
 Range("Print_Area").Select
 Selection.CopyPicture Appearance:=xlScreen, Format:=xlPicture
 Sheets("Claim Check").Protect Password:="spike"
 ActiveWindow.WindowState = xlMinimized
 
 wdDoc.ActiveWindow.Selection.Paste
 wdDoc.ActiveWindow.LargeScroll Up:=6
 wdDoc.WindowState = xlMaximized
 Application.ScreenUpdating = True
 
 Set wdDoc = Nothing
 Set wdApp = Nothing
 wdApp.Quit

   
        ActiveWorkbook.Saved = True
        Application.Quit

End Sub