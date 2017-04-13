Option Explicit

'--------------------------------------------------------------------
'Known bugs:
'Crashes sometimes with a large number of mail items
'--------------------------------------------------------------------

Sub SaveScanAttachments()
Dim i As Integer

Dim OlApp As Outlook.Application, OlMsg As Outlook.MailItem, _
        OlAttachment As Outlook.Attachments, OlSelection As Outlook.Selection
Dim XlApp As Excel.Application, FdFolder As Office.FileDialog
Dim file As String, folderPath As String, strPosition As Integer, folderName As String
Dim attachCount As Integer, numberEmails As Integer

'Lets user pick folder to save files in
Set XlApp = New Excel.Application
XlApp.Visible = False
Set FdFolder = XlApp.Application.FileDialog(msoFileDialogFolderPicker)
FdFolder.Show
folderPath = FdFolder.SelectedItems(1)
XlApp.Quit
Set XlApp = Nothing
Set FdFolder = Nothing

'Seperates name of folder chosen
For i = 1 To Len(folderPath)
    If Mid(folderPath, i, 1) = "\" Then
        strPosition = i
    End If
Next i
folderName = Right(folderPath, (Len(folderPath) - strPosition))

Set OlApp = CreateObject("Outlook.Application")
Set OlSelection = OlApp.ActiveExplorer.Selection
numberEmails = OlSelection.Count

If OlSelection.Count = 0 Then
    MsgBox ("Please select an e-mail.")
    Exit Sub
End If

'Loops through each message in selection
For Each OlMsg In OlSelection
    Set OlAttachment = OlMsg.Attachments
    
    'Saves the corresponding attachment
    'If OlAttachment.Count = 1 Then
    If OlAttachment.Count > 0 Then
        file = "\" + folderName + " pt " + CStr(numberEmails) + ".pdf"
        file = folderPath + file
        OlAttachment.Item(1).SaveAsFile file
        numberEmails = numberEmails - 1
        
    'Error message boxes
    ElseIf OlAttachment.Count = 0 Then
        MsgBox "One of the selected e-mails doesn't contain an attachment."
    End If
    
Next OlMsg

Set OlApp = Nothing
Set OlSelection = Nothing
Set OlAttachment = Nothing
Set OlMsg = Nothing

'Displays when macro finished running
MsgBox "Done."

End Sub
