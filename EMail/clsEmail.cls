Option Explicit

Public Sub EmailFile(Recipients As String, Subject As String, Body As String, _
        Optional Attachments As String, Optional CC As String)

Dim OutApp As Object, MailItem As Object
Dim Signature As String

Set OutApp = CreateObject("Outlook.Application")
Set MailItem = OutApp.CreateItem(0)

Signature = ReturnSignature
If Signature <> "" Then Body = Body & "<br><br>" & Signature

With MailItem
    .Subject = Subject
    .To = Recipients
    .HTMLBody = Body
    If Attachments <> "" Then
        .Attachments.Add Attachments
    End If
'    .Display
    .Send
End With

Set OutApp = Nothing
Set MailItem = Nothing

End Sub

Private Function ReturnSignature() As String

Dim path As String, extension As String
'Dim signatures() As String, tempSig As String
Dim sigFile As String

path = "C:\Users\" & Environ$("UserName") & "\AppData\Roaming\Microsoft\Signatures\"
extension = "*.htm"

sigFile = Dir$(path & extension, vbNormal)
If sigFile = "" Then
    Debug.Print "NO SIGNATURES"
    Exit Function
End If

'Needs form implementation in order to pick possible choices
''-------------------------------------------------------------------------
'Do While tempSig <> ""
'    If Not IsArrayAllocated(signatures) Then
'        ReDim signatures(0) As String
'    Else
'        ReDim Preserve signatures(UBound(signatures) + 1) As String
'    End If
'    signatures(UBound(signatures)) = tempSig
'    'Look for next file with matching extension in folder
'    tempSig = Dir
'Loop
''-------------------------------------------------------------------------

ReturnSignature = ReadFile(path & sigFile)

End Function
