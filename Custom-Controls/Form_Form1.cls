Option Compare Database
Option Explicit

Private Handler As clsHandler

Private Sub Form_Close()
    Set Handler = Nothing
End Sub

Private Sub Form_Open(cancel As Integer)

Set Handler = New clsHandler
Handler.Initialize Me

End Sub
