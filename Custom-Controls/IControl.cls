Option Compare Database
Option Explicit

Public Enum ActiveStates
    'ComboBox and CommandButtum
    Active = 2 ^ 1
    Inactive = 2 ^ 2
    
    'TextBox
    Standard = 2 ^ 3
    Custom = 2 ^ 4
    Assumed = 2 ^ 5
End Enum

Public Function Initialize(ctl As Access.Control, delegateObject As clsDelegate) As Access.Control

End Function

Public Function DefaultAvailable(Optional StandardState As ActiveStates = ActiveStates.Custom) As Boolean

End Function

Public Sub SetFocus()

End Sub

Public Sub Hide()

End Sub

Public Sub Enable()

End Sub

Public Sub Reset()

End Sub

Public Property Get Top() As Long

End Property

Public Property Get Name() As String

End Property

Public Property Let TabIndex(index As Long)

End Property

Public Property Get ActiveState() As ActiveStates

End Property

