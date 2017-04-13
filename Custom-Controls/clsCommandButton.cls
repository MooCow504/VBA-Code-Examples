Option Compare Database
Option Explicit

Implements IControl

Private Enum udtCursorNames
    IDC_HELP = 32651
    IDC_WAIT = 32514
    IDC_ARROW = 32512
    IDC_HAND = 32649
    IDC_APPSTARTING = 32650
    IDC_CROSS = 32515
    IDC_IBEAM = 32513
    IDC_NO = 32648
    IDC_SIZEALL = 32646
    IDC_SIZENESW = 32643
    IDC_SIZENS = 32645
    IDC_SIZENWSE = 32642
    IDC_SIZEWE = 32644
    IDC_UP = 32516
End Enum

Private Declare Function SetCursor Lib "user32.dll" (ByVal hCursor As Long) As Long
Private Declare Function LoadCursor Lib "user32.dll" Alias "LoadCursorA" _
        (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long

Private Delegate As clsDelegate

Private WithEvents CB As Access.CommandButton
Private Const EVENTED As String = "[Event Procedure]"

Private Top As Long, Name As String, pThis As Access.Control, ActiveState As ActiveStates

Public Function IControl_Initialize(ctl As Access.Control, delegateObject As clsDelegate) As Access.Control

Set CB = ctl

CB.OnClick = EVENTED
CB.OnMouseMove = EVENTED

Top = ctl.Top
Name = ctl.Name
Set pThis = ctl
Set Delegate = delegateObject

Set IControl_Initialize = CB

End Function

Private Sub CB_Click()

'Raise Event to Delegate

End Sub

Private Sub CB_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

Dim cursor As Long
cursor = DisplayCustomCursor(IDC_HAND)

End Sub

Private Function DisplayCustomCursor(Optional udtCursorName As udtCursorNames = udtCursorNames.IDC_HAND) As Long
      
Dim cursorHandle As Long

cursorHandle = LoadCursor(0, udtCursorName)
DisplayCustomCursor = SetCursor(cursorHandle)

End Function

Public Function IControl_DefaultAvailable(Optional StandardState As ActiveStates = ActiveStates.Custom) As Boolean

'In case Button is above an Input Box
'Will be skipped by Delegate.NextTab
IControl_DefaultAvailable = True

End Function

Public Sub IControl_SetFocus()
    This.SetFocus
End Sub

Public Sub IControl_Hide()
    This.Visible = False
End Sub

Public Sub IControl_Reset()

End Sub

Public Sub IControl_Enable()

End Sub

Public Property Get IControl_Top() As Long
    IControl_Top = Top
End Property

Public Property Get IControl_Name() As String
    IControl_Name = Name
End Property

Public Property Let IControl_TabIndex(index As Long)
    IControl_This.TabIndex = index
End Property

Public Property Get IControl_ActiveState() As ActiveStates
    IControl_ActiveState = ActiveState
End Property

Private Property Get This() As Access.Control
    Set This = pThis
End Property
