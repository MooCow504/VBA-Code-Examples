Option Compare Database
Option Explicit

Implements IControl

Private Enum InputTypes
    Text = 2 ^ 1
    Numeric = 2 ^ 2
    Fraction = 2 ^ 3
End Enum

Private Delegate As clsDelegate

Private WithEvents TB As Access.TextBox
Private Const EVENTED As String = "[Event Procedure]"

Private Top As Long, Name As String, pThis As Access.Control, pInputType As InputTypes, ActiveState As ActiveStates 

Public Function IControl_Initialize(ctl As Access.Control, delegateObject As clsDelegate) As Access.Control

Set TB = ctl

TB.OnChange = EVENTED
TB.AfterUpdate = EVENTED

ActiveState = Custom
Top = ctl.Top
Name = ctl.Name
pInputType = Fraction
Set pThis = ctl
Set Delegate = delegateObject

Set IControl_Initialize = TB

End Function

Private Sub TB_Change()

If InputType = Fraction Then
    If InvalidFraction Then InvalidCharacterWarning
ElseIf InputType = Numeric Then
    
ElseIf InputType = Text Then

End If

End Sub

Private Function InvalidFraction() As Boolean
Dim i As Long

Dim inputText As String, char As String, prevChar As String
Dim period As Boolean, slash As Boolean, dash As Boolean

inputText = This.Text

If inputText <> "" Then
    char = Mid(inputText, Len(inputText))
    period = IsPeriod(char)
    slash = IsSlash(char)
    dash = IsDash(char)
    
    'If user enters a "SPACE" replace it with a dash
    If char = " " Then
        This.Text = Mid(inputText, 1, (Len(inputText) - 1)) & "-"
        char = "-"
        'Kepps cursor at end of input string
        SetCursorAtEnd
'        'SelStart fires another OnChange Event, exit sub to prevent running rest of sub twice
'        Exit Function
    End If
    
    'Check to make sure character entered is a valid character
    If Not IsNumeric(char) _
            And Not period _
            And Not dash _
            And Not slash Then
        InvalidFraction = True
    ElseIf Not IsNumeric(char) Then
        If Len(inputText) = 1 Then
            'If first character is not numeric
            InvalidFraction = True
        ElseIf Not IsNumeric(Mid(inputText, (Len(inputText) - 1), 1)) Then
            'If two non numeric characters are entered back to back
            InvalidFraction = True
        End If
        
        For i = 1 To (Len(inputText) - 1)
            prevChar = Mid(inputText, i, 1)
            
            'If non numeric characters are duplicated, mixed and matched, or in the wrong order
            If IsPeriod(prevChar) _
                    And (period _
                        Or slash _
                        Or dash) Then
                InvalidFraction = True
            ElseIf IsDash(prevChar) _
                    And (period _
                        Or dash) Then
                InvalidFraction = True
            ElseIf IsSlash(prevChar) _
                    And (period _
                        Or dash _
                        Or slash) Then
                InvalidFraction = True
            End If
        Next i
    End If
End If

End Function

Private Sub SetCursorAtEnd()

This.SetFocus
This.SelStart = Len(This.Text)

End Sub

Private Sub InvalidCharacterWarning()

MsgBox "Please enter a valid character.", vbExclamation
This.Text = Mid(This.Text, 1, (Len(This.Text) - 1))
SetCursorAtEnd

End Sub

Private Sub TB_AfterUpdate()

'Delegate.PassData Me, CreateDataArray
PassData
Delegate.NextTab Me

End Sub

Private Function PassData()
    Delegate.PassData Me, CreateDataArray
End Function

Private Function CreateDataArray() As Variant

Dim dataArr(1, 0) As Variant

dataArr(0, 0) = Mid(Name, 4)
dataArr(1, 0) = This.Value

CreateDataArray = dataArr

End Function

Public Function IControl_DefaultAvailable(Optional StandardState As ActiveStates = ActiveStates.Custom) As Boolean

SetOriginalState
'ActiveState = StandardState
If ActiveState = Standard Then
    LockDefault
    IControl_DefaultAvailable = True
ElseIf ActiveState = Assumed Then
    IControl_Enable
    IControl_DefaultAvailable = True
ElseIf ActiveState = Custom Then
    IControl_Enable
    IControl_DefaultAvailable = False
ElseIf ActiveState = Inactive Then
'    SetOriginalState
    Debug.Print "Inactive"
End If

End Function

Private Sub SetOriginalState()

This.Locked = False
This.Enabled = False
This.BackColor = RGB(255, 255, 255)

End Sub

Public Sub IControl_Reset()

If ActiveState <> Custom Then
    This.Value = ""
    SetOriginalState
End If

End Sub

Public Sub IControl_SetFocus()

This.Enabled = True
This.SetFocus

End Sub

Public Sub IControl_Hide()
    This.Visible = False
End Sub

Private Sub LockDefault()

This.Locked = True
This.BackColor = RGB(231, 230, 230)

End Sub

Public Sub IControl_Enable()

'ActiveState = Custom
This.Visible = True
This.Enabled = True
This.Locked = False
This.BackColor = RGB(255, 255, 255)

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

Private Property Get InputType() As InputTypes
    InputType = pInputType
End Property

Private Property Let InputType(flag As InputTypes)
    pInputType = flag
End Property
