Option Compare Database
Option Explicit

Implements IControl

Private Enum States
    Unselected
    Selected
End Enum

Private Delegate As clsDelegate

Private WithEvents CBO As Access.ComboBox
Private Const EVENTED As String = "[Event Procedure]"

Private Top As Long, Name As String, pThis As Access.Control, pState As States, ActiveState As ActiveStates

Public Function IControl_Initialize(ctl As Access.Control, delegateObject As clsDelegate) As Access.Control

Set CBO = ctl

CBO.AfterUpdate = EVENTED
CBO.OnDblClick = EVENTED
CBO.OnGotFocus = EVENTED

State = Unselected
Top = ctl.Top
Name = ctl.Name
ActiveState = Inactive
Set pThis = ctl
Set Delegate = delegateObject

Set IControl_Initialize = CBO

End Function

Private Sub CBO_DblClick(cancel As Integer)

If Delegate.CancelCustom Then Exit Sub
Delegate.ChangeState Me, CreateColumnList

End Sub

Private Sub CBO_AfterUpdate()

State = Selected
Delegate.RequeryForm Me
PassData
Delegate.NextTab Me

End Sub

Private Sub CBO_GotFocus()

ActiveState = Active
This.Requery

End Sub

Private Function PassData()
    Delegate.PassData Me, CreateDataArray
End Function

Private Function CreateDataArray() As Variant
Dim i As Long

Dim dataArr() As Variant, columnList As Variant
columnList = CreateColumnList

ReDim dataArr(1, UBound(columnList)) As Variant
For i = LBound(dataArr, 2) To UBound(dataArr, 2)
    dataArr(0, i) = columnList(i)
    dataArr(1, i) = This.Column(i)
Next i

CreateDataArray = dataArr

End Function

Private Function CreateColumnList() As Variant
Dim i As Long

Dim colArr As Variant
Dim columnList As String

columnList = This.RowSource
columnList = Replace(columnList, "select ", "")

If InStr(columnList, " from") Then
    columnList = Left(columnList, InStr(columnList, "from") - 2)
Else
    columnList = Left(columnList, InStr(columnList, "from") - 1)
End If
    
columnList = Replace(columnList, "[", "")
columnList = Replace(columnList, "]", "")
colArr = Split(columnList, ", ")

For i = LBound(colArr) To UBound(colArr)
    colArr(i) = Mid(colArr(i), InStr(colArr(i), ".") + 1)
    'If 'as' Find Alias
    If InStr(colArr(i), " as ") Then
        colArr(i) = Mid(colArr(i), InStr(colArr(i), " as ") + 4)
    End If
    'Remove Line Breaks from string going on to long
    If Right(colArr(i), 2) = vbCrLf _
            Or Right(colArr(i), 2) = vbNewLine Then
        colArr(i) = Left(colArr(i), Len(colArr(i)) - 2)
    End If
Next i

CreateColumnList = colArr

End Function

Public Function IControl_DefaultAvailable(Optional StandardState As ActiveStates = ActiveStates.Custom) As Boolean
Dim i As Long

If State = Unselected Then
    'If Combobox has not been selected yet
    With This
        .Enabled = True
        .Requery
        
        'If only one option available
        If .ListCount = 1 Then
            .Value = .ItemData(0)
            LockDefault
        Else
            'Set 0 To (.ListCount - 1) because ListCount is 'Base 1' and Column reference is 'Base 0'
            For i = 0 To (.ListCount - 1)
                'Will Pick value based if Default value available
                If InStr(.Column(1, i), "Standard") _
                        Or .Column(1, i) = "Default" _
                        Or .Column(1, i) = "Normal" Then
                    .Value = .ItemData(i)
                End If
                        
            Next i
        End If
    End With
Else
    'If Combobox has been Selected and is being requeried
    Dim currentValue As Variant
    
    With This
        currentValue = .Column(1)
        .Value = Null
        State = Unselected
        .Requery
        
        'Check to see if ComboBox original choice is availabled
        For i = 0 To (.ListCount - 1)
            If .Column(1, i) = currentValue Then
                .Value = .ItemData(i)
            End If
        Next i
    End With
End If

IControl_DefaultAvailable = Not IsNull(This.Value)
    
If IControl_DefaultAvailable Then
    PassData
    State = Selected
    ActiveState = Active
Else
    SetOriginalState
    ActiveState = Inactive
End If

End Function

Private Sub LockDefault()

This.Locked = True
This.BackColor = RGB(231, 230, 230)

End Sub

Private Sub SetOriginalState()

This.Locked = False
This.Enabled = False
This.BackColor = RGB(255, 255, 255)

End Sub

Public Sub IControl_Reset()

With This
    .Value = Null
    .Requery
End With
SetOriginalState

End Sub

Public Sub IControl_SetFocus()

This.Enabled = True
This.SetFocus

End Sub

Public Sub IControl_Hide()
    This.Visible = False
End Sub

Public Sub IControl_Enable()

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

Private Property Get State() As States
    State = pState
End Property

Private Property Let State(flag As States)
    pState = flag
End Property
