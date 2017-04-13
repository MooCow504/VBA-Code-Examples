Option Compare Database
Option Explicit

'**************************************************************************************
'Delegate class needed to act as intermediary between custom controls and handler class
'    do to the way class with interfaces work in VBA
'**************************************************************************************

Public Event PassData(ctl As IControl, dataArr As Variant)
Public Event NextTab(ctl As IControl)
Public Event DisplayData()
Public Event ChangeState(ctl As IControl, columnList As Variant)
Public Event CancelCustom(ByRef cancel As Boolean)
Public Event RequeryForm(ctl As IControl)

Public Sub PassData(ctl As IControl, dataArr As Variant)
    RaiseEvent PassData(ctl, dataArr)
End Sub

Public Sub NextTab(ctl As IControl)
    RaiseEvent NextTab(ctl)
End Sub

Public Sub DisplayData()
    RaiseEvent DisplayData
End Sub

Public Sub ChangeState(ctl As IControl, columnList As Variant)
    RaiseEvent ChangeState(ctl, columnList)
End Sub

Public Function CancelCustom() As Boolean

Dim cancel As Boolean

RaiseEvent CancelCustom(cancel)
CancelCustom = cancel
    
End Function

Public Function RequeryForm(ctl As IControl)
    RaiseEvent RequeryForm(ctl)
End Function
