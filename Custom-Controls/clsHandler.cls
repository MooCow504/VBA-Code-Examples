Option Compare Database
Option Explicit

'Event Delegate
Private WithEvents Delegate As clsDelegate
'Data Object
Private dataCol As clsDataCol
'View
Private view As Access.Form
'Custom Controls
Private Control As IControl, InputArray() As IControl

Private Sub Class_Terminate()

'If InputArray not erased memory issues cause Access to crash
Erase InputArray
Set Control = Nothing
Set dataCol = Nothing

End Sub

Public Sub Initialize(frm As Access.Form)
Dim i As Long

Dim ctl As Access.Control

Set view = frm
Set dataCol = New clsDataCol
Set Delegate = New clsDelegate

For Each ctl In view.Controls
    If ctl.Tag = "Input" _
            Or ctl.ControlType = acCommandButton Then
        If ctl.ControlType = acComboBox Then
            Set Control = New clsComboBox
        ElseIf ctl.ControlType = acTextBox Then
            Set Control = New clsTextBox
        ElseIf ctl.ControlType = acCommandButton Then
            Set Control = New clsCommandButton
        End If
    
        Control.Initialize ctl, Delegate
        
        If Not IsArrayAllocated(InputArray) Then
            ReDim InputArray(0)
        Else
            ReDim Preserve InputArray(UBound(InputArray) + 1)
        End If
        Set InputArray(UBound(InputArray)) = Control
    End If
Next ctl

Dim Sorter As ISorter, Comparer as IComparer
Set Sorter = New SortControls
Set Comparer = New CompareNumeric

InputArray = Sorter.Sort(InputArray, Comparer)

Set Sorter = Nothing
Set Comparer = Nothing

InputArray(LBound(InputArray)).SetFocus

End Sub

'Example 'Delegate_Events'
Private Sub Delegate_PassData(ctl As IControl, dataArr As Variant)
Dim i As Long

For i = LBound(dataArr, 2) To UBound(dataArr, 2)
    dataCol.Store CStr(dataArr(0, i)), dataArr(1, i)
Next i

End Sub

Private Sub Delegate_NextTab(ctl As IControl)
Dim i As Long

For i = LBound(InputArray) To (UBound(InputArray) - 1)
    If InputArray(i).Name = ctl.Name Then
        InputArray(i + 1).SetFocus
    End If
Next i

End Sub

Private Sub Delegate_CancelCustom(ctl as IControl, ByRef cancel As Boolean)

If ctl.State = States.Standard Then
    cancel = MsgBox("Would you like to enter a custom value?", vbYesNo + vbExclamation) = vbNo
    If Not cancel Then
        ctl.State = States.Custom
    End If
End If

End Sub

Private Sub Delegate_RequeryForm(ctl As IControl)
Dim i As Long, j As Long, K As Long

For i = LBound(InputArray) To UBound(InputArray)
    If InputArray(i).Name = ctl.Name Then
        For j = (i + 1) To UBound(InputArray)
            If TypeOf InputArray(j) Is clsComboBox Then
                If InputArray(j).ActiveState = Active Then
                    If Not InputArray(j).DefaultAvailable Then
                        'Resets all boxes down the line
                        For K = (j + 1) To UBound(InputArray)
                            InputArray(K).Reset
                        Next K
                    
                        'Reset dataCol if Necassary
                        Set dataCol = Nothing
                        Set dataCol = New clsDataCol
                    End If
                End If
            End If
        Next j
    End If
Next i

End Sub
