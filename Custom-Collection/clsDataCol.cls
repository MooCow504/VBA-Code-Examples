Option Explicit

Private dataCol As Collection, kv As KeyValue

Private Sub Class_Terminate()

Set dataCol = Nothing

End Sub

Public Sub Class_Initialize()

Set dataCol = New Collection

End Sub

Public Function Retrieve(Key As String) As String
'Returns String instead of Variant so empty values = "", instead of = Empty

If Contains(Key) Then
    Set kv = dataCol(Key)
    Retrieve = kv.cValue
End If

End Function

Public Sub Remove(Key As String)

If Contains(Key) Then
    dataCol.Remove (Key)
End If

End Sub

Public Sub Store(Key As String, Value As Variant)

If Contains(Key) Then
    Set kv = dataCol(Key)
    kv.cValue = Value
Else
    Set kv = New KeyValue
    kv.Init Key, Value
    dataCol.Add kv, Key
End If

End Sub

Private Function Contains(Key As String) As Boolean 'col As Collection, Key As String) As Boolean

'Recieve compile error at line 'col (Key)' if collection used not scoped to function
Dim col As Collection
Set col = dataCol

On Error Resume Next
col (Key)
Contains = (Err.Number = 0)
Err.Clear

Set col = Nothing

End Function
