Option Explicit

Implements IComparer

Private Function IComparer_Compare(X As Variant, Y As Variant, Ascending As Boolean) As Long

Dim N As Integer

If Ascending = True Then
    N = 1&
Else
    N = -1&
End If

If X > Y Then
    IComparer_Compare = N * 1&
ElseIf Y > X Then
    IComparer_Compare = N * -1&
Else
    IComparer_Compare = 0&
End If

End Function
