Option Explicit

Implements IComparer

Private Function IComparer_Compare(X As Variant, Y As Variant, Ascending As Boolean) As Long

Dim N As Long

If Ascending = True Then
    N = 1&
Else
    N = -1&
End If

IComparer_Compare = N * StrComp(X, Y, vbTextCompare)

End Function
