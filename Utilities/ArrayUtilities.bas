Option Explicit

Public Function FirstEmpty(Arr As Variant, Optional Element As Long = 1, Optional Dimension As Long = 1) As Long
Dim i As Long

If Dimension = 1 Then
    For i = LBound(Arr) To UBound(Arr)
        If Arr(i, Element) = "" Then
            FirstEmpty = i
            Exit Function
        End If
    Next i
ElseIf Dimension = 2 Then
    For i = LBound(Arr, 2) To UBound(Arr, 2)
        If Arr(Element, i) = "" Then
            FirstEmpty = i
            Exit Function
        End If
    Next i
Else
    'Throw some kind of error
End If

End Function

Public Function RemoveElement(Arr As Variant, Element As Long, Optional Dimension As Long = 1) As Variant
Dim i As Long, j As Long

Dim outputArray() As Variant
Dim outElement As Long, inElement As Long

If Dimension = 1 Then
    ReDim outputArray(LBound(Arr) To (UBound(Arr) - 1), LBound(Arr, 2) To UBound(Arr, 2)) As Variant
   
    For i = LBound(Arr) To UBound(Arr)
        For j = LBound(Arr, 2) To UBound(Arr, 2)
            If i <> Element Then
                outElement = i
                If i > Element Then outElement = outElement - 1
                outputArray(outElement, j) = Arr(i, j)
            End If
        Next j
    Next i
ElseIf Dimension = 2 Then
    ReDim outputArray(LBound(Arr) To UBound(Arr), LBound(Arr, 2) To (UBound(Arr, 2) - 1)) As Variant

    For i = LBound(Arr) To UBound(Arr)
        For j = LBound(Arr, 2) To UBound(Arr, 2)
            If j <> Element Then
                outElement = j
                If j > Element Then outElement = outElement - 1
                outputArray(i, outElement) = Arr(i, j)
            End If
        Next j
    Next i
Else
    'Throw some kind of error
End If

RemoveElement = outputArray

End Function

Public Function IsArrayAllocated(Arr As Variant) As Boolean
Dim N As Long

On Error Resume Next
If Not IsArray(Arr) Then
    IsArrayAllocated = False
    Exit Function
End If

N = UBound(Arr)
If (Err.Number = 0) Then
    'In certain situations no error is thrown when checking UBound
    'Check UBound against LBound to verify
    If LBound(Arr) <= UBound(Arr) Then
        IsArrayAllocated = True
    Else
        IsArrayAllocated = False
    End If
Else
    IsArrayAllocated = False
End If

End Function

Public Function ReturnUniqueValues(Arr() As String) As String()
Dim i As Long

Dim uniqueCol As New Collection, uniqueValues() As String

On Error Resume Next
For i = LBound(Arr) To UBound(Arr)
    uniqueCol.Add Arr(i), Arr(i)
Next i
On Error GoTo 0

ReDim uniqueValues(1 To uniqueCol.Count) As String
For i = 1 To uniqueCol.Count
    uniqueValues(i) = uniqueCol(i)
Next i

ReturnUniqueValues = uniqueValues

End Function

