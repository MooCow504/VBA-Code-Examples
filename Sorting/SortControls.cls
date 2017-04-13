Implements ISorter

Private Function ISorter_Sort(Arr As Variant, Comparer As IComparer, _
        Optional LB As Long = -1&, Optional UB As Long = -1&, Optional Ascending As Boolean = True, _
        Optional sortCol As Long = 1&) As Variant

Dim varArr As Variant
varArr = Arr

SortValues varArr, Comparer, LB, UB, Ascending
ISorter_Sort = varArr
  
End Function

Private Sub SortValues(ByRef inputArray As Variant, Comparer As IComparer, _
        ByVal LB As Long, ByVal UB As Long, Ascending As Boolean)

Dim pivot As Variant
Dim tmpSwap As Variant
Dim tmpLow As Long
Dim tmpHigh As Long

If LB < 0 Then LB = LBound(inputArray)
If UB < 0 Then UB = UBound(inputArray)
tmpLow = LB
tmpHigh = UB

pivot = inputArray((tmpLow + tmpHigh) \ 2).Top

Do While tmpLow <= tmpHigh
    Do While (Comparer.Compare(inputArray(tmpLow).Top, pivot, Ascending) < 0) _
            And (tmpLow < UB)
        tmpLow = tmpLow + 1
    Loop
    
    Do While (Comparer.Compare(pivot, inputArray(tmpHigh).Top, Ascending) < 0) _
            And (tmpHigh > LB)
        tmpHigh = tmpHigh - 1
    Loop
    
    If tmpLow <= tmpHigh Then
        tmpSwap = inputArray(tmpLow)
        inputArray(tmpLow) = inputArray(tmpHigh)
        inputArray(tmpHigh) = tmpSwap
        tmpLow = tmpLow + 1
        tmpHigh = tmpHigh - 1
    End If
Loop

If LB < tmpHigh Then SortValues inputArray, Comparer, LB, tmpHigh, Ascending
If tmpLow < UB Then SortValues inputArray, Comparer, tmpLow, UB, Ascending

End Sub
