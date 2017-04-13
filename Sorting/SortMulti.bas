Option Explicit

Implements ISorter

Private Function ISorter_Sort(Arr As Variant, Comparer As IComparer, _
        Optional LB As Long = -1&, Optional UB As Long = -1&, Optional Ascending As Boolean = True, _
        Optional sortCol As Long = 1&) As Variant

SortValues Arr, sortCol, Comparer, LB, UB, Ascending
ISorter_Sort = Arr
    
End Function

Private Sub SortValues(inputArray As Variant, col As Long, Comparer As IComparer, _
        ByVal LB As Long, ByVal UB As Long, Ascending As Boolean)

Dim pivot As Variant
Dim tmpSwap As Variant
Dim tmpLow As Long
Dim tmpHigh As Long
Dim tmpRow As Variant
Dim tmpCol As Long

If LB < 0 Then LB = LBound(inputArray)
If UB < 0 Then UB = UBound(inputArray)
tmpLow = LB
tmpHigh = UB

pivot = inputArray((tmpLow + tmpHigh) \ 2, col)

Do While tmpLow <= tmpHigh
    Do While (Comparer.Compare(inputArray(tmpLow, col), pivot, Ascending) < 0) _
            And (tmpLow < UB)
        tmpLow = tmpLow + 1
    Loop
    
    Do While (Comparer.Compare(pivot, inputArray(tmpHigh, col), Ascending) < 0) _
            And (tmpHigh > LB)
        tmpHigh = tmpHigh - 1
    Loop
    
    If tmpLow <= tmpHigh Then
        ReDim tmpRow(LBound(inputArray, 2) To UBound(inputArray, 2)) As Variant
        For tmpCol = LBound(inputArray, 2) To UBound(inputArray, 2)
            tmpRow(tmpCol) = inputArray(tmpLow, tmpCol)
            inputArray(tmpLow, tmpCol) = inputArray(tmpHigh, tmpCol)
            inputArray(tmpHigh, tmpCol) = tmpRow(tmpCol)
        Next tmpCol
        
        tmpLow = tmpLow + 1
        tmpHigh = tmpHigh - 1
    End If
Loop

If LB < tmpHigh Then SortValues inputArray, col, Comparer, LB, tmpHigh, Ascending
If tmpLow < UB Then SortValues inputArray, col, Comparer, tmpLow, UB, Ascending

End Sub
