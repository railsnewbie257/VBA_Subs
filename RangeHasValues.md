<pre>
Function RangeHasValues(inRange) As range
Dim numRange As range, txtRange As range
    Set RangeHasValues = inRange.SpecialCells(xlCellTypeConstants)
End Function
</pre>
