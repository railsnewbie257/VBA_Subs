<b>Dependencies</b>
- [MyUnion](https://github.com/ppihoge/VBA_Subs/blob/master/MyUnion.md)

<pre>
Function FindRangeErrors(useRange As range) As range
Dim aRange As range, bRange As range

    On Error Resume Next
    Set aRange = Nothing
    Set aRange = useRange.SpecialCells(xlCellTypeConstants, xlErrors)
    Set bRange = Nothing
    Set bRange = useRange.SpecialCells(xlCellTypeFormulas, xlErrors)
    Set FindRangeErrors = MyUnion(aRange, bRange)
    If IsEmpty(FindRangeErrors) Then Set FindRangeErrors = Nothing
    
    Set aRange = Nothing
    Set bRange = Nothing
    
End Function
</pre>
