<pre>

Function FindRangeNonEmpty(searchRange)
    
    On Error Resume Next
    Set aRange = Nothing
    Set aRange = searchRange.SpecialCells(xlCellTypeConstants)
    Set bRange = Nothing
    Set bRange = searchRange.SpecialCells(xlCellTypeFormulas)
    Set FindRangeNonEmpty = MyUnion(aRange, bRange)
    If IsEmpty(FindRangeNonEmpty) Then Set FindRangeNonEmpty = Nothing
    
    Set aRange = Nothing
    Set bRange = Nothing
End Function

</pre>
