<pre>
    If IsMissing(SHUse) Then SHUse = ActiveSheet.Name
    If IsMissing(WBUse) Then WBUse = ActiveWorkbook.Name
    
    NextRow = LastRow(SHUse, WBUse) + 1
End Function
</pre>
