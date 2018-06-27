<pre>
Function RowLastColumn(Optional useRow, Optional SHUse, Optional WBUse) As Long

    If IsMissing(SHUse) Then SHUse = ActiveSheet.Name
    If IsMissing(WBUse) Then WBUse = ActiveWorkbook.Name
    
    With Workbooks(WBUse).Worksheets(SHUse)
        RowLastColumn = .Cells(useRow, .Columns.count).End(xlToLeft).Column
        If (IsEmpty(.Cells(useRow, RowLastColumn))) Then RowLastColumn = 0
    End With
End Function
</pre>
