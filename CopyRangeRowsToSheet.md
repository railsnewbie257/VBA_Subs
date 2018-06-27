<pre>
Sub CopyRangeRowsToSheet(rowRange As range, SHTo As String)
Dim copyRange As range, rw As range
Dim SHFrom As String
Dim i As Integer, rightCol As Integer

    'rowRange.Copy
    'Worksheets(SHTo).Activate
    
    'Cells(2, 1).PasteSpecial xlPasteAll
    
    SHFrom = rowRange.Parent.Name
    
    i = 2
    For Each rw In rowRange
        rightCol = RowLastColumn(rw.Row, SHFrom)
        Set copyRange = range(Worksheets(SHFrom).Cells(rw.Row, 1), Worksheets(SHFrom).Cells(rw.Row, rightCol))
        copyRange.Copy Destination:=Worksheets(SHTo).Cells(i, 1)
        i = i + 1
    Next rw
End Sub
</pre>
