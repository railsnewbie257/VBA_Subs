<pre>
Sub DeleteColumnUseHeader(headerName, Optional SHUse, Optional WBUse)
Dim aCol As Integer

    If IsMissing(SHUse) Then SHUse = ActiveSheet.Name
    If IsMissing(WBUse) Then WBUse = ActiveWorkbook.Name
    
    aCol = FindColumnHeader(headerName, SHUse, WBUse)
    Workbooks(WBUse).Worksheets(SHUse).Columns(aCol).Delete
End Sub
</pre>
