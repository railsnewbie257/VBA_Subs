<pre>
Function FindColumnHeader(columnName, Optional SHUse, Optional WBUse) As Long
Dim sRange As range, fRange As range

    If IsMissing(SHUse) Then SHUse = ActiveSheet.Name
    If IsMissing(WBUse) Then WBUse = ActiveWorkbook.Name
    '
    ' expanded since top row may have the FullTableName
    '
    With Workbooks(WBUse).Worksheets(SHUse)
        Set sRange = range(.Rows(1), .Rows(1))
        'sRange.Copy
        Set fRange = FindInRange(columnName, sRange)
        If fRange Is Nothing Then       ' may be 2 row header
            Set sRange = range(.Rows(1), .Rows(2))
            'sRange.Copy
            Set fRange = FindInRange(columnName, sRange)
        End If
    End With
    
    If Not fRange Is Nothing Then
        FindColumnHeader = fRange(1).Column
        Exit Function
    Else
        FindColumnHeader = -1
    End If
    
    Set sRange = Nothing
    Set fRange = Nothing
    
End Function
</pre>
