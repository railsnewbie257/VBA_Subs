<pre>
Sub QueryNewCondition(SHQuery, WBQuery, useCondition, Optional newValue, Optional newCondition)
Dim sRange As range, fRange As range
Dim botRow As Long
Dim s As String
Dim dataCol As Integer

10  On Error GoTo gotError
20  dataCol = QUERYDATACOL
30  If IsColumnEmpty(dataCol, SHQuery, WBQuery) Then dataCol = 1  ' this is a cheat incase the query is only in the first column
40  botRow = ColumnLastRow(dataCol, SHQuery, WBQuery)
    
50  Set sRange = range(Workbooks(WBQuery).Worksheets(SHQuery).Cells(1, dataCol), _
                       Workbooks(WBQuery).Worksheets(SHQuery).Cells(botRow + 2, dataCol))
    
60  Set fRange = FindInRange(useCondition, sRange)
    
    If Not IsMissing(newCondition) Then
        s = newCondition
    Else
70      s = useCondition & " " & newValue
    End If
    
80  't = sRange.Cells(1, 1)
90  't = sRange.Cells(93, 1)
100 If Not fRange Is Nothing Then
110     fRange.Value = s
120 Else
130     Debug_Print useCondition & " not found"  ' warning message
140 End If
    
150 Exit Sub

gotError:
    MsgBox Err.Number & " " & Err.Description & vbNewLine & vbNewLine & "Error on line: " & Erl
    Stop
    Resume Next
    
End Sub
</pre>
