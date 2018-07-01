<b>Dependency</b>
- [LastRow](https://github.com/ppihoge/VBA_Subs/blob/master/LastRow.md)

<pre>
Function ColumnLastRow(Optional useCol, Optional SHUse, Optional WBUse) ' problem ?

10  On Error GoTo gotError

20  If IsMissing(SHUse) Then SHUse = ActiveSheet.Name
30  If IsMissing(WBUse) Then WBUse = ActiveWorkbook.Name
40  If IsMissing(useCol) Then useCol = 1
50  k = LastRow(SHUse, WBUse)
    
60  With Workbooks(WBUse).Worksheets(SHUse)
70      ColumnLastRow = .Cells(k + 1, useCol).End(xlUp).Row
80      If IsEmpty(.Cells(ColumnLastRow, useCol)) Then ColumnLastRow = 0
90  End With

100 Exit Function

gotError:
    MsgBox Err.Number & " " & Err.Description & vbNewLine & vbNewLine & "Error on line: " & Erl, Title:="ColumnLastRow"
    Stop
    Resume Next
End Function
</pre>
