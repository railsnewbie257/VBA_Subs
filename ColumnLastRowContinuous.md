<pre>
Function ColumnLastRowContinuous(Optional topRow, Optional useCol, Optional SHUse, Optional WBUse) ' problem ?

10  On Error GoTo gotError

20  If IsMissing(SHUse) Then SHUse = ActiveSheet.Name
30  If IsMissing(WBUse) Then WBUse = ActiveWorkbook.Name
40  If IsMissing(useCol) Then useCol = 1
    If IsMissing(topRow) Then topRow = 2
50  k = Workbooks(WBUse).Sheets(SHUse).Cells(topRow, useCol).End(xlDown).Row
    
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
