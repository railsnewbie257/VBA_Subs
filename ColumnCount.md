<pre>
Function ColumnCountA(useCol)
    Set aRange = range(Cells(2, useCol), Cells(ColumnLastRow(useCol), useCol))
    ColumnCountA = WorksheetFunction.CountA(aRange)
    ColumnCountA = ActiveSheet.Columns(useCol).Cells.SpecialCells(xlCellTypeConstants).count
End Function
'
' similar to filterMultipleWorkOrders
'
Sub ColumnCountValues(Optional useCol)
Dim colRange As range
Dim botRow As Long, i As Long, k As Long

    If IsMissing(useCol) Then useCol = ActiveCell.Column
    Set colRange = range(Cells(1, useCol), Cells(1, useCol))
    
    Call SortSheetUp(colRange.Column)
    
    countCol = ColumnInsertLeft(colRange.Column)
    Cells(1, countCol) = "Count of " & Cells(1, colRange.Column)
    
    botRow = ColumnLastRow(colRange.Column) + 1
    
    i = 2
    k = 1
    While i < botRow
        While (Cells(i, colRange.Column) = Cells(i + 1, colRange.Column))
            k = k + 1
            Rows(i + 1).Delete
            botRow = botRow - 1
        Wend
        Cells(i, countCol) = k
        
        i = i + 1
        k = 1
    Wend

End Sub
</pre>
