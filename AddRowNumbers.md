<h2>Add A Column Of Row Numbers</h2>

<pre>
Function AddRowNumbers(Optional useCol, Optional SHUse, Optional WBUse) As Integer
Dim useRange, numberRange As range
Dim botRow As Long
Dim useHeader As String
Dim newCol As Long

    If IsMissing(SHUse) Then SHUse = ActiveSheet.Name
    If IsMissing(WBUse) Then WBUse = ActiveWorkbook.Name
    
    Call ScreenOff
    If IsMissing(useCol) Then
        On Error Resume Next
            Set numberRange = Nothing
            Set numberRange = Application.InputBox("Add Row Numbers for which column?", Title:="AddRowNumbers", Type:=8)
            On Error GoTo 0
            If numberRange Is Nothing Then Exit Function
        useCol = numberRange.Column
    End If
    
    With Workbooks(WBUse).Worksheets(SHUse)
        ' ordering of the next 3 steps is important
        botRow = ColumnLastRow(useCol, SHUse, WBUse)
        useHeader = .Cells(1, useCol).Value
    ' If (useCol = 1) Then
    '     useCol = 0
    '     useHeader = "ThisSheet"
    ' End If
    
        newCol = ColumnInsertRight(useCol, SHUse, WBUse)
    
        Set useRange = range(.Cells(DATASTARTROW, newCol), .Cells(botRow, newCol))
        useRange.NumberFormat = "General"
        useRange.Formula = "=ROW()"
        Call RangeToValues(useRange)
        useRange.NumberFormat = "0"
        Workbooks(WBUse).Worksheets(SHUse).Cells(1, newCol).Value = useHeader & "-RowIndex"
        Call ColorRange(Workbooks(WBUse).Worksheets(SHUse).Cells(1, newCol), LIGHTGREEN)
    End With
    Call ScreenOn
    AddRowNumbers = newCol
End Function
</pre>
