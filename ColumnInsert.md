<pre>
Function ColumnInsertRight(useCol, Optional SHUse, Optional WBUse)
    If IsMissing(SHUse) Then SHUse = ActiveSheet.Name
    If IsMissing(WBUse) Then WBUse = ActiveWorkbook.Name
    Call ClearClipboard
    Call CalculationOff
    Workbooks(WBUse).Worksheets(SHUse).Columns(useCol + 1).Insert Shift:=xlToRight
    Columns(useCol + 1).NumberFormat = "General"
    ColumnInsertRight = useCol + 1
    Call CalculationOn
End Function

Function ColumnInsertLeft(useCol, Optional SHUse, Optional WBUse)
    If IsMissing(SHUse) Then SHUse = ActiveSheet.Name
    If IsMissing(WBUse) Then WBUse = ActiveWorkbook.Name
    Call ClearClipboard
    Call CalculationOff
    Workbooks(WBUse).Worksheets(SHUse).Columns(useCol).Insert Shift:=xlToRight
    Columns(useCol).NumberFormat = "General"
    ColumnInsertLeft = useCol
    Call CalculationOn
End Function
</pre>
