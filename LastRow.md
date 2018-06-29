<pre>
Function LastRow(Optional SHUse, Optional WBUse) As Long

    If (IsMissing(SHUse)) Then SHUse = ActiveSheet.Name
    If (IsMissing(WBUse)) Then WBUse = ActiveWorkbook.Name

    If WorksheetFunction.CountA(Workbooks(WBUse).Sheets(SHUse).Cells) > 0 Then

        'Search for any entry, by searching backwards by Rows.

        LastRow = Workbooks(WBUse).Sheets(SHUse).Cells.Find(What:="*", _
            After:=[A1], _
            SearchOrder:=xlByRows, _
            SearchDirection:=xlPrevious).Row
              
        'LastRow = Cells.SpecialCells(xlCellTypeLastCell).Row
    End If

End Function
</pre>

--------------------------------------------
Possible Alternative


<pre>
'
' This function will return 0 if nothing there
'
Function LastRow2(Optional SHUse, Optional WBUse) As Long
Dim useRow As Long, useCol As Long
Dim t As Variant

On Error GoTo Err1:

    If (IsMissing(SHUse)) Then SHUse = ActiveSheet.Name
    If (IsMissing(WBUse)) Then WBUse = ActiveWorkbook.Name
    'LastRow = Workbooks(WBUse).Worksheets(SHUse).Cells.Find(What:="*", _
    '                after:=Workbooks(WBUse).Worksheets(SHUse).Cells(1, 1), _
    '                LookAt:=xlPart, _
    '                LookIn:=xlFormulas, _
    '                SearchOrder:=xlByRows, _
    '                SearchDirection:=xlPrevious, _
    '                MatchCase:=False).Row
    
    useRow = Workbooks(WBUse).Worksheets(SHUse).Cells.SpecialCells(xlCellTypeLastCell).Row
    useCol = Workbooks(WBUse).Worksheets(SHUse).Cells.SpecialCells(xlCellTypeLastCell).Column
    LastRow = useRow
    If useRow = 1 Then
        useCol = Workbooks(WBUse).Worksheets(SHUse).Cells.SpecialCells(xlCellTypeLastCell).Column
        If IsEmpty(Cells(useRow, useCol)) Then LastRow = 0
    End If
    Exit Function

Err1:
    LastRow = 0
    
End Function
</pre>
