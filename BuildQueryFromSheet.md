<b>Dependencies</b>
- ColumnLastRow

<pre>
Function BuildQueryFromSheet(SHUse)

    botRow = ColumnLastRow(1)
    
    t = ""
    With ThisWorkbook.Sheets(SHUse)
        For i = 1 To botRow
             t = t & Trim(.Cells(i, 1)) & " "
        Next i
    End With

    BuildQueryFromSheet = t
End Function
</pre>
