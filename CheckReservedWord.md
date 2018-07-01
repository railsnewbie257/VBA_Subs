<pre>
Function CheckReservedWord(word)

    word = TrimReplace(word)
    ThisWorkbook.Worksheets("SQLReservedWords").Cells(1, 1) = word
    If Not IsError(ThisWorkbook.Worksheets("SQLReservedWords").Cells(2, 1)) Then
        CheckReservedWord = "a_" & word
    Else
        CheckReservedWord = word
    End If
    
End Function
</pre>

<pre>
Function IsReservedWord(word)

    ThisWorkbook.Worksheets("SQLReservedWords").Cells(1, 1) = word
    If Not IsError(ThisWorkbook.Worksheets("SQLReservedWords").Cells(2, 1)) Then
        IsReservedWord = True
    Else
        IsReservedWord = False
    End If
    
End Function
</pre>
