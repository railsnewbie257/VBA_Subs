<pre>
Sub jeenifer()
    botRow = ColumnLastRow(1)
    
    For i = 2 To botRow Step 4
        Cells(i, 2) = Cells(i + 1, 1)
        Cells(i, 3) = Cells(i + 2, 1)
        'Rows(i + 2).Delete
        'Rows(i + 1).Delete
    Next i
End Sub
</pre>

<pre>
Sub jeenifer2()
    botRow = ColumnLastRow(1)
    
    For i = 3 To botRow Step 4
        Cells(i, 1) = Cells(i, 1) & Cells(i, 2)
        Cells(i, 2) = ""
        'Rows(i + 2).Delete
        'Rows(i + 1).Delete
    Next i
End Sub
</pre>

<pre>
Sub jeenifer3()
    botRow = ColumnLastRow(1)
    
    For i = 4 To botRow Step 4
        Cells(i, 2) = Cells(i, 2) & Cells(i, 3)
        Cells(i, 3) = ""
        'Rows(i + 2).Delete
        'Rows(i + 1).Delete
    Next i
End Sub
</pre>

<pre>
Sub jeenifer4()
    botRow = ColumnLastRow(1)
    </pre>
    For i = 2 To botRow Step 4
        Cells(i, 4) = Cells(i + 1, 1)
        Cells(i, 5) = Cells(i + 1, 3)
        Cells(i + 1, 1) = ""
        Cells(i + 1, 3) = ""
        
        Cells(i, 6) = Cells(i + 2, 1)
        Cells(i, 7) = Cells(i + 2, 2)
        Cells(i + 2, 1) = ""
        Cells(i + 2, 2) = ""
        'Rows(i + 2).Delete
        'Rows(i + 1).Delete
    Next i
End Sub
</pre>
