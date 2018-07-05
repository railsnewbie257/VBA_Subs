<pre>
Sub kwh_profile_ranks()
Dim x(20)
Dim y(20)
  
  'select first row  to rank
  
    Set aRange = Selection
    useCol = aRange.Column
    useRow = aRange.Row
    nCols = aRange.Columns.count
    
    botRow = ColumnLastRow(useCol)
    
    For i = useRow To botRow
        For j = 1 To nCols
            x(j) = Cells(i, useCol + j - 1)
            y(j) = j
        Next j
        
        For n = 1 To nCols - 1
            For m = 1 To nCols - n
                If x(m) > x(m + 1) Then
                    tx = x(m + 1)
                    ty = y(m + 1)
                    x(m + 1) = x(m)
                    y(m + 1) = y(m)
                    x(m) = tx
                    y(m) = ty
                End If
            Next m
        Next n
        
        Cells(i + botRow + 2, useCol - 1) = Cells(i, useCol - 1)
        For j = 1 To nCols
            Cells(i + botRow + 2, useCol + y(j) - 1) = j
        Next j
    Next i
    
End Sub
</pre>
