<pre>
Sub kwh_profile_ranks()
Dim x(20)
Dim y(20)

    Set aRange = Selection
    useCol = aRange.Column
    useRow = aRange.Row
    nCols = aRange.Columns.count
    
    botRow = ColumnLastRow(useCol)
    
    totalCol = FindColumnHeader("totalkwh")
    
    k = 0
    For i = useRow To botRow
    
        If (Cells(i, totalCol) = 0) Then
            k = k - 1
            GoTo skip
        End If
        
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
        
        Cells(i + botRow + 2 + k, useCol - 1) = Cells(i + k, useCol - 1)
        For j = 1 To nCols
            Cells(i + botRow + 2 + k, useCol + y(j) - 1) = j
        Next j
        'k = k + 1
skip:
    Next i
    
End Sub
</pre>
