<pre>
Function QueryBuilder(Optional SHQuery, Optional WBQuery)
Dim hasGroupBy As Boolean
Dim twoColumn As Boolean

    If IsMissing(SHQuery) Then SHQuery = ActiveSheet.Name
    If IsMissing(WBQuery) Then WBQuery = ActiveWorkbook.Name

    Call StatusbarDisplay("QueryBuilder: Start")
    'Workbooks(WBQuery).Worksheets(SHQuery).Activate
    
    botRow = LastRow(SHQuery, WBQuery)
    
    With Workbooks(WBQuery).Worksheets(SHQuery)
        
        GLBUserQuery = ""
        fieldCount = 0
        groupBy = ""
        getFields = True
        inSelect = False
        hasGroupBy = False
        twoColumn = False
        
        For i = 1 To botRow
            token1 = Trim(.Cells(i, 1))
            token1 = Replace(token1, Chr(160), " ")
            If Not token1 = "" Then
                If left(token1, 2) = "--" Then GoTo continue

                k = InStr(token1, "--")
                If k > 0 Then token1 = left(token1, k - 2)
                
                GLBUserQuery = GLBUserQuery & token1 & " "

                If UCase(Trim(token1)) = "SELECT" Then
                    inSelect = True
                Else
                    inSelect = False
                End If
                If UCase(Trim(token1)) = "GROUP BY" Then
                    inGroupBy = True
                    hasGroupBy = True
                Else
                    inGroupBy = False
                End If
            End If
            
            token2 = Trim(.Cells(i, 2))
            If Not token2 = "" Then
            
                twoColumn = True
                If Not IsEmpty(.Cells(i + 1, 2)) Then ' one line look ahead
                    c = "," & vbNewLine
                Else
                    c = " " & vbNewLine
                   End If
                GLBUserQuery = GLBUserQuery & token2 & c
                If inSelect Then  ' only group fields in the SELECT
                    '
                    ' Exclude aggregates
                    '
                    If UCase(left(token2, 5)) = "COUNT" Or _
                        UCase(left(token2, 3)) = "MIN" Or _
                        UCase(left(token2, 3)) = "MAX" Or _
                        UCase(left(token2, 3)) = "SUM" Then
                        
                        fieldCount = fieldCount + 1
                    Else
                        fieldCount = fieldCount + 1
                        If Len(groupBy) > 0 Then groupBy = groupBy & ","
                        groupBy = groupBy & format(fieldCount, "0")
                    End If
                End If
            End If
continue:
        Next i
        
    End With 'Workbooks(WBQuery).Worksheets(SHQuery)
    
    If Not hasGroupBy And twoColumn Then GLBUserQuery = GLBUserQuery & " GROUP BY " & groupBy
    Debug_Print GLBUserQuery
    
    GLBUserQuery = Replace(GLBUserQuery, ",,", ",")
    
    QueryBuilder = GLBUserQuery
    
    Call StatusbarDisplay("QueryBuilder: Done")
    'Workbooks(WBQuery).Worksheets(SHQuery).Activate
    
End Function
</pre>
