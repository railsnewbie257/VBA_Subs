<pre>
Sub FastloadMerge(fullTableName)
    
Dim DBCn As ADODB.Connection
Dim DBRs As ADODB.Recordset

On Error GoTo gotError

10    qq = "INSERT INTO " & left(fullTableName, Len(fullTableName) - 3) & " SELECT * from " & fullTableName

        Debug_Print qq
20    Set DBCn = DBCheckConnection(DBCn)
30    Set DBRs = DBCheckRecordset(DBRs)

40    With DBRs
50        .CursorLocation = adUseServer ' adUseClient
60        .CursorType = adOpenDynamic ' adUseClient ' adOpenStatic ' adOpenDynamic ' adOpenForwardOnly
70        .LockType = adLockOptimistic  ' adLockReadOnly
80        Set .ActiveConnection = DBCn
90    End With

100   DBRs.Open qq, DBCn

110   Exit Sub
    
gotError:
    MsgBox Err.Number & " " & Err.Description & vbNewLine & vbNewLine & "Error on line: " & Erl, Title:="FastloadMerge"
    Stop
    Resume Next
End Sub
</pre>
