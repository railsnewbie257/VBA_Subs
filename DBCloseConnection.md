<pre>
Function DBCloseConnection(Optional DBConn)
    If IsMissing(DBConn) Then Set DBConn = DBGlbConnection
    If Not DBConn Is Nothing Then
        If DBConn.State <> 0 Then DBConn.Close
        Set DBConn = Nothing
        If DBGlbConnection.State <> 0 Then DBGlbConnection.Close
        Set DBGlbConnection = Nothing
        On Error Resume Next
        ' DBConn.Close
        ' Set DBConn = Nothing
    End If
    'MsgBox "Database Connection Reset", Title:="DBCloseConnection"
End Function
</pre>
