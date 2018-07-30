<h2>DBCheckRecordset</h2>

<pre>
Function DBCheckRecordset(DBRecordset)
    Call StatusbarDisplay("DBCheckRecordset: Check for Nothing.")
    If DBRecordset Is Nothing Then
        Call StatusbarDisplay("DBCheckRecordset: Allocate New.")
        Set DBCheckRecordset = New ADODB.Recordset
    Else
        Set DBCheckRecordset = DBRecordset
    End If
    Call StatusbarDisplay("DBCheckRecordset: Return.")
End Function
</pre>

<h2>DBCloseRecordset</h2>

<pre>
Function DBCloseRecordset(DBRecordset)
    If Not DBRecordset Is Nothing Then
        DBRecordset.Close
        Set DBRecordset = Nothing
    End If
End Function
</pre>
