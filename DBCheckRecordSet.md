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
