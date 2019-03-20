This needs to go into <b><em>ThisWorkbook</em></b>

It is dependent on **Util_StoreRetrieve.bas**

<pre>
Private Sub Workbook_Activate()
    WBUse = Application.Windows(2).Caption
    Call StoreCurrentWorkbook(WBUse)
    Call StoreCurrentSheet(Workbooks(WBUse).ActiveSheet.Name)
End Sub
</pre>
