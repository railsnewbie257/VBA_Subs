<pre>
Public targetWorkbook As String
Public targetWorksheet As String

Function ChooseTargetSheet()
    
    'If targetWorkbook <> "" And targetWorksheet <> "" Then Exit Function
    
    On Error Resume Next
    Set aRange = Nothing
    Set aRange = Application.InputBox("Select Workbook and Sheet", Type:=8)
    If aRange Is Nothing Then End
    
    targetWorkbook = aRange.Parent.Parent.Name
    targetWorksheet = aRange.Parent.Name
End Function
</pre>
