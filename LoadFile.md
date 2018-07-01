<pre>
Sub LoadFile()
Dim txtFilename As Variant
Dim t As String

    LoadFileDirectory.Show
    If formCancel Then Exit Sub
    
    lastFile = LatestFile(GLBFilePath)
    t = InStr(1, GLBFilePath, "SSN", vbTextCompare)
    'If InStr(1, GLBFilePath, "SSN", vbTextCompare) > 0 Then
    '    lastFile = "SSN" & lastFile
    'End If

    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    With fd
        .AllowMultiSelect = False
        .Title = "Please Select File To Load"
        .Filters.Clear
        .Filters.Add "Excel", "*.xls?"
        .InitialFileName = GLBFilePath & lastFile
        
        If .Show = False Then Exit Sub
        
        useFile = .SelectedItems(1)

        If GLBOpenReadOnly Then
            Workbooks.Open useFile, ReadOnly:=True
            Debug_Print "readonly"
        Else
            Workbooks.Open useFile
        End If
    End With

    Set fd = Nothing
End Sub
</pre>
