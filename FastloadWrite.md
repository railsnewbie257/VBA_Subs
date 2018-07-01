<pre>
Sub FastLoadWrite(filePath, str)
Dim fso As Object
Dim oFile As Object

    Set fso = CreateObject("Scripting.FileSystemObject")
    '
    ' 1 - readonly
    ' 2 - writing
    ' 8 - append
    '
    ' 0 - Ascii format
    Set oFile = fso.OpenTextFile(filePath, 8, True, 0)
    
    oFile.WriteLine str
    oFile.Close

    Set fso = Nothing  ' for garbage collector
    Set oFile = Nothing

End Sub
</pre>
