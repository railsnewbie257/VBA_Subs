<pre>
Public Function DirExists(s_directory As String) As Boolean

Set OFSO = CreateObject("Scripting.FileSystemObject")
DirExists = OFSO.FolderExists(s_directory)

End Function
</pre>

<pre>
Public Function FolderExists(strFolderPath As String) As Boolean
    On Error Resume Next
    FolderExists = ((GetAttr(strFolderPath) And vbDirectory) = vbDirectory)
    On Error GoTo 0
End Function
</pre>
