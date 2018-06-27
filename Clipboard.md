<pre>
'
' Manipulate the Clipboard, also has general purpose EmptyClipboard() to deallocate Clipboard
'
Sub test_IsClipboardEmpty()
    If IsClipboardEmpty Then
        MsgBox "Clipboard Empty"
    Else
        MsgBox "Clipboard NOT Empty"
    End If
End Sub

Sub EmptyClipboard()
    Application.CutCopyMode = False
End Sub

Function IsClipboardEmpty()
Dim myDataObject As DataObject
Dim A As String

    Set myDataObject = New DataObject
    myDataObject.GetFromClipboard
    A = myDataObject.GetFormat(1)
    MsgBox A
    MsgBox "Records = " & UBound(Split(A, Chr(13) & Chr(10)))
    If myDataObject.GetFormat(1) = True Then
        IsClipboardEmpty = False
    Else
        IsClipboardEmpty = True
    End If

End Function

Sub CopyToClipboard(s As String)
Dim DataObj As New MSForms.DataObject

    DataObj.SetText s
    DataObj.PutInClipboard
End Sub
</pre>
