<pre>
Sub RangeToValues(inRange)
    inRange.Copy
    inRange.PasteSpecial _
        Paste:=xlPasteValues, _
        Operation:=xlNone, _
        SkipBlanks:=False, _
        Transpose:=False
    Call ClearClipboard
End Sub
</pre>
