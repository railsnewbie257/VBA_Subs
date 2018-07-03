Resets AutoParsing for copying from SQL Assistant (Tab delimited)

<pre>
Function SetAutoParse()
    t = Range("A1")
    Range("A1") = 1
    Range("A1").TextToColumns Destination:=Range("A1"), _
        DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, _
        ConsecutiveDelimiter:=False, _
        Tab:=True, _
        Semicolon:=False, _
        Comma:=False, _
        Space:=False, _
        Other:=False, _
        OtherChar:="", _
        FieldInfo:=Array(1, 1), _
        TrailingMinusNumbers:=False
    Range("A1") = t
End Function
</pre>
