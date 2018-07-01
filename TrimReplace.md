
Makes sure identifiers (column nmaes) are legal for Fastload

<pre>
Function TrimReplace(t)
Dim s As String
    s = Trim(t)
    s = Replace(s, ".", "")
    s = Replace(s, " ", "_")
    s = Replace(s, "(", "_")
    s = Replace(s, ")", "_")
    s = Replace(s, "/", "_")
    s = Replace(s, "-", "_")
    s = Replace(s, "?", "_")
    s = Replace(s, ">", "_")
    s = Replace(s, "<", "_")
    s = Replace(s, "'", "")
    s = Replace(s, """", "")
    s = Replace(s, "%", "pct")
    t1 = Len(s)
    s = Replace(s, "__", "_")
    If Len(s) <> t1 Then s = Replace(s, "__", "_")
    t1 = Len(s)
    If Len(s) <> t1 Then s = Replace(s, "__", "_")
    t1 = Len(s)
    If Len(s) <> t1 Then s = Replace(s, "__", "_")
    TrimReplace = s
End Function
</pre>
