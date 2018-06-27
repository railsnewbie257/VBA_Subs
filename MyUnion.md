<pre>
Function MyUnion(aRange, bRange) As range
    If Not (aRange Is Nothing) And Not (bRange Is Nothing) Then
        Set MyUnion = Application.Union(aRange, bRange)
    ElseIf aRange Is Nothing Then
        Set MyUnion = bRange
    Else
        Set MyUnion = aRange
    End If
End Function
</pre>
