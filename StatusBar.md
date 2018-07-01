<pre>
Function <b>StatusBarOn</b>()
    Application.DisplayStatusBar = True
End Function
</pre>

<pre>
Function <b>StatusBarOff</b>()
    Application.DisplayStatusBar = False
End Function
</pre>

<pre>
Sub <b>StatusbarDisplay</b>(Optional s)
    Application.DisplayStatusBar = True
    If IsMissing(s) Then s = "testing..."
        Application.StatusBar = s
        'DoEvents
End Sub
</pre>
