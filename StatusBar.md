<pre>
Function StatusBarOn()
    Application.DisplayStatusBar = True
End Function

Function StatusBarOff()
    Application.DisplayStatusBar = False
End Function

Sub StatusbarDisplay(Optional s)
    Application.DisplayStatusBar = True
    If IsMissing(s) Then s = "testing..."
        Application.StatusBar = s
        'DoEvents
End Sub
</pre>
