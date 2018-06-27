<pre>
Sub SpeedupOn()
    With Application
        .Calculation = xlCalculationManual
        .ScreenUpdating = False
        .DisplayStatusBar = False
        .EnableEvents = False
        '.DisplayPageBreaks = False
    End With
End Sub

Sub SpeedupOff()
    With Application
        .Calculation = xlCalculationAutomatic
        .ScreenUpdating = True
        .DisplayStatusBar = True
        .EnableEvents = True
        '.DisplayPageBreaks = True
    End With
End Sub
</pre>
