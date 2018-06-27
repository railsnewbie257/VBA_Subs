<pre>
Dim startTime As Date, stopTime As Date

Sub StartTimer()
    startTime = Now() ' global
    Debug_Print format(startTime, "HH:MM:SS")
End Sub

Sub StopTimer()
    stopTime = Now()  ' global
    Debug_Print format(stopTime, "HH:MM:SS")
End Sub

Function ElapsedTime()
    stopTime = Now()
    ElapsedTime = "Elapsed Time: " & format(stopTime - startTime, "HH:MM:SS")
End Function

Sub testTimer()
    Call StartTimer
    Debug_Print Now()
    Call Application.Wait(Now + TimeValue("0:01:00"))
    Debug_Print Now()
    Call StopTimer
    
    Debug_Print ElapsedTime & " " & format(startTime, "HH:MM:SS") & " " & format(stopTime, "HH:MM:SS")
End Sub

Sub FillData()
    Call StartTimer
    For i = 1 To 10000
        Cells(i, 1) = 100
    Next i
    Call StopTimer
    Debug_Print ElapsedTime
End Sub
</pre>
