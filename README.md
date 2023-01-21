# VBA-PERFORMANCE
 Classes to speed up and monitor the performance of VBA code

--------------------------
CLASS cPerformanceMonitor
--------------------------
Can be tested with
Public Sub TestP()
    Dim cPM As cPerformanceMonitor
    Set cPM = New cPerformanceMonitor
    
    cPM.StartCounter (4)
        cPM.Pause 1
    Debug.Print cPM.TimeElapsed(4)
    
    Set cPM = Nothing
End Sub
