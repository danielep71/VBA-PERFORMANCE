VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cPerformanceMonitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
'                         CLASS PERFORMANCE MONITOR                           =
'==============================================================================
'
'************
'DESCRIPTION*   High-precision timer used to benchmark procedures' performance.
'************
'
'************
'AUTHOR     *   Daniele Penza
'************
'
'************
'VERSION    *   1.0.1
'************
'
'************
'LAST UPDATE*   21/01/2023
'************
'
'==============================================================================
'************
'PROCEDURES *   FUNCTIONS AND SUBROUTINES
'************
'
'  1. PRI - SUB - Class_Initialize
'  2. PUB - SUB - StartCounter
'  3. PUB - PRO - TimeElapsed
'  4. PUB - SUB - Pause
'
'==============================================================================
'
'------------------------------------------------------------------------------
'MODULE SETTINGS
'------------------------------------------------------------------------------
    Option Explicit     'Force to declare all variables
'------------------------------------------------------------------------------
'MODULE DECLARATIONS
'------------------------------------------------------------------------------
    '
    '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    'API DECLARATIONS (work with every VBA (VBA6 and VBA7) and bitness (32/64)
    '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    '
    'GETTICKCOUNT API
        #If VBA7 Then
            'Use LongPtr as return value type for both declarations to avoid
            'platform dependencies in the calling code
            #If Win64 Then
                Private Declare PtrSafe Function GetTickCount Lib "kernel32" _
                    Alias "GetTickCount64" () As LongPtr    'LongLong
            #Else
                Private Declare PtrSafe Function GetTickCount Lib "kernel32" _
                    () As LongPtr                           'Long
            #End If
        #Else
            Private Declare Function GetTickCount Lib "kernel32" _
                    () As Long
        #End If

    'TIMEGETTIME API
        #If VBA7 Then
            Private Declare PtrSafe Function timeGetTime Lib "winmm" () As Long
        #Else
            Private Declare Function timeGetTime Lib "winmm.dll" Alias "timeGetTime" () As Long
        #End If
        
    'TIMEGESTSYSTEMTIME API
        Private Type SMPTE
            Hour        As Byte
            Min         As Byte
            Sec         As Byte
            Frame       As Byte
            Fps         As Byte
            Dummy       As Byte
            Pad(2)      As Byte
        End Type
        Private Type MMTIME
            wType       As Long
            Units       As Long
            smpteVal    As Long
            songPtrPos  As Long
        End Type
        
        #If VBA7 Then
            Private Declare PtrSafe Function timeGetSystemTime Lib "winmm.dll" _
                (lpTime As MMTIME, ByVal uSize As Long) As Long
        #Else
            Private Declare Function timeGetSystemTime Lib "winmm.dll" Alias _
            "timeGetSystemTime" (lpTime As MMTIME, ByVal uSize As Long) As Long
        #End If
    
    'QUERYPERFORMANCECOUNTER API
        #If VBA7 Then
            Private Declare PtrSafe Function GetFrequency Lib "kernel32" _
                Alias "QueryPerformanceFrequency" (ByRef Frequency As _
                Currency) As LongPtr
            Private Declare PtrSafe Function GetTime Lib "kernel32" _
                Alias "QueryPerformanceCounter" (ByRef Counter As Currency) _
                As LongPtr
        #Else
            Private Declare  Function GetFrequency Lib "kernel32" _
                Alias "QueryPerformanceFrequency" (ByRef Frequency As _
                Currency) As Long
            Private Declare Function GetTime Lib "kernel32" _
                Alias "QueryPerformanceCounter" (ByRef Counter As Currency) _
                As Long
        #End If

    'SLEEP
        #If VBA7 Then
            Private Declare PtrSafe Sub Sleep Lib "kernel32" ( _
                        ByVal dwMilliseconds As LongPtr)
        #Else
            Private Declare Sub Sleep Lib "kernel32" ( _
                        ByVal dwMilliseconds As Long)
        #End If
    '
    '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    'VARIABLES
    '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    '
    'Timer, GetTickCount, TimeGetTime, TimeGetSystemTime method
        Private T1              As Double       'Start Time
        Private T2              As Double       'End time
        Private ET              As Double       'Elapsed Time
    'QueryPerformance method
        Dim ccyStartTime        As Currency
        Dim ccyEndTime          As Currency
        Dim ccyFrequency        As Currency
        Dim ccyTimeElapsed      As Currency
Private Sub Class_Initialize()
    GetFrequency ccyFrequency
End Sub
Public Sub StartCounter(Optional Method As Integer = 4)
'------------------------------------------------------------------------------
'ROUTINE
'------------------------------------------------------------------------------
    Select Case Method
        '------------------------------
        'VBA.TIMER
        '------------------------------
            'IT RETURNS THE NUMBER OF SECONDS ELAPSED SINCE MIDNIGHT using a
            'single-precision floating point value. It is not a threaded
            'control so, it may not be triggered at exactly the correct time.
            'Delays can be caused by other applications or system processes.
        Case 1
            T1 = Timer
        
        '------------------------------
        'GETTICKCOUNT API
        '------------------------------
            'IT RETURNS THE NUMBER OF MILLISECONDS ELAPSED SINCE WINDOWS
            'WAS STARTED
            'It will run for 49 days before resetting back to zero.
        Case 2
            T1 = GetTickCount
        
        '------------------------------
        'TIMEGETTIME API
        '------------------------------
            'IT RETURNS THE NUMBER OF MILLISECONDS ELAPSED SINCE WINDOWS
            'WAS STARTED
        Case 3
            T1 = timeGetTime()
        
        '------------------------------
        'QUERYPERFORMANCECOUNTER API
        '------------------------------
            'QueryPerformanceCounter returns the current value of the computer’s
            'performance counter. This will be the whole number representing the
            'number of “counts” (similar to the number of ticks on a very fast
            'clock).
            'QueryPerformanceFrequency returns the number of “counts” there
            'are per second, which is fixed at system boot.
            'Both of these functions take a LARGE_INTEGER as an argument,
            'which is a 64-bit integer. VBA doesn’t have a 64-bit integer
            '(the VBA long integer has only 32-bits). Luckily, we can use the
            'VBA currency data type, which is 64-bit.
        Case 4
            GetTime ccyStartTime
            
        '------------------------------
        'TIMEGETSYSTEMTIME API
        '------------------------------
            'IT RETURNS THE CURRENT TIME IN MILLISECONDS
        Case 5
            T1 = GetCurrentTime()
            
        'WRONG INPUT
        Case Else
            Exit Sub
    End Select
End Sub
Public Property Get TimeElapsed(Optional iMethod As Integer = 5) As Variant
'------------------------------------------------------------------------------
'ROUTINE
'------------------------------------------------------------------------------
    Select Case iMethod
        
        'TIMER
        Case 1
            T2 = Timer
            ET = T2 - T1 + Abs((T2 < T1) * 86400)
        'TICKCOUNT
        Case 2
            T2 = GetTickCount
            If T2 < T1 Then
            'If T2 is smaller than T1 then the tickcount has been reset, use
            'reset Tick value + T2 - T1
                ET = ((4294967295# + T2) - T1)
             Else
                'If T2 is bigger than T1 then just use T2 - T1
                ET = (T2 - T1) / 1000
            End If
        
        'TIMEGETTIME
        Case 3
            T2 = timeGetTime()
            ET = (T2 - T1) / 1000
        
        'TIMEGESTSYSTEMTIME
        Case 4
            T2 = GetCurrentTime
            ET = (T2 - T1) / 1000
        
        'QUERYPERFORMANCE
        Case 5
            GetTime ccyEndTime
            ccyTimeElapsed = ((ccyEndTime - ccyStartTime) * 1000 / _
                ccyFrequency)
            ET = ccyTimeElapsed / 1000
        

            
        'WRONG INPUT
        Case Else
            Exit Property
    
    End Select
'------------------------------------------------------------------------------
'ASSIGN RESULT
'------------------------------------------------------------------------------
    TimeElapsed = Strings.Format(Int(ET) / (86400), "HH:mm:ss") & " and " & _
            Strings.Right(Strings.Format(ET, "#0.000"), 3) & " ms"
End Property
Private Function GetCurrentTime() As Long
    'TIMEGESTSYSTEMTIME
    Const TIME_MS = 1
    Dim Mmt As MMTIME
    Mmt.wType = TIME_MS
    timeGetSystemTime Mmt, LenB(Mmt)
    GetCurrentTime = Mmt.Units
End Function
Public Sub Pause(ByVal lSeconds As Byte, Optional iMethod As Integer = 1)
'
'This SUB can be used for testing
'The number of seconds is limited by the data type (0-256)
'
'------------------------------------------------------------------------------
'CONTROL
'------------------------------------------------------------------------------
    If iMethod < 1 Or iMethod > 3 Then GoTo Proc_Err
'------------------------------------------------------------------------------
'INITIALIZE
'------------------------------------------------------------------------------
    On Error GoTo Proc_Err
'------------------------------------------------------------------------------
'MAIN ROUTINE
'------------------------------------------------------------------------------
    Select Case iMethod
        Case 1
        '------------------------------
        'SLEEP API
        '------------------------------
            'Sleep statement depends on the processor’s ticks to calculate the
            'time delays, which may vary slightly on different machines.
            'The advantage of the Sleep statement over the Wait function is that
            'it is more flexible as you can give the time delays in milliseconds.
            'SLEEP releases the CPU and allows other applications to run on real
            'multitask systems
        Sleep lSeconds * 1000
        '------------------------------
        'APPLICATION.WAIT
        '------------------------------
            'The Wait function can only delay the application by whole seconds
            'Does not use API functions
        Case 2
            Application.Wait DateAdd("s", lSeconds, Now)
        '------------------------------
        'WORKAROUND
        '------------------------------
            'Does not use API
            'Does not use built-in Wait function
        Case 3
            Dim WaitTill As Double
            WaitTill = (Now + TimeValue("0:00:" & lSeconds))
            While Now() < WaitTill
                DoEvents
            Wend
    End Select
'------------------------------------------------------------------------------
'EXIT
'------------------------------------------------------------------------------
Proc_Exit:
    Exit Sub
'------------------------------------------------------------------------------
'HANDLE ERRORS
'------------------------------------------------------------------------------
Proc_Err:
    Debug.Print "???"
End Sub
'
'..............................................................................
'                                 TEST
'..............................................................................
'
'Copy this routine in a standard module
Public Sub TestPerformance()
'------------------------------------------------------------------------------
'DECLARE
'------------------------------------------------------------------------------
    Dim cPM As cPerformanceMonitor
    Dim i As Long
'------------------------------------------------------------------------------
'INITIALIZE
'------------------------------------------------------------------------------
    Set cPM = New cPerformanceMonitor
'------------------------------------------------------------------------------
'TEST CLASS MODULE
'------------------------------------------------------------------------------
    cPM.StartCounter                    'Default 5
        cPM.Pause 1                     '1 second
    Debug.Print cPM.TimeElapsed
'------------------------------------------------------------------------------
'TEST YOUR CODE
'------------------------------------------------------------------------------
    cPM.StartCounter (1)
        'Your code here (example)
        For i = 1 To 10000000: Next i   'Empty loop
    Debug.Print cPM.TimeElapsed(1)
'------------------------------------------------------------------------------
'EXIT
'------------------------------------------------------------------------------
    Set cPM = Nothing
End Sub
'
'//////////////////////////////////////////////////////////////////////////////
'                               RELEASE NOTE
'//////////////////////////////////////////////////////////////////////////////
'
'1.0.1 24/01/2023
'   - Added example for testing
'
