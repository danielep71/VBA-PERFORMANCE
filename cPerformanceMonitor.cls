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
'DESCRIPTION*   High-precision timer
'************
'
'************
'AUTHOR     *   Daniele Penza
'************
'
'************
'VERSION    *   1.2.4
'************
'
'************
'LAST UPDATE*   30/01/2023
'************
'
'==============================================================================
'************
'PROCEDURES *   FUNCTIONS AND SUBROUTINES
'************
'
'  1. PRI - SUB - Class_Initialize
'  2. PUB - SUB - StartTimer
'  3. PUB - PRO - ElapsedTime
'  4. PUB - SUB - Pause
'  5. PRI - FUN - NextTick
'  6. PRI - FUN - GetCurrentTime
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
    'API DECLARATIONS
    '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    '
    'COMPILER CONSTANTS:
    '___________________
    '#If VBA7   ==> Excel 2010 or later else Excel 2007 or earlier
    '#If Win64  ==> Excel 64-bit else Excel 32-bit
    '
    'KEYWORDS FOR 32/64-BIT
    '_______________________
    'PtrSafe    ==> After Declare
    'LongPtr    ==> returns LONG for 32-bit and LONGLONG for 64-bit
    '
    
    'GETTICKCOUNT API               - LIB kernel21
        #If VBA7 Then
            #If Win64 Then
                Private Declare PtrSafe Function GetTickCount Lib "kernel32" _
                    Alias "GetTickCount64" () As LongPtr    'LongLong
            #Else
                Private Declare PtrSafe Function GetTickCount Lib "kernel32" () _
                    As LongPtr                           'Long
            #End If
        #Else
            Private Declare Function GetTickCount Lib "kernel32" Alias _
                "timeGetTime"() As Long
        #End If

    'TIMEGETTIME API                - LIB winmm
        #If VBA7 Then
            Private Declare PtrSafe Function timeGetTime Lib "winmm.dll" () As _
                LongPtr
            Private Declare PtrSafe Function timeBeginPeriod Lib "winmm.dll" _
                (ByVal uPeriod As Integer) As Integer
            Private Declare PtrSafe Function timeEndPeriod Lib "winmm.dll" _
                (ByVal uPeriod As Integer) As Integer
        #Else
            Private Declare Function timeGetTime Lib "winmm.dll" Alias _
                "timeGetTime" () As Long
            Private Declare Function timeBeginPeriod Lib "winmm.dll" (ByVal _
                uPeriod As Integer) As Integer
            Private Declare Function timeEndPeriod Lib "winmm.dll" (ByVal _
                uPeriod As Integer) As Integer
        #End If
        
    'TIMEGESTSYSTEMTIME API         - LIB winmm
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
                (lpTime As MMTIME, ByVal uSize As Long) As LongPtr
        #Else
            Private Declare Function timeGetSystemTime Lib "winmm.dll" Alias _
                "timeGetSystemTime" (lpTime As MMTIME, ByVal uSize As Long) As Long
        #End If
    
    'QUERYPERFORMANCECOUNTER API    - LIB kernel32
        #If VBA7 Then
            Private Declare PtrSafe Function GetFrequency Lib "kernel32" Alias _
                "QueryPerformanceFrequency" (ByRef Frequency As Currency) As LongPtr
            Private Declare PtrSafe Function GetTime Lib "kernel32" Alias _
                "QueryPerformanceCounter" (ByRef counter As Currency) As LongPtr
        #Else
            Private Declare  Function GetFrequency Lib "kernel32" Alias _
                "QueryPerformanceFrequency" (ByRef Frequency As Currency) As Long
            Private Declare Function GetTime Lib "kernel32" Alias _
                "QueryPerformanceCounter" (ByRef Counter As Currency) As Long
        #End If

    'LIB user32
        #If VBA7 Then
                Private Declare PtrSafe Function SetTimer Lib "user32" (ByVal _
                    hwnd As LongPtr, ByVal nIDEvent As LongPtr, ByVal uElapse As _
                    Long, ByVal lpTimerFunc As LongPtr) As LongPtr
    
                'Private Declare PtrSafe Function KillTimer Lib "user32" (ByVal _
                    hwnd As LongPtr, ByVal nIDEvent As LongPtr) As Long
                'Private SetTimer_Result As LongPtr  'Result value of SetTimer.
        #Else
                Private Declare Function SetTimer Lib "user32" (ByVal hwnd As _
                    Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal _
                    lpTimerFunc As Long) As Long
                'Private Declare Function KillTimer Lib "user32" (ByVal hwnd As _
                    Long, ByVal nIDEvent As Long) As Long
                'Private SetTimer_Result As Long     'Result value of SetTimer.
        #End If
    
    'SLEEP
        #If VBA7 Then
            Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal _
                dwMilliseconds As LongPtr)
        #Else
            Private Declare Sub Sleep Lib "kernel32" ( ByVal dwMilliseconds As _
                Long)
        #End If
    
    'SYSTEM TIME ADJUSTMENT
        Private Declare PtrSafe Function GetSystemTimeAdjustment Lib "kernel32" _
            (lpTimeAdjustment As Long, lpTimeIncrement As Long, _
            lpTimeAdjustmentDisabled As Long) As Long
    
    
    '
    '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    'VARIABLES
    '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    '
    'Timer, GetTickCount, TimeGetTime, TimeGetSystemTime method
        Private m_T1                As Double       'Start Time
        Private m_T2                As Double       'End time
        Private m_ET                As Double       'Elapsed Time
        Private m_Freq              As Double       'Frequency
    'QueryPerformance method
        Private m_ccyStartTime      As Currency
        Private m_ccyEndTime        As Currency
        Private m_ccyFrequency      As Currency
    'Arrays
        Private arrMethods          As Variant
        Private Mmt                 As MMTIME

'
'______________________________________________________________________________
'
Private Sub Class_Initialize()
    GetFrequency m_ccyFrequency
    m_Freq = m_ccyFrequency
    Mmt.wType = 1       ' 1 Millisecond
    ReDim arrMethods(1 To 6)
    arrMethods(1) = "Timer"
    arrMethods(2) = "GetTickCount"
    arrMethods(3) = "TimeGetTime"
    arrMethods(4) = "TimeGetSystemTime"
    arrMethods(5) = "QPC"
    arrMethods(6) = "Now()"
End Sub

'
'______________________________________________________________________________
'
Public Property Get ElapsedTime(Optional iMethod As Integer = 5) As Variant
'------------------------------------------------------------------------------
'DECLARE
'------------------------------------------------------------------------------
    'Fractions of a second                  Metric name
    '0,000 000 001 [billionth]              nanosecond [ns]
    '0,000 001 [millionth]                  microsecond [µs]
    '0,001 [thousandth]                     millisecond [ms]
    '1.0                                    second [s]
    Dim dSeconds        As Double
    Dim sDecimals       As String
    Dim sMilliSeconds   As String
    Dim sMicroSeconds   As String
    Dim sNanoSeconds    As String
'------------------------------------------------------------------------------
'ROUTINE
'------------------------------------------------------------------------------
    Select Case iMethod
        
        'TIMER
        Case 1
            m_T2 = Timer
            'Manage rollover
            m_ET = m_T2 - m_T1 + Abs((m_T2 < m_T1) * 86400)
        
        'TICKCOUNT
        Case 2
            m_T2 = GetTickCount
            'Manage rollover
            If m_T2 < m_T1 Then
            'If m_T2 is smaller than m_m_T1 then the tickcount has been reset, use
            'reset Tick value + m_T2 - m_m_T1
                m_ET = ((4294967295# + m_T2) - m_T1)
             Else
                'If m_T2 is bigger than m_m_T1 then just use m_T2 - m_T1
                m_ET = (m_T2 - m_T1) / 1000
            End If
        
        'TIMEGETTIME
        Case 3
            m_T2 = timeGetTime()
            m_ET = (m_T2 - m_T1) / 1000
        
        'TIMEGESTSYSTEMTIME
        Case 4
            m_T2 = GetCurrentTime
            m_ET = (m_T2 - m_T1) / 1000
        
        'QUERYPERFORMANCE
        Case 5
            GetTime m_ccyEndTime
            m_T2 = m_ccyEndTime
            m_ET = ((m_T2 - m_T1) / m_Freq)
        
        'NOW
        Case 6
            m_T2 = Now() * 86400
            m_ET = m_T2 - m_T1
    End Select
'------------------------------------------------------------------------------
'ASSIGN RESULT
'------------------------------------------------------------------------------
    dSeconds = Fix(m_ET)
    sDecimals = GetDecimalString(m_ET)
    sMilliSeconds = Mid$(sDecimals, 1, 3)
    sMicroSeconds = Mid$(sDecimals, 4, 3)
    sNanoSeconds = Mid$(sDecimals, 7, 3)
    
    ElapsedTime = Format(dSeconds / (86400), "HH:nn:ss") & " - " & sMilliSeconds _
        & " ms - " & sMicroSeconds & " µs - " & sNanoSeconds & " ns"
End Property
'
'______________________________________________________________________________
'
Public Property Get ET() As Double
    ET = m_ET
End Property
'
'______________________________________________________________________________
'
Public Property Get Overhead(Optional iMethod As Integer = 5) As String
    'The API call itself takes a small amount of time to complete
    StartTimer iMethod
    Sleep 0
    Overhead = "Overhead method - " & arrMethods(iMethod - 1) & " - " & _
        ElapsedTime(iMethod)
End Property
'
'______________________________________________________________________________
'
Public Property Get T1() As Double
    T1 = m_T1
End Property
'
'______________________________________________________________________________
'
Public Property Get T2() As Double
    T2 = m_T2
End Property

'
'______________________________________________________________________________
'
Public Property Get TickInterval() As Variant
    'For kernel32
    Dim lA As Long
    Dim lB As Long
    Dim lC As Long
    GetSystemTimeAdjustment lA, lB, lC
    TickInterval = "Tick Interval = " & lB / 10000 & " milliseconds"
End Property
'
'______________________________________________________________________________
'
Public Property Get HRT_FrequencyPerSecond() As Variant
    HRT_FrequencyPerSecond = "HRT Tick frequency is: " & m_ccyFrequency / 10000
End Property
'
'______________________________________________________________________________
'
Public Property Get HRT_TickInterval() As Variant
    HRT_TickInterval = "HRT Tick interval is: " & 1 / (m_ccyFrequency * 10000)
End Property
'
'______________________________________________________________________________
'
Public Property Get TimerMethod(Idx As Integer) As String
    TimerMethod = arrMethods(Idx)
End Property
'
'______________________________________________________________________________
'
Private Function GetCurrentTime() As Long
    timeGetSystemTime Mmt, LenB(Mmt)
    GetCurrentTime = Mmt.Units
End Function
'
'______________________________________________________________________________
'
Private Function GetDecimalString(ByVal dNumber As Double, Optional iDigits As _
    Integer = 9) As String
    On Error Resume Next
'Declare
    Dim strDecimals     As String
'Calculate decimals, round to iDigits and convert to string
    strDecimals = CStr(Round(dNumber - Fix(dNumber), iDigits)) & String$(2, "0")
'Get only the decimal part and fill with zeros for a lenght = to iDigits
    GetDecimalString = Right$(strDecimals, Len(strDecimals) - 2)
    GetDecimalString = GetDecimalString & String$(iDigits + 4 - Len(strDecimals), _
        "0")
End Function

'
'______________________________________________________________________________
'
Private Function NextTick(Optional iMethod As Integer = 5) As Double
'The starttime may be assigned when the system tick is about to expire
'Your routine could appear to take longer that it actually does.
'That's why we try to wait for the next Timer tick

'Declare
    Dim Tick1           As Double
    Dim Tick2           As Double
'Main Routine
    Select Case iMethod
        Case 1          'TIMER
            Tick1 = GetTickCount
            Do
                Tick2 = GetTickCount
                DoEvents
            Loop While Tick1 = GetTickCount
            NextTick = Timer

            
        Case 2          'GET TICK COUNT
            Tick1 = GetTickCount
            Do
                Tick2 = GetTickCount
            Loop Until Tick2 <> Tick1
            NextTick = Tick2
        
        Case 3          'TIME GET TIME
            timeBeginPeriod 1
'            Tick1 = timeGetTime()
'            Do Until Tick2 <> Tick1: Tick2 = timeGetTime(): Loop
            NextTick = timeGetTime
        Case 4          'TIME GET TIME
            Tick1 = GetCurrentTime()
            Do
                Tick2 = GetCurrentTime()
            Loop Until Tick2 <> Tick1
            NextTick = Tick2
        Case 6          'NOW()
            NextTick = Now() * 86400 'Now() is in 'days'
    End Select
End Function

'
'______________________________________________________________________________
'
Private Function NextTickQPC() As Currency
'Declare
    Dim Tick1           As Currency
    Dim Tick2           As Currency
    
    GetTime Tick1
    Do
        GetTime Tick2
    Loop Until Tick2 > Tick1
    NextTickQPC = Tick2
End Function

'
'______________________________________________________________________________
'
Public Sub Pause(ByVal dSeconds As Double, Optional iMethod As Integer = 1)
'
'This routine will pause the system for a given number of milliseconds.
'The pause is limited to 1 hour
'Can be used for testing
'
    'Too long
    If dSeconds > 3600 Then Exit Sub
    
    Select Case iMethod
        Case 1
            'The Sleep function can delay by millisecond
            'Use Sleep API
            Sleep dSeconds * 1000
        Case 2  '
            'The Timer function can only delay by whole seconds
            'Does not use API functions
            Dim PauseTime As Double, Start As Double, Finish As Double
            Dim TotalTime As Double
            PauseTime = dSeconds
            Start = Timer    ' Set start time.
            Do While Timer < Start + PauseTime
                DoEvents    ' Yield to other processes.
            Loop
            Finish = Timer    ' Set end time.
            TotalTime = Finish - Start    ' Calculate total time.
            Debug.Print TotalTime
        Case 3
            'The Wait function can only delay the application by whole seconds
            'Does not use API functions
            Application.Wait DateAdd("s", dSeconds, Now)
        Case 4
            'Does not use API
            'Does not use the built-in Wait function
            Dim WaitTill As Double
            WaitTill = (Now + TimeValue("0:00:" & dSeconds))
            While Now() < WaitTill
                DoEvents
            Wend
    End Select
End Sub

'
'______________________________________________________________________________
'
' *****************************************************************************
' SUB StartTimer
' -----------------------------------------------------------------------------
' Activates the counter
' -----------------------------------------------------------------------------
' INPUT:
'       iMethod     - (Optional). QPC is the default
' -----------------------------------------------------------------------------
' USAGE:
'
' *****************************************************************************
Public Sub StartTimer(Optional iMethod As Integer = 5)
'Select
    If iMethod = 5 Then
        m_ccyStartTime = NextTickQPC
        m_T1 = m_ccyStartTime
    Else
        m_T1 = NextTick(iMethod)
    End If
End Sub

'
'//////////////////////////////////////////////////////////////////////////////
'                                 TEST
'//////////////////////////////////////////////////////////////////////////////
'
'Copy this routine in a standard module
'______________________________________________________________________________
'
Public Sub Test_PerformanceMonitor()
'------------------------------------------------------------------------------
'DECLARE
'------------------------------------------------------------------------------
    Dim cPM     As cPerformanceMonitor
    Dim i       As Integer              'Loop counter
    Dim j       As Long                 'Loop counter
'------------------------------------------------------------------------------
'INITIALIZE
'------------------------------------------------------------------------------
    Set cPM = New cPerformanceMonitor
'------------------------------------------------------------------------------
'TEST CLASS MODULE
'------------------------------------------------------------------------------
    For i = 1 To 20
        cPM.StartTimer (1)
            cPM.Pause 1              '1 second
        Debug.Print "Method " & "1" & " - "; cPM.ElapsedTime(1)
    Next i
'------------------------------------------------------------------------------
'TEST YOUR CODE
'------------------------------------------------------------------------------
    For i = 1 To 5
        cPM.StartTimer (i)
             'Your code here (example)
             For j = 1 To 10000000: Next j   'Empty loop
         Debug.Print cPM.ElapsedTime(i)
    Next i
'------------------------------------------------------------------------------
'TEST OVERHEAD
'------------------------------------------------------------------------------
    For i = 1 To 5
        Debug.Print cPM.Overhead(i)
    Next i
'------------------------------------------------------------------------------
'TEST TICK INTERVAL
'------------------------------------------------------------------------------
    Debug.Print cPM.TickInterval
'------------------------------------------------------------------------------
'TEST HRT TICK FREQUENCY
'------------------------------------------------------------------------------
    Debug.Print cPM.HRT_FrequencyPerSecond
'------------------------------------------------------------------------------
'TEST HRT TICK FREQUENCY
'------------------------------------------------------------------------------
    Debug.Print cPM.HRT_TickInterval
'------------------------------------------------------------------------------
'TEST METHODS NAME
'------------------------------------------------------------------------------
    For i = 1 To 6
        Debug.Print i & " - " & cPM.TimerMethod(i)
    Next i
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
'  1.2.1    - 26/01/2023
'
'           - Added array methods
'           - Property Overhead
'           - Property TickFrequency
'______________________________________________________________________________
'  1.2.0    - 25/01/2023
'
'           - Increased precision to nano-seconds (when available)
'           - Added function GetDecimalString
'           - More comments
'______________________________________________________________________________
'  1.1.0    - 24/01/2023
'
'           - Improved accuracy for Timer method (Function NextTick)
'           - Added example for testing
'______________________________________________________________________________
'







