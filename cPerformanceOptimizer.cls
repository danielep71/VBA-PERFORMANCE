VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cPerformanceOptimizer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
'                        CLASS PERFORMANCE OPTIMIZER                          =
'==============================================================================
'
'************
'DESCRIPTION*   Speed up VBA code by disabling time-wasters at the beginning
'************   of the macro and re-enabling them at the end.
'
'               1) Turn-off Automatic Calculations
'               2) Disable Screen Updating
'               3) Disable Event Notifications
'               4) Disable Alerts
'               5) Set the cursor in wait status
'               6) Manage Cancel key
'
'************
'AUTHOR     *   Daniele Penza (adapted from Juan Pablo Gonzàlez Ruiz of Bogotà)
'************
'
'************
'VERSION    *   1.0.0
'************
'
'************
'LAST UPDATE*   05/02/2023
'************
'
'==============================================================================
'************
'PROCEDURES *   FUNCTIONS AND SUBROUTINES
'************
'
'  1. PRI - SUB - Class_Initialize
'  2. PUB - PRO - Class_Terminate
'  3. PUB - SUB - Optimize
'
'==============================================================================
'
'------------------------------------------------------------------------------
'MODULE SETTINGS
'------------------------------------------------------------------------------
'
    Option Explicit             'Force to declare all variables
'------------------------------------------------------------------------------
'MODULE DECLARATIONS
'------------------------------------------------------------------------------
    Private m_SU        As Boolean  'Screen Updating
    Private m_EE        As Boolean  'Enable Events
    Private m_DA        As Boolean  'Display Alerts
    Private m_CALC      As Long     'Calculation mode
    Private m_CURSOR    As Long     'Cursor state
    Private m_CANCELKEY As Long     'EnableCancelKey
    Private m_Except    As PO_Enum  'Exceptions
    
    Public Enum PO_Enum
        None = 0
        Calculation = 1
        ScreenUpdating = 2
        EnableEvents = 2
        DisplayAlerts = 4
        Cursor = 16
        CancelKey = 32
    End Enum
'
Private Sub Class_Initialize()
'Save current settings
'------------------------------------------------------------------------------
    With Application
        m_SU = .ScreenUpdating
        m_EE = .EnableEvents
        m_DA = .DisplayAlerts
        m_CALC = .Calculation
        m_CURSOR = .Cursor
        m_CANCELKEY = .EnableCancelKey
    End With
End Sub
'
Private Sub Class_Terminate()
'Restore previous settings
'------------------------------------------------------------------------------
    With Application
        If Not m_Except And PO_Enum.Calculation Then
            .Calculation = m_CALC
        End If
        If Not m_Except And PO_Enum.ScreenUpdating Then
            .ScreenUpdating = m_SU
        End If
        If Not m_Except And PO_Enum.EnableEvents Then
            .EnableEvents = m_EE
        End If
        If Not m_Except And PO_Enum.DisplayAlerts Then
            .DisplayAlerts = m_DA
        End If
        If Not m_Except And PO_Enum.Cursor Then
            .Cursor = m_CURSOR
        End If
        If Not m_CANCELKEY And PO_Enum.m_CANCELKEY Then
            .EnableCancelKey = m_CANCELKEY
        End If
    End With
End Sub
'
Public Sub Optimize(Optional ByVal Except As PO_Enum = PO_Enum.None)
'Initialize
'------------------------------------------------------------------------------
    m_Except = Except
'Main routine
'------------------------------------------------------------------------------
    With Application
        If Not m_Except And PO_Enum.Calculation Then
            .Calculation = xlCalculationManual
        End If
        If Not m_Except And PO_Enum.ScreenUpdating Then
            .ScreenUpdating = False
        End If
        If Not m_Except And PO_Enum.EnableEvents Then
            .EnableEvents = False
        End If
        If Not m_Except And PO_Enum.DisplayAlerts Then
            .DisplayAlerts = False
        End If

        If Not m_Except And PO_Enum.Cursor Then
            .Cursor = xlWait
        End If
        If Not m_CANCELKEY And PO_Enum.m_CANCELKEY Then
            .EnableCancelKey = xlInterrupt
        End If
    End With
End Sub
'
'//////////////////////////////////////////////////////////////////////////////
'                               HOW TO USE
'//////////////////////////////////////////////////////////////////////////////
'
'Copy this routine in a standard module and remove comment block
'______________________________________________________________________________
'
' Sub RunCode()
'     Dim cPO As cPerformanceOptimizer          'Declare variable
'     Set cPO = New cPerformanceOptimizer       'Create object
'     cPO.Optimize                              'Run Subroutine
'       'Your code here
'     Set cPO = Nothing                         'Destroy the object
' End Sub
