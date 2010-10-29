VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CStopWatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
' Win32 API declarations.
'
Private Declare Function timeGetTime Lib "winmm.dll" () As Long
Private Declare Function timeGetDevCaps Lib "winmm.dll" (lpTimeCaps As TIMECAPS, ByVal uSize As Long) As Long
'
' API Structure definitions.
'
Private Type TIMECAPS
   wPeriodMin As Long
   wPeriodMax As Long
End Type
'
' Set aside storage for private member variables.
'
Private m_StartTime As Long
Private m_PeriodMin As Long
Private m_PeriodMax As Long

' ********************************************
'  Initialize
' ********************************************
Private Sub Class_Initialize()
   '
   ' Retrieve system timer resolution.
   '
   Dim tc As TIMECAPS
   Call timeGetDevCaps(tc, Len(tc))
   m_PeriodMin = tc.wPeriodMin
   m_PeriodMax = tc.wPeriodMax
   '
   ' Initialize starting time.
   '
   m_StartTime = timeGetTime()
End Sub

' ********************************************
'  Public Properties
' ********************************************
Public Property Get Elapsed() As Long
   '
   ' Read-Only: return elapsed time in milliseconds
   '            since stopwatch was reset.
   '
   Elapsed = timeGetTime() - m_StartTime
End Property

Public Property Get MinimumResolution() As Long
   '
   ' Read-Only: return minimum number of milliseconds
   '            timer is capable of resolving.
   '
   MinimumResolution = m_PeriodMin
End Property

Public Property Get MaximumResolution() As Long
   '
   ' Read-Only: return maximum number of milliseconds
   '            timer is capable of resolving.
   '
   MaximumResolution = m_PeriodMax
End Property

' ********************************************
'  Public Methods
' ********************************************
Public Sub Reset()
   '
   ' Reinitialize starting time.
   '
   m_StartTime = timeGetTime()
End Sub
