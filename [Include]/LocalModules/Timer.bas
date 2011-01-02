Attribute VB_Name = "MTimer"
Option Explicit

' Bug reported by Steve McMahon fixed 2/28/98. TimerDestroy left blank
' entries in the timer object array.

' Bug reported by Andy Hopper fixed 10/1/90. Two copies of each CTimer--one
' in internal array and one created by user. If user tried to destroy public
' one, the internal one would live on and never be destroyed. Fix by storing
' internal version as an object pointer.

Const cTimerMax = 100

Private Type TTimerData
    idTimer  As Long
    pTimer As Long
End Type

' Array of timers
Public atdata(1 To cTimerMax) As TTimerData

Private cTimers As Long

Function TimerCreate(Timer As CTimer) As Boolean
    On Error Resume Next
    ' Make sure there's room
    If cTimers + 1 = cTimerMax Then Timer.ErrRaise eeTooManyTimers
    ' Create the timer
    Timer.TimerID = SetTimer(0&, 0&, Timer.Interval, AddressOf TimerProc)
    If Timer.TimerID Then
        TimerCreate = True
        cTimers = cTimers + 1
        atdata(cTimers).idTimer = Timer.TimerID
        atdata(cTimers).pTimer = ObjPtr(Timer)
    Else
        ' TimerCreate = False
        Timer.TimerID = 0
        Timer.Interval = 0
    End If
    
End Function

Public Function TimerDestroy(Timer As CTimer) As Long
    ' TimerDestroy = False
    ' Find and remove this timer
    On Error Resume Next
    Dim i As Long, iDead As Long, tdata As TTimerData ' = zeros
    For i = 1 To cTimers
        ' Find timer in array
        If Timer.TimerID = atdata(i).idTimer Then
            ' Kill the timer
            Call KillTimer(hNull, Timer.TimerID)
            cTimers = cTimers - 1
            ' Overwrite dead timers so there is no blank space
            ' by moving remaining timers down one space
            For iDead = i To cTimers
                atdata(iDead) = atdata(iDead + 1)
            Next
            atdata(cTimers + 1) = tdata
            TimerDestroy = True
            Exit Function
        End If
    Next
    BugAssert True          ' Should never happen
End Function


Sub TimerProc(ByVal hwnd As Long, ByVal uMsg As Long, _
              ByVal idEvent As Long, ByVal dwTime As Long)
    Dim i As Integer, Timer As CTimer
    ' Find the timer with this ID
    On Error Resume Next
    For i = 1 To cTimers
        If idEvent = atdata(i).idTimer Then
            ' First create a timer from the unreference timer pointer
            CopyMemory Timer, atdata(i).pTimer, 4
            ' Generate the event
            Timer.PulseTimer
            ' Destroy the temporary timer so it won't be reference counted
            CopyMemory Timer, 0&, 4
            Exit Sub
        End If
    Next
    
End Sub



