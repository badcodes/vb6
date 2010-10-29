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

Function TimerCreate(timer As CTimer) As Boolean
    ' Make sure there's room
    If cTimers + 1 = cTimerMax Then timer.ErrRaise eeTooManyTimers
    ' Create the timer
    timer.TimerID = SetTimer(0&, 0&, timer.Interval, AddressOf TimerProc)
    If timer.TimerID Then
        TimerCreate = True
        cTimers = cTimers + 1
        atdata(cTimers).idTimer = timer.TimerID
        atdata(cTimers).pTimer = ObjPtr(timer)
    Else
        ' TimerCreate = False
        timer.TimerID = 0
        timer.Interval = 0
    End If
    
End Function

Public Function TimerDestroy(timer As CTimer) As Long
    ' TimerDestroy = False
    ' Find and remove this timer
    Dim i As Long, iDead As Long, tdata As TTimerData ' = zeros
    For i = 1 To cTimers
        ' Find timer in array
        If timer.TimerID = atdata(i).idTimer Then
            ' Kill the timer
            Call KillTimer(hNull, timer.TimerID)
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


Sub TimerProc(ByVal hWnd As Long, ByVal uMsg As Long, _
              ByVal idEvent As Long, ByVal dwTime As Long)
    Dim i As Integer, timer As CTimer
    ' Find the timer with this ID
    For i = 1 To cTimers
        If idEvent = atdata(i).idTimer Then
            ' First create a timer from the unreference timer pointer
            CopyMemory timer, atdata(i).pTimer, 4
            ' Generate the event
            timer.PulseTimer
            ' Destroy the temporary timer so it won't be reference counted
            CopyMemory timer, 0&, 4
            Exit Sub
        End If
    Next
    
End Sub



