Attribute VB_Name = "MThread"
Option Explicit

Declare Sub ExitThread Lib "KERNEL32" ( _
    ByVal dwExitCode As Long)
    
Declare Sub CloseHandle Lib "KERNEL32" ( _
    ByVal h As Long)
    
Declare Function GetExitCodeThread Lib "KERNEL32" ( _
    ByVal hThread As Long, _
    ByRef lpExitCode As Long) As Long

Declare Function CreateThread Lib "KERNEL32" ( _
    ByRef lpThreadAttributes As Any, _
    ByVal dwStackSize As Long, _
    ByVal lpStartAddress As Long, _
    ByRef lpParameter As Any, _
    ByVal dwCreationFlags As Long, _
    ByRef lpThreadId As Long) As Long

Declare Sub Sleep Lib "KERNEL32" (ByVal dwMilliseconds As Long)

Declare Function GetTickCount Lib "KERNEL32" () As Long

Const STILL_ACTIVE = 259
Const pNull As Long = 0

Private fRunning As Boolean
Private cCalc As Long
Private cAPI As Long
Private datBasic As Date
Private hThread As Long
Private idThread As Long

Sub StartThread(ByVal i As Long)
    ' Signal that thread is starting
    fRunning = True
    ' Create new thread
    hThread = CreateThread(ByVal pNull, 0, AddressOf ThreadProc, _
                           ByVal i, 0, idThread)
    If hThread = 0 Then MsgBox "Can't start thread"
End Sub

Function StopThread() As Long
    ' Signal thread to stop
    fRunning = False
    ' Make sure thread is dead before returning exit code
    Do
        Call GetExitCodeThread(hThread, StopThread)
    Loop While StopThread = STILL_ACTIVE
    CloseHandle hThread
    hThread = 0
End Function

Function ThreadRunning() As Boolean
    ThreadRunning = fRunning
End Function

Function CalcCount() As Long
    CalcCount = cCalc
End Function

Function APICount() As Long
    APICount = cAPI
End Function

Function BasicTime() As Date
    BasicTime = datBasic
End Function

Sub ThreadProc(ByVal i As Long)
    ' Use parameter
    cCalc = i
    Do While fRunning
        ' Calculate something
        cCalc = cCalc + 1
        ' Use an API call
        cAPI = GetTickCount
        ' Use a Basic function
        datBasic = Now
        ' Switch immediately to another thread
        Sleep 1
    Loop
    ' Return a value
    ExitThread cCalc
End Sub
'
