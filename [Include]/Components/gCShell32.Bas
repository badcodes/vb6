Attribute VB_Name = "MgCShell32"
Option Explicit


' ****************************************************************
'  Shell32.Bas, Copyright ?996-97 Karl E. Peterson
' ****************************************************************
'  You are free to use this code within your own applications, but you
'  are expressly forbidden from selling or otherwise distributing this
'  source code without prior written consent.
' ****************************************************************
'  Three methods to "Shell and Wait" under Win32.
'  One deals with the infamous "Finished" behavior of Win95.
'  Fourth method that simply shells and returns top-level hWnd.
' ****************************************************************

Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Private Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function pShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Const STILL_ACTIVE = &H103
Private Const PROCESS_QUERY_INFORMATION = &H400
Private Const SYNCHRONIZE = &H100000


Private Const WAIT_FAILED = -1&        'Error on call
Private Const WAIT_OBJECT_0 = 0        'Normal completion
Private Const WAIT_ABANDONED = &H80&   '
Private Const WAIT_TIMEOUT = &H102&    'Timeout period elapsed
Private Const IGNORE = 0               'Ignore signal
Private Const INFINITE = -1&           'Infinite timeout

Public Enum SWCMD
 SW_HIDE = 0
 SW_SHOWNORMAL = 1
 SW_SHOWMINIMIZED = 2
 SW_SHOWMAXIMIZED = 3
 SW_SHOWNOACTIVATE = 4
 SW_SHOW = 5
 SW_MINIMIZE = 6
 SW_SHOWMINNOACTIVE = 7
 SW_SHOWNA = 8
 SW_RESTORE = 9
End Enum

Private Const WM_CLOSE = &H10
Private Const GW_HWNDNEXT = 2
Private Const GW_OWNER = 4

Public Function ShellExecute(ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As SWCMD) As Long

    ShellExecute = pShellExecute(hWnd, lpOperation, lpFile, lpParameters, lpDirectory, nShowCmd)

End Function

Public Function ShellAndWait(ByVal JobToDo As String, Optional ExecMode As VbAppWinStyle = vbMinimizedNoFocus, Optional TimeOut As Long = INFINITE) As Long

    ShellAndWait = pShellAndWait(JobToDo, ExecMode, TimeOut)

End Function

Private Function pShellAndWait(ByVal JobToDo As String, Optional ExecMode, Optional TimeOut) As Long

    '
    ' Shells a new process and waits for it to complete.
    ' Calling application is totally non-responsive while
    ' new process executes.
    '
    Dim ProcessID As Long
    Dim hProcess As Long
    Dim nRet As Long
    Const fdwAccess = SYNCHRONIZE

    If IsMissing(ExecMode) Then
        ExecMode = vbMinimizedNoFocus
    Else

        If ExecMode < vbHide Or ExecMode > vbMinimizedNoFocus Then
            ExecMode = vbMinimizedNoFocus
        End If

    End If

    On Error Resume Next
    ProcessID = Shell(JobToDo, CLng(ExecMode))

    If Err Then
        pShellAndWait = vbObjectError + Err.Number
        Exit Function
    End If

    On Error GoTo 0

    If IsMissing(TimeOut) Then
        TimeOut = INFINITE
    End If

    hProcess = OpenProcess(fdwAccess, False, ProcessID)
    nRet = WaitForSingleObject(hProcess, CLng(TimeOut))
    Call CloseHandle(hProcess)

    Select Case nRet
    Case WAIT_TIMEOUT: Debug.Print "Timed out!"
    Case WAIT_OBJECT_0: Debug.Print "Normal completion."
    Case WAIT_ABANDONED: Debug.Print "Wait Abandoned!"
    Case WAIT_FAILED: Debug.Print "Wait Error:"; Err.LastDllError
    End Select

    pShellAndWait = nRet

End Function

Public Function ShellAndLoop(ByVal JobToDo As String, Optional ExecMode As VbAppWinStyle = vbMinimizedNoFocus) As Long

    ShellAndLoop = pShellAndLoop(JobToDo, ExecMode)

End Function

Private Function pShellAndLoop(ByVal JobToDo As String, Optional ExecMode) As Long

    '
    ' Shells a new process and waits for it to complete.
    ' Calling application is responsive while new process
    ' executes. It will react to new events, though execution
    ' of the current thread will not continue.
    '
    Dim ProcessID As Long
    Dim hProcess As Long
    Dim nRet As Long
    Const fdwAccess = PROCESS_QUERY_INFORMATION

    If IsMissing(ExecMode) Then
        ExecMode = vbMinimizedNoFocus
    Else

        If ExecMode < vbHide Or ExecMode > vbMinimizedNoFocus Then
            ExecMode = vbMinimizedNoFocus
        End If

    End If

    On Error Resume Next
    ProcessID = Shell(JobToDo, CLng(ExecMode))

    If Err Then
        pShellAndLoop = vbObjectError + Err.Number
        Exit Function
    End If

    On Error GoTo 0
    hProcess = OpenProcess(fdwAccess, False, ProcessID)

    Do
        GetExitCodeProcess hProcess, nRet
        DoEvents
        Sleep 100
    Loop While nRet = STILL_ACTIVE

    Call CloseHandle(hProcess)
    pShellAndLoop = nRet

End Function

Public Function ShellAndClose(ByVal JobToDo As String, Optional ExecMode As VbAppWinStyle = vbMinimizedNoFocus) As Long

    ShellAndClose = pShellAndClose(JobToDo, ExecMode)

End Function

Private Function pShellAndClose(ByVal JobToDo As String, Optional ExecMode) As Long

    '
    ' Shells a new process and waits for it to complete.
    ' Calling application is responsive while new process
    ' executes. It will react to new events, though execution
    ' of the current thread will not continue.
    '
    ' Will close a DOS box when Win95 doesn't. More overhead
    ' than pShellAndLoop but useful when needed.
    '
    Dim ProcessID As Long
    Dim PID As Long
    Dim hProcess As Long
    Dim hWndJob As Long
    Dim nRet As Long
    Dim TitleTmp As String
    Const fdwAccess = PROCESS_QUERY_INFORMATION

    If IsMissing(ExecMode) Then
        ExecMode = vbMinimizedNoFocus
    Else

        If ExecMode < vbHide Or ExecMode > vbMinimizedNoFocus Then
            ExecMode = vbMinimizedNoFocus
        End If

    End If

    On Error Resume Next
    ProcessID = Shell(JobToDo, CLng(ExecMode))

    If Err Then
        pShellAndClose = vbObjectError + Err.Number
        Exit Function
    End If

    On Error GoTo 0
    hWndJob = FindWindow(vbNullString, vbNullString)

    Do Until hWndJob = 0

        If GetParent(hWndJob) = 0 Then
            Call GetWindowThreadProcessId(hWndJob, PID)

            If PID = ProcessID Then Exit Do
        End If

        hWndJob = GetWindow(hWndJob, GW_HWNDNEXT)
    Loop

    hProcess = OpenProcess(fdwAccess, False, ProcessID)

    Do
        TitleTmp = Space$(256)
        nRet = GetWindowText(hWndJob, TitleTmp, Len(TitleTmp))

        If nRet Then
            TitleTmp = UCase$(Left$(TitleTmp, nRet))

            If InStr(TitleTmp, "FINISHED") = 1 Then
                Call SendMessage(hWndJob, WM_CLOSE, 0, 0)
            End If

        End If

        GetExitCodeProcess hProcess, nRet
        DoEvents
        Sleep 100
    Loop While nRet = STILL_ACTIVE

    Call CloseHandle(hProcess)
    pShellAndClose = nRet

End Function

Public Function hWndShell(ByVal JobToDo As String, Optional ExecMode As VbAppWinStyle = vbMinimizedNoFocus) As Long

    hWndShell = phWndShell(JobToDo, ExecMode)

End Function

Private Function phWndShell(ByVal JobToDo As String, Optional ExecMode) As Long

    '
    ' Shells a new process and returns the hWnd
    ' of its main window.
    '
    Dim ProcessID As Long
    Dim PID As Long
    'Dim hProcess As Long
    Dim hWndJob As Long

    If IsMissing(ExecMode) Then
        ExecMode = vbMinimizedNoFocus
    Else

        If ExecMode < vbHide Or ExecMode > vbMinimizedNoFocus Then
            ExecMode = vbMinimizedNoFocus
        End If

    End If

    On Error Resume Next
    ProcessID = Shell(JobToDo, CLng(ExecMode))

    If Err Then
        phWndShell = 0
        Exit Function
    End If

    On Error GoTo 0
    hWndJob = FindWindow(vbNullString, vbNullString)

    Do While hWndJob <> 0

        If GetParent(hWndJob) = 0 Then
            Call GetWindowThreadProcessId(hWndJob, PID)

            If PID = ProcessID Then
                phWndShell = hWndJob
                Exit Do
            End If

        End If

        hWndJob = GetWindow(hWndJob, GW_HWNDNEXT)
    Loop

End Function
Function GetCommandLine(Optional MaxArgs)
   '声明变量。
   Dim C, CmdLine, CmdLnLen, InArg, I, NumArgs
   '检查是否提供了 MaxArgs 参数。
   If IsMissing(MaxArgs) Then MaxArgs = 10
   ' 使数组的大小合适。
   ReDim ArgArray(MaxArgs)
   NumArgs = 0: InArg = False
   '取得命令行参数。
   CmdLine = Command()
   CmdLnLen = Len(CmdLine)
   '以一次一个字符的方式取出命令行参数。
   For I = 1 To CmdLnLen
      C = Mid(CmdLine, I, 1)
      '检测是否为 space 或 tab。
      If (C <> " " And C <> vbTab) Then
         '若既不是 space 键，也不是 tab 键，
         '则检测是否为参数内含之字符。
         If Not InArg Then
         '新的参数。
         '检测参数是否过多。
            If NumArgs = MaxArgs Then Exit For
               NumArgs = NumArgs + 1
InArg = True
            End If
         '将字符连接到当前参数中。
         ArgArray(NumArgs) = ArgArray(NumArgs) & C
      Else
         '找到 space 或 tab。
         '将 InArg 标志设置成 False。
         InArg = False
      End If
   Next I
   '调整数组大小使其刚好符合参数个数。
   ReDim Preserve ArgArray(NumArgs)
   '将数组返回。
   GetCommandLine = ArgArray()
End Function

