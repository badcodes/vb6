VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CExecutive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum EWaitMode
    ewmWaitIdle = -1
    ewmNoWait
    ewmWaitDead
End Enum

Public Enum EErrorExecutive
    eeBaseExecutive = 13060     ' CExecutive
End Enum

Private proc As PROCESS_INFORMATION
Private start As STARTUPINFO
Private sInitDir As String
Private sProg As String
Private iExit As Long
Private ewm As EWaitMode
Private sPipedInText As String, hReadStdIn As Long, hWriteStdIn As Long
Private sPipedOutText As String, hReadStdOut As Long, hWriteStdOut As Long
Private sPipedErrText As String, hReadStdErr As Long, hWriteStdErr As Long

Private Sub Class_Initialize()
    Reset
End Sub

Private Sub Class_Terminate()
    CloseAll
End Sub

Private Sub CloseAll()
    CloseHandleNull hReadStdIn
    CloseHandleNull hWriteStdIn
    CloseHandleNull hReadStdOut
    CloseHandleNull hReadStdErr
End Sub

Sub Reset()
    With start
        .cb = LenB(start)
        .dwFlags = 0
        .lpTitle = sNullStr
        .dwX = -1
        .dwY = -1
        .dwXSize = -1
        .dwYSize = -1
        .dwFillAttribute = -1
        .dwXCountChars = -1
        .dwYCountChars = -1
        .wShowWindow = -1
    End With
    With proc
        .dwProcessId = 0
        .dwThreadId = 0
        .hProcess = 0
        .hThread = 0
    End With
    sInitDir = sNullStr
    iExit = -1
    ewm = ewmNoWait
    CloseAll
End Sub

Property Get WaitMode() As EWaitMode
    WaitMode = ewm
End Property

Property Let WaitMode(ByVal ewmA As EWaitMode)
    Select Case ewmA
    Case ewmWaitDead, ewmWaitIdle
        ewm = ewmA
    Case Else ' Unrecognized same as ewmNoWait
        ewm = ewmNoWait
    End Select
End Property

Property Get ProcessID() As Long
    ProcessID = proc.dwProcessId
End Property

Property Get ThreadID() As Long
    ThreadID = proc.dwThreadId
End Property

Property Get ExitCode() As Long
    If ewm <> ewmNoWait Then
        ' If a WaitMode was on, we already have the exit code
        ExitCode = iExit
    Else
        ' Otherwise open a handle and get exit code
        Dim hProcess As Long
        hProcess = OpenProcess(PROCESS_SET_INFORMATION, False, _
                               proc.dwProcessId)
        GetExitCodeProcess hProcess, ExitCode
    End If
End Property

Property Get Completed() As Boolean
    Completed = (ExitCode <> STILL_ACTIVE)
End Property

Property Get Show() As VbAppWinStyle
    Show = start.wShowWindow
End Property

Property Let Show(eswShow As VbAppWinStyle)
With start
    .wShowWindow = eswShow
    If eswShow = -1 Then
        .dwFlags = .dwFlags And (Not STARTF_USESHOWWINDOW)
    Else
        .dwFlags = .dwFlags Or STARTF_USESHOWWINDOW
    End If
End With
End Property

Property Get Title() As String
    If UnicodeTypeLib Then
        Title = start.lpTitle
    Else
        Title = StrConv(start.lpTitle, vbUnicode)
    End If
End Property

Property Let Title(sTitle As String)
    If UnicodeTypeLib Then
        start.lpTitle = sTitle
    Else
        start.lpTitle = StrConv(sTitle, vbFromUnicode)
    End If
End Property

Property Get InitDir() As String
    InitDir = sInitDir
End Property

Property Let InitDir(sInitDirA As String)
    If MUtility.ExistFileDir(sInitDirA) Then sInitDir = sInitDirA
End Property

Property Get Left() As Long
    Left = start.dwX
End Property

Property Let Left(ByVal xLeft As Long)
With start
    .dwX = xLeft
    If .dwX = -1 Then
        .dwY = -1
        .dwFlags = .dwFlags And (Not STARTF_USEPOSITION)
    Else
        .dwFlags = .dwFlags Or STARTF_USEPOSITION
    End If
End With
End Property

Property Get Top() As Long
    Top = start.dwY
End Property

Property Let Top(ByVal yTop As Long)
With start
    .dwY = yTop
    If .dwX = -1 Then
        .dwY = -1
        .dwFlags = .dwFlags And (Not STARTF_USEPOSITION)
    Else
        .dwFlags = .dwFlags Or STARTF_USEPOSITION
    End If
End With
End Property

Property Get Width() As Long
    Width = start.dwXSize
End Property

Property Let Width(ByVal xWidth As Long)
With start
    .dwXSize = xWidth
    If .dwXSize = -1 Then
        .dwYSize = -1
        .dwFlags = .dwFlags And (Not STARTF_USESIZE)
    Else
        .dwFlags = .dwFlags Or STARTF_USESIZE
    End If
End With
End Property

Property Get Height() As Long
    Height = start.dwYSize
End Property

Property Let Height(ByVal yHeight As Long)
With start
    .dwYSize = yHeight
    If .dwYSize = -1 Then
        .dwXSize = -1
        .dwFlags = .dwFlags And (Not STARTF_USESIZE)
    Else
        .dwFlags = .dwFlags Or STARTF_USESIZE
    End If
End With
End Property

Property Get Columns() As Long
    Columns = start.dwXCountChars
End Property

Property Let Columns(ByVal xColumns As Long)
With start
    .dwXCountChars = xColumns
    If .dwXCountChars = -1 Then
        .dwYCountChars = -1
        .dwFlags = .dwFlags And (Not STARTF_USECOUNTCHARS)
    Else
        .dwFlags = .dwFlags Or STARTF_USECOUNTCHARS
    End If
End With
End Property

Property Get Rows() As Long
    Rows = start.dwYCountChars
End Property

Property Let Rows(ByVal yRows As Long)
With start
    .dwYCountChars = yRows
    If .dwYCountChars = -1 Then
        .dwXCountChars = -1
        .dwFlags = .dwFlags And (Not STARTF_USECOUNTCHARS)
    Else
        .dwFlags = .dwFlags Or STARTF_USECOUNTCHARS
    End If
End With
End Property

Property Get BackColor() As Long
    BackColor = MBytes.LoByte(start.dwFillAttribute)
End Property

Property Let BackColor(ByVal atrBackColor As Long)
With start
    If atrBackColor = -1 Then
        .dwFillAttribute = -1
        .dwFlags = .dwFlags And (Not STARTF_USEFILLATTRIBUTE)
    Else
        atrBackColor = MBytes.LShiftWord(atrBackColor, 4)
        .dwFillAttribute = .dwFillAttribute And &HF Or atrBackColor
        .dwFlags = .dwFlags Or STARTF_USEFILLATTRIBUTE
    End If
End With
End Property

Property Get ForeColor() As Long
    ForeColor = MBytes.HiByte(MBytes.LoWord(start.dwFillAttribute))
End Property

Property Let ForeColor(ByVal atrForeColor As Long)
With start
    If atrForeColor = -1 Then
        .dwFillAttribute = -1
        .dwFlags = .dwFlags And (Not STARTF_USEFILLATTRIBUTE)
    Else
        .dwFillAttribute = .dwFillAttribute And &HF0 Or atrForeColor
        .dwFlags = .dwFlags Or STARTF_USEFILLATTRIBUTE
    End If
End With
End Property

Property Get FullScreen() As Boolean
    FullScreen = start.dwFlags And STARTF_RUNFULLSCREEN
End Property

Property Let FullScreen(fFullScreen As Boolean)
With start
    If fFullScreen Then
        .dwFlags = .dwFlags Or STARTF_RUNFULLSCREEN
    Else
        .dwFlags = .dwFlags And (Not STARTF_RUNFULLSCREEN)
    End If
End With
End Property

Property Get IsWindowed() As Boolean
    Dim ept As EProgramType
    ept = MExeType.ExeType(ProgPath)
    IsWindowed = ept <> eptMSDOS And ept <> eptWin32Console And _
                 ept <> eptDOSUnknown And ept <> eptOS2_1
End Property

Property Get ProgName() As String
    ProgName = MUtility.GetFileBaseExt(MUtility.SearchForExe(sProg))
End Property

Property Get ProgPath() As String
    ProgPath = MUtility.GetFullPath(MUtility.SearchForExe(sProg))
End Property

Property Get PipedInText() As String
    PipedInText = sPipedInText
End Property

Property Let PipedInText(sPipedInTextA As String)
    sPipedInText = sPipedInTextA
    ' Close any open handles
    CloseHandleNull hReadStdIn
    CloseHandleNull hWriteStdIn
End Property

Property Get PipedOutText() As String
    PipedOutText = sPipedOutText
End Property

Property Get PipedErrText() As String
    PipedErrText = sPipedErrText
End Property

Sub Run(sCmd As String)

    ' Process any environment variables
    Dim sCmdLine As String, sPipeOut As String, sPipeErr As String
    sCmdLine = MUtility.ExpandEnvStr(sCmd)
    sProg = MParse.GetQToken(sCmdLine, " ")
    
    ' Create standard input, output, and error pipes
    CreatePipes
    
    ' Create process and run it
    If CreateProcess(sNullStr, sCmdLine, ByVal pNull, ByVal pNull, _
                     APITRUE, 0&, pNull, sInitDir, start, proc) Then
        
        ' Must close write end of out and err handles before you can read
        CloseHandleNull hWriteStdOut
        CloseHandleNull hWriteStdErr
        
        Select Case ewm
        Case ewmWaitIdle
            ' Wait, but allow painting and other processing
            Do
                GetExitCodeProcess proc.hProcess, iExit
                DoEvents
            Loop Until ReadPipeChunk And ReadPipeErrChunk And Completed
        Case ewmWaitDead
            ' Stop dead until process terminates
            Dim iResult As Long
            iResult = WaitForSingleObject(proc.hProcess, INFINITE)
            If iResult = WAIT_FAILED Then ErrRaise Err.LastDllError
            ' Get the return value
            GetExitCodeProcess proc.hProcess, iExit
            Do
            Loop Until ReadPipeChunk And ReadPipeErrChunk And Completed
        Case Else
            ' Caller must call use ExitCode and pipe chunks directly
        End Select
        CloseHandleNull proc.hProcess
        CloseHandleNull proc.hThread
    Else
        ApiRaise Err.LastDllError
    End If
End Sub
'

Private Sub CloseHandleNull(h As Long)
    If h <> hNull Then CloseHandle h
    h = hNull
End Sub

Sub CreatePipes()
With start
    Dim saPipe As SECURITY_ATTRIBUTES, f As Long, c As Long
    saPipe.nLength = LenB(saPipe)
    saPipe.bInheritHandle = APITRUE
    saPipe.lpSecurityDescriptor = pNull
    .dwFlags = .dwFlags Or STARTF_USESTDHANDLES
    ' Create anonymous pipe for standard output
    f = CreatePipe(hReadStdOut, hWriteStdOut, saPipe, 0)
    If f = False Then ApiRaise Err.LastDllError
    .hStdOutput = hWriteStdOut
    ' Create anonymous pipe for standard error
    f = CreatePipe(hReadStdErr, hWriteStdErr, saPipe, 0)
    If f = False Then ApiRaise Err.LastDllError
    .hStdError = hWriteStdErr
    If sPipedInText = sEmpty Then
        .hStdInput = hNull
    Else
        ' Create anonymous pipe for standard input
        f = CreatePipe(hReadStdIn, hWriteStdIn, saPipe, 0)
        If f = False Then ApiRaise Err.LastDllError
        .hStdInput = hReadStdIn
        ' Write input string to handle
        Dim abPipedInText() As Byte
        MBytes.StrToBytes abPipedInText, sPipedInText
        f = WriteFile(hWriteStdIn, abPipedInText(0), _
                      MBytes.LenBytes(abPipedInText), c, ByVal pNull)
        If f = False Then ApiRaise Err.LastDllError
        CloseHandleNull hWriteStdIn
    End If
End With
End Sub

Function ReadPipeChunk(Optional sChunk As String, _
                       Optional ByVal cWant As Long = 512) As Boolean
    Dim f As Long, cGot As Long, cPeek As Long, fDone As Boolean, abChunk() As Byte
    ' See if program is done
    fDone = Completed
    ' See if there's more in the pipe
    f = PeekNamedPipe(hReadStdOut, pNull, 0, 0, cPeek, 0)
    If (f <> 0) And (cPeek <> 0) Then
        ' Read a chunk of bytes
        ReDim abChunk(0 To cWant - 1)
        Call ReadFile(hReadStdOut, abChunk(0), cWant, cGot, ByVal pNull)
        sChunk = MBytes.LeftBytes(abChunk, cGot)
        sPipedOutText = sPipedOutText & sChunk
    End If
    ' Strange difference between WinNT and Win95
    If MUtility.IsNT Then
        If (cGot = 0) And fDone Then ReadPipeChunk = True
    Else
        If (cPeek = 0) And fDone Then ReadPipeChunk = True
    End If
End Function

Function ReadPipeErrChunk(Optional sChunk As String, _
                          Optional ByVal cWant As Long = 512) As Boolean
    Dim f As Long, cGot As Long, cPeek As Long, fDone As Boolean, abChunk() As Byte
    ' See if program is done
    fDone = Completed
    ' See if there's more in the pipe
    f = PeekNamedPipe(hReadStdErr, pNull, 0, 0, cPeek, 0)
    If (f <> 0) And (cPeek <> 0) Then
        ' Read a chunk of bytes
        ReDim abChunk(0 To cWant - 1)
        Call ReadFile(hReadStdErr, abChunk(0), cWant, cGot, ByVal pNull)
        sChunk = MBytes.LeftBytes(abChunk, cGot)
        sPipedErrText = sPipedErrText & sChunk
    End If
    ' Strange difference between WinNT and Win95
    If MUtility.IsNT Then
        If (cGot = 0) And fDone Then ReadPipeErrChunk = True
    Else
        If (cPeek = 0) And fDone Then ReadPipeErrChunk = True
    End If
End Function

#If fComponent = 0 Then
Private Sub ErrRaise(e As Long)
    Dim sText As String, sSource As String
    If e > 1000 Then
        sSource = App.ExeName & ".Executive"
        Select Case e
        Case eeBaseExecutive
            BugAssert True
       ' Case ee...
       '     Add additional errors
        End Select
        Err.Raise COMError(e), sSource, sText
    Else
        ' Raise standard Visual Basic error
        sSource = App.ExeName & ".VBError"
        Err.Raise e, sSource
    End If
End Sub
#End If

