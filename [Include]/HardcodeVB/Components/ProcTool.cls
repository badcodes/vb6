VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GProcTool"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum EErrorProcTool
    eeBaseProcTool = 13570  ' ProcTool
End Enum

Function TopWndFromProcID(idProcA As Long) As Long
    Dim idProc As Long, hWnd As Long
    
    ' Get first window
    hWnd = GetWindow(GetDesktopWindow(), GW_CHILD)
    Do While hWnd <> hNull
        ' Check instance until it matches
        Dim sTitle As String
        sTitle = MWinTool.WindowTextLineFromWnd(hWnd)
        idProc = MWinTool.ProcIDFromWnd(hWnd)
        If idProcA = idProc Then
            If MWinTool.IsVisibleTopWnd(hWnd) Then Exit Do
        End If

        ' Get next sibling
        hWnd = GetWindow(hWnd, GW_HWNDNEXT)
    Loop
    TopWndFromProcID = hWnd
End Function
    
Public Function GetProcInfo(ByVal ID As Long, Optional TabStop As Integer = 0) As String
    Dim sStart As String, s As String, sTemp As String
    
    ' Nested starting position
    sStart = Space$(TabStop * 4)
    ' Module information
    s = sStart & "Program: " & MModTool.ExeNameFromProcID(ID) & sCrLf
    s = s & sStart & "Module: " & Hex$(MModTool.ModFromProcID(ID)) & sCrLf
    s = s & sStart & "Instance: " & Hex$(MModTool.InstFromProcID(ID)) & sCrLf
    s = s & sStart & "PID: " & ID & sCrLf

    GetProcInfo = s
End Function

' Pass idProg returned by Shell or ShellPlus
Function IsRunning(ByVal idProg As Long, _
                   Optional ExitCode As Long) As Boolean
    Static hProg As Long
    ' Get process handle first time through and save it
    If hProg = hNull Then
        hProg = OpenProcess(PROCESS_QUERY_INFORMATION, False, idProg)
    End If
    If hProg = hNull Then
        ' Invalid idProc because program completed before first call
        ExitCode = 0
    Else
        ' Got a valid handle so use it to check process status
        GetExitCodeProcess hProg, ExitCode
    End If
    If ExitCode = STILL_ACTIVE Then
        IsRunning = True
    Else
        CloseHandle hProg
    End If
End Function

Function WaitOnProgram(ByVal idProg As Long, _
                       Optional ByVal WaitDead As Boolean) As Long
    Dim cRead As Long, iExit As Long, hProg As Long
    ' Get process handle
    hProg = OpenProcess(PROCESS_ALL_ACCESS, False, idProg)
    ' Failure probably means process has terminated or ID was invalid
    If hProg = hNull Then
        WaitOnProgram = -1
        Exit Function
    End If
    
    If WaitDead Then
        ' Stop dead until process terminates
        Dim iResult As Long
        iResult = WaitForSingleObject(hProg, INFINITE)
        If iResult = WAIT_FAILED Then ErrRaise Err.LastDllError
        ' Get the return value
        GetExitCodeProcess hProg, iExit
    Else
        ' Get the return value
        GetExitCodeProcess hProg, iExit
        ' Wait, but allow painting and other processing
        Do While iExit = STILL_ACTIVE
            DoEvents
            GetExitCodeProcess hProg, iExit
        Loop
    End If
    CloseHandle hProg
    WaitOnProgram = iExit
End Function

' Combine foreground and background console color attributes
Function ColorAttr(ByVal atrFore As Byte, ByVal atrBack As Byte) As Long
    ColorAttr = MBytes.LShiftWord((&HF And atrBack), 4) Or (&HF And atrFore)
End Function

Function VBShellExecute(sFile As String, _
                        Optional Args As String, _
                        Optional Show As Long = vbNormalFocus, _
                        Optional InitDir As String, _
                        Optional Verb As String, _
                        Optional hWnd As Long = hNull) As Long
    Dim ID As Long
    ID = ShellExecute(hWnd, Verb, sFile, Args, InitDir, Show)
    ' Translate weird ShellExecute errors into normal errors
    Select Case ID
    Case 0
        ID = ERROR_NOT_ENOUGH_MEMORY
    Case SE_ERR_SHARE                                       ' 26
        ID = ERROR_SHARING_VIOLATION
    Case SE_ERR_ASSOCINCOMPLETE                             ' 27
        ID = ERROR_NO_ASSOCIATION
    Case SE_ERR_DDETIMEOUT, SE_ERR_DDEFAIL, SE_ERR_DDEBUSY  ' 28, 29, 30
        ID = ERROR_DDE_FAIL
    Case SE_ERR_NOASSOC                                     ' 31
        ID = ERROR_NO_ASSOCIATION
    Case SE_ERR_DLLNOTFOUND                                 ' 32
        ID = ERROR_DLL_NOT_FOUND
    Case Is > 32
        VBShellExecute = ID
        Exit Function
    End Select
    ApiRaise ID
End Function

#If fComponent = 0 Then
Private Sub ErrRaise(e As Long)
    Dim sText As String, sSource As String
    If e > 1000 Then
        sSource = App.ExeName & ".ProcTool"
        Select Case e
        Case eeBaseProcTool
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

