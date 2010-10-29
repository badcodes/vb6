Attribute VB_Name = "MModTool"
Option Explicit

'$ Uses UTILITY.BAS DEBUG.BAS WINTOOL.BAS

'' ToolHelp functions for Windows 95 and Windows NT

' Windows 95 Private Declares, constants, and Private Types
' Use the ToolHelp functions found in KERNEL32.DLL

Public Enum EErrorModTool
    eeBaseModTool = 13530   ' ModTool
End Enum

Const MAX_MODULE_NAME32 = 255

' ****** Shapshot function *****

Private Declare Function CreateToolhelp32Snapshot Lib "KERNEL32" ( _
    ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long

'
' The th32ProcessID argument is only used if TH32CS_SNAPHEAPLIST or
' TH32CS_SNAPMODULE is specified. th32ProcessID == 0 means the current
' process.
'
' NOTE that all of the snapshots are global except for the heap and module
'       lists which are process specific. To enumerate the heap or module
'       state for all WIN32 processes call with TH32CS_SNAPALL and the
'       current process. Then for each process in the TH32CS_SNAPPROCESS
'       list that isn't the current process, do a call with just
'       TH32CS_SNAPHEAPLIST and/or TH32CS_SNAPMODULE.
'
' dwFlags
'
Const TH32CS_SNAPHEAPLIST = &H1&
Const TH32CS_SNAPPROCESS = &H2&
Const TH32CS_SNAPTHREAD = &H4&
Const TH32CS_SNAPMODULE = &H8&
Const TH32CS_SNAPALL = (TH32CS_SNAPHEAPLIST Or TH32CS_SNAPPROCESS Or _
                        TH32CS_SNAPTHREAD Or TH32CS_SNAPMODULE)
Const TH32CS_INHERIT = &H80000000
'
' Use CloseHandle to destroy the snapshot
'

' ****** heap walking ******
#If 0 Then  ' In Visual Basic? No way!

Private Type HEAPLIST32
    dwSize As Long
    th32ProcessID As Long    ' owning process
    th32HeapID As Long       ' heap (in owning process's context!)
    dwFlags As Long
End Type

'
' dwFlags
'
Const HF32_DEFAULT = 1        ' process's default heap
Const HF32_SHARED = 2         ' is shared heap

Private Declare Function Heap32ListFirst Lib "TOOLHELP32" ( _
    ByVal hSnapshot As Long, lphl As HEAPLIST32) As Long
Private Declare Function Heap32ListNext Lib "TOOLHELP32" ( _
    ByVal hSnapshot As Long, lphl As HEAPLIST32) As Long

Private Type HEAPENTRY32
    dwSize As Long
    hHandle As Long         ' Handle of this heap block
    dwAddress As Long       ' Linear address of start of block
    dwBlockSize As Long     ' Size of block in bytes
    dwFlags As Long
    dwLockCount As Long
    dwResvd As Long
    th32ProcessID As Long   ' Owning process
    th32HeapID As Long      ' Heap block is in
End Type

'
' dwFlags
'
Const LF32_FIXED = &H1&
Const LF32_FREE = &H2&
Const LF32_MOVEABLE = &H4&

Private Declare Function Heap32First Lib "KERNEL32" (lphe As HEAPENTRY32, _
    ByVal th32ProcessID As Long, ByVal th32HeapID As Long) As Long
Private Declare Function Heap32Next Lib "KERNEL32" (lphe As HEAPENTRY32) As Long
Private Declare Function Toolhelp32ReadProcessMemory Lib "KERNEL32" ( _
    ByVal th32ProcessID As Long, ByVal lpBaseAddress As Long, _
    ByVal lpBuffer As Long, ByVal cbRead As Long, _
    lpNumberOfBytesRead As Long) As Long
#End If

' ***** Process walking ****

Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long           ' This process
    th32DefaultHeapID As Long
    th32ModuleID As Long            ' Associated exe
    cntThreads As Long
    th32ParentProcessID As Long     ' This process's parent process
    pcPriClassBase As Long          ' Base priority of process's threads
    dwFlags As Long
    szExeFile As String * 260       ' MAX_PATH
End Type

Private Declare Function Process32First Lib "KERNEL32" ( _
    ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long
Private Declare Function Process32Next Lib "KERNEL32" ( _
    ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long

' ***** Thread walking *****
#If 0 Then          ' Don't bother!
Private Type THREADENTRY32
    dwSize As Long
    cntUsage As Long
    th32ThreadID As Long        ' this thread
    th32OwnerProcessID As Long  ' Process this thread is associated with
    tpBasePri As Long
    tpDeltaPri As Long
    dwFlags As Long
End Type

Private Declare Function Thread32First Lib "KERNEL32" ( _
    ByVal hSnapshot As Long, lpte As THREADENTRY32) As Long
Private Declare Function Thread32Next Lib "KERNEL32" ( _
    ByVal hSnapshot As Long, lpte As THREADENTRY32) As Long
#End If

' ***** Module walking *****

Private Type MODULEENTRY32
    dwSize As Long
    th32ModuleID As Long        ' This module
    th32ProcessID As Long       ' owning process
    GlblcntUsage As Long        ' Global usage count on the module
    ProccntUsage As Long        ' Module usage count in th32ProcessID's context
    modBaseAddr As Long         ' Base address of module in th32ProcessID's context
    modBaseSize As Long         ' Size in bytes of module starting at modBaseAddr
    hModule As Long             ' The hModule of this module in th32ProcessID's context
    szModule As String * 256    ' MAX_MODULE_NAME32 + 1
    szExePath As String * 260   ' MAX_PATH
End Type

'
' NOTE CAREFULLY that the modBaseAddr and hModule fields are valid ONLY
' in th32ProcessID's process context.
'

Private Declare Function Module32First Lib "KERNEL32" ( _
    ByVal hSnapshot As Long, lpme As MODULEENTRY32) As Long
Private Declare Function Module32Next Lib "KERNEL32" ( _
    ByVal hSnapshot As Long, lpme As MODULEENTRY32) As Long
    

' Windows NT Private Declares, constants, and Private Types
' Use the PSAPI functions found in PSAPI.DLL
    
Private Declare Function EnumProcesses Lib "PSAPI" ( _
    lpidProcess As Long, ByVal cb As Long, cbNeeded As Long) As Long

Private Declare Function EnumProcessModules Lib "PSAPI" ( _
    ByVal hProcess As Long, lphModule As Long, _
    ByVal cb As Long, lpcbNeeded As Long) As Long

Private Declare Function GetModuleBaseName Lib "PSAPI" Alias "GetModuleBaseNameA" ( _
    ByVal hProcess As Long, ByVal hModule As Long, _
    ByVal lpBaseName As String, ByVal nSize As Long) As Long

Private Declare Function GetModuleBaseNameW Lib "PSAPI" ( _
    ByVal hProcess As Long, ByVal hModule As Long, _
    lpBaseName As Byte, ByVal nSize As Long) As Long

Private Declare Function GetModuleFileNameEx Lib "PSAPI" Alias "GetModuleFileNameExA" ( _
    ByVal hProcess As Long, ByVal hModule As Long, _
    ByVal lpFileName As String, ByVal nSize As Long) As Long

Private Declare Function GetModuleFileNameExW Lib "PSAPI" ( _
    ByVal hProcess As Long, ByVal hModule As Long, _
    lpFileName As Byte, ByVal nSize As Long) As Long

Private Type MODULEINFO
    lpBaseOfDll As Long
    SizeOfImage As Long
    EntryPoint As Long
End Type

Private Declare Function GetModuleInformation Lib "PSAPI" ( _
    ByVal hProcess As Long, ByVal hModule As Long, _
    lpmodinfo As MODULEINFO, ByVal cb As Long) As Long

' Additional PSAPI functions that I don't use from Visual Basic
#If 0 Then
Private Declare Function EmptyWorkingSet Lib "PSAPI" ( _
    ByVal hProcess As Long) As Long

Private Declare Function QueryWorkingSet Lib "PSAPI" ( _
    ByVal hProcess As Long, ByVal pv As Long, _
    ByVal cb As Long) As Long

Private Declare Function InitializeProcessForWsWatch Lib "PSAPI" ( _
    ByVal hProcess As Long) As Long

Private Type PSAPI_WS_WATCH_INFORMATION
    FaultingPc As Long
    FaultingVa As Long
End Type

Private Declare Function GetWsChanges Lib "PSAPI" ( _
    ByVal hProcess As Long, _
    lpWatchInfo As PSAPI_WS_WATCH_INFORMATION) As Long

Private Declare Function GetMappedFileNameA Lib "PSAPI" ( _
    ByVal hProcess As Long, lpv As Long, _
    lpFileName As Byte, ByVal nSize As Long) As Long

Private Declare Function GetMappedFileNameW Lib "PSAPI" ( _
    ByVal hProcess As Long, lpv As Long, _
    ByVal lpFileName As String, ByVal nSize As Long) As Long

Private Declare Function EnumDeviceDrivers Lib "PSAPI" ( _
    ByVal lpImageBase As Long, ByVal cb As Long, _
   lpcbNeeded As Long) As Long

Private Declare Function GetDeviceDriverBaseNameA Lib "PSAPI" ( _
    ByVal lpImageBase As Long, ByVal lpBaseName As String, _
    ByVal nSize As Long) As Long

Private Declare Function GetDeviceDriverBaseNameW Lib "PSAPI" ( _
    ByVal lpImageBase As Long, lpBaseName As Byte, _
    ByVal nSize As Long) As Long

Private Declare Function GetDeviceDriverFileNameA Lib "PSAPI" ( _
    ByVal lpImageBase As Long, ByVal lpFileName As String, _
    ByVal nSize As Long) As Long

Private Declare Function GetDeviceDriverFileNameW Lib "PSAPI" ( _
    ByVal lpImageBase As Long, lpFileName As Byte, _
    ByVal nSize As Long) As Long

' Structure for GetProcessMemoryInfo()

Private Type PROCESS_MEMORY_COUNTERS
    cb As Long
    PageFaultCount As Long
    PeakWorkingSetSize As Long
    WorkingSetSize As Long
    QuotaPeakPagedPoolUsage As Long
    QuotaPagedPoolUsage As Long
    QuotaPeakNonPagedPoolUsage As Long
    QuotaNonPagedPoolUsage As Long
    PagefileUsage As Long
    PeakPagefileUsage As Long
End Type

Private Declare Function GetProcessMemoryInfo Lib "PSAPI" ( _
    ByVal hProcess As Long, ppsmemCounters As PROCESS_MEMORY_COUNTERS, _
    ByVal cb As Long) As Long
#End If

Function CreateProcessList() As CVector
    Dim c As Long, f As Long, sName As String
    Dim vec As CVector, process As CProcess
    Set vec = New CVector
    
    If MUtility.IsNT = False Then
        ' Windows 95 uses ToolHelp32 functions
        Dim hSnap As Long, proc As PROCESSENTRY32
        ' Take a picture of current process list
        hSnap = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
        If hSnap = hNull Then Exit Function
        proc.dwSize = Len(proc)
        ' Iterate through the processes
        f = Process32First(hSnap, proc)
        Do While f
            ' Put this process in vector and count it
            sName = MUtility.StrZToStr(proc.szExeFile)
            Set process = New CProcess
            process.Create proc.th32ProcessID, MUtility.GetFileBaseExt(sName)
            c = c + 1
            Set vec(c) = process
            f = Process32Next(hSnap, proc)
        Loop
        CloseHandle hSnap
    Else
        ' Windows NT uses PSAPI functions
        Dim i As Long, iCur As Long, cRequest As Long, cGot As Long
        Dim aProcesses() As Long, hProcess As Long, hModule As Long
        ' Guess at maximum number and loop until guess is enough
        cRequest = 96       ' Request in bytes for 24 processes
        Do
            ReDim aProcesses(0 To (cRequest / 4) - 1) As Long
            ' Fill an array with process IDs
            f = EnumProcesses(aProcesses(0), cRequest, cGot)
            If f = 0 Then Exit Function
            If cGot < cRequest Then Exit Do
            cRequest = cRequest * 2
        Loop
        cGot = cGot / 4     ' From bytes to processes
        If cGot Then ReDim Preserve aProcesses(0 To cGot - 1) As Long
        
        ' Create CProcess object for each process
        For i = 0 To cGot - 1
            hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or _
                                   PROCESS_VM_READ, 0, _
                                   aProcesses(i))
            ' Processes that fail probably don't have security rights
            If hProcess Then
                ' Get first module only
                f = EnumProcessModules(hProcess, hModule, 4, c)
                If f Then
                    sName = String$(cMaxPath, 0)
                    c = GetModuleFileNameEx(hProcess, hModule, _
                                            sName, cMaxPath)
                    ' Put this process in vector and count it
                    Set process = New CProcess
                    process.Create aProcesses(i), Left$(sName, c)
                    iCur = iCur + 1
                    Set vec(iCur) = process
                End If
                CloseHandle hProcess
            End If
        Next
    End If
    Set CreateProcessList = vec
End Function

Function CreateModuleList(idProcessA As Long) As CVector
    Dim sName As String, f As Long, c As Long, i As Long, iCur As Long
    Dim vec As CVector, module As CModule
    Set vec = New CVector
    
    If MUtility.IsNT = False Then
        ' Windows 95 uses ToolHelp functions
        Dim modu As MODULEENTRY32, hSnap As Long
        hSnap = CreateToolhelp32Snapshot(TH32CS_SNAPMODULE, idProcessA)
        If hSnap = hNull Then Exit Function
        modu.dwSize = Len(modu)
        f = Module32First(hSnap, modu)
        Do While f
            Set module = New CModule
            sName = MUtility.GetFileBaseExt(MUtility.StrZToStr(modu.szExePath))
            ' Validate module handle
            If sName = ExeNameFromMod(modu.hModule) Then
                module.Create modu.th32ProcessID, modu.modBaseAddr, sName
                i = i + 1
                Set vec(i) = module
            End If
            f = Module32Next(hSnap, modu)
        Loop
        CloseHandle hSnap
    Else
        ' Windows NT uses PSAPI functions
        Dim cRequest As Long, cGot As Long
        Dim aModules() As Long, hProcess As Long, hModule As Long
        ' Get a handle
        hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or _
                               PROCESS_VM_READ, 0, idProcessA)
        If hProcess = 0 Then Exit Function
        ' Guess at maximum number and loop until guess is enough
        cRequest = 48       ' Request in bytes (16 modules)
        Do
            ReDim aModules(0 To (cRequest / 4) - 1) As Long
            ' Fill an array with module handles
            f = EnumProcessModules(hProcess, aModules(0), cRequest, cGot)
            If f = 0 Then Exit Function
            If cGot < cRequest Then Exit Do
            cRequest = cRequest * 2
        Loop
        cGot = cGot / 4     ' From bytes to modules
        If cGot Then ReDim Preserve aModules(0 To cGot - 1) As Long

        ' Create CModule object for each module of process
        For i = 0 To cGot - 1
            sName = String$(cMaxPath, 0)
            c = GetModuleFileNameEx(hProcess, aModules(i), sName, cMaxPath)
            If c Then
                sName = Left$(sName, c)
                Set module = New CModule
                module.Create idProcessA, aModules(i), sName
                iCur = iCur + 1
                Set vec(iCur) = module
            End If
        Next
        CloseHandle hProcess
    End If
    Set CreateModuleList = vec
End Function

Function ExeNameFromMod(ByVal hMod As Long) As String
    Dim st As String, cT As Long
    cT = 256: st = String$(256, 0)
    cT = GetModuleFileName(hMod, st, cT)
    st = Left$(st, cT)
    ExeNameFromMod = MUtility.GetFileBaseExt(st)
End Function

Function ModFromWnd(ByVal hWnd As Long) As Long
    BugAssert hWnd <> hNull
    ModFromWnd = ModFromProcID(MWinTool.ProcIDFromWnd(hWnd))
End Function

Function ModFromProcID(ByVal idProc As Long) As Long
    If Not MUtility.IsNT Then
        Dim process As PROCESSENTRY32, module As MODULEENTRY32
        Dim hSnap As Long, f As Long, idModule As Long
        hSnap = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
        If hSnap = hNull Then Exit Function
        ' Loop through to find matching process
        process.dwSize = Len(process)
        f = Process32First(hSnap, process)
        Do While f
            If process.th32ProcessID = idProc Then
                ' Save module ID
                idModule = process.th32ModuleID
                Exit Do
            End If
            f = Process32Next(hSnap, process)
        Loop
        CloseHandle hSnap

        hSnap = CreateToolhelp32Snapshot(TH32CS_SNAPMODULE, idProc)
        If hSnap = hNull Then Exit Function
        ' Loop through to find matching module
        module.dwSize = Len(module)
        f = Module32First(hSnap, module)
        Do While f
            If module.th32ModuleID = idModule Then
                ModFromProcID = module.hModule
                Exit Function
            End If
            f = Module32Next(hSnap, module)
        Loop
        CloseHandle hSnap
    Else
        Dim hModule As Long, c As Long
        ' First module is the main executable
        f = EnumProcessModules(ProcFromProcID(idProc), hModule, 4, c)
        ' Ignore errors (probably you have no security access)
        If f Then ModFromProcID = hModule
    End If
End Function

Function InstFromProcID(ByVal idProc As Long) As Long
    Dim f As Long, hModule As Long, c As Long
    If Not MUtility.IsNT Then
        Dim process As PROCESSENTRY32, module As MODULEENTRY32
        Dim hSnap As Long, idModule As Long
        hSnap = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
        If hSnap = hNull Then Exit Function
        ' Loop through to find matching process
        process.dwSize = Len(process)
        f = Process32First(hSnap, process)
        Do While f
            If process.th32ProcessID = idProc Then
                ' Save module ID
                idModule = process.th32ModuleID
                Exit Do
            End If
            f = Process32Next(hSnap, process)
        Loop
        CloseHandle hSnap
        
        ' Loop through modules
        hSnap = CreateToolhelp32Snapshot(TH32CS_SNAPMODULE, idProc)
        If hSnap = hNull Then Exit Function
        module.dwSize = Len(module)
        f = Module32First(hSnap, module)
        Do While f
            If module.th32ModuleID = idModule Then
                InstFromProcID = module.modBaseAddr
                Exit Function
            End If
            f = Module32Next(hSnap, module)
        Loop
        CloseHandle hSnap
    Else
        ' First module is the main executable
        f = EnumProcessModules(ProcFromProcID(idProc), hModule, 4, c)
        If f = 0 Then Exit Function
        Dim modinfo As MODULEINFO
        f = GetModuleInformation(ProcFromProcID(idProc), hModule, modinfo, c)
        If f Then InstFromProcID = modinfo.lpBaseOfDll
    End If
End Function

Function ProcIDFromInst(ByVal hInst As Long) As Long
    Dim f As Long, c As Long, idProc As Long
    If Not MUtility.IsNT Then
        Dim process As PROCESSENTRY32, hSnap As Long
        hSnap = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
        If hSnap = hNull Then Exit Function
        process.dwSize = Len(process)
        f = Process32First(hSnap, process)
        Do While f
            If InstFromProcID(process.th32ProcessID) = hInst Then
                ProcIDFromInst = process.th32ProcessID
                Exit Function
            End If
            f = Process32Next(hSnap, process)
        Loop
        CloseHandle hSnap
    Else
        Dim i As Long, iCur As Long, cRequest As Long, cGot As Long, modinfo As MODULEINFO
        Dim aProcesses() As Long, hProcess As Long, hModule As Long
        ' Guess at maximum number and loop until guess is enough
        cRequest = 96       ' Request in bytes (32 processes)
        Do
            ReDim aProcesses(0 To (cRequest / 4) - 1) As Long
            ' Fill an array with process IDs
            f = EnumProcesses(aProcesses(0), cRequest, cGot)
            If f = 0 Then Exit Function
            If cGot < cRequest Then Exit Do
            cRequest = cRequest * 2
        Loop
        cGot = cGot / 4     ' From bytes to processes
        If cGot Then ReDim Preserve aProcesses(0 To cGot - 1) As Long
        
        For i = 0 To cGot - 1
            hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or _
                                   PROCESS_VM_READ, 0, aProcesses(i))
            ' Processes that fail probably don't have security rights
            If hProcess Then
                ' Get first module only
                f = EnumProcessModules(hProcess, hModule, 4, c)
                If f Then
                    f = GetModuleInformation(hProcess, _
                                             hModule, modinfo, c)
                    ' If instance handle matches, return ID
                    If f And (hInst = modinfo.lpBaseOfDll) Then
                        ProcIDFromInst = aProcesses(i)
                        CloseHandle hProcess
                        Exit Function
                    End If
                    iCur = iCur + 1
                End If
                CloseHandle hProcess
            End If
        Next
    End If
End Function

' Must close result with CloseHandle when done
Function ProcFromInst(ByVal hInst As Long) As Long
    ProcFromInst = ProcFromProcID(ProcIDFromInst(hInst))
End Function

Function ModFromInst(ByVal hInst As Long) As Long
    ModFromInst = ModFromProcID(ProcIDFromInst(hInst))
End Function

Function ProcFromProcID(idProc As Long)
    ProcFromProcID = OpenProcess(PROCESS_QUERY_INFORMATION Or _
                                 PROCESS_VM_READ, 0, idProc)
End Function

Function ExeNameFromWnd(ByVal hWnd As Long) As String
    BugAssert hWnd <> hNull
    ExeNameFromWnd = MUtility.GetFileBaseExt(ExePathFromWnd(hWnd))
End Function

Function ExePathFromWnd(ByVal hWnd As Long) As String
    ExePathFromWnd = ExePathFromProcID(MWinTool.ProcIDFromWnd(hWnd))
End Function

Function ExePathFromProcID(idProc As Long) As String
    If Not MUtility.IsNT Then
        Dim process As PROCESSENTRY32, hSnap As Long, f As Long
        hSnap = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
        If hSnap = hNull Then Exit Function
        process.dwSize = Len(process)
        f = Process32First(hSnap, process)
        Do While f
            If process.th32ProcessID = idProc Then
                ExePathFromProcID = MUtility.StrZToStr(process.szExeFile)
                Exit Function
            End If
            f = Process32Next(hSnap, process)
        Loop
        CloseHandle hSnap
    Else
        Dim s As String, c As Long
        s = String$(cMaxPath, 0)
        c = GetModuleFileNameEx(ProcFromProcID(idProc), _
                                ModFromProcID(idProc), s, cMaxPath)
        If c Then ExePathFromProcID = Left$(s, c)
    End If
End Function

Function ExeNameFromProcID(idProc As Long) As String
    ExeNameFromProcID = MUtility.GetFileBaseExt(ExePathFromProcID(idProc))
End Function

Function ModFromExePath(sExe As String) As Long
    ModFromExePath = GetModuleHandle(sExe)
End Function

Function GetFirstInstWnd(hWndMe As Long) As Long
    Dim hWndYou As Long, idMe As Long, sExeMe As String
    
    ' Get my own process ID and executable name
    idMe = MWinTool.ProcIDFromWnd(hWndMe)
    sExeMe = ExeNameFromWnd(hWndMe)
    ' Get first sibling to start iterating top-level windows
    hWndYou = GetWindow(hWndMe, GW_HWNDFIRST)
    Do While hWndYou <> hNull
        ' Ignore if process ID of target is same
        If idMe <> MWinTool.ProcIDFromWnd(hWndYou) Then
            ' Ignore if module name is different
            If sExeMe = ExeNameFromWnd(hWndYou) Then
                ' Return first with same module, different process
                GetFirstInstWnd = hWndYou
                Exit Function
            End If
        End If
        ' Get next sibling
        hWndYou = GetWindow(hWndYou, GW_HWNDNEXT)
    Loop
End Function

Function GetAllInstWnd(hWndMe As Long) As CVector
    Dim hWndYou As Long, idMe As Long, sExeMe As String
    Dim vecWnds As CVector, i As Long
    Set vecWnds = New CVector
    
    ' Get my own process ID and executable name
    idMe = MWinTool.ProcIDFromWnd(hWndMe)
    sExeMe = ExeNameFromWnd(hWndMe)
    ' Get first sibling to start iterating top level windows
    hWndYou = GetWindow(hWndMe, GW_HWNDFIRST)
    Do While hWndYou <> hNull
        ' Ignore if process ID of target is same
        If idMe <> MWinTool.ProcIDFromWnd(hWndYou) Then
            ' Ignore if module name is different
            If sExeMe = ExeNameFromWnd(hWndYou) Then
                ' Return all with same module, different process
                i = i + 1
                vecWnds(i) = hWndYou
            End If
        End If
        ' Get next sibling
        hWndYou = GetWindow(hWndYou, GW_HWNDNEXT)
    Loop
    Set GetAllInstWnd = vecWnds
End Function
'

#If fComponent = 0 Then
Private Sub ErrRaise(e As Long)
    Dim sText As String, sSource As String
    If e > 1000 Then
        sSource = App.ExeName & ".ModTool"
        Select Case e
        Case eeBaseModTool
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

