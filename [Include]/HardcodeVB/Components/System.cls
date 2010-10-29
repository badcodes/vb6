VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum EErrorSystem
    eeBaseSystem = 13250    ' CSystem
End Enum

Private Type SYSTEM_INFO
    wProcessorArchitecture  As Integer
    wReserved As Integer
    dwPageSize As Long
    lpMinimumApplicationAddress As Long
    lpMaximumApplicationAddress As Long
    dwActiveProcessorMask As Long
    dwNumberOfProcessors As Long
    dwProcessorType As Long
    dwAllocationGranularity As Long
    wProcessorLevel As Integer
    wProcessorRevision As Integer
End Type

Private Declare Sub GetSystemInfo Lib "KERNEL32" (lpSystemInfo As SYSTEM_INFO)

Private iWinMajor As Integer
Private iWinMinor As Integer
Private sMode As String
Private sys As SYSTEM_INFO

Private Sub Class_Initialize()
    Dim dw As Long, c As Integer
    dw = GetVersion()
    iWinMajor = dw And &HFF&
    iWinMinor = (dw And &HFF00&) / &H100&
    sMode = IIf(dw And &H80000000, "Windows 95", "Windows NT")
    GetSystemInfo sys
End Sub

Property Get FreePhysicalMemory() As Long
    Dim mem As MEMORYSTATUS
    mem.dwLength = Len(mem)
    GlobalMemoryStatus mem
    FreePhysicalMemory = mem.dwAvailPhys \ 1024
End Property

Property Get TotalPhysicalMemory() As Long
    Dim mem As MEMORYSTATUS
    mem.dwLength = Len(mem)
    GlobalMemoryStatus mem
    TotalPhysicalMemory = mem.dwTotalPhys \ 1024
End Property

Property Get FreeVirtualMemory() As Long
    Dim mem As MEMORYSTATUS
    mem.dwLength = Len(mem)
    GlobalMemoryStatus mem
    FreeVirtualMemory = mem.dwAvailVirtual \ 1024
End Property

Property Get TotalVirtualMemory() As Long
    Dim mem As MEMORYSTATUS
    mem.dwLength = Len(mem)
    GlobalMemoryStatus mem
    TotalVirtualMemory = mem.dwTotalVirtual \ 1024
End Property

Property Get FreePageFile() As Long
    Dim mem As MEMORYSTATUS
    mem.dwLength = Len(mem)
    GlobalMemoryStatus mem
    FreePageFile = mem.dwAvailPageFile \ 1024
End Property

Property Get TotalPageFile() As Long
    Dim mem As MEMORYSTATUS
    mem.dwLength = Len(mem)
    GlobalMemoryStatus mem
    TotalPageFile = mem.dwTotalPageFile \ 1024
End Property

Property Get MemoryLoad() As Long
    Dim mem As MEMORYSTATUS
    mem.dwLength = Len(mem)
    GlobalMemoryStatus mem
    MemoryLoad = mem.dwMemoryLoad
End Property

Property Get WinMajor() As Integer
    WinMajor = iWinMajor
End Property

Property Get WinMinor() As Integer
    WinMinor = iWinMinor
End Property

Property Get WinVersion() As Single
    WinVersion = iWinMajor + (iWinMinor / 100)
End Property

Property Get Processor() As String
    If sMode = "Windows 95" Then
        Processor = "Intel "
        Select Case sys.dwProcessorType
        Case 386
            Processor = Processor & "386"
        Case 486
            Processor = Processor & "486"
        Case 586
            Processor = Processor & "586"
        End Select
    Else
        Select Case sys.wProcessorArchitecture
        Case PROCESSOR_ARCHITECTURE_INTEL
            Processor = "Intel "
            Select Case sys.wProcessorLevel
            Case 3, 4
                Processor = Processor & sys.wProcessorLevel & "86"
            Case 5
                Processor = Processor & "Pentium"
            Case Else
                Processor = Processor & "Level " & sys.wProcessorLevel
            End Select
        Case PROCESSOR_ARCHITECTURE_MIPS
            Processor = "MIPS R" & sys.wProcessorLevel & "000"
        Case PROCESSOR_ARCHITECTURE_ALPHA
            Processor = "Alpha " & sys.wProcessorLevel
        Case PROCESSOR_ARCHITECTURE_PPC
            Processor = "Power PC " & IIf(sys.wProcessorLevel > 9, "6", "60") & _
                        sys.wProcessorLevel
        Case PROCESSOR_ARCHITECTURE_UNKNOWN
            Processor = "Unknown"
        Case Else
            Processor = "Other " & sys.wProcessorArchitecture & " " & sys.wProcessorLevel
        End Select
    End If
End Property

Property Get ProcessorCount() As String
    ProcessorCount = sys.dwNumberOfProcessors
End Property

Property Get Mode() As String
    Mode = sMode
End Property

Property Get WindowsDir() As String
    Dim s As String, c As Long
    s = String$(cMaxPath, 0)
    c = GetWindowsDirectory(s, cMaxPath)
    WindowsDir = Left(s, c)
End Property

Property Get SystemDir() As String
    Dim s As String, c As Long
    s = String$(cMaxPath, 0)
    c = GetSystemDirectory(s, cMaxPath)
    SystemDir = Left(s, c)
End Property

Property Get User() As String
    Dim s As String, c As Long
    c = 80: s = String$(c + 1, 0)
    ' Includes null in returned length, unlike all other API functions
    If GetUserName(s, c) Then User = Left$(s, c - 1)
End Property

Property Get Machine() As String
    Dim s As String, c As Long
    c = 16: s = String$(16, 0)
    If GetComputerName(s, c) Then Machine = Left$(s, c)
End Property

#If fComponent = 0 Then
Private Sub ErrRaise(e As Long)
    Dim sText As String, sSource As String
    If e > 1000 Then
        sSource = App.ExeName & ".System"
        Select Case e
        Case eeBaseSystem
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

