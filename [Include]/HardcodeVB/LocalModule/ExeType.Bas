Attribute VB_Name = "MExeType"
Option Explicit

Public Enum EErrorExeType
    eeBaseExeType = 13470   ' ExeType
End Enum

' Valid Exe types (for ExeType function)

Public Enum EProgramType
    ' Unknown - could still be .BAT, .CMD, .COM, or .PIF
    eptNotExe = 0
    ' Recognized executable types
    eptMSDOS = 1
    eptWin16 = 2
    eptOS2_1 = 3
    eptWin32 = 4
    eptWin32Console = 5
    eptDOSUnknown = 7
    ' Errors
    eptNoFile = -1
    eptOS2_2 = -2
    eptWinOS2DLL = -3
    eptNEUnknown = -4
    eptNTNonIntel = -5
    eptWin32DLL = -6
    eptAccessFail = -7
End Enum

' Check to see if specified file is executable, and if so, what kind
Function ExeType(sSpec As String) As EProgramType
    On Error GoTo ExeTypeFail
    Dim nFile As Integer
    nFile = FreeFile
    If MUtility.ExistFile(sSpec) Then
        Open sSpec For Binary Access Read Shared As nFile
    Else
        ExeType = eptNoFile
        Exit Function
    End If

    Dim abHeader() As Byte
    ReDim abHeader(128)
    Get nFile, 1, abHeader

    ' MS-DOS headers start with magic header "MZ"
    Dim sMagic As String, bData As Byte, wData As Integer
    sMagic = MBytes.LeftBytes(abHeader, 2)
    If sMagic <> "MZ" Then
        ' Could still be a .BAT, .CMD, .PIF, or .COM file
        ExeType = eptNotExe
        Close nFile
        Exit Function
    End If

    ' If word at offset &H18 does not point beyond DOS header
    ' (length &H40), file is MS-DOS EXE
    If MBytes.BytesToWord(abHeader, &H18) < &H40 Then
        ExeType = eptMSDOS
        Close nFile
        Exit Function
    End If

    ' Get offset of new EXE header
    wData = MBytes.BytesToWord(abHeader, &H3C)
    Get nFile, wData + 1, abHeader
    Close nFile

    ' New .EXE headers start with magic header "NE"
    sMagic = MBytes.LeftBytes(abHeader, 2)
    ' Check for Windows/OS2 format
    If sMagic = "NE" Then

        ' Get the executable file flags to check for DLL
        If abHeader(&HD) And &H80 Then
            ' This is a DLL (executable but not by us)
            ExeType = eptWinOS2DLL
        Else
            ' Get the operating system flags (byte, not word)
            bData = abHeader(&H36)
            If bData And &H2 Then
                ExeType = eptWin16 ' Windows
            ElseIf bData And &H1 Then
                ExeType = eptOS2_1 ' OS/2 1.x
            Else
                ' Unknown NE system, probably bound, but call it MS-DOS
                ExeType = eptMSDOS
            End If
        End If

    ' Check for OS/2 2.x format (can't execute from Windows or NT)
    ElseIf sMagic = "LE" Then
        ExeType = eptOS2_2 ' OS/2 LE
    ' Check for NT format
    ElseIf sMagic = "PE" And MBytes.BytesToWord(abHeader, &H2) = 0 Then
        ' Get processor flags
        bData = abHeader(&H4)
        Select Case bData
        Case &H4C, &H4D, &H4E, &H4F ' NT for intel 386, 486, 586, 686
            ExeType = eptWin32 ' NT Windows
        Case Else
            ExeType = eptNTNonIntel ' Some sort of RISC or other
            Exit Function
        End Select

        ' Get the Exe type flags
        If abHeader(&H17) And &H20 Then
            ExeType = eptWin32DLL ' Executable, but not by us
            Exit Function
        End If

        ' Get the subsystem flags to identify NT character
        If abHeader(&H5C) = 3 Then ExeType = eptWin32Console
        ' Could also identify Posix here

    Else
        ' MS-DOS file with a header, but notNE file
        ' (Some 16-bit DOS-extended executables fall through here, or
        ' could be non-EXE file with "MZ" as first two bytes)
        ExeType = eptDOSUnknown ' Probably DOS extended
    End If
    Exit Function
    
ExeTypeFail:
    ExeType = eptAccessFail
End Function

Function ExeTypeStr(sFile As String) As String
    Select Case ExeType(sFile)
    ' Valid Exe types (for ExeType function)
    Case eptMSDOS
        ExeTypeStr = "MS-DOS"
    Case eptWin16
        ExeTypeStr = "Windows 16-bit"
    Case eptOS2_1
        ExeTypeStr = "OS/2 1.x"
    Case eptWin32
        ExeTypeStr = "Windows 32-bit"
    Case eptWin32Console
        ExeTypeStr = "Windows 32-bit Console"
    Case eptDOSUnknown
        ExeTypeStr = "Unknown MS-DOS Compatible"
    Case eptNotExe
        ExeTypeStr = "Not EXE File"
    Case eptNoFile
        ExeTypeStr = "No File"
    Case eptOS2_2
        ExeTypeStr = "OS/2 2.x"
    Case eptWinOS2DLL
        ExeTypeStr = "Windows 3.x or OS/2 DLL"
    Case eptNEUnknown
        ExeTypeStr = "Unknown Format"
    Case eptNTNonIntel
        ExeTypeStr = "Non-Intel Windows"
    Case eptWin32DLL
        ExeTypeStr = "Windows 32-bit DLL"
    End Select
End Function

#If fComponent = 0 Then
Private Sub ErrRaise(e As Long)
    Dim sText As String, sSource As String
    If e > 1000 Then
        sSource = App.ExeName & ".ExeType"
        Select Case e
        Case eeBaseExeType
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

