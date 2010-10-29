VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CVersion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum EErrorVersion
    eeBaseVersion = 13410   ' CVersion
End Enum

' Internal class data
Private sExeName As String
'Private sVer As String     ' Used in original version
' James Musial correctly noted that Version APIs should read data
' into an array of bytes rather than a string, so I changed it
Private abVer() As Byte
Private pData As Long, cData As Long
Private fixed As VS_FIXEDFILEINFO
Private sPrefix As String

' Create new Version object
Sub Create(sExeNameA As String)

    ' If it already exists, destroy it
    If Not MUtility.IsArrayEmpty(abVer) Then ReInit

    Dim hVer As Long, cVer As Long
    ' Get size and handle of version data
    cVer = GetFileVersionInfoSize(sExeNameA, hVer)
    If cVer = 0 Then
        ' No error means file has no resources
        If Err.LastDllError = 0 Then Exit Sub
        ErrRaise Err.LastDllError
    End If
    ReDim abVer(0 To cVer - 1) As Byte
    Dim f As Long
    f = GetFileVersionInfo(sExeNameA, hVer, cVer, abVer(0))
    If f = 0 Then ErrRaise Err.LastDllError

    ' Get fixed portion of data
    f = VerQueryValue(abVer(0), "\", pData, cData)
    If f = 0 Then ErrRaise Err.LastDllError
    BugAssert cData = Len(fixed)
    CopyMemory fixed, ByVal pData, Len(fixed)

    ' Get version as hex number with low and high reversed
    f = VerQueryValue(abVer(0), "\VarFileInfo\Translation", pData, cData)
    If f = 0 Then ErrRaise Err.LastDllError
    If cData = 0 Then
        ' No version information
        ReInit
        Exit Sub
    End If
    ' This is where you would enhance to support multiple languages--cData
    ' could be greater than 4, but we use only the first entry (thanks
    ' to Martin Naughton for identifying this problem)
    cData = 4

    ' Success - copy internal data to prefix and exe name
    CopyMemory cVer, ByVal pData, cData
    ' Version APIs are extremely picky about format of this string
    sPrefix = "\StringFileInfo\" & _
              MUtility.FmtHex(MBytes.LoWord(cVer), 4) & _
              MUtility.FmtHex(MBytes.HiWord(cVer), 4) & "\"
    ' Language string from cVer has this format:
    '   04------ = SUBLANG such as SUBLANG_ENGLISH_USA
    '   --09---- = LANG such as LANG_ENGLISH
    '   ----04E4 = Codepage such 1252 for Windows:Multilingual
              
    sExeName = sExeNameA
    
End Sub

' Change file associated with an existing version object

Property Let ExeName(sExeNameA As String)
Attribute ExeName.VB_UserMemId = 0
    ReInit
    Create sExeNameA
End Property

Property Get ExeName() As String
    ExeName = sExeName
End Property

Private Sub ReInit()
With fixed
    sExeName = sEmpty
    Erase abVer
    .dwSignature = 0&
    .dwStrucVersion = 0&
    .dwFileVersionMS = 0&
    .dwFileVersionLS = 0&
    .dwProductVersionMS = 0&
    .dwProductVersionLS = 0&
    .dwFileFlagsMask = 0&
    .dwFileFlags = 0&
    .dwFileOS = 0&
    .dwFileType = 0&
    .dwFileSubtype = 0&
    .dwFileDateMS = 0&
    .dwFileDateLS = 0&
End With
End Sub

Private Function BufToStr(pBuf As Long, cBuf) As String
    Dim s As String
    s = String$(cBuf + 1, 0)
    If UnicodeTypeLib Then
        CopyMemoryToStr s, ByVal pBuf, cBuf * 2
    Else
        CopyMemoryToStr s, ByVal pBuf, cBuf
    End If
    BufToStr = MUtility.StrZToStr(s)
End Function

Property Get FullFileVersion() As String
    If Not MUtility.IsArrayEmpty(abVer) Then
        FullFileVersion = MBytes.HiWord(fixed.dwFileVersionMS) & "." & _
                          MBytes.LoWord(fixed.dwFileVersionMS) & "." & _
                          MBytes.HiWord(fixed.dwFileVersionLS) & "." & _
                          MBytes.LoWord(fixed.dwFileVersionLS)
    End If
End Property

Property Get FullProductVersion() As String
    If Not MUtility.IsArrayEmpty(abVer) Then
        FullProductVersion = MBytes.HiWord(fixed.dwProductVersionMS) & "." & _
                             MBytes.LoWord(fixed.dwProductVersionMS) & "." & _
                             MBytes.HiWord(fixed.dwProductVersionLS) & "." & _
                             MBytes.LoWord(fixed.dwProductVersionLS)
    End If
End Property

Property Get FileVersionMajor() As Long
    FileVersionMajor = fixed.dwFileVersionMS
End Property

Property Get FileVersionMinor() As Long
    FileVersionMinor = fixed.dwFileVersionLS
End Property

Property Get ProductVersionMajor() As Long
    ProductVersionMajor = fixed.dwProductVersionMS
End Property

Property Get ProductVersionMinor() As Long
    ProductVersionMinor = fixed.dwProductVersionLS
End Property

Property Get BuildOptions() As Long
    BuildOptions = fixed.dwFileFlags
End Property

Property Get BuildString() As String
With fixed
    If MUtility.IsArrayEmpty(abVer) Then Exit Property
    Dim s As String
    If .dwFileFlags And VS_FF_DEBUG Then s = s & "Debug "
    If .dwFileFlags And VS_FF_PRERELEASE Then s = s & "Prerelease "
    If .dwFileFlags And VS_FF_PATCHED Then s = s & "Patched "
    If .dwFileFlags And VS_FF_PRIVATEBUILD Then s = s & "PrivateBuild "
    If .dwFileFlags And VS_FF_INFOINFERRED Then s = s & "InfoInferred "
    If .dwFileFlags And VS_FF_SPECIALBUILD Then s = s & "SpecialBuild "
    BuildString = s
End With
End Property

Property Get Environment() As String
    If MUtility.IsArrayEmpty(abVer) Then Exit Property
    Dim s As String
    Select Case MBytes.LoWord(fixed.dwFileOS)
    Case VOS__WINDOWS16
        s = "Windows (16-bit)"
    Case VOS__PM16
        s = "PM (16-bit)"
    Case VOS__PM32
        s = "PM (32-bit)"
    Case VOS__WINDOWS32
        s = "Windows (32-bit)"
    End Select
    Select Case MBytes.HiWord(fixed.dwFileOS)
    Case MBytes.HiWord(VOS_DOS)
        s = s & " under MS-DOS"
    Case MBytes.HiWord(VOS_OS216)
        s = s & " under OS/2 (16-bit)"
    Case MBytes.HiWord(VOS_OS232)
        s = s & " under OS/2 (32-bit)"
    Case MBytes.HiWord(VOS_NT)
        s = s & " under Windows-NT"
    End Select
    Environment = s
End Property

Property Get ExeType() As String
    If MUtility.IsArrayEmpty(abVer) Then Exit Property
    Dim s As String
    Select Case fixed.dwFileType
    Case VFT_APP:
        s = "Application"
    Case VFT_DLL:
        s = "Dynamic Link Library"
    Case VFT_DRV:
        s = "Driver"
        Select Case fixed.dwFileSubtype
        Case VFT2_DRV_PRINTER
            s = s & ":Printer"
        Case VFT2_DRV_KEYBOARD
            s = s & ":Keyboard"
        Case VFT2_DRV_LANGUAGE
            s = s & ":Language"
        Case VFT2_DRV_DISPLAY
            s = s & ":Display"
        Case VFT2_DRV_MOUSE
            s = s & ":Mouse"
        Case VFT2_DRV_NETWORK
            s = s & ":Network"
        Case VFT2_DRV_SYSTEM
            s = s & ":System"
        Case VFT2_DRV_INSTALLABLE
            s = s & ":Installable"
        Case VFT2_DRV_SOUND
            s = s & ":Sound"
        Case VFT2_DRV_COMM
            s = s & ":Communications"
        End Select
    Case VFT_FONT:
        s = "Font"
        Select Case fixed.dwFileSubtype
        Case VFT2_FONT_RASTER
            s = s & ":Raster"
        Case VFT2_FONT_VECTOR
            s = s & ":Vector"
        Case VFT2_FONT_TRUETYPE
            s = s & ":TrueType"
        End Select
    Case VFT_VXD:
        s = "VXD"
    Case VFT_STATIC_LIB:
        s = "Static Library"
    Case Else
        s = "Unknown"
    End Select
    ExeType = s
End Property

Property Get TimeStamp() As Date
    Dim f As Boolean, ft As FILETIME, ftl As FILETIME, st As SYSTEMTIME
    If fixed.dwFileDateMS = 0 And fixed.dwFileDateLS = 0 Then Exit Property
    ft.dwHighDateTime = fixed.dwFileDateMS
    ft.dwLowDateTime = fixed.dwFileDateLS
    f = FileTimeToLocalFileTime(ft, ftl)
    If f Then f = FileTimeToSystemTime(ftl, st)
    If f = False Then Exit Property
    TimeStamp = DateSerial(st.wYear, st.wMonth, st.wDay) + _
                TimeSerial(st.wHour, st.wMinute, st.wSecond)
End Property

Property Get Company() As String
    If MUtility.IsArrayEmpty(abVer) Then Exit Property
    If VerQueryValue(abVer(0), sPrefix & "CompanyName", pData, cData) Then
        Company = BufToStr(pData, cData)
    End If
End Property

Property Get Comments() As String
    If MUtility.IsArrayEmpty(abVer) Then Exit Property
    If VerQueryValue(abVer(0), sPrefix & "Comments", pData, cData) Then
        Comments = BufToStr(pData, cData)
    End If
End Property

Property Get Description() As String
    If MUtility.IsArrayEmpty(abVer) Then Exit Property
    If VerQueryValue(abVer(0), sPrefix & "FileDescription", pData, cData) Then
        Description = BufToStr(pData, cData)
    End If
End Property

Property Get ProductVersionString() As String
    If MUtility.IsArrayEmpty(abVer) Then Exit Property
    If VerQueryValue(abVer(0), sPrefix & "ProductVersion", pData, cData) Then
        ProductVersionString = BufToStr(pData, cData)
    End If
End Property

Property Get FileVersionString() As String
    If MUtility.IsArrayEmpty(abVer) Then Exit Property
    If VerQueryValue(abVer(0), sPrefix & "FileVersion", pData, cData) Then
        FileVersionString = BufToStr(pData, cData)
    End If
End Property


Property Get InternalName() As String
    If MUtility.IsArrayEmpty(abVer) Then Exit Property
    If VerQueryValue(abVer(0), sPrefix & "InternalName", pData, cData) Then
        InternalName = BufToStr(pData, cData)
    End If
End Property

Property Get Copyright() As String
    If MUtility.IsArrayEmpty(abVer) Then Exit Property
    If VerQueryValue(abVer(0), sPrefix & "LegalCopyright", pData, cData) Then
        Copyright = BufToStr(pData, cData)
    End If
End Property

Property Get Trademarks() As String
    If MUtility.IsArrayEmpty(abVer) Then Exit Property
    If VerQueryValue(abVer(0), sPrefix & "LegalTradeMarks", pData, cData) Then
        Trademarks = BufToStr(pData, cData)
    End If
End Property

Property Get OriginalFilename() As String
    If MUtility.IsArrayEmpty(abVer) Then Exit Property
    If VerQueryValue(abVer(0), sPrefix & "OriginalFilename", pData, cData) Then
        OriginalFilename = BufToStr(pData, cData)
    End If
End Property

Property Get PrivateBuild() As String
    If MUtility.IsArrayEmpty(abVer) Then Exit Property
    If VerQueryValue(abVer(0), sPrefix & "PrivateBuild", pData, cData) Then
        PrivateBuild = BufToStr(pData, cData)
    End If
End Property

Property Get ProductName() As String
    If MUtility.IsArrayEmpty(abVer) Then Exit Property
    If VerQueryValue(abVer(0), sPrefix & "ProductName", pData, cData) Then
        ProductName = BufToStr(pData, cData)
    End If
End Property

Property Get ProductVersion() As String
    If MUtility.IsArrayEmpty(abVer) Then Exit Property
    If VerQueryValue(abVer(0), sPrefix & "ProductVersion", pData, cData) Then
        ProductVersion = BufToStr(pData, cData)
    End If
End Property

Property Get SpecialBuild() As String
    If MUtility.IsArrayEmpty(abVer) Then Exit Property
    If VerQueryValue(abVer(0), sPrefix & "SpecialBuild", pData, cData) Then
        SpecialBuild = BufToStr(pData, cData)
    End If
End Property

Property Get Custom(sCustom As String) As String
    If MUtility.IsArrayEmpty(abVer) Then Exit Property
    If VerQueryValue(abVer(0), sPrefix & sCustom, pData, cData) Then
        Custom = BufToStr(pData, cData)
    End If
End Property

#If fComponent = 0 Then
Private Sub ErrRaise(e As Long)
    Dim sText As String, sSource As String
    If e > 1000 Then
        sSource = App.ExeName & ".Version"
        Select Case e
        Case eeBaseVersion
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

