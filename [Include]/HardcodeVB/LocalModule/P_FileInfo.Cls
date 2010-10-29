VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFileInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum EErrorFileInfo
    eeBaseFileInfo = 13070      ' CFileInfo
End Enum

Public Enum EItemType
    eitFile = 1                 ' File or directory
    eitDrive                    ' Drive
    eitID                       ' PIDL passed to us
    eitFolder                   ' PIDL created by us from special folder
End Enum

Private Enum EItemState
    eisNotCreated
    eisFile                     ' File or directory
    eisDrive                    ' Drive
    eisID                       ' PIDL passed to us
    eisFolder                   ' PIDL created by us from special folder
End Enum
Private eis As EItemState       ' How object was created
Private vItem As Variant        ' File name or PIDL
Private shfi As SHFILEINFO      ' Info from SHGetFileInfo
Private fd As WIN32_FIND_DATA   ' Info from FindFirstFile
Private afAttr As Long          ' File attributes
Private afOption As Long        ' Options for SHGetFileInfo

Public Enum EExeType
    eetWin16Exe = &H454E
    eetDosExe = &H5A4D
    eetWin32Exe = &H4550
    eetWin32Console = &H4543
End Enum

Property Get item() As Variant
Attribute item.VB_UserMemId = 0
    item = vItem
End Property

Property Let item(vItemA As Variant)
    Dim h As Long, f As Long, af As Long
    Destroy     ' Clear any previous assignment
    If VarType(vItemA) = vbString Then
        ' String item is a file, directory, or drive
        If Len(vItemA) <= 3 And Mid$(vItemA, 2, 1) = ":" Then
            ' Must be drive, get attributes
            afAttr = 0: afOption = 0
        Else
            ' No terminating backslashes
            MUtility.DenormalizePath vItemA
            ' For file, get information in advance
            h = FindFirstFile(vItemA, fd)
            If h = hInvalid Then ApiRaise Err.LastDllError
            FindClose h
            afAttr = fd.dwFileAttributes
            afOption = SHGFI_USEFILEATTRIBUTES
        End If
        eis = eisFile
        af = afOption And (Not SHGFI_PIDL) Or _
             SHGFI_DISPLAYNAME Or SHGFI_TYPENAME
        f = SHGetFileInfo(vItemA, afAttr, shfi, LenB(shfi), af)
    Else
        ' Integer item is a special folder constant or pidl
        If vItemA < 32767 Then
            ' Turn special folder location into a pidl
            Dim pidl As Long
            SHGetSpecialFolderLocation 0, CLng(vItemA), pidl
            vItemA = pidl
            eis = eisFolder
        Else
            eis = eisID
            pidl = vItemA
        End If
        ' For special folders or other PIDLs, everything comes from system
        afAttr = 0: afOption = 0
        ' Get item ID pointer, but don't use attributes
        af = SHGFI_PIDL Or SHGFI_DISPLAYNAME Or _
             SHGFI_TYPENAME
        f = SHGetItemInfo(pidl, afAttr, shfi, Len(shfi), af)
    End If
    If f Then
        vItem = vItemA
    Else
        eis = eisNotCreated
    End If
End Property

' In a drive loop it's more efficient to create from drive data
Sub CreateFromDrive(sRootA As String, sKind As String, _
                    rTotal As Double, rFree As Double)
    Dim f As Long
    Destroy
    afAttr = 0
    fd.dwFileAttributes = 0
    fd.ftLastAccessTime = 0
    fd.ftLastWriteTime = 0
    fd.ftCreationTime = 0
    fd.nFileSizeLow = 0
    fd.dwReserved0 = CLng(rTotal / 1000000)
    fd.dwReserved1 = CLng(rFree / 1000000)
    f = SHGetFileInfo(sRootA, afAttr, shfi, LenB(shfi), _
                      SHGFI_DISPLAYNAME Or SHGFI_TYPENAME Or SHGFI_ATTRIBUTES)
    MBytes.StrToBytes fd.cAlternateFileName, sKind
    If f Then
        vItem = sRootA
        eis = eisDrive
    End If
End Sub

' In a FindFirstFile loop it's more efficient to create from file data
Sub CreateFromFile(sFileA As String, ByVal afAttrA As Long, _
                   ByVal cLenA As Long, ftModifiedA As Currency, _
                   ftAccessedA As Currency, ftCreatedA As Currency)
    Dim f As Long
    Destroy
    afAttr = afAttrA
    fd.dwFileAttributes = afAttrA
    fd.ftLastAccessTime = ftAccessedA
    fd.ftLastWriteTime = ftModifiedA
    fd.ftCreationTime = ftCreatedA
    fd.nFileSizeLow = cLenA
    f = SHGetFileInfo(sFileA, afAttr, shfi, LenB(shfi), _
                      SHGFI_DISPLAYNAME Or SHGFI_TYPENAME Or SHGFI_ATTRIBUTES)
    If f Then
        vItem = sFileA
        eis = eisFile
    End If
End Sub

Sub CreateFromNamePidl(sNameA As String, ByVal pidl As Long)
    Destroy
    afAttr = 0
    fd.dwFileAttributes = 0
    fd.ftLastAccessTime = 0
    fd.ftLastWriteTime = 0
    fd.ftCreationTime = 0
    fd.nFileSizeLow = 0
    vItem = pidl
    eis = eisID
    Dim c As Long
    If UnicodeTypeLib Then
        c = Len(sNameA) + 1
        CopyMemoryStr shfi.szDisplayName(0), sNameA, c * 2
    Else
        c = Len(sNameA) + 1
        CopyMemoryStr shfi.szDisplayName(0), sNameA, c
    End If
End Sub

Private Sub Destroy()
    ' Free any pidl we created from special folder
    If eis = eisFolder Then Allocator.Free vItem
    eis = eisNotCreated
    vItem = Empty
End Sub

Property Get DisplayName() As String
    If eis Then DisplayName = MBytes.ByteZToStr(shfi.szDisplayName)
    If DisplayName = sEmpty Then DisplayName = "Unknown"
End Property

' Returns file type
Property Get TypeName() As String
    If eis Then TypeName = MBytes.ByteZToStr(shfi.szTypeName)
    If TypeName = sEmpty Then TypeName = "None"
End Property

Function SmallIcon(Optional afOverlay As Long = 0) As Picture
    Dim shfiT As SHFILEINFO
    If eis = eisNotCreated Then Exit Function
    ' Filter out any invalid flags -- only overlays allowed
    afOverlay = afOverlay And (SHGFI_LINKOVERLAY Or SHGFI_SELECTED _
                               Or SHGFI_OPENICON)
    ' Add in standard and small icon flags
    afOverlay = afOverlay Or afOption Or SHGFI_ICON Or SHGFI_SMALLICON
    GetFileItemInfo vItem, shfiT, afOverlay, afAttr
    Set SmallIcon = MPicTool.IconToPicture(shfiT.hIcon)
End Function

Function LargeIcon(Optional afOverlay As Long = 0) As Picture
    Dim shfiT As SHFILEINFO
    If eis = eisNotCreated Then Exit Function
    ' Filter out any invalid flags -- only overlays allowed
    afOverlay = afOverlay And (SHGFI_LINKOVERLAY Or SHGFI_SELECTED _
                               Or SHGFI_OPENICON)
    ' Add in standard and large icon flags
    afOverlay = afOverlay Or afOption Or SHGFI_ICON Or SHGFI_LARGEICON
    GetFileItemInfo vItem, shfiT, afOverlay, afAttr
    Set LargeIcon = MPicTool.IconToPicture(shfiT.hIcon)
End Function

Function ShellIcon(Optional afOverlay As Long = 0) As Picture
    Dim shfiT As SHFILEINFO
    If eis = eisNotCreated Then Exit Function
    ' Filter out any invalid flags -- only overlays allowed
    afOverlay = afOverlay And (SHGFI_LINKOVERLAY Or SHGFI_SELECTED _
                               Or SHGFI_OPENICON)
    ' Add in standard and large icon flags
    afOverlay = afOverlay Or afOption Or SHGFI_ICON Or SHGFI_SHELLICONSIZE
    GetFileItemInfo vItem, shfiT, afOverlay, afAttr
    Set ShellIcon = MPicTool.IconToPicture(shfiT.hIcon)
End Function

Function Icon(afKind As Long) As Picture
    Dim shfiT As SHFILEINFO
    If eis = eisNotCreated Then Exit Function
    GetFileItemInfo vItem, shfiT, afOption Or SHGFI_ICON Or afKind, afAttr
    Set Icon = MPicTool.IconToPicture(shfiT.hIcon)
End Function

' Retrieves file attribute flags:
'       ReadOnly        Hidden      System      Directory
'       Archive         Normal      Temporary   Compressed
Function Attributes() As Long
    If eis = eisFile Then Attributes = fd.dwFileAttributes
End Function

Function length() As Long
    If eis = eisFile Then length = fd.nFileSizeLow
End Function

Function Modified() As Date
    If eis = eisFile Then Modified = MFileTool.Win32ToVbTime(fd.ftLastWriteTime)
End Function

Function Created() As Date
    If eis = eisFile Then Created = MFileTool.Win32ToVbTime(fd.ftCreationTime)
End Function

Function Accessed() As Date
    If eis = eisFile Then Accessed = MFileTool.Win32ToVbTime(fd.ftLastAccessTime)
End Function

Function TotalKilo() As Long
    If eis = eisDrive Then TotalKilo = fd.dwReserved0
End Function

Function FreeKilo() As Long
    If eis = eisDrive Then FreeKilo = fd.dwReserved1
End Function

Function DriveType() As String
    If eis = eisDrive Then DriveType = MBytes.ByteZToStr(fd.cAlternateFileName)
End Function

Function ItemType() As EItemType
    ItemType = eis
End Function

Private Function GetFileItemInfo(vFileItem As Variant, fi As SHFILEINFO, _
                                 ByVal afOption As Long, _
                                 Optional afAttr As Long = 0) As Long
    Dim f As Long
    If VarType(vFileItem) = vbString Then
        afOption = afOption And (Not SHGFI_PIDL)
        f = SHGetFileInfo(CStr(vFileItem), afAttr, fi, LenB(fi), afOption)
    Else
        afOption = afOption Or SHGFI_PIDL
        f = SHGetItemInfo(CLng(vFileItem), 0, fi, LenB(fi), afOption)
    End If
    GetFileItemInfo = f
End Function

#If fComponent = 0 Then
Private Sub ErrRaise(e As Long)
    Dim sText As String, sSource As String
    If e > 1000 Then
        sSource = App.ExeName & ".FileInfo"
        Select Case e
        Case eeBaseFileInfo
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


