Attribute VB_Name = "MFileTool"
Option Explicit

Public Enum EErrorFileTool
    eeBaseFileTool = 13480      ' FileTool
End Enum

Public Enum EWalkModeFile
    ewmfDirs = &H20
    ewmfFiles = &H40
    ewmfBoth = &H20 Or &H40
End Enum

' Use Declare instead of type library because of Strings in UDT

Private Declare Function SHFileOperation Lib "shell32.dll" _
    Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long

Private Type SHFILEOPSTRUCT
    hWnd As Long                ' Window owner of any dialogs
    wFunc As Long               ' Copy, move, rename, or delete code
    pFrom As String             ' Source file
    pTo As String               ' Destination file or directory
    fFlags As Integer           ' Options to control the operations
    fAnyOperationsAbortedLo As Integer ' Indicates partial failure
    fAnyOperationsAbortedHi As Integer
    hNameMappingsLo As Long     ' Array indicating each success
    hNameMappingsHi As Long
    lpszProgressTitleLo As Long ' Title for progress dialog
    lpszProgressTitleHi As Long
End Type

Const datMin As Date = #1/1/100#
Const datMax  As Date = #12/31/9999 11:59:59 PM#

' Difference between day zero for VB dates and Win32 dates
' (or #12-30-1899# - #01-01-1601#)
Const rDayZeroBias As Double = 109205#   ' Abs(CDbl(#01-01-1601#))

' 10000000 nanoseconds * 60 seconds * 60 minutes * 24 hours / 10000
' comes to 86400000 (the 10000 adjusts for fixed point in Currency)
Const rMillisecondPerDay As Double = 10000000# * 60# * 60# * 24# / 10000#

Function Win32ToVbTime(ft As Currency) As Date
    Dim ftl As Currency
    ' Call API to convert from UTC time to local time
    If FileTimeToLocalFileTime(ft, ftl) Then
        ' Local time is nanoseconds since 01-01-1601
        ' In Currency that comes out as milliseconds
        ' Divide by milliseconds per day to get days since 1601
        ' Subtract days from 1601 to 1899 to get VB Date equivalent
        Win32ToVbTime = CDate((ftl / rMillisecondPerDay) - rDayZeroBias)
    Else
        ApiRaise Err.LastDllError
    End If
End Function

Function VbToWin32Time(dat As Date) As Currency
    Dim ftl As Currency
    ' Date is days since 1899
    ' Add days from 1601 to 1899 to get Win32 days
    ' Multiply by milliseconds per day to get milliseconds since 1601
    ' That would be nanoseconds if it weren't in Currency
    ftl = CCur((CDbl(dat) + rDayZeroBias) * rMillisecondPerDay)
    ' Call API to convert from local time to UTC time
    If LocalFileTimeToFileTime(ftl, VbToWin32Time) = 0 Then
        ApiRaise Err.LastDllError
    End If
End Function

Function FileAnyDateTime(sPath As String, _
                         Optional datCreation As Date = datMin, _
                         Optional datAccess As Date = datMin) As Date
    ' Take the easy way if no optional arguments
    If datCreation = datMin And datAccess = datMin Then
        FileAnyDateTime = VBA.FileDateTime(sPath)
        Exit Function
    End If
            
    Dim fnd As WIN32_FIND_DATA
    Dim ftCreate As FILETIME, ftAccess As FILETIME, ftModify As FILETIME
    Dim hFind As Long, f As Boolean, stime As SYSTEMTIME
    ' Get all three times in UDT
    hFind = FindFirstFile(sPath, fnd)
    If hFind = hInvalid Then ApiRaise Err.LastDllError
    FindClose hFind
    ' Convert them to Visual Basic format
    datCreation = Win32ToVbTime(fnd.ftCreationTime)
    datAccess = Win32ToVbTime(fnd.ftLastAccessTime)
    FileAnyDateTime = Win32ToVbTime(fnd.ftLastWriteTime)
End Function

Sub ReplaceFile(sOld As String, sTmp As String)
    Dim fnd As WIN32_FIND_DATA, hFind As Long, hOld As Long, f As Boolean
    ' Get file time and attributes of old file
    hFind = FindFirstFile(sOld, fnd)
    If hFind = hInvalid Then ApiRaise Err.LastDllError
    ' Replace by deleting old and renaming new to old
    Kill sOld
    Name sTmp As sOld
    ' Assign old attributes and time to new file
    hOld = lopen(sOld, OF_WRITE Or OF_SHARE_DENY_WRITE)
    If hOld = hInvalid Then ApiRaise Err.LastDllError
    f = SetFileTime(hOld, fnd.ftCreationTime, _
                    fnd.ftLastAccessTime, fnd.ftLastWriteTime)
    If f = 0 Then ApiRaise Err.LastDllError
    lclose hOld
    f = SetFileAttributes(sOld, fnd.dwFileAttributes)
    If f = 0 Then ApiRaise Err.LastDllError
End Sub

' Better version of FileCopy (CopyAnyFile) and matching MoveAnyFile,
' DeleteAnyFile, and RenameAnyFile

Function CopyAnyFile(sSrc As String, sDst As String, _
                     Optional Options As Long = 0, _
                     Optional Owner As Long = hNull) As Boolean
    If MUtility.HasShell Then
        Dim fo As SHFILEOPSTRUCT, f As Long
        fo.wFunc = FO_COPY
        ' Make sure all strings are double-null-terminated as required
        fo.pFrom = sSrc & vbNullChar
        fo.pTo = sDst & vbNullChar
        fo.fFlags = Options
        fo.hWnd = Owner
        ' Mask out invalid flags
        fo.fFlags = fo.fFlags And FOF_COPYFLAGS
        f = SHFileOperation(fo)
        CopyAnyFile = (f = 0)
    Else
        ' For Windows NT 3.51
        On Error Resume Next
        ' FileCopy expects full name of destination file
        FileCopy sSrc, sDst
        If Err Then
            Err = 0
            ' CopyAnyFile can handle destination directory
            sDst = MUtility.NormalizePath(sDst) & _
                   MUtility.GetFileBaseExt(sSrc)
            FileCopy sSrc, sDst
        End If
        ' Enhance further to emulate SHFileOperation options
        ' such as validation and wild cards
        CopyAnyFile = (Err = 0)
    End If
End Function

Function MoveAnyFile(sSrc As String, sDst As String, _
                  Optional afOptions As Long = 0, _
                  Optional Owner As Long = hNull) As Boolean
    If MUtility.HasShell Then
        Dim fo As SHFILEOPSTRUCT, f As Long
        fo.wFunc = FO_MOVE
        ' Make sure all strings are double-null-terminated as required
        fo.pFrom = sSrc & vbNullChar
        fo.pTo = sDst & vbNullChar
        fo.fFlags = afOptions
        fo.hWnd = Owner
        ' Mask out invalid flags
        fo.fFlags = fo.fFlags And FOF_COPYFLAGS
        f = SHFileOperation(fo)
        MoveAnyFile = (f = 0)
    Else
        ' Windows NT 3.51
        On Error Resume Next
        ' Name actually moves files, contrary to documentation
        Name sSrc As sDst
        If Err Then ' Probably you gave directory destination
            Err = 0
            sDst = MUtility.NormalizePath(sDst) & _
                   MUtility.GetFileBaseExt(sSrc)
            Name sSrc As sDst
        End If
        ' Enhance further to emulate SHFileOperation options
        ' such as validation and wild cards
        MoveAnyFile = (Err = 0)
    End If
End Function

Function RenameAnyFile(sSrc As String, sDst As String, _
                       Optional Options As Long = 0, _
                       Optional Owner As Long = hNull) As Boolean
    If MUtility.HasShell Then
        Dim fo As SHFILEOPSTRUCT, f As Long
        ' Make sure all strings are double-null-terminated as required
        fo.wFunc = FO_RENAME
        fo.pFrom = sSrc & vbNullChar
        fo.pTo = sDst & vbNullChar
        fo.fFlags = Options
        fo.hWnd = Owner
        ' Mask out invalid flags
        fo.fFlags = fo.fFlags And FOF_RENAMEFLAGS
        f = SHFileOperation(fo)
        RenameAnyFile = (f = 0)
    Else
        ' Windows NT 3.51
        On Error Resume Next
        Name sSrc As sDst
        RenameAnyFile = (Err = 0)
        ' Enhance further to emulate SHFileOperation options
        ' such as validation and wild cards
    End If
End Function

Function DeleteAnyFile(sSrc As String, _
                    Optional Options As Long = 0, _
                    Optional Owner As Long = hNull) As Boolean
    If MUtility.HasShell Then
        Dim fo As SHFILEOPSTRUCT, f As Long
        fo.wFunc = FO_DELETE
        ' Make sure all strings are double-null-terminated as required
        fo.pFrom = sSrc & vbNullChar
        ' fo.pTo = sNullStr
        fo.fFlags = Options
        fo.hWnd = Owner
        ' Mask out invalid flags
        fo.fFlags = fo.fFlags And FOF_DELETEFLAGS
        f = SHFileOperation(fo)
        DeleteAnyFile = (f = 0)
    Else
        ' Windows NT 3.51
        On Error Resume Next
        Kill sSrc
        DeleteAnyFile = (Err = 0)
        ' Enhance further to emulate SHFileOperation options
        ' such as validation and wild cards
    End If
End Function

Function Files(hFiles As Long, fi As CFileInfo, _
               ByVal sSpec As String, _
               Optional afAttr As Long = 0) As String
    Dim fd As WIN32_FIND_DATA, sName As String, f As Boolean, sPath As String
    
    ' Stop finding and close handle early
    If afAttr = -1 Then
        f = FindClose(hFiles)
        hFiles = 0: Exit Function
    End If
    f = True
    Do
        ' Get first or next file
        If hFiles = 0 Then
            hFiles = FindFirstFile(sSpec, fd)
        Else
            f = FindNextFile(hFiles, fd)
        End If
        If (f = False Or hFiles = INVALID_HANDLE_VALUE) Then
            If Err.LastDllError = ERROR_NO_MORE_FILES Then
                f = FindClose(hFiles)
            End If
            hFiles = 0: Exit Function
        End If
        ' Keep looping until something matches attributes
    Loop While (afAttr <> vbNormal) And _
               ((afAttr And fd.dwFileAttributes) = 0)
    ' Get file data and return through reference
    sPath = MUtility.GetFileDir(sSpec)
    sName = MUtility.StrZToStr(MBytes.BytesToStr(fd.cFileName))
    fi.CreateFromFile sPath & sName, fd.dwFileAttributes, _
                      fd.nFileSizeLow, fd.ftLastWriteTime, _
                      fd.ftLastAccessTime, fd.ftCreationTime
    Files = sName
End Function

' Efficient find files function
Function FindFiles(sTarget As String, _
                   Optional ByVal start As String) As Collection

    ' Statics for less memory use in recursive procedure
    Static sName As String, sSpec As String, nFound As New Collection
    Static fd As WIN32_FIND_DATA, iLevel As Long
    Dim hFiles As Long, f As Boolean
    If start = sEmpty Then start = CurDir$
    ' Maintain level to ensure collection is cleared first time
    If iLevel = 0 Then
        Set nFound = Nothing
        start = MUtility.NormalizePath(start)
    End If
    iLevel = iLevel + 1
    
    ' Find first file (get handle to find)
    hFiles = FindFirstFile(start & "*.*", fd)
    f = (hFiles <> INVALID_HANDLE_VALUE)
    Do While f
        sName = MBytes.ByteZToStr(fd.cFileName)
        ' Skip . and ..
        If Left$(sName, 1) <> "." Then
            sSpec = start & sName
            If fd.dwFileAttributes And vbDirectory Then
                DoEvents
                ' Call recursively on each directory
                FindFiles sTarget, sSpec & "\"
            ElseIf StrComp(sName, sTarget, 1) = 0 Then ' Text comparison
                ' Store found files in collection
                nFound.Add sSpec
            End If
        End If
        ' Keep looping until no more files
        f = FindNextFile(hFiles, fd)
    Loop
    f = FindClose(hFiles)
    ' Return the matching files in collection
    Set FindFiles = nFound
    iLevel = iLevel - 1
End Function

Function WalkAllFiles(fileit As IUseFile, _
                      Optional ByVal ewmf As EWalkModeFile = ewmfBoth, _
                      Optional ByVal start As String) As Boolean

    ' Statics for less memory use in recursive procedure
    Static sName As String, fd As WIN32_FIND_DATA, iLevel As Long
    Static fi As New CFileInfo
    Dim hFiles As Long, f As Boolean
    If start = sEmpty Then start = CurDir$
    ' Maintain level to ensure collection is cleared first time
    If iLevel = 0 Then start = MUtility.NormalizePath(start)
    iLevel = iLevel + 1
    
    ' Find first file (get handle to find)
    hFiles = FindFirstFile(start & "*.*", fd)
    f = (hFiles <> INVALID_HANDLE_VALUE)
    Do While f
        sName = MBytes.ByteZToStr(fd.cFileName)
        ' Skip . and ..
        If Left$(sName, 1) <> "." Then
            ' Create a file info object from file data
            fi.CreateFromFile start & sName, fd.dwFileAttributes, _
                              fd.nFileSizeLow, fd.ftLastWriteTime, _
                              fd.ftLastAccessTime, fd.ftCreationTime
            If fd.dwFileAttributes And vbDirectory Then
                If ewmf And ewmfDirs Then
                    ' Let client use directory data
                    WalkAllFiles = fileit.UseFile(iLevel, start, fi)
                    ' If client returns True, walk terminates
                    If WalkAllFiles Then Exit Function
                End If
                ' Call recursively on each directory
                WalkAllFiles = WalkAllFiles(fileit, ewmf, _
                                            start & sName & "\")
            Else
                If ewmf And ewmfFiles Then
                    ' Let client use file data
                    WalkAllFiles = fileit.UseFile(iLevel, start, fi)
                    ' If client returns True, walk terminates
                    If WalkAllFiles Then Exit Function
                End If
            End If
        End If
        ' Keep looping until no more files
        f = FindNextFile(hFiles, fd)
    Loop
    f = FindClose(hFiles)
    ' Return the matching files in collection
    iLevel = iLevel - 1
End Function

Function WalkFiles(fileit As IUseFile, _
                   Optional ByVal ewmf As EWalkModeFile = ewmfBoth, _
                   Optional ByVal start As String, _
                   Optional UserData As Variant) As Boolean

    Dim sName As String, sSpec As String, fd As WIN32_FIND_DATA
    Dim hFiles As Long, f As Boolean, fi As New CFileInfo
    If start = sEmpty Then start = CurDir$
    start = MUtility.NormalizePath(start)
    
    ' Find first file (get handle to find)
    hFiles = FindFirstFile(start & "*.*", fd)
    f = (hFiles <> INVALID_HANDLE_VALUE)
    Do While f
        sName = MBytes.ByteZToStr(fd.cFileName)
        ' Skip . and ..
        If Left$(sName, 1) <> "." Then
            ' Create a file info object from file data
            fi.CreateFromFile start & sName, fd.dwFileAttributes, _
                              fd.nFileSizeLow, fd.ftLastWriteTime, _
                              fd.ftLastAccessTime, fd.ftCreationTime
            If fd.dwFileAttributes And vbDirectory Then
                If ewmf And ewmfDirs Then
                    ' Let client use directory data
                    WalkFiles = fileit.UseFile(UserData, start, fi)
                End If
            Else
                If ewmf And ewmfFiles Then
                    ' Let client use file data
                    WalkFiles = fileit.UseFile(UserData, start, fi)
                End If
            End If
            ' If client returns True, walk terminates
            If WalkFiles Then Exit Function
        End If
        ' Keep looping until no more files
        f = FindNextFile(hFiles, fd)
    Loop
    f = FindClose(hFiles)
End Function

#If fComponent = 0 Then
Private Sub ErrRaise(e As Long)
    Dim sText As String, sSource As String
    If e > 1000 Then
        sSource = App.ExeName & ".FileTool"
        Select Case e
        Case eeBaseFileTool
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

