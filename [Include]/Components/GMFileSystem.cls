VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GFileSystem"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Depends
'MApiTime
'MApiFileSystem


Public Enum LNFileType
    ftUnKnown = 0
    ftIE = 2
    ftExE = 4
    ftCHM = 8
    ftIMG = 16
    ftAUDIO = 32
    ftVIDEO = 64
    ftHTML = 128
    ftZIP = 256
    ftTxt = 512
    ftZhtm = 1024
    ftRTF = 3
End Enum
Private Enum LNIfStringNotFound
    ReturnOriginalStr = 1
    ReturnEmptyStr = 0
End Enum
Public Enum LNPathType
    LNUnKnown = 0
    LNFolder = 1
    LNFile = 2
End Enum
Public Enum LNPathStyle
    lnpsDos = 0
    lnpsUnix = 1
End Enum
Public Enum LNLOOKFOR
    LN_FILE_prev
    LN_FILE_next
    LN_FILE_RAND
End Enum


Public Type FindFileData
    Attributes As enumFileAttribute
    CreateTime As Date
    AccessTime As Date
    WriteTime As Date
    FileSize As Long
    filename As String
    ShortName As String
End Type

Private Const cMaxPath = MApiFileSystem.MAX_PATH




Private Declare Function GetFullPathName Lib "kernel32" Alias "GetFullPathNameA" (ByVal lpFileName As String, ByVal nBufferLength As Long, ByVal lpBuffer As String, ByVal lpFilePart As String) As Long



Private Sub WIN32FINDDATA_TO_FINDFILEDATA(ByRef wfdSrc As WIN32_FIND_DATA, ByRef ffdDest As FindFileData)
    With ffdDest
        .filename = CStringToVBString(wfdSrc.cFileName)
        .ShortName = CStringToVBString(wfdSrc.cAlternate)
        If (.ShortName = "") Then .ShortName = .filename
        .Attributes = wfdSrc.dwFileAttributes
        .AccessTime = FileTimeToDate(wfdSrc.ftLastAccessTime)
        .CreateTime = FileTimeToDate(wfdSrc.ftCreationTime)
        .WriteTime = FileTimeToDate(wfdSrc.ftLastWriteTime)
        .FileSize = DWORDPairToDouble(wfdSrc.nFileSizeHigh, wfdSrc.nFileSizeLow)
    End With
End Sub
Public Function QFindFirstFile(ByRef sFilename As String, ByRef sResult As String) As Long
    Dim wfdFind As WIN32_FIND_DATA
    Dim hndFindFile As Long
    hndFindFile = API_FindFirstFile(sFilename, wfdFind)

    If (hndFindFile <> INVALID_HANDLE_VALUE) Then
        QFindFirstFile = hndFindFile
        sResult = CStringToVBString(wfdFind.cFileName)
    Else
        QFindFirstFile = 0
        sResult = ""
    End If
End Function
Public Function QFindNextFile(ByVal hFindFile As Long, ByRef sResult As String) As Boolean
    Dim wfdFind As WIN32_FIND_DATA
    Dim result As Long
    result = API_FindNextFile(hFindFile, wfdFind)
    If (result = 0) Then
        QFindNextFile = False
    Else
        QFindNextFile = True
        sResult = CStringToVBString(wfdFind.cFileName)
    End If
End Function
Public Function QFindClose(ByVal hFindFile As Long) As Boolean
    QFindClose = (Not (API_FindClose(hFindFile) = 0))
End Function
Public Function FindClose(ByVal hFindFile As Long) As Long
    FindClose = API_FindClose(hFindFile)
End Function
Public Function FindFirstFile(ByRef sFilename As String, ByRef hndFindFile As Long, ByRef ffdResult As FindFileData) As Boolean
    Dim wfdFind As WIN32_FIND_DATA
    hndFindFile = API_FindFirstFile(sFilename, wfdFind)
    If (hndFindFile <> INVALID_HANDLE_VALUE) Then
        FindFirstFile = True
        WIN32FINDDATA_TO_FINDFILEDATA wfdFind, ffdResult
    Else
        FindFirstFile = False
    End If
End Function

Public Function FindNextFile(ByVal hFindFile As Long, ByRef ffdResult As FindFileData) As Boolean
    Dim wfdFind As WIN32_FIND_DATA
    Dim result As Long
    result = API_FindNextFile(hFindFile, wfdFind)
    If (result = 0) Then
        FindNextFile = False
    Else
        FindNextFile = True
        WIN32FINDDATA_TO_FINDFILEDATA wfdFind, ffdResult
    End If
End Function



Public Function FileExists(ByRef strPath As String) As Boolean
On Error Resume Next
FileExists = False
If GetAttr(strPath) And vbArchive Then
If Err = 0 Then FileExists = True
End If
End Function
Public Function FolderExists(ByRef strPath As String) As Boolean
On Error Resume Next
FolderExists = False
If GetAttr(strPath) And vbDirectory Then
    If Err = 0 Then FolderExists = True
End If
End Function

Function PathExists(ByRef PathName As String) As Boolean

    Dim Temp$
    'Set Default
    PathExists = True
    Temp$ = Replace$(PathName, "/", "\")

    If Right$(Temp$, 1) = "\" Then Temp$ = Left$(Temp$, Len(Temp$) - 1)
    'Set up error handler
    On Error Resume Next
    'Attempt to grab date and time
    Temp$ = GetAttr(Temp$)
    'Process errors

    If Err <> 0 Then PathExists = False
    '    Select Case Err
    '    Case 53, 76, 68   'File Does Not Exist
    '        modFile_FileExists = False
    '        Err = 0
    '    Case Else
    '
    '        If Err <> 0 Then
    '            MsgBox "Error Number: " & Err & Chr$(10) & Chr$(13) & " " & Error, vbOKOnly, "Error"
    '            End
    '        End If
    '
    '    End Select

End Function

Function AppendSlash(ByVal strPath As String) As String
    Dim c As String
    c = Right$(strPath, 1)
    If Not (c = "/" Or c = "\") Then
        AppendSlash = strPath & "\"
    Else
        AppendSlash = strPath
    End If
End Function

Function refAppendSlash(ByRef strPath As String) As String
    Dim c As String
    c = Right$(strPath, 1)
    If Not (c = "/" Or c = "\") Then
        strPath = strPath & "\"
    End If
    refAppendSlash = strPath
End Function

Function BuildPath(ByVal sPathIn As String, Optional ByVal sFileNameIn As String = "", Optional lnps As LNPathStyle = lnpsDos) As String

    '*******************************************************************
    '
    '  PURPOSE: Takes a path (including Drive letter and any subdirs) and
    '           concatenates the file name to path. Path may be empty, path
    '           may or may not have an ending backslash '\'.  No validation
    '           or existance is check on path or file.
    '
    '  INPUTS:  sPathIn - Path to use
    '           sFileNameIn - Filename to use
    '
    '
    '  OUTPUTS:  N/A
    '
    '  RETURNS:  Path concatenated to File.
    '
    '*******************************************************************
    '    Dim sPath As String
    '    Dim sFilename As String
    '    'Remove any leading or trailing spaces
    '    sPath = Trim$(sPathIn)
    '    sFilename = Trim$(sFileNameIn)
    Dim sSlash As String

    If lnps = lnpsDos Then
        sSlash = "\"
        sPathIn = Replace$(sPathIn, "/", "\")
        sFileNameIn = Replace$(sFileNameIn, "/", "\")
    Else
        sSlash = "/"
        sPathIn = Replace$(sPathIn, "\", "/")
        sFileNameIn = Replace$(sPathIn, "\", "/")
    End If

    If sPathIn = "" Then
        BuildPath = sFileNameIn
    Else

        If Right$(sPathIn, 1) = sSlash Then
            BuildPath = sPathIn & sFileNameIn
        Else
            BuildPath = sPathIn & sSlash & sFileNameIn
        End If

    End If

End Function

Function GetFileName(ByRef sFilename As String) As String

    GetFileName = sFilename
    GetFileName = Replace$(GetFileName, "/", "\")

    If Right$(GetFileName, 1) = "\" Then GetFileName = Left$(GetFileName, Len(GetFileName) - 1)
    GetFileName = RightRight(GetFileName, "\", vbTextCompare, ReturnOriginalStr)

End Function

Function GetParentFolderName(ByRef sFilename As String) As String

    Dim lF As Long
    Dim pos As Long
    lF = Len(sFilename)
    GetParentFolderName = sFilename
    pos = InStrRev(GetParentFolderName, "/")

    If pos = 0 Then pos = InStrRev(GetParentFolderName, "\")

    If pos = lF Then
        GetParentFolderName = Left$(GetParentFolderName, lF - 1)
        pos = InStrRev(GetParentFolderName, "/")

        If pos = 0 Then pos = InStrRev(GetParentFolderName, "\")
    End If

    If pos = 0 Then
        GetParentFolderName = ""
    Else
        GetParentFolderName = Mid$(sFilename, 1, pos - 1) & "\"
    End If

    '
    '    pos = InStrRev(GetParentFolder, "/")
    '    If pos = 0 Then pos = InStrRev(GetParentFolder, "\")
    '    If pos = 0 Then GetParentFolder = ""

End Function

Public Function GetBaseName(ByVal sPath As String) As String

    Dim pos As Long
    Dim ptThis As LNPathType
    ptThis = PathType(sPath)

    If sPath = "" Then Exit Function

    If ptThis = LNUnKnown Then ptThis = LNFile
    GetBaseName = GetFileName(sPath)

    If ptThis = LNFile Then
        pos = InStrRev(GetBaseName, ".")

        If pos > 0 Then GetBaseName = Mid$(GetBaseName, 1, pos - 1)
    End If

End Function

Public Function GetExtensionName(ByRef sPath As String) As String

    If sPath = "" Then Exit Function
    GetExtensionName = RightRight(sPath, ".", vbTextCompare, ReturnEmptyStr)

End Function

Private Function RightRight(ByRef Str As String, RFind As String, Optional Compare As VbCompareMethod = vbBinaryCompare, Optional RetError As LNIfStringNotFound = ReturnEmptyStr) As String

    Dim K As Long
    K = InStrRev(Str, RFind, , Compare)

    If K = 0 Then
        RightRight = IIf(RetError = ReturnOriginalStr, Str, "")
    Else
        RightRight = Mid$(Str, K + 1, Len(Str))
    End If

End Function

Public Function GetTempFileName(Optional sPrefix As String = "lTmp", Optional sExt As String) As String

    Randomize Timer

    If sExt <> "" Then sExt = "." & sExt
    GetTempFileName = sPrefix & Hex$(Int(Rnd(Timer) * 10000 + 1)) & sExt

    Do Until PathExists(GetTempFileName) = False
        GetTempFileName = sPrefix & Hex$(Int(Rnd(Timer) * 10000 + 1)) & sExt
    Loop

End Function

Public Function GetFullPath(sFilename As String) As String

    Dim c As Long, p As Long, sRet As String
    GetFullPath = sFilename

    If sFilename = Empty Then Exit Function
    ' Get the path size, then create string of that size
    sRet = String$(cMaxPath, 0)
    c = GetFullPathName(sFilename, cMaxPath, sRet, p)

    If c = 0 Then Exit Function
    sRet = Left$(sRet, c)
    c = InStr(sRet, Chr$(0))

    If c = 0 Then Exit Function
    sRet = Left$(sRet, c - 1)
    GetFullPath = sRet

End Function

Public Function PathType(sPath As String) As LNPathType

    PathType = LNUnKnown
    On Error GoTo Herr

    If sPath = "" Then Exit Function

    If InStr(sPath, ":") < 1 Then sPath = GetFullPath(sPath)
    Dim PathAttr As VbFileAttribute
    PathAttr = GetAttr(sPath)

    If (PathAttr And vbDirectory) Then
        PathType = LNFolder
    ElseIf (PathAttr And vbArchive) Then
        PathType = LNFile
    End If

Herr:

End Function

Public Function subCount(ByVal spathName As String, Optional ByRef lFolders As Long, Optional ByRef lFiles As Long) As Long

    Dim subName As String

    If PathType(spathName) <> LNFolder Then Exit Function
    spathName = GetFullPath(spathName)
    subName = Dir(spathName, vbDirectory Or vbArchive Or vbHidden Or vbNormal Or vbSystem Or vbReadOnly)

    Do Until subName = ""

        If subName = "." Or subName = ".." Then
        Else
            subCount = subCount + 1
            subName = BuildPath(spathName, subName)

            If PathType(subName) = LNFolder Then
                lFolders = lFolders + 1
            Else
                lFiles = lFiles + 1
            End If

        End If

        subName = Dir()
    Loop

End Function

Public Function testGetSubFilenames(Optional andattr As VbFileAttribute, Optional notattr As VbFileAttribute)
    Dim hnd As Long
    Dim result As FindFileData
    If FindFirstFile("c:\*.*", hnd, result) Then
        Do
            Debug.Print result.filename & " " & CStr(result.CreateTime)
        Loop While FindNextFile(hnd, result)
    End If
End Function

Public Function DirEx(ByVal spathName As String, ByRef strResult() As String, Optional andattr As VbFileAttribute = -1, Optional notattr As VbFileAttribute = -1) As Long
    Dim fdCount As Long
    Dim subName As String
    Dim fAttr As VbFileAttribute
    'spathName = GetFullPath(spathName)

    fAttr = vbReadOnly Or vbHidden Or vbNormal Or vbArchive Or vbDirectory Or vbSystem
    spathName = MFileSystem.BuildPath(spathName, "")
    subName = Dir$(spathName, fAttr)

    'refAppendSlash spathName
    On Error GoTo ERROR_GETATTR
    Do Until subName = ""
        If subName = "." Then GoTo NEXT_NAME
        If subName = ".." Then GoTo NEXT_NAME
        On Error GoTo ERROR_GETATTR
        fAttr = GetAttr(spathName & subName)
             
        If (Not andattr = -1) Then
            If (fAttr And andattr) = 0 Then GoTo NEXT_NAME
        End If
        If (Not notattr = -1) Then
            If (fAttr And notattr) <> 0 Then GoTo NEXT_NAME
        End If
        
     
        ReDim Preserve strResult(0 To fdCount) As String
        strResult(fdCount) = subName
        fdCount = fdCount + 1

        'Debug.Print subName
        
NEXT_NAME:
    subName = Dir$()

    Loop
    DirEx = fdCount
    
Exit Function
ERROR_GETATTR:
    fAttr = 0
    GoTo NEXT_NAME

    
End Function



Public Function subFolders(ByVal spathName As String, ByRef strFolder() As String) As Long
    Dim fdCount As Long
    Dim subName As String
    
    spathName = GetFullPath(spathName)
    subName = Dir$(spathName, vbDirectory)
    Do Until subName = ""
        If subName <> "." And subName <> ".." Then
            fdCount = fdCount + 1
            ReDim Preserve strFolder(1 To fdCount) As String
            strFolder(fdCount) = BuildPath(spathName, subName)
        End If
        subName = Dir$()
    Loop
    subFolders = fdCount
    
End Function
Public Function subFiles(ByVal spathName As String, ByRef strFile() As String) As Long
    Dim fCount As Long
    Dim subName As String
    
    spathName = GetFullPath(spathName)
    subName = Dir$(spathName, vbDirectory)
    Do Until subName = ""
        If subName <> "." And subName <> ".." Then
            fCount = fCount + 1
            ReDim Preserve strFile(1 To fCount) As String
            strFile(fCount) = subName
        End If
        subName = Dir$()
    Loop
    subFiles = fCount
 
End Function

Public Sub xMkdir(sPath As String)
    Dim parentFolder As String
    If FolderExists(sPath) Then Exit Sub
    parentFolder = GetParentFolderName(sPath)
    If parentFolder <> "" And FolderExists(parentFolder) = False Then xMkdir parentFolder
    MkDir sPath
End Sub



Public Function chkFileType(chkfile As String) As LNFileType
    Dim ext As String
    Dim K As Long
    K = InStrRev(chkfile, ".", , vbTextCompare)

    If K > 0 Then
        ext = LCase$(Mid$(chkfile, K + 1, Len(chkfile)))
    End If

    Select Case ext
    Case "rtf"
        chkFileType = ftRTF
    Case "zhtm", "zip"
        chkFileType = ftZIP
    Case "txt", "ini", "bat", "cmd", "css", "log", "cfg", "txtindex"
        chkFileType = ftTxt
    Case "jpg", "jpeg", "gif", "bmp", "png", "ico"
        chkFileType = ftIMG
    Case "htm", "html", "shtml"
        chkFileType = ftIE
    Case "exe", "com"
        chkFileType = ftExE
    Case "chm"
        chkFileType = ftCHM
    Case "mp3", "wav", "wma"
        chkFileType = ftAUDIO
    Case "wmv", "rm", "rmvb", "avi", "mpg", "mpeg"
        chkFileType = ftVIDEO
    End Select

End Function

Public Function LookFor(sCurFile As String, Optional lookForWhat As LNLOOKFOR = LN_FILE_next, Optional sWildcard As String = "*")

Dim sCurFilename As String
Dim sCurFolder As String
Dim i As Long
Dim iCount As Long
Dim sFileList() As String
Dim Index As String

If PathExists(sCurFile) = False Then Exit Function

If PathType(sCurFile) = LNFolder Then
    sCurFolder = sCurFile
ElseIf PathType(sCurFile) = LNFile Then
    sCurFolder = GetParentFolderName(sCurFile)
    sCurFilename = GetFileName(sCurFile)
Else
    Exit Function
End If

iCount = subFiles(BuildPath(sCurFolder, sWildcard), sFileList())
If iCount < 1 Then Exit Function
Index = 0
If lookForWhat = LN_FILE_RAND Then
    Index = Int(Rnd(Timer) * iCount) + 1
ElseIf sCurFilename = "" Then
        Index = 1
Else
    For i = 1 To iCount
        If StrComp(sCurFilename, sFileList(i), vbTextCompare) = 0 Then
            Index = i: Exit For
        End If
    Next
End If

If lookForWhat = LN_FILE_next Then
    Index = Index + 1
    If Index > iCount Then Index = 1
ElseIf lookForWhat = LN_FILE_prev Then
    Index = Index - 1
    If Index < 1 Then Index = iCount
End If

LookFor = BuildPath(sCurFolder, sFileList(Index))

End Function


Public Function SyncDirectory(ByVal vSrc As String, ByVal vDest As String, ByRef bDelete As Boolean, Optional ByRef ObjCallBack As Object, Optional ByRef CopyingMethod As String, Optional ByRef CheckingMethod As String) As Long
        
    Dim Dires() As String
    Dim Files() As String
    Dim cFiles As Long
    Dim cDires As Long
    Dim count As Long
    Dim doCopyCallBack As Boolean
    Dim doCheckCallback As Boolean
    
    count = -1
    SyncDirectory = count
    
    Dim FileA As String
    Dim FileB As String
    Dim DateA As Date
    Dim DateB As Date
    Dim SizeA As Long
    Dim SizeB As Long
    
    
    If Not FolderExists(vSrc) Then Exit Function
    If Not FolderExists(vDest) Then Exit Function
    
    refAppendSlash vSrc
    refAppendSlash vDest
    
    cFiles = DirEx(vSrc, Files(), , vbDirectory)
    cDires = DirEx(vSrc, Dires(), vbDirectory)
    

    count = 0
    doCopyCallBack = (IsObject(ObjCallBack) And (CopyingMethod <> ""))
    doCheckCallback = (IsObject(ObjCallBack) And (CheckingMethod <> ""))
    
    Dim i As Long
    For i = 1 To cFiles
        FileA = vSrc & Files(i)
        FileB = vDest & Files(i)
        If doCheckCallback Then CallByName ObjCallBack, CheckingMethod, VbMethod, FileB
        If Not FileExists(FileB) Then GoTo COPYFILE
        DateA = FileDateTime(FileA)
        DateB = FileDateTime(FileB)
        If (DateA <> DateB) Then GoTo COPYFILE
        SizeA = FileLen(FileA)
        SizeB = FileLen(FileB)
        If (SizeA <> SizeB) Then GoTo COPYFILE
NEXTFILE:
    Next
    
    For i = 1 To cDires
        FileA = vSrc & Dires(i)
        FileB = vDest & Dires(i)
        If doCheckCallback Then CallByName ObjCallBack, CheckingMethod, VbMethod, FileB
        If Not FolderExists(FileB) Then GoTo COPYDIR
NEXTDIR:
    Next
    
    For i = 1 To cDires
        count = count + SyncDirectory(vSrc & Dires(i), vDest & Dires(i), bDelete, ObjCallBack, CopyingMethod, CheckingMethod)
    Next
    
    SyncDirectory = count
Exit Function
COPYFILE:
    If doCopyCallBack Then CallByName ObjCallBack, CopyingMethod, VbMethod, FileB
    FileCopy FileA, FileB
    count = count + 1
    GoTo NEXTFILE
COPYDIR:
    If doCopyCallBack Then CallByName ObjCallBack, CopyingMethod, VbMethod, FileB
    MkDir FileB
    count = count + 1
    GoTo NEXTDIR
End Function


Public Function RGetFolderContent(ByRef sFiles() As String, ByVal sFolder As String, Optional YesAttrs As VbFileAttribute = -1, Optional NoAttrs As VbFileAttribute = -1, Optional nLevel As Integer = -1) As Long
    Dim count As Long
    Dim i As Long, j As Long
    Dim sSubFolders() As String
    Dim sSubFiles() As String
    Dim cSubFiles As Long
    'Dim cSubFolders As Long
    sFolder = BuildPath(sFolder)
    count = DirEx(sFolder, sFiles(), YesAttrs, NoAttrs)
    If (DirEx(sFolder, sSubFolders, vbDirectory) > 0) Then
        If nLevel <> 0 Then
            nLevel = nLevel - 1
            For i = LBound(sSubFolders) To UBound(sSubFolders)
                cSubFiles = RGetFolderContent(sSubFiles, sFolder & sSubFolders(i), YesAttrs, NoAttrs, nLevel)
                If cSubFiles > 0 Then
                    ReDim Preserve sFiles(0 To count + cSubFiles - 1)
                    For j = 0 To cSubFiles - 1
                        sFiles(count + j) = sSubFolders(i) & "\" & sSubFiles(j)
                    Next
                    count = count + cSubFiles
                End If
            Next
        End If
    End If
    RGetFolderContent = count
End Function

Public Sub TEST_RGFC()
    Dim a() As String
    RGetFolderContent a(), "X:\ShortCuts", , vbDirectory
    Debug_DumpArray (a())
End Sub


