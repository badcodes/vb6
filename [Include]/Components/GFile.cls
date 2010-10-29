VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GFile"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Const CryptKey1 = 29 '|(Asc ("L") + Asc("I") + Asc("N")-256|
Public Const CryptKey2 = 49  '|Asc("X") + Asc("I") + Asc("A") + Asc("O") - 256|
Public Const CryptKey3 = 31 '|Asc ("R") + Asc("A") + Asc("N")-256|
Public Const CryptFlag = "LCF" 'Lin Crypt File
Public CryptProgress As Integer

Enum chkFileType
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
Public Const MAXTEXTBLOCK = 10240
Public Const cMaxPath = 260
Public Declare Function GetFullPathName Lib "kernel32" Alias "GetFullPathNameA" (ByVal lpFileName As String, ByVal nBufferLength As Long, ByVal lpBuffer As String, ByVal lpFilePart As String) As Long

Sub rebuildfile(filepath As String, skipline As Integer)

    Dim fpath As String
    Dim fso As New Scripting.FileSystemObject
    Dim FileList() As String
    Dim FileCount As Integer
    Dim fs As Files
    Dim ff As File
    fpath = filepath
    Set fs = fso.GetFolder(fpath).Files
    FileCount = fs.count

    If FileCount < 1 Then Exit Sub
    ReDim FileList(FileCount) As String
    Dim i As Long

    For Each ff In fs
        i = i + 1
        FileList(i) = ff.Name
    Next

    fpath = bddir(fpath)
    Dim srcTS As Scripting.TextStream
    Dim dstTS As Scripting.TextStream
    Dim norb As Boolean
    Dim tmpstr As String
    Dim j As Long

    For i = 1 To FileCount
        norb = False
        Set srcTS = fso.OpenTextFile(fpath + FileList(i), ForReading)

        For j = 1 To skipline

            If srcTS.AtEndOfStream Then
                norb = True
                Exit For
            End If

            srcTS.skipline
        Next

        If srcTS.AtEndOfStream Then norb = True

        If norb = False Then
            tmpstr = srcTS.ReadLine
            tmpstr = LTrim$(tmpstr)
            tmpstr = RTrim$(tmpstr)

            If Right$(tmpstr, 1) = Chr$(13) Then tmpstr = Left$(tmpstr, Len(tmpstr) - 1)
            tmpstr = StrConv(tmpstr, vbWide)
            Set dstTS = fso.CreateTextFile(fpath + tmpstr + ".txt", True)
            dstTS.WriteLine tmpstr

            Do Until srcTS.AtEndOfStream
                tmpstr = srcTS.ReadLine
                dstTS.WriteLine tmpstr
            Loop

            dstTS.Close
            srcTS.Close
            fso.DeleteFile fpath + FileList(i), True
        End If

    Next

End Sub


Public Function MyFileEncrypt(srcfile As String, dstfile As String) As Boolean

    Dim tmpFile As String
    Dim thebyte As Byte
    MyFileEncrypt = True

    If Dir$(srcfile) = "" Then MyFileEncrypt = False: Exit Function
    tmpFile = "~$$$CRfile.tmp"

    If Dir$(tmpFile) <> "" Then Kill tmpFile
    Open srcfile For Binary As #1
    Open tmpFile For Binary As #2
    Put #2, , CryptFlag '标识符

    Do Until Loc(1) = LOF(1)
        CryptProgress = Int(Loc(1) * 100 / LOF(1))
        Get #1, , thebyte
        thebyte = thebyte Xor CryptKey1
        thebyte = thebyte Xor CryptKey2
        thebyte = thebyte Xor CryptKey3
        Put #2, , thebyte
    Loop

    Close #1
    Close #2

    If Dir$(dstfile) <> "" Then Kill dstfile
    FileCopy tmpFile, dstfile
    Kill tmpFile

End Function

Public Function MyFileDecrypt(srcfile As String, dstfile As String) As Boolean

    Dim tmpFile As String
    Dim thebyte As Byte
    Dim skipflag
    MyFileDecrypt = True

    If Dir$(srcfile) = "" Then MyFileDecrypt = False: Exit Function

    If isLXTfile(srcfile) = False Then MyFileDecrypt = False: Exit Function
    Open srcfile For Binary As #1
    tmpFile = "~$$$CRfile.tmp"

    If Dir$(tmpFile) <> "" Then Kill tmpFile
    Open tmpFile For Binary As #2
    skipflag = Input(Len(CryptFlag), #1)

    Do Until Loc(1) = LOF(1)
        CryptProgress = Int(Loc(1) * 100 / LOF(1))
        Get #1, , thebyte
        thebyte = thebyte Xor CryptKey3
        thebyte = thebyte Xor CryptKey2
        thebyte = thebyte Xor CryptKey1
        Put #2, , thebyte
    Loop

    Close #1
    Close #2

    If Dir$(dstfile) <> "" Then Kill dstfile
    FileCopy tmpFile, dstfile
    Kill tmpFile

End Function

Public Function isLXTfile(thefile As String) As Boolean

    Dim fso As New Scripting.FileSystemObject
    Dim f As File
    isLXTfile = False

    If fso.FileExists(thefile) = False Then Exit Function
    Set f = fso.GetFile(thefile)

    If f.Size < Len(CryptFlag) Then Exit Function

    If f.OpenAsTextStream(ForReading).Read(Len(CryptFlag)) = CryptFlag Then isLXTfile = True

End Function

Public Function chkFileType(chkfile As String) As chkFileType

    Dim ext As String
    ext = LCase$(RightRight(chkfile, ".", vbBinaryCompare, ReturnEmptyStr))

    Select Case ext
    Case "rtf"
        chkFileType = ftRTF
    Case "zhtm", "zip"
        chkFileType = ftZIP
    Case "txt", "ini", "bat", "cmd", "css", "log", "cfg"
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
    Case "wma", "rm", "rmvb", "avi", "mpg", "mpeg"
        chkFileType = ftVIDEO
    End Select

End Function



Function opsget(filenum As Integer) As String

    Dim thebyte As Byte
    Dim tempstr As String
    Get 1, , thebyte

    If thebyte > 127 Then
        Seek 1, Loc(1) - 1
        tempstr = Input(1, 1)
        Seek 1, Loc(1) - 2
    Else
        tempstr = Chr$(thebyte)
        Seek 1, Loc(1) - 1
    End If

    opsget = tempstr

End Function



Sub splitfile(thefile As String, SplitFlag As String)

    Dim fso As New Scripting.FileSystemObject
    Dim ts As Scripting.TextStream
    Dim tempts As Scripting.TextStream
    Dim tempstr As String
    Dim thefolder As String
    Dim tempfile As String
    Dim SplitTS As Scripting.TextStream
    Dim n As Integer

    If fso.FolderExists(bddir(fso.GetParentFolderName(thefile)) + fso.GetBaseName(thefile)) = False Then
        fso.CreateFolder bddir(fso.GetParentFolderName(thefile)) + fso.GetBaseName(thefile)
    End If

    thefolder = bddir(fso.GetParentFolderName(thefile)) + fso.GetBaseName(thefile)
    tempfile = bddir(thefolder) + SplitFlag + StrNum(0, 3) + "." + fso.GetExtensionName(thefile)
    Set tempts = fso.OpenTextFile(tempfile, ForWriting, True)
    Set ts = fso.OpenTextFile(thefile, ForReading)
    tempstr = ts.ReadLine

    Do Until ts.AtEndOfStream

        Do Until Left$(LTrim$(tempstr), Len(SplitFlag)) = SplitFlag Or ts.AtEndOfStream
            tempts.WriteLine tempstr
            tempstr = ts.ReadLine
        Loop

        If ts.AtEndOfStream = False Then
            n = n + 1
            tempts.WriteLine tempstr
            Set SplitTS = fso.OpenTextFile(bddir(thefolder) + SplitFlag + StrNum(n, 3) + "." + fso.GetExtensionName(thefile), ForWriting, True)
            SplitTS.WriteLine tempstr
            tempstr = ts.ReadLine

            Do Until Left$(LTrim$(tempstr), Len(SplitFlag)) = SplitFlag Or ts.AtEndOfStream
                SplitTS.WriteLine tempstr
                tempstr = ts.ReadLine
            Loop

            If ts.AtEndOfStream Then SplitTS.WriteLine tempstr
            SplitTS.Close
        End If

    Loop

    tempts.Close
    ts.Close

End Sub

Public Function treeSearch(ByVal sPath As String, ByVal SFileSpec As String, sSubDirs() As String, lDirCount As Long, sFiles() As String, lFileCount As Long) As Boolean

    Dim fstFiles As Long
    Dim fstIndex As Long '文件数目
    Dim sDir As String
    Dim i As Long
    Dim IndexBegin As Long
    Dim IndexStop As Long
    fstFiles = lFileCount
    fstIndex = lDirCount
    IndexBegin = fstIndex

    If Right$(sPath, 1) <> "\" Then sPath = sPath + "\"
    sDir = Dir$(sPath + SFileSpec)
    '获得当前目录下文件名和数目

    Do While Len(sDir)
        ReDim Preserve sFiles(0 To fstFiles)
        sFiles(fstFiles) = sPath + sDir
        fstFiles = fstFiles + 1
        sDir = Dir
    Loop

    '获得当前目录下的子目录名称
    sDir = Dir$(sPath + "*.*", 16)

    Do While Len(sDir)

        If Left$(sDir, 1) <> "." Then 'skip.and..
            '找出子目录名

            If (GetAttr(sPath + sDir) And vbDirectory) <> 0 Then
                '保存子目录名
                ReDim Preserve sSubDirs(0 To fstIndex)
                sSubDirs(fstIndex) = sPath + sDir + "\"
                fstIndex = fstIndex + 1
            End If

        End If

        sDir = Dir
    Loop

    lDirCount = fstIndex
    lFileCount = fstFiles
    IndexStop = fstIndex - 1

    For i = IndexBegin To IndexStop '查找每一个子目录下文件，这里利用了递归
        Call treeSearch(sSubDirs(i), SFileSpec, sSubDirs(), lDirCount, sFiles(), lFileCount)
    Next

    treeSearch = True

End Function

Public Function treeSearchFiles(ByVal sPath As String, ByVal SFileSpec As String, sFiles() As String, lFileCount As Long) As Boolean
    
    Dim sSubDirs() As String
    Dim lDirCount As Long
    Dim fstFiles As Long
    Dim fstIndex As Long '文件数目
    Dim sDir As String
    Dim i As Long
    Dim IndexBegin As Long
    Dim IndexStop As Long
    fstFiles = lFileCount
    fstIndex = lDirCount
    IndexBegin = fstIndex

    If Right$(sPath, 1) <> "\" Then sPath = sPath + "\"
    sDir = Dir$(sPath + SFileSpec)
    '获得当前目录下文件名和数目

    Do While Len(sDir)
        ReDim Preserve sFiles(0 To fstFiles)
        sFiles(fstFiles) = sPath + sDir
        fstFiles = fstFiles + 1
        sDir = Dir
    Loop

    '获得当前目录下的子目录名称
    sDir = Dir$(sPath + "*.*", 16)

    Do While Len(sDir)

        If Left$(sDir, 1) <> "." Then 'skip.and..
            '找出子目录名

            If (GetAttr(sPath + sDir) And vbDirectory) <> 0 Then
                '保存子目录名
                ReDim Preserve sSubDirs(0 To fstIndex)
                sSubDirs(fstIndex) = sPath + sDir + "\"
                fstIndex = fstIndex + 1
            End If

        End If

        sDir = Dir
    Loop

    lDirCount = fstIndex
    lFileCount = fstFiles
    IndexStop = fstIndex - 1

    For i = IndexBegin To IndexStop '查找每一个子目录下文件，这里利用了递归
        Call treeSearchFiles(sSubDirs(i), SFileSpec, sFiles(), lFileCount)
    Next

    treeSearchFiles = True

End Function

Public Function treeSearchAll(ByVal sPath As String, ByVal SFileSpec As String, sAll() As String, lCount As Long) As Boolean
    
    Dim sSubDirs() As String
    Dim lDirCount As Long
    Dim fstFiles As Long
    Dim fstIndex As Long '文件数目
    Dim sDir As String
    Dim i As Long
    Dim IndexBegin As Long
    Dim IndexStop As Long
    fstFiles = lCount
    fstIndex = lDirCount
    IndexBegin = fstIndex

    If Right$(sPath, 1) <> "\" Then sPath = sPath + "\"
    sDir = Dir$(sPath + SFileSpec)
    '获得当前目录下文件名和数目

    Do While Len(sDir)
        ReDim Preserve sAll(0 To fstFiles)
        sAll(fstFiles) = sPath + sDir
        fstFiles = fstFiles + 1
        sDir = Dir
    Loop

    '获得当前目录下的子目录名称
    sDir = Dir$(sPath + "*.*", 16)

    Do While Len(sDir)

        If Left$(sDir, 1) <> "." Then 'skip.and..
            '找出子目录名

            If (GetAttr(sPath + sDir) And vbDirectory) <> 0 Then
                '保存子目录名
                ReDim Preserve sSubDirs(0 To fstIndex)
                sSubDirs(fstIndex) = sPath + sDir + "\"
                fstIndex = fstIndex + 1
            End If

        End If

        sDir = Dir
    Loop

    lDirCount = fstIndex
    lCount = fstFiles
    IndexStop = fstIndex - 1

    For i = IndexBegin To IndexStop '查找每一个子目录下文件，这里利用了递归
        Call treeSearchAll(sSubDirs(i), SFileSpec, sAll(), lCount)
    Next

    treeSearchAll = True

End Function

Sub delline(filepath As String, skipline As Integer)

    Dim fpath As String
    Dim fso As New Scripting.FileSystemObject
    Dim FileList() As String
    Dim FileCount As Integer
    Dim fs As Files
    Dim ff As File
    fpath = filepath
    Set fs = fso.GetFolder(fpath).Files
    FileCount = fs.count

    If FileCount < 1 Then Exit Sub
    ReDim FileList(FileCount) As String
    Dim i As Long

    For Each ff In fs
        i = i + 1
        FileList(i) = ff.Name
    Next

    fpath = bddir(fpath)
    Dim srcTS As Scripting.TextStream
    Dim dstTS As Scripting.TextStream
    Dim norb As Boolean
    Dim tmpstr As String

    For i = 1 To FileCount
        norb = False
        Set srcTS = fso.OpenTextFile(fpath + FileList(i), ForReading)
        Dim j As Long

        For j = 1 To skipline

            If srcTS.AtEndOfStream Then
                norb = True
                Exit For
            End If

            srcTS.skipline
        Next

        If srcTS.AtEndOfStream Then norb = True

        If norb = False Then
            Dim dstfile As String
            dstfile = fso.GetTempName
            Set dstTS = fso.CreateTextFile(dstfile, True)

            Do Until srcTS.AtEndOfStream
                tmpstr = srcTS.ReadLine
                dstTS.WriteLine tmpstr
            Loop

            dstTS.Close
            srcTS.Close
            fso.DeleteFile fpath + FileList(i), True
            fso.MoveFile dstfile, fpath + FileList(i)
        End If

    Next

End Sub

Public Sub RenameBat(thedir As String, renameflag As String)

    Dim fso As New Scripting.FileSystemObject
    Dim fs As Files
    Dim f As File
    Dim tmpline As String
    Dim ts As Scripting.TextStream

    If fso.FolderExists(thedir) = False Then Exit Sub
    Set fs = fso.GetFolder(thedir).Files

    For Each f In fs
        Set ts = f.OpenAsTextStream(ForReading)
        Dim m As Long
        m = 0

        Do Until m > 20

            If ts.AtEndOfStream Then Exit Do
            m = m + 1
            tmpline = ts.ReadLine

            If InStr(tmpline, renameflag) > 0 Then
                ts.Close
                tmpline = ldel(rdel(tmpline))
                Dim dstfile As String
                dstfile = bddir(fso.GetParentFolderName(f.Path)) + StrConv(tmpline, vbWide) + "." + fso.GetExtensionName(f.Path)

                If fso.FileExists(dstfile) Then
                    fso.DeleteFile f.Path
                Else
                    fso.MoveFile f.Path, dstfile
                End If

                m = 21
            End If

        Loop

    Next

End Sub

Public Sub BatRenamebyFile(thedir As String, thefile As String, SeperateFlag As String)

    Dim fso As New Scripting.FileSystemObject
    Dim ts As Scripting.TextStream
    Dim tempstr As String
    Dim srcfile As String
    Dim dstfile As String
    Dim pos As Integer

    If fso.FileExists(bddir(thedir) + thefile) = False Then Exit Sub
    Set ts = fso.OpenTextFile(bddir(thedir) + thefile, ForReading)

    Do Until ts.AtEndOfStream
        tempstr = ts.ReadLine
        pos = InStr(tempstr, SeperateFlag)

        If pos > 0 Then
            srcfile = Left$(tempstr, pos - 1)
            dstfile = Right$(tempstr, Len(tempstr) - pos + Len(SeperateFlag) - 1)

            If srcfile <> dstfile And fso.FileExists(bddir(thedir) + srcfile) = True Then
                srcfile = bddir(thedir) + srcfile
                dstfile = bddir(thedir) + StrConv(fso.GetBaseName(dstfile), vbWide) + "." + fso.GetExtensionName(dstfile)

                If fso.FileExists(dstfile) = False Then fso.MoveFile srcfile, dstfile
            End If

        End If

    Loop

End Sub

Public Function delblankline(thefile As String, Optional dstfile As String = "") As Boolean

    Dim fso As New Scripting.FileSystemObject

    If fso.FileExists(thefile) = False Then Exit Function

    If dstfile = "" Then dstfile = thefile
    Dim ts As Scripting.TextStream
    Dim tempts As Scripting.TextStream
    Dim tempfile As String
    Dim tempstr As String
    Dim realstr As String
    Dim blankline As Boolean
    tempfile = fso.GetTempName
    Set ts = fso.OpenTextFile(thefile, ForReading)
    Set tempts = fso.OpenTextFile(tempfile, ForWriting, True)

    Do Until ts.AtEndOfStream
        tempstr = ts.ReadLine
        realstr = RTrim$(LTrim$(tempstr))
        blankline = False

        If realstr = "" Then blankline = True

        If realstr = Chr$(13) Then blankline = True

        If realstr = Chr$(10) Then blankline = True

        If realstr = Chr$(13) + Chr$(10) Then blankline = True

        If Not blankline Then tempts.WriteLine tempstr
    Loop

    ts.Close
    tempts.Close
    fso.DeleteFile thefile
    fso.MoveFile tempfile, dstfile
    delblankline = True

End Function

Public Function BATdelblankline(thedir As String) As Boolean

    Dim fso As New Scripting.FileSystemObject

    If fso.FolderExists(thedir) = False Then Exit Function
    Dim f As File
    Dim fs As Files
    Set fs = fso.GetFolder(thedir).Files
    Dim m As Long

    For Each f In fs
        m = m + 1
        Debug.Print Str$(m) + "/" + Str$(fs.count) + ":" + f.Path
        delblankline f.Path
    Next

End Function

Public Sub RenameByFstLine(thefile As String)

    Dim fso As New Scripting.FileSystemObject

    If fso.FileExists(thefile) = False Then Exit Sub
    Dim ts As Scripting.TextStream
    Dim tempstr As String
    Dim dstfile As String
    Set ts = fso.OpenTextFile(thefile, ForReading)

    Do Until ts.AtEndOfStream
        tempstr = ts.ReadLine
        tempstr = rdel(ldel(tempstr))

        If tempstr <> "" Then
            tempstr = StrConv(tempstr, vbWide)
            dstfile = bddir(fso.GetParentFolderName(thefile)) + tempstr + "." + fso.GetExtensionName(thefile)

            If dstfile <> thefile And fso.FileExists(dstfile) = False Then
                ts.Close
                fso.MoveFile thefile, dstfile
            End If

            Exit Do
        End If

    Loop

End Sub


Function GetFullPath(sFilename As String) As String

    Dim c As Long, p As Long, sRet As String
    GetFullPath = sFilename

    If sFilename = Empty Then Exit Function
    ' Get the path size, then create string of that size
    sRet = String$(cMaxPath, 0)
    c = GetFullPathName(sFilename, cMaxPath, sRet, p)

    If c = 0 Then Exit Function
    sRet = Left$(sRet, c)
    GetFullPath = sRet

End Function

Sub FileStrReplace(thefile As String, thetext As String, RPText As String)

    Const MAXSTRING = 28800

    If thetext = RPText Then Exit Sub

    If thetext = "" Then Exit Sub

    If Len(thetext) >= MAXSTRING \ 2 Then MsgBox ("The text to replace is too large!"): Exit Sub
    Dim fso As New Scripting.FileSystemObject
    'Dim MatchNum As Integer

    If fso.FileExists(thefile) = False Then Exit Sub
    Dim ff As File
    Set ff = fso.GetFile(thefile)

    If ff.Size < Len(thetext) Then Exit Sub
    Dim BlockSize As Long
    Dim textSize As Long
    Dim blocknum As Long
    BlockSize = MAXSTRING

    If ff.Size < BlockSize Then BlockSize = ff.Size
    textSize = Len(thetext)
    blocknum = (ff.Size - 1) \ (BlockSize) + 1
    Dim tempstring As String
    Dim reststring As String
    Dim srcTS As Scripting.TextStream
    Dim dstTS As Scripting.TextStream
    Dim tempfile As String
    Dim iLastPos As Integer
    tempfile = fso.GetTempName
    Set srcTS = ff.OpenAsTextStream(ForReading)
    Set dstTS = fso.CreateTextFile(tempfile, True)
    Dim lEnd As Long
    Dim i As Long
    lEnd = blocknum + 1

    For i = 1 To lEnd

        If srcTS.AtEndOfStream Then Exit For
        tempstring = reststring + srcTS.Read(BlockSize)
        iLastPos = InStrRev(tempstring, thetext)

        If iLastPos > 0 Then
            iLastPos = Len(tempstring) - iLastPos - Len(thetext)
            tempstring = Replace(tempstring, thetext, RPText, , , vbTextCompare)

            If iLastPos > textSize Then iLastPos = textSize

            If iLastPos = 0 Then
                reststring = ""
            Else
                reststring = Right$(tempstring, iLastPos)
            End If

            tempstring = Left$(tempstring, Len(tempstring) - iLastPos)
            dstTS.Write tempstring
        Else

            If Len(tempstring) < textSize Then
                reststring = tempstring
                tempstring = ""
            Else
                reststring = Right$(tempstring, textSize)
                tempstring = Left$(tempstring, Len(tempstring) - textSize)
                dstTS.Write tempstring
            End If

        End If

    Next

    dstTS.Write reststring
    dstTS.Close
    srcTS.Close
    fso.DeleteFile thefile
    fso.MoveFile tempfile, thefile

End Sub

Function FileInStr(thefile As String, thetext As String, Min_MatchTimes As Integer, Optional CompMethod As VbCompareMethod = vbBinaryCompare) As Boolean

    If CompMethod = 0 Then CompMethod = vbBinaryCompare
    Const MAXSTRING = 32768

    If thetext = "" Then Exit Function

    If Min_MatchTimes = 0 Then Exit Function

    If Len(thetext) >= MAXSTRING \ 2 Then MsgBox ("The text to Search is too large!"): Exit Function
    Dim fso As New Scripting.FileSystemObject

    If fso.FileExists(thefile) = False Then Exit Function
    Dim ff As File
    Set ff = fso.GetFile(thefile)

    If ff.Size < Len(thetext) Then Exit Function
    Dim BlockSize As Long
    Dim textSize As Long
    Dim blocknum As Long
    BlockSize = MAXSTRING

    If ff.Size < BlockSize Then BlockSize = ff.Size
    textSize = Len(thetext)
    blocknum = ff.Size \ (BlockSize) + 1
    Dim tempstring As String
    Dim reststring As String
    Dim srcTS As Scripting.TextStream
    Dim MatchTimes As Integer
    Dim pos As Integer
    Set srcTS = ff.OpenAsTextStream(ForReading)
    reststring = Space$(textSize)
    Dim i As Long

    For i = 1 To blocknum

        If srcTS.AtEndOfStream Then Exit For
        tempstring = reststring + srcTS.Read(BlockSize)
        pos = InStr(1, tempstring, thetext, CompMethod)

        Do Until pos = 0
            MatchTimes = MatchTimes + 1

            If MatchTimes >= Min_MatchTimes Then FileInStr = True: srcTS.Close: Exit Function
            pos = InStr(pos + textSize, tempstring, thetext, CompMethod)
        Loop

        If LCase$(Right$(tempstring, textSize)) <> LCase$(thetext) Then
            reststring = Right$(tempstring, textSize)
        End If

    Next

    srcTS.Close

End Function

Function FileInStrTimes(thefile As String, thetext As String, Optional CountToStop As Integer, Optional CompMethod As VbCompareMethod = vbBinaryCompare, Optional t As Tristate = TristateMixed) As Integer

    If CompMethod = 0 Then CompMethod = vbBinaryCompare
    Const MAXSTRING = 32768
    
    If CountToStop = 0 Then CountToStop = 1000
    
    If thetext = "" Then Exit Function

    If Len(thetext) >= MAXSTRING \ 2 Then MsgBox ("The text to Search is too large!"): Exit Function
    Dim fso As New Scripting.FileSystemObject

    If fso.FileExists(thefile) = False Then Exit Function
    Dim ff As File
    Set ff = fso.GetFile(thefile)

    If ff.Size < Len(thetext) Then Exit Function
    Dim BlockSize As Long
    Dim textSize As Long
    Dim blocknum As Long
    BlockSize = MAXSTRING

    If ff.Size < BlockSize Then BlockSize = ff.Size
    textSize = Len(thetext)
    blocknum = ff.Size \ (BlockSize) + 1
    Dim tempstring As String
    Dim reststring As String
    Dim srcTS As Scripting.TextStream
    Dim MatchTimes As Integer
    Dim pos As Integer
    Set srcTS = ff.OpenAsTextStream(ForReading, t)
    reststring = Space$(textSize)
    Dim i As Long

    For i = 1 To blocknum

        If srcTS.AtEndOfStream Then Exit For
        tempstring = reststring + srcTS.Read(BlockSize)
        pos = InStr(1, tempstring, thetext, CompMethod)

        Do Until pos = 0
            MatchTimes = MatchTimes + 1
            If MatchTimes > CountToStop Then Exit Do
            pos = InStr(pos + textSize, tempstring, thetext, CompMethod)
        Loop

        If LCase$(Right$(tempstring, textSize)) <> LCase$(thetext) Then
            reststring = Right$(tempstring, textSize)
        End If
        
        If MatchTimes > CountToStop Then Exit For
    Next

    srcTS.Close
    FileInStrTimes = MatchTimes

End Function

Function modFile_FileExists(ByVal FileName As String) As Boolean

    Dim Temp$
    'Set Default
    modFile_FileExists = True
    'Set up error handler
    On Error Resume Next
    'Attempt to grab date and time
    Temp$ = FileDateTime(FileName)
    'Process errors

    Select Case Err
    Case 53, 76, 68   'File Does Not Exist
        modFile_FileExists = False
        Err = 0
    Case Else

        If Err <> 0 Then
            MsgBox "Error Number: " & Err & Chr$(10) & Chr$(13) & " " & Error, vbOKOnly, "Error"
            End
        End If

    End Select

End Function

Function modFile_buildpath(ByVal sPathIn As String, ByVal sFileNameIn As String) As String

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
    Dim sPath As String
    Dim sFilename As String
    'Remove any leading or trailing spaces
    sPath = Trim$(sPathIn)
    sFilename = Trim$(sFileNameIn)

    If sPath = "" Then
        modFile_buildpath = sFilename
    Else

        If Right$(sPath, 1) = "\" Then
            modFile_buildpath = sPath & sFilename
        Else
            modFile_buildpath = sPath & "\" & sFilename
        End If

    End If

End Function

Function ExtractFileName(sFilename As String) As String

    '*******************************************************************
    '
    '  PURPOSE: This returns just a file name from a full/partial path.
    '
    '  INPUTS:  sFileName - String Data to remove path from.
    '
    '  OUTPUTS: N/A
    '
    '  RETURNS: This function returns all the characters from right to the
    '           first \.  Does NOT check validity of the filename....
    '
    '*******************************************************************
    Dim nIdx As Integer
    Dim lF As Long
    lF = Len(sFilename)

    For nIdx = lF To 1 Step -1

        If Mid$(sFilename, nIdx, 1) = "\" Then
            ExtractFileName = Mid$(sFilename, nIdx + 1)
            Exit Function
        End If

    Next

    ExtractFileName = sFilename

End Function

Function ExtractPath(sFilename As String) As String

    '*******************************************************************
    '
    '  PURPOSE: This returns just a path name from a full/partial path.
    '
    '  INPUTS:  sFileName - String Data to remove file from.
    '
    '  OUTPUTS: N/A
    '
    '  RETURNS: This function returns all the characters from left to the last
    '           first \.  Does NOT check validity of the filename/Path....
    '*******************************************************************
    Dim nIdx As Integer
    Dim lF As Long
    lF = Len(sFilename)

    For nIdx = lF To 1 Step -1

        If Mid$(sFilename, nIdx, 1) = "\" Then
            ExtractPath = Mid$(sFilename, 1, nIdx)
            Exit Function
        End If

    Next

    ExtractPath = sFilename

End Function



Public Sub xMkdir(sPath As String)

    Dim sPathPart() As String
    Dim lPartCount As Long
    Dim curMkdir As String
    Dim lfor As Long
    Dim fso As New gCFileSystem
    sPath = Replace(sPath, "/", "\")
    sPathPart = Split(sPath, "\")
    lPartCount = UBound(sPathPart) + 1

    If lPartCount <= 1 Then
        MkDir sPath
        Exit Sub
    End If

    curMkdir = sPathPart(0)
    Dim lEnd As Long
    lEnd = lPartCount - 1

    For lfor = 1 To lEnd
        curMkdir = curMkdir & "\" & sPathPart(lfor)

        If fso.PathExists(curMkdir) = False Then MkDir curMkdir
    Next

End Sub

Public Function expandStr(ByVal systemString As String) As String
Dim stmp As String
Dim sMass As String

Dim pos1 As Long
Dim pos2 As Long

expandStr = systemString
Do
    pos1 = InStr(expandStr, "%")
    If pos1 = 0 Then Exit Do
    pos2 = InStr(pos1 + 1, expandStr, "%")
    If pos2 = 0 Then Exit Do
    sMass = Mid$(expandStr, pos1 + 1, pos2 - pos1 - 1)
    sMass = Environ$(sMass)
    stmp = Left$(expandStr, pos1 - 1) & sMass & Right$(expandStr, Len(expandStr) - pos2)
    expandStr = stmp

Loop

End Function


Public Function queryPdgLib(ByVal strQuery As String) As String
    Dim Lib() As String
    Dim libCount As Long
    Dim Cata() As String
    Dim cataCount As Long
    Dim Book() As String
    Dim bookCount As String
    Const libList = "D:\Read\SSREADER39\remote112\liblist.dat"
    libCount = pdgLibList(libList, Lib())
    
    Dim i As Long
    For i = 1 To libCount
        cataCount = pdgCatalist(Lib(2, i), Cata())
        Dim j As Long
        For j = 1 To cataCount
            bookCount = pdgBookList(Cata(2, j), Book())
        Next
    Next
    
    
        
End Function

Public Function pdgLibList(ByVal rootTree As String, ByRef Lib() As String) As Long
    Dim fso As New FileSystemObject
    Dim ts As TextStream
    Dim tmp() As String
    Set ts = fso.OpenTextFile(rootTree, ForReading, False)
    Do Until ts.AtEndOfStream
        tmp = Split(ts.ReadLine, "|")
        If UBound(tmp) > 1 Then
            pdgLibList = pdgLibList + 1
            ReDim Preserve Lib(1 To 2, 1 To pdgLibList) As String
            Lib(1, pdgLibList) = tmp(0)
            Lib(2, pdgLibList) = fso.BuildPath(fso.GetParentFolderName(rootTree), tmp(1)) & "bktree.dat"
            Debug.Print Lib(1, pdgLibList) & " - " & Lib(2, pdgLibList)
        End If
    Loop
    ts.Close
    Set fso = Nothing
End Function

Public Function pdgCatalist(ByVal cataTree As String, ByRef Cata() As String) As Long
    Dim fso As New FileSystemObject
    Dim ts As TextStream
    Dim tmp() As String
    If fso.FileExists(cataTree) = False Then Exit Function
    Set ts = fso.OpenTextFile(cataTree, ForReading, False)
    Do Until ts.AtEndOfStream
        tmp = Split(ts.ReadLine, "|")
        If UBound(tmp) > 1 Then
            pdgCatalist = pdgCatalist + 1
            ReDim Preserve Cata(1 To 2, 1 To pdgCatalist) As String
            Cata(1, pdgCatalist) = tmp(0)
            Cata(2, pdgCatalist) = fso.BuildPath(fso.GetParentFolderName(cataTree), tmp(2))
            Debug.Print Cata(1, pdgCatalist) & " - " & Cata(2, pdgCatalist)
        End If
    Loop
    ts.Close
    Set fso = Nothing
End Function


Public Function pdgBookList(ByVal bookTree As String, ByRef Book() As String) As Long
    Dim fso As New FileSystemObject
    Dim ts As TextStream
    Dim tmp() As String
    If fso.FileExists(bookTree) = False Then Exit Function
    Set ts = fso.OpenTextFile(bookTree, ForReading, False)
    Do Until ts.AtEndOfStream
        tmp = Split(ts.ReadLine, "|")
        If UBound(tmp) > 1 Then
            pdgBookList = pdgBookList + 1
            ReDim Preserve Book(1 To 2, 1 To pdgBookList) As String
            Book(1, pdgBookList) = tmp(0)
            Book(2, pdgBookList) = tmp(3)
            Debug.Print Book(1, pdgBookList) & " - " & Book(2, pdgBookList)
        End If
    Loop
    ts.Close
    Set fso = Nothing
End Function
Public Sub splitFileByText(ByRef sFilename As String, ByRef sTextSign As String)
Dim fso As FileSystemObject
Dim srcTS As TextStream
Dim dstTS As TextStream
Dim iCount As Long
Dim sFolder As String
Dim sName As String
Dim sExt As String
Dim sLine As String
Dim sDstFile As String
Dim iLineCount As Long

Set fso = New FileSystemObject
Debug.Print sFilename
If Not fso.FileExists(sFilename) Then Exit Sub
sName = fso.GetBaseName(sFilename)
sFolder = fso.GetParentFolderName(sFilename)
sFolder = fso.BuildPath(sFolder, sName)
sExt = fso.GetExtensionName(sFilename)
If sExt <> "" Then sExt = "." & sExt
If Not fso.FolderExists(sFolder) Then fso.CreateFolder sFolder


Set srcTS = fso.OpenTextFile(sFilename, ForReading, False)
sName = getLineNotEmpty(srcTS)
If sName = "" Then Exit Sub
sDstFile = fso.BuildPath(sFolder, cleanFilename(sName) & sExt)
Debug.Print sDstFile

Set dstTS = fso.CreateTextFile(sDstFile, True)
dstTS.WriteLine sName

Do Until srcTS.AtEndOfStream
    sLine = srcTS.ReadLine
    If InStr(sLine, sTextSign) And iLineCount > 30 Then
        iLineCount = 0
        dstTS.WriteLine sLine
        dstTS.Close
        sName = getLineNotEmpty(srcTS)
        If sName = "" Then Exit Sub
        sDstFile = fso.BuildPath(sFolder, cleanFilename(sName) & sExt)
        Set dstTS = fso.CreateTextFile(sDstFile, True)
        dstTS.WriteLine sName
        Debug.Print sDstFile
    Else
        If InStr(sLine, sTextSign) Then
            iLineCount = 0
        Else
            iLineCount = iLineCount + 1
        End If
        dstTS.WriteLine sLine
    End If
Loop

dstTS.Close
srcTS.Close
Set dstTS = Nothing
Set srcTS = Nothing
Set fso = Nothing

If sLine = "" Then sLine = LTrim(RTrim(srcTS.ReadLine))

End Sub
Public Function getLineNotEmpty(ByRef ts As TextStream) As String
    If ts Is Nothing Then Exit Function
    Do Until ts.AtEndOfStream
        getLineNotEmpty = LTrim$(RTrim$(ts.ReadLine))
        If getLineNotEmpty <> "" Then Exit Function
    Loop
End Function
                                                                                                                                                                                          

