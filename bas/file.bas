Attribute VB_Name = "fileasist"
Public Const CryptKey1 = 29 '|(Asc ("L") + Asc("I") + Asc("N")-256|
Public Const CryptKey2 = 49  '|Asc("X") + Asc("I") + Asc("A") + Asc("O") - 256|
Public Const CryptKey3 = 31 '|Asc ("R") + Asc("A") + Asc("N")-256|
Public Const CryptFlag = "LCF" 'Lin Crypt File
Public CryptProgress As Integer
Public Const ftIE = 1
Public Const ftExE = 2
Public Const ftCHM = 3
Public Const ftIMG = 4
Public Const ftAUDIO = 5
Public Const ftVIDEO = 6
Public Const ftHTML = 7
Public Const ftZIP = 8
Public Const MAXTEXTBLOCK = 10240
Public Const cMaxPath = 260
Public Declare Function GetFullPathName Lib "kernel32" Alias "GetFullPathNameA" (ByVal lpFileName As String, ByVal nBufferLength As Long, ByVal lpBuffer As String, ByVal lpFilePart As String) As Long





Sub rebuildfile(filepath As String, skipline As Integer)

Dim fpath As String
Dim fso As New FileSystemObject
Dim filelist() As String
Dim FileCount As Integer

fpath = filepath
Set fs = fso.GetFolder(fpath).Files

FileCount = fs.Count
If FileCount < 1 Then Exit Sub
ReDim filelist(FileCount) As String

For Each ff In fs
i = i + 1
filelist(i) = ff.Name
Next

fpath = bddir(fpath)

Dim srcTS As TextStream
Dim dstTS As TextStream
Dim norb As Boolean
Dim tmpstr As String

For i = 1 To FileCount

norb = False
Set srcTS = fso.OpenTextFile(fpath + filelist(i), ForReading)
    
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
    tmpstr = LTrim(tmpstr)
    tmpstr = RTrim(tmpstr)
    If Right(tmpstr, 1) = Chr(13) Then tmpstr = Left(tmpstr, Len(tmpstr) - 1)
    tmpstr = StrConv(tmpstr, vbWide)
    Set dstTS = fso.CreateTextFile(fpath + tmpstr + ".txt", True)
    dstTS.WriteLine tmpstr
    Do Until srcTS.AtEndOfStream
    tmpstr = srcTS.ReadLine
    dstTS.WriteLine tmpstr
    Loop
    dstTS.Close
    srcTS.Close
    fso.DeleteFile fpath + filelist(i), True
End If


Next


End Sub

Public Function bddir(dirname As String) As String
bddir = dirname
If Right(bddir, 1) <> "\" Then bddir = bddir + "\"
End Function

Sub insertline(srcfile As String, txtLINE As String, insTYPE As Integer)
Dim fso As FileSystemObject
Dim srcf As File, dstf As File
Dim srcfs As TextStream, dstfs As TextStream
Dim tmpstr As String
Set fso = New FileSystemObject
Set srcf = fso.GetFile(srcfile)
fso.CreateTextFile "$1$2$3.$$$", True
Set dstf = fso.GetFile("$1$2$3.$$$")
Set srcfs = srcf.OpenAsTextStream(ForReading)
Set dstfs = dstf.OpenAsTextStream(ForWriting)
Select Case insTYPE
Case 1
    Do Until srcfs.AtEndOfStream
    tmpstr = srcfs.ReadLine
    dstfs.WriteLine (txtLINE + tmpstr)
    Loop
Case Else
    Do Until srcfs.AtEndOfStream
    tmpstr = srcfs.ReadLine
    dstfs.WriteLine (tmpstr + txtLINE)
    Loop
End Select
srcfs.Close
dstfs.Close
fso.DeleteFile srcfile, True
fso.MoveFile "$1$2$3.$$$", srcfile

End Sub

Public Function MyFileCrypt(srcfile As String, dstfile As String) As Boolean


Dim tmpFile As String
Dim thebyte As Byte
MyFileCrypt = True
If Dir(srcfile) = "" Then MyFileCrypt = False: Exit Function

tmpFile = "~$$$CRfile.tmp"
If Dir(tmpFile) <> "" Then Kill tmpFile

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
If Dir(dstfile) <> "" Then Kill dstfile
FileCopy tmpFile, dstfile
Kill tmpFile

End Function

Public Function MyFileENCrypt(srcfile As String, dstfile As String) As Boolean


Dim tmpFile As String
Dim thebyte As Byte
MyFileENCrypt = True
If Dir(srcfile) = "" Then MyFileENCrypt = False: Exit Function
If isLXTfile(srcfile) = False Then MyFileENCrypt = False: Exit Function
Open srcfile For Binary As #1
tmpFile = "~$$$CRfile.tmp"
If Dir(tmpFile) <> "" Then Kill tmpFile
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
If Dir(dstfile) <> "" Then Kill dstfile
FileCopy tmpFile, dstfile
Kill tmpFile

End Function

Public Function isLXTfile(thefile As String) As Boolean
Dim fso As New FileSystemObject
Dim f As File
isLXTfile = False
If fso.FileExists(thefile) = False Then Exit Function
Set f = fso.GetFile(thefile)
If f.Size < Len(CryptFlag) Then Exit Function
If f.OpenAsTextStream(ForReading).Read(Len(CryptFlag)) = CryptFlag Then isLXTfile = True
End Function

Public Function testfso(thefile As String)
Dim fso As New FileSystemObject
Dim ff As File
Dim ts As TextStream

Set ff = fso.GetFile(thefile)
Set ts = ff.OpenAsTextStream
Debug.Print ts.Read(ff.Size)


End Function


Public Function chkFileType(chkfile As String) As Integer
Dim fso As New FileSystemObject
Dim ext As String
ext = LCase(fso.GetExtensionName(chkfile))
Select Case ext

    Case "jpg"
        chkFileType = ftIMG
        Exit Function
    Case "jpeg"
        chkFileType = ftIMG
        Exit Function
    Case "gif"
        chkFileType = ftIMG
        Exit Function
    Case "bmp"
        chkFileType = ftIMG
        Exit Function
    Case "png"
        chkFileType = ftIMG
        Exit Function
    Case "ico"
        chkFileType = ftIMG
    Case "swf"
        chkFileType = ftIE
        Exit Function
    Case "htm"
        chkFileType = ftHTML
        Exit Function
    Case "html"
        chkFileType = ftHTML
        Exit Function
    Case "shtml"
        chkFileType = ftHTML
        Exit Function
    Case "mht"
        chkFileType = ftIE
        Exit Function
    Case "pdf"
        chkFileType = ftIE
        Exit Function
    Case "exe"
        chkFileType = ftExE
        Exit Function
    Case "com"
        chkFileType = ftExE
        Exit Function
    Case "bat"
        chkFileType = ftExE
        Exit Function
    Case "cmd"
        chkFileType = ftExE
        Exit Function
    Case "chm"
        chkFileType = ftCHM
        Exit Function
    Case "mp3"
        chkFileType = ftAUDIO
        Exit Function
    Case "wav"
        chkFileType = ftAUDIO
        Exit Function
    Case "wma"
        chkFileType = ftAUDIO
        Exit Function
    Case "rm"
        chkFileType = ftVIDEO
        Exit Function
    Case "rmvb"
        chkFileType = ftVIDEO
        Exit Function
    Case "avi"
        chkFileType = ftVIDEO
        Exit Function
    Case "mpg"
        chkFileType = ftVIDEO
        Exit Function
    Case "mpeg"
        chkFileType = ftVIDEO
        Exit Function
    Case "zip"
        chkFileType = ftZIP
        Exit Function
End Select


End Function

Function appenddata(thefile As String, thedata As String)
n = FreeFile()
Open thefile For Binary As #n

endpos = LOF(n)
Seek #n, endpos + 1
Put #n, , thedata
Put #n, , endpos + 1


Close #n
End Function

Function readappenddata(thefile As String) As String
n = FreeFile()
Open thefile For Binary As #n
Seek #n, LOF(n) - 3
Dim a As Long
Get #n, , a
If a <= 0 Then
Close #n
Exit Function
End If

Seek #n, a
Do Until Loc(n) = LOF(n) - 6
readappenddata = readappenddata + Input(1, #n)
Loop
Close #n
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
        tempstr = Chr(thebyte)
        Seek 1, Loc(1) - 1
        End If
opsget = tempstr

End Function

Function lenchar(thechar As String) As String
If Asc(thechar) < 0 Then
lenchar = 2
Else
lenchar = 1
End If
End Function

Function delformat(thefile As String)
Dim fso As New FileSystemObject
Dim ts As TextStream
Dim tempts As TextStream
Dim tempstr As String
TempFile = fso.GetTempName
Set tempts = fso.OpenTextFile(TempFile, ForWriting, True)
Set ts = fso.OpenTextFile(thefile, ForReading)
Do Until ts.AtEndOfStream
tempstr = ts.ReadLine
If Trim(tempstr) = "" Then
    tempts.Write Chr(13) + Chr(10)
Else
    tempts.Write tempstr
End If
Loop
ts.Close
tempts.Close
fso.DeleteFile thefile
fso.MoveFile TempFile, thefile
End Function

Function splitfile(thefile As String, SplitFlag As String)
Dim fso As New FileSystemObject
Dim ts As TextStream
Dim tempts As TextStream
Dim tempstr As String
Dim thefolder As String
Dim SplitTS As TextStream
Dim n As Integer
If fso.FolderExists(bddir(fso.GetParentFolderName(thefile)) + fso.GetBaseName(thefile)) = False Then
fso.CreateFolder bddir(fso.GetParentFolderName(thefile)) + fso.GetBaseName(thefile)
End If
thefolder = bddir(fso.GetParentFolderName(thefile)) + fso.GetBaseName(thefile)
TempFile = bddir(thefolder) + SplitFlag + Strnum(0, 3) + "." + fso.GetExtensionName(thefile)
Set tempts = fso.OpenTextFile(TempFile, ForWriting, True)
Set ts = fso.OpenTextFile(thefile, ForReading)
tempstr = ts.ReadLine
Do Until ts.AtEndOfStream

    Do Until Left(LTrim(tempstr), Len(SplitFlag)) = SplitFlag Or ts.AtEndOfStream
    tempts.WriteLine tempstr
    tempstr = ts.ReadLine
    Loop
    If ts.AtEndOfStream = False Then
    n = n + 1
    tempts.WriteLine tempstr
    Set SplitTS = fso.OpenTextFile(bddir(thefolder) + SplitFlag + Strnum(n, 3) + "." + fso.GetExtensionName(thefile), ForWriting, True)
    SplitTS.WriteLine tempstr
    tempstr = ts.ReadLine
    Do Until Left(LTrim(tempstr), Len(SplitFlag)) = SplitFlag Or ts.AtEndOfStream
    SplitTS.WriteLine tempstr
    tempstr = ts.ReadLine
    Loop
    If ts.AtEndOfStream Then SplitTS.WriteLine tempstr
    SplitTS.Close
    End If
Loop


tempts.Close
ts.Close

End Function

Public Function treeSearch(ByVal Spath As String, ByVal SFileSpec As String, SFiles() As String) As Long
Static fstFiles As Long '文件数目
Dim sDir As String
Dim sSubDirs() As String '存放子目录名称
Dim fstIndex As Long
If Right(Spath, 1) <> "\" Then Spath = Spath + "\"
sDir = Dir(Spath + SFileSpec)
'获得当前目录下文件名和数目
Do While Len(sDir)
fstFiles = fstFiles + 1
ReDim Preserve SFiles(1 To fstFiles)
SFiles(fstFiles) = Spath + sDir
sDir = Dir
Loop
'获得当前目录下的子目录名称
fstIndex = 0
sDir = Dir(Spath + "*.*", 16)
Do While Len(sDir)
If Left(sDir, 1) <> "." Then 'skip.and..
'找出子目录名
If GetAttr(Spath + sDir) = vbDirectory Then
fstIndex = fstIndex + 1
'保存子目录名
ReDim Preserve sSubDirs(1 To fstIndex)
sSubDirs(fstIndex) = Spath + sDir + "\"
End If
End If
sDir = Dir
Loop
For fstIndex = 1 To fstIndex '查找每一个子目录下文件，这里利用了递归
Call treeSearch(sSubDirs(fstIndex), SFileSpec, SFiles())
Next fstIndex
treeSearch = fstFiles
End Function

Sub delline(filepath As String, skipline As Integer)

Dim fpath As String
Dim fso As New FileSystemObject
Dim filelist() As String
Dim FileCount As Integer
Dim fs As Files
fpath = filepath
Set fs = fso.GetFolder(fpath).Files



FileCount = fs.Count
If FileCount < 1 Then Exit Sub
ReDim filelist(FileCount) As String

For Each ff In fs
i = i + 1
filelist(i) = ff.Name
Next

fpath = bddir(fpath)

Dim srcTS As TextStream
Dim dstTS As TextStream
Dim norb As Boolean
Dim tmpstr As String

For i = 1 To FileCount

norb = False
Set srcTS = fso.OpenTextFile(fpath + filelist(i), ForReading)
    
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
    fso.DeleteFile fpath + filelist(i), True
    fso.MoveFile dstfile, fpath + filelist(i)
End If


Next


End Sub

Public Sub RenameBat(thedir As String, renameflag As String)

Dim fso As New FileSystemObject
Dim fs As Files
Dim f As File
Dim tmpline As String
Dim ts As TextStream
If fso.FolderExists(thedir) = False Then Exit Sub
Set fs = fso.GetFolder(thedir).Files
For Each f In fs
Set ts = f.OpenAsTextStream(ForReading)
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
Dim fso As New FileSystemObject
Dim ts As TextStream
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
srcfile = Left(tempstr, pos - 1)
dstfile = Right(tempstr, Len(tempstr) - pos + Len(SeperateFlag) - 1)
If srcfile <> dstfile And fso.FileExists(bddir(thedir) + srcfile) = True Then
srcfile = bddir(thedir) + srcfile
dstfile = bddir(thedir) + StrConv(fso.GetBaseName(dstfile), vbWide) + "." + fso.GetExtensionName(dstfile)
If fso.FileExists(dstfile) = False Then fso.MoveFile srcfile, dstfile
End If
End If

Loop



End Sub

Public Function delblankline(thefile As String, Optional dstfile As String) As Boolean
Dim fso As New FileSystemObject
If fso.FileExists(thefile) = False Then Exit Function
If dstfile = "" Then dstfile = thefile
Dim ts As TextStream
Dim tempts As TextStream
Dim TempFile As String
Dim tempstr As String
Dim realstr As String
Dim blankline As Boolean
TempFile = fso.GetTempName
Set ts = fso.OpenTextFile(thefile, ForReading)
Set tempts = fso.OpenTextFile(TempFile, ForWriting, True)
Do Until ts.AtEndOfStream
tempstr = ts.ReadLine
realstr = RTrim(LTrim(tempstr))
blankline = False
If realstr = "" Then blankline = True
If realstr = Chr(13) Then blankline = True
If realstr = Chr(10) Then blankline = True
If realstr = Chr(13) + Chr(10) Then blankline = True
If Not blankline Then tempts.WriteLine tempstr
Loop

ts.Close
tempts.Close
fso.DeleteFile thefile
fso.MoveFile TempFile, dstfile
delblankline = True
End Function

Public Function BATdelblankline(thedir As String) As Boolean
Dim fso As New FileSystemObject
If fso.FolderExists(thedir) = False Then Exit Function
Dim f As File
Dim fs As Files
Set fs = fso.GetFolder(thedir).Files
For Each f In fs
m = m + 1
Debug.Print Str(m) + "/" + Str(fs.Count) + ":" + f.Path
delblankline f.Path
Next
End Function
Public Function RenameByFstLine(thefile As String)
Dim fso As New FileSystemObject
If fso.FileExists(thefile) = False Then Exit Function

Dim ts As TextStream
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

End Function
Public Function FilestrBetween(thefile, strStart As String, strEnd As String, strResult() As String, strcount As Integer)

If strStart = "" Then Exit Function
If strEnd = "" Then Exit Function
Dim fso As New FileSystemObject
If fso.FileExists(thefile) = False Then Exit Function
Dim ts As TextStream
Dim tempstrA(10240) As String
Dim tempstr As String
Dim pos1 As Integer
Dim pos2 As Integer
Set ts = fso.OpenTextFile(thefile, ForReading)
Do Until ts.AtEndOfStream
tempstr = ts.ReadLine
pos1 = InStr(1, tempstr, strStart, vbTextCompare)
If pos1 > 0 Then
    pos2 = InStr(pos1 + Len(strStart), tempstr, strEnd, vbTextCompare)
    If pos2 > 0 Then
    strcount = strcount + 1
    ReDim Preserve strResult(strcount) As String
    strResult(strcount) = Mid(tempstr, pos1 + Len(strStart), pos2 - pos1 - Len(strStart))
    End If
End If
Loop
ts.Close

End Function

Function GetFullPath(sFileName As String, Optional FilePart As Long, Optional ExtPart As Long, Optional DirPart As Long) As String
Dim c As Long, p As Long, sRet As String
GetFullPath = sFileName
If sFileName = sEmpty Then Exit Function
' Get the path size, then create string of that size
 sRet = String(cMaxPath, 0)
 c = GetFullPathName(sFileName, cMaxPath, sRet, p)
 If c = 0 Then Exit Function
 sRet = Left$(sRet, c)
 GetFullPath = sRet
End Function

Function FileStrReplace(thefile As String, thetext As String, RPText As String)

Const MAXSTRING = 28800

If thetext = RPText Then Exit Function
If thetext = "" Then Exit Function
If Len(thetext) >= MAXSTRING \ 2 Then MsgBox ("The text to replace is too large!"): Exit Function

Dim fso As New FileSystemObject
Dim MatchNum As Integer

If fso.FileExists(thefile) = False Then Exit Function
Dim ff As File
Set ff = fso.GetFile(thefile)
If ff.Size < Len(thetext) Then Exit Function

Dim BlockSize As Long

BlockSize = MAXSTRING

If ff.Size < BlockSize Then BlockSize = ff.Size
textsize = Len(thetext)
blocknum = (ff.Size - 1) \ (BlockSize) + 1

Dim tempstring As String
Dim reststring As String
Dim srcTS As TextStream
Dim dstTS As TextStream
Dim TempFile As String
Dim iLastPos As Integer
TempFile = fso.GetTempName


Set srcTS = ff.OpenAsTextStream(ForReading)
Set dstTS = fso.CreateTextFile(TempFile, True)

  
For i = 1 To blocknum + 1

If srcTS.AtEndOfStream Then Exit For

tempstring = reststring + srcTS.Read(BlockSize)
iLastPos = InStrRev(tempstring, thetext)

If iLastPos > 0 Then
    iLastPos = Len(tempstring) - iLastPos - Len(thetext)
    tempstring = Replace(tempstring, thetext, RPText, , , vbTextCompare)
    If iLastPos > textsize Then iLastPos = textsize
    If iLastPos = 0 Then
        reststring = ""
    Else
        reststring = Right(tempstring, iLastPos)
    End If
    tempstring = Left(tempstring, Len(tempstring) - iLastPos)
    dstTS.Write tempstring
Else
    If Len(tempstring) < textsize Then
        reststring = tempstring
        tempstring = ""
    Else
        reststring = Right(tempstring, textsize)
        tempstring = Left(tempstring, Len(tempstring) - textsize)
        dstTS.Write tempstring
    End If
End If
   
Next

    dstTS.Write reststring

    dstTS.Close
    srcTS.Close

fso.DeleteFile thefile
fso.MoveFile TempFile, thefile

End Function

Function FileInStr(thefile As String, thetext As String, Min_MatchTimes As Integer, Optional CompMethod As VbCompareMethod) As Boolean

If CompMethod = 0 Then CompMethod = vbBinaryCompare

Const MAXSTRING = 32768
If thetext = "" Then Exit Function
If Min_MatchTimes = 0 Then Exit Function
If Len(thetext) >= MAXSTRING \ 2 Then MsgBox ("The text to Search is too large!"): Exit Function

Dim fso As New FileSystemObject

If fso.FileExists(thefile) = False Then Exit Function

Dim ff As File
Set ff = fso.GetFile(thefile)
If ff.Size < Len(thetext) Then Exit Function

Dim BlockSize As Long

BlockSize = MAXSTRING

If ff.Size < BlockSize Then BlockSize = ff.Size
textsize = Len(thetext)
blocknum = ff.Size \ (BlockSize) + 1

Dim tempstring As String
Dim reststring As String
Dim srcTS As TextStream
Dim MatchTimes As Integer
Dim pos As Integer

Set srcTS = ff.OpenAsTextStream(ForReading)
reststring = space(textsize)
  
  
For i = 1 To blocknum

If srcTS.AtEndOfStream Then Exit For

tempstring = reststring + srcTS.Read(BlockSize)
pos = InStr(1, tempstring, thetext, CompMethod)

Do Until pos = 0
MatchTimes = MatchTimes + 1
If MatchTimes >= Min_MatchTimes Then FileInStr = True: srcTS.Close: Exit Function
pos = InStr(pos + textsize, tempstring, thetext, CompMethod)
Loop


If LCase(Right(tempstring, textsize)) <> LCase(thetext) Then
reststring = Right(tempstring, textsize)

End If

    
Next


srcTS.Close

End Function

Function FileInStrTimes(thefile As String, thetext As String, Optional CompMethod As VbCompareMethod) As Integer

If CompMethod = 0 Then CompMethod = vbBinaryCompare

Const MAXSTRING = 32768

If thetext = "" Then Exit Function
If Len(thetext) >= MAXSTRING \ 2 Then MsgBox ("The text to Search is too large!"): Exit Function

Dim fso As New FileSystemObject

If fso.FileExists(thefile) = False Then Exit Function

Dim ff As File
Set ff = fso.GetFile(thefile)
If ff.Size < Len(thetext) Then Exit Function

Dim BlockSize As Long

BlockSize = MAXSTRING

If ff.Size < BlockSize Then BlockSize = ff.Size
textsize = Len(thetext)
blocknum = ff.Size \ (BlockSize) + 1

Dim tempstring As String
Dim reststring As String
Dim srcTS As TextStream
Dim MatchTimes As Integer
Dim pos As Integer

Set srcTS = ff.OpenAsTextStream(ForReading)
reststring = space(textsize)
    
For i = 1 To blocknum

If srcTS.AtEndOfStream Then Exit For

tempstring = reststring + srcTS.Read(BlockSize)
pos = InStr(1, tempstring, thetext, CompMethod)

Do Until pos = 0
MatchTimes = MatchTimes + 1
pos = InStr(pos + textsize, tempstring, thetext, CompMethod)
Loop

If LCase(Right(tempstring, textsize)) <> LCase(thetext) Then
reststring = Right(tempstring, textsize)
End If

Next

srcTS.Close
FileInStrTimes = MatchTimes
End Function

Sub printfolder(thedir As String)
Dim fso As New FileSystemObject
Dim ff As File
Dim ffs As Files
Dim ts As TextStream
If fso.FolderExists(thedir) = False Then Exit Sub
Set ffs = fso.GetFolder(thedir).Files
c = ffs.Count
i = 0
For Each ff In ffs
i = i + 1
Debug.Print Chr(34) + ff.Name + Chr(34) + "," + Chr(34) + "File:[" + Str(i) + " of" + Str(c) + " ]" + Chr(34) + ","
Next
End Sub
