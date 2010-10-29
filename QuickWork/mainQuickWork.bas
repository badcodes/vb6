Attribute VB_Name = "mainQuickWork"

Public Sub moveByFilename()

    Dim fso As New FileSystemObject
    Dim fsoFiles As Files
    Dim fsoF As File
    Dim sFilename As String
    Dim sDirMoveto As String

    Set fsoFiles = fso.GetFolder("E:\Document\Literature\English Original").Files

    For Each fsoF In fsoFiles
        sFilename = fsoF.Name
        sDirMoveto = LeftRange(fsoF.Name, "[", "]")

        If sDirMoveto = "" Then GoTo forContune
        sFilename = Replace(sFilename, "[" & sDirMoveto & "] ", "")
        sDirMoveto = fso.BuildPath("E:\Document\Literature\English Original", sDirMoveto)

        If fso.FolderExists(sDirMoveto) = False Then fso.CreateFolder (sDirMoveto)
        sFilename = fso.BuildPath(sDirMoveto, sFilename)
        fso.MoveFile fsoF.Path, sFilename

forContune:
    Next

End Sub
Public Sub RestoreIt()

    Dim fso As New FileSystemObject
    Dim fsoFiles As Files
    Dim fsoF As File
    Dim sFilename As String
    Dim sDirMoveto As String
    Dim fsofolders As folders
    Dim fsoFolder As Folder

    Set fsofolders = fso.GetFolder("E:\Document\Literature\English Original").subFolders

    For Each fsoFolder In fsofolders

        If fsoFolder.Files.count < 2 Then

            For Each fsoF In fsoFolder.Files
                sFilename = "[" & fsoFolder.Name & "] " & fsoF.Name
                sFilename = fso.BuildPath(fsoFolder.parentFolder.Path, sFilename)
                fso.MoveFile fsoF.Path, sFilename

            Next

            fso.DeleteFolder fsoFolder

        End If

    Next

End Sub
Public Sub MoveLonelyDir(ByRef mainFolder As String)

    Dim fso As New FileSystemObject
    Dim fsoFiles As Files
    Dim fsoF As File
    Dim sDirMoveto As String
    Dim fsofolders As folders
    Dim fsoFolder As Folder

    Set fsofolders = fso.GetFolder(mainFolder).subFolders

    For Each fsoFolder In fsofolders

        If fsoFolder.Files.count < 2 And fsoFolder.subFolders.count = 0 Then

            For Each fsoF In fsoFolder.Files
                 sDirMoveto = fsoFolder.parentFolder & "\" & fsoF.Name
                 If fso.FileExists(sDirMoveto) Then
                    fso.DeleteFile fsoF.Path
                 Else
                    fso.MoveFile fsoF.Path, sDirMoveto
                 End If
            Next

            fso.DeleteFolder fsoFolder, True

        End If

    Next

End Sub


Public Sub CreateIndex(sPath As String)

    Dim fso As New FileSystemObject
    Dim fsofile As File
    Dim ts As TextStream
    Dim fTmp As String
    Dim fReal As String
    Dim tsContent As String
    fReal = fso.BuildPath(sPath, "index.txt")

    If fso.FileExists(fReal) Then fso.DeleteFile fReal, True
    fTmp = fso.BuildPath(fso.GetSpecialFolder(TemporaryFolder), fso.GetTempName)
    Set ts = fso.OpenTextFile(fTmp, ForWriting, True)
    tsContent = "<table width=_100%_ border=0 >"
    tsContent = tsContent & "<tr><td align=_center_>"
    tsContent = tsContent & "<table><tr><td style=_line-height: 150%_>"

    For Each fsofile In fso.GetFolder(sPath).Files
        tsContent = tsContent & "&gt;&gt;&nbsp;<a href=_" & fsofile.Name & "_>" & fso.GetBaseName(fsofile.Name) & "</a>" & vbCrLf
    Next

    tsContent = tsContent & "</td></tr></table></td></tr></table>"
    tsContent = Replace(tsContent, "_", Chr(34))
    ts.Write tsContent
    ts.Close
    fso.MoveFile fTmp, fReal

End Sub

Public Sub replaceFirstLine(sPath As String)

    Dim fso As New FileSystemObject
    Dim fsofile As File
    Dim tsR As TextStream
    Dim tsW As TextStream
    Dim fTmp As String
    Dim fReal As String
    Dim stmp As String
    Dim arrFile() As String
    Dim fCount As Long
    Dim l As Long
    
    For Each fsofile In fso.GetFolder(sPath).Files
        ReDim Preserve arrFile(fCount) As String
        arrFile(fCount) = fsofile.Path
        fCount = fCount + 1
    Next
    
    fTmp = fso.GetTempName

    For l = 0 To fCount - 1
        fReal = arrFile(l)
        Set tsW = fso.CreateTextFile(fTmp, True)
        Set tsR = fso.OpenTextFile(fReal, ForReading)

        If tsR.AtEndOfStream = False Then
            stmp = Replace(tsR.ReadLine, "章", "章    ")
            stmp = RTrim(stmp)
            tsW.WriteLine stmp
        End If

        Do Until tsR.AtEndOfStream
            tsW.WriteLine tsR.ReadLine
        Loop

        tsW.Close
        tsR.Close
        fso.DeleteFile fReal
        fso.MoveFile fTmp, fReal
    
    Next

End Sub

Public Sub unix2dosMain()

    Dim sPath As String
    Dim a As VbMsgBoxResult
    sPath = CurDir$
    a = MsgBox(bddir(sPath) & "..." & vbCrLf & vbCrLf & "Will Convert All Files To Dos Text Format. " & vbCrLf _
       & "Contune ? ", vbYesNo, "Unix2Dos")

    If a = vbNo Then Exit Sub
    unix2dos sPath

End Sub

Public Sub unix2dos(sPath As String)

    Dim fTmp As String
    Dim fReal As String
    Dim arrFile() As String
    Dim fCount As Long
    Dim l As Long
    
    treeSearchFiles sPath, "*.*", arrFile(), fCount

    fTmp = "kfldjsaifoejaklfds.tejpfdsfd"

    For l = 0 To fCount - 1

        fReal = arrFile(l)
        Dim f%, ff%, t As String * 2048, ct&, x%, nxt%

        'On Error GoTo File_Err
        f% = FreeFile
        Open fReal For Binary As #f%
        ff% = FreeFile
        Open fTmp For Output As #ff%
    
        Do While Not EOF(f%)
            Get #f%, , t$
            x% = 1

            Do While x% < Len(t$)
                ct& = ct& + 1
                x% = InStr(x%, t$, Chr$(10))

                If x% = 0 Then Exit Do
                Print #ff%, Left$(t$, x% - 1)
                t$ = Right$(t$, Len(t$) - x%)
            Loop

        Loop

        Close #ff%
        Close #f%
        Kill fReal
        FileCopy fTmp, fReal
        Kill fTmp
    Next

    MsgBox "Done!"
Exit_Sub:
    Exit Sub

File_Err:
    Resume Exit_Sub

End Sub

Public Sub MoveEmptyDir(pFolder As String)

    Dim fso As New FileSystemObject
    Dim f As Folder
    Dim ff As Folder
    Dim fs As folders
    Dim ffs As folders
    Dim src As String
    Dim dst As String
    Dim ad As String
    Dim adf As String
    Dim dstFolder As String
    dstFolder = "E:\download\" & fso.GetBaseName(pFolder) & "\"

    If fso.FolderExists(dstFolder) = False Then fso.CreateFolder dstFolder

    Set fs = fso.GetFolder(pFolder).subFolders

    For Each f In fs
        src = f.Path & "\" & 0
        ad = src & "\ad"
        adf = src & "\ad.htm"
        dst = dstFolder & f.Name

        If fso.FolderExists(src) Then

            If fso.FolderExists(ad) Then fso.DeleteFolder ad

            If fso.FileExists(adf) Then fso.DeleteFile adf
            ' Debug.Print "Delete " & ad
            ' Debug.Print "Detete " & adf
            Debug.Print "Move " & src; " -> " & dst
            fso.MoveFolder src, dst
        End If

    Next

End Sub
Public Sub renameHtmlFileByTitle(pFolder As String, Optional tFolder As String = "")
    Dim fso As New FileSystemObject
    Dim fs As Files
    Dim f As File
    Dim Ext As String
    Dim src As String
    Dim dst As String
    Dim fCount As Long
    
    If fso.FolderExists(pFolder) = False Then Exit Sub
    If tFolder = "" Then tFolder = pFolder
    If fso.FolderExists(tFolder) = False Then fso.CreateFolder (tFolder)
    If fso.FolderExists(tFolder) = False Then Exit Sub
    
    
    On Error Resume Next
    Set fs = fso.GetFolder(pFolder).Files
    For Each f In fs
        Ext = LCase(fso.GetExtensionName(f.Name))
        If Ext = "htm" Or Ext = "html" Then
            src = f.Path
            dst = getHtmlTitle(src, 300)
            dst = linvblib.cleanFilename(dst)
            If dst <> "" Then
                dst = fso.BuildPath(tFolder, dst & "." & Ext)
                fCount = fCount + 1
                Debug.Print "[" & fCount & "]" & "src:" & src
                Debug.Print "[" & fCount & "]" & "dst:" & dst
                fso.MoveFile src, dst
            End If
        End If
    Next
    Set f = Nothing
    Set fs = Nothing
    Set fso = Nothing
End Sub
Public Sub renameFolderByIndexHtmlTitle(pFolder As String, Optional tFolder As String = "")

    Dim fso As New FileSystemObject
    Dim fd As Folder
    Dim fds As folders
    Dim f As File
    Dim fs As Files
    
    Dim HtmlFile As String
    
    Dim l As Long
    Dim lc As Long
    Dim src As String
    Dim dst As String
    
    'Delete OE file
    '=========================================================
    Dim oeFile As String
    oeFile = fso.BuildPath(pFolder, "Descr.WD3")
    If fso.FileExists(oeFile) Then fso.DeleteFile oeFile, True
    '=========================================================
    
    
    Set fds = fso.GetFolder(pFolder).subFolders
    For Each fd In fds
        Call renameFolderByIndexHtmlTitle(fd.Path, tFolder)
    Next
    
    If tFolder = "" Then tFolder = fso.GetParentFolderName(pFolder)
    
    HtmlFile = fso.BuildPath(pFolder, "index.htm")
    
    If Not fso.FileExists(HtmlFile) Then
        HtmlFile = fso.BuildPath(pFolder, "index.html")
    End If
    
    If Not fso.FileExists(HtmlFile) Then
        Set fs = fso.GetFolder(pFolder).Files
        For Each f In fs
            Dim Ext As String
            Ext = fso.GetExtensionName(f.Path)
            If Ext = "htm" Or Ext = "html" Then
                HtmlFile = f.Path
                Exit For
            End If
        Next
    End If
    
    If Not fso.FileExists(HtmlFile) Then Exit Sub
    
    dst = getHtmlTitle(HtmlFile, 300)
    dst = linvblib.cleanFilename(dst)
    '文心阁
    dst = Replace$(dst, "--文心阁制作", "")
    If dst <> "" Then
        src = pFolder
        dst = fso.BuildPath(tFolder, dst)
        Debug.Print "Move " & Right$(src, 40)
        Debug.Print " ->  " & Right$(dst, 40)
        DoEvents
        Name src As dst
    End If

    Set fso = Nothing
    
End Sub

Public Sub testHy()

    Dim a As New clsHzToPY

    Debug.Print a.HzToPyASC(Asc("fdsf"))

End Sub

Public Sub orgClass()

    Dim classPath As String
    classPath = "E:\WorkBench\VB\[Class]"

    Dim fso As New FileSystemObject
    Dim fsofile As File
    Dim tsR As TextStream
    Dim tsW As TextStream
    Dim fTmp As String
    Dim fReal As String
    Dim stmp As String
    Dim arrFile() As String
    Dim fCount As Long
    Dim l As Long

    For Each fsofile In fso.GetFolder(classPath).Files
    
        ReDim Preserve arrFile(fCount) As String
        arrFile(fCount) = fsofile.Path
        fCount = fCount + 1
    Next
    
    fTmp = fso.GetTempName

    For l = 0 To fCount - 1
    
        fReal = fso.GetParentFolderName(arrFile(l)) & "\" & UpperChar(fso.GetFileName(arrFile(l)), 2)
        Set tsW = fso.CreateTextFile(fTmp, True)
        Set tsR = fso.OpenTextFile(fReal, ForReading)
    
        Do Until tsR.AtEndOfStream
            stmp = tsR.ReadLine

            If InStr(stmp, "Attribute VB_Name") > 0 Then
                Debug.Print stmp & "->";
                stmp = "Attribute VB_Name = " & Chr$(34) & fso.GetBaseName(fReal) & Chr$(34)
                Debug.Print stmp
                Exit Do
            End If

            tsW.WriteLine stmp
        Loop
    
        tsW.WriteLine stmp

        Do Until tsR.AtEndOfStream
            tsW.WriteLine tsR.ReadLine
        Loop
    
        tsW.Close
        tsR.Close
        fso.DeleteFile fReal
        fso.MoveFile fTmp, fReal
    
    Next

End Sub

Public Function UpperChar(strComing As String, pos As Long) As String

    UpperChar = strComing
    Mid$(UpperChar, pos, 1) = UCase$(Mid$(strComing, pos, 1))

End Function

Public Function HhcToZhc(spathName As String) As String

    h = "C:\1.HHC"

    Dim HHCText As String
    Dim hdoc As New HTMLDocument
    Dim ThisChild As Object
    Dim sAll() As String
    Dim lCount As Integer
    Dim fNum As Integer

    fNum = FreeFile
    Open spathName For Binary Access Read As #fNum
    HHCText = String(LOF(fNum), " ")
    Get fNum, , HHCText
    Close #fNum

    hdoc.body.innerHTML = HHCText

    ReDim sAll(1, 0) As String

    For Each ThisChild In hdoc.body.childNodes

        If ThisChild.nodeName = "UL" Then getLI ThisChild, sAll, lCount, ""
    Next

    'For i = 1 To LCount
    'Debug.Print sAll(0, i) & "=" & sAll(1, i)
    'Next

End Function

Private Sub getLI(ByVal ULE As HTMLUListElement, ByRef sAll() As String, ByRef iStart As Integer, ByVal sParent As String)

    Dim LI As HTMLLIElement
    Dim oChild As Object
    Dim p As HTMLParamElement
    Dim LIName As String
    Dim LILocal As String

    For Each LI In ULE.childNodes
        iStart = iStart + 1
        ReDim Preserve sAll(1, iStart) As String
        LIName = ""
        LILocal = ""

        For Each oChild In LI.childNodes

            Select Case oChild.nodeName
            Case "OBJECT"

                For Each p In oChild.childNodes

                    If p.Name = "Name" Then LIName = p.Value

                    If p.Name = "Local" Then LILocal = p.Value
                Next

                'If LILocal = "" Then LIName = LIName & "\"
                LIName = bddir(sParent & LIName)
                'If LILocal <> "" Then LILocal = bddir(LILocal)
                sAll(0, iStart) = LIName
                sAll(1, iStart) = LILocal
            Case "UL"
                LIName = bddir(LIName)
                getLI oChild, sAll, iStart, LIName
            End Select

        Next

    Next

End Sub

Public Function reNamePsc()

    Dim pFolder As String
    Dim tmpFolder As String
    Dim fso As New FileSystemObject
    Dim fs As Files
    Dim f As File

    Dim lunzip As New cUnzip
    Dim fs2 As Files
    Dim f2 As File
    Dim ts As TextStream
    Dim firstLine As String

    pFolder = "X:\codes\vb\Planet Source Code"
    'tmpFolder = pFolder & "\temp"
    'MkDir pFolder & "ReNamed"

    With lunzip
        .CaseSensitiveFileNames = False
        .UseFolderNames = False
        .OverwriteExisting = True
        .FileToProcess = "@PSC_ReadMe*.txt"
        .UnzipFolder = tmpFolder
    End With

    Set fs = fso.GetFolder(pFolder).Files

    For Each f In fs

        If LCase$(Right$(f.Name, 3)) = "zip" Then
            On Error Resume Next
            fso.DeleteFolder tmpFolder, True
            lunzip.ZipFile = f.Path
            lunzip.unzip
            Set fs2 = fso.GetFolder(tmpFolder).Files
    
            firstLine = ""

            For Each f2 In fs2
                Set ts = Nothing
                Set ts = f2.OpenAsTextStream
                firstLine = ts.ReadLine
                ts.Close
                pos = InStr(firstLine, "Title: ")

                If pos > 0 Then firstLine = Right$(firstLine, Len(firstLine) - pos - Len("Title: ") + 1)
                firstLine = cleanFilename(firstLine)
                firstLine = firstLine & ".zip"
               
                Exit For
            Next
    
            If firstLine <> "" And firstLine <> f.Name Then
                Debug.Print f.Name, "->", firstLine
                firstLine = fso.BuildPath(f.parentFolder.Path, firstLine)
                f.Move firstLine
            End If
    
        End If

    Next

End Function

Public Sub reName_hzStrToNum(sFolder As String)
    Dim fso As New FileSystemObject
    Dim fs As Files
    Dim f As File
    Dim sORG As String
    Dim sCHN As String
    Set fs = fso.GetFolder(sFolder).Files
    For Each f In fs
        sORG = fso.GetBaseName(f.Name)
        sCHN = hzStrToNum(sORG)
        If sORG <> sCHN Then
        sORG = f.Path
        sCHN = fso.BuildPath(fso.GetParentFolderName(f.Path), sCHN & "." & fso.GetExtensionName(f.Name))
        'Debug.Print sORG & "->"
        'Debug.Print "   " & sCHN
        Name sORG As sCHN
        End If
    Next
End Sub
Public Function hzStrToNum(ByRef sHz As String) As String

    Dim charNow As String
    'Dim m As Long
    Dim l As Long
    Dim lStr As Long
    'Dim lHNStart As Long
    'Dim lHNEnd As Long
    'Dim lCurValue As Long
    'Dim lTotalValue As Long
    hzStrToNum = sHz
    lStr = Len(hzStrToNum)
    'hzStrToNum = hzstrTonum

    For l = 1 To lStr
        charNow = Mid$(hzStrToNum, l, 1)

        Select Case charNow
        Case "一"
            Mid$(hzStrToNum, l, 1) = "1"
        Case "二"
            Mid$(hzStrToNum, l, 1) = "2"
        Case "三"
            Mid$(hzStrToNum, l, 1) = "3"
        Case "四"
            Mid$(hzStrToNum, l, 1) = "4"
        Case "五"
            Mid$(hzStrToNum, l, 1) = "5"
        Case "六"
            Mid$(hzStrToNum, l, 1) = "6"
        Case "七"
            Mid$(hzStrToNum, l, 1) = "7"
        Case "八"
            Mid$(hzStrToNum, l, 1) = "8"
        Case "九"
            Mid$(hzStrToNum, l, 1) = "9"
        Case "零"
            Mid$(hzStrToNum, l, 1) = "0"
        Case Else
            If isHZDigit(charNow) Then Mid$(hzStrToNum, l, 1) = Chr$(0)
        End Select
        Next
        hzStrToNum = Replace$(hzStrToNum, Chr$(0), "")
        
        'For l = 1 To lStr
        '    charNow = Mid$(hzstrTonum, l, 1)
        '    If ihzstrTonumNum(charNow) Then lHNStart = l: Exit For
        'Next

        'For l = lnstart To lStr
        '    charNow = Mid$(hzstrTonum, l, 1)
        '    If ihzstrTonumNum(charNow) = False And ihzstrTonumDigit(charNow) = False Then lHNEnd = l: Exit For
        'Next
        '
        'If lHNStart = 0 Then Exit Function
        'If lHNEnd = 0 Then lHNEnd = lStr
        '
        'Dim bLoopingHZNum As Boolean, bLoopingHZDigit As Boolean
        'For l = lHNStart To lHNEnd
        '    charNow = Mid$(hzstrTonum, l, 1)
        '    If ihzstrTonumNum(charNow) Then
        '        bLoopingHZNum = True
        '        bLoopingHZDigit = False
        '    ElseIf ihzstrTonumDigit(charNow) Then
        '        bLoopingHZDigit = True
        '        bLoopingHZNum = False
        '    End If
        'Next

    End Function

Public Function isHzNum(ByVal sHz As String) As Boolean

    sHz = Left$(sHz, 1)

    Select Case sHz
    Case "一", "二", "三", "四", "五", "六", "七", "八", "九", "零"
        isHzNum = True
    Case Else
        isHzNum = False
    End Select

End Function

Public Function isHZDigit(ByVal sHz As String) As Boolean

    sHz = Left$(sHz, 1)

    Select Case sHz
    Case "十", "百", "千", "万", "亿"
        isHZDigit = True
    Case Else
        isHZDigit = False
    End Select

End Function

Public Sub CDListor()
Dim fso As New FileSystemObject
Dim fsoDrive As Drive
Dim sSerial As String
Dim stmp As String
Dim sListFile As String
Dim LastTime As Long
Dim sLastTime As String
Dim curTime As Long
Dim sCurTime As String
sSerial = "00000000000000000000000000000000000000000"
LastTime = Timer
Do
curTime = Timer
If curTime - LastTime > 0.8 Then
   ' Debug.Print "curTime:" & vbTab & Now
    LastTime = Timer
    Set fsoDrive = fso.GetDrive("g:")
    If fsoDrive.IsReady Then
        stmp = fsoDrive.SerialNumber
        If stmp <> sSerial Then
            sSerial = stmp
            sListFile = VBA.Chr$(34) & fso.BuildPath("E:\Document\CDList", fsoDrive.VolumeName & "(" & fsoDrive.SerialNumber & ")" & ".txt") & VBA.Chr$(34)
            Shell "cmd.exe /C Dir /ad/s/b G:\ /b >" & sListFile, vbHide
            Debug.Print "SerialNumber : " & sSerial
            Debug.Print "VolumeName : " & fsoDrive.VolumeName
            Debug.Print "FileLIst Saved : " & sListFile
        End If
    End If
    Set fsoDrive = Nothing
End If
    DoEvents
Loop

End Sub

Public Sub RenamePdg(sDir As String)
Dim fso As New FileSystemObject
Dim fd As Folder
Dim f As File
Dim SArrFile() As String
Dim i As Integer
Dim l As Integer
Dim pos As Integer
Dim oName As String, nName As String, Ext As String
Set fd = fso.GetFolder(sDir)
ReDim SArrFile(1 To fd.Files.count) As String
For Each f In fd.Files
    i = i + 1
    SArrFile(i) = f.Name
Next
For l = 1 To i
    oName = fso.GetBaseName(SArrFile(l))
    Ext = fso.GetExtensionName(SArrFile(l))
    pos = InStr(oName, "]")
    If pos > 0 Then
        nName = Right$(oName, Len(oName) - pos)
        nName = nName & Left$(oName, pos)
        oName = fso.BuildPath(sDir, oName) & "." & Ext
        nName = fso.BuildPath(sDir, nName) & "." & Ext
        Name oName As nName
    End If
Next
End Sub

Public Sub testcfsi()
        
Debug.Print MClassicIO.ReadAll("c:\1.htm")

End Sub
Public Sub CreateFolderIndex(sFolder As String, Optional sParent As String = "")

    Dim fso As New FileSystemObject
    Dim fds As folders
    Dim fd As Folder
    Dim fs As Files
    Dim f As File
    Dim ts As TextStream
    Dim i As Long
    Dim fdCount As Long
    Dim fCount As Long
    Dim subFolders() As String
    Dim subFiles() As String
    
    Set fds = fso.GetFolder(sFolder).subFolders
    Set fs = fso.GetFolder(sFolder).Files
    fdCount = fds.count
    fCount = fs.count
    
    If fdCount > 0 Then ReDim subFolders(1 To fds.count) As String
    If fCount > 0 Then ReDim subFiles(1 To fs.count) As String
    
    i = 0
    For Each fd In fds
    i = i + 1
    subFolders(i) = fd.Path
    Next
    
    i = 0
    For Each f In fs
    i = i + 1
    subFiles(i) = f.Name
    Next
    
    Set ts = fso.CreateTextFile(fso.BuildPath(sFolder, "index.htm"), True, True)
    ts.WriteLine "<html><head>"
    ts.WriteLine "<meta http-equiv='Content-Type' content='text/html;charset=utf-8'>"
    ts.WriteLine "<title>" & fso.GetBaseName(sFolder) & "</title>"
    ts.WriteLine "</head><body>"
    ts.WriteLine "<table class='listtable'>"
    If sParent <> "" Then
        ts.WriteLine "<tr><td>"
        ts.WriteLine "<img src='folder.gif'>"
        ts.WriteLine "<a href='../index.htm' alt=' " & fso.GetFileName(sParent) & "'>..</a>"
        ts.WriteLine "</td></tr>"
    End If
    For i = 1 To fdCount
        ts.WriteLine "<tr><td>"
        ts.WriteLine "<img src='folder.gif'>"
        ts.WriteLine "<a href='" & fso.GetFileName(subFolders(i)) & "/index.htm' >" & fso.GetFileName(subFolders(i)) & "</a>"
        ts.WriteLine "</td></tr>"
        CreateFolderIndex subFolders(i), sFolder
    Next
    For i = 1 To fCount
        ts.WriteLine "<tr><td>"
        ts.WriteLine "<img src='file.gif'>"
        ts.WriteLine "<a href='" & subFiles(i) & "' >" & fso.GetBaseName(subFiles(i)) & "</a>"
        ts.WriteLine "</td></tr>"
    Next
    ts.WriteLine "</table>"
    ts.WriteLine "</body></html>"
    ts.Close
    
    Set ts = Nothing
    Set fds = Nothing
    Set fd = Nothing
    Set fs = Nothing
    Set f = Nothing
    Set fso = Nothing
    
End Sub

Public Sub imagegarden()
Dim srcUrl As String
Dim iFrom As Integer
Dim iTo As Integer
Dim leftUrl As String
Dim rightUrl As String
Dim i As Integer
Dim alink As String
Dim fso As New FileSystemObject
Dim ts As TextStream
Dim stmp As String

srcUrl = InputBox("输入ImageGarden URL,例如:" & vbCrLf & " http://www.imagegarden.net/" & vbCrLf & "image.php?from=white+paper&cataid=2&albumid=2234" & vbCrLf & "&imageid=1&image=jpeg", "提示")
If srcUrl = "" Then Exit Sub
leftUrl = LeftLeft(srcUrl, "imageid=", vbTextCompare, ReturnEmptyStr)
If leftUrl = "" Then Exit Sub

leftUrl = leftUrl & "imageid="
rightUrl = "&type=jpeg"
iFrom = InputBox("From:如1", "提示")
iTo = InputBox("To:如100", "提示")
If iFrom < 1 Or iTo < 1 Then Exit Sub

stmp = fso.BuildPath(Environ$("temp"), "IGfileforOE.htm")
Set ts = fso.CreateTextFile(stmp, True)
ts.WriteLine "<body>"
For i = iFrom To iTo
alink = leftUrl & LTrim$(CStr(i)) & rightUrl
ts.WriteLine "<a href='" & alink & "'>noname</a><br>"
Next
ts.WriteLine "</body>"
ts.Close

'Dim htmdoc As New HTMLDocument
'Dim ihtm As IHTMLCommentElement2
'
'Set ihtm = htmdoc.createDocumentFromUrl(stmp)
'Do Until htmdoc.readyState = "complete"
'DoEvents
'Loop

Set ts = Nothing
Set fso = Nothing

Dim iOE As New OELib.OfflineExplorerAddUrl
iOE.AddUrl "file:///" & stmp, srcUrl, srcUrl

'ShellAndClose "explorer.exe " & stmp, vbMaximizedFocus




End Sub

Public Sub renameByIndex(fileIndex As String, Optional Ext As String = "")
Dim mainFolder As String
Dim fso As New FileSystemObject
Dim ts As TextStream
Dim tempLine As String
Dim arrWord() As String
Dim i As Long
Dim nameSrc As String
Dim nameDst As String

If fso.FileExists(fileIndex) = False Then Exit Sub
Set ts = fso.OpenTextFile(fileIndex, ForReading, False)
mainFolder = fso.GetParentFolderName(fileIndex)
ChDrive (Left$(mainFolder, 1))
ChDir (mainFolder)

Do Until ts.AtEndOfStream
tempLine = ts.ReadLine
i = splitToWord(tempLine, arrWord, 2)
If i = 2 Then
    nameSrc = arrWord(1)
    nameDst = arrWord(2)
    If Ext <> "" Then
        nameSrc = nameSrc & "." & Ext
        nameDst = nameDst & "." & Ext
    End If
    nameDst = cleanFilename(nameDst)
    If fso.FileExists(nameSrc) And Not fso.FileExists(nameDst) Then
        Name nameSrc As nameDst
    End If
End If
Loop
Set ts = Nothing
Set fso = Nothing
End Sub

Public Function splitToWord(ByRef strSource As String, ByRef arrWord() As String, Optional maxWord As Long = -1) As Long
Dim c As String
Dim i As Long
Dim l As Long
Dim inWord As Boolean
Dim word As String
Dim count As Long

Debug.Print "Split " & strSource
inWord = False
l = Len(strSource)
For i = 1 To l
    c = Mid$(strSource, i, 1)
    If isSpace(c) Then
        If inWord Then
            count = count + 1
            ReDim Preserve arrWord(1 To count) As String
            arrWord(count) = word
            Debug.Print count & ":" & word
            inWord = False
            word = ""
        End If
    Else
        inWord = True
        If maxWord <= 0 Or count < maxWord - 1 Then
            word = word & c
        ElseIf count >= maxWord - 1 Then
            word = Right$(strSource, l - i + 1)
            Exit For
        End If
    End If
Next
If inWord Then
    count = count + 1
    ReDim Preserve arrWord(1 To count) As String
    arrWord(count) = word
    Debug.Print count & ":" & word
End If
splitToWord = count
End Function

Public Function isSpace(ByRef c As String) As Boolean
Dim keyCode As Integer
keyCode = Asc(c)
If keyCode = vbKeySpace Or _
   keyCode = vbKeyTab _
    Then _
isSpace = True

End Function

Public Sub testSplitWord(ByRef strSource As String, Optional maxWord As Long = -1)
    Dim arrWord() As String
    Dim count As Long
    count = splitToWord(strSource, arrWord, maxWord)
End Sub

Public Function makeWenxinPageforOE(ByRef baseUrl As String, Optional iStart As Integer = 1, Optional iEnd As Integer = 8)
Dim i As Integer
Dim href As String
Dim hOE As New OELib.OfflineExplorerAddUrl
For i = iStart To iEnd
href = Replace(baseUrl, "$*$", LTrim(Str(i)))
hOE.AddUrl href, href, href
Next
Set hOE = Nothing
End Function

Public Sub testMakeZHM(ByRef pFolder As String)
Dim hZH As New CMakeZhComment
Dim flist() As String




End Sub


Public Sub testMDB()

   Dim db As New DBEngine
   Dim dbase As Database
   Dim tdef As TableDef
   Dim fs As Fields
   Dim f As Field
   Dim rs As Recordset
   
   Set dbase = db.OpenDatabase("c:\sslib.mdb", , False)

   Set tdef = dbase.CreateTableDef("fdsfd", dbAttachExclusive)
    
    'tdef.Updatable = True
    With tdef
        .Fields.Append .CreateField("ssid", dbLong)
        .Fields.Append .CreateField("title", dbText)
        .Fields.Append .CreateField("author", dbText)
        .Fields.Append .CreateField("pages", dbText)
        .Fields.Append .CreateField("date", dbText)
        .Fields.Append .CreateField("catalog", dbText)
        .Fields.Append .CreateField("lib", dbText)
        .Fields.Append .CreateField("link", dbText)
    End With
   
   dbase.TableDefs.Append tdef
   
   Set dbase = Nothing
   Set db = Nothing

   
End Sub

Public Sub DelelteJPGPage(ByRef TopParent As String)
    Dim folders() As String
    Dim count As Long
    count = MFileSystem.subFolders(TopParent, folders())
    On Error Resume Next
    For i = LBound(folders()) To UBound(folders())
        folders(i) = BuildPath(folders(i))
        Kill folders(i) & "cov*.pdg"
        Kill folders(i) & "bok*.pdg"
        Kill folders(i) & "leg*.pdg"
        Kill folders(i) & "bac*.pdg"
        Kill folders(i) & "att*.pdg"
    Next
End Sub


