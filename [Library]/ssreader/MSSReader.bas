Attribute VB_Name = "MSSReader"
Option Explicit

Public Type pdginfo
    vailable As Boolean
    isreadonly As Boolean
    Name As String '%t
    Author As String '%a
    totalpage As String '%p
    download As String '%u
    Publisher As String '%c
    pdate As String '%d
    SSID As String '%s
End Type

Public Type PDG

    iszip As Boolean
    isreadonly As Boolean
    infolder As String
    unzipfolder As String
    zipfile As String
    infofile As String
    info As pdginfo

End Type

Public Enum SSLIBTaskStatus
    STS_START = 1
    STS_PAUSE = 0
    STS_COMPLETE = 2
    STS_PENDING = 3
End Enum

Public Type SSLIB_BOOKINFO
    title As String
    Author As String
    SSID As String
    Publisher As String
    PublishedDate As String
    Subject As String
    Header As String
    PagesCount As String
    AddInfo As String
    URL As String
    About As String
End Type

Private SSLIB_FLAGS_INIT As Boolean

Public Const CST_SSLIB_FIELDS_LBound As Long = 1

Public Enum SSLIBFields
        SSF_Title = CST_SSLIB_FIELDS_LBound
        SSF_Author
        SSF_SSID
        SSF_PagesCount
        SSF_Publisher
        SSF_PublishDate
        SSF_StartPage
        
        SSF_Subject
        SSF_Comments
        
        SSF_URL
        SSF_SAVEDIN
        SSF_HEADER
        'SSF_FULLNAME
        
        SSF_Downloader
        SSF_DownloadDate
        'SSF_STATUS
        'SSF_FILES_DOWNLOADED
        
        SSF_FIELDS_END

End Enum
Public Const CST_SSLIB_FIELDS_UBound As Long = SSF_FIELDS_END - 1
Public Const CST_SSLIB_FIELDS_IMPORTANT_UBOUND As Long = SSF_Comments
Public Const CST_SSLIB_FIELDS_TASKS_UBOUND As Long = SSF_HEADER

Public SSLIB_FIELDS_NAME(CST_SSLIB_FIELDS_LBound To CST_SSLIB_FIELDS_UBound, 1 To 2) As String

'Public Function SSLIB_BOOKINFO_TO_CStringMap(ByRef vBookInfo As SSLIB_BOOKINFO) As CStringMap
'    Dim pMap As CStringMap
'    Set pMap = New CStringMap
'    With vBookInfo
'        pMap.Map(SSLIB_ChnFieldName(SSF_Author)) = .Author
'        pMap.Map(SSLIB_ChnFieldName(SSF_Comments)) = .About
'        pMap.Map(SSLIB_ChnFieldName(SSF_HEADER)) = .Header
'        pMap.Map(SSLIB_ChnFieldName(SSF_PagesCount)) = .PagesCount
'        pMap.Map(SSLIB_ChnFieldName(SSF_PublishDate)) = .PublishedDate
'        pMap.Map(SSLIB_ChnFieldName(SSF_Publisher)) = .Publisher
'        'pMap.map(SSLIB_ChnFieldName(SSF_SAVEDIN ))
'        pMap.Map(SSLIB_ChnFieldName(SSF_SSID)) = .SSID
'        'pMap.map(SSLIB_ChnFieldName(SSF_STATUS ))
'        pMap.Map(SSLIB_ChnFieldName(SSF_Subject)) = .Subject
'        pMap.Map(SSLIB_ChnFieldName(SSF_Title)) = .title
'        pMap.Map(SSLIB_ChnFieldName(SSF_URL)) = .URL
'
'    End With
'
'    Set SSLIB_BOOKINFO_TO_CStringMap = pMap
'End Function


Public Sub SSLIB_Init()

    If SSLIB_FLAGS_INIT Then Exit Sub
    
    SSLIB_FIELDS_NAME(SSLIBFields.SSF_Author, 1) = "Author"
    SSLIB_FIELDS_NAME(SSLIBFields.SSF_Author, 2) = "作者"
    
'    SSLIB_FIELDS_NAME(SSLIBFields.SSF_FULLNAME, 1) = "Fullname"
'    SSLIB_FIELDS_NAME(SSLIBFields.SSF_FULLNAME, 2) = "全名"
    
    SSLIB_FIELDS_NAME(SSLIBFields.SSF_HEADER, 1) = "Header"
    SSLIB_FIELDS_NAME(SSLIBFields.SSF_HEADER, 2) = "报头"
    
    SSLIB_FIELDS_NAME(SSLIBFields.SSF_PagesCount, 1) = "PagesCount"
    SSLIB_FIELDS_NAME(SSLIBFields.SSF_PagesCount, 2) = "页数"
    
    SSLIB_FIELDS_NAME(SSLIBFields.SSF_PublishDate, 1) = "PublishedData"
    SSLIB_FIELDS_NAME(SSLIBFields.SSF_PublishDate, 2) = "出版日期"
    
    SSLIB_FIELDS_NAME(SSLIBFields.SSF_Publisher, 1) = "Publisher"
    SSLIB_FIELDS_NAME(SSLIBFields.SSF_Publisher, 2) = "出版社"
    
    SSLIB_FIELDS_NAME(SSLIBFields.SSF_SAVEDIN, 1) = "SavedIn"
    SSLIB_FIELDS_NAME(SSLIBFields.SSF_SAVEDIN, 2) = "保存位置"
    
    SSLIB_FIELDS_NAME(SSLIBFields.SSF_SSID, 1) = "SSID"
    SSLIB_FIELDS_NAME(SSLIBFields.SSF_SSID, 2) = "SS号"
    
'    SSLIB_FIELDS_NAME(SSLIBFields.SSF_STATUS, 1) = "Status"
'    SSLIB_FIELDS_NAME(SSLIBFields.SSF_STATUS, 2) = "状态"
    
    SSLIB_FIELDS_NAME(SSLIBFields.SSF_Subject, 1) = "Subject"
    SSLIB_FIELDS_NAME(SSLIBFields.SSF_Subject, 2) = "主题词"
    
    SSLIB_FIELDS_NAME(SSLIBFields.SSF_Title, 1) = "Title"
    SSLIB_FIELDS_NAME(SSLIBFields.SSF_Title, 2) = "书名"
    
    SSLIB_FIELDS_NAME(SSLIBFields.SSF_URL, 1) = "Url"
    SSLIB_FIELDS_NAME(SSLIBFields.SSF_URL, 2) = "下载位置"

    SSLIB_FIELDS_NAME(SSLIBFields.SSF_Comments, 1) = "Comments"
    SSLIB_FIELDS_NAME(SSLIBFields.SSF_Comments, 2) = "简介"
    
'    SSLIB_FIELDS_NAME(SSLIBFields.SSF_FILES_DOWNLOADED, 1) = "FilesDownloaded"
'    SSLIB_FIELDS_NAME(SSLIBFields.SSF_FILES_DOWNLOADED, 2) = "已下载页数"
    
    SSLIB_FIELDS_NAME(SSLIBFields.SSF_StartPage, 1) = "StartPage"
    SSLIB_FIELDS_NAME(SSLIBFields.SSF_StartPage, 2) = "开始页"
    
    SSLIB_FIELDS_NAME(SSLIBFields.SSF_Downloader, 1) = "Downloader"
    SSLIB_FIELDS_NAME(SSLIBFields.SSF_Downloader, 2) = "下载工具"
    
    SSLIB_FIELDS_NAME(SSLIBFields.SSF_DownloadDate, 1) = "DownloadDate"
    SSLIB_FIELDS_NAME(SSLIBFields.SSF_DownloadDate, 2) = "下载日期"
    
        
    SSLIB_FLAGS_INIT = True
    
End Sub

Public Function SSLIB_EngFieldName(ByRef vField As SSLIBFields) As String
    On Error Resume Next
    
    SSLIB_EngFieldName = SSLIB_FIELDS_NAME(vField, 1)
End Function

Public Function SSLIB_ChnFieldName(ByRef vField As SSLIBFields) As String
    On Error Resume Next
    
    SSLIB_ChnFieldName = SSLIB_FIELDS_NAME(vField, 2)
End Function

'4. 《世界儿童文学名著 小妇人》
'电信阅读  |  网通阅读
'作者:[美]奥尔科特（Alcott，L.M.）著 宋丽军 宋颖军译
'页数:330   出版日期:2001年08月第3版
'主题词:长篇小说 美国 近代 CT S007141 长篇小说
 

'Public Function SSLIB_ParseInfoText(Optional ByRef vText As String) As SSLIB_BOOKINFO
'    On Error Resume Next
'    If vText = "" Then vText = Clipboard.GetText
'    If vText = "" Then Exit Function
'    vText = vText & vbCrLf
'    If InStr(1, vText, vbCrLf & "SSCT:", vbTextCompare) > 0 Then
'        SSLIB_ParseInfoText.Header = Replace$(vText, "(Request-Line):", "")
'    Else
'        With SSLIB_ParseInfoText
'            .SSID = SubStringBetween(vText, "SS号:", vbCrLf, True)
'            .Author = SubStringBetween(vText, "作者:", vbCrLf, True)
'            If Right$(.Author, 1) = "著" Then .Author = Left$(.Author, Len(.Author) - 1)
'            .Title = SubStringBetween(vText, "《", "》", True)
'            .PagesCount = SubStringBetween(vText, "页数:", " ", True)
'            .PublishedDate = SubStringBetween(vText, "出版日期:", vbCrLf, True)
'            .Subject = SubStringBetween(vText, "主题词:", vbCrLf, True)
'            .About = SubStringBetween(vText, "简介:", vbCrLf, True)
'            '.AddInfo = "简介=" & .About & vbCrLf & "主题词=" & .Subject
'        End With
'    End If
'End Function

Public Function SSLIB_CreateBookInfoArray() As String()
    Dim Result(CST_SSLIB_FIELDS_LBound To CST_SSLIB_FIELDS_UBound) As String
    SSLIB_CreateBookInfoArray = Result
End Function

Public Function SSLIB_GetField(ByRef vArray() As String, ByVal vField As SSLIBFields) As String
    On Error Resume Next
    SSLIB_GetField = vArray(vField)
End Function

Public Sub SSLIB_SetField(ByRef vArray() As String, ByVal vField As SSLIBFields, ByVal vValue As String)
    On Error Resume Next
    vArray(vField) = vValue
End Sub

Public Function SSLIB_ParseInfoText(Optional ByRef vText As String) As String()
    On Error Resume Next
    Dim Result() As String
    Result = SSLIB_CreateBookInfoArray()
    Dim TmpStr As String
    If vText = "" Then vText = Clipboard.GetText
    If vText = "" Then Exit Function
    vText = vText & vbCrLf
    If InStr(1, vText, vbCrLf & "SSCT:", vbTextCompare) > 0 Then
        vText = Replace$(vText, "(Request-Line):", "")
        Result(SSF_HEADER) = vText
        Result(SSF_URL) = SubStringBetween(vText, "Host:", vbCrLf)
        If Result(SSF_URL) <> "" Then Result(SSF_URL) = "http://" & Result(SSF_URL) & "/"
        vText = SubStringBetween(vText, "Get /", " ")
        If vText <> "" Then Result(SSF_URL) = Result(SSF_URL) & vText
        
        
    Else
        Result(SSF_SSID) = SubStringBetween(vText, "SS号:", vbCrLf, True)
        TmpStr = SubStringBetween(vText, "作者:", vbCrLf, True)
        If Right$(TmpStr, 1) = "著" Then TmpStr = Left$(TmpStr, Len(TmpStr) - 1)
        Result(SSF_Author) = TmpStr
        Result(SSF_Title) = SubStringBetween(vText, "《", "》", True)
        Result(SSF_PagesCount) = SubStringBetween(vText, "页数:", " ", True)
        Result(SSF_PublishDate) = SubStringBetween(vText, "出版日期:", vbCrLf, True)
        Result(SSF_Subject) = SubStringBetween(vText, "主题词:", vbCrLf, True)
        Result(SSF_Comments) = SubStringBetween(vText, "简介:", vbCrLf, True)
    End If
    SSLIB_ParseInfoText = Result
End Function



Public Function SSLIB_ParseInfoRule(ByRef vFilename As String, ByRef vUrls() As String) As Long
    On Error GoTo Error_ParseInfoRule
    
    Dim srcLen As Long
    Dim fNUM As Integer
    Dim Result() As Byte
    
    fNUM = FreeFile
    Open vFilename For Binary Access Read Shared As #fNUM
    
    srcLen = LOF(fNUM) - &H44
    ReDim Source(0 To srcLen) As Byte
    Seek #fNUM, &H44 + 1
    Get #fNUM, , Source()
    Close #fNUM
    
    Dim CharZero As String
    CharZero = Chr$(0)
    If (UncompressBuf(Source, Result)) Then
        Dim vSource As String
        Dim iCount As Long
        vSource = StrConv(Result, vbUnicode)
        Dim iStart As Long
        Dim iEnd As Long
        Dim iPos As Long
        iPos = 1
        Do
            iEnd = InStr(iPos, vSource, ".pdg", vbTextCompare)
            If iEnd > 0 Then
                iStart = InStrRev(vSource, CharZero, iEnd)
                If iStart > 0 Then
                    ReDim Preserve vUrls(0 To iCount)
                    vUrls(iCount) = Mid$(vSource, iStart + 1, iEnd - iStart + 3)
                    iCount = iCount + 1
                    iPos = iEnd + 4
                    'vSource = Mid$(vSource, iEnd + 4)
                Else
                    Exit Do
                End If
            Else
                Exit Do
            End If
        Loop
        
        SSLIB_ParseInfoRule = iCount
    End If
    

    Exit Function
Error_ParseInfoRule:
    SSLIB_ParseInfoRule = -1
End Function


Public Function GETpdginfo(infofile As String) As pdginfo

    On Error GoTo ErrorGetPdgInfo
    
    Dim thisbook As pdginfo
    Dim Num As Integer
    
    Dim TmpStr As String
    Dim str1 As String, str2 As String

    
    If Not FileExists(infofile) Then Exit Function

    
        thisbook.vailable = True
  
    Dim fNUM As Integer
    fNUM = FreeFile
    Open infofile For Input As #fNUM
    Do While Not EOF(fNUM)
        Line Input #fNUM, TmpStr
        Num = InStr(1, TmpStr, "=")
        If Num > 0 Then
        str1 = RTrim(LTrim(Left(TmpStr, Num - 1)))
        str2 = RTrim(LTrim(Right(TmpStr, Len(TmpStr) - Num)))
        
        Select Case str1
        Case "作者"
            thisbook.Author = str2
        Case "书名"
            thisbook.Name = str2
        Case "下载位置"
            thisbook.download = str2
        Case "页数"
            thisbook.totalpage = str2
        Case "出版社"
            thisbook.Publisher = str2
        Case "出版日期"
            thisbook.pdate = str2
        Case "SS号"
            thisbook.SSID = str2
        End Select
            
        End If
    
    Loop
    If (GetAttr(infofile) Mod 2) = 1 Then thisbook.isreadonly = True Else thisbook.isreadonly = False
    Close #fNUM
    
'        If thisbook.isreadonly = False Then
'        If thisbook.author = "" Or thisbook.author = "BEXP" Then
'            thisbook.author = InputBox("No author information for file: " + thisbook.name + Chr(13) + Chr(10) + "Enter it below:", "PDGinfo")
'            Set ft = ff.OpenAsTextStream(ForWriting)
'            ft.WriteLine "书名=" + thisbook.name
'            ft.WriteLine "作者=" + thisbook.author
'            ft.WriteLine "下载位置=" + thisbook.download
'            ft.Close
'
'        End If
'        End If
    GETpdginfo = thisbook
    Exit Function
ErrorGetPdgInfo:
    On Error Resume Next
        Close #fNUM
        GETpdginfo = thisbook
        Err.Clear
End Function



Public Function getpdg(strcatch As String) As PDG

Dim thispdg As PDG
    With thispdg
        .infofile = ""
        .infolder = ""
        .iszip = False
        .unzipfolder = ""
        .zipfile = ""
        .isreadonly = False
    End With
    
    With thispdg.info
        .Author = ""
        .download = "'"
        .Name = ""
        .totalpage = 0
        .vailable = False
        .isreadonly = False
    End With

    Dim strtype As VbFileAttribute
   ' Dim fso As New FileSystemObject

    strtype = GetAttr(strcatch)
    
    If strtype = vbDirectory Or strtype = vbDirectory + vbReadOnly Then
        
        If strtype = vbDirectory + vbReadOnly Then thispdg.isreadonly = True
                
        thispdg.infolder = strcatch
        strcatch = BuildPath(strcatch)
        
        If Dir(strcatch + "*.pdg") <> "" Then
            thispdg.infofile = strcatch + "BOOKINFO.DAT"
        ElseIf Dir(strcatch + "*.ZIP") <> "" Then
            thispdg.unzipfolder = Environ("temp") + "\PdgZF"
            thispdg.zipfile = strcatch + Dir(strcatch + "*.zip")
            thispdg.infofile = strcatch + "bookinfo.dat"
            thispdg.iszip = True
        Else
            MsgBox ("Error:Not pdg folder or zipfile")
            End
        End If
        
    End If
    If strtype = 32 Or strtype = vbArchive + vbReadOnly Or strtype = vbReadOnly Then
    If strtype = vbArchive + vbReadOnly Then thispdg.isreadonly = True
    If strtype = vbReadOnly Then thispdg.isreadonly = True
    If LCase(GetExtensionName(strcatch)) <> "zip" Then
    MsgBox ("Error:NOT pdg folder or zipfile")
    End
    End If
    thispdg.infolder = GetParentFolderName(strcatch)
    thispdg.unzipfolder = Environ("temp") + "\PdgZF"
    thispdg.zipfile = strcatch
    thispdg.infofile = BuildPath(thispdg.infolder, "bookinfo.dat")
    thispdg.iszip = True
    End If
    
    thispdg.info = GETpdginfo(thispdg.infofile)
    Dir Environ("temp")
    

    
    
    getpdg = thispdg
 
End Function



Public Sub checkpdg(thispdg As PDG)

With thispdg

If .infolder = "" Then MsgBox "CHECK PDG ERROR": Exit Sub
If .infofile = "" Then MsgBox "CHECK PDG ERROR": Exit Sub
If Not FolderExists(.infolder) Then MsgBox "CHECK PDG ERROR": Exit Sub
If Not FileExists(.infofile) Then MsgBox "CHECK PDG ERROR": Exit Sub
If .iszip Then
    If .zipfile = "" Then MsgBox "CHECK PDG ERROR": Exit Sub
    If .unzipfolder = "" Then MsgBox "CHECK PDG ERROR": Exit Sub
    If Not FileExists(.zipfile) Then MsgBox "CHECK PDG ERROR": Exit Sub
End If


End With
End Sub

Public Function pdgformat(thispdg As pdginfo, formatstr As String) As String
Dim TmpStr As String
TmpStr = formatstr
If MyInstr(formatstr, "%t,%a,%p,%c,%d") = False Then Exit Function

'If InStr(tmpstr, "%title") = 0 And InStr(tmpstr, "%author") = 0 And InStr(tmpstr, "%pages") = 0 Then Exit Function

TmpStr = Replace(TmpStr, "%t", thispdg.Name)
TmpStr = Replace(TmpStr, "%a", thispdg.Author)
TmpStr = Replace(TmpStr, "%p", thispdg.totalpage)
TmpStr = Replace(TmpStr, "%c", thispdg.Publisher)
TmpStr = Replace(TmpStr, "%d", thispdg.pdate)
TmpStr = Replace(TmpStr, "%s", thispdg.SSID)

Dim vStr() As String
Dim i As Long
Dim iL As Long
Dim iU As Long
Dim sPart As String
vStr = Split(TmpStr, "-")
iL = LBound(vStr)
iU = UBound(vStr)

pdgformat = LTrim$(RTrim$(vStr(iL)))
iL = iL + 1
For i = iL To iU
    sPart = LTrim$(RTrim$(vStr(i)))
    If sPart <> "" Then pdgformat = pdgformat & " - " & sPart
Next

pdgformat = Replace$(pdgformat, "()", "")
pdgformat = Replace$(pdgformat, "[]", "")
pdgformat = Replace$(pdgformat, "《》", "")
pdgformat = Replace$(pdgformat, "［］", "")
pdgformat = Replace$(pdgformat, "“”", "")
pdgformat = Replace$(pdgformat, Chr$(34) & Chr$(34), "")
If (Right$(pdgformat, 3) = " - ") Then pdgformat = Mid$(pdgformat, 1, Len(pdgformat) - 3)
If (Left$(pdgformat, 3) = " - ") Then pdgformat = Left$(pdgformat, Len(pdgformat) - 3)
pdgformat = Trim$(pdgformat)

'pdgformat = Replace$(pdgformat, "()", "")


'pdgformat = tmpstr

End Function


Public Function TextPdgCount(ByVal vURL As String) As Long
    On Error Resume Next
    TextPdgCount = 1
    vURL = GetBaseName(vURL)
    If StringToLong(Left$(vURL, 1)) < 1 Then Exit Function
    Dim i As Long
    i = InStrRev(vURL, "_")
    If i > 0 Then
        TextPdgCount = CLng(Right$(vURL, Len(vURL) - i))
    End If
End Function
