Attribute VB_Name = "MSSReader"
'CSEH: ErrDebugPrint
Option Explicit
#Const ZLIB = 1

Private Const Z_OK As Long = &H0

Private Declare Function compress Lib "ZLibWAPI.dll" ( _
    ByRef dest As Any, ByRef destLen As Long, _
    ByRef Source As Any, ByVal sourceLen As Long) As Long
    
Private Declare Function compressBound Lib "ZLibWAPI.dll" ( _
    ByVal sourceLen As Long) As Long
    
Private Declare Function uncompress Lib "ZLibWAPI.dll" ( _
    ByRef dest As Any, ByRef destLen As Long, _
    ByRef Source As Any, ByVal sourceLen As Long) As Long
    
Private Declare Function adler32 Lib "ZLibWAPI.dll" ( _
    ByVal adler As Long, ByRef buf As Any, ByVal length As Long) As Long
    
Private Declare Function crc32 Lib "ZLibWAPI.dll" ( _
    ByVal crc As Long, ByRef buf As Any, ByVal length As Long) As Long

Private Declare Function zlibCompileFlags Lib "ZLibWAPI.dll" () As Long

Private Const Z_NO_COMPRESSION As Long = 0
Private Const Z_BEST_SPEED As Long = 1
Private Const Z_BEST_COMPRESSION As Long = 9
Private Const Z_DEFAULT_COMPRESSION As Long = (-1)
 
Private Declare Function compress2 Lib "ZLibWAPI.dll" ( _
    ByRef dest As Any, ByRef destLen As Long, _
    ByRef Source As Any, ByVal sourceLen As Long, _
    ByVal level As Long) As Long
Public Type pdginfo
    vailable As Boolean
    isreadonly As Boolean
    Name As String '%t
    Author As String '%a
    totalpage As String '%p
    Download As String '%u
    Publisher As String '%c
    pdate As String '%d
    ssid As String '%s
End Type

Public Type PDG

    iszip As Boolean
    isreadonly As Boolean
    infolder As String
    unzipfolder As String
    zipfile As String
    infofile As String
    Info As pdginfo

End Type



Public Type SSLIB_BOOKINFO
    Title As String
    Author As String
    ssid As String
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

#If Not afComponent = 1 Then

Public Enum SSLIBTaskStatus
    STS_START = 1
    STS_PAUSE = 0
    STS_COMPLETE = 2
    STS_PENDING = 3
    STS_ERRORS = 4
End Enum

Public Const CST_SSLIB_FIELDS_LBound As Long = 1
Public Enum SSLIBFields
        SSF_Title = CST_SSLIB_FIELDS_LBound
        SSF_AUTHOR
        SSF_SSID
        SSF_ISJPGBOOK
        SSF_PagesCount
        SSF_Publisher
        SSF_PublishDate
        SSF_StartPage
        
        SSF_Subject
        SSF_Comments
        
        SSF_PAGEURL
        SSF_URL
        SSF_SSURL
        SSF_JPGURL
        SSF_IEJPGURL
        
        SSF_PARAMS
        SSF_SAVEDIN
        SSF_HEADER
        'SSF_FULLNAME
        
        SSF_Downloader
        SSF_DownloadDate
        'SSF_STATUS
        'SSF_FILES_DOWNLOADED
        
        
        SSF_HTMLContent
        SSF_FIELDS_END

End Enum

Public Const CST_SSLIB_FIELDS_UBound As Long = SSF_FIELDS_END - 1
Public Const CST_SSLIB_FIELDS_IMPORTANT_UBOUND As Long = SSF_Comments
Public Const CST_SSLIB_FIELDS_TASKS_UBOUND As Long = SSF_HEADER

#Else
Public Const CST_SSLIB_FIELDS_LBound As Long = 1
Public Const CST_SSLIB_FIELDS_UBound As Long = SSF_FIELDS_END - 1
Public Const CST_SSLIB_FIELDS_IMPORTANT_UBOUND As Long = SSF_Comments
Public Const CST_SSLIB_FIELDS_TASKS_UBOUND As Long = SSF_HEADER
#End If


Private SSLIB_FIELDS_NAME(CST_SSLIB_FIELDS_LBound To CST_SSLIB_FIELDS_UBound, 1 To 2) As String
Private Const cst_sslibrary_main As String = "http://pds.sslibrary.com"
Private Const cst_sslibrary_gojpg As String = "gojpgRead.jsp?ssnum="
#If Not afNoCWinHTTP = 1 Then
    Private mLastCookie As String
#End If
  
Public Function UncompressString(ByRef Source() As Byte) As String
Dim result() As Byte
If (UncompressBuf(Source, result)) Then
    UncompressString = StrConv(result, vbUnicode)
End If
End Function

Public Function UncompressFile(ByRef FileName As String, ByRef result() As Byte, Optional seekPos As Long = 1) As Boolean
    '<EhHeader>
    On Error GoTo UncompressFile_Err
    '</EhHeader>
    
    Dim srcLen As Long
    Dim fNum As Integer
    fNum = FreeFile
    Open FileName For Binary Access Read As #fNum
    srcLen = LOF(fNum) - seekPos + 1
    ReDim Source(0 To srcLen - 1) As Byte
    Seek #fNum, seekPos
    Get #fNum, , Source()
    Close #fNum
    UncompressFile = UncompressBuf(Source, result)
    '<EhFooter>
    Exit Function

UncompressFile_Err:
    Debug.Print "GetSSLib.MSSReader.UncompressFile:Error " & Err.Description
    
    '</EhFooter>
End Function

Public Function UncompressFileTo(ByRef srcFile As String, ByRef destFile As String, Optional seekPos As Long = 1) As Boolean
    '<EhHeader>
    On Error GoTo UncompressFileTo_Err
    '</EhHeader>

    Dim result() As Byte
    If (UncompressFile(srcFile, result, seekPos)) Then
        Dim fNum As Integer
        fNum = FreeFile
        Open destFile For Binary Access Write Lock Read As #fNum
        Put #fNum, , result
        Close #fNum
        UncompressFileTo = True
    End If
   
    '<EhFooter>
    Exit Function

UncompressFileTo_Err:
    Debug.Print "GetSSLib.MSSReader.UncompressFileTo:Error " & Err.Description
    
    '</EhFooter>
End Function

Public Function UncompressBuf(Source() As Byte, ByRef result() As Byte) As Boolean
    On Error GoTo invalidUsage
        Dim destLen As Long
        Dim sourceLen As Long
        Dim ret As Long
        sourceLen = UBound(Source()) + 1
        destLen = sourceLen
        Do
            destLen = destLen * 2
            ReDim result(0 To destLen - 1)
            ret = uncompress(result(0), destLen, Source(0), sourceLen)
            
         '   If ret = -3 Then Exit Do
        Loop While ret = -5
        'Until ret = Z_OK
        ReDim Preserve result(0 To destLen - 1)
        If ret = Z_OK Then UncompressBuf = True
        'UncompressBuf = result
        'Dim result(0 to ubound(source())
    Exit Function
invalidUsage:
    Debug.Print Err.Description
    UncompressBuf = False
End Function

Public Function UncompressFileToString(srcFile As String, Optional ByVal seekPos As Long = 1) As String
    Dim result() As Byte
    If (UncompressFile(srcFile, result, seekPos)) Then
        UncompressFileToString = StrConv(result, vbUnicode)
    End If
End Function

Public Function UncompressBufToString(Source() As Byte) As String
    Dim result() As Byte
    If (UncompressBuf(Source, result)) Then
        UncompressBufToString = StrConv(result, vbUnicode)
    End If
End Function
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
    
    SSLIB_FIELDS_NAME(SSLIBFields.SSF_AUTHOR, 1) = "Author"
    SSLIB_FIELDS_NAME(SSLIBFields.SSF_AUTHOR, 2) = "作者"
    
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
    
    SSLIB_FIELDS_NAME(SSLIBFields.SSF_PAGEURL, 1) = "PageURL"
    SSLIB_FIELDS_NAME(SSLIBFields.SSF_PAGEURL, 2) = "页面链接"
    
    SSLIB_FIELDS_NAME(SSLIBFields.SSF_HTMLContent, 1) = "HTMLContent"
    SSLIB_FIELDS_NAME(SSLIBFields.SSF_HTMLContent, 2) = "HTML目录"
    
    SSLIB_FIELDS_NAME(SSLIBFields.SSF_ISJPGBOOK, 1) = "JPGBook"
    SSLIB_FIELDS_NAME(SSLIBFields.SSF_ISJPGBOOK, 2) = "JPG大图"
    
    SSLIB_FIELDS_NAME(SSLIBFields.SSF_JPGURL, 1) = "JPGURL"
    SSLIB_FIELDS_NAME(SSLIBFields.SSF_JPGURL, 2) = "JPGURL"
    
        
    SSLIB_FIELDS_NAME(SSLIBFields.SSF_SSURL, 1) = "SSReaderURL"
    SSLIB_FIELDS_NAME(SSLIBFields.SSF_SSURL, 2) = "超星阅读器URL"
    
    SSLIB_FIELDS_NAME(SSLIBFields.SSF_PARAMS, 1) = "Params"
    SSLIB_FIELDS_NAME(SSLIBFields.SSF_PARAMS, 2) = "Params"
    
    
    SSLIB_FIELDS_NAME(SSLIBFields.SSF_IEJPGURL, 1) = "IEJPGURL"
    SSLIB_FIELDS_NAME(SSLIBFields.SSF_IEJPGURL, 2) = "IEJPGURL"
    
    SSLIB_FLAGS_INIT = True
    
End Sub

Public Function SSLIB_GetFiledID(ByVal vFieldname As String) As Long
    Dim i As Long
    i = -1
    For i = CST_SSLIB_FIELDS_LBound To CST_SSLIB_FIELDS_UBound
        If vFieldname = SSLIB_FIELDS_NAME(i, 1) Then
            SSLIB_GetFiledID = i: Exit Function
        ElseIf vFieldname = SSLIB_FIELDS_NAME(i, 2) Then
            SSLIB_GetFiledID = i: Exit Function
        End If
    Next
End Function

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
'    If vText = vbNULLSTRING Then vText = Clipboard.GetText
'    If vText = vbNULLSTRING Then Exit Function
'    vText = vText & vbCrLf
'    If InStr(1, vText, vbCrLf & "SSCT:", vbTextCompare) > 0 Then
'        SSLIB_ParseInfoText.Header = Replace$(vText, "(Request-Line):", vbNULLSTRING)
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
    Dim result(CST_SSLIB_FIELDS_LBound To CST_SSLIB_FIELDS_UBound) As String
    SSLIB_CreateBookInfoArray = result
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
    Dim result() As String
    result = SSLIB_CreateBookInfoArray()
    Dim TmpStr As String
    If vText = vbNullString Then vText = Clipboard.GetText
    If vText = vbNullString Then Exit Function
    vText = vText & vbCrLf
    If InStr(1, vText, vbCrLf & "SSCT:", vbTextCompare) > 0 Then
        vText = Replace$(vText, "(Request-Line):", vbNullString)
        result(SSF_HEADER) = vText
        result(SSF_URL) = SubStringBetween(vText, "Host:", vbCrLf, True)
        If result(SSF_URL) <> vbNullString Then result(SSF_URL) = "http://" & result(SSF_URL) & "/"
        vText = SubStringBetween(vText, "Get /", " ", True)
        If vText <> vbNullString Then result(SSF_URL) = result(SSF_URL) & vText
        
        
    Else
        vText = Replace$(vText, "</DIV>", vbNullString, , , vbTextCompare)
        result(SSF_SSID) = SubStringBetween(vText, "SS号:", vbCrLf, True)
        TmpStr = SubStringBetween(vText, "作者:", " ", True)
        If Right$(TmpStr, 1) = "著" Then TmpStr = Left$(TmpStr, Len(TmpStr) - 1)
        result(SSF_AUTHOR) = TmpStr
        TmpStr = SubStringUntilMatch(TmpStr, 1, " /")
        If TmpStr <> vbNullString Then result(SSF_AUTHOR) = TmpStr
        result(SSF_Title) = SubStringBetween(vText, "《", "》", True)
        If InStr(result(SSF_Title), "《") > 0 Then
            TmpStr = SubStringBetween(vText, "》", "》", True)
            If result(SSF_Title) <> vbNullString And TmpStr <> vbNullString Then
                result(SSF_Title) = result(SSF_Title) & "》" & TmpStr
            End If
        End If
        result(SSF_PagesCount) = SubStringBetween(vText, "页数:", " ", True)
        result(SSF_PublishDate) = SubStringBetween(vText, "出版日期:", " ", True)
        result(SSF_Publisher) = SubStringBetween(vText, "出版社:", " ", True)
        result(SSF_Subject) = SubStringBetween(vText, "主题词:", vbCrLf, True)
        result(SSF_Comments) = SubStringBetween(vText, "简介:", vbCrLf, True)
    End If
    SSLIB_ParseInfoText = result
End Function


#If ZLIB Then
Public Function SSLIB_ParseInfoRule(ByRef vFilename As String, ByRef vUrls() As String) As Long
    On Error GoTo Error_ParseInfoRule
    
    Dim srcLen As Long
    Dim fNum As Integer
    Dim result() As Byte
    
    fNum = FreeFile
    Open vFilename For Binary Access Read Shared As #fNum
    
    srcLen = LOF(fNum) - &H44
    ReDim Source(0 To srcLen) As Byte
    Seek #fNum, &H44 + 1
    Get #fNum, , Source()
    Close #fNum
    
    Dim CharZero As String
    CharZero = Chr$(0)
    If (UncompressBuf(Source, result)) Then
        Dim vSource As String
        Dim iCount As Long
        vSource = StrConv(result, vbUnicode)
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

#End If

Public Function GETpdginfo(infofile As String) As pdginfo

    On Error GoTo ErrorGetPdgInfo
    
    Dim thisbook As pdginfo
    Dim Num As Integer
    
    Dim TmpStr As String
    Dim str1 As String, str2 As String

    
    If Not FileExists(infofile) Then Exit Function

    
        thisbook.vailable = True
  
    Dim fNum As Integer
    fNum = FreeFile
    Open infofile For Input As #fNum
    Do While Not EOF(fNum)
        Line Input #fNum, TmpStr
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
            thisbook.Download = str2
        Case "页数"
            thisbook.totalpage = str2
        Case "出版社"
            thisbook.Publisher = str2
        Case "出版日期"
            thisbook.pdate = str2
        Case "SS号"
            thisbook.ssid = str2
        End Select
            
        End If
    
    Loop
    If (GetAttr(infofile) Mod 2) = 1 Then thisbook.isreadonly = True Else thisbook.isreadonly = False
    Close #fNum
    
'        If thisbook.isreadonly = False Then
'        If thisbook.author = vbNULLSTRING Or thisbook.author = "BEXP" Then
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
        Close #fNum
        GETpdginfo = thisbook
        Err.Clear
End Function



Public Function getpdg(strcatch As String) As PDG

'Dim thispdg As PDG
'    With thispdg
'        .infofile = vbNULLSTRING
'        .infolder = vbNULLSTRING
'        .iszip = False
'        .unzipfolder = vbNULLSTRING
'        .zipfile = vbNULLSTRING
'        .isreadonly = False
'    End With
'
'    With thispdg.Info
'        .Author = vbNULLSTRING
'        .Download = "'"
'        .Name = vbNULLSTRING
'        .totalpage = 0
'        .vailable = False
'        .isreadonly = False
'    End With
'
'    Dim strtype As VbFileAttribute
'   ' Dim fso As New FileSystemObject
'
'    strtype = GetAttr(strcatch)
'
'    If strtype = vbDirectory Or strtype = vbDirectory + vbReadOnly Then
'
'        If strtype = vbDirectory + vbReadOnly Then thispdg.isreadonly = True
'
'        thispdg.infolder = strcatch
'        strcatch = BuildPath(strcatch)
'
'        If Dir(strcatch + "*.pdg") <> vbNULLSTRING Then
'            thispdg.infofile = strcatch + "BOOKINFO.DAT"
'        ElseIf Dir(strcatch + "*.ZIP") <> vbNULLSTRING Then
'            thispdg.unzipfolder = Environ("temp") + "\PdgZF"
'            thispdg.zipfile = strcatch + Dir(strcatch + "*.zip")
'            thispdg.infofile = strcatch + "bookinfo.dat"
'            thispdg.iszip = True
'        Else
'            'MsgBox ("Error:Not pdg folder or zipfile")
'            Exit Function
'        End If
'
'    End If
'    If strtype = 32 Or strtype = vbArchive + vbReadOnly Or strtype = vbReadOnly Then
'    If strtype = vbArchive + vbReadOnly Then thispdg.isreadonly = True
'    If strtype = vbReadOnly Then thispdg.isreadonly = True
'    If LCase(GetExtensionName(strcatch)) <> "zip" Then
'    MsgBox ("Error:NOT pdg folder or zipfile")
'    End
'    End If
'    thispdg.infolder = GetParentFolderName(strcatch)
'    thispdg.unzipfolder = Environ("temp") + "\PdgZF"
'    thispdg.zipfile = strcatch
'    thispdg.infofile = BuildPath(thispdg.infolder, "bookinfo.dat")
'    thispdg.iszip = True
'    End If
'
'    thispdg.Info = GETpdginfo(thispdg.infofile)
'    Dir Environ("temp")
'
'
'
'
'    getpdg = thispdg
 
End Function



Public Sub checkpdg(thispdg As PDG)

With thispdg

If .infolder = vbNullString Then MsgBox "CHECK PDG ERROR": Exit Sub
If .infofile = vbNullString Then MsgBox "CHECK PDG ERROR": Exit Sub
If Not FolderExists(.infolder) Then MsgBox "CHECK PDG ERROR": Exit Sub
If Not FileExists(.infofile) Then MsgBox "CHECK PDG ERROR": Exit Sub
If .iszip Then
    If .zipfile = vbNullString Then MsgBox "CHECK PDG ERROR": Exit Sub
    If .unzipfolder = vbNullString Then MsgBox "CHECK PDG ERROR": Exit Sub
    If Not FileExists(.zipfile) Then MsgBox "CHECK PDG ERROR": Exit Sub
End If


End With
End Sub

Public Function pdgformat(thispdg As CBookInfo, formatstr As String) As String
Dim TmpStr As String
TmpStr = formatstr
If MyInstr(formatstr, "%t,%a,%p,%c,%d") = False Then Exit Function

'If InStr(tmpstr, "%title") = 0 And InStr(tmpstr, "%author") = 0 And InStr(tmpstr, "%pages") = 0 Then Exit Function

TmpStr = Replace(TmpStr, "%t", thispdg(SSF_Title))
TmpStr = Replace(TmpStr, "%a", thispdg(SSF_AUTHOR))
TmpStr = Replace(TmpStr, "%p", thispdg(SSF_PagesCount))
TmpStr = Replace(TmpStr, "%c", thispdg(SSF_Publisher))
TmpStr = Replace(TmpStr, "%d", thispdg(SSF_PublishDate))
TmpStr = Replace(TmpStr, "%s", thispdg(SSF_SSID))

Dim vStr() As String
Dim i As Long
Dim iL As Long
Dim iU As Long
Dim sPart As String
vStr = Split(TmpStr, "-")
iL = LBound(vStr)
iU = UBound(vStr)

pdgformat = vStr(iL)
iL = iL + 1
For i = iL To iU
    sPart = LTrim$(RTrim$(vStr(i)))
    If sPart <> vbNullString Then pdgformat = pdgformat & "-" & vStr(i)
Next

pdgformat = Replace$(pdgformat, "()", vbNullString)
pdgformat = Replace$(pdgformat, "[]", vbNullString)
pdgformat = Replace$(pdgformat, "《》", vbNullString)
pdgformat = Replace$(pdgformat, "［］", vbNullString)
pdgformat = Replace$(pdgformat, "“”", vbNullString)
pdgformat = Replace$(pdgformat, Chr$(34) & Chr$(34), vbNullString)
 Do While (Right$(pdgformat, 3) = " - ")
 pdgformat = Left$(pdgformat, Len(pdgformat) - 3)
 Loop
Do While (Left$(pdgformat, 3) = " - ")
    pdgformat = Right$(pdgformat, Len(pdgformat) - 3)
Loop
pdgformat = Trim$(pdgformat)

'pdgformat = Replace$(pdgformat, "()", vbNULLSTRING)


'pdgformat = tmpstr

End Function


Public Function TextPdgCount(ByVal vUrl As String) As Long
    On Error Resume Next
    TextPdgCount = 1
    vUrl = GetBaseName(vUrl)
    If StringToLong(Left$(vUrl, 1)) < 1 Then Exit Function
    Dim i As Long
    i = InStrRev(vUrl, "_")
    If i > 0 Then
        TextPdgCount = CLng(Right$(vUrl, Len(vUrl) - i))
    End If
End Function

#If Not afNoCWinHTTP = 1 Then
Public Sub GetJpgBook(ByVal vFilename As String, ByVal vSaveAS As String, Optional vHeader As String)
    If vHeader = vbNullString Then vHeader = Clipboard.GetText
    vSaveAS = BuildPath(vSaveAS, vFilename)
    'If Right$(vSaveAs, 1) = "\" Then vSaveAs = vSaveAs & vFilename
Dim pHeader As CHttpHeader
   Set pHeader = New CHttpHeader
   pHeader.Init vHeader

Dim pUrl As String

        Dim Host As String
        Dim action As String
        Host = pHeader.GetField("host")
        action = pHeader.GetField(vbNullString)
        action = SubStringBetween(action, "POST ", " HTTP/1", False)
        If Host <> vbNullString And action <> vbNullString Then
            pUrl = "http://" & Host & action
        End If
        pHeader.DeleteField (vbNullString)
    Debug.Print pUrl
   'pHeader.DeleteField (vbNULLSTRING)
   'mBookInfo(SSF_HEADER) = pHeader.HeaderString
   If pUrl = vbNullString Then Exit Sub
   pUrl = SubStringUntilMatch(pUrl, 1, "%26jid%3D")
   pUrl = pUrl & "%26jid%3D" & "/" & EscapeUrl(vFilename)
   Dim pHttp As IWinHttp
   Set pHttp = New CWinHttpSimple
   pHttp.URL = pUrl
   pHttp.method = "POST"
   pHttp.Header = pHeader.HeaderString
   pHttp.OpenConnect False
   pHttp.Options(WinHttpRequestOption_EnableHttpsToHttpRedirects) = 1
   pHttp.send
   Do Until pHttp.IsFree
    DoEvents
   Loop
   
   If pHttp.Status <> 200 Then Exit Sub
   
   Dim pHeader2 As String
   pHeader2 = pHttp.ResponseHeader
   pHeader.SetField "Cookie", pHeader.GetField("Cookie") & "; " & SubStringBetween(pHeader2, "Set-Cookie:", vbCrLf, True)
   pHeader.DeleteField ("host")
   pHeader.DeleteField vbNullString
   
   Dim pResponse As String
   pResponse = StrConv(pHttp.responseBody, vbUnicode)
   'pHeader.SetField
   'Debug.Print pHeader.HeaderString
   pUrl = UnescapeUrl(SubStringBetween(pUrl, "=", vbNullString, True))
   If pUrl <> vbNullString Then
        pUrl = pUrl & "&a=" & pResponse & "&uf=ssr&zoom=2"
   Set pHttp = New CWinHttpSimple
   pHttp.Init
   pHttp.URL = pUrl
   pHttp.method = "GET"
   pHttp.Header = pHeader.HeaderString
   pHttp.Destination = vSaveAS
   pHttp.OpenConnect False
   pHttp.send
   Do Until pHttp.IsFree
   DoEvents
   Loop
   Debug.Print pHttp.ResponseHeader
   End If
End Sub


Public Function GetJpgBookInfoUrl2(ByRef pUrl As String, Optional ByRef vCookie As String = vbNullString, Optional ByRef vHttp As IWinHttp)
If vCookie = vbNullString Then vCookie = mLastCookie
    If vCookie = vbNullString Then vCookie = "bkname=ssgpgdjy; UID=18523; state=1; lib=all; AID=161; tbExist=; exp=" & Chr$(34) & "2009-01-23 00:00:00.0" & Chr$(34) & "; allBooks=0; send=35534BF01F5A60E8BCE2EA46D1F624AF; userLogo=ssgpgdjy.gif; company=%u5e7f%u4e1c%u6559%u80b2%u5b66%u9662; marking=all; bnp=1.0; showDuxiu=0; showTopbooks=0; goWhere=1; JSESSIONID=75E5DC21B512ECEA267711A7A80360BA.tomcat2"

    If vHttp Is Nothing Then Set vHttp = New CWinHttpSimple
    vHttp.Init
    vHttp.URL = pUrl
    vHttp.method = "GET"
    'vHttp.SetTimeouts 20, 20, 20, 20
    vHttp.Options(WinHttpRequestOption_EnableRedirects) = 0
    vHttp.OpenConnect False
    
    vHttp.setRequestHeader "Cookie", vCookie
    
    vHttp.send
    
    Do Until vHttp.IsFree
        DoEvents
    Loop
    
    Dim pHeader As String
    pHeader = vHttp.ResponseHeader
    
    'If vHttp.Status <> 200 Then Exit Sub
    pUrl = HttpHeaderGetField(vHttp.ResponseHeader, "Location")

    If pUrl = vbNullString Then
        pUrl = SubStringBetween(vHttp.responseText, "window.location.href='", "'", False)
    End If
    If InStr(1, pUrl, "kid=", vbTextCompare) > 0 Then
        vCookie = HttpHeaderMergeCookie(vCookie, HttpHeaderSetCookie(vHttp.ResponseHeader))
        mLastCookie = vCookie
        GetJpgBookInfoUrl2 = pUrl
    Else
        'MsgBox "Cant'not get jpgurl from ssid,maybe param D needed", vbCritical
    End If
    Debug.Print GetJpgBookInfoUrl2
End Function

Public Function GetJpgBookInfoUrl(ByRef vSSID As String, Optional ByRef vCookie As String = vbNullString, Optional vMainSite As String = cst_sslibrary_main, Optional ByRef vHttp As IWinHttp, Optional d As String) As String
    Dim pUrl As String
    pUrl = vMainSite & "/" & cst_sslibrary_gojpg & vSSID
    If (d <> vbNullString) Then
        If (InStr(d, "=") > 0) Then
            pUrl = pUrl & "&" & d
        Else
            pUrl = pUrl & "&d=" & d
        End If
    End If
    GetJpgBookInfoUrl = GetJpgBookInfoUrl2(pUrl, vCookie, vHttp)
End Function


Public Function IEURL_To_BookInfo(ByVal vUrl As String, ByRef vCookie As String) As String()
    Dim pUrl As String
    pUrl = GetJpgBookInfoUrl2(vUrl, vCookie)
    If pUrl <> vbNullString Then IEURL_To_BookInfo = ParseJPGBookInfoText(GetJpgBookInfo(pUrl, vCookie))
End Function



Public Function SSID_To_BookInfo(ByVal vSSID As String, ByRef vCookie As String, Optional vMainSite As String = cst_sslibrary_main, Optional d As String) As String()
    Dim pUrl As String
    pUrl = GetJpgBookInfoUrl(vSSID, vCookie, vMainSite, Nothing, d)
    If pUrl <> vbNullString Then SSID_To_BookInfo = ParseJPGBookInfoText(GetJpgBookInfo(pUrl, vCookie))
End Function

Public Function GetJpgBookInfo(ByVal vUrl As String, Optional ByRef vCookie As String = vbNullString, Optional ByRef vHttp As IWinHttp = Nothing) As String
    If vCookie = vbNullString Then vCookie = mLastCookie
    If vCookie = vbNullString Then vCookie = "bkname=ssgpgdjy; UID=18523; state=1; lib=all; AID=161; tbExist=; exp=" & Chr$(34) & "2009-01-23 00:00:00.0" & Chr$(34) & "; allBooks=0; send=35534BF01F5A60E8BCE2EA46D1F624AF; userLogo=ssgpgdjy.gif; company=%u5e7f%u4e1c%u6559%u80b2%u5b66%u9662; marking=all; bnp=1.0; showDuxiu=0; showTopbooks=0; goWhere=1; JSESSIONID=75E5DC21B512ECEA267711A7A80360BA.tomcat2"
    If vUrl = vbNullString Then Exit Function
    If vHttp Is Nothing Then Set vHttp = New CWinHttpSimple
    vHttp.Init
    vHttp.URL = vUrl
    vHttp.method = "GET"
    vHttp.Options(WinHttpRequestOption_EnableRedirects) = 0
    vHttp.OpenConnect False
    vHttp.setRequestHeader "Cookie", vCookie
    vHttp.send
    
    Do Until vHttp.IsFree
        DoEvents
    Loop
    
    Dim ret As String
    ret = Trim$(StrConv(vHttp.responseBody, vbUnicode))
    If ret <> vbNullString Then
       vCookie = HttpHeaderMergeCookie(vCookie, HttpHeaderGetField(vHttp.ResponseHeader, "Set-Cookie"))
        mLastCookie = vCookie
       GetJpgBookInfo = ret
    End If


End Function

Public Function SSLIB_GetReadBookURL(vShowBookUrl As String, Optional vSite As String = "http://pds.sslibrary.com", Optional vCookie As String) As String
    If vCookie = vbNullString Then vCookie = Clipboard.GetText
    If vShowBookUrl = vbNullString Then vShowBookUrl = "showbook.do?dxNumber=10108311&d=99653DE6D2EBC329B6D3F914FECB92E4&fenleiID=0I20100140&nettype=wangtong&username=ssgpgdjy"
    Dim pHttp As IWinHttp
    Set pHttp = New CWinHttpSimple
    With pHttp
        .Init
        .URL = BuildPath(vSite, vShowBookUrl, lnpsUnix)
        .method = "GET"
        .OpenConnect False
        .setRequestHeader "Cookie", vCookie
        .send
    End With
    Do Until pHttp.IsFree
        DoEvents
    Loop
    Debug.Print UnescapeUrl(pHttp.responseText)
End Function
Public Function SSLibraryLogin(ByVal vUsername As String, ByVal vPassWord As String, ByVal vLoginUrl As String, ByVal vPostUrl As String, Optional ByVal vDisplayFailData As Boolean = False) As String

    If vUsername = vbNullString Then vUsername = "ssgpgdjy"
    If vPassWord = vbNullString Then vPassWord = "shngdjy"
    If vLoginUrl = vbNullString Then vLoginUrl = "http://pds.sslibrary.com/userlogon.jsp"
    If vPostUrl = vbNullString Then vPostUrl = "http://pds.sslibrary.com/loginhl.jsp"
    
    Dim Http As IWinHttp
    Set Http = New CWinHttpSimple
    Http.Init
    Http.URL = vLoginUrl
    Http.method = "GET"
    Http.Options(WinHttpRequestOption_EnableRedirects) = 0
    Http.OpenConnect False
    Http.send
    Do Until Http.IsFree
        DoEvents
    Loop
    Dim ret As String
    ret = HttpHeaderGetField(Http.ResponseHeader, "Set-Cookie")
    
    Dim postBody As String
    postBody = "send=true&UserName=" & vUsername & "&PassWord=" & vPassWord & "&rd=0&Submit3322=%B5%C7%C2%BC&backurl="
'    Dim postBytes() As Byte
'    postBytes = StrConv(postBody, vbFromUnicode)
'    postBody = postBytes
    Http.Init
    Http.URL = vPostUrl
    Http.method = "POST"
    Http.Options(WinHttpRequestOption_EnableRedirects) = 0
    Http.OpenConnect False
    Http.setRequestHeader "Cookie", ret
    Http.setRequestHeader "Referer", vLoginUrl
    Http.setRequestHeader "Content-Length", Len(postBody)
    Http.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"

    Http.send postBody
    Do Until Http.IsFree
    DoEvents
    Loop
    Dim Header As String
    Header = Http.ResponseHeader
    If InStr(1, Header, "UID=", vbTextCompare) Then
        SSLibraryLogin = HttpHeaderMergeCookie(ret, HttpHeaderSetCookie(Header))
    ElseIf vDisplayFailData Then
        MsgBox StrConv(Http.responseBody, vbUnicode), vbCritical, vPostUrl
    End If
End Function
Public Sub GetJpgBook2(ByVal vCookie As String, ByVal vSSID As String)
'    vCookie = "goWhere=1; bkname=ssgpgdjy; UID=18523; state=1; lib=all; AID=161; tbExist=; allBooks=0; send=E8B93BE35A15D6A3CD27EB5A83D91BF0; userLogo=ssgpgdjy.gif; company=%u5e7f%u4e1c%u6559%u80b2%u5b66%u9662; marking=all; bnp=1.0; showDuxiu=0; showTopbooks=0; JSESSIONID=2E1CA4433C3528DEF6F8FD2860D4566E.tomcat2"
'    vSSID = "11159512"
'
'    Dim pUrl As String
'    Dim pHttp AS IWinHttp
'    'Dim pRefer As String
'
'    Set pHttp = New CWinHttpSimple
'    pUrl = GetJpgBookInfoUrl(vSSID, vCookie, pHttp)
'    If pUrl = vbNULLSTRING Then Exit Sub
'    Dim pContent As String
'    pContent = GetJpgBookInfo(pUrl, vCookie, pHttp)
'    'pRefer = pUrl
'    pUrl = SubStringBetween(pContent, "var str='", "&jid=", True)
'    If pUrl <> vbNULLSTRING Then pUrl = pUrl & "&jid=/"
'
'    Debug.Print pUrl
'
'
'    DownloadPage pUrl, "cov001.jpg", "X:\", vCookie, pHttp
'
'    'DownloadPage vCookie, pRefer, pUrl, "!00001.jpg", "X:\"
'    'DownloadPage vCookie, pRefer, pUrl, "000001.jpg", "X:\"
'    'Debug.Print GetJpgParamA(vCookie, pUrl)
'    'http://pds.sslibrary.com/gojpgRead.jsp?ssnum=10844001
End Sub

Public Function GetJpgBookPageUrl(ByVal vRootUrl As String, ByVal vFilename As String, Optional ByRef vCookie As String, Optional vHttp As IWinHttp = Nothing, Optional vSizeType As Integer = 1) As String
    If vCookie = vbNullString Then vCookie = mLastCookie
    If vCookie = vbNullString Then vCookie = "bkname=ssgpgdjy; UID=18523; state=1; lib=all; AID=161; tbExist=; exp=" & Chr$(34) & "2009-01-23 00:00:00.0" & Chr$(34) & "; allBooks=0; send=35534BF01F5A60E8BCE2EA46D1F624AF; userLogo=ssgpgdjy.gif; company=%u5e7f%u4e1c%u6559%u80b2%u5b66%u9662; marking=all; bnp=1.0; showDuxiu=0; showTopbooks=0; goWhere=1; JSESSIONID=75E5DC21B512ECEA267711A7A80360BA.tomcat2"
    If vHttp Is Nothing Then Set vHttp = New CWinHttpSimple
    
    Dim vUrl As String
    vUrl = vRootUrl & vFilename
    
    vUrl = Replace$(vUrl, ":", "%3A")
    vUrl = Replace$(vUrl, "?", "%3F")
    vUrl = Replace$(vUrl, "&", "%26")
    vUrl = Replace$(vUrl, "=", "%3D")
    vUrl = Replace$(vUrl, "!", "%21")
    If vHttp Is Nothing Then Set vHttp = New CWinHttpSimple
    With vHttp
        .Init
        .URL = "http://img.sslibrary.com/jpgssrurl.jsp?h=" & vUrl
        .method = "POST"
        .OpenConnect False
        .setRequestHeader "Cookie", vCookie
        .send
    End With
    Do Until vHttp.IsFree
        DoEvents
    Loop
    Dim pRet As String
    pRet = StrConv(vHttp.responseBody, vbUnicode)
    If pRet <> vbNullString Then
        If vSizeType < 0 Then vSizeType = 0
        vCookie = HttpHeaderMergeCookie(vCookie, HttpHeaderGetField(vHttp.ResponseHeader, "Set-Cookie"))
        mLastCookie = vCookie
        GetJpgBookPageUrl = vRootUrl & vFilename & "&a=" & pRet & "&uf=ssr&zoom=" & vSizeType
    End If
    
End Function
Private Function DownloadJpgBookPage(vRootUrl As String, vFilename As String, ByVal vSaveTo As String, Optional ByRef vCookie As String, Optional vHttp As IWinHttp) As Long
    If vCookie = vbNullString Then vCookie = mLastCookie
    If vCookie = vbNullString Then vCookie = "bkname=ssgpgdjy; UID=18523; state=1; lib=all; AID=161; tbExist=; exp=" & Chr$(34) & "2009-01-23 00:00:00.0" & Chr$(34) & "; allBooks=0; send=35534BF01F5A60E8BCE2EA46D1F624AF; userLogo=ssgpgdjy.gif; company=%u5e7f%u4e1c%u6559%u80b2%u5b66%u9662; marking=all; bnp=1.0; showDuxiu=0; showTopbooks=0; goWhere=1; JSESSIONID=75E5DC21B512ECEA267711A7A80360BA.tomcat2"
    If vHttp Is Nothing Then Set vHttp = New CWinHttpSimple
    
    
    vSaveTo = BuildPath(vSaveTo, vFilename)
    Dim pUrl As String
    pUrl = GetJpgBookPageUrl(vRootUrl, vFilename, vCookie, vHttp)
        
    If pUrl = vbNullString Then Exit Function
    
    With vHttp
        .Init
        .URL = pUrl
        .method = "GET"
        .Destination = vSaveTo
        .OpenConnect False
        .send
    End With
    
   Do Until vHttp.IsFree
    DoEvents
   Loop
   
   Debug.Print vHttp.Status
   DownloadJpgBookPage = vHttp.BytesDownloaded
   
   'vCookie = HttpHeaderMergeCookie(vCookie, HttpHeaderGetField(pHttp.ResponseHeader, "Set-Cookie"))
   'Debug.Print pHttp.ResponseHeader
   'Debug.Print pHttp.ResponseBody

End Function

Public Function GetJpgBookParamA(ByRef vCookie As String, ByVal vUrl As String) As String

    vUrl = Replace$(vUrl, ":", "%3A")
    vUrl = Replace$(vUrl, "?", "%3F")
    vUrl = Replace$(vUrl, "&", "%26")
    vUrl = Replace$(vUrl, "=", "%3D")
    vUrl = Replace$(vUrl, "!", "%21")
    
    Dim pHttp As IWinHttp
    Set pHttp = New CWinHttpSimple
    pHttp.Init
    pHttp.URL = "http://img.sslibrary.com/jpgssrurl.jsp?h=" & vUrl
    pHttp.method = "POST"
    pHttp.OpenConnect False
    pHttp.setRequestHeader "Cookie", vCookie
    pHttp.send
    Do Until pHttp.IsFree
        DoEvents
    Loop
    vCookie = HttpHeaderMergeCookie(vCookie, HttpHeaderGetField(pHttp.ResponseHeader, "Set-Cookie"))
    GetJpgBookParamA = StrConv(pHttp.responseBody, vbUnicode)
End Function

Public Function SSLIB_UpdateAllTaskFolders(ByVal vParentFolder As String, ByRef vCookie As String, Optional vForce As Boolean) As Boolean
    Dim folders() As String
    Dim count As Long
    count = MFileSystem.subFolders(vParentFolder, folders())
    Dim i As Long
    For i = 1 To count
        Debug.Print "Processing " & folders(i) & "..."
        SSLIB_UpdateTaskFolder folders(i), vCookie, vForce
    Next
End Function

Public Function SSLIB_UpdateTaskFolder(ByVal vFolder As String, ByRef vCookie As String, Optional vForce As Boolean) As Boolean
    Dim bookInfo As CBookInfo
    Set bookInfo = New CBookInfo
    If vCookie = vbNullString Then vCookie = Clipboard.GetText
    bookInfo.LoadFromFile BuildPath(vFolder, "GetSSLib.ini"), , "TaskInfo"
    SSLIB_UpdateTaskFolder = SSLIB_UpdateTaskOnNeed(bookInfo, vFolder, vCookie, vForce)
End Function

'CSEH: ErrExit
Public Function SSLIB_UpdateTaskOnNeed(ByRef vBookInfo As CBookInfo, ByVal vdirectory As String, ByRef mJpgBookCookie As String, Optional vForceUpdate As Boolean = False, Optional vMainSite As String = cst_sslibrary_main) As Boolean
    '<EhHeader>
    On Error GoTo SSLIB_UpdateTaskOnNeed_Err
    '</EhHeader>
    Dim pBookInfo As CBookInfo
    Set pBookInfo = vBookInfo
    If pBookInfo Is Nothing Then Exit Function
    
    
    Dim pIsJpgBook As Boolean
    pIsJpgBook = pBookInfo(SSF_ISJPGBOOK) <> vbNullString
    
'    If pBookInfo(SSF_SSID) <> vbNULLSTRING And InStr(pBookInfo(SSF_HEADER), "SSCT:") < 1 And pBookInfo(SSF_URL) = vbNULLSTRING Then
'        pIsJpgBook = True
'    ElseIf InStr(pBookInfo(SSF_URL), "ss2jpg.dll") > 1 Then
'        pIsJpgBook = True
'    Else
'        pIsJpgBook = False
'    End If
' vTask.IsJpgBook = pIsJpgBook
    If pIsJpgBook Then
        If InStr(pBookInfo(SSF_URL), "ss2jpg.dll") > 1 Then
            pBookInfo(SSF_JPGURL) = pBookInfo(SSF_URL)
        End If
        
       If vForceUpdate Or pBookInfo(SSF_JPGURL) = vbNullString Or pBookInfo(SSF_HTMLContent) = vbNullString Then
            Dim vNewBookInfo() As String
            
            If pBookInfo(SSF_IEJPGURL) <> vbNullString Then
                vNewBookInfo = IEURL_To_BookInfo(pBookInfo(SSF_IEJPGURL), mJpgBookCookie)
            Else
                Dim d As String
                d = pBookInfo(SSF_PARAMS)
            'If d <> vbNULLSTRING Then d = SubStringBetween(d, "&d=", "&")
                vNewBookInfo = SSID_To_BookInfo(pBookInfo(SSLIBFields.SSF_SSID), mJpgBookCookie, vMainSite, d)
            End If
            If UBound(vNewBookInfo) > 0 Then
            If vNewBookInfo(SSLIBFields.SSF_JPGURL) <> vbNullString Then
                If vNewBookInfo(SSLIBFields.SSF_HTMLContent) = vbNullString Then
                    vNewBookInfo(SSLIBFields.SSF_HTMLContent) = "empty"
                End If
                Debug.Print "Update BookInfo of [" & pBookInfo(SSF_SSID) & "_" & pBookInfo(SSF_Title) & "]"
                    Dim i As Long
                    For i = CST_SSLIB_FIELDS_LBound To CST_SSLIB_FIELDS_UBound
                        If pBookInfo(i) = vbNullString Then pBookInfo(i) = vNewBookInfo(i)
                    Next
            End If
            SSLIB_UpdateTaskOnNeed = True
            End If
        End If
        If pBookInfo(SSF_HTMLContent) <> vbNullString And pBookInfo(SSF_HTMLContent) <> "empty" Then
            Dim folder As String
            Dim Data() As Byte
            folder = BuildPath(vdirectory)
            If FolderExists(folder) = True Then
                If vForceUpdate Or FileExists(folder & "BookContents.htm") = False Then
                    Debug.Print "Update BookcContents.htm of [" & pBookInfo(SSF_SSID) & "_" & pBookInfo(SSF_Title) & "]"
                    Data() = StrConv(pBookInfo(SSF_HTMLContent), vbFromUnicode)
                    FS_WriteFile folder & "BookContents.htm", Data()
                    'SSLIB_UpdateTaskOnNeed = True
                End If
                If vForceUpdate Or FileExists(folder & "Contents.txt") = False Then
                    Debug.Print "Update Content.txt of [" & pBookInfo(SSF_SSID) & "_" & pBookInfo(SSF_Title) & "]"
                    Data() = StrConv(SSLIB_ParseHTMLContent(pBookInfo(SSF_HTMLContent), True), vbFromUnicode)
                    FS_WriteFile folder & "Contents.txt", Data()
                    'SSLIB_UpdateTaskOnNeed = True
                End If
                If vForceUpdate Or FileExists(folder & "BookContents.txt") = False Then
                    Debug.Print "Update BookcContents.txt of [" & pBookInfo(SSF_SSID) & "_" & pBookInfo(SSF_Title) & "]"
                    Data() = StrConv(SSLIB_ParseHTMLContent(pBookInfo(SSF_HTMLContent), False), vbFromUnicode)
                    FS_WriteFile folder & "BookContents.txt", Data()
                End If
                If vForceUpdate Or FileExists(folder & "BookContents.txt") = True And _
                    FileExists(folder & "BookContents.dat") = False Then
                    Debug.Print "Update BookcContents.dat of [" & pBookInfo(SSF_SSID) & "_" & pBookInfo(SSF_Title) & "]"
                    SSLIB_ConvertBookContentFile folder & "BookContents.txt", folder & "BookContents.dat", False
                End If
            End If
        End If
    End If

    '<EhFooter>
    Exit Function

SSLIB_UpdateTaskOnNeed_Err:
    Debug.Print "GetSSLibX.MSSReader.SSLIB_UpdateTaskOnNeed:Error " & Err.Description
    Err.Clear

    '</EhFooter>
End Function

#End If
Public Function ParseJPGBookInfoText(ByVal vText As String) As String()
    
    Dim fNum As Integer
    fNum = FreeFile
    Dim vLine As String
    Dim vTag As String
    Dim fContent As Boolean
    Dim fScript As Boolean
    Dim pContent As String
    Dim pScript As String
    Dim vBookInfo() As String
    vBookInfo = SSLIB_CreateBookInfoArray()
    If vText = vbNullString Then ParseJPGBookInfoText = vBookInfo: Exit Function
    
    Dim i As Long
    Dim u As Long
    Dim vLines() As String
    vLines = Split(vText, vbCrLf)
    u = UBound(vLines)
    For i = 0 To u
        vLine = Trim$(vLines(i))
        If vLine = vbNullString Then GoTo Continue
        vTag = LCase$(SubStringBetween(vLine, "<", ">"))
        If fContent Then
            If vTag = "/div" Then
                fContent = False
            Else
                pContent = pContent & vbCrLf & vLine
            End If
        ElseIf fScript Then
            If vTag = "/script" Then
                fScript = False
                Exit For
            Else
                pScript = pScript & vbCrLf & vLine
            End If
        Else
            If vTag = "h2" Then
                vBookInfo(SSLIBFields.SSF_Title) = SubStringBetween(vLine, "<h2>", "</h2>", True)
            ElseIf vTag = "h5" Then
                vBookInfo(SSLIBFields.SSF_AUTHOR) = SubStringBetween(vLine, "作者：", " ", True)
                vBookInfo(SSLIBFields.SSF_PagesCount) = SubStringBetween(vLine, "页数：", " ", True)
                vBookInfo(SSLIBFields.SSF_PublishDate) = SubStringBetween(vLine, "出版日期：", "<", True)
            ElseIf vTag = "div class=" & Chr$(34) & "book_m" & Chr$(34) & " id=" & Chr$(34) & "div1" & Chr$(34) Then
                fContent = True
            ElseIf Left(vTag, Len("script")) = "script" Then
                fScript = True
            End If
        End If
Continue:
    Next
    'Close #fNum
    vBookInfo(SSLIBFields.SSF_HTMLContent) = pContent
    vBookInfo(SSLIBFields.SSF_SSID) = SubStringBetween(pScript, "ssNo = " & Chr$(34), Chr$(34))
    vBookInfo(SSLIBFields.SSF_JPGURL) = SubStringBetween(pScript, "var str='", "&jid=/")
    If vBookInfo(SSLIBFields.SSF_JPGURL) <> vbNullString Then vBookInfo(SSLIBFields.SSF_JPGURL) = vBookInfo(SSLIBFields.SSF_JPGURL) & "&jid=/"
    ParseJPGBookInfoText = vBookInfo
End Function


Public Function SSLIB_ParseHTMLBookInfo(ByVal vFilename As String) As String()
    If vFilename = vbNullString Then vFilename = "X:\bookinfo.html"
    Dim fNum As Integer
    fNum = FreeFile
    Dim vLine As String
    Dim vTag As String
    Dim fContent As Boolean
    Dim fScript As Boolean
    Dim pContent As String
    Dim pScript As String
    Dim vBookInfo() As String
    vBookInfo = SSLIB_CreateBookInfoArray()
    Open vFilename For Input As #fNum
    Do Until EOF(fNum)
    Line Input #fNum, vLine
        vLine = Trim$(vLine)
        If vLine = vbNullString Then GoTo Continue
        vTag = LCase$(SubStringBetween(vLine, "<", ">"))
        If fContent Then
            If vTag = "/div" Then
                fContent = False
            Else
                pContent = pContent & vbCrLf & vLine
            End If
        ElseIf fScript Then
            If vTag = "/script" Then
                fScript = False
                Exit Do
            Else
                pScript = pScript & vbCrLf & vLine
            End If
        Else
            If vTag = "h2" Then
                vBookInfo(SSLIBFields.SSF_Title) = SubStringBetween(vLine, "<h2>", "</h2>", True)
            ElseIf vTag = "h5" Then
                vBookInfo(SSLIBFields.SSF_AUTHOR) = SubStringBetween(vLine, "作者：", " ", True)
                vBookInfo(SSLIBFields.SSF_PagesCount) = SubStringBetween(vLine, "页数：", " ", True)
                vBookInfo(SSLIBFields.SSF_PublishDate) = SubStringBetween(vLine, "出版日期：", "<", True)
            ElseIf vTag = "div class=" & Chr$(34) & "book_m" & Chr$(34) & " id=" & Chr$(34) & "div1" & Chr$(34) Then
                fContent = True
            ElseIf vTag = "script" Then
                fScript = True
            End If
        End If
Continue:
    Loop
    Close #fNum
    vBookInfo(SSLIBFields.SSF_HTMLContent) = pContent
    vBookInfo(SSLIBFields.SSF_SSID) = SubStringBetween(pScript, "ssNo = " & Chr$(34), Chr$(34))
    vBookInfo(SSLIBFields.SSF_URL) = SubStringBetween(pScript, "var str='", "&jid=/")
    If vBookInfo(SSLIBFields.SSF_URL) <> vbNullString Then vBookInfo(SSLIBFields.SSF_URL) = vBookInfo(SSLIBFields.SSF_URL) & "&jid=/"
    SSLIB_ParseHTMLBookInfo = vBookInfo
End Function

#If Not afNoHtmlObject = 1 Then
'CSEH: ErrExit
    Public Function SSLIB_HTMLContentToBookContent(ByVal vHtmlFile As String, ByVal vBookFile As String, Optional vTabMode As Boolean = False) As Boolean
    '<EhHeader>
    On Error GoTo SSLIB_HTMLContentToBookContent_Err
    '</EhHeader>
        Dim fNum As String
        fNum = FreeFile
        Dim pData() As Byte
        Open vHtmlFile For Binary Access Read As #fNum
        ReDim pData(0 To LOF(fNum) - 1)
        Get #fNum, , pData
        Close #fNum
        fNum = 0
        Dim pHtml As String
        pHtml = StrConv(pData, vbUnicode)
        pHtml = SSLIB_ParseHTMLContent(pHtml, vTabMode)
        If pHtml <> vbNullString Then
            fNum = FreeFile
            Open vBookFile For Output As #fNum
            Print #fNum, pHtml
            Close #fNum
            fNum = 0
        End If
        SSLIB_HTMLContentToBookContent = True
    '<EhFooter>
    Exit Function

SSLIB_HTMLContentToBookContent_Err:
    SSLIB_HTMLContentToBookContent = False
    If fNum <> 0 Then Close #fNum
    'Err.Clear

    '</EhFooter>
    End Function
    Public Function SSLIB_ParseHTMLContent(ByVal vText As String, Optional vTabMode As Boolean = False) As String
        If vText = vbNullString Then vText = Clipboard.GetText
        If vText = vbNullString Then Exit Function
        vText = SSLIB_HtmlContentCleanSource(vText)
        Dim pDoc As HTMLDocument
        Set pDoc = New HTMLDocument
        Dim pElm As HTMLDivElement
        Set pElm = pDoc.createElement("Div")
        pElm.innerHTML = vText
        Dim pContent As String
        pContent = SSLIB_HtmlContentRootOutput(pElm, "05", vTabMode)
        If pContent <> vbNullString Then pContent = pContent & vbCrLf
        If vTabMode = False Then
            If pContent <> vbNullString Then pContent = pContent & vbCrLf
            SSLIB_ParseHTMLContent = _
                "封面|00|1|0|1|" & vbCrLf & _
                "书名|01|1|0|2|" & vbCrLf & _
                "版权|02|1|0|3|" & vbCrLf & _
                "前言|03|1|0|4|" & vbCrLf & _
                "目录|04|1|0|5|" & vbCrLf & _
                pContent & _
                "索引|06|1|0|7|" & vbCrLf & _
                "附录|07|1|0|8|" & vbCrLf & _
                "封底|08|1|0|9|" & vbCrLf
        Else
             If pContent <> vbNullString Then pContent = vbTab & pContent & vbCrLf
            SSLIB_ParseHTMLContent = _
                "封面|cov1" & vbCrLf & _
                "书名|bok1" & vbCrLf & _
                "版权|leg1" & vbCrLf & _
                "前言|fow1" & vbCrLf & _
                "目录|!1|" & vbCrLf & _
                pContent & _
                "索引|ins1" & vbCrLf & _
                "附录|att1" & vbCrLf & _
                "封底|bac1" & vbCrLf
        End If
       ' Debug.Print SSLIB_ParseHTMLContent
    End Function
    

    
    Public Function SSLIB_HTMLContentEntryOutput(ByRef vEntry As IHTMLElement, ByVal vLevel As String, Optional vTabMode As Boolean = False)
        Dim pPage As String
        Dim pText As String
        Dim pType As Integer
        pText = Trim$((vEntry.innerText))
        If pText = vbNullString Then Exit Function
        If pText = vbLf Then Exit Function
        If pText = vbCrLf Then Exit Function
        On Error Resume Next
        pPage = SubStringBetween(vEntry.getAttribute("onclick"), "Page(", ")", True)
        If pPage = vbNullString Then Exit Function
        
        Dim pTypeStr As String
        pType = CInt(SubStringBetween(pPage, ",", vbNullString))
        If vTabMode Then
            Select Case pType
                Case 0
                Case 1
                Case 2
                Case 3
                Case 4
                    pTypeStr = vbNullString
            End Select
        Else
            pTypeStr = "6"
            Select Case pType
                Case 0
                Case 1
                Case 2
                Case 3
                Case 4
                    pTypeStr = "6"
            End Select
        End If
        pPage = SubStringUntilMatch(pPage, 1, ",")
        If vTabMode Then
            SSLIB_HTMLContentEntryOutput = Mid$(vLevel, 2) & pText & vbTab & pTypeStr & pPage
        Else
            SSLIB_HTMLContentEntryOutput = pText & "|" & vLevel & "|" & pPage & "|0|" & pTypeStr & "|"
        End If
        'Debug.Print ":::" & SSLIB_HTMLContentEntryOutput
    End Function
    Public Function SSLIB_HtmlContentRootOutput(ByRef vRoot As IHTMLElement, ByVal vLevel As String, Optional vTabMode As Boolean = False)
        Dim pElm As IHTMLElement
        Dim ret As String
        Dim pCount As Integer
        For Each pElm In vRoot.children
            If pElm.id = "limore" Then
                ret = SSLIB_HtmlContentRootOutput(pElm, vLevel, vTabMode)
            ElseIf pElm.children.length > 1 Then
                If vTabMode Then
                    ret = SSLIB_HtmlContentRootOutput(pElm, vLevel & vbTab, True)
                Else
                    ret = SSLIB_HtmlContentRootOutput(pElm, vLevel & StrNum(pCount, 2, True), False)
                End If
            Else
                ret = SSLIB_HTMLContentEntryOutput(pElm, vLevel, vTabMode)
            End If
            If ret <> vbNullString Then
                'Debug.Print ":::" & ret
                SSLIB_HtmlContentRootOutput = SSLIB_HtmlContentRootOutput & vbCrLf & ret
                pCount = pCount + 1
            End If
        Next
        Do While Left$(SSLIB_HtmlContentRootOutput, 2) = vbCrLf
            SSLIB_HtmlContentRootOutput = Mid$(SSLIB_HtmlContentRootOutput, 3)
        Loop
    End Function
    Public Function SSLIB_HtmlContentCleanSource(ByRef vText As String) As String
        Dim vRet As String
        vRet = vText
        Dim pStart As Long
        Dim pEnd As Long
        Dim pNext As Long
        Do
            pStart = InStr(1, vRet, "<img", vbTextCompare)
            If pStart > 0 Then
                pEnd = InStr(pStart + 1, vRet, "/>", vbBinaryCompare)
                If pEnd > 0 Then
                    vRet = Left$(vRet, pStart - 1) + Mid$(vRet, pEnd + 2)
                End If
            End If
        Loop While pStart > 0
        SSLIB_HtmlContentCleanSource = Replace$(vRet, "<", vbCrLf & "<")
        SSLIB_HtmlContentCleanSource = Replace$(SSLIB_HtmlContentCleanSource, "_Main.GoTo", "Page")
    End Function

#End If

#If Not afNoZlib = 1 Then
    Public Function SSLIB_DecodeBookContents(ByVal vFilename As String) As String
        Dim vSource() As Byte
        If FS_ReadFile(vFilename, vSource(), &H29) > 0 Then
            SSLIB_DecodeBookContents = Zlib_UncompressToString(vSource())
        End If
    End Function
    Public Function SSLIB_EncodeBookContents(ByVal vFilename As String, ByRef vResult As String) As Long
        Dim fHead As String * 40
        'fHead = String$("_", 40)
        Mid$(fHead, 1, 6) = "SSCNTS"
        Mid$(fHead, 13) = "created by xiaoranzzz"
        fHead = StrConv(fHead, vbFromUnicode)
        
        Dim fHeadByte() As Byte
        fHeadByte = fHead
        ReDim Preserve fHeadByte(0 To 39)
        
        Dim vContent() As Byte
        vResult = fHeadByte
        vResult = vResult & Zlib_CompressFile(vFilename)
        
        'vResult = fHeadByte & vContent
        SSLIB_EncodeBookContents = Len(vResult)
    End Function
    Public Function SSLIB_ConvertBookContentFile(ByVal vFromFile As String, ByVal vToFile As String, Optional vDecodeMode As Boolean = True) As Boolean
        Dim result As String
        Dim resultArray() As Byte
        If vDecodeMode = True Then
            resultArray = StrConv(SSLIB_DecodeBookContents(vFromFile), vbFromUnicode)
        Else
            SSLIB_EncodeBookContents vFromFile, result
            resultArray = result
        End If
        SSLIB_ConvertBookContentFile = FS_WriteFile(vToFile, resultArray)
    End Function
#End If
