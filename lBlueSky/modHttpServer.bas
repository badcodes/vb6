Attribute VB_Name = "modHttpServer"
Option Explicit
Public Type HttpServerSet 'SectionName "[HttpServer]"
    sName As String
    sVersion As String
    sIP As String
    sPort As String
    sHostName As String
    sTemplateFile As String
    bUseTemplate As Boolean
    sRootPath As String
End Type

Public Enum HSUrlType
    nullHSUrl = 0
    zipHSUrl = 1
    fileHSUrl = 2
    folderHSUrl = 3
End Enum
Public Enum HSFileType
    hsFTJustLink = 1
    hsFTElse = 2
End Enum
Public Type HSUrl
    mainPart As String
    secondPart As String
    urlType As HSUrlType
    IsFakeHtml As Boolean
    fileType As HSFileType
End Type

Public Const sZipUrlSep As String = "|"
Public Const sFakeHtmlTrail As String = ".zhFake.Html"
Public Const sTempFileTrail As String = ".$zhTemp$"
Public Sub hs_getServerSetting(sIniFilename As String, hssSaveTo As HttpServerSet)

With hssSaveTo
    .sName = iniGetSetting(sIniFilename, "HttpServer", "Name")
    .sHostName = iniGetSetting(sIniFilename, "HttpServer", "HostName")
    .sVersion = iniGetSetting(sIniFilename, "HttpServer", "Version")
    .sIP = iniGetSetting(sIniFilename, "HttpServer", "IP")
    .sPort = iniGetSetting(sIniFilename, "HttpServer", "Port")
    .sTemplateFile = iniGetSetting(sIniFilename, "ViewerStyle", "TemplateFile")
    .bUseTemplate = CBoolStr(iniGetSetting(sIniFilename, "ViewerStyle", "UseTemplate"))
    .sRootPath = iniGetSetting(sIniFilename, "HttpServer", "RootPath")
End With
    

End Sub
Public Sub hs_saveServerSetting(sIniFilename As String, hssToSave As HttpServerSet)

With hssToSave
    iniSaveSetting sIniFilename, "HttpServer", "Name", .sName
    iniSaveSetting sIniFilename, "HttpServer", "HostName", .sHostName
    iniSaveSetting sIniFilename, "HttpServer", "Version", .sVersion
    iniSaveSetting sIniFilename, "HttpServer", "IP", .sIP
    iniSaveSetting sIniFilename, "HttpServer", "Port", .sPort
    iniSaveSetting sIniFilename, "ViewerStyle", "TemplateFile", .sTemplateFile
    iniSaveSetting sIniFilename, "ViewerStyle", "UseTemplate", .bUseTemplate
    iniSaveSetting sIniFilename, "HttpServer", "RootPath", .sRootPath
End With

End Sub
Public Function hs_ParseUrl(ByVal sUrl As String) As HSUrl

    'sUrl = DecodeUrl(sUrl, CP_UTF8)
    sUrl = Replace(Replace(sUrl, "/", "\"), "\\", "\")
    sUrl = RTrim(sUrl)
    sUrl = RightDelete(sUrl, Chr(0))
    'sUrl = RightDelete(sUrl, "\")
    sUrl = LeftDelete(sUrl, "\")
    
    Dim fso As New FileSystemObject
  
    With hs_ParseUrl
        .IsFakeHtml = hs_isFakeHtmlUrl(sUrl, sUrl)
        .mainPart = LeftLeft(sUrl, sZipUrlSep, vbBinaryCompare, ReturnOriginalStr)
        .secondPart = LeftDelete(LeftRight(sUrl, sZipUrlSep, vbBinaryCompare, ReturnEmptyStr), "\")
        If fso.FolderExists(.mainPart) Then
            .urlType = folderHSUrl
        ElseIf .secondPart <> "" Then
            .urlType = zipHSUrl
        ElseIf .mainPart <> "" Then
            .urlType = fileHSUrl
        Else
            .urlType = nullHSUrl
        End If
        
        Dim chkUrlType As mFile_FileType
        chkUrlType = chkFileType(sUrl)
        .fileType = hsFTElse
        If chkUrlType = ftAUDIO Or chkUrlType = ftIMG Or chkUrlType = ftVIDEO Then
        .fileType = hsFTJustLink
        End If
   
        
    End With
    
End Function

Public Function hs_isFakeHtmlUrl(ByVal sUrl As String, ByRef sRealUrl As String) As Boolean
    If Right(sUrl, Len(sFakeHtmlTrail)) = sFakeHtmlTrail Then
        sRealUrl = RightDelete(sUrl, sFakeHtmlTrail)
        hs_isFakeHtmlUrl = True
    Else
        sRealUrl = sUrl
        hs_isFakeHtmlUrl = False
    End If
End Function

Public Function hs_isTempFile(ByVal sFilename As String) As Boolean
    If Right(sFilename, Len(sTempFileTrail)) = sTempFileTrail Then
        hs_isTempFile = True
    Else
        hs_isTempFile = False
    End If
End Function

Public Function hs_getTempFileName(Optional sTempName As String = "") As String

    If sTempName = "" Then sTempName = "$kjlfieu$.temp"
    hs_getTempFileName = sTempName & sTempFileTrail
    
End Function


Public Function hs_DecodeUrl(ByVal sUrl As String) As String

Dim sUtf8Url As String
Dim sESCUrl As String
Dim lOUrl As Long
Dim lUUrl As Long
Dim lEUrl As Long
Dim minLen As Long

sUtf8Url = DecodeUrl(sUrl, CP_UTF8)

lOUrl = Len(sUrl)
lUUrl = Len(sUtf8Url)

minLen = lOUrl - charCountInStr(sUrl, "%") * 8 / 3

If lUUrl >= minLen Then
    hs_DecodeUrl = sUtf8Url
Else
    hs_DecodeUrl = DecodeUrl(sUrl, 0)
End If

End Function

Public Function hs_CreateIndex(ByVal sPath As String, ByVal sHttpServerHead As String, ByVal sIdxFilename As String) As Boolean
        
        Dim fso As New FileSystemObject
        Dim ts As TextStream
        Dim fReal As String
        Dim tsContent As String
        Dim sHref As String
        Dim fsof As File
        Dim fsofs As Files
        Dim fsofd As Folder
        Dim fsofds As Folders
        Dim stmp As String
        
        If fso.FolderExists(sPath) = False Then Exit Function
        
        fReal = sIdxFilename
        If fso.FileExists(fReal) Then fso.DeleteFile fReal, True
        Set ts = fso.OpenTextFile(fReal, ForWriting, True)
        tsContent = "<table width=!100%! border=0 >"
        tsContent = tsContent & "<tr><td align=!center!>"
        tsContent = tsContent & "<table><tr><td style=!line-height: 150%!>"
        
        stmp = fso.GetParentFolderName(sPath)
        If stmp <> "" Then
        sHref = toUnixPath(stmp)
             tsContent = tsContent & "&gt;&gt;&nbsp;<a href=!" & _
                sHttpServerHead & sHref & "!>..</a>" & vbCrLf
        End If
        
        Set fsofds = fso.GetFolder(sPath).SubFolders
        For Each fsofd In fsofds
        sHref = toUnixPath(fsofd.Path)
             tsContent = tsContent & "&gt;&gt;&nbsp;<a href=!" & _
                sHttpServerHead & sHref & "!>" & fsofd.Name & "</a>" & vbCrLf
        Next
                        
        Dim cft As mFile_FileType
        Set fsofs = fso.GetFolder(sPath).Files
        For Each fsof In fsofs
        sHref = toUnixPath(fsof.Path)
        cft = chkFileType(sHref)
        If cft = ftIE Or cft = ftZIP Or cft = ftZhtm Then
        Else
        sHref = sHref & sFakeHtmlTrail
        End If
            tsContent = tsContent & "&gt;&gt;&nbsp;<a href=!" & _
                sHttpServerHead & sHref & "!>" & fsof.Name & "</a>" & vbCrLf
        Next
        
        tsContent = tsContent & "</td></tr></table></td></tr></table>"
        tsContent = Replace(tsContent, "!", Chr$(34))
        ts.Write tsContent
        ts.Close
        
        hs_CreateIndex = True
        
End Function

Public Function hs_createHtmlFromTemplate(sSourcePath As String, sTemplate As String, sHtmlPath As String) As Boolean

    Const htmlTemplateDir = "#####TEMPLATEDIR#####"
    Const htmlTitle = "#####TITLE#####"
    Const htmlContent = "#####CONTENT#####"
    Const htmlLinkPrevious = "#####LINKPREVIOUS#####"
    Const htmlLinkNext = "#####LINKNEXT#####"
    Const htmlLinkIndex = "#####LINKINDEX#####"
    Const htmlLinkScript = "#####LINKSCRIPT#####"
    Const htmlLinkStyle = "#####LINKSTYLE#####"
    Const htmlHrefPrevious = "#####HREFPREVIOUS#####"
    Const htmlHrefNext = "#####HREFNEXT#####"
    Const htmlHrefIndex = "#####HREFINDEX#####"
    Const hrefPrevious = "zhCmd://mnuGo_previous/"
    Const hrefNext = "zhCmd://mnuGo_next/"
    Const hrefIndex = "zhCmd://mnuGo_home/"
    Const LinkPrevious = "<a href='" & hrefPrevious & "' >Previous</a>"
    Const LinkNext = "<a href='" & hrefNext & "' >Next</a>"
    Const LinkIndex = "<a href='" & hrefIndex & "' >Index</a>"
    Dim fso As New FileSystemObject
    Dim textTs As TextStream
    Dim htmlSourceTs As TextStream
    Dim htmlToWriteTS As TextStream
    Dim stmp As String
    'Dim posStart As String
    'Dim posEnd As String
    Dim htmlSplit() As String
    Dim sTitle As String
    hs_createHtmlFromTemplate = False

    If fso.FileExists(sTemplate) = False Then Exit Function
    sTitle = fso.GetBaseName(sSourcePath)
    Set htmlSourceTs = fso.OpenTextFile(sTemplate, ForReading)
    stmp = htmlSourceTs.ReadAll
    htmlSourceTs.Close
    stmp = Replace(stmp, htmlTitle, sTitle)
    stmp = Replace(stmp, htmlLinkPrevious, "") ' LinkPrevious)
    stmp = Replace(stmp, htmlLinkNext, "") 'LinkNext)
    stmp = Replace(stmp, htmlLinkIndex, "") ' LinkIndex)
    stmp = Replace(stmp, htmlHrefPrevious, "#") ' hrefPrevious)
    stmp = Replace(stmp, htmlHrefNext, "#") 'hrefNext)
    stmp = Replace(stmp, htmlHrefIndex, "#") 'hrefIndex)
    stmp = Replace(stmp, htmlTemplateDir, fso.GetParentFolderName(sTemplate))
    htmlSplit = Split(stmp, htmlContent)

    If UBound(htmlSplit) <> 1 Then Exit Function
    Set htmlToWriteTS = fso.OpenTextFile(sHtmlPath, ForWriting, True)

    With htmlToWriteTS
        .Write htmlSplit(0)

        Select Case chkFileType(sSourcePath)
            Case ftIMG
                .WriteLine "<center>"
                .WriteLine "<TABLE cellSpacing=0 cellPadding=0 >"
                .WriteLine "<TR ><TD  align=center>"
                .WriteLine "<img src=" + Chr$(34) + sSourcePath + Chr$(34) + ">"
                .WriteLine "</td></tr></table>"
            Case ftAUDIO, ftVIDEO
                .WriteLine "<center>"
                .WriteLine "<TABLE  width=" + Chr$(34) + "96%" + Chr$(34) + " cellSpacing=4 cellPadding=4 width=" + Chr$(34) + "100%" + Chr$(34) + ">"
                .WriteLine "<TR ><TD  align=center class=m_text>"
                .WriteLine "<object classid=" + Chr$(34) + "Mid:22D6F312-B0F6-11D0-94AB-0080C74C7E95" + Chr$(34) + " id=" + Chr$(34) + "MediaPlayer1" + Chr$(34) + ">"
                .WriteLine "<param name=" + Chr$(34) + "Filename" + Chr$(34) + " value=" + Chr$(34) + sSourcePath + Chr$(34) + ">"
                .WriteLine "</object>"
                .WriteLine "</td></tr></table>"
            Case Else
                If fso.FileExists(sSourcePath) = False Then Exit Function
                Set textTs = fso.OpenTextFile(sSourcePath, ForReading)

                Do Until textTs.AtEndOfStream
                    .WriteLine textTs.ReadLine & "<br>"
                Loop

                textTs.Close

        End Select

        .Write htmlSplit(1)
        .Close
    End With

    hs_createHtmlFromTemplate = True
End Function

Public Function hs_createDefaultHtml(sSource As String, sHtmlPath As String) As Boolean
    Dim stmp As String
    Dim sTmpTemplate As String
    Dim fso As New FileSystemObject
    Dim ts As TextStream
    stmp = stmp & "<html>" & vbCrLf
    stmp = stmp & "<head>" & vbCrLf
    stmp = stmp & "<meta http-equiv=" & Chr$(34) & "Content-Type" & Chr$(34) & " content=" & Chr$(34) & "text/html; charset=gb2312" & Chr$(34) & ">" & vbCrLf
    stmp = stmp & "<title>" & vbCrLf
    stmp = stmp & "#####TITLE#####" & vbCrLf
    stmp = stmp & "</title>" & vbCrLf
    stmp = stmp & "<link REL=" & Chr$(34) & "stylesheet" & Chr$(34) & " href=" & Chr$(34) & "#####TEMPLATEDIR#####/style.css" & Chr$(34) & " type=" & Chr$(34) & "text/css" & Chr$(34) & ">" & vbCrLf
    stmp = stmp & "<script src=" & Chr$(34) & "#####TEMPLATEDIR#####/script.js" & Chr$(34) & "></script>" & vbCrLf
    stmp = stmp & "</head>" & vbCrLf
    stmp = stmp & "<body>" & vbCrLf
    stmp = stmp & "<div align=" & Chr$(34) & "center" & Chr$(34) & ">" & vbCrLf
    stmp = stmp & "<center>" & vbCrLf
    stmp = stmp & "<table border=" & Chr$(34) & "0" & Chr$(34) & " cellpadding=" & Chr$(34) & "0" & Chr$(34) & " cellspacing=" & Chr$(34) & "0" & Chr$(34) & " width=" & Chr$(34) & "90%" & Chr$(34) & ">" & vbCrLf
    stmp = stmp & "<tr> " & vbCrLf
    stmp = stmp & "<td valign=" & Chr$(34) & "top" & Chr$(34) & "><DIV class=" & Chr$(34) & "m_text" & Chr$(34) & " align=" & Chr$(34) & "left" & Chr$(34) & "> " & vbCrLf
    stmp = stmp & "#####CONTENT#####" & vbCrLf
    stmp = stmp & "</DIV></td>" & vbCrLf
    stmp = stmp & "</tr>" & vbCrLf
    stmp = stmp & "</table>" & vbCrLf
    stmp = stmp & "</center>" & vbCrLf
    stmp = stmp & "</div>" & vbCrLf
    stmp = stmp & "</body>" & vbCrLf
    stmp = stmp & "</html>" & vbCrLf
    sTmpTemplate = fso.BuildPath(App.Path, fso.GetTempName)
    Set ts = fso.CreateTextFile(sTmpTemplate, True)
    ts.Write stmp
    ts.Close
    hs_createDefaultHtml = hs_createHtmlFromTemplate(sSource, sTmpTemplate, sHtmlPath)
    fso.DeleteFile sTmpTemplate
End Function



