Attribute VB_Name = "MZhtmTemplate"
Option Explicit

Public Function createHtmlFromTemplate(sSourcePath As String, sTemplate As String, sHtmlPath As String) As Boolean

    Const htmlTemplateDir = "#####TEMPLATEDIR#####"
    Const htmlTitle = "#####TITLE#####"
    Const htmlContent = "#####CONTENT#####"
    Const htmlLinkPrevious = "#####LINKPREVIOUS#####"
    Const htmlLinkNext = "#####LINKNEXT#####"
    Const htmlLinkIndex = "#####LINKINDEX#####"
    Const htmlHrefPrevious = "#####HREFPREVIOUS#####"
    Const htmlHrefNext = "#####HREFNEXT#####"
    Const htmlHrefIndex = "#####HREFINDEX#####"
    
    Const tmplTitle = "[PART1]"
    Const tmplContent = "[PART0]"
    Const tmplLinkPrevious = "[GOPREV]"
    Const tmplLinkNext = "[GONEXT]"
    Const tmplLinkIndex = "[GOINDEX]"
    Const tmplHrefPrevious = "[PREVPAGE]"
    Const tmplHrefNext = "[NEXTPAGE]"
    Const tmplHrefIndex = "[INDEXPAGE]"
    Const tmplPageType = "[PAGETYPE]"
    
    Const PageType = "page"
    Const hrefPrevious = "zhCmd://mnuGo_previous_Click"
    Const hrefNext = "zhCmd://mnuGo_next_Click"
    Const hrefIndex = "zhCmd://mnuGo_home_Click"
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
    createHtmlFromTemplate = False

    If fso.FileExists(sTemplate) = False Then Exit Function
    sTitle = fso.GetBaseName(sSourcePath)
    Set htmlSourceTs = fso.OpenTextFile(sTemplate, ForReading)
    stmp = htmlSourceTs.ReadAll
    htmlSourceTs.Close
    
    stmp = Replace(stmp, htmlTitle, sTitle)
    stmp = Replace(stmp, htmlLinkPrevious, LinkPrevious)
    stmp = Replace(stmp, htmlLinkNext, LinkNext)
    stmp = Replace(stmp, htmlLinkIndex, LinkIndex)
    stmp = Replace(stmp, htmlHrefPrevious, hrefPrevious)
    stmp = Replace(stmp, htmlHrefNext, hrefNext)
    stmp = Replace(stmp, htmlHrefIndex, hrefIndex)
    stmp = Replace(stmp, htmlTemplateDir, fso.GetParentFolderName(sTemplate))
    
    
    stmp = Replace(stmp, tmplTitle, sTitle)
    stmp = Replace(stmp, tmplLinkPrevious, LinkPrevious)
    stmp = Replace(stmp, tmplLinkNext, LinkNext)
    stmp = Replace(stmp, tmplLinkIndex, LinkIndex)
    stmp = Replace(stmp, tmplHrefPrevious, hrefPrevious)
    stmp = Replace(stmp, tmplHrefNext, hrefNext)
    stmp = Replace(stmp, tmplHrefIndex, hrefIndex)
    stmp = Replace(stmp, tmplPageType, PageType)
    
    
    htmlSplit = Split(stmp, htmlContent)

    If UBound(htmlSplit) <> 1 Then htmlSplit = Split(stmp, tmplContent)
    If UBound(htmlSplit) <> 1 Then Exit Function
    
    Set htmlToWriteTS = fso.OpenTextFile(sHtmlPath, ForWriting, True)

    With htmlToWriteTS
        .Write "<base url=" & Chr$(34) & fso.GetParentFolderName(sTemplate) & Chr$(34) & " >"
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
            .WriteLine "<object classid=" + Chr$(34) + "clsid:22D6F312-B0F6-11D0-94AB-0080C74C7E95" + Chr$(34) + " id=" + Chr$(34) + "MediaPlayer1" + Chr$(34) + ">"
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

    createHtmlFromTemplate = True

End Function

Public Function createDefaultHtml(sSource As String, sHtmlPath As String) As Boolean

    Dim stmp As String
    Dim sTmpTemplate As String
    Dim fso As New FileSystemObject
    Dim ts As TextStream
    stmp = stmp & "<html>"
    stmp = stmp & "<head>"
    stmp = stmp & "<meta http-equiv=" & Chr$(34) & "Content-Type" & Chr$(34) & " content=" & Chr$(34) & "text/html; charset=gb2312" & Chr$(34) & ">"
    stmp = stmp & "<title>"
    stmp = stmp & "#####TITLE#####"
    stmp = stmp & "</title>"
    stmp = stmp & "<link REL=" & Chr$(34) & "stylesheet" & Chr$(34) & " href=" & Chr$(34) & "#####TEMPLATEDIR#####/style.css" & Chr$(34) & " type=" & Chr$(34) & "text/css" & Chr$(34) & ">"
    stmp = stmp & "<script src=" & Chr$(34) & "#####TEMPLATEDIR#####/script.js" & Chr$(34) & "></script>"
    stmp = stmp & "</head>"
    stmp = stmp & "<body>"
    stmp = stmp & "<div align=" & Chr$(34) & "center" & Chr$(34) & ">"
    stmp = stmp & "<center>"
    stmp = stmp & "<table border=" & Chr$(34) & "0" & Chr$(34) & " cellpadding=" & Chr$(34) & "0" & Chr$(34) & " cellspacing=" & Chr$(34) & "0" & Chr$(34) & " width=" & Chr$(34) & "100%" & Chr$(34) & ">"
    stmp = stmp & "<tr> "
    stmp = stmp & "<td colspan=3 valign=" & Chr$(34) & "top" & Chr$(34) & "><DIV class=" & Chr$(34) & "m_text" & Chr$(34) & " align=" & Chr$(34) & "left" & Chr$(34) & "> "
    stmp = stmp & "#####CONTENT#####"
    stmp = stmp & "</DIV></td>"
    stmp = stmp & "</tr>"
    stmp = stmp & "<tr><td colspan=3><hr width=60%></td></tr>"
    stmp = stmp & "<tr><Td align=right width='33%'>[GOPREV]</TD><TD width='33%' align=middle>[GOINDEX]</TD><TD align=left width='33%'>[GONEXT]</TD></TR>"
    stmp = stmp & "</table>"
    stmp = stmp & "</center>"
    stmp = stmp & "</div>"
    stmp = stmp & "</body>"
    stmp = stmp & "</html>"
    sTmpTemplate = fso.BuildPath(App.Path, fso.GetTempName)
    Set ts = fso.CreateTextFile(sTmpTemplate, True)
    ts.Write stmp
    ts.Close
    createDefaultHtml = createHtmlFromTemplate(sSource, sTmpTemplate, sHtmlPath)
    fso.DeleteFile sTmpTemplate

End Function

Public Function IndexFromFileList(sFList() As String, sFileOut As String) As Boolean

Dim fso As New clsFileSystem
Dim fList(26) As String
Dim fdList() As String
Dim fdCount As Long
Dim sTmpChar As String
Dim lAsc As Long
Dim l As Long
Dim lStart As Long
Dim lEnd As Long
Dim sTmpFile As String


lStart = LBound(sFList())
lEnd = UBound(sFList())

On Error GoTo Herr

For l = lStart To lEnd
    If Right$(sFList(l), 1) = "\" Then
        fdCount = fdCount + 1
        ReDim Preserve fdList(fdCount) As String
        fdList(fdCount) = sFList(l)
    Else
        sTmpChar = LCase(Left(fso.GetBaseName(sFList(l)), 1))
        lAsc = Asc(sTmpChar)
        If lAsc < 97 Or lAsc > 122 Then
            sTmpChar = ToPY(lAsc)
            If sTmpChar = "" Then
                lAsc = 96
            Else
                lAsc = LCase(Asc(Left(sTmpChar, 1)))
            End If
        End If
        fList(lAsc - 96) = fList(lAsc - 96) & Chr(0) & sFList(l)
    End If
Next

Dim fnum As Long
fnum = FreeFile

Open sFileOut For Output As #fnum

Dim sArr() As String
Dim lCount As Long
Dim lPart As Long
Dim lRest As Long
Dim j As Long
Dim K As Long
Dim fN As String

Print #fnum, "<table width='100%' border='1' >";

    If fdCount > 0 Then
Print #fnum, "<tr><td colspan='3' class='sTitle' bgcolor='#CCCCCC' align='center' > <b>文件夹</b></td></tr>";
            lPart = fdCount \ 3
            lRest = fdCount Mod 3
            For j = 1 To lPart
Print #fnum, "<tr>";
                For K = 1 To 3
                fN = fdList((j - 1) * 3 + K)
Print #fnum, "<td class='sContent' width='33%' align='center' ><a href='" & fN & "'>" & fso.GetBaseName(fN) & "</a></td>";
                Next
Print #fnum, "</tr>";
            Next
            If lRest > 0 Then
Print #fnum, "<tr>";
                For j = 1 To lRest
                fN = fdList(lPart * 3 + j)
Print #fnum, "<td class='sContent' width='33%' align='center' ><a  href='" & fN & "'>" & fso.GetBaseName(fN) & "</a></td>";
                Next
                For j = lRest + 1 To 3
Print #fnum, "<td class='sContent' width='33%' align='center'> </td>";
                Next
Print #fnum, "</tr>";
            End If
End If

    For l = 0 To 26
    sArr = Split(fList(l), Chr(0))
    lCount = UBound(sArr)
        If lCount > 0 Then
Print #fnum, "<tr><td colspan='3' class='sTitle' bgcolor='#CCCCCC' align='center' > <b>" & Chr(64 + l) & "</b></td></tr>";
            lPart = lCount \ 3
            lRest = lCount Mod 3
            For j = 1 To lPart
Print #fnum, "<tr>";
                For K = 1 To 3
                fN = sArr((j - 1) * 3 + K)
Print #fnum, "<td class='sContent' width='33%' align='center' ><a title='" & fso.GetExtensionName(fN) & " 文件'  href='" & fN & "'>" & fso.GetBaseName(fN) & "</a></td>";
                Next
Print #fnum, "</tr>";
            Next
            If lRest > 0 Then
Print #fnum, "<tr>";
                For j = 1 To lRest
                fN = sArr(lPart * 3 + j)
Print #fnum, "<td class='sContent' width='33%' align='center' ><a title='" & fso.GetExtensionName(fN) & " 文件'  href='" & fN & "'>" & fso.GetBaseName(fN) & "</a></td>";
                Next
                For j = lRest + 1 To 3
Print #fnum, "<td class='sContent' width='33%' align='center'> </td>";
                Next
Print #fnum, "</tr>";
            End If
        End If
    Next

Print #fnum, "</table>"

Close #fnum

IndexFromFileList = True
Exit Function

Herr:

End Function
Public Function DefaultImgTemplate(sSource As String, sHtmlPath As String) As Boolean

    Dim stmp As String
    Dim sTmpTemplate As String
    Dim fso As New FileSystemObject
    Dim ts As TextStream
    stmp = stmp & "<html>"
    stmp = stmp & "<head>"
    stmp = stmp & "<meta http-equiv=" & Chr$(34) & "Content-Type" & Chr$(34) & " content=" & Chr$(34) & "text/html; charset=gb2312" & Chr$(34) & ">"
    stmp = stmp & "<title>"
    stmp = stmp & "#####TITLE#####"
    stmp = stmp & "</title>"
    stmp = stmp & "<link REL=" & Chr$(34) & "stylesheet" & Chr$(34) & " href=" & Chr$(34) & "#####TEMPLATEDIR#####/style.css" & Chr$(34) & " type=" & Chr$(34) & "text/css" & Chr$(34) & ">"
    stmp = stmp & "<script src=" & Chr$(34) & "#####TEMPLATEDIR#####/script.js" & Chr$(34) & "></script>"
    stmp = stmp & "</head>"
    stmp = stmp & "<body>"
    stmp = stmp & "<div align=" & Chr$(34) & "center" & Chr$(34) & ">"
    stmp = stmp & "<center>"
    stmp = stmp & "<table border=" & Chr$(34) & "0" & Chr$(34) & " cellpadding=" & Chr$(34) & "0" & Chr$(34) & " cellspacing=" & Chr$(34) & "0" & Chr$(34) & " width=" & Chr$(34) & "100%" & Chr$(34) & ">"
    stmp = stmp & "<tr> "
    stmp = stmp & "<td colspan=3 valign=" & Chr$(34) & "top" & Chr$(34) & "><DIV class=" & Chr$(34) & "m_text" & Chr$(34) & " align=" & Chr$(34) & "left" & Chr$(34) & "> "
    stmp = stmp & "#####CONTENT#####"
    stmp = stmp & "</DIV></td>"
    stmp = stmp & "</tr>"
    stmp = stmp & "<tr><td colspan=3><hr width=60%></td></tr>"
    stmp = stmp & "<tr><Td align=right width='33%'>[GOPREV]</TD><TD width='33%' align=middle>[GOINDEX]</TD><TD align=left width='33%'>[GONEXT]</TD></TR>"
    stmp = stmp & "</table>"
    stmp = stmp & "</center>"
    stmp = stmp & "</div>"
    stmp = stmp & "</body>"
    stmp = stmp & "</html>"
    sTmpTemplate = fso.BuildPath(App.Path, fso.GetTempName)
    Set ts = fso.CreateTextFile(sTmpTemplate, True)
    ts.Write stmp
    ts.Close
    createDefaultHtml = createHtmlFromTemplate(sSource, sTmpTemplate, sHtmlPath)
    fso.DeleteFile sTmpTemplate

End Function
