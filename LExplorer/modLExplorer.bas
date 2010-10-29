Attribute VB_Name = "modLExplorer"

Option Explicit
Type MYPoS
    Top As Long
    Left As Long
    Height As Long
    Width As Long
End Type


Type ReaderStyle
    WindowState As FormWindowStateConstants
    formPos As MYPoS
    LeftWidth As Long
    LastPath As String
    ShowMenu As Boolean
    ShowLeft As Boolean
    ShowStatusBar As Boolean
    FullScreenMode As Boolean
    TextEditor As String
End Type

Type ViewerStyle
    Viewfont As MYFont
    ForeColor As OLE_COLOR
    BackColor As OLE_COLOR
    LineHeight As Integer
    UseTemplate As Boolean
    TemplateFile As String
    RecentMax As Integer
End Type

Public Type zhReaderStatus
 sCur_zhFile As String
 sCur_zhSubFile As String
 bMenuShowed As Boolean
 bLeftShowed As Boolean
 bStatusBarShowed As Boolean
End Type


Public Const TempHtm = "$$TEMP$$.HTM"
Public Const cHtmlAboutFilename = "about.htm"
'Public Const szhHtmlTemplate = "html\book.htm"
Public zhrStatus As zhReaderStatus
Private Const zhMemorySplit = vbCrLf


'Public Const TempHtm = "$$HTML$$.HTM"


Public Sub loadBookmark(inifiletodo As String, ByRef mnuBookmark As Object)

    Dim i As Integer
    Dim bCount As Long

        bCount = Val(iniGetSetting(inifiletodo, "Bookmark", "Count")) - 1
        On Error Resume Next
        For i = 0 To bCount
            Load mnuBookmark(i + 1)
            mnuBookmark(i + 1).Caption = iniGetSetting(inifiletodo, "Bookmark", "Name" & Str$(i))
            mnuBookmark(i + 1).Tag = iniGetSetting(inifiletodo, "Bookmark", "Location" & Str$(i))
            mnuBookmark(i + 1).Visible = True
        Next


End Sub

Public Sub saveBookmark(inifiletodo As String, ByRef mnuBookmark As Object)

    Dim i As Integer
    iniDeleteSection inifiletodo, "Bookmark"
    Dim fnum As Integer
    fnum = FreeFile
    Open inifiletodo For Append As fnum
    Print #fnum, "[Bookmark]"
    Print #fnum, "Count=" & mnuBookmark.Count - 1
    For i = 1 To mnuBookmark.Count - 1
        Print #fnum, "Name" & Str$(i - 1) & "=" & mnuBookmark(i).Caption
        Print #fnum, "Location" & Str$(i - 1) & "=" & mnuBookmark(i).Tag
    Next
    Close #fnum
    
End Sub

Public Sub GetReaderStyle(inifiletodo As String, RS As ReaderStyle)

    With RS.formPos
        .Height = CLngStr(iniGetSetting(inifiletodo, "ReaderStyle", "FormHeight"))
        .Width = CLngStr(iniGetSetting(inifiletodo, "ReaderStyle", "FormWidth"))
        .Top = CLngStr(iniGetSetting(inifiletodo, "ReaderStyle", "FormTop"))
        .Left = CLngStr(iniGetSetting(inifiletodo, "ReaderStyle", "FormLeft"))
    End With

    RS.WindowState = CLngStr(iniGetSetting(inifiletodo, "ReaderStyle", "WindowState"))
    RS.LeftWidth = CLngStr(iniGetSetting(inifiletodo, "ReaderStyle", "LeftWidth"))
    RS.LastPath = iniGetSetting(inifiletodo, "ReaderStyle", "LastPath")
    RS.ShowMenu = CBoolStr(iniGetSetting(inifiletodo, "ReaderStyle", "ShowMenu"))
    RS.ShowLeft = CBoolStr(iniGetSetting(inifiletodo, "ReaderStyle", "ShowLeft"))
    RS.ShowStatusBar = CBoolStr(iniGetSetting(inifiletodo, "ReaderStyle", "ShowStatusBar"))
    RS.FullScreenMode = CBoolStr(iniGetSetting(inifiletodo, "ReaderStyle", "FullScreenMode"))
    RS.TextEditor = iniGetSetting(inifiletodo, "ReaderStyle", "TextEditor")
End Sub

Public Sub SaveReaderStyle(inifiletodo As String, RS As ReaderStyle)

    iniDeleteSection inifiletodo, "ReaderStyle"
    Dim fnum As Integer
    fnum = FreeFile
    Open inifiletodo For Append As fnum
    Print #fnum, "[ReaderStyle]"
    With RS.formPos
        Print #fnum, "FormHeight=" & CStr(.Height)
        Print #fnum, "FormWidth=" & CStr(.Width)
        Print #fnum, "FormTop=" & CStr(.Top)
        Print #fnum, "FormLeft=" & CStr(.Left)
    End With
    Print #fnum, "LastPath=" & RS.LastPath
    Print #fnum, "WindowState=" & CStr(RS.WindowState)
    Print #fnum, "LeftWidth=" & CStr(RS.LeftWidth)
    Print #fnum, "ShowMenu=" & CStr(RS.ShowMenu)
    Print #fnum, "ShowLeft=" & CStr(RS.ShowLeft)
    Print #fnum, "ShowStatusBar=" & CStr(RS.ShowStatusBar)
    Print #fnum, "FullScreenMode=" & CStr(RS.FullScreenMode)
    Print #fnum, "TextEditor=" & RS.TextEditor
    
    Close #fnum
End Sub

Public Sub GetViewerStyle(inifiletodo As String, VS As ViewerStyle)

    With VS.Viewfont
        .Bold = (Val(iniGetSetting(inifiletodo, "ViewStyle", "Bold")) > 0)
        .Italic = (Val(iniGetSetting(inifiletodo, "ViewStyle", "Italic")) > 0)
        .Underline = (Val(iniGetSetting(inifiletodo, "ViewStyle", "Underline")) > 0)
        .Strikethrough = (Val(iniGetSetting(inifiletodo, "ViewStyle", "Strikethrough")) > 0)
        .name = iniGetSetting(inifiletodo, "ViewStyle", "Name")
        .Size = Val(iniGetSetting(inifiletodo, "ViewStyle", "Size"))

        If .Size = 0 Then .Size = 9
    End With

    With VS
        .ForeColor = Val(iniGetSetting(inifiletodo, "ViewStyle", "ForeColor"))
        .BackColor = Val(iniGetSetting(inifiletodo, "ViewStyle", "BackColor"))
        .LineHeight = Val(iniGetSetting(inifiletodo, "ViewStyle", "LineHeight"))

        If .LineHeight = 0 Then .LineHeight = 100
    End With
    
    VS.RecentMax = Val(iniGetSetting(inifiletodo, "Viewstyle", "RecentMax"))
    VS.TemplateFile = iniGetSetting(inifiletodo, "Viewstyle", "TemplateFile")
    VS.UseTemplate = CBoolStr(iniGetSetting(inifiletodo, "ViewStyle", "UseTemplate"))

End Sub

Public Sub SaveViewerStyle(inifiletodo As String, VS As ViewerStyle)
    
    iniDeleteSection inifiletodo, "ViewStyle"
    
    Dim fnum As Integer
    fnum = FreeFile
    Open inifiletodo For Append As fnum
    Print #fnum, "[ViewStyle]"
    
    Dim a As Integer

    With VS.Viewfont

        If .Bold Then a = 1 Else a = 0
        Print #fnum, "Bold=" & CStr(a)

        If .Italic Then a = 1 Else a = 0
        Print #fnum, "Italic=" & CStr(a)

        If .Underline Then a = 1 Else a = 0
        Print #fnum, "Underline=" & CStr(a)

        If .Strikethrough Then a = 1 Else a = 0
        Print #fnum, "Strikethrough=" & CStr(a)
        Print #fnum, "Name=" & .name
        Print #fnum, "Size=" & CStr(.Size)
    End With

    With VS
        Print #fnum, "ForeColor=" & CStr(.ForeColor)
        Print #fnum, "Backcolor=" & CStr(.BackColor)
        Print #fnum, "LineHeight=" & CStr(.LineHeight)
        Print #fnum, "UseTemplate=" & CStr(.UseTemplate)
        Print #fnum, "TemplateFile=" & .TemplateFile
        Print #fnum, "RecentMax=" & CStr(.RecentMax)
    End With
    
    Close #fnum
End Sub

Public Sub rememberNew(ByRef zhMemoryIn As String, ByVal szhFilename As String, ByVal ssecondPart As String)

    Dim fso As New Scripting.FileSystemObject
    Dim fsoMemoryTS As Scripting.TextStream
    Dim sMemoryText As String
    Dim stmp As String
    Dim posStart As Long
    Dim posEnd As Long
    Dim fMemoryDecrypted As String

    If szhFilename = "" Then Exit Sub
    If ssecondPart = "" Then Exit Sub

    fMemoryDecrypted = fso.BuildPath(Environ$("temp"), fso.GetTempName)
    MyFileDecrypt zhMemoryIn, fMemoryDecrypted
    Set fsoMemoryTS = fso.OpenTextFile(fMemoryDecrypted, ForReading, True)

    If fsoMemoryTS.AtEndOfStream = False Then sMemoryText = fsoMemoryTS.ReadAll
    fsoMemoryTS.Close
    Set fsoMemoryTS = fso.OpenTextFile(fMemoryDecrypted, ForWriting, True)
    posStart = InStr(sMemoryText, szhFilename & "|")

    If posStart > 0 Then posEnd = InStr(posStart, sMemoryText, zhMemorySplit, vbTextCompare)

    If posStart > 0 And posEnd > posStart Then
        stmp = Left$(sMemoryText, posStart - 1)
        stmp = stmp & szhFilename & "|" & ssecondPart & zhMemorySplit
        stmp = stmp & Right$(sMemoryText, Len(sMemoryText) - posEnd - Len(zhMemorySplit) + 1)
        sMemoryText = stmp
    Else
        sMemoryText = sMemoryText & szhFilename & "|" & ssecondPart & zhMemorySplit
    End If

    If Left$(sMemoryText, Len(zhMemorySplit)) = zhMemorySplit Then sMemoryText = Right$(sMemoryText, Len(sMemoryText) - Len(zhMemorySplit))
    fsoMemoryTS.Write sMemoryText
    fsoMemoryTS.Close
    MyFileEncrypt fMemoryDecrypted, zhMemoryIn
    fso.DeleteFile fMemoryDecrypted

End Sub

Public Function searchMemory(ByRef zhMemoryIn As String, ByRef szhFilename As String) As String

    Dim fso As New Scripting.FileSystemObject
    Dim fsoMemoryTS As Scripting.TextStream
    Dim sMemoryText As String
    Dim fMemoryDecrypted As String
    Dim posStart As Long
    Dim posEnd As Long

    fMemoryDecrypted = fso.BuildPath(Environ$("temp"), fso.GetTempName)
    MyFileDecrypt zhMemoryIn, fMemoryDecrypted
    Set fsoMemoryTS = fso.OpenTextFile(fMemoryDecrypted, ForReading, True)

    If fsoMemoryTS.AtEndOfStream = False Then sMemoryText = fsoMemoryTS.ReadAll
    fsoMemoryTS.Close
    posStart = InStr(sMemoryText, szhFilename & "|")

    If posStart > 0 Then posEnd = InStr(posStart, sMemoryText, zhMemorySplit, vbTextCompare)

    If posStart > 0 And posEnd > posStart + 1 Then
        searchMemory = Mid$(sMemoryText, posStart, posEnd - posStart)
        searchMemory = Replace(searchMemory, szhFilename & "|", "")
    End If

    fso.DeleteFile fMemoryDecrypted

End Function


Public Sub GetFileFilter(ByRef inifiletodo As String, ByRef cmbFilter As ComboBox)

Dim i As Integer
Dim ffNum As Long

ffNum = CLngStr(iniGetSetting(inifiletodo, "FileFilter", "Count")) - 1

For i = 0 To ffNum

cmbFilter.AddItem iniGetSetting(inifiletodo, "FileFilter", "F" + Str$(i)), i

Next

End Sub
Public Sub SaveFileFilter(ByRef inifiletodo As String, ByRef cmbFilter As ComboBox)

Dim ffNum As Long
Dim fnum As Integer
Dim i As Integer

iniDeleteSection inifiletodo, "FileFilter"

fnum = FreeFile
Open inifiletodo For Append As fnum
Print #fnum, "[FileFilter]"
Print #fnum, "Count=" & Str$(cmbFilter.ListCount)
ffNum = cmbFilter.ListCount - 1
For i = 0 To ffNum
Print #fnum, "F" & Str$(i) & "=" & cmbFilter.List(i)
Next
Close #fnum

End Sub

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

Dim fso As New gCFileSystem
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
