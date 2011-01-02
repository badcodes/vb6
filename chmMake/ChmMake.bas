Attribute VB_Name = "LchmMake"

Public Function GetContentEntry(indexhtm As String, chmContent() As String, chmContentNum As Integer) As String
Dim links() As String
Dim linknum As Long
linknum = LiNVBLib.gCHtmlWeb.getTagsProperty(indexhtm, "", "href", links())
'getlinks indexhtm, links(), linknum
ReDim chmContent(linknum) As String



For i = 0 To linknum - 1
chmContent(i) = chmContent(i) + "       <LI><OBJECT type=" + Chr(34) + "text/sitemap" + Chr(34) + ">" + Chr(13) + Chr(10)
chmContent(i) = chmContent(i) + "           <param name=" + Chr(34) + "Name" + Chr(34) + " value=" + Chr(34) + "####" + Chr(34) + ">" + Chr(13) + Chr(10)
chmContent(i) = chmContent(i) + "           <param name=" + Chr(34) + "Local" + Chr(34) + " value=" + Chr(34) + links(i) + Chr(34) + ">" + Chr(13) + Chr(10)
chmContent(i) = chmContent(i) + "       </OBJECT>" + Chr(13) + Chr(10)
Next
chmContentNum = linknum
End Function
Public Function CreateContentFile(indexhtm As String, ContentFile As String) As Boolean

Dim fso As New FileSystemObject
If fso.FileExists(indexhtm) = False Then Exit Function
Dim ts As TextStream

'Content Head
    Set ts = fso.OpenTextFile(ContentFile, ForWriting, True)
ts.WriteLine "<!DOCTYPE HTML PUBLIC " + Chr(34) + "-//IETF//DTD HTML//EN" + Chr(34) + ">"
ts.WriteLine "<HTML><HEAD>"
ts.WriteLine "<meta name=" + Chr(34) + "GENERATOR" + Chr(34) + " content=" + Chr(34) + "Microsoft&reg; HTML Help Workshop 4.1" + Chr(34) + ">"
ts.WriteLine "<!-- Sitemap 1.0 -->"
ts.WriteLine "</HEAD><BODY>"
ts.WriteLine "<OBJECT type=" + Chr(34) + "text/site properties" + Chr(34) + ">"
ts.WriteLine "<param name=" + Chr(34) + "ImageType" + Chr(34) + " value=" + Chr(34) + "Folder" + Chr(34) + ">"
ts.WriteLine "</OBJECT>"
ts.WriteLine "<UL>"

'Content Entry
Dim chmContent() As String
Dim chmContentNum As Integer
GetContentEntry indexhtm, chmContent(), chmContentNum

chmContent(0) = chmContent(0) + "   <LI><OBJECT type=" + Chr(34) + "text/sitemap" + Chr(34) + ">" + Chr(13) + Chr(10)
chmContent(0) = chmContent(0) + "       <param name=" + Chr(34) + "Name" + Chr(34) + " value=" + Chr(34) + fso.GetBaseName(ContentFile) + Chr(34) + ">" + Chr(13) + Chr(10)
chmContent(0) = chmContent(0) + "       <param name=" + Chr(34) + "Local" + Chr(34) + " value=" + Chr(34) + fso.GetFileName(indexhtm) + Chr(34) + ">" + Chr(13) + Chr(10)
chmContent(0) = chmContent(0) + "   </OBJECT>" + Chr(13) + Chr(10)
chmContent(0) = chmContent(0) + "   <UL>" + Chr(13) + Chr(10)

For i = 0 To chmContentNum - 1
ts.Write chmContent(i)
Next
'Content trailer
ts.WriteLine "   </UL>"
ts.WriteLine "</UL>"
ts.WriteLine "</BODY></HTML>"
ts.Close
CreateContentFile = True
End Function

Public Function CreateProject(indexhtm As String, projectname As String) As Boolean
Dim fso As New FileSystemObject
Dim ts As TextStream
Dim chmts As TextStream
Dim ContentFilename As String
Dim TopicFile As String
ContentFilename = projectname + ".hhc"
TopicFile = fso.GetFileName(indexhtm)

  Set ts = fso.CreateTextFile(bddir(fso.GetParentFolderName(indexhtm)) + projectname + ".hhp", True)
If fso.FileExists(bddir(App.Path) + "#####PROJECTNAME#####.hhp") = False Then
    ts.WriteLine "[Options]"
    ts.WriteLine "Binary Index = No"
    ts.WriteLine "Compatibility = 1.1 Or later"
    ts.WriteLine "Compiled File = " + projectname + ".chm"
    ts.WriteLine "Contents File = " + ContentFilename
    ts.WriteLine "Default Window = lin"
    ts.WriteLine "Default topic=" + TopicFile
    ts.WriteLine "Display compile progress=No"
    ts.WriteLine "Language=0x804 中文(中国)"
    ts.WriteLine "Title = " + projectname + " - Complied by LiN MYPLACE 2004"
    ts.WriteLine "[WINDOWS]"
    ts.WriteLine "lin=," + ContentFilename + ",," + Chr(34) + TopicFile + Chr(34) + "," + Chr(34) + TopicFile + Chr(34) + ",,,,,0x3020,,0x307e,,0x18a0000,,,,,,0"
    ts.WriteLine "[Files]"
    ts.WriteLine TopicFile
    ts.WriteLine "[INFOTYPES]"
    ts.WriteLine ""
    ts.Close
Else
Set chmts = fso.OpenTextFile(bddir(App.Path) + "#####PROJECTNAME#####.hhp", ForReading)
Dim tempstr As String
tempstr = chmts.ReadAll
tempstr = strreplace(tempstr, "#####PROJECTNAME#####", projectname)
tempstr = strreplace(tempstr, "#####TOPICFILE#####", TopicFile)
ts.Write tempstr
ts.Close


End If
CreateContentFile indexhtm, bddir(fso.GetParentFolderName(indexhtm)) + ContentFilename
fso.CopyFile bddir(fso.GetParentFolderName(indexhtm)) + ContentFilename, bddir(fso.GetParentFolderName(indexhtm)) + projectname + ".hhk"
MsgBox "Project: " + projectname + " done!", , "MakeChm"
End Function

Private Sub BatchMake()
Dim fso As New FileSystemObject
Dim ffolder As Folder
Dim ffff As Folder
Dim ts As TextStream
For Each ffolder In fso.GetFolder(Text1.Text).SubFolders
    For Each ffff In ffolder.SubFolders
    TopicFile = ffolder.Name + "\" + ffff.Name + "\index.htm"
    Label1.Caption = TopicFile
    MainFrm.Refresh
    projectname = ffolder.Name
    Set ts = fso.CreateTextFile(Text1.Text + "\" + ffolder.Name + ".hhp")
    ts.WriteLine "[Options]"
    ts.WriteLine "Binary Index = No"
    ts.WriteLine "Compatibility = 1.1 Or later"
    ts.WriteLine "Compiled File = " + projectname + ".chm"
    ts.WriteLine "Default Window = lin"
    ts.WriteLine "Default topic=" + TopicFile
    ts.WriteLine "Display compile progress=No"
    ts.WriteLine "Language=0x804 中文(中国)"
    ts.WriteLine "Title = " + projectname + " - Complied by LiN MYPLACE 2004"
    ts.WriteLine "[WINDOWS]"
    ts.WriteLine "lin=,,," + Chr(34) + TopicFile + Chr(34) + "," + Chr(34) + TopicFile + Chr(34) + ",,,,,0x2020,,0x307e,,0x1890000,,,,,,0"
    ts.WriteLine "[Files]"
    ts.WriteLine TopicFile
    ts.WriteLine "[INFOTYPES]"
    ts.WriteLine ""
    ts.Close
    Next
Next
Set ts = fso.CreateTextFile(Text1.Text + "\makechm.bat")
ts.WriteLine "for " + "%" + "%" + "f in (*.hhp) do complie.bat " + "%" + "%" + "f"
ts.WriteLine "del makechm.bat"
ts.WriteLine "del complie.bat"
ts.Close
Set ts = fso.CreateTextFile(Text1.Text + "\complie.bat")
ts.WriteLine "hhc " + "%" + "1"
ts.WriteLine "del " + "%" + "1"
ts.Close
End Sub

Public Sub Main()
If Command$ = "" Then Exit Sub
Dim cmdline As String
Dim projectname As String
cmdline = Command$


projectname = InputBox("Input the Name of the Porject:")
If projectname = "" Then Exit Sub
Dim fso As New FileSystemObject
If fso.FolderExists(cmdline) Then
CreatePROJECTFromDir cmdline, projectname
Else
CreateProject cmdline, projectname
End If
End Sub

Public Function CreatePROJECTFromDir(dirname As String, projectname As String) As Boolean


Dim fso As New FileSystemObject
Dim fsoF As File
Dim ts As TextStream
Dim indexts As TextStream
Dim tempstr As String
If fso.FolderExists(dirname) = False Then Exit Function
If fso.FileExists(bddir(App.Path) + "index.htm") = False Then MsgBox ("TemplatE of IndeX FilE NoT ExisT"): Exit Function
Set indexts = fso.OpenTextFile(bddir(App.Path) + "index.htm", ForReading)
Set ts = fso.OpenTextFile(bddir(fso.GetParentFolderName(dirname)) + fso.GetBaseName(dirname) + "index.htm", ForWriting, True)
Do Until indexts.AtEndOfStream
tempstr = indexts.ReadLine
    If tempstr = "#####TITLE#####" Then
ts.WriteLine projectname
    ElseIf tempstr = "#####LINKTR#####" Then
    
    m = 0
    trline = 3
    ts.WriteLine "<tr>"
    For Each fsoF In fso.GetFolder(dirname).Files
    m = m + 1
    If m > trline Then
    m = 1
ts.WriteLine "</tr>"
ts.WriteLine "<tr>"
    End If
    
    Dim Fext As String
    Dim linkname As String
    Dim linkhref As String
    Fext = LCase(fso.GetExtensionName(fsoF.Path))
    If Fext = "htm" Or Fext = "html" Then
        Dim HtmDoc As New HTMLDocument
        Dim theHtm As IHTMLDocument2
        Set theHtm = HtmDoc.createDocumentFromUrl(fsoF.Path, "")
        Do Until theHtm.Title <> "" Or theHtm.readyState = "complete"
        DoEvents
        Loop
        linkname = theHtm.Title
        If linkname = "" Then linkname = fso.GetBaseName(fsoF.Path)
     Else
        linkname = fso.GetBaseName(fsoF.Path)
    End If
    linkhref = bddir(fso.GetFileName(dirname)) + fsoF.Name
ts.WriteLine "<td width=" + Chr(34) + Str(100 \ 3) + "%" + Chr(34) + ">"
ts.WriteLine "<a href=" + Chr(34) + linkhref + Chr(34) + ">" + linkname + "</a></td>"
    Next
ts.WriteLine "</tr>"
    Else
ts.WriteLine tempstr
    End If
Loop

indexts.Close
ts.Close

CreateProject bddir(fso.GetParentFolderName(dirname)) + fso.GetBaseName(dirname) + "index.htm", projectname
End Function
