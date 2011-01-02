Attribute VB_Name = "ZtmMake"
Dim ZipAPP As String


Public Function GetContentEntry(indexhtm As String, chmContent() As String, chmContentNum As Integer) As String
Dim links() As String
Dim linknum As Integer
getlinks indexhtm, links(), linknum
ReDim chmContent(linknum) As String
For i = 1 To linknum
chmContent(i) = links(i, 1) + "|" + links(i, 0)
Next
chmContentNum = linknum
End Function
Public Function CreateContentFile(indexhtm As String, ContentFile As String, projectname As String) As Boolean

Dim fso As New FileSystemObject
If fso.FileExists(indexhtm) = False Then Exit Function
Dim ts As TextStream


    Set ts = fso.OpenTextFile(ContentFile, ForAppending, True)
    ts.WriteLine "[CONTENT]"
'Content Entry
Dim chmContent() As String
Dim chmContentNum As Integer
GetContentEntry indexhtm, chmContent(), chmContentNum
chmContent(0) = "|" + fso.GetFileName(indexhtm)
For i = 0 To chmContentNum
ts.WriteLine projectname + "\" + chmContent(i)
Next
ts.Close
CreateContentFile = True
End Function

Public Function CreateProject(indexhtm As String, projectname As String) As Boolean
Dim fso As New FileSystemObject
Dim ts As TextStream
Dim chmts As TextStream
Dim ContentFilename As String
Dim TopicFile As String
Dim thePPath As String
Dim theCfile As String

ContentFilename = ztmInfo

TopicFile = GetSetting(App.ProductName, "Preference", "TopicFile")
TopicFile = InputBox("Type TopicFile name.", "MakeZtm", fso.GetFileName(indexhtm) + " OR " + TopicFile)
SaveSetting App.ProductName, "Preference", "TopicFile", TopicFile

thePPath = fso.GetParentFolderName(indexhtm)
theCfile = fso.BuildPath(thePPath, ztmInfo)
Set ts = fso.CreateTextFile(theCfile, True)

    ts.WriteLine "[INFO]"
    ts.WriteLine "listshow=0"
    ts.WriteLine "menushow=1"
    ts.WriteLine "title=" + projectname
    ts.WriteLine "defaultfile=" + TopicFile
    ts.Close

CreateContentFile indexhtm, bddir(fso.GetParentFolderName(indexhtm)) + ContentFilename, projectname

Dim ContuneMake As VbMsgBoxResult
ContuneMake = MsgBox("Project: " + projectname + " done!" + vbCrLf + "Pack it as ZtmFile?", vbYesNo, "MakeZhm")
If ContuneMake = vbYes Then
    thePPath = fso.GetFolder(thePPath).ShortPath
    CreateZtm projectname, thePPath
End If
End Function

Public Sub Main()

Dim fso As New FileSystemObject
ZipAPP = fso.BuildPath(App.Path, "pkzip.exe")
If fso.FileExists(ZipAPP) = False Then
    MsgBox "File(s) below not exist" + vbCrLf + ZipAPP, vbExclamation, "Error"
    Exit Sub
End If
If Command$ = "" Then Exit Sub
Dim cmdLine As String
Dim projectname As String
cmdLine = Command$

If fso.FolderExists(cmdLine) Then
projectname = InputBox("Input the Name of the Porject:", cmdLine, fso.GetBaseName(cmdLine))
ElseIf fso.FileExists(cmdLine) Then
projectname = InputBox("Input the Name of the Porject:", cmdLine, fso.GetBaseName(fso.GetParentFolderName(cmdLine)))
Else
Exit Sub
End If

If projectname = "" Then Exit Sub
If fso.FolderExists(cmdLine) Then
CreatePROJECTFromDir cmdLine, projectname
Else
CreateProject cmdLine, projectname
End If
End
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
    trlinE = 1
    ts.WriteLine "<tr>"
    For Each fsoF In fso.GetFolder(dirname).Files
    m = m + 1
    If m > trlinE Then
    m = 1
ts.WriteLine "</tr>"
ts.WriteLine "<tr>"
    End If
    

    Dim linkname As String
    Dim linkhref As String
 
    linkname = fso.GetBaseName(fsoF.Path)
    linkhref = bddir(fso.GetFileName(dirname)) + fsoF.name
ts.WriteLine "<td width=" + Chr(34) + Str(100 \ trlinE) + "%" + Chr(34) + ">"
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

 Function CreateZtm(theZtmfile As String, theZtmPath As String)
 ChDrive Left(theZtmPath, 1)
 ChDir theZtmPath
 Dim cmdLine As String
 cmdLine = ZipAPP + " -a -r -p " + theZtmfile + ".ztm " + "*.*"
 Shell cmdLine, vbNormalFocus
 End Function
