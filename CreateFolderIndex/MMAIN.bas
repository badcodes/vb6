Attribute VB_Name = "MMAIN"
Option Explicit
Public Sub Main()
Dim fso As New FileSystemObject
Dim sFolder As String
Dim mR As VbMsgBoxResult
sFolder = Command$
If fso.FolderExists(sFolder) = False Then
sFolder = fso.BuildPath(CurDir$, sFolder)
End If
sFolder = fso.GetAbsolutePathName(sFolder)
If fso.FolderExists(sFolder) = False Then Exit Sub
mR = MsgBox("将在目录" & Chr$(34) & sFolder & Chr$(34) & "及其子目录下创建index.htm文件", vbOKCancel)
If mR <> vbOK Then Exit Sub
Set fso = Nothing
CreateFolderIndex sFolder
End Sub
Public Sub CreateFolderIndex(sFolder As String, Optional sParent As String = "")

    Dim fso As New FileSystemObject
    Dim fds As Folders
    Dim fd As Folder
    Dim fs As Files
    Dim f As File
    Dim ts As TextStream
    Dim i As Long
    Dim fdcount As Long
    Dim fcount As Long
    Dim subFolders() As String
    Dim subFiles() As String
    
    Set fds = fso.GetFolder(sFolder).subFolders
    Set fs = fso.GetFolder(sFolder).Files
    fdcount = fds.Count
    fcount = fs.Count
    
    If fdcount > 0 Then ReDim subFolders(1 To fds.Count) As String
    If fcount > 0 Then ReDim subFiles(1 To fs.Count) As String
    
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
        'ts.WriteLine "<img src='folder.gif'>"
        ts.WriteLine "[DIR]<a href='../index.htm' alt=' " & fso.GetFileName(sParent) & "'>..</a>"
        ts.WriteLine "</td></tr>"
    End If
    For i = 1 To fdcount
        ts.WriteLine "<tr><td>"
        'ts.WriteLine "<img src='folder.gif'>"
        ts.WriteLine "[DIR]<a href='" & fso.GetFileName(subFolders(i)) & "/index.htm' >" & fso.GetFileName(subFolders(i)) & "</a>"
        ts.WriteLine "</td></tr>"
        CreateFolderIndex subFolders(i), sFolder
    Next
    For i = 1 To fcount
        ts.WriteLine "<tr><td>"
        'ts.WriteLine "<img src='file.gif'>"
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

