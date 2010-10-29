Attribute VB_Name = "MCHM"


Function newProject(projectName As String, contentFilename As String, topicFile As String, ByRef sFileList() As String) As String

   
    newProject = _
        "[Options]" & _
        vbCrLf & "Binary Index=No" & _
        vbCrLf & "Compatibility=1.1 Or later" & _
        vbCrLf & "Compiled File=" & projectName & ".chm" & _
        vbCrLf & "Contents File=" & contentFilename & _
        vbCrLf & "Default Window=lin" & _
        vbCrLf & "Default topic=" & topicFile & _
        vbCrLf & "Display compile progress=Yes" & _
        vbCrLf & "Full-text search=Yes" & _
        vbCrLf & "Language=0x409 Ó¢Óï(ÃÀ¹ú)" & _
        vbCrLf & "Title=" & projectName & " - Complied by xiaoranzzz@MYPLACE " & Year(DateTime.Date) & _
        vbCrLf & "[WINDOWS]" & _
        vbCrLf & "lin=," & Quote(contentFilename) & ",," & Quote(topicFile) & "," & Quote(topicFile) & ",,,,,0x43520,,0x384e,[172,163,1081,920],,,,,,,0" & _
        vbCrLf & "[INFOTYPES]" & _
        vbCrLf & "[Files]"

    Dim i As Long
    Dim u As Long
    On Error Resume Next
    u = -1
    u = UBound(sFileList)
    For i = 0 To u
        newProject = newProject & vbCrLf & sFileList(i)
    Next

End Function

 
Function getTopicFiles(ByRef pathHHC As String) As String()
 
 Dim nFileNum As Integer
 Dim sBuffer As String
 Dim sTopic As String
 nFileNum = FreeFile
 On Error GoTo Error_FileAccess
 
 Open pathHHC For Input As #nFileNum
 Dim nPos As Integer
 Dim cQuote As String
 Dim nCount As Integer
 Dim sTopics() As String
 
 cQuote = Chr$(34)
 
 Dim nEndPos As Integer
 Do Until EOF(nFileNum)
    Line Input #nFileNum, sBuffer
    nPos = InStr(1, sBuffer, "<param name=" & cQuote & _
        "Local" & cQuote, vbTextCompare)
    If (nPos > 0) Then
        nPos = InStr(1, sBuffer, "value=" & cQuote, vbTextCompare)
        If (nPos > 0) Then
            nPos = nPos + 7
            nEndPos = InStr(nPos, sBuffer, cQuote, vbTextCompare)
            sTopic = Mid$(sBuffer, nPos, nEndPos - nPos)
            If (sTopic <> "") Then
                ReDim Preserve sTopics(0 To nCount)
                sTopics(nCount) = sTopic
                nCount = nCount + 1
                'Debug.Print sTopic
            End If
        End If
    End If
 Loop
 
 Close #nFileNum
 getTopicFiles = sTopics

 
 Exit Function
Error_FileAccess:
    On Error Resume Next
    Close #nFileNum
    Err.Raise Err.Number
    Exit Function
End Function

Public Sub CHM_InsertLinksTable(ByVal pathFolder As String, ByVal pathHHC As String)
    Dim sTopicFiles() As String
    sTopicFiles = MCHM.getTopicFiles(pathHHC)
    Dim nCount As Integer
    nCount = UBound(sTopicFiles) + 1
    If (nCount < 1) Then Exit Sub
    If (MFileSystem.FolderExists(pathFolder) = False) Then Exit Sub
    Dim sScript As String
    
    
    pathFolder = MFileSystem.BuildPath(pathFolder, "")
    sScript = pathFolder & "pn_prevnext.js"
    FileCopy "X:\Workspace\VB\HtmlPrevNext\pn_prevnext.js", sScript
    FileCopy "X:\Workspace\VB\HtmlPrevNext\pn_prev.gif", pathFolder & "pn_prev.gif"
    FileCopy "X:\Workspace\VB\HtmlPrevNext\pn_next.gif", pathFolder & "pn_next.gif"
    Dim nFileNum As Integer
    nFileNum = FreeFile()
    Open sScript For Append As #nFileNum
    Print #nFileNum, "pnFileList = new Array("
    
    Dim cQuote As String
    Dim i As Integer
    cQuote = Chr$(34)
    For i = 0 To nCount - 2
        Print #nFileNum, "    " & cQuote & MFileSystem.GetFileName(sTopicFiles(i)) & cQuote & ","
    Next
    Print #nFileNum, "    " & cQuote & MFileSystem.GetFileName(sTopicFiles(nCount - 1)) & cQuote
    Print #nFileNum, ");"
    Close #nFileNum
    
    If MFileSystem.DirEx(pathFolder, sTopicFiles, , vbDirectory) Then
    For i = LBound(sTopicFiles) To UBound(sTopicFiles)
        If (Left$(MFileSystem.GetExtensionName(sTopicFiles(i)), 3) = "htm") Then
            nFileNum = FreeFile()
            Open MFileSystem.BuildPath(pathFolder, sTopicFiles(i)) For Append As #nFileNum
            'Debug.Print sTopicFiles(i)
            Print #nFileNum, "<script language=" & cQuote & "JavaScript" & cQuote & " type=" & cQuote & "text/javascript" & cQuote & " src=pn_prevnext.js></script>"
            Close nFileNum
        End If
    Next
    
    End If
    
    
End Sub
