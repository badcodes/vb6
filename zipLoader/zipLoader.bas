Attribute VB_Name = "zipLoaderMain"
Option Explicit
Private PASSWORD As String
Private InvaildPassword As Boolean
Const modHtmlWeb_SplitSymbol = "|"
Const modHtmlWeb_WebsiteDefaultFile = "·âÃæ|cover|Ê×Ò³|index|default|start|home|Ä¿Â¼|content|contents|aaa|bbb|00"


Public Function getZipFirstFile(ByVal thisfile As String) As String

    Dim fso As New gCFileSystem
    Dim firstfile As String
    Dim sTmpFile As String
    Dim lUnzip As New cUnzip
    
    
    If fso.PathExists(thisfile) = False Then Exit Function
    

    Dim sTmpText As String
    
    With lUnzip
    .ZipFile = thisfile
    End With
    
    sTmpText = lUnzip.GetComment
    
    
    firstfile = LeftRange(sTmpText, "defaultfile", vbCrLf, vbTextCompare, ReturnEmptyStr)
    If firstfile = "" Then firstfile = LeftRange(sTmpText, "defaultfile", vbLf, vbTextCompare, ReturnEmptyStr)
    firstfile = LeftRight(firstfile, "=", vbTextCompare, ReturnEmptyStr)
    firstfile = Trim(firstfile)
        
    If firstfile = "" Then
        
    
        Dim zipFileList As New CZipItems
        
        lUnzip.getZipItems zipFileList
        
        Set lUnzip = Nothing
    
        Dim sZipFiles() As String
        Dim lzipFilescount As Long
        Dim sArrHtmfile() As String
        Dim lHtmFileCount As Long
        Dim sExtName As String
        Dim sDefaultfile As String
        Dim lEnd As Long
        Dim m As Long
        
        lEnd = zipFileList.Count
        For m = 1 To lEnd
        If zipFileList(m).FileType <> vbDirectory Then
            ReDim Preserve sZipFiles(lzipFilescount) As String
            sZipFiles(lzipFilescount) = zipFileList(m).FileName
            lzipFilescount = lzipFilescount + 1
        End If
        Next
        
        Set zipFileList = Nothing
        
        lEnd = lzipFilescount - 1
        For m = 0 To lEnd
            sExtName = LCase$(fso.GetExtensionName(sZipFiles(m)))
            
        '    If sExtName = LCase$(cTxtIndex) Then
        '        sDefaultfile = sZipFiles(m)
        '        loadCmdLine = starthttp(thisfile, sHttpServerHead, sDefaultfile)
        '        Exit Function
        '    End If
            
            If sExtName = "htm" Or sExtName = "html" Then
                If IsWebsiteDefaultFile(sZipFiles(m)) Then
                    ReDim Preserve sArrHtmfile(lHtmFileCount) As String
                    sArrHtmfile(lHtmFileCount) = sZipFiles(m)
                    lHtmFileCount = lHtmFileCount + 1
                End If
            End If
        Next
        
        If lHtmFileCount > 1 Then
            sDefaultfile = sArrHtmfile(0)
            QuickSortFiles sArrHtmfile, 0, lHtmFileCount - 1
            firstfile = findDefaultHtml(sArrHtmfile)
        ElseIf lHtmFileCount = 1 Then
            firstfile = sArrHtmfile(0)
        End If

    End If

    getZipFirstFile = firstfile
End Function



Sub Main()
Dim sArgu As String
Dim sFirstFile As String
sArgu = Command$
sArgu = RightDelete(sArgu, Chr(34))
sArgu = LeftDelete(sArgu, Chr(34))
sFirstFile = getZipFirstFile(sArgu)
If sFirstFile = "" Then
    MsgBox sArgu & vbCrLf & "not a valid zipPacked Html File!"
Else
    ShellExecute 1, "open", toUnixPath(zipProtocolHead & sArgu & zipSep & sFirstFile), "", "", 1
End If

End Sub
