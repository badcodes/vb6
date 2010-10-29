Attribute VB_Name = "MQuickWork"
Public Sub batRenameByTextfile()

    Dim fso As New FileSystemObject
    Dim ts As TextStream
    Dim pdfname As String
    Dim pdfAuthor As String
    Dim pdfTitle As String
    Dim pdfLine() As String
    Dim pdfRealfile As String
    Dim pdfCopyto As String
    Set ts = fso.GetFile("e:\t06.dbl").OpenAsTextStream

    Do Until ts.AtEndOfStream
        pdfLine = Split(ts.ReadLine, "|")
        pdfTitle = pdfLine(0)
        pdfname = pdfLine(1)
        pdfAuthor = pdfLine(2)
        pdfRealfile = fso.BuildPath("H:\dbook", pdfname)
        pdfCopyto = pdfAuthor & " - " & pdfTitle
        pdfCopyto = cleanFilename(pdfCopyto)
        pdfCopyto = fso.BuildPath("e:\iso\", pdfCopyto & ".pdf")

        If fso.FileExists(pdfRealfile) Then
            fso.CopyFile pdfRealfile, pdfCopyto
        End If

    Loop

End Sub
