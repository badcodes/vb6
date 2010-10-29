VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GQuickWork"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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

