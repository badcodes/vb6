VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ZTM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public showlist As Boolean
Public showmenu As Boolean
Public contentcount As Integer
Public defaultfile As String
Public title As String
Public author As String

Sub getcontent(ContentFile As String, content() As String)

Dim fso As New FileSystemObject
Dim ts As TextStream
Dim theline As String
Dim pos As Integer
Dim pos2 As Integer

If fso.FileExists(ContentFile) = False Then contentcount = 0: Exit Sub

Set ts = fso.OpenTextFile(ContentFile)

Do Until ts.AtEndOfStream
theline = ts.ReadLine
If theline <> "" Then
    contentcount = contentcount + 1
    ReDim Preserve content(1, contentcount) As String
    pos = InStrRev(theline, "|")
    pos2 = InStrRev(theline, ",")
    If pos > 0 Then
        content(0, contentcount) = Left(theline, pos - 1)
        content(1, contentcount) = Right(theline, Len(theline) - pos)
    ElseIf pos2 > 0 Then
        content(0, contentcount) = Left(theline, pos2 - 1)
        content(1, contentcount) = Right(theline, Len(theline) - pos2)
    Else
        content(0, contentcount) = theline
        content(1, contentcount) = ""
    End If
End If

Loop
ts.Close
End Sub

Sub loadinfo(stylefile As String)

Dim fso As New FileSystemObject
Dim ts As TextStream
Dim theline(1) As String
If fso.FileExists(stylefile) = False Then Exit Sub

Set ts = fso.OpenTextFile(stylefile, ForReading, False)
Do Until ts.AtEndOfStream
theline(0) = ts.ReadLine
pos = InStr(theline(0), "=")
If pos > 0 Then
    theline(1) = Right(theline(0), Len(theline(0)) - pos)
    theline(0) = LCase(Left(theline(0), pos - 1))
    Select Case theline(0)
        Case "listshow"
            If Val(theline(1)) > 0 Then showlist = True
        Case "menushow"
            If Val(theline(1)) > 0 Then showmenu = True
        Case "defaultfile"
            defaultfile = theline(1)
        Case "author"
            author = theline(1)
        Case "title"
            title = theline(1)
    End Select
End If
Loop
ts.Close

End Sub

