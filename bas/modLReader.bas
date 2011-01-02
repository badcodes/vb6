Attribute VB_Name = "modLReader"


Type MYFont
    Bold As Boolean
    Italic As Boolean
    Charset As Integer
    Name As String
    Size As Currency
    Strikethrough As Boolean
    Underline As Boolean
End Type
Type MYPoS
    Top As Long
    Left As Long
    Height As Long
    Width As Long
End Type

Type ReaderStyle
    FormPOS As MYPoS
    LeftPos As Long
End Type

Type ViewerStyle
    Viewfont As MYFont
    ForeColor As OLE_COLOR
    BackColor As OLE_COLOR
    LineHeight As Integer
End Type

Public Const BACKWARDVIEW = -1
Public Const FORWARDVIEW = -2
Public Const PREVIOUSVIEW = -3
Public Const NEXTVIEW = -4
Public Const RANDOMVIEW = -100
Public Const MAXFILTER = 99
Public NumToNodeIndex() As Integer
Public NodeIndexToNum() As Integer
Public PathReading As String
Public FileFilter(99, 1) As String
Public FilterCount As Integer
Public FilterNum As Integer
Public FileCount As Integer
Public FolderCount As Integer
Public CurFile As Integer
Public Filezip As Boolean
Public FileIe As Boolean
Public Tempdir As String
Public Const TempHtm = "$$HTML$$.HTM"
Public MenuShow As Boolean
Public ListShow As Boolean
Public theRS As ReaderStyle
Public theVS As ViewerStyle




Public Sub CopytoIfont(srcfont As MYFont, dstfont As IFontDisp)

With dstfont
.Bold = srcfont.Bold
.Italic = srcfont.Italic
.Name = srcfont.Name
.Size = srcfont.Size
.Strikethrough = srcfont.Strikethrough
.Underline = srcfont.Underline
.Charset = srcfont.Charset
End With
End Sub

Public Sub Setfont(dstobject As Object, srcfont As MYFont)

With dstobject
.FontBold = srcfont.Bold
.FontItalic = srcfont.Italic
If srcfont.Name <> "" Then .FontName = srcfont.Name
.FontSize = srcfont.Size
.FontStrikethru = srcfont.Strikethrough
.FontUnderline = srcfont.Underline
End With


End Sub

Public Function toMYfont(srcfont As IFontDisp) As MYFont
With toMYfont
.Bold = srcfont.Bold
.Italic = srcfont.Italic
.Name = srcfont.Name
.Size = srcfont.Size
.Strikethrough = srcfont.Strikethrough
.Underline = srcfont.Underline
.Charset = srcfont.Charset
End With
End Function

Public Sub webview(IEView As WebBrowser, viewfile As String)

Dim fso As New FileSystemObject
Dim ts As TextStream
Dim TempFile As String

FileIe = False
TempFile = viewfile

Select Case chkFileType(TempFile) 'file.bas
    Case ftZIP
         Filezip = True
         MainFrm.LoadlistFromZipFile TempFile
         Exit Sub
    Case ftHTML
         FileIe = True
         IEView.Navigate TempFile
         Exit Sub
    Case ftIE
         FileIe = True
         IEView.Navigate TempFile
         Exit Sub
    Case ftExE
         Shell TempFile + " " + InputBox("Ready to run :" + Chr(13) + Chr(10) + TempFile + Chr(13) + Chr(10) + "Input the Arguments :" + Chr(13) + Chr(10), "RUN")
         Exit Sub
    Case ftCHM
         ListShow = False
         frmList.Hide
         Shell Environ("systemroot") + "\hh.exe " + TempFile, vbNormalFocus
         Exit Sub
    Case ftIMG
         tempfile2 = Tempdir + "\" + TempHtm
         Set ts = fso.OpenTextFile(tempfile2, ForWriting, True)
         ts.WriteLine "<HTML><HEAD>"
         ts.WriteLine "<link REL=stylesheet href=" + Chr(34) + bddir(App.Path) + "style.css" + Chr(34) + " type=" + Chr(34) + "text/css" + Chr(34) + ">"
         ts.WriteLine "<script src=" + Chr(34) + bddir(App.Path) + "scroll.js" + Chr(34) + "></script>"
         ts.WriteLine "</HEAD><BODY>"
         ts.WriteLine "<center>"
         ts.WriteLine "<TABLE  width=" + Chr(34) + "96%" + Chr(34) + " cellSpacing=4 cellPadding=4 width=" + Chr(34) + "100%" + Chr(34) + ">"
         ts.WriteLine "<TR ><TD  align=center class=m_text>"
         ts.WriteLine "<img src=" + Chr(34) + TempFile + Chr(34) + ">"
         ts.WriteLine "</td></tr></table>"
         ts.WriteLine "</body></html>"
         ts.Close
         IEView.Navigate tempfile2
         Exit Sub
    Case ftAUDIO
         tempfile2 = Tempdir + "\" + TempHtm
         Set ts = fso.OpenTextFile(tempfile2, ForWriting, True)
         ts.WriteLine "<HTML><HEAD>"
         ts.WriteLine "<link REL=stylesheet href=" + Chr(34) + bddir(App.Path) + "style.css" + Chr(34) + " type=" + Chr(34) + "text/css" + Chr(34) + ">"
         ts.WriteLine "</HEAD><BODY>"
         ts.WriteLine "<center>"
         ts.WriteLine "<TABLE  width=" + Chr(34) + "96%" + Chr(34) + " cellSpacing=4 cellPadding=4 width=" + Chr(34) + "100%" + Chr(34) + ">"
         ts.WriteLine "<TR ><TD  align=center class=m_text>"
         ts.WriteLine "<object classid=" + Chr(34) + "clsid:22D6F312-B0F6-11D0-94AB-0080C74C7E95" + Chr(34) + " id=" + Chr(34) + "MediaPlayer1" + Chr(34) + ">"
         ts.WriteLine "<param name=" + Chr(34) + "Filename" + Chr(34) + " value=" + Chr(34) + TempFile + Chr(34) + ">"
         ts.WriteLine "</object>"
         ts.WriteLine "</td></tr></table>"
         ts.WriteLine "</body></html>"
         ts.Close
         IEView.Navigate tempfile2
         Exit Sub
    Case ftVIDEO
         tempfile2 = Tempdir + "\" + TempHtm
         Set ts = fso.OpenTextFile(tempfile2, ForWriting, True)
         ts.WriteLine "<HTML><HEAD>"
         ts.WriteLine "<link REL=stylesheet href=" + Chr(34) + bddir(App.Path) + "style.css" + Chr(34) + " type=" + Chr(34) + "text/css" + Chr(34) + ">"
         ts.WriteLine "</HEAD><BODY>"
         ts.WriteLine "<center>"
         ts.WriteLine "<TABLE  width=" + Chr(34) + "96%" + Chr(34) + " cellSpacing=4 cellPadding=4 width=" + Chr(34) + "100%" + Chr(34) + ">"
         ts.WriteLine "<TR ><TD  align=center class=m_text>"
         ts.WriteLine "<object classid=" + Chr(34) + "clsid:22D6F312-B0F6-11D0-94AB-0080C74C7E95" + Chr(34) + " id=" + Chr(34) + "MediaPlayer1" + Chr(34) + ">"
         ts.WriteLine "<param name=" + Chr(34) + "Filename" + Chr(34) + " value=" + Chr(34) + TempFile + Chr(34) + ">"
         ts.WriteLine "</object>"
         ts.WriteLine "</td></tr></table>"
         ts.WriteLine "</body></html>"
         ts.Close
         IEView.Navigate tempfile2
         Exit Sub
    Case Else
         tempfile2 = Tempdir + "\" + TempHtm
         Set ts = fso.OpenTextFile(tempfile2, ForWriting, True)
         ts.WriteLine "<HTML><HEAD>"
         ts.WriteLine "<link REL=stylesheet href=" + Chr(34) + bddir(App.Path) + "style.css" + Chr(34) + " type=" + Chr(34) + "text/css" + Chr(34) + ">"
         ts.WriteLine "<script src=" + Chr(34) + bddir(App.Path) + "scroll.js" + Chr(34) + "></script>"
         ts.WriteLine "</HEAD><BODY>"
         ts.WriteLine "<center>"
         ts.WriteLine "<TABLE  width=" + Chr(34) + "96%" + Chr(34) + " cellSpacing=4 cellPadding=4 width=" + Chr(34) + "100%" + Chr(34) + ">"
         ts.WriteLine "<TR align=center><TD  align=left class=m_text>"
         Dim tempts As TextStream
         Set tempts = fso.OpenTextFile(TempFile, ForReading)
         Do Until tempts.AtEndOfStream
         ts.WriteLine tempts.ReadLine + "<br>" 'HTMLBODY
         Loop
         ts.WriteLine "</TD></TR></TABLE></center></body></html>"
         ts.Close
         tempts.Close
         IEView.Navigate tempfile2
End Select
    

End Sub

Public Sub GetViewerStyle(VS As ViewerStyle)
    With VS.Viewfont
    .Bold = (Val(GetSetting(App.ProductName, "ViewStyle", "Bold")) > 0)
    .Italic = (Val(GetSetting(App.ProductName, "ViewStyle", "Italic")) > 0)
    .Underline = (Val(GetSetting(App.ProductName, "ViewStyle", "Underline")) > 0)
    .Strikethrough = (Val(GetSetting(App.ProductName, "ViewStyle", "Strikethrough")) > 0)
    .Name = GetSetting(App.ProductName, "ViewStyle", "Name")
    .Size = Val(GetSetting(App.ProductName, "ViewStyle", "Size"))
    If .Size = 0 Then .Size = 9
    End With
    With VS
    .ForeColor = Val(GetSetting(App.ProductName, "ViewStyle", "ForeColor"))
    .BackColor = Val(GetSetting(App.ProductName, "ViewStyle", "BackColor"))
    .LineHeight = Val(GetSetting(App.ProductName, "ViewStyle", "LineHeight"))
    If .LineHeight = 0 Then .LineHeight = 100
    End With
End Sub

Public Sub SaveViewerStyle(VS As ViewerStyle)
Dim a As Integer
With VS.Viewfont

    If .Bold Then a = 1 Else a = 0
    SaveSetting App.ProductName, "ViewStyle", "Bold", Str(a)
    If .Italic Then a = 1 Else a = 0
    SaveSetting App.ProductName, "ViewStyle", "Italic", Str(a)
    If .Underline Then a = 1 Else a = 0
    SaveSetting App.ProductName, "ViewStyle", "Underline", Str(a)
    If .Strikethrough Then a = 1 Else a = 0
    SaveSetting App.ProductName, "ViewStyle", "Strikethrough", Str(a)
    SaveSetting App.ProductName, "ViewStyle", "Name", .Name
    SaveSetting App.ProductName, "ViewStyle", "Size", Str(.Size)
End With
With VS
    SaveSetting App.ProductName, "ViewStyle", "ForeColor", Str(.ForeColor)
    SaveSetting App.ProductName, "ViewStyle", "Backcolor", Str(.BackColor)
    SaveSetting App.ProductName, "ViewStyle", "LineHeight", Str(.LineHeight)
End With
End Sub

Public Sub GetReaderStyle(RS As ReaderStyle)

'    With RS.ListFont
'    .Bold = (Val(GetSetting(App.ProductName, "ReaderStyle", "Bold")) > 0)
'    .Italic = (Val(GetSetting(App.ProductName, "ReaderStyle", "Italic")) > 0)
'    .Underline = (Val(GetSetting(App.ProductName, "ReaderStyle", "Underline")) > 0)
'    .Strikethrough = (Val(GetSetting(App.ProductName, "ReaderStyle", "Strikethrough")) > 0)
'    .Name = GetSetting(App.ProductName, "ReaderStyle", "Name")
'    .Size = Val(GetSetting(App.ProductName, "ReaderStyle", "Size"))
'    If .Size = 0 Then .Size = 9
'    End With
    
'    RS.FormBC = Val(GetSetting(App.ProductName, "ReaderStyle", "FormBackColor"))
    
    With RS.FormPOS
    .Height = Val(GetSetting(App.ProductName, "ReaderStyle", "FormHeight"))
    .Width = Val(GetSetting(App.ProductName, "ReaderStyle", "FormWidth"))
    .Top = Val(GetSetting(App.ProductName, "ReaderStyle", "FormTop"))
    .Left = Val(GetSetting(App.ProductName, "ReaderStyle", "FormLeft"))
    End With
    
    
    RS.LeftPos = Val(GetSetting(App.ProductName, "ReaderStyle", "LeftPos"))

 

End Sub

Public Sub SaveReaderStyle(RS As ReaderStyle)

'Dim a As Integer
'With RS.ListFont
'    If .Bold Then a = 1 Else a = 0
'    SaveSetting App.ProductName, "ReaderStyle", "Bold", Str(a)
'    If .Italic Then a = 1 Else a = 0
'    SaveSetting App.ProductName, "ReaderStyle", "Italic", Str(a)
'    If .Underline Then a = 1 Else a = 0
'    SaveSetting App.ProductName, "ReaderStyle", "Underline", Str(a)
'    If .Strikethrough Then a = 1 Else a = 0
'    SaveSetting App.ProductName, "ReaderStyle", "Strikethrough", Str(a)
'    SaveSetting App.ProductName, "ReaderStyle", "Name", .Name
'    SaveSetting App.ProductName, "ReaderStyle", "Size", Str(.Size)
'End With

'
'With RS
'    SaveSetting App.ProductName, "ReaderStyle", "FormBackColor", Str(.FormBC)
'End With

With RS.FormPOS
    SaveSetting App.ProductName, "ReaderStyle", "FormHeight", Str(.Height)
    SaveSetting App.ProductName, "ReaderStyle", "FormWidth", Str(.Width)
    SaveSetting App.ProductName, "ReaderStyle", "FormTop", Str(.Top)
    SaveSetting App.ProductName, "ReaderStyle", "FormLeft", Str(.Left)
End With

    SaveSetting App.ProductName, "ReaderStyle", "ListHeight", Str(RS.LeftPos)
 

End Sub

Public Sub GetFileFilter(FFilter() As String, ffnum As Integer)
Dim tempstr As String
For i = 0 To MAXFILTER
tempstr = GetSetting(App.ProductName, "FileFilter", "F" + Str(i))
If tempstr <> "" Then
 Dim pos As Integer
    pos = InStr(tempstr, "|")
    If pos > 0 Then
    FFilter(ffnum, 0) = Left(tempstr, pos - 1)
    FFilter(ffnum, 1) = Right(tempstr, Len(tempstr) - pos)
    ffnum = ffnum + 1
    End If
End If
Next

End Sub

Public Sub SaveFileFilter(FFilter() As String, ffnum As Integer)
Dim tempstr As String

DeleteSetting App.ProductName, "FileFilter"
For i = 0 To ffnum - 1
tempstr = FFilter(i, 0) + "|" + FFilter(i, 1)
SaveSetting App.ProductName, "FileFilter", "F" + Str(i), tempstr
Next
End Sub
