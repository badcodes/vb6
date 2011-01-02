Attribute VB_Name = "modZtmReader"

Type MYFont
    Bold As Boolean
    Italic As Boolean
    Charset As Integer
    name As String
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
End Type

Type ViewerStyle
    Viewfont As MYFont
    ForeColor As OLE_COLOR
    BackColor As OLE_COLOR
    LineHeight As Integer
End Type

Type ZtmSetting
ZtmLast As String
ZtmLastFile As String
End Type

Public Const ztmInfo = "ztm.nfo"





Public Sub getZtmSetting(ZtmS As ZtmSetting)
With ZtmS
.ZtmLast = GetSetting(App.ProductName, "LatelyUsed", "ztmFile")
.ZtmLastFile = GetSetting(App.ProductName, "LatelyUsed", "ztmLastFile")
End With
End Sub
Public Sub saveZtmSetting(ZtmS As ZtmSetting)
With ZtmS
SaveSetting App.ProductName, "LatelyUsed", "ztmFile", .ZtmLast
SaveSetting App.ProductName, "LatelyUsed", "ztmLastFile", .ZtmLastFile

End With
End Sub





Public Sub GetReaderStyle(RS As ReaderStyle)


    With RS.FormPOS
    .Height = Val(GetSetting(App.ProductName, "ReaderStyle", "FormHeight"))
    .Width = Val(GetSetting(App.ProductName, "ReaderStyle", "FormWidth"))
    .Top = Val(GetSetting(App.ProductName, "ReaderStyle", "FormTop"))
    .Left = Val(GetSetting(App.ProductName, "ReaderStyle", "FormLeft"))
    End With


End Sub

Public Sub SaveReaderStyle(RS As ReaderStyle)


With RS.FormPOS
    SaveSetting App.ProductName, "ReaderStyle", "FormHeight", Str(.Height)
    SaveSetting App.ProductName, "ReaderStyle", "FormWidth", Str(.Width)
    SaveSetting App.ProductName, "ReaderStyle", "FormTop", Str(.Top)
    SaveSetting App.ProductName, "ReaderStyle", "FormLeft", Str(.Left)
End With

 
End Sub
Public Sub GetViewerStyle(VS As ViewerStyle)
    With VS.Viewfont
    .Bold = (Val(GetSetting(App.ProductName, "ViewStyle", "Bold")) > 0)
    .Italic = (Val(GetSetting(App.ProductName, "ViewStyle", "Italic")) > 0)
    .Underline = (Val(GetSetting(App.ProductName, "ViewStyle", "Underline")) > 0)
    .Strikethrough = (Val(GetSetting(App.ProductName, "ViewStyle", "Strikethrough")) > 0)
    .name = GetSetting(App.ProductName, "ViewStyle", "Name")
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
    SaveSetting App.ProductName, "ViewStyle", "Name", .name
    SaveSetting App.ProductName, "ViewStyle", "Size", Str(.Size)
End With
With VS
    SaveSetting App.ProductName, "ViewStyle", "ForeColor", Str(.ForeColor)
    SaveSetting App.ProductName, "ViewStyle", "Backcolor", Str(.BackColor)
    SaveSetting App.ProductName, "ViewStyle", "LineHeight", Str(.LineHeight)
End With
End Sub

Public Sub CopytoIfont(srcfont As MYFont, dstfont As IFontDisp)

With dstfont
.Bold = srcfont.Bold
.Italic = srcfont.Italic
.name = srcfont.name
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
If srcfont.name = "" Then .FontName = "System" Else .FontName = srcfont.name
.FontSize = srcfont.Size
.FontStrikethru = srcfont.Strikethrough
.FontUnderline = srcfont.Underline
End With


End Sub

Public Function toMYfont(srcfont As IFontDisp) As MYFont
With toMYfont
.Bold = srcfont.Bold
.Italic = srcfont.Italic
.name = srcfont.name
.Size = srcfont.Size
.Strikethrough = srcfont.Strikethrough
.Underline = srcfont.Underline
.Charset = srcfont.Charset
End With
End Function



