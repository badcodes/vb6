Attribute VB_Name = "ZIPReaderPUBLIC"
Public NumToNodeIndex() As Integer
Public NodeIndexToNum() As Integer
Public FileCount As Integer
Public FolderCount As Integer
Public Filezip As Boolean
Public FileIe As Boolean

Public Type lpfInfo   '12
    FileCount As Long '4
    adrfiletable As Long '4
    adrfiledata As Long '4
End Type
Public Type lpfFileTable '264
    filename As String * 256 '256
    shortname As String * 256 '256
    adrfilebegin As Long '4
    filesize As Long '4
End Type

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
    FormBC As OLE_COLOR
    FormPOS As MYPoS
    ListFont As MYFont
    ListPos As MYPoS
End Type
Public Const BACKWARDVIEW = -1
Public Const FORWARDVIEW = -2
Public Const PREVIOUSVIEW = -3
Public Const NEXTVIEW = -4
Public Const RANDOMVIEW = -100




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
