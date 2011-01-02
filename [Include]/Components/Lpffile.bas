Attribute VB_Name = "MLpffile"

Public Type lpfInfo   '12
    filecount As Long '4
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
    ViewFont As MYFont
    ViewBC As OLE_COLOR
    ViewFC As OLE_COLOR
End Type




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

Public Sub loadtable(m_info As lpfInfo, m_table() As lpfFileTable, tv As TreeView, cb As ComboBox)
Dim tempcata(100) As String
catanum = 0
Dim cataloged As Boolean

For i = 1 To m_info.filecount
tempstr = rdel(m_table(i).filename)
pos = InStr(tempstr, "\")
If pos > 0 Then
    thecata = Left(tempstr, pos - 1)
    thename = Right(tempstr, Len(tempstr) - pos)
    cataloged = False
    For c = 1 To catanum
    If thecata = tempcata(c) Then cataloged = True: Exit For
    Next
    If cataloged = False Then
    catanum = catanum + 1
    tempcata(catanum) = thecata
    tv.Nodes.Add , , thecata, thecata, 1
    End If
    tv.Nodes.Add thecata, tvwChild, "林晓然" + Str(i), thename, 2
Else
    tv.Nodes.Add , , "林晓然" + Str(i), tempstr
End If
cb.AddItem tempstr
Next
End Sub
