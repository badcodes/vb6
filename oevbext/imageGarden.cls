VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "imageGarden"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Sub downFullAlbum(cataid As Integer, albumid As Integer, Optional sComment As String)
'Example:  "http://www.imagegarden.net/image.php?from=white+paper&cataid=2&albumid=2234&imageid=1&image=jpeg"
Dim headUrl As String
Dim tailUrl As String
Dim i As Integer
Dim tmpFile As String
Dim fnum As Integer
Dim link As String
Dim iOE As OELib.OfflineExplorerAddUrl
Dim imageCount As Integer

imageCount = InputBox("该专辑共有多少图片?", "输入")
If imageCount < 1 Then Exit Sub

If sComment = "" Then sComment = "cataid" & intToStr(cataid) & "-" & "albumid" & intToStr(albumid)

tmpFile = linvblib.bddir(App.Path) & linvblib.cleanFilename(sComment) & ".htm"
fnum = FreeFile
Open tmpFile For Output As #fnum

headUrl = "http://www.imagegarden.net/image.php?from=white+paper"
tailUrl = "&image=jpeg"

headUrl = headUrl & "&cataid=" & intToStr(cataid) _
             & "&albumid=" & intToStr(albumid)
             
Print #fnum, "<html><body>"
For i = 1 To imageCount
link = headUrl & "&imageid=" & intToStr(i) & tailUrl
Print #fnum, "<a href='" & link & "'>" & "ImageId" & intToStr(i) & "</a><br>"
Next
Print #fnum, "</body></html>"
Close #fnum

Set iOE = New OELib.OfflineExplorerAddUrl
iOE.AddUrl "file:///" & tmpFile, sComment, headUrl
Set iOE = Nothing

End Sub

Private Function intToStr(n As Integer) As String
intToStr = LTrim$(CStr(n))
End Function
