Attribute VB_Name = "mMusic"
Option Explicit
Function getInfo_VeryCD(ByVal SArchiveName As String, ByRef sArtist As String, ByRef sAlbum As String) As Boolean
SArchiveName = CleanFileName(SArchiveName)
sArtist = linvblib.LeftLeft(SArchiveName, " - ", vbTextCompare, ReturnEmptyStr)
If sArtist = "" Then sArtist = linvblib.LeftLeft(SArchiveName, "-", vbTextCompare, ReturnEmptyStr)
sAlbum = linvblib.LeftRange(SArchiveName, "[", "]", vbTextCompare, ReturnEmptyStr)
If sAlbum = "" Then sAlbum = linvblib.LeftRange(SArchiveName, " - ", " - ", vbTextCompare, ReturnEmptyStr)
If sAlbum = "" Then sAlbum = linvblib.LeftRange(SArchiveName, "-", "-", vbTextCompare, ReturnEmptyStr)


Debug.Print sArtist
Debug.Print sAlbum
If sArtist <> "" And sAlbum <> "" Then getInfo_VeryCD = True
End Function
Function CleanFileName(SArchiveName As String) As String
CleanFileName = Replace(SArchiveName, ".", " ")
CleanFileName = Replace(CleanFileName, "_", " ")
End Function
