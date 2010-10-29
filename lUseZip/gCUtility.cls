VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "gCUtility"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Function CleanZipFilename(sInCome) As String

Dim sFilenameClean As String
Dim iPos As Long, iStart As Long, iEnd As Long
Dim charNow As String
iEnd = Len(sInCome)
iStart = 1
For iPos = iStart To iEnd
    charNow = Mid$(sInCome, iPos, 1)
    Select Case charNow
        Case "\"
            sFilenameClean = sFilenameClean & "/"
        Case "["
            sFilenameClean = sFilenameClean & "[[]"
       ' Case "]"
           ' sFilenameClean = sFilenameClean & "[]]"
        Case Else
            sFilenameClean = sFilenameClean & charNow
    End Select
 Next

CleanZipFilename = sFilenameClean

End Function

Public Function CBytesToStr(ByRef CBytes() As Byte) As String
    CBytesToStr = mShareFunction.CBytesToStr(CBytes())
End Function
Public Sub StrToCBytes(ByVal strUnicode As String, ByRef CBytes() As Byte)
    Call mShareFunction.StrToCBytes(strUnicode, CBytes())
End Sub
