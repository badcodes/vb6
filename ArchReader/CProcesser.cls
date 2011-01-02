VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CProcesser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function OpenFile(ByRef path As String) As CArchive
End Function

Public Function OpenFolder(ByRef path As String) As CArchive
End Function

Public Function OpenUrl(ByRef URL As String) As CArchive
End Function

Public Function OpenArchive(ByRef archive As CArchive, ByRef entry As String) As String
End Function

Public Function GetUrlInfo(ByRef URL As String) As CURLInfo
End Function

Public Function CanHandlerURL(ByRef URL As String) As Boolean
End Function

