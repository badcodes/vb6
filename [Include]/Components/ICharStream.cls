VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ICharStream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "A vb Interface of char stream made by LiN"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Public Property Get IsOver() As Boolean
End Property
Public Property Get Length() As Long
End Property
Public Property Get CharLeft() As Long
End Property
Public Property Get CharRead() As Long
End Property
Public Function Read(Optional charCount As Long = 1) As String
End Function
Public Function ReadUntil(ByRef charEnd As String) As String
End Function
Public Sub Skip(Optional ByVal charCount As Long = 1)
End Sub
Public Sub SkipUntil(ByRef charEnd As String)
End Sub
Public Function OpenStream(streamName As String) As Boolean
End Function
Public Function CloseStream() As Boolean
End Function

Public Function ReadAll() As String
End Function

