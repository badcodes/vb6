VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ILString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Property Get Length() As Long
End Property

Public Property Get VBString() As String
End Property

Public Property Let VBString(ByVal strValue As String)
End Property

Public Function SubString(ByVal lStartPos As Long, ByVal lLength As Long) As String
End Function

Public Function Append(ByVal strValue As String) As String
End Function

Public Function DelString(ByVal lStartPos As Long, ByVal lLength As Long) As String
End Function

Public Function Replace(Optional ByVal lStartPos As Long = 1, Optional ByRef strFind As String = "", Optional ByRef strReplace As String = "", Optional cmpMethod As VbCompareMethod = vbBinaryCompare) As String
End Function

