VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CNull"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private iProp As Integer
Public PubProp As Integer
Public StrProp As String

Public Property Let ProcProp(iPropA As Integer)
    iProp = iPropA
End Property

Public Property Get ProcProp() As Integer
    ProcProp = iProp
End Property

Public Sub SubMethod(i As Integer)
    i = 1
End Sub

Public Function FuncMethod()
    FuncMethod = 2
End Function

Friend Property Get FriendProp() As Long
    FriendProp = 1
End Property

