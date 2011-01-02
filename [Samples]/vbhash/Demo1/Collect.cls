VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCollect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' collection class wrapper

Private Node As Collection

Public Sub Insert(ByVal KeyVal, ByRef RecVal As Variant)
'   inputs:
'       KeyVal                key of node to insert
'       RecVal                record associated with key
'   action:
'       Inserts record RecVal with key KeyVal.
'   error:
'       errDuplicateKey
'
    On Error GoTo Handler
    Node.Add RecVal, CStr(KeyVal)
    Exit Sub
Handler:
    If Err.Number = 457 Then Raise errDuplicateKey, "CCollect.Insert"
    Err.Raise Number:=Err.Number
End Sub

Public Sub Delete(ByVal KeyVal As Variant)
'   inputs:
'       KeyVal                key of node to delete
'   action:
'       Deletes record with key KeyVal.
'   error:
'       errKeyNotFound
'
    On Error GoTo Handler
    Node.Remove CStr(KeyVal)
    Exit Sub
Handler:
    If Err.Number = 5 Then Raise errKeyNotFound, "CCollect.Delete"
    Err.Raise Number:=Err.Number
End Sub

Public Function Find(ByVal KeyVal As Variant) As Variant
'   inputs:
'       KeyVal                key of node to delete
'   returns:
'       record associated with key
'   action:
'       Finds record with key KeyVal, and returns record in RecVal.
'   errors:
'       errKeyNotFound
'
    On Error GoTo Handler
    Find = Node(CStr(KeyVal))
    Exit Function
Handler:
    If Err.Number = 5 Then Raise errKeyNotFound, "CCollect.Find"
    Err.Raise Number:=Err.Number
End Function

Private Sub Class_Initialize()
    Set Node = New Collection
End Sub
