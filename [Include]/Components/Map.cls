VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private mData As Collection
Private mKeys As Collection

Public Property Get Value(ByRef vKey As Variant) As Variant
    On Error Resume Next
    Value = mData.Item(vKey)
End Property

Public Property Let Value(ByRef vKey As Variant, ByRef vValue As Variant)
    mData.Add vValue, vKey
    mKeys.Add vKey, vKey
End Property

Public Property Set Value(ByRef vKey As Variant, ByRef vValue As Variant)
    mData.Add vValue, vKey
    mKeys.Add vKey, vKey
End Property

Public Function Exists(ByRef vKey As Variant) As Boolean
    On Error GoTo NotExists
    mData.Item vKey
    Exists = True
NotExists:
    Exists = False
End Function


Public Sub Delete(ByRef vKey As Variant)
    On Error Resume Next
    mData.Remove vKey
    mKeys.Remove vKey
End Sub

Public Sub Clear()
    Set mData = New Collection
    Set mKeys = New Collection
End Sub


Public Function Count() As Long
    Count = mData.Count
End Function
Public Function Values() As Collection
    Set Values = mData
End Function

Public Function Keys() As Collection
    Set Keys = mKeys
End Function
Private Sub Class_Initialize()
    Set mData = New Collection
    Set mKeys = New Collection
End Sub

