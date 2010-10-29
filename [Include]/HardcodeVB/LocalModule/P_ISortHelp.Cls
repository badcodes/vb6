VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ISortHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ISortHelper interface class

Function Compare(v1 As Variant, v2 As Variant) As Integer
End Function

Sub Swap(v1 As Variant, v2 As Variant)
End Sub

Sub CollectionSwap(n As Collection, _
                   i1 As Variant, _
                   i2 As Variant, _
                   Optional key1 As Variant, _
                   Optional key2 As Variant)
End Sub
'


