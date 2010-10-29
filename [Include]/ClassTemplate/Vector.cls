VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CVector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Implements CArray
'Template header:
'================================================================================

#Const ObjectType = 1
#Const NormalType = 2
#Const VariantType = (ObjectType Or NormalType)
#Const DefaultType = VariantType

#Const fVariant = VariantType


'================================================================================

Private mData As CArray

Public Sub Add(vItem As Variant) 
    mData.Add vItem
End Sub

Public Property Let ChunkSize(RHS As Long)
    mData.ChunkSize = RHS
End Property

Public Property Get ChunkSize() As Long
    ChunkSize = mData.ChunkSize
End Property

Public Property Get Count() As Long
    Count = mData.Count
End Property

Public Function Find(vItem As Variant, Optional ByVal nStart As Long = 0&) As Long 
    Find = mData.Find(vItem, nStart)
End Function

Public Sub Init(Optional nSize As Long = -1&, Optional nChunkSize As Long = -1&)
    mData.Init nSize, nChunkSize
End Sub

Public Sub Insert(ByVal nIndex As Long, vItem As Variant) 
    mData.Insert nIndex, vItem
End Sub
#If fVariant And ObjectType Then 
Public Property Set Item(nIndex As Long, RHS As Variant) 
    Set mData.Item(nIndex) = RHS
End Property
#End If
#If fVariant And NormalType Then 
Public Property Let Item(nIndex As Long, RHS As Variant) 
    mData.Item(nIndex) = RHS
End Property
#End If

Public Property Get Item(nIndex As Long) As Variant 
    If IsObject(mData.Item(nIndex))  then Set Item = mData.Item(nIndex) Else Item = mData.Item(nIndex) 
End Property

Public Sub Remove(ByVal nIndex As Long, Optional ByVal nCount As Long = 1&)
    mData.Remove nIndex, nCount
End Sub

Public Sub Reset()
    mData.Reset
End Sub

Public Sub Shink(Optional nCount As Long = -1&)
    mData.Shink nCount
End Sub


Public Sub Class_Initialize()
    Set mData = New CArray
End Sub

Private Sub Class_Terminate()
    Set mData = Nothing
End Sub
