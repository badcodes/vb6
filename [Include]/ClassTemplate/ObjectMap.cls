VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CObjectMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Template header:
'================================================================================

#Const ObjectType = 1
#Const NormalType = 2
#Const VariantType = (ObjectType Or NormalType)
#Const DefaultType = VariantType

#Const fObject = ObjectType


'================================================================================


Private Const DEFAULT_CHUNKSIZE As Long = 128
Private Const DEFAULT_SIZE As Long = 128

Dim mData() As Object 
Private mCount As Long
Private mSize As Long
Private mChunkSize As Long
Private mIndexTable As CIndexTable

Private Sub Class_Initialize()
    mChunkSize = DEFAULT_CHUNKSIZE
    Set mIndexTable = New CIndexTable
End Sub

Public Sub Reset()
    mCount = 0
End Sub

Public Property Get Count() As Long
    Count = mCount
End Property

Public Sub Add(ByRef vItem As Object) 
    Dim nIndex As Long
    nIndex = mCount
    If nIndex > mSize - 1 Then EnlargeData nIndex
    Set mData(nIndex) = vItem 
    mCount = mCount + 1
End Sub

Private Sub EnlargeData(ByRef nMinBound As Long)
    mSize = nMinBound + 1 + mChunkSize
    If mCount = 0 Then
        ReDim mData(0 To mSize - 1) As Object 
    Else
        ReDim Preserve mData(0 To mSize - 1) As Object 
    End If
End Sub


Public Sub Shink(Optional ByRef nCount As Long = -1)
    If nCount = -1 Then nCount = mCount
    If (nCount < 0) Then Exit Sub
    mSize = nCount
    If (mSize = 0) Then
        Erase mData
    Else
        ReDim Preserve mData(0 To mSize - 1)
    End If
    mCount = mSize
End Sub

Public Property Get ChunkSize() As Long
    rfChunkSize = mChunkSize
End Property

Public Property Let ChunkSize(ByRef nSize As Long)
    If nSize < 1 Then Exit Property
    mChunkSize = nSize
End Property


Public Sub Remove(ByVal vKey As String)
    Dim idx As Long
    idx = mIndexTable.Remove(vKey)
    If idx >= 0 Then
        Set mData(idx) = Nothing 
    End If
End Sub


Public Sub Map(ByVal vKey As String, ByRef vItem As Object) 
    Dim idx As Long
    idx = mIndexTable.Add(vKey)
    If idx >= mCount Then
        mCount = mCount + 1
    End If
    
    If idx >= mSize Then
        EnlargeData idx
    End If
    Set mData(idx) = vItem 
End Sub
#If fObject = NormalType Or fObject = VariantType Then 
Public Property Let Item(ByVal vKey As String, ByRef vItem As Object) 
    Map vKey, vItem
End Property
#End If
#If fObject = ObjectType Or fObject = VariantType Then 
Public Property Set Item(ByVal vKey As String, ByRef vItem As Object) 
    Map vKey, vItem
End Property
#End If
Public Property Get Item(ByVal vKey As String) As Object 
Attribute Item.VB_UserMemId = 0
    Dim idx As Long
    idx = mIndexTable.IndexOf(vKey)
    If idx >= 0 Then
        Set Item = mData(idx) 
    End If
End Property
Public Function ValueOf(ByRef vKey As String) As Object 
    Dim idx As Long
    idx = mIndexTable.IndexOf(vKey)
    If idx >= 0 Then
        Set ValueOf = mData(idx) 
    End If
End Function
Public Function KeyOf(ByRef vValue As Object) As String 
    Dim i As Long
    For i = 0 To mCount - 1
        If (mData(i) Is vValue) = True Then 
            KeyOf = mIndexTable.ItemOf(i)
        End If
    Next
End Function

Public Property Get Keys() As String()
    Keys = mIndexTable.Keys
End Property

Public Property Get Values() As Object() 
    If mCount < 1 Then Exit Property
    ReDim result(0 To mCount - 1) As Object 
    Dim i As Long
    For i = 0 To mCount - 1
        Set result(i) = mData(i) 
    Next
    Values = result
End Property
