VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CStringStack"
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

#Const fString = NormalType


'================================================================================


Private Const DEFAULT_CHUNKSIZE As Long = 128
Private Const DEFAULT_SIZE As Long = 128

Dim mData() As String
Private mCount As Long
Private mSize As Long
Private mChunkSize As Long


Private Sub Class_Initialize()
    mChunkSize = DEFAULT_CHUNKSIZE
    mSize = 0
End Sub

Public Sub Reset()
    mCount = 0
End Sub

Public Property Get Count() As Long
    Count = mCount
End Property

Public Sub Push(ByRef vItem As String)
    Dim nIndex As Long
    nIndex = mCount
    If nIndex > mSize - 1 Then EnlargeData nIndex
    mData(nIndex) = vItem
    mCount = mCount + 1
End Sub

Public Function Pop() As String
    If mCount < 1 Then Exit Function
    mCount = mCount - 1
    Pop = mData(mCount)
     
End Function

Public Function Peek(Optional vIndex As Long = -1) As String
    If vIndex < 0 Then
        vIndex = mCount - 1
    End If
    If vIndex < 0 Then Exit Function
    Peek = mData(vIndex)
End Function



Private Sub EnlargeData(ByRef nMinBound As Long)
    mSize = nMinBound + 1 + mChunkSize
    If mCount = 0 Then
        ReDim mData(0 To mSize - 1) As String
    Else
        ReDim Preserve mData(0 To mSize - 1) As String
    End If
End Sub

Public Sub Init(Optional ByRef nSize As Long = -1, Optional nChunkSize As Long = -1)
    If (nChunkSize > 0) Then mChunkSize = nChunkSize
    If (nSize > 0) Then mSize = nSize Else mSize = DEFAULT_SIZE
    ReDim mData(0 To mSize - 1)
End Sub

Public Property Get ChunkSize() As Long
    rfChunkSize = mChunkSize
End Property

Public Property Let ChunkSize(ByRef nSize As Long)
    If nSize < 1 Then Exit Property
    mChunkSize = nSize
End Property

