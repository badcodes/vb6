VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CUniqueStringVector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const cst_HashTable_Size As Long = 256
Private Const CST_DEFAULT_CHUNKSIZE As Long = 128
Private Const CST_DEFAULT_SIZE As Long = 128

Private mTableSize As Long

Dim mData() As String
Private mCount As Long
Private mSize As Long
Private mChunkSize As Long
Private mHashTableSize As Long
Private mHashTable() As CLongLink

Public Property Get HashTableSize() As Long
    HashTableSize = mHashTableSize
End Property

Public Property Let HashTableSize(ByVal LonValue As Long)
    If LonValue < 0 Then LonValue = 0
    If LonValue = mHashTableSize Then Exit Property
    mHashTableSize = LonValue
    If mHashTableSize > 0 Then
        ReDim mHashTable(0 To mHashTableSize - 1)
        If mCount > 0 Then ReHashData
    End If
End Property
Private Function HashTableCompute(ByRef vString As String) As Long
    If mHashTableSize = 0 Then HashTableCompute = -1: Exit Function
    Dim iKey As Long
    iKey = 0
    Dim bKey() As Byte
    bKey = StrConv(vString, vbFromUnicode)
    Dim i As Long
    For i = LBound(bKey) To UBound(bKey)
        iKey = iKey + bKey(i)
    Next
    iKey = iKey Mod mHashTableSize
    HashTableCompute = iKey
End Function

Private Sub HashTableAdd(ByVal vHashIdx, ByVal vIndex As Long)
    If vHashIdx >= 0 Then
        Dim c As CLongLink
        Set c = mHashTable(vHashIdx)
        If c Is Nothing Then
            Set c = New CLongLink
            c.Data = vIndex
        Else
            Do Until c.NextLink Is Nothing
                Set c = c.NextLink
            Loop
            Set c.NextLink = New CLongLink
            c.NextLink.Data = vIndex
        End If
    End If
End Sub

Public Sub ReHashData()
    If mHashTableSize < 0 Then Exit Sub
    Dim i As Long
    Dim pHash As Long
    For i = 0 To mCount - 1
        pHash = HashTableCompute(mData(i))
        HashTableAdd pHash, i
    Next
End Sub

Private Sub Class_Initialize()
    mChunkSize = CST_DEFAULT_CHUNKSIZE
    HashTableSize = cst_HashTable_Size
End Sub

Public Sub Reset()
    mCount = 0
    ReDim mHashTable(0 To mHashTableSize - 1)
End Sub

Public Property Get Count() As Long
    Count = mCount
End Property

Public Function ItemOf(ByRef nIndex As Long) As String
    If nIndex < 0 Or nIndex > mCount - 1 Then Exit Function
    Item = mData(nIndex)
End Function

Public Function IndexOf(ByRef vString As String) As Long
    Dim pHash As Long
    pHash = HashTableCompute(vString)
    IndexOf = SearchIndex(pHash, vString)
End Function

Private Function SearchIndex(ByVal vHashIdx As Long, ByRef vItem As String) As Long
SearchIndex = -1
If vHashIdx > 0 Then
        Dim c As CLongLink
        Set c = mHashTable(vHashIdx)
        Do Until c Is Nothing
            If mData(c.Data) = vItem Then SearchIndex = c.Data: Exit For
            Set c = c.NextLink
        Loop
Else
        Dim i As Long
        For i = 0 To mCount - 1
            If mData(i) = vItem Then SearchIndex = i: Exit For
        Next
End If
End Function

Public Sub Add(ByRef vItem As String)
    Dim pHash As Long
    Dim nIndex As Long
    pHash = HashTableCompute(vItem)
    nIndex = SearchIndex(pHash, vItem)
    If nIndex > 0 Then
        Exit Sub
    Else
        nIndex = mCount
        If nIndex > mSize - 1 Then EnlargeData nIndex
        mData(nIndex) = vItem
        mCount = mCount + 1
        HashTableAdd pHash, nIndex
    End If
End Sub

Private Sub EnlargeData(ByRef nMinBound As Long)
    mSize = nMinBound + 1 + mChunkSize
    If mCount = 0 Then
        ReDim mData(0 To mSize - 1) As String
    Else
        ReDim Preserve mData(0 To mSize - 1) As String
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


Public Sub Init(Optional ByRef nSize As Long = -1, Optional nChunkSize As Long = -1)
    mCount = 0
    If (nChunkSize > 0) Then mChunkSize = nChunkSize
    If (nSize >= 0) Then mSize = nSize Else mSize = CST_DEFAULT_SIZE
    If (nSize = 0) Then Exit Sub
    ReDim mData(0 To mSize - 1)
    If mHashTableSize > 0 Then ReDim mHashTable(0 To mHashTableSize - 1)
End Sub

Public Property Get ChunkSize() As Long
    rfChunkSize = mChunkSize
End Property

Public Property Let ChunkSize(ByRef nSize As Long)
    If nSize < 1 Then Exit Property
    mChunkSize = nSize
End Property




