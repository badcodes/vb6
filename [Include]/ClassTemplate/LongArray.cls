VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLongArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Template header:
'================================================================================

#Const ObjectType = 1
#Const NormalType = 2
#Const VariantType = (ObjectType Or NormalType)
#Const DefaultType = VariantType

#Const fLong = NormalType


'================================================================================


Private Const DEFAULT_CHUNKSIZE As Long = 128 
Private Const DEFAULT_SIZE As Long = 128 

Dim mData() As Long 
Private mCount As Long 
Private mSize As Long 
Private mChunkSize As Long 
Private mLIndex As Long 
Private mUIndex As Long 

Public Property Get UIndex() As Long 
    UIndex = mCount - 1
End Property

Public Property Get LIndex() As Long 
    LIndex = 0
End Property

Private Sub Class_Initialize()
    mChunkSize = DEFAULT_CHUNKSIZE
End Sub

Public Sub Reset()
    mCount = 0
End Sub

Public Property Get Count() As Long 
    Count = mCount
End Property

Public Property Get Item(ByRef nIndex As Long) As Long 
Attribute Item.VB_UserMemId = 0
    If nIndex < 0 Or nIndex > mCount - 1 Then Exit Property
    Item = mData(nIndex) 
End Property

#If fLong And ObjectType Then 
Public Property Set Item(ByRef nIndex As Long, ByRef vItem As Long) 
    If nIndex < 0 Then Exit Property
    If nIndex > mSize - 1 Then EnlargeData nIndex
    mData(nIndex) = vItem 
    If (nIndex >= mCount) Then mCount = nIndex + 1
End Property
#End If


#If fLong And NormalType Then 
Public Property Let Item(ByRef nIndex As Long, ByRef vItem As Long) 
    If nIndex < 0 Then Exit Property
    If nIndex > mSize - 1 Then EnlargeData nIndex
    mData(nIndex) = vItem 
    If (nIndex >= mCount) Then mCount = nIndex + 1
End Property
#End If

Public Sub Add(ByRef vItem As Long) 
    Dim nIndex As Long 
    nIndex = mCount
    If nIndex > mSize - 1 Then EnlargeData nIndex
    mData(nIndex) = vItem 
    mCount = mCount + 1
End Sub

Private Sub EnlargeData(ByRef nMinBound As Long) 
    mSize = nMinBound + 1 + mChunkSize
    If mCount = 0 Then
        ReDim mData(0 To mSize - 1) As Long 
    Else
        ReDim Preserve mData(0 To mSize - 1) As Long 
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
    If (nChunkSize > 0) Then mChunkSize = nChunkSize
    If (nSize >= 0) Then mSize = nSize Else mSize = DEFAULT_SIZE
    If (nSize = 0) Then Exit Sub
    ReDim mData(0 To mSize - 1)
End Sub

Public Property Get ChunkSize() As Long 
    rfChunkSize = mChunkSize
End Property

Public Property Let ChunkSize(ByRef nSize As Long) 
    If nSize < 1 Then Exit Property
    mChunkSize = nSize
End Property

Public Sub Insert(ByVal nIndex As Long, ByRef vItem As Long) 
    If nIndex < 0 Then nIndex = 0
    If nIndex > mSize - 1 Then EnlargeData nIndex
    
    If nIndex > mCount - 1 Then
        mData(nIndex) = vItem 
        mCount = nIndex + 1
    Else
        Dim i As Long 
        Dim iPrev As Long 
        For i = nIndex + 1 To mCount
            iPrev = i - 1
            mData(iPrev) = mData(i) 
        Next
        mData(nIndex) = vItem 
        mCount = mCount + 1
    End If
End Sub

Public Sub Remove(ByVal nIndex As Long, Optional ByVal nCount As Long = 1) 
    If nIndex < 0 Then Exit Sub
    If nIndex > mCount - 1 Then Exit Sub
    If nCount = 0 Then Exit Sub
    
    Dim nEnd As Long 
    If nCount < 0 Then
        nEnd = nIndex + nCount + 1
    Else
        nEnd = nIndex + nCount - 1
    End If
    If nEnd < 0 Then nEnd = 0
    If nEnd > mCount - 1 Then nEnd = mCount - 1
    
    Dim nStart As Long 
    nStart = nIndex
    If nStart > nEnd Then
        nStart = nEnd
        nEnd = nIndex
    End If
    
    'Reuse Varaibles nCount,nIndex
    nCount = nEnd - nStart + 1
    If nCount = 0 Then Exit Sub
    
     
    Dim iPrev As Long 
    For nIndex = nEnd + 1 To mCount - 1
        iPrev = nIndex - nCount
        mData(iPrev) = mData(nIndex) 
         
    Next
    mCount = mCount - nCount
    
End Sub

Public Function Find(ByRef vItem As Long, Optional ByVal nStart As Long = 0) As Long 
    Find = -1
    If nStart < 0 Then nStart = 0
    Dim i As Long 
    For i = nStart To mCount - 1
        If (mData(i) = vItem) Then Find = i: Exit Function 
    Next
End Function

Public Function ToArray() As Long() 
    If mCount < 1 Then Exit Function
    Dim aResult() As Long 
    ReDim aResult(0 To mCount - 1) As Long 
    Dim i As Long 
    For i = 0 To mCount - 1
        aResult(i) = mData(i) 
    Next
    ToArray = aResult
End Function

Private Sub Class_Terminate()
    
     
    
End Sub
