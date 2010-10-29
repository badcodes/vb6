VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CStringBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
      (ByVal pDst As Long, ByVal pSrc As Long, ByVal ByteLen As Long)
      

Const CONST_DEFAULT_CHUNK_SIZE As Long = 2048

Private mBuffer As String
Private mChunkSize As Long
Private mCount As Long
Private mSize As Long
Private mLPBuffer As Long

Public Sub AppendLine(ByRef sSrc As String)
    Dim nLen As Long
    nLen = PrepareSpaceFor(sSrc, 2)
    If nLen > 0 Then
        CopyMemory ByVal mLPBuffer + mCount + mCount, ByVal StrPtr(sSrc), nLen + nLen
        mCount = mCount + nLen
    End If
    CopyMemory ByVal mLPBuffer + mCount + mCount, ByVal StrPtr(vbCrLf), 4
    mCount = mCount + 2
End Sub

Public Sub AppendLineV(ByVal sStr As String)
    AppendLine sStr
End Sub

Public Sub Append(ByRef sSrc As String)
    Dim nLen As Long
    nLen = PrepareSpaceFor(sSrc)
    If nLen > 0 Then
        CopyMemory ByVal StrPtr(mBuffer) + mCount * 2, ByVal StrPtr(sSrc), nLen + nLen
        mCount = mCount + nLen
    End If
End Sub

Public Sub AppendV(ByVal sSrc As String)
    Append sSrc
End Sub
      
Private Function PrepareSpaceFor(ByRef sIncome As String, Optional nAdd As Long = 0) As Long
    
    PrepareSpaceFor = Len(sIncome)
    If (PrepareSpaceFor + mCount + nAdd > mSize) Then
        mSize = PrepareSpaceFor + mCount + mChunkSize + nAdd
        If mCount > 0 Then
            Dim sTemp As String
            sTemp = String$(mCount, vbNull)
            CopyMemory ByVal StrPtr(sTemp), ByVal mLPBuffer, mCount + mCount
            mBuffer = String$(mSize, vbNull)
            mLPBuffer = StrPtr(mBuffer)
            CopyMemory ByVal mLPBuffer, ByVal StrPtr(sTemp), mCount + mCount
        Else
            mBuffer = String$(mSize, vbNull)
            mLPBuffer = StrPtr(mBuffer)
        End If
        
    End If
    
End Function

Public Property Let ChunkSize(ByRef nChunkSize As Long)
    If nChunkSize < 1 Then
        Err.Raise 1 + vbObjectError + 512, , "Invalid integer value"
    End If
    mChunkSize = nChunkSize
End Property

Public Property Get ChunkSize() As Long
    ChunkSize = mChunkSize
End Property
Private Sub Class_Initialize()
    mChunkSize = CONST_DEFAULT_CHUNK_SIZE
End Sub

Public Sub Clear()
    mCount = 0
End Sub

Public Function Value() As String
Attribute Value.VB_UserMemId = 0
    If mCount > 0 Then Value = Left$(mBuffer, mCount)
End Function

Public Property Get Length() As Long
    Length = mCount
End Property

