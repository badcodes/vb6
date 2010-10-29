VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFileStreamIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements ICharStream
Private mOpened As Boolean
Private mCharLeft As Long
Private mCharRead As Long
Private mCharLength As Long
Private mFNUM As Long
'Private Const BUFSIZE = 500
'Private mBufPos As Long
'Private mBufSize As Long
'Private mStrStream As String

Private Property Get ICharStream_CharLeft() As Long
ICharStream_CharLeft = mCharLeft
End Property

Private Property Get ICharStream_CharRead() As Long
ICharStream_CharRead = mCharRead
End Property

Private Function ICharStream_CloseStream() As Boolean
'<EhHeader>
On Error GoTo ICharStream_CloseStream_EXIT
'</EhHeader>
ICharStream_CloseStream = False
If Not mOpened Then Exit Function
Close mFNUM
mOpened = False
ICharStream_CloseStream = True
'<EhFooter>
Exit Function
ICharStream_CloseStream_EXIT:
'</EhFooter>
End Function

Private Property Get ICharStream_IsOver() As Boolean
If mCharRead = mLength Then ICharStream_IsOver = True
End Property

Private Property Get ICharStream_Length() As Long
ICharStream_Length = mLength
End Property

Private Function ICharStream_OpenStream(streamName As String) As Boolean
Call ICharStream_CloseStream
If linvblib.FileExists(streamName) = False Then Exit Function
mFNUM = FreeFile()
Open streamName For Input As #mFNUM
mLength = LOF(mFNUM)
mCharRead = 0
mCharLeft = mLength
'mBufSize = 0
'mBufPos = 0
'mStrStream = ""
mOpened = True
'Call ReadBuffer
End Function
'Private Function BufferEmpty() As Boolean
'If mBufPos > mBufSize Then BufferEmpty = True
'End Function
'Private Function ReadBuffer() As Boolean
'
'End Function

Private Function ICharStream_Read(Optional charCount As Long = 1&) As String
If charCount > mCharLeft Then charCount = mCharLeft
If charCount < 1 Then Exit Function
ICharStream_Read = Input$(charCount, mFNUM)
mCharRead = mCharRead + charCount
mCharLeft = mLength - mCharRead
End Function

Private Function ICharStream_ReadAll() As String
Dim curPos As Long
curPos = Seek(mFNUM)
Seek mFNUM, 1
ICharStream_ReadAll = StrConv(InputB(mLength - 1, mFNUM), vbUnicode)
Seek mFNUM, curPos
End Function

Private Function ICharStream_ReadUntil(charEnd As String) As String

    Dim CC As String
    If charEnd = "" Then Exit Function
    Do While Not EOF(mFNUM)
        CC = Input$(1, #mFNUM)
        mCharRead = mCharRead + 1
        If InStr(c, CC) > 0 Then
            UnRead 1
            Exit Do
        End If
        ICharStream_ReadUntil = ICharStream_ReadUntil & CC
    Loop
    mCharLeft = mLength - mCharRead
End Function

Private Sub ICharStream_Skip(Optional ByVal charCount As Long = 1&)
If charCount > mCharLeft Then charCount = mCharLeft
If charCount < 1 Then Exit Sub
charCount = charCount + Seek(mFNUM)
Seek mFNUM, charCount
End Sub

Private Sub ICharStream_SkipUntil(charEnd As String)

    Dim CC As String
    If charEnd = "" Then Exit Function
    Do While Not EOF(mFNUM)
        CC = Input$(1, #mFNUM)
        mCharRead = mCharRead + 1
        If InStr(c, CC) > 0 Then
            UnRead 1
            Exit Do
        End If
    Loop
    
    mCharLeft = mLength - mCharRead


End Sub

Private Sub UnRead(Optional lBytes As Long = 1)

    Dim lCurPos As Long
    Dim lSeekTo As Long
    lCurPos = Seek(mFNUM)
    lSeekTo = lCurPos - lBytes
    If lSeekTo >= 0 Then
        Seek mFNUM, lSeekTo
    End If

    mCharRead = mCharRead + lBytes
    mCharLeft = mLength - mCharRead
    

End Sub


