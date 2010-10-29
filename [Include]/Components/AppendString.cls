VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CAppendString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
''This is version 2. 2001-11-24
''Changes:
''- Twice as fast at least.
''- Added a cAppendString class instead of using Mid$
''- Now uses Regular Expressions to color the 'inside' of tags (Thanks to Gary aka RegX on PSC)
''- Added a RT control to speed up viewing

''CREDITS AND COPYRIGHT:
''MartijnB (bambi@crackdealer.com)
''Money, rewards, bugs and .. can be send to me (please!)

Option Explicit
Option Compare Binary

Private Declare Sub RtlMoveMemory Lib "kernel32" (dst As Any, src As Any, ByVal nBytes&)
Private Declare Function SysAllocStringByteLen& Lib "oleaut32" (ByVal olestr&, ByVal BLen&)

Private plngStringLen   As Long
Private plngBufferLen   As Long
Private pstrBuffer      As String

Public Sub Append(Text As String)

    Dim lngText          As Long
    Dim strTemp          As String
    Dim lngVPointr       As Long
    lngText = Len(Text)

    If lngText > 0 Then

        If (plngStringLen + lngText) > plngBufferLen Then
            plngBufferLen = (plngStringLen + lngText) * 2&
            strTemp = AllocString04(plngBufferLen)
            '***  copymemory might be faster than this
            Mid$(strTemp, 1&) = pstrBuffer
            '***  Alternate pstrBuffer = strTemp
            '***  switch pointers instead of slow =
            lngVPointr = StrPtr(pstrBuffer)
            RtlMoveMemory ByVal VarPtr(pstrBuffer), ByVal VarPtr(strTemp), 4&
            RtlMoveMemory ByVal VarPtr(strTemp), lngVPointr, 4&
            '  Debug.Print "plngBufferLen: " & plngBufferLen
        End If

        Mid$(pstrBuffer, plngStringLen + 1&) = Text
        plngStringLen = plngStringLen + lngText
    End If

End Sub



Public Sub AppendLine(Text As String)

    Append Text
    Append vbCrLf

End Sub

Public Property Get Value() As String
Attribute Value.VB_UserMemId = 0

    Value = Left$(pstrBuffer, plngStringLen)

End Property

Private Function AllocString04(ByVal lSize As Long) As String

    ' http://www.xbeat.net/vbspeed/
    ' by Jory, jory@joryanick.com, 20011023
    RtlMoveMemory ByVal VarPtr(AllocString04), SysAllocStringByteLen(0&, lSize + lSize), 4&

End Function

Public Sub Clear()

    '***  do not clear the buffer to save allocation time
    '***  if you use the function multiple times
    plngStringLen = 0&
    plngBufferLen = 0& 'clear the buffer
    pstrBuffer = vbNullString 'clear the buffer

End Sub

