VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GBytes"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum EErrorBytes
    eeBaseBytes = 13430     ' Bytes
End Enum

Private aPower2(0 To 31) As Long

Sub StrToBytes(ab() As Byte, s As String)
    If MUtility.IsArrayEmpty(ab) Then
        ' Assign to empty array
        ab = StrConv(s, vbFromUnicode)
    Else
        Dim cab As Long
        ' Copy to existing array, padding or truncating if necessary
        cab = UBound(ab) - LBound(ab) + 1
        If Len(s) < cab Then s = s & String$(cab - Len(s), 0)
        If UnicodeTypeLib Then
            Dim st As String
            st = StrConv(s, vbFromUnicode)
            CopyMemoryStr ab(LBound(ab)), st, cab
        Else
            CopyMemoryStr ab(LBound(ab)), s, cab
        End If
    End If
End Sub

Function StrToBytesV(s As String) As Variant
    ' Copy to array
    StrToBytesV = StrConv(s, vbFromUnicode)
End Function

' Convert an ANSI string in a byte array to a VB Unicode string
Function BytesToStr(ab() As Byte) As String
    BytesToStr = StrConv(ab, vbUnicode)
End Function

' Convert a null-terminated string in a padded byte array buffer
' to a VB string with no padding.
Function ByteZToStr(ab() As Byte) As String
    If UnicodeTypeLib Then
        ByteZToStr = ab
    Else
        ByteZToStr = StrConv(ab, vbUnicode)
    End If
    ByteZToStr = Left$(ByteZToStr, lstrlen(ByteZToStr))
End Function

' ANSI only version
Function AByteZToStr(ab() As Byte) As String
    AByteZToStr = StrConv(ab, vbUnicode)
    AByteZToStr = Left$(AByteZToStr, lstrlenAByte(ab(0)))
End Function

Function BytesToWord(abBuf() As Byte, iOffset As Long) As Integer
    BugAssert iOffset <= UBound(abBuf) + 1 - 2
    Dim w As Integer
    CopyMemory w, abBuf(iOffset), 2
    BytesToWord = w
End Function

Function BytesToDWord(abBuf() As Byte, iOffset As Long) As Long
    BugAssert iOffset <= UBound(abBuf) + 1 - 4
    Dim dw As Long
    CopyMemory dw, abBuf(iOffset), 4
    BytesToDWord = dw
End Function

Sub BytesFromWord(w As Integer, abBuf() As Byte, iOffset As Long)
    BugAssert iOffset <= UBound(abBuf)
    CopyMemory abBuf(iOffset), w, 2
End Sub

' Read string with length in first byte
Function BytesToPStr(ab() As Byte, iOffset As Long) As String
    BugAssert iOffset <= UBound(ab)
    BytesToPStr = MidBytes(ab, iOffset + 1, ab(iOffset))
End Function

Sub BytesFromDWord(dw As Long, abBuf() As Byte, iOffset As Long)
    BugAssert iOffset <= UBound(abBuf) + 1 - 4
    CopyMemory abBuf(iOffset), dw, 4
End Sub

'' Emulate relevant Basic string functions for arrays of bytes:
''     Len$             LenBytes
''     Mid$ function    MidBytes
''     Mid$ statement   InsBytes sub
''     Left$            LeftBytes
''     Right$           RightBytes

' LenBytes - Emulates Len for array of bytes
Function LenBytes(ab() As Byte) As Long
    LenBytes = UBound(ab) - LBound(ab) + 1
End Function

' MidBytes - emulates Mid$ function for array of bytes
' (Note that MidBytes does not emulate Mid$ exactly--string fields
' in byte arrays are often null-padded, and MidBytes can extract
' non-null portion)
Function MidBytes(ab() As Byte, ByVal iOffset As Long, _
                  Optional ByVal iLen As Long = 0, _
                  Optional fToNull As Boolean = False) As String
    BugAssert iOffset < LenBytes(ab) And iOffset >= 0
    Dim s As String, cab As Long
    ' Calculate length
    If iLen <= 0 Then
        cab = LenBytes(ab) - iOffset
    Else
        cab = iLen
    End If
    ' Assign and return string
    s = String$(cab, 0)
    CopyMemoryToStr s, ab(iOffset), cab
    If UnicodeTypeLib Then s = MUtility.StrZToStr(StrConv(s, vbUnicode))
    If fToNull Then
        cab = InStr(s, vbNullChar)
        If cab Then
            MidBytes = Left$(s, cab - 1)
        Else
            MidBytes = s
        End If
    Else
        MidBytes = s
    End If
End Function

' InsBytes - Emulates Mid$ statement for array of bytes
' (Note that InsBytes does not emulate Mid$ exactly--it inserts
' a null-padded string into a fixed-size field in order to work
' better with common use of byte arrays.)
Sub InsBytes(sIns As String, ab() As Byte, ByVal iOffset As Long, _
             Optional iLen As Long = 0)
    BugAssert iOffset < LenBytes(ab) And iOffset >= 0
    Dim cab As Long
    ' Calculate length
    If iLen <= 0 Then
        cab = Len(sIns)
    Else
        cab = iLen
        ' Null-pad insertion string if too short
        If (Len(sIns) < cab) Then
            sIns = sIns & String$(cab - Len(sIns), 0)
        End If
    End If
    BugAssert (Len(sIns) <= (LenBytes(ab) - iOffset))
    ' Insert string
    If UnicodeTypeLib Then
        Dim s As String
        s = StrConv(sIns, vbFromUnicode)
        CopyMemoryStr ab(iOffset), s, cab
    Else
        CopyMemoryStr ab(iOffset), sIns, cab
    End If
End Sub

' LeftBytes - Emulates Left$ function for array of bytes
Function LeftBytes(ab() As Byte, ByVal iLen As Long) As String
    Dim s As String
    s = String$(iLen, 0)
    CopyMemoryToStr s, ab(LBound(ab)), iLen
    If UnicodeTypeLib Then s = MUtility.StrZToStr(StrConv(s, vbUnicode))
    LeftBytes = s
End Function

' RightBytes - Emulates Right$ function for array of bytes
Function RightBytes(ab() As Byte, ByVal iLen As Long) As String
    Dim s As String
    s = String$(iLen, 0)
    CopyMemoryToStr s, ab(UBound(ab) - iLen + 1), iLen
    If UnicodeTypeLib Then s = MUtility.StrZToStr(StrConv(s, vbUnicode))
    RightBytes = s
End Function

' FillBytes - Fills field in array of bytes with given byte
Sub FillBytes(ab() As Byte, ByVal b As Byte, _
              ByVal iOffset As Long, ByVal iLen As Long)
    BugAssert (iOffset < LenBytes(ab)) And (iOffset >= 0)
    BugAssert iOffset - 1 + iLen <= LenBytes(ab)
    Dim i As Long
    For i = iOffset To iOffset + iLen - 1
        ab(i) = b
    Next
End Sub

' InStrBytes is not implemented because a simple version would
' simply be equivalent to InStr(ab(), s). This creates a temporary
' string for ab() on every call. An efficient version that works
' directly on arrays of bytes could be written in C.

Function LoWord(ByVal dw As Long) As Integer
    If dw And &H8000& Then
        LoWord = dw Or &HFFFF0000
    Else
        LoWord = dw And &HFFFF&
    End If
End Function

Function HiWord(ByVal dw As Long) As Integer
    HiWord = (dw And &HFFFF0000) \ 65536
End Function

Function LoByte(ByVal w As Integer) As Byte
    LoByte = w And &HFF
End Function

Function HiByte(ByVal w As Integer) As Byte
    HiByte = (w And &HFF00&) \ 256
End Function

Function MakeWord(ByVal bLo As Byte, ByVal bHi As Byte) As Integer
    'CopyMemory MakeWord, bLo, 1
    'CopyMemory ByVal VarPtr(MakeWord) + 1, bHi, 1
    If bHi And &H80 Then
        MakeWord = ((bHi * 256&) + bLo) Or &HFFFF0000
    Else
        MakeWord = (bHi * 256) + bLo
    End If
End Function

Function MakeDWord(ByVal wLo As Integer, ByVal wHi As Integer) As Long
    'CopyMemory MakeDWord, wLo, 2
    'CopyMemory ByVal VarPtr(MakeDWord) + 2, wHi, 2
    MakeDWord = (wHi * 65536) + (wLo And &HFFFF&)
End Function

Function LShiftWord(ByVal w As Integer, ByVal c As Integer) As Integer
    BugAssert c >= 0 And c <= 15
    Dim dw As Long
    dw = w * Power2(c)
    If dw And &H8000& Then
        LShiftWord = CInt(dw And &H7FFF&) Or &H8000
    Else
        LShiftWord = dw And &HFFFF&
    End If
End Function

Function RShiftWord(ByVal w As Integer, ByVal c As Integer) As Integer
    BugAssert c >= 0 And c <= 15
    Dim dw As Long
    If c = 0 Then
        RShiftWord = w
    Else
        dw = w And &HFFFF&
        dw = dw \ Power2(c)
        RShiftWord = dw And &HFFFF&
    End If
End Function

Function LShiftDWord(ByVal dw As Long, ByVal c As Integer) As Long
    BugAssert c >= 0 And c <= 31
    Dim dwT As Long
    On Error GoTo FailLShiftDWord
    dwT = dw * Power2(c)
    If dwT And &H80000000 Then
        LShiftDWord = CLng(dwT And &H7FFFFFFF) Or &H80000000
    Else
        LShiftDWord = dwT
    End If
    Exit Function
FailLShiftDWord:
    LShiftDWord = &HFFFFFFFF
End Function

Function RShiftDWord(ByVal dw As Long, ByVal c As Integer) As Long
    BugAssert c >= 0 And c <= 31
    On Error GoTo FailRShiftDWord
    If c = 0 Then
        RShiftDWord = dw
    Else
        RShiftDWord = dw \ Power2(c)
    End If
    Exit Function
FailRShiftDWord:
    RShiftDWord = 0
End Function

' Set or clear iBitPos bit in iValue according to whether
' iTest expression is true.
Sub SetBitWord(ByVal iTest As Boolean, iValue As Integer, _
               ByVal iBitPos As Integer)
    BugAssert iBitPos >= 0 And iBitPos <= 15
    If iTest Then
        iValue = LoWord(iValue Or Power2(iBitPos))
    Else
        iValue = LoWord(iValue And Not Power2(iBitPos))
    End If
End Sub

Sub SetBitDWord(ByVal iTest As Boolean, iValue As Long, _
                ByVal iBitPos As Integer)
    BugAssert iBitPos >= 0 And iBitPos <= 31
    If iTest Then
        iValue = iValue Or Power2(iBitPos)
    Else
        iValue = iValue And Not Power2(iBitPos)
    End If
End Sub

' Get state of iBitPos bit in iValue
Function GetBit(ByVal iValue As Long, ByVal iBitPos As Integer) As Boolean
    BugAssert iBitPos >= 0 And iBitPos <= 31
    GetBit = iValue And Power2(iBitPos)
End Function

Function SwapWordBytes(ByVal w As Integer) As Integer
    CopyMemory ByVal VarPtr(SwapWordBytes) + 1, w, 1
    CopyMemory SwapWordBytes, ByVal VarPtr(w) + 1, 1
End Function

Function SwapDWordWords(ByVal dw As Long) As Long
    CopyMemory ByVal VarPtr(SwapDWordWords) + 2, dw, 2
    CopyMemory SwapDWordWords, ByVal VarPtr(dw) + 2, 2
End Function

' Swap a little endian DWORD to big endian, or vice versa
Function SwapEndian(ByVal dw As Long) As Long
    CopyMemory ByVal VarPtr(SwapEndian) + 3, dw, 1
    CopyMemory ByVal VarPtr(SwapEndian) + 2, ByVal VarPtr(dw) + 1, 1
    CopyMemory ByVal VarPtr(SwapEndian) + 1, ByVal VarPtr(dw) + 2, 1
    CopyMemory SwapEndian, ByVal VarPtr(dw) + 3, 1
End Function

Function VBGetLogicalDrives() As String

    Dim f32  As Long, i As Integer, s As String
    f32 = GetLogicalDrives()
    For i = 0 To 25
        s = s & IIf(f32 And 1, "+", "-")
        f32 = MBytes.RShiftDWord(f32, 1)
    Next
    VBGetLogicalDrives = s
    
End Function

Property Get Power2(ByVal i As Integer) As Long
    BugAssert i >= 0 And i <= 31
#If fComponent = 0 Then
    If aPower2(0) = 0 Then
        aPower2(0) = &H1&
        aPower2(1) = &H2&
        aPower2(2) = &H4&
        aPower2(3) = &H8&
        aPower2(4) = &H10&
        aPower2(5) = &H20&
        aPower2(6) = &H40&
        aPower2(7) = &H80&
        aPower2(8) = &H100&
        aPower2(9) = &H200&
        aPower2(10) = &H400&
        aPower2(11) = &H800&
        aPower2(12) = &H1000&
        aPower2(13) = &H2000&
        aPower2(14) = &H4000&
        aPower2(15) = &H8000&
        aPower2(16) = &H10000
        aPower2(17) = &H20000
        aPower2(18) = &H40000
        aPower2(19) = &H80000
        aPower2(20) = &H100000
        aPower2(21) = &H200000
        aPower2(22) = &H400000
        aPower2(23) = &H800000
        aPower2(24) = &H1000000
        aPower2(25) = &H2000000
        aPower2(26) = &H4000000
        aPower2(27) = &H8000000
        aPower2(28) = &H10000000
        aPower2(29) = &H20000000
        aPower2(30) = &H40000000
        aPower2(31) = &H80000000
    End If
#End If
    Power2 = aPower2(i)
End Property

#If fComponent Then
Private Sub Class_Initialize()
    aPower2(0) = &H1&
    aPower2(1) = &H2&
    aPower2(2) = &H4&
    aPower2(3) = &H8&
    aPower2(4) = &H10&
    aPower2(5) = &H20&
    aPower2(6) = &H40&
    aPower2(7) = &H80&
    aPower2(8) = &H100&
    aPower2(9) = &H200&
    aPower2(10) = &H400&
    aPower2(11) = &H800&
    aPower2(12) = &H1000&
    aPower2(13) = &H2000&
    aPower2(14) = &H4000&
    aPower2(15) = &H8000&
    aPower2(16) = &H10000
    aPower2(17) = &H20000
    aPower2(18) = &H40000
    aPower2(19) = &H80000
    aPower2(20) = &H100000
    aPower2(21) = &H200000
    aPower2(22) = &H400000
    aPower2(23) = &H800000
    aPower2(24) = &H1000000
    aPower2(25) = &H2000000
    aPower2(26) = &H4000000
    aPower2(27) = &H8000000
    aPower2(28) = &H10000000
    aPower2(29) = &H20000000
    aPower2(30) = &H40000000
    aPower2(31) = &H80000000
End Sub
#End If
'

#If fComponent = 0 Then
Private Sub ErrRaise(e As Long)
    Dim sText As String, sSource As String
    If e > 1000 Then
        sSource = App.ExeName & ".Bytes"
        Select Case e
        Case eeBaseBytes
            BugAssert True
       ' Case ee...
       '     Add additional errors
        End Select
        Err.Raise COMError(e), sSource, sText
    Else
        ' Raise standard Visual Basic error
        sSource = App.ExeName & ".VBError"
        Err.Raise e, sSource
    End If
End Sub
#End If

