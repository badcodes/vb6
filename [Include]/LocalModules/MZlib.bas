Attribute VB_Name = "MZlib"
Option Explicit



Private Const Z_OK As Long = &H0

Private Declare Function compress Lib "ZLibWAPI.dll" ( _
    ByRef dest As Any, ByRef destLen As Long, _
    ByRef source As Any, ByVal sourceLen As Long) As Long
    
Private Declare Function compressBound Lib "ZLibWAPI.dll" ( _
    ByVal sourceLen As Long) As Long
    
Private Declare Function Uncompress Lib "ZLibWAPI.dll" Alias _
    "uncompress" (ByRef dest As Any, ByRef destLen As _
    Long, ByRef source As Any, ByVal sourceLen As Long) As Long
    
Private Declare Function adler32 Lib "ZLibWAPI.dll" ( _
    ByVal adler As Long, ByRef buf As Any, ByVal length As Long) As Long
    
Private Declare Function crc32 Lib "ZLibWAPI.dll" ( _
    ByVal crc As Long, ByRef buf As Any, ByVal length As Long) As Long

Private Declare Function zlibCompileFlags Lib "ZLibWAPI.dll" () As Long

Private Const Z_NO_COMPRESSION As Long = 0
Private Const Z_BEST_SPEED As Long = 1
Private Const Z_BEST_COMPRESSION As Long = 9
Private Const Z_DEFAULT_COMPRESSION As Long = (-1)
 
Private Declare Function compress2 Lib "ZLibWAPI.dll" ( _
    ByRef dest As Any, ByRef destLen As Long, _
    ByRef source As Any, ByVal sourceLen As Long, _
    ByVal level As Long) As Long
  
Public Function UncompressString(ByRef source() As Byte) As String
Dim Result() As Byte
If (UncompressBuf(source, Result)) Then
    UncompressString = StrConv(Result, vbUnicode)
End If
End Function

Public Function UncompressFile(ByRef filename As String, ByRef Result() As Byte, Optional seekPos As Long = 1) As Boolean
    Dim srcLen As Long
    Dim fNUM As Integer
    fNUM = FreeFile
    Open filename For Binary Access Read As #fNUM
    srcLen = LOF(fNUM) - seekPos + 1
    ReDim source(0 To srcLen - 1) As Byte
    Seek #fNUM, seekPos
    Get #fNUM, , source()
    Close #fNUM
    UncompressFile = UncompressBuf(source, Result)
End Function

Public Function UncompressFileTo(ByRef srcFile As String, ByRef destFile As String, Optional seekPos As Long = 1) As Boolean

    Dim Result() As Byte
    If (UncompressFile(srcFile, Result, seekPos)) Then
        Dim fNUM As Integer
        fNUM = FreeFile
        Open destFile For Binary Access Write Lock Read As #fNUM
        Put #fNUM, , Result
        Close #fNUM
        UncompressFileTo = True
    End If
   
End Function

Public Function UncompressBuf(ByRef source() As Byte, ByRef Result() As Byte) As Boolean
    On Error GoTo invalidUsage
        Dim destLen As Long
        Dim sourceLen As Long
        Dim ret As Long
        sourceLen = UBound(source()) + 1
        destLen = sourceLen
        Do
            destLen = destLen * 2
            ReDim Result(0 To destLen - 1)
            ret = Uncompress(Result(0), destLen, source(0), sourceLen)
            
         '   If ret = -3 Then Exit Do
        Loop While ret = -5
        'Until ret = Z_OK
        ReDim Preserve Result(0 To destLen - 1)
        If ret = Z_OK Then UncompressBuf = True
        'UncompressBuf = result
        'Dim result(0 to ubound(source())
    Exit Function
invalidUsage:
    Debug.Print Err.Description
    UncompressBuf = False
End Function

Public Function UncompressFileToString(ByRef srcFile As String, Optional seekPos As Long = 1) As String
    Dim Result() As Byte
    If (UncompressFile(srcFile, Result, seekPos)) Then
        UncompressFileToString = StrConv(Result, vbUnicode)
    End If
End Function

Public Function UncompressBufToString(ByRef source() As Byte) As String
    Dim Result() As Byte
    If (UncompressBuf(source, Result)) Then
        UncompressBufToString = StrConv(Result, vbUnicode)
    End If
End Function

Public Sub Test()
'Dim fNum As Integer
'fNum = FreeFile
'Dim srcLen As Long
'Dim temp As String
'
'Open "X:\download\pdg\文学\龙文鞭影_11239938\InfoRule.dat" For Binary Access Read As #fNum
'srcLen = LOF(fNum) - &H44
'ReDim source(0 To srcLen) As Byte
'
'Seek #fNum, &H44 + 1
'Get #fNum, , source()
'Close #fNum
'
'Dim result() As Byte
' UncompressBuf source, result
'
'fNum = FreeFile
'Open "X:\download\pdg\文学\龙文鞭影_11239938\InfoRule.txt" For Binary Access Write As #fNum
'Put #fNum, , result()
'Close #fNum


Dim vUrls() As String
SSLIB_ParseInfoRule "X:\download\pdg\文学\龙文鞭影_11239938\InfoRule.dat", vUrls


End Sub


