VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CZlib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
  

Public Function UncompressFile(ByRef source As String, ByRef dest As String) As Boolean

End Function

Public Function UncompressBuf(ByRef source() As Byte) As Byte()
    On Error GoTo invalidUsage
        Dim destLen As Long
        Dim sourceLen As Long
        Dim ret As Long
        sourceLen = UBound(source())
        destLen = sourceLen
        Do
            ReDim result(0 To destLen - 1)
            destLen = destLen * 2
            ret = Uncompress(result(0), destLen, source(0), sourceLen)
        Loop Until ret = Z_OK
        ReDim Preserve result(0 To destLen - 1)
        Uncompress = result
        'Dim result(0 to ubound(source())
    Exit Function
invalidUsage:
End Function


Public Sub Test()
Dim fnum As Integer
fnum = FreeFile
Dim srcLen As Long
srcLen = LOF(fnum)
ReDim source(0 To srcLen - 1) As Byte

Open "X:\download\pdg\文学\龙文鞭影_11239938\InfoRule.dat" For Binary Access Read As #fnum
Get #fnum, srcLen, source()
Close #fnum

Dim result() As Byte
result = UncompressBuf(source)

End Sub
