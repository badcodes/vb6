Attribute VB_Name = "MZlib"
Option Explicit

Private Const Z_OK As Long = &H0

Private Declare Function compress Lib "ZLibWAPI.dll" ( _
    ByRef dest As Any, ByRef destLen As Long, _
    ByRef Source As Any, ByVal sourceLen As Long) As Long
    
Private Declare Function compressBound Lib "ZLibWAPI.dll" ( _
    ByVal sourceLen As Long) As Long
    
Private Declare Function uncompress Lib "ZLibWAPI.dll" ( _
    ByRef dest As Any, ByRef destLen As _
    Long, ByRef Source As Any, ByVal sourceLen As Long) As Long
    
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
    ByRef Source As Any, ByVal sourceLen As Long, _
    ByVal level As Long) As Long
  
Public Function Zlib_UncompressString(ByVal vSource As String) As String
Dim Source() As Byte
Source = vSource
Dim result() As Byte
If (Zlib_Uncompress(Source(), result())) Then
    Zlib_UncompressString = StrConv(result(), vbUnicode)
End If
End Function

Public Function Zlib_UncompressFile(ByRef FileName As String, ByRef result() As Byte, Optional seekPos As Long = 1) As Boolean
    Dim srcLen As Long
    Dim fNum As Integer
    fNum = FreeFile
    Open FileName For Binary Access Read As #fNum
    srcLen = LOF(fNum) - seekPos + 1
    ReDim Source(0 To srcLen - 1) As Byte
    Seek #fNum, seekPos
    Get #fNum, , Source()
    Close #fNum
    Zlib_UncompressFile = (Zlib_Uncompress(Source, result) > 0)
End Function

Public Function Zlib_UncompressFileAs(ByRef srcFile As String, ByRef destFile As String, Optional seekPos As Long = 1) As Boolean
    '<EhHeader>
    On Error GoTo Zlib_UncompressFileAs_Err
    '</EhHeader>

    Dim result() As Byte
    If Zlib_UncompressFile(srcFile, result, seekPos) = True Then
        Dim fNum As Integer
        fNum = FreeFile
        Open destFile For Binary Access Write Lock Read As #fNum
        Put #fNum, , result
        Close #fNum
        fNum = 0
        Zlib_UncompressFileAs = True
    End If
   
    '<EhFooter>
    Exit Function

Zlib_UncompressFileAs_Err:
    Debug.Print "GetSSLib.MZlib.Zlib_UncompressFileAs:Error " & Err.Description
    Err.Clear
    Zlib_UncompressFileAs = False
    On Error Resume Next
    If fNum > 0 Then Close #fNum
    '</EhFooter>
End Function

Public Function Zlib_FileExist(ByVal vFilename As String) As Boolean
    '<EhHeader>
    On Error GoTo Zlib_FileExist_Err
    '</EhHeader>
    GetAttr vFilename
    Zlib_FileExist = True
    '<EhFooter>
    Exit Function

Zlib_FileExist_Err:
    'Debug.Print "GetSSLib.MZlib.Zlib_FileExist:Error " & Err.Description
    Err.Clear
    Zlib_FileExist = False
    '</EhFooter>
End Function

Public Function Zlib_CompressFileAs(ByVal vFilename As String, ByVal vDestname As String) As Boolean
    '<EhHeader>
    On Error GoTo Zlib_CompressFileAs_Err
    '</EhHeader>
    Dim result() As Byte
    result = Zlib_CompressFile(vFilename)
    'If result <> "" Then
        If Zlib_FileExist(vDestname) Then Kill vDestname
        Dim fNum As Integer
        fNum = FreeFile
        Open vDestname For Binary Access Write As #fNum
        Put #fNum, , result()
        Close #fNum
        fNum = 0
        Zlib_CompressFileAs = True
    'End If
    '<EhFooter>
    Exit Function

Zlib_CompressFileAs_Err:
    Debug.Print "GetSSLib.MZlib.Zlib_CompressFileAs:Error " & Err.Description
    Err.Clear
    Zlib_CompressFileAs = False
    On Error Resume Next
    If fNum > 0 Then Close #fNum
    '</EhFooter>
End Function
Public Function Zlib_CompressFile(ByVal vFilename As String) As String
    '<EhHeader>
    On Error GoTo Zlib_CompressFile_Err
    '</EhHeader>
    Dim Src() As Byte
    Dim fNum As Integer
    fNum = FreeFile
    Open vFilename For Binary Access Read As #fNum
    If LOF(fNum) > 0 Then
        ReDim Src(0 To LOF(fNum) - 1)
        Get #fNum, , Src()
        Dim result() As Byte
        Dim Size As Long
        Size = Zlib_Compress(Src(), result())
        If Size > 0 Then Zlib_CompressFile = result()
    End If
    Close #fNum
    fNum = 0
    '<EhFooter>
    Exit Function

Zlib_CompressFile_Err:
    Debug.Print "GetSSLib.MZlib.Zlib_CompressFile:Error " & Err.Description
    Err.Clear
    Zlib_CompressFile = ""
    On Error Resume Next
    If fNum > 0 Then Close #fNum
    '</EhFooter>
End Function
Public Function Zlib_CompressString(ByVal vSource As String) As String
    Dim Src() As Byte
    Dim result() As Byte
    Dim Size As Long
    Src = StrConv(vSource, vbFromUnicode)
    Size = Zlib_Compress(Src(), result())
    If Size > 0 Then
        Zlib_CompressString = result
    End If
End Function
Public Function Zlib_Compress(ByRef vSource() As Byte, ByRef vResult() As Byte) As Long
    '<EhHeader>
    On Error GoTo Zlib_Compress_Err
    '</EhHeader>
    Dim srcSize As Long
    Dim dstSize As Long
    Dim ret As Long
    srcSize = UBound(vSource) + 1
    If srcSize > 0 Then
        dstSize = srcSize / 2
        Do
            dstSize = 2 * dstSize
            ReDim vResult(0 To dstSize - 1)
            ret = compress(vResult(0), dstSize, vSource(0), srcSize)
            'If ret = Z_OK Then Exit Do
            
        Loop While ret = -5
        If ret = Z_OK Then
            ReDim Preserve vResult(0 To dstSize - 1)
            Zlib_Compress = dstSize
        Else
            Zlib_Compress = -1
        End If
    Else
        Zlib_Compress = 0
    End If
    
    '<EhFooter>
    Exit Function

Zlib_Compress_Err:
    Zlib_Compress = -1
    Debug.Print "GetSSLib.MZlib.Zlib_Compress:Error " & Err.Description
    Err.Clear

    '</EhFooter>
End Function

Public Function Zlib_Uncompress(ByRef Source() As Byte, ByRef result() As Byte) As Long
    On Error GoTo invalidUsage
        Dim destLen As Long
        Dim sourceLen As Long
        Dim ret As Long
        sourceLen = UBound(Source()) + 1
        destLen = sourceLen
        Do
            destLen = destLen * 2
            ReDim result(0 To destLen - 1)
            ret = uncompress(result(0), destLen, Source(0), sourceLen)
            
         '   If ret = -3 Then Exit Do
        Loop While ret = -5
        'Until ret = Z_OK
        
        If ret = Z_OK And destLen > 0 Then
            ReDim Preserve result(0 To destLen - 1)
            Zlib_Uncompress = destLen
        Else
            Zlib_Uncompress = -1
        End If
        'Uncompress = result
        'Dim result(0 to ubound(source())
    Exit Function
invalidUsage:
    Debug.Print Err.Description
    Zlib_Uncompress = -1
End Function

Public Function Zlib_UncompressFileToString(ByRef srcFile As String, Optional seekPos As Long = 1) As String
    Dim result() As Byte
    If (Zlib_UncompressFile(srcFile, result, seekPos)) Then
        Zlib_UncompressFileToString = StrConv(result, vbUnicode)
    End If
End Function

Public Function Zlib_UncompressToString(ByRef Source() As Byte) As String
    Dim result() As Byte
    If (Zlib_Uncompress(Source, result) > 0) Then
        Zlib_UncompressToString = StrConv(result, vbUnicode)
    End If
End Function

'Public Sub Test()
''Dim fNum As Integer
''fNum = FreeFile
''Dim srcLen As Long
''Dim temp As String
''
''Open "X:\download\pdg\文学\龙文鞭影_11239938\InfoRule.dat" For Binary Access Read As #fNum
''srcLen = LOF(fNum) - &H44
''ReDim source(0 To srcLen) As Byte
''
''Seek #fNum, &H44 + 1
''Get #fNum, , source()
''Close #fNum
''
''Dim result() As Byte
'' Uncompress source, result
''
''fNum = FreeFile
''Open "X:\download\pdg\文学\龙文鞭影_11239938\InfoRule.txt" For Binary Access Write As #fNum
''Put #fNum, , result()
''Close #fNum
'
'
'Dim vUrls() As String
'SSLIB_ParseInfoRule "X:\download\pdg\文学\龙文鞭影_11239938\InfoRule.dat", vUrls
'
'
'End Sub
'
'
