Attribute VB_Name = "MFileEncrypt"
Option Explicit
Private Const CryptKey1 = 29 '|(Asc ("L") + Asc("I") + Asc("N")-256|
Private Const CryptKey2 = 49  '|Asc("X") + Asc("I") + Asc("A") + Asc("O") - 256|
Private Const CryptKey3 = 31 '|Asc ("R") + Asc("A") + Asc("N")-256|
Private Const CryptFlag = "LCF" 'Lin Crypt File
Public Function MyFileEncrypt(srcfile As String, dstfile As String) As Boolean

    Dim tmpFile As String
    Dim thebyte As Byte
    MyFileEncrypt = True

    If Dir$(srcfile) = "" Then MyFileEncrypt = False: Exit Function
    tmpFile = "~$$$CRfile.tmp"

    If Dir$(tmpFile) <> "" Then Kill tmpFile
    Open srcfile For Binary As #1
    Open tmpFile For Binary As #2
    Put #2, , CryptFlag '±êÊ¶·û

    Do Until Loc(1) = LOF(1)
        Get #1, , thebyte
        thebyte = thebyte Xor CryptKey1
        thebyte = thebyte Xor CryptKey2
        thebyte = thebyte Xor CryptKey3
        Put #2, , thebyte
    Loop

    Close #1
    Close #2

    If Dir$(dstfile) <> "" Then Kill dstfile
    FileCopy tmpFile, dstfile
    Kill tmpFile

End Function

Public Function MyFileDecrypt(srcfile As String, dstfile As String) As Boolean

    Dim tmpFile As String
    Dim thebyte As Byte
    Dim skipflag
    MyFileDecrypt = True

    If Dir$(srcfile) = "" Then MyFileDecrypt = False: Exit Function

    If isLXTfile(srcfile) = False Then MyFileDecrypt = False: Exit Function
    Open srcfile For Binary As #1
    tmpFile = "~$$$CRfile.tmp"

    If Dir$(tmpFile) <> "" Then Kill tmpFile
    Open tmpFile For Binary As #2
    skipflag = Input(Len(CryptFlag), #1)

    Do Until Loc(1) = LOF(1)
        Get #1, , thebyte
        thebyte = thebyte Xor CryptKey3
        thebyte = thebyte Xor CryptKey2
        thebyte = thebyte Xor CryptKey1
        Put #2, , thebyte
    Loop

    Close #1
    Close #2

    If Dir$(dstfile) <> "" Then Kill dstfile
    FileCopy tmpFile, dstfile
    Kill tmpFile

End Function

Public Function isLXTfile(thefile As String) As Boolean

    Dim fso As New gCFileSystem
    Dim fNum As Integer
    Dim sTest As String
    isLXTfile = False

    If fso.PathExists(thefile) = False Then Exit Function
    fNum = FreeFile
    Open thefile For Binary Access Read As fNum
    sTest = String$(Len(CryptFlag), " ")
    Get fNum, , sTest
    Close fNum

    If sTest = CryptFlag Then isLXTfile = True

End Function
