Attribute VB_Name = "stringasist"
Option Explicit
Const MAXLINE = 10240
Public Function strreplace(theSTR As String, rpstr As String, dststr As String) As String

Dim tempstr As String
tempstr = theSTR
Dim n As Integer
n = InStr(tempstr, rpstr)
If rpstr = dststr Then n = 0
Do Until n = 0
    strreplace = strreplace + Left(tempstr, n - 1) + dststr
    tempstr = Right(tempstr, Len(tempstr) - n - Len(rpstr) + 1)
    n = InStr(tempstr, rpstr)
Loop
    strreplace = strreplace + tempstr

End Function

Function fmstr(SrcStr As String) As String
Const space = "-"
Dim tmpstr As String
Dim i As Long

For i = 1 To Len(SrcStr)
If Mid(SrcStr, i, 1) = " " Then
    tmpstr = Left$(SrcStr, i - 1) + space + Right$(SrcStr, Len(SrcStr) - i)
    SrcStr = tmpstr
End If
Next
fmstr = SrcStr

End Function




Public Function rdel(theSTR As String) As String
rdel = theSTR
If rdel = "" Then Exit Function
Dim A As String
A = Right(rdel, 1)
Do Until A <> Chr(0) And A <> Chr(32) And A <> Chr(10) And A <> Chr(13)
rdel = Left(rdel, Len(rdel) - 1)
A = Right(rdel, 1)
Loop
End Function

Public Function ldel(theSTR As String) As String
ldel = theSTR
If ldel = "" Then Exit Function
Dim A As String
A = Left(ldel, 1)
Do Until A <> Chr(0) And A <> Chr(32) And A <> Chr(10) And A <> Chr(13)
ldel = Right(ldel, Len(ldel) - 1)
A = Left(ldel, 1)
Loop
End Function

Function FileStrFind(thefile As String, thetext As String) As Long
Const MAXSTRING = 28800
'If thetext = RPText Then Exit Function
If thetext = "" Then Exit Function
If Len(thetext) >= MAXSTRING \ 2 Then MsgBox ("The text to replace is too large!"): Exit Function
Dim fso As New FileSystemObject
If fso.FileExists(thefile) = False Then Exit Function
Dim ff As File
Set ff = fso.GetFile(thefile)
If ff.Size < Len(thetext) Then
End If

Dim BlockSize As Long
Dim textsize As Long
Dim blocknum As Long
Dim restchar As Long

BlockSize = MAXSTRING

If ff.Size < BlockSize Then BlockSize = ff.Size
textsize = Len(thetext)
blocknum = ff.Size \ (BlockSize)
restchar = ff.Size Mod (BlockSize)
Dim tempstring As String
Dim srcTS As TextStream
Set srcTS = fso.OpenTextFile(thefile, ForReading)
Dim n As Long
Dim textPos As Long

n = 0
tempstring = srcTS.Read(BlockSize)
textPos = InStr(String(textsize, " ") + tempstring, thetext)

If textPos = 0 Then
For n = 1 To blocknum - 1
tempstring = Right(tempstring, textsize) + srcTS.Read(BlockSize)
textPos = InStr(tempstring, thetext)
If textPos > 0 Then Exit For
Next
End If

If textPos > 0 Then
    textPos = textPos + n * BlockSize - textsize
End If

FileStrFind = textPos

End Function

Function StrNum(Num As Integer, numnum As Integer) As String
StrNum = LTrim(Str(Num))
If Len(StrNum) >= numnum Then Exit Function
StrNum = String(numnum - Len(StrNum), "0") + StrNum
End Function

Public Function strRepl(theSTR As String, strOLD As String, strNEW As String) As String
strRepl = Replace$(theSTR, strOLD, strNEW)
End Function

Public Function MyInstr(strBig, strSmall) As Boolean
Dim strcount As Integer
Dim strSmallOne() As String
Dim i As Long

MyInstr = False
If strSmall = "" Then Exit Function
strSmallOne = Split(strSmall, ",")
strcount = UBound(strSmallOne)
For i = 0 To strcount
If InStr(1, strBig, strSmallOne(i), vbTextCompare) > 0 Then MyInstr = True: Exit Function
Next
End Function


Public Function strBetween(theSTR, strStart As String, strEnd As String) As String

If strStart = "" Then Exit Function
If strEnd = "" Then Exit Function

Dim pos1 As Integer
Dim pos2 As Integer

pos1 = InStr(1, theSTR, strStart, vbTextCompare)
If pos1 > 0 Then
    pos2 = InStr(pos1 + Len(strStart), theSTR, strEnd, vbTextCompare)
    If pos2 > 0 Then
    strBetween = Mid(theSTR, pos1 + Len(strStart), pos2 - pos1 - Len(strStart))
    End If
End If


End Function
Public Function cleanFilename(ByRef sFilenameDirty As String) As String

    Dim iLoop As Long, iEnd As Long
    Dim charCur As String * 1
    iEnd = Len(sFilenameDirty)

    For iLoop = 1 To iEnd
        charCur = Mid$(sFilenameDirty, iLoop, 1)

        Select Case charCur
        Case ":", "?"
            cleanFilename = cleanFilename & StrConv(charCur, vbWide)
        Case "\", "/", "|", ">", "<", "*", Chr$(34)
        Case Else
            cleanFilename = cleanFilename & charCur
        End Select

    Next

End Function
