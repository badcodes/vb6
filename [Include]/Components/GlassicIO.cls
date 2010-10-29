VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GClassicIO"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'CSEH: ErrRaise
Option Explicit

'CSEH: ErrRaise
Public Function Read(ByRef fNum, Optional lChar As Long = 1) As String
'<EhHeader>
On Error GoTo Read_Err
'</EhHeader>

    If Not EOF(fNum) Then Read = Input$(lChar, fNum)

'<EhFooter>
Exit Function

Read_Err:
    Err.Raise vbObjectError + 100, _
              "htmlParser.MClassicIO.Read", _
              "MClassicIO component failure" & _
              vbCrLf & Err.Description
'</EhFooter>
End Function

Public Function UnRead(ByRef fNum As Integer, Optional lBytes As Long = 1) As Boolean
'<EhHeader>
On Error GoTo unRead_Err
'</EhHeader>

    UnRead = False
    Dim lCurPos As Long
    Dim lSeekTo As Long
    lCurPos = Seek(fNum)
    lSeekTo = lCurPos - lBytes
    If lSeekTo >= 0 Then
        Seek fNum, lSeekTo
        UnRead = True
    End If

'<EhFooter>
Exit Function

unRead_Err:
    Err.Raise vbObjectError + 100, _
              "htmlParser.MClassicIO.unRead", _
              "MClassicIO component failure" & _
              vbCrLf & Err.Description
'</EhFooter>
End Function

Public Function SkipUntil(ByRef c As String, ByRef fNum As Integer) As Boolean
'<EhHeader>
On Error GoTo skipUntil_Err
'</EhHeader>

    SkipUntil = False
    Dim CC As String
    If c = "" Then Exit Function
    Do While Not EOF(fNum)
        CC = Input$(1, #fNum)
        If InStr(c, CC) > 0 Then
            If UnRead(fNum) Then SkipUntil = True
            Exit Function
        End If
    Loop

'<EhFooter>
Exit Function

skipUntil_Err:
    Err.Raise vbObjectError + 100, _
              "htmlParser.MClassicIO.skipUntil", _
              "MClassicIO component failure" & _
              vbCrLf & Err.Description
'</EhFooter>
End Function

Public Function StrUntil(ByRef c As String, ByRef fNum As Integer) As String
'<EhHeader>
On Error GoTo strUntil_Err
'</EhHeader>

    Dim strResult As String
    Dim CC As String

    If c = "" Then Exit Function
    Do While Not EOF(fNum)
        CC = Input$(1, #fNum)
        If InStr(c, CC) > 0 Then UnRead (fNum): Exit Function
        StrUntil = StrUntil & CC
    Loop

'<EhFooter>
Exit Function

strUntil_Err:
    Err.Raise vbObjectError + 100, _
              "htmlParser.MClassicIO.strUntil", _
              "MClassicIO component failure" & _
              vbCrLf & Err.Description
'</EhFooter>
End Function

Public Function SkipChar(ByRef c As String, ByRef fNum As Integer) As Boolean
'<EhHeader>
On Error GoTo skipChar_Err
'</EhHeader>

    SkipChar = False
    Dim CC As String
    If c = "" Then Exit Function
    Do While Not EOF(fNum)
        CC = Input$(1, #fNum)
        If InStr(c, CC) <= 0 Then
            If UnRead(fNum) Then SkipChar = True
            Exit Function
        End If
    Loop

'<EhFooter>
Exit Function

skipChar_Err:
    Err.Raise vbObjectError + 100, _
              "htmlParser.MClassicIO.skipChar", _
              "MClassicIO component failure" & _
              vbCrLf & Err.Description
'</EhFooter>
End Function

Public Function FileStr( _
    ByRef strToSearch As String, _
    ByRef fNum As Integer, _
    Optional startAt As Long = 0, _
    Optional cmp As VbCompareMethod = vbBinaryCompare _
    ) As Long
    
'<EhHeader>
On Error GoTo fileStr_Err
'</EhHeader>
Dim filePos As Long
Dim cfile As String
Dim cSearch As String
Dim posSearch As Long
Dim posEnd As Long
Dim found As Boolean


posEnd = Len(strToSearch)
If posEnd < 1 Then Exit Function
filePos = Seek(fNum)
posSearch = 1
If startAt > 0 Then Seek fNum, startAt
FileStr = Seek(fNum)

Do While Not EOF(fNum)
    cfile = Read(fNum)
    cSearch = Mid$(strToSearch, posSearch, 1)
    If StrComp(cfile, cSearch, cmp) = 0 Then
        If posSearch >= posEnd Then found = True: Exit Do
        posSearch = posSearch + 1
    Else
        FileStr = Seek(fNum)
        posSearch = 1
    End If
Loop

Seek fNum, filePos
If found = False Then FileStr = 0

'<EhFooter>
Exit Function

fileStr_Err:
    Err.Raise vbObjectError + 100, _
              "Project1.MClassicIO.fileStr", _
              "MClassicIO component failure" & _
              vbCrLf & Err.Description
'</EhFooter>
End Function

'CSEH: ErrExit
Public Function StrBetween( _
                ByRef fNum As Integer, _
                ByRef strBegin As String, _
                ByRef strEnd As String, _
                ByRef strResult() As String, _
                Optional cmp As VbCompareMethod = vbBinaryCompare, _
                Optional returnEmpStr As Boolean = True _
                ) As Long
'<EhHeader>
On Error GoTo StrBetween_EXIT
'</EhHeader>
    If strBegin = "" Or strEnd = "" Then Exit Function
    Dim posStart As Long
    Dim posEnd As Long
    Dim startAt As Long
    Dim bytesCount As Long
    Dim tmpStr As String
    startAt = Seek(fNum)
    Do While Not EOF(fNum)
        posStart = 0
        posEnd = 0
        tmpStr = ""
        posStart = FileStr(strBegin, fNum, , cmp)
        If posStart < 1 Then Exit Do
        posStart = posStart + LenB(StrConv(strBegin, vbFromUnicode))
        Seek fNum, posStart
        posEnd = FileStr(strEnd, fNum, , cmp)
        If posEnd < 1 Then Exit Do

        bytesCount = posEnd - posStart
        If bytesCount > 0 Then tmpStr = StrConv(InputB(bytesCount, fNum), vbUnicode)
        
        If returnEmpStr = True Or tmpStr <> "" Then
            StrBetween = StrBetween + 1
            ReDim Preserve strResult(1 To StrBetween) As String
            strResult(StrBetween) = tmpStr
        End If
        
        posEnd = posEnd + LenB(StrConv(strEnd, vbFromUnicode)) + 1
        Seek fNum, posEnd
    Loop
    

'<EhFooter>
Exit Function
StrBetween_EXIT:
'</EhFooter>
End Function

'CSEH: ErrExit
Public Function LenOfFile(ByRef filename As String) As Long
'<EhHeader>
On Error GoTo lenOfFile_EXIT
'</EhHeader>
If linvblib.FileExists(filename) = False Then Exit Function
Dim tmpStr As String
Dim fNum As Integer
fNum = FreeFile()
Open filename For Binary Access Read Shared As fNum
tmpStr = StrConv(InputB(LOF(fNum), fNum), vbUnicode)
Close fNum
LenOfFile = Len(tmpStr)
'<EhFooter>
Exit Function
lenOfFile_EXIT:
'</EhFooter>
End Function

'CSEH: ErrExit
Public Function ReadAll(ByRef filename As String) As String
'<EhHeader>
'On Error GoTo ReadAll_EXIT
'</EhHeader>
If linvblib.FileExists(filename) = False Then Exit Function
Dim fNum As Integer
fNum = FreeFile
Open filename For Binary Access Read Shared As #fNum
ReadAll = StrConv(InputB(LOF(fNum), fNum), vbUnicode)
Close #fNum

'<EhFooter>
Exit Function
ReadAll_EXIT:
'</EhFooter>
End Function


