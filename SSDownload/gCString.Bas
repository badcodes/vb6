Attribute VB_Name = "MgCString"
'--------------------------------------------------------------------------------
'    Component  : MgCString
'
'    Description: String Function Modules
'
'    Author   : xrLin
'
'    Date     : 2005-2006
'--------------------------------------------------------------------------------
Option Explicit

Private Const MAX_PATH                   As Long = 260
Private Const ERROR_SUCCESS              As Long = 0

'Treat entire URL param as one URL segment
Private Const URL_ESCAPE_SEGMENT_ONLY    As Long = &H2000
Private Const URL_ESCAPE_PERCENT         As Long = &H1000
Private Const URL_UNESCAPE_INPLACE       As Long = &H100000

'escape #'s in paths
Private Const URL_INTERNAL_PATH          As Long = &H800000
Private Const URL_DONT_ESCAPE_EXTRA_INFO As Long = &H2000000
Private Const URL_ESCAPE_SPACES_ONLY     As Long = &H4000000
Private Const URL_DONT_SIMPLIFY          As Long = &H8000000

'Converts unsafe characters,
'such as spaces, into their
'corresponding escape sequences.
Private Declare Function UrlEscape Lib "shlwapi" _
   Alias "UrlEscapeA" _
  (ByVal pszURL As String, _
   ByVal pszEscaped As String, _
   pcchEscaped As Long, _
   ByVal dwFlags As Long) As Long

'Converts escape sequences back into
'ordinary characters.
Private Declare Function UrlUnescape Lib "shlwapi" _
   Alias "UrlUnescapeA" _
  (ByVal pszURL As String, _
   ByVal pszUnescaped As String, _
   pcchUnescaped As Long, _
   ByVal dwFlags As Long) As Long


Private Declare Function WideCharToMultiByte Lib "kernel32" ( _
            ByVal CodePage As Long, ByVal dwFlags As Long, _
            ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, _
            ByRef lpMultiByteStr As Any, ByVal cchMultiByte As Long, _
            ByVal lpDefaultChar As String, ByVal lpUsedDefaultChar As Long _
            ) As Long
Private Declare Function MultiByteToWideChar Lib "kernel32" ( _
            ByVal CodePage As Long, ByVal dwFlags As Long, _
            ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, _
            ByVal lpWideCharStr As Long, ByVal cchWideChar As Long _
            ) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
            lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long _
            )
Public Enum CodePage
     CP_UTF8 = 65001
     CP_Default = 0
End Enum

Public Enum IfStringNotFound
    ReturnOriginalStr = 1
    ReturnEmptyStr = 0
End Enum

Public Function rdel(ByRef theSTR As String) As String

    Dim A As String
    rdel = theSTR

    If rdel = "" Then Exit Function
    A = Right$(rdel, 1)

    Do Until A <> Chr$(0) And A <> Chr$(32) And A <> Chr$(10) And A <> Chr$(13)
        rdel = Left$(rdel, Len(rdel) - 1)
        A = Right$(rdel, 1)
    Loop

End Function

Public Function ldel(ByRef theSTR As String) As String

    Dim A As String
    ldel = theSTR

    If ldel = "" Then Exit Function
    A = Left$(ldel, 1)

    Do Until A <> Chr$(0) And A <> Chr$(32) And A <> Chr$(10) And A <> Chr$(13)
        ldel = Right$(ldel, Len(ldel) - 1)
        A = Left$(ldel, 1)
    Loop

End Function

Public Function LeftDelete(theSTR As String, sDel As String) As String

    LeftDelete = theSTR

    If LeftDelete = "" Then Exit Function

    Do Until Left$(LeftDelete, Len(sDel)) <> sDel
        LeftDelete = Right$(LeftDelete, Len(LeftDelete) - Len(sDel))
    Loop

End Function

Public Function RightDelete(theSTR As String, sDel As String) As String

    RightDelete = theSTR

    If RightDelete = "" Then Exit Function

    Do Until Right$(RightDelete, Len(sDel)) <> sDel
        RightDelete = Left$(RightDelete, Len(RightDelete) - Len(sDel))
    Loop

End Function

Function StrNum(Num As Integer, lenNum As Integer) As String

    StrNum = LTrim$(Str$(Num))

    If Len(StrNum) >= lenNum Then
        StrNum = Left$(StrNum, lenNum)
    Else
        StrNum = String$(lenNum - Len(StrNum), "0") + StrNum
    End If

End Function

Public Function MyInstr(strBig As String, strList As String, Optional strListSep As String = ",", Optional cmp As VbCompareMethod = vbBinaryCompare) As Boolean

    Dim i As Long
    Dim strcount As Integer
    Dim strSmallOne() As String

    If strList = "" Then MyInstr = True: Exit Function
    strSmallOne = Split(strList, strListSep)
    strcount = UBound(strSmallOne)

    For i = 0 To strcount

        If InStr(1, strBig, strSmallOne(i), cmp) > 0 Then MyInstr = True: Exit Function
    Next

End Function

Public Function bddir(dirname As String) As String

    Dim slash As String
    bddir = dirname

    If InStr(dirname, "/") > 0 Then slash = "/" Else slash = "\"

    If Right$(bddir, 1) <> slash Then bddir = bddir & slash

End Function

Public Function bdUnixDir(ByVal FirstPart As String, ByVal SecPart As String) As String

    Const sSlash = "/"
    FirstPart = toUnixPath(FirstPart)
    SecPart = toUnixPath(SecPart)

    If Right$(FirstPart, 1) <> sSlash Then FirstPart = FirstPart & sSlash

    If Left$(SecPart, 1) = sSlash Then SecPart = Left$(SecPart, Len(SecPart) - 1)
    bdUnixDir = FirstPart & SecPart

End Function

Public Function bdDosDir(ByVal FirstPart As String, ByVal SecPart As String) As String

    Const sSlash = "\"
    FirstPart = toDosPath(FirstPart)
    SecPart = toDosPath(SecPart)

    If Right$(FirstPart, 1) <> sSlash Then FirstPart = FirstPart & sSlash

    If Left$(SecPart, 1) = sSlash Then SecPart = Left$(SecPart, Len(SecPart) - 1)
    bdDosDir = FirstPart & SecPart

End Function

Public Function VBColorToRGB(vbcolor As Long) As String

    Dim colorstr As String
    colorstr = Hex$(vbcolor)

    If Len(colorstr) > 6 Then VBColorToRGB = colorstr: Exit Function
    colorstr = String$(6 - Len(colorstr), "0") + colorstr
    VBColorToRGB = Right$(colorstr, 2) + Mid$(colorstr, 3, 2) + Left$(colorstr, 2)

End Function

Public Function InStrCount(ByRef strBig As String, ByRef strSmall As String, Optional cmp As VbCompareMethod = vbBinaryCompare) As Long

    Dim lenBig As Long, lenSmall As Long
    Dim posStart As Long, nextPos As Long
    lenBig = Len(strBig)
    lenSmall = Len(strSmall)

    If lenBig < lenSmall Or lenSmall = 0 Then Exit Function
    posStart = InStr(1, strBig, strSmall, cmp)

    Do Until posStart < 1
        InStrCount = InStrCount + 1
        nextPos = posStart + 1

        If nextPos > lenBig Then Exit Do
        posStart = InStr(nextPos, strBig, strSmall, cmp)
    Loop

End Function

Public Function charCountInStr(ByRef strSource As String, ByVal charSearchFor As String, Optional cmp As VbCompareMethod = vbBinaryCompare) As Long

    charSearchFor = Left$(charSearchFor, 1)
    charCountInStr = InStrCount(strSource, charSearchFor, cmp)

End Function

Public Function slashCountInstr(ByRef strSource As String) As Long

    'count "\" and "/" in the  strSource
    slashCountInstr = charCountInStr(strSource, "\")
    slashCountInstr = slashCountInstr + charCountInStr(strSource, "/")

End Function

Public Function EncodeURI(ByVal s As String) As String

    Dim i As Long
    Dim lLength As Long
    Dim lBufferSize As Long
    Dim lResult As Long
    Dim abUTF8() As Byte
    EncodeURI = ""
    lLength = Len(s)

    If lLength = 0 Then Exit Function
    lBufferSize = lLength * 3 + 1
    ReDim abUTF8(lBufferSize - 1)
    lResult = WideCharToMultiByte(CP_UTF8, 0, StrPtr(s), lLength, abUTF8(0), lBufferSize, vbNullString, 0)

    If lResult <> 0 Then
        lResult = lResult - 1
        ReDim Preserve abUTF8(lResult)
        Dim lStart As Long
        Dim lEnd As Long
        lStart = LBound(abUTF8)
        lEnd = UBound(abUTF8)

        For i = lStart To lEnd
            EncodeURI = EncodeURI & "%" & Hex$(abUTF8(i))
        Next

    End If

End Function

Public Function DecodeUrl(ByVal s As String, lCodePage As CodePage) As String

    On Error Resume Next
    Dim lRet As Long
    Dim lLength As Long
    Dim sL As Long
    Dim sDecode As String
    Dim lBufferSize As Long
    Dim abUTF8() As Byte
    Dim i As Long
    Dim v As Variant
    v = Split(s, "%")
    lLength = UBound(v)

    If lLength <= 0 Then
        DecodeUrl = s
        Exit Function
    End If

    DecodeUrl = v(0)
    sL = -1

    For i = 1 To lLength

        If Len(v(i)) = 2 Then
            sL = sL + 1
            ReDim Preserve abUTF8(sL)
            abUTF8(sL) = CByte("&H" & v(i))
        Else
            sL = sL + 1
            ReDim Preserve abUTF8(sL)
            abUTF8(sL) = CByte("&H" & Left$(v(i), 2))
            lBufferSize = (sL + 1) * 2
            sDecode = String$(lBufferSize, Chr$(0))
            lRet = MultiByteToWideChar(lCodePage, 0, VarPtr(abUTF8(0)), sL + 1, StrPtr(sDecode), lBufferSize)

            If lRet <> 0 Then DecodeUrl = DecodeUrl & Left$(sDecode, lRet)
            sL = -1
            sDecode = ""
            DecodeUrl = DecodeUrl & Right$(v(i), Len(v(i)) - 2)
            Erase abUTF8
        End If

    Next

    If sL > 0 Then
        lBufferSize = (sL + 1) * 2
        sDecode = String$(lBufferSize, Chr$(0))
        lRet = MultiByteToWideChar(lCodePage, 0, VarPtr(abUTF8(0)), sL + 1, StrPtr(sDecode), lBufferSize)

        If lRet <> 0 Then DecodeUrl = DecodeUrl & Left$(sDecode, lRet)
    End If

End Function

' Search from end to beginning, and return the left side of the string
Public Function RightLeft(ByRef Str As String, RFind As String, Optional Compare As VbCompareMethod = vbBinaryCompare, Optional RetError As IfStringNotFound = ReturnOriginalStr) As String

    Dim K As Long
    K = InStrRev(Str, RFind, , Compare)

    If K = 0 Then
        RightLeft = IIf(RetError = ReturnOriginalStr, Str, "")
    Else
        RightLeft = Left$(Str, K - 1)
    End If

End Function

' Search from end to beginning and return the right side of the string
Public Function RightRight(ByRef Str As String, RFind As String, Optional Compare As VbCompareMethod = vbBinaryCompare, Optional RetError As IfStringNotFound = ReturnOriginalStr) As String

    Dim K As Long
    K = InStrRev(Str, RFind, , Compare)

    If K = 0 Then
        RightRight = IIf(RetError = ReturnOriginalStr, Str, "")
    Else
        RightRight = Mid$(Str, K + 1, Len(Str))
    End If

End Function

' Search from the beginning to end and return the left size of the string
Public Function LeftLeft(ByRef Str As String, LFind As String, Optional Compare As VbCompareMethod = vbBinaryCompare, Optional RetError As IfStringNotFound = ReturnOriginalStr) As String

    Dim K As Long
    K = InStr(1, Str, LFind, Compare)

    If K = 0 Then
        LeftLeft = IIf(RetError = ReturnOriginalStr, Str, "")
    Else
        LeftLeft = Left$(Str, K - 1)
    End If

End Function

' Search from the beginning to end and return the right size of the string
Public Function LeftRight(ByRef Str As String, LFind As String, Optional Compare As VbCompareMethod = vbBinaryCompare, Optional RetError As IfStringNotFound = ReturnOriginalStr) As String

    Dim K As Long
    K = InStr(1, Str, LFind, Compare)

    If K = 0 Then
        LeftRight = IIf(RetError = ReturnOriginalStr, Str, "")
    Else
        LeftRight = Right$(Str, (Len(Str) - Len(LFind)) - K + 1)
    End If

End Function

' Search from the beginning to end and return from StrFrom string to StrTo string
' both strings (StrFrom and StrTo) must be found in order to be successfull
Public Function LeftRange(ByRef Str As String, strFrom As String, strTo As String, Optional Compare As VbCompareMethod = vbBinaryCompare, Optional RetError As IfStringNotFound = ReturnOriginalStr) As String

    Dim K As Long, Q As Long
    K = InStr(1, Str, strFrom, Compare)

    If K > 0 Then
        Q = InStr(K + Len(strFrom), Str, strTo, Compare)

        If Q > K Then
            LeftRange = Mid$(Str, K + Len(strFrom), (Q - K) - Len(strFrom))
        Else
            LeftRange = IIf(RetError = ReturnOriginalStr, Str, "")
        End If

    Else
        LeftRange = IIf(RetError = ReturnOriginalStr, Str, "")
    End If

End Function

' Search from the end to beginning and return from StrFrom string to StrTo string
' both strings (StrFrom and StrTo) must be found in order to be successfull
Public Function RightRange(ByRef Str As String, strFrom As String, strTo As String, Optional Compare As VbCompareMethod = vbBinaryCompare, Optional RetError As IfStringNotFound = ReturnOriginalStr) As String

    Dim K As Long, Q As Long
    K = InStrRev(Str, strTo, , Compare)

    If K > 0 Then
        Q = InStrRev(Str, strFrom, K, Compare)

        If Q > 0 Then
            RightRange = Mid$(Str, Q + Len(strFrom), (K - Q) - Len(strFrom))
        Else
            RightRange = IIf(RetError = ReturnOriginalStr, Str, "")
        End If

    Else
        RightRange = IIf(RetError = ReturnOriginalStr, Str, "")
    End If

End Function

Public Function EscapeUrl(ByVal sUrl As String) As String

    Dim buff As String
    Dim dwSize As Long
    Dim dwFlags As Long

    If Len(sUrl) > 0 Then
        buff = Space$(MAX_PATH)
        dwSize = Len(buff)
        dwFlags = URL_ESCAPE_PERCENT

        If UrlEscape(sUrl, _
           buff, _
           dwSize, _
           dwFlags) = ERROR_SUCCESS Then
            EscapeUrl = Left$(buff, dwSize)
        End If  'UrlEscape

    End If  'Len(sUrl)

End Function

Public Function UnescapeUrl(ByVal sUrl As String) As String

    Dim buff As String
    Dim dwSize As Long
    Dim dwFlags As Long

    If Len(sUrl) > 0 Then
        buff = Space$(MAX_PATH)
        dwSize = Len(buff)
        dwFlags = URL_ESCAPE_PERCENT

        If UrlUnescape(sUrl, _
           buff, _
           dwSize, _
           dwFlags) = ERROR_SUCCESS Then
            UnescapeUrl = LeftLeft(buff, Chr(0))
        End If  'UrlUnescape

    End If  'Len(sUrl)

End Function

Public Function CBoolStr(s As String) As Boolean

    If s = "" Then s = "False"
    CBoolStr = CBool(s)

End Function

Public Function CLngStr(s As String) As Long

    If s = "" Then s = "0"
    CLngStr = CLng(s)

End Function

Public Function toUnixPath(sDosPath As String) As String

    toUnixPath = Replace(sDosPath, "\", "/")

End Function

Public Function toDosPath(sUnixPath As String) As String

    toDosPath = Replace(sUnixPath, "/", "\")

End Function

Public Function expandStr(ByVal systegCString As String) As String

    Dim stmp As String
    Dim sMass As String
    Dim pos1 As Long
    Dim pos2 As Long
    expandStr = systegCString

    Do
        pos1 = InStr(expandStr, "%")

        If pos1 = 0 Then Exit Do
        pos2 = InStr(pos1 + 1, expandStr, "%")

        If pos2 = 0 Then Exit Do
        sMass = Mid$(expandStr, pos1 + 1, pos2 - pos1 - 1)
        sMass = Environ$(sMass)
        stmp = Left$(expandStr, pos1 - 1) & sMass & Right$(expandStr, Len(expandStr) - pos2)
        expandStr = stmp
    Loop

End Function

Public Function CBytesToStr(ByRef CBytes() As Byte) As String

    Dim lUB As Long, lLb As Long
    Dim iPos As Long
    Dim bTemp() As Byte
    Dim l As Long
    lUB = UBound(CBytes)
    lLb = LBound(CBytes)

    For iPos = lLb To lUB

        If CBytes(iPos) = 0 Then Exit For
    Next

    If iPos = 0 Then
        CBytesToStr = StrConv(CBytes, vbUnicode)
    ElseIf iPos = lLb Then
        CBytesToStr = ""
    Else
        ReDim bTemp(lLb To iPos - 1)
        CopyMemory bTemp(lLb), CBytes(lLb), iPos - lLb
        CBytesToStr = StrConv(bTemp, vbUnicode)
    End If

End Function

Public Sub StrToCBytes(ByVal strUnicode As String, ByRef CBytes() As Byte)

    Dim lUB As Long, lLb As Long
    Dim bTemp() As Byte
    Dim lSize As Long
    lUB = UBound(CBytes)
    lLb = LBound(CBytes)
    bTemp = StrConv(strUnicode, vbFromUnicode)
    lSize = UBound(bTemp) + 1
    ReDim Preserve bTemp(lSize)
    bTemp(lSize) = 0

    If lSize > lUB - lLb Then
        lSize = lUB - lLb
        bTemp(lSize) = 0
    End If

    CopyMemory CBytes(lLb), bTemp(0), lSize + 1

End Sub
Public Function cleanFilename(sFilenameDirty As String) As String

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

Public Function isTypicalFileName(ByVal strTest) As Boolean
    Dim pos As Long
    Dim l As Long
    Dim c As String
    isTypicalFileName = False
    If strTest = "" Then Exit Function
    strTest = Replace$(strTest, "/", "\")
    l = Len(strTest)
    For pos = 1 To l
        c = Mid$(strTest, pos, 1)
        If c = "|" Then Exit Function
        If c = ">" Then Exit Function
        If c = "<" Then Exit Function
        If c = "*" Then Exit Function
        If c = Chr$(34) Then Exit Function
        If c = "?" Then Exit Function
    Next
    pos = InStr(strTest, ".")
    l = InStr(strTest, "\")
    If pos > l Then isTypicalFileName = True
End Function
