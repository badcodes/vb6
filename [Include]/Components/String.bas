Attribute VB_Name = "MString"
Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Copyright ?1996-2005 VBnet, Randy Birch, All Rights Reserved.
' Some pages may also contain other copyrights by the author.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Distribution: You can freely use this code in your own
'               applications, but you may not reproduce
'               or publish this code on any web site,
'               online service, or distribute as source
'               on any media without express permission.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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

Public Const VBQuote As String = """"

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


Public Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByRef lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpDefaultChar As String, ByVal lpUsedDefaultChar As Long) As Long
Public Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
Public Const CP_UTF8 = 65001


Public Enum IfStringNotFound
    ReturnOriginalStr = 1
    ReturnEmptyStr = 0
End Enum

Public Function rdel(ByRef theSTR As String) As String

    Dim a As String
    rdel = theSTR

    If rdel = "" Then Exit Function
    a = Right$(rdel, 1)

    Do Until a <> Chr$(0) And a <> Chr$(32) And a <> Chr$(10) And a <> Chr$(13)
        rdel = Left$(rdel, Len(rdel) - 1)
        a = Right$(rdel, 1)
    Loop

End Function

Public Function ldel(ByRef theSTR As String) As String

    Dim a As String
    ldel = theSTR

    If ldel = "" Then Exit Function
    a = Left$(ldel, 1)

    Do Until a <> Chr$(0) And a <> Chr$(32) And a <> Chr$(10) And a <> Chr$(13)
        ldel = Right$(ldel, Len(ldel) - 1)
        a = Left$(ldel, 1)
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


Function Strnum(Num As Integer, numnum As Integer) As String

    Strnum = LTrim$(str$(Num))

    If Len(Strnum) >= numnum Then Exit Function
    Strnum = String$(numnum - Len(Strnum), "0") + Strnum

End Function

Public Function MYinstr(strBig As String, strSmall As String) As Boolean

    Dim i As Long
    Dim strcount As Integer
    Dim strSmallOne() As String

    If strSmall = "" Then MYinstr = True: Exit Function
    strSmallOne = Split(strSmall, ",")
    strcount = UBound(strSmallOne)

    For i = 0 To strcount

        If InStr(1, strBig, strSmallOne(i), vbTextCompare) > 0 Then MYinstr = True: Exit Function
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
            strBetween = Mid$(theSTR, pos1 + Len(strStart), pos2 - pos1 - Len(strStart))
        End If

    End If

End Function

Public Function bddir(dirname As String) As String

    bddir = dirname

    If Right$(bddir, 1) <> "\" Then bddir = bddir + "\"

End Function

Public Function VBColorToRGB(vbcolor As Long) As String

    Dim colorstr As String
    colorstr = Hex$(vbcolor)

    If Len(colorstr) > 6 Then VBColorToRGB = colorstr: Exit Function
    colorstr = String$(6 - Len(colorstr), "0") + colorstr
    VBColorToRGB = Right$(colorstr, 2) + Mid$(colorstr, 3, 2) + Left$(colorstr, 2)

End Function

Public Function charCountInStr(ByRef strSource, ByVal charSearchFor) As Long

    Dim lsSourceLen As Long
    Dim lsSearchForLen As Long
    Dim lfor As Long
    lsSourceLen = Len(strSource)
    lsSearchForLen = Len(charSearchFor)

    If lsSearchForLen < 1 Then Exit Function

    If lsSourceLen < 1 Then Exit Function
    charSearchFor = Left$(charSearchFor, 1)

    For lfor = 1 To lsSourceLen

        If Mid$(strSource, lfor, 1) = charSearchFor Then charCountInStr = charCountInStr + 1
    Next

End Function

Public Function slashCountInstr(ByRef strSource) As Long

    'count "\" and "/" in the  strSource
    slashCountInstr = charCountInStr(strSource, "\")
    slashCountInstr = slashCountInstr + charCountInStr(strSource, "/")

End Function

Public Function UTF8Encoding(ByRef szString As String) As String

    Dim szChar As String
    Dim szTemp As String
    Dim szCode As String
    Dim szHex As String
    Dim szBin As String
    Dim iCount1 As Integer
    Dim iCount2 As Integer
    Dim iStrLen1 As Integer
    Dim iStrLen2 As Integer
    Dim lResult As Long
    Dim lAscVal As Long
    szString = Trim$(szString)
    iStrLen1 = Len(szString)

    For iCount1 = 1 To iStrLen1
        szChar = Mid$(szString, iCount1, 1)
        lAscVal = AscW(szChar)

        If lAscVal >= &H0 And lAscVal <= &HFF Then

            If (lAscVal >= &H30 And lAscVal <= &H39) Or _
               (lAscVal >= &H41 And lAscVal <= &H5A) Or _
               (lAscVal >= &H61 And lAscVal <= &H7A) Then
                szCode = szCode & szChar
            Else
                szCode = szCode & "%" & Hex$(AscW(szChar))
            End If

        Else
            szHex = Hex$(AscW(szChar))
            iStrLen2 = Len(szHex)

            For iCount2 = 1 To iStrLen2
                szChar = Mid$(szHex, iCount2, 1)

                Select Case szChar
                Case Is = "0"
                    szBin = szBin & "0000"
                Case Is = "1"
                    szBin = szBin & "0001"
                Case Is = "2"
                    szBin = szBin & "0010"
                Case Is = "3"
                    szBin = szBin & "0011"
                Case Is = "4"
                    szBin = szBin & "0100"
                Case Is = "5"
                    szBin = szBin & "0101"
                Case Is = "6"
                    szBin = szBin & "0110"
                Case Is = "7"
                    szBin = szBin & "0111"
                Case Is = "8"
                    szBin = szBin & "1000"
                Case Is = "9"
                    szBin = szBin & "1001"
                Case Is = "A"
                    szBin = szBin & "1010"
                Case Is = "B"
                    szBin = szBin & "1011"
                Case Is = "C"
                    szBin = szBin & "1100"
                Case Is = "D"
                    szBin = szBin & "1101"
                Case Is = "E"
                    szBin = szBin & "1110"
                Case Is = "F"
                    szBin = szBin & "1111"
                Case Else
                End Select

            Next

            szTemp = "1110" & Left$(szBin, 4) & "10" & Mid$(szBin, 5, 6) & "10" & Right$(szBin, 6)

            For iCount2 = 1 To 24

                If Mid$(szTemp, iCount2, 1) = "1" Then
                    lResult = lResult + 1 * 2 ^ (24 - iCount2)
                Else
                    lResult = lResult + 0 * 2 ^ (24 - iCount2)
                End If

            Next

            szTemp = Hex$(lResult)
            szCode = szCode & "%" & Left$(szTemp, 2) & "%" & Mid$(szTemp, 3, 2) & "%" & Right$(szTemp, 2)
        End If

        szBin = vbNullString
        lResult = 0
    Next

    UTF8Encoding = szCode

End Function

Public Function EncodeURI(ByVal S As String) As String

    Dim i As Long
    Dim lLength As Long
    Dim lBufferSize As Long
    Dim lResult As Long
    Dim abUTF8() As Byte
    EncodeURI = ""
    lLength = Len(S)

    If lLength = 0 Then Exit Function
    lBufferSize = lLength * 3 + 1
    ReDim abUTF8(lBufferSize - 1)
    lResult = WideCharToMultiByte(CP_UTF8, 0, StrPtr(S), lLength, abUTF8(0), lBufferSize, vbNullString, 0)

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

Public Function DecodeUrl(ByVal S As String, lCodePage As Long) As String

    On Error Resume Next
    Dim lRet As Long
    Dim lLength As Long
    Dim sL As Long
    Dim sDecode As String
    Dim lBufferSize As Long
    Dim abUTF8() As Byte
    Dim i As Long
    Dim v As Variant
    v = Split(S, "%")
    lLength = UBound(v)

    If lLength <= 0 Then
        DecodeUrl = S
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
Public Function RightLeft(ByRef str As String, RFind As String, Optional Compare As VbCompareMethod = vbBinaryCompare, Optional RetError As IfStringNotFound = ReturnOriginalStr) As String

    Dim K As Long
    K = InStrRev(str, RFind, , Compare)

    If K = 0 Then
        RightLeft = IIf(RetError = ReturnOriginalStr, str, "")
    Else
        RightLeft = Left$(str, K - 1)
    End If

End Function

' Search from end to beginning and return the right side of the string
Public Function RightRight(ByRef str As String, RFind As String, Optional Compare As VbCompareMethod = vbBinaryCompare, Optional RetError As IfStringNotFound = ReturnOriginalStr) As String

    Dim K As Long
    K = InStrRev(str, RFind, , Compare)

    If K = 0 Then
        RightRight = IIf(RetError = ReturnOriginalStr, str, "")
    Else
        RightRight = Mid$(str, K + 1, Len(str))
    End If

End Function

' Search from the beginning to end and return the left size of the string
Public Function LeftLeft(ByRef str As String, LFind As String, Optional Compare As VbCompareMethod = vbBinaryCompare, Optional RetError As IfStringNotFound = ReturnOriginalStr) As String

    Dim K As Long
    K = InStr(1, str, LFind, Compare)

    If K = 0 Then
        LeftLeft = IIf(RetError = ReturnOriginalStr, str, "")
    Else
        LeftLeft = Left$(str, K - 1)
    End If

End Function

' Search from the beginning to end and return the right size of the string
Public Function LeftRight(ByRef str As String, LFind As String, Optional Compare As VbCompareMethod = vbBinaryCompare, Optional RetError As IfStringNotFound = ReturnOriginalStr) As String

    Dim K As Long
    K = InStr(1, str, LFind, Compare)

    If K = 0 Then
        LeftRight = IIf(RetError = ReturnOriginalStr, str, "")
    Else
        LeftRight = Right$(str, (Len(str) - Len(LFind)) - K + 1)
    End If

End Function

' Search from the beginning to end and return from StrFrom string to StrTo string
' both strings (StrFrom and StrTo) must be found in order to be successfull
Public Function LeftRange(ByRef str As String, StrFrom As String, StrTo As String, Optional Compare As VbCompareMethod = vbBinaryCompare, Optional RetError As IfStringNotFound = ReturnOriginalStr) As String

    Dim K As Long, Q As Long
    K = InStr(1, str, StrFrom, Compare)

    If K > 0 Then
        Q = InStr(K + Len(StrFrom), str, StrTo, Compare)

        If Q > K Then
            LeftRange = Mid$(str, K + Len(StrFrom), (Q - K) - Len(StrFrom))
        Else
            LeftRange = IIf(RetError = ReturnOriginalStr, str, "")
        End If

    Else
        LeftRange = IIf(RetError = ReturnOriginalStr, str, "")
    End If

End Function

' Search from the end to beginning and return from StrFrom string to StrTo string
' both strings (StrFrom and StrTo) must be found in order to be successfull
Public Function RightRange(ByRef str As String, StrFrom As String, StrTo As String, Optional Compare As VbCompareMethod = vbBinaryCompare, Optional RetError As IfStringNotFound = ReturnOriginalStr) As String

    Dim K As Long, Q As Long
    K = InStrRev(str, StrTo, , Compare)

    If K > 0 Then
        Q = InStrRev(str, StrFrom, K, Compare)

        If Q > 0 Then
            RightRange = Mid$(str, Q + Len(StrFrom), (K - Q) - Len(StrFrom))
        Else
            RightRange = IIf(RetError = ReturnOriginalStr, str, "")
        End If

    Else
        RightRange = IIf(RetError = ReturnOriginalStr, str, "")
    End If

End Function

' SOUNDEX, used in SQL mostly, and dictionaries
' useful to find words that sound the same
Public Function SOUNDEX(Word As String) As String

    Dim K As Integer, PrevNum As Integer, Num As Integer, LLetter As String
    Dim SoundX As String
    Dim wl As Long
    wl = Len(Word)

    For K = 2 To wl
        LLetter = LCase$(Mid$(Word, K, 1))

        Select Case LLetter
        Case "b", "f", "p", "v"
            Num = 1
        Case "c", "e", "g", "j", "k", "q", "s", "x", "z"
            Num = 2
        Case "d", "t"
            Num = 3
        Case "l"
            Num = 4
        Case "m", "n"
            Num = 5
        Case "r"
            Num = 6
        Case "a", "e", "i", "o", "u"
            Num = 7
        End Select

        If PrevNum <> Num Then
            PrevNum = Num
            SoundX = SoundX & Num
        End If

    Next

    SoundX = Replace(SoundX, "7", "", , , vbBinaryCompare)
    SOUNDEX = UCase$(Left$(Word, 1)) & Left$(SoundX & "000", 3)

End Function

Public Function SOUNDEX2(Word As String, Optional StrLength As Integer = 8) As String

    Dim K As Integer, PrevNum As Integer, Num As Integer, LLetter As String
    Dim SoundX As String
    Dim lw As Long
    lw = Len(Word)

    For K = 2 To lw
        LLetter = LCase$(Mid$(Word, K, 1))

        Select Case LLetter
        Case "b", "f", "p", "v"
            Num = 1
        Case "c", "e", "g", "j", "k", "q", "s", "x", "z"
            Num = 2
        Case "d", "t"
            Num = 3
        Case "l"
            Num = 4
        Case "m", "n"
            Num = 5
        Case "r"
            Num = 6
        Case "a", "e", "i", "o", "u"
            Num = 7
        End Select

        If PrevNum <> Num Then
            PrevNum = Num
            SoundX = SoundX & Num
        End If

    Next

    SoundX = Replace(SoundX, "7", "", , , vbBinaryCompare)
    SOUNDEX2 = UCase$(Left$(Word, 1)) & Left$(SoundX & String$(StrLength - 1, "0"), StrLength - 1)

End Function


Public Function cleanFilename(sFilenameDirty As String) As String

    cleanFilename = Replace(sFilenameDirty, ":", "£º")
    cleanFilename = Replace(cleanFilename, "?", "£¿")
    cleanFilename = Replace(cleanFilename, "\", "")
    cleanFilename = Replace(cleanFilename, "/", "")
    cleanFilename = Replace(cleanFilename, "|", "")
    cleanFilename = Replace(cleanFilename, ">", "")
    cleanFilename = Replace(cleanFilename, "<", "")
    cleanFilename = Replace(cleanFilename, "*", "")
    cleanFilename = Replace(cleanFilename, Chr$(34), "")

End Function

Public Function replaceSlash(ByVal sSource As String) As String

    replaceSlash = Replace(sSource, "/", "\")

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

Public Function CBoolStr(S As String) As Boolean

    If S = "" Then S = "False"
    CBoolStr = CBool(S)

End Function

Public Function CLngStr(S As String) As Long

    If S = "" Then S = "0"
    CLngStr = CLng(S)

End Function

Public Function toUnixPath(sDosPath As String) As String

    toUnixPath = Replace(sDosPath, "\", "/")

End Function

Public Function toDosPath(sUnixPath As String) As String

    toDosPath = Replace(sUnixPath, "/", "\")

End Function

Function lenchar(thechar As String) As Long

    If Asc(thechar) < 0 Then
        lenchar = 2
    Else
        lenchar = 1
    End If

End Function

Public Sub ParseHTML(HTML As String)
'We go through the HTML, character by character
'checking first for <, then for spaces, then
'quotation marks, and finally /. As we find
'them we fire events and continue parsing.
'
'Clean code with few relevant comments is better than
'unwieldy code commented to death, IMHO
'
Dim IsValue, IsProperty, IsTag, RaisedTagBegin As Boolean
Dim i As Long
Dim CurrentChar As String
Dim CurrentProperty As String
Dim CurrentPropertyValue As String
Dim CurrentTag As String
Dim CurrentText As String
'Remove tabs and returns, they have no place in HTML
HTML = Replace(HTML, vbCrLf, "")
HTML = Replace(HTML, vbTab, "")
'Start our searching
For i = 1 To Len(HTML)
    CurrentChar = Mid(HTML, i, 1)
    If IsTag = True Then
        If IsProperty = True Then
            If IsValue = True Then
                If CurrentChar = Chr(34) Then
                    IsValue = False
                    IsProperty = False
                    CurrentPropertyValue = Trim(CurrentPropertyValue)
                    CurrentProperty = Trim(CurrentProperty)
                    CurrentPropertyValue = ""
                    CurrentProperty = ""
                Else
                    CurrentPropertyValue = CurrentPropertyValue & CurrentChar
                End If
            ElseIf CurrentChar = Chr(34) Then
                IsValue = True
            Else
                CurrentProperty = CurrentProperty & CurrentChar
            End If
        Else
            If CurrentChar = " " Then
                IsProperty = True
                CurrentTag = Trim(CurrentTag)
                CurrentTag = CurrentTag
                If RaisedTagBegin = False Then
                     RaisedTagBegin = True
                End If
            ElseIf CurrentChar = ">" Then
                IsTag = False
                If Left(CurrentTag, 1) = "/" Then
                 ElseIf RaisedTagBegin = False Then
                    RaisedTagBegin = True
                Else
                 End If
                CurrentTag = ""
                
            Else
                CurrentTag = CurrentTag & CurrentChar
            End If
        End If
    Else
        If CurrentChar = "<" Then
            IsTag = True
            RaisedTagBegin = False
            If Trim(CurrentText) <> "" Then
                 CurrentText = ""
            End If
        Else
            CurrentText = CurrentText & CurrentChar
        End If
    End If
Next i
End Sub


Public Function Quote(ByVal sNaked As String, Optional fSingleQuote As Boolean = False) As String
    Dim c As String
    If fSingleQuote Then c = "'" Else c = Chr$(34)
    Quote = c & sNaked & c
End Function


Public Function LineToWords(ByRef sLine As String) As String()

End Function

Public Function Rotate13(ByRef str As String) As String
    Static CODE_A_U As Integer
    Static CODE_A_L As Integer
    Static CODE_Z_U As Integer
    Static CODE_Z_L As Integer
    Static L_TO_U As Integer
    If Not CODE_A_U > 0 Then
        CODE_A_U = Asc("A")
        CODE_A_L = Asc("a")
        CODE_Z_U = Asc("Z")
        CODE_Z_L = Asc("z")
        L_TO_U = CODE_A_U - CODE_A_L
    End If
    Dim iLen As Long
    Dim i As Long
    Dim c As Integer
    Dim sResult As String
    iLen = Len(str)
    sResult = Space(iLen)
    For i = 1 To iLen
        c = Asc(Mid$(str, i, 1))
        If c >= CODE_A_L And c <= CODE_Z_L Then
            c = c + 13
            If (c > CODE_Z_L) Then c = c - 26
        ElseIf c >= CODE_A_U And c <= CODE_Z_U Then
            c = c + 13
            If (c > CODE_Z_U) Then c = c - 26
        End If
        Mid(sResult, i, 1) = Chr(c)
    Next
    Rotate13 = sResult
End Function
