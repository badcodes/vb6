Attribute VB_Name = "MStrings"
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
Private Declare Function UrlEscape Lib "SHLWAPI" _
   Alias "UrlEscapeA" _
  (ByVal pszURL As String, _
   ByVal pszEscaped As String, _
   pcchEscaped As Long, _
   ByVal dwFlags As Long) As Long

'Converts escape sequences back into
'ordinary characters.
Private Declare Function UrlUnescape Lib "SHLWAPI" _
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

Private Const ZCharDoubleQuote As String = """"
Private Const ZCharSingleQuote As String = "'"
Private Declare Function StrSpn Lib "SHLWAPI" Alias "StrSpnW" ( _
        ByVal psz As Long, ByVal pszSet As Long) As Long
Private Declare Function StrCSpn Lib "SHLWAPI" Alias "StrCSpnW" ( _
        ByVal lpStr As Long, ByVal lpSet As Long) As Long

Private Const ZCharSpace As String = " "
Private Const ZCharEmpty As String = vbNullString

#If fComponent Then
' In component fNoShlWapi is a variable initialized in Class_Initialize
Private fNoShlWapi As Boolean ' = False

Private Sub Class_Initialize()
    On Error GoTo Fail
    ' Dummy call to any SHLWAPI function fails if no DLL
    Call StrSpn(StrPtr("a"), StrPtr("a"))
    Exit Sub
Fail:
    fNoShlWapi = True
End Sub

#Else
' In standard module fNoShlWapi is a Property Get that checks for DLL
Private fNotFirstTime As Boolean, fNoShlWapiI As Boolean ' = False

Private Property Get fNoShlWapi() As Boolean
    If fNotFirstTime = False Then
        fNotFirstTime = True
        On Error Resume Next
        Call StrSpn(StrPtr("a"), StrPtr("a"))
        If Err Then fNoShlWapiI = True
    End If
    fNoShlWapi = fNoShlWapiI
End Property
#End If

' GetQTokenO uses our StrSpan and StrBreak
Private Function GetQTokenO(sTarget As String, sSeps As String) As String
    ' GetQTokenO = ZCharEmpty

    ' Note that sSave and iStart must be static from call to call
    ' If first call, make copy of string
    Static sSave As String, iStart As Integer, cSave As Integer
    Dim iNew As Integer, fQuote As Integer
    If sTarget <> ZCharEmpty Then
        iStart = 1
        sSave = sTarget
        cSave = Len(sSave)
    Else
        If sSave = ZCharEmpty Then Exit Function
    End If
    ' Make sure separators includes quote
    sSeps = sSeps & ZCharDoubleQuote

    ' Find start of next token
    iNew = StrSpan(sSave, iStart, sSeps)
    If iNew Then
        ' Set position to start of token
        iStart = iNew
    Else
        ' If no new token, return empty string
        sSave = ZCharEmpty
        Exit Function
    End If
    
    ' Find end of token
    If iStart = 1 Then
        iNew = StrBreak(sSave, iStart, sSeps)
    ElseIf Mid$(sSave, iStart - 1, 1) = ZCharDoubleQuote Then
        iNew = StrBreak(sSave, iStart, ZCharDoubleQuote)
    Else
        iNew = StrBreak(sSave, iStart, sSeps)
    End If

    If iNew = 0 Then
        ' If no end of token, set to end of string
        iNew = cSave + 1
    End If
    ' Cut token out of sTarget string
    GetQTokenO = Mid$(sSave, iStart, iNew - iStart)
    
    ' Set new starting position
    iStart = iNew

End Function

' New GetQToken uses faster StrSpn and StrCSpn from SHLWAPI.DLL
Function GetQToken(sTarget As String, sSeps As String) As String
    ' GetQToken = ZCharEmpty
    If fNoShlWapi Then
        GetQToken = GetQTokenO(sTarget, sSeps)
        Exit Function
    End If
    
    ' Note that sSave, pSave, pCur, and cSave must be static between calls
    Static sSave As String, pSave As Long, pCur As Long, cSave As Long
    ' First time through save start and length of string
    If sTarget <> ZCharEmpty Then
        ' Save in case sTarget is moveable string (Command$)
        sSave = sTarget
        pSave = StrPtr(sSave)
        pCur = pSave
        cSave = Len(sSave)
    Else
        ' Quit if past end (also catches null or empty target)
        If pCur >= pSave + (cSave * 2) Then Exit Function
    End If
    ' Make sure separators includes quote
    Dim sSepsNew As String, pSeps As Long
    sSepsNew = sSeps & ZCharDoubleQuote
    pSeps = StrPtr(sSepsNew)

    ' Get current character
    Dim pNew As Long, C As Long
    ' Find start of next token
    C = StrSpn(pCur, pSeps)
    ' Set position to start of token
    If C Then pCur = pCur + (C * 2)
    
    Dim ch As Integer
    Const chQuote = 34  ' Asc("""")
    CopyMemory ch, ByVal pCur - 2, 2
    ' Check first character for quote, then find end of token
    If ch = chQuote Then
        C = StrCSpn(pCur, StrPtr(ZCharDoubleQuote))
    Else
        C = StrCSpn(pCur, pSeps)
    End If
    ' If token length is zero, we're at end
    If C = 0 Then Exit Function
    
    ' Cut token out of target string
    GetQToken = String$(C, 0)
    CopyMemory ByVal StrPtr(GetQToken), ByVal pCur, C * 2
    ' Set new starting position
    pCur = pCur + (C * 2)

End Function

' Old GetTokenO uses our StrSpan and StrBreak
Private Function GetTokenO(sTarget As String, sSeps As String) As String
    ' GetTokenO = ZCharEmpty
    
    ' Note that sSave and iStart must be static from call to call
    ' If first call, make copy of string
    Static sSave As String, iStart As Integer, cSave As Integer
    
    If sTarget <> ZCharEmpty Then
        iStart = 1
        sSave = sTarget
        cSave = Len(sSave)
    Else
        If sSave = ZCharEmpty Then Exit Function
    End If
    
    ' Find start of next token
    Dim iNew As Integer
    iNew = StrSpan(sSave, iStart, sSeps)
    If iNew Then
        ' Set position to start of token
        iStart = iNew
    Else
        ' If no new token, return empty string
        sSave = ZCharEmpty
        Exit Function
    End If
    
    ' Find end of token
    iNew = StrBreak(sSave, iStart, sSeps)
    If iNew = 0 Then
        ' If no end of token, set to end of string
        iNew = cSave + 1
    End If
    
    ' Cut token out of sTarget string
    GetTokenO = Mid$(sSave, iStart, iNew - iStart)
    ' Set new starting position
    iStart = iNew

End Function

' New GetToken uses faster StrSpn and StrCSpn from SHLWAPI.DLL
Function GetToken(sTarget As String, sSeps As String) As String
    
    ' If SHLWAPI.DLL not available, delegate to Basic-only GetToken
    If fNoShlWapi Then
        GetToken = GetTokenO(sTarget, sSeps)
        Exit Function
    End If
    
    ' Note that sSave, pSave, pCur, and cSave must be static between calls
    Static sSave As String, pSave As Long, pCur As Long, cSave As Long
    ' First time through save start and length of string
    If sTarget <> ZCharEmpty Then
        ' Save in case sTarget is moveable string (Command$)
        sSave = sTarget
        pSave = StrPtr(sSave)
        pCur = pSave
        cSave = Len(sSave)
    Else
        ' Quit if past end (also catches null or empty target)
        If pCur >= pSave + (cSave * 2) Then Exit Function
    End If
    
    ' Find start of next token
    Dim pNew As Long, C As Long, pSeps As Long
    pSeps = StrPtr(sSeps)
    C = StrSpn(pCur, pSeps)
    ' Set position to start of token
    If C Then pCur = pCur + (C * 2)
    
    ' Find end of token
    C = StrCSpn(pCur, pSeps)
    ' If token length is zero, we're at end
    If C = 0 Then Exit Function
    
    ' Cut token out of target string
    GetToken = String$(C, 0)
    CopyMemory ByVal StrPtr(GetToken), ByVal pCur, C * 2
    ' Set new starting position
    pCur = pCur + (C * 2)

End Function

' StrBreak and StrSpan are used by GetTokenO, but can be called by clients
Function StrBreak(sTarget As String, ByVal iStart As Integer, _
                  sSeps As String) As Integer
    
    Dim cTarget As Integer
    cTarget = Len(sTarget)
   
    ' Look for end of token (first character that is a separator)
    Do While InStr(sSeps, Mid$(sTarget, iStart, 1)) = 0
        If iStart > cTarget Then
            StrBreak = 0
            Exit Function
        Else
            iStart = iStart + 1
        End If
    Loop
    StrBreak = iStart

End Function

Function StrSpan(sTarget As String, ByVal iStart As Integer, _
                 sSeps As String) As Integer
    
    Dim cTarget As Integer
    cTarget = Len(sTarget)
    ' Look for start of token (character that isn't a separator)
    Do While InStr(sSeps, Mid$(sTarget, iStart, 1))
        If iStart > cTarget Then
            StrSpan = 0
            Exit Function
        Else
            iStart = iStart + 1
        End If
    Loop
    StrSpan = iStart

End Function

' This replaces the wretched Split function with its simple-minded design
' and careless implementation. This version allows more than one separator
' and thus can be used for parsing.

' Note that unlike Split, this function has (and needs) no case-sensitivity
' parameter. Split would normally not be used to split on letters, but if
' by some chance you really wanted to do this and wanted the compare to be
' case-insensitive, you could do this:
'     Splits("ThisaisAtest", "Aa")
' Rather than this:
'     Split("ThisaisAtest", "A", ,vbCompareText)
'Function Splits(sExpression As String, _
'                Optional sDelimiters As String = ZCharSpace, _
'                Optional cMax As Long = -1) As Variant
'    Dim sToken As String, avsRet() As Variant, c As Long
'    ' Error trap to resize on overflow
'    On Error GoTo SplitResize
'    ' Break into tokens and put in an array
'    sToken = GetToken(sExpression, sDelimiters)
'    Do While sToken <> ZCharEmpty
'        If cMax <> -1 Then If c >= cMax Then Exit Do
'        avsRet(c) = sToken
'        c = c + 1
'        sToken = GetToken(ZCharEmpty, sDelimiters)
'    Loop
'    ' Size is an estimate, so resize to counted number of tokens
'    If c Then ReDim Preserve avsRet(0 To c - 1) As Variant
'    Splits = avsRet
'    Exit Function
'
'SplitResize:
'    ' Resize on overflow
'    Const cChunk As Long = 20
'    If Err.Number = eeOutOfBounds Then
'        ReDim Preserve avsRet(0 To c + cChunk) As Variant
'        Resume              ' Try again
'    End If
'    ErrRaise Err.Number     ' Other VB error for client
'End Function

' Same as Splits, but handles quoted strings
'Function QSplits(sExpression As String, _
'                 Optional sDelimiters As String = ZCharSpace, _
'                 Optional cMax As Long = -1) As Variant
'    Dim sToken As String, avsRet() As Variant, c As Long
'    ' Error trap to resize on overflow
'    On Error GoTo QSplitResize
'    ' Break into tokens and put in an array
'    sToken = GetQToken(sExpression, sDelimiters)
'    Do While sToken <> ZCharEmpty
'        If cMax <> -1 Then If c >= cMax Then Exit Do
'        avsRet(c) = sToken
'        c = c + 1
'        sToken = GetQToken(ZCharEmpty, sDelimiters)
'    Loop
'    ' Size is an estimate, so resize to counted number of tokens
'    If c Then ReDim Preserve avsRet(0 To c - 1) As Variant
'    QSplits = avsRet
'    Exit Function
'
'QSplitResize:
'    ' Resize on overflow
'    Const cChunk As Long = 20
'    If Err.Number = eeOutOfBounds Then
'        ReDim Preserve avsRet(0 To c + cChunk) As Variant
'        Resume              ' Try again
'    End If
'    ErrRaise Err.Number     ' Other VB error for client
'End Function


'
'#If fComponent = 0 Then
'Private Sub ErrRaise(e As Long)
'    Dim sText As String, sSource As String
'    If e > 1000 Then
'        sSource = App.EXEName & ".Parse"
'        Select Case e
'        Case eeBaseParse
'            BugAssert True
'       ' Case ee...
'       '     Add additional errors
'        End Select
'        Err.Raise COMError(e), sSource, sText
'    Else
'        ' Raise standard Visual Basic error
'        sSource = App.EXEName & ".VBError"
'        Err.Raise e, sSource
'    End If
'End Sub
'#End If

Public Function rdel(ByRef theSTR As String) As String

    Dim a As String
    rdel = theSTR

    If rdel = vbNullString Then Exit Function
    a = Right$(rdel, 1)

    Do Until a <> Chr$(0) And a <> Chr$(32) And a <> Chr$(10) And a <> Chr$(13)
        rdel = Left$(rdel, Len(rdel) - 1)
        a = Right$(rdel, 1)
    Loop

End Function

Public Function ldel(ByRef theSTR As String) As String

    Dim a As String
    ldel = theSTR

    If ldel = vbNullString Then Exit Function
    a = Left$(ldel, 1)

    Do Until a <> Chr$(0) And a <> Chr$(32) And a <> Chr$(10) And a <> Chr$(13)
        ldel = Right$(ldel, Len(ldel) - 1)
        a = Left$(ldel, 1)
    Loop

End Function

Public Function LeftDelete(theSTR As String, sDel As String) As String

    LeftDelete = theSTR

    If LeftDelete = vbNullString Then Exit Function

    Do Until Left$(LeftDelete, Len(sDel)) <> sDel
        LeftDelete = Right$(LeftDelete, Len(LeftDelete) - Len(sDel))
    Loop

End Function

Public Function RightDelete(theSTR As String, sDel As String) As String

    RightDelete = theSTR

    If RightDelete = vbNullString Then Exit Function

    Do Until Right$(RightDelete, Len(sDel)) <> sDel
        RightDelete = Left$(RightDelete, Len(RightDelete) - Len(sDel))
    Loop

End Function

Function StrNum(ByVal Num As Long, ByVal lenNum As Long, Optional vToHex As Boolean = False) As String

    If vToHex Then
        StrNum = LCase$(Hex$(Num))
    Else
        StrNum = LTrim$(Str$(Num))
    End If
    
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

    If strList = vbNullString Then MyInstr = True: Exit Function
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
    EncodeURI = vbNullString
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
            sDecode = vbNullString
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
        RightLeft = IIf(RetError = ReturnOriginalStr, Str, vbNullString)
    Else
        RightLeft = Left$(Str, K - 1)
    End If

End Function

' Search from end to beginning and return the right side of the string
Public Function RightRight(ByRef Str As String, RFind As String, Optional Compare As VbCompareMethod = vbBinaryCompare, Optional RetError As IfStringNotFound = ReturnOriginalStr) As String

    Dim K As Long
    K = InStrRev(Str, RFind, , Compare)

    If K = 0 Then
        RightRight = IIf(RetError = ReturnOriginalStr, Str, vbNullString)
    Else
        RightRight = Mid$(Str, K + 1, Len(Str))
    End If

End Function

' Search from the beginning to end and return the left size of the string
Public Function LeftLeft(ByRef Str As String, LFind As String, Optional Compare As VbCompareMethod = vbBinaryCompare, Optional RetError As IfStringNotFound = ReturnOriginalStr) As String

    Dim K As Long
    K = InStr(1, Str, LFind, Compare)

    If K = 0 Then
        LeftLeft = IIf(RetError = ReturnOriginalStr, Str, vbNullString)
    Else
        LeftLeft = Left$(Str, K - 1)
    End If

End Function

' Search from the beginning to end and return the right size of the string
Public Function LeftRight(ByRef Str As String, LFind As String, Optional Compare As VbCompareMethod = vbBinaryCompare, Optional RetError As IfStringNotFound = ReturnOriginalStr) As String

    Dim K As Long
    K = InStr(1, Str, LFind, Compare)

    If K = 0 Then
        LeftRight = IIf(RetError = ReturnOriginalStr, Str, vbNullString)
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
            LeftRange = IIf(RetError = ReturnOriginalStr, Str, vbNullString)
        End If

    Else
        LeftRange = IIf(RetError = ReturnOriginalStr, Str, vbNullString)
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
            RightRange = IIf(RetError = ReturnOriginalStr, Str, vbNullString)
        End If

    Else
        RightRange = IIf(RetError = ReturnOriginalStr, Str, vbNullString)
    End If

End Function

Public Function EscapeUrl(ByVal sUrl As String) As String

    Dim buff As String
    Dim dwSize As Long
    Dim dwFlags As Long
    EscapeUrl = sUrl
    If Len(sUrl) > 0 Then
        buff = Space$(4 * Len(sUrl))
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
        buff = Space$(Len(sUrl))
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

    If s = vbNullString Then s = "False"
    CBoolStr = CBool(s)

End Function

Public Function CLngStr(s As String) As Long

    If s = vbNullString Then s = "0"
    CLngStr = CLng(s)

End Function

Public Function toUnixPath(sDosPath As String) As String

    toUnixPath = Replace(sDosPath, "\", "/")

End Function

Public Function toDosPath(sUnixPath As String) As String

    toDosPath = Replace(sUnixPath, "/", "\")

End Function

Public Function ToPY(HZCode As Long) As String
'
'    Const WordSep = "|"
'    Const m_lMaxValue As Long = -10254
'    Const m_lMinValue As Long = -20319
'    Static sPYDCT As String
'    Dim sHzCode As String
'    Dim lchkCode As Long
'    Dim lposEnd As Long
'    Dim lposStart As Long
'
'    If HZCode > m_lMaxValue Then Exit Function
'
'    If HZCode < m_lMinValue Then Exit Function
'
'    If sPYDCT = vbNULLSTRING Then
'        Dim NewCAppendString As New CAppendString
'
'        With NewCAppendString
'            .Append "|a=-20319|ai=-20317|an=-20304|ang=-20295|ao=-20292|ba=-20283|bai=-20265|ban=-20257|bang=-20242|bao=-20230|bei=-20051|ben=-20036|beng=-20032|bi=-20026|bian=-20002|biao=-19990|bie=-19986|bin=-19982|bing=-19976|"
'            .Append "bo=-19805|bu=-19784|ca=-19775|cai=-19774|can=-19763|cang=-19756|cao=-19751|ce=-19746|ceng=-19741|cha=-19739|chai=-19728|chan=-19725|chang=-19715|chao=-19540|che=-19531|chen=-19525|cheng=-19515|chi=-19500|chong=-19484|chou=-19479|"
'            .Append "chu=-19467|chuai=-19289|chuan=-19288|chuang=-19281|chui=-19275|chun=-19270|chuo=-19263|ci=-19261|cong=-19249|cou=-19243|cu=-19242|cuan=-19238|cui=-19235|cun=-19227|cuo=-19224|da=-19218|dai=-19212|dan=-19038|dang=-19023|dao=-19018|"
'            .Append "de=-19006|deng=-19003|di=-18996|dian=-18977|diao=-18961|die=-18952|ding=-18783|diu=-18774|dong=-18773|dou=-18763|du=-18756|duan=-18741|dui=-18735|dun=-18731|duo=-18722|e=-18710|en=-18697|er=-18696|fa=-18526|fan=-18518|"
'            .Append "fang=-18501|fei=-18490|fen=-18478|feng=-18463|fo=-18448|fou=-18447|fu=-18446|ga=-18239|gai=-18237|gan=-18231|gang=-18220|gao=-18211|ge=-18201|gei=-18184|gen=-18183|geng=-18181|gong=-18012|gou=-17997|gu=-17988|gua=-17970|"
'            .Append "guai=-17964|guan=-17961|guang=-17950|gui=-17947|gun=-17931|guo=-17928|ha=-17922|hai=-17759|han=-17752|hang=-17733|hao=-17730|he=-17721|hei=-17703|hen=-17701|heng=-17697|hong=-17692|hou=-17683|hu=-17676|hua=-17496|huai=-17487|"
'            .Append "huan=-17482|huang=-17468|hui=-17454|hun=-17433|huo=-17427|ji=-17417|jia=-17202|jian=-17185|jiang=-16983|jiao=-16970|jie=-16942|jin=-16915|jing=-16733|jiong=-16708|jiu=-16706|ju=-16689|juan=-16664|jue=-16657|jun=-16647|ka=-16474|"
'            .Append "kai=-16470|kan=-16465|kang=-16459|kao=-16452|ke=-16448|ken=-16433|keng=-16429|kong=-16427|kou=-16423|ku=-16419|kua=-16412|kuai=-16407|kuan=-16403|kuang=-16401|kui=-16393|kun=-16220|kuo=-16216|la=-16212|lai=-16205|lan=-16202|"
'            .Append "lang=-16187|lao=-16180|le=-16171|lei=-16169|leng=-16158|li=-16155|lia=-15959|lian=-15958|liang=-15944|liao=-15933|lie=-15920|lin=-15915|ling=-15903|liu=-15889|long=-15878|lou=-15707|lu=-15701|lv=-15681|luan=-15667|lue=-15661|"
'            .Append "lun=-15659|luo=-15652|ma=-15640|mai=-15631|man=-15625|mang=-15454|mao=-15448|me=-15436|mei=-15435|men=-15419|meng=-15416|mi=-15408|mian=-15394|miao=-15385|mie=-15377|min=-15375|ming=-15369|miu=-15363|mo=-15362|mou=-15183|"
'            .Append "mu=-15180|na=-15165|nai=-15158|nan=-15153|nang=-15150|nao=-15149|ne=-15144|nei=-15143|nen=-15141|neng=-15140|ni=-15139|nian=-15128|niang=-15121|niao=-15119|nie=-15117|nin=-15110|ning=-15109|niu=-14941|nong=-14937|nu=-14933|"
'            .Append "nv=-14930|nuan=-14929|nue=-14928|nuo=-14926|o=-14922|ou=-14921|pa=-14914|pai=-14908|pan=-14902|pang=-14894|pao=-14889|pei=-14882|pen=-14873|peng=-14871|pi=-14857|pian=-14678|piao=-14674|pie=-14670|pin=-14668|ping=-14663|"
'            .Append "po=-14654|pu=-14645|qi=-14630|qia=-14594|qian=-14429|qiang=-14407|qiao=-14399|qie=-14384|qin=-14379|qing=-14368|qiong=-14355|qiu=-14353|qu=-14345|quan=-14170|que=-14159|qun=-14151|ran=-14149|rang=-14145|rao=-14140|re=-14137|"
'            .Append "ren=-14135|reng=-14125|ri=-14123|rong=-14122|rou=-14112|ru=-14109|ruan=-14099|rui=-14097|run=-14094|ruo=-14092|sa=-14090|sai=-14087|san=-14083|sang=-13917|sao=-13914|se=-13910|sen=-13907|seng=-13906|sha=-13905|shai=-13896|"
'            .Append "shan=-13894|shang=-13878|shao=-13870|she=-13859|shen=-13847|sheng=-13831|shi=-13658|shou=-13611|shu=-13601|shua=-13406|shuai=-13404|shuan=-13400|shuang=-13398|shui=-13395|shun=-13391|shuo=-13387|si=-13383|song=-13367|sou=-13359|su=-13356|"
'            .Append "suan=-13343|sui=-13340|sun=-13329|suo=-13326|ta=-13318|tai=-13147|tan=-13138|tang=-13120|tao=-13107|te=-13096|teng=-13095|ti=-13091|tian=-13076|tiao=-13068|tie=-13063|ting=-13060|tong=-12888|tou=-12875|tu=-12871|tuan=-12860|"
'            .Append "tui=-12858|tun=-12852|tuo=-12849|wa=-12838|wai=-12831|wan=-12829|wang=-12812|wei=-12802|wen=-12607|weng=-12597|wo=-12594|wu=-12585|xi=-12556|xia=-12359|xian=-12346|xiang=-12320|xiao=-12300|xie=-12120|xin=-12099|xing=-12089|"
'            .Append "xiong=-12074|xiu=-12067|xu=-12058|xuan=-12039|xue=-11867|xun=-11861|ya=-11847|yan=-11831|yang=-11798|yao=-11781|ye=-11604|yi=-11589|yin=-11536|ying=-11358|yo=-11340|yong=-11339|you=-11324|yu=-11303|yuan=-11097|yue=-11077|"
'            .Append "yun=-11067|za=-11055|zai=-11052|zan=-11045|zang=-11041|zao=-11038|ze=-11024|zei=-11020|zen=-11019|zeng=-11018|zha=-11014|zhai=-10838|zhan=-10832|zhang=-10815|zhao=-10800|zhe=-10790|zhen=-10780|zheng=-10764|zhi=-10587|zhong=-10544|"
'            .Append "zhou=-10533|zhu=-10519|zhua=-10331|zhuai=-10329|zhuan=-10328|zhuang=-10322|zhui=-10315|zhun=-10309|zhuo=-10307|zi=-10296|zong=-10281|zou=-10274|zu=-10270|zuan=-10262|zui=-10260|zun=-10256|zuo=-10254|"
'        End With
'
'        sPYDCT = NewCAppendString.Value
'        Set NewCAppendString = Nothing
'    End If
'
'    On Error GoTo Herr
'
'    For lchkCode = HZCode To m_lMinValue Step -1
'        sHzCode = "=" & LTrim$(Str$(lchkCode))
'        lposEnd = InStr(sPYDCT, sHzCode)
'
'        If lposEnd > 0 Then
'            lposStart = InStrRev(sPYDCT, WordSep, lposEnd)
'
'            If lposStart > 0 Then
'                ToPY = Mid(sPYDCT, lposStart + 1, lposEnd - lposStart - 1)
'                Exit Function
'            End If
'
'        End If
'
'    Next
'
'Herr:

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
        CBytesToStr = vbNullString
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
    Dim C As String
    isTypicalFileName = False
    If strTest = vbNullString Then Exit Function
    strTest = Replace$(strTest, "/", "\")
    l = Len(strTest)
    For pos = 1 To l
        C = Mid$(strTest, pos, 1)
        If C = "|" Then Exit Function
        If C = ">" Then Exit Function
        If C = "<" Then Exit Function
        If C = "*" Then Exit Function
        If C = Chr$(34) Then Exit Function
        If C = "?" Then Exit Function
    Next
    pos = InStr(strTest, ".")
    l = InStr(strTest, "\")
    If pos > l Then isTypicalFileName = True
End Function

Public Function startAt(ByRef strTest As String, ByRef strStart As String, Optional ByRef cmp As VbCompareMethod = vbTextCompare) As Boolean
    Dim lstrTest As Long
    Dim lstrStart As Long
    
    startAt = False
    lstrTest = Len(strTest)
    lstrStart = Len(strStart)
    
    If lstrTest < lstrStart Then Exit Function
    If StrComp(Left$(strTest, lstrStart), strStart, cmp) <> 0 Then Exit Function
    
    startAt = True

End Function

Public Function StringToLong(ByRef vString As String) As Long
    On Error Resume Next
    StringToLong = CLng(vString)
    Err.Clear
End Function

Public Function StringToInteger(ByRef vString As String) As Integer
    On Error Resume Next
    StringToInteger = CInt(vString)
    Err.Clear
End Function

Public Function SubStringUntilMatch(ByRef Source As String, ByRef Start As Long, ByRef Target As String) As String
    On Error GoTo NoMatch
    Dim iEnd As Long
    If Start < 1 Then Start = 1
    iEnd = InStr(Start, Source, Target, vbTextCompare)
    If (iEnd > Start) Then
        SubStringUntilMatch = Mid$(Source, Start, iEnd - Start)
    End If
    Exit Function
    
NoMatch:
    
End Function

Public Function SubStringBetween(ByRef Source As String, ByRef vLeft As String, Optional ByRef vRight As String, Optional vTrimResult As Boolean = False) As String
    On Error GoTo NoMatch
    Dim pStart As Long
    pStart = InStr(1, Source, vLeft, vbTextCompare)
    
    If (pStart > 0) Then
        If vRight = vbNullString Then
            SubStringBetween = Mid$(Source, pStart + 1)
        Else
        SubStringBetween = SubStringUntilMatch(Source, pStart + Len(vLeft), vRight)
        End If
        If vTrimResult Then SubStringBetween = Trim$(SubStringBetween)
    End If
    Exit Function
NoMatch:
End Function


Public Function QuoteString(ByRef vString As String) As String
    QuoteString = ZCharDoubleQuote & vString & ZCharDoubleQuote
End Function

'Public Function XXor(Optional ByVal vStart As Long = 1) As String
'    Dim fNum As Integer
'    fNum = FreeFile
'    Open App.Path & "\ss2path.dat" For Binary Access Read As #fNum
'    Seek #fNum, vStart
'
'    Dim vData(0 To 7) As Byte
'    Get #fNum, , vData()
'    Dim xByte() As Byte
''    vData(0) = &H72
''    vData(1) = &HF5
''    vData(2) = &HAD
''    vData(3) = &HE7
''    vData(4) = &H3F
''    vData(5) = &HD1
''    vData(6) = &HD0
''    vData(7) = &HBA
'Close #fNum
'    xByte() = StrConv("3.1415926", vbFromUnicode)
'    Dim i As Long
'    For i = 0 To 7
'        vData(i) = vData(i) Xor xByte(i)
'    Next
'    XXor = vData
'End Function

Public Function BuildRangeString(vPre As String, vNumLength As Long, vSuf As String, Optional vStart As Long = 1) As String
    Dim s As String
    Dim l As Long
    
    s = CStr(vStart)
    l = vNumLength - Len(s)
    If (l > 0) Then
        BuildRangeString = vPre & String$(l, "0") & s & vSuf
    Else
        BuildRangeString = vPre & s & vSuf
    End If

End Function

'Public Function StringFromCPtr(ByVal vPtr As Long) As String
'    Dim pLen As Long
'    pLen = lstrlenPtr(vPtr)
'    StringFromCPtr = String$(pLen, " ")
'    lstrcpyStrPtr StringFromCPtr, vPtr
'End Function

