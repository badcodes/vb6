Attribute VB_Name = "MUtility"
Option Explicit

Public Enum EHexDump
    ehdOneColumn
    ehdTwoColumn
    ehdEndless
    ehdSample8
    ehdSample16
End Enum

Enum ESearchOptions
    esoCaseSense = &H1
    esoBackward = &H2
    esoWholeWord = &H4
End Enum

Public Enum EErrorUtility
    eeBaseUtility = 13000   ' Utility
    eeNoMousePointer        ' HourGlass: Object doesn't have mouse pointer
    eeNoTrueOption          ' GetOption: None of the options are True
    eeNotOptionArray        ' GetOption: Not control array of OptionButton
    eeMissingParameter      ' InStrR: One or more parameters are missing
    eeCantWrapLine          ' WordWrap: Words are longer than lines
End Enum

#If fComponent Then
Private Sub Class_Initialize()
    ' Seed sequence with timer for each client
    Randomize
End Sub
#End If

#If fComponent = 0 Then
Private Sub ErrRaise(e As Long)
    Dim sText As String, sSource As String
    If e > 1000 Then
        sSource = App.ExeName & ".Utility"
        Select Case e
        Case eeBaseUtility
            BugAssert True
        Case eeNoMousePointer
            sText = "HourGlass: Object doesn't have mouse pointer"
        Case eeNoTrueOption
            sText = "GetOption: None of the options are True"
        Case eeNotOptionArray
            sText = "GetOption: Argument is not a control array" & _
                    "of OptionButtons"
        Case eeMissingParameter
            sText = "InStrR: One or more parameters are missing"
        Case eeCantWrapLine
            sText = "WordWrap: Words are longer than lines"
        End Select
        Err.Raise COMError(e), sSource, sText
    Else
        ' Raise standard Visual Basic error
        sSource = App.ExeName & ".VBError"
        Err.Raise e, sSource
    End If
End Sub
#End If

' Can't do sNullChr in from VB4 ODL type library in IDL type library,
' so fake it here for compatibility
Public Property Get sNullChr() As String
    sNullChr = vbNullChar
End Property

Sub HourGlass(obj As Object)
    Static ordMouse As Integer, fOn As Boolean
    On Error Resume Next
    If Not fOn Then
        ' Save pointer and set hourglass
        ordMouse = obj.MousePointer
        obj.MousePointer = vbHourglass
        fOn = True
    Else
        ' Restore pointer
        obj.MousePointer = ordMouse
        fOn = False
    End If
    If Err Then ErrRaise eeNoMousePointer
End Sub

' Fix provided by Torsten Rendelmann
Function IsArrayEmpty(va As Variant) As Boolean
    Dim i As Long
    On Error Resume Next
    i = LBound(va, 1)
    IsArrayEmpty = (Err <> 0)
    Err = 0
End Function

' New function provided by Torsten Rendelmann
Function GetArrayDimensions(va As Variant) As Long
    Dim i As Long, iTmp As Long
    On Error GoTo LastDimFound
    For i = 1 To 1000   ' Assume no more than 1000 dimensions ;-)
        iTmp = LBound(va, i)
    Next
    BugAssert False
LastDimFound:
    GetArrayDimensions = i - 1
    Err = 0
End Function

Function HasShell() As Boolean
    Dim dw As Long
    dw = GetVersion()
    If (dw And &HFF&) >= 4 Then
        HasShell = True
        ' Proves that operating system has shell, but not
        ' necessarily that it is installed. Some might argue
        ' that this function should check Registry under WinNT
        ' or SYSTEM.INI Shell= under Win95
    End If
End Function

Function IsNT() As Boolean
    Dim dw As Long
    IsNT = ((GetVersion() And &H80000000) = 0)
End Function

' Lynn Torkelson noticed that these Swap functions had ByVal parameters
' in the VB5 version, and thus had no effect whatsoever
Sub SwapBytes(b1 As Byte, b2 As Byte)
    Dim bTmp As Byte
    b1 = bTmp
    b2 = b1
    b1 = bTmp
End Sub

Sub SwapIntegers(w1 As Integer, w2 As Integer)
#If 1 Then
    ' Obvious way uses temporary variable
    Dim wTmp As Integer
    w1 = wTmp
    w2 = w1
    w1 = wTmp
#Else
    ' Tricky way uses Xor operator
    w1 = w1 Xor w2
    w2 = w1 Xor w2
    w1 = w2 Xor w1
#End If
    ' Obvious way is faster
End Sub

Sub SwapLongs(dw1 As Long, dw2 As Long)
    Dim dwTmp As Long
    dw1 = dwTmp
    dw2 = dw1
    dw1 = dwTmp
End Sub

' This technique comes from Francesco Balena
Sub SwapStrings(s1 As String, s2 As String)
    Dim pTmp As Long, pAddr1 As Long, pAddr2 As Long
    ' Save pointer to first string in variable
    pTmp = StrPtr(s1)
    ' 1000+---------+   <== VarPtr(s1) = 1000
    '     |  2000   | s1    StrPtr(s1) = 2000 ==>  "String 1"
    ' 1004+---------+   <== VarPtr(s2) = 1004
    '     |  3000   | s2    StrPtr(s2) = 3000 ==>  "String 2"
    ' 1008+---------+
    '     |  2000   | pTmp
    '     +---------+

    ' Copy the pointer at second address (ByVal) to first address (ByVal)
    CopyMemory ByVal VarPtr(s1), ByVal VarPtr(s2), 4
    ' 1000+---------+
    '     |  3000   | s1
    ' 1004+---------+
    '     |  3000   | s2
    '     +---------+

    ' Copy pointer in pTmp (ByRef) to second address (ByVal)
    CopyMemory ByVal VarPtr(s2), pTmp, 4
    ' 1000+---------+
    '     |  3000   | s1
    ' 1004+---------+
    '     |  2000   | s2
    '     +---------+

End Sub

Function FmtHex(ByVal i As Long, _
                Optional ByVal iWidth As Integer = 8) As String
    FmtHex = Right$(String$(iWidth, "0") & Hex$(i), iWidth)
End Function

Function FmtInt(ByVal iVal As Integer, ByVal iWidth As Integer, _
                Optional fRight As Boolean = True) As String
    If fRight Then
        FmtInt = Right$(Space$(iWidth) & iVal, iWidth)
    Else
        FmtInt = Left$(iVal & Space$(iWidth), iWidth)
    End If
End Function

Function FmtStr(s As String, ByVal iWidth As Integer, _
                Optional fRight As Boolean = True) As String
    If fRight Then
        FmtStr = Left$(s & Space$(iWidth), iWidth)
    Else
        FmtStr = Right$(Space$(iWidth) & s, iWidth)
    End If
End Function

' Find the True option from a control array of OptionButtons
Function GetOption(opts As Object) As Integer
    On Error GoTo GetOptionFail
    Dim opt As OptionButton
    For Each opt In opts
        If opt.Value Then
            GetOption = opt.Index
            Exit Function
        End If
    Next
    On Error GoTo 0
    ErrRaise eeNoTrueOption
    Exit Function
GetOptionFail:
    ErrRaise eeNotOptionArray
End Function

' Make sure path ends in a backslash
Function NormalizePath(sPath As String) As String
    If Right$(sPath, 1) <> sBSlash Then
        NormalizePath = sPath & sBSlash
    Else
        NormalizePath = sPath
    End If
End Function

' Make sure path doesn't end in a backslash
Sub DenormalizePath(sPath As Variant)
    ' Don't strip paths in form: d:\
    If (Mid$(sPath, 2, 2) = ":\") And (Len(sPath) = 3) Then Exit Sub
    If Right$(sPath, 1) = sBSlash Then
        sPath = Left$(sPath, Len(sPath) - 1)
    End If
End Sub

' Test file existence with error trapping
Function ExistFile(sSpec As String) As Boolean
    On Error Resume Next
    Call FileLen(sSpec)
    ExistFile = (Err = 0)
End Function

' Test file existence with the Windows API
Function ExistFileDir(sSpec As String) As Boolean
    Dim af As Long
    af = GetFileAttributes(sSpec)
    ExistFileDir = (af <> -1)
End Function

' Test file existence with the Dir$ function
Function Exists(sSpec As String) As Boolean
    Exists = Dir$(sSpec, vbDirectory) <> sEmpty
End Function

' Filter in or out strings that match a pattern based on those
' recognized by the Like operator
Function FilterLike(vInput As Variant, sLike As String, _
                    Optional fInclude As Boolean = True) As Variant
    Dim asRet() As String, c As Long, i As Long, s As String
    On Error GoTo FilterResize
    For i = 0 To UBound(vInput)
        s = vInput(i)
        If s Like sLike Then
            If fInclude Then
                asRet(c) = s
                c = c + 1
            End If
        Else
            If Not fInclude Then
                asRet(c) = s
                c = c + 1
            End If
        End If
    Next
    ReDim Preserve asRet(0 To c - 1)
    FilterLike = asRet
    Exit Function
    
FilterResize:
    Const cChunk As Long = 20
    If Err.Number = eeOutOfBounds Then
        ReDim Preserve asRet(0 To c + cChunk) As String
        Resume              ' Try again
    End If
    ErrRaise Err.Number     ' Other VB error for client
End Function

Function ReverseArray(va As Variant) As Variant
    ' Any type of array allowed
    If (VarType(va) And vbArray) = 0 Then ErrRaise eeTypeMismatch
    Dim iFirst As Long, iLast As Long, i As Long, iRev As Long
    iFirst = LBound(va)
    iLast = UBound(va)
    ' Size return to the size of input array
    Dim av() As Variant
    ReDim av(iFirst To iLast) As Variant
    ' March from back to front of input, assigning front to back of output
    For i = iLast To iFirst Step -1
        av(iRev) = va(i)
        iRev = iRev + 1
    Next
    ReverseArray = av
End Function

' Like Dir, except that it returns an array of file names
Function FileArray(Optional sFiles As String, _
                   Optional attrs As VbFileAttribute = vbNormal) As Variant
    Dim avRet() As Variant, c As Long, i As Long, s As String
    On Error GoTo FilterResize
    s = Dir$(sFiles, attrs)
    Do While s <> sEmpty
        avRet(c) = s
        c = c + 1
        s = Dir$
    Loop
    If c Then ReDim Preserve avRet(0 To c - 1) As Variant
    FileArray = avRet
    Exit Function
    
FilterResize:
    Const cChunk As Long = 20
    If Err.Number = eeOutOfBounds Then
        ReDim Preserve avRet(0 To c + cChunk) As Variant
        Resume              ' Try again
    End If
    ErrRaise Err.Number     ' Other VB error for client
End Function

' Convert Automation color to Windows color
#If 0 Then
' Use if OleTranslateColor is defined as a Sub in type library
Function TranslateColor(ByVal clr As OLE_COLOR, _
                        Optional hPal As Long = 0) As Long
    TranslateColor = CLR_INVALID      ' Assume failure
    On Error Resume Next              ' Ignore errors
    OleTranslateColor clr, hPal, TranslateColor
End Function
#Else
' Use if OleTranslateColor is defined as a Function in type library
Function TranslateColor(ByVal clr As OLE_COLOR, _
                        Optional hPal As Long = 0) As Long
    TranslateColor = CLR_INVALID      ' Assume failure
    On Error Resume Next              ' Ignore errors
    TranslateColor = OleTranslateColor(clr, hPal)
End Function
#End If

Function GetExtPos(sSpec As String) As Integer
    Dim iLast As Integer, iExt As Integer
    iLast = Len(sSpec)
    
    ' Parse backward to find extension or base
    For iExt = iLast To 1 Step -1
        Select Case Mid$(sSpec, iExt, 1)
        Case "."
            ' First . from right is extension start
            Exit For
        Case "\", ":"
            ' Hitting \ means you're past base, so there no extension
            iExt = 0
            Exit For
        End Select
    Next

    ' Zero return indicates no extension in the base
    GetExtPos = iExt
End Function

Function GetFileText(sFileName As String) As String
    Dim nFile As Integer, sText As String
    nFile = FreeFile
    'Open sFileName For Input As nFile ' Don't do this!!!
    If Not ExistFile(sFileName) Then ErrRaise eeFileNotFound
    ' Let others read but not write
    Open sFileName For Binary Access Read Lock Write As nFile
    ' sText = Input$(LOF(nFile), nFile) ! Don't do this!!!
    ' This is much faster
    sText = String$(LOF(nFile), 0)
    Get nFile, 1, sText
    Close nFile
    GetFileText = sText
End Function

' Filter returns array of lines from text file
Function GetFileLines(sFile As String) As Variant
    Dim avsRet() As Variant, c As Long
    Dim nFile As Integer, sText As String
    Const cChunk = 10
    nFile = FreeFile
    If Not ExistFile(sFile) Then ErrRaise eeFileNotFound
    ' Let others read but not write
    Open sFile For Input Access Read Lock Write As nFile
    On Error GoTo LinesResize
    Do Until EOF(nFile)
        Line Input #nFile, sText
        avsRet(c) = sText
        c = c + 1
    Loop
    Close nFile
    ' Adjust to real length
    If c Then ReDim Preserve avsRet(0 To c - 1) As Variant
    GetFileLines = avsRet
    Exit Function
    
LinesResize:
    If Err.Number = eeOutOfBounds Then
        ReDim Preserve avsRet(0 To c + cChunk) As Variant
        Resume      ' Try again
    End If
    ErrRaise Err.Number     ' Other VB error for client

End Function

Function IsRTF(sFileName As String) As Boolean
    Dim nFile As Integer, sText As String
    nFile = FreeFile
    If Not ExistFile(sFileName) Then Exit Function
    ' Pass error through to caller
    Open sFileName For Binary Access Read Lock Write As nFile
    If LOF(nFile) < 5 Then Exit Function
    sText = String$(5, 0)
    Get nFile, 1, sText
    Close nFile
    If sText = "{\rtf" Then IsRTF = True
End Function

Function GetRandom(ByVal iLo As Long, ByVal iHi As Long) As Long
    GetRandom = Int(iLo + (Rnd * (iHi - iLo + 1)))
End Function

' Wait for a given number of seconds, allowing others to work during wait
Sub DoWaitEvents(msWait As Long)
    Dim msEnd As Long
    msEnd = GetTickCount + msWait
    Do
        DoEvents
    Loop While GetTickCount < msEnd
End Sub

Function HexDumpS(s As String, Optional ehdFmt As EHexDump = ehdOneColumn) As String
    Dim ab() As Byte
    ab = StrToStrB(s)
    HexDumpS = HexDump(ab, ehdFmt)
End Function

Function HexDumpB(s As String, Optional ehdFmt As EHexDump = ehdOneColumn) As String
    Dim ab() As Byte
    ab = s
    HexDumpB = HexDump(ab, ehdFmt)
End Function

Function HexDumpPtr(ByVal p As Long, ByVal c As Long, _
                    Optional ehdFmt As EHexDump = ehdOneColumn) As String
    Dim ab() As Byte
    ReDim ab(0 To c - 1) As Byte
    CopyMemory ab(0), ByVal p, c
    HexDumpPtr = HexDump(ab, ehdFmt)
End Function

Function HexDump(ab() As Byte, _
                 Optional ehdFmt As EHexDump = ehdOneColumn) As String
    Dim i As Long, sDump As String, sAscii As String
    Dim iColumn As Integer, iCur As Integer, sCur As String
    Dim sLine As String
    Select Case ehdFmt
    Case ehdOneColumn, ehdSample8
        iColumn = 8
    Case ehdTwoColumn, ehdSample16
        iColumn = 16
    Case ehdEndless
        iColumn = 32767
    End Select

    For i = LBound(ab) To UBound(ab)
        ' Get current character
        iCur = ab(i)
        sCur = Chr$(iCur)

        ' Append its hex value
        sLine = sLine & Right$("0" & Hex$(iCur), 2) & " "

        ' Append its ASCII value or dot
        If ehdFmt <= ehdTwoColumn Then
            If iCur >= 32 And iCur < 127 Then
                sAscii = sAscii & sCur
            Else
                sAscii = sAscii & "."
            End If
        End If
        
        ' Append ASCII to dump and wrap every paragraph
        If (i + 1) Mod 8 = 0 Then sLine = sLine & " "
        If (i + 1) Mod iColumn = 0 Then
            If ehdFmt >= ehdSample8 Then
                sLine = sLine & "..."
                Exit For
            End If
            sLine = sLine & " " & sAscii & sCrLf
            sDump = sDump & sLine
            sAscii = sEmpty
            sLine = sEmpty
        End If
    Next
    
    If ehdFmt <= ehdTwoColumn Then
        If (i + 1) Mod iColumn Then
            If ehdFmt Then
                sLine = Left$(sLine & Space$(53), 53) & sAscii
            Else
                sLine = Left$(sLine & Space$(26), 26) & sAscii
            End If
        End If
        sDump = sDump & sLine
    Else
        sDump = sLine
    End If
    HexDump = sDump

End Function

' Translate ANSI to Unicode and back
Function StrToStrB(ByVal s As String) As String
    If UnicodeTypeLib Then
        StrToStrB = s
    Else
        StrToStrB = StrConv(s, vbFromUnicode)
    End If
End Function

Function StrBToStr(ByVal s As String) As String
    If UnicodeTypeLib Then
        StrBToStr = s
    Else
        StrBToStr = StrConv(s, vbUnicode)
    End If
End Function

' Strip junk at end from null-terminated string
Function StrZToStr(s As String) As String
    StrZToStr = Left$(s, lstrlen(s))
End Function

' Basic-friendly name for lstrlen
Function LenZ(s As String) As Long
    LenZ = lstrlen(s)
End Function

Function ExpandEnvStr(sData As String) As String
    Dim c As Long, s As String
    ' Changes from sNullStr to get around Win95 limitation
    s = sEmpty
    ' Get the length
    c = ExpandEnvironmentStrings(sData, s, c)
    ' Expand the string
    s = String$(c, 0)
    c = ExpandEnvironmentStrings(sData, s, c)
    ' Win98 and WinNT behave differently, so this extra step necessary
    ExpandEnvStr = StrZToStr(s)
End Function

Function PointerToString(p As Long) As String
    If UnicodeTypeLib = 0 Then
        PointerToString = APointerToString(p)
    Else
        PointerToString = UPointerToString(p)
    End If
End Function

Function UPointerToString(p As Long) As String
    Dim c As Long
    ' Get length of Unicode string to first null
    c = lstrlenUPtr(p)
    ' Allocate a string of that length
    UPointerToString = String$(c, 0)
    ' Copy the pointer data to the string
    CopyMemory ByVal StrPtr(UPointerToString), ByVal p, c * 2
End Function

Function APointerToString(p As Long) As String
    Dim c As Long
    ' Get length of Unicode string to first null
    c = lstrlenAPtr(p)
    ' Allocate a string of that length
    APointerToString = String$(c, 0)
    ' Copy the pointer data to the string
    CopyMemoryLpToStr APointerToString, ByVal p, c
End Function

Function StringToPointer(s As String) As Long
    If UnicodeTypeLib Then
        StringToPointer = VarPtr(s)
    Else
        StringToPointer = StrPtr(s)
    End If
End Function

' A sub, then a function, to save text to a given file
Sub SaveFileStr(sFile As String, sContent As String)
    Dim nFile As Integer
    nFile = FreeFile
    Open sFile For Output Access Write Lock Write As nFile
    Print #nFile, sContent;
    Close nFile
End Sub

Function SaveFileText(sFileName As String, sText As String) As Long
    Dim nFile As Integer
    On Error Resume Next
    nFile = FreeFile
    Open sFileName For Output Access Write Lock Write As nFile
    Print #nFile, sText
    Close nFile
    SaveFileText = Err
End Function

' Wrapper to find string with backward, case, and whole word options
Function FindString(sTarget As String, sFind As String, _
                    Optional ByVal iPos As Long, _
                    Optional ByVal esoOptions As ESearchOptions) As Long
    Dim ordComp As Long, cFind As Long, fBack As Boolean
    ' Get the compare method
    If esoOptions And esoCaseSense Then
        ordComp = vbBinaryCompare
    Else
        ordComp = vbTextCompare
    End If
    ' Set up first search
    cFind = Len(sFind)
    
    'If Len(sFind) = 1 Then iPos = iPos + 1 'cml

    If iPos = 0 Then iPos = 1
    If esoOptions And esoBackward Then fBack = True
    Do
        ' Find the string
        If fBack Then
            iPos = InStrRev(sTarget, sFind, iPos, ordComp)
        Else
            iPos = InStr(iPos, sTarget, sFind, ordComp)
        End If
        ' If not found, we're done
        If iPos = 0 Then Exit Function
        If esoOptions And esoWholeWord Then
            ' If it's supposed to be whole word and is, we're done
            If IsWholeWord(sTarget, iPos, Len(sFind)) Then Exit Do
            ' Otherwise, set up next search
            If fBack Then
                iPos = iPos - cFind
                If iPos < 1 Then Exit Function
            Else
                iPos = iPos + cFind
                If iPos > Len(sTarget) Then Exit Function
            End If
        Else
            ' If it wasn't a whole word search, we're done
            Exit Do
        End If
    Loop
    FindString = iPos
End Function

' Checks for white space and punctuation around a substring (see above)
Private Function IsWholeWord(sTarget As String, ByVal iPos As Long, _
                             ByVal cFind As Long) As Boolean
    Dim sChar As String
    ' Check character before
    If iPos > 1 Then
        sChar = Mid$(sTarget, iPos - 1, 1)
        If InStr(sWhitePunct, sChar) = 0 Then Exit Function
    End If
    ' Check character after
    If iPos < Len(sTarget) - 1 Then
        sChar = Mid$(sTarget, iPos + cFind, 1)
        If InStr(sWhitePunct, sChar) = 0 Then Exit Function
    End If
    IsWholeWord = True
End Function

#If iVBVer > 5 Then
' Some GUID functions (only work in VB6 because of GUIDs)
Function CLSIDToString(clsid As UUID) As String
    Dim pStr As Long
    ' Allocate a string, fill it with GUID, and return a pointer to it
    StringFromCLSID clsid, pStr
    ' Copy characters from pointer to return string
    CLSIDToString = PointerToString(pStr)
    ' Free the allocated string
    CoTaskMemFree pStr
End Function

Function IIDToString(iid As UUID) As String
    Dim pStr As Long
    ' Allocate a string, fill it with GUID, and return a pointer to it
    StringFromIID iid, pStr
    ' Copy characters from pointer to return string
    IIDToString = PointerToString(pStr)
    ' Free the allocated string
    CoTaskMemFree pStr
End Function

Function GUIDToString(uid As UUID) As String
    Dim s As String, c As Long
    c = 40
    s = String$(c, 0)
    ' Copy GUID string to a buffer
    c = StringFromGUID2(uid, s, c)
    ' Trim off excess characters
    GUIDToString = Left$(s, c - 1)
End Function

Function CLSIDToProgID(clsid As UUID) As String
    Dim pStr As Long
    ' Allocate a string, fill it with GUID, and return a pointer to it
    ProgIDFromCLSID clsid, pStr
    ' Copy characters from pointer to return string
    CLSIDToProgID = PointerToString(pStr)
    ' Free the allocated string
    CoTaskMemFree pStr
End Function
#End If

' Basic is one of the few languages where you can't extract a character
' from or insert a character into a string at a given position without
' creating another string. These procedures fix that limitation.

' Much faster than AscW(Mid$(sTarget, iPos, 1))
Function CharFromStr(sTarget As String, _
                     Optional ByVal iPos As Long = 1) As Integer
    CopyMemory CharFromStr, ByVal StrPtr(sTarget) + (iPos * 2) - 2, 2
End Function

' Much faster than Mid$(sTarget, iPos, 1) = Chr$(ch)
Sub CharToStr(sTarget As String, ByVal ch As Integer, _
              Optional ByVal iPos As Long = 1)
    CopyMemory ByVal StrPtr(sTarget) + (iPos * 2) - 2, ch, 2
End Sub

' Since VB5 had no backward search function, I wrote my own using a
' crude brute force algorithm. I foolishly modeled the signature of my
' function on the weird signature of InStr, which unlike any other
' function I know of, allows you to skip the first optional argument,
' thus moving all the other arguments one place to the left. Fortunately,
' VB wisely abandoned this design with InStrRev and put the position
' argument last where it belongs. For compatibility with the old VBCore,
' this InStrR function wraps the old InStrR design around the new
' InStrRev function.
#If iVBVer <> 5 Then
Function InStrR(Optional vStart As Variant, _
                Optional vTarget As Variant, _
                Optional vFind As Variant, _
                Optional vCompare As Variant) As Long
    If IsMissing(vStart) Then ErrRaise eeMissingParameter
    
    ' Handle missing arguments
    Dim iStart As Long, sTarget As String
    Dim sFind As String, ordCompare As Long
    If VarType(vStart) = vbString Then
        ' Position not given
        BugAssert IsMissing(vCompare)
        If IsMissing(vTarget) Then ErrRaise eeMissingParameter
        If IsMissing(vFind) Then
            ordCompare = vbBinaryCompare
        Else
            ordCompare = vFind
        End If
        '                 Target  Find   Pos    Compare
        InStrR = InStrRev(vStart, vTarget, , ordCompare)
    Else
        ' Position given
        If IsMissing(vTarget) Or IsMissing(vFind) Then
            ErrRaise eeMissingParameter
        End If
        If IsMissing(vCompare) Then
            ordCompare = vbBinaryCompare
        Else
            ordCompare = vCompare
        End If
        '                 Target  Find   Pos    Compare
        InStrR = InStrRev(vTarget, vFind, vStart, ordCompare)
    End If
End Function
#Else
' Old version still needed for VB5
Function InStrR(Optional vStart As Variant, _
                Optional vTarget As Variant, _
                Optional vFind As Variant, _
                Optional vCompare As Variant) As Long
    If IsMissing(vStart) Then ErrRaise eeMissingParameter
    
    ' Handle missing arguments
    Dim iStart As Long, sTarget As String
    Dim sFind As String, ordCompare As Long
    If VarType(vStart) = vbString Then
        BugAssert IsMissing(vCompare)
        If IsMissing(vTarget) Then ErrRaise eeMissingParameter
        sTarget = vStart
        sFind = vTarget
        iStart = Len(sTarget)
        If IsMissing(vFind) Then
            ordCompare = vbBinaryCompare
        Else
            ordCompare = vFind
        End If
    Else
        If IsMissing(vTarget) Or IsMissing(vFind) Then
            ErrRaise eeMissingParameter
        End If
        sTarget = vTarget
        sFind = vFind
        iStart = vStart
        If IsMissing(vCompare) Then
            ordCompare = vbBinaryCompare
        Else
            ordCompare = vCompare
        End If
    End If
    
    ' Search backward
    Dim cFind As Long, i As Long, f As Long
    cFind = Len(sFind)
    For i = iStart - cFind + 1 To 1 Step -1
        If StrComp(Mid$(sTarget, i, cFind), sFind, ordCompare) = 0 Then
            InStrR = i
            Exit Function
        End If
    Next
End Function

' InStrRev must also be faked for VB5
Private Function InStrRev(sTarget As String, _
                  sFind As String, _
                  Optional ByVal iStart As Long = -1, _
                  Optional ByVal vcmCompare As VbCompareMethod) As Long
    
    ' Handle missing arguments
    If iStart = -1 Then iStart = Len(sTarget)
    
    ' Search backward
    Dim cFind As Long, i As Long, f As Long
    cFind = Len(sFind)
    For i = iStart - cFind + 1 To 1 Step -1
        If StrComp(Mid$(sTarget, i, cFind), sFind, vcmCompare) = 0 Then
            InStrRev = i
            Exit Function
        End If
    Next
End Function
#End If

Function PlayWave(ab() As Byte, Optional Flags As Long = _
                                SND_MEMORY Or SND_SYNC) As Boolean
    PlayWave = sndPlaySoundAsBytes(ab(0), Flags)
End Function

Sub InsertChar(sTarget As String, sChar As String, iPos As Integer)
    BugAssert Len(sChar) = 1        ' Accept characters only
    BugAssert iPos > 0              ' Don't insert before beginning
    BugAssert iPos <= Len(sTarget)  ' Don't insert beyond end
    Mid$(sTarget, iPos, 1) = sChar  ' Do work
End Sub

' Pascal:    if ch in ['a', 'f', 'g'] then
' Basic:     If Among(ch, "a", "f", "g") Then
Function Among(vTarget As Variant, ParamArray A() As Variant) As Boolean
    Among = True    ' Assume found
    Dim v As Variant
    For Each v In A()
        If v = vTarget Then Exit Function
    Next
    Among = False
End Function

' Work around limitation of AddressOf
'    Call like this: procVar = GetProc(AddressOf ProcName)
Function GetProc(proc As Long) As Long
    GetProc = proc
End Function

' I accidentally left two wrap functions, this and WordWrap, in the
' module for the second edition. I can't remove one without breaking
' compatibility, so make this one a wrapper for the other.
Function LineWrap(sText As String, cMax As Integer)
    LineWrap = WordWrap(sText, cMax)
End Function

Function WordWrap(sText As String, ByVal cMax As Long) As String
    Dim iStart As Long, iEnd As Long, cText As Long, sSep As String
    cText = Len(sText)
    iStart = 1
    iEnd = cMax
    sSep = " " & sTab & sCrLf
    Do While iEnd < cText
        ' Parse back to white space
        Do While InStr(sSep, Mid$(sText, iEnd, 1)) = 0
            iEnd = iEnd - 1
            ' Don't send us text with words longer than the lines!
            If iEnd = iStart Then ErrRaise eeCantWrapLine
            ' Throw away line separators
            Do While InStr(sCrLf, Mid$(sText, iEnd, 1))
                 iEnd = iEnd - 1
            Loop
        Loop
        WordWrap = WordWrap & Mid$(sText, iStart, iEnd - iStart + 1) & sCrLf
        iStart = iEnd + 1
        ' Parse forward to throw away white space
        Do While InStr(sSep, Mid$(sText, iStart, 1))
            iStart = iStart + 1
        Loop
        
        iEnd = iStart + cMax
    Loop
    WordWrap = WordWrap + Mid$(sText, iStart)
End Function

Sub CollectionReplace(n As Collection, vIndex As Variant, _
                      vVal As Variant)
    If VarType(vIndex) = vbString Then
        n.Remove vIndex
        n.Add vVal, vIndex
    Else
        n.Add vVal, , vIndex
        n.Remove vIndex + 1
    End If
End Sub

Function GetLabel(sRoot As String) As String
    GetLabel = Dir$(sRoot & "*.*", vbVolume)
End Function

Function GetFileBase(sFile As String) As String
    Dim iBase As Long, iExt As Long, s As String
    If sFile = sEmpty Then Exit Function
    s = GetFullPath(sFile, iBase, iExt)
    GetFileBase = Mid$(s, iBase, iExt - iBase)
End Function

Function GetFileBaseExt(sFile As String) As String
    Dim iBase As Long, s As String
    If sFile = sEmpty Then Exit Function
    s = GetFullPath(sFile, iBase)
    GetFileBaseExt = Mid$(s, iBase)
End Function

Function GetFileExt(sFile As String) As String
    Dim iExt As Long, s As String
    If sFile = sEmpty Then Exit Function
    s = GetFullPath(sFile, , iExt)
    GetFileExt = Mid$(s, iExt)
End Function

Function GetFileDir(sFile As String) As String
    Dim iBase As Long, s As String
    If sFile = sEmpty Then Exit Function
    s = GetFullPath(sFile, iBase)
    GetFileDir = Left$(s, iBase - 1)
End Function

Function GetFileFullSpec(sFile As String) As String
    If sFile = sEmpty Then Exit Function
    GetFileFullSpec = GetFullPath(sFile)
End Function

Function SearchForExe(sName As String) As String
    Dim sSpec As String, asExt(1 To 5) As String, i As Integer
    asExt(1) = ".EXE": asExt(2) = ".COM": asExt(3) = ".PIF":
    asExt(4) = ".BAT": asExt(5) = ".CMD"
    For i = 1 To 5
        sSpec = SearchDirs(sName, asExt(i))
        If sSpec <> sEmpty Then Exit For
    Next
    SearchForExe = sSpec
End Function

Function IsExe() As Boolean
    Dim sExe As String, vVer As Variant
    sExe = RealExeName
    vVer = Mid$(sExe, 3, InStr(sExe, ".") - 3)
    ' Pretty good test that recognizes VB7.EXE or VB10.EXE
    If (Right$(sExe, 4) <> ".EXE") Or _
       (Left$(sExe, 2) <> "VB") And _
       (Not IsNumeric(vVer)) Then
       IsExe = True
    End If
End Function

' Gets the name of the executable even if it is VB in the IDE
Function RealExeName() As String
    Dim sExe As String, c As Long
    sExe = String$(255, 0)
    c = GetModuleFileName(hNull, sExe, 255)
    sExe = UCase$(Left$(sExe, c))
    RealExeName = GetFileBaseExt(sExe)
End Function

' Use with API functions that use RECT, which likes right and bottom
' rather than width and height
Function xRight(obj As Object) As Single
    xRight = obj.Left + obj.Width
End Function

Function yBottom(obj As Object) As Single
    yBottom = obj.Top + obj.Height
End Function

' Win32 functions with Basic interface

' GetFullPath - Basic version of Win32 API emulation routine. It returns a
' BSTR, and indexes to the file name, directory, and extension parts of the
' full name.
'
' Input:  sFileName - file to be qualified in one of these formats:
'
'              [relpath\]file.ext
'              \[path\]file.ext
'              .\[path\]file.ext
'              d:\[path\]file.ext
'              ..\[path\]file.ext
'              \\server\machine\[path\]file.ext
'          iName - variable to receive file name position
'          iDir - variable to receive directory position
'          iExt - variable to receive extension position
'
' Return: Full path name, or an empty string on failure
'
' Errors: Any of the following:
'              ERROR_BUFFER_OVERFLOW      = 111
'              ERROR_INVALID_DRIVE        = 15
'              ERROR_CALL_NOT_IMPLEMENTED = 120
'              ERROR_BAD_PATHNAME         = 161


Function GetFullPath(sFileName As String, _
                     Optional FilePart As Long, _
                     Optional ExtPart As Long, _
                     Optional DirPart As Long) As String

    Dim c As Long, p As Long, sRet As String
    If sFileName = sEmpty Then Exit Function
    
    ' Get the path size, then create string of that size
    sRet = String(cMaxPath, 0)
    c = GetFullPathName(sFileName, cMaxPath, sRet, p)
    If c = 0 Then ApiRaise Err.LastDllError
    BugAssert c <= cMaxPath
    sRet = Left$(sRet, c)

    ' Get the directory, file, and extension positions
    GetDirExt sRet, FilePart, DirPart, ExtPart
    GetFullPath = sRet
    
End Function

Function GetTempFile(Optional Prefix As String, _
                     Optional PathName As String) As String
    
    If Prefix = sEmpty Then Prefix = sEmpty
    If PathName = sEmpty Then PathName = GetTempDir
    
    Dim sRet As String
    sRet = String(cMaxPath, 0)
    GetTempFileName PathName, Prefix, 0, sRet
    ApiRaiseIf Err.LastDllError
    GetTempFile = GetFullPath(StrZToStr(sRet))
End Function

Function GetTempDir() As String
    Dim sRet As String, c As Long
    sRet = String(cMaxPath, 0)
    c = GetTempPath(cMaxPath, sRet)
    If c = 0 Then ApiRaise Err.LastDllError
    GetTempDir = Left$(sRet, c)
End Function

Function SearchDirs(sFileName As String, _
                    Optional Ext As String, _
                    Optional Path As String, _
                    Optional FilePart As Long, _
                    Optional ExtPart As Long, _
                    Optional DirPart As Long) As String

    Dim p As Long, c As Long, sRet As String

    If sFileName = sEmpty Then ApiRaise ERROR_INVALID_PARAMETER

    ' Handle missing or invalid extension or path
    If Ext = sEmpty Then Ext = sNullStr
    If Path = sEmpty Then Path = sNullStr
    
    ' Get the file (treating empty strings as NULL pointers)
    sRet = String$(cMaxPath, 0)
    c = SearchPath(Path, sFileName, Ext, cMaxPath, sRet, p)
    If c = 0 Then
        If Err.LastDllError = ERROR_FILE_NOT_FOUND Then Exit Function
        ApiRaise Err.LastDllError
    End If
    BugAssert c <= cMaxPath
    sRet = Left$(sRet, c)

    ' Get the directory, file, and extension positions
    GetDirExt sRet, FilePart, DirPart, ExtPart
    SearchDirs = sRet
   
End Function

Private Sub GetDirExt(sFull As String, iFilePart As Long, _
                      iDirPart As Long, iExtPart As Long)

    Dim iDrv As Long, i As Long, cMax As Long
    cMax = Len(sFull)

    iDrv = Asc(UCase$(Left$(sFull, 1)))

    ' If in format d:\path\name.ext, return 3
    If iDrv <= 90 Then                          ' Less than Z
        If iDrv >= 65 Then                      ' Greater than A
            If Mid$(sFull, 2, 1) = ":" Then     ' Second character is :
                If Mid$(sFull, 3, 1) = "\" Then ' Third character is \
                    iDirPart = 3
                End If
            End If
        End If
    Else

        ' If in format \\machine\share\path\name.ext, return position of \path
        ' First and second character must be \
        If Mid$(sFull, 1, 2) <> "\\" Then ApiRaise ERROR_BAD_PATHNAME

        Dim fFirst As Boolean
        i = 3
        Do
            If Mid$(sFull, i, 1) = "\" Then
                If Not fFirst Then
                    fFirst = True
                Else
                    iDirPart = i
                    Exit Do
                End If
            End If
            i = i + 1
        Loop Until i = cMax
    End If

    ' Start from end and find extension
    iExtPart = cMax + 1       ' Assume no extension
    fFirst = False
    Dim sChar As String
    For i = cMax To iDirPart Step -1
        sChar = Mid$(sFull, i, 1)
        If Not fFirst Then
            If sChar = "." Then
                iExtPart = i
                fFirst = True
            End If
        End If
        If sChar = "\" Then
            iFilePart = i + 1
            Exit For
        End If
    Next
    Exit Sub
FailGetDirExt:
    iFilePart = 0
    iDirPart = 0
    iExtPart = 0
End Sub

#If fComponent Then
' Seed the component's copy of the random number generator
Sub CoreRandomize(Optional Number As Long)
    Randomize Number
End Sub

Function CoreRnd(Optional Number As Long)
    CoreRnd = Rnd(Number)
End Function
#End If

' Returns a line from a string, where a "line" is all characters
' up to and including a carriage return/line feed. GetNextLine
' works the same way as GetToken. The first call to GetNextLine
' should pass the string to parse; subsequent calls should pass
' an empty string. GetNextLine returns an empty string after all lines
' have been read from the source string.
Function GetNextLine(Optional sSource As String) As String
    Static sSave As String, iStart As Long, cSave As Long
    Dim iEnd As Long
    
    ' Initialize GetNextLine
    If (sSource <> sEmpty) Then
        iStart = 1
        sSave = sSource
        cSave = Len(sSave)
    Else
        If sSave = sEmpty Then Exit Function
    End If
    
    ' iStart points to first character after the previous sCrLf
    iEnd = InStr(iStart, sSave, sCrLf)
    
    If iEnd > 0 Then
        ' Return line
        GetNextLine = Mid$(sSave, iStart, iEnd - iStart + 2)
        iStart = iEnd + 2
        If iStart > cSave Then sSave = sEmpty
    Else
        ' Return remainder of string as a line
        GetNextLine = Mid$(sSave, iStart) & sCrLf
        sSave = sEmpty
    End If
End Function

' Strips off trailing carriage return/line feed
Function RTrimLine(sLine As String) As String
    If Right$(sLine, 2) = sCrLf Then
        RTrimLine = Left$(sLine, Len(sLine) - 2)
    Else
        RTrimLine = sLine
    End If
End Function



