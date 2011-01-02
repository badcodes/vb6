VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GParse"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum EErrorParse
    eeBaseParse = 13550
End Enum

Private Declare Function StrSpn Lib "SHLWAPI" Alias "StrSpnW" ( _
        ByVal psz As Long, ByVal pszSet As Long) As Long
Private Declare Function StrCSpn Lib "SHLWAPI" Alias "StrCSpnW" ( _
        ByVal lpStr As Long, ByVal lpSet As Long) As Long

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
        On Error GoTo Fail
        Call StrSpn(StrPtr("a"), StrPtr("a"))
    End If
    Exit Property
Fail:
    fNoShlWapiI = True
End Property
#End If

' GetQTokenO uses our StrSpan and StrBreak
Private Function GetQTokenO(sTarget As String, sSeps As String) As String
    ' GetQTokenO = sEmpty

    ' Note that sSave and iStart must be static from call to call
    ' If first call, make copy of string
    Static sSave As String, iStart As Integer, cSave As Integer
    Dim iNew As Integer, fQuote As Integer
    If sTarget <> sEmpty Then
        iStart = 1
        sSave = sTarget
        cSave = Len(sSave)
    Else
        If sSave = sEmpty Then Exit Function
    End If
    ' Make sure separators includes quote
    sSeps = sSeps & sQuote2

    ' Find start of next token
    iNew = StrSpan(sSave, iStart, sSeps)
    If iNew Then
        ' Set position to start of token
        iStart = iNew
    Else
        ' If no new token, return empty string
        sSave = sEmpty
        Exit Function
    End If
    
    ' Find end of token
    If iStart = 1 Then
        iNew = StrBreak(sSave, iStart, sSeps)
    ElseIf Mid$(sSave, iStart - 1, 1) = sQuote2 Then
        iNew = StrBreak(sSave, iStart, sQuote2)
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
    ' GetQToken = sEmpty
    If fNoShlWapi Then
        GetQToken = GetQTokenO(sTarget, sSeps)
        Exit Function
    End If
    
    ' Note that sSave, pSave, pCur, and cSave must be static between calls
    Static sSave As String, pSave As Long, pCur As Long, cSave As Long
    ' First time through save start and length of string
    If sTarget <> sEmpty Then
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
    sSepsNew = sSeps & sQuote2
    pSeps = StrPtr(sSepsNew)

    ' Get current character
    Dim pNew As Long, c As Long
    ' Find start of next token
    c = StrSpn(pCur, pSeps)
    ' Set position to start of token
    If c Then pCur = pCur + (c * 2)
    
    Dim ch As Integer
    Const chQuote = 34  ' Asc("""")
    CopyMemory ch, ByVal pCur - 2, 2
    ' Check first character for quote, then find end of token
    If ch = chQuote Then
        c = StrCSpn(pCur, StrPtr(sQuote2))
    Else
        c = StrCSpn(pCur, pSeps)
    End If
    ' If token length is zero, we're at end
    If c = 0 Then Exit Function
    
    ' Cut token out of target string
    GetQToken = String$(c, 0)
    CopyMemory ByVal StrPtr(GetQToken), ByVal pCur, c * 2
    ' Set new starting position
    pCur = pCur + (c * 2)

End Function

' Old GetTokenO uses our StrSpan and StrBreak
Private Function GetTokenO(sTarget As String, sSeps As String) As String
    ' GetTokenO = sEmpty
    
    ' Note that sSave and iStart must be static from call to call
    ' If first call, make copy of string
    Static sSave As String, iStart As Integer, cSave As Integer
    
    If sTarget <> sEmpty Then
        iStart = 1
        sSave = sTarget
        cSave = Len(sSave)
    Else
        If sSave = sEmpty Then Exit Function
    End If
    
    ' Find start of next token
    Dim iNew As Integer
    iNew = StrSpan(sSave, iStart, sSeps)
    If iNew Then
        ' Set position to start of token
        iStart = iNew
    Else
        ' If no new token, return empty string
        sSave = sEmpty
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
    If sTarget <> sEmpty Then
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
    Dim pNew As Long, c As Long, pSeps As Long
    pSeps = StrPtr(sSeps)
    c = StrSpn(pCur, pSeps)
    ' Set position to start of token
    If c Then pCur = pCur + (c * 2)
    
    ' Find end of token
    c = StrCSpn(pCur, pSeps)
    ' If token length is zero, we're at end
    If c = 0 Then Exit Function
    
    ' Cut token out of target string
    GetToken = String$(c, 0)
    CopyMemory ByVal StrPtr(GetToken), ByVal pCur, c * 2
    ' Set new starting position
    pCur = pCur + (c * 2)

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
Function Splits(sExpression As String, _
                Optional sDelimiters As String = sWhiteSpace, _
                Optional cMax As Long = -1) As Variant
    Dim sToken As String, avsRet() As Variant, c As Long
    ' Error trap to resize on overflow
    On Error GoTo SplitResize
    ' Break into tokens and put in an array
    sToken = GetToken(sExpression, sDelimiters)
    Do While sToken <> sEmpty
        If cMax <> -1 Then If c >= cMax Then Exit Do
        avsRet(c) = sToken
        c = c + 1
        sToken = GetToken(sEmpty, sDelimiters)
    Loop
    ' Size is an estimate, so resize to counted number of tokens
    If c Then ReDim Preserve avsRet(0 To c - 1) As Variant
    Splits = avsRet
    Exit Function
    
SplitResize:
    ' Resize on overflow
    Const cChunk As Long = 20
    If Err.Number = eeOutOfBounds Then
        ReDim Preserve avsRet(0 To c + cChunk) As Variant
        Resume              ' Try again
    End If
    ErrRaise Err.Number     ' Other VB error for client
End Function

' Same as Splits, but handles quoted strings
Function QSplits(sExpression As String, _
                 Optional sDelimiters As String = sWhiteSpace, _
                 Optional cMax As Long = -1) As Variant
    Dim sToken As String, avsRet() As Variant, c As Long
    ' Error trap to resize on overflow
    On Error GoTo QSplitResize
    ' Break into tokens and put in an array
    sToken = GetQToken(sExpression, sDelimiters)
    Do While sToken <> sEmpty
        If cMax <> -1 Then If c >= cMax Then Exit Do
        avsRet(c) = sToken
        c = c + 1
        sToken = GetQToken(sEmpty, sDelimiters)
    Loop
    ' Size is an estimate, so resize to counted number of tokens
    If c Then ReDim Preserve avsRet(0 To c - 1) As Variant
    QSplits = avsRet
    Exit Function
    
QSplitResize:
    ' Resize on overflow
    Const cChunk As Long = 20
    If Err.Number = eeOutOfBounds Then
        ReDim Preserve avsRet(0 To c + cChunk) As Variant
        Resume              ' Try again
    End If
    ErrRaise Err.Number     ' Other VB error for client
End Function

Function GetFileTokens(sFile As String, _
                       Optional sDelimiters As String = sWhiteSpace _
                       ) As Variant
    ' File to string to token array
    GetFileTokens = Splits(MUtility.GetFileText(sFile), sDelimiters)
End Function

#If fComponent = 0 Then
Private Sub ErrRaise(e As Long)
    Dim sText As String, sSource As String
    If e > 1000 Then
        sSource = App.ExeName & ".Parse"
        Select Case e
        Case eeBaseParse
            BugAssert True
       ' Case ee...
       '     Add additional errors
        End Select
        Err.Raise COMError(e), sSource, sText
    Else
        ' Raise standard Visual Basic error
        sSource = App.ExeName & ".VBError"
        Err.Raise e, sSource
    End If
End Sub
#End If

