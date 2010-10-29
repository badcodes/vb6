Attribute VB_Name = "MParse"
Option Explicit

Public Enum EErrorParse
    eeBaseParse = 13550
End Enum
Private Const sEmpty  As String = ""
Private Const sQuote2 As String = """"

Function GetQToken(sTarget As String, sSeps As String) As String
    ' Assume failure
    GetQToken = sEmpty

    ' Note that sSave and iStart must be static from call to call
    ' If first call, make copy of string
    Static sSave As String, iStart As Integer, cSave As Integer
    Dim iNew As Integer, fQuote As Integer
    If (sTarget <> sEmpty) Then
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
    If (iStart = 1) Then
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
    GetQToken = Mid$(sSave, iStart, iNew - iStart)
    
    ' Set new starting position
    iStart = iNew

End Function

Function GetToken(sTarget As String, sSeps As String) As String
    
    ' Assume failure
    GetToken = sEmpty
    
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
    GetToken = Mid$(sSave, iStart, iNew - iStart)
    ' Set new starting position
    iStart = iNew

End Function

Function StrBreak(sTarget As String, ByVal iStart As Integer, sSeps As String) As Integer
    
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

Function StrSpan(sTarget As String, ByVal iStart As Integer, sSeps As String) As Integer
    
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

