Attribute VB_Name = "MParseO"
Option Explicit

'
' GetQToken1:
'  Extracts tokens, including quoted tokens, from a string. A token
'  is a word that is surrounded by separators, such as spaces or commas.
'  Tokens are extracted and analyzed when parsing sentences or commands.
'  A quoted token is a group of characters surrounded by double quotes.
'  The quoted portion may contain the separator characters. To use the
'  GetQToken1 function, pass the string to be parsed on the first call,
'  then pass an empty string on subsequent calls until the function
'  returns an empty string to indicate that the entire string has been
'  parsed.
' Input:
'  sTarget = String to search
'  sSeps  = String of separators
' Output:
'  GetQToken1 = next token
'
Function GetQToken1(sTarget As String, sToken As String, sSeps As String) As Long

    ' Note that sSave and iStart must be static from call to call
    ' If first call, make copy of string
    Static sSave As String, iStart As Long, cSave As Long
    Dim iNew As Long, fQuote As Long
    If (sTarget <> sSave) Or (sSave = sEmpty) Then
        iStart = 1
        sSave = sTarget
        cSave = Len(sSave)
    End If
    sSeps = sSeps & """"

    GetQToken1 = False

    ' Find start of next token
    iNew = StrSpan1(Mid$(sSave, iStart, cSave), sSeps)
    If iNew Then
        ' Set position to start of token
        iStart = iNew + iStart - 1
    Else
        ' If no new token, quit and return empty string
        sToken = sEmpty
        Exit Function
    End If
    
    ' Find end of token
    If (iStart = 1) Then
        iNew = StrBreak1(Mid$(sSave, iStart, cSave), sSeps)
    ElseIf Mid$(sSave, iStart - 1, 1) = """" Then
        iNew = StrBreak1(Mid$(sSave, iStart, cSave), """")
    Else
        iNew = StrBreak1(Mid$(sSave, iStart, cSave), sSeps)
    End If

    If iNew Then
        ' Set position to end of token
        iNew = iStart + iNew - 1
    Else
        ' If no end of token, return set to end a value
        iNew = cSave + 1
    End If
    ' Cut token out of sTarget string
    GetQToken1 = True
    sToken = Mid$(sSave, iStart, iNew - iStart)
    
    ' Set new starting position
    iStart = iNew

End Function

''@B GetToken1
Function GetToken1(sTarget As String, sSeps As String) As String
    
    ' Note that sSave and iStart must be static from call to call
    ' If first call, make copy of string
    Static sSave As String, iStart As Long
    If sTarget <> sEmpty Then
        iStart = 1
        sSave = sTarget
    End If

    ' Find start of next token
    Dim iNew As Long
    iNew = StrSpan1(Mid$(sSave, iStart, Len(sSave)), sSeps)
    If iNew Then
        ' Set position to start of token
        iStart = iNew + iStart - 1
    Else
        ' If no new token, return empty string
        GetToken1 = sEmpty
        Exit Function
    End If
    
    ' Find end of token
    iNew = StrBreak1(Mid$(sSave, iStart, Len(sSave)), sSeps)
    If iNew Then
        ' Set position to end of token
        iNew = iStart + iNew - 1
    Else
        ' If no end of token, set to end of string
        iNew = Len(sSave) + 1
    End If
    ' Cut token out of sTarget string
    GetToken1 = Mid$(sSave, iStart, iNew - iStart)
    ' Set new starting position
    iStart = iNew

End Function ''@E GetToken1

Function GetToken2(sTarget As String, sSeps As String) As String
    
    ' Note that sSave and iStart must be static from call to call
    ' If first call, make copy of string
    Static sSave As String, iStart As Long, cSave As Long
    Dim iNew As Long
    If sTarget <> sEmpty Then
        iStart = 1
        sSave = sTarget
        cSave = Len(sSave)
    End If
    
    ' Find start of next token
    iNew = StrSpan1(Mid$(sSave, iStart, cSave), sSeps)
    If iNew Then
        ' Set position to start of token
        iStart = iNew + iStart - 1
    Else
        ' If no new token, return empty string
        GetToken2 = sEmpty
        Exit Function
    End If
    
    ' Find end of token
    iNew = StrBreak1(Mid$(sSave, iStart, cSave), sSeps)
    If iNew Then
        ' Set position to end of token
        iNew = iStart + iNew - 1
    Else
        ' If no end of token, set to end of string
        iNew = cSave + 1
    End If
    ' Cut token out of sTarget string
    GetToken2 = Mid$(sSave, iStart, iNew - iStart)
    ' Set new starting position
    iStart = iNew

End Function

Function GetToken3(sTarget As String, sSeps As String) As String
    
    ' Note that sSave and iStart must be static from call to call
    ' If first call, make copy of string
    Static sSave As String, iStart As Long, cSave As Long
    Dim iNew As Long
    If sTarget <> sEmpty Then
        iStart = 1
        sSave = sTarget
    End If

    ' Find start of next token
    ''@B CallStrSpan1
    iNew = StrSpan1(Mid$(sSave, iStart), sSeps)
    ''@E CallStrSpan1
    If iNew Then
        ' Set position to start of token
        iStart = iNew + iStart - 1
    Else
        ' If no new token, return empty string
        GetToken3 = sEmpty
        Exit Function
    End If
    
    ' Find end of token
    iNew = StrBreak1(Mid$(sSave, iStart), sSeps)
    If iNew Then
        ' Set position to end of token
        iNew = iStart + iNew - 1
    Else
        ' If no end of token, set to end of string
        iNew = cSave + 1
    End If
    ' Cut token out of sTarget string
    GetToken3 = Mid$(sSave, iStart, iNew - iStart)
    ' Set new starting position
    iStart = iNew

End Function

Function GetToken4(sTarget As String, sSeps As String) As String
    
    ' Note that sSave and iStart must be static from call to call
    ' If first call, make copy of string
    Static sSave As String, iStart As Long, cSave As Long
    Dim iNew As Long
    If sTarget <> sEmpty Then
        iStart = 1
        sSave = sTarget
        cSave = Len(sSave)
    End If
    
    ' Find start of next token
    ''@B CallStrSpan2
    iNew = StrSpan2(sSave, iStart, sSeps)
    ''@E CallStrSpan2
    If iNew Then
        ' Set position to start of token
        iStart = iNew
    Else
        ' If no new token, return empty string
        GetToken4 = sEmpty
        Exit Function
    End If
    
    ' Find end of token
    iNew = StrBreak2(sSave, iStart, sSeps)
    If iNew = 0 Then
        ' If no end of token, set to end of string
        iNew = cSave + 1
    End If
    ' Cut token out of sTarget string
    GetToken4 = Mid$(sSave, iStart, iNew - iStart)
    ' Set new starting position
    iStart = iNew

End Function

''@B GetToken5
Function GetToken5(sTarget As String, sSeps As String) As String
    
    ' Note that sSave and iStart must be static from call to call
    ' If first call, make copy of string
    Static sSave As String, iStart As Long, cSave As Long
    
    ' Assume failure
    GetToken5 = sEmpty
    
    If sTarget <> sEmpty Then
        iStart = 1
        sSave = sTarget
        cSave = Len(sSave)
    Else
        If sSave = sEmpty Then Exit Function
    End If
    
    ' Find start of next token
    Dim iNew As Long
    ''@B CallStrSpan2
    iNew = StrSpan2(sSave, iStart, sSeps)
    ''@E CallStrSpan2
    If iNew Then
        ' Set position to start of token
        iStart = iNew
    Else
        ' If no new token, return empty string
        Exit Function
    End If
    
    ' Find end of token
    iNew = StrBreak2(sSave, iStart, sSeps)
    If iNew = 0 Then
        ' If no end of token, set to end of string
        iNew = cSave + 1
    End If
    ' Cut token out of sTarget string
    GetToken5 = Mid$(sSave, iStart, iNew - iStart)
    ' Set new starting position
    iStart = iNew

End Function ''@E GetToken5

#If 0 Then
''@B GetToken
Function GetToken(sTarget As String, sSeps As String) As String
    ' GetToken = sEmpty
    
    ' Note that pSave, pCur, and cSave static from call to call
    Static pSave As Long, pCur As Long, cSave As Long
    ' First time through save start and length of string
    If sTarget <> sEmpty Then
        pSave = StrPtr(sTarget)
        pCur = pSave
        cSave = Len(sTarget)
    Else
        ' Quit if past end (also catches null or empty target)
        If pCur >= pSave + (cSave * 2) Then Exit Function
    End If
    
    ' Find start of next token
    Dim pNew As Long, c As Long
    c = StrSpn(pCur, sSeps)
    ' Set position to start of token
    If c Then pCur = pCur + (c * 2)
    
    ' Find end of token
    c = StrCSpn(pCur, sSeps)
    ' If token length is zero, we're at end
    If c = 0 Then Exit Function
    
    ' Cut token out of target string
    GetToken = String$(c, 0)
    CopyMemory ByVal StrPtr(GetToken), ByVal pCur, c * 2
    ' Set new starting position
    pCur = pCur + (c * 2)

End Function
''@E
#End If

'
' StrBreak1:
'  Searches sTarget to find the first character from among those in
'  sSeps. Returns the index of that character. This function can
'  be used to find the end of a token.
' Input:
'  sTarget = string to search
'  sSeps = characters to search for
' Output:
'  StrBreak1 = index to first match in sTarget or 0 if no match
'
Function StrBreak1(sTarget As String, sSeps As String) As Long

    Dim cTarget As Long, iStart As Long
    cTarget = Len(sTarget)
    iStart = 1
   
    ''@B StrBreak1
    ' Look for end of token (first character that is a separator)
    Do While InStr(sSeps, Mid$(sTarget, iStart, 1)) = 0
    ''@E StrBreak1
        If iStart > cTarget Then
            StrBreak1 = 0
            Exit Function
        Else
            iStart = iStart + 1
        End If
    Loop
    StrBreak1 = iStart
  
End Function

'
' StrSpan1:
'  Searches sTarget to find the first character that is not one of
'  those in sSeps. Returns the index of that character. This
'  function can be used to find the start of a token.
' Input:
'  sTarget = string to sTarget
'  sSeps = characters to sTarget for
' Output:
'  StrSpan1 = index to first nonmatch in sTarget or 0 if all match
'
''@B StrSpan1
Function StrSpan1(sTarget As String, sSeps As String) As Long

    Dim cTarget As Long, iStart As Long
    cTarget = Len(sTarget)
    iStart = 1
    ' Look for start of token (character that isn't a separator)
    Do While InStr(sSeps, Mid$(sTarget, iStart, 1))
        If iStart > cTarget Then
            StrSpan1 = 0
            Exit Function
        Else
            iStart = iStart + 1
        End If
    Loop
    StrSpan1 = iStart

End Function
''@E StrSpan1

Function StrBreak2(sTarget As String, ByVal iStart As Long, _
                   sSeps As String) As Long
    
    Dim cTarget As Long
    cTarget = Len(sTarget)
   
    ''@B StrBreak2
    ' Look for end of token (first character that is a separator)
    Do While InStr(sSeps, Mid$(sTarget, iStart, 1)) = 0
    ''@E StrBreak2
        If iStart > cTarget Then
            StrBreak2 = 0
            Exit Function
        Else
            iStart = iStart + 1
        End If
    Loop
    StrBreak2 = iStart

End Function

''@B StrSpan2
Function StrSpan2(sTarget As String, ByVal iStart As Long, _
                  sSeps As String) As Long
    
    Dim cTarget As Long
    cTarget = Len(sTarget)
    ' Look for start of token (character that isn't a separator)
    ''@B CallMid
    Do While InStr(sSeps, Mid$(sTarget, iStart, 1))
    ''@E CallMid
        If iStart > cTarget Then
            StrSpan2 = 0
            Exit Function
        Else
            iStart = iStart + 1
        End If
    Loop
    StrSpan2 = iStart

End Function
''@E StrSpan2

