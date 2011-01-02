Attribute VB_Name = "MRegTypeLib"
Option Explicit

' This program uses no type library because I didn't want to require
' a type library to register a type library. It uses no other modules
' because most of my modules require a type library. I've violated
' my general reuse principles in a good cause.

Declare Function GetFullPathName Lib "kernel32" _
    Alias "GetFullPathNameA" (ByVal lpFileName As String, _
    ByVal nBufferLength As Long, ByVal lpBuffer As String, _
    ByVal lpFilePart As Long) As Long
    
Declare Sub ExitProcess Lib "kernel32" (ByVal ExitCode As Long)
    

Const sEmpty = ""
Const sQuote2 = """"
Const cMaxPath = 260

Sub Main()
    Dim fSilent As Boolean, fVerbose As Boolean, fUnregister As Boolean
    Dim sCmd As String, i As Integer, sFail As String
    Dim sSep As String, sToken As String, sLib As String, sUn As String
    sCmd = Command$
    If sCmd = sEmpty Then
        sCmd = InputBox("Enter type library name and path: ")
        If sCmd = sEmpty Then End
    End If
    sSep = " " & sTab
    
    ' Parse command line
    sToken = GetQToken(sCmd, sSep)
    Do While sToken <> sEmpty
        If InStr("/-", Left$(sToken, 1)) Then
            Select Case UCase$(Mid$(sToken, 2, 1))
            Case "S"
                fSilent = True
            Case "V"
                fVerbose = True
            Case "U"
                fUnregister = True
                sUn = "un"
            Case Else
                ShowSyntax "Unknown option", fSilent
                End
            End Select
        Else
            sLib = GetWholePath(sToken)
        End If
        sToken = GetQToken(sEmpty, sSep)
    Loop
    
    Dim sExt As String
    Dim sBase As String, sFull As String
    Dim iExt As Long, iBase As Long
    ' Validate extension
    iExt = GetExtPos(sLib)
    iBase = GetBasePos(sLib)
    sFull = sLib
    sExt = Mid$(sFull, iExt)
    sBase = Mid$(sFull, iBase, iExt - iBase)
    Select Case UCase$(sExt)
    Case sEmpty
        ShowSyntax "No extension given", fSilent
        End
    Case ".TLB", ".OLB", ".DLL"
    Case Else
        ShowSyntax "Unknown extension", fSilent
        End
    End Select
        
    ' Register full name if given, or try to create 16/32 names
    If sFull = sEmpty Then
        ShowSyntax "File not found", fSilent
    Else
        If fUnregister Then
            sFail = UnRegTypeLib(sFull)
        Else
            sFail = RegTypeLib(sFull)
        End If
        If sFail <> sEmpty Then
            If Not fSilent Then
                MsgBox "Can't " & sUn & "register type library: " & _
                        sLib & vbCrLf & sFail
            End If
            ExitProcess 1
        Else
            If fVerbose Then
                MsgBox "Type library " & sUn & "registered: " & sLib
            End If
            ExitProcess 0
        End If
    End If
End Sub

Function RegTypeLib(sLib As String) As String
    Dim mgr As TLIApplication
    On Error GoTo FailRegTypeLib
    Set mgr = New TLIApplication
    mgr.TypeLibInfoFromFile(sLib).Register
    Exit Function
FailRegTypeLib:
    ' Pass error message back to caller
    RegTypeLib = Err.Description
End Function

Function UnRegTypeLib(sLib As String) As String
    Dim mgr As TLIApplication
    On Error GoTo FailUnRegTypeLib
    Set mgr = New TLIApplication
    mgr.TypeLibInfoFromFile(sLib).UnRegister
    Exit Function
FailUnRegTypeLib:
    UnRegTypeLib = Err.Description
End Function

Sub ShowSyntax(sErr As String, fSilent As Boolean)
    If fSilent Then Exit Sub
    Dim sMsg As String
    Const sProg = "REGTLB32"
    sMsg = sErr & sCr & sCr
    sMsg = sMsg & _
        sTab & "Syntax: " & sProg & " [/s] libname.ext" & sCr & sCr & _
        sTab & "/s - Silent (don't show this message box)" & sCr & _
        sTab & "/v - Verbose (report success)" & sCr & sCr
    sMsg = sMsg & sProg & _
        " will attempt to register both 16-bit and 32-bit libraries." & sCr & _
        "For example, to register WIN16.TLB and WIN32.TLB, give any " & sCr & _
        "of these commands: " & sCr & sCr
    sMsg = sMsg & sTab & sProg & " WIN.TLB" & sCr
    sMsg = sMsg & sTab & sProg & " WIN32.TLB" & sCr
    sMsg = sMsg & sTab & sProg & " WIN16.TLB" & sCr
    MsgBox sMsg
End Sub

' Some functions duplicated from other modules, but we don't want to use
' the Windows API type library in this program.

Function GetQToken(sTarget As String, sSeps As String) As String
    ' GetQToken = sEmpty

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
    GetQToken = Mid$(sSave, iStart, iNew - iStart)
    
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

Function GetExtPos(sSpec As String) As Integer
    Dim iLast As Integer, iExt As Integer
    iLast = Len(sSpec)
    
    ' Assume no extension
    GetExtPos = iLast + 1
    ' Parse backward to find extension or base
    For iExt = iLast + 1 To 1 Step -1
        Select Case Mid$(sSpec, iExt, 1)
        Case "."
            ' First . from right is extension start
            GetExtPos = iExt
            Exit Function
        Case "\", ":"
            ' First \ or : from right is base start, so no extension
            Exit Function
        End Select
    Next
    ' Fall through means no extension
End Function

Function GetBasePos(sFile As String) As Integer
    Dim iLast As Integer, iBase As Integer
    iLast = Len(sFile)
    
    ' Assume no directory
    GetBasePos = 1
    
    ' Parse backward to find base
    For iBase = iLast + 1 To 1 Step -1
        Select Case Mid$(sFile, iBase, 1)
        Case "\", ":"
            ' First \ or : from right is base start
            GetBasePos = iBase + 1
            Exit For
        End Select
    Next
End Function

' Defined in type library, but we must define for others
Property Get sCr() As String
    sCr = Chr$(13)
End Property

Property Get sTab() As String
    sTab = Chr$(9)
End Property

Function GetWholePath(sFileName As String)

    Dim c As Long, p As Long, sRet As String
    If sFileName = sEmpty Then Exit Function
    
    ' Get the path size, then create string of that size
    sRet = String(cMaxPath, 0)
    c = GetFullPathName(sFileName, cMaxPath, sRet, p)
    If c = 0 Then Exit Function
    sRet = Left$(sRet, c)

    GetWholePath = sRet
    
End Function


