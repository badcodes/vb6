VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFilterTypeName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Implements ITemplateFilter
'Private Const SpaceChar = " "
'Private Const TOKEN_SEPARATOR As String = vbTab & SpaceChar
Private Const KeywordsBeforeTypeName As String = " as new type enum "
Private Const KeywordsBeforeInterface As String = " implements "
Private Const KeywordsBeforeFunction As String = " function sub get let set "


'Private Function ITemplateFilter_Process(vSrc As String, vDest As String, ByRef tInfo As ITypeInfo, ByRef Reportor As IFilterReporter) As Boolean
'
'    Dim expType As RegExp
'    Set expType = New RegExp
'    expType.Global = True
'    expType.IgnoreCase = True
'    expType.MultiLine = False
'    expType.Pattern = "\s+as\s+(\w+)"
'
'    Dim regMatches As MatchCollection
'    Dim regMatch As Match
'    Set regMatches = expType.Execute(vSrc)
'
'    ITemplateFilter_Process = False
'    If regMatches.Count < 1 Then Exit Function
'
'    Dim sTypeName As String
'    Dim nStart As Long
'    Dim nStop As Long
'    Dim nId As Long
'    nStart = 1
'
'    vDest = ""
'    For Each regMatch In regMatches
'        sTypeName = regMatch.SubMatches.Item(0)
'        nId = tInfo.IdOf(sTypeName)
'        If (nId >= 0) Then
'            nStop = regMatch.FirstIndex
'            vDest = vDest & Mid$(vSrc, nStart, nStop - nStart + 1) & " AS " & tInfo.NameOf(nId)
'            nStart = nStop + regMatch.Length + 1
'            Reportor.NotifyChange "As " & sTypeName, "As " & tInfo.NameOf(nId)
'            ITemplateFilter_Process = True
'        End If
'    Next
'    nStop = Len(vSrc)
'    If nStop >= nStart Then
'        vDest = vDest & Mid$(vSrc, nStart, nStop - nStart + 1)
'    End If
'
'End Function

Private Function ITemplateFilter_Process(Lines() As String, ByVal count As Long, Types As ITypeInfo, Reportor As IFilterReporter) As Boolean

    Dim i As Long
    Dim Token As String
    Dim fChanged As Boolean
    Dim fEspectTypename As Boolean
    Dim fEspectInterface As Boolean
    Dim fEspectFunctionName As Boolean
    Dim sPadToken As String
    Dim nId As Long
    Dim Interfaces() As String
    Dim cIFace As Long
    Dim j As Long
    Dim sSuffix As String
    Dim sPrefix As String
    Dim sOld As String
    Dim sNew As String
    Dim fModified As Boolean
    Dim ch As String
    
    ITemplateFilter_Process = False
    If count < 1 Then Exit Function
    If Types Is Nothing Then Exit Function
    If Types.count < 1 Then Exit Function
    
    For i = 0 To count - 1
        Token = GetToken(Lines(i), TOKEN_SEPARATOR)
        If Left$(Token, 1) = "'" Then GoTo Line_Next
        fModified = False
        fEspectTypename = False
        fEspectInterface = False
        fEspectFunctionName = False
        sNew = ""
        sPrefix = ""
        sSuffix = ""

        Do While Token <> ""
Token_Continue:
            fChanged = False
            If fEspectTypename Then
                Dim idot As Integer
                idot = InStr(Token, ".")
                If idot > 0 Then
                    nId = Types.IdOf(Left$(Token, idot - 1))
                    sSuffix = Mid$(Token, idot)
                ElseIf Right$(Token, 2) = "()" Then
                    nId = Types.IdOf(Left$(Token, Len(Token) - 2))
                    ch = "()"
                ElseIf Right$(Token, 1) = "," Then
                    ch = ","
                    nId = Types.IdOf(Left$(Token, Len(Token) - 1))
                ElseIf Right$(Token, 1) = ")" Then
                    nId = Types.IdOf(Left$(Token, Len(Token) - 1))
                    ch = ")"
                ElseIf InStr(Token, ".") Then
                Else
                    ch = ""
                    nId = Types.IdOf(Token)
                End If
                If (nId >= 0) Then
                    If ch <> "" Then sSuffix = ch
                    fChanged = True
                    Token = Types.NameOf(nId)
                    GoTo Token_Next
                End If
            ElseIf fEspectInterface Then
                ReDim Preserve Interfaces(0 To cIFace) As String
                Interfaces(cIFace) = Token
                fEspectTypename = True
                fEspectInterface = False
                GoTo Token_Continue
            ElseIf fEspectFunctionName Then
                nId = InStr(1, Token, "_")
                If (nId > 0) Then
                    sPadToken = Left$(Token, nId - 1)
                    For j = 0 To cIFace - 1
                        If Interfaces(j) = sPadToken Then
                            sSuffix = Mid$(Token, nId)
                            Token = sPadToken
                            fEspectTypename = True
                            GoTo Token_Continue
                        End If
                    Next
                End If
            End If
            sPadToken = SpaceChar & LCase$(Token) & SpaceChar
            If InStr(1, KeywordsBeforeTypeName, sPadToken) Then
               fEspectTypename = True
            ElseIf InStr(1, KeywordsBeforeInterface, sPadToken) Then
                fEspectInterface = True
            ElseIf InStr(1, KeywordsBeforeFunction, sPadToken) Then
                fEspectFunctionName = True
            End If
Token_Next:
            sNew = sNew & sPrefix & Token & sSuffix & " "
            If (fChanged) Then
                fModified = True
                sPrefix = ""
                sSuffix = ""
            End If
            Token = GetToken("", TOKEN_SEPARATOR)
        Loop
        If (fModified) Then
            sOld = Lines(i)
            MFilterHelper.AssignString Lines(i), sNew
            Reportor.NotifyChange sOld, Lines(i), i
        End If
Line_Next:
    Next
End Function

