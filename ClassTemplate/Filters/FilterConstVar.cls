VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFilterConstVar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ITemplateFilter

'Private Function ITemplateFilter_Process(vSrc As String, vDest As String, Types As ITypeInfo, ByRef Reportor As IFilterReporter) As Boolean
'
'    Dim Lines() As String
'    Dim sToken As String
'    Dim cPre As Long
'    Dim cSuf As Long
'    Dim sPre As String
'    Dim sSuf As String
'    Dim sTypeName As String
'    Dim sNew As String
'    Dim nId As Long
'    Dim fChanged As Boolean
'    Dim fInConst As Boolean
'    Dim fAssignConst As Boolean
'    Dim fModified As Boolean
'
'    ITemplateFilter_Process = False
'
'    If vSrc = "" Then Exit Function
'    If Types Is Nothing Then Exit Function
'    If Types.Count < 1 Then Exit Function
'
'    sPre = Types.ConstTypePrefix
'    sSuf = Types.ConstTypeSuffix
'    cPre = Len(sPre)
'    cSuf = Len(sSuf)
'
'    Lines = Split(vSrc, vbCrLf)
'    Dim i As Long
'    For i = LBound(Lines) To UBound(Lines)
'        sToken = MParse.GetToken(Lines(i), " ")
'        If Left$(sToken, 1) = "#" Then
'            If StrComp(sToken, "#const", vbTextCompare) = 0 Then fInConst = True Else fInConst = False
'            fAssignConst = False
'            fChanged = False
'            nId = -1
'            sNew = sToken & " "
'            sToken = MParse.GetToken("", " ")
'            Do While sToken <> ""
'                If fInConst And nId >= 0 And sToken = "=" Then
'                    fAssignConst = True
'                    sNew = sNew & sToken & " "
'                    GoTo NextToken
'                End If
'                If fAssignConst Then
'                    fAssignConst = False
'                    fChanged = True
'                    sNew = sNew & Types.ConstVarOf(Types.StyleOf(nId))
'                    nId = -1
'                    GoTo NextToken
'                End If
'                If (Left$(sToken, cPre) = sPre And Right$(sToken, cSuf) = sSuf) Then
'                    sTypeName = Mid$(sToken, cPre + 1, Len(sToken) - cPre - cSuf)
'                Else
'                    sTypeName = sToken
'                End If
'                nId = Types.IdOf(sTypeName)
'                If (nId >= 0) Then
'                    fChanged = True
'                    sToken = sPre & Types.NameOf(nId) & sSuf
'                End If
'                sNew = sNew & sToken & " "
'NextToken:
'                sToken = MParse.GetToken("", " ")
'            Loop
'            If fChanged Then
'                fModified = True
'                Reportor.NotifyChange Lines(i), sNew
'                Lines(i) = sNew
'            End If
'        End If
'    Next
'
'    If fModified Then
'        vDest = ""
'        For i = LBound(Lines) To UBound(Lines)
'            vDest = vDest & Lines(i) & vbCrLf
'        Next
'        ITemplateFilter_Process = True
'    End If
'
'End Function


Private Function ITemplateFilter_Process(Lines() As String, ByVal Count As Long, Types As ITypeInfo, Reportor As IFilterReporter) As Boolean

    Dim sToken As String
    Dim cPre As Long
    Dim cSuf As Long
    Dim sPre As String
    Dim sSuf As String
    Dim sTypeName As String
    Dim sNew As String
    Dim nId As Long
    Dim fChanged As Boolean
    Dim fInConst As Boolean
    Dim fAssignConst As Boolean
    Dim fModified As Boolean
    Dim sOld As String
    ITemplateFilter_Process = False
    
    If Count < 1 Then Exit Function
    If Types Is Nothing Then Exit Function
    If Types.Count < 1 Then Exit Function
    
    sPre = Types.ConstTypePrefix
    sSuf = Types.ConstTypeSuffix
    cPre = Len(sPre)
    cSuf = Len(sSuf)
    
    'Dim c As String
    Dim i As Long
    For i = 0 To Count - 1
        sToken = MParse.GetToken(Lines(i), TOKEN_SEPARATOR)
'        c =
'        If c = "'" Then GoTo Line_Next
        If Left$(sToken, 1) = "#" Then
            If StrComp(sToken, "#const", vbTextCompare) = 0 Then fInConst = True Else fInConst = False
            fAssignConst = False
            fChanged = False
            nId = -1
            sNew = sToken & " "
            sToken = MParse.GetToken("", TOKEN_SEPARATOR)
            Do While sToken <> ""
                If fInConst And nId >= 0 And sToken = "=" Then
                    fAssignConst = True
                    sNew = sNew & sToken & " "
                    GoTo NextToken
                End If
                If fAssignConst Then
                    fAssignConst = False
                    fChanged = True
                    sNew = sNew & Types.ConstVarOf(Types.StyleOf(nId))
                    nId = -1
                    GoTo NextToken
                End If
                If (Left$(sToken, cPre) = sPre And Right$(sToken, cSuf) = sSuf) Then
                    sTypeName = Mid$(sToken, cPre + 1, Len(sToken) - cPre - cSuf)
                Else
                    sTypeName = sToken
                End If
                nId = Types.IdOf(sTypeName)
                If (nId >= 0) Then
                    fChanged = True
                    sToken = sPre & Types.NameOf(nId) & sSuf
                End If
                sNew = sNew & sToken & " "
NextToken:
                sToken = MParse.GetToken("", TOKEN_SEPARATOR)
            Loop
            If fChanged Then
                fModified = True
                sOld = Lines(i)
                MFilterHelper.AssignString Lines(i), sNew
                Reportor.NotifyChange sOld, Lines(i), i
                ITemplateFilter_Process = True
            End If
        End If
Line_Next:
    Next
    
End Function
