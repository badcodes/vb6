Attribute VB_Name = "MHttpHeader"
Option Explicit
'Private Const zCharZero As String = Chr$(0)

Public Function HttpHeader_ParseRequest(ByVal vAction As String) As String
    Dim iPos As Long
    iPos = InStr(1, vAction, "GET ", vbTextCompare)
    If iPos < 1 Then iPos = InStr(1, vAction, "POST ", vbTextCompare)
    If iPos < 1 Then Exit Function
    HttpHeader_ParseRequest = Trim$(StripString(vAction, " ", " ", iPos))
End Function

Public Function HttpHeader_ParseResponse(ByVal vAction As String) As Long
    On Error Resume Next
    Dim iPos As Long
    iPos = InStr(1, vAction, "HTTP/", vbTextCompare)
    If iPos < 1 Then Exit Function
    HttpHeader_ParseResponse = CLng(Trim$(StripString(vAction, " ", " ", iPos)))
End Function
Public Function HttpHeaderGetField(ByRef vHeader As String, ByVal vFieldname As String) As String
    Dim ret As String
    Dim pPos As Long
    pPos = InStr(1, vHeader, vFieldname & ":", vbTextCompare)
    Dim pSkip As Long
    pSkip = Len(vFieldname) + 1
    If pPos < 1 Then Exit Function
    ret = SubStringUntilMatch(vHeader, pPos + Len(vFieldname), Chr$(0))
    If ret = "" Then ret = SubStringUntilMatch(vHeader, pPos + pSkip, vbCrLf)
    If ret = "" Then
        ret = Mid$(vHeader, pPos + pSkip)
    Else
        ret = ret & HttpHeaderGetField(Mid$(vHeader, pPos + pSkip), vFieldname)
    End If
    HttpHeaderGetField = Trim$(ret)
End Function
Public Function HttpHeaderSetCookie(ByVal vHeader As String) As String
    Dim ret As String
    Dim pPos As Long
    Dim pSkip As Long
    pPos = InStr(1, vHeader, "Set-Cookie:", vbTextCompare)
    pSkip = Len("Set-Cookie:")
    Do While pPos > 1
        ret = Trim$(SubStringUntilMatch(vHeader, pPos + pSkip, ";"))
        If ret <> "" Then HttpHeaderSetCookie = HttpHeaderSetCookie & ret & "; "
        pPos = InStr(pPos + pSkip + Len(ret), vHeader, "Set-Cookie:", vbTextCompare)
    Loop
    If Right$(HttpHeaderSetCookie, 2) = "; " Then HttpHeaderSetCookie = Left$(HttpHeaderSetCookie, Len(HttpHeaderSetCookie) - 2)
End Function
Public Function HttpHeaderMergeCookie(ByVal vCookieOld As String, ByVal vCookieNew As String) As String
'HttpHeaderMergeCookie = IIf((vCookieNew = ""), vCookieOld, vCookieNew & "; " & vCookieOld): Exit Function
    
    HttpHeaderMergeCookie = vCookieNew
    Dim pPairs() As String
    Dim pPairsCount As Long
    On Error Resume Next
    pPairs = Split(vCookieOld, ";")
    pPairsCount = UBound(pPairs) + 1
    If pPairsCount < 1 Then Exit Function
    Dim I As Long
    For I = 0 To pPairsCount - 1
        pPairs(I) = Trim$(pPairs(I))
        Dim pName As String
        pName = SubStringUntilMatch(pPairs(I), 1, "=")
        If pName <> "" And InStr(1, HttpHeaderMergeCookie, pName & "=", vbTextCompare) < 1 Then
            HttpHeaderMergeCookie = HttpHeaderMergeCookie & "; " & pPairs(I)
        End If
    Next
    If Left$(HttpHeaderMergeCookie, 2) = "; " Then HttpHeaderMergeCookie = Mid$(HttpHeaderMergeCookie, 3)
End Function

Public Function HttpHeader_Parse(ByVal vText As String) As String()
On Error Resume Next
    
    Dim result() As String
    
    If vText = "" Then GoTo Exit_HttpHeader_Parse
    
    Dim iEnd As Long
    Dim sHeaderMap() As String
    If InStr(vText, Chr$(0)) < 1 Then
        sHeaderMap = Split(vText, vbCrLf)
    Else
        sHeaderMap = Split(vText, Chr$(0))
    End If
    
    iEnd = ArrayUbound(sHeaderMap)
    If iEnd < 0 Then GoTo Exit_HttpHeader_Parse
    
    ReDim result(0 To 1, 0 To iEnd)
    
    Dim iStart As Long
    Dim fAction As Boolean
    Dim iPos As Long
    Dim I As Long
    
    For iStart = 0 To iEnd
        iPos = InStr(sHeaderMap(iStart), ":")
        I = I + 1
        If iPos > 1 Then
            result(0, I) = Trim$(Left$(sHeaderMap(iStart), iPos - 1))
            result(1, I) = Trim$(Mid$(sHeaderMap(iStart), iPos + 1))
        ElseIf Not fAction Then
            fAction = True
            result(0, 0) = Trim$(sHeaderMap(iStart))
            I = I - 1
        Else
            result(0, I) = Trim$(sHeaderMap(iStart))
        End If
    Next
    
    If result(0, I) = "" Then ReDim Preserve result(0 To 1, 0 To I - 1)
Exit_HttpHeader_Parse:
        HttpHeader_Parse = result
End Function


Private Function StripString(ByRef vString As String, ByRef vStart As String, Optional ByRef vEnd As String = vbNullString, Optional ByVal Start As Long = 1) As String
Dim iStart As Long
Dim iEnd As Long

iStart = InStr(Start, vString, vStart)
If iStart < 1 Then Exit Function

iStart = iStart + Len(vStart)
If vEnd = vbNullString Then
    StripString = Mid$(vString, iStart)
    Exit Function
End If

iEnd = InStr(iStart, vString, vEnd)
If iEnd < 1 Then Exit Function

StripString = Mid$(vString, iStart, iEnd - iStart)


End Function


Public Function ArrayUbound(ByRef vArr() As String) As Long
    On Error Resume Next
    
    ArrayUbound = -2
    ArrayUbound = UBound(vArr())
    
End Function


