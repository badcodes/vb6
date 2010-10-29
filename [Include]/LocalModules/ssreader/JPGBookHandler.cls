VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CJPGBookHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mHost As String
Private mLoginUrl As String
Private mPostUrl As String
Private mCookie As String
Private mHttp As IWinHttp
Private mInit As Boolean

Public Property Get Cookie() As String
    Cookie = mCookie
End Property

Public Function InitHandler( _
    ByVal vCookie As String, _
    Optional ByVal vHost As String = "http://pds.sslibrary.com/", _
    Optional ByVal vLoginUrl As String = "userlogon.jsp", _
    Optional ByVal vPostUrl As String = "loginhl.jsp" _
    )
    mCookie = vCookie
    mHost = LCase$(vHost)
    If Left$(mHost, 7) <> "http://" Then mHost = "http://" & mHost
    If Right$(mHost, 1) <> "/" Then mHost = mHost & "/"
    
    mLoginUrl = vLoginUrl
    If Left$(mLoginUrl, 1) = "/" Then mLoginUrl = Mid$(mLoginUrl, 2)
    If LCase$(Left$(mLoginUrl, 7)) <> "http://" Then mLoginUrl = mHost & mLoginUrl
    mPostUrl = vPostUrl
    If Left$(mPostUrl, 1) = "/" Then mPostUrl = Mid$(mPostUrl, 2)
    If LCase$(Left$(mPostUrl, 7)) <> "http://" Then mPostUrl = mHost & mPostUrl

    If mHttp Is Nothing Then Set mHttp = New CWinHttpSimple
    InitHandler = True
    mInit = True
End Function

Public Function CheckInitState() As Boolean
    CheckInitState = True
    If mInit Then Exit Function
    CheckInitState = False
    MsgBox "CJPGBookHandler not be initialized!", vbCritical
End Function


Public Sub UpdateBookInfo(ByVal vSSID As String, ByRef vBookInfo As CBookInfo)
End Sub

Public Function GetBookInfo(ByVal vSSID As String) As String()
End Function


'
'
'Public Function GetInfoUrl(ByRef vSSID As String, Optional ByRef vCookie As String = "", Optional vMainSite As String = cst_sslibrary_main, Optional ByRef vHttp As IWinHttp) As String
'    Dim pUrl As String
'    pUrl = vMainSite & "/" & cst_sslibrary_gojpg & vSSID
'    If vCookie = "" Then vCookie = mLastCookie
'    If vCookie = "" Then vCookie = "bkname=ssgpgdjy; UID=18523; state=1; lib=all; AID=161; tbExist=; exp=" & Chr$(34) & "2009-01-23 00:00:00.0" & Chr$(34) & "; allBooks=0; send=35534BF01F5A60E8BCE2EA46D1F624AF; userLogo=ssgpgdjy.gif; company=%u5e7f%u4e1c%u6559%u80b2%u5b66%u9662; marking=all; bnp=1.0; showDuxiu=0; showTopbooks=0; goWhere=1; JSESSIONID=75E5DC21B512ECEA267711A7A80360BA.tomcat2"
'
'    If vHttp Is Nothing Then Set vHttp = New CWinHttpSimple
'    vHttp.Init
'    vHttp.URL = pUrl
'    vHttp.Method = "GET"
'    'vHttp.SetTimeouts 20, 20, 20, 20
'    vHttp.Options(WinHttpRequestOption_EnableRedirects) = 0
'    vHttp.OpenConnect False
'    vHttp.SetRequestHeader "Cookie", vCookie
'    vHttp.Send
'
'    Do Until vHttp.IsFree
'        DoEvents
'    Loop
'
'    Dim pHeader As String
'    pHeader = vHttp.ResponseHeader
'
'    'If vHttp.Status <> 200 Then Exit Sub
'    pUrl = HttpHeaderGetField(vHttp.ResponseHeader, "Location")
'
'    If InStr(1, pUrl, "kid=", vbTextCompare) > 0 Then
'        vCookie = HttpHeaderMergeCookie(vCookie, HttpHeaderSetCookie(vHttp.ResponseHeader))
'        mLastCookie = vCookie
'        GetJpgBookInfoUrl = pUrl
'    End If
'    Debug.Print GetJpgBookInfoUrl
'End Function
'
'Public Function GetJpgBookInfo(ByVal vUrl As String, Optional ByRef vCookie As String = "", Optional ByRef vHttp As IWinHttp = Nothing) As String
'    If vCookie = "" Then vCookie = mLastCookie
'    If vCookie = "" Then vCookie = "bkname=ssgpgdjy; UID=18523; state=1; lib=all; AID=161; tbExist=; exp=" & Chr$(34) & "2009-01-23 00:00:00.0" & Chr$(34) & "; allBooks=0; send=35534BF01F5A60E8BCE2EA46D1F624AF; userLogo=ssgpgdjy.gif; company=%u5e7f%u4e1c%u6559%u80b2%u5b66%u9662; marking=all; bnp=1.0; showDuxiu=0; showTopbooks=0; goWhere=1; JSESSIONID=75E5DC21B512ECEA267711A7A80360BA.tomcat2"
'    If vUrl = "" Then Exit Function
'    If vHttp Is Nothing Then Set vHttp = New CWinHttpSimple
'    vHttp.Init
'    vHttp.URL = vUrl
'    vHttp.Method = "GET"
'    vHttp.Options(WinHttpRequestOption_EnableRedirects) = 0
'    vHttp.OpenConnect False
'    vHttp.SetRequestHeader "Cookie", vCookie
'    vHttp.Send
'
'    Do Until vHttp.IsFree
'        DoEvents
'    Loop
'
'    Dim ret As String
'    ret = Trim$(StrConv(vHttp.ResponseBody, vbUnicode))
'    If ret <> "" Then
'       vCookie = HttpHeaderMergeCookie(vCookie, HttpHeaderGetField(vHttp.ResponseHeader, "Set-Cookie"))
'        mLastCookie = vCookie
'       GetJpgBookInfo = ret
'    End If
'
'
'End Function
'
'Public Function SSLIB_GetReadBookURL(vShowBookUrl As String, Optional vSite As String = "http://pds.sslibrary.com", Optional vCookie As String) As String
'    If vCookie = "" Then vCookie = Clipboard.GetText
'    If vShowBookUrl = "" Then vShowBookUrl = "showbook.do?dxNumber=10108311&d=99653DE6D2EBC329B6D3F914FECB92E4&fenleiID=0I20100140&nettype=wangtong&username=ssgpgdjy"
'    Dim pHttp As IWinHttp
'    Set pHttp = New CWinHttpSimple
'    With pHttp
'        .Init
'        .URL = BuildPath(vSite, vShowBookUrl, lnpsUnix)
'        .Method = "GET"
'        .OpenConnect False
'        .SetRequestHeader "Cookie", vCookie
'        .Send
'    End With
'    Do Until pHttp.IsFree
'        DoEvents
'    Loop
'    Debug.Print UnescapeUrl(pHttp.ResponseText)
'End Function
'
'Public Function SSLibraryLogin(ByVal vUsername As String, ByVal vPassWord As String, ByVal vLoginUrl As String, ByVal vPostUrl As String, Optional ByVal vDisplayFailData As Boolean = False) As String
'
'    If vUsername = "" Then vUsername = "ssgpgdjy"
'    If vPassWord = "" Then vPassWord = "shngdjy"
'    If vLoginUrl = "" Then vLoginUrl = "http://pds.sslibrary.com/userlogon.jsp"
'    If vPostUrl = "" Then vPostUrl = "http://pds.sslibrary.com/loginhl.jsp"
'
'    Dim Http As IWinHttp
'    Set Http = New CWinHttpSimple
'    Http.Init
'    Http.URL = vLoginUrl
'    Http.Method = "GET"
'    Http.Options(WinHttpRequestOption_EnableRedirects) = 0
'    Http.OpenConnect False
'    Http.Send
'    Do Until Http.IsFree
'        DoEvents
'    Loop
'    Dim ret As String
'    ret = HttpHeaderGetField(Http.ResponseHeader, "Set-Cookie")
'
'    Dim postBody As String
'    postBody = "send=true&UserName=" & vUsername & "&PassWord=" & vPassWord & "&rd=0&Submit3322=%B5%C7%C2%BC&backurl="
''    Dim postBytes() As Byte
''    postBytes = StrConv(postBody, vbFromUnicode)
''    postBody = postBytes
'    Http.Init
'    Http.URL = vPostUrl
'    Http.Method = "POST"
'    Http.Options(WinHttpRequestOption_EnableRedirects) = 0
'    Http.OpenConnect False
'    Http.SetRequestHeader "Cookie", ret
'    Http.SetRequestHeader "Referer", vLoginUrl
'    Http.SetRequestHeader "Content-Length", Len(postBody)
'    Http.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
'
'    Http.Send postBody
'    Do Until Http.IsFree
'    DoEvents
'    Loop
'    Dim Header As String
'    Header = Http.ResponseHeader
'    If InStr(1, Header, "UID=", vbTextCompare) Then
'        SSLibraryLogin = HttpHeaderMergeCookie(ret, HttpHeaderSetCookie(Header))
'    ElseIf vDisplayFailData Then
'        MsgBox StrConv(Http.ResponseBody, vbUnicode), vbCritical, vPostUrl
'    End If
'End Function
'
'Public Function GetJpgBookPageUrl(ByVal vRootUrl As String, ByVal vFilename As String, Optional ByRef vCookie As String, Optional vHttp As IWinHttp = Nothing, Optional vSizeType As Integer = 1) As String
'    If vCookie = "" Then vCookie = mLastCookie
'    If vCookie = "" Then vCookie = "bkname=ssgpgdjy; UID=18523; state=1; lib=all; AID=161; tbExist=; exp=" & Chr$(34) & "2009-01-23 00:00:00.0" & Chr$(34) & "; allBooks=0; send=35534BF01F5A60E8BCE2EA46D1F624AF; userLogo=ssgpgdjy.gif; company=%u5e7f%u4e1c%u6559%u80b2%u5b66%u9662; marking=all; bnp=1.0; showDuxiu=0; showTopbooks=0; goWhere=1; JSESSIONID=75E5DC21B512ECEA267711A7A80360BA.tomcat2"
'    If vHttp Is Nothing Then Set vHttp = New CWinHttpSimple
'
'    Dim vUrl As String
'    vUrl = vRootUrl & vFilename
'
'    vUrl = Replace$(vUrl, ":", "%3A")
'    vUrl = Replace$(vUrl, "?", "%3F")
'    vUrl = Replace$(vUrl, "&", "%26")
'    vUrl = Replace$(vUrl, "=", "%3D")
'    vUrl = Replace$(vUrl, "!", "%21")
'    If vHttp Is Nothing Then Set vHttp = New CWinHttpSimple
'    With vHttp
'        .Init
'        .URL = "http://img.sslibrary.com/jpgssrurl.jsp?h=" & vUrl
'        .Method = "POST"
'        .OpenConnect False
'        .SetRequestHeader "Cookie", vCookie
'        .Send
'    End With
'    Do Until vHttp.IsFree
'        DoEvents
'    Loop
'    Dim pRet As String
'    pRet = StrConv(vHttp.ResponseBody, vbUnicode)
'    If pRet <> "" Then
'        If vSizeType < 0 Then vSizeType = 0
'        vCookie = HttpHeaderMergeCookie(vCookie, HttpHeaderGetField(vHttp.ResponseHeader, "Set-Cookie"))
'        mLastCookie = vCookie
'        GetJpgBookPageUrl = vRootUrl & vFilename & "&a=" & pRet & "&uf=ssr&zoom=" & vSizeType
'    End If
'
'End Function
'
'Public Function GetJpgBookParamA(ByRef vCookie As String, ByVal vUrl As String) As String
'
'    vUrl = Replace$(vUrl, ":", "%3A")
'    vUrl = Replace$(vUrl, "?", "%3F")
'    vUrl = Replace$(vUrl, "&", "%26")
'    vUrl = Replace$(vUrl, "=", "%3D")
'    vUrl = Replace$(vUrl, "!", "%21")
'
'    Dim pHttp As IWinHttp
'    Set pHttp = New CWinHttpSimple
'    pHttp.Init
'    pHttp.URL = "http://img.sslibrary.com/jpgssrurl.jsp?h=" & vUrl
'    pHttp.Method = "POST"
'    pHttp.OpenConnect False
'    pHttp.SetRequestHeader "Cookie", vCookie
'    pHttp.Send
'    Do Until pHttp.IsFree
'        DoEvents
'    Loop
'    vCookie = HttpHeaderMergeCookie(vCookie, HttpHeaderGetField(pHttp.ResponseHeader, "Set-Cookie"))
'    GetJpgBookParamA = StrConv(pHttp.ResponseBody, vbUnicode)
'End Function


    

