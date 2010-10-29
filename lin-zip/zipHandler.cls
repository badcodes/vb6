VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsZipHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

Implements olelib2.IInternetProtocol
Implements olelib.IInternetProtocolInfo

Private m_oSink As olelib2.IInternetProtocolSink  ' Callback object used to informthe progress

Private m_lFile As Long ' File handle
Private m_bDownloadDone As Boolean        ' Flag used to know if all the data has been readed
Private m_sFile As String
Private InvaildPassword As Boolean
Private PASSWORD As String
Private WithEvents lUnzip As cUnzip
Attribute lUnzip.VB_VarHelpID = -1
Private sTempFolder As String

Private Sub Class_Initialize()
'<CSCustomCode> 1
#If DebugNow = 1 Then
MDebug.DebugFPrint "Start linzip.clsZipHandler.Class_Initialize"
#End If
'</CSCustomCode> 1
    Dim oPH As IInternetProtocol
    Set lUnzip = New cUnzip
    Set oPH = Me
    ReplaceVTableEntry ObjPtr(oPH), 10, AddressOf modReturnfunc.IInternetProtocol_Read
    ReplaceVTableEntry ObjPtr(oPH), 4, AddressOf modReturnfunc.IInternetProtocol_Start
    
    Dim oPH2 As IInternetProtocolInfo
    Set oPH2 = Me
    ReplaceVTableEntry ObjPtr(oPH2), 4, AddressOf modReturnfunc.IInternetProtocolInfo_ParseUrl
    ReplaceVTableEntry ObjPtr(oPH2), 5, AddressOf modReturnfunc.IInternetProtocolInfo_CombineUrl
    ReplaceVTableEntry ObjPtr(oPH2), 7, AddressOf modReturnfunc.IInternetProtocolInfo_QueryInfo

    sTempFolder = Environ$("temp")
    modReturnfunc.instanceCount = modReturnfunc.instanceCount + 1
    
    
    
End Sub
Private Sub freeUsedObject()
'<CSCustomCode> 1
#If DebugNow = 1 Then
MDebug.DebugFPrint "Start linzip.clsZipHandler.freeUsedObject"
#End If
'</CSCustomCode> 1
    Set lUnzip = Nothing
    Set m_oSink = Nothing
End Sub
Private Sub Class_Terminate()
'<CSCustomCode> 1
#If DebugNow = 1 Then
MDebug.DebugFPrint "Start linzip.clsZipHandler.Class_Terminate"
#End If
'</CSCustomCode> 1
    freeUsedObject
    modReturnfunc.instanceCount = modReturnfunc.instanceCount - 1
    If modReturnfunc.instanceCount < 0 Then modReturnfunc.instanceCount = 0
    On Error Resume Next
    If modReturnfunc.instanceCount = 0 Then RmDir sTempFolder
End Sub
Private Sub IInternetProtocol_Abort(ByVal hrReason As Long, ByVal dwOptions As Long)
'<CSCustomCode> 1
#If DebugNow = 1 Then
MDebug.DebugFPrint "Start linzip.clsZipHandler.IInternetProtocol_Abort"
#End If
'</CSCustomCode> 1
    If m_lFile <> 0 Then Close m_lFile
    Kill m_sFile
    m_lFile = 0
    freeUsedObject
End Sub
Private Sub IInternetProtocol_Continue(pProtocolData As olelib.PROTOCOLDATA)
'<CSCustomCode> 1
#If DebugNow = 1 Then
MDebug.DebugFPrint "Start linzip.clsZipHandler.IInternetProtocol_Continue"
#End If
'</CSCustomCode> 1
End Sub
Private Sub IInternetProtocol_LockRequest(ByVal dwOptions As Long)
'<CSCustomCode> 1
#If DebugNow = 1 Then
MDebug.DebugFPrint "Start linzip.clsZipHandler.IInternetProtocol_LockRequest"
#End If
'</CSCustomCode> 1
End Sub
Private Sub IInternetProtocol_Read(ByVal pv As Long, ByVal cb As Long, pcbRead As Long)
'<CSCustomCode> 1
#If DebugNow = 1 Then
MDebug.DebugFPrint "Start linzip.clsZipHandler.IInternetProtocol_Read"
#End If
'</CSCustomCode> 1
End Sub
Private Sub IInternetProtocol_Resume()
'<CSCustomCode> 1
#If DebugNow = 1 Then
MDebug.DebugFPrint "Start linzip.clsZipHandler.IInternetProtocol_Resume"
#End If
'</CSCustomCode> 1
End Sub
Private Sub IInternetProtocol_Seek(ByVal dlibMove As Currency, ByVal dwOrigin As Long, plibNewPosition As Currency)
'<CSCustomCode> 1
#If DebugNow = 1 Then
MDebug.DebugFPrint "Start linzip.clsZipHandler.IInternetProtocol_Seek"
#End If
'</CSCustomCode> 1
End Sub
Private Sub IInternetProtocol_Start(ByVal szUrl As Long, ByVal pOIProtSink As olelib2.IInternetProtocolSink, ByVal pOIBindInfo As olelib.IInternetBindInfo, ByVal grfPI As olelib.PI_FLAGS, dwReserved As olelib.PROTOCOLFILTERDATA)
'<CSCustomCode> 1
#If DebugNow = 1 Then
MDebug.DebugFPrint "Start linzip.clsZipHandler.IInternetProtocol_Start"
#End If
'</CSCustomCode> 1
End Sub
Private Sub IInternetProtocol_Suspend()
'<CSCustomCode> 1
#If DebugNow = 1 Then
MDebug.DebugFPrint "Start linzip.clsZipHandler.IInternetProtocol_Suspend"
#End If
'</CSCustomCode> 1
End Sub
Private Sub IInternetProtocol_Terminate(ByVal dwOptions As Long)
'<CSCustomCode> 1
#If DebugNow = 1 Then
MDebug.DebugFPrint "Start linzip.clsZipHandler.IInternetProtocol_Terminate"
#End If
'</CSCustomCode> 1
    freeUsedObject
End Sub
Private Sub IInternetProtocol_UnlockRequest()
'<CSCustomCode> 1
#If DebugNow = 1 Then
MDebug.DebugFPrint "Start linzip.clsZipHandler.IInternetProtocol_UnlockRequest"
#End If
'</CSCustomCode> 1
    freeUsedObject
End Sub

Private Sub IInternetProtocolInfo_CombineUrl(ByVal pwzBaseUrl As Long, ByVal pwzRelativeUrl As Long, ByVal dwCombineFlags As Long, ByVal pwzResult As Long, ByVal cchResult As Long, pcchResult As Long, ByVal dwReserved As Long)
'<CSCustomCode> 1
#If DebugNow = 1 Then
MDebug.DebugFPrint "Start linzip.clsZipHandler.IInternetProtocolInfo_CombineUrl"
#End If
'</CSCustomCode> 1

End Sub

Private Sub IInternetProtocolInfo_CompareUrl(ByVal pwzUrl1 As Long, ByVal pwzUrl2 As Long, ByVal dwCompareFlags As Long)
'<CSCustomCode> 1
#If DebugNow = 1 Then
MDebug.DebugFPrint "Start linzip.clsZipHandler.IInternetProtocolInfo_CompareUrl"
#End If
'</CSCustomCode> 1

End Sub

Private Sub IInternetProtocolInfo_ParseUrl(ByVal pwzUrl As Long, ByVal PARSEACTION As olelib.PARSEACTION, ByVal dwParseFlags As Long, ByVal pwzResult As Long, ByVal cchResult As Long, pcchResult As Long, ByVal dwReserved As Long)
'<CSCustomCode> 1
#If DebugNow = 1 Then
MDebug.DebugFPrint "Start linzip.clsZipHandler.IInternetProtocolInfo_ParseUrl"
#End If
'</CSCustomCode> 1

End Sub

Private Sub IInternetProtocolInfo_QueryInfo(ByVal pwzUrl As Long, ByVal OueryOption As olelib.QUERYOPTION, ByVal dwQueryFlags As Long, ByVal pBuffer As Long, ByVal cbBuffer As Long, pcbBuf As Long, ByVal dwReserved As Long)
'<CSCustomCode> 1
#If DebugNow = 1 Then
MDebug.DebugFPrint "Start linzip.clsZipHandler.IInternetProtocolInfo_QueryInfo"
#End If
'</CSCustomCode> 1

End Sub

Private Sub lUnzip_PasswordRequest(sPassword As String, ByVal sName As String, bCancel As Boolean)
'<CSCustomCode> 1
#If DebugNow = 1 Then
MDebug.DebugFPrint "Start linzip.clsZipHandler.lUnzip_PasswordRequest"
#End If
'</CSCustomCode> 1
    bCancel = False
    Static lastName As String
    If InvaildPassword = False And PASSWORD <> "" Then
        sPassword = PASSWORD
        If sName = lastName Then
            InvaildPassword = True
        Else
            lastName = sName
        End If
    Else
        sPassword = InputBox(lUnzip.ZipFile & vbCrLf & sName & " Request For Password", "Password", "")
        If sPassword <> "" Then
            InvaildPassword = False
            PASSWORD = sPassword
        Else
            bCancel = True
        End If
    End If
End Sub
Private Function myXUnzip(ByVal sZipfilename As String, ByVal sFilesToProcess As String, ByVal sUnzipTo As String) As unzReturnCode
'<CSCustomCode> 1
#If DebugNow = 1 Then
MDebug.DebugFPrint "Start linzip.clsZipHandler.myXUnzip"
#End If
'</CSCustomCode> 1
    With lUnzip
        .CaseSensitiveFileNames = False
        .PromptToOverwrite = False
        .OverwriteExisting = True
        .UseFolderNames = False
        .ZipFile = sZipfilename
        .FileToProcess = lusezipdll.CleanZipFilename(sFilesToProcess)
        .UnzipFolder = sUnzipTo
    End With
    myXUnzip = lUnzip.unzip
End Function
Friend Function Read(ByVal pv As Long, ByVal cb As Long, pcbRead As Long) As Long
'<CSCustomCode> 1
#If DebugNow = 1 Then
MDebug.DebugFPrint "Start linzip.clsZipHandler.Read"
#End If
'</CSCustomCode> 1
    Dim abBuffer() As Byte
    Dim curPos As Long
    Dim lastPos As Long
    Dim cbReal As Long
    On Error GoTo hell
    If Not m_bDownloadDone Then
        If m_lFile = 0 Then
            pcbRead = 0
            Read = S_FALSE
            Exit Function
        End If
        curPos = Loc(m_lFile)
        lastPos = LOF(m_lFile)
        If curPos >= lastPos Then
            pcbRead = 0
            Read = S_FALSE
            Exit Function
        End If
        cbReal = cb
        If curPos + cb > lastPos Then cbReal = lastPos - curPos
        ReDim abBuffer(0 To cbReal - 1)
        Get m_lFile, , abBuffer
        MoveMemory ByVal pv, abBuffer(0), cbReal
        pcbRead = cbReal
        Read = S_OK
        If Loc(m_lFile) < LOF(m_lFile) Then
            m_oSink.ReportData BSCF_INTERMEDIATEDATANOTIFICATION, Loc(m_lFile), LOF(m_lFile)
        Else
            m_oSink.ReportData BSCF_LASTDATANOTIFICATION, 100, 100
            m_oSink.ReportResult S_OK, 0, vbNullString
            m_bDownloadDone = True
            Close m_lFile
            Kill m_sFile
            Read = S_FALSE
        End If
    Else
        pcbRead = 0
        Read = S_FALSE
    End If
    Exit Function
hell:
End Function
Friend Function Start(ByVal szUrl As Long, ByVal pOIProtSink As olelib2.IInternetProtocolSink, ByVal pOIBindInfo As olelib.IInternetBindInfo, ByVal grfPI As olelib.PI_FLAGS, dwReserved As olelib.PROTOCOLFILTERDATA) As Long
'<CSCustomCode> 1
#If DebugNow = 1 Then
MDebug.DebugFPrint "Start linzip.clsZipHandler.Start"
#End If
'</CSCustomCode> 1
    Dim sRealUrl As String
    Dim sUrl As zipUrl
    Dim sMimeType As String
    
    sMimeType = "text/html"
    Start = S_OK
    m_bDownloadDone = False
    Set m_oSink = pOIProtSink
    
    sRealUrl = SysAllocString(szUrl)
    sUrl = zipProtocol_ParseURL(sRealUrl)
    If sUrl.sZipName = "" Or sUrl.sHtmlPath = "" Then
        Start = INET_E_INVALID_URL
        Exit Function
    End If
    Dim sRealPath As String
    If FakeToReal(sUrl.sHtmlPath, sRealPath) Then
        m_sFile = BuildPath(BuildPath(Environ$("temp"), zipTempName), sRealPath)
    Else
        m_sFile = BuildPath(sTempFolder, GetFileName(sUrl.sHtmlPath))
        myXUnzip sUrl.sZipName, sUrl.sHtmlPath, sTempFolder
    End If
    If PathExists(m_sFile) = False Then
        Start = INET_E_DATA_NOT_AVAILABLE
        Exit Function
    End If
    sMimeType = MimeType(m_sFile)
    m_lFile = FreeFile
    Open m_sFile For Binary As m_lFile
ReportProgress:
    m_oSink.ReportProgress BINDSTATUS_MIMETYPEAVAILABLE, sMimeType
    m_oSink.ReportProgress BINDSTATUS_VERIFIEDMIMETYPEAVAILABLE, sMimeType
    m_oSink.ReportData BSCF_FIRSTDATANOTIFICATION, 0, LOF(m_lFile)
End Function

Friend Function ParseUrl(ByVal pwzUrl As Long, ByVal PARSEACTION As olelib.PARSEACTION, ByVal dwParseFlags As Long, ByVal pwzResult As Long, ByVal cchResult As Long, pcchResult As Long, ByVal dwReserved As Long) As HRESULTS
'<CSCustomCode> 1
#If DebugNow = 1 Then
MDebug.DebugFPrint "Start linzip.clsZipHandler.ParseUrl"
#End If
'</CSCustomCode> 1

    Dim sUrl As String
    Dim sResult As String
    Dim parseType As String
    
    sUrl = olelib.SysAllocString(pwzUrl)
    sResult = ""
    ParseUrl = E_NOTIMPL
    
    Select Case PARSEACTION
    Case PARSE_ANCHOR
        sResult = linvblib.RightLeft(sUrl, "#", vbBinaryCompare, ReturnEmptyStr)
        parseType = "PARSE_ANCHOR"
    Case PARSE_CANONICALIZE
        sResult = modReturnfunc.CanonicalizeUrl(sUrl, dwParseFlags)
        parseType = "PARSE_CANONICALIZE"
    Case PARSE_DECODE
        sResult = sUrl
        parseType = "PARSE_DECODE"
    Case PARSE_DOCUMENT
        parseType = "PARSE_DOCUMENT"
    Case PARSE_DOMAIN
        sResult = "LocalHost"
        'sResult = LiNVBLib.LeftRange(sUrl, "//", "/", vbBinaryCompare, ReturnEmptyStr)
        parseType = "PARSE_DOMAIN"
    Case PARSE_ENCODE
        sResult = sUrl
        parseType = "PARSE_ENCODE"
    Case PARSE_ESCAPE
        sResult = linvblib.EscapeUrl(sUrl)
        parseType = "PARSE_ESCAPE"
    Case PARSE_FRIENDLY
        sResult = linvblib.EscapeUrl(sUrl)
        parseType = "PARSE_FRIENDLY"
    Case PARSE_LOCATION
        sResult = "" 'RightRight(sUrl, "/", vbBinaryCompare, ReturnOriginalStr)
        parseType = "PARSE_LOCATION"
    Case PARSE_MIME
        parseType = "PARSE_MIME"
    Case PARSE_PATH_FROM_URL
        sResult = "" 'linvblib.LeftRight(sUrl, "://", vbBinaryCompare, ReturnOriginalStr)
        sResult = "" 'linvblib.UnescapeUrl(sResult)
        parseType = "PARSE_PATH_FROM_URL"
    Case PARSE_ROOTDOCUMENT
        parseType = "PARSE_ROOTDOCUMENT"
    Case PARSE_SCHEMA
        sResult = linvblib.LeftLeft(sUrl, ":", vbBinaryCompare, ReturnEmptyStr)
        parseType = "PARSE_SCHEMA"
    Case PARSE_SECURITY_DOMAIN
        sResult = linvblib.LeftRange(sUrl, ":", zipSep, vbBinaryCompare, ReturnEmptyStr)
         parseType = "PARSE_SECURITY_DOMAIN"
    Case PARSE_SECURITY_URL
        sResult = linvblib.LeftLeft(sUrl, zipSep, vbBinaryCompare, ReturnEmptyStr)
        parseType = "PARSE_SECURITY_URL"
    Case PARSE_SERVER
        sResult = linvblib.LeftRange(sUrl, ":", zipSep, vbBinaryCompare, ReturnEmptyStr)
        parseType = "PARSE_SERVER"
    Case PARSE_SITE
        sResult = linvblib.LeftLeft(sUrl, zipSep, vbBinaryCompare, ReturnEmptyStr)
        parseType = "PARSE_SITE"
    Case PARSE_UNESCAPE
        sResult = linvblib.UnescapeUrl(sUrl)
        parseType = "PARSE_UNESCAPE"
    Case PARSE_URL_FROM_PATH
        sResult = zipProtocolHead & sUrl
        parseType = "PARSE_URL_FROM_PATH"
    End Select
    'debug
    'MDebug.DebugFPrint "ParseUrl of " & parseType & "(" & PARSEACTION & ")" & " of " & sUrl & vbCrLf & _
                     "Parse result : " & sResult
 
    If sResult <> "" Then
        pcchResult = Len(sResult) * 2
        olelib.MoveMemory ByVal pwzResult, ByVal sResult, pcchResult + 2
        ParseUrl = S_OK '  INET_E_DEFAULT_ACTION ' = S_OK 'S_FALSE
    End If

End Function

Friend Function QueryInfo(ByVal pwzUrl As Long, ByVal QUERYOPTION As olelib.QUERYOPTION, ByVal dwQueryFlags As Long, ByVal pBuffer As Long, ByVal cbBuffer As Long, pcbBuf As Long, ByVal dwReserved As Long) As HRESULTS
'<CSCustomCode> 1
#If DebugNow = 1 Then
MDebug.DebugFPrint "Start linzip.clsZipHandler.QueryInfo"
#End If
'</CSCustomCode> 1
'QueryInfo = INET_E_DEFAULT_ACTION: Exit Function


    Dim sReturn As String
    Dim sUrl As String
    Dim sQueryType As String
    
    sUrl = olelib.SysAllocString(pwzUrl)
    sReturn = ""
    
    QueryInfo = INET_E_DEFAULT_ACTION
    
    Select Case QUERYOPTION
    Case QUERY_EXPIRATION_DATE
        sQueryType = "QUERY_EXPIRATION_DATE"
    Case QUERY_TIME_OF_LAST_CHANGE
        sQueryType = "QUERY_TIME_OF_LAST_CHANGE"
    Case QUERY_CONTENT_ENCODING
        sQueryType = "QUERY_CONTENT_ENCODING"
    Case QUERY_CONTENT_TYPE
        sReturn = MimeType(sUrl)
        sQueryType = "QUERY_CONTENT_TYPE"
    Case QUERY_REFRESH
        sQueryType = "QUERY_REFRESH"
    Case QUERY_RECOMBINE
        sQueryType = "QUERY_RECOMBINE"
    Case QUERY_CAN_NAVIGATE
            pcbBuf = 4
            olelib.MoveMemory ByVal pBuffer, 1, pcbBuf
            QueryInfo = S_OK
            sQueryType = "QUERY_CAN_NAVIGATE"
    Case QUERY_USES_NETWORK
            pcbBuf = 4
            olelib.MoveMemory ByVal pBuffer, 0, pcbBuf
            QueryInfo = S_OK
            sQueryType = "QUERY_USES_NETWORK"
    Case QUERY_IS_CACHED
             pcbBuf = 4
        olelib.MoveMemory ByVal pBuffer, 0, pcbBuf
        QueryInfo = S_OK
        sQueryType = "QUERY_IS_CACHED"
    Case QUERY_IS_INSTALLEDENTRY
        sQueryType = "QUERY_IS_INSTALLEDENTRY"
    Case QUERY_IS_CACHED_OR_MAPPED
        pcbBuf = 4
        olelib.MoveMemory ByVal pBuffer, 0, pcbBuf
        QueryInfo = S_OK
        sQueryType = "QUERY_IS_CACHED_OR_MAPPED"
    Case QUERY_USES_CACHE
        pcbBuf = 4
        olelib.MoveMemory ByVal pBuffer, 0, pcbBuf
        QueryInfo = S_OK
        sQueryType = "QUERY_USES_CACHE"
    Case QUERY_IS_SECURE
        pcbBuf = 4
        olelib.MoveMemory ByVal pBuffer, 1, pcbBuf
        QueryInfo = S_OK
        sQueryType = "QUERY_IS_SECURE"
    Case QUERY_IS_SAFE
        pcbBuf = 4
        olelib.MoveMemory ByVal pBuffer, 1, pcbBuf
        QueryInfo = S_OK
        sQueryType = "QUERY_IS_SAFE"
    End Select
    
    If sReturn <> "" Then
        pcbBuf = LenB(sReturn) + 2
        olelib.MoveMemory ByVal pBuffer, ByVal sReturn, pcbBuf
    End If
    
   'MDebug.DebugFPrint "Queryinfo of " & sQueryType & "(" & CStr(QUERYOPTION) & ")" & " For " & olelib.SysAllocString(pwzUrl) & vbCrLf & _
                "Result :" & sReturn
 
 End Function
Friend Function CombineUrl(ByVal pwzBaseUrl As Long, ByVal pwzRelativeUrl As Long, ByVal dwCombineFlags As Long, ByVal pwzResult As Long, ByVal cchResult As Long, pcchResult As Long, ByVal dwReserved As Long) As HRESULTS
'<CSCustomCode> 1
#If DebugNow = 1 Then
MDebug.DebugFPrint "Start linzip.clsZipHandler.CombineUrl"
#End If
'</CSCustomCode> 1

    Dim sBaseUrl As String
    Dim sRelativeUrl As String
    Dim sResult As String

    sBaseUrl = olelib.SysAllocString(pwzBaseUrl)
    sBaseUrl = linvblib.RightLeft(sBaseUrl, "/", vbBinaryCompare, ReturnOriginalStr) & "/"
    sRelativeUrl = olelib.SysAllocString(pwzRelativeUrl)

    If InStr(sRelativeUrl, ":") > 0 Then
        sResult = sRelativeUrl
    Else
        sResult = sBaseUrl & sRelativeUrl
    End If

    'sResult = modReturnfunc.CanonicalizeUrl(sResult, dwCombineFlags)
    pcchResult = Len(sResult) * 2
    If pcchResult <= cchResult Then
        olelib.MoveMemory ByVal pwzResult, ByVal sResult, pcchResult + 2
        CombineUrl = S_OK
    Else
        pcchResult = pcchResult + 1
        CombineUrl = S_FALSE
    End If
    
  'MDebug.DebugFPrint "BaseUrl : " & sBaseUrl & vbCrLf & _
                "RelativeUrl :" & sRelativeUrl & vbCrLf & _
                "CombineFlag :" & dwCombineFlags & vbCrLf & _
                "Url Return :" & sResult



End Function
