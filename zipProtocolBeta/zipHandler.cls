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
'*********************************************************************************************
'
' Protocol Handlers
'
' Protocol Handler class
'
'*********************************************************************************************
'
' Author: Eduardo A. Morcillo
' E-Mail: e_morcillo@yahoo.com
' Web Page: http://www.domaindlx.com/e_morcillo
'
' Distribution: You can freely use this code in your own applications but you
'               can't publish this code in a web site, online service, or any
'               other media, without my express permission.
'
' Usage: at your own risk.
'
' Tested with:
'              VB6 SP5
'              Windows XP Pro SP1
'              IE 6.0
'
' History:
'           08/19/2001 - This code was released
'
'*********************************************************************************************
Option Explicit

' Implement IInternetProtocol
Implements olelib2.IInternetProtocol
Implements olelib.IInternetProtocolInfo

#Const TESTBETA = 1
' Private members
Private m_oSink As IInternetProtocolSink  ' Callback object used to informthe progress
Private m_oBindInfo As olelib.IInternetBindInfo
Private m_lFile As Long ' File handle
Private m_bDownloadDone As Boolean        ' Flag used to know if all the data has been readed
Private m_sFile As String
Private InvaildPassword As Boolean
Private PASSWORD As String
Private WithEvents lUnzip As cUnzip
Attribute lUnzip.VB_VarHelpID = -1
Private sTempFolder As String
'Private WithEvents lXZ As XceedZip
Private Enum Errors
   errCannotOpenConnection
   errCannotOpenRecordset
   errInvalidURL
   errCannotOpenFile
End Enum

'
'
' ParseURL
'
' Parses the URL. Returns False if the URL can't be parsed.
'

'
' WriteError
'
' Writes an error file that will be returned by this handler.
'

Private Sub Class_Initialize()

    'DoEvents

    Dim oPH As IInternetProtocol
    Set lUnzip = New cUnzip
    'Set lXZ = New XceedZip
    ' Get the IInternetProtocol interface
    Set oPH = Me
     'Replace the Read method
    ReplaceVTableEntry ObjPtr(oPH), 10, AddressOf modReturnfunc.IInternetProtocol_Read
    ReplaceVTableEntry ObjPtr(oPH), 4, AddressOf modReturnfunc.IInternetProtocol_Start


    Dim oPH2 As IInternetProtocolInfo
    Set oPH2 = Me
    ReplaceVTableEntry ObjPtr(oPH2), 4, AddressOf modReturnfunc.IInternetProtocolInfo_ParseUrl
    ReplaceVTableEntry ObjPtr(oPH2), 5, AddressOf modReturnfunc.IInternetProtocolInfo_CombineUrl
    ReplaceVTableEntry ObjPtr(oPH2), 7, AddressOf modReturnfunc.IInternetProtocolInfo_QueryInfo




    sTempFolder = Environ$("temp")

   
End Sub

Private Sub Class_Terminate()

    'zipProtocolCount = zipProtocolCount - 1
    Set lUnzip = Nothing
    'Dim fso As New FileSystemObject
    'On Error Resume Next
    'If zipProtocolCount <= 0 Then fso.DeleteFolder sTempFolder, True

End Sub

Private Sub IInternetBindInfo_GetBindInfo(grfBINDF As Long, pbindinfo As olelib.BINDINFO)

End Sub

Private Sub IInternetBindInfo_GetBindString(ByVal ulStringType As Long, ppwzStr As Long, ByVal cEl As Long, pcElFetched As Long)

End Sub

Private Sub IInternetProtocol_Abort(ByVal hrReason As Long, ByVal dwOptions As Long)
       
    'msgbox "Abort " & m_sFile
    ' Delete the temporary file

    If m_lFile <> 0 Then Close m_lFile

    Kill m_sFile
   
    ' Reset all variables

    m_lFile = 0
   
    ' Release the objects
    Set m_oSink = Nothing
   
End Sub

Private Sub IInternetProtocol_Continue(pProtocolData As olelib.PROTOCOLDATA)

    'msgbox "Continue " & m_sFile
    Err.Raise E_NOTIMPL

End Sub

Private Sub IInternetProtocol_LockRequest(ByVal dwOptions As Long)

    'msgbox "lock " & m_sFile

End Sub

Private Sub IInternetProtocol_Read(ByVal pv As Long, ByVal cb As Long, pcbRead As Long)

End Sub

Private Sub IInternetProtocol_Resume()

    'msgbox "resume " & m_sFile
    Err.Raise E_NOTIMPL

End Sub

Private Sub IInternetProtocol_Seek(ByVal dlibMove As Currency, ByVal dwOrigin As Long, plibNewPosition As Currency)

    'msgbox "seek " & m_sFile
    Err.Raise E_NOTIMPL

End Sub



Private Sub IInternetProtocol_Start(ByVal szURL As Long, ByVal pOIProtSink As olelib2.IInternetProtocolSink, ByVal pOIBindInfo As olelib.IInternetBindInfo, ByVal grfPI As olelib.PI_FLAGS, dwReserved As olelib.PROTOCOLFILTERDATA)

End Sub

Private Sub IInternetProtocol_Suspend()

    'msgbox "suspend " & m_sFile
    Err.Raise E_NOTIMPL

End Sub

Private Sub IInternetProtocol_Terminate(ByVal dwOptions As Long)

    'MsgBox "terminate " & m_sFile
    ' Delete the temporary file
    Set m_oSink = Nothing
   
End Sub

Private Sub IInternetProtocol_UnlockRequest()

    'msgbox "Unlock " & m_sFile
    Set m_oSink = Nothing

End Sub

'
Private Sub IInternetProtocolInfo_CombineUrl(ByVal pwzBaseUrl As Long, ByVal pwzRelativeUrl As Long, ByVal dwCombineFlags As Long, ByVal pwzResult As Long, ByVal cchResult As Long, pcchResult As Long, ByVal dwReserved As Long)

End Sub

Private Sub IInternetProtocolInfo_CompareUrl(ByVal pwzUrl1 As Long, ByVal pwzUrl2 As Long, ByVal dwCompareFlags As Long)

End Sub

Private Sub IInternetProtocolInfo_ParseUrl(ByVal pwzUrl As Long, ByVal PARSEACTION As olelib.PARSEACTION, ByVal dwParseFlags As Long, ByVal pwzResult As Long, ByVal cchResult As Long, pcchResult As Long, ByVal dwReserved As Long)

End Sub

Private Sub IInternetProtocolInfo_QueryInfo(ByVal pwzUrl As Long, ByVal OueryOption As olelib.QueryOption, ByVal dwQueryFlags As Long, ByVal pBuffer As Long, ByVal cbBuffer As Long, pcbBuf As Long, ByVal dwReserved As Long)
MsgBox "clsZipHandler::IInternetProtocolInfo_QueryInfo Begin"
End Sub

Private Sub IInternetProtocolRoot_Abort(ByVal hrReason As Long, ByVal dwOptions As Long)

End Sub

Private Sub IInternetProtocolRoot_Continue(pProtocolData As olelib.PROTOCOLDATA)

End Sub

Private Sub IInternetProtocolRoot_Resume()

End Sub

Private Sub IInternetProtocolRoot_Start(ByVal szURL As Long, ByVal pOIProtSink As olelib.IInternetProtocolSink, ByVal pOIBindInfo As olelib.IInternetBindInfo, ByVal grfPI As olelib.PI_FLAGS, dwReserved As olelib.PROTOCOLFILTERDATA)

End Sub

Private Sub IInternetProtocolRoot_Suspend()

End Sub

Private Sub IInternetProtocolRoot_Terminate(ByVal dwOptions As Long)

End Sub

Private Sub lUnzip_PasswordRequest(sPassword As String, ByVal sName As String, bCancel As Boolean)

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

Friend Function CombineUrl(ByVal pwzBaseUrl As Long, ByVal pwzRelativeUrl As Long, ByVal dwCombineFlags As Long, ByVal pwzResult As Long, ByVal cchResult As Long, pcchResult As Long, ByVal dwReserved As Long) As HRESULTS


    Dim sBaseUrl As String
    Dim sRelativeUrl As String
    Dim sResult As String
    Dim bResult() As Byte

    sBaseUrl = olelib.SysAllocString(pwzBaseUrl)
    If InStr(sBaseUrl, "\") > 0 Then sBaseUrl = Replace(sBaseUrl, "\", "/")
    sBaseUrl = LiNVBLib.RightLeft(sBaseUrl, "/", vbBinaryCompare, ReturnOriginalStr)
    sRelativeUrl = olelib.SysAllocString(pwzRelativeUrl)

    If InStr(sRelativeUrl, ":") > 0 Then
        'byteResult = StrConv(sRelativeUrl, vbUnicode)
        sResult = sRelativeUrl
    Else
        sResult = sBaseUrl & "/" & sRelativeUrl
        'byteResult = StrConv(sBaseUrl & "/" & sRelativeUrl, vbUnicode)
    End If

    'pcchResult = UBound(byteResult) + 1
    'sResult = StrConv(sResult, vbFromUnicode)
    sResult = modReturnfunc.CanonicalizeUrl(sResult, dwCombineFlags)
    pcchResult = Len(sResult) * 2
    olelib.MoveMemory ByVal pwzResult, ByVal sResult, pcchResult + 2

  MDebug.DebugFPrint "BaseUrl : " & sBaseUrl & vbCrLf & _
                "RelativeUrl :" & sRelativeUrl & vbCrLf & _
                "CombineFlag :" & dwCombineFlags & vbCrLf & _
                "Url Return :" & sResult

CombineUrl = S_OK

              

End Function

Private Function myXUnzip(ByVal sZipfilename As String, ByVal sFilesToProcess As String, ByVal sUnzipTo As String)

    'MsgBox sZipfilename & vbCrLf & sFilesToProcess
 
    Static BUSY As Boolean

    Do Until BUSY = False
        DoEvents
    Loop

    BUSY = True

    With lUnzip
        .CaseSensitiveFileNames = False
        .PromptToOverwrite = False
        .OverwriteExisting = True
        .UseFolderNames = False
        .ZipFile = sZipfilename
        .FileToProcess = lusezipdll.CleanZipFilename(sFilesToProcess)
        .UnzipFolder = sUnzipTo
    End With

    lUnzip.unzip
    
    BUSY = False

    'With lXZ
    '.ZipFilename = sZipfilename
    '.FilesToProcess = sFilesToProcess
    '.PreservePaths = False
    '.UnzipToFolder = sUnzipTo
    '.SkipIfExisting = False
    '.unzip
    'End With

End Function

Friend Function ParseUrl(ByVal pwzUrl As Long, ByVal PARSEACTION As olelib.PARSEACTION, ByVal dwParseFlags As Long, ByVal pwzResult As Long, ByVal cchResult As Long, pcchResult As Long, ByVal dwReserved As Long) As HRESULTS

    Dim sUrl As String
    Dim sResult As String
    Dim parseType As String
    
    sUrl = olelib.SysAllocString(pwzUrl)
    sResult = ""
    ParseUrl = E_NOTIMPL
    
    Select Case PARSEACTION
    Case PARSE_ANCHOR
        sResult = LiNVBLib.RightLeft(sUrl, "#", vbBinaryCompare, ReturnEmptyStr)
        parseType = "PARSE_ANCHOR"
    Case PARSE_CANONICALIZE
        sResult = modReturnfunc.CanonicalizeUrl(sUrl, dwParseFlags)
        parseType = "PARSE_CANONICALIZE"
    Case PARSE_DECODE
        'sResult = linvblib.DecodeUrl(sUrl, CP_Default)
        parseType = "PARSE_DECODE"
    Case PARSE_DOCUMENT
        parseType = "PARSE_DOCUMENT"
    Case PARSE_DOMAIN
        sResult = "LocalHost"
        'sResult = LiNVBLib.LeftRange(sUrl, "//", "/", vbBinaryCompare, ReturnEmptyStr)
        parseType = "PARSE_DOMAIN"
    Case PARSE_ENCODE
        'sResult = linvblib.EncodeURI(sUrl)
        parseType = "PARSE_ENCODE"
    Case PARSE_ESCAPE
        'sResult = linvblib.EscapeUrl(sUrl)
        parseType = "PARSE_ESCAPE"
    Case PARSE_FRIENDLY
        sResult = LiNVBLib.LeftRange(sUrl, "//", "/", vbBinaryCompare, ReturnOriginalStr)
        parseType = "PARSE_FRIENDLY"
    Case PARSE_LOCATION
        sResult = LiNVBLib.RightRight(sUrl, "/", vbBinaryCompare, ReturnOriginalStr)
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
        sResult = LiNVBLib.LeftLeft(sUrl, "://", vbBinaryCompare, ReturnOriginalStr)
        parseType = "PARSE_SCHEMA"
    Case PARSE_SECURITY_DOMAIN
        sResult = "" 'LiNVBLib.LeftLeft(sUrl, "|/", vbBinaryCompare, ReturnOriginalStr)
         parseType = "PARSE_SECURITY_DOMAIN"
    Case PARSE_SECURITY_URL
        sResult = "" 'LiNVBLib.LeftLeft(sUrl, "|/", vbBinaryCompare, ReturnOriginalStr)
        parseType = "PARSE_SECURITY_URL"
    Case PARSE_SERVER
        parseType = "PARSE_SERVER"
    Case PARSE_SITE
        parseType = "PARSE_SITE"
    Case PARSE_UNESCAPE
        'sResult = linvblib.UnescapeUrl(sUrl)
        parseType = "PARSE_UNESCAPE"
    Case PARSE_URL_FROM_PATH
        sResult = "zip://" & sUrl
        parseType = "PARSE_URL_FROM_PATH"
    End Select
    'debug
    MDebug.DebugFPrint "ParseUrl of " & parseType & "(" & PARSEACTION & ")" & " of " & sUrl & vbCrLf & _
                     "Parse result : " & sResult
 
    If sResult <> "" Then
        pcchResult = LenB(sResult) + 2
        olelib.MoveMemory ByVal pwzResult, ByVal sResult, pcchResult
        ParseUrl = S_OK '  INET_E_DEFAULT_ACTION ' = S_OK 'S_FALSE
    Else
        ParseUrl = E_NOTIMPL
    End If

End Function

Friend Function QueryInfo(ByVal pwzUrl As Long, ByVal QueryOption As olelib.QueryOption, ByVal dwQueryFlags As Long, ByVal pBuffer As Long, ByVal cbBuffer As Long, pcbBuf As Long, ByVal dwReserved As Long) As HRESULTS
'QueryInfo = INET_E_DEFAULT_ACTION: Exit Function


    Dim sReturn As String
    Dim sUrl As String
    Dim sQueryType As String
    
    sUrl = olelib.SysAllocString(pwzUrl)
    sReturn = ""
    
    QueryInfo = INET_E_DEFAULT_ACTION
    
    Select Case QueryOption
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
        sQueryType = "QUERY_USES_NETWORK"
    Case QUERY_IS_CACHED
        sQueryType = "QUERY_IS_CACHED"
    Case QUERY_IS_INSTALLEDENTRY
        sQueryType = "QUERY_IS_INSTALLEDENTRY"
    Case QUERY_IS_CACHED_OR_MAPPED
        sQueryType = "QUERY_IS_CACHED_OR_MAPPED"
    Case QUERY_USES_CACHE
        sQueryType = "QUERY_USES_CACHE"
    Case QUERY_IS_SECURE
        pcbBuf = 4
        olelib.MoveMemory ByVal pBuffer, 0, pcbBuf
        QueryInfo = S_OK
        sQueryType = "QUERY_IS_SECURE"
    Case QUERY_IS_SAFE
        sQueryType = "QUERY_IS_SAFE"
    End Select
    
    If sReturn <> "" Then
        pcbBuf = LenB(sReturn) + 2
        olelib.MoveMemory ByVal pBuffer, ByVal sReturn, pcbBuf
        QueryInfo = S_OK
    End If
    
   MDebug.DebugFPrint "Queryinfo of " & sQueryType & "(" & CStr(QueryOption) & ")" & " For " & olelib.SysAllocString(pwzUrl) & vbCrLf & _
                "Result :" & sReturn

    
    
End Function

'
' Read
'
' Reads the data
'
Friend Function Read(ByVal pv As Long, ByVal cb As Long, pcbRead As Long) As Long

    Dim abBuffer() As Byte
    Dim curPos As Long
    Dim lastPos As Long
    Dim cbReal As Long
   
    'MsgBox "read " & m_sFile

    On Error GoTo hell

    If Not m_bDownloadDone Then
         
        'msgbox "not Download"
        
        'On Error Resume Next
         
        ' Read from the file

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
         
        'Put lTempFile, , abBuffer
        ' Copy the data to the buffer
        MoveMemory ByVal pv, abBuffer(0), cbReal
         
        ' Return how many bytes were readed

        pcbRead = cbReal
         
        ' Report progress
         
        Read = S_OK

        If Loc(m_lFile) < LOF(m_lFile) Then

            m_oSink.ReportData BSCF_INTERMEDIATEDATANOTIFICATION, Loc(m_lFile), LOF(m_lFile)

        Else

            m_oSink.ReportData BSCF_LASTDATANOTIFICATION, 100, 100
            m_oSink.ReportResult S_OK, 0, vbNullString

            m_bDownloadDone = True
            ' Close the file
            Close m_lFile
            'Close lTempFile
            Kill m_sFile
            'Kill sTempFile
            Read = S_FALSE
            
        End If
   
    Else
   
        ' Nothing readed
        pcbRead = 0
        ' Report the result code
        'm_oSink.ReportResult S_OK, 0, 0
        ' Return S_FALSE to notify the
        ' caller that all the data has been readed
         
        Read = S_FALSE
      
    End If
      
    Exit Function
      
hell:
    'MsgBox "error"
   
End Function

Friend Function Start(ByVal szURL As Long, ByVal pOIProtSink As olelib2.IInternetProtocolSink, ByVal pOIBindInfo As olelib.IInternetBindInfo, ByVal grfPI As olelib.PI_FLAGS, dwReserved As olelib.PROTOCOLFILTERDATA) As Long

    Dim sRealUrl As String
    Dim sUrl As zipUrl
    Dim sMimeType As String
    Dim pFolder As String

    sMimeType = "text/html"

    m_bDownloadDone = False

    ' Store the sink object
    Set m_oSink = pOIProtSink
    
'    Dim dBindinfo As BINDINFO
'    Set m_oBindInfo = pOIBindInfo
'    m_oBindInfo.GetBindInfo brfpi, dBindinfo
'    m_oBindInfo.GetBindString
    
    
    ' Parse the URL

    sRealUrl = SysAllocString(szURL)

    'MsgBox "Start " & vbCrLf & sRealUrl

    sUrl = zipProtocol_ParseURL(sRealUrl)

    If sUrl.sZipName = "" Or sUrl.sHtmlPath = "" Then

        'MsgBox "ERROR 1"
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

        MsgBox m_sFile
        MsgBox "ERROR2"
        Start = INET_E_DATA_NOT_AVAILABLE
        Exit Function
    
    End If

    sMimeType = MimeType(m_sFile)
    m_lFile = FreeFile
    Open m_sFile For Binary As m_lFile

ReportProgress:
    'Send the MIME type to the caller
 
    m_oSink.ReportProgress BINDSTATUS_MIMETYPEAVAILABLE, sMimeType
    m_oSink.ReportProgress BINDSTATUS_VERIFIEDMIMETYPEAVAILABLE, sMimeType

    'm_oSink.ReportProgress BINDSTATUS_ENDDOWNLOADDATA, ""

    ' Send the BEGINDOWNLOADDATA message to the caller

    'm_oSink.ReportProgress BINDSTATUS_BEGINDOWNLOADDATA, "Downloading " & m_sFile
    m_oSink.ReportData BSCF_FIRSTDATANOTIFICATION, 0, LOF(m_lFile)

End Function


