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
'Implements olelib.IInternetProtocolInfo

' Private members
Private m_oSink As IInternetProtocolSink  ' Callback object used to informthe progress
Private m_lFile As Long ' File handle
Private m_bDownloadDone As Boolean        ' Flag used to know if all the data has been readed
Private m_sFile As String
Private InvaildPassword As Boolean
Private PASSWORD As String
Private WithEvents lUnzip As cUnzip
Attribute lUnzip.VB_VarHelpID = -1
Private sTempFolder As String
'Private WithEvents lXZ As XceedZip
'Private Enum Errors
'   errCannotOpenConnection
'   errCannotOpenRecordset
'   errInvalidURL
'   errCannotOpenFile
'End Enum
Private Sub Class_Initialize()

    Dim oPH As IInternetProtocol
    Set lUnzip = New cUnzip
    ' Get the IInternetProtocol interface
    Set oPH = Me
    ' Replace the Read method
    ReplaceVTableEntry ObjPtr(oPH), 10, AddressOf modReturnfunc.IInternetProtocol_Read
    ReplaceVTableEntry ObjPtr(oPH), 4, AddressOf modReturnfunc.IInternetProtocol_Start
    sTempFolder = Environ$("temp")
    modReturnfunc.instanceCount = modReturnfunc.instanceCount + 1
    
End Sub
Private Sub Class_Terminate()

    'zipProtocolCount = zipProtocolCount - 1
    Set lUnzip = Nothing
    Set m_oSink = Nothing
    modReturnfunc.instanceCount = modReturnfunc.instanceCount - 1
    If modReturnfunc.instanceCount < 0 Then modReturnfunc.instanceCount = 0
    'Dim fso As New FileSystemObject
    On Error Resume Next
    If modReturnfunc.instanceCount = 0 Then RmDir sTempFolder
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
    'Err.Raise E_NOTIMPL
End Sub
Private Sub IInternetProtocol_LockRequest(ByVal dwOptions As Long)
    'msgbox "lock " & m_sFile
End Sub
Private Sub IInternetProtocol_Read(ByVal pv As Long, ByVal cb As Long, pcbRead As Long)
End Sub
Private Sub IInternetProtocol_Resume()
    'msgbox "resume " & m_sFile
    'Err.Raise E_NOTIMPL
End Sub
Private Sub IInternetProtocol_Seek(ByVal dlibMove As Currency, ByVal dwOrigin As Long, plibNewPosition As Currency)
    'msgbox "seek " & m_sFile
    'Err.Raise E_NOTIMPL
End Sub
Private Sub IInternetProtocol_Start(ByVal szUrl As Long, ByVal pOIProtSink As olelib2.IInternetProtocolSink, ByVal pOIBindInfo As olelib.IInternetBindInfo, ByVal grfPI As olelib.PI_FLAGS, dwReserved As olelib.PROTOCOLFILTERDATA)
End Sub
Private Sub IInternetProtocol_Suspend()
    'msgbox "suspend " & m_sFile
    'Err.Raise E_NOTIMPL
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
Private Function myXUnzip(ByVal sZipfilename As String, ByVal sFilesToProcess As String, ByVal sUnzipTo As String) As unzReturnCode
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
Friend Function Start(ByVal szUrl As Long, ByVal pOIProtSink As olelib2.IInternetProtocolSink, ByVal pOIBindInfo As olelib.IInternetBindInfo, ByVal grfPI As olelib.PI_FLAGS, dwReserved As olelib.PROTOCOLFILTERDATA) As Long
    Dim sRealUrl As String
    Dim sUrl As zipUrl
    Dim sMimeType As String
    sMimeType = "text/html"
    Start = S_OK
    m_bDownloadDone = False
    ' Store the sink object
    Set m_oSink = pOIProtSink
    ' Parse the URL
    sRealUrl = SysAllocString(szUrl)
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
        'MsgBox m_sFile
        'MsgBox "ERROR2"
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
