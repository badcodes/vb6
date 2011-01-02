VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CWinHttpSimple"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IWinHttp

'Declaration Start
Private WithEvents mNet As WinHttpRequest
Attribute mNet.VB_VarHelpID = -1

Private mResponse As CHttpHeader
Private MHttpHeader As String

Private mUrl As String
Private mDestination As String
Private mvarConnect As Boolean
Private mMethod As String
Private mIsFree As Boolean
'Public Event ProgressChange(ByVal vTag As Variant, ByRef vMin As Long, ByRef vMax As Long)
'Public Event StatusChange(ByVal vTag As Variant, ByVal vStatus As String)
'Public Event OnResponseFinished(ByVal vTag As Variant)
'Public Event OnResponseStart(ByVal vTag As Variant, ByVal Status As Long)
'Public Event OnDataAvailable(ByVal vTag As Variant, ByRef vData() As Byte)
Public Event OnTestCompleted(ByVal vTag As Variant, ByVal vStatus As Long, ByVal vText As String)

'Public Event OnError(ByVal vTag As Variant, ByVal ErrorNum As Long, ByVal ErrorDescription As String)

Private Const CST_BASE_ERROR_NUM As Long = 120500

Private mTotalBytes As Long
Private mRetBytes As Long
'Private mFileHandler As Integer
Private mGetData As Boolean
'Declaration End
Private mForceGetData As Boolean
Private mStatus As Long
Private mContentType As String
Private mAborting As Boolean
Private mTestMode As Boolean
Private mErrorCount As Long
Private mStatusText As String
Private mResponseHeader As String
Private mClient As IWinHttpClient
Private mTag As Variant
Private mBufferSize As Long
Private mBuffer() As Byte
Private mBufferLength As Long
Private mAsyn As Boolean
Private Sub ResetStatus()

    MHttpHeader = ""
    mUrl = ""
    mMethod = "GET"
    mDestination = ""
    mIsFree = True
    mTotalBytes = 0
    mRetBytes = 0
    mGetData = False
    mForceGetData = False
    mStatus = 0
    mContentType = ""
    mAborting = False
    mTestMode = False
    mErrorCount = 0
    mStatusText = ""
    mResponseHeader = ""
    Set mClient = Nothing
    
    
    On Error Resume Next
    'If mFileHandler <> 0 Then Close #mFileHandler
    Set mResponse = New CHttpHeader

    If mBufferSize < 1 Then mBufferSize = 1024 * 10
    ReDim mBuffer(0 To mBufferSize - 1)
    mBufferLength = 0
End Sub

Public Property Get IWinHttp_BufferSize() As Long
    IWinHttp_BufferSize = mBufferSize
End Property

Public Property Let IWinHttp_BufferSize(ByVal LonValue As Long)
    mBufferSize = LonValue
End Property




Public Property Get IWinHttp_ResponseBody() As Variant
    On Error Resume Next
    IWinHttp_ResponseBody = mNet.ResponseBody
End Property


Private Property Get IWinHttp_ResponseText() As String
    IWinHttp_ResponseText = mNet.ResponseText
End Property

'Public Property Get IWinHttp_ResponseStream() As Variant
'    ResponseStream = mNet.ResponseStream
'End Property


Public Property Get IWinHttp_Tag() As Variant
    IWinHttp_Tag = mTag
End Property

Public Property Let IWinHttp_Tag(ByVal VarValue As Variant)
    mTag = VarValue
End Property

Public Property Get IWinHttp_Client() As IWinHttpClient
    Set IWinHttp_Client = mClient
End Property

Public Property Set IWinHttp_Client(ByVal IWiValue As IWinHttpClient)
    Set mClient = IWiValue
End Property



'CSEH: ErrExit
Public Property Get IWinHttp_ResponseHeader(Optional vHeader As String) As String
    '<EhHeader>
    On Error GoTo IWinHttp_ResponseHeader_Err
    '</EhHeader>
    If vHeader <> "" Then
        IWinHttp_ResponseHeader = mNet.GetResponseHeader(vHeader)
    Else
        IWinHttp_ResponseHeader = mResponseHeader
    End If
    '<EhFooter>
    Exit Property

IWinHttp_ResponseHeader_Err:
    Debug.Print "GetSSLibX.CWinHttpSimple.IWinHttp_ResponseHeader:Error " & Err.Description
    Err.Clear

    '</EhFooter>
End Property

Public Property Get IWinHttp_StatusText() As String
    IWinHttp_StatusText = mStatusText
End Property
Public Sub IWinHttp_ClearError()
    mErrorCount = 0
End Sub
Public Property Get IWinHttp_ErrorCount() As Long
    IWinHttp_ErrorCount = mErrorCount
End Property

Public Property Get IWinHttp_TestMode() As Boolean
    IWinHttp_TestMode = mTestMode
End Property

Public Property Let IWinHttp_TestMode(ByVal bValue As Boolean)
    mTestMode = bValue
End Property
Public Property Get IWinHttp_BytesTotal() As Long
    IWinHttp_BytesTotal = mTotalBytes
End Property
Public Property Get IWinHttp_BytesDownloaded() As Long
    IWinHttp_BytesDownloaded = mRetBytes
End Property
Public Property Get IWinHttp_ContentType() As String
    IWinHttp_ContentType = mContentType
End Property
Public Property Get IWinHttp_Status() As Long
    IWinHttp_Status = mStatus
End Property

Public Property Get IWinHttp_ForceGetData() As Boolean
    IWinHttp_ForceGetData = mForceGetData
End Property

Public Property Let IWinHttp_ForceGetData(ByVal bValue As Boolean)
    mForceGetData = bValue
End Property

Public Sub IWinHttp_ResetHeader()
    Set mNet = New WinHttpRequest
End Sub

Public Property Get IWinHttp_Method() As String
    IWinHttp_Method = mMethod
End Property

Public Property Let IWinHttp_Method(ByVal StrValue As String)
    mMethod = StrValue
End Property


Property Let IWinHttp_Header(ByVal vHeader As String)
    MHttpHeader = vHeader
End Property

'Public Property Get IWinHttp_Response() As CHttpHeader
'    Set Response = mResponse
'End Property


Public Property Get IWinHttp_IsFree() As Boolean
    IWinHttp_IsFree = mIsFree
End Property

Private Sub SetHttpHeader(ByRef vText As String)
On Error Resume Next
            Dim vMap() As String
            Dim u As Long
            vMap = HttpHeader_Parse(MHttpHeader)
            u = UBound(vMap, 2)
            Dim I As Long
            Dim pUrl As String
            pUrl = HttpHeader_ParseRequest(vMap(0, 0))
            If pUrl <> "" And mUrl = "" Then
                mUrl = pUrl
            End If
            For I = 1 To u
                If vMap(0, I) <> "" And vMap(1, I) <> "" Then
                mNet.SetRequestHeader vMap(0, I), vMap(1, I)
                End If
                'Debug.Print vMap(0, i) & ":" & vMap(1, i)
            Next
Err.Clear
End Sub


Public Sub IWinHttp_Init()
On Error Resume Next
        'mIsFree = False
        '<EhHeader>

        '</EhHeader>
        If Not mNet Is Nothing Then
            mAborting = True
            IWinHttp_Abort
            mAborting = False
        Else
            Set mNet = New WinHttpRequest
        End If
        ResetStatus
End Sub

Public Sub IWinHttp_OpenConnect(Optional ByVal vAsyn As Boolean = True)
On Error GoTo Connect_Err
    mIsFree = False
    mAsyn = vAsyn
    mNet.Open mMethod, mUrl, vAsyn
    If MHttpHeader <> "" Then SetHttpHeader (MHttpHeader)
    If Not mClient Is Nothing Then mClient.URL = mUrl
    Exit Sub
Connect_Err:
        
        mNet_OnError Err.Number, Err.Description

End Sub

Public Sub IWinHttp_Send(Optional vBody As Variant)
On Error GoTo Connect_Err

    mNet.Send vBody
    Exit Sub
    
Connect_Err:
        
        mNet_OnError Err.Number, Err.Description
End Sub

Public Sub IWinHttp_Connect(Optional ByVal vAsyn As Boolean = True)
    mIsFree = False
        '<EhHeader>
        On Error GoTo Connect_Err
        '</EhHeader>
        If Not mNet Is Nothing Then
            mAborting = True
            mNet.Abort
        End If
        Set mNet = New WinHttpRequest
        'mNet.Option(WinHttpRequestOption_EnableRedirects) = 1
        ResetStatus
        

        'If mFileHandler <> 0 Then Close mFileHandler
        If Not mClient Is Nothing Then mClient.URL = mUrl
100     mNet.Open mMethod, mUrl, vAsyn
    If MHttpHeader <> "" Then SetHttpHeader (MHttpHeader)
        mNet.Send
        'mNet.WaitForResponse
        
        '<EhFooter>
        Exit Sub

Connect_Err:
        
        mNet_OnError Err.Number, Err.Description
        
End Sub


Public Property Let IWinHttp_Destination(ByVal filename As String)

    mDestination = filename

End Property
Public Property Get IWinHttp_Destination() As String
    IWinHttp_Destination = mDestination
End Property
Public Property Let IWinHttp_Url(ByVal URL As String)


    mUrl = URL

End Property

Public Property Get IWinHttp_Url() As String
    IWinHttp_Url = mUrl
End Property


Private Sub Class_Initialize()
    'Set mNet = New WinHttpRequest
    'Set mResponse = New CHttpHeader
    'mNet.Option WinHttpRequestOption_URL
    mMethod = "GET"
    mIsFree = True
    mStatus = 0
    mBufferSize = 1024 * 10
End Sub

Private Sub Class_Terminate()
    'If mFileHandler <> 0 Then Close #mFileHandler
    Set mResponse = Nothing
    Set mNet = Nothing
End Sub

Public Sub IWinHttp_Abort()
    On Error Resume Next
    If mIsFree Then Exit Sub
    
    If mAsyn Then
        mAborting = True
        mNet.Abort
        mIsFree = True
    'If mFileHandler <> 0 Then Close #mFileHandler
    'mFileHandler = 0
        mAborting = False
    End If
End Sub


Public Property Let IWinHttp_Options(ByVal vOption As WinHttp.WinHttpRequestOption, ByVal RHS As Variant)
    mNet.Option(vOption) = RHS
End Property

Public Property Get IWinHttp_Options(ByVal vOption As WinHttp.WinHttpRequestOption) As Variant
     IWinHttp_Options = mNet.Option(vOption)
End Property

''CSEH: ErrResumeNext
'Public Sub IWinHttp_SetAutoLogonPolicy(ByVal AutoLogonPolicy As WinHttp.WinHttpRequestAutoLogonPolicy)
'    '<EhHeader>
'    On Error Resume Next
'    '</EhHeader>
'    mNet.SetAutoLogonPolicy AutoLogonPolicy
'End Sub
'
''CSEH: ErrResumeNext
'Public Sub IWinHttp_SetClientCertificate(ByVal ClientCertificate As String)
'    '<EhHeader>
'    On Error Resume Next
'    '</EhHeader>
'    mNet.SetClientCertificate ClientCertificate
'End Sub
'
''CSEH: ErrResumeNext
'Public Sub IWinHttp_SetCredentials(ByVal UserName As String, ByVal Password As String, ByVal Flags As WinHttp.HTTPREQUEST_SETCREDENTIALS_FLAGS)
'    '<EhHeader>
'    On Error Resume Next
'    '</EhHeader>
'    mNet.SetCredentials UserName, Password, Flags
'End Sub
'
'Public Sub IWinHttp_SetProxy(ByVal ProxySetting As WinHttp.HTTPREQUEST_PROXY_SETTING, Optional ByVal ProxyServer As Variant, Optional ByVal BypassList As Variant)
'    On Error Resume Next
'    mNet.SetProxy ProxySetting, ProxyServer
'End Sub

Public Sub IWinHttp_SetRequestHeader(ByVal Header As String, ByVal Value As String)
    On Error Resume Next
    mNet.SetRequestHeader Header, Value
End Sub

Public Sub IWinHttp_SetTimeouts(ByVal ResolveTimeout As Long, ByVal ConnectTimeout As Long, ByVal SendTimeout As Long, ByVal ReceiveTimeout As Long)
    mNet.SetTimeouts ResolveTimeout, ConnectTimeout, SendTimeout, ReceiveTimeout
End Sub


Public Function IWinHttp_WaitForResponse(Optional ByVal TimeOut As Variant) As Boolean
    mNet.WaitForResponse TimeOut
End Function



Private Sub mNet_OnError(ByVal ErrorNumber As Long, ByVal ErrorDescription As String)
    
    '<CSCustomCode> 1
        Debug.Print "Entered - " & "LiNNetWork.CWinHTTPSimple.mNet_OnError"
    '</CSCustomCode> 1
    
    'RaiseEvent StatusChange(mTag, "Error " & ErrorNumber & " :" & ErrorDescription)
    'RaiseEvent OnError(mTag, ErrorNumber, ErrorDescription)
    If mTestMode Then
        RaiseEvent OnTestCompleted(mTag, ErrorNumber, ErrorDescription)
    End If
    mErrorCount = mErrorCount + 1
    mStatus = ErrorNumber
    mStatusText = ErrorDescription
    'Err.Raise CST_BASE_ERROR_NUM + ErrorNumber, ErrorDescription
    
    On Error Resume Next
    
    'If mFileHandler <> 0 Then Close #mFileHandler
    'mFileHandler = 0
    
    If Not mClient Is Nothing Then mClient.ReportError mTag, ErrorNumber, ErrorDescription
    'If mAborting Then Exit Sub
    mIsFree = True
End Sub

'CSEH: ErrExit
Private Function ArrayUbound(ByRef vData() As Byte) As Long
    '<EhHeader>
    On Error GoTo ArrayUbound_Err
    '</EhHeader>
        ArrayUbound = UBound(vData, 1)
    '<EhFooter>
    Exit Function

ArrayUbound_Err:
    Err.Clear
    ArrayUbound = -1
    '</EhFooter>
End Function

'Private Sub FlushBuffer(Optional vReport As Boolean = False)
'    Dim i As Long
'    ReDim pData(0 To mBufferLength - 1) As Byte
'    For i = 0 To mBufferLength - 1
'        pData(i) = mBuffer(i)
'    Next
'    mBufferLength = 0
'    WriteData pData, vReport
'End Sub
'CSEH: ErrExit
Private Sub mNet_OnResponseDataAvailable(Data() As Byte)

'Static pThreadBusy As Boolean
'Do Until pThreadBusy = False
'    DoEvents
'Loop
'pThreadBusy = True

    '<CSCustomCode> 1
        Debug.Print "Entered - " & "LiNNetWork.CWinHTTPSimple.mNet_OnResponseDataAvailable"
    '</CSCustomCode> 1
    '<EhHeader>
    On Error GoTo mNet_OnResponseDataAvailable_Err
    '</EhHeader>
    If mAborting Then Exit Sub
    If Not mGetData Then Exit Sub
    
    'Dim pData() As Byte
    'pData = Data
    
    Dim pBytes As Long
    pBytes = ArrayUbound(Data())
    If pBytes = 0 Then Exit Sub

    mRetBytes = mRetBytes + pBytes + 1
    If mRetBytes > mTotalBytes Then mTotalBytes = mRetBytes
'
    If Not mClient Is Nothing Then
        mClient.ReportDataAvail mTag, Data, mTotalBytes
        mClient.ReportProgress mTag, mRetBytes, mTotalBytes
    End If
            
    
'    WriteData pData, True
'
'
'    Static sCount As Integer
'    sCount = sCount + 1
'    If sCount > 10 Then
'        sCount = 0
'        WriteData Data, True
'    Else
'        WriteData Data, False
'    End If
'
'
'
'    If mBufferSize - mBufferLength > pBytes Then
'        mBufferLength = mBufferLength + pBytes
'        Dim i As Long
'        For i = 0 To pBytes - 1
'            mBuffer(mBufferLength - pBytes + i) = Data(i)
'        Next
'    Else
'        If mBufferLength > 0 Then FlushBuffer False
'        WriteData Data, True
'    End If
'pThreadBusy = False
                
    '<EhFooter>
    Exit Sub

mNet_OnResponseDataAvailable_Err:
    mNet_OnError Err.Number, Err.Description
    On Error Resume Next
'    Close #mFileHandler
'    mFileHandler = 0
    mGetData = False
   
    '</EhFooter>
End Sub

'Private Sub WriteData(ByRef vData() As Byte, Optional vReport As Boolean = False)
'    On Error GoTo mNet_OnResponseDataAvailable_Err
'    Static pIsBusy As Boolean
'    If mDestination <> "" Then
'        If mFileHandler = 0 Then
'            Do Until pIsBusy = False
'            Loop
'            pIsBusy = True
'            If FileExists(mDestination) Then Kill mDestination
'            mFileHandler = FreeFile
'            Open mDestination For Binary Access Write As #mFileHandler
'            pIsBusy = False
'        End If
'        Put #mFileHandler, , vData()
'    End If
'
'    If vReport Then
''        If mDestination <> "" Then
''            RaiseEvent ProgressChange(mTag, mRetBytes, mTotalBytes)
''        Else
''            RaiseEvent OnDataAvailable(mTag, vData)
''        End If
'        If Not mClient Is Nothing Then
'            mClient.ReportDataAvail mTag, vData, mTotalBytes
'            mClient.ReportProgress mTag, mRetBytes, mTotalBytes
'        End If
'    End If
'
'    Exit Sub
'mNet_OnResponseDataAvailable_Err:
'    mNet_OnError Err.Number, Err.Description
'    On Error Resume Next
'    Close #mFileHandler
'    mFileHandler = 0
'    mGetData = False
'    pIsBusy = False
'    '</EhFooter>
'End Sub

Private Sub mNet_OnResponseFinished()
    On Error GoTo Error_ResponseFinished
    '<CSCustomCode> 1
        Debug.Print "Entered - " & "LiNNetWork.CWinHTTPSimple.mNet_OnResponseFinished"
    '</CSCustomCode> 1
    If mAborting Then Exit Sub
    'If mGetData = False Then Exit Sub
    'Debug.Print mUrl
    Dim fNum As Integer
    If mGetData = True And mDestination <> "" And IsEmpty(mNet.ResponseBody) = False Then
       If FileExists(mDestination) Then Kill mDestination
            Dim pData() As Byte
            pData = mNet.ResponseBody
            fNum = FreeFile
            Open mDestination For Binary Access Write As #fNum
            Put #fNum, , pData
            Close #fNum
            fNum = 0
            mGetData = False
    End If
    
    
    'RaiseEvent OnResponseFinished(mTag)
    If Not mClient Is Nothing Then mClient.ReportComplete mTag, mStatus
    mIsFree = True
    Exit Sub
    
Error_ResponseFinished:
    On Error Resume Next
    mGetData = False
    mIsFree = True
    If fNum <> 0 Then Close #fNum
    'mFileHandler = 0
'    mResponse.Init mNet.GetAllResponseHeaders
'    mTotalBytes = mResponse.ConentLength
'    mRetBytes = 0
End Sub

Private Sub mNet_OnResponseStart(ByVal Status As Long, ByVal ContentType As String)
    If mAborting Then Exit Sub
    'Set mResponse = New CHttpHeader
    On Error Resume Next
    '<CSCustomCode> 1
        Debug.Print "Entered - " & "LiNNetWork.CWinHTTPSimple.mNet_OnResponseStart"
    '</CSCustomCode> 1
    mIsFree = False
    mResponseHeader = mNet.GetAllResponseHeaders
    mResponse.Init mResponseHeader
    mTotalBytes = mResponse.ConentLength
    mResponse.HttpStatus = Status
    mResponse.ContentType = ContentType
    mStatus = Status
    mContentType = ContentType
    mRetBytes = 0
    mStatusText = mResponse.HttpResponse
    If Status = 200 Then mGetData = True
    If mForceGetData Then mGetData = True
    If mTestMode Then
        IWinHttp_Abort
        RaiseEvent OnTestCompleted(mTag, Status, mStatusText)
    Else
        If Not mClient Is Nothing Then mClient.ReportResponse mTag, mStatus, mStatusText
    End If
    'RaiseEvent OnResponseStart(mTag, Status)
    
End Sub



