VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CWinHTTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Declaration Start
Private WithEvents mNet As WinHttpRequest
Attribute mNet.VB_VarHelpID = -1

Private mRespone As CHttpHeader
Private MHttpHeader As String

Private mUrl As String
Private mDestination As String
Private mvarConnect As Boolean
Private mMethod As String
Private mIsFree As Boolean
Public Event ProgressChange(ByVal vTag As Variant, ByRef vMin As Long, ByRef vMax As Long)
Public Event StatusChange(ByVal vTag As Variant, ByVal vStatus As String)
Public Event OnResponseFinished(ByVal vTag As Variant)
Public Event OnResponseStart(ByVal vTag As Variant, ByVal Status As Long)
Public Event OnDataAvailable(ByVal vTag As Variant, ByRef vData() As Byte)
Public Event OnTestCompleted(ByVal vTag As Variant, ByVal vStatus As Long, ByVal vText As String)

Public Event OnError(ByVal vTag As Variant, ByVal ErrorNum As Long, ByVal ErrorDescription As String)
Private Const CST_BASE_ERROR_NUM As Long = 120500

Private mTotalBytes As Long
Private mRetBytes As Long
Private mFileHandler As Integer
Private mGetData As Boolean
'Declaration End
Private mForceGetData As Boolean
Private mStatus As Long
Private mContentType As String
Private mAborting As Boolean
Private mTestMode As Boolean
Private mErrorCount As Long
Private mStatusText As String
Private mResponeHeader As String
Private mClient As IWinHttpClient
Private mTag As Variant
Private mBufferSize As Long
Private mBuffer() As Byte
Private mBufferLength As Long

Public Property Get BufferSize() As Long
    BufferSize = mBufferSize
End Property

Public Property Let BufferSize(ByVal LonValue As Long)
    mBufferSize = LonValue
End Property


Public Property Get ResponeBody() As Variant
    On Error Resume Next
    ResponeBody = mNet.ResponseBody
End Property


'Public Property Get ResponeStream() As Variant
'    ResponeStream = mNet.ResponseStream
'End Property


Public Property Get Tag() As Variant
    Tag = mTag
End Property

Public Property Let Tag(ByVal VarValue As Variant)
    mTag = VarValue
End Property

Public Property Get Client() As IWinHttpClient
    Set Client = mClient
End Property

Public Property Set Client(ByVal IWiValue As IWinHttpClient)
    Set mClient = IWiValue
End Property

Private Sub ResetStatus()
    mAborting = False
    mRetBytes = 0
    mTotalBytes = 0
    On Error Resume Next
    If mFileHandler <> 0 Then Close #mFileHandler
    mFileHandler = 0
    mStatus = 0
    mContentType = ""
    mStatusText = 0
    mResponeHeader = ""
    'mIsFree = True
    mErrorCount = 0
    Set mRespone = New CHttpHeader
    mAborting = False
    If mBufferSize < 1 Then mBufferSize = 1024 * 10
    ReDim mBuffer(0 To mBufferSize - 1)
    mBufferLength = 0
End Sub

Public Property Get ResponeHeader() As String
    ResponeHeader = mResponeHeader
End Property

Public Property Get StatusText() As String
    StatusText = mStatusText
End Property
Public Sub ClearError()
    mErrorCount = 0
End Sub
Public Property Get ErrorCount() As Long
    ErrorCount = mErrorCount
End Property

Public Property Get TestMode() As Boolean
    TestMode = mTestMode
End Property

Public Property Let TestMode(ByVal bValue As Boolean)
    mTestMode = bValue
End Property
Public Property Get BytesTotal() As Long
    BytesTotal = mTotalBytes
End Property
Public Property Get BytesDownloaded() As Long
    BytesDownloaded = mRetBytes
End Property
Public Property Get ContentType() As String
    ContentType = mContentType
End Property
Public Property Get Status() As Long
    Status = mStatus
End Property

Public Property Get ForceGetData() As Boolean
    ForceGetData = mForceGetData
End Property

Public Property Let ForceGetData(ByVal bValue As Boolean)
    mForceGetData = bValue
End Property

Public Sub ResetHeader()
    Set mNet = New WinHttpRequest
End Sub

Public Property Get Method() As String
    Method = mMethod
End Property

Public Property Let Method(ByVal StrValue As String)
    mMethod = StrValue
End Property


Property Let Header(ByVal vHeader As String)
    MHttpHeader = vHeader
End Property

'Public Property Get Respone() As CHttpHeader
'    Set Respone = mRespone
'End Property


Public Property Get IsFree() As Boolean
    IsFree = mIsFree
End Property

Private Sub SetHttpHeader(ByRef vText As String)
On Error Resume Next
            Dim vMap() As String
            Dim u As Long
            vMap = HttpHeader_Parse(MHttpHeader)
            u = UBound(vMap, 2)
            Dim i As Long
            Dim pUrl As String
            pUrl = HttpHeader_ParseRequest(vMap(0, 0))
            If pUrl <> "" And mUrl = "" Then
                mUrl = pUrl
            End If
            For i = 1 To u
                If vMap(0, i) <> "" And vMap(1, i) <> "" Then
                mNet.SetRequestHeader vMap(0, i), vMap(1, i)
                End If
                'Debug.Print vMap(0, i) & ":" & vMap(1, i)
            Next
Err.Clear
End Sub


Public Sub Init()
        mIsFree = False
        '<EhHeader>

        '</EhHeader>
        If Not mNet Is Nothing Then
            mAborting = True
            mNet.Abort
        End If
        Set mNet = New WinHttpRequest
        ResetStatus
End Sub

Public Sub OpenConnect()
On Error GoTo Connect_Err

    mNet.Open mMethod, mUrl, 1
    If MHttpHeader <> "" Then SetHttpHeader (MHttpHeader)
    If Not mClient Is Nothing Then mClient.Url = mUrl
    Exit Sub
Connect_Err:
        
        mNet_OnError Err.Number, Err.Description

End Sub

Public Sub Send(Optional vBody As Variant)
On Error GoTo Connect_Err

    mNet.Send vBody
    Exit Sub
    
Connect_Err:
        
        mNet_OnError Err.Number, Err.Description
End Sub

Public Sub Connect()
    mIsFree = False
        '<EhHeader>
        On Error GoTo Connect_Err
        '</EhHeader>
        If Not mNet Is Nothing Then
            mAborting = True
            mNet.Abort
        End If
        Set mNet = New WinHttpRequest
        mNet.Option(WinHttpRequestOption_EnableRedirects) = 1
        ResetStatus
        

        'If mFileHandler <> 0 Then Close mFileHandler
        If Not mClient Is Nothing Then mClient.Url = mUrl
100     mNet.Open mMethod, mUrl, 1
    If MHttpHeader <> "" Then SetHttpHeader (MHttpHeader)
        mNet.Send
        'mNet.WaitForResponse
        
        '<EhFooter>
        Exit Sub

Connect_Err:
        
        mNet_OnError Err.Number, Err.Description
        
End Sub


Public Property Let Destination(ByVal filename As String)

    mDestination = filename

End Property
Public Property Get Destination() As String
    Destination = mDestination
End Property
Public Property Let Url(ByVal Url As String)


    mUrl = Url

End Property

Public Property Get Url() As String
    Url = mUrl
End Property


Private Sub Class_Initialize()
    'Set mNet = New WinHttpRequest
    'Set mRespone = New CHttpHeader
    'mNet.Option WinHttpRequestOption_URL
    mMethod = "GET"
    mIsFree = True
    mStatus = 0
    mBufferSize = 1024 * 10
End Sub

Private Sub Class_Terminate()
    If mFileHandler <> 0 Then Close #mFileHandler
    Set mRespone = Nothing
    Set mNet = Nothing
End Sub

Public Sub Abort()
    On Error Resume Next
    If mIsFree Then Exit Sub
    mAborting = True
    mNet.Abort
    mIsFree = True
    If mFileHandler <> 0 Then Close #mFileHandler
    mFileHandler = 0
End Sub


Public Property Let Options(ByVal vOption As WinHttp.WinHttpRequestOption, ByVal RHS As Variant)
    mNet.Option(vOption) = RHS
End Property

Public Property Get Options(ByVal vOption As WinHttp.WinHttpRequestOption) As Variant
    Options = mNet.Option(vOption)
End Property

'CSEH: ErrResumeNext
Public Sub SetAutoLogonPolicy(ByVal AutoLogonPolicy As WinHttp.WinHttpRequestAutoLogonPolicy)
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>
    mNet.SetAutoLogonPolicy AutoLogonPolicy
End Sub

'CSEH: ErrResumeNext
Public Sub SetClientCertificate(ByVal ClientCertificate As String)
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>
    mNet.SetClientCertificate ClientCertificate
End Sub

'CSEH: ErrResumeNext
Public Sub SetCredentials(ByVal UserName As String, ByVal Password As String, ByVal Flags As WinHttp.HTTPREQUEST_SETCREDENTIALS_FLAGS)
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>
    mNet.SetCredentials UserName, Password, Flags
End Sub

Public Sub SetProxy(ByVal ProxySetting As WinHttp.HTTPREQUEST_PROXY_SETTING, Optional ByVal ProxyServer As Variant, Optional ByVal BypassList As Variant)
    On Error Resume Next
    mNet.SetProxy ProxySetting, ProxyServer
End Sub

Public Sub SetRequestHeader(ByVal Header As String, ByVal Value As String)
    On Error Resume Next
    mNet.SetRequestHeader Header, Value
End Sub

Public Sub SetTimeouts(ByVal ResolveTimeout As Long, ByVal ConnectTimeout As Long, ByVal SendTimeout As Long, ByVal ReceiveTimeout As Long)
    mNet.SetTimeouts ResolveTimeout, ConnectTimeout, SendTimeout, ReceiveTimeout
End Sub


Public Function WaitForResponse(Optional ByVal Timeout As Variant) As Boolean
    mNet.WaitForResponse Timeout
End Function

Private Sub mNet_OnError(ByVal ErrorNumber As Long, ByVal ErrorDescription As String)
    
    '<CSCustomCode> 1
        Debug.Print "Entered - " & "LiNNetWork.CWinHTTP.mNet_OnError"
    '</CSCustomCode> 1
    
    RaiseEvent StatusChange(mTag, "Error " & ErrorNumber & " :" & ErrorDescription)
    RaiseEvent OnError(mTag, ErrorNumber, ErrorDescription)
    If mTestMode Then
        RaiseEvent OnTestCompleted(mTag, ErrorNumber, ErrorDescription)
    End If
    mErrorCount = mErrorCount + 1
    mStatus = ErrorNumber
    mStatusText = ErrorDescription
    'Err.Raise CST_BASE_ERROR_NUM + ErrorNumber, ErrorDescription
    
    On Error Resume Next
    
    If mFileHandler <> 0 Then Close #mFileHandler
    mFileHandler = 0
    
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

Private Sub FlushBuffer(Optional vReport As Boolean = False)
    Dim i As Long
    ReDim pData(0 To mBufferLength - 1) As Byte
    For i = 0 To mBufferLength - 1
        pData(i) = mBuffer(i)
    Next
    mBufferLength = 0
    WriteData pData, vReport
End Sub
'CSEH: ErrExit
Private Sub mNet_OnResponseDataAvailable(Data() As Byte)

'Static pThreadBusy As Boolean
'Do Until pThreadBusy = False
'    DoEvents
'Loop
'pThreadBusy = True

    '<CSCustomCode> 1
        Debug.Print "Entered - " & "LiNNetWork.CWinHTTP.mNet_OnResponseDataAvailable"
    '</CSCustomCode> 1
    '<EhHeader>
    On Error GoTo mNet_OnResponseDataAvailable_Err
    '</EhHeader>
    If mAborting Then Exit Sub
    If Not mGetData Then Exit Sub
    Dim pBytes As Long
    pBytes = ArrayUbound(Data())
    If pBytes = 0 Then Exit Sub

    mRetBytes = mRetBytes + pBytes + 1
    If mRetBytes > mTotalBytes Then mTotalBytes = mRetBytes
    
    WriteData Data, True
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
    Close #mFileHandler
    mFileHandler = 0
    mGetData = False
   
    '</EhFooter>
End Sub

Private Sub WriteData(ByRef vData() As Byte, Optional vReport As Boolean = False)
    On Error GoTo mNet_OnResponseDataAvailable_Err
    
    If mDestination <> "" Then
        If mFileHandler = 0 Then
            If FileExists(mDestination) Then Kill mDestination
            mFileHandler = FreeFile
            Open mDestination For Binary Access Write As #mFileHandler
        End If
        Put #mFileHandler, , vData()
    End If
    
    If vReport Then
        If mDestination <> "" Then
            RaiseEvent ProgressChange(mTag, mRetBytes, mTotalBytes)
        Else
            RaiseEvent OnDataAvailable(mTag, vData)
        End If
        If Not mClient Is Nothing Then
            mClient.ReportDataAvail mTag, vData, mTotalBytes
            mClient.ReportProgress mTag, mRetBytes, mTotalBytes
        End If
    End If

    Exit Sub
mNet_OnResponseDataAvailable_Err:
    mNet_OnError Err.Number, Err.Description
    On Error Resume Next
    Close #mFileHandler
    mFileHandler = 0
    mGetData = False
    '</EhFooter>
End Sub

Private Sub mNet_OnResponseFinished()
    '<CSCustomCode> 1
        Debug.Print "Entered - " & "LiNNetWork.CWinHTTP.mNet_OnResponseFinished"
    '</CSCustomCode> 1
    If mAborting Then Exit Sub
    Debug.Print mUrl
    If mBufferLength > 0 Then FlushBuffer True
    
    'If Not mGetData Then Exit Sub
    On Error Resume Next
    If mFileHandler <> 0 Then Close #mFileHandler
    mFileHandler = 0
    
    mGetData = False
    RaiseEvent OnResponseFinished(mTag)
    If Not mClient Is Nothing Then mClient.ReportComplete mTag, mStatus
    mIsFree = True
'    mRespone.Init mNet.GetAllResponseHeaders
'    mTotalBytes = mRespone.ConentLength
'    mRetBytes = 0
End Sub

Private Sub mNet_OnResponseStart(ByVal Status As Long, ByVal ContentType As String)
    If mAborting Then Exit Sub
    'Set mRespone = New CHttpHeader
    On Error Resume Next
    '<CSCustomCode> 1
        Debug.Print "Entered - " & "LiNNetWork.CWinHTTP.mNet_OnResponseStart"
    '</CSCustomCode> 1
    mIsFree = False
    mResponeHeader = mNet.GetAllResponseHeaders
    mRespone.Init mResponeHeader
    mTotalBytes = mRespone.ConentLength
    mRespone.HttpStatus = Status
    mRespone.ContentType = ContentType
    mStatus = Status
    mContentType = ContentType
    mRetBytes = 0
    mStatusText = mRespone.HttpRespone
    If Status = 200 Then mGetData = True
    If mForceGetData Then mGetData = True
    If mTestMode Then
        RaiseEvent OnTestCompleted(mTag, Status, mStatusText)
        mNet.Abort
        mGetData = False
    End If
    If Not mClient Is Nothing Then mClient.ReportRespone mTag, mStatus, mStatusText
    RaiseEvent OnResponseStart(mTag, Status)
    
End Sub


