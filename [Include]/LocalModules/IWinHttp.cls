VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IWinHttp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Public Event ProgressChange(ByVal vTag As Variant, ByRef vMin As Long, ByRef vMax As Long)
Public Event StatusChange(ByVal vTag As Variant, ByVal vStatus As String)
Public Event OnResponseFinished(ByVal vTag As Variant)
Public Event OnResponseStart(ByVal vTag As Variant, ByVal Status As Long)
Public Event OnDataAvailable(ByVal vTag As Variant, ByRef vData() As Byte)
Public Event OnTestCompleted(ByVal vTag As Variant, ByVal vStatus As Long, ByVal vText As String)
Public Event OnError(ByVal vTag As Variant, ByVal ErrorNum As Long, ByVal ErrorDescription As String)

Public Property Get BufferSize() As Long
  
End Property

Public Property Let BufferSize(ByVal LonValue As Long)
   
End Property

Public Property Get ResponseText() As String

End Property
Public Property Get ResponseBody() As Variant
    
End Property


Public Property Get Tag() As Variant
End Property

Public Property Let Tag(ByVal VarValue As Variant)
End Property

Public Property Get Client() As IWinHttpClient
End Property

Public Property Set Client(ByVal IWiValue As IWinHttpClient)
End Property

Private Sub ResetStatus()
End Sub

Public Property Get ResponseHeader(Optional vHeader As String) As String
End Property

Public Property Get StatusText() As String
End Property
Public Sub ClearError()
End Sub
Public Property Get ErrorCount() As Long
End Property

Public Property Get TestMode() As Boolean
End Property

Public Property Let TestMode(ByVal bValue As Boolean)
End Property
Public Property Get BytesTotal() As Long
End Property
Public Property Get BytesDownloaded() As Long
End Property
Public Property Get ContentType() As String
End Property
Public Property Get Status() As Long
End Property

Public Property Get ForceGetData() As Boolean
End Property

Public Property Let ForceGetData(ByVal bValue As Boolean)
End Property

Public Sub ResetHeader()
End Sub

Public Property Get Method() As String
End Property

Public Property Let Method(ByVal StrValue As String)
End Property


Property Let Header(ByVal vHeader As String)
End Property

'Public Property Get Response() As CHttpHeader
'    Set Response = mResponse
'End Property


Public Property Get IsFree() As Boolean
End Property




Public Sub Init()
End Sub

Public Sub OpenConnect(Optional ByVal vAsyn As Boolean = True)


End Sub

Public Sub Send(Optional vBody As Variant)
End Sub




Public Property Let Destination(ByVal filename As String)



End Property
Public Property Get Destination() As String

End Property
Public Property Let URL(ByVal URL As String)




End Property

Public Property Get URL() As String

End Property



Public Sub Abort()

End Sub


Public Property Let Options(ByVal vOption As WinHttp.WinHttpRequestOption, ByVal RHS As Variant)
End Property

Public Property Get Options(ByVal vOption As WinHttp.WinHttpRequestOption) As Variant
End Property

'CSEH: ErrResumeNext
'Public Sub SetAutoLogonPolicy(ByVal AutoLogonPolicy As WinHttp.WinHttpRequestAutoLogonPolicy)
'
'End Sub
'
''CSEH: ErrResumeNext
'Public Sub SetClientCertificate(ByVal ClientCertificate As String)
'
'End Sub
'
''CSEH: ErrResumeNext
'Public Sub SetCredentials(ByVal UserName As String, ByVal Password As String, ByVal Flags As WinHttp.HTTPREQUEST_SETCREDENTIALS_FLAGS)
'
'End Sub
'
'Public Sub SetProxy(ByVal ProxySetting As WinHttp.HTTPREQUEST_PROXY_SETTING, Optional ByVal ProxyServer As Variant, Optional ByVal BypassList As Variant)
'
'End Sub

Public Sub SetRequestHeader(ByVal Header As String, ByVal Value As String)

End Sub

Public Sub SetTimeouts(ByVal ResolveTimeout As Long, ByVal ConnectTimeout As Long, ByVal SendTimeout As Long, ByVal ReceiveTimeout As Long)
End Sub


Public Function WaitForResponse(Optional ByVal Timeout As Variant) As Boolean
End Function



