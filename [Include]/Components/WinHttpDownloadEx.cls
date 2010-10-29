VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CWinHttpDownloadEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements IWinHttpClient

Private mMaxConnection As Long
Private mIsFree As Boolean
Private mDownloader() As IWinHttp

Private mConnectionsCount As Long

Public Event Completed(ByVal vId As Long, vStatus As Long)
Public Event DataAvailable(ByVal vId As Long, ByRef vData() As Byte, vTotalByte As Long)
Public Event OnError(ByVal vId As Long, ByVal vErrorNum As Long, ByVal vDesc As String)
Public Event ResponeStart(ByVal vId As Long, ByVal vStatus As Long, ByVal vText As String)
Public Event ProgressChange(ByVal vId As Long, ByVal vMin As Long, ByVal vMax As Long)

Public Event BusyNow()

Public Property Get FreeId() As Long
    FreeId = GetNewDownloader
End Property
Public Property Get IsFree() As Boolean
Static imbusy As Boolean
If imbusy Then Exit Property
imbusy = True
    If mConnectionsCount < mMaxConnection Then
        IsFree = True
    Else
        Dim i As Long
        For i = 1 To mConnectionsCount
            If mDownloader(i) Is Nothing Then IsFree = True: Exit For
            If mDownloader(i).IsFree Then IsFree = True: Exit For
        Next
    End If
imbusy = False
    'IsFree = False
    'IsFree = (GetNewDownloader > 0)
End Property

Public Property Get MaxConnection() As Long
    MaxConnection = mMaxConnection
End Property

Public Property Let MaxConnection(ByVal LonValue As Long)
    If LonValue < 1 Then LonValue = 1
    mMaxConnection = LonValue
End Property

Public Property Get Threads(ByVal vId As Long) As IWinHttp
    On Error Resume Next
    Set Threads = mDownloader(vId)
End Property
Private Sub Class_Initialize()
    mMaxConnection = 1
    mIsFree = True
End Sub
Private Function GetNewDownloader() As Long
Static imbusy As Boolean
Do Until imbusy = False
    DoEvents
Loop
imbusy = True
    Dim i As Long
    Dim ret As Long
    For i = 1 To mConnectionsCount
        If mDownloader(i) Is Nothing Then ret = i: Exit For
        If mDownloader(i).IsFree Then
            ret = i
            Exit For
        End If
    Next
    If ret < 1 And mConnectionsCount < mMaxConnection Then
            mConnectionsCount = mConnectionsCount + 1
            ReDim Preserve mDownloader(1 To mConnectionsCount)
            ret = mConnectionsCount
    End If
    If ret > 0 Then
            If mDownloader(ret) Is Nothing Then Set mDownloader(ret) = New CWinHttpSimple
            mDownloader(ret).Init
    End If
    GetNewDownloader = ret
imbusy = False
End Function
Private Sub ReleaseDownloader()
    Dim i As Long
    For i = 1 To mConnectionsCount
        Set mDownloader(i).Client = Nothing
        mDownloader(i).Abort
        Set mDownloader(i) = Nothing
    Next
End Sub
Public Function Connect(vUrl As String, _
        Optional vMethod As String = "GET", _
        Optional vSaveAS As String = "", _
        Optional vHeader As String = "" _
    ) As Boolean
    Dim pId As Long
    pId = GetNewDownloader()
    If pId > 0 Then
        Connect = True
        With mDownloader(pId)
            Set .Client = Me
            .Tag = pId
            .Method = vMethod
            .Header = vHeader
            .URL = vUrl
            .Destination = vSaveAS
            .OpenConnect True
            .Send
        End With
    Else
        Connect = False
        RaiseEvent BusyNow
    End If
End Function

Private Function CheckId(ByVal vId As Long) As Boolean
    If vId < 1 Then Exit Function
    If vId > mMaxConnection Then Exit Function
    If mDownloader(vId) Is Nothing Then Exit Function
    CheckId = True
End Function
Public Sub Abort(ByVal vId As Long)
    If Not CheckId(vId) Then Exit Sub
    mDownloader(vId).Abort
End Sub


Public Sub AbortAll()
    On Error Resume Next
    Dim i As Long
    For i = 1 To mMaxConnection
        If Not mDownloader(i) Is Nothing Then mDownloader(i).Abort
    Next
End Sub


Private Sub IWinHttpClient_ReportComplete(ByVal vTag As Variant, vStatus As Long)
    On Error Resume Next
    RaiseEvent Completed(vTag, vStatus)
    'Set mDownloader(vTag) = Nothing
End Sub


Private Sub IWinHttpClient_ReportDataAvail(ByVal vTag As Variant, vData() As Byte, ByVal vTotalByte As Long)
End Sub

Private Sub IWinHttpClient_ReportError(ByVal vTag As Variant, ByVal vErrorNum As Long, ByVal vErrorDesc As String)
    On Error Resume Next
    RaiseEvent OnError(vTag, vErrorNum, vErrorDesc)
    'Set mDownloader(vTag) = Nothing
End Sub

Private Sub IWinHttpClient_ReportProgress(ByVal vTag As Variant, ByVal vMin As Long, ByVal vMax As Long)
    RaiseEvent ProgressChange(vTag, vMin, vMax)
End Sub

Private Sub IWinHttpClient_ReportRespone(ByVal vTag As Variant, ByVal vStatus As Long, ByVal vStatusText As String)
    On Error Resume Next
    RaiseEvent ResponeStart(vTag, vStatus, vStatusText)
End Sub

Private Property Let IWinHttpClient_Url(ByVal RHS As String)
End Property

