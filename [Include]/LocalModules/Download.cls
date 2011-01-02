VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDownload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private WithEvents m_Connection As CHTTPConnection
Attribute m_Connection.VB_VarHelpID = -1
Private m_Progress As CDownloadProgress
Private Const cstUserAgent As String = "Internet Explorer 5.x"
Private m_StrURL As String
Private m_StrSaveAS As String
Private m_Client As IDownloadClient
Public Event ProgressChanged(ByRef Progress As CDownloadProgress)
Private m_LonID As Long
Private m_bIsFree As Boolean

Public Property Get IsFree() As Boolean
    IsFree = m_bIsFree
End Property

Public Property Let ID(ByVal LonValue As Long)
    m_LonID = LonValue
End Property

Public Property Get ID() As Long
    ID = m_LonID
End Property
Public Property Get saveAs() As String
    saveAs = m_StrSaveAS
End Property

Public Property Get Url() As String
    Url = m_StrURL
End Property

Public Function Connect(ByRef Url As String, ByRef vHeader As String, Optional Reportor As CDownloadProgress, Optional UserAgent As String = cstUserAgent) As CHttpHeader
    On Error GoTo ERRORConnect
    
    If Not Reportor Is Nothing Then Set m_Progress = Reportor
    m_StrURL = Url
    If Not m_Connection Is Nothing Then Set m_Connection = Nothing
    Set m_Connection = New CHTTPConnection
    
    m_Progress.Url = Url
    m_Progress.TextInfo = "建立连接..."

    m_Connection.Url = Url
    m_Connection.Header = vHeader
    m_Connection.Connect
    Set Connect = m_Connection.Respone
    NotifyClient
    
    Exit Function
ERRORConnect:
    On Error Resume Next
    m_Progress.TextInfo = Err.Description
    Err.Clear
    'Set m_Connection = Nothing
    NotifyClient
End Function

Public Sub StartDownload(ByRef saveAs As String)
On Error GoTo ErrorStartDownload
    If m_Connection Is Nothing Then Exit Sub
    
    
    
    m_bIsFree = False
    m_StrSaveAS = saveAs
    m_Connection.Destination = saveAs
    m_Progress.Enable = True
    m_Progress.TextInfo = "开始下载"
    m_Progress.CurrentBytes = 0
    m_Progress.TotalBytes = 0
    m_Progress.saveAs = saveAs
    NotifyClient
    
    m_Connection.StartDownload
    m_bIsFree = True
    Exit Sub
ErrorStartDownload:
    m_Progress.TextInfo = Err.Description
    Err.Clear
    NotifyClient
    m_bIsFree = True
End Sub

Public Property Get Progress() As CDownloadProgress
    Set Progress = m_Progress
End Property

Private Sub Class_Initialize()
    Set m_Progress = New CDownloadProgress
    m_bIsFree = True
End Sub

Private Sub Class_Terminate()
    Set m_Progress = Nothing
    Set m_Client = Nothing
    Set m_Connection = Nothing
    
End Sub

Private Sub m_Connection_DownLoadOver()
    m_Progress.TextInfo = "下载完毕"
NotifyClient
End Sub

Private Sub m_Connection_ProgressChange(vMin As Long, vMax As Long)
    
    m_Progress.CurrentBytes = vMin
    m_Progress.TotalBytes = vMax
NotifyClient
End Sub

Private Sub m_Connection_StatusChange(vStatus As String)
    m_Progress.TextInfo = vStatus
    NotifyClient
End Sub

Public Property Set Client(ByRef vClient As IDownloadClient)
    Set m_Client = vClient
End Property

Public Property Let Client(ByRef vClient As IDownloadClient)
    Set m_Client = vClient
End Property

Private Sub NotifyClient()
    RaiseEvent ProgressChanged(m_Progress)
    If Not m_Client Is Nothing Then
        m_Client.ProgressChanged m_LonID, m_Progress
        m_Client.ProgressCompleted m_LonID
    End If
End Sub

