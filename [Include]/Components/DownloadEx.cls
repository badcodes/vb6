VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDownloadEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IDownloadClient

Private Const cstMaxConnections As Long = 5
Private m_Connections() As CDownload
Private m_bIsFree As Boolean
Private m_MaxConnection As Long
Private m_LonErrors As Long

Public Event ProgressChanged(ByVal vId As Long, ByRef Progress As CDownloadProgress)


Public Property Let MaxConnection(ByVal vMax As Long)
    If vMax < 1 Then vMax = 1
    m_MaxConnection = vMax
    ReDim Preserve m_Connections(1 To vMax)
End Property

Public Property Get IsFree() As Boolean
    IsFree = m_bIsFree
End Property


Public Function Download( _
    ByRef vSaveAS As String, _
    ByRef vUrl As String, _
    Optional ByRef vHeader As String = "", _
    Optional ByRef vOverWrite As Boolean = True) As Boolean
    
On Error GoTo ErrorDownload
    
    Dim Downloader As CDownload
    Set Downloader = CreateNewConnection()
    
    Downloader.Progress.CurrentBytes = 0
    Downloader.Progress.TotalBytes = 0
    Downloader.Progress.Url = vUrl
    Downloader.Progress.saveAs = vSaveAS
    
    If FileExists(vSaveAS) Then
        If vOverWrite = False Then
            Downloader.Progress.TextInfo = "文件已经存在，跳过"
            Download = True
            RaiseEvent ProgressChanged(Downloader.ID, Downloader.Progress)
            DoEvents
            Exit Function
        Else
            Downloader.Progress.TextInfo = "重新下载"
            Kill vSaveAS
        End If
    End If
    
    
    Dim Respone As CHttpHeader
    Set Respone = Downloader.Connect(vUrl, vHeader)
    If Respone Is Nothing Then
        Downloader.Progress.TextInfo = "连接错误"
        GoTo ErrorDownload
    End If
    Downloader.Progress.TextInfo = Respone.HttpStatus & " " & Respone.HttpRespone
    If Not Respone.HttpStatus = 200 Then
        If Not Respone.HttpStatus = 404 Then m_LonErrors = m_LonErrors + 1
        Download = False
        Exit Function
    End If
    Downloader.StartDownload (vSaveAS)
    Download = True
    'RaiseEvent ProgressChanged(Downloader.ID, Downloader.Progress)
    Exit Function
    
ErrorDownload:
    On Error Resume Next
    m_LonErrors = m_LonErrors + 1
    Download = False
    Downloader.Progress.TextInfo = Err.Description
    Err.Clear
End Function

Public Property Get count() As Long
    count = m_MaxConnection
End Property

Public Property Get Progress(ByRef Index As Long) As CDownloadProgress
    On Error Resume Next
    Set Progress = m_Connections(Index).Progress
End Property

Private Sub Class_Initialize()
    Me.MaxConnection = cstMaxConnections
    'Set m_Progress = New CDownloadProgress
    m_bIsFree = True
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    Dim i As Long
    For i = 1 To m_MaxConnection
        Set m_Connections(i).Progress = Nothing
        Set m_Connections(i) = Nothing
    Next
End Sub


Private Function CreateNewConnection() As CDownload
    
    Dim Index As Long
    Dim i As Long
    Index = -1

    For i = 1 To m_MaxConnection
        If m_Connections(i) Is Nothing Then
            Set m_Connections(i) = New CDownload
            Index = i
            Exit For
        ElseIf m_Connections(i).IsFree Then
            Index = i
            Exit For
        End If
    Next

    If Index = -1 Then
        Set CreateNewConnection = Nothing
    End If
    
    m_Connections(Index).ID = Index
    Set m_Connections(Index).Client = Me
    Set CreateNewConnection = m_Connections(Index)
    
End Function

Private Function ReleaseConnection(ByRef vDownloader As CDownload)
    Set m_Connections(vDownloader.ID) = Nothing
    Set vDownloader = Nothing
End Function

Private Sub IDownloadClient_ProgressChanged(ByVal vId As Long, vProgress As CDownloadProgress)
    RaiseEvent ProgressChanged(vId, vProgress)
End Sub

Public Property Get ErrorsCount() As Long
    ErrorsCount = m_LonErrors
End Property

Public Sub ClearErrors()
    m_LonErrors = 0
End Sub

Private Sub IDownloadClient_ProgressCompleted(ByVal vId As Long)

End Sub
