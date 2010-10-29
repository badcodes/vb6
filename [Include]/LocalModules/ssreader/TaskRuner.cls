VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTaskRunner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private mNotify As ITaskNotify
Private WithEvents mDownloader As CTaskDownloader2
Attribute mDownloader.VB_VarHelpID = -1
Private mProgress() As CDownloadProgress
Private mTask As CTask
Private mThreadCount As Long
Private mIsFree As Boolean
Private mStatusText  As String
Private mFiles As Long

Private mSkipContentPages As Boolean
Private mSkipCatalogPages As Boolean
Private mSkipStaticPages As Boolean
Private mJPGBookQuality As Integer
Private mRenameJPGBook As Boolean
Private mTaskName As String
Private mSavedType As Boolean
Private mSavedFiles As Long
Private mRunTimes As Long
Private mThreadProgressCount As Long


Public Sub Runonce()
    If Not mDownloader Is Nothing Then
        Debug.Print "Time " & mTaskName & " mDownloader"
        mDownloader.TimerMe
    End If
End Sub

Public Property Get ThreadProgressCount() As Long
    ThreadProgressCount = mThreadCount + 1
End Property

Public Property Get RenameJPGBook() As Boolean
    RenameJPGBook = mRenameJPGBook
End Property

Public Property Let RenameJPGBook(ByVal bValue As Boolean)
    mRenameJPGBook = bValue
End Property
Public Property Get JPGBookQuality() As Integer
    JPGBookQuality = mJPGBookQuality
End Property

Public Property Let JPGBookQuality(ByVal intValue As Integer)
    mJPGBookQuality = intValue
End Property

Public Property Get SkipStaticPages() As Boolean
    SkipStaticPages = mSkipStaticPages
End Property

Public Property Let SkipStaticPages(ByVal bValue As Boolean)
    mSkipStaticPages = bValue
End Property

Public Property Get SkipCatalogPages() As Boolean
    SkipCatalogPages = mSkipCatalogPages
End Property

Public Property Let SkipCatalogPages(ByVal bValue As Boolean)
    mSkipCatalogPages = bValue
End Property

Public Property Get SkipContentPages() As Boolean
    SkipContentPages = mSkipContentPages
End Property

Public Property Let SkipContentPages(ByVal bValue As Boolean)
    mSkipContentPages = bValue
End Property

Public Property Get Files() As Long
    Files = mFiles
End Property
Public Property Let Files(ByVal vCount As Long)
    mFiles = vCount
End Property
Public Property Get Client() As ITaskNotify
    Set Client = mNotify
End Property

Public Property Let Client(ByVal ITaValue As ITaskNotify)
    Set mNotify = ITaValue
End Property

Public Sub Execute(ByRef vClient As ITaskNotify, ByRef vTask As CTask, Optional ByVal vDownloadThread As Long = 3)
    
        If vTask Is Nothing Then Exit Sub
        If Not vTask.InitForDownload Then
            mDownloader_OnError
            Exit Sub
        End If
        mRunTimes = 0
        vTask.Changed = True
        mIsFree = False
        mTaskName = vTask.ToString()
        mSavedType = vTask.IsJpgBook
        
        If vDownloadThread < 1 Then vDownloadThread = 1
        Set mDownloader = New CTaskDownloader2
        mDownloader.SkipCatalogPages = mSkipCatalogPages
        mDownloader.SkipContentPages = mSkipContentPages
        mDownloader.SkipStaticPages = mSkipStaticPages
        mDownloader.JPGBookQuality = mJPGBookQuality
        mDownloader.RenameJPGBook = mRenameJPGBook
        mDownloader.Threads = vDownloadThread
        ReDim mProgress(0 To vDownloadThread)
        Dim I As Long
        For I = 0 To vDownloadThread
            Set mProgress(I) = New CDownloadProgress
            mProgress(I).TextInfo = "准备下载..."
        Next
        mThreadCount = vDownloadThread
        Set mNotify = vClient
        Set mTask = vTask
        
        Execute2
        
End Sub

Private Sub Execute2()
    mFiles = 0
    mRunTimes = mRunTimes + 1
    mDownloader.Download mTask
End Sub

Public Property Let ThreadCount(ByVal vCount As Long)
    If vCount < 1 Then vCount = 1
    mThreadCount = vCount
End Property


Public Property Get ThreadCount() As Long
    ThreadCount = mThreadCount
End Property

Public Property Get ThreadProgress(ByVal vId As Long) As CDownloadProgress
    On Error Resume Next
    Set ThreadProgress = mProgress(vId)
End Property

'Private Sub If Not mNotify Is Nothing Then mNotify.TaskStatusChange Me()
'    If Not mNotify Is Nothing Then mNotify.TaskStatusChange Me
'End Sub
'
'Private Sub If Not mNotify Is Nothing Then mNotify.DownloadStatusChange Me()
'    If Not mNotify Is Nothing Then mNotify.DownloadStatusChange Me
'End Sub
'
'Private Sub If Not mNotify Is Nothing Then mNotify.TaskComplete Me()
'    If Not mNotify Is Nothing Then mNotify.TaskComplete Me
'End Sub

Public Sub Abort()
    On Error Resume Next
    mDownloader.Abort
    mIsFree = True
    If Not mTask Is Nothing Then mTask.IsJpgBook = mSavedType
End Sub

Private Sub Class_Initialize()
    mIsFree = True
    mJPGBookQuality = 1
End Sub

Private Sub Class_Terminate()
    Set mTask = Nothing
    Set mNotify = Nothing
    Dim I As Long
    For I = 1 To mThreadCount
        Set mProgress(I) = Nothing
    Next
    Set mDownloader = Nothing
End Sub

Private Sub mDownloader_Finish()
    
    If mRunTimes = 1 And mSavedType = False Then
        mRunTimes = 2
        mTask.IsJpgBook = True
        mSavedFiles = mTask.FilesCount
        If Not mNotify Is Nothing Then
            mNotify.UpdateJpgBook mTask
            mNotify.StatusChange mTask.Directory & "：补页中..."
        End If
        If mTask.bookInfo(SSF_JPGURL) = vbNullString Then
            mNotify.StatusChange mTask.Directory & ": 无效的JPGURL信息!"
            mDownloader_Finish
            Exit Sub
        End If
        With mDownloader
            .SkipStaticPages = True
            .SkipContentPages = True
            .SkipCatalogPages = True
        End With
        Execute2
        Exit Sub
    Else
        If mSavedType = False Then
            mTask.FilesCount = mSavedFiles
            mTask.IsJpgBook = mSavedType
        End If
        mTask.Status = STS_COMPLETE
        If Not mNotify Is Nothing Then
            mNotify.StatusChange mTask.Directory & ": 完成"
            mNotify.TaskComplete Me
        End If
        'Set mtask = Nothing
        'Set mDownloader = Nothing
        mIsFree = True
    End If
End Sub

Private Sub mDownloader_OnError()
    mTask.Status = STS_ERRORS
    If mRunTimes > 1 And mSavedType = False Then
        mTask.IsJpgBook = mSavedType
        mTask.FilesCount = mSavedFiles
    End If
    If Not mNotify Is Nothing Then
        
        mNotify.TaskComplete Me
        'mnotify.StatusChange "
    End If
    'Set mTask = Nothing
    mIsFree = True
    'Set mDownloader = Nothing
End Sub

Private Sub mDownloader_ProgressChange(ByVal vId As Long, vUrl As String, vMin As Long, vMax As Long)
On Error Resume Next
With mProgress(vId)
    .URL = vUrl
    .CurrentBytes = vMin
    .TotalBytes = vMax
    .TextInfo = "正在下载..."
End With
If Not mNotify Is Nothing Then mNotify.DownloadStatusChange Me
End Sub

Private Sub mDownloader_ProgressDone(ByVal vId As Long, vUrl As String, vFileCount As Long)
On Error Resume Next
mFiles = mFiles + vFileCount
mTask.FilesCount = mFiles
With mProgress(vId)
    .URL = vUrl
    .TextInfo = "下载完成"
End With
If Not mNotify Is Nothing Then mNotify.DownloadStatusChange Me
If Not mNotify Is Nothing Then mNotify.TaskStatusChange Me
End Sub

Private Sub mDownloader_ProgressError(ByVal vId As Long, vUrl As String, vError As String)
On Error Resume Next
With mProgress(vId)
    .URL = vUrl
    .CurrentBytes = 0
    .TotalBytes = 0
    .TextInfo = "错误:" & vError
End With
If Not mNotify Is Nothing Then mNotify.DownloadStatusChange Me
End Sub

Private Sub mDownloader_StatusChange(vText As String)
    'mStatusText = vText
    If Not mNotify Is Nothing Then mNotify.StatusChange mTaskName & ":" & vText
    'If Not mNotify Is Nothing Then mNotify.TaskStatusChange Me
End Sub

Public Property Get StatusText()
    StatusText = mStatusText
    If Not mNotify Is Nothing Then mNotify.TaskStatusChange Me
End Property

Public Property Get IsFree() As Boolean
    IsFree = mIsFree
End Property

Public Property Get task() As CTask
    Set task = mTask
End Property
