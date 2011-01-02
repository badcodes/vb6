VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTaskDownloader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Private WithEvents myTimer As CTimer

Private mThreads As Long
'Private mLastStatus As Long
'Private mLastStatusText As String
Private WithEvents mHttpConnection As CWinHttpDownloadEx
Attribute mHttpConnection.VB_VarHelpID = -1
'Private mStopNow As Boolean
Private mStep As Integer
Private mTextPdg As Boolean
Private mTaskName As String
Private mTask As CTask
Private mHeader As String
Private mRootUrl As String
Private mFilename As String
Private mDirectory As String

'Private mLastUrl As String

Private WithEvents mTester As CBookTester
Attribute mTester.VB_VarHelpID = -1

Public Event OnError()
Public Event Finish()
Public Event StatusChange(ByRef vText As String)

Public Event ProgressChange(ByVal vId As Long, vUrl As String, vMin As Long, vMax As Long)
Public Event ProgressDone(ByVal vId As Long, vUrl As String, vCount As Long)
Public Event ProgressError(ByVal vId As Long, vUrl As String, vError As String)
'Public Event ProgressStart(ByVal vId As Long, vUrl As String)

'Private mTestingTextPdg As Boolean
'Private mTestingStartPoint As Boolean
'Private WithEvents mHttpTester As CWinHTTP
'Private mTask As CTask
'Private mInfoRule As String
Private mUrls As CBookURL
Private mStartPoint As Long
Private mBusy As Boolean
Private mReChecked As Boolean

Private mTaskDone As Boolean
Private mResult As Boolean
Private mAborting As Boolean
Private mSkipContentPages As Boolean
Private mSkipCatalogPages As Boolean
Private mSkipStaticPages As Boolean
Private mJPGBookQuality As Integer
Private mRenameJPGBook As Boolean

Private Const cst_timer_interval As Long = 900

Private WithEvents myTimer As CTimer
Attribute myTimer.VB_VarHelpID = -1

Public Sub SetTimer(ByVal vInteval As Long)
    myTimer.Interval = vInteval
End Sub

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

'Private mPageCount As Long
'
'Public Property Get LastStatusText() As String
'    LastStatusText = mLastStatusText
'End Property
'
'Public Property Get LastStatus() As Long
'    LastStatus = mLastStatus
'End Property

Public Property Get Threads() As Long
    Threads = mThreads
End Property

Public Property Let Threads(ByVal LonValue As Long)
    mThreads = LonValue
    If mThreads < 1 Then mThreads = 1
    If Not mHttpConnection Is Nothing Then mHttpConnection.AbortAll
    Set mHttpConnection = New CWinHttpDownloadEx
    mHttpConnection.MaxConnection = mThreads
End Property

Private Sub Class_Initialize()
    mThreads = 1
   'mLastStatus =  0
    Set myTimer = New CTimer
    mJPGBookQuality = 1
End Sub
'Private Function GetInfoRule() As String
'    mInfoRule = BuildPath(mTask.Directory, "InfoRule.dat")
'    GetInfoRule = mInfoRule
'    If FileExists(mInfoRule) Then
'        mTextPdg = True
'        mUrls.SetTextPdg mInfoRule
'        GetStartPoint
'    Else
'        mTestingTextPdg = True
'        Set mHttpTester = New CWinHTTP
'        With mHttpTester
'            .Method = "GET"
'            .Destination = pSaveAs
'            .URL = vTask.bookInfo(SSF_URL) & "InfoRule.dat"
'            .Header = vTask.bookInfo(SSF_HEADER)
'            .Connect
'        End With
'    End If
'End Function
'
'
'Private Sub GetStartPoint(Optional vIdx As Long = 0)
'    mTestingStartPoint = True
'    Dim pSaveAs As String
'    If vIdx = 0 Then
'        pSaveAs = BuildPath(mTask.Directory, BuildRangeString("", 6, ".pdg", 1))
'    Else
'
'        pSaveAs = BuildPath(mTask.Directory, BuildRangeString("", 6, ".pdg", vIdx))
'    End If
'    mSaveas
'End Sub

Private Sub WaitConnections()
Static sWaiting As Boolean
If sWaiting Then Exit Sub
sWaiting = True
    If mTaskDone Then
        SetTimer 0
        mTaskDone = False
        Dim i As Long
        For i = 1 To mThreads
            If Not mHttpConnection.Threads(i) Is Nothing Then
                If Not mHttpConnection.Threads(i).IsFree Then GoTo ExitWait
            End If
        Next
        If mResult Then
            RaiseEvent StatusChange("完成下载：" & mTaskName)
            RaiseEvent Finish
        Else
            RaiseEvent StatusChange("下载出现错误：" & mTaskName)
            RaiseEvent OnError
        End If
    End If
ExitWait:
    sWaiting = False
End Sub



Public Function Download(ByRef vTask As CTask) As Boolean
    SetTimer 0
    mReChecked = False
    mResult = False
    mTaskDone = False
    mAborting = False
    mStep = 0
    mBusy = False
    
    Set mTask = vTask
    Set mTester = New CBookTester
    
    mTaskName = "《" & mTask.bookInfo(SSF_Title) & "》"
    If mTask.bookInfo(SSF_SSID) Then mTaskName = "[" & mTask.bookInfo(SSF_SSID) & "]" & mTaskName
    mHeader = vTask.bookInfo(SSF_HEADER)
    
    If mTask.IsJpgBook Then
        mRootUrl = vTask.bookInfo(SSF_JPGURL)
    Else
        mRootUrl = vTask.bookInfo(SSF_URL)
    End If
    
    mDirectory = BuildPath(vTask.Directory)
    
'    If mTask.IsJpgBook Then
'
'    Else
        RaiseEvent StatusChange("正在测试" & mTaskName & "...")
        mTester.Test vTask

        
    'End If
End Function

Private Sub TestComplete()
   
    If mTester.IsValid = False Then
        'mTask.Status = STS_ERRORS
        'RaiseEvent Finish
        RaiseEvent StatusChange(mTaskName & " 不能下载!")
        mResult = False
        mTaskDone = True
        'RaiseEvent OnError
       Exit Sub
    End If
    
    If mHttpConnection Is Nothing Then
        Set mHttpConnection = New CWinHttpDownloadEx
        mHttpConnection.MaxConnection = mThreads
    End If
    mTask.FilesCount = 0
    If mTester.IsJpgBook Then
        GetJpgBook
    ElseIf mTester.IsTextPdg Then
        GetTextPdg
    Else
        If mTester.StartPage > 0 Then
            mTask.bookInfo(SSF_StartPage) = mTester.StartPage
            mStartPoint = mTester.StartPage
            GetNormalPdg
        Else
            'mTask.Status = STS_ERRORS
            'RaiseEvent Finish
            mResult = False
            mTaskDone = True
            'RaiseEvent OnError
        End If
    End If
End Sub
Public Sub Abort()
    mAborting = True
    SetTimer 0
    Dim i As Long
    
    On Error Resume Next
    Dim vFile As String
    For i = 1 To mThreads
        If mHttpConnection.Threads(i) Is Nothing Then GoTo continue
        'If mHttpConnection.Threads(i).IsFree Then GoTo Continue
        vFile = mHttpConnection.Threads(i).Destination
        If vFile = "" Then GoTo continue
        Err.Clear
        Kill vFile
        If Err.Number = 0 Then RaiseEvent ProgressError(i, vFile, "删除正在下载的文件")
        'Err.Clear
        
        Debug.Print "Kill " & vFile
continue:
    Next
    
    mHttpConnection.AbortAll
    
    
    'mTask.Status = STS_PAUSE
    'RaiseEvent Finish
End Sub

'Private Sub mHttpConnection_BusyNow()
'    Dim i As Long
'    For i = 1 To 10000
'        DoEvents
'    Next
'    GetUrl ""
'End Sub

Private Sub UrlCompleted(ByVal vUrl As String, ByVal vId As Long)
   Debug.Print vUrl & "[Completed]"
    Dim vCount As Long
    If mTester.IsTextPdg Then
        vCount = MSSReader.TextPdgCount(vUrl)
    ElseIf Left$(vUrl, 1) = "0" Then
        vCount = 1
    End If
    RaiseEvent ProgressDone(vId, vUrl, vCount)
    DownloadNextUrl
    If mStep = -1 Then
        GetTextPdg2
    End If
End Sub

Private Sub Class_Terminate()
    'Set myTimer = Nothing
End Sub

Private Sub mHttpConnection_Completed(ByVal vId As Long, vStatus As Long)
    If vStatus = 200 Then
        Dim pUrl As String
        Dim pFile As String
        Dim pRename As String
        pUrl = ThreadUrlBaseName(vId)
        pFile = mDirectory & TmpFileName(pUrl)
        pRename = pUrl
        If mTask.IsJpgBook Then
            If FileExists(pFile) = False Then
                mUrls.ReportFail pUrl
                RaiseEvent ProgressError(vId, pUrl, "下载失败!")
                'DownloadNextUrl
                Exit Sub
            ElseIf FileLen(pFile) < 200 Then
                Kill pFile
                mUrls.ReportFail pUrl
                RaiseEvent ProgressError(vId, pUrl, "文件太小，放弃！")
                'DownloadNextUrl
                Exit Sub
            End If
            If mRenameJPGBook = True Then
                pRename = Left$(pUrl, Len(pUrl) - 3) & "pdg"
            End If
        End If
        On Error Resume Next
        Name pFile As mDirectory & pRename
        UrlCompleted pUrl, vId
    End If
End Sub
Private Function ThreadUrlBaseName(ByVal vId As Long) As String
    On Error Resume Next
    
    Dim iPos As Long
    Dim pUrl As String
    pUrl = mHttpConnection.Threads(vId).URL
    iPos = InStrRev(pUrl, "/")
    If iPos > 0 Then
        ThreadUrlBaseName = SubStringUntilMatch(pUrl, iPos + 1, "&")
        If ThreadUrlBaseName = "" Then ThreadUrlBaseName = Mid$(pUrl, iPos + 1)
    Else
        ThreadUrlBaseName = pUrl
    End If
End Function

Private Sub mHttpConnection_OnError(ByVal vId As Long, ByVal vErrorNum As Long, ByVal vDesc As String)
    If mStep = -1 Then
        SetTimer 0
        mStep = 0
        'mTask.Status = STS_ERRORS
        'RaiseEvent Finish
        mResult = False
        mTaskDone = True
        'RaiseEvent OnError
        WaitConnections
        Exit Sub
    End If
    Dim vUrl As String
    vUrl = ThreadUrlBaseName(vId)
    Debug.Print vUrl & "[Error]"
    mUrls.ReportFail vUrl
    RaiseEvent ProgressChange(vId, vUrl, 0, 0)
    RaiseEvent ProgressError(vId, vUrl, vDesc)
    'DownloadNextUrl
End Sub

Private Sub mHttpConnection_ProgressChange(ByVal vId As Long, ByVal vMin As Long, ByVal vMax As Long)
    RaiseEvent ProgressChange(vId, ThreadUrlBaseName(vId), vMin, vMax)
End Sub

Private Sub mHttpConnection_ResponeStart(ByVal vId As Long, ByVal vStatus As Long, ByVal vText As String)
    Dim vUrl As String
    vUrl = ThreadUrlBaseName(vId)
    Debug.Print vUrl & "[Start...]"
    If vStatus = 200 Then
        RaiseEvent ProgressChange(vId, vUrl, 0, 0)
    Else
        mUrls.ReportFail vUrl
        RaiseEvent ProgressError(vId, vUrl, vText & "(" & vStatus & ")")
        'DownloadNextUrl
    End If
    
End Sub

Private Sub mTester_StatusChange(vText As String)
    RaiseEvent StatusChange("正在测试" & mTaskName & ": " & vText)
End Sub

Private Function TmpFileName(ByVal vForm As String, Optional vReverseMode As Boolean = False)
    If vReverseMode Then
        TmpFileName = Left$(vForm, Len(vForm) - 5)
    Else
        TmpFileName = vForm & ".part"
    End If
End Function

Private Sub GetUrl(ByVal vUrl As String)
    Debug.Print vUrl & "[TRY]"
    'mBusy = True
    On Error GoTo GetUrl_Error
    Static lastURL As String
    If vUrl = "" Then vUrl = lastURL
    lastURL = vUrl
    Dim vSaveAS As String
    If mTask.IsJpgBook And mRenameJPGBook Then
        vSaveAS = mDirectory & Left$(vUrl, Len(vUrl) - 3) & "pdg"
    Else
        vSaveAS = mDirectory & vUrl
    End If
    
    
    If FileExists(vSaveAS) = True Then
        Dim vFreeId As Long
        Dim vFileLen As Long
        vFreeId = mHttpConnection.FreeId
        vFileLen = FileLen(vSaveAS)
        RaiseEvent ProgressChange(vFreeId, vUrl, vFileLen, vFileLen)
        UrlCompleted vUrl, vFreeId
        mBusy = False
        DownloadNextUrl
        Exit Sub
    End If
    
    vSaveAS = mDirectory & TmpFileName(vUrl)
    If FileExists(vSaveAS) Then Kill vSaveAS
    
    If mTask.IsJpgBook Then
        Dim pUrl As String
        pUrl = GetJpgBookPageUrl(mRootUrl, vUrl, HttpHeaderGetField(mHeader, "Cookie"), , mJPGBookQuality)
        If pUrl = "" Then
            mResult = False
            mTaskDone = True
            Exit Sub
        End If
        mHttpConnection.Connect pUrl, "GET", vSaveAS
    Else
        mHttpConnection.Connect BuildPath(mRootUrl, vUrl, lnpsUnix), "GET", vSaveAS, mHeader
    End If
    
    Exit Sub
GetUrl_Error:
    'WaitConnections
End Sub

Private Sub GetTextPdg()
    'Set mTester = Nothing
    RaiseEvent StatusChange("开始下载" & mTaskName & "...")

    mStep = -1
    GetUrl "InfoRule.dat"
    
End Sub
Private Sub GetNormalPdg()
    'Set mTester = Nothing
    RaiseEvent StatusChange("开始下载" & mTaskName & "...")

    mStep = 0
    Set mUrls = New CBookURL
    mUrls.SkipCatalogPages = mSkipCatalogPages
    mUrls.SkipContentPage = mSkipContentPages
    mUrls.SkipStaticPages = mSkipStaticPages
    mUrls.SetStartPoint mStartPoint
    StartDownload
    SetTimer cst_timer_interval
End Sub
Private Sub GetTextPdg2()
    mStep = 0
    Set mUrls = New CBookURL
    mUrls.SkipCatalogPages = mSkipCatalogPages
    mUrls.SkipContentPage = mSkipContentPages
    mUrls.SkipStaticPages = mSkipStaticPages
    mUrls.SetTextPdg BuildPath(mDirectory, "InfoRule.dat")
    StartDownload
    SetTimer cst_timer_interval
End Sub
Private Sub GetJpgBook()
    RaiseEvent StatusChange("开始下载" & mTaskName & "...")
    Set mUrls = New CBookURL
    mUrls.SkipStaticPages = True
    mUrls.SkipCatalogPages = mSkipCatalogPages
    mUrls.SkipContentPage = mSkipContentPages
    'mUrls.SkipStaticPages = mSkipStaticPages
    mUrls.Ext = ".jpg"
    StartDownload
    SetTimer cst_timer_interval
End Sub
Private Sub mTester_TestCompleted()
    RaiseEvent StatusChange("测试" & mTaskName & "完成。")
    TestComplete
End Sub

Private Sub StartDownload()
    
    'dim pC
    
    'Dim i As Long
    'For i = 1 To mThreads
        DownloadNextUrl
    'Next
End Sub
Private Sub DownloadNextUrl()
Static imbusy As Boolean
If imbusy Then Exit Sub
imbusy = True
    If mAborting Then Exit Sub
    If mTaskDone Then Exit Sub

    Dim vUrl As String
    If mHttpConnection.IsFree Then
        vUrl = mUrls.GetNextUrl
        If vUrl <> "" Then
            imbusy = False
            GetUrl vUrl
'        ElseIf mReChecked = False Then
'            mReChecked = True
'            RaiseEvent StatusChange("检查" & mTaskName & "...")
'            TestComplete
        Else
            mResult = True
            mTaskDone = True
            'WaitConnections
            'mTask.Status = STS_COMPLETE
            WaitConnections
            'SetTimer cst_timer_interval
            'RaiseEvent Finish
        End If
    End If
imbusy = False
End Sub

Private Sub mytimer_ThatTime()
    'DownloadNextUrl
    If mTaskDone Then
        WaitConnections
    ElseIf mAborting Then
        Exit Sub
    Else
        DownloadNextUrl
    End If
End Sub
