VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTaskDownloader2"
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


Private mResult As Boolean
Private mAborting As Boolean
Private mSkipContentPages As Boolean
Private mSkipCatalogPages As Boolean
Private mSkipStaticPages As Boolean
Private mJPGBookQuality As Integer
Private mRenameJPGBook As Boolean

Private Const cst_timer_interval As Long = 300

Private mTimerEnabled As Boolean
Private mTaskDone As Boolean
'Private mTimerAction As Boolean
Private mStartDownload As Boolean
Private mStartWaiting As Long
Private WithEvents taskTimer As CTimer
Attribute taskTimer.VB_VarHelpID = -1
Private mIsTextPdg As Boolean
Private mHeartBeat As Boolean
Private mWaiting As Boolean

Private Enum TIMER_ACTIONS
    TIMER_ACTION_NOTHING = 0
    TIMER_ACTION_TASK_HEARTBEAT
    TIMER_ACTION_TASK_START
    TIMER_ACTION_TASK_ABORT
    TIMER_ACTION_TASK_COMPLETED
    TIMER_ACTION_SELF_DESTROY
End Enum
Private mTimerAction As TIMER_ACTIONS

Private Sub TestComplete()
   
   mIsTextPdg = mTester.IsTextPdg
    If mTester.IsValid = False Then
        'mTask.Status = STS_ERRORS
        'RaiseEvent Finish
        RaiseEvent StatusChange("不能下载!")
        mTaskDone = True
        mResult = False
        TaskComplete
       Exit Sub
    End If
    
    If mHttpConnection Is Nothing Then
        Set mHttpConnection = New CWinHttpDownloadEx
        mHttpConnection.MaxConnection = mThreads
    End If
    mTask.FilesCount = 0
    If mTester.IsJpgBook Then
        mStep = 0
        Set mUrls = New CBookURL
        mUrls.SkipStaticPages = True
        mUrls.SkipCatalogPages = mSkipCatalogPages
        mUrls.SkipContentPage = mSkipContentPages
        If mTester.IsTextPdg = True Then
            mUrls.SkipCatalogPages = False
            mUrls.SkipContentPage = False
        End If
        'mUrls.SkipStaticPages = mSkipStaticPages
        mUrls.Ext = ".jpg"
        taskTimer.Interval = 2 * cst_timer_interval
    ElseIf mTester.IsTextPdg Then
            Set mTester = Nothing
            If FileExists(mDirectory & "InfoRule.dat") Then
                GetTextPdg2
                Exit Sub
            Else
                mStep = -1
                mHttpConnection.Connect mRootUrl & "InfoRule.dat", "GET", mDirectory & "Inforule.dat", mHeader
                Exit Sub
            End If
    Else
        If mTester.StartPage > 0 Then
            mTask.bookInfo(SSF_StartPage) = mTester.StartPage
            mStartPoint = mTester.StartPage
            mStep = 0
           Set mUrls = New CBookURL
           mUrls.SkipCatalogPages = mSkipCatalogPages
           mUrls.SkipContentPage = mSkipContentPages
           mUrls.SkipStaticPages = mSkipStaticPages
           mUrls.SetStartPoint mStartPoint
        Else
            mTaskDone = True
            mResult = False
            TaskComplete
            Set mTester = Nothing
            Exit Sub
        End If
    End If
    Set mTester = Nothing
    mTimerAction = TIMER_ACTION_TASK_START
    'StartDownload
End Sub

Private Sub StartTimer()
    
    mTimerAction = TIMER_ACTION_NOTHING
    If taskTimer Is Nothing Then
        Set taskTimer = New CTimer
    End If
        taskTimer.Interval = cst_timer_interval


End Sub


Private Sub StartDownload()
    
        
    '========================
    'Download Loop Start
    '========================
    
    'IF MHTTPCONNECTION.IsFree
    
    Do
        If mTaskDone Then
            TaskComplete
            Exit Do
        End If
        mTimerAction = TIMER_ACTION_TASK_HEARTBEAT
        If mHeartBeat Then
            mHeartBeat = False
            If mHttpConnection.IsFree Then DownloadNextURL
        End If
        DoEvents
    Loop Until mAborting
    
    mTimerAction = TIMER_ACTION_NOTHING
    On Error Resume Next
    mHttpConnection.AbortAll
    
    
End Sub

Private Sub taskTimer_ThatTime()
Static BLOCK_ME As Boolean
If BLOCK_ME Then Exit Sub
BLOCK_ME = True
Dim THIS As TIMER_ACTIONS
THIS = mTimerAction
mTimerAction = TIMER_ACTION_NOTHING
    Select Case THIS
        Case TIMER_ACTIONS.TIMER_ACTION_NOTHING
        Case TIMER_ACTIONS.TIMER_ACTION_SELF_DESTROY
            On Error Resume Next
            taskTimer.Interval = 0
            Set taskTimer = Nothing
        Case TIMER_ACTIONS.TIMER_ACTION_TASK_ABORT
            Call Abort
        Case TIMER_ACTIONS.TIMER_ACTION_TASK_COMPLETED
            Call TaskComplete
        Case TIMER_ACTIONS.TIMER_ACTION_TASK_HEARTBEAT
            mHeartBeat = True
        Case TIMER_ACTIONS.TIMER_ACTION_TASK_START
            BLOCK_ME = False
            Call StartDownload
    End Select
BLOCK_ME = False
End Sub

Private Sub DownloadNextURL()
Static ImBusy As Boolean
    If ImBusy Then Exit Sub
    ImBusy = True
    If Not mHttpConnection.IsFree Then ImBusy = False: Exit Sub
    If mTaskDone Then ImBusy = False: Exit Sub
    If mAborting Then ImBusy = False: Exit Sub
    
    'ImBusy = True

ReEnter:
    ImBusy = True
   
    Dim vUrl As String
    Dim vSaveAS As String
    Dim vFileLen As Long
        
    vUrl = mUrls.GetNextUrl
    If vUrl = "" Then
        mTaskDone = True
        mResult = True
        ImBusy = False
        Exit Sub
    End If
    
    
    
    Dim pUrl As String
    Dim pSaveAs As String
    Dim pHeader As String

    
    RaiseEvent StatusChange("正在下载...")
    Debug.Print mTask.ToString & vUrl & "[TRY]"
    
        If mTask.IsJpgBook Then
            If mRenameJPGBook Then
                vSaveAS = mDirectory & Left$(vUrl, Len(vUrl) - 3) & "pdg"
            Else
                vSaveAS = mDirectory & vUrl
            End If
            pHeader = ""
        Else
            pUrl = mRootUrl & vUrl
            vSaveAS = mDirectory & vUrl
            pHeader = mHeader
        End If
        If FileExists(vSaveAS) = True Then
            'Dim flen As Long
            'flen = FileLen(vSaveAS)
            UrlCompleted vUrl, 0
            RaiseEvent ProgressError(0, vUrl, "文件已经存在")
            DoEvents
            GoTo ReEnter
        End If
        vSaveAS = mDirectory & TmpFileName(vUrl)
        If FileExists(vSaveAS) Then Kill vSaveAS
        If mTask.IsJpgBook Then
            pUrl = GetJpgBookPageUrl(mRootUrl, vUrl, HttpHeaderGetField(mHeader, "Cookie"), , mJPGBookQuality)
            If pUrl = "" Then
                RaiseEvent StatusChange("错误：不能得到" & vUrl & "的JPG大图链接!")
                mTaskDone = True
                mResult = False
                ImBusy = False
                Exit Sub
            End If
        End If
        mHttpConnection.Connect pUrl, "GET", vSaveAS, pHeader
        ImBusy = False
End Sub

'Private Sub StartDownload2()
''    mStartDownload = False
''    taskTimer.Interval = cst_timer_interval
''    mTimerEnabled = True
''    mTaskDone = False
''    mResult = False
''    mAborting = False
''    mBusy = False
''    Exit Sub
''
'    Dim vUrl As String
'    Dim vSaveAS As String
'    Dim vFileLen As Long
'    vUrl = mUrls.GetNextUrl
'
'    Dim pUrl As String
'    Dim pSaveAs As String
'    Dim pHeader As String
'    mResult = True
'
'    RaiseEvent StatusChange("正在下载...")
'Static ImBusy As Boolean
'On Error GoTo ErrorDownload
'    Do Until vUrl = ""
'        If mAborting Then Exit Sub
'        Debug.Print mTask.ToString & vUrl & "[TRY]"
'        If mTask.IsJpgBook Then
'            If mRenameJPGBook Then
'                vSaveAS = mDirectory & Left$(vUrl, Len(vUrl) - 3) & "pdg"
'            Else
'                vSaveAS = mDirectory & vUrl
'            End If
'            pHeader = ""
'        Else
'            pUrl = mRootUrl & vUrl
'            vSaveAS = mDirectory & vUrl
'            pHeader = mHeader
'        End If
'        If FileExists(vSaveAS) = True Then
'            'Dim flen As Long
'            'flen = FileLen(vSaveAS)
'            UrlCompleted vUrl, 0
'            RaiseEvent ProgressError(0, vUrl, "文件已经存在")
'            mTimerAction = True ''DoEvents
'            GoTo Continue
'        End If
'        vSaveAS = mDirectory & TmpFileName(vUrl)
'        If FileExists(vSaveAS) Then Kill vSaveAS
'
'        If mTask.IsJpgBook Then
'            If mAborting Then Exit Sub
'            DoEvents
'            pUrl = GetJpgBookPageUrl(mRootUrl, vUrl, HttpHeaderGetField(mHeader, "Cookie"), , mJPGBookQuality)
'            If pUrl = "" Then
'                RaiseEvent StatusChange("错误：不能得到" & vUrl & "的JPG大图链接!")
'                mResult = False
'                Exit Do
'            End If
'        End If
'
'        Do Until ImBusy = False
'            DoEvents
'        Loop
'        ImBusy = True
'        Do Until mHttpConnection.IsFree = True
'            Debug.Print mTask.ToString & "Waiting..."
'            DoEvents
'        Loop
'
'        If mTask.IsJpgBook Then
'            RaiseEvent StatusChange("Sleeping...")
'            WinBase.Sleep 2000
'        End If
'
'        ImBusy = False
'
'        If mAborting Then Exit Sub
'        mHttpConnection.Connect pUrl, "GET", vSaveAS, pHeader
'    GoTo Continue
'ErrorDownload:
'        RaiseEvent ProgressError(0, vUrl, Err.Description)
'Continue:
'        Do Until mTimerAction = True Or mAborting
'            Debug.Print mTask.ToString & "Waiting..."
'            DoEvents
'        Loop
'        mTimerAction = False
'
'        vUrl = mUrls.GetNextUrl
'        'imbusy = False
'    Loop
'
'
'    TaskComplete
'
'
'End Sub

Public Sub TaskComplete()
    If Not mTaskDone Then Exit Sub
    
    Dim I As Long
    RaiseEvent StatusChange("等待所有下载进程结束")
    If Not mHttpConnection Is Nothing Then mHttpConnection.AbortAll
    For I = 1 To mThreads
        If Not mHttpConnection.Threads(I) Is Nothing Then
            If mHttpConnection.Threads(I).IsFree = False Then
                mTimerAction = TIMER_ACTION_TASK_COMPLETED
                Exit Sub
            End If
        End If
    Next
    mTimerAction = TIMER_ACTION_NOTHING
    Set mHttpConnection = Nothing
    'Set MTimer = Nothing
    Set mTask = Nothing
    taskTimer.Interval = 0
    Set taskTimer = Nothing
    
    'mTaskDone = False
    If mResult Then
        RaiseEvent StatusChange("完成下载")
        RaiseEvent Finish
    Else
        RaiseEvent StatusChange("下载出现错误")
        RaiseEvent OnError
    End If
End Sub

Private Sub Class_Initialize()
    mThreads = 1
    mJPGBookQuality = 1
   'mLastStatus =  0
   mTimerAction = TIMER_ACTION_NOTHING
   StartTimer
    'Set taskTimer = New CTimer
    
    
    
End Sub
Public Sub TimerMe()
'    If mStartDownload Then
'        mStartDownload = False
'        StartDownload2
'    Else
'        mTimerAction = True
'    End If
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



Public Function Download(ByRef vTask As CTask) As Boolean
    'settimer 0
    mReChecked = False
    mResult = False
    mTaskDone = False
    mAborting = False
    mStep = 0
    mBusy = False
    
    Set mTask = vTask
    Set mTester = New CBookTester
    
'    mTaskName = "《" & mTask.bookInfo(SSF_Title) & "》"
'    If mTask.bookInfo(SSF_SSID) Then mTaskName = "[" & mTask.bookInfo(SSF_SSID) & "]" & mTaskName
    mHeader = vTask.bookInfo(SSF_HEADER)
    
    If mTask.IsJpgBook Then
        mRootUrl = vTask.bookInfo(SSF_JPGURL)
    Else
        mRootUrl = vTask.bookInfo(SSF_URL)
    End If
    mRootUrl = BuildPath(mRootUrl, , lnpsUnix)
    mDirectory = BuildPath(vTask.Directory)
    
'    If mTask.IsJpgBook Then
'
'    Else
        RaiseEvent StatusChange("正在测试...")
        
        'If mTask.IsJpgBook = True Then
        mTester.Test vTask

        
    'End If
End Function


Public Sub Abort()
    mAborting = True
    'settimer 0
    Dim I As Long
    
    On Error Resume Next
    Dim vFile As String
    For I = 1 To mThreads
        If mHttpConnection.Threads(I) Is Nothing Then GoTo Continue
        'If mHttpConnection.Threads(i).IsFree Then GoTo Continue
        vFile = mHttpConnection.Threads(I).Destination
        If vFile = "" Then GoTo Continue
        Err.Clear
        Kill vFile
        If Err.Number = 0 Then RaiseEvent ProgressError(I, vFile, "删除正在下载的文件")
        'Err.Clear
        
        Debug.Print mTask.ToString & "Kill " & vFile
Continue:
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

    Debug.Print mTask.ToString & vUrl & "[Completed]"

    If mStep = -1 Then
        GetTextPdg2
        Exit Sub
    End If
    
   
    Dim vCount As Long
    If mIsTextPdg Then
        vCount = MSSReader.TextPdgCount(vUrl)
    ElseIf Left$(vUrl, 1) = "0" Then
        vCount = 1
    End If
    RaiseEvent ProgressDone(vId, vUrl, vCount)
    

End Sub

Private Sub Class_Terminate()
   ' taskTimer.Interval = 0
    'Set taskTimer = Nothing
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
        'settimer 0
        mStep = 0
        mTaskDone = True
        mResult = False
        TaskComplete
        Exit Sub
    End If
    Dim vUrl As String
    vUrl = ThreadUrlBaseName(vId)
    Debug.Print mTask.ToString & vUrl & "[Error]"
    mUrls.ReportFail vUrl
    RaiseEvent ProgressChange(vId, vUrl, 0, 0)
    RaiseEvent ProgressError(vId, vUrl, vDesc)
    'DownloadNextUrl
End Sub

Private Sub mHttpConnection_ProgressChange(ByVal vId As Long, ByVal vMin As Long, ByVal vMax As Long)
    RaiseEvent ProgressChange(vId, ThreadUrlBaseName(vId), vMin, vMax)
End Sub

Private Sub mHttpConnection_ResponseStart(ByVal vId As Long, ByVal vStatus As Long, ByVal vText As String)
    Dim vUrl As String
    vUrl = ThreadUrlBaseName(vId)
    Debug.Print mTask.ToString & vUrl & "[Start...]"
    If vStatus = 200 Then
        RaiseEvent ProgressChange(vId, vUrl, 0, 0)
    Else
        mUrls.ReportFail vUrl
        RaiseEvent ProgressError(vId, vUrl, vText & "(" & vStatus & ")")
    End If
    
End Sub

Private Sub mTester_StatusChange(vText As String)
    RaiseEvent StatusChange("正在测试: " & vText)
End Sub

Private Function TmpFileName(ByVal vForm As String, Optional vReverseMode As Boolean = False)
    If vReverseMode Then
        TmpFileName = Left$(vForm, Len(vForm) - 5)
    Else
        TmpFileName = vForm & ".part"
    End If
End Function

'Private Sub GetUrl(ByVal vUrl As String)
'    Debug.Print vUrl & "[TRY]"
'    'mBusy = True
'    On Error GoTo GetUrl_Error
'    Static lastURL As String
'    If vUrl = "" Then vUrl = lastURL
'    lastURL = vUrl
'    Dim vSaveAS As String
'    If mTask.IsJpgBook And mRenameJPGBook Then
'        vSaveAS = mDirectory & Left$(vUrl, Len(vUrl) - 3) & "pdg"
'    Else
'        vSaveAS = mDirectory & vUrl
'    End If
'
'
'    If FileExists(vSaveAS) = True Then
'        Dim vFreeId As Long
'        Dim vFileLen As Long
'        vFreeId = mHttpConnection.FreeId
'        vFileLen = FileLen(vSaveAS)
'        RaiseEvent ProgressChange(vFreeId, vUrl, vFileLen, vFileLen)
'        UrlCompleted vUrl, vFreeId
'        mBusy = False
'        DownloadNextUrl
'        Exit Sub
'    End If
'
'    vSaveAS = mDirectory & TmpFileName(vUrl)
'    If FileExists(vSaveAS) Then Kill vSaveAS
'
'    If mTask.IsJpgBook Then
'        Dim pUrl As String
'        pUrl = GetJpgBookPageUrl(mRootUrl, vUrl, HttpHeaderGetField(mHeader, "Cookie"), , mJPGBookQuality)
'        If pUrl = "" Then
'            mResult = False
'            mTaskDone = True
'            Exit Sub
'        End If
'        mHttpConnection.Connect pUrl, "GET", vSaveAS
'    Else
'        mHttpConnection.Connect BuildPath(mRootUrl, vUrl, lnpsUnix), "GET", vSaveAS, mHeader
'    End If
'
'    Exit Sub
'GetUrl_Error:
'    'WaitConnections
'End Sub



Private Sub GetTextPdg2()
    mStep = 0
    Set mUrls = New CBookURL
    mUrls.SkipCatalogPages = mSkipCatalogPages
    mUrls.SkipContentPage = mSkipContentPages
    mUrls.SkipStaticPages = mSkipStaticPages
    mUrls.SetTextPdg BuildPath(mDirectory, "InfoRule.dat")
    mTimerAction = TIMER_ACTION_TASK_START
    'StartDownload
    'settimer cst_timer_interval
End Sub
Private Sub mTester_TestCompleted()
    RaiseEvent StatusChange("测试完成。")
    TestComplete
End Sub









