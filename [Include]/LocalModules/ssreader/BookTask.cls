VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTask"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
#Const ZLIB = 1
#If afCTaskBrief = 0 Then

Private mNotify As ITaskNotify
Private mDownloader As CTaskRunner
Private m_bStopNow As Boolean

Private mStep As Long


Private m_LonStartPage As Long

#End If

Private m_LonStaticCount As Long
Public Status As SSLIBTaskStatus
Private m_LonFilesCount As Long
'Private m_StrTitle As String
'Private m_StrAuthor As String
'Private m_StrSSID As String
'Private m_StrPublisher As String
'Private m_StrHttpHeader As String
'Private m_StrSavedIN As String

'Private m_StrRootURL As String
'Private m_StrCurrentFile As String

Private m_strFolder As String
'Private m_StrPublishedDate As String
'Private m_StrPagesCount As String
'Private m_StrAdditionalText As String

Private Const CSTTaskConfigFilename As String = "GetSSLib.ini"
Private Const CSTDefaultNameFormat As String = "%s_%t_%a"
Private m_StrTaskID As String
Private m_bChanged As Boolean
Private mBookInfo As CBookInfo
Private mName As String
Private mDownloadId As Long
Private mIsJPGBook As Boolean
Private mGetJPGPage As Boolean

Public Property Get GetJPGPage() As Boolean
    GetJPGPage = mGetJPGPage
End Property

Public Property Let GetJPGPage(ByVal bValue As Boolean)
    mGetJPGPage = bValue
End Property

Public Property Get IsJpgBook() As Boolean
    If mBookInfo(SSF_ISJPGBOOK) = "1" Then IsJpgBook = True
    If mGetJPGPage Then IsJpgBook = True
End Property

Public Property Let IsJpgBook(ByVal bValue As Boolean)
    If bValue Then
        mBookInfo(SSF_ISJPGBOOK) = "1"
    Else
        mBookInfo(SSF_ISJPGBOOK) = vbNullString
    End If
End Property


Public Property Get Name() As String
    Name = mName
End Property

Public Property Let Name(ByVal StrValue As String)
    If StrComp(StrValue, mName, vbTextCompare) <> 0 Then
        mName = StrValue
        UpdateFolder
    End If
End Property


Public Property Get bookInfo() As CBookInfo
   Set bookInfo = mBookInfo
End Property

Public Property Let bookInfo(ByVal cboValue As CBookInfo)
   If Not cboValue Is Nothing Then Set mBookInfo = cboValue
End Property

Public Property Set bookInfo(ByRef vBookInfo As CBookInfo)
    If Not vBookInfo Is Nothing Then Set mBookInfo = vBookInfo
End Property

Public Property Get Changed() As Boolean
    Changed = m_bChanged
End Property

Public Property Let Changed(ByVal bValue As Boolean)
    m_bChanged = bValue
End Property


Public Property Get DownloadId() As Long
    DownloadId = mDownloadId
End Property

Public Property Let DownloadId(ByVal vId As Long)
    mDownloadId = vId
End Property
#If afCTaskBrief = 0 Then
    Public Property Get Downloader() As CTaskRunner
        Set Downloader = mDownloader
    End Property
#End If

Public Sub Reset()
    m_bChanged = True
    #If afCTaskBrief = 0 Then
        mStep = 0
        Set mDownloader = New CDownloadEx
    
        Set mNotify = Nothing
    #End If
End Sub

Private Sub Class_Initialize()
On Error GoTo ErrorInit
Set mBookInfo = New CBookInfo
mBookInfo(SSF_Downloader) = App.ProductName
m_bChanged = True
#If afCTaskBrief = 0 Then
    mStep = 0
    Set mDownloader = New CTaskRunner
#End If
Exit Sub
ErrorInit:
Err.Raise Err.Number, Err.Description
'Set mHelper = New CHttpConnection
End Sub
Public Property Get PagesCount() As Long
    'PagesCount = mBookInfo.Field(SSF_PAGESCOUNT)
    PagesCount = StringToLong(mBookInfo(SSF_PagesCount))
End Property

Public Property Get taskId() As String
    taskId = m_StrTaskID
End Property

Public Property Let taskId(ByVal StringValue As String)
    m_StrTaskID = StringValue
End Property

Public Property Get FilesCount() As Long
    If m_LonFilesCount = 0 Then FilesCount = m_LonStaticCount Else FilesCount = m_LonFilesCount
End Property

Public Property Let FilesCount(ByVal vCount As Long)
    m_LonFilesCount = vCount
'    If mDownloader Is Nothing Then Exit Property
'    mDownloader.Files = vCount
End Property

Public Function InitForDownload() As Boolean
    
    If Not CheckFolder(mBookInfo(SSF_SAVEDIN)) Then
        MsgBox "不能创建" & mBookInfo(SSF_SAVEDIN) & vbCrLf & "请检查配置", vbCritical, "Task.StartDownload"
        GoTo ErrorChecked
    End If
    
    'UpdateFolder
    If m_strFolder = vbNullString Then UpdateFolder
    If Not CheckFolder(m_strFolder) Then
        MsgBox "不能创建" & m_strFolder & vbCrLf & "请检查配置", vbCritical, "Task.StartDownload"
        GoTo ErrorChecked
    End If
    
    'pCheckPassed = True
    
    
    'm_LonStartPage = StringToLong(mBookInfo(SSF_StartPage))
    InitForDownload = True
    Exit Function

ErrorChecked:
    InitForDownload = False
End Function

#If afCTaskBrief = 0 Then

Private Sub Class_Terminate()
'    Set mHelper = Nothing
'    Set mNotify = Nothing
'    Set mDownloader = Nothing
End Sub

'Public Property Get Connections(ByVal vIndex As Long) As CDownloadProgress
'    On Error Resume Next
'    Set Connections = mDownloader.ThreadProgress(vIndex)
'End Property
'
'Public Property Get ConnectionsCount() As Long
'    ConnectionsCount = mDownloader.ThreadCount
'End Property
'
'Public Property Let ConnectionsCount(ByVal intValue As Long)
'    mDownloader.ThreadCount = intValue
'    'm_intConnectionsCount = intValue
'End Property
'Public Property Let StopNow(ByVal bValue As Boolean)
'    m_bStopNow = bValue
'End Property




Public Sub Init(Optional ByRef Notify As ITaskNotify = Nothing)
    Set mNotify = Notify
End Sub


'
'Private Sub GetTextPdg(ByRef vFilename As String)
'    Dim sUrls() As String
'    Dim iCount As Long
'    iCount = SSLIB_ParseInfoRule(vFilename, sUrls)
'    Dim i As Long
'    For i = 0 To iCount - 1
'        If GetUrl(sUrls(i), , False) Then
'            m_LonFilesCount = m_LonFilesCount + TextPdgCount(sUrls(i))
'        End If
'
'        If m_bStopNow Then StopMe: Exit Sub
'    Next
'    Me.Status = STS_COMPLETE
'    Call NotifyStatus
'End Sub


'Public Sub StartDownload()
'
'm_bChanged = True
'
'   Dim pHeader As CHttpHeader
'   Set pHeader = New CHttpHeader
'   pHeader.Init mBookInfo(SSF_HEADER)
'
'   If mBookInfo(SSF_URL) = vbNULLSTRING Then
'        Dim host As String
'        Dim action As String
'        host = pHeader.GetField("host")
'        action = pHeader.HttpGet
'        If host <> vbNULLSTRING And action <> vbNULLSTRING Then
'            mBookInfo(SSF_URL) = "http://" & host & action
'
'        End If
'   End If
'
'   pHeader.DeleteField (vbNULLSTRING)
'   mBookInfo(SSF_HEADER) = pHeader.HeaderString
''
''    'Dim pCheckPassed As Boolean
''    'pCheckPassed = False
''
''
''
''    If Not CheckConnection(mBookInfo(SSF_URL), "test.pdg", mBookInfo(SSF_HEADER)) Then
''        GoTo ErrorChecked
''    End If
''
'    If Not CheckFolder(mBookInfo(SSF_SAVEDIN)) Then
'        MsgBox "不能创建" & mBookInfo(SSF_SAVEDIN) & vbCrLf & "请检查配置", vbCritical, "Task.StartDownload"
'        GoTo ErrorChecked
'    End If
'
'    'UpdateFolder
'    If m_strFolder = vbNULLSTRING Then UpdateFolder
'    If Not CheckFolder(m_strFolder) Then
'        MsgBox "不能创建" & m_strFolder & vbCrLf & "请检查配置", vbCritical, "Task.StartDownload"
'        GoTo ErrorChecked
'    End If
'
'    'pCheckPassed = True
'
'
'    m_LonStartPage = StringToLong(mBookInfo(SSF_StartPage))
'
'    Me.Status = STS_START
'    Me.StopNow = False
'    m_LonFilesCount = 0
'    Call NotifyTaskStatus
'
'    Set mDownloader = New CTaskRunner
'    mDownloader.Execute mNotify, Me
'
'    Exit Sub
'ErrorChecked:
'    Me.Status = STS_ERRORS
'    Exit Sub
'End Sub
'Private Sub StartDownload2()
'
'
'    If mStep <= 1 Then
'    'dat
'        GetUrl "bookinfo.dat", , False
'        GetUrl "BookContents.dat", , False
'        GetUrl "BookContents.xml", , False
'        GetUrl "catalog.dat", , False
'        GetUrl "Catalog.dat", , False
'        If m_bStopNow Then StopMe: Exit Sub
'        mStep = 2
'    End If
'    Call WriteBookInfo
'
'    If GetUrl("InfoRule.dat", , False) Then
'        GetTextPdg m_strFolder & "\InfoRule.dat"
'        Exit Sub
'    End If
'
'    GetRangeUrl "bac", 3, ".pdg", , False
'    If m_bStopNow Then StopMe: Exit Sub
'    GetRangeUrl "att", 3, ".pdg", , False
'    If m_bStopNow Then StopMe: Exit Sub
'
'    If mStep <= 2 Then
'    '!...pdg
'        GetRangeUrl "!", 5, ".pdg", , False
'        If m_bStopNow Then StopMe: Exit Sub
'        mStep = 3
'    End If
'
'    If mStep <= 3 Then
'    '...pdg
'        'If Not GetURL("000001.pdg", False, False) Then
'
'        'Else
'            Dim p As Long
'            p = GetStartPoint()
'            If p > 1 Then m_LonStartPage = p
'            GetRangeUrl vbNULLSTRING, 6, ".pdg", p
'        'End If
'        If m_bStopNow Then StopMe: Exit Sub
'        mStep = 4
'    End If
'
'    If mStep <= 4 Then
'    'cov...pdg
'        GetRangeUrl "cov", 3, ".pdg", , False
'        If m_bStopNow Then StopMe: Exit Sub
'        mStep = 5
'    End If
'
'    If mStep <= 5 Then
'    'fow...pdg
'        GetRangeUrl "fow", 3, ".pdg", , False
'        If m_bStopNow Then StopMe: Exit Sub
'        mStep = 6
'    End If
'
'    If mStep <= 6 Then
'    'bok...pdg
'        GetRangeUrl "bok", 3, ".pdg", , False
'        If m_bStopNow Then StopMe: Exit Sub
'        mStep = 7
'    End If
'
'    If mStep <= 7 Then
'    'leg...pdg
'        GetRangeUrl "leg", 3, ".pdg", , False
'        If m_bStopNow Then StopMe: Exit Sub
'        mStep = 8
'    End If
'
'    Me.Status = STS_COMPLETE
'
'    mBookInfo(SSF_DownloadDate) = DateTime.Date$ & " " & DateTime.Time$
'
'    Me.AutoSave
'
'If Not mNotify Is Nothing Then mNotify.TaskStatusChange Me
'
'End Sub
'Public Sub StopMe()
'    Me.Status = STS_PAUSE
'    m_bStopNow = False
'    'Call NotifyStatus
'    If mDownloader Is Nothing Then Exit Sub
'    mDownloader.Abort
'End Sub

'Private Function BuildRangeString(vPre As String, vNumLength As Long, vSuf As String, Optional vStart As Long = 1)
'    Dim s As String
'    Dim l As Long
'
'    s = CStr(vStart)
'    l = vNumLength - Len(s)
'    If (l > 0) Then
'        BuildRangeString = vPre & String$(l, "0") & s & vSuf
'    Else
'        BuildRangeString = vPre & s & vSuf
'    End If
'
'End Function
'Private Function GetRangeUrl(vPre As String, vNumLength As Long, vSuf As String, Optional vStart As Long = 1, Optional fCount As Boolean = True)
'    Dim sUrl As String
'    Dim i As Long
'    i = vStart - 1
'    Do
'        If m_bStopNow Then Exit Function
'        i = i + 1
'        sUrl = BuildRangeString(vPre, vNumLength, vSuf, i)
'        If m_bStopNow Then Exit Function
'    Loop While GetUrl(sUrl, , fCount)
'End Function
'
'Private Function download(ByRef vUrl As String, Optional vOverWrite As Boolean = False) As Boolean
'    DoEvents
'    Dim saveAs As String
'    saveAs = m_strFolder & "\" & vUrl
'    Dim URL As String
'    URL = mBookInfo(SSF_URL) & vUrl
'    If Not mDownloader.download(saveAs, URL, mBookInfo(SSF_HEADER), vOverWrite) Then
'        download = mDownloader.download(saveAs, URL, mBookInfo(SSF_HEADER), vOverWrite)
'    Else
'        download = True
'    End If
'
'End Function
'
'Private Function GetUrl(vUrl As String, Optional ByVal vOverWrite As Boolean = False, Optional ByVal fCount As Boolean = True) As Boolean
' On Error GoTo ErrorDownload
'    'm_LonCurrentLength = 0
'    'm_LonCurrentPos = 0
'    'm_StrCurrentFile = vURL
'    Static lastURL As String
'
''    If mDownloader.ErrorsCount > 100 Then
''        MsgBox "错误次数多于100次", vbCritical
''        Exit Function
''    End If
'
'    If download(vUrl, vOverWrite) Then
'        If fCount Then m_LonFilesCount = m_LonFilesCount + 1
'        GetUrl = True
'    Else
'        GetUrl = False
'    End If
'    Call NotifyStatus
'
'
'
'
'    Exit Function
'
'ErrorDownload:
'    On Error Resume Next
'    Err.Clear
'    GetUrl = False
'    'Kill saveAs
'    Call NotifyStatus
'End Function

Public Sub Restart()
    mStep = 0
    m_LonFilesCount = 0
End Sub

'
'Private Sub NotifyTaskStatus()
'    If Not mNotify Is Nothing Then mNotify.TaskStatusChange mDownloader
'End Sub
'
'Private Sub NotifyDownloadStatus()
'    If Not mNotify Is Nothing Then mNotify.DownloadStatusChange mDownloader
'End Sub
'
'Private Sub NotifyStatus()
'If Not mNotify Is Nothing Then
'    mNotify.TaskStatusChange mDownloader
'    mNotify.DownloadStatusChange mDownloader
'End If
'End Sub

'Private Sub mDownloader_ProgressChanged(ByVal vId As Long, Progress As CDownloadProgress)
'    Call NotifyDownloadStatus
'End Sub
'
'Private Function GetDefaultStartPoint() As Long
'    If download(BuildRangeString(vbNULLSTRING, 6, ".pdg", 1)) Then
'        GetDefaultStartPoint = True
'    Else
'        GetDefaultStartPoint = False
'    End If
'End Function
'
'Private Function GetStartPoint() As Long
'    If GetDefaultStartPoint Then GetStartPoint = 1: Exit Function
'    mDownloader.ClearErrors
'    'If m_LonStartPage > 0 Then GetStartPoint = m_LonStartPage: Exit Function
'    Dim u As Long
'    Dim i As Long
'    Dim URL As String
'    If m_LonStartPage > 0 Then i = m_LonStartPage - 1
'    If mBookInfo(SSF_PagesCount) <> vbNULLSTRING Then u = StringToLong(mBookInfo(SSF_PagesCount))
'    If u < 1 Then u = 9999
'    Do
'        DoEvents
'        If i > u Then
'            i = 1
'            Exit Do
'        End If
'        i = i + 1
'        URL = BuildRangeString(vbNULLSTRING, 6, ".pdg", i)
'        If m_bStopNow Then
'            StopMe
'            GetStartPoint = 1
'            Exit Function
'        End If
'    Loop Until download(URL, False)
'    GetStartPoint = i
'    If i > 1 Then mBookInfo(SSF_StartPage) = i
'End Function
'
'Private Function CheckConnection(ByVal vRootUrl As String, ByVal vFilename As String, ByVal vHeader As String) As Boolean
'    Dim c As CHTTPConnection
'    Set c = New CHTTPConnection
'    c.URL = vRootUrl & vFilename
'    c.Header = vHeader
'    c.Connect
'    Dim ret As Long
'    ret = c.Response.HttpStatus
'    If ret = 200 Or ret = 404 Then CheckConnection = True Else CheckConnection = False
'End Function


#End If



Private Function FileExist(ByRef sFilename As String) As Boolean
On Error GoTo Error_test
GetAttr sFilename
FileExist = True
    
Exit Function
'FileExist = True
Error_test:
FileExist = False

End Function



Private Function CheckFolder(ByRef PathName As String) As Boolean
    CheckFolder = True
    On Error GoTo Errorfolder
    If FileExist(PathName) = False Then MkDir PathName
    Exit Function
Errorfolder:
    CheckFolder = False
End Function

'CSEH: ErrMsgBox
Private Sub WriteBookInfo()
        '<EhHeader>
        On Error GoTo WriteBookInfo_Err
        '</EhHeader>

100     mBookInfo.SaveToFile m_strFolder & "\bookinfo.dat", Nothing, , True
        
        
        '<EhFooter>
        Exit Sub

WriteBookInfo_Err:
        MsgBox Err.Description & vbCrLf & _
               "in GetSSLib.CTask.WriteBookInfo " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        '</EhFooter>
End Sub



Public Sub PersistTo(ByRef iniHnd As CLiNInI, ByRef vSection As String)
    
    
'        iniHnd.SaveSetting vSection, "Title", m_StrTitle
'        iniHnd.SaveSetting vSection, "Author", m_StrAuthor
'        iniHnd.SaveSetting vSection, "SSId", m_StrSSID
'        iniHnd.SaveSetting vSection, "Publisher", m_StrPublisher
'        iniHnd.SaveSetting vSection, "Header", m_StrHttpHeader
'        iniHnd.SaveSetting vSection, "URL", m_StrRootURL
'        iniHnd.SaveSetting vSection, "SaveIn", m_StrSavedIN
'        iniHnd.SaveSetting vSection, "PagesCount", m_StrPagesCount
'        iniHnd.SaveSetting vSection, "PDate", m_StrPublishedDate
'        iniHnd.SaveSetting vSection, "AddInfo", m_StrAdditionalText
        iniHnd.SaveSetting vSection, "Status", Status
        iniHnd.SaveSetting vSection, "Downloaded", m_LonFilesCount
        mBookInfo.SaveToFile vbNullString, iniHnd, vSection
        'iniHnd.SaveSetting vSection, "StartPage", m_LonStartPage
        'm_LonStartPage = CLng(iniHnd.GetSetting(vSection, "StartPage"))
End Sub

Public Sub LoadFrom(ByRef iniHnd As CLiNInI, ByRef vSection As String)
On Error Resume Next
    Dim Title As String
    Title = iniHnd.GetSetting(vSection, "Title")
    If Title <> vbNullString Then
        mBookInfo(SSF_Title) = Title
        mBookInfo(SSF_AUTHOR) = iniHnd.GetSetting(vSection, "Author")
        mBookInfo(SSF_SSID) = iniHnd.GetSetting(vSection, "SSId")
        mBookInfo(SSF_Publisher) = iniHnd.GetSetting(vSection, "Publisher")
        mBookInfo(SSF_HEADER) = iniHnd.GetSetting(vSection, "Header")
        mBookInfo(SSF_URL) = iniHnd.GetSetting(vSection, "URL")
        mBookInfo(SSF_SAVEDIN) = iniHnd.GetSetting(vSection, "SaveIn")
        mBookInfo(SSF_PagesCount) = iniHnd.GetSetting(vSection, "PagesCount")
        mBookInfo(SSF_PublishDate) = iniHnd.GetSetting(vSection, "PDate")
        mBookInfo(SSF_Comments) = iniHnd.GetSetting(vSection, "AddInfo")
    Else
        mBookInfo.LoadFromFile vbNullString, iniHnd, vSection
    End If
            
    If mBookInfo(SSF_HTMLContent) <> vbNullString Then mIsJPGBook = True
        

        Status = CLng(iniHnd.GetSetting(vSection, "Status"))
        If Status = STS_START Then Status = STS_PENDING
        m_LonFilesCount = CLng(iniHnd.GetSetting(vSection, "Downloaded"))
        'm_LonStartPage = CLng(iniHnd.GetSetting(vSection, "StartPage"))
End Sub

Public Property Get Directory() As String
    Directory = m_strFolder
End Property

Public Property Let Directory(ByVal vPath As String)

    'If vPath = m_strFolder Then Exit Property
    
    Dim pFolder As String
    pFolder = GetParentFolderName(vPath)
    Dim pName As String
    pName = GetBaseName(vPath)
    
    mBookInfo(SSF_SAVEDIN) = pFolder
    UpdateFolder
    If vPath <> m_strFolder Then
        mName = pName
        UpdateFolder
    End If
    'm_strFolder = vPath
    
End Property

Public Sub AutoLoad()
    Dim taskconfig As CLiNInI
    Dim taskDir As String
    Set taskconfig = New CLiNInI
    taskDir = Me.Directory
    taskconfig.File = BuildPath(taskDir, CSTTaskConfigFilename)
    Me.LoadFrom taskconfig, "TaskInfo"
    Set taskconfig = Nothing
End Sub

Public Sub AutoSave()
    Dim taskconfig As CLiNInI
    Dim taskDir As String
    Set taskconfig = New CLiNInI
    
    If Not CheckFolder(mBookInfo(SSF_SAVEDIN)) Then
        MsgBox "不能创建" & mBookInfo(SSF_SAVEDIN) & vbCrLf & "请检查配置", vbCritical, "Task.AutoSave"
        
        Exit Sub
    End If
    
    If m_strFolder = vbNullString Then UpdateFolder
'    If Not fSolidDirectory Then UpdateFolder
    If Not CheckFolder(m_strFolder) Then
        MsgBox "不能创建" & m_strFolder & vbCrLf & "请检查配置", vbCritical, "Task.AutoSave"
        Exit Sub
    End If
    
    
    taskDir = Me.Directory
    Me.PersistTo taskconfig, "TaskInfo"
    taskconfig.WriteTo BuildPath(taskDir, CSTTaskConfigFilename)
    Set taskconfig = Nothing
    WriteBookInfo
    Me.Changed = False
End Sub

Public Function ToString() As String
    
    Dim pTitle As String
    If mBookInfo(SSF_Title) <> vbNullString Then pTitle = "《" & mBookInfo(SSF_Title) & "》"
    If mBookInfo(SSF_SSID) <> vbNullString Then pTitle = "[" & mBookInfo(SSF_SSID) & "]" & pTitle
    ToString = pTitle
End Function

Public Sub UpdateFolder()
    Dim OldFolder As String
    Dim NewFolder As String
    OldFolder = m_strFolder
    If mName = vbNullString Then
        NewFolder = MSSReader.pdgformat(mBookInfo, CSTDefaultNameFormat)
        NewFolder = BuildPath(mBookInfo(SSF_SAVEDIN)) & NewFolder
    Else
        NewFolder = BuildPath(mBookInfo(SSF_SAVEDIN), mName)
    End If
    m_strFolder = ReplaceInvalidChars(NewFolder, " ")
    
    If OldFolder <> vbNullString And OldFolder <> m_strFolder Then
        If FolderExists(OldFolder) And FolderExists(m_strFolder) = False Then
            If (MsgBox("将" & vbCrLf & OldFolder & vbCrLf & "重命名为" & vbCrLf & m_strFolder & "？", vbInformation + vbYesNo + vbDefaultButton2, "Task.UpdateFolder") = vbYes) Then
                MoveFile OldFolder, m_strFolder
            End If
        End If
    End If
End Sub
'Private Sub IDownloadClient_ProgressChanged(ByVal vId As Long, vProgress As CDownloadProgress)
'    Call NotifyDownloadStatus
'End Sub

'Public Sub TaskComplete()
'    'Set mDownloader = Nothing
'End Sub
