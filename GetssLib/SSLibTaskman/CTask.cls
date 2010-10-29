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
Private mNotify As ITaskNotify
Private mStep As Long
Public Status As SSLIBTaskStatus

'Private m_StrTitle As String
'Private m_StrAuthor As String
'Private m_StrSSID As String
'Private m_StrPublisher As String
'Private m_StrHttpHeader As String
'Private m_StrSavedIN As String
Private m_LonFilesCount As Long
Private m_StrTaskID As String
'Private m_StrRootURL As String
'Private m_StrCurrentFile As String
Private m_bStopNow As Boolean
'Private m_LonCurrentPos As Long
'Private m_LonCurrentLength As Long
Private m_LonStaticCount As Long
Private m_strFolder As String
'Private m_StrPublishedDate As String
'Private m_StrPagesCount As String
'Private m_StrAdditionalText As String
Private m_LonStartPage As Long


Private Const CSTTaskConfigFilename As String = "GetSSLib.ini"

Private WithEvents mDownloader As CDownloadEx
Attribute mDownloader.VB_VarHelpID = -1
Private m_bChanged As Boolean
Private mBookInfo As CBookInfo


Public Property Get bookInfo() As CBookInfo
   Set bookInfo = mBookInfo
End Property

Public Property Let bookInfo(ByVal cboValue As CBookInfo)
   If Not cboValue Is Nothing Then Set mBookInfo = cboValue
End Property


Public Property Get Changed() As Boolean
    Changed = m_bChanged
End Property

Public Property Let Changed(ByVal bValue As Boolean)
    m_bChanged = bValue
End Property


Public Property Get Downloader() As CDownloadEx
    Set Downloader = mDownloader
End Property

Public Sub Reset()
    m_bChanged = True
    mStep = 0
    Set mDownloader = New CDownloadEx

    Set mNotify = Nothing
End Sub

Private Sub Class_Initialize()
On Error GoTo ErrorInit
Set mBookInfo = New CBookInfo
mBookInfo(SSF_Downloader) = App.ProductName
m_bChanged = True
mStep = 0
Set mDownloader = New CDownloadEx
Exit Sub
ErrorInit:
Err.Raise Err.Number, Err.Description
'Set mHelper = New CHttpConnection
End Sub

Private Sub Class_Terminate()
'    Set mHelper = Nothing
    Set mNotify = Nothing
    Set mDownloader = Nothing
End Sub

Public Property Get Connections(ByVal vIndex As Long) As CDownloadProgress
    On Error Resume Next
    Set Connections = mDownloader.Progress(vIndex)
End Property

Public Property Get ConnectionsCount() As Long
    ConnectionsCount = mDownloader.Count
End Property

Public Property Let ConnectionsCount(ByVal intValue As Long)
    mDownloader.MaxConnection = intValue
    'm_intConnectionsCount = intValue
End Property

'Public Property Get AdditionalText() As String
'    AdditionalText = m_StrAdditionalText
'
'End Property
'
'Public Property Let AdditionalText(ByVal StrValue As String)
'    m_StrAdditionalText = StrValue
'
'    'mBookInfo.Field(SSF_COMMENTS) = StrValue
'End Property
'
Public Property Get PagesCount() As Long
    'PagesCount = mBookInfo.Field(SSF_PAGESCOUNT)
    PagesCount = StringToLong(mBookInfo(SSF_PagesCount))
End Property

'Public Property Let PagesCount(ByVal LonValue As String)
'
'    m_StrPagesCount = LonValue
'End Property
'
'Public Property Get PublishedDate() As String
'    PublishedDate = m_StrPublishedDate
'
'End Property
'
'Public Property Let PublishedDate(ByVal StrValue As String)
'    'mBookInfo.Field(SSF_PublishDate) = StrValue
'    m_StrPublishedDate = StrValue
'End Property
'
'
''Public Property Get CurrentLength() As Long
''    CurrentLength = m_LonCurrentLength
''End Property
''
''Public Property Get CurrentPos() As Long
''    CurrentPos = m_LonCurrentPos
''End Property

Public Property Let StopNow(ByVal bValue As Boolean)
    m_bStopNow = bValue
End Property

'Public Property Get CurrentFile() As String
'    CurrentFile = m_StrCurrentFile
'End Property

'Public Property Let CurrentFile(ByVal StrValue As String)
'    m_StrCurrentFile = StrValue
'End Property

'Public Property Get RootURL() As String
'    RootURL = m_StrRootURL
'
'End Property
'
'Public Property Let RootURL(ByVal StrValue As String)
'        Dim iPos As Long
'        iPos = InStrRev(StrValue, "/")
'        If iPos <> Len(StrValue) Then
'            m_StrRootURL = Left$(StrValue, iPos)
'        Else
'            m_StrRootURL = StrValue
'        End If
'       ' mBookInfo.Field(SSF_URL) = m_StrRootURL
'End Property

Public Property Get taskId() As String
    taskId = m_StrTaskID
End Property

Public Property Let taskId(ByVal StringValue As String)
    m_StrTaskID = StringValue
End Property

Public Property Get FilesCount() As Long
    If m_LonFilesCount = 0 Then FilesCount = m_LonStaticCount Else FilesCount = m_LonFilesCount
End Property

'Public Property Get SavedIN() As String
'    SavedIN = m_StrSavedIN
'End Property
'
'Public Property Let SavedIN(ByVal StrValue As String)
'    On Error Resume Next
'    If Right$(StrValue, 1) <> "\" Then StrValue = StrValue & "\"
'    'If FileExist(StrValue) = False Then MkDir StrValue
'    m_StrSavedIN = StrValue
'    'mBookInfo.Field(SSF_SAVEDIN) = m_StrSavedIN
'End Property
'
'Public Property Get HttpHeader() As String
'    HttpHeader = m_StrHttpHeader
'End Property
'
'Public Property Let HttpHeader(ByVal StrValue As String)
'   Dim pHeader As CHttpHeader
'   Set pHeader = New CHttpHeader
'   pHeader.Init StrValue
'   If m_StrRootURL = "" Then
'        Dim host As String
'        Dim action As String
'        host = pHeader.GetField("host")
'        action = pHeader.HttpGet
'        If host <> "" And action <> "" Then
'            Me.RootURL = "http://" & host & action
'        End If
'   End If
'
'   pHeader.DeleteField ("")
'
'    m_StrHttpHeader = pHeader.HeaderString
'End Property
'
'Public Property Get Publisher() As String
'    Publisher = m_StrPublisher
'End Property
'
'Public Property Let Publisher(ByVal StrValue As String)
'    m_StrPublisher = StrValue
'End Property
'
'Public Property Get SSID() As String
'    SSID = m_StrSSID
'End Property
'
'Public Property Let SSID(ByVal StrValue As String)
'    If StrValue = "" Then SSID = "0"
'    m_StrSSID = StrValue
'End Property
'
'Public Property Get Author() As String
'    Author = m_StrAuthor
'End Property
'
'Public Property Let Author(ByVal StrValue As String)
'    m_StrAuthor = StrValue
'End Property
'
'Public Property Get Title() As String
'    Title = m_StrTitle
'End Property
'
'Public Property Let Title(ByVal StrValue As String)
'    If StrValue = "" Then StrValue = "未命名"
'    m_StrTitle = StrValue
'End Property
Public Sub Init(Optional ByRef Notify As ITaskNotify = Nothing)
    Set mNotify = Notify
    UpdateFolder
    
'    On Error GoTo error_init
'    Dim sName As String
'    m_LonStaticCount = 0
'    sName = Dir$(m_strFolder & "\*.pdg")
'    Do While sName <> ""
'        m_LonStaticCount = m_LonStaticCount + 1
'        sName = Dir$()
'    Loop
'
'
'error_init:
    
End Sub
    


'Private Sub mHelper_DownLoadOver()
'
'm_LonFilesCount = m_LonFilesCount + 1
'
'
'If mNotify Is Nothing Then Exit Sub
'mNotify.StatusChange Me
'
'
'
'End Sub

'Private Sub mHelper_GetData(progress As Long)
'If mNotify Is Nothing Then Exit Sub
'mNotify.StatusChange Me
'End Sub
'
'Private Sub mHelper_ProgressChange(vMin As Long, vMax As Long)
'
'm_LonCurrentPos = vMin
'If mNotify Is Nothing Then Exit Sub
'mNotify.StatusChange Me
'
'End Sub
'
'Private Sub mHelper_StatusChange(vStatus As String)
'If mNotify Is Nothing Then Exit Sub
'mNotify.StatusChange Me
'End Sub

Private Sub GetTextPdg(ByRef vFilename As String)
    Dim sUrls() As String
    Dim iCount As Long
    iCount = SSLIB_ParseInfoRule(vFilename, sUrls)
    Dim i As Long
    For i = 0 To iCount - 1
        If GetURL(sUrls(i), , False) Then
            m_LonFilesCount = m_LonFilesCount + TextPdgCount(sUrls(i))
        End If
        
        If m_bStopNow Then StopMe: Exit Sub
    Next
    Me.Status = STS_COMPLETE
    Call NotifyStatus
End Sub


Public Sub StartDownload()

   Dim pHeader As CHttpHeader
   Set pHeader = New CHttpHeader
   pHeader.Init mBookInfo(SSF_HEADER)
   
   If mBookInfo(SSF_URL) = "" Then
        Dim host As String
        Dim action As String
        host = pHeader.GetField("host")
        action = pHeader.HttpGet
        If host <> "" And action <> "" Then
            mBookInfo(SSF_URL) = "http://" & host & action
            
        End If
   End If
       
   pHeader.DeleteField ("")
   mBookInfo(SSF_HEADER) = pHeader.HeaderString
    

    m_bChanged = True
  
    If Not CheckFolder(mBookInfo(SSF_SAVEDIN)) Then
        MsgBox "不能创建" & mBookInfo(SSF_SAVEDIN) & vbCrLf & "请检查配置", vbCritical, "Task.StartDownload"
        Exit Sub
    End If
    
    UpdateFolder
    If Not CheckFolder(m_strFolder) Then
        MsgBox "不能创建" & m_strFolder & vbCrLf & "请检查配置", vbCritical, "Task.StartDownload"
        Exit Sub
    End If
    
    m_LonStartPage = StringToLong(mBookInfo(SSF_StartPage))
    
    Me.Status = STS_START
    Me.StopNow = False
    mDownloader.ClearErrors
    mStep = 0
    m_LonFilesCount = 0
    Call NotifyTaskStatus
    
    StartDownload2
    
    On Error Resume Next
    'If mBookInfo(SSF_PagesCount) < m_LonFilesCount Then mBookInfo(SSF_PagesCount) = m_LonFilesCount
    
    If Not mNotify Is Nothing Then mNotify.TaskComplete Me
End Sub
Private Sub StartDownload2()

    
    If mStep <= 1 Then
    'dat
        GetURL "bookinfo.dat", , False
        GetURL "BookContents.dat", , False
        GetURL "BookContents.xml", , False
        GetURL "catalog.dat", , False
        GetURL "Catalog.dat", , False
        If m_bStopNow Then StopMe: Exit Sub
        mStep = 2
    End If
    Call WriteBookInfo
    
    If GetURL("InfoRule.dat", , False) Then
        GetTextPdg m_strFolder & "\InfoRule.dat"
        Exit Sub
    End If

    GetRangeUrl "bac", 3, ".pdg", , False
    If m_bStopNow Then StopMe: Exit Sub
    GetRangeUrl "att", 3, ".pdg", , False
    If m_bStopNow Then StopMe: Exit Sub
    
    If mStep <= 2 Then
    '!...pdg
        GetRangeUrl "!", 5, ".pdg", , False
        If m_bStopNow Then StopMe: Exit Sub
        mStep = 3
    End If
    
    If mStep <= 3 Then
    '...pdg
        'If Not GetURL("000001.pdg", False, False) Then
        
        'Else
            Dim p As Long
            p = GetStartPoint()
            If p > 1 Then m_LonStartPage = p
            GetRangeUrl "", 6, ".pdg", p
        'End If
        If m_bStopNow Then StopMe: Exit Sub
        mStep = 4
    End If
    
    If mStep <= 4 Then
    'cov...pdg
        GetRangeUrl "cov", 3, ".pdg", , False
        If m_bStopNow Then StopMe: Exit Sub
        mStep = 5
    End If
    
    If mStep <= 5 Then
    'fow...pdg
        GetRangeUrl "fow", 3, ".pdg", , False
        If m_bStopNow Then StopMe: Exit Sub
        mStep = 6
    End If
    
    If mStep <= 6 Then
    'bok...pdg
        GetRangeUrl "bok", 3, ".pdg", , False
        If m_bStopNow Then StopMe: Exit Sub
        mStep = 7
    End If
    
    If mStep <= 7 Then
    'leg...pdg
        GetRangeUrl "leg", 3, ".pdg", , False
        If m_bStopNow Then StopMe: Exit Sub
        mStep = 8
    End If
       
    Me.Status = STS_COMPLETE
    
    mBookInfo(SSF_DownloadDate) = DateTime.Date$ & " " & DateTime.Time$
  
    Me.AutoSave
    
If Not mNotify Is Nothing Then mNotify.TaskStatusChange Me

End Sub
Private Sub StopMe()
    Me.Status = STS_PAUSE
    m_bStopNow = False
    Call NotifyStatus
End Sub

Private Function BuildRangeString(vPre As String, vNumLength As Long, vSuf As String, Optional vStart As Long = 1)
    Dim s As String
    Dim l As Long
    
    s = CStr(vStart)
    l = vNumLength - Len(s)
    If (l > 0) Then
        BuildRangeString = vPre & String$(l, "0") & s & vSuf
    Else
        BuildRangeString = vPre & s & vSuf
    End If

End Function
Private Function GetRangeUrl(vPre As String, vNumLength As Long, vSuf As String, Optional vStart As Long = 1, Optional fCount As Boolean = True)
    Dim sUrl As String
    Dim i As Long
    i = vStart - 1
    Do
        If m_bStopNow Then Exit Function
        i = i + 1
        sUrl = BuildRangeString(vPre, vNumLength, vSuf, i)
        If m_bStopNow Then Exit Function
    Loop While GetURL(sUrl, , fCount)
End Function

Private Function FileExist(ByRef sFilename As String) As Boolean
On Error GoTo Error_test
GetAttr sFilename
FileExist = True
    
Exit Function
'FileExist = True
Error_test:
FileExist = False

End Function

'Private Function CreateNewConnection() As CDownload
'    Dim index As long
'    Dim i As long
'    index = -1
'    If m_intConnectionsCount > 0 Then
'        For i = 1 To m_intConnectionsCount
'            If m_Connections(i) Is Nothing Then
'                Set m_Connections(i) = New CDownload
'                index = i
'                Exit For
'            ElseIf m_Connections(i).IsFree Then
'                index = i
'                Exit For
'            End If
'        Next
'    End If
'    If index = -1 Then
'        m_intConnectionsCount = m_intConnectionsCount + 1
'        ReDim Preserve m_Connections(1 To m_intConnectionsCount)
'        Set m_Connections(m_intConnectionsCount) = New CDownload
'        index = m_intConnectionsCount
'    End If
'
'    m_Connections(index).ID = index
'    Set m_Connections(index).Client = Me
'    Set CreateNewConnection = m_Connections(index)
'
'End Function

'Private Function ReleaseConnection(ByRef vDownloader As CDownload)
'    'Set m_Connections(vDownloader.ID) = Nothing
'    'Set vDownloader = Nothing
'End Function

Private Function download(ByRef vURL As String, Optional vOverWrite As Boolean = False) As Boolean
    Dim saveAs As String
    saveAs = m_strFolder & "\" & vURL
    Dim URL As String
    URL = mBookInfo(SSF_URL) & vURL
    If Not mDownloader.download(saveAs, URL, mBookInfo(SSF_HEADER), vOverWrite) Then
        download = mDownloader.download(saveAs, URL, mBookInfo(SSF_HEADER), vOverWrite)
    Else
        download = True
    End If
    
End Function

Private Function GetURL(vURL As String, Optional ByVal vOverWrite As Boolean = False, Optional ByVal fCount As Boolean = True) As Boolean
 On Error GoTo ErrorDownload
    'm_LonCurrentLength = 0
    'm_LonCurrentPos = 0
    'm_StrCurrentFile = vURL
    Static LastUrl As String
    
    
    If download(vURL, vOverWrite) Then
        If fCount Then m_LonFilesCount = m_LonFilesCount + 1
        GetURL = True
    Else
        GetURL = False
    End If
    Call NotifyStatus
    
    
    
    
    Exit Function
    
ErrorDownload:
    On Error Resume Next
    Err.Clear
    GetURL = False
    'Kill saveAs
    Call NotifyStatus
End Function

Public Sub Restart()
    mStep = 0
    m_LonFilesCount = 0
End Sub

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

Private Sub UpdateFolder()
    Dim OldFolder As String
    OldFolder = m_strFolder
    m_strFolder = mBookInfo(SSF_Title)
    If m_strFolder = "" Then m_strFolder = "未命名"
    If (mBookInfo(SSF_Author) <> "") Then m_strFolder = m_strFolder & "_" & mBookInfo(SSF_Author)
    If (mBookInfo(SSF_SSID) <> "") Then m_strFolder = m_strFolder & "_" & mBookInfo(SSF_SSID)
    m_strFolder = BuildPath(mBookInfo(SSF_SAVEDIN)) & m_strFolder
    m_strFolder = ReplaceInvalidChars(m_strFolder, " ")
    If OldFolder <> "" And OldFolder <> m_strFolder Then
        If FolderExists(OldFolder) And FolderExists(m_strFolder) = False Then
            If (MsgBox("将" & vbCrLf & OldFolder & vbCrLf & "重命名为" & vbCrLf & m_strFolder & "？", vbInformation + vbYesNo + vbDefaultButton2, "Task.UpdateFolder") = vbYes) Then
                MoveFile OldFolder, m_strFolder
            End If
        End If
    End If
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
        mBookInfo.SaveToFile "", iniHnd, vSection
        'iniHnd.SaveSetting vSection, "StartPage", m_LonStartPage
        'm_LonStartPage = CLng(iniHnd.GetSetting(vSection, "StartPage"))
End Sub

Public Sub LoadFrom(ByRef iniHnd As CLiNInI, ByRef vSection As String)
On Error Resume Next
    Dim title As String
    title = iniHnd.GetSetting(vSection, "Title")
    If title <> "" Then
        mBookInfo(SSF_Title) = title
        mBookInfo(SSF_Author) = iniHnd.GetSetting(vSection, "Author")
        mBookInfo(SSF_SSID) = iniHnd.GetSetting(vSection, "SSId")
        mBookInfo(SSF_Publisher) = iniHnd.GetSetting(vSection, "Publisher")
        mBookInfo(SSF_HEADER) = iniHnd.GetSetting(vSection, "Header")
        mBookInfo(SSF_URL) = iniHnd.GetSetting(vSection, "URL")
        mBookInfo(SSF_SAVEDIN) = iniHnd.GetSetting(vSection, "SaveIn")
        mBookInfo(SSF_PagesCount) = iniHnd.GetSetting(vSection, "PagesCount")
        mBookInfo(SSF_PublishDate) = iniHnd.GetSetting(vSection, "PDate")
        mBookInfo(SSF_Comments) = iniHnd.GetSetting(vSection, "AddInfo")
    Else
        mBookInfo.LoadFromFile "", iniHnd, vSection
    End If
        
        

        Status = CLng(iniHnd.GetSetting(vSection, "Status"))
        m_LonFilesCount = CLng(iniHnd.GetSetting(vSection, "Downloaded"))
        'm_LonStartPage = CLng(iniHnd.GetSetting(vSection, "StartPage"))
End Sub

Public Property Get Directory() As String
    UpdateFolder
    Directory = m_strFolder
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
    
    UpdateFolder
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


'Private Sub IDownloadClient_ProgressChanged(ByVal vId As Long, vProgress As CDownloadProgress)
'    Call NotifyDownloadStatus
'End Sub


Private Sub NotifyTaskStatus()
    If Not mNotify Is Nothing Then mNotify.TaskStatusChange Me
End Sub

Private Sub NotifyDownloadStatus()
    If Not mNotify Is Nothing Then mNotify.DownloadStatusChange Me
End Sub

Private Sub NotifyStatus()
If Not mNotify Is Nothing Then
    mNotify.TaskStatusChange Me
    mNotify.DownloadStatusChange Me
End If
End Sub

Private Sub mDownloader_ProgressChanged(ByVal vId As Long, Progress As CDownloadProgress)
    Call NotifyDownloadStatus
End Sub

Private Function GetDefaultStartPoint() As Long
    If download(BuildRangeString("", 6, ".pdg", 1)) Then
        GetDefaultStartPoint = True
    Else
        GetDefaultStartPoint = False
    End If
End Function

Private Function GetStartPoint() As Long
    If GetDefaultStartPoint Then GetStartPoint = 1: Exit Function
    'If m_LonStartPage > 0 Then GetStartPoint = m_LonStartPage: Exit Function
    Dim u As Long
    Dim i As Long
    Dim URL As String
    If m_LonStartPage > 0 Then i = m_LonStartPage - 1
    If mBookInfo(SSF_PagesCount) <> "" Then u = StringToLong(mBookInfo(SSF_PagesCount))
    If u < 1 Then u = 9999
    Do
        DoEvents
        If i > u Then
            i = 1
            Exit Do
        End If
        i = i + 1
        URL = BuildRangeString("", 6, ".pdg", i)
        If m_bStopNow Then
            StopMe
            GetStartPoint = 1
            Exit Function
        End If
    Loop Until download(URL, False)
    GetStartPoint = i
    If i > 1 Then mBookInfo(SSF_StartPage) = i
End Function
