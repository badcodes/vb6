VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "sslib Explorer"
   ClientHeight    =   5310
   ClientLeft      =   180
   ClientTop       =   855
   ClientWidth     =   16575
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "MainV2.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5310
   ScaleWidth      =   16575
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraIE 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4485
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   6960
      Begin VB.PictureBox picAddress 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   0
         ScaleHeight     =   495
         ScaleWidth      =   10935
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   0
         Width           =   10935
         Begin VB.ComboBox cboAddress 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3000
            TabIndex        =   9
            Top             =   120
            Width           =   3795
         End
         Begin MSComctlLib.Toolbar tbToolBar 
            Height          =   450
            Left            =   0
            TabIndex        =   10
            Top             =   0
            Width           =   2865
            _ExtentX        =   5054
            _ExtentY        =   794
            ButtonWidth     =   820
            ButtonHeight    =   794
            Wrappable       =   0   'False
            Appearance      =   1
            Style           =   1
            ImageList       =   "imlToolbarIcons"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   6
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Back"
                  Object.ToolTipText     =   "Back"
                  ImageIndex      =   1
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Forward"
                  Object.ToolTipText     =   "Forward"
                  ImageIndex      =   2
               EndProperty
               BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Stop"
                  Object.ToolTipText     =   "Stop"
                  ImageIndex      =   3
               EndProperty
               BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Refresh"
                  Object.ToolTipText     =   "Refresh"
                  ImageIndex      =   4
               EndProperty
               BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Home"
                  Object.ToolTipText     =   "Home"
                  ImageIndex      =   5
               EndProperty
               BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Search"
                  Object.ToolTipText     =   "Search"
                  ImageIndex      =   6
               EndProperty
            EndProperty
         End
      End
      Begin SHDocVwCtl.WebBrowser IE 
         Height          =   3735
         Left            =   0
         TabIndex        =   11
         Top             =   495
         Width           =   5400
         ExtentX         =   9513
         ExtentY         =   6586
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   1
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   "http:///"
      End
   End
   Begin VB.Frame fraBookList 
      BorderStyle     =   0  'None
      Height          =   3030
      Left            =   8160
      TabIndex        =   0
      Top             =   240
      Width           =   8160
      Begin VB.Frame FraButtons 
         BorderStyle     =   0  'None
         Height          =   3525
         Left            =   480
         TabIndex        =   1
         Top             =   360
         Width           =   3330
         Begin VB.CommandButton cmd 
            Caption         =   "Set Cookie"
            Height          =   360
            Index           =   8
            Left            =   1800
            TabIndex        =   17
            Top             =   1560
            Width           =   1125
         End
         Begin VB.CommandButton cmd 
            Caption         =   "下一页"
            Height          =   360
            Index           =   7
            Left            =   1680
            TabIndex        =   16
            Top             =   1080
            Width           =   1125
         End
         Begin VB.CommandButton cmd 
            Caption         =   "上一页"
            Height          =   360
            Index           =   6
            Left            =   1800
            TabIndex        =   15
            Top             =   720
            Width           =   1125
         End
         Begin VB.CommandButton cmd 
            Caption         =   "阅读|下载"
            Height          =   360
            Index           =   1
            Left            =   1680
            TabIndex        =   14
            Top             =   120
            Width           =   1125
         End
         Begin VB.CommandButton cmd 
            Caption         =   "Cookie"
            Height          =   360
            Index           =   5
            Left            =   375
            TabIndex        =   12
            Top             =   2325
            Width           =   1125
         End
         Begin VB.CommandButton cmd 
            Caption         =   "删除"
            Height          =   360
            Index           =   2
            Left            =   270
            TabIndex        =   5
            Top             =   885
            Width           =   1125
         End
         Begin VB.CommandButton cmd 
            Caption         =   "下载"
            Height          =   360
            Index           =   0
            Left            =   240
            TabIndex        =   4
            Top             =   360
            Width           =   1125
         End
         Begin VB.CommandButton cmd 
            Caption         =   "清除列表"
            Height          =   360
            Index           =   3
            Left            =   315
            TabIndex        =   3
            Top             =   1335
            Width           =   1125
         End
         Begin VB.CommandButton cmd 
            Caption         =   "更新"
            Height          =   360
            Index           =   4
            Left            =   345
            TabIndex        =   2
            Top             =   1830
            Width           =   1125
         End
      End
      Begin MSComctlLib.ListView lstBooks 
         Height          =   1035
         Left            =   2880
         TabIndex        =   6
         Top             =   600
         Width           =   5310
         _ExtentX        =   9366
         _ExtentY        =   1826
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ID"
            Object.Width           =   452
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "SSID"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "书名"
            Object.Width           =   5080
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "页数"
            Object.Width           =   1806
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "作者"
            Object.Width           =   2992
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "出版社"
            Object.Width           =   2992
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "出版日期"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "简介"
            Object.Width           =   8819
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   2340
      Top             =   1785
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainV2.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainV2.frx":0B54
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainV2.frx":1266
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainV2.frx":1978
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainV2.frx":208A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainV2.frx":279C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin sslibExplorer.ctlSplitterEx ctlSplitterEx 
      Height          =   5295
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   15135
      _ExtentX        =   26696
      _ExtentY        =   9340
   End
   Begin VB.Menu mnuPreference 
      Caption         =   "设置(&P)"
   End
   Begin VB.Menu mnuBookList 
      Caption         =   "BookList"
      Visible         =   0   'False
      Begin VB.Menu mnuBookListCopy 
         Caption         =   ""
         Index           =   0
      End
      Begin VB.Menu mnuBookListCopy 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuBookListCopy 
         Caption         =   "复制SSID"
         Index           =   2
      End
      Begin VB.Menu mnuBookListCopy 
         Caption         =   "下载"
         Index           =   3
      End
      Begin VB.Menu mnuBookListCopy 
         Caption         =   "阅读|下载"
         Index           =   4
      End
      Begin VB.Menu mnuBookListCopy 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuBookListCopy 
         Caption         =   "删除"
         Index           =   6
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Const CSTTaskIncomeFilename As String = "Incoming.lst"
Private Const CSTTaskIncomeFilename2 As String = "Incoming.tmp"



Private mDefaultCaption As String

Private mIncomingFile As String
Private mIncomingFile2 As String


Public mLastAddress As String
Dim mbDontNavigateNow As Boolean
Private Const CST_SSLIBRARY_HOST As String = "sslibrary.com"
Private Const CST_DUXIU_HOST As String = "duxiu.com"
Private mBookListUrl As String
Private Const cstLeft As Long = 120
Private Const cstTop As Long = 120
'Private WithEvents timTimer As CTimer
Private Const cstTimerInterval As Long = 400

Private mScriptHost As IHTMLWindow2

Public LastAddress As String
'Private mSkipAddress As Boolean
Private Const configFile As String = "sslibExplorer.ini"
Private mPassword As String
Private mUsername As String
Private mHomePage As String

Private mTaskList As String
Private mCookie As String
Private mUnloading As Boolean
Private WithEvents OLDIE As WebBrowser_V1
Attribute OLDIE.VB_VarHelpID = -1


Private Function GetSelectBook() As CBookInfo
    If lstBooks.SelectedItem Is Nothing Then Exit Function
    If IsObject(lstBooks.SelectedItem.Tag) Then
        Set GetSelectBook = lstBooks.SelectedItem.Tag
    Else
        Set GetSelectBook = Nothing
    End If
End Function

Private Sub ctlSplitterEx_Resized()
    Form_Resize
End Sub

Private Sub Form_Load()

 On Error Resume Next


    Set OLDIE = IE.Object

    mDefaultCaption = App.ProductName & " " & App.Major & "." & App.Minor & " Build " & App.Revision
    Me.Caption = mDefaultCaption

    'Timer.Enabled = False
    
    
    Me.Show
    tbToolBar.Refresh
    

    Me.ctlSplitterEx.AttachObjects fraIE, fraBookList, True
    Me.ctlSplitterEx.TileMode = TILE_HORIZONTALLY
    Form_Resize
    
 SSLIB_Init

    'cboAddress.Move 50, lblAddress.Top + lblAddress.Height + 15
    
    
    Dim configHnd As CLiNInI
    Set configHnd = New CLiNInI
    With configHnd
        .Source = App.Path & "\" & configFile
        FormStateFromString Me, .GetSetting("Form", "State")
        ComboxItemsFromString cboAddress, .GetSetting("Browser", "History")
        mLastAddress = .GetSetting("Browser", "LastAddress")
        mHomePage = .GetSetting("Browser", "HomePage")
        mUsername = .GetSetting("Browser", "Username")
        mPassword = .GetSetting("Browser", "Password")
        mTaskList = .GetSetting("Browser", "TaskListFile")
    End With
    If mTaskList = vbNullString Then mTaskList = App.Path & "\" & CSTTaskIncomeFilename
    Set configHnd = Nothing
    
    
    If mLastAddress = vbNullString Then mLastAddress = mHomePage

    'Set timTimer = New CTimer
    'timTimer.Interval = 0
    

        IE.Navigate2 mLastAddress
    
'    If Len(mLastAddress) > 0 Then
'        cboAddress.text = mLastAddress
'        cboAddress.AddItem cboAddress.text
'        'try to navigate to the starting address
'
'        'IE.Navigate mLastAddress
'    End If

   
   ' AutoLogin
End Sub



Public Sub CallBack_AddTask(ByVal vName As String, vBookInfoArray() As String, Optional vPending As Boolean)
        
   On Error GoTo ErrorAddTask
   
        Dim vtask As CTask
        Set vtask = New CTask
        vtask.bookInfo.LoadFromArray vBookInfoArray
        vtask.Name = vName
        vtask.InitForDownload
        If vPending Then vtask.Status = STS_PENDING
        vtask.Changed = True
        vtask.AutoSave
        
        Dim fNum As Integer
        fNum = FreeFile
        Open mTaskList For Append As #fNum
        Print #fNum, vtask.Directory
        Close #fNum
        Exit Sub
        
ErrorAddTask:
    MsgBox "添加任务失败: 错误" & Err.Number & vbCrLf & Err.Description, vbCritical
    On Error Resume Next
    Close #fNum
End Sub
'
'Public Sub CallBack_AddTask(bookInfo As CBookInfo)
'    Dim vtask As CTask
'    Set vtask = New CTask
'    Set vtask.bookInfo = bookInfo
'End Sub


Private Sub Form_Terminate()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Unload frmOptions
    
    mUnloading = True
    
    Dim i As Integer
    For i = 0 To 4
        Sleep 200
        DoEvents
    Next
    Set frmTask = Nothing
    'Unload frmTask
    'Set timTimer = Nothing
    On Error Resume Next
     Dim configHnd As CLiNInI
    Set configHnd = New CLiNInI
    With configHnd
        .Source = App.Path & "\" & configFile
        .SaveSetting "Browser", "History", ComboxItemsToString(cboAddress)
        .SaveSetting "Form", "State", FormStateToString(Me)
        .SaveSetting "Browser", "LastAddress", IE.LocationURL
        .SaveSetting "Browser", "HomePage", mHomePage
        .SaveSetting "Browser", "UserName", mUsername
        .SaveSetting "Browser", "Password", mPassword
        .SaveSetting "Browser", "TaskListFile", mTaskList
        .Save
    End With
    Set configHnd = Nothing

End Sub




Private Function QuoteString(ByRef vString As String) As String
    QuoteString = Chr$(34) & vString & Chr$(34)
End Function


Private Sub mnuExit_Click()
    Unload Me
End Sub





Private Sub IE_DocumentComplete(ByVal pDisp As Object, URL As Variant)
    TimerCheck
End Sub

Private Sub lstBooks_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Dim bookInfo As CBookInfo
        Set bookInfo = GetSelectBook()
        If Not bookInfo Is Nothing Then mnuBookListCopy(0).Caption = bookInfo(SSF_Title)
        PopupMenu mnuBookList, , fraBookList.Left + lstBooks.Left + X, fraBookList.Top + lstBooks.Top + Y
    End If
End Sub

Private Sub mnuBookListCopy_Click(Index As Integer)
    If Index > 0 Then
        Dim vAction As String
        vAction = mnuBookListCopy(Index).Caption
        If vAction <> "-" Then ActionOnItem mnuBookListCopy(Index).Caption
    End If
End Sub

Private Sub mnuPreference_Click()
    Load frmOptions
    With frmOptions
        .HomePage = mHomePage
        .UserName = mUsername
        .Password = mPassword
        .TaskListFile = mTaskList
        .Show 1, Me
        mHomePage = .HomePage
        mUsername = .UserName
        mPassword = .Password
        mTaskList = .TaskListFile
    End With
    Unload frmOptions
End Sub


'CSEH: ErrExit
Private Sub ClickLinkByText(ByRef vWindow As IHTMLWindow2, ByVal vLinkText As String)
    '<EhHeader>
    On Error GoTo ClickLinkByText_Err
    '</EhHeader>
    If vWindow Is Nothing Then Exit Sub
    Dim link As IHTMLElement
    Dim links As IHTMLElementCollection
    Set links = vWindow.Document.links
    For Each link In links
        Debug.Print link.innerText
        If StrComp(link.innerText, vLinkText, vbTextCompare) = 0 Then
            link.Click
            Exit Sub
        End If
    Next
    MsgBox "没有发现名称为" & QuoteString(vLinkText) & "的链接", vbInformation + vbOKOnly
    '<EhFooter>
    Exit Sub

ClickLinkByText_Err:
    Debug.Print "sslibExplorer.frmMain.ClickLinkByText:Error " & Err.Description
    Err.Clear

    '</EhFooter>
End Sub

Sub inputCookie()
    Dim vCookie As String
    mCookie = InputBox("Set Cookie below", "Set Cookie", mCookie)
End Sub

Private Sub ActionOnItem(ByVal vAction As String)
    
    If vAction = "清除列表" Then
        mBookListUrl = vbNullString
        lstBooks.ListItems.Clear
        Exit Sub
    ElseIf vAction = "更新" Then
        mBookListUrl = vbNullString
        TimerCheck
        Exit Sub
    ElseIf vAction = "下载" Then
        frmTask.Show 0, Me
    ElseIf vAction = "阅读|下载" Then
        frmTask.Show 0, Me
    ElseIf vAction = "Cookie" Then
        Clipboard.SetText mCookie
        MsgBox mCookie, vbOKOnly, "Cookie已经复制到剪贴板"
        Exit Sub
    ElseIf vAction = "Set Cookie" Then
        inputCookie
        Exit Sub
    ElseIf vAction = "上一页" Then
        ClickLinkByText mScriptHost, vAction
        Exit Sub
    ElseIf vAction = "下一页" Then
        ClickLinkByText mScriptHost, vAction
        Exit Sub
    End If
    
    Dim bookInfo As CBookInfo
    Set bookInfo = GetSelectBook()
    If bookInfo Is Nothing Then
        MsgBox "No tasks selected.", vbInformation, vAction
        Exit Sub
    End If
    
    
    
    Select Case vAction
        Case "下载"
            AddTask bookInfo, False
            Exit Sub
        Case "阅读|下载"
            AddTask bookInfo, True
            Exit Sub
         Case "复制SSID"
            Clipboard.SetText bookInfo(SSF_SSID)
            Exit Sub

    End Select
    
    
    Dim selectListItem As ListItem
    Dim selected() As ListItem
    Dim count As Long
    For Each selectListItem In lstBooks.ListItems
        If selectListItem.selected Then
            count = count + 1
            ReDim Preserve selected(1 To count)
            Set selected(count) = selectListItem
        End If
    Next
    
    Dim i As Long
    
    Select Case vAction
        Case "删除"
            For i = count To 1 Step -1
                lstBooks.ListItems.Remove selected(i).Key
            Next
            For Each selectListItem In lstBooks.ListItems
                selectListItem.text = selectListItem.Index
            Next
    End Select
End Sub

Private Sub AddTask(ByRef vBookInfo As CBookInfo, Optional vOpen As Boolean = False)
    On Error Resume Next
    
    If vOpen Then
        
        If Left$(vBookInfo(SSF_SSURL), 7) = "book://" Then
            win.ShellExecute Me.HWND, "open", vBookInfo(SSF_SSURL), vbNull, vbNull, SW_NORMAL
        Else
        
        Dim pWin As IHTMLWindow2
        If mScriptHost Is Nothing Then Set pWin = IE.Document.parentWindow Else Set pWin = mScriptHost
        'Set pWin = lstBooks.Tag
        If pWin Is Nothing Then Exit Sub
        Dim pUrl As String
        
        Dim pScript As String
        'pUrl = "/" & SubStringBetween(vBookInfo(SSF_PageURL), "('", "')", True)
        'pWin.Navigate pUrl
        pScript = "readbook('" & vBookInfo(SSF_SSURL) & "');"
        pWin.execScript pScript
        End If
    End If
    
    frmTask.EditTaskMode = True
    frmTask.InitWithBookInfo vBookInfo
    frmTask.Show 0, Me
    
End Sub
Private Sub cmd_Click(Index As Integer)
    ActionOnItem cmd(Index).Caption
    
'    Select Case Index
'
'        Case 0
'            ActionOnItem "Download"
'        Case 1
'            ActionOnItem "Delete"
'
'    End Select
End Sub

Public Sub Init(vHomePage As String, vUsername As String, vPassWord As String)
    mHomePage = vHomePage
    mPassword = vPassWord
    mUsername = vUsername
End Sub




'CSEH: ErrReport
Private Sub TimerCheck()
        '<EhHeader>
        On Error GoTo TimerCheck_Err
        '</EhHeader>
        On Error Resume Next
        If mUnloading Then Exit Sub
        Dim elm As IHTMLElement
        Dim elms As IHTMLElementCollection
100     Set elms = IE.Document.getElementsByTagName("iframe")
102     For Each elm In elms
104         CheckUrl elm.contentWindow.location.href, elm.contentWindow
        Next
106     CheckUrl IE.LocationURL, IE.Document.parentWindow
110     Err.Clear
        '<EhFooter>
        Exit Sub

TimerCheck_Err:
        Debug.Print "sslibExplorer.frmMain.TimerCheck" & ":line:" & Erl & ":" & Err.Description,
        Resume Next
        '</EhFooter>
End Sub




Private Sub IE_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
'    If mSkipAddress Then mSkipAddress = False: Exit Sub
'    LastAddress = URL
    
End Sub

Private Sub IE_DownloadComplete()
    On Error Resume Next
    Me.Caption = IE.LocationName
    
    'Dim i As Long
    

End Sub

Private Sub CheckUrl(ByVal vUrl As String, ByRef vWindow As IHTMLWindow2) '  As HTMLDDElement)
On Error Resume Next
        If mUnloading Then Exit Sub
    Debug.Print "checkURL : " & vUrl
    Dim pUrl As String
    pUrl = LCase$(vUrl)
    If vWindow Is Nothing Then Set vWindow = IE
    
    Dim vDocument As HTMLDocument
    Dim vBody As IHTMLBodyElement2
    Dim vLocation As IHTMLLocation
    Dim vDocElement As IHTMLHtmlElement
    
    Set vDocument = vWindow.Document
    Set vBody = vDocument.body
    Set vLocation = vDocument.location
    Set vDocElement = vDocument.documentElement
    
    Dim pHref As String
    Dim pHost As String
    Dim pPath As String
    pHref = LCase$(vLocation.href)
    pHost = LCase$(vLocation.Host)
    pPath = LCase$(vLocation.PathName)
    
    AutoLogin vDocument
    
    If Right$(pHost, Len(CST_SSLIBRARY_HOST)) = CST_SSLIBRARY_HOST Then
        
        If pPath = "/userlogon.jsp" Then
            'AutoLogin vDocument
        ElseIf pPath = "/search.jsp" Then
            AddSearchResult pHref, vDocument
        ElseIf pPath = "/books.jsp" Then
            AddSearchResult pHref, vDocument
        ElseIf pPath = "/advshow.jsp" Then
            AddSearchResult pHref, vDocument
        ElseIf pPath = "/showbook_duxiudsr.jsp" Then
            AddSearchResult pHref, vDocument
        End If
        
    ElseIf Right$(pHost, Len(CST_DUXIU_HOST)) = CST_DUXIU_HOST Then
        Select Case pPath
            Case "/search"
                DoxiuSearch pHref, vDocument
            Case Else
        End Select
        
        
    End If
    
    
    
End Sub

Private Sub DoxiuSearch(ByRef pHref As String, ByRef vDoc As IHTMLDocument3)
    On Error Resume Next
    If mUnloading Then Exit Sub
    
    Dim tables As IHTMLElementCollection
    
    Dim table As HTMLTable
    
    Set tables = vDoc.getElementsByTagName("table")
    
    
    For Each table In tables
        Debug.Print "Table:Class->" & table.className
        If LCase$(table.className) <> "book1" Then GoTo nextTable
        
        
        Dim elms As IHTMLElementCollection
        Dim elm As IHTMLAnchorElement
        
        Set elms = table.getElementsByTagName("a")
        
        For Each elm In elms
            If mUnloading Then Exit Sub
            If InStr(elm.href, "/gobaoku.jsp?") > 0 Then
                
                
                Dim newDoc As HTMLDocument
                Dim newDocELM As IHTMLDocument2
                Set newDoc = New HTMLDocument
                newDoc.cookie = vDoc.cookie
                'Set newDocELM = newDoc.open(elm.href)
                Set newDocELM = newDoc.createDocumentFromUrl(Replace$(elm.href, "pds.sslibrary.com", "hn.sslibrary.com", , , vbTextCompare), vbNullString)
                Do Until newDocELM.ReadyState = "complete"
                    If mUnloading Then Exit Sub
                    DoEvents
                    'Sleep 100
                Loop
                
                    Dim newBook As CBookInfo
                    Set newBook = New CBookInfo
                    FillBookInfoFromElement newBook, newDocELM.body, newDocELM.location.href
                    FillBookInfoFromElement newBook, table, ""
                                        
                    AddBookInfoToList newBook
                    Set newBook = Nothing
                    GoTo nextTable

            End If
        Next
nextTable:
    Next
    
    
End Sub

Private Sub FillBookInfoFromElement(ByRef vBookInfo As CBookInfo, ByRef vElm As IHTMLElement, Optional vRefer As String)
    '<EhHeader>
    On Error GoTo FillBookInfoFromElement_Err
    '</EhHeader>
    Dim vInfo() As String
    vInfo = MSSReader.SSLIB_CreateBookInfoArray
    vInfo = MSSReader.SSLIB_ParseInfoText(vElm.innerText)
    
    Debug.Print vInfo(SSLIBFields.SSF_Title)
    
        If vInfo(SSF_Title) <> vbNullString Then
            If vRefer <> vbNullString Then vInfo(SSF_PAGEURL) = Replace(vRefer, "pds.sslibrary.com", "hn.sslibrary.com", , , vbTextCompare)
            Dim elm_a As IHTMLAnchorElement
            Dim elm_as As IHTMLElementCollection
            Set elm_as = vElm.getElementsByTagName("a")
            Dim pos As Integer
            Dim href As String
            For Each elm_a In elm_as
                href = Replace$(elm_a.href, "pds.sslibrary.com", "hn.sslibrary.com", , , vbTextCompare)
                
                pos = InStr(1, href, "gojpgRead.jsp?", vbTextCompare)
                If pos < 1 Then pos = InStr(1, href, "fromduxiutoJpg.jsp?", vbTextCompare)
                If (pos > 0) Then
                    vInfo(SSF_PARAMS) = MStrings.SubStringBetween(href, "?")
                    If vInfo(SSF_SSID) = vbNullString Then vInfo(SSF_SSID) = SubStringBetween(vInfo(SSF_PARAMS), "ssnum=", "&")
                    vInfo(SSF_IEJPGURL) = href
                ElseIf InStr(1, href, "book://", vbTextCompare) = 1 Then
                    If vInfo(SSF_SSURL) = vbNullString Then vInfo(SSF_SSURL) = href
                    If vInfo(SSF_SSID) = vbNullString Then vInfo(SSF_SSID) = SubStringBetween(vInfo(SSF_SSURL), "ssid=", "&")
                ElseIf InStr(1, href, "javascript:readbook(", vbTextCompare) = 1 Then
                    If vInfo(SSF_SSURL) = vbNullString Then vInfo(SSF_SSURL) = SubStringBetween(href, "('", "')")
                    If vInfo(SSF_SSID) = vbNullString Then vInfo(SSF_SSID) = SubStringBetween(href, "dxNumber=", "&")
                End If
Continue:
            Next
        End If
    
    
    vBookInfo.LoadFromArray vInfo, False
    
    
    '<EhFooter>
    Exit Sub

FillBookInfoFromElement_Err:
    Debug.Print "sslibExplorer.frmMain.FillBookInfoFromElement:Error " & Err.Description
    On Error Resume Next
    vBookInfo.LoadFromArray vInfo, False
    Err.Clear
    '</EhFooter>
End Sub

Private Sub AddBookInfoToList(ByRef vBookInfo As CBookInfo)
    On Error Resume Next
    Dim pItem As ListItem
    Set pItem = lstBooks.ListItems.Add(, vBookInfo(SSF_Title) & vBookInfo(SSF_SSID))
    If Not pItem Is Nothing Then
        Set pItem.Tag = vBookInfo
        pItem.text = pItem.Index
        pItem.SubItems(1) = vBookInfo(SSF_SSID)
        pItem.SubItems(2) = vBookInfo(SSF_Title)
        pItem.SubItems(3) = vBookInfo(SSF_PagesCount)
        pItem.SubItems(4) = vBookInfo(SSF_AUTHOR)
        pItem.SubItems(5) = vBookInfo(SSF_Publisher)
        pItem.SubItems(6) = vBookInfo(SSF_PublishDate)
        pItem.SubItems(7) = vBookInfo(SSF_Comments)
    End If
    Err.Clear
End Sub
Private Sub AddSearchResult(ByRef pHref As String, ByRef vDoc As IHTMLDocument2)
    On Error Resume Next
    If pHref = mBookListUrl Then Exit Sub
    Dim pBooks As Collection
    Set pBooks = GetBooksInfo(vDoc)
    If pBooks Is Nothing Then Exit Sub
    If pBooks.count > 0 Then
        Debug.Print "AddSearchResult Get pBooks " & pBooks.count
        Set mScriptHost = vDoc.parentWindow '  .parentWindow
        
        
        Dim pBook As CBookInfo
       ' Dim pItem As ListItem
        For Each pBook In pBooks
            AddBookInfoToList pBook
        Next

    End If
    Err.Clear
End Sub

Private Function GetBooksInfo(ByRef vDoc As HTMLDocument) As Collection
    Dim result As Collection
    Set result = New Collection
    On Error Resume Next
    'Dim divResult As IHTMLDivElement
    'If Not FindElement(divResult, vDoc, "divResult") Then Exit Function
    'lstBooks.ListItems.Clear
    Dim elms As IHTMLElementCollection
    Set elms = vDoc.documentElement.getElementsByTagName("TABLE")
    Dim elm As IHTMLElement
    
    'http://pds.sslibrary.com/gojpgRead.jsp?ssnum=10000374&d=02506F1905D12774A97AFBDD803C629A&fenleiID=0I20407050&ssreaderurl=http%3A%2F%2Fpds.sslibrary.com%3A80%2FgopdgRead.jsp%3FdxNumber%3D10000374%26d%3DA5835EB61949C2E9D5D95D17E6DAB3A5%26fenleiID%3D0I20407050%26username%3Dhntsg%26pdgcode%3D85F61AA57A801B35A2CEE7106C988B2D%26jpathkey%3D3756268969727
    
    Dim vBookInfo As CBookInfo
    For Each elm In elms
        Set vBookInfo = New CBookInfo
        FillBookInfoFromElement vBookInfo, elm, vDoc.location.href
        If vBookInfo(SSF_SSID) <> "" Or vBookInfo(SSF_Title) <> "" Then
            result.Add vBookInfo
        End If
    Next
    
    Set GetBooksInfo = result
End Function

Private Sub AutoLogin(ByRef vDoc As HTMLDocument)
    
    Debug.Print "AutoLogin..."
    
    Dim elms As IHTMLElementCollection
    Dim elm As IHTMLInputElement
    Set elms = vDoc.getElementsByTagName("input")
    For Each elm In elms
        Dim vName As String
        vName = LCase$(elm.Name)
        If vName = "username" Or vName = "userid" Then
            elm.Value = mUsername
        ElseIf vName = "password" Or vName = "userpwd" Then
            elm.Value = mPassword
        End If
    Next
        
End Sub

Private Sub FillField(ByRef vDoc As IHTMLDocument2, ByVal vName As String, ByVal vValue As String)
    Dim elms As IHTMLElementCollection
    Dim elm As IHTMLInputElement
    Set elms = vDoc.getElementsByTagName("input")
    For Each elm In elms
        If StrComp(elm.Name, vName, vbTextCompare) = 0 Then
            elm.Value = vValue
        End If
    Next
End Sub

Private Sub IE_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
    
    On Error Resume Next
    Dim i As Integer
    Dim bFound As Boolean
    Me.Caption = IE.LocationName
    For i = 0 To cboAddress.ListCount - 1
        If cboAddress.List(i) = IE.LocationURL Then
            bFound = True
            Exit For
        End If
    Next i
    mbDontNavigateNow = True
    If bFound Then
        cboAddress.RemoveItem i
    End If
    cboAddress.AddItem IE.LocationURL, 0
    cboAddress.ListIndex = 0
    mbDontNavigateNow = False
    'timTimer.Interval = cstTimerInterval
    mCookie = IE.Document.cookie
    Debug.Print mCookie
    'TimerCheck
    'CheckUrl URL, pDisp
End Sub


Private Sub cboAddress_Click()
    If mbDontNavigateNow Then Exit Sub
    IE.Navigate cboAddress.text
End Sub


Private Sub cboAddress_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = vbKeyReturn Then
        cboAddress_Click
    End If
End Sub


Private Sub Form_Resize()
    On Error Resume Next
    Dim pH As Long
    Dim pW As Long
    pH = Me.ScaleHeight / 2 - 3 * 120
    pW = Me.ScaleWidth - 2 * 120
    
    ctlSplitterEx.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
    
    'fraIE.Move 120, 120, pW, pH
    'fraBookList.Move 120, fraIE.Top + fraIE.Height + 120, pW, pH
    Dim i As Long

    With fraBookList

'        FraButtons.Move 0, 0
'        FraButtons.Width = cmd(0).Width + cstLeft
'        FraButtons.Height = .Height - FraButtons.Top
'
'        lstBooks.Move FraButtons.Left + FraButtons.Width + cstLeft, 0
'        lstBooks.Width = .Width - lstBooks.Left - cstLeft
'        lstBooks.Height = .Height
'
'
'        cmd(0).Move 0, 0
'        For i = 1 To cmd.UBound
'            cmd(i).Move cmd(i - 1).Left, cmd(i - 1).Top + cmd(i - 1).Height + 2 * cstTop
'        Next

        FraButtons.Move cstLeft, cstTop
        FraButtons.Width = .Width - 2 * cstLeft ' - 2 * cstLeft ' cmd(0).Width + cstLeft
        FraButtons.Height = cmd(0).Height + cstTop '.Height - FraButtons.Top

        lstBooks.Move cstLeft, FraButtons.Top + FraButtons.Height + cstTop
        lstBooks.Width = .Width - 2 * cstLeft
        lstBooks.Height = .Height - lstBooks.Top - cstTop


        cmd(0).Move 0, 0
        For i = 1 To cmd.UBound
            cmd(i).Move cmd(i - 1).Left + cmd(i - 1).Width + cstLeft, 0 ', cmd(i - 1).Top + cmd(i - 1).Height + 2 * cstTop
        Next




        'FrameTaskInfo.Left =
        'FrameTaskInfo.Width = fraTabs(1).Width - FrameTaskInfo.Left - 2 * cstLeft



    End With

    With fraIE
        picAddress.Width = .Width - 120
        cboAddress.Width = picAddress.Width - cboAddress.Left
        IE.Width = .Width - 120
        IE.Height = .Height - (picAddress.Top + picAddress.Height) - 120
    End With


    
End Sub




Private Function GetBookInfo(ByRef vText As String) As CBookInfo
    Dim vBook As CBookInfo
    Dim vResult() As String
    vResult = SSLIB_ParseInfoText(vText)
    Set vBook = New CBookInfo
    vBook.LoadFromArray vResult(), True
    Set GetBookInfo = vBook
End Function



Private Sub lstBooks_DblClick()
    ActionOnItem "阅读|下载"
End Sub



Private Sub OLDIE_NewWindow(ByVal URL As String, ByVal flags As Long, ByVal TargetFrameName As String, PostData As Variant, ByVal Headers As String, Processed As Boolean)
    IE.Navigate2 URL
    Processed = True
End Sub

Private Sub tbToolBar_ButtonClick(ByVal Button As Button)
    On Error Resume Next
      

   ' timTimer.Enabled = True
      

    Select Case Button.Key
        Case "Back"
            IE.GoBack
        Case "Forward"
            IE.GoForward
        Case "Refresh"
            IE.Refresh
        Case "Home"
            IE.Navigate2 mHomePage
        Case "Search"
            IE.GoSearch
        Case "Stop"
            'timTimer.Enabled = False
            IE.Stop
            Me.Caption = IE.LocationName
    End Select


End Sub


Private Sub timTimer_ThatTime()
   ' TimerCheck
End Sub

