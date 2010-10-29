VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBrowser 
   ClientHeight    =   8415
   ClientLeft      =   3060
   ClientTop       =   3450
   ClientWidth     =   10770
   LinkTopic       =   "Form1"
   ScaleHeight     =   8415
   ScaleWidth      =   10770
   Begin VB.Frame fraBookList 
      BorderStyle     =   0  'None
      Height          =   3030
      Left            =   465
      TabIndex        =   5
      Top             =   5220
      Width           =   9360
      Begin VB.Frame FraButtons 
         BorderStyle     =   0  'None
         Height          =   3060
         Left            =   135
         TabIndex        =   7
         Top             =   195
         Width           =   1650
         Begin VB.CommandButton cmd 
            Caption         =   "更新"
            Height          =   360
            Index           =   3
            Left            =   345
            TabIndex        =   11
            Top             =   1830
            Width           =   1125
         End
         Begin VB.CommandButton cmd 
            Caption         =   "清除列表"
            Height          =   360
            Index           =   2
            Left            =   315
            TabIndex        =   10
            Top             =   1335
            Width           =   1125
         End
         Begin VB.CommandButton cmd 
            Caption         =   "下载"
            Height          =   360
            Index           =   0
            Left            =   240
            TabIndex        =   9
            Top             =   360
            Width           =   1125
         End
         Begin VB.CommandButton cmd 
            Caption         =   "删除"
            Height          =   360
            Index           =   1
            Left            =   270
            TabIndex        =   8
            Top             =   885
            Width           =   1125
         End
      End
      Begin MSComctlLib.ListView lstBooks 
         Height          =   1035
         Left            =   2895
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
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ID"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "SSID"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Title"
            Object.Width           =   5080
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Author"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Publisher"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "PageURL"
            Object.Width           =   14111
         EndProperty
      End
   End
   Begin VB.Frame fraIE 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4485
      Left            =   330
      TabIndex        =   0
      Top             =   480
      Width           =   11280
      Begin VB.PictureBox picAddress 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   0
         ScaleHeight     =   495
         ScaleWidth      =   10935
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   0
         Width           =   10935
         Begin VB.ComboBox cboAddress 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3000
            TabIndex        =   2
            Top             =   120
            Width           =   3795
         End
         Begin MSComctlLib.Toolbar tbToolBar 
            Height          =   450
            Left            =   0
            TabIndex        =   4
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
         TabIndex        =   3
         Top             =   495
         Width           =   5400
         ExtentX         =   9513
         ExtentY         =   6586
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
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
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   2670
      Top             =   2265
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
            Picture         =   "frmBrowser.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":0712
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":0E24
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":1536
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":1C48
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":235A
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public StartingAddress As String
Dim mbDontNavigateNow As Boolean
Private Const CST_SSLIBRARY_HOST As String = "sslibrary.com"
Private mBookListUrl As String
Private Const cstLeft As Long = 120
Private Const cstTop As Long = 120
Private WithEvents timTimer As CTimer
Attribute timTimer.VB_VarHelpID = -1
Private Const cstTimerInterval As Long = 400

Private mScriptHost As IHTMLWindow2

Public LastAddress As String
Private mSkipAddress As Boolean
Private Const configFile As String = "Browser.ini"
Private mPassword As String
Private mUsername As String
Private mHomePage As String

Private Sub ActionOnItem(ByVal vAction As String)
    
    If vAction = "清除列表" Then
        mBookListUrl = ""
        lstBooks.ListItems.Clear
        Exit Sub
    ElseIf vAction = "更新" Then
        mBookListUrl = ""
        AddSearchResult IE.LocationURL, IE.Document
        Exit Sub
    End If
    If lstBooks.SelectedItem Is Nothing Then
        MsgBox "No tasks selected.", vbInformation, vAction
        Exit Sub
    End If
    
    
    
    Dim selectListItem As ListItem
    Dim selected() As ListItem
    Dim Count As Long
    For Each selectListItem In lstBooks.ListItems
        If selectListItem.selected Then
            Count = Count + 1
            ReDim Preserve selected(1 To Count)
            Set selected(Count) = selectListItem
        End If
    Next
    
    Dim i As Long
    
    Select Case vAction
    
        Case "下载"
            Dim bookInfo As CBookInfo
            For i = 1 To Count
                Set bookInfo = selected(i).Tag
                If Not bookInfo Is Nothing Then
                    AddTask bookInfo
                    Exit Sub
                End If
            Next
        Case "删除"
            For i = Count To 1 Step -1
                lstBooks.ListItems.Remove selected(i).Key
            Next
            For Each selectListItem In lstBooks.ListItems
                selectListItem.text = selectListItem.Index
            Next
    End Select
End Sub

Private Sub AddTask(ByRef vBookInfo As CBookInfo)
    On Error Resume Next
    Dim pWin As IHTMLWindow2
    If mScriptHost Is Nothing Then Set pWin = IE.Document.parentWindow Else Set pWin = mScriptHost
    'Set pWin = lstBooks.Tag
    If pWin Is Nothing Then Exit Sub
    Dim pUrl As String
    Dim pScript As String
    'pUrl = "/" & SubStringBetween(vBookInfo(SSF_PageURL), "('", "')", True)
    'pWin.Navigate pUrl
    pScript = Replace(Mid$(vBookInfo(SSF_PageURL), Len("javascript:") + 1), "readbook2", "readbook")
    pWin.execScript pScript
    
    
    frmTask.InitWithBookInfo vBookInfo
    frmTask.Show
    
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

Public Sub Init(vHomePage As String, vUsername As String, vPassword As String)
    mHomePage = vHomePage
    mPassword = vPassword
    mUsername = vUsername
End Sub

Private Sub Form_Load()
    On Error Resume Next
    
    
    
    Me.Show
    tbToolBar.Refresh
    Form_Resize
    
 SSLIB_Init

    'cboAddress.Move 50, lblAddress.Top + lblAddress.Height + 15
    
    
    Dim configHnd As CLiNInI
    Set configHnd = New CLiNInI
    With configHnd
        .source = App.Path & "\" & configFile
        ComboxItemsFromString cboAddress, .GetSetting("Browser", "History")
        FormStateFromString Me, .GetSetting("Form", "State")
        StartingAddress = .GetSetting("Browser", "LastAddress")
    End With
    Set configHnd = Nothing
    
    
    If StartingAddress = "" Then StartingAddress = mHomePage

    Set timTimer = New CTimer
    timTimer.Interval = 0
    

        IE.Navigate2 StartingAddress
    
'    If Len(StartingAddress) > 0 Then
'        cboAddress.text = StartingAddress
'        cboAddress.AddItem cboAddress.text
'        'try to navigate to the starting address
'
'        'IE.Navigate StartingAddress
'    End If

   
   ' AutoLogin
    
End Sub


Private Sub TimerCheck()
    On Error Resume Next
    Dim elm As IHTMLElement
    Dim elms As IHTMLElementCollection
    Set elms = IE.Document.getElementsByTagName("iframe")
    For Each elm In elms
        CheckUrl elm.contentWindow.location.href, elm.contentWindow
    Next
    CheckUrl IE.LocationURL, IE.Document.parentWindow
    Debug.Print Err.Number & ":" & Err.Description
    Err.Clear
End Sub


Private Sub Form_Unload(Cancel As Integer)

    Set timTimer = Nothing
    On Error Resume Next
     Dim configHnd As CLiNInI
    Set configHnd = New CLiNInI
    With configHnd
        .source = App.Path & "\" & configFile
        .SaveSetting "Browser", "History", ComboxItemsToString(cboAddress)
        .SaveSetting "Form", "State", FormStateToString(Me)
        .SaveSetting "Browser", "LastAddress", IE.LocationURL
        .Save
    End With
    Set configHnd = Nothing
    
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
    'Debug.Print vURL
    Dim pUrl As String
    pUrl = LCase$(vUrl)
    If vWindow Is Nothing Then Set vWindow = IE
    
    Dim vDocument As HTMLDocument
    Dim vBody As IHTMLBodyElement2
    Dim vLocation As HTMLLocation
    Dim vDocElement As HTMLHtmlElement
    
    Set vDocument = vWindow.Document
    Set vBody = vDocument.body
    Set vLocation = vDocument.location
    Set vDocElement = vDocument.documentElement
    
    Dim pHref As String
    Dim pHost As String
    Dim pPath As String
    pHref = LCase$(vLocation.href)
    pHost = LCase$(vLocation.host)
    pPath = LCase$(vLocation.PathName)
    
    If Right$(pHost, Len(CST_SSLIBRARY_HOST)) = CST_SSLIBRARY_HOST Then
        
        If pPath = "/userlogon.jsp" Then
            AutoLogin vDocument
        ElseIf pPath = "/search.jsp" Then
            AddSearchResult pHref, vDocument
        ElseIf pPath = "/books.jsp" Then
            AddSearchResult pHref, vDocument
        End If
        
    End If
    
    
    
End Sub

Private Sub AddSearchResult(ByRef pHref As String, ByRef vDoc As HTMLDocument)
    On Error Resume Next
    If pHref = mBookListUrl Then Exit Sub
    Dim pBooks As Collection
    Set pBooks = GetBooksInfo(vDoc)
    If pBooks.Count > 0 Then
        Set mScriptHost = vDoc.parentWindow '  .parentWindow
        mBookListUrl = pHref
        Dim pBook As CBookInfo
        Dim pItem As ListItem
        For Each pBook In pBooks
            Set pItem = Nothing
            Set pItem = lstBooks.ListItems.Add(, pBook(SSF_Title) & pBook(SSF_SSID))
            If Not pItem Is Nothing Then
                Set pItem.Tag = pBook
                pItem.text = pItem.Index
                pItem.SubItems(1) = pBook(SSF_SSID)
                pItem.SubItems(2) = pBook(SSF_Title)
                pItem.SubItems(3) = pBook(SSF_AUTHOR)
                pItem.SubItems(4) = pBook(SSF_PublishDate)
                pItem.SubItems(5) = pBook(SSF_PageURL)
            End If
        Next
        timTimer.Interval = 0
    End If
    Err.Clear
End Sub

Private Function GetBooksInfo(ByRef vDoc As HTMLDocument) As Collection
    Dim Result As Collection
    Set Result = New Collection
    On Error Resume Next
    Dim divResult As IHTMLDivElement
    If Not FindElement(divResult, vDoc, "divResult") Then Exit Function
    'lstBooks.ListItems.Clear
    Dim elms As IHTMLElementCollection
    Set elms = divResult.getElementsByTagName("TABLE")
    Dim elm As IHTMLElement
    
    Dim vBookInfo As CBookInfo
    For Each elm In elms
        Set vBookInfo = GetBookInfo(elm.innerText)
        Debug.Print vBookInfo(SSF_Title)
        If vBookInfo(SSF_Title) <> "" Then
            
            Dim elm_a As IHTMLAnchorElement
            Dim elm_as As IHTMLElementCollection
            Set elm_as = elm.getElementsByTagName("a")
            For Each elm_a In elm_as
                If InStr(1, elm_a.href, "javascript:readbook(", vbTextCompare) = 1 Then
                    vBookInfo(SSF_SSID) = SubStringBetween(elm_a.href, "showbook.do?dxNumber=", "&", True)
                    vBookInfo(SSF_PageURL) = elm_a.href
                    Exit For
                End If
            Next
            Result.Add vBookInfo
            'lstBooks.ListItems.Add , , vBookInfo(SSF_Title)
        End If
        'Debug.Print elm.innerText
        'lstBooks.ListItems.Add , , elm.innerText
    Next
    
    Set GetBooksInfo = Result
End Function

Private Sub AutoLogin(ByRef vDoc As HTMLDocument)
    
    Debug.Print "AutoLogin..."
    
    
    'Dim pHtmlElm As IHTMLElement
    
    If Not SetAttribute(vDoc, "UserName", "value", mUsername) Then Exit Sub
    If Not SetAttribute(vDoc, "PassWord", "value", mPassword) Then Exit Sub
    
'    If FindElement(pHtmlElm, vDoc, "Submit3322") Then
'        pHtmlElm.Click
'        timTimer.Interval = 0
'    End If
    
    'mSkipAddress = True
    'vDoc.parentWindow.execScript
    'LastAddress = vDoc.location.href
    
    'IE.Navigate2 "http://edu.sslibrary.com/"
End Sub
Private Sub IE_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
    timTimer.Interval = cstTimerInterval
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
    fraIE.Move 120, 120, pW, pH
    fraBookList.Move 120, fraIE.Top + fraIE.Height + 120, pW, pH
    Dim i As Long
    
    With fraBookList
        
        FraButtons.Move 0, 0
        FraButtons.Width = cmd(0).Width + cstLeft
        FraButtons.Height = .Height - FraButtons.Top
    
        lstBooks.Move FraButtons.Left + FraButtons.Width + cstLeft, 0
        lstBooks.Width = .Width - lstBooks.Left - cstLeft
        lstBooks.Height = .Height
        
        
        cmd(0).Move 0, 0
        For i = 1 To cmd.UBound
            cmd(i).Move cmd(i - 1).Left, cmd(i - 1).Top + cmd(i - 1).Height + 2 * cstTop
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
    ActionOnItem "下载"
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
    TimerCheck
End Sub
