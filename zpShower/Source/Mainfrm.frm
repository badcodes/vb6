VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form MainFrm 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   ClientHeight    =   3720
   ClientLeft      =   132
   ClientTop       =   720
   ClientWidth     =   7812
   Icon            =   "Mainfrm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   NegotiateMenus  =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   7812
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Interval        =   2500
      Left            =   4068
      Top             =   3036
   End
   Begin VB.PictureBox picSplitter 
      Appearance      =   0  'Flat
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   4800
      Left            =   3084
      ScaleHeight     =   2079.675
      ScaleMode       =   0  'User
      ScaleWidth      =   520
      TabIndex        =   6
      Top             =   -1596
      Visible         =   0   'False
      Width           =   72
   End
   Begin MSComctlLib.StatusBar StsBar 
      Height          =   372
      Left            =   -288
      TabIndex        =   1
      Top             =   3384
      Width           =   8076
      _ExtentX        =   14245
      _ExtentY        =   656
      ShowTips        =   0   'False
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9292
            Key             =   "ie"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3440
            MinWidth        =   3440
            Key             =   "reading"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1418
            MinWidth        =   1411
            Key             =   "order"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList Listimg 
      Left            =   4848
      Top             =   2844
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mainfrm.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mainfrm.frx":11A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mainfrm.frx":1A7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mainfrm.frx":2358
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame theShow 
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   2784
      Left            =   3264
      TabIndex        =   0
      Top             =   60
      Width           =   3924
      Begin VB.HScrollBar HBoxScroll 
         CausesValidation=   0   'False
         Height          =   252
         LargeChange     =   100
         Left            =   108
         SmallChange     =   10
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   2472
         Width           =   3180
      End
      Begin VB.VScrollBar VBOXScroll 
         CausesValidation=   0   'False
         Height          =   2592
         LargeChange     =   100
         Left            =   3684
         SmallChange     =   10
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   132
         Width           =   252
      End
      Begin VB.Image ImageBox 
         Appearance      =   0  'Flat
         Height          =   7200
         Left            =   84
         Stretch         =   -1  'True
         Top             =   180
         Width           =   9600
      End
   End
   Begin VB.PictureBox LeftFrame 
      Height          =   2772
      Left            =   300
      ScaleHeight     =   2724
      ScaleWidth      =   2520
      TabIndex        =   2
      Top             =   84
      Width           =   2568
      Begin VB.Frame ListFrame 
         BorderStyle     =   0  'None
         Height          =   2436
         Left            =   228
         TabIndex        =   3
         Top             =   324
         Width           =   2532
         Begin MSComctlLib.TreeView List 
            Height          =   1452
            Left            =   -48
            TabIndex        =   4
            Top             =   444
            Width           =   2292
            _ExtentX        =   4043
            _ExtentY        =   2561
            _Version        =   393217
            HideSelection   =   0   'False
            Indentation     =   423
            LabelEdit       =   1
            LineStyle       =   1
            PathSeparator   =   "/"
            Style           =   7
            ImageList       =   "Listimg"
            BorderStyle     =   1
            Appearance      =   1
         End
      End
      Begin MSComctlLib.TabStrip LeftStrip 
         Height          =   4812
         Left            =   108
         TabIndex        =   5
         Top             =   24
         Width           =   2412
         _ExtentX        =   4255
         _ExtentY        =   8488
         MultiRow        =   -1  'True
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   1
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Files"
               Key             =   "TABfile"
               Object.ToolTipText     =   "FileList"
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
      Begin VB.Image imgSplitter 
         Appearance      =   0  'Flat
         Height          =   5172
         Left            =   0
         MousePointer    =   9  'Size W E
         Top             =   96
         Width           =   104
      End
   End
   Begin VB.Menu mnu 
      Caption         =   "&File"
      Index           =   0
      Begin VB.Menu mnuFile_Open 
         Caption         =   "&Open File"
      End
      Begin VB.Menu mnuFile_Close 
         Caption         =   "&Close"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuFile_Preference 
         Caption         =   "&Preference"
      End
      Begin VB.Menu mnuFile_Recent 
         Caption         =   "-"
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFile_Exit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnu 
      Caption         =   "&Mode"
      Index           =   1
      Begin VB.Menu mnu_mode_fit_width 
         Caption         =   "Fit Width"
      End
      Begin VB.Menu mnu_mode_fit_height 
         Caption         =   "Fit Height"
      End
      Begin VB.Menu mnu_mode_fit_viewer 
         Caption         =   "Fit Viewer"
      End
      Begin VB.Menu mnu_mode_zoom_in 
         Caption         =   "Zoom In"
      End
      Begin VB.Menu mnu_mode_zoom_out 
         Caption         =   "Zoom Out"
      End
      Begin VB.Menu mnu_mode_keep_original 
         Caption         =   "Keep Original"
      End
   End
   Begin VB.Menu mnu 
      Caption         =   "&View"
      Index           =   2
      Begin VB.Menu mnuView_Left 
         Caption         =   "&Left"
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnuView_Menu 
         Caption         =   "&Menu"
         Checked         =   -1  'True
         Shortcut        =   {F8}
      End
      Begin VB.Menu mnuView_StatusBar 
         Caption         =   "&StatusBar"
         Checked         =   -1  'True
         Shortcut        =   {F9}
      End
      Begin VB.Menu mnuView_FullScreen 
         Caption         =   "&FullScreen"
         Shortcut        =   {F11}
      End
      Begin VB.Menu mnuView_TopMost 
         Caption         =   "&TopMost Made"
      End
   End
   Begin VB.Menu mnu 
      Caption         =   "&Go"
      Index           =   3
      Begin VB.Menu mnuGo_Back 
         Caption         =   "&Back         Alt+Left"
      End
      Begin VB.Menu mnuGo_Forward 
         Caption         =   "&Forward         Alt+Right"
      End
      Begin VB.Menu mnuGo_Previous 
         Caption         =   "&Previous         Alt+Down"
      End
      Begin VB.Menu mnuGo_Next 
         Caption         =   "&Next         Alt+Up"
      End
      Begin VB.Menu mnuGo_Random 
         Caption         =   "&Random         Alt+Z"
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGo_AutoNext 
         Caption         =   "AutoNext"
      End
      Begin VB.Menu mnuGo_AutoRandom 
         Caption         =   "&AutoRandom"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuSep6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGo_Home 
         Caption         =   "&Home         Alt+Home"
      End
   End
   Begin VB.Menu mnu 
      Caption         =   "&Directory"
      Index           =   4
      Begin VB.Menu mnuDir_readPrev 
         Caption         =   "Previous File"
      End
      Begin VB.Menu mnuDir_readNext 
         Caption         =   "Next File"
      End
      Begin VB.Menu mnuDir_random 
         Caption         =   "Random File"
      End
      Begin VB.Menu mnuDir_delete 
         Caption         =   "Delete File"
      End
   End
   Begin VB.Menu mnu 
      Caption         =   "&Bookmark"
      Index           =   5
      Begin VB.Menu mnuBookmark_Add 
         Caption         =   "&Add"
      End
      Begin VB.Menu mnuBookmark_Manage 
         Caption         =   "&Manage"
      End
      Begin VB.Menu mnuBookmark 
         Caption         =   "-"
         Index           =   0
      End
   End
   Begin VB.Menu mnu 
      Caption         =   "&Help"
      Index           =   6
      Begin VB.Menu mnuHelp_BookInfo 
         Caption         =   "&Book Info"
      End
      Begin VB.Menu mnuHelp_About 
         Caption         =   "&About This"
      End
   End
   Begin VB.Menu mnu 
      Caption         =   "<<<"
      Index           =   7
   End
   Begin VB.Menu mnu 
      Caption         =   ">>>"
      Index           =   8
   End
End
Attribute VB_Name = "MainFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''
'Method_description.<BR>
'Can_be_more_than_one_line.
'
'@author xrLiN
'@version 1.0
'@date 2006-07-19
Option Explicit


Private Enum ListStatusConstant
lstLoaded = 1
lstNotLoaded = 2
lstNotExists = 0
End Enum

'Gobal Var
Private Const sTaglstLoaded = "ListLoaded"
Private Const staglstNotLoaded = "ListNotLoaded"
'Private Const sTaglstNotExists = "ListNotExists"
Private Const sglSplitLimit = 500
Private Const minFormHeight = 2000
Private Const minFormWidth = 3000
'Private Const minLeftWidth = 1500
'Private Const icstDefaultFormHeight = 6000
'Private Const icstDefaultFormWidth = 8000
Private Const icstDefaultLeftWidth = 1500
Private Const lcstFittedListItemsNum = 3000
Private Const defaultRecentFileList = 5
Private Const defaultAutoViewTime = 2500


'Private InvalidPWD As Boolean
'Public Navigated As Boolean
Public NotPreOperate As Boolean
'Private StopNow As Boolean
Private NotResize As Boolean
'Private bReloadContent As Boolean

'Public iCurIEView As Integer
'Public bIsZhtm As Boolean

Private sFilesInZip As New CStringVentor 'Collection
'Private sFoldersInZip As cStringventor  ' CStringCollection
'Private sFilesINContent(1 To 2) As New CStringVentor  'CStringCollection
'Private sfilesinzip.count As Long
'Private lFoldersIZcount As Long
'Public bFullScreen As Boolean
'Private WithEvents ieViewV1  As SHDocVwCtl.WebBrowser_V1
'Private WithEvents ieview.document As HTMLDocument
Public WithEvents lUnzip As cUnzip
Attribute lUnzip.VB_VarHelpID = -1
Public WithEvents lZip As cZip
Attribute lZip.VB_VarHelpID = -1
Private bInValidPassword  As Boolean
Private bAutoShowNow As Boolean
Private bRandomShow As Boolean

Private Const zhMemFile = "memory.ini"

Private bScroll As Boolean
Private scrollX As Single
Private scrollY As Single
Public ZOOM_FACTOR As Single

Enum ViewMode
    FIT_WIDTH
    FIT_HEIGHT
    FIT_VIEWER
    ZOOM_OUT
    ZOOM_IN
    KEEP_ORIGIN
End Enum


Public Function curFileIndex(sCurFileInZip As String) As Long

    'If zhrStatus.iListIndex = lwContent Then
    '    curFileIndex = sFilesINContent(2).Index(sCurFileInZip)
    'Else
        curFileIndex = sFilesInZip.Index(sCurFileInZip)
    'End If

End Function

Public Function execZhcommand(ByVal sCommandLine As String) As Boolean

    Dim cmdName As String
    Dim cmdArgu As String
    On Error GoTo Herr
    execZhcommand = True
    sCommandLine = Replace$(sCommandLine, "\", "/")
    sCommandLine = LeftRight(sCommandLine, "://", vbTextCompare, ReturnEmptyStr)
    cmdName = LeftLeft(sCommandLine, "|", vbTextCompare, ReturnOriginalStr)
    cmdName = RightDelete(cmdName, "/")
    cmdArgu = LeftRight(sCommandLine, "|", vbTextCompare, ReturnEmptyStr)
    cmdArgu = LeftDelete(cmdArgu, "/")
    cmdArgu = RightDelete(cmdArgu, "/")

    If cmdArgu = "" Then
        CallByName Me, cmdName, VbMethod
    Else
        CallByName Me, cmdName, VbMethod, cmdArgu
    End If

    Exit Function
Herr:
    execZhcommand = False

End Function

'Function getZhCommentText(ByVal sZipfilename As String) As String
'
'    'Set lUnzip = New cUnzip
'    lUnzip.ZipFile = sZipfilename
'    lUnzip.Comment = ""
'    getZhCommentText = toUnixPath(lUnzip.GetComment)
'    'Set lUnzip = Nothing
'
'End Function

'Private Function ieview.document_oncontextmenu() As Boolean
'
'    If bAutoShowNow = False Then
'        ieview.document_oncontextmenu = True
'    End If
'
'End Function

Private Function isListLoaded(trvwlist As TreeView) As ListStatusConstant

    On Error GoTo trvwListnotExists

    If trvwlist.Tag = sTaglstLoaded Then
        isListLoaded = lstLoaded
    Else
        isListLoaded = lstNotLoaded
    End If

    Exit Function
trvwListnotExists:
    isListLoaded = lstNotExists

End Function

Public Function isZhcommand(ByVal sUrl As String) As Boolean

    If LCase$(Left$(sUrl, 6)) = "zhcmd:" Then isZhcommand = True

End Function

Public Function randomView(Optional bRestart As Boolean = False) As Boolean

    Static curZhtm As String
    Static iViewNow As Long
    Static iViewLast As Long
    Static iRandomArr() As Long

    Dim i As Long
    

    If zhrStatus.sCur_zhFile = "" Then Exit Function

    If curZhtm <> zhrStatus.sCur_zhFile Or bRestart Then
        
        curZhtm = zhrStatus.sCur_zhFile
        iViewNow = 0

        'If lastFileList = lwContent Then
        '    iViewLast = sFilesINContent(1).Count
        'Else
            iViewLast = sFilesInZip.Count
        'End If
        
        If iViewLast < 1 Then
            randomView = False
            Exit Function
        End If
        
        ReDim iRandomArr(1 To iViewLast) As Long
        For i = 1 To iViewLast
            iRandomArr(i) = i
        Next
        MAlgorithms.BedlamArr iRandomArr, 1, iViewLast '打乱数组
    End If

    iViewNow = iViewNow + 1

    If iViewNow > iViewLast Then iViewNow = 1

    'If lastFileList = lwContent Then
    '    GetView sFilesINContent(1).Value(iRandomArr(iViewNow))
    'Else
        GetView sFilesInZip.value(iRandomArr(iViewNow))
   ' End If

    randomView = True

End Function

Private Function setListStatus(trvwlist As TreeView, lstStatus As ListStatusConstant) As Boolean

    On Error GoTo trvwListnotExists

    If lstStatus = lstLoaded Then
        trvwlist.Tag = sTaglstLoaded
    ElseIf lstStatus = lstNotLoaded Then
        trvwlist.Tag = staglstNotLoaded
    End If

    setListStatus = True
    Exit Function
trvwListnotExists:
    setListStatus = False

End Function

Public Function StrLocalize(sEnglish As String) As String

    Dim zhLocalize As New CLocalize
    zhLocalize.Install Me, LanguageIni
    StrLocalize = zhLocalize.loadLangStr(sEnglish)
    Set zhLocalize = Nothing

End Function

Public Sub appAbout()

'    Dim fNum As Integer
'    Dim sAppHtmlAboutFile As String
'    Dim sAppend As New CAppendString
'
'    sAppHtmlAboutFile = linvblib.BuildPath(Tempdir, cHtmlAboutFilename)
'
'    If linvblib.PathExists(sAppHtmlAboutFile) = False Then
'
'        With sAppend
'            .AppendLine "<html>"
'            .AppendLine htmlline("<link REL=!stylesheet! href=!" & sConfigDir & "/style.css! type=!text/css!>")
'            .AppendLine htmlline("<head>")
'            .AppendLine htmlline("<Title>Zippacked Html Reader</title>")
'            .AppendLine htmlline("<meta http-equiv=!Content-Type! content=!text/html; charset=us-ascii!>")
'            .AppendLine htmlline("</head>")
'            .AppendLine htmlline("<body class=!m_text!><p><br>")
'            .AppendLine htmlline("<p align=right ><span lang=EN-US style=!font-size:24.0pt;font-family:TAHOMA,Courier New! class=!m_text!>" & App.ProductName & " (Build" & Str$(App.Major) + "." + Str$(App.Minor) & "." & Str$(App.Revision) & ")</span></p>")
'            .AppendLine htmlline("<p align=right ><span lang=EN-US style=!font-size:24.0pt;font-family:TAHOMA,Courier New! class=!m_text!>" & App.LegalCopyright & "</span></span></p>")
'            .AppendLine htmlline("</body>")
'            .AppendLine htmlline("</html>")
'        End With
'
'        fNum = FreeFile
'        Open sAppHtmlAboutFile For Output As #fNum
'        Print #fNum, sAppend.Value
'        Close #fNum
'        Set sAppend = Nothing
'    End If
'
'    NotPreOperate = True
'    IEView.Navigate2 sAppHtmlAboutFile

End Sub


'Private Sub ApplyDefaultStyle(fApply As Boolean)
'
'    Const CSSID = "zhReaderCSS"
'    Dim curCss As String
'    Dim iIndex As Long
'    Dim ALLCSS As HTMLStyleSheetsCollection
'    Dim ICSS As IHTMLStyleSheet
'
'
'    'On Error GoTo 0
'
'    Set ALLCSS = IEView.Document.styleSheets
'    For Each ICSS In ALLCSS
'    If ICSS.Title = CSSID Then ICSS.href = "": ICSS.Title = ""
'    Next
'
'    If fApply = False Then Exit Sub
'
'    iIndex = Val(mnuView_ApplyStyleSheet.Tag)
'    If iIndex = 0 Then Exit Sub
'    If iIndex = 1 Then
'        curCss = linvblib.bdUnixDir(sConfigDir, "Style.css")
'        If linvblib.PathExists(curCss) = False Then
'                Load frmOptions
'                frmOptions.MakeCss curCss
'                Unload frmOptions
'        End If
'    Else
'        curCss = mnuView_ApplyStyleSheet_List(iIndex).Tag
'    End If
'
'    'if linvblib.PathExists (curcss)=False then
'    Set ICSS = IEView.Document.createStyleSheet(curCss)
'    ICSS.Title = CSSID
'
'
'End Sub

Public Sub ChangeMainFrmCaption(sTitle As String)

    If mnuView_FullScreen.Checked Then Exit Sub
    MainFrm.Caption = sTitle

End Sub

'Public Sub Load_StyleSheetList()
'    Dim fso As New FileSystemObject
'    Dim fs As Files
'    Dim f As file
'    Dim iIndex As Long
'    Dim iLast As Long
'    On Error Resume Next
'
'    iLast = mnuView_ApplyStyleSheet_List.UBound
'    For iIndex = iLast To 3 Step -1
'    Unload mnuView_ApplyStyleSheet_List(iIndex)
'    Next
'
'    iIndex = 2
'
'    Set fs = fso.GetFolder(fso.BuildPath(App.Path, "CSS")).Files
'
'    If TypeName(fs) <> "Nothing" Then
'        For Each f In fs
'            iIndex = iIndex + 1
'            Load mnuView_ApplyStyleSheet_List(iIndex)
'            mnuView_ApplyStyleSheet_List(iIndex).Checked = False
'            mnuView_ApplyStyleSheet_List(iIndex).Visible = True
'            mnuView_ApplyStyleSheet_List(iIndex).Caption = f.name
'            mnuView_ApplyStyleSheet_List(iIndex).Tag = f.Path
'        Next
'    End If
'
'    Set f = Nothing
'    Set fs = Nothing
'    Set fs = fso.GetFolder(fso.BuildPath(sConfigDir, "CSS")).Files
'    If TypeName(fs) <> "Nothing" Then
'        For Each f In fs
'            iIndex = iIndex + 1
'            Load mnuView_ApplyStyleSheet_List(iIndex)
'            mnuView_ApplyStyleSheet_List(iIndex).Checked = False
'            mnuView_ApplyStyleSheet_List(iIndex).Visible = True
'            mnuView_ApplyStyleSheet_List(iIndex).Caption = f.name
'            mnuView_ApplyStyleSheet_List(iIndex).Tag = f.Path
'        Next
'    End If
'    Set fso = Nothing
'End Sub

'Private Sub cmbAddress_Click()
'
'    Dim txtCmb As String
'    txtCmb = cmbAddress.text
'
'    If txtCmb <> "" Then MNavigate txtCmb, IEView
'
'End Sub

'Private Sub cmbAddress_KeyPress(KeyAscii As Integer)
'
'    Select Case KeyAscii
'    Case Asc(vbCr)
'        Dim txtCmb As String
'        txtCmb = cmbAddress.text
'
'        If txtCmb <> "" Then MNavigate txtCmb, IEView
'        AddUniqueItem cmbAddress, txtCmb
'        '        Dim iIndex As Long
'        '        Dim iEnd As Long
'        '        iEnd = cmbAddress.ListCount - 1
'        '        For iIndex = 0 To iEnd
'        '        If StrComp(cmbAddress.List(iIndex), txtCmb, vbTextCompare) = 0 Then Exit Sub
'        '        Next
'        '        cmbAddress.AddItem txtCmb
'    Case Asc(vbKeyEscape)
'
'        If cmbAddress.ListCount > 1 Then cmbAddress.text = cmbAddress.List(cmbAddress.ListCount - 1) '  IEView.LocationURL
'    End Select
'
'End Sub
Private Sub ScrollBy(scrollObject As Object, value As Long)
    Dim cV As Long
    Dim mV As Long
    Dim nV As Long
    
    On Error GoTo ErrScrollBye
    
    If scrollObject.Visible = False Then Exit Sub
    With scrollObject
 
        cV = .value
        mV = .Max
        nV = cV + value
        
        If nV > mV Then nV = mV
        If nV < .Min Then nV = .Min
        .value = nV
    
    End With

ErrScrollBye:

End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim iKeyCode As Integer
    iKeyCode = KeyCode
    KeyCode = 0

    If Shift = 0 Then

        Select Case iKeyCode
        Case vbKeyF7
            mnuView_Left_Click
        Case vbKeyF8
            mnuView_Menu_Click
        Case vbKeyF9
            mnuView_StatusBar_Click
        Case vbKeyF11
            mnuView_fullscreen_Click
        Case vbKeyF12
            mnuGo_AutoRandom_Click
        Case vbKeyF4
            mnufile_Close_Click
        Case vbKeyF1
            mnuHelp_BookInfo_Click
        Case vbKeyEscape
            bAutoShowNow = False
            'IEView.Stop
            Timer.Enabled = False
            mnuGo_AutoRandom.Checked = False
        Case vbKeyPageDown
             ScrollBy VBOXScroll, VBOXScroll.LargeChange
        Case vbKeyPageUp
              ScrollBy VBOXScroll, -VBOXScroll.LargeChange
        Case vbKeyUp
             ScrollBy VBOXScroll, -VBOXScroll.SmallChange
        Case vbKeyDown
             ScrollBy VBOXScroll, VBOXScroll.SmallChange
        Case vbKeyLeft
             ScrollBy HBoxScroll, -HBoxScroll.SmallChange
        Case vbKeyRight
            ScrollBy HBoxScroll, HBoxScroll.SmallChange
        Case Else
            KeyCode = iKeyCode
            
        End Select

    ElseIf Shift = vbAltMask Then

        Select Case iKeyCode
        Case vbKeyUp
            mnuGo_Previous_Click
        Case vbKeyDown
            mnuGo_Next_Click
        Case vbKeyHome
            mnuGo_Home_Click
        Case vbKeyF4
            mnufile_exit_Click
        Case vbKeyQ
            mnufile_exit_Click
        Case vbKeyO
            mnufile_Open_Click
        Case vbKeyA
            mnuBookmark_add_Click
        Case vbKeyM
            mnuBookmark_manage_Click
        Case vbKeyP
            mnuFile_PReFerence_Click
        Case vbKeyZ
            mnuGo_Random_Click
'        Case vbKeyAdd
'            zoomBodyFontSize (1)
'        Case vbKeySubtract
'            zoomBodyFontSize (-1)
        Case Else
            KeyCode = iKeyCode
        End Select

    ElseIf Shift = vbCtrlMask Then

        Select Case iKeyCode
        Case vbKeyAdd
            mnu_mode_zoom_in_Click 'IEZoom IEView, 1
        Case vbKeySubtract
            mnu_mode_zoom_out_Click 'IEZoom IEView, -1
        Case Else
            KeyCode = iKeyCode
        End Select

    End If

End Sub
'Public Sub IEZoom(iehost As WebBrowser, inOrOut As Integer)
'    Dim iLevel As Variant
'
'    Const ZoomLevelMin As Integer = 0
'    Const ZoomLevelMax As Integer = 4
'
'
'    iehost.ExecWB OLECMDID_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, Null, iLevel
'
'    iLevel = iLevel + inOrOut
'    If iLevel < ZoomLevelMin Then iLevel = 0
'    If iLevel > ZoomLevelMax Then iLevel = 4
'
'    iehost.ExecWB OLECMDID_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, iLevel, Null
'
'End Sub

'Private Sub Form_KeyPress(KeyAscii As Integer)
'
'    'KeyAscii = 0
'
'End Sub

'
Private Sub Form_Load()

    Dim fso As New FileSystemObject

    
    NotResize = True
    
    'Public String
    sConfigDir = fso.BuildPath(Environ$("APPDATA"), App.EXEName)
    If fso.FolderExists(sConfigDir) = False Then fso.CreateFolder sConfigDir
    zhtmIni = fso.BuildPath(sConfigDir, "config.ini")
    LanguageIni = fso.BuildPath(App.Path, "Language.ini")
    Tempdir = fso.BuildPath(sConfigDir, "Cache")
    If fso.FolderExists(Tempdir) Then fso.DeleteFolder Tempdir, True
    If fso.FolderExists(Tempdir) = False Then fso.CreateFolder Tempdir
    Tempdir = linvblib.toUnixPath(Tempdir)
    
    
    Dim zhLocalize As New CLocalize
    zhLocalize.Install Me, LanguageIni
    zhLocalize.loadFormStr
    Set zhLocalize = Nothing

    Set lUnzip = New cUnzip
     
    On Error Resume Next
    

    If fso.FileExists(zhtmIni) Then
    '-----------------------------------------------------------------------------------
    '-------------------------- Call LoadSetting Start-----------------------------
        Dim hSetting As New CSetting
        hSetting.iniFile = zhtmIni
        hSetting.Load Me, SF_FORM                                               'Postion
        'hsetting.Load Me, SF_FONT                                               'Font For ViewStyle
        'hsetting.Load Me, SF_COLOR                                             'Color For ViewStyle
        hSetting.Load mnuFile_Recent, SF_MENUARRAY                'Recent File
        hSetting.Load mnuBookmark, SF_MENUARRAY                   'BookMark
        hSetting.Load mnuView_Left, SF_CHECKED                        'ShowLeft Check
        hSetting.Load mnuView_Menu, SF_CHECKED                       'ShowMenu Check
        hSetting.Load mnuView_StatusBar, SF_CHECKED                'ShowstatusBar Check
        hSetting.Load mnuView_FullScreen, SF_CHECKED               'FullScreen Check
        hSetting.Load mnuView_TopMost, SF_CHECKED                  'OnTopMost Check
        hSetting.Load Timer, SF_Tag                                             'Time InterVal
        hSetting.Load mnuFile_Open, SF_Tag                                 'InitDir for Dialog
        hSetting.Load Me, SF_Tag                                                  'UseTemplate ?
        hSetting.Load ImageBox, SF_Tag                                            'TemplateHtml
        hSetting.Load Me.picSplitter, SF_Tag 'LeftWidth
        hSetting.Load Me.theShow, SF_COLOR
        hSetting.Load mnu_mode_fit_width, SF_CHECKED
        hSetting.Load mnu_mode_fit_height, SF_CHECKED
        hSetting.Load mnu_mode_fit_viewer, SF_CHECKED
        hSetting.Load mnu_mode_zoom_in, SF_CHECKED
        hSetting.Load mnu_mode_zoom_out, SF_CHECKED
        hSetting.Load mnu_mode_keep_original, SF_CHECKED
        
        Set hSetting = Nothing
    '-------------------------- Call LoadSetting END----------------------------
    '----------------------------------------------------------------------------------
    End If
    mnuView_Left.Checked = Not mnuView_Left.Checked
    mnuView_Menu.Checked = Not mnuView_Menu.Checked
    mnuView_StatusBar.Checked = Not mnuView_StatusBar.Checked
    mnuView_TopMost.Checked = Not mnuView_TopMost.Checked
    mnuView_Left_Click
    mnuView_Menu_Click
    mnuView_StatusBar_Click
    If mnuView_FullScreen.Checked Then
        mnuView_FullScreen.Checked = False
        mnuView_fullscreen_Click
    End If
    mnuView_topmost_Click
    
    
    If picSplitter.Tag = 0 Then picSplitter.Tag = icstDefaultLeftWidth
    picSplitter.Left = picSplitter.Tag '= thers.LeftWidth
    
    If Timer.Tag = 0 Then Timer.Tag = defaultAutoViewTime
    Timer.Interval = Timer.Tag
    If MainFrm.mnuFile_Recent(0).Tag = 0 Then mnuFile_Recent(0).Tag = defaultRecentFileList
    
    If mnuFile_Recent.Count > 1 Then
    mnuFile_Recent(0).Visible = True
    Else
    mnuFile_Recent(0).Visible = False
    End If
    
   
'   Dim hMRU As New CMenuArrHandle
'   With hMRU
'        .maxItem = CLng(mnuFile_Recent(0).Tag)
'         .LoadFromMenus mnuFile_Recent '(0) ' Ini zhtmIni
'        .FillinMenu mnuFile_Recent
'    End With
'    Set hMRU = Nothing
    
'
'    Dim icofile(4) As String
'    Dim i As Integer
'    icofile(1) = fso.BuildPath(App.Path, "images\foldercl.ico")
'    icofile(2) = fso.BuildPath(App.Path, "images\folderop.ico")
'    icofile(3) = fso.BuildPath(App.Path, "images\file.ico")
'    icofile(4) = fso.BuildPath(App.Path, "images\drive.ico")
'
'    For i = 1 To 4
'
'        If fso.PathExists(icofile(i)) Then
'            Listimg.ListImages.Remove (i)
'            Listimg.ListImages.Add i, , LoadPicture(icofile(i))
'        End If
'
'    Next
   
    Dim sInitDir As String
    sInitDir = mnuFile_Open.Tag
    If fso.FolderExists(sInitDir) Then
    ChDir (sInitDir)
    ChDrive (Left$(sInitDir, 1))
    End If
    
    With HBoxScroll
        .SmallChange = 40 * Screen.TwipsPerPixelX
        .LargeChange = 10 * .SmallChange
    End With
    
    With VBOXScroll
        .SmallChange = 40 * Screen.TwipsPerPixelY
        .LargeChange = 10 * .SmallChange
    End With
    NotResize = False


    
End Sub

Private Sub Form_Resize()

    If NotResize Then Exit Sub
    Dim tempint As Integer

    If MainFrm.WindowState = 1 Then Exit Sub
    

        With MainFrm

           If .ScaleHeight < minFormHeight Then .Height = minFormHeight

           If .ScaleWidth < minFormWidth Then .Width = minFormWidth
        End With

        With StsBar
            .Left = 0
            .Width = MainFrm.ScaleWidth
            .Top = MainFrm.ScaleHeight - .Height
        End With

        With LeftFrame
            .Left = 0
            .Top = 0
            .Height = StsBar.Top
            .Width = picSplitter.Left ' - 60 '+ imgSplitter.Width
        End With

        With imgSplitter
            .Top = 0
            .Height = LeftFrame.Height
            .Left = LeftFrame.Width - .Width
        End With
        
        With theShow

        If mnuView_Left.Checked Then
            .Left = LeftFrame.Left + LeftFrame.Width '- 30
            .Top = 0
            tempint = MainFrm.ScaleWidth - .Left

            If tempint < 0 Then
                theShow.Visible = False
            Else
                theShow.Visible = True
                .Width = tempint
            End If

            .Height = LeftFrame.Height
        Else
            .Left = 0
            .Top = 0
            .Width = MainFrm.ScaleWidth
            .Height = LeftFrame.Height
        End If

    End With

    With LeftStrip
        .Left = 30
        .Top = 60
        tempint = LeftFrame.Height - 120
        .Height = Abs(tempint)
        tempint = LeftFrame.Width - 120
        .Width = Abs(tempint)
    End With

    With ListFrame
        .Left = LeftStrip.Left + 30
        .Top = LeftStrip.Top + 360
        tempint = LeftStrip.Width - 60

        If tempint < 0 Then
            ListFrame.Visible = False
        Else
            ListFrame.Visible = True
            .Width = tempint
        End If

        tempint = LeftStrip.Height - 420

        If tempint < 0 Then
            ListFrame.Visible = False
        Else
            ListFrame.Visible = True
            .Height = tempint
        End If

    End With


        With List
            .Top = 0
            .Left = 0
            .Width = ListFrame.Width
            .Height = ListFrame.Height
        End With





    'For i = 1 To IEView.Count

    

    
    
    
    With HBoxScroll
        .Left = 0
        .Top = theShow.Height - .Height
        .Width = theShow.Width
    End With
    
    With VBOXScroll
        .Top = 0
        .Left = theShow.Width - .Width
        .Height = theShow.Height - .Width
    End With
    
    
    With ImageBox '(i - 1)
    
        If .Width - 100 < theShow.Width Then
            .Left = (theShow.Width - .Width) / 2
            HBoxScroll.Visible = False
            VBOXScroll.Height = VBOXScroll.Height + VBOXScroll.Width
        Else
            HBoxScroll.Visible = True
            .Left = 0
            HBoxScroll.Min = 0
            HBoxScroll.Max = ImageBox.Width - theShow.Width
        End If
        
        If .Height - 100 < theShow.Height Then
            .Top = (theShow.Height - .Height) / 2
            VBOXScroll.Visible = False
            'HBoxScroll.Width = HBoxScroll.Height + HBoxScroll.Width
        Else
            .Top = 0
            VBOXScroll.Visible = True
            VBOXScroll.Min = 0
            VBOXScroll.Max = ImageBox.Height - theShow.Height

        End If
        
    End With
    
    

    'Next
    MainFrm.Refresh

End Sub

Private Sub Form_Unload(Cancel As Integer)

    
    
     '-----------------------------------------------------------------------------------
    '-------------------------- Call SaveSetting Start-----------------------------
        Dim hSetting As New CSetting
        hSetting.iniFile = zhtmIni
        hSetting.Save Me, SF_FORM                                               'Postion
        'hsetting.Save Me, SF_FONT                                               'Font For ViewStyle
        'hsetting.Save Me, SF_COLOR                                             'Color For ViewStyle
        hSetting.Save mnuFile_Recent, SF_MENUARRAY                'Recent File
        hSetting.Save mnuBookmark, SF_MENUARRAY                   'BookMark
        hSetting.Save mnuView_Left, SF_CHECKED                        'ShowLeft Check
        hSetting.Save mnuView_Menu, SF_CHECKED                       'ShowMenu Check
        hSetting.Save mnuView_StatusBar, SF_CHECKED                'ShowstatusBar Check
        hSetting.Save mnuView_FullScreen, SF_CHECKED               'FullScreen Check
        hSetting.Save mnuView_TopMost, SF_CHECKED                  'OnTopMost Check
        hSetting.Save Timer, SF_Tag                                             'Time InterVal
        hSetting.Save mnuFile_Open, SF_Tag                                 'InitDir for Dialog
        hSetting.Save Me, SF_Tag                                                  'UseTemplate ?
        hSetting.Save ImageBox, SF_Tag                                            'TemplateHtml
        hSetting.Save Me.picSplitter, SF_Tag 'LeftWidth
        hSetting.Save Me.theShow, SF_COLOR
        hSetting.Save mnu_mode_fit_width, SF_CHECKED
        hSetting.Save mnu_mode_fit_height, SF_CHECKED
        hSetting.Save mnu_mode_fit_viewer, SF_CHECKED
        hSetting.Save mnu_mode_zoom_in, SF_CHECKED
        hSetting.Save mnu_mode_zoom_out, SF_CHECKED
        hSetting.Save mnu_mode_keep_original, SF_CHECKED
        Set hSetting = Nothing
    '-------------------------- Call SaveSetting END----------------------------
    '----------------------------------------------------------------------------------
    
    Dim fso As New scripting.FileSystemObject

    
    Call saveReadingStatus
    On Error Resume Next
    If fso.FolderExists(sTempZH) Then fso.DeleteFolder (sTempZH), True
'    If App.PrevInstance = False Then
'        If fso.FolderExists(Tempdir) Then fso.DeleteFolder Tempdir, True
'    End If
    
    Set fso = Nothing
    Set lUnzip = Nothing
    Set lZip = Nothing
    Set sFilesInZip = Nothing
    Unload frmBookmark
    Unload frmOptions
    'Call endUP

End Sub

Public Sub GetView(ByVal shortfile As String)

    Dim tempfile As String
    tempfile = BuildPath(sTempZH, shortfile)
    
    If linvblib.PathExists(tempfile) = False Then
        MainFrm.myXUnzip zhrStatus.sCur_zhFile, shortfile, sTempZH, zhrStatus.sPWD
    End If
    
    If linvblib.PathExists(tempfile) = False Then Exit Sub
    
    DisplayProgress 1, 3
    
    
    VBOXScroll.value = 0
    HBoxScroll.value = 0
    StsBar.Panels(3).text = Me.curFileIndex(shortfile) & "\" & sFilesInZip.Count
    
    
    Dim oH As Single
    Dim oW As Single
    
    Dim imgToLoad As Picture
    Dim imgHeight As Long
    Dim imgWidth As Long
    
    
    On Error GoTo ErrLoadingPicture
    
    Set imgToLoad = LoadPicture(tempfile)
    
    DisplayProgress 2, 3
    
    imgHeight = MainFrm.ScaleX(imgToLoad.Height, 8, 3) * Screen.TwipsPerPixelY
    imgWidth = MainFrm.ScaleX(imgToLoad.Width, 8, 3) * Screen.TwipsPerPixelX
    
    NotResize = True
    
    Dim vm  As ViewMode
    Dim R As Double
    Dim W As Long
    Dim H As Long
    
    W = theShow.Width
    H = theShow.Height
    
    vm = checkMode()
    R = 1
    
    Select Case vm
    
        Case KEEP_ORIGIN
            R = 1
        Case FIT_WIDTH
            If imgWidth > W Then R = W / imgWidth
        Case FIT_HEIGHT
            If imgHeight > H Then R = H / imgHeight
        Case FIT_VIEWER
            Dim R1 As Single
            Dim R2 As Single
            R1 = 1
            R2 = 1
            If imgWidth > W Then R1 = W / imgWidth
            If imgHeight > H Then R2 = H / imgHeight
            If R1 > R2 Then R = R2 Else R = R1
        Case ZOOM_IN
            R = 1 + ZOOM_FACTOR
        Case ZOOM_OUT
            R = 1 + ZOOM_FACTOR
    End Select
        
        ImageBox.Width = imgWidth * R
        ImageBox.Height = imgHeight * R
    
    NotResize = False
    

    Set ImageBox.Picture = imgToLoad
    
    DisplayProgress 3, 3
    Form_Resize
    

'    screenHeight = (IEView.Height - 360) \ Screen.TwipsPerPixelY
'    screenWidth = (IEView.Width - 360) \ Screen.TwipsPerPixelX
'    Set imgToLoad = Nothing
'    resizeRate = 1
'    resizeRateY = 1
'    resizeRateX = 1
'    If imgHeight > screenHeight Then resizeRateY = screenHeight / imgHeight
'    If imgWidth > screenWidth Then resizeRateX = screenWidth / imgWidth
'    resizeRate = resizeRateY
'    If resizeRateY > resizeRateX Then resizeRate = resizeRateX
'    If resizeRate < 1 Then
'    imgHeight = Int(imgHeight * resizeRate)
'    imgWidth = Int(imgWidth * resizeRate)
'    Else
'    imgHeight = 0
'    imgWidth = 0
'    End If


zhrStatus.sCur_zhSubFile = shortfile

If bAutoShowNow Or bRandomShow Then Timer.Enabled = True

ErrLoadingPicture:

DisplayProgress 0, 0
StsBar.Panels(1).text = ""



End Sub

Public Sub getZIPContent(ByVal sZipfilename As String)
    
    Dim lfor As Long
    Dim lEnd As Long
    Dim sFilename As String
    Dim sExt As String
    Dim zipFileList As New CZipItems
    
    lUnzip.ZipFile = sZipfilename
    lUnzip.getZipItems zipFileList
    
    
    Set sFilesInZip = New CStringVentor
    sFilesInZip.initSize = 500
    
    
    lEnd = zipFileList.Count
    StsBar.Panels("ie").text = "正在扫描 " & lUnzip.ZipFile & "..."
    
    For lfor = 1 To lEnd
        'StsBar.Panels(1).text = "已扫描到" & lfor & "个文件..."
        DisplayProgress lfor, lEnd
        sFilename = zipFileList(lfor).filename
        sExt = LCase$(Right$(sFilename, 3))
        If sExt = "jpg" Or _
           sExt = "gif" Or _
           sExt = "png" Or _
           sExt = "bmp" Or _
           sExt = "jpeg" Or _
           sExt = "jpe" Then
                sFilesInZip.assign sFilename
        End If
    Next
    
    DisplayProgress 0, 0
    
    StsBar.Panels(1).text = "扫描完毕"
      
    MainFrm.Enabled = True

    Set zipFileList = Nothing


End Sub






Public Sub DisplayProgress(ByRef iCur As Long, ByRef iMax As Long)
'▓
'█
'Const ps = "▓"
Const ps = "█"
Dim oldText As String
Dim newText As String
oldText = StsBar.Panels(2).text
If iMax <= 0 Then
    StsBar.Panels(2).text = ""
Else
    newText = String$(Int(iCur / iMax * 9) + 1, ps)
    StsBar.Panels(2).text = newText
End If

End Sub
'Private Sub IEView_ProgressChange(Index As Integer, ByVal Progress As Long, ByVal ProgressMax As Long)
' Call
'End Sub




Private Sub HBoxScroll_Scroll()
ImageBox.Move -HBoxScroll.value, ImageBox.Top
End Sub




Private Sub ImageBox_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        mnuGo_Next_Click
    ElseIf Button = 2 Then
        mnuGo_Previous_Click
    End If
    
End Sub

Private Sub imgSplitter_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    With imgSplitter
        picSplitter.Move .Left, .Top, .Width \ 2, .Height - 20
    End With

    imgSplitter.Tag = "Moving"
    picSplitter.Visible = True

End Sub

Private Sub imgSplitter_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    Dim sglPos As Single

    If imgSplitter.Tag = "Moving" Then
        sglPos = x + imgSplitter.Left

        If sglPos < sglSplitLimit Then
            picSplitter.Left = sglSplitLimit
        ElseIf sglPos > Me.Width - sglSplitLimit Then
            picSplitter.Left = Me.Width - sglSplitLimit
        Else
            picSplitter.Left = sglPos
        End If

    End If

End Sub

Private Sub imgSplitter_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    picSplitter.Visible = False
    imgSplitter.Tag = ""
    LeftFrame.Width = picSplitter.Left
    picSplitter.Tag = picSplitter.Left
    Form_Resize

End Sub

'Private Sub IEView_NewWindow2(Index As Integer, ppDisp As Object, Cancel As Boolean)
'Cancel = True
'End Sub
Private Sub LeftStrip_Click()

    Dim indexList As Integer
    
    If LeftStrip.Enabled = False Then Exit Sub
    On Error Resume Next
    
    If isListLoaded(List) <> lstLoaded Then
        readyToLoadList
    End If
    
    Call selectListItem
    
End Sub

Private Sub List_Collapse(ByVal Node As MSComctlLib.Node)

    Node.Image = 1

End Sub

Private Sub List_Expand(ByVal Node As MSComctlLib.Node)

    Node.Image = 2

End Sub

Private Sub list_NodeClick(ByVal Node As MSComctlLib.Node)

    Dim sTag As String
    sTag = Node.Tag
    If Right$(sTag, 1) <> "/" Then GetView sTag

End Sub

'Public lcurFileIndex As Long
Sub Loadlist(thelist As TreeView, LContent() As String, lCount As Long)

    thelist.Visible = False
    thelist.Nodes.Clear
    
    Dim Krelative As String
    Dim Krelationship As TreeRelationshipConstants
    Dim kKey As String
    Dim Ktext As String
    Dim Kimageindex As Integer
    Dim Ktag As String
    Dim thename As String
    Dim i As Integer, pos As Integer
    
  On Error GoTo CatalogError

    For i = 1 To lCount
        
        StsBar.SimpleText = "载入列表: " & (i + 1) & "/" & lCount & " (" & Format$((i + 1) / lCount, "00%") & ")"
        thename = LContent(1, i)
        Ktag = LContent(2, i)
        kKey = "ZTM" + LContent(1, i)

        If Right$(thename, 1) = "/" Then thename = Left$(thename, Len(thename) - 1)
        pos = InStrRev(thename, "/")

        If pos = 0 Then
            Ktext = thename
            If Right$(LContent(1, i), 2) = ":/" Then
                Kimageindex = 4
            ElseIf Right$(LContent(1, i), 1) = "/" Then
                Kimageindex = 1
            Else
                Kimageindex = 3
            End If
            thelist.Nodes.Add(, , kKey, Ktext, Kimageindex).Tag = Ktag
        Else
            Krelative = "ZTM" + Left$(LContent(1, i), pos)
            Krelationship = tvwChild
            Ktext = Right$(thename, Len(thename) - pos)
            If Right$(LContent(1, i), 1) = "/" Then
                Kimageindex = 1
            Else
                Kimageindex = 3
            End If
            On Error Resume Next
            If nodeExist(thelist.Nodes, Krelative) = False Then
            xAddNode thelist.Nodes, Left$(LContent(1, i), pos), "ZTM", "", 1
            End If
            thelist.Nodes.Add(Krelative, Krelationship, kKey, Ktext, Kimageindex).Tag = Ktag
            
            
        End If

CatalogError:
    Next

    '---------------------------------------------
    setListStatus thelist, lstLoaded
    thelist.Visible = True

End Sub

Private Sub saveReadingStatus()
    If zhrStatus.sCur_zhFile <> "" And zhrStatus.sCur_zhSubFile <> "" Then
        Dim nowAt As ReadingStatus
        With nowAt
            .page = zhrStatus.sCur_zhSubFile
            '.perOfScrollTop = IEView.Document.body.scrollTop / IEView.Document.body.scrollHeight
            '.perOfScrollLeft = IEView.Document.body.scrollLeft / IEView.Document.body.scrollWidth
        End With
        rememberBook linvblib.BuildPath(sConfigDir, zhMemFile), zhrStatus.sCur_zhFile, nowAt
    End If
End Sub

Public Sub loadzh(ByVal thisfile As String, Optional ByVal firstfile As String = "", Optional Reloadit As Boolean = False)
    
       
    If linvblib.PathExists(thisfile) = False Then Exit Sub
    If linvblib.PathType(thisfile) <> LNFile Then Exit Sub
    
    thisfile = toUnixPath(thisfile)
    firstfile = toUnixPath(firstfile)

    With zhrStatus

        If Reloadit = False And thisfile = .sCur_zhFile Then

            If firstfile <> "" Then
                GetView firstfile
                Exit Sub
            ElseIf .sCur_zhSubFile <> "" Then
                GetView .sCur_zhSubFile
                Exit Sub
            End If

        End If

    End With


    Call saveReadingStatus
    Call zhReaderReset
    
    getZIPContent thisfile
    
    If sFilesInZip.Count <= 0 Then Exit Sub
    
    
    Me.ChangeMainFrmCaption thisfile
    
    If thisfile <> zhrStatus.sCur_zhFile Then
        sTempZH = linvblib.GetBaseName(linvblib.GetTempFileName)
        sTempZH = linvblib.BuildPath(Tempdir, sTempZH)

        Do Until linvblib.PathExists(sTempZH) = False
            sTempZH = linvblib.GetBaseName(linvblib.GetTempFileName)
            sTempZH = linvblib.BuildPath(Tempdir, sTempZH)
        Loop

        MkDir sTempZH
        sTempZH = toUnixPath(sTempZH)
        zhrStatus.sCur_zhFile = thisfile
    End If


    
    Dim hMRU As New CMenuArrHandle
    hMRU.Menus = mnuFile_Recent
    hMRU.maxItem = Val(mnuFile_Recent(0).Tag)
    hMRU.AddUnique toDosPath(thisfile) ', thisfile
    Set hMRU = Nothing
    
    NotResize = True

'    With zhInfo
'
'        If .zvShowLeft = zhtmVisiableTrue Then
'            ShowMenu True
'        ElseIf .zvShowLeft = zhtmVisiableFalse Then
'            ShowLeft False
'        End If
'
'        If .zvShowMenu = zhtmVisiableTrue Then
'            ShowMenu True
'        ElseIf .zvShowMenu = zhtmVisiableFalse Then
'            ShowMenu False
'        End If
'
'        If .zvShowStatusBar = zhtmVisiableTrue Then
'            ShowStatusBar True
'        ElseIf .zvShowStatusBar = zhtmVisiableFalse Then
'            ShowStatusBar False
'        End If
'
'    End With

    If mnuView_Left.Checked Then
        LeftStrip.Enabled = False
            Dim i As Integer
            For i = 1 To LeftStrip.Tabs.Count
                LeftStrip.Tabs(i).Selected = False
            Next
        LeftStrip.Enabled = True
        LeftStrip.Tabs(1).Selected = True
    End If

    NotResize = False
    'StsBar.Panels("reading").text = linvblib.GetBaseName(thisfile)
    Form_Resize
    
    If firstfile <> "" Then
        GetView firstfile
    Else
        Dim nowAt As ReadingStatus
        nowAt = searchMem(linvblib.BuildPath(sConfigDir, zhMemFile), zhrStatus.sCur_zhFile)
        If nowAt.page <> "" Then
            scrollX = nowAt.perOfScrollLeft
            scrollY = nowAt.perOfScrollTop
            bScroll = True
            GetView nowAt.page
        Else
            GetView sFilesInZip.value(1)
        End If
    End If

End Sub



Private Sub readyToLoadList()

    Dim zipContent() As String
    Dim lzipCount As Long
    Dim i As Long

    'lzipCount = lFoldersIZcount + sfilesinzip.count
    
'    If listIndex = lwContent Then
'        lzipCount = sFilesINContent(1).Count
'    Else

    lzipCount = sFilesInZip.Count


    If lzipCount > lcstFittedListItemsNum Then
        Dim YESORNO As VbMsgBoxResult
        YESORNO = MsgBox("文件列表有" & lzipCount & "项之多,继续吗?", vbYesNo, "打开文件列表")
        If YESORNO = vbNo Then Exit Sub
    End If

    ReDim zipContent(1 To 2, 1 To lzipCount)
'    lEnd = lFoldersIZcount - 1
'
'    For i = 0 To lEnd
'        zipContent(0, i) = sFoldersInZip.value(i + 1)
'        zipContent(1, i) = zipContent(0, i)
'    Next


        For i = 1 To lzipCount
            zipContent(1, i) = sFilesInZip.value(i)
            zipContent(2, i) = zipContent(1, i)
        Next
'    Else
'        For i = 1 To lzipCount
'            zipContent(1, i) = sFilesINContent(1).Value(i)
'            zipContent(2, i) = sFilesINContent(2).Value(i)
'        Next

    'Set sFoldersInZip = Nothing
    '    zhrStatus.iListIndex = lwFiles
    '    trvwlist.ZOrder 0
    Loadlist List, zipContent(), lzipCount

End Sub




Private Sub lUnzip_PasswordRequest(sPassword As String, ByVal sName As String, bCancel As Boolean)

    bCancel = False
    Static lastName As String

    If bInValidPassword = False And zhrStatus.sPWD <> "" Then
        sPassword = zhrStatus.sPWD

        If sName = lastName Then
            bInValidPassword = True
        Else
            lastName = sName
        End If

    Else
        sPassword = InputBox(lUnzip.ZipFile & vbCrLf & sName & " Request For Password", "Password", "")

        If sPassword <> "" Then
            bInValidPassword = False
            zhrStatus.sPWD = sPassword
        Else
            bCancel = True
        End If

    End If

End Sub

Private Sub lZip_PasswordRequest(sPassword As String, bCancel As Boolean)

    sPassword = InputBox("Type the password of " + vbCrLf + lUnzip.ZipFile + ":", "Invaild Password")

    If sPassword <> "" Then
        bCancel = False
        zhrStatus.sPWD = sPassword
    Else
        bCancel = True
    End If

End Sub


'Private Sub listFav_DblClick()
'If IsNull(listFav.SelectedItem) Then Exit Sub
'Dim tempstr As String
'Dim pos As Integer
'tempstr = favlist.locate(listFav.SelectedItem.Index)
'pos = InStr(tempstr, "|")
'If pos = 0 Then Exit Sub
'loadztm left$(tempstr, pos - 1), right$(tempstr, Len(tempstr) - pos)
'End Sub
'
'Private Sub listFav_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'If Button = 1 Then Exit Sub
'MainFrm.PopupMenu mnuFav, , x + ListFrame.Left, y + ListFrame.Top
'End Sub
Public Sub mnu_Click(Index As Integer)
    
    If Index = 7 Then
        mnuGo_Previous_Click
        Exit Sub
    ElseIf Index = 8 Then
        mnuGo_Next_Click
        Exit Sub
    End If

    If zhrStatus.sCur_zhFile = "" Then
        mnuFile_Close.Enabled = False
        mnuHelp_BookInfo.Enabled = False
    Else
        mnuFile_Close.Enabled = True
        mnuHelp_BookInfo.Enabled = True
    End If


End Sub

Private Function checkMode() As ViewMode

    checkMode = KEEP_ORIGIN
    
    If mnu_mode_fit_width.Checked Then
        checkMode = FIT_WIDTH
    ElseIf mnu_mode_fit_height.Checked Then
        checkMode = FIT_HEIGHT
    ElseIf mnu_mode_fit_viewer.Checked Then
        checkMode = FIT_VIEWER
    ElseIf mnu_mode_zoom_in.Checked Then
        checkMode = ZOOM_IN
    ElseIf mnu_mode_zoom_out.Checked Then
        checkMode = ZOOM_OUT
    ElseIf mnu_mode_fit_height.Checked Then
        checkMode = KEEP_ORIGIN
    End If

    If checkMode = KEEP_ORIGIN Then mnu_mode_keep_original.Checked = True
    
End Function

Private Sub disable_all_mode()
    mnu_mode_fit_width.Checked = False
    mnu_mode_fit_height.Checked = False
    mnu_mode_fit_viewer.Checked = False
    mnu_mode_zoom_in.Checked = False
    mnu_mode_zoom_out.Checked = False
    mnu_mode_keep_original.Checked = False
End Sub

Private Sub mnu_mode_fit_height_Click()
    Call disable_all_mode
    mnu_mode_fit_height.Checked = True
    GetView zhrStatus.sCur_zhSubFile
End Sub

Private Sub mnu_mode_fit_viewer_Click()
    Call disable_all_mode
    mnu_mode_fit_viewer.Checked = True
     GetView zhrStatus.sCur_zhSubFile
End Sub

Private Sub mnu_mode_fit_width_Click()

    Call disable_all_mode
    mnu_mode_fit_width.Checked = True
     GetView zhrStatus.sCur_zhSubFile
End Sub

Private Sub mnu_mode_keep_original_Click()
    Call disable_all_mode
    mnu_mode_keep_original.Checked = True
     GetView zhrStatus.sCur_zhSubFile
End Sub

Private Sub mnu_mode_zoom_in_Click()
    Call disable_all_mode
    mnu_mode_zoom_in.Checked = True
    ZOOM_FACTOR = ZOOM_FACTOR + 0.08
     GetView zhrStatus.sCur_zhSubFile
End Sub

Private Sub mnu_mode_zoom_out_Click()
    Call disable_all_mode
    mnu_mode_zoom_out.Checked = True
    ZOOM_FACTOR = ZOOM_FACTOR - 0.08
     GetView zhrStatus.sCur_zhSubFile
End Sub

Public Sub mnuBookmark_add_Click()

    'Dim i As Integer
    On Error Resume Next
    Dim sCaption As String
    If zhrStatus.sCur_zhFile = "" Then Exit Sub
    sCaption = linvblib.GetBaseName(zhrStatus.sCur_zhFile)
    
    Dim hMNU As New CMenuArrHandle
    With hMNU
     .Menus = mnuBookmark
    .maxItem = 100
    .maxCaptionLength = 100
    .JustAdd sCaption, zhrStatus.sCur_zhFile & "|" & zhrStatus.sCur_zhSubFile
    End With
    Set hMNU = Nothing
    
End Sub

Public Sub mnuBookmark_Click(Index As Integer)

    Dim sBMZhfile As String
    Dim sBMZhsubfile As String
    Dim i As Integer
    Dim pos As Integer
    i = Index
    pos = InStr(mnuBookmark(i).Tag, "|")

    If pos > 0 Then
        sBMZhfile = Left$(mnuBookmark(i).Tag, pos - 1)
        sBMZhsubfile = Right$(mnuBookmark(i).Tag, Len(mnuBookmark(i).Tag) - pos)
        loadzh sBMZhfile, sBMZhsubfile
    End If

End Sub

Public Sub mnuBookmark_manage_Click()

    If mnuBookmark.Count = 1 Then Exit Sub
    frmBookmark.Show 1
    Unload frmBookmark
    'saveMNUBookmark

End Sub

Private Sub mnuDir_delete_Click()
    If zhrStatus.sCur_zhFile <> "" Then
        Dim sBackUp As String
        sBackUp = zhrStatus.sCur_zhFile
        Dim msgConfirm As VbMsgBoxResult
        msgConfirm = MsgBox("Delete" & " " & sBackUp, vbOKCancel, "Make Sure ?")
        If msgConfirm = vbOK Then
            mnuDir_readNext_Click
            Kill sBackUp
            If zhrStatus.sCur_zhFile = sBackUp Then
                mnufile_Close_Click
            End If
        End If
    End If
End Sub



Private Sub mnuDir_random_Click()
    Dim sCur As String
    Dim sPath As String
    sPath = zhrStatus.sCur_zhFile
    If sPath = "" Then sPath = CurDir$
    sCur = linvblib.gCFileSystem.LookFor(sPath, LN_FILE_RAND, "*." & TakeCare_EXT)
    If sCur <> "" Then MainFrm.loadzh sCur
End Sub

Private Sub mnuDir_readNext_Click()
    Dim sCur As String
    Dim sPath As String
    sPath = zhrStatus.sCur_zhFile
    If sPath = "" Then sPath = CurDir$
    sCur = linvblib.gCFileSystem.LookFor(sPath, LN_FILE_next, "*." & TakeCare_EXT)
    If sCur <> "" Then MainFrm.loadzh sCur
End Sub

Private Sub mnuDir_readPrev_Click()
    
    Dim sCur As String
    Dim sPath As String
    sPath = zhrStatus.sCur_zhFile
    If sPath = "" Then sPath = CurDir$
    sCur = linvblib.gCFileSystem.LookFor(sPath, LN_FILE_prev, "*." & TakeCare_EXT)
    If sCur <> "" Then MainFrm.loadzh sCur
        
    
End Sub





Public Sub mnufile_Close_Click()

    Dim fso As New FileSystemObject
    'Dim i As Integer
    'rememberNew zhrStatus.sCur_zhFile, zhrStatus.sCur_zhSubFile
    zhReaderReset
    On Error Resume Next

    If fso.FolderExists(sTempZH) Then fso.DeleteFolder sTempZH, False
    MainFrm.appAbout
End Sub

Public Sub mnufile_exit_Click()

    Unload Me

End Sub

Public Sub mnufile_Open_Click()

    Dim thisfile As String
    Dim sInitDir As String
    Dim fso As New gCFileSystem

    'rememberNew zhrStatus.sCur_zhFile, zhrStatus.sCur_zhSubFile

    If zhrStatus.sCur_zhFile <> "" Then
        If fso.PathExists(zhrStatus.sCur_zhFile) = True Then sInitDir = fso.GetParentFolderName(zhrStatus.sCur_zhFile)
    Else
        sInitDir = mnuFile_Open.Tag
    End If
    mnuFile_Open.Tag = sInitDir
    
    Dim fResult As Boolean
    Dim cDLG As New CCommonDialogLite

    If fso.PathExists(sInitDir) Then sInitDir = linvblib.toDosPath(sInitDir)
    fResult = cDLG.VBGetOpenFileName( _
       filename:=thisfile, _
       Filter:="Zippacked of pictures|*." & TakeCare_EXT & ";*.zip|所有文件|*.*", _
       InitDir:=sInitDir, _
       DlgTitle:=Me.StrLocalize(mnuFile_Open.Caption), _
       Owner:=Me.hwnd)
    Set cDLG = Nothing

    If fResult Then
        If thisfile = "" Then Exit Sub
        loadzh thisfile, "", False
    End If

End Sub

Public Sub mnuFile_PReFerence_Click()

    Load frmOptions
    frmOptions.Show

End Sub

Public Sub mnuFile_Recent_Click(Index As Integer)

    'rememberNew zhrStatus.sCur_zhFile, zhrStatus.sCur_zhSubFile
    Dim fname As String
    fname = mnuFile_Recent(Index).Tag
    If linvblib.FileExists(fname) = False Then
        Dim HM As New CMenuArrHandle
        HM.Menus = mnuFile_Recent
        HM.Remove Index
        Set HM = Nothing
        MsgBox "文件不存在:" & vbCrLf & fname, vbInformation, "错误"
    Else
        loadzh mnuFile_Recent(Index).Tag
    End If
End Sub

Private Sub mnuGo_AutoNext_Click()

    If sFilesInZip.Count = 0 Then Exit Sub

    If mnuGo_AutoNext.Checked = False Then
        bAutoShowNow = True
        bRandomShow = False
        mnuGo_AutoNext.Checked = True
        mnuGo_AutoRandom.Checked = False
        Timer_Timer
        Timer.Enabled = True
    Else
        bAutoShowNow = False
        mnuGo_AutoNext.Checked = False
        Timer.Enabled = False
    End If

End Sub

Private Sub mnuGo_AutoRandom_Click()

    If sFilesInZip.Count = 0 Then Exit Sub

    If mnuGo_AutoRandom.Checked = False Then
        bAutoShowNow = True
        bRandomShow = True
        mnuGo_AutoRandom.Checked = True
        mnuGo_AutoNext.Checked = False
        Timer_Timer
        Timer.Enabled = True
    Else
        bAutoShowNow = False
        mnuGo_AutoRandom.Checked = False
        Timer.Enabled = False
    End If

End Sub

Public Function GoBack()

End Function
Public Function GoForward()

End Function

Public Sub mnuGo_Back_Click()

    On Error Resume Next
    Timer.Enabled = False
    
    Call GoBack

    If bAutoShowNow Then Timer.Enabled = True

End Sub

Public Sub mnuGo_Forward_Click()

    On Error Resume Next
    Timer.Enabled = False
    Call GoForward

    If bAutoShowNow Then Timer.Enabled = True

End Sub

Public Sub mnuGo_Home_Click()

    On Error GoTo 0
    
    If zhrStatus.sCur_zhFile = "" Or sFilesInZip.Count < 1 Then
        appAbout
    Else
        GetView sFilesInZip(1)
    End If

    '
    '        Dim sTmpFile As String
    '        sTmpFile = BuildPath(sTempZH, "index." & cTxtIndex)
    '        If IndexFromFileList(sFilesInZip(), sTmpFile) Then
    '            zhInfo.sDefaultfile = "index." & cTxtIndex
    '        End If
    '        GetView zhInfo.sDefaultfile
    '    End If

End Sub

Public Sub mnuGo_Next_Click()

    On Error GoTo Herr
    Timer.Enabled = False
    Dim lcurPage As Long
    

        
        If sFilesInZip.Count < 1 Then GoTo Herr

        If zhrStatus.sCur_zhSubFile = "" Then
            lcurPage = 0
        Else
            lcurPage = curFileIndex(zhrStatus.sCur_zhSubFile)
        End If

        If lcurPage >= sFilesInZip.Count Then lcurPage = 0
        GetView sFilesInZip.value(lcurPage + 1)


Herr:

    If bAutoShowNow Then Timer.Enabled = True

End Sub

Public Sub mnuGo_Previous_Click()

    On Error GoTo Herr
    Timer.Enabled = False
    Dim lcurPage As Long


        If sFilesInZip.Count < 1 Then GoTo Herr

        If zhrStatus.sCur_zhSubFile = "" Then
            lcurPage = 2
        Else
            lcurPage = curFileIndex(zhrStatus.sCur_zhSubFile)
        End If

        If lcurPage <= 1 Then lcurPage = sFilesInZip.Count + 1
        GetView sFilesInZip.value(lcurPage - 1)


Herr:

    If bAutoShowNow Then Timer.Enabled = True

End Sub

Private Sub mnuGo_Random_Click()

    Timer.Enabled = False
    randomView

    If bAutoShowNow Then Timer.Enabled = True

End Sub

Public Sub mnuhelp_About_Click()

    Dim sAbout As String
    sAbout = sAbout + Space$(4) + App.ProductName + " (Build" + Str$(App.Major) + "." + Str$(App.Minor) + "." + Str$(App.Revision) + ")" + vbCrLf
    sAbout = sAbout + Space$(4) + App.LegalCopyright
    MsgBox sAbout, vbInformation, "About"

End Sub

Public Sub mnuHelp_BookInfo_Click()

'    Dim sAbout As String
'
'    If zhrStatus.sCur_zhFile <> "" Then
'        sAbout = sAbout + Space$(4) + "Title:" + zhInfo.sTitle + vbCrLf
'        sAbout = sAbout + Space$(4) + "Author:" + zhInfo.sAuthor + vbCrLf
'        sAbout = sAbout + Space$(4) + "Catalog:" + zhInfo.sCatalog + vbCrLf
'        sAbout = sAbout + Space$(4) + "Publisher:" + zhInfo.sPublisher + vbCrLf
'        sAbout = sAbout + Space$(4) + "Date:" + zhInfo.sDate + vbCrLf
'        MsgBox sAbout, vbInformation, "BookInfo of [" & zhrStatus.sCur_zhFile & "]"
'    End If

End Sub






Public Sub mnuView_Left_Click()

    If mnuView_Left.Checked Then
        ShowLeft False
    Else
        ShowLeft True
    End If

End Sub

Public Sub mnuView_Menu_Click()

    If mnuView_Menu.Checked Then
        ShowMenu False
    Else
        ShowMenu True
    End If

End Sub

'Public Sub mnuIe_AddBookmark_Click()
'mnuBookmark_add_Click
'End Sub
'
'Public Sub mnuIe_Backward_Click()
'On Error Resume Next
'IEView.GoBack
'End Sub
'
'Public Sub mnuIe_copy_Click()
'On Error Resume Next
'IEView.ExecWB OLECMDID_COPY, OLECMDEXECOPT_DODEFAULT
'End Sub
'
'Public Sub mnuIe_Forward_Click()
'On Error Resume Next
'IEView.GoForward
'End Sub
'
'Public Sub mnuIe_Print_Click()
'On Error Resume Next
'IEView.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_DODEFAULT
'End Sub
'
'Public Sub mnuIe_property_Click()
'Dim sAbout As String
'Dim indentString  As String
'indentString = String$(10, Chr(32))
'sAbout = "当前文件:" & vbCrLf & indentString
'If zhrStatus.sCur_zhFile <> "" Then
'    sAbout = sAbout & zhrStatus.sCur_zhFile & "|" & zhrStatus.sCur_zhSubFile & vbCrLf & vbCrLf
'Else
'    sAbout = sAbout & vbCrLf & vbCrLf
'End If
'sAbout = sAbout & "书名:" & vbCrLf & indentString & zhInfo.sTitle & vbCrLf & vbCrLf
'sAbout = sAbout & "作者:" & vbCrLf & indentString & zhInfo.sAuthor & vbCrLf & vbCrLf
'sAbout = sAbout & "分类:" & vbCrLf & indentString & zhInfo.sCatalog & vbCrLf & vbCrLf
'sAbout = sAbout & "出版:" & vbCrLf & indentString & zhInfo.sPublisher & vbCrLf & vbCrLf
'sAbout = sAbout & "日期:" & vbCrLf & indentString & zhInfo.sDate
'Load dlgProperty
'dlgProperty.lblProperty.Caption = sAbout
'dlgProperty.Show 1
'End Sub
'
'Public Sub mnuIe_refresh_Click()
'IEView.Refresh2
'End Sub
'
'Public Sub mnuIe_SelectAll_Click()
'On Error Resume Next
'IEView.ExecWB OLECMDID_SELECTALL, OLECMDEXECOPT_DODEFAULT
'End Sub
'
'Public Sub mnuIe_ViewSource_Click()
'Dim clsRegOpen As New clsRegView
'Dim Arrsettings() As Variant
'Dim sViewer As String
'Dim sTmpFile As String
'Dim iNum As Long
'If zhrStatus.sCur_zhSubFile = "" Then Exit Sub
'sTmpFile = ModFile.modFile_buildpath(sTempZH, ModFile.ExtractFileName(zhrStatus.sCur_zhSubFile))
'
'With clsRegOpen
'    .m_root = HKEY_CURRENT_USER
'    .m_key = "Software\Microsoft\Internet Explorer\Default HTML Editor\shell\edit\command"
'    If .GetAllSettings(Arrsettings) = ERROR_SUCCESS Then
'        sViewer = Arrsettings(0, 1)
'        sViewer = expandStr(sViewer)
'    End If
'End With
'If sViewer <> "" Then
'    iNum = FreeFile
'    Open sTmpFile For Binary Access Write As iNum
'    Put #iNum, , IEView.Document.documentElement.outerHTML
'    Close iNum
'    sViewer = Replace$(sViewer, "%1", sTmpFile, , , vbTextCompare)
'    sViewer = Replace$(sViewer, "%l", sTmpFile, , , vbTextCompare)
'    Shell sViewer, vbNormalFocus
'End If
'End Sub
Public Sub mnuView_fullscreen_Click()

    Static oPos As MYPoS
    NotResize = True

    If mnuView_FullScreen.Checked Then
        mnuView_FullScreen.Checked = False
        MWindows.setBorderStyle Me, bsResizable
        Me.WindowState = vbNormal
        Me.Move oPos.Left, oPos.Top, oPos.Width, oPos.Height
    Else
        mnuView_FullScreen.Checked = True

        With oPos
            .Left = Me.Left
            .Top = Me.Top
            .Height = Me.Height
            .Width = Me.Width
        End With

        Me.WindowState = vbMaximized
        MWindows.setBorderStyle Me, bsNone
    End If

    NotResize = False
    Form_Resize

End Sub

Public Sub mnuView_StatusBar_Click()

    If mnuView_StatusBar.Checked Then
        ShowStatusBar False
    Else
        ShowStatusBar True
    End If

End Sub

Public Sub mnuView_topmost_Click()

    If mnuView_TopMost.Checked = False Then
        mnuView_TopMost.Checked = True
        MWindows.setPosition Me, HWND_TOPMOST
    Else
        mnuView_TopMost.Checked = False
        MWindows.setPosition Me, HWND_NOTOPMOST
    End If

End Sub

Public Sub myXUnzip(ByVal sZipfilename As String, ByVal sFilesToProcess As String, ByVal sUnzipTo As String, Optional ByVal sPWD As String, Optional ByVal bUseFolderNames As Boolean = True)

    MainFrm.Enabled = False
    sFilesToProcess = LUseZipDll.CleanZipFilename(sFilesToProcess)

    With lUnzip
        .CaseSensitiveFileNames = False
        .OverwriteExisting = True
        .PromptToOverwrite = False
        .UseFolderNames = bUseFolderNames
        .ZipFile = sZipfilename
        '        .ZipFilename = sZipfilename
        .FileToProcess = sFilesToProcess
        .UnzipFolder = sUnzipTo
    End With

    lUnzip.unzip
    MainFrm.Enabled = True
    StsBar.Panels("ie").text = "[" & sZipfilename & "] " & sFilesToProcess & " ->Loaded."

End Sub

Public Sub myXZip(ByVal sZipfilename As String, ByVal sFilesToProcess As String, ByVal sPWD As String, Optional ByVal sBasePath As String = "")

    'Dim fso As New FileSystemObject
    StsBar.Panels("ie").text = "Writing " & "[" & sZipfilename & "]..."
    MainFrm.Enabled = False
    'sFilesToProcess = Replace(sFilesToProcess, "\", "/")
    sFilesToProcess = LUseZipDll.CleanZipFilename(sFilesToProcess)
    Set lZip = New cZip

    With lZip
        .ZipFile = sZipfilename
        .FileToProcess = sFilesToProcess
        .StoreDirectories = True
        .StoreFolderNames = False
        '       .FreshenFiles = True
        '.AllowAppend = False
        '.FilesToExclude = ""
        .BasePath = sBasePath
        '.EncryptionPassword = sPWD
    End With

    lZip.Zip
    Set lZip = Nothing
    MainFrm.Enabled = True
    StsBar.Panels("ie").text = "[" & sZipfilename & "] ->Saved."

End Sub

Public Sub saveCommentToZipfile(sComment As String, sZipFile As String)

    StsBar.Panels("ie").text = "正在重新压缩..."
    Set lZip = New cZip
    lZip.ZipFile = sZipFile
    lZip.ZipComment (sComment)
    Set lZip = Nothing
    StsBar.Panels("ie").text = "重新压缩完成。"

End Sub

'Sub saveMNUBookmark()
'
'    Dim bmCollection As typeZhBookmarkCollection
'    Dim i As Integer
'    Dim pos As Integer
'
'    With bmCollection
'        .Count = mnuBookmark.Count - 1
'
'        If .Count > 0 Then ReDim bmCollection.zhBookmark(.Count - 1) As typeZhBookmark
'
'        For i = 1 To mnuBookmark.Count - 1
'            .zhBookmark(i - 1).sName = mnuBookmark(i).Caption
'            pos = InStr(mnuBookmark(i).Tag, "|")
'
'            If pos > 0 Then
'                .zhBookmark(i - 1).sZhfile = Left$(mnuBookmark(i).Tag, pos - 1)
'                .zhBookmark(i - 1).sZhsubfile = Right$(mnuBookmark(i).Tag, Len(mnuBookmark(i).Tag) - pos)
'            End If
'
'        Next
'
'    End With
'
'    saveBookmark zhtmIni, bmCollection
'
'End Sub


Public Sub selectListItem()

    If zhrStatus.sCur_zhSubFile = "" Then Exit Sub
    Dim fIndex As Long
    Dim fcount As Long
    Dim fKey As String
    
    On Error Resume Next
    
    
    
'     If zhrStatus.iListIndex = lwContent Then
'        fIndex = sFilesINContent(2).Index(zhrStatus.sCur_zhSubFile)
'        fcount = sFilesINContent(2).Count
'        fKey = "ZTM" & sFilesINContent(1).Value(fIndex)
'    Else
        fIndex = sFilesInZip.Index(zhrStatus.sCur_zhSubFile)
        fcount = sFilesInZip.Count
        fKey = "ZTM" & sFilesInZip.value(fIndex)
    'End If
    StsBar.Panels("order").text = fIndex & "\" & fcount
    List.Nodes(fKey).Selected = True

End Sub

Private Sub ShowLeft(showit As Boolean)

    If showit Then
        mnuView_Left.Checked = True
        'zhrStatus.bLeftShowed = True

        LeftStrip.Tabs(1).Selected = True

   
    Else
        mnuView_Left.Checked = False
        'zhrStatus.bLeftShowed = False
    End If

    Form_Resize

End Sub

Private Sub ShowMenu(showit As Boolean)

    Dim i As Integer

    If showit = False Then
        mnuView_Menu.Checked = False
        'zhrStatus.bMenuShowed = False

        For i = 0 To mnu.Count - 1
            mnu(i).Visible = False
        Next

    Else
        mnuView_Menu.Checked = True
        'zhrStatus.bMenuShowed = True

        For i = 0 To mnu.Count - 1
            mnu(i).Visible = True
        Next

    End If

End Sub

Private Sub ShowStatusBar(showit As Boolean)

    If showit Then
        mnuView_StatusBar.Checked = True
        'zhrStatus.bStatusBarShowed = True
        StsBar.Height = 375
        Form_Resize
    Else
        mnuView_StatusBar.Checked = False
        'zhrStatus.bStatusBarShowed = False
        StsBar.Height = 0
        Form_Resize
    End If

End Sub





Private Sub Timer_Timer()

    Timer.Enabled = False

    If bRandomShow Then

        If randomView = False Then mnuGo_AutoRandom_Click
    Else
        mnuGo_Next_Click
    End If

End Sub

Sub zhReaderReset()

'    zhInfo.selfReset
    Dim i As Integer
    bInValidPassword = False
    bAutoShowNow = False
    bRandomShow = False
    mnuGo_AutoNext.Checked = False
    mnuGo_AutoRandom.Checked = False
    Timer.Enabled = False

'    For i = 1 To List.Count
'        List(i).Visible = False
'        List(i).Nodes.Clear
'        List(i).Tag = ""
'        List(i).Visible = True
'    Next
    
    With List
        .Visible = False
        .Nodes.Clear
        .Tag = ""
        .Visible = True
    End With
    'setListStatus List(0), lstNotloaded
    'zhInfo.selfReset

    With zhrStatus
        'If bIsZhtm Then .iListIndex = lwContent Else .iListIndex = lwFiles
        .sCur_zhFile = ""
        .sCur_zhSubFile = ""
    End With

    's_AI_DefaultFile = ""

'    Set sFoldersInZip = New cStringventor ' CStringCollection
'    lFoldersIZcount = 0
    'Set sFilesINContent(1) = New CStringVentor ' CStringCollection
    'Set sFilesINContent(2) = New CStringVentor ' CStringCollection
    'LeftStrip.Tabs(1).Selected = True
    'Navigated = False
    'appHtmlAbout
    '    Do
    '        DoEvents
    '    Loop While IEView.ReadyState = READYSTATE_LOADING

End Sub


'Public Function LoadSetting()
'
'    Dim hSetting As New CSetting
'    hSetting.iniFile = zhtmini
'    hSetting.Load cmbAddress, SF_LISTTEXT
'    hSetting.Load Me, SF_FORM
'    hSetting.Load Me, SF_FONT
'    hSetting.Load Me, SF_COLOR
'    hSetting.Load mnuFile_Recent, SF_MENUARRAY
'    hSetting.Load mnuBookmark, SF_MENUARRAY
'    hSetting.Load mnuView_Left, SF_CHECKED
'    hSetting.Load mnuView_Menu, SF_CHECKED
'    hSetting.Load mnuView_StatusBar, SF_CHECKED
'    hSetting.Load mnuView_FullScreen, SF_CHECKED
'    hSetting.Load mnuView_AddressBar, SF_CHECKED
'    hSetting.Load mnuView_TopMost, SF_CHECKED
'    hSetting.Load mnuView_ApplyStyleSheet, SF_CHECKED
'    hSetting.Load mnuEdit_SelectEditor, SF_CHECKED
'    hSetting.Load Timer, SF_Tag
'    Set hSetting = Nothing
'
'End Function
'Public Function WriteSetting()
'    Dim hSetting As New CSetting
'    hSetting.iniFile = zhtmini
'    hSetting.Save cmbAddress, SF_LISTTEXT
'    hSetting.Save Me, SF_FORM
'    hSetting.Save mnuFile_Recent, SF_MENUARRAY
'    hSetting.Save mnuBookmark, SF_MENUARRAY
'    hSetting.Save mnuView_Left, SF_CHECKED
'    hSetting.Save mnuView_Menu, SF_CHECKED
'    hSetting.Save mnuView_StatusBar, SF_CHECKED
'    hSetting.Save mnuView_FullScreen, SF_CHECKED
'    hSetting.Save mnuView_AddressBar, SF_CHECKED
'    hSetting.Save mnuView_TopMost, SF_CHECKED
'    hSetting.Save mnuView_ApplyStyleSheet, SF_CHECKED
'    hSetting.Save mnuEdit_SelectEditor, SF_CHECKED
'    Set hSetting = Nothing
'End Function

'Public Function KeyExistInNodes(nNodes As Nodes, sKey As String) As Boolean
'    KeyExistInNodes = False
'    On Error GoTo 0
'    nNodes.Item(sKey).Tag
'    KeyExistInNodes = True
'End Function

Private Function nodeExist(pNodes As Nodes, key As String) As Boolean
nodeExist = False
Dim tmp As Node
On Error Resume Next
Err.Clear
Set tmp = pNodes(key)
If Err.Number = 0 Then nodeExist = True
'Dim lastIndex As Integer
'Dim i As Integer
'lastIndex = pNodes.Count
'For i = 1 To lastIndex
'    If pNodes(i).key = key Then nodeExist = True: Exit Function
'Next
End Function

Private Sub xAddNode(pNodes As Nodes, folderName As String, keyPrfix As String, textPrfix As String, indexImg As Integer)
Dim key As String
Dim text As String
Dim pfdName As String
key = keyPrfix & folderName
text = linvblib.GetBaseName(folderName)
pfdName = linvblib.GetParentFolderName(folderName)
pfdName = linvblib.bdUnixDir(pfdName, "")

If nodeExist(pNodes, key) Then Exit Sub
If pfdName = "/" Then
    pNodes.Add , , key, text, indexImg
Else
    xAddNode pNodes, pfdName, keyPrfix, textPrfix, indexImg
    pNodes.Add keyPrfix & pfdName, tvwChild, key, text, indexImg
End If
End Sub


Private Sub VBOXScroll_Change()
    ImageBox.Move ImageBox.Left, -VBOXScroll.value
End Sub

Private Sub hboxscroll_change()
    ImageBox.Move -HBoxScroll.value, ImageBox.Top
End Sub

Private Sub VBOXScroll_Scroll()
    ImageBox.Move ImageBox.Left, -VBOXScroll.value
End Sub

