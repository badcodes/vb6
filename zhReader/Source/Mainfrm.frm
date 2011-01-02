VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form MainFrm 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   ClientHeight    =   3720
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   7815
   Icon            =   "Mainfrm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   NegotiateMenus  =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   7815
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Interval        =   2500
      Left            =   4080
      Top             =   2520
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
      TabIndex        =   7
      Top             =   -1248
      Visible         =   0   'False
      Width           =   72
   End
   Begin ComctlLib.StatusBar StsBar 
      Height          =   372
      Left            =   -288
      TabIndex        =   1
      Top             =   3384
      Width           =   8076
      _ExtentX        =   14235
      _ExtentY        =   635
      SimpleText      =   ""
      ShowTips        =   0   'False
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   9260
            Key             =   "ie"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   3440
            MinWidth        =   3440
            Key             =   "reading"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1429
            MinWidth        =   1411
            Key             =   "order"
            Object.Tag             =   ""
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
   Begin VB.Frame theShow 
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   2244
      Left            =   3240
      TabIndex        =   0
      Top             =   120
      Width           =   3696
      Begin SHDocVwCtl.WebBrowser IEView 
         Height          =   1704
         Left            =   276
         TabIndex        =   9
         Top             =   408
         Width           =   3300
         ExtentX         =   5821
         ExtentY         =   3006
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   1
         RegisterAsDropTarget=   0
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
         Location        =   ""
      End
      Begin VB.ComboBox cmbAddress 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   312
         ItemData        =   "Mainfrm.frx":08CA
         Left            =   240
         List            =   "Mainfrm.frx":08CC
         TabIndex        =   8
         Top             =   0
         Width           =   3492
      End
   End
   Begin VB.PictureBox LeftFrame 
      Height          =   2772
      Left            =   300
      ScaleHeight     =   2715
      ScaleWidth      =   2505
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
         Begin ComctlLib.TreeView List 
            Height          =   1452
            Index           =   1
            Left            =   -36
            TabIndex        =   4
            Top             =   600
            Width           =   2292
            _ExtentX        =   4022
            _ExtentY        =   2566
            _Version        =   327682
            HideSelection   =   0   'False
            Indentation     =   423
            LabelEdit       =   1
            LineStyle       =   1
            Style           =   7
            ImageList       =   "Listimg"
            BorderStyle     =   1
            Appearance      =   1
         End
         Begin ComctlLib.TreeView List 
            Height          =   1452
            Index           =   2
            Left            =   96
            TabIndex        =   5
            Top             =   -528
            Width           =   2292
            _ExtentX        =   4048
            _ExtentY        =   2566
            _Version        =   327682
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
      Begin ComctlLib.TabStrip LeftStrip 
         Height          =   4812
         Left            =   108
         TabIndex        =   6
         Top             =   0
         Width           =   2412
         _ExtentX        =   4260
         _ExtentY        =   8493
         MultiRow        =   -1  'True
         _Version        =   327682
         BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
            NumTabs         =   2
            BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Content"
               Key             =   "TabList"
               Object.Tag             =   ""
               Object.ToolTipText     =   "ContentList"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Files"
               Key             =   "TABfile"
               Object.Tag             =   ""
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
   Begin ComctlLib.ImageList Listimg 
      Left            =   5640
      Top             =   2520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
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
      Caption         =   "&Edit"
      Index           =   1
      Begin VB.Menu mnuEdit_EditCurPage 
         Caption         =   "&Edit Current Page"
      End
      Begin VB.Menu mnuEdit_EditInfo 
         Caption         =   "Edit zhFile &Info"
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit_Delete 
         Caption         =   "&Delete This Page"
      End
      Begin VB.Menu mnuSep21 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit_SelectEditor 
         Caption         =   "Select Text Editor..."
      End
      Begin VB.Menu mnuEdit_SetDefault 
         Caption         =   "Set Current Page As Default"
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
      Begin VB.Menu mnuView_AddressBar 
         Caption         =   "&AddressBar"
      End
      Begin VB.Menu mnuView_TopMost 
         Caption         =   "&TopMost Made"
      End
      Begin VB.Menu mnuViewsep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuView_ApplyStyleSheet 
         Caption         =   "Apply StyleSheet"
         Begin VB.Menu mnuView_ApplyStyleSheet_List 
            Caption         =   "None"
            Index           =   0
         End
         Begin VB.Menu mnuView_ApplyStyleSheet_List 
            Caption         =   "Default Style"
            Index           =   1
         End
         Begin VB.Menu mnuView_ApplyStyleSheet_List 
            Caption         =   "-"
            Index           =   2
         End
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

Private sFilesInZip As New CStringArray 'Collection
'Private sFoldersInZip As CStringArray  ' CStringCollection
'FIXIT: Non Zero lowerbound arrays are not supported in Visual Basic .NET                  FixIT90210ae-R9815-H1984
Private sFilesINContent(1 To 2) As New CStringArray  'CStringCollection
'Private sfilesinzip.count As Long
'Private lFoldersIZcount As Long
'Public bFullScreen As Boolean
Private WithEvents ieViewV1  As SHDocVwCtl.WebBrowser_V1
Attribute ieViewV1.VB_VarHelpID = -1
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

Private m_PortableMode As String

Public Function curFileIndex(sCurFileInZip As String) As Long

    If zhrStatus.iListIndex = lwContent Then
        curFileIndex = sFilesINContent(2).Find(sCurFileInZip)
    Else
        curFileIndex = sFilesInZip.Find(sCurFileInZip)
    End If

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

Function getZhCommentText(ByVal sZipfilename As String) As String

    'Set lUnzip = New cUnzip
    lUnzip.ZipFile = sZipfilename
    lUnzip.Comment = ""
    getZhCommentText = toUnixPath(lUnzip.GetComment)
    'Set lUnzip = Nothing

End Function

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
    Static lastFileList As ListWhat
    Dim i As Long
    

    If zhrStatus.sCur_zhFile = "" Then Exit Function

    If curZhtm <> zhrStatus.sCur_zhFile Or bRestart Or lastFileList <> zhrStatus.iListIndex Then
        lastFileList = zhrStatus.iListIndex
        curZhtm = zhrStatus.sCur_zhFile
        iViewNow = 0

        If lastFileList = lwContent Then
            iViewLast = sFilesINContent(1).Count
        Else
            iViewLast = sFilesInZip.Count
        End If
        
        If iViewLast < 1 Then
            randomView = False
            Exit Function
        End If
        
'FIXIT: Non Zero lowerbound arrays are not supported in Visual Basic .NET                  FixIT90210ae-R9815-H1984
        ReDim iRandomArr(1 To iViewLast) As Long
        For i = 1 To iViewLast
            iRandomArr(i) = i
        Next
        MAlgorithms.BedlamArr iRandomArr, 1, iViewLast '打乱数组
    End If

    iViewNow = iViewNow + 1

    If iViewNow > iViewLast Then iViewNow = 1

    If lastFileList = lwContent Then
        GetView sFilesINContent(1).Item(iRandomArr(iViewNow))
    Else
        GetView sFilesInZip.Item(iRandomArr(iViewNow))
    End If

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

    Dim zhLocalize As New CLocalizer
    zhLocalize.Install Me, LanguageIni
    StrLocalize = zhLocalize.loadLangStr(sEnglish)
    Set zhLocalize = Nothing

End Function

Public Sub AppHtmlAbout()

    Dim fNum As Integer
    Dim sAppHtmlAboutFile As String
    Dim sAppend As New CAppendString

    sAppHtmlAboutFile = linvblib.BuildPath(Tempdir, cHtmlAboutFilename)

    If linvblib.PathExists(sAppHtmlAboutFile) = False Then

        With sAppend
            .AppendLine "<html>"
            .AppendLine htmlline("<link REL=!stylesheet! href=!" & sConfigDir & "/style.css! type=!text/css!>")
            .AppendLine htmlline("<head>")
            .AppendLine htmlline("<Title>Zippacked Html Reader</title>")
            .AppendLine htmlline("<meta http-equiv=!Content-Type! content=!text/html; charset=us-ascii!>")
            .AppendLine htmlline("</head>")
            .AppendLine htmlline("<body class=!m_text!><p><br>")
            .AppendLine htmlline("<p align=right ><span lang=EN-US style=!font-size:24.0pt;font-family:TAHOMA,Courier New! class=!m_text!>" & App.ProductName)
            If LenB(m_PortableMode) > 0 Then .AppendLine htmlline(" Portable")
            
'FIXIT: App.Revision property has no Visual Basic .NET equivalent and will not be upgraded.     FixIT90210ae-R7593-R67265
            .AppendLine htmlline(" (Build" & Str$(App.Major) + "." + Str$(App.Minor) & "." & Str$(App.Revision) & ")</span></p>")
            .AppendLine htmlline("<p align=right ><span lang=EN-US style=!font-size:24.0pt;font-family:TAHOMA,Courier New! class=!m_text!>" & App.LegalCopyright & "</span></span></p>")
            .AppendLine htmlline("</body>")
            .AppendLine htmlline("</html>")
        End With

        fNum = FreeFile
        Open sAppHtmlAboutFile For Output As #fNum
'FIXIT: Print method has no Visual Basic .NET equivalent and will not be upgraded.         FixIT90210ae-R7593-R67265
        Print #fNum, sAppend.Value
        Close #fNum
        Set sAppend = Nothing
    End If

    NotPreOperate = True
    IEView.Navigate2 sAppHtmlAboutFile

End Sub


Private Sub ApplyDefaultStyle(fApply As Boolean)

    Const CSSID = "zhReaderCSS"
    Dim curCss As String
    Dim iIndex As Long
    Dim ALLCSS As HTMLStyleSheetsCollection
    Dim ICSS As IHTMLStyleSheet
    

    'On Error GoTo 0

    Set ALLCSS = IEView.Document.styleSheets
    For Each ICSS In ALLCSS
    If ICSS.Title = CSSID Then ICSS.href = "": ICSS.Title = ""
    Next
         
    If fApply = False Then Exit Sub
        
'FIXIT: mnuView_ApplyStyleSheet.Tag property has no Visual Basic .NET equivalent and will not be upgraded.     FixIT90210ae-R7593-R67265
    iIndex = Val(mnuView_ApplyStyleSheet.Tag)
    If iIndex = 0 Then Exit Sub
    If iIndex = 1 Then
        curCss = linvblib.bdUnixDir(sConfigDir, "Style.css")
        If linvblib.PathExists(curCss) = False Then
                Load frmOptions
                frmOptions.MakeCss curCss
                Unload frmOptions
        End If
    Else
'FIXIT: mnuView_ApplyStyleSheet_List(iIndex).Tag property has no Visual Basic .NET equivalent and will not be upgraded.     FixIT90210ae-R7593-R67265
        curCss = mnuView_ApplyStyleSheet_List(iIndex).Tag
    End If
    
    'if linvblib.PathExists (curcss)=False then
    Set ICSS = IEView.Document.createStyleSheet(curCss)
    ICSS.Title = CSSID
    

End Sub

Public Sub ChangeMainFrmCaption(sTitle As String)

    If mnuView_FullScreen.Checked Then Exit Sub
    MainFrm.Caption = sTitle

End Sub

Public Sub Load_StyleSheetList()
    Dim fso As New FileSystemObject
    Dim fs As Files
    Dim f As File
    Dim iIndex As Long
    Dim iLast As Long
    On Error Resume Next

    iLast = mnuView_ApplyStyleSheet_List.UBound
    For iIndex = iLast To 3 Step -1
    Unload mnuView_ApplyStyleSheet_List(iIndex)
    Next
    
    iIndex = 2
    
    Set fs = fso.GetFolder(fso.BuildPath(App.Path, "CSS")).Files
    
    If TypeName(fs) <> "Nothing" Then
        For Each f In fs
            iIndex = iIndex + 1
            Load mnuView_ApplyStyleSheet_List(iIndex)
            mnuView_ApplyStyleSheet_List(iIndex).Checked = False
            mnuView_ApplyStyleSheet_List(iIndex).Visible = True
            mnuView_ApplyStyleSheet_List(iIndex).Caption = f.name
'FIXIT: mnuView_ApplyStyleSheet_List(iIndex).Tag property has no Visual Basic .NET equivalent and will not be upgraded.     FixIT90210ae-R7593-R67265
            mnuView_ApplyStyleSheet_List(iIndex).Tag = f.Path
        Next
    End If
    
    Set f = Nothing
    Set fs = Nothing
    Set fs = fso.GetFolder(fso.BuildPath(sConfigDir, "CSS")).Files
    If TypeName(fs) <> "Nothing" Then
        For Each f In fs
            iIndex = iIndex + 1
            Load mnuView_ApplyStyleSheet_List(iIndex)
            mnuView_ApplyStyleSheet_List(iIndex).Checked = False
            mnuView_ApplyStyleSheet_List(iIndex).Visible = True
            mnuView_ApplyStyleSheet_List(iIndex).Caption = f.name
'FIXIT: mnuView_ApplyStyleSheet_List(iIndex).Tag property has no Visual Basic .NET equivalent and will not be upgraded.     FixIT90210ae-R7593-R67265
            mnuView_ApplyStyleSheet_List(iIndex).Tag = f.Path
        Next
    End If
    Set fso = Nothing
End Sub

Private Sub cmbAddress_Click()

    Dim txtCmb As String
    txtCmb = cmbAddress.text

    If txtCmb <> "" Then MNavigate txtCmb, IEView

End Sub

Private Sub cmbAddress_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii
    Case Asc(vbCr)
        Dim txtCmb As String
        txtCmb = cmbAddress.text

        If txtCmb <> "" Then MNavigate txtCmb, IEView
        AddUniqueItem cmbAddress, txtCmb
        '        Dim iIndex As Long
        '        Dim iEnd As Long
        '        iEnd = cmbAddress.ListCount - 1
        '        For iIndex = 0 To iEnd
        '        If StrComp(cmbAddress.List(iIndex), txtCmb, vbTextCompare) = 0 Then Exit Sub
        '        Next
        '        cmbAddress.AddItem txtCmb
    Case Asc(vbKeyEscape)

        If cmbAddress.ListCount > 1 Then cmbAddress.text = cmbAddress.List(cmbAddress.ListCount - 1) '  IEView.LocationURL
    End Select

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
            IEView.Stop
            Timer.Enabled = False
            mnuGo_AutoRandom.Checked = False
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
            IEZoom IEView, 1
        Case vbKeySubtract
            IEZoom IEView, -1
        Case Else
            KeyCode = iKeyCode
        End Select

    End If

End Sub
Public Sub IEZoom(iehost As WebBrowser, inOrOut As Integer)
'FIXIT: Declare 'iLevel' with an early-bound data type                                     FixIT90210ae-R1672-R1B8ZE
    Dim iLevel As Variant
    
    Const ZoomLevelMin As Integer = 0
    Const ZoomLevelMax As Integer = 4
    
    
    iehost.ExecWB OLECMDID_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, Null, iLevel
    
    iLevel = iLevel + inOrOut
    If iLevel < ZoomLevelMin Then iLevel = 0
    If iLevel > ZoomLevelMax Then iLevel = 4
    
    iehost.ExecWB OLECMDID_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, iLevel, Null
        
End Sub


'
Private Sub Form_Load()

    MWindows.InitCommonControlsVB
    'Dim fso As LiNVBLib.GFileSystem
    'Set fso = New LiNVBLib.GFileSystem
    NotResize = True
    
    On Error Resume Next
    'Public String
    
    
    Dim GlobalConfig As String
    
    GlobalConfig = linvblib.BuildPath(App.Path, "zhReader.ini")
'    MsgBox GlobalConfig
    If FileExists(GlobalConfig) Then
        Dim GlobalIni As CLiNInI
        Set GlobalIni = New CLiNInI
        GlobalIni.File = GlobalConfig
        m_PortableMode = GlobalIni.GetSetting("zhReader", "PortableMode")
        Set GlobalIni = Nothing
    End If
    
    
    If LenB(m_PortableMode) > 0 Then
        sConfigDir = BuildPath(App.Path, "Profile")
    Else
        sConfigDir = linvblib.BuildPath(Environ$("APPDATA"), "zhReader")
    End If
    If linvblib.FolderExists(sConfigDir) = False Then linvblib.CreateFolder (sConfigDir)
    zhtmIni = linvblib.BuildPath(sConfigDir, "config.ini")
    
    LanguageIni = linvblib.BuildPath(sConfigDir, "Language.ini")
    Tempdir = linvblib.BuildPath(sConfigDir, "Cache")
    If linvblib.FolderExists(Tempdir) Then linvblib.DeleteFolder Tempdir, True, True
    If linvblib.FolderExists(Tempdir) = False Then linvblib.CreateFolder Tempdir
    Tempdir = linvblib.toUnixPath(Tempdir)
    
    
    Dim zhLocalize As CLocalizer
    Set zhLocalize = New CLocalizer
    zhLocalize.Install Me, LanguageIni
    zhLocalize.loadFormStr
    Set zhLocalize = Nothing

    Set lUnzip = New cUnzip
    
    IEView.Navigate2 "about:blank"
    Set ieViewV1 = IEView.object
    
    On Error Resume Next
    

    If linvblib.FileExists(zhtmIni) Then
    '-----------------------------------------------------------------------------------
    '-------------------------- Call LoadSetting Start-----------------------------
        Dim hSetting As CSetting
        Set hSetting = New CSetting
        hSetting.iniFile = zhtmIni
        hSetting.Load cmbAddress, SF_LISTTEXT                           'Text
        hSetting.Load Me, SF_FORM                                               'Postion
        'hsetting.Load Me, SF_FONT                                               'Font For ViewStyle
        'hsetting.Load Me, SF_COLOR                                             'Color For ViewStyle
        hSetting.Load mnuFile_Recent, SF_MENUARRAY, "RecentFiles"               'Recent File
        hSetting.Load mnuBookmark, SF_MENUARRAY, "Bookmarks"                  'BookMark
        hSetting.Load mnuView_Left, SF_CHECKED                        'ShowLeft Check
        hSetting.Load mnuView_Menu, SF_CHECKED                       'ShowMenu Check
        hSetting.Load mnuView_StatusBar, SF_CHECKED                'ShowstatusBar Check
        hSetting.Load mnuView_FullScreen, SF_CHECKED               'FullScreen Check
        hSetting.Load mnuView_AddressBar, SF_CHECKED             'ShowAddressBar Check
        hSetting.Load mnuView_TopMost, SF_CHECKED                  'OnTopMost Check
        'hSetting.Load mnuView_ApplyStyleSheet, SF_CHECKED    'ApplyDefaultStyle Check
        hSetting.Load mnuEdit_SelectEditor, SF_Tag              'TextEditor
        hSetting.Load mnuView_ApplyStyleSheet, SF_Tag            'StyleSheet Path
        hSetting.Load Timer, SF_Tag                                             'Time InterVal
        hSetting.Load mnuFile_Open, SF_Tag                                 'InitDir for Dialog
        hSetting.Load Me, SF_Tag                                                  'UseTemplate ?
        hSetting.Load IEView, SF_Tag                                            'TemplateHtml
        hSetting.Load Me.picSplitter, SF_Tag                                 'LeftWidth
        Set hSetting = Nothing
    '-------------------------- Call LoadSetting END----------------------------
    '----------------------------------------------------------------------------------
    End If
    mnuView_Left.Checked = Not mnuView_Left.Checked
    mnuView_Menu.Checked = Not mnuView_Menu.Checked
    mnuView_StatusBar.Checked = Not mnuView_StatusBar.Checked
    mnuView_AddressBar.Checked = Not mnuView_AddressBar.Checked
    mnuView_TopMost.Checked = Not mnuView_TopMost.Checked
    mnuView_ApplyStyleSheet.Checked = Not mnuView_ApplyStyleSheet.Checked
    mnuView_Left_Click
    mnuView_Menu_Click
    mnuView_StatusBar_Click
    If mnuView_FullScreen.Checked Then
        mnuView_FullScreen.Checked = False
        mnuView_fullscreen_Click
    End If
    mnuView_AddressBar_Click
    mnuView_topmost_Click
    
    
    If picSplitter.Tag = 0 Then picSplitter.Tag = icstDefaultLeftWidth
    picSplitter.Left = picSplitter.Tag '= thers.LeftWidth
    
'FIXIT: Timer.Tag property has no Visual Basic .NET equivalent and will not be upgraded.     FixIT90210ae-R7593-R67265
'FIXIT: Timer.Tag property has no Visual Basic .NET equivalent and will not be upgraded.     FixIT90210ae-R7593-R67265
    If Timer.Tag = 0 Then Timer.Tag = defaultAutoViewTime
'FIXIT: Timer.Tag property has no Visual Basic .NET equivalent and will not be upgraded.     FixIT90210ae-R7593-R67265
    Timer.Interval = Timer.Tag
'FIXIT: mnuFile_Recent(0).Tag property has no Visual Basic .NET equivalent and will not be upgraded.     FixIT90210ae-R7593-R67265
'FIXIT: mnuFile_Recent(0).Tag property has no Visual Basic .NET equivalent and will not be upgraded.     FixIT90210ae-R7593-R67265
    If MainFrm.mnuFile_Recent(0).Tag = 0 Then mnuFile_Recent(0).Tag = defaultRecentFileList
    
    If mnuFile_Recent.Count > 1 Then
    mnuFile_Recent(0).Visible = True
    Else
    mnuFile_Recent(0).Visible = False
    End If
    
    Load_StyleSheetList
    Dim iIndex As Integer
'FIXIT: mnuView_ApplyStyleSheet.Tag property has no Visual Basic .NET equivalent and will not be upgraded.     FixIT90210ae-R7593-R67265
    iIndex = Val(mnuView_ApplyStyleSheet.Tag)
    mnuView_ApplyStyleSheet_List(iIndex).Checked = True
    
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
    Dim resID As Integer
    Randomize Timer
    resID = MAlgorithms.getRndnum(101, 104)
    Set MainFrm.Icon = LoadResPicture(resID, vbResIcon)
    
    Listimg.ListImages.Clear
    For resID = 105 To 107
        Listimg.ListImages.Add resID - 104, , LoadResPicture(resID, vbResIcon)
    Next
'    For i = 1 To 4
'
'        If fso.PathExists(icofile(i)) Then
'            Listimg.ListImages.Remove (i)
'            Listimg.ListImages.Add i, , LoadPicture(icofile(i))
'        End If
'
'    Next
   
    Dim sInitDir As String
'FIXIT: mnuFile_Open.Tag property has no Visual Basic .NET equivalent and will not be upgraded.     FixIT90210ae-R7593-R67265
    sInitDir = mnuFile_Open.Tag
    If linvblib.FolderExists(sInitDir) Then
    ChDir (sInitDir)
    ChDrive (Left$(sInitDir, 1))
    End If
    
    NotResize = False


    
End Sub

Private Sub Form_Resize()

    If NotResize Then Exit Sub
    Dim tempint As Integer

    If MainFrm.WindowState <> 1 Then

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

    Dim i As Integer

    For i = 1 To List.Count

        With List(i)
            .Top = 0
            .Left = 0
            .Width = ListFrame.Width
            .Height = ListFrame.Height
        End With

    Next

    With cmbAddress
        .Left = 0
        .Top = 0
        .Width = theShow.Width
    End With

    'For i = 1 To IEView.Count

    With IEView '(i - 1)
        .Left = 0

        If cmbAddress.Visible Then .Top = cmbAddress.Height Else .Top = 0
        .Height = theShow.Height - .Top
        .Width = theShow.Width
    End With

    End If

    

    'Next
    MainFrm.Refresh

End Sub

Private Sub Form_Unload(Cancel As Integer)

    
    
     '-----------------------------------------------------------------------------------
    '-------------------------- Call SaveSetting Start-----------------------------
        Dim hSetting As CSetting
        Set hSetting = New CSetting
        hSetting.iniFile = zhtmIni
        hSetting.Save cmbAddress, SF_LISTTEXT                           'Text
        hSetting.Save Me, SF_FORM                                               'Postion
        'hsetting.Save Me, SF_FONT                                               'Font For ViewStyle
        'hsetting.Save Me, SF_COLOR                                             'Color For ViewStyle
        hSetting.Save mnuFile_Recent, SF_MENUARRAY, "RecentFiles"              'Recent File
        hSetting.Save mnuBookmark, SF_MENUARRAY, "Bookmarks"                  'BookMark
        hSetting.Save mnuView_Left, SF_CHECKED                        'ShowLeft Check
        hSetting.Save mnuView_Menu, SF_CHECKED                       'ShowMenu Check
        hSetting.Save mnuView_StatusBar, SF_CHECKED                'ShowstatusBar Check
        hSetting.Save mnuView_FullScreen, SF_CHECKED               'FullScreen Check
        hSetting.Save mnuView_AddressBar, SF_CHECKED             'ShowAddressBar Check
        hSetting.Save mnuView_TopMost, SF_CHECKED                  'OnTopMost Check
        'hSetting.Save mnuView_ApplyStyleSheet, SF_CHECKED    'ApplyDefaultStyle Check
        hSetting.Save mnuEdit_SelectEditor, SF_Tag                     'TextEditor
        hSetting.Save mnuView_ApplyStyleSheet, SF_Tag            'StyleSheet Path
        hSetting.Save Timer, SF_Tag                                             'Time InterVal
        hSetting.Save mnuFile_Open, SF_Tag                                 'InitDir for Dialog
        hSetting.Save Me, SF_Tag                                                  'UseTemplate ?
        hSetting.Save IEView, SF_Tag                                            'TemplateHtml
        hSetting.Save Me.picSplitter, SF_Tag                                 'LeftWidth
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
    Set zhInfo = Nothing
    Set lUnzip = Nothing
    Set lZip = Nothing
    Set sFilesInZip = Nothing
    'Set sFoldersInZip = Nothing
    Set sFilesINContent(1) = Nothing
    Set sFilesINContent(2) = Nothing
    Erase sFilesINContent
    Unload frmBookmark
    Unload frmOptions
    Call endUP

End Sub

Public Sub GetView(ByVal shortfile As String)

    Call MGetView(shortfile, IEView)

End Sub
Public Sub getZHContent(ByRef cZhCommentToLoad As CZhComment)

Dim sArrZhContent() As String
Dim tmpfile As String

    If cZhCommentToLoad.lContentCount > 0 Then
        cZhCommentToLoad.CopyContentTo sArrZhContent
    End If

    If cZhCommentToLoad.lContentCount < 1 And cZhCommentToLoad.sContentFile <> "" Then
        tmpfile = cZhCommentToLoad.sContentFile
        myXUnzip zhrStatus.sCur_zhFile, tmpfile, sTempZH, zhrStatus.sPWD
        tmpfile = linvblib.bdUnixDir(sTempZH, tmpfile)

        If linvblib.PathExists(tmpfile) Then
            cZhCommentToLoad.parseZhComment (tmpfile)

            If cZhCommentToLoad.lContentCount > 0 Then
                cZhCommentToLoad.CopyContentTo sArrZhContent
            End If

        End If

    End If

    If cZhCommentToLoad.lContentCount < 1 And cZhCommentToLoad.sHHCFile <> "" And cZhCommentToLoad.sHHCFile <> "none" Then
        tmpfile = cZhCommentToLoad.sHHCFile
        myXUnzip zhrStatus.sCur_zhFile, tmpfile, sTempZH, zhrStatus.sPWD
        tmpfile = linvblib.bdUnixDir(sTempZH, tmpfile)

        If linvblib.PathExists(tmpfile) Then
            cZhCommentToLoad.parseHHC tmpfile, linvblib.GetParentFolderName(cZhCommentToLoad.sHHCFile)

            If cZhCommentToLoad.lContentCount > 0 Then
                cZhCommentToLoad.CopyContentTo sArrZhContent
                saveZhInfo '!!!!!!!!!!!!!!!!!!!!!!
            End If

        End If

    End If

    Dim i As Long
    Dim iEnd As Long
    iEnd = cZhCommentToLoad.lContentCount - 1

    For i = 0 To iEnd

        If sArrZhContent(1, i) <> "" And Right$(sArrZhContent(1, i), 1) <> "/" Then
            sFilesINContent(1).Add sArrZhContent(0, i)
            sFilesINContent(2).Add sArrZhContent(1, i)
        End If
    Next

End Sub
Public Sub getZIPContent(ByVal sZipfilename As String)

    Dim lfor As Long
    Dim sFilename As String
    Dim sExtname As String
    Dim bRequireDefaultFile As Boolean
    Dim bRequireHHC As Boolean
    Dim sDefaultfile As String
    Dim iCountSL As Integer ' Count "\" in sFilename
    Dim iMinSL As Integer
    iMinSL = 100 ' 设为最大值
    
    With zhInfo
    If .sDefaultfile = "" Then bRequireDefaultFile = True
    If .sHHCFile = "" And .sContentFile = "" And .lContentCount <= 0 Then bRequireHHC = True
    End With
    
    Dim zipFileList As New CZipItems
    lUnzip.ZipFile = sZipfilename
    lUnzip.getZipItems zipFileList
    
    Dim lEnd As Long
    
    lEnd = zipFileList.Count
    StsBar.Panels("ie").text = "Scanning " & lUnzip.ZipFile & "..."
    
    If bRequireDefaultFile = False Then
        For lfor = 1 To lEnd
            'StsBar.Panels(1).text = "已扫描到" & lfor & "个文件..."
            DisplayProgress lfor, lEnd
            sFilename = zipFileList(lfor).FileName
            If Right$(sFilename, 1) <> "/" Then
                    sFilesInZip.Add sFilename
                    If bRequireHHC Then
                        sExtname = LCase$(linvblib.GetExtensionName(sFilename))
                         If sExtname = "hhc" Then
                            zhInfo.sHHCFile = sFilename
                            bRequireHHC = False
                        End If
                    End If
            End If
        Next
    Else
        For lfor = 1 To lEnd
            'StsBar.Panels(1).text = "已扫描到" & lfor & "个文件..."
            DisplayProgress lfor, lEnd
            sFilename = zipFileList(lfor).FileName
            sExtname = LCase$(linvblib.GetExtensionName(sFilename))
            If Right$(sFilename, 1) <> "/" Then
                sFilesInZip.Add sFilename
                If bRequireDefaultFile And _
                   (sExtname = "htm" Or sExtname = "html") And _
                   IsWebsiteDefaultFile(sFilename) Then
                        iCountSL = Len(sFilename)  'linvblib.charCountInStr(sFilename, "\", vbBinaryCompare)
                        If iCountSL < iMinSL Then
                            iMinSL = iCountSL
                            sDefaultfile = sFilename
                            If iMinSL = 0 Then bRequireDefaultFile = False
                        End If
                End If
                If bRequireHHC And sExtname = "hhc" Then
                    zhInfo.sHHCFile = sFilename
                    bRequireHHC = False
                End If
            End If

        Next
        
         If sDefaultfile = "" And sFilesInZip.Count > 0 Then
            sDefaultfile = sFilesInZip.Item(1)
        End If
              
    End If
    Set zipFileList = Nothing
    
        If sDefaultfile <> "" Or zhInfo.sHHCFile = "" Or zhInfo.sTitle = "" Then
            If sDefaultfile <> "" Then zhInfo.sDefaultfile = sDefaultfile
            If zhInfo.sHHCFile = "" Then zhInfo.sHHCFile = "none"
            If zhInfo.sTitle = "" Then zhInfo.sTitle = linvblib.GetBaseName(sZipfilename)
            'saveCommentToZipfile zhInfo.ToString, zhrStatus.sCur_zhFile
        End If
    

    StsBar.Panels(1).text = "Scanning done!"
      
    MainFrm.Enabled = True



End Sub

'FIXIT: Declare 'pDisp' and 'URL' and 'flags' and 'targetFrameName' and 'PostData' and 'Headers' with an early-bound data type     FixIT90210ae-R1672-R1B8ZE
Private Sub IEView_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, flags As Variant, targetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)

    If isZhcommand(URL) Then
        Cancel = True
        execZhcommand URL
        Exit Sub
    End If

    If NotPreOperate Then
        NotPreOperate = False
        Exit Sub
    End If

    'iCurIEView = Index
    Call eBeforeNavigate(URL, Cancel, IEView, CStr(targetFrameName))
    
    Call selectListItem

End Sub





'FIXIT: Declare 'pDisp' and 'URL' with an early-bound data type                            FixIT90210ae-R1672-R1B8ZE
Private Sub IEView_DocumentComplete(ByVal pDisp As Object, URL As Variant)

    On Error Resume Next
    Call eNavigateComplete(URL, IEView)
    'Set IEView.Document = IEView.Document
    MainFrm.LeftFrame.Enabled = True
    'MainFrm.Navigated = True
    Dim sTitle As String

    sTitle = IEView.Document.Title
    If zhrStatus.sCur_zhFile <> "" Or zhrStatus.sCur_zhSubFile <> "" Then
    
        If IEView.Document.Title = "" Then
            sTitle = GetFileName(zhrStatus.sCur_zhSubFile)
        Else
            sTitle = IEView.Document.Title
        End If
        
        If zhInfo.sTitle = "" Then
            sTitle = GetBaseName(zhrStatus.sCur_zhFile) & " - " & sTitle
        Else
            sTitle = zhInfo.sTitle & " - " & sTitle
        End If
    
    End If
    
    ChangeMainFrmCaption sTitle
    
    IEView.Document.focus
    ApplyDefaultStyle True

    If bScroll Then
        bScroll = False
        Dim x As Long, y As Long
        x = CLng(scrollX * IEView.Document.body.scrollWidth)
        y = CLng(scrollY * IEView.Document.body.scrollHeight)
        IEView.Document.parentWindow.scrollTo x, y
    End If
    
    If bAutoShowNow Then Timer.Enabled = True
    
'    Dim testDoc As IHTMLDocument2
'    Dim eAll As IHTMLElementCollection
'    Dim e As IHTMLElement
'    Set testDoc = IEView.Document
'    Set eAll = testDoc.All
'    For Each e In eAll
'    Debug.Print e.tagName & ":" & e.Style.FontSize
'    Next
    
    
    End Sub




'FIXIT: Declare 'pDisp' and 'URL' and 'Frame' and 'StatusCode' with an early-bound data type     FixIT90210ae-R1672-R1B8ZE
Private Sub IEView_NavigateError(ByVal pDisp As Object, URL As Variant, Frame As Variant, StatusCode As Variant, Cancel As Boolean)

    LeftFrame.Enabled = True

    If bAutoShowNow Then Timer.Enabled = True

End Sub

Private Sub IEView_ProgressChange(ByVal Progress As Long, ByVal ProgressMax As Long)
    DisplayProgress Progress, ProgressMax
End Sub
Public Sub DisplayProgress(ByRef iCur As Long, ByRef iMax As Long)
'▓
'█
'Const ps = "▓"
Dim newtext As String
Const sts_max As Integer = 28
newtext = String$(sts_max, "-")

If iMax <= 0 Then
    StsBar.Panels(2).text = ""
ElseIf iCur < 0 Then
    'Mid(newtext, 1, 1) = ">"
    StsBar.Panels(2).text = ">"
Else
    'Mid(newtext, Int(iCur / iMax * (sts_max - 1)) + 1, 1) = ">"
    StsBar.Panels(2).text = String$(Int(iCur / iMax * (sts_max - 1)) + 1, "-") & ">"
End If

End Sub
'Private Sub IEView_ProgressChange(Index As Integer, ByVal Progress As Long, ByVal ProgressMax As Long)
' Call
'End Sub
Private Sub IEView_StatusTextChange(ByVal text As String)

    'If text = "" Then Exit Sub
    Call eStatusTextChange(text, IEView)

End Sub

'FIXIT: Declare 'PostData' with an early-bound data type                                   FixIT90210ae-R1672-R1B8ZE
Private Sub ieViewV1_NewWindow(ByVal URL As String, ByVal flags As Long, ByVal targetFrameName As String, PostData As Variant, ByVal Headers As String, Processed As Boolean)

    Processed = True
  IEView.Navigate2 URL ', , targetFrameName, PostData, Headers

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
    
    indexList = LeftStrip.SelectedItem.Index
    zhrStatus.iListIndex = indexList
    List(indexList).ZOrder 0
    
    If isListLoaded(List(indexList)) <> lstLoaded Then
        readyToLoadList (indexList)
    End If
    
    Call selectListItem
    
End Sub

Private Sub List_Collapse(Index As Integer, ByVal Node As ComctlLib.Node)

    Node.Image = 1

End Sub

Private Sub List_Expand(Index As Integer, ByVal Node As ComctlLib.Node)

    Node.Image = 2

End Sub

Private Sub list_NodeClick(Index As Integer, ByVal Node As ComctlLib.Node)

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
    Dim i As Long, pos As Integer
    
  On Error GoTo CatalogError

    For i = 1 To lCount
        
        'StsBar.Panels(1).text = "Loading: " & (i + 1) & "/" & lCount & " (" & Format$((i + 1) / lCount, "00%") & ")"
        DisplayProgress i, lCount
        thename = LContent(1, i)
        Ktag = LContent(2, i)
        kKey = "ZTM" + LContent(1, i)

        If Right$(thename, 1) = "/" Then thename = Left$(thename, Len(thename) - 1)
        pos = InStrRev(thename, "/")

        If pos = 0 Then
            Ktext = thename
            If Right$(LContent(1, i), 2) = ":/" Then
                Kimageindex = 1
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
    DisplayProgress -1, -1
    thelist.Visible = True
    
End Sub

Private Sub saveReadingStatus()
    If zhrStatus.sCur_zhFile <> "" And zhrStatus.sCur_zhSubFile <> "" Then
        Dim nowAt As ReadingStatus
        With nowAt
            .page = zhrStatus.sCur_zhSubFile
            .perOfScrollTop = IEView.Document.body.scrollTop / IEView.Document.body.scrollHeight
            .perOfScrollLeft = IEView.Document.body.scrollLeft / IEView.Document.body.scrollWidth
        End With
        rememberBook linvblib.BuildPath(sConfigDir, zhMemFile), zhrStatus.sCur_zhFile, nowAt
    End If
End Sub

Public Sub loadzh(ByVal thisfile As String, Optional ByVal firstfile As String = "", Optional Reloadit As Boolean = False)
    
       
    If linvblib.PathExists(thisfile) = False Then
        PopupMessage "File not exist: " & thisfile
        Exit Sub
    End If
    If linvblib.PathType(thisfile) <> LNFile Then
        PopupMessage "Invaild filename: " & thisfile
        Exit Sub
    End If
    
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
            ElseIf zhInfo.sDefaultfile <> "" Then
                GetView zhInfo.sDefaultfile
                Exit Sub
            End If

        End If

    End With


    Call saveReadingStatus
    Call zhReaderReset

    If thisfile <> zhrStatus.sCur_zhFile Then
        sTempZH = linvblib.GetBaseName(linvblib.GetTempFilename)
        sTempZH = linvblib.BuildPath(Tempdir, sTempZH)

        Do Until linvblib.PathExists(sTempZH) = False
            sTempZH = linvblib.GetBaseName(linvblib.GetTempFilename)
            sTempZH = linvblib.BuildPath(Tempdir, sTempZH)
        Loop

        MkDir sTempZH
        sTempZH = toUnixPath(sTempZH)
        zhrStatus.sCur_zhFile = thisfile
    End If

    zhInfo.selfReset
    zhInfo.parseZhCommentText getZhCommentText(thisfile)
    getZIPContent zhrStatus.sCur_zhFile
    getZHContent zhInfo
    
    'loadZHList List(lwContent), zhInfo


    If zhInfo.lContentCount > 0 Then
        zhrStatus.iListIndex = lwContent
        List(zhrStatus.iListIndex).ZOrder 0
    Else
        zhrStatus.iListIndex = lwFiles
        List(zhrStatus.iListIndex).ZOrder 0
    End If

   'StsBar.Panels(2).text = LiNVBLib.GetFileName(thisfile)
    
    Dim hMRU As New CMenuArrHandle
    hMRU.Menus = mnuFile_Recent
'FIXIT: mnuFile_Recent(0).Tag property has no Visual Basic .NET equivalent and will not be upgraded.     FixIT90210ae-R7593-R67265
    hMRU.maxItem = Val(mnuFile_Recent(0).Tag)
    hMRU.AddUnique toDosPath(thisfile) ', thisfile
    Set hMRU = Nothing
    
    NotResize = True

    With zhInfo

        If .zvShowLeft = zhtmVisiableTrue Then
            ShowMenu True
        ElseIf .zvShowLeft = zhtmVisiableFalse Then
            ShowLeft False
        End If

        If .zvShowMenu = zhtmVisiableTrue Then
            ShowMenu True
        ElseIf .zvShowMenu = zhtmVisiableFalse Then
            ShowMenu False
        End If

        If .zvShowStatusBar = zhtmVisiableTrue Then
            ShowStatusBar True
        ElseIf .zvShowStatusBar = zhtmVisiableFalse Then
            ShowStatusBar False
        End If

    End With

    If mnuView_Left.Checked Then
        LeftStrip.Enabled = False
        LeftStrip.Tabs(1).Selected = False
        LeftStrip.Tabs(2).Selected = False
        LeftStrip.Enabled = True
        LeftStrip.Tabs(zhrStatus.iListIndex).Selected = True
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
        ElseIf zhInfo.sDefaultfile <> "" Then
            GetView zhInfo.sDefaultfile
        End If
    End If

End Sub



Private Sub readyToLoadList(ByRef listIndex As ListWhat)

    Dim zipContent() As String
    Dim lzipCount As Long
    Dim i As Long

    'lzipCount = lFoldersIZcount + sfilesinzip.count
    
    If listIndex = lwContent Then
        lzipCount = sFilesINContent(1).Count
    ElseIf listIndex = lwFiles Then
        lzipCount = sFilesInZip.Count
    Else
        Exit Sub
    End If

    If lzipCount > lcstFittedListItemsNum Then
        If PopupMessage("File list contains " & lzipCount & " items, It maybe hang the application, Countinue?", "File list", vbYesNo) = vbNo Then Exit Sub
    End If

'FIXIT: Non Zero lowerbound arrays are not supported in Visual Basic .NET                  FixIT90210ae-R9815-H1984
    ReDim zipContent(1 To 2, 1 To lzipCount)
'    lEnd = lFoldersIZcount - 1
'
'    For i = 0 To lEnd
'        zipContent(0, i) = sFoldersInZip.Item(i + 1)
'        zipContent(1, i) = zipContent(0, i)
'    Next

    If listIndex = lwFiles Then
        For i = 1 To lzipCount
            zipContent(1, i) = sFilesInZip.Item(i - 1)
            zipContent(2, i) = zipContent(1, i)
        Next
    Else
        For i = 1 To lzipCount
            zipContent(1, i) = sFilesINContent(1).Item(i - 1)
            zipContent(2, i) = sFilesINContent(2).Item(i - 1)
        Next
    End If
    'Set sFoldersInZip = Nothing
    '    zhrStatus.iListIndex = lwFiles
    '    trvwlist.ZOrder 0
    Loadlist List(listIndex), zipContent(), lzipCount

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
    If Index = 4 Then
        If zhrStatus.sCur_zhFile = "" Then
            mnuDir_delete.Enabled = False
        Else
            mnuDir_delete.Enabled = True
        End If
        Exit Sub
    ElseIf Index = 7 Then
        mnuGo_Previous_Click
        Exit Sub
    ElseIf Index = 8 Then
        mnuGo_Next_Click
        Exit Sub
    End If

    If zhrStatus.sCur_zhFile = "" Then
        mnuFile_Close.Enabled = False
        mnuEdit_EditInfo.Enabled = False
        mnuHelp_BookInfo.Enabled = False
    Else
        mnuFile_Close.Enabled = True
        mnuEdit_EditInfo.Enabled = True
        mnuHelp_BookInfo.Enabled = True
    End If

    If zhrStatus.sCur_zhSubFile = "" Then
        mnuEdit_EditCurPage.Enabled = False
        mnuEdit_SetDefault.Enabled = False
    Else
        mnuEdit_EditCurPage.Enabled = True
        mnuEdit_SetDefault.Enabled = True
    End If

End Sub



Public Sub mnuBookmark_add_Click()

    'Dim i As Integer
    On Error Resume Next
    Dim sCaption As String
    If zhrStatus.sCur_zhFile = "" Then Exit Sub
    sCaption = MainFrm.IEView.Document.Title
    If sCaption = "" Then sCaption = linvblib.GetBaseName(zhrStatus.sCur_zhFile)
    
    Dim hMNU As CMenuArrHandle
    Set hMNU = New CMenuArrHandle
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
'FIXIT: mnuBookmark(i).Tag property has no Visual Basic .NET equivalent and will not be upgraded.     FixIT90210ae-R7593-R67265
    pos = InStr(mnuBookmark(i).Tag, "|")

    If pos > 0 Then
'FIXIT: mnuBookmark(i).Tag property has no Visual Basic .NET equivalent and will not be upgraded.     FixIT90210ae-R7593-R67265
        sBMZhfile = Left$(mnuBookmark(i).Tag, pos - 1)
'FIXIT: mnuBookmark(i).Tag property has no Visual Basic .NET equivalent and will not be upgraded.     FixIT90210ae-R7593-R67265
'FIXIT: mnuBookmark(i).Tag property has no Visual Basic .NET equivalent and will not be upgraded.     FixIT90210ae-R7593-R67265
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
'FIXIT: The use of 'vbOK' is not valid for the property being assigned.                    FixIT90210ae-R5049-R62328
        If msgConfirm = vbOK Then
            mnuDir_readNext_Click
            Kill sBackUp
            If zhrStatus.sCur_zhFile = sBackUp Then
                mnufile_Close_Click
            End If
        End If
    End If
End Sub

'Private Sub mnuDir_label_Click()
'    Call mnufile_Open_Click
'End Sub

Private Sub mnuDir_random_Click()
    Dim sCur As String
    Dim sPath As String
    sPath = zhrStatus.sCur_zhFile
    If sPath = "" Then sPath = CurDir$
    sCur = linvblib.GFileSystem.LookFor(sPath, LN_FILE_RAND, "*.zhtm")
    If sCur <> "" Then MainFrm.loadzh sCur
End Sub

Private Sub mnuDir_readNext_Click()
    Dim sCur As String
    Dim sPath As String
    sPath = zhrStatus.sCur_zhFile
    If sPath = "" Then sPath = CurDir$
    sCur = linvblib.GFileSystem.LookFor(sPath, LN_FILE_next, "*.zhtm")
    If sCur <> "" Then MainFrm.loadzh sCur
End Sub

Private Sub mnuDir_readPrev_Click()
    
    Dim sCur As String
    Dim sPath As String
    sPath = zhrStatus.sCur_zhFile
    If sPath = "" Then sPath = CurDir$
    sCur = linvblib.GFileSystem.LookFor(sPath, LN_FILE_prev, "*.zhtm")
    If sCur <> "" Then MainFrm.loadzh sCur
        
    
End Sub

Private Sub mnuEdit_Delete_Click()

    Dim askConfirm As VbMsgBoxResult
    Dim fileToDelete As String
    fileToDelete = zhrStatus.sCur_zhSubFile

    If fileToDelete = "" Then Exit Sub
    askConfirm = MsgBox(StrLocalize("Delete") & " " & fileToDelete & "?", vbOKCancel, StrLocalize("Confirm"))

    If askConfirm <> vbOK Then Exit Sub
    StsBar.Panels("ie").text = StrLocalize("Deleting ") & fileToDelete & " ..."
    Set lZip = New cZip

    With lZip
        .ZipFile = zhrStatus.sCur_zhFile
        .AddFileToProcess zhrStatus.sCur_zhSubFile
    End With

    lZip.Delete
    StsBar.Panels("ie").text = StrLocalize("Deleting ") & fileToDelete & " Done!"
    zhrStatus.sCur_zhSubFile = sFilesInZip.Item(curFileIndex(zhrStatus.sCur_zhSubFile) + 1)
    loadzh zhrStatus.sCur_zhFile, zhrStatus.sCur_zhSubFile, True

End Sub

Public Sub mnuEdit_Editcurpage_Click()

    Dim sShellTextEditor As String
    If zhrStatus.sCur_zhSubFile = "" Then Exit Sub
'FIXIT: mnuEdit_SelectEditor.Tag property has no Visual Basic .NET equivalent and will not be upgraded.     FixIT90210ae-R7593-R67265
    sShellTextEditor = mnuEdit_SelectEditor.Tag
    If sShellTextEditor = "" Then
        mnuEdit_SelectEditor_Click
'FIXIT: mnuEdit_SelectEditor.Tag property has no Visual Basic .NET equivalent and will not be upgraded.     FixIT90210ae-R7593-R67265
        sShellTextEditor = mnuEdit_SelectEditor.Tag
    End If
    If sShellTextEditor = "" Then Exit Sub
    
    Dim fso As New GFileSystem
    Dim sTmpFile As String
    '    If bIsZhtm Then
    sTmpFile = fso.BuildPath(sTempZH, zhrStatus.sCur_zhSubFile)

    If fso.PathExists(sTmpFile) = False Then
        myXUnzip zhrStatus.sCur_zhFile, zhrStatus.sCur_zhSubFile, sTempZH, zhrStatus.sPWD
    End If

    If fso.PathExists(sTmpFile) = False Then Exit Sub
    '    Else
    '        sTmpFile = fso.BuildPath(zhrStatus.sCur_zhFile, zhrStatus.sCur_zhSubFile)
    '    End If
    ShellAndClose sShellTextEditor & " " & Chr$(34) & sTmpFile & Chr$(34), vbNormalFocus
    '    If bIsZhtm Then
    myXZip zhrStatus.sCur_zhFile, zhrStatus.sCur_zhSubFile, zhrStatus.sPWD, sTempZH
    '    End If
    IEView.Refresh2

End Sub

Public Sub mnuEdit_EditInfo_Click()

    Dim sShellTextEditor As String


    If zhrStatus.sCur_zhSubFile = "" Then Exit Sub
    
'FIXIT: mnuEdit_SelectEditor.Tag property has no Visual Basic .NET equivalent and will not be upgraded.     FixIT90210ae-R7593-R67265
    sShellTextEditor = mnuEdit_SelectEditor.Tag 'hINI.GetSetting("ReaderStyle", "TextEditor")

    If sShellTextEditor = "" Then
        mnuEdit_SelectEditor_Click
'FIXIT: mnuEdit_SelectEditor.Tag property has no Visual Basic .NET equivalent and will not be upgraded.     FixIT90210ae-R7593-R67265
        sShellTextEditor = mnuEdit_SelectEditor.Tag
    End If

    If sShellTextEditor = "" Then Exit Sub
    Dim fso As New GFileSystem
    Dim sTmpFile As String
    Dim fNum As Integer
    Dim stmp As String
    Dim sBackUp As String
    sTmpFile = fso.BuildPath(sTempZH, "zhInfo")
    On Error Resume Next

    If fso.PathExists(sTmpFile) Then Kill sTmpFile

    If fso.PathExists(sTmpFile) Then RmDir sTmpFile
    sBackUp = rdel(zhInfo.ToString)
    fNum = FreeFile
    Open sTmpFile For Output As #fNum
'FIXIT: Print method has no Visual Basic .NET equivalent and will not be upgraded.         FixIT90210ae-R7593-R67265
    Print #fNum, sBackUp ' getZhCommentText(zhrStatus.sCur_zhFile)
    Close #fNum
    ShellAndClose sShellTextEditor & " " & sTmpFile
    fNum = FreeFile
    Open sTmpFile For Binary As #fNum
    stmp = String$(LOF(fNum), " ")
    Get #fNum, , stmp
    Close #fNum
    Kill sTmpFile
    stmp = rdel(stmp)

    If stmp <> sBackUp Then
        zhInfo.parseZhCommentText stmp
        saveZhInfo
        loadzh zhrStatus.sCur_zhFile, zhrStatus.sCur_zhSubFile, True
    End If

End Sub

Public Sub mnuEdit_SelectEditor_Click()

    Dim fso As New GFileSystem
    Dim sShellTextEditor As String

    Dim sInitDir As String

'FIXIT: mnuEdit_SelectEditor.Tag property has no Visual Basic .NET equivalent and will not be upgraded.     FixIT90210ae-R7593-R67265
    sShellTextEditor = mnuEdit_SelectEditor.Tag 'hINI.GetSetting("ReaderStyle", "TextEditor")

    If sShellTextEditor <> "" Then
        sInitDir = fso.GetParentFolderName(sShellTextEditor)
    End If

    Set fso = Nothing
    Dim fResult As Boolean
    Dim cDLG As New CCommonDialogLite
    fResult = cDLG.VBGetOpenFileName( _
       FileName:=sShellTextEditor, _
       Filter:="EXE File|*.exe|All Files|*.*", _
       InitDir:=sInitDir, _
       DlgTitle:=mnuEdit_SelectEditor.Caption, _
       Owner:=Me.hwnd)
    Set cDLG = Nothing

    If fResult Then

        If sShellTextEditor <> "" Then
'FIXIT: mnuEdit_SelectEditor.Tag property has no Visual Basic .NET equivalent and will not be upgraded.     FixIT90210ae-R7593-R67265
            mnuEdit_SelectEditor.Tag = sShellTextEditor
        End If

    End If



End Sub

Public Sub mnuEdit_SetDefault_Click()

    '    If bIsZhtm = False Then Exit Sub
    'Dim sTmpFile As String
    'Dim fso As New scripting.FileSystemObject
    'Dim ts As scripting.TextStream

    If zhrStatus.sCur_zhFile <> "" And zhrStatus.sCur_zhSubFile <> "" Then
        zhInfo.sDefaultfile = zhrStatus.sCur_zhSubFile
        saveZhInfo
        'saveCommentToZipfile zhInfo.toString, zhrStatus.sCur_zhFile
    End If

End Sub

Public Sub mnufile_Close_Click()

    Dim fso As New FileSystemObject
    'Dim i As Integer
    'rememberNew zhrStatus.sCur_zhFile, zhrStatus.sCur_zhSubFile
    zhReaderReset
    On Error Resume Next

    If fso.FolderExists(sTempZH) Then fso.DeleteFolder sTempZH, False
    MainFrm.AppHtmlAbout
End Sub

Public Sub mnufile_exit_Click()

    Unload Me

End Sub

Public Sub mnufile_Open_Click()

    Dim thisfile As String
    Dim sInitDir As String
    Dim fso As New GFileSystem

    'rememberNew zhrStatus.sCur_zhFile, zhrStatus.sCur_zhSubFile

    If zhrStatus.sCur_zhFile <> "" Then
        If fso.PathExists(zhrStatus.sCur_zhFile) = True Then sInitDir = fso.GetParentFolderName(zhrStatus.sCur_zhFile)
    Else
'FIXIT: mnuFile_Open.Tag property has no Visual Basic .NET equivalent and will not be upgraded.     FixIT90210ae-R7593-R67265
        sInitDir = mnuFile_Open.Tag
    End If
'FIXIT: mnuFile_Open.Tag property has no Visual Basic .NET equivalent and will not be upgraded.     FixIT90210ae-R7593-R67265
    mnuFile_Open.Tag = sInitDir
    
    Dim fResult As Boolean
    Dim cDLG As New CCommonDialogLite

    If fso.PathExists(sInitDir) Then sInitDir = linvblib.toDosPath(sInitDir)
    fResult = cDLG.VBGetOpenFileName( _
       FileName:=thisfile, _
       Filter:="Zippacked Html File|*.zhtm;*.zbook;*.zip|所有文件|*.*", _
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
'FIXIT: mnuFile_Recent(Index).Tag property has no Visual Basic .NET equivalent and will not be upgraded.     FixIT90210ae-R7593-R67265
    fname = mnuFile_Recent(Index).Tag
    If linvblib.FileExists(fname) = False Then
        Dim HM As New CMenuArrHandle
        HM.Menus = mnuFile_Recent
        HM.Remove Index
        Set HM = Nothing
        MsgBox "File not exist: " & fname, vbInformation, "Error..."
    Else
'FIXIT: mnuFile_Recent(Index).Tag property has no Visual Basic .NET equivalent and will not be upgraded.     FixIT90210ae-R7593-R67265
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

Public Sub mnuGo_Back_Click()

    On Error Resume Next
    Timer.Enabled = False
    IEView.GoBack

    If bAutoShowNow Then Timer.Enabled = True

End Sub

Public Sub mnuGo_Forward_Click()

    On Error Resume Next
    Timer.Enabled = False
    IEView.GoForward

    If bAutoShowNow Then Timer.Enabled = True

End Sub

Public Sub mnuGo_Home_Click()

    On Error GoTo 0
    
    If zhrStatus.sCur_zhFile = "" Or sFilesInZip.Count < 1 Then
        AppHtmlAbout
    ElseIf zhInfo.sDefaultfile <> "" Then
        GetView zhInfo.sDefaultfile
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
    

        
    If zhrStatus.iListIndex = lwContent Then

        If sFilesINContent(2).Count < 1 Then GoTo Herr

        If zhrStatus.sCur_zhSubFile = "" Then
            lcurPage = 0
        Else
            lcurPage = curFileIndex(zhrStatus.sCur_zhSubFile)
        End If

        If lcurPage >= sFilesINContent(2).Count Then lcurPage = 0
        GetView sFilesINContent(2).Item(lcurPage + 1)
    Else

        If sFilesInZip.Count < 1 Then GoTo Herr

        If zhrStatus.sCur_zhSubFile = "" Then
            lcurPage = 0
        Else
            lcurPage = curFileIndex(zhrStatus.sCur_zhSubFile)
        End If

        If lcurPage >= sFilesInZip.Count Then lcurPage = 0
        GetView sFilesInZip.Item(lcurPage + 1)
    End If

Herr:

    If bAutoShowNow Then Timer.Enabled = True

End Sub

Public Sub mnuGo_Previous_Click()

    On Error GoTo Herr
    Timer.Enabled = False
    Dim lcurPage As Long

    If zhrStatus.iListIndex = lwContent Then

        If sFilesINContent(2).Count < 1 Then GoTo Herr

        If zhrStatus.sCur_zhSubFile = "" Then
            lcurPage = 2
        Else
            lcurPage = curFileIndex(zhrStatus.sCur_zhSubFile)
        End If

        If lcurPage <= 1 Then lcurPage = sFilesINContent(2).Count + 1
        GetView sFilesINContent(2).Item(lcurPage - 1)
    Else

        If sFilesInZip.Count < 1 Then GoTo Herr

        If zhrStatus.sCur_zhSubFile = "" Then
            lcurPage = 2
        Else
            lcurPage = curFileIndex(zhrStatus.sCur_zhSubFile)
        End If

        If lcurPage <= 1 Then lcurPage = sFilesInZip.Count + 1
        GetView sFilesInZip.Item(lcurPage - 1)
    End If

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
'FIXIT: App.Revision property has no Visual Basic .NET equivalent and will not be upgraded.     FixIT90210ae-R7593-R67265
    sAbout = sAbout + Space$(4) + App.ProductName + " (Build" + Str$(App.Major) + "." + Str$(App.Minor) + "." + Str$(App.Revision) + ")" + vbCrLf
    sAbout = sAbout + Space$(4) + App.LegalCopyright
    MsgBox sAbout, vbInformation, "About"

End Sub

Public Sub mnuHelp_BookInfo_Click()

    Dim sAbout As String

    If zhrStatus.sCur_zhFile <> "" Then
        sAbout = sAbout + Space$(4) + "Title:" + zhInfo.sTitle + vbCrLf
        sAbout = sAbout + Space$(4) + "Author:" + zhInfo.sAuthor + vbCrLf
        sAbout = sAbout + Space$(4) + "Catalog:" + zhInfo.sCatalog + vbCrLf
        sAbout = sAbout + Space$(4) + "Publisher:" + zhInfo.sPublisher + vbCrLf
        sAbout = sAbout + Space$(4) + "Date:" + zhInfo.sDate + vbCrLf
        MsgBox sAbout, vbInformation, "BookInfo of [" & zhrStatus.sCur_zhFile & "]"
    End If

End Sub

Public Sub mnuView_AddressBar_Click()

    If mnuView_AddressBar.Checked Then
        mnuView_AddressBar.Checked = False
        cmbAddress.Visible = False '.Top = -cmbAddress.Height  '= 0
    Else
        mnuView_AddressBar.Checked = True
        cmbAddress.Visible = True
    End If

    Form_Resize

End Sub


Private Sub mnuView_ApplyStyleSheet_List_Click(Index As Integer)

    On Error Resume Next
    
    Dim iLBound As Long
    Dim iUBound As Long
    Dim iFor As Long
    iLBound = mnuView_ApplyStyleSheet_List.LBound
    iUBound = mnuView_ApplyStyleSheet_List.UBound
    For iFor = iLBound To Index - 1 'iUBound
        mnuView_ApplyStyleSheet_List(iFor).Checked = False
    Next
    For iFor = Index + 1 To iUBound
        mnuView_ApplyStyleSheet_List(iFor).Checked = False
    Next
    
    If mnuView_ApplyStyleSheet_List(Index).Checked Then
        mnuView_ApplyStyleSheet_List(Index).Checked = False
'FIXIT: mnuView_ApplyStyleSheet.Tag property has no Visual Basic .NET equivalent and will not be upgraded.     FixIT90210ae-R7593-R67265
        mnuView_ApplyStyleSheet.Tag = 0
        mnuView_ApplyStyleSheet_List(0).Checked = True
        Call ApplyDefaultStyle(False)
    Else
        mnuView_ApplyStyleSheet_List(Index).Checked = True
'FIXIT: mnuView_ApplyStyleSheet.Tag property has no Visual Basic .NET equivalent and will not be upgraded.     FixIT90210ae-R7593-R67265
        mnuView_ApplyStyleSheet.Tag = Index
        Call ApplyDefaultStyle(True)
    End If

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
    sFilesToProcess = CleanZipFilename(sFilesToProcess)

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
    sFilesToProcess = CleanZipFilename(sFilesToProcess)
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

Public Sub saveZhInfo()

    Dim sComment As String
    Dim sContent As String
    Dim sTmpFile As String
    Dim fNum As Integer

    If zhrStatus.sCur_zhFile = "" Then Exit Sub

    If linvblib.PathExists(zhrStatus.sCur_zhFile) = False Then Exit Sub
    StsBar.Panels("ie").text = "Saving Info of " & zhrStatus.sCur_zhFile & " ..."

    If zhInfo.lContentCount > 0 Then
        sContent = zhInfo.ContentText
        sTmpFile = linvblib.bdUnixDir(sTempZH, zhCommentFileName)
        On Error Resume Next

        If linvblib.PathExists(sTmpFile) Then Kill sTmpFile
        fNum = FreeFile
        Open sTmpFile For Binary Access Write As #fNum
        Put #fNum, , sContent
        Close #fNum
        myXZip zhrStatus.sCur_zhFile, sTmpFile, zhrStatus.sPWD, sTempZH
        zhInfo.sContentFile = zhCommentFileName
    ElseIf zhInfo.sContentFile <> "" Then
        Set lZip = New cZip

        With lZip
            .ZipFile = zhrStatus.sCur_zhFile
            .AddFileToProcess zhInfo.sContentFile
        End With

        lZip.Delete
        Set lZip = Nothing
    End If

    sComment = zhInfo.InfoText
    saveCommentToZipfile sComment, zhrStatus.sCur_zhFile
    StsBar.Panels("ie").text = "Info of " & zhrStatus.sCur_zhFile & " saved."

End Sub

Public Sub selectListItem()

    If zhrStatus.sCur_zhSubFile = "" Then Exit Sub
    Dim fIndex As Long
    Dim fcount As Long
    Dim fKey As String
    
    On Error Resume Next
    
    
    
     If zhrStatus.iListIndex = lwContent Then
        fIndex = sFilesINContent(2).Find(zhrStatus.sCur_zhSubFile)
        fcount = sFilesINContent(2).Count
        fKey = "ZTM" & sFilesINContent(1).Item(fIndex)
    Else
        fIndex = sFilesInZip.Find(zhrStatus.sCur_zhSubFile)
        fcount = sFilesInZip.Count
        fKey = "ZTM" & sFilesInZip.Item(fIndex)
    End If
    StsBar.Panels("order").text = fIndex & "\" & fcount
    List(zhrStatus.iListIndex).Nodes(fKey).Selected = True

End Sub

Private Sub ShowLeft(showit As Boolean)

    If showit Then
        mnuView_Left.Checked = True
        'zhrStatus.bLeftShowed = True
        If zhrStatus.iListIndex > 0 Then
            LeftStrip.Tabs(zhrStatus.iListIndex).Selected = True
        End If
        'If zhrStatus.sCur_zhFile = "" Then Form_Resize: Exit Sub
        '
        '    If zhrStatus.iListIndex = lwContent Then
        '
        '        List(zhrStatus.iListIndex).ZOrder 0
        '        If isListLoaded(List(lwContent)) <> lstloaded Or bReloadContent Then
        '        setListStatus List(zhrStatus.iListIndex), lstloaded
        '        loadZHList List(zhrStatus.iListIndex), zhInfo
        '        bReloadContent = False
        '        End If
        '
        '    ElseIf zhrStatus.iListIndex = lwFiles Then
        '
        '        If List.count = 1 Then
        '            Load List(lwFiles)
        '            List(lwFiles).Tag = ""
        '        End If
        '        List(zhrStatus.iListIndex).Visible = True
        '        List(zhrStatus.iListIndex).ZOrder 0
        '        If isListLoaded(List(lwFiles)) <> lstloaded Or bReloadContent Then
        '        loadZIPContent List(zhrStatus.iListIndex), zhrStatus.sCur_zhFile
        '        setListStatus List(zhrStatus.iListIndex), lstloaded
        '        End If
        '
        '    End If
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

    zhInfo.selfReset
    Dim i As Integer
    bInValidPassword = False
    bAutoShowNow = False
    bRandomShow = False
    mnuGo_AutoNext.Checked = False
    mnuGo_AutoRandom.Checked = False
    Timer.Enabled = False

    For i = 1 To List.Count
        List(i).Visible = False
        List(i).Nodes.Clear
        List(i).Tag = ""
        List(i).Visible = True
    Next

    'setListStatus List(0), lstNotloaded
    'zhInfo.selfReset

    With zhrStatus
        'If bIsZhtm Then .iListIndex = lwContent Else .iListIndex = lwFiles
        .sCur_zhFile = ""
        .sCur_zhSubFile = ""
    End With

    's_AI_DefaultFile = ""
    Set sFilesInZip = New CStringArray   'CStringCollection
    sFilesInZip.ChunkSize = 500
'    Set sFoldersInZip = New CStringArray ' CStringCollection
'    lFoldersIZcount = 0
    Set sFilesINContent(1) = New CStringArray ' CStringCollection
    Set sFilesINContent(2) = New CStringArray ' CStringCollection
    'LeftStrip.Tabs(1).Selected = True
    'Navigated = False
    'appHtmlAbout
    '    Do
    '        DoEvents
    '    Loop While IEView.ReadyState = READYSTATE_LOADING

End Sub

Public Sub AddUniqueItem(cmbBoxToAdd As ComboBox, sItem As String)

    On Error GoTo 0
    'Dim txtCmb As String
    'txtCmb = cmbBoxToAdd.text
    Dim iIndex As Long
    Dim iEnd As Long
    iEnd = cmbBoxToAdd.ListCount - 1

    For iIndex = 0 To iEnd
        If StrComp(cmbBoxToAdd.List(iIndex), sItem, vbTextCompare) = 0 Then Exit Sub
    Next

    cmbBoxToAdd.AddItem sItem

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


Public Function PopupMessage(ByRef vMessage As String, Optional vTitle As String = "", Optional vStyle As VbMsgBoxStyle = vbOKOnly) As VbMsgBoxResult
    PopupMessage = MsgBox(vMessage, vStyle, vTitle)
End Function


Public Sub WriteMessage(ByRef vMessage As String, Optional HTMLPrefix As String = "", Optional HTMLSuffix As String = "")
    Dim doc As IHTMLDocument
    Set doc = IEView.Document
    On Error Resume Next
    doc.body.innerHTML = ""
    doc.write HTMLPrefix & vMessage & HTMLSuffix & "<BR>"
End Sub

Public Sub AppendMessage(ByRef vMessage As String, Optional HTMLPrefix As String = "", Optional HTMLSuffix As String = "")
    Dim doc As IHTMLDocument
    Set doc = IEView.Document
    On Error Resume Next
    doc.write HTMLPrefix & vMessage & HTMLSuffix & "<BR>"
End Sub

