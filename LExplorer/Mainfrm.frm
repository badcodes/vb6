VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form MainFrm 
   ClientHeight    =   3408
   ClientLeft      =   132
   ClientTop       =   744
   ClientWidth     =   7056
   Icon            =   "Mainfrm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3408
   ScaleWidth      =   7056
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picSplitter 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillColor       =   &H00808080&
      Height          =   4800
      Left            =   4776
      ScaleHeight     =   2090.126
      ScaleMode       =   0  'User
      ScaleWidth      =   390
      TabIndex        =   1
      Top             =   -1692
      Visible         =   0   'False
      Width           =   40
   End
   Begin VB.Frame theShow 
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   2112
      Left            =   3072
      TabIndex        =   2
      Top             =   156
      Width           =   3384
      Begin SHDocVwCtl.WebBrowser IEView 
         Height          =   1536
         Left            =   204
         TabIndex        =   0
         Top             =   288
         Width           =   2904
         ExtentX         =   5122
         ExtentY         =   2709
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
         Location        =   "http:///"
      End
   End
   Begin MSComDlg.CommonDialog cDlg 
      Left            =   5592
      Top             =   2436
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox LeftFrame 
      Height          =   2868
      Left            =   360
      ScaleHeight     =   2820
      ScaleWidth      =   2592
      TabIndex        =   3
      Top             =   24
      Width           =   2640
      Begin VB.DirListBox DirList 
         CausesValidation=   0   'False
         Height          =   1584
         Left            =   48
         TabIndex        =   7
         Top             =   924
         Width           =   2316
      End
      Begin VB.ComboBox cmbFilter 
         Appearance      =   0  'Flat
         Height          =   288
         Left            =   24
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   540
         Width           =   2352
      End
      Begin VB.DriveListBox DriveList 
         Height          =   288
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   2415
      End
   End
   Begin VB.Image imgSplitter 
      Appearance      =   0  'Flat
      Height          =   5172
      Left            =   5364
      MousePointer    =   9  'Size W E
      Top             =   -2040
      Width           =   20
   End
   Begin VB.Label stsBar 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   288
      Left            =   12
      TabIndex        =   6
      Top             =   3120
      Width           =   5052
   End
   Begin VB.Menu mnu 
      Caption         =   "&File"
      Index           =   0
      Begin VB.Menu mnuFile_Open 
         Caption         =   "&Open File"
      End
      Begin VB.Menu mnuFile_OpenDir 
         Caption         =   "Open &Directory"
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
      Visible         =   0   'False
      Begin VB.Menu mnuEdit_EditCurPage 
         Caption         =   "&Edit Current Page"
      End
      Begin VB.Menu mnuEdit_EditInfo 
         Caption         =   "Edit zhFile &Info"
      End
      Begin VB.Menu mnuSep2 
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
         Shortcut        =   {F8}
      End
      Begin VB.Menu mnuView_StatusBar 
         Caption         =   "&StatusBar"
         Shortcut        =   {F9}
      End
      Begin VB.Menu mnuView_FullScreen 
         Caption         =   "&FullScreen"
         Shortcut        =   {F11}
      End
   End
   Begin VB.Menu mnu 
      Caption         =   "&Go"
      Index           =   3
      Begin VB.Menu mnuGo_Back 
         Caption         =   "&Back         Alt+Left"
      End
      Begin VB.Menu mnuGo_Forward 
         Caption         =   "&Forward     Alt+Right"
      End
      Begin VB.Menu mnuGo_Previous 
         Caption         =   "&Previous    Alt+Down"
      End
      Begin VB.Menu mnuGo_Next 
         Caption         =   "&Next         Alt+Up"
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGo_Home 
         Caption         =   "&Home       Alt+Home"
      End
   End
   Begin VB.Menu mnu 
      Caption         =   "&Bookmark"
      Index           =   4
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
      Caption         =   "Fi&lter"
      Index           =   5
      Begin VB.Menu mnuFilter_Add 
         Caption         =   "&Add"
         Index           =   5
      End
      Begin VB.Menu mnuFilter_Delete 
         Caption         =   "Delete"
      End
   End
   Begin VB.Menu mnu 
      Caption         =   "&Help"
      Index           =   6
      Begin VB.Menu mnuHelp_BookInfo 
         Caption         =   "&Book Info"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuHelp_About 
         Caption         =   "&About This"
      End
   End
   Begin VB.Menu mnuIe 
      Caption         =   "&IE"
      Visible         =   0   'False
      Begin VB.Menu mnuIe_Backward 
         Caption         =   "&Backward"
      End
      Begin VB.Menu mnuIe_Forward 
         Caption         =   "F&orward"
      End
      Begin VB.Menu mnuSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuIe_copy 
         Caption         =   "&Copy"
      End
      Begin VB.Menu mnuIe_SelectAll 
         Caption         =   "Select &All"
      End
      Begin VB.Menu mnuSep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuIe_AddBookmark 
         Caption         =   "Add Book&mark"
      End
      Begin VB.Menu mnuIe_ViewSource 
         Caption         =   "&View Source"
      End
      Begin VB.Menu mnuSep6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuIe_Print 
         Caption         =   "P&rint"
      End
      Begin VB.Menu mnuIe_refresh 
         Caption         =   "&Refresh"
      End
      Begin VB.Menu mnuSep7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuIe_property 
         Caption         =   "&Property"
      End
   End
End
Attribute VB_Name = "MainFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Use XCEEDZIP.tag To save ZIPComment

Enum ListStatusConstant
lstloaded = 1
lstNotloaded = 2
lstNotExists = 0
End Enum

'Gobal Var
Private Const sTaglstLoaded = "ListLoaded"
Private Const staglstNotLoaded = "ListNotLoaded"

Private Const sglSplitLimit = 500
Private Const minFormHeight = 2000
Private Const minFormWidth = 3000

Private Const icstDefaultFormHeight = 6000
Private Const icstDefaultFormWidth = 8000
Private Const icstDefaultLeftWidth = 1500

Private mbMoving As Boolean
Private Tempdir As String
Public sTempZH As String
'Private InvalidPWD As Boolean
Public Navigated As Boolean
Public NotPreOperate As Boolean
'Private StopNow As Boolean
Private NotResize As Boolean
'Private bReloadContent As Boolean
Public zhtmIni As String

Public iCurIEView As Integer
Public bIsZhtm As Boolean
Public bFullScreen As Boolean
Private WithEvents ieViewV1  As SHDocVwCtl.WebBrowser_V1
Attribute ieViewV1.VB_VarHelpID = -1
'Private WithEvents ieDocument As HTMLDocument
Public zhRecentFile As New CRecentFile
Private zhFileList As New CFiles
Private set_drive As Boolean
Private sMemoryFile As String

Public Sub cmbFilter_Click()
If cmbFilter.Enabled = False Then Exit Sub
loadzh zhrStatus.sCur_zhFile
End Sub

Public Sub DirList_Change()
If DirList.Enabled = False Then Exit Sub
loadzh DirList.Path
End Sub

'Private Function ieDocument_oncontextmenu() As Boolean
'
'If App.ProductName <> "zhReader" Then Exit Function
'ieDocument_oncontextmenu = False
'MainFrm.PopupMenu mnuIe
'ieDocument_oncontextmenu = True
'End Function

Public Sub IEView_StatusTextChange(ByVal text As String)

    MainFrm.stsBar.Caption = text
    'Call eStatusTextChange(text, IEView(Index))

End Sub

Public Sub ieViewV1_NewWindow(ByVal URL As String, ByVal flags As Long, ByVal TargetFrameName As String, PostData As Variant, ByVal Headers As String, Processed As Boolean)

    Processed = True
    IEView.Navigate2 URL

End Sub

Public Sub Form_Load()
    
    loadFormStr MainFrm
    Dim fso As New gCFileSystem
    Dim appPath As String
    
    appPath = fso.BuildPath(Environ("APPDATA"), App.ProductName)
    If fso.PathExists(appPath) = False Then MkDir appPath
    zhtmIni = fso.BuildPath(appPath, "config.ini")
    sMemoryFile = fso.BuildPath(appPath, "memory.encrypt")
    
    appHtmlAbout
    
    Set ieViewV1 = IEView.object
    
    NotResize = True
   
    On Error Resume Next
     Tempdir = fso.BuildPath(App.Path, "Cache")

    If fso.PathExists(Tempdir) Then RmDir Tempdir

    If fso.PathExists(Tempdir) = False Then MkDir Tempdir

    If fso.PathExists(fso.BuildPath(App.Path, cHtmlAboutFilename)) Then _
       Kill fso.BuildPath(App.Path, cHtmlAboutFilename)
       
    Dim theRS As ReaderStyle
    GetReaderStyle zhtmIni, theRS
   
    If theRS.WindowState = vbNormal Then

        With theRS.formPos

            If .Width = 0 Then .Width = icstDefaultFormWidth

            If .Height = 0 Then .Height = icstDefaultFormHeight
        End With

        With theRS.formPos
            MainFrm.Move .Left, .Top, .Width, .Height
        End With

    Else
        MainFrm.WindowState = vbMaximized
    End If

    If theRS.LeftWidth = 0 Then theRS.LeftWidth = icstDefaultLeftWidth
    imgSplitter.Left = theRS.LeftWidth
    ShowMenu theRS.ShowMenu
    ShowLeft theRS.ShowLeft
    ShowStatusBar theRS.ShowStatusBar
    '载入bookmark
    loadMNUBookmark
      
    cmbFilter.Enabled = False
    GetFileFilter zhtmIni, cmbFilter
    If cmbFilter.ListCount > 0 Then
        cmbFilter.ListIndex = 0
    End If
    cmbFilter.Enabled = True
    
    'If fso.FileExists(icofile(1)) And fso.FileExists(icofile(2)) And fso.FileExists(icofile(3)) Then
    '    Listimg.ListImages.Clear
    '    Listimg.ListImages.Add , , LoadPicture(icofile(1))
    '    Listimg.ListImages.Add , , LoadPicture(icofile(2))
    '    Listimg.ListImages.Add , , LoadPicture(icofile(3))
    ''    Listimg.ImageHeight = 16
    ''    Listimg.ImageWidth = 16
    '    List(0).ImageList = Listimg
    '    List(1).ImageList = Listimg
    'End If
    Dim thisfile As String
    thisfile = Command$

    If Left$(thisfile, "1") = Chr$(34) And Right$(thisfile, 1) = Chr$(34) And Len(thisfile) > 1 Then
        thisfile = Right$(thisfile, Len(thisfile) - 1)
        thisfile = Left$(thisfile, Len(thisfile) - 1)
    End If

    If thisfile = "" Then thisfile = theRS.LastPath
    loadzh thisfile

    NotResize = False
    
    If theRS.FullScreenMode Then mnuView_FullScreen_Click
    
    zhRecentFile.maxItem = Val(iniGetSetting(zhtmIni, "ViewStyle", "RecentMax"))
    zhRecentFile.maxCaptionLength = 30
    zhRecentFile.LoadFromIni zhtmIni
    zhRecentFile.FillinMenu mnuFile_Recent
    
'    If mnuFile_Recent.count = 1 Then
'        mnufile_s2.Visible = False
'    Else
'        mnufile_s2.Visible = True
'    End If

End Sub

Public Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

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
            mnuView_FullScreen_Click
        Case vbKeyF4
            mnufile_Close_Click
        'Case vbKeyF1
            'mnuHelp_BookInfo_Click
        Case Else
            '                KeyCode = iKeyCode
        End Select

    ElseIf Shift = vbAltMask Then

        Select Case iKeyCode
        Case vbKeyUp
            mnuGo_Previous_Click
        Case vbKeyDown
            mnuGo_Next_Click
 '       Case vbKeyHome
  '          mnuGo_Home_Click
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
        Case Else
            '               KeyCode = iKeyCode
        End Select

    ElseIf Shift = vbCtrlMask Then
        
        Select Case iKeyCode
        End Select

    End If

End Sub

Public Sub Form_KeyPress(KeyAscii As Integer)

    KeyAscii = 0

End Sub

Public Sub Form_Unload(Cancel As Integer)

    Dim fso As New Scripting.FileSystemObject
    Dim ListPath As String
    Dim Listfile As String
    Dim theRS As ReaderStyle

    With theRS.formPos
        .Height = MainFrm.Height
        .Width = MainFrm.Width
        .Top = MainFrm.Top
        .Left = MainFrm.Left
    End With

    With theRS
        .WindowState = MainFrm.WindowState
        .LeftWidth = LeftFrame.Width
    End With
        
    theRS.FullScreenMode = bFullScreen
    theRS.LastPath = fso.GetParentFolderName(zhrStatus.sCur_zhFile)
    theRS.ShowLeft = zhrStatus.bLeftShowed
    theRS.ShowMenu = zhrStatus.bMenuShowed
    theRS.ShowStatusBar = zhrStatus.bStatusBarShowed
    theRS.LastPath = zhrStatus.sCur_zhFile
    SaveReaderStyle zhtmIni, theRS
    rememberNew sMemoryFile, zhrStatus.sCur_zhFile, zhrStatus.sCur_zhSubFile
    zhRecentFile.SaveToIni zhtmIni
    SaveFileFilter zhtmIni, cmbFilter
    'StopNow = True
    'IEView.Stop
    'StopNow = False
    '
    'Navigated = False
    'IEView.Navigate2 "about:blank"
    'Do Until Navigated = True
    'DoEvents
    'Loop
    On Error Resume Next

    If fso.FolderExists(Tempdir) Then fso.DeleteFolder Tempdir, True

    If fso.FolderExists(sTempZH) Then fso.DeleteFolder (sTempZH), True
    'Call endUP

End Sub

Public Sub Form_Resize()
    
    Const ControlMargin = 16
    
    If NotResize Then Exit Sub
    Dim tempint As Long

    If MainFrm.WindowState = 1 Then Exit Sub

    With MainFrm
        If .ScaleHeight < minFormHeight Then .ScaleHeight = minFormHeight
        If .ScaleWidth < minFormWidth Then .ScaleWidth = minFormWidth
    End With

    With stsBar
        .Left = 0
        .Width = MainFrm.ScaleWidth
        .Top = MainFrm.ScaleHeight - .Height
    End With

    With imgSplitter
        .Top = 85
        .Height = stsBar.Top
    End With

    With LeftFrame
        .Left = ControlMargin
        .Top = ControlMargin
        .Height = Abs(stsBar.Top - ControlMargin * 2)
        .Width = Abs(imgSplitter.Left - ControlMargin * 2)
    End With
    
    '该死的这个控件会自己调整高度...使我加了九行代码
    With DirList
        tempint = Abs(LeftFrame.Height - 3 * ControlMargin - DriveList.Height - cmbFilter.Height)
        .Height = tempint
        tempint = tempint - .Height
        'If tempint < 0 Then tempint = 0
    End With
    
    With DriveList
        .Top = ControlMargin + tempint / 3
        .Left = ControlMargin
        .Width = Abs(LeftFrame.Width - ControlMargin * 3)
    End With
    
    With cmbFilter
        .Top = DriveList.Top + DriveList.Height + tempint / 3
        .Left = ControlMargin
        .Width = Abs(LeftFrame.Width - ControlMargin * 3)
    End With
    
    With DirList
        .Top = cmbFilter.Top + cmbFilter.Height + tempint / 3
        .Left = ControlMargin
        .Width = Abs(LeftFrame.Width - ControlMargin * 3)
    End With



    With theShow

        If zhrStatus.bLeftShowed = True Then
            .Left = imgSplitter.Left + imgSplitter.Width
            .Top = LeftFrame.Top
            tempint = MainFrm.ScaleWidth - .Left - ControlMargin

            If tempint < 0 Then
                theShow.Visible = False
            Else
                theShow.Visible = True
                .Width = tempint
            End If

            .Height = LeftFrame.Height
        Else
            .Left = ControlMargin
            .Top = LeftFrame.Top
            .Width = MainFrm.ScaleWidth - ControlMargin * 2
            .Height = LeftFrame.Height
        End If

    End With
    

    With IEView
        .Left = 0
        .Top = 0
        .Height = theShow.Height
        .Width = theShow.Width
    End With


    MainFrm.Refresh

End Sub

Public Sub IEView_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)

    If isZhcommand(URL) Then
        Cancel = True
        execZhcommand URL
        Exit Sub
    End If

'    If NotPreOperate Then
'        NotPreOperate = False
'        Exit Sub
'    End If

    
    'zhrStatus.sCur_zhSubFile = ""

    Dim fso As New gCFileSystem
    Dim sBaseDir As String
    Dim sLocalUrl As String
    Dim thefile As String
    Dim sUrl As String
    
    'Dim htmlfile As String
    
    



    If fso.PathExists(zhrStatus.sCur_zhFile) = False Then Exit Sub
    
    sUrl = URL
    sLocalUrl = LCase$(toUnixPath(sUrl))
    sBaseDir = LCase$(toUnixPath(MainFrm.sTempZH))
    
    If bddir(fso.GetParentFolderName(sLocalUrl)) = bddir(LCase(zhrStatus.sCur_zhFile)) Then
        thefile = fso.GetFileName(sUrl)
    
    ElseIf InStr(1, sLocalUrl, sBaseDir, vbTextCompare) = 1 Then

        If Left$(sLocalUrl, 5) = "file:" And Len(sLocalUrl) > 7 Then sLocalUrl = Right$(sUrl, Len(sUrl) - 8)
        thefile = Right$(sUrl, Len(sUrl) - Len(sBaseDir) - 1)
       
    End If
    
    If chkFileType(thefile) <> ftIE Then
        GetView thefile
        Cancel = True
        Exit Sub
    End If
    
    zhrStatus.sCur_zhSubFile = thefile
    If Right$(thefile, Len(TempHtm)) = TempHtm Then
     zhrStatus.sCur_zhSubFile = Replace(thefile, TempHtm, "")
    End If

    
'    Dim curIndex As Long
'    curIndex = getCurIndex
'    If curIndex > 0 Then List.Nodes(curIndex).Selected = True
    
    'zhRecentFile.add bddir(zhrStatus.sCur_zhFile) & zhrStatus.sCur_zhSubFile, bddir(zhrStatus.sCur_zhFile) & zhrStatus.sCur_zhSubFile
    'zhRecentFile.FillinMenu mnuFile_Recent
     
     
'    Call eBeforeNavigate(URL, Cancel, IEView)

End Sub

Public Sub IEView_DocumentComplete(ByVal pDisp As Object, URL As Variant)

    LeftFrame.Enabled = True


End Sub

Public Sub IEView_NavigateError(ByVal pDisp As Object, URL As Variant, Frame As Variant, StatusCode As Variant, Cancel As Boolean)

    LeftFrame.Enabled = True

End Sub



'Public Sub listFav_DblClick()
'If IsNull(listFav.SelectedItem) Then Exit Sub
'Dim tempstr As String
'Dim pos As Integer
'tempstr = favlist.locate(listFav.SelectedItem.Index)
'pos = InStr(tempstr, "|")
'If pos = 0 Then Exit Sub
'loadztm left$(tempstr, pos - 1), right$(tempstr, Len(tempstr) - pos)
'End Sub
'
'Public Sub listFav_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'If Button = 1 Then Exit Sub
'MainFrm.PopupMenu mnuFav, , x + ListFrame.Left, y + ListFrame.Top
'End Sub
Public Sub mnu_Click(Index As Integer)
    
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

    If bIsZhtm = False Then
        mnuEdit_EditInfo.Enabled = False
        mnuEdit_SetDefault.Enabled = False
        mnuHelp_BookInfo.Enabled = False
    Else
        mnuEdit_EditInfo.Enabled = True
        mnuEdit_SetDefault.Enabled = True
        mnuHelp_BookInfo.Enabled = True
    End If

End Sub

'Public Sub mnuEdit_SelectEditor_Click()
'
'    Dim fso As New gcFilesystem
'    Dim sShellTextEditor As String
'
'    sShellTextEditor = iniGetSetting(MainFrm.zhtmIni, "ReaderStyle", "TextEditor")
'
'    If sShellTextEditor <> "" Then
'        cDlg.InitDir = fso.GetParentFolderName(sShellTextEditor)
'        cDlg.FileName = sShellTextEditor
'    End If
'
'    cDlg.ShowOpen
'
'    sShellTextEditor = cDlg.FileName
'    If sShellTextEditor <> "" Then iniSaveSetting MainFrm.zhtmIni, "ReaderStyle", "TextEditor", sShellTextEditor
'
'End Sub
'
'Public Sub mnuEdit_SetDefault_Click()
'
'    If bIsZhtm = False Then Exit Sub
'    Dim sTmpFile As String
'    Dim fso As New Scripting.FileSystemObject
'    Dim ts As Scripting.TextStream
'
'    If zhrStatus.sCur_zhFile <> "" And zhrStatus.sCur_zhSubFile <> "" Then
'        sTmpFile = fso.BuildPath(sTempZH, fso.GetTempName)
'        zhInfo.sDefaultfile = zhrStatus.sCur_zhSubFile
'        zhInfo.saveZhCommentToFile sTmpFile
'        Set ts = fso.OpenTextFile(sTmpFile, ForReading)
'        xz(0).Tag = ts.ReadAll
'        ts.Close
'        fso.DeleteFile sTmpFile
'        cDoZipComment = dZCwriteComment
'        myXZip zhrStatus.sCur_zhFile, "", zhrStatus.sPWD, sTempZH
'        cDoZipComment = dZCNothing
'        xz(0).Tag = ""
'    End If
'
'End Sub

Public Sub mnuFile_OpenDir_Click()
    
    Dim sPath As String
    rememberNew sMemoryFile, zhrStatus.sCur_zhFile, zhrStatus.sCur_zhSubFile
    sPath = openDirDialog(Me.hwnd)

    If sPath <> "" Then loadzh sPath

End Sub

Public Sub mnuFile_Recent_Click(Index As Integer)

    rememberNew sMemoryFile, zhrStatus.sCur_zhFile, zhrStatus.sCur_zhSubFile
    loadzh mnuFile_Recent(Index).Tag
End Sub

Public Sub mnuFilter_Add_Click(Index As Integer)
Dim tempstr As String
tempstr = InputBox("输入文件名过滤字符串" & vbCrLf)
cmbFilter.AddItem tempstr

End Sub

Public Sub mnuFilter_Delete_Click()

    If cmbFilter.ListIndex < 0 Then Exit Sub
    
    Dim a As VbMsgBoxResult
    Dim theindex As Integer
    Dim thetext As String
    theindex = cmbFilter.ListIndex
    thetext = cmbFilter.text
    
    a = MsgBox("Delete " + thetext + " ?", vbOKCancel)
    
    If a = vbCancel Then Exit Sub
    cmbFilter.RemoveItem theindex
End Sub

Public Sub mnuGo_Back_Click()

    IEView.GoBack

End Sub

Public Sub mnuGo_Forward_Click()

    IEView.GoForward

End Sub

Public Sub mnuGo_Home_Click()
    viewIndex zhrStatus.sCur_zhFile
End Sub

'Public Sub mnuGo_Home_Click()
'
'    If zhrStatus.sCur_zhFile = "" Then
'        appHtmlAbout
'    Else
'        GetView zhInfo.sDefaultfile, IEView
'    End If
'
'End Sub

Public Sub mnuGo_Next_Click()

    Dim curIndex As String
    Dim nextFile As String
    
    curIndex = zhFileList.Index(zhrStatus.sCur_zhSubFile)
    nextFile = zhFileList.Item(curIndex + 1)
    If nextFile = "" Then nextFile = zhFileList.Item(0)
    If nextFile = "" Then Exit Sub
    
    GetView nextFile


End Sub

Public Sub mnuGo_Previous_Click()
    Dim curIndex As String
    Dim preFile As String
    
    curIndex = zhFileList.Index(zhrStatus.sCur_zhSubFile)
    preFile = zhFileList.Item(curIndex - 1)
    If preFile = "" Then preFile = zhFileList.Item(zhFileList.Count - 1)
    If preFile = "" Then Exit Sub
    
    GetView preFile
    
End Sub

Public Sub mnuhelp_About_Click()

    Dim sAbout As String
    sAbout = sAbout + Space$(4) + App.ProductName + " (Build" + Str$(App.Major) + "." + Str$(App.Minor) + "." + Str$(App.Revision) + ")" + vbCrLf
    sAbout = sAbout + Space$(4) + App.LegalCopyright
    MsgBox sAbout, vbInformation, "About"

End Sub


Public Sub mnuBookmark_add_Click()

    If zhrStatus.sCur_zhFile = "" Then Exit Sub
    Load mnuBookmark(mnuBookmark.Count)

    With mnuBookmark(mnuBookmark.Count - 1)
        .Caption = MainFrm.IEView.Document.Title
        .Tag = bddir(zhrStatus.sCur_zhFile) & zhrStatus.sCur_zhSubFile
    End With

End Sub

Public Sub mnuBookmark_Click(Index As Integer)
loadzh mnuBookmark(Index).Tag
End Sub

Public Sub mnuBookmark_manage_Click()

    If mnuBookmark.Count = 1 Then Exit Sub
    frmBookmark.Show 1
    Unload frmBookmark
    saveMNUBookmark

End Sub

Public Sub mnufile_Close_Click()

    Dim fso As New FileSystemObject
    'Dim i As Integer
    rememberNew sMemoryFile, zhrStatus.sCur_zhFile, zhrStatus.sCur_zhSubFile
    'zhReaderReset
    On Error Resume Next

    If fso.FolderExists(sTempZH) Then fso.DeleteFolder sTempZH, False

End Sub

Public Sub mnufile_exit_Click()

    Unload Me

End Sub

Public Sub mnufile_Open_Click()


    Dim sCD1InitDir As String
    Dim fso As New FileSystemObject
    
    rememberNew sMemoryFile, zhrStatus.sCur_zhFile, zhrStatus.sCur_zhSubFile

    If zhrStatus.sCur_zhFile <> "" Then

        If fso.FileExists(zhrStatus.sCur_zhFile) = True Then sCD1InitDir = fso.GetParentFolderName(zhrStatus.sCur_zhFile)
    Else
        sCD1InitDir = iniGetSetting(zhtmIni, "ReaderStyle", "LastPath")
    End If

    If fso.FolderExists(sCD1InitDir) Then cDlg.InitDir = sCD1InitDir
    cDlg.FileName = ""
    cDlg.Filter = "所有文件|*.*"
    cDlg.ShowOpen

    If cDlg.FileName <> "" Then
        loadzh cDlg.FileName
    End If

End Sub

Public Sub IEView_NavigateComplete2(ByVal pDisp As Object, URL As Variant)

    'Set ieDocument = IEView.Document
    
    MainFrm.LeftFrame.Enabled = True
    MainFrm.Navigated = True

    If zhrStatus.sCur_zhFile = "" Then
        ChangeMainFrmCaption App.ProductName
    Else
        ChangeMainFrmCaption zhrStatus.sCur_zhFile & " - " & App.ProductName
    End If
    
    MainFrm.stsBar.Caption = bddir(zhrStatus.sCur_zhFile) & zhrStatus.sCur_zhSubFile
    
End Sub

Public Sub ShowMenu(showit As Boolean)

    Dim i As Integer

    If showit = False Then
        mnuView_Menu.Checked = False
        zhrStatus.bMenuShowed = False

        For i = 0 To mnu.Count - 1
            mnu(i).Visible = False
        Next

    Else
        mnuView_Menu.Checked = True
        zhrStatus.bMenuShowed = True

        For i = 0 To mnu.Count - 1
            mnu(i).Visible = True
        Next

    End If

End Sub

Public Sub ShowStatusBar(showit As Boolean)

    If showit Then
        mnuView_StatusBar.Checked = True
        zhrStatus.bStatusBarShowed = True
        stsBar.Height = 375
        Form_Resize
    Else
        mnuView_StatusBar.Checked = False
        zhrStatus.bStatusBarShowed = False
        stsBar.Height = 0
        Form_Resize
    End If

End Sub

Public Sub ShowLeft(showit As Boolean)

    If showit Then
        mnuView_Left.Checked = True
        zhrStatus.bLeftShowed = True
        '    End If
    Else
        mnuView_Left.Checked = False
        zhrStatus.bLeftShowed = False
    End If

    Form_Resize

End Sub

Public Sub mnuFile_PReFerence_Click()

    Load frmOptions
    frmOptions.Show

End Sub


Public Sub mnuView_FullScreen_Click()

    NotResize = True
    bFullScreen = MFullScreen.switch_FullScreen(Me)
    mnuView_FullScreen.Checked = bFullScreen
    NotResize = False
    Form_Resize

End Sub

Public Sub mnuView_Left_Click()

    If zhrStatus.bLeftShowed Then
        ShowLeft False
    Else
        ShowLeft True
    End If

End Sub

Public Sub mnuView_Menu_Click()

    If zhrStatus.bMenuShowed Then
        ShowMenu False
    Else
        ShowMenu True
    End If

End Sub

Public Sub mnuView_StatusBar_Click()

    If zhrStatus.bStatusBarShowed Then
        ShowStatusBar False
    Else
        ShowStatusBar True
    End If

End Sub


Public Sub imgSplitter_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    With imgSplitter
        picSplitter.Move .Left, .Top, picSplitter.Width, .Height
    End With

    picSplitter.Visible = True
    mbMoving = True

End Sub

Public Sub imgSplitter_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    Dim sglPos As Single

    If mbMoving Then
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

Public Sub imgSplitter_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    picSplitter.Visible = False
    mbMoving = False
    imgSplitter.Left = picSplitter.Left
    Form_Resize

End Sub









Public Sub appHtmlAbout()

    Dim fso As New Scripting.FileSystemObject
    Dim fsoTS As Scripting.TextStream
    Dim sAppHtmlAboutFile As String
    sAppHtmlAboutFile = fso.BuildPath(App.Path, cHtmlAboutFilename)

    If fso.FileExists(sAppHtmlAboutFile) = False Then
        Set fsoTS = fso.CreateTextFile(sAppHtmlAboutFile, True)

        With fsoTS
            .WriteLine "<html>"
            .WriteLine "<head>"
            .WriteLine "<Title>Zippacked Html Reader</title>"
            .WriteLine "<meta http-equiv=Content-Type content=" & Chr$(34) & "text/html; charset=us-ascii" & Chr$(34) & ">"
            .WriteLine "</head>"
            .WriteLine "<body  background=images\bg.jpg >"
            .WriteLine "<p align=right ><span lang=EN-US style='font-size:24.0pt;font-family:TAHOMA,Courier New'>" & App.ProductName & " (Build" & Str$(App.Major) + "." + Str$(App.Minor) & "." & Str$(App.Revision) & ")</span></p>"
            .WriteLine "<p align=right ><span lang=EN-US style='font-size:24.0pt;font-family:TAHOMA,Courier New'>" & App.LegalCopyright & "</span></span></p>"
            .WriteLine "</body>"
            .WriteLine "</html>"
        End With

        fsoTS.Close
    End If

    NotPreOperate = True
    IEView.Navigate2 sAppHtmlAboutFile

End Sub

Sub loadMNUBookmark()

 
    loadBookmark zhtmIni, mnuBookmark

End Sub

Sub saveMNUBookmark()

     saveBookmark zhtmIni, mnuBookmark

End Sub


Public Sub loadzh(ByVal thisfile As String)

    Dim fso As New Scripting.FileSystemObject
    Dim folderToLoad As String
    Dim fileToRead As String
    
    thisfile = Replace(thisfile, "/", "\")
    thisfile = LeftDelete(thisfile, "\")
    
    If thisfile = "<上一级目录>" And zhrStatus.sCur_zhFile <> "" Then
        thisfile = fso.GetParentFolderName(zhrStatus.sCur_zhFile)
    End If
    
    If fso.FolderExists(thisfile) = True Then
        folderToLoad = thisfile
    ElseIf fso.FileExists(thisfile) = True Then
        folderToLoad = fso.GetParentFolderName(thisfile)
        fileToRead = fso.GetFileName(thisfile)
    End If
    
    
    'zhReaderReset
    
    sTempZH = fso.GetBaseName(fso.GetTempName)
    sTempZH = fso.BuildPath(Tempdir, sTempZH)
    Do Until fso.FolderExists(sTempZH) = False
        sTempZH = fso.GetBaseName(fso.GetTempName)
        sTempZH = fso.BuildPath(Tempdir, sTempZH)
    Loop
    fso.CreateFolder sTempZH
    
    zhrStatus.sCur_zhFile = folderToLoad
    
    DirList.Enabled = False
    If folderToLoad = "" Then
        DriveList.Drive = DriveList.List(0)
        DirList.Path = Left$(DriveList.Drive, 2)
        ChangeMainFrmCaption "My Computer:"
    Else
        DriveList.Drive = Left$(folderToLoad, 2)
        DirList.Path = folderToLoad
        ChangeMainFrmCaption folderToLoad
    End If
    DirList.Enabled = True
   
    If fileToRead <> "" Then
        GetView fileToRead
    ElseIf folderToLoad <> "" Then
        viewIndex folderToLoad
    End If
    
    If folderToLoad <> "" Then
        zhRecentFile.Add folderToLoad, folderToLoad
        zhRecentFile.FillinMenu mnuFile_Recent
    End If
    
End Sub




Public Function execZhcommand(ByVal sCommandLine As String) As Boolean

    Dim cmdName As String
    Dim cmdArgu As String
    
    On Error GoTo Herr
    execZhcommand = True
    sCommandLine = Replace$(sCommandLine, "\", "/")
    sCommandLine = LeftRight(sCommandLine, "://", ReturnEmptyStr)
    
    
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

Public Function isZhcommand(ByVal sUrl As String) As Boolean

    If LCase$(Left$(sUrl, 6)) = "zhcmd:" Then isZhcommand = True

End Function



Public Sub ChangeMainFrmCaption(sTitle As String)

    If bFullScreen Then Exit Sub
    MainFrm.Caption = sTitle

End Sub
Public Sub DriveList_change()

Dim fso As New FileSystemObject
Dim thedrive As String
thedrive = Left$(DriveList.Drive, 1)
If fso.GetDrive(thedrive).IsReady = False Then
    MsgBox "Drive " + Chr$(34) + thedrive + ":" + Chr$(34) + " is not ready.", vbCritical, "Alert"
    Exit Sub
End If

DirList.Path = thedrive & ":"

End Sub


Public Sub GetView(shortfile As String)

    If shortfile = "" Then MainFrm.appHtmlAbout: Exit Sub
    shortfile = UnescapeUrl(shortfile)
    Dim fso As New gCFileSystem
    'Dim ts As scripting.TextStream
    Dim tempfile As String
    Dim tempFile2 As String
    'Dim tempVS As ViewerStyle
    Dim bUseTemplate As Boolean
    Dim sTemplateFile As String

    tempfile = fso.BuildPath(zhrStatus.sCur_zhFile, shortfile)
    If fso.PathExists(tempfile) = False Then Exit Sub
    tempFile2 = fso.BuildPath(MainFrm.sTempZH, shortfile & TempHtm)
    xMkdir fso.GetParentFolderName(tempFile2)

    zhrStatus.sCur_zhSubFile = shortfile
    MainFrm.NotPreOperate = True

    Select Case chkFileType(tempfile) 'file.bas

    Case ftIE
        MainFrm.NotPreOperate = False
        IEView.Navigate2 tempfile
        Exit Sub
    Case ftTxt, ftVIDEO, ftAUDIO, ftIMG
        sTemplateFile = iniGetSetting(MainFrm.zhtmIni, "Viewstyle", "TemplateFile")
        bUseTemplate = CBoolStr(iniGetSetting(MainFrm.zhtmIni, "ViewStyle", "UseTemplate"))

        If bUseTemplate Then

            If createHtmlFromTemplate(tempfile, sTemplateFile, tempFile2) Then
                IEView.Navigate2 tempFile2
            ElseIf createDefaultHtml(tempfile, tempFile2) Then
                IEView.Navigate2 tempFile2
            Else
                IEView.Navigate2 tempfile
            End If

        ElseIf createDefaultHtml(tempfile, tempFile2) Then
            IEView.Navigate2 tempFile2
        Else
            IEView.Navigate2 tempfile
        End If

    Case Else
    
        ShellExecute Me.hwnd, "open", tempfile, "", "", 1
       
    End Select

End Sub

Public Sub viewIndex(sPath As String)
    
    Dim sHtmFile As String
    Dim sTmpFile As String
    Dim sFileList() As String
    Dim lFileCount As Long
    Dim fso As New FileSystemObject
    Dim fd As Folder
    Dim fds As Folders
    
    If fso.FolderExists(sPath) = False Then Exit Sub
    
    
'    If fso.FileExists(sHtmFile) = True Then
'        IEView.Navigate2 sHtmFile
'        Exit Sub
'    End If

    zhFileList.Parentfolder = sPath
    zhFileList.Create cmbFilter.text
    
    lFileCount = zhFileList.Count - 1
    
     If lFileCount >= 0 Then
     
        ReDim sFileList(lFileCount) As String
        Dim l As Long
        For l = 0 To lFileCount
        sFileList(l) = "zhcmd://GetView|/" & zhFileList.Item(l)
        Next
    End If
    
    If fso.GetParentFolderName(sPath) <> "" Then
        lFileCount = lFileCount + 1
        ReDim Preserve sFileList(lFileCount) As String
        sFileList(lFileCount) = "zhcmd://loadzh|/" & "<上一级目录>\"
    End If
    Set fds = fso.GetFolder(sPath).SubFolders
    For Each fd In fds
        lFileCount = lFileCount + 1
        ReDim Preserve sFileList(lFileCount) As String
        sFileList(lFileCount) = "zhcmd://loadzh|/" & bddir(fd.Path)
    Next
        
    If lFileCount < 0 Then Exit Sub
    
    sHtmFile = fso.BuildPath(sTempZH, "index.htm")
    sTmpFile = fso.BuildPath(sTempZH, "index.txt")
    
    Set fso = Nothing
    
    If IndexFromFileList(sFileList, sTmpFile) Then

        Dim sTemplateFile As String
        Dim bUseTemplate As Boolean
        sTemplateFile = iniGetSetting(MainFrm.zhtmIni, "Viewstyle", "TemplateFile")
        bUseTemplate = CBoolStr(iniGetSetting(MainFrm.zhtmIni, "ViewStyle", "UseTemplate"))
        
        If bUseTemplate Then
            If createHtmlFromTemplate(sTmpFile, sTemplateFile, sHtmFile) Then
                IEView.Navigate2 sHtmFile
            ElseIf createDefaultHtml(sTmpFile, sHtmFile) Then
                IEView.Navigate2 sHtmFile
            Else
                IEView.Navigate2 sTmpFile
            End If
        ElseIf createDefaultHtml(sTmpFile, sHtmFile) Then
            IEView.Navigate2 sHtmFile
        Else
            IEView.Navigate2 sTmpFile
        End If
        
    End If
    
    
    
    
End Sub

