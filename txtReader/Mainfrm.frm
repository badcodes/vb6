VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form MainFrm 
   Caption         =   "txtReader"
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
   Begin MSComctlLib.StatusBar stsBar 
      Height          =   372
      Left            =   -12
      TabIndex        =   7
      Top             =   3036
      Width           =   7056
      _ExtentX        =   12446
      _ExtentY        =   656
      ShowTips        =   0   'False
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9843
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            AutoSize        =   2
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picSplitter 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillColor       =   &H00808080&
      Height          =   4800
      Left            =   4776
      ScaleHeight     =   2090.126
      ScaleMode       =   0  'User
      ScaleWidth      =   390
      TabIndex        =   0
      Top             =   -1692
      Visible         =   0   'False
      Width           =   40
   End
   Begin VB.Frame theShow 
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   2532
      Left            =   3312
      TabIndex        =   1
      Top             =   156
      Width           =   3384
      Begin RichTextLib.RichTextBox rtxtView 
         Height          =   2280
         Left            =   108
         TabIndex        =   5
         Top             =   156
         Width           =   3060
         _ExtentX        =   5398
         _ExtentY        =   4022
         _Version        =   393217
         BorderStyle     =   0
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         MousePointer    =   1
         AutoVerbMenu    =   -1  'True
         OLEDragMode     =   0
         OLEDropMode     =   0
         TextRTF         =   $"Mainfrm.frx":08CA
      End
   End
   Begin MSComDlg.CommonDialog cDlg 
      Left            =   3756
      Top             =   2712
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox LeftFrame 
      Height          =   2868
      Left            =   360
      ScaleHeight     =   2820
      ScaleWidth      =   2592
      TabIndex        =   2
      Top             =   24
      Width           =   2640
      Begin VB.ComboBox cmbFilter 
         Appearance      =   0  'Flat
         Height          =   288
         Left            =   84
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   348
         Width           =   2352
      End
      Begin VB.DriveListBox DriveList 
         Height          =   288
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   2415
      End
      Begin MSComctlLib.TreeView FileList 
         Height          =   1944
         Left            =   132
         TabIndex        =   6
         Top             =   732
         Width           =   2244
         _ExtentX        =   3958
         _ExtentY        =   3429
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   423
         LabelEdit       =   1
         Style           =   7
         ImageList       =   "Listimg"
         BorderStyle     =   1
         Appearance      =   1
      End
   End
   Begin MSComctlLib.ImageList Listimg 
      Left            =   6384
      Top             =   2784
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
            Picture         =   "Mainfrm.frx":0962
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mainfrm.frx":123C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mainfrm.frx":1B16
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mainfrm.frx":23F0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image imgSplitter 
      Appearance      =   0  'Flat
      Height          =   5172
      Left            =   5364
      MousePointer    =   9  'Size W E
      Top             =   -2040
      Width           =   20
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
      Begin VB.Menu mnuEdit_EditCurPage 
         Caption         =   "&Edit Current Page"
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit_SelectEditor 
         Caption         =   "Select Text Editor..."
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
      Begin VB.Menu mnuGo_Previous 
         Caption         =   "&Previous    Alt+Down"
      End
      Begin VB.Menu mnuGo_Next 
         Caption         =   "&Next         Alt+Up"
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
      Begin VB.Menu mnuHelp_About 
         Caption         =   "&About This"
      End
   End
End
Attribute VB_Name = "MainFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const sglSplitLimit = 500
Private Const minFormHeight = 2000
Private Const minFormWidth = 3000
Private Const icstDefaultFormHeight = 6000
Private Const icstDefaultFormWidth = 8000
Private Const icstDefaultLeftWidth = 1500
Private Const lcstFittedListItemsNum = 2500
Private mbMoving As Boolean
Private Tempdir As String
Public sTempZH As String
Private NotResize As Boolean
Public zhtmIni As String
Public bIsZhtm As Boolean
Public bFullScreen As Boolean
Public WithEvents lUnzip As cUnzip
Attribute lUnzip.VB_VarHelpID = -1
Public WithEvents lZip As cZip
Attribute lZip.VB_VarHelpID = -1
Public zhRecentFile As New CRecentFile
Private sMemoryFile As String
Private PASSWORD As String
Private InvaildPassword As Boolean

Sub Loadlist(LContent() As String, lCount As Long)
    
    FileList.Visible = False
    FileList.Nodes.clear
    
    Dim tempNode As Node
    Dim Krelative As String
    Dim Krelationship As TreeRelationshipConstants
    Dim Kkey As String
    Dim Ktext As String
    Dim Kimageindex As Integer
    Dim Ktag As String
    Dim thename As String
    Dim i As Integer, pos As Integer

    Dim lEnd As Long
    lEnd = lCount - 1

    For i = 0 To lEnd
        On Error GoTo CatalogError
        stsBar.SimpleText = "Loading File List " & Format$(((i + 1) / lCount), "0.0%") & " (" & Str$(i + 1) & " of" & Str$(lCount) & ")"
        thename = LContent(0, i)
        Ktag = LContent(1, i)
        Kkey = "ZTM" + LContent(0, i)

        If Right$(thename, 1) = "\" Then thename = Left$(thename, Len(thename) - 1)
        pos = InStrRev(thename, "\")

        If pos = 0 Then
            Ktext = thename

            If Right$(LContent(0, i), 2) = ":\" Then
                Kimageindex = 4
            ElseIf Right$(LContent(0, i), 1) = "\" Then
                Kimageindex = 1
            Else
                Kimageindex = 3
            End If

            Set tempNode = FileList.Nodes.Add(, , Kkey, Ktext, Kimageindex)
            tempNode.Tag = Ktag
        Else
            Krelative = "ZTM" + Left$(LContent(0, i), pos)
            Krelationship = tvwChild
            Ktext = Right$(thename, Len(thename) - pos)

            If Right$(LContent(0, i), 1) = "\" Then
                Kimageindex = 1
            Else
                Kimageindex = 3
            End If
            Set tempNode = FileList.Nodes.Add(Krelative, Krelationship, Kkey, Ktext, Kimageindex)
            tempNode.Tag = Ktag

        End If

CatalogError:
    Next
    
    FileList.Visible = True

End Sub

Private Sub cmbFilter_Change()
If cmbFilter.Enabled = False Then Exit Sub
loadzh zhrStatus.sCur_zhFile, "", True
End Sub

'Private Function ieDocument_oncontextmenu() As Boolean
'
'If App.ProductName <> "zhReader" Then Exit Function
'ieDocument_oncontextmenu = False
'MainFrm.PopupMenu mnuIe
'ieDocument_oncontextmenu = True
'End Function


Private Sub DriveList_Change()
If DriveList.Enabled = False Then Exit Sub

Dim fso As New FileSystemObject
Dim thedrive As String
thedrive = Left$(DriveList.Drive, 1)
If fso.GetDrive(thedrive).IsReady = False Then
    MsgBox "Drive " + Chr$(34) + thedrive + ":" + Chr$(34) + " is not ready.", vbCritical, "Alert"
    Exit Sub
End If

loadzh fso.GetAbsolutePathName(thedrive & ":")

End Sub



Private Sub mnuEdit_Editcurpage_Click()
    
    Dim sShellTextEditor As String
    If zhrStatus.sCur_zhSubFile = "" Then Exit Sub
    sShellTextEditor = iniGetSetting(MainFrm.zhtmIni, "ReaderStyle", "TextEditor")
    If sShellTextEditor = "" Then
        mnuEdit_SelectEditor_Click
        sShellTextEditor = iniGetSetting(MainFrm.zhtmIni, "ReaderStyle", "TextEditor")
    End If
    If sShellTextEditor = "" Then Exit Sub
        
    Dim fso As New gCFileSystem
    
    Dim sTmpFile As String

    If bIsZhtm Then
        sTmpFile = fso.BuildPath(sTempZH, zhrStatus.sCur_zhSubFile)

        If fso.PathExists(sTmpFile) = False Then
            myXUnzip zhrStatus.sCur_zhFile, zhrStatus.sCur_zhSubFile, sTempZH, zhrStatus.sPWD
        End If

        If fso.PathExists(sTmpFile) = False Then Exit Sub
    Else
        sTmpFile = fso.BuildPath(zhrStatus.sCur_zhFile, zhrStatus.sCur_zhSubFile)
    End If

    ShellAndClose sShellTextEditor & " " & Chr(34) & sTmpFile & Chr(34), vbNormalFocus
    

    If bIsZhtm Then
        myXZip zhrStatus.sCur_zhFile, zhrStatus.sCur_zhSubFile, zhrStatus.sPWD, sTempZH
    End If
    
   GetView zhrStatus.sCur_zhSubFile

End Sub

'
Private Sub Form_Load()
       
    MainFrm.Caption = App.ProductName & " [" & App.LegalCopyright & "]"
    loadFormStr MainFrm
    Dim fso As New gCFileSystem
    Dim appPath As String
    
    appPath = fso.BuildPath(Environ("APPDATA"), App.ProductName)
    If fso.PathExists(appPath) = False Then MkDir appPath
    zhtmIni = fso.BuildPath(appPath, "config.ini")
    sMemoryFile = fso.BuildPath(appPath, "memory.encrypt")
    

    
    
    NotResize = True
   
    On Error Resume Next
     Tempdir = fso.BuildPath(App.Path, "Cache")

    If fso.PathExists(Tempdir) Then RmDir Tempdir

    If fso.PathExists(Tempdir) = False Then MkDir Tempdir

    If fso.PathExists(fso.BuildPath(App.Path, cHtmlAboutFilename)) Then _
       Kill fso.BuildPath(App.Path, cHtmlAboutFilename)
       
    Dim theRS As ReaderStyle
    Dim theVs As ViewerStyle
    GetViewerStyle zhtmIni, theVs
    GetReaderStyle zhtmIni, theRS
   
    With rtxtView
        .BackColor = theVs.BackColor
        CopytoIfont theVs.Viewfont, .Font
        .Tag = CStr(theVs.ForeColor)
    End With
    
    
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
    
    appHtmlAbout
    
    If theRS.FullScreenMode Then mnuView_FullScreen_Click
    
    zhRecentFile.maxItem = Val(iniGetSetting(zhtmIni, "ViewStyle", "RecentMax"))
    zhRecentFile.maxCaptionLength = 30
    zhRecentFile.LoadFromIni zhtmIni
    zhRecentFile.FillinMenu mnuFile_Recent
        
    loadzh thisfile

    NotResize = False
    

    
    
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
            mnuView_FullScreen_Click
        Case vbKeyF4
            mnufile_Close_Click
        Case Else
            '                KeyCode = iKeyCode
        End Select

    ElseIf Shift = vbAltMask Then

        Select Case iKeyCode
        Case vbKeyUp
            mnuGo_Previous_Click
        Case vbKeyDown
            mnuGo_Next_Click
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

Private Sub Form_KeyPress(KeyAscii As Integer)

    KeyAscii = 0

End Sub

Private Sub Form_Unload(Cancel As Integer)

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
        .LeftWidth = imgSplitter.Left
    End With
        
    theRS.FullScreenMode = bFullScreen
    theRS.LastPath = fso.GetParentFolderName(zhrStatus.sCur_zhFile)
    theRS.ShowLeft = zhrStatus.bLeftShowed
    theRS.ShowMenu = zhrStatus.bMenuShowed
    theRS.ShowStatusBar = zhrStatus.bStatusBarShowed
    theRS.TextEditor = iniGetSetting(zhtmIni, "ReaderStyle", "TextEditor")
    SaveReaderStyle zhtmIni, theRS
    rememberNew sMemoryFile, zhrStatus.sCur_zhFile, zhrStatus.sCur_zhSubFile
    zhRecentFile.SaveToIni zhtmIni
    SaveFileFilter zhtmIni, cmbFilter
    
    On Error Resume Next

    If fso.FolderExists(Tempdir) Then fso.DeleteFolder Tempdir, True

    If fso.FolderExists(sTempZH) Then fso.DeleteFolder (sTempZH), True
    

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
      
    With DriveList
        .Top = ControlMargin
        .Left = ControlMargin
        .Width = Abs(LeftFrame.Width - ControlMargin * 3)
    End With
    
    With cmbFilter
        .Top = DriveList.Top + DriveList.Height
        .Left = ControlMargin
        .Width = Abs(LeftFrame.Width - ControlMargin * 3)
    End With
    
    With FileList
        .Top = cmbFilter.Top + cmbFilter.Height
        .Left = ControlMargin
        .Width = Abs(LeftFrame.Width - ControlMargin * 3)
        .Height = Abs(LeftFrame.Height - .Top - ControlMargin)
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
    

    With rtxtView
        .Left = 0
        .Top = 0
        .Height = theShow.Height
        .Width = theShow.Width
    End With


    MainFrm.Refresh

End Sub




Private Sub fileList_Collapse(ByVal Node As MSComctlLib.Node)

    Node.Image = 1

End Sub

Private Sub fileList_Expand(ByVal Node As MSComctlLib.Node)

    Node.Image = 2

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
Private Sub mnu_Click(Index As Integer)
    
    If zhrStatus.sCur_zhFile = "" Then
        mnuFile_Close.Enabled = False
    Else
        mnuFile_Close.Enabled = True
    End If

    If zhrStatus.sCur_zhSubFile = "" Then
        mnuEdit_EditCurPage.Enabled = False
    Else
        mnuEdit_EditCurPage.Enabled = True
    End If


End Sub

Private Sub mnuEdit_SelectEditor_Click()

    Dim fso As New gCFileSystem
    Dim sShellTextEditor As String
    
    sShellTextEditor = iniGetSetting(MainFrm.zhtmIni, "ReaderStyle", "TextEditor")
    
    If sShellTextEditor <> "" Then
        cDlg.InitDir = fso.GetParentFolderName(sShellTextEditor)
        cDlg.FileName = sShellTextEditor
    End If
    
        cDlg.Filter = "EXE File|*.exe|All Files|*.*"
        cDlg.ShowOpen
    
    sShellTextEditor = cDlg.FileName
    If sShellTextEditor <> "" Then iniSaveSetting MainFrm.zhtmIni, "ReaderStyle", "TextEditor", sShellTextEditor

End Sub



Private Sub mnuFile_OpenDir_Click()
    
    Dim sPath As String
    rememberNew sMemoryFile, zhrStatus.sCur_zhFile, zhrStatus.sCur_zhSubFile
    sPath = openDirDialog(Me.hwnd)

    If sPath <> "" Then loadzh sPath

End Sub

Private Sub mnuFile_Recent_Click(Index As Integer)
    rememberNew sMemoryFile, zhrStatus.sCur_zhFile, zhrStatus.sCur_zhSubFile
    Dim sFstPart As String, sSecPart As String
    sFstPart = LeftLeft(mnuFile_Recent(Index).Tag, "|", vbTextCompare, ReturnOriginalStr)
    sSecPart = LeftRight(mnuFile_Recent(Index).Tag, "|", vbTextCompare, ReturnEmptyStr)
    loadzh sFstPart, sSecPart
End Sub


Private Sub mnuFilter_Add_Click(Index As Integer)
Dim tempstr As String
tempstr = InputBox("输入文件名过滤字符串" & vbCrLf)
cmbFilter.AddItem tempstr

End Sub

Private Sub mnuFilter_Delete_Click()
    If cmbFilter.ListIndex < 0 Then Exit Sub
    
    Dim a As VbMsgBoxResult
    Dim theindex As Integer
    Dim thetext As String
    theindex = cmbFilter.ListIndex
    thetext = cmbFilter.Text
    
    a = MsgBox("Delete " + thetext + " ?", vbOKCancel)
    
    If a = vbCancel Then Exit Sub
    cmbFilter.RemoveItem theindex
End Sub

Private Sub mnuGo_Next_Click()
Dim nextPos As Long
nextPos = getCurIndex + 1
If nextPos > getLastIndex Then nextPos = getFirstIndex
If nextPos = 0 Then Exit Sub


 GetView FileList.Nodes.Item(nextPos).Tag

End Sub
Public Function getCurIndex() As Long
On Error GoTo Herr
getCurIndex = FileList.Nodes("ZTM" & zhrStatus.sCur_zhSubFile).Index
Herr:
End Function
Public Function getFirstIndex() As Long
Dim tempNode As Node
For Each tempNode In FileList.Nodes
    If tempNode.Children = 0 And Right$(tempNode.Tag, 1) <> "\" Then
        getFirstIndex = tempNode.Index
        Exit For
    End If
Next
End Function
Public Function getLastIndex() As Long
Dim tempNode As Node
If FileList.Nodes.Count < 1 Then Exit Function
getLastIndex = FileList.Nodes.Count
Set tempNode = FileList.Nodes(getLastIndex)
If tempNode.Children > 0 Or Right$(tempNode.Tag, 1) = "\" Then getLastIndex = 0

End Function



Private Sub mnuGo_Previous_Click()
Dim prePos As Long
prePos = getCurIndex - 1
If prePos < getFirstIndex Then prePos = getLastIndex
If prePos = 0 Then Exit Sub

GetView FileList.Nodes.Item(prePos).Tag
End Sub

Private Sub mnuhelp_About_Click()

    Dim sAbout As String
    sAbout = sAbout + Space$(4) + App.ProductName + " (Build" + Str$(App.Major) + "." + Str$(App.Minor) + "." + Str$(App.Revision) + ")" + vbCrLf
    sAbout = sAbout + Space$(4) + App.LegalCopyright
    MsgBox sAbout, vbInformation, "About"

End Sub


Private Sub mnuBookmark_add_Click()

    Dim cfs As New gCFileSystem
    
    If zhrStatus.sCur_zhFile = "" Then Exit Sub
    Load mnuBookmark(mnuBookmark.Count)

    With mnuBookmark(mnuBookmark.Count - 1)
        .Caption = cfs.GetBaseName(zhrStatus.sCur_zhFile) & "-" & cfs.GetBaseName(zhrStatus.sCur_zhSubFile)
        .Tag = zhrStatus.sCur_zhFile & "|" & zhrStatus.sCur_zhSubFile
    End With


End Sub

Private Sub mnuBookmark_Click(Index As Integer)
Dim sFstPart As String, sSecPart As String
sFstPart = LeftLeft(mnuBookmark(Index).Tag, "|", vbTextCompare, ReturnOriginalStr)
sSecPart = LeftRight(mnuBookmark(Index).Tag, "|", vbTextCompare, ReturnEmptyStr)
loadzh sFstPart, sSecPart

End Sub

Private Sub mnuBookmark_manage_Click()

    If mnuBookmark.Count = 1 Then Exit Sub
    frmBookmark.Show 1
    Unload frmBookmark
    saveMNUBookmark

End Sub

Private Sub mnufile_Close_Click()

    Dim fso As New FileSystemObject
    'Dim i As Integer
    rememberNew sMemoryFile, zhrStatus.sCur_zhFile, zhrStatus.sCur_zhSubFile
    zhReaderReset
    On Error Resume Next

    If fso.FolderExists(sTempZH) Then fso.DeleteFolder sTempZH, False

End Sub

Private Sub mnufile_exit_Click()

    Unload Me

End Sub

Public Sub mnufile_Open_Click()

    Dim thisfile As String
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
    cDlg.Filter = "Zippacked Html File|*.zhtm|Zip File|*.zip|所有文件|*.*"
    cDlg.ShowOpen

    If cDlg.FileName <> "" Then
        loadzh cDlg.FileName
    End If

End Sub


Private Sub ShowMenu(showit As Boolean)

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

Private Sub ShowStatusBar(showit As Boolean)

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

Private Sub ShowLeft(showit As Boolean)

    If showit Then
        mnuView_Left.Checked = True
        zhrStatus.bLeftShowed = True
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
        zhrStatus.bLeftShowed = False
    End If

    Form_Resize

End Sub

Private Sub mnuFile_PReFerence_Click()

    Load frmOptions
    frmOptions.Show

End Sub


'Private Sub mnuIe_AddBookmark_Click()
'mnuBookmark_add_Click
'End Sub
'
'Private Sub mnuIe_Backward_Click()
'On Error Resume Next
'IEView(iCurIEView).GoBack
'End Sub
'
'Private Sub mnuIe_copy_Click()
'On Error Resume Next
'IEView(iCurIEView).ExecWB OLECMDID_COPY, OLECMDEXECOPT_DODEFAULT
'End Sub
'
'Private Sub mnuIe_Forward_Click()
'On Error Resume Next
'IEView(iCurIEView).GoForward
'End Sub
'
'Private Sub mnuIe_Print_Click()
'On Error Resume Next
'IEView(iCurIEView).ExecWB OLECMDID_PRINT, OLECMDEXECOPT_DODEFAULT
'End Sub
'
'Private Sub mnuIe_property_Click()
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
'Private Sub mnuIe_refresh_Click()
'IEView(iCurIEView).Refresh2
'End Sub
'
'Private Sub mnuIe_SelectAll_Click()
'On Error Resume Next
'IEView(iCurIEView).ExecWB OLECMDID_SELECTALL, OLECMDEXECOPT_DODEFAULT
'End Sub
'
'Private Sub mnuIe_ViewSource_Click()
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
'    Put #iNum, , IEView(iCurIEView).Document.documentElement.outerHTML
'    Close iNum
'    sViewer = Replace$(sViewer, "%1", sTmpFile, , , vbTextCompare)
'    sViewer = Replace$(sViewer, "%l", sTmpFile, , , vbTextCompare)
'    Shell sViewer, vbNormalFocus
'End If
'End Sub

Private Sub mnuView_FullScreen_Click()

    NotResize = True
    bFullScreen = MFullScreen.switch_FullScreen(Me)
    mnuView_FullScreen.Checked = bFullScreen
    NotResize = False
    Form_Resize

End Sub

Private Sub mnuView_Left_Click()

    If zhrStatus.bLeftShowed Then
        ShowLeft False
    Else
        ShowLeft True
    End If

End Sub

Private Sub mnuView_Menu_Click()

    If zhrStatus.bMenuShowed Then
        ShowMenu False
    Else
        ShowMenu True
    End If

End Sub

Private Sub mnuView_StatusBar_Click()

    If zhrStatus.bStatusBarShowed Then
        ShowStatusBar False
    Else
        ShowStatusBar True
    End If

End Sub



Private Sub filelist_NodeClick(ByVal Node As MSComctlLib.Node)

    Dim fso As New gCFileSystem
    'Dim nodeTag As String

    'If Node.Tag = "" Then Exit Sub

    If bIsZhtm Then

        If Node.Tag = "..\" Then
            loadzh fso.GetParentFolderName(zhrStatus.sCur_zhFile), "", True
        ElseIf Right$(Node.Tag, 1) <> "\" Then
            GetView Node.Tag
        End If

        Exit Sub
    ElseIf bIsZhtm = False Then
    
        If Node.Tag = "..\" Then
            loadzh fso.GetParentFolderName(zhrStatus.sCur_zhFile), "", True
        ElseIf InStr(Node.Tag, ":") > 0 Then
            loadzh Node.Tag, "", True
        ElseIf Right$(Node.Tag, 1) = "\" Then
            loadzh fso.BuildPath(zhrStatus.sCur_zhFile, Node.Tag), "", True
        Else
            GetView Node.Tag
        End If


        Exit Sub
    End If

End Sub

Private Sub imgSplitter_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    With imgSplitter
        picSplitter.Move .Left, .Top, .Width \ 2, .Height - 20
    End With

    picSplitter.Visible = True
    mbMoving = True

End Sub

Private Sub imgSplitter_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

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

Private Sub imgSplitter_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    picSplitter.Visible = False
    mbMoving = False
    imgSplitter.Left = picSplitter.Left
    Form_Resize

End Sub

Public Sub getZIPContent(ByVal sZipfilename As String)


    Dim lfor As Long
    Dim fso As New FileSystemObject
    Dim fsofolder As Folder
    Dim fsofile As File
    Dim fsofolders As Folders
    Dim fsoFiles As Files
    Dim sFilename As String
    Dim sFilesInZip() As String
    Dim lFilesIZcount As Long
    Dim sFoldersInZip() As String
    Dim lFoldersIZcount As Long
    
    If bIsZhtm Then
    
        Set lUnzip = New cUnzip
        With lUnzip
        .ZipFile = sZipfilename
        .Directory
        End With
        If fso.GetParentFolderName(sZipfilename) <> "" Then
                ReDim sFoldersInZip(0) As String
                sFoldersInZip(0) = "..\"
                lFoldersIZcount = 1
            End If
        For lfor = 1 To lUnzip.FileCount
            sFilename = lUnzip.FileName(lfor)
            sFilename = toDosPath(sFilename)
            If Right(sFilename, 1) = "\" Then
                ReDim Preserve sFoldersInZip(lFoldersIZcount) As String
                sFoldersInZip(lFoldersIZcount) = sFilename
                lFoldersIZcount = lFoldersIZcount + 1
            ElseIf MYinstr(sFilename, cmbFilter.Text) Then
                ReDim Preserve sFilesInZip(lFilesIZcount) As String
                sFilesInZip(lFilesIZcount) = sFilename
                lFilesIZcount = lFilesIZcount + 1
            End If
        Next
        Set lUnzip = Nothing

    Else
        lfor = 0
        lFoldersIZcount = fso.GetFolder(sZipfilename).SubFolders.Count

        If fso.GetParentFolderName(sZipfilename) <> "" Then
            ReDim sFoldersInZip(lfor) As String
            sFoldersInZip(lfor) = "..\"
            lfor = lfor + 1
        Else
            Dim fsodrive As Drive

            For Each fsodrive In fso.Drives

                If fsodrive.IsReady Then
                    ReDim Preserve sFoldersInZip(lfor) As String
                    sFoldersInZip(lfor) = fsodrive.DriveLetter & ":\"
                    lfor = lfor + 1
                End If

            Next

        End If

        Set fsofolders = fso.GetFolder(sZipfilename).SubFolders

        For Each fsofolder In fsofolders
            ReDim Preserve sFoldersInZip(lfor) As String
            sFoldersInZip(lfor) = fsofolder.name & "\"
            lfor = lfor + 1
        Next

        lFoldersIZcount = lfor
        lfor = 0
        Set fsoFiles = fso.GetFolder(sZipfilename).Files

        For Each fsofile In fsoFiles
            If MYinstr(fsofile.name, cmbFilter.Text) Then
                ReDim Preserve sFilesInZip(lfor) As String
                sFilesInZip(lfor) = fsofile.name
                lfor = lfor + 1
            End If
        Next

        lFilesIZcount = lfor
    End If

    Dim zipContent() As String
    Dim lzipCount As Long
    Dim i As Long
    lzipCount = lFoldersIZcount + lFilesIZcount
    
    
    ReDim zipContent(1, lzipCount)
    Dim lEnd As Long
    lEnd = lFoldersIZcount - 1

    For i = 0 To lEnd
        zipContent(0, i) = sFoldersInZip(i)
        zipContent(1, i) = sFoldersInZip(i)
    Next

    lEnd = lFilesIZcount - 1

    For i = 0 To lEnd
        zipContent(0, i + lFoldersIZcount) = sFilesInZip(i)
        zipContent(1, i + lFoldersIZcount) = sFilesInZip(i)
    Next

    Loadlist zipContent(), lzipCount
    

    MainFrm.Enabled = True
    stsBar.SimpleText = "[" & sZipfilename & "] ->Read."
    'Loadlist trvwList, ZipContent(), ZipCcount
    'loadZipList List(1)

End Sub

Public Sub myXZip(ByVal sZipfilename As String, ByVal sFilesToProcess As String, ByVal sPWD As String, Optional ByVal sBasePath As String = "")

    'Dim fso As New FileSystemObject

    MainFrm.Enabled = False
    sFilesToProcess = Replace(sFilesToProcess, "\", "/")
    
    Set lZip = New cZip
    With lZip
        .ZipFile = sZipfilename
        .FileToProcess = sFilesToProcess
        '.FilesToExclude = ""
        .BasePath = sBasePath
        '.EncryptionPassword = sPWD
    End With
    Set lZip = Nothing

    MainFrm.Enabled = True


End Sub

Public Sub myXUnzip(ByVal sZipfilename As String, ByVal sFilesToProcess As String, ByVal sUnzipTo As String, Optional ByVal sPWD As String)
  

    
    MainFrm.Enabled = False
    sFilesToProcess = Replace(sFilesToProcess, "\", "/")
    
    Set lUnzip = New cUnzip
    With lUnzip
        .CaseSensitiveFileNames = False
        .PromptToOverwrite = True
        .UseFolderNames = True
        .ZipFile = sZipfilename
'        .ZipFilename = sZipfilename
        .FileToProcess = sFilesToProcess
        .UnzipFolder = sUnzipTo
    End With
    
    lUnzip.unzip
    
    Set lUnzip = Nothing

    MainFrm.Enabled = True
    

End Sub

Public Sub appHtmlAbout()

rtxtView.Visible = False
rtxtView.Text = aboutApp
'Me.RtxtView_ChangeAppearance
rtxtView.SelLength = Len(rtxtView.Text)
rtxtView.SelBold = 1
rtxtView.SelFontName = rtxtView.Font.name
rtxtView.SelFontSize = 18
rtxtView.SelAlignment = 1
rtxtView.SelStart = 0
rtxtView.SelLength = 0
rtxtView.Visible = True

'    Dim fso As New Scripting.FileSystemObject
'    Dim fsoTS As Scripting.TextStream
'    Dim sAppHtmlAboutFile As String
'    sAppHtmlAboutFile = fso.BuildPath(App.Path, cHtmlAboutFilename)
'
'    If fso.FileExists(sAppHtmlAboutFile) = False Then
'        Set fsoTS = fso.CreateTextFile(sAppHtmlAboutFile, True)
'
'        With fsoTS
'            .WriteLine "<html>"
'            .WriteLine "<head>"
'            .WriteLine "<Title>Zippacked Html Reader</title>"
'            .WriteLine "<meta http-equiv=Content-Type content=" & Chr$(34) & "text/html; charset=us-ascii" & Chr$(34) & ">"
'            .WriteLine "</head>"
'            .WriteLine "<body  background=images\bg.jpg >"
'            .WriteLine "<p align=right ><span lang=EN-US style='font-size:24.0pt;font-family:TAHOMA,Courier New'>" & App.ProductName & " (Build" & Str$(App.Major) + "." + Str$(App.Minor) & "." & Str$(App.Revision) & ")</span></p>"
'            .WriteLine "<p align=right ><span lang=EN-US style='font-size:24.0pt;font-family:TAHOMA,Courier New'>" & App.LegalCopyright & "</span></span></p>"
'            .WriteLine "</body>"
'            .WriteLine "</html>"
'        End With
'
'        fsoTS.Close
'    End If
'
'    NotPreOperate = True
'    IEView(iCurIEView).Navigate2 sAppHtmlAboutFile

End Sub

Sub loadMNUBookmark()

 
    loadBookmark zhtmIni, mnuBookmark

End Sub

Sub saveMNUBookmark()

     saveBookmark zhtmIni, mnuBookmark

End Sub

Sub zhReaderReset()

    Dim i As Integer

    With FileList
        .Visible = False
        .Nodes.clear
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
'    Erase sFilesInZip
'    lFilesIZcount = 0
'    Erase sFoldersInZip
'    lFoldersIZcount = 0
    'LeftStrip.Tabs(1).Selected = True
    appHtmlAbout


End Sub

Public Sub loadzh(ByVal thisfile As String, Optional ByVal firstfile As String = "", Optional Reloadit As Boolean = False)

    Dim fso As New Scripting.FileSystemObject
    Dim fE As Boolean
    Dim fdE As Boolean
    fE = fso.FileExists(thisfile)
    fdE = fso.FolderExists(thisfile)
    If fE = False And fdE = False Then Exit Sub
    

    
    If fE And firstfile <> "" Then
        loadZip thisfile, firstfile, Reloadit
        Exit Sub
    ElseIf fE Then
        firstfile = fso.GetFileName(thisfile)
        thisfile = fso.GetParentFolderName(thisfile)
    End If
    
    bIsZhtm = False
    
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

    zhReaderReset
    
    sTempZH = fso.GetBaseName(fso.GetTempName)
    sTempZH = fso.BuildPath(Tempdir, sTempZH)
    Do Until fso.FolderExists(sTempZH) = False
        sTempZH = fso.GetBaseName(fso.GetTempName)
        sTempZH = fso.BuildPath(Tempdir, sTempZH)
    Loop
    fso.CreateFolder sTempZH
    
    zhrStatus.sCur_zhFile = thisfile

    If firstfile = "" Then firstfile = searchMemory(sMemoryFile, zhrStatus.sCur_zhFile)
    'Dim sTmpText As String

    getZIPContent zhrStatus.sCur_zhFile


'    If mnuFile_Recent.count = 1 Then
'        mnufile_s2.Visible = False
'    Else
'        mnufile_s2.Visible = True
'    End If
    ShowMenu zhrStatus.bMenuShowed
    ShowLeft zhrStatus.bLeftShowed
    ShowStatusBar zhrStatus.bStatusBarShowed
    
    DriveList.Enabled = False
    If Left$(zhrStatus.sCur_zhFile, 2) <> "" Then
        DriveList.Drive = Left$(zhrStatus.sCur_zhFile, 2)
    End If
    DriveList.Enabled = True
    
    If firstfile <> "" Then
        GetView firstfile
    End If
    

End Sub

Public Sub loadZip(ByVal thisfile As String, Optional ByVal firstfile As String = "", Optional Reloadit As Boolean = False)

    Dim fso As New Scripting.FileSystemObject
    
    If fso.FileExists(thisfile) = False Then Exit Sub
    
    bIsZhtm = True
    
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

    zhReaderReset
    
    sTempZH = fso.GetBaseName(fso.GetTempName)
    sTempZH = fso.BuildPath(Tempdir, sTempZH)
    Do Until fso.FolderExists(sTempZH) = False
        sTempZH = fso.GetBaseName(fso.GetTempName)
        sTempZH = fso.BuildPath(Tempdir, sTempZH)
    Loop
    fso.CreateFolder sTempZH
    
    zhrStatus.sCur_zhFile = thisfile

    If firstfile = "" Then firstfile = searchMemory(sMemoryFile, zhrStatus.sCur_zhFile)
    'Dim sTmpText As String

    getZIPContent zhrStatus.sCur_zhFile

    If firstfile <> "" Then
        GetView firstfile
    End If
    
'    If mnuFile_Recent.count = 1 Then
'        mnufile_s2.Visible = False
'    Else
'        mnufile_s2.Visible = True
'    End If
    ShowMenu zhrStatus.bMenuShowed
    ShowLeft zhrStatus.bLeftShowed
    ShowStatusBar zhrStatus.bStatusBarShowed
    
    DriveList.Enabled = False
    If Left$(zhrStatus.sCur_zhFile, 2) <> "" Then
        DriveList.Drive = Left$(zhrStatus.sCur_zhFile, 2)
    End If
    DriveList.Enabled = True
End Sub

Public Sub GetView(shortfile As String)

    If shortfile = "" Then MainFrm.appHtmlAbout: Exit Sub
    Dim fso As New gCFileSystem
    'Dim ts As scripting.TextStream
    Dim tempfile As String
    'Dim tempVS As ViewerStyle
    Dim bUseTemplate As Boolean
    Dim sTemplateFile As String

    If MainFrm.bIsZhtm Then
        MainFrm.myXUnzip zhrStatus.sCur_zhFile, shortfile, MainFrm.sTempZH, zhrStatus.sPWD
        tempfile = fso.BuildPath(MainFrm.sTempZH, shortfile)
        If fso.PathExists(tempfile) = False Then Exit Sub
    Else
        tempfile = fso.BuildPath(zhrStatus.sCur_zhFile, shortfile)
        If fso.PathExists(tempfile) = False Then Exit Sub
    End If

    zhrStatus.sCur_zhSubFile = shortfile
    FileList.Nodes(getCurIndex).Selected = True
    
    If zhrStatus.sCur_zhFile <> "" Then
    stsBar.Panels(1).Text = zhrStatus.sCur_zhFile
    Else
    stsBar.Panels(1).Text = "My Computer"
    End If
    stsBar.Panels(2).Text = zhrStatus.sCur_zhSubFile
    

    Select Case chkFileType(tempfile) 'file.bas

    Case ftRTF
        MainFrm.rtxtView.LoadFile tempfile, rtfRTF
        RtxtView_ChangeAppearance
    Case ftZIP
        MainFrm.loadZip tempfile
        Exit Sub
    Case ftZhtm
        MainFrm.loadZip tempfile
        Exit Sub
    Case ftIMG, ftVideo, ftAudio
        ShellExecute Me.hwnd, "open", tempfile, "", fso.GetParentFolderName(tempfile), 1
        Exit Sub
    Case Else
        MainFrm.rtxtView.LoadFile tempfile, rtfText
        RtxtView_ChangeAppearance
    End Select
    
    zhRecentFile.Add zhrStatus.sCur_zhFile & "|" & zhrStatus.sCur_zhSubFile, zhrStatus.sCur_zhFile & "|" & zhrStatus.sCur_zhSubFile
    zhRecentFile.FillinMenu mnuFile_Recent
    
End Sub
Public Sub RtxtView_ChangeAppearance()
With rtxtView
.Visible = False
.SelLength = Len(.Text)
If .SelLength <= 0 Then .SelLength = Len(.TextRTF)
.SelFontName = .Font.name
.SelFontSize = .Font.Size
.SelColor = CLng(.Tag)
.SelBold = .Font.Bold
.SelStrikeThru = .Font.Strikethrough
.SelItalic = .Font.Italic
.SelStart = 0
.SelLength = 0
.Visible = True
End With
End Sub
Private Sub lUnzip_PasswordRequest(sPassword As String, ByVal sName As String, bCancel As Boolean)

bCancel = False
Static lastName As String

If InvaildPassword = False And PASSWORD <> "" Then
    sPassword = PASSWORD
    If sName = lastName Then
        InvaildPassword = True
    Else
        lastName = sName
    End If
Else
    sPassword = InputBox(lUnzip.ZipFile & vbCrLf & sName & " Request For Password", "Password", "")
    If sPassword <> "" Then
        InvaildPassword = False
        PASSWORD = sPassword
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

