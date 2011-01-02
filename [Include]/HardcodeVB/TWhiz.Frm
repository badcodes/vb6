VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Begin VB.Form FTestWhizzyStuff 
   Caption         =   "Windows Interface Tricks"
   ClientHeight    =   6495
   ClientLeft      =   1110
   ClientTop       =   1815
   ClientWidth     =   8595
   Icon            =   "TWhiz.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6495
   ScaleWidth      =   8595
   Begin VB.TextBox txtSpecLoc 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5850
      TabIndex        =   27
      Text            =   "0"
      Top             =   156
      Width           =   510
   End
   Begin VB.TextBox txtDst 
      Height          =   330
      Left            =   1995
      TabIndex        =   7
      Top             =   4080
      Width           =   2115
   End
   Begin ComCtl2.UpDown udSysIcon 
      Height          =   375
      Left            =   6379
      TabIndex        =   26
      Top             =   2055
      Width           =   192
      _ExtentX        =   423
      _ExtentY        =   661
      _Version        =   327681
      BuddyControl    =   "txtSysIcon"
      BuddyDispid     =   196611
      OrigLeft        =   5760
      OrigTop         =   2047
      OrigRight       =   6000
      OrigBottom      =   2437
      Max             =   100
      SyncBuddy       =   -1  'True
      Wrap            =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin VB.TextBox txtSysIcon 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5850
      TabIndex        =   22
      Text            =   "0"
      Top             =   2055
      Width           =   510
   End
   Begin VB.CheckBox chkFileOnly 
      Caption         =   "Files Only"
      Height          =   225
      Left            =   135
      TabIndex        =   21
      Top             =   4095
      Width           =   1764
   End
   Begin VB.CheckBox chkMultiDest 
      Caption         =   "Multiple Destinations"
      Height          =   225
      Left            =   135
      TabIndex        =   20
      Top             =   4410
      Width           =   1764
   End
   Begin VB.CheckBox chkNoConfirm 
      Caption         =   "No Confirm"
      Height          =   225
      Left            =   135
      TabIndex        =   19
      Top             =   4725
      Width           =   1764
   End
   Begin VB.CheckBox chkMkDir 
      Caption         =   "Make Directories"
      Height          =   252
      Left            =   135
      TabIndex        =   18
      Top             =   5028
      Width           =   1764
   End
   Begin VB.CheckBox chkRename 
      Caption         =   "Rename on Collision"
      Height          =   225
      Left            =   135
      TabIndex        =   17
      Top             =   5364
      Width           =   1764
   End
   Begin VB.CheckBox chkSilent 
      Caption         =   "No Progress"
      Height          =   225
      Left            =   135
      TabIndex        =   16
      Top             =   5685
      Width           =   1764
   End
   Begin VB.CheckBox chkProgress 
      Caption         =   "Simple Progress"
      Height          =   225
      Left            =   135
      TabIndex        =   15
      Top             =   6000
      Width           =   1764
   End
   Begin VB.CheckBox chkUndo 
      Caption         =   "Recycle Bin"
      Height          =   225
      Left            =   135
      TabIndex        =   14
      Top             =   3780
      Width           =   1764
   End
   Begin VB.CommandButton cmdDocs 
      Caption         =   "Add To Docs"
      Height          =   435
      Left            =   135
      TabIndex        =   12
      Top             =   2220
      Width           =   1275
   End
   Begin VB.CommandButton cmdClearDoc 
      Caption         =   "Clear Docs"
      Height          =   435
      Left            =   135
      TabIndex        =   11
      Top             =   2745
      Width           =   1275
   End
   Begin VB.CommandButton cmdDst 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4152
      TabIndex        =   9
      Top             =   4080
      Width           =   435
   End
   Begin VB.CommandButton cmdSrc 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4152
      TabIndex        =   6
      Top             =   1155
      Width           =   435
   End
   Begin VB.TextBox txtSrc 
      Height          =   330
      Left            =   1995
      TabIndex        =   4
      Top             =   1155
      Width           =   2115
   End
   Begin VB.CommandButton cmdRename 
      Caption         =   "Rename"
      Height          =   435
      Left            =   135
      TabIndex        =   3
      Top             =   1695
      Width           =   1275
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   435
      Left            =   135
      TabIndex        =   2
      Top             =   1170
      Width           =   1275
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "Move"
      Height          =   435
      Left            =   135
      TabIndex        =   1
      Top             =   156
      Width           =   1275
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "Copy"
      Height          =   435
      Left            =   135
      TabIndex        =   0
      Top             =   645
      Width           =   1275
   End
   Begin ComCtl2.UpDown udSpecLoc 
      Height          =   375
      Left            =   6379
      TabIndex        =   28
      Top             =   156
      Width           =   192
      _ExtentX        =   423
      _ExtentY        =   661
      _Version        =   327681
      BuddyControl    =   "txtSpecLoc"
      BuddyDispid     =   196609
      OrigLeft        =   5760
      OrigTop         =   150
      OrigRight       =   6000
      OrigBottom      =   540
      SyncBuddy       =   -1  'True
      Wrap            =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin VB.Label lbl 
      Caption         =   "Right-click almost anywhere to get a context menu"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   636
      Index           =   3
      Left            =   5280
      TabIndex        =   29
      Top             =   3480
      Width           =   2616
   End
   Begin VB.Label lbl 
      Caption         =   "Destination:"
      Height          =   330
      Index           =   1
      Left            =   1995
      TabIndex        =   8
      Top             =   3765
      Width           =   1275
   End
   Begin VB.Label lblDst 
      Height          =   1380
      Left            =   1995
      TabIndex        =   13
      Top             =   4605
      UseMnemonic     =   0   'False
      Width           =   3795
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblSpecLoc 
      Height          =   810
      Left            =   5715
      TabIndex        =   25
      Top             =   825
      UseMnemonic     =   0   'False
      Width           =   2745
      WordWrap        =   -1  'True
   End
   Begin VB.Label lbl 
      Caption         =   "System Image Lists"
      Height          =   288
      Index           =   4
      Left            =   6036
      TabIndex        =   24
      Top             =   2616
      Width           =   2196
   End
   Begin VB.Image imgSSpecLoc 
      Height          =   240
      Left            =   7332
      Top             =   156
      Width           =   240
   End
   Begin VB.Image imgLSpecLoc 
      Height          =   480
      Left            =   6732
      Top             =   156
      Width           =   480
   End
   Begin VB.Label lbl 
      Caption         =   "System Images"
      Height          =   255
      Index           =   2
      Left            =   45
      TabIndex        =   23
      Top             =   -255
      Width           =   2160
   End
   Begin VB.Image imgLSysIcon 
      Height          =   480
      Left            =   6732
      Top             =   2055
      Width           =   480
   End
   Begin VB.Image imgSSysIcon 
      Height          =   240
      Left            =   7332
      Top             =   2055
      Width           =   240
   End
   Begin VB.Image imgLDIcon 
      Height          =   480
      Left            =   1995
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   480
   End
   Begin VB.Image imgSDIcon 
      Height          =   240
      Left            =   2565
      Top             =   3240
      Width           =   240
   End
   Begin VB.Image imgSDIconSel 
      Height          =   240
      Left            =   3504
      Top             =   3240
      Width           =   240
   End
   Begin VB.Image imgLDIconSel 
      Height          =   480
      Left            =   2925
      Top             =   3240
      Width           =   480
   End
   Begin VB.Image imgSDIconOpen 
      Height          =   240
      Left            =   4455
      Top             =   3240
      Width           =   240
   End
   Begin VB.Image imgLDIconOpen 
      Height          =   480
      Left            =   3870
      Top             =   3240
      Width           =   480
   End
   Begin VB.Image imgLIconLink 
      Height          =   480
      Left            =   3870
      Top             =   156
      Width           =   480
   End
   Begin VB.Image imgSIconLink 
      Height          =   240
      Left            =   4455
      Top             =   156
      Width           =   240
   End
   Begin VB.Image imgLIconSel 
      Height          =   480
      Left            =   2925
      Top             =   156
      Width           =   480
   End
   Begin VB.Image imgSIconSel 
      Height          =   240
      Left            =   3504
      Top             =   156
      Width           =   240
   End
   Begin VB.Image imgSIcon 
      Height          =   240
      Left            =   2565
      Top             =   156
      Width           =   240
   End
   Begin VB.Image imgLIcon 
      Height          =   480
      Left            =   1995
      Top             =   156
      Width           =   480
   End
   Begin VB.Label lblSrc 
      Height          =   1545
      Left            =   1995
      TabIndex        =   10
      Top             =   1575
      UseMnemonic     =   0   'False
      Width           =   3795
      WordWrap        =   -1  'True
   End
   Begin VB.Label lbl 
      Caption         =   "Source File:"
      Height          =   330
      Index           =   0
      Left            =   1995
      TabIndex        =   5
      Top             =   840
      Width           =   1275
   End
End
Attribute VB_Name = "FTestWhizzyStuff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private afOpFlags As Long
Private hLSysImages As Long, cLSysImages As Long
Private hSSysImages As Long, cSSysImages As Long
Private fiFolder As New CFileInfo
Attribute fiFolder.VB_VarHelpID = -1
Private fiSrc As New CFileInfo, fiDst As New CFileInfo
Private WithEvents ico As CTrayIcon
Attribute ico.VB_VarHelpID = -1
Private sTmpSrc As String

Private sExeName As String

Private Sub Form_Load()
    
    If Not HasShell Then
        MsgBox "This program requires the enhanced Windows interface"
        End
    End If
    If IsExe Then
        sExeName = App.Path & sBSlash & App.EXEName & ".EXE"
    Else
        sExeName = App.Path & sBSlash & App.EXEName & ".VBP"
    End If
    
    ' Create temporary file for initial source
    sTmpSrc = Environ$("WINDIR") & "\Cathistan.bas"
    SaveFileText sTmpSrc, "' The land of Basic"
    txtSrc = sTmpSrc
    NewSource
    ' Set destination to temp directory
    txtDst = GetTempDir
    NewDestination
    
    ' Set up UpDown control for special folders
    udSpecLoc.Min = CSIDL_DESKTOP
    ' Largest reasonable maximum, which may change depending on OS
    udSpecLoc.Max = 60
    udSpecLoc = 60
    NewSpecialFolder -1
    
    ' Set up UpDown control for system images
    hLSysImages = GetSysImageList(cLSysImages, True)
    hSSysImages = GetSysImageList(cSSysImages, False)
    udSysIcon.Min = 0
    udSysIcon.Max = cLSysImages - 1
    udSysIcon = 0
    DrawSysIcons udSysIcon
    
    ' Set up tray icon
    Set ico = New CTrayIcon
    ico.Create Me.hWnd, imgSSpecLoc.Picture, "What the heck"
                   
End Sub

Private Sub Form_Unload(Cancel As Integer)
    BugMessage "Unloading"
    Set ico = Nothing
    If ExistFileDir(sTmpSrc) Then Kill sTmpSrc
End Sub

Private Sub Form_Terminate()
    BugMessage "Terminating"
End Sub

' Handle source and destination

Private Sub cmdSrc_Click()
    Dim sFile As String, f As Boolean
    Static sDir As String
    If sDir = sEmpty Then sDir = CurDir$
    
    f = VBGetOpenFileName(FileName:=sFile, _
                          InitDir:=sDir, _
                          flags:=OFN_FILEMUSTEXIST Or OFN_HIDEREADONLY, _
                          Filter:="All files | *.*")
    If f And sFile <> sEmpty Then txtSrc.Text = sFile
    sDir = GetFileDir(sFile)
    NewSource
End Sub

Private Sub txtSrc_LostFocus()
    NewSource
End Sub

Private Sub NewSource()
With fiSrc
    Dim s As String, sOld As String
    On Error GoTo FailNewSource
    sOld = fiSrc
    ' Assign a file name to a file information object
    fiSrc = txtSrc.Text
    ' Get back all the information about the file
    s = s & "Display name: " & .DisplayName & sCrLf
    s = s & "Type name: " & .TypeName & sCrLf
    s = s & "Length: " & .Length & " bytes" & sCrLf
    s = s & "Created: " & .Created & sCrLf
    s = s & "Last modified: " & .Modified & sCrLf
    s = s & "Last accessed: " & .Accessed & sCrLf
    lblSrc.Caption = s
    Set imgLIcon.Picture = .ShellIcon()
    Set imgSIcon.Picture = .SmallIcon()
    Set imgLIconSel.Picture = .ShellIcon(SHGFI_SELECTED)
    Set imgSIconSel.Picture = .SmallIcon(SHGFI_SELECTED)
    Set imgLIconLink.Picture = .ShellIcon(SHGFI_LINKOVERLAY)
    Set imgSIconLink.Picture = .SmallIcon(SHGFI_LINKOVERLAY)
    Exit Sub
FailNewSource:
    txtSrc = sOld
    fiSrc = sOld
    MsgBox "Can't set new source"
End With
End Sub

Private Sub cmdDst_Click()
    Dim sDisplay As String, sResult As String
    Static sStart  As String
    If sStart = sEmpty Then sStart = GetTempDir
    sResult = BrowseForFolder(Owner:=hWnd, DisplayName:=sDisplay, _
                              Options:=BIF_RETURNONLYFSDIRS, _
                              Title:="Select directory:", _
                              Root:=CSIDL_DRIVES, _
                              Default:=sStart)
    If sResult <> sEmpty Then txtDst = sResult
    sStart = sResult
    NewDestination
End Sub

Private Sub txtDst_LostFocus()
    NewDestination
End Sub

Private Sub NewDestination()
With fiDst
    Dim s As String, sOld As String
    On Error GoTo FailNewDestination
    sOld = fiDst
    fiDst = txtDst.Text
    s = s & "Display name: " & .DisplayName & sCrLf
    s = s & "Type name: " & .TypeName & sCrLf
    ' Creation and last accessed times invalid for directories
    s = s & "Last modified: " & .Modified & sCrLf
    lblDst.Caption = s
    Set imgLDIcon.Picture = .ShellIcon()
    Set imgSDIcon.Picture = .SmallIcon()
    Set imgLDIconSel.Picture = .ShellIcon(SHGFI_SELECTED)
    Set imgSDIconSel.Picture = .SmallIcon(SHGFI_SELECTED)
    Set imgLDIconOpen.Picture = .ShellIcon(SHGFI_OPENICON)
    Set imgSDIconOpen.Picture = .SmallIcon(SHGFI_OPENICON)
    Exit Sub
FailNewDestination:
    txtDst = sOld
    fiDst = sOld
    MsgBox "Can't set new destination"
End With
End Sub

' Handle command buttons

Private Sub cmdMove_Click()
    SetOpFlags
    MoveAnyFile txtSrc.Text, txtDst.Text, afOpFlags
    txtSrc.Text = txtDst.Text & GetFileBaseExt(txtSrc.Text)
End Sub

Private Sub cmdCopy_Click()
    SetOpFlags
    CopyAnyFile txtSrc, txtDst, afOpFlags
End Sub

Private Sub cmdDelete_Click()
    SetOpFlags
    DeleteAnyFile txtSrc.Text, afOpFlags
    txtSrc.Text = sEmpty
End Sub

Private Sub cmdRename_Click()
    Dim s As String
    s = InputBox("New name for " & txtSrc.Text)
    SetOpFlags
    RenameAnyFile txtSrc.Text, s, afOpFlags
End Sub

Private Sub SetOpFlags()
    afOpFlags = 0
    If chkMultiDest.Value = vbChecked Then
        afOpFlags = afOpFlags Or FOF_MULTIDESTFILES
    End If
    If chkSilent.Value = vbChecked Then
        afOpFlags = afOpFlags Or FOF_SILENT
    End If
    If chkRename.Value = vbChecked Then
        afOpFlags = afOpFlags Or FOF_RENAMEONCOLLISION
    End If
    If chkNoConfirm.Value = vbChecked Then
        afOpFlags = afOpFlags Or FOF_NOCONFIRMATION
    End If
    If chkUndo.Value = vbChecked Then
        afOpFlags = afOpFlags Or FOF_ALLOWUNDO
    End If
    If chkFileOnly.Value = vbChecked Then
        afOpFlags = afOpFlags Or FOF_FILESONLY
    End If
    If chkProgress.Value = vbChecked Then
        afOpFlags = afOpFlags Or FOF_SIMPLEPROGRESS
    End If
    If chkMkDir.Value = vbChecked Then
        afOpFlags = afOpFlags Or FOF_NOCONFIRMMKDIR
    End If
    
End Sub

Private Sub cmdDocs_Click()
    AddToRecentDocs txtSrc.Text
End Sub

Private Sub cmdClearDoc_Click()
    ClearRecentDocs
End Sub

' Handle special folder locations

Private Sub udSpecLoc_UpClick()
    ' Increment on every UpClick
    NewSpecialFolder 1
End Sub

Private Sub udSpecLoc_DownClick()
    ' Decrement on every DownClick
    NewSpecialFolder -1
End Sub

Private Sub NewSpecialFolder(ByVal iInc As Long)
    Do
        On Error Resume Next
        fiFolder = udSpecLoc                ' Folder to CFileInfo object
        If Err = 0 Then Exit Do
        udSpecLoc.Value = udSpecLoc + iInc  ' Skip missing numbers
        ' Wrap property doesn't work on assignment
        If udSpecLoc = udSpecLoc.Min Then udSpecLoc = udSpecLoc.Max
        If udSpecLoc = udSpecLoc.Max Then udSpecLoc = udSpecLoc.Min
    Loop
    lblSpecLoc.Caption = "Special folder: " & fiFolder.DisplayName
    imgLSpecLoc = fiFolder.ShellIcon
    imgSSpecLoc = fiFolder.SmallIcon
End Sub

' Handle system icons
Private Sub udSysIcon_UpClick()
    DrawSysIcons udSysIcon
End Sub

Private Sub udSysIcon_DownClick()
    DrawSysIcons udSysIcon
End Sub

Private Sub DrawSysIcons(i As Integer)
    imgLSysIcon.Picture = GetSysIcon(hLSysImages, i)
    imgSSysIcon.Picture = GetSysIcon(hSSysImages, i)
End Sub

' Handle tray icon messages

Private Sub ico_TrayMessage(ByVal Message As Long, ByVal ID As Long)
    ' Handle left double click or right click
    Select Case Message
    Case WM_LBUTTONDBLCLK
        MsgBox "You double-clicked me"
    Case WM_RBUTTONUP
        MsgBox "You right clicked me"
    End Select
End Sub

' Right click for context menus

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button And 2 Then ContextPopMenu hWnd, sExeName, x, y
End Sub

Private Sub cmdDst_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button And 2 Then ContextPopMenu cmdDst.hWnd, CStr(fiDst), x, y
End Sub

Private Sub imgLDIcon_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
With imgLDIcon
    If Button And 2 Then ContextPopMenu hWnd, CStr(fiDst), x + .Left, y + .Top
End With
End Sub

Private Sub imgLDIconOpen_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
With imgLDIconOpen
    If Button And 2 Then ContextPopMenu hWnd, CStr(fiDst), x + .Left, y + .Top
End With
End Sub

Private Sub imgLDIconSel_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
With imgLDIconSel
    If Button And 2 Then ContextPopMenu hWnd, CStr(fiDst), x + .Left, y + .Top
End With
End Sub

Private Sub imgLIcon_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
With imgLIcon
    If Button And 2 Then ContextPopMenu hWnd, CStr(fiSrc), x + .Left, y + .Top
End With
End Sub
Private Sub imgLIconLink_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
With imgLIconLink
    If Button And 2 Then ContextPopMenu hWnd, CStr(fiSrc), x + .Left, y + .Top
End With
End Sub

Private Sub imgLIconSel_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
With imgLIconSel
    If Button And 2 Then ContextPopMenu hWnd, CStr(fiSrc), x + .Left, y + .Top
End With
End Sub

Private Sub imgLSpecLoc_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
With imgLSpecLoc
    If Button And 2 Then ContextPopMenu hWnd, fiFolder, x + .Left, y + .Top
End With
End Sub

Private Sub imgSDIcon_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
With imgSDIcon
    If Button And 2 Then ContextPopMenu hWnd, CStr(fiDst), x + .Left, y + .Top
End With
End Sub

Private Sub imgSDIconOpen_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
With imgSDIconOpen
    If Button And 2 Then ContextPopMenu hWnd, CStr(fiDst), x + .Left, y + .Top
End With
End Sub

Private Sub imgSDIconSel_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
With imgSDIconSel
    If Button And 2 Then ContextPopMenu hWnd, CStr(fiDst), x + .Left, y + .Top
End With
End Sub

Private Sub imgSIcon_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
With imgSIcon
    If Button And 2 Then ContextPopMenu hWnd, CStr(fiSrc), x + .Left, y + .Top
End With
End Sub

Private Sub imgSIconLink_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
With imgSIconLink
    If Button And 2 Then ContextPopMenu hWnd, CStr(fiSrc), x + .Left, y + .Top
End With
End Sub

Private Sub imgSIconSel_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
With imgSIconSel
    If Button And 2 Then ContextPopMenu hWnd, CStr(fiSrc), x + .Left, y + .Top
End With
End Sub

Private Sub imgSSpecLoc_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
With imgSSpecLoc
    If Button And 2 Then ContextPopMenu hWnd, fiFolder, x + .Left, y + .Top
End With
End Sub

Private Sub txtSpecLoc_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button And 2 Then ContextPopMenu txtSpecLoc.hWnd, fiFolder, x, y
End Sub

Private Sub udSpecLoc_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button And 2 Then ContextPopMenu udSpecLoc.hWnd, fiFolder, x, y
End Sub

Private Sub lblDst_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
With lblDst
    If Button And 2 Then ContextPopMenu hWnd, CStr(fiDst), x + .Left, y + .Top
End With
End Sub

Private Sub lblSrc_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
With lblSrc
    If Button And 2 Then ContextPopMenu hWnd, CStr(fiSrc), x + .Left, y + .Top
End With
    If Button And 2 Then ContextPopMenu hWnd, CStr(fiSrc), x, y
End Sub

Private Sub txtDst_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button And 2 Then ContextPopMenu txtDst.hWnd, CStr(fiDst), x, y
End Sub

Private Sub cmdSrc_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button And 2 Then ContextPopMenu cmdSrc.hWnd, fiSrc, x, y
End Sub

Private Sub txtSrc_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button And 2 Then ContextPopMenu txtSrc.hWnd, CStr(fiSrc), x, y
End Sub

