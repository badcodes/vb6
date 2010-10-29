VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FWatch 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "WinWatch"
   ClientHeight    =   6036
   ClientLeft      =   4080
   ClientTop       =   2436
   ClientWidth     =   9360
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   7.8
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Winwatch.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6036
   ScaleWidth      =   9360
   Begin MSComctlLib.TreeView tvwWin 
      Height          =   1728
      Left            =   4932
      TabIndex        =   30
      Top             =   2544
      Width           =   4260
      _ExtentX        =   7514
      _ExtentY        =   3048
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   529
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   6
      Appearance      =   1
   End
   Begin VB.ListBox lstTopWin 
      Height          =   1392
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   29
      Top             =   4530
      Width           =   2175
   End
   Begin VB.ListBox lstResource 
      Height          =   1392
      Left            =   7080
      Sorted          =   -1  'True
      TabIndex        =   28
      Top             =   4532
      Width           =   2175
   End
   Begin VB.ListBox lstModule 
      Height          =   1392
      Left            =   4760
      TabIndex        =   27
      Top             =   4534
      Width           =   2175
   End
   Begin VB.ListBox lstProcess 
      Height          =   1392
      Left            =   2440
      Sorted          =   -1  'True
      TabIndex        =   26
      Top             =   4536
      Width           =   2175
   End
   Begin VB.CheckBox chkBlank 
      Caption         =   "Show Blank"
      Height          =   255
      Left            =   1200
      TabIndex        =   21
      Top             =   3756
      Width           =   1695
   End
   Begin VB.CheckBox chkOwned 
      Caption         =   "Show Owned"
      Height          =   255
      Left            =   1200
      TabIndex        =   20
      Top             =   3516
      Width           =   1695
   End
   Begin VB.PictureBox bstMenu 
      Height          =   480
      Left            =   8445
      ScaleHeight     =   432
      ScaleWidth      =   1152
      TabIndex        =   19
      Top             =   6840
      Width           =   1200
   End
   Begin VB.CheckBox chkFilter 
      Caption         =   "Filter Resources"
      Height          =   255
      Left            =   3000
      TabIndex        =   18
      Top             =   3996
      Value           =   1  'Checked
      Width           =   1845
   End
   Begin VB.PictureBox pbDump 
      AutoRedraw      =   -1  'True
      Height          =   396
      Left            =   165
      ScaleHeight     =   348
      ScaleWidth      =   504
      TabIndex        =   17
      Top             =   3720
      Visible         =   0   'False
      Width           =   552
   End
   Begin VB.CommandButton cmdDump 
      Caption         =   "&Dump"
      Height          =   372
      Left            =   120
      TabIndex        =   16
      Top             =   2770
      Width           =   972
   End
   Begin VB.CommandButton cmdLogFile 
      Caption         =   "&Log File"
      Height          =   372
      Left            =   120
      TabIndex        =   13
      Top             =   360
      Width           =   972
   End
   Begin VB.CheckBox chkInvisible 
      Caption         =   "Show Invisible"
      Height          =   255
      Left            =   1200
      TabIndex        =   12
      Top             =   3996
      Width           =   1695
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   372
      Left            =   120
      TabIndex        =   9
      Top             =   842
      Width           =   972
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "&Refresh"
      Height          =   372
      Left            =   120
      TabIndex        =   5
      Top             =   2288
      Width           =   972
   End
   Begin VB.CommandButton cmdPoint 
      Caption         =   "&Point"
      Height          =   372
      Left            =   120
      TabIndex        =   1
      Top             =   1324
      Width           =   972
   End
   Begin VB.CommandButton cmdActivate 
      Caption         =   "&Activate"
      Height          =   372
      Left            =   120
      TabIndex        =   2
      Top             =   1806
      Width           =   972
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   372
      Left            =   120
      TabIndex        =   3
      Top             =   3255
      Width           =   972
   End
   Begin VB.PictureBox pbResource 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      DragIcon        =   "Winwatch.frx":0CFA
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3264
      Left            =   1200
      ScaleHeight     =   3264
      ScaleWidth      =   3624
      TabIndex        =   4
      Top             =   216
      Width           =   3624
   End
   Begin VB.Image imgCloud 
      Height          =   1536
      Left            =   2976
      Picture         =   "Winwatch.frx":0E4C
      Top             =   1752
      Visible         =   0   'False
      Width           =   1536
   End
   Begin VB.Label lbl 
      Caption         =   "Module:"
      Height          =   228
      Index           =   4
      Left            =   7008
      TabIndex        =   25
      Top             =   1236
      Width           =   1896
   End
   Begin VB.Label lblMod 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   768
      Left            =   6984
      TabIndex        =   24
      Top             =   1476
      Width           =   2028
   End
   Begin VB.Label lblProc 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Left            =   4968
      TabIndex        =   23
      Top             =   1464
      Width           =   2028
   End
   Begin VB.Label lbl 
      Caption         =   "Process:"
      Height          =   228
      Index           =   2
      Left            =   4956
      TabIndex        =   22
      Top             =   1224
      Width           =   1896
   End
   Begin VB.Label lbl 
      Caption         =   "Window Hierarchy:"
      Height          =   228
      Index           =   3
      Left            =   4956
      TabIndex        =   15
      Top             =   2304
      Width           =   2220
   End
   Begin VB.Label lbl 
      Caption         =   "Window:"
      Height          =   228
      Index           =   1
      Left            =   4932
      TabIndex        =   14
      Top             =   0
      Width           =   4212
   End
   Begin VB.Label lblWin 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1044
      Left            =   4944
      TabIndex        =   10
      Top             =   228
      Width           =   4116
   End
   Begin VB.Label lblResource 
      Caption         =   "Resources:"
      Height          =   252
      Left            =   7080
      TabIndex        =   11
      Top             =   4320
      Width           =   1140
   End
   Begin VB.Label lblProcess 
      Caption         =   "Processes:"
      Height          =   255
      Left            =   2440
      TabIndex        =   8
      Top             =   4320
      Width           =   1935
   End
   Begin VB.Label lblModule 
      Caption         =   "Modules:"
      Height          =   255
      Left            =   4760
      TabIndex        =   7
      Top             =   4320
      Width           =   2175
   End
   Begin VB.Label lbl 
      Caption         =   "Top Windows:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   4320
      Width           =   2175
   End
   Begin VB.Label lblMsg 
      Caption         =   "Resource Information:"
      Height          =   255
      Left            =   1200
      TabIndex        =   0
      Top             =   0
      Width           =   3255
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuLogFile 
         Caption         =   "&Log File"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnuPoint 
         Caption         =   "&Point"
      End
      Begin VB.Menu mnuActivate 
         Caption         =   "&Activate"
      End
      Begin VB.Menu mnuRefresh 
         Caption         =   "&Refresh"
      End
      Begin VB.Menu mnuDump 
         Caption         =   "&Dump"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "FWatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private idProcCur As Long
Private hModCur As Long
Private hModFree As Long
Private hTopWndCur As Long
Private hWndCur As Long
Private hInstCur As Long
Private sModCur As String
Private hResourceCur As Long
Private hResourceLast As Long
Private fCapture As Boolean
Private ordResourceLast As Integer
Private ordPointerLast As Integer
Private dxPicMax As Long, dyPicMax As Long
Private nFileCur As Integer

Const sMsg = "Resource Information:"

Public Enum EUpdateType
    eutTopWindow
    eutWindow
    eutProcess
    eutModule
End Enum

' Constants for accessing icon directory and entry structures
Enum EIconDirEntryImage
    ' Group Directory
    wReserved = 0
    wType = 2
    wCount = 4
    entFirst = 6
        ' Icon Group Entry
        bWidth = 0
        bHeight = 1
        bColorCount = 2
        bReserved = 3
        wPlanes = 4
        wBitCount = 6
        dwBytesInRes = 8
        wID = 12
        cEntrySize = 14
        ' Cursor Group Entry
        wWidth = 0
        wHeight = 2
        ' Rest same as Icon
End Enum

' Flag to prevent recursion in
Private fInClick As Boolean

Private Sub Form_Load()
    
    Dim hWndOther As Long
    hWndOther = GetFirstInstWnd(Me.hWnd)
    If hWndOther <> hNull Then
        ' Uncomment this line for debugging
        'MsgBox "Activating first instance"
        SetForegroundWindow hWndOther
        End
    End If
    dxPicMax = pbResource.Width
    dyPicMax = pbResource.Height
    ChDrive App.Path
    ChDir App.Path
    PaletteMode = vbPaletteModeCustom
    Palette = pbResource.Picture
       
    Show
    RefreshAllLists

End Sub

Private Sub Form_Paint()
    pbResource.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    BugTerm
    ClearResource
End Sub

Private Sub RefreshAllLists(Optional hWnd As Long)
    ' Prevent calling again until this one finishes
    Static fInside As Integer
    If fInside Then Exit Sub
    fInside = True
    
    RefreshFullWinList False
    RefreshTopWinList
    RefreshProcessList
    
    ' Update entire display
    fInClick = True
    hWndCur = 0: hTopWndCur = 0: idProcCur = 0
    hWnd = IIf(hWnd, hWnd, lstTopWin.ItemData(0))
    UpdateDisplay eutWindow, hWnd
    fInClick = False
    
    fInside = False
End Sub

Private Sub RefreshFullWinList(fLogFile As Boolean)
    Const sLog = "WINLIST.TXT"

    Call LockWindowUpdate(tvwWin.hWnd)
    tvwWin.Nodes.Clear
    HourGlass Me
    If fLogFile Then
        lblMsg.Caption = "Creating log file " & sLog & "..."
        nFileCur = FreeFile
        Open sLog For Output As nFileCur
        Print #nFileCur, sEmpty
        Print #nFileCur, "Window List " & sCrLf
        Dim helperFile As CWindowToFile
        Set helperFile = New CWindowToFile
        helperFile.FileNumber = nFileCur
        Call IterateChildWindows(-1, GetDesktopWindow(), helperFile)
        Close nFileCur
    Else
        lblMsg.Caption = "Building window list..."
        Dim helperForm As CWindowToForm
        Set helperForm = New CWindowToForm
        Set helperForm.TreeViewControl = tvwWin
        helperForm.ShowInvisible = chkInvisible
        Call IterateChildWindows(-1, GetDesktopWindow(), helperForm)
    End If
    lblMsg.Caption = sMsg
    HourGlass Me
    tvwWin.Refresh
    Call LockWindowUpdate(hNull)
End Sub

Private Sub RefreshTopWinList()
    Dim sTitle As String, hWnd As Long
    
    SetRedraw lstTopWin, False
    lstTopWin.Clear
    ' Get first top-level window
    hWnd = GetWindow(GetDesktopWindow(), GW_CHILD)
    BugAssert hWnd <> hNull
    ' Iterate through remaining windows
    Do While hWnd <> hNull
        sTitle = WindowTextLineFromWnd(hWnd)
        ' Determine whether to display titled, visible, and unowned
        If IsVisibleTopWnd(hWnd, chkBlank, _
                           chkInvisible, chkOwned) Then
            lstTopWin.AddItem sTitle
            lstTopWin.ItemData(lstTopWin.NewIndex) = hWnd
        End If
        ' Get next child
        hWnd = GetWindow(hWnd, GW_HWNDNEXT)
    Loop
    SetRedraw lstTopWin, True
End Sub

Private Sub RefreshProcessList()
    Dim processes As CVector, process As CProcess, i As Long
    Set processes = CreateProcessList
    SetRedraw lstProcess, False
    lstProcess.Clear
    For i = 1 To processes.Last
        lstProcess.AddItem processes(i).EXEName
        lstProcess.ItemData(lstProcess.NewIndex) = processes(i).ID
    Next
    SetRedraw lstProcess, True
End Sub

Private Sub RefreshModuleList(idProc As Long)
    Dim modules As CVector, module As CModule, i As Long
    Set modules = CreateModuleList(idProc)
    
    ' Illustrate three ways to prevent visible window update
#Const ordQuiet = 0
#If ordQuiet = 0 Then
    SetRedraw lstModule, False
#ElseIf ordQuiet = 1 Then
    Call LockWindowUpdate(lstModule.hWnd)
#ElseIf ordQuiet = 2 Then
    lstModule.Visible = False
#End If
    lstModule.Clear

    ' Add module names and handles
    For i = 1 To modules.Last
        lstModule.AddItem modules(i).ExeFile
        lstModule.ItemData(lstModule.NewIndex) = modules(i).Handle
    Next
    ' Look up main executable file
    lstModule.ListIndex = LookupItem(lstModule, ExeNameFromProcID(idProc))
    If lstModule.ListIndex = -1 Then lstModule.ListIndex = 0
    
#If ordQuiet = 0 Then
    SetRedraw lstModule, True
#ElseIf ordQuiet = 1 Then
    Call LockWindowUpdate(hNull)
#ElseIf ordQuiet = 2 Then
    lstModule.Visible = True
#End If

End Sub
 
Private Sub lstTopWin_DblClick()
#Const fWindowsWay = 0
#If fWindowsWay Then
    SetForegroundWindow hTopWndCur
#Else
    ' Ignore errors
    On Error Resume Next
    AppActivate lstTopWin.Text
    If Err Then BugMessage "AppActivate error: " & Err
#End If
End Sub

Private Sub lstTopWin_Click()

    ' Module-level flag to prevent circular references
    If fInClick Then Exit Sub
    fInClick = True
    
    ' Look up window handle
    Dim hWnd As Long
    hWnd = lstTopWin.ItemData(lstTopWin.ListIndex)
    UpdateDisplay eutTopWindow, hWnd
    
    fInClick = False
End Sub

Private Sub lstProcess_Click()

    ' Module-level flag to prevent circular references
    If fInClick Then Exit Sub
    fInClick = True

    ' Load process ID
    Dim idProc As Long
    idProc = lstProcess.ItemData(lstProcess.ListIndex)
    UpdateDisplay eutProcess, idProc
        
    fInClick = False
End Sub

Private Sub lstModule_Click()

    ' Module-level flag to prevent circular references
    If fInClick Then Exit Sub
    fInClick = True
    
    UpdateDisplay eutModule, lstModule.ItemData(lstModule.ListIndex)
    
    fInClick = False
End Sub

Private Sub chkFilter_Click()
    UpdateResources hModCur
End Sub

Private Sub chkInvisible_Click()
    RefreshAllLists
End Sub

Private Sub chkOwned_Click()
    RefreshAllLists
End Sub

Private Sub chkBlank_Click()
    RefreshAllLists
End Sub

Private Sub cmdActivate_Click()
    SetForegroundWindow hWndCur
End Sub

Private Sub cmdDump_Click()
    Dim hDCCur As Long, hWndOld As Long
    Dim tim As Double
    Dim RECT As RECT, dx As Long, dy As Long
    
    ' Save current window, and switch to capture window
    hWndOld = GetActiveWindow()
    SetForegroundWindow hWndCur
    ' Give window time to repaint
    tim = Timer + 0.5
    Do
        DoEvents
    Loop Until Timer >= tim
    ' Borrow window DC
    hDCCur = GetWindowDC(hWndCur)
    Call GetWindowRect(hWndCur, RECT)
    dx = RECT.Right - RECT.Left + 2: dy = RECT.bottom - RECT.Top + 2
    ' Blit window DC to hidden picture box
    With pbDump
        .Width = Screen.TwipsPerPixelX * dx
        .Height = Screen.TwipsPerPixelY * dy
        Call BitBlt(.hDC, 0, 0, dx, dy, hDCCur, 0, 0, vbSrcCopy)
        ' Copy from DC to Picture
        .Picture = .Image
    End With
    ' Give DC back
    Call ReleaseDC(hWndCur, hDCCur)
    SetForegroundWindow hWndOld
    ' Save Picture property in file
    Dim sFile As String, sDirCur As String
    sDirCur = CurDir
    If VBGetSaveFileName(FileName:="*.BMP", _
                         FileTitle:=sFile, _
                         InitDir:=sDirCur, _
                         DlgTitle:="Save Window As", _
                         Filter:="Bitmaps (*.BMP) | *.BMP)", _
                         DefaultExt:="BMP") Then
            SavePicture pbDump.Picture, sFile
    End If
    ChDir sDirCur
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdLogFile_Click()
    RefreshFullWinList True
End Sub

Private Sub cmdPoint_Click()
    If cmdPoint.Caption = "&Point" Then
        fCapture = True
        cmdPoint.Caption = "End &Point"
        Call SetCapture(Me.hWnd)
        lblMsg.Caption = "Move mouse for window information"
    Else
        fCapture = False
        cmdPoint.Caption = "&Point"
        ReleaseCapture
        lblMsg.Caption = sMsg
    End If
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, _
                           X As Single, Y As Single)
    If fCapture Then
        Dim pt As POINTL
        Dim hWnd As Long
        Dim idProc As Long
        pt.X = X / Screen.TwipsPerPixelX
        pt.Y = Y / Screen.TwipsPerPixelY
        ClientToScreen Me.hWnd, pt
        hWnd = WindowFromPoint(pt.X, pt.Y)
        If hWnd <> hNull Then
            ' Turn point mode off
            fCapture = False
            cmdPoint.Caption = "&Point"
            On Error Resume Next
            Dim nodX As Node
            Set nodX = tvwWin.Nodes.Item("W" & hWnd)
            ' If window isn't in list, refresh the display
            fInClick = True
            If nodX Is Nothing Then
                RefreshAllLists hWnd
            Else
                UpdateDisplay eutWindow, hWnd
                lblMsg.Caption = sMsg
            End If
            fInClick = False
        End If
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, _
                           X As Single, Y As Single)
    If fCapture Then
        Dim pt As POINTL, hWnd As Long
        Static hWndLast As Long
        ' Set point and convert it to screen coordinates
        pt.X = X / Screen.TwipsPerPixelX
        pt.Y = Y / Screen.TwipsPerPixelY
        ClientToScreen Me.hWnd, pt
        ' Find window under it
        hWnd = WindowFromPoint(pt.X, pt.Y)
        ' Update display only if window has changed
        If hWnd <> hWndLast Then
            lblWin.Caption = GetWndInfo(hWnd)
            hWndLast = hWnd
        End If
    End If
End Sub

Private Sub cmdRefresh_Click()
    ' Maintain currently selected item
    RefreshAllLists lstTopWin.ItemData(lstTopWin.ListIndex)
End Sub

Private Sub cmdSave_Click()
    Dim sDirCur As String, sFileTitle As String, sFile As String
    Dim sDlgTitle As String, sFilter As String, sDefaultExt As String
    sDirCur = CurDir$

    Select Case ordResourceLast
    Case RT_BITMAP
        sFile = "*.BMP"
        sDlgTitle = "Save Bitmap As"
        sFilter = "Bitmaps (*.BMP) | *.BMP)"
        sDefaultExt = "BMP"
                   
    Case RT_ICON, RT_GROUP_ICON
        sFile = "*.ICO"
        sDlgTitle = "Save Icon As"
        sFilter = "Icons (*.ICO) | *.ICO)"
        sDefaultExt = "ICO"
        
    Case RT_CURSOR, RT_GROUP_CURSOR
        sFile = "*.CUR"
        sDlgTitle = "Save Cursor As"
        sFilter = "Cursors (*.CUR) | *.CUR)"
        sDefaultExt = "CUR"
    
    Case Else
        Exit Sub
    End Select
    If VBGetSaveFileName(FileName:=sFile, _
                         FileTitle:=sFileTitle, _
                         InitDir:=sDirCur, _
                         DlgTitle:=sDlgTitle, _
                         Filter:=sFilter, _
                         DefaultExt:=sDefaultExt) Then
        If ordResourceLast <> RT_CURSOR And ordResourceLast <> RT_GROUP_CURSOR Then
            SavePicture pbResource.Picture, sFileTitle
        End If
    End If
    ChDir sDirCur
End Sub

Private Sub mnuActivate_Click()
    cmdActivate_Click
End Sub

Private Sub mnuDump_Click()
    cmdDump_Click
End Sub

Private Sub mnuExit_Click()
    cmdExit_Click
End Sub

Private Sub mnuLogFile_Click()
    cmdLogFile_Click
End Sub

Private Sub mnuPoint_Click()
    cmdPoint_Click
End Sub

Private Sub mnuRefresh_Click()
    cmdRefresh_Click
End Sub

Private Sub mnuSave_Click()
    cmdSave_Click
End Sub

#If Win32 = 0 Then
Private Sub bstMenu_Message(MsgVal As Integer, wParam As Integer, _
                            lParam As Long, ReturnVal As Long)
    If MsgVal = WM_COMMAND Then
        ReturnVal = SendMessage(hTopWndCur, MsgVal, ByVal wParam, ByVal lParam)
    ElseIf MsgVal = WM_INITMENU Then
        SyncMenu hResourceCur, GetMenu(hTopWndCur)
        ReturnVal = 0&
    End If
End Sub
#End If

Private Sub lstResource_Click()
    Dim sType As String, sName As String, i As Integer

    sType = lstResource.Text
    BugAssert sType <> sEmpty
    ' Extract resource ID and type
    If Left$(sType, 1) = "0" Then
        ' Append # so Windows will recognize numbers as strings
        sName = "#" & Left$(sType, 5)
        sType = Trim$(Mid$(sType, 7))
    Else
        i = InStr(sType, " ")
        sName = Trim$(Left$(sType, i - 1))
        sType = Trim$(Mid$(sType, i + 1))
    End If
    
    ' Clear last resource and handle new one
    ClearResource
    pbResource.AutoRedraw = False
    If UCase$(sType) <> "BITMAP" Then
        BmpTile pbResource, imgCloud.Picture
    End If
    
    Select Case UCase$(sType)
    Case "CURSOR"
        ShowCursor hModCur, sName
    Case "GROUP_CURSOR", "GROUP CURSOR"
        ShowCursors hModCur, sName
    Case "BITMAP"
        ShowBitmap hModCur, sName
    Case "ICON"
        ShowIcon hModCur, sName
    Case "GROUP_ICON", "GROUP ICON"
        ShowIcons hModCur, sName
    Case "MENU"
        ShowMenu hModCur, sName
    Case "STRING", "STRINGTABLE"
        ShowString hModCur, sName
    Case "WAVE"
        PlayWave hModCur, sName
    'Case "AVI"
    '    PlayAvi hModCur, sName
    Case "FONTDIR", "FONT", "DIALOG", "ACCELERATOR"
        pbResource.Print sType & " selected"
    Case "VERSION"
        pbResource.Print GetVersionData(sModCur, 26)
    Case Else
        ShowData hModCur, sName, sType
    End Select
    pbResource.AutoRedraw = True
End Sub

Private Sub UpdateResources(ByVal hMod As Long)
    ' Turn on hourglass, turn off redrawing
    HourGlass Me
    Call LockWindowUpdate(lstResource.hWnd)
    lstResource.Clear
    
    Call EnumResourceTypes(hMod, AddressOf ResTypeProc, Me)
    
    Call LockWindowUpdate(hNull)
    HourGlass Me
End Sub

Sub SaveIcon(pb As PictureBox, sName As String)
#If 0 Then
    Static cbIconBits As Integer

    ' Set byte size of an icon only one time
    If cbIconBits = 0 Then
        Dim hWnd As Long, hDC As Long
        hWnd = GetDesktopWindow()
        hDC = GetDC(hWnd)
        If GetDeviceCaps(hDC, BITSPIXEL) = 8 Then
            cbIconBits = 1024
        Else
            cbIconBits = 512
        End If
        Call ReleaseDC(hWnd, hDC)
    End If

    ' Lock dummy icon so that we can write bits to it
    Dim pIcon As Long
    'pbTemp.Picture = pbBlank.Picture
    pIcon = GlobalLock(pbTemp.Picture)

    ' Copy bits from picture to icon dummy, skipping icon header
    Dim iBits As Long, cHeader   As Integer, cColors As Integer
    cHeader = 12
    cColors = 128
    iBits = GetBitmapBits(pb.Image, cbIconBits, pIcon + cHeader + cColors)
    pb.Refresh
    
    ' Unlock icon dummy
    Call GlobalUnlock(pbTemp.Picture)

    ' Save icon
    SavePicture pbTemp.Picture, sName
#End If
End Sub

Private Sub ClearResource()
    With pbResource

    Select Case ordResourceLast
    Case RT_MENU
        Call SetMenu(Me.hWnd, hResourceLast)
        Call DestroyMenu(hResourceCur)

    Case RT_GROUP_CURSOR, RT_CURSOR
        MousePointer = ordPointerLast

    Case RT_BITMAP
        
    Case RT_GROUP_ICON, RT_ICON, RT_STRING, RT_RCDATA
        ' No restore needed
        
    End Select
    
    .CurrentX = 0
    .CurrentY = 0
    ordResourceLast = 0
    hResourceCur = hNull
    hResourceLast = hNull
'    .Picture = LoadPicture()
    
End With
End Sub

Sub ShowBitmap(ByVal hMod As Long, sBitmap As String)
With pbResource
    
    Dim hPal As Long, hPal2 As Long
    ' Convert resource into bitmap handle
    hResourceCur = LoadBitmapPalette(hMod, sBitmap, hPal)
    If hResourceCur = hNull Then
        pbResource.Print "Can't load bitmap: " & sCrLf & sCrLf & _
                         WordWrap(ApiError(Err.LastDllError), 25)
        Exit Sub
    End If
    ' Convert hBitmap to Picture (clip anything larger than picture box)
    .Picture = BitmapToPicture(hResourceCur, hPal)
    ' Set the form palette to use this picture's palette
    Palette = .Picture
    ' Make sure palette is realized
    Refresh
    DoEvents
    ' Draw the palette
    DrawPalette pbResource, hPal, .Width, .Height * 0.1, 0, .Height * 0.9
    ' Record the type for cleanup
    ordResourceLast = RT_BITMAP
End With
End Sub

Sub ShowCursor(ByVal hMod As Long, sCursor As String)
    ' Get cursor handle
    hResourceCur = LoadImage(hMod, sCursor, IMAGE_CURSOR, 0, 0, 0)
    If hResourceCur <> hNull Then
        ordPointerLast = MousePointer
        MousePointer = vbCustom
        MouseIcon = CursorToPicture(hResourceCur)
        ordResourceLast = RT_CURSOR
        Call DrawIconEx(pbResource.hDC, 0, 0, hResourceCur, _
                        0, 0, 0, hNull, DI_NORMAL)
    Else
        pbResource.Print "Can't display cursor: " & sCrLf & sCrLf & _
                         WordWrap(ApiError(Err.LastDllError), 25)
    End If
End Sub

Sub ShowCursors(ByVal hMod As Long, sCursor As String)
    BugAssert (hMod <> hNull) And (sCursor <> sEmpty)
    
    ' Find the resource
    Dim hRes As Long, hmemRes As Long, cRes As Long, pRes As Long
    Dim abGroup() As Byte, abEntry() As Byte
    hRes = FindResourceStrId(hMod, sCursor, RT_GROUP_CURSOR)
    If hRes = hNull Then
        pbResource.Print "Can't display data: " & sCrLf & sCrLf & _
                         WordWrap(ApiError(Err.LastDllError), 25)
        Exit Sub
    End If
    pbResource.ScaleMode = vbPixels
    ' Allocate memory block, get size, get pointer, and allocate array
    hmemRes = LoadResource(hMod, hRes)
    cRes = SizeofResource(hMod, hRes)
    pRes = LockResource(hmemRes)
    ReDim abGroup(cRes)
    ' Copy memory block to bytes and free
    CopyMemory abGroup(0), ByVal pRes, cRes
    Call FreeResource(hmemRes)
    
    Dim cImage As Integer, i As Integer, iImage As Integer
    Dim dxCursor As Integer, dyCursor As Integer, hCursor As Long, s As String
    ' Validate entry
    BugAssert BytesToWord(abGroup, wType) = vbResCursor
    ' Get image count
    cImage = BytesToWord(abGroup, wCount)
    ' Set up first entry
    iImage = entFirst
    pbResource.CurrentX = 75
    pbResource.CurrentY = 0
    For i = 0 To cImage - 1
        ' Get size and colors
        dxCursor = abGroup(iImage + wWidth)
        dyCursor = abGroup(iImage + wHeight)
        ' For reasons unknown height always comes out twice real size,
        ' so since all cursors I've seen are square, reuse width as height
        s = dxCursor & "x" & dxCursor
        ' Find, load, size, allocate, and copy entry
        hRes = FindResourceIdId(hMod, BytesToWord(abGroup, iImage + wID), RT_CURSOR)
        BugAssert hRes
        hmemRes = LoadResource(hMod, hRes)
        cRes = SizeofResource(hMod, hRes)
        pRes = LockResource(hmemRes)
        ReDim abEntry(cRes)
        CopyMemory abEntry(0), ByVal pRes, cRes
        Call FreeResource(hmemRes)
        ' Create an Cursor from resource data
        hCursor = CreateIconFromResource(abEntry(0), cRes, False, &H30000)
        ' Draw Cursor and print description
        s = s & " (" & BytesToWord(abEntry, 0) & "," & _
                       BytesToWord(abEntry, 2) & ")"
        Call DrawIconEx(pbResource.hDC, 0, pbResource.CurrentY, hCursor, _
                        dxCursor, dxCursor, 0, hNull, DI_NORMAL)
        pbResource.Print s
        ' Move to next entry
        pbResource.CurrentY = pbResource.CurrentY + dxCursor
        pbResource.CurrentX = 75
        iImage = iImage + cEntrySize
    Next
    pbResource.ScaleMode = vbTwips
    ' Use the best cursor
    hResourceCur = LoadImage(hMod, sCursor, IMAGE_CURSOR, 0, 0, 0)
    BugAssert hResourceCur <> hNull
    ordPointerLast = MousePointer
    MousePointer = vbCustom
    MouseIcon = IconToPicture(hResourceCur)
    ordResourceLast = RT_CURSOR
    
End Sub

Sub ShowData(ByVal hMod As Long, sData As String, _
             Optional sDataType As String = "RCDATA")
    
    Dim hRes As Long, hmemRes As Long, cRes As Long
    Dim pRes As Long, abRes() As Byte
    If sDataType = "RCDATA" Then
        hRes = FindResourceStrId(hMod, sData, RT_RCDATA)
    Else
        hRes = FindResourceStrStr(hMod, sData, sDataType)
    End If
    If hRes = hNull Then
        pbResource.Print "Can't display data: " & sCrLf & sCrLf & _
                         WordWrap(ApiError(Err.LastDllError), 25)
        Exit Sub
    End If
    ' Allocate memory block and get its size
    hmemRes = LoadResource(hMod, hRes)
    cRes = SizeofResource(hMod, hRes)
    ' Don't dump more than 500 bytes
    If cRes > 500 Then cRes = 500
    ' Lock it to get pointer
    pRes = LockResource(hmemRes)
    ' Allocate byte array of right size
    ReDim abRes(cRes)
    ' Copy memory block to array
    CopyMemory abRes(0), ByVal pRes, cRes
    ' Free resource (no need to unlock)
    Call FreeResource(hmemRes)
    pbResource.Print HexDump(abRes, False)

End Sub

Sub ShowIcon(ByVal hMod As Long, sIcon As String)
    BugAssert (hMod <> hNull) And (sIcon <> sEmpty)
    
    ' Load icon resource
    hResourceCur = LoadImage(hMod, sIcon, IMAGE_ICON, 0, 0, 0)
    With pbResource
        If hResourceCur <> hNull Then
            ' Convert icon handle to Picture
            Dim pic As New StdPicture
            Set pic = IconToPicture(hResourceCur)
            pbResource.PaintPicture pic, 0, 0
            ordResourceLast = RT_ICON
        Else
            pbResource.Print "Can't display icon: " & sCrLf & sCrLf & _
                             WordWrap(ApiError(Err.LastDllError), 25)
        End If
    End With

End Sub

Sub ShowIcons(ByVal hMod As Long, sIcon As String)
    BugAssert (hMod <> hNull) And (sIcon <> sEmpty)
    
    ' Find the resource
    Dim hRes As Long, hmemRes As Long, cRes As Long, pRes As Long
    Dim abGroup() As Byte, abEntry() As Byte
    hRes = FindResourceStrId(hMod, sIcon, RT_GROUP_ICON)
    If hRes = hNull Then
        pbResource.Print "Can't display data: " & sCrLf & sCrLf & _
                         WordWrap(ApiError(Err.LastDllError), 25)
        Exit Sub
    End If
    pbResource.ScaleMode = vbPixels
    ' Allocate memory block, get size, get pointer, and allocate array
    hmemRes = LoadResource(hMod, hRes)
    cRes = SizeofResource(hMod, hRes)
    pRes = LockResource(hmemRes)
    ReDim abGroup(cRes)
    ' Copy memory block to bytes and free
    CopyMemory abGroup(0), ByVal pRes, cRes
    Call FreeResource(hmemRes)
    
    Dim cImage As Integer, i As Integer, iImage As Integer
    Dim dxIcon As Byte, dyIcon As Byte, hIcon As Long, s As String
    ' Validate entry
    BugAssert BytesToWord(abGroup, wType) = vbResIcon
    ' Get image count
    cImage = BytesToWord(abGroup, wCount)
    ' Set up first entry
    iImage = entFirst
    pbResource.CurrentX = 75
    pbResource.CurrentY = 0
    For i = 0 To cImage - 1
        ' Get size and colors
        dxIcon = abGroup(iImage + bWidth)
        dyIcon = abGroup(iImage + bHeight)
        s = dxIcon & "x" & dyIcon & ", " & _
            abGroup(iImage + bColorCount) & " color"
        ' Find, load, size, allocate, and copy entry
        hRes = FindResourceIdId(hMod, _
                                BytesToWord(abGroup, iImage + wID), _
                                RT_ICON)
        BugAssert hRes
        hmemRes = LoadResource(hMod, hRes)
        cRes = SizeofResource(hMod, hRes)
        pRes = LockResource(hmemRes)
        ReDim abEntry(cRes)
        CopyMemory abEntry(0), ByVal pRes, cRes
        Call FreeResource(hmemRes)
        ' Create an icon from resource data
        hIcon = CreateIconFromResource(abEntry(0), cRes, True, &H30000)
        ' Draw icon and print description
        Call DrawIconEx(pbResource.hDC, 0, pbResource.CurrentY, hIcon, _
                        dxIcon, dyIcon, 0, hNull, DI_NORMAL)
        pbResource.Print s
        ' Move to next entry
        pbResource.CurrentY = pbResource.CurrentY + dyIcon
        pbResource.CurrentX = 75
        iImage = iImage + cEntrySize
    Next
    pbResource.ScaleMode = vbTwips
    hResourceCur = hIcon
    ordResourceLast = RT_ICON
    
End Sub

Sub ShowString(ByVal hMod As Long, sString As String)
    Dim hRes As Long, hmemRes As Long, cRes As Long
    Dim pRes As Long, abRes() As Byte, i As Long
    hRes = FindResourceStrId(hMod, sString, RT_STRING)
    If hRes = hNull Then
        pbResource.Print "Can't display string: " & sCrLf & sCrLf & _
                         WordWrap(ApiError(Err.LastDllError), 25)
        Exit Sub
    End If
    ' Allocate memory block and get its size
    hmemRes = LoadResource(hMod, hRes)
    cRes = SizeofResource(hMod, hRes)
    ' Don't dump more than 500 bytes
    If cRes > 500 Then cRes = 500
    ' Lock it to get pointer
    pRes = LockResource(hmemRes)
    ' Allocate byte array of right size
    ReDim abRes(cRes)
    ' Copy memory block to array
    CopyMemory abRes(0), ByVal pRes, cRes
    ' Free resource (no need to unlock)
    Call FreeResource(hmemRes)
    pbResource.Print HexDump(abRes, False)
End Sub

Sub ShowMenu(ByVal hMod As Long, sMenu As String)

    hResourceCur = LoadMenu(hMod, sMenu)
    If hResourceCur <> 0 Then
        pbResource.Print "Menu set to: "
        pbResource.Print lstTopWin.Text
        hResourceLast = GetMenu(Me.hWnd)
        Call SetMenu(Me.hWnd, hResourceCur)
        ordResourceLast = RT_MENU
    Else
        pbResource.Print "Can't display menu: " & sCrLf & sCrLf & _
                         WordWrap(ApiError(Err.LastDllError), 25)
    End If
End Sub

Sub PlayWave(ByVal hMod As Long, sWave As String)
    ' Convert wave resource to memory
    Dim hWave As Long, hmemWave As Long, pWave As Long
    hWave = FindResourceStrStr(hMod, sWave, "WAVE")
    hmemWave = LoadResource(hMod, hWave)
    pWave = LockResource(hmemWave)
    Call FreeResource(hmemWave)
    ' Play it
    If sndPlaySoundAsLp(pWave, SND_MEMORY Or SND_NODEFAULT) Then
        pbResource.Print "Sound played"
    Else
        pbResource.Print "Can't play sound: " & sCrLf & sCrLf & _
                         WordWrap(ApiError(Err.LastDllError), 25)
    End If
End Sub

Sub PlayAvi(ByVal hMod As Long, sWave As String)
    Dim hWave As Long, hmemWave As Long, pWave As Long
    hWave = FindResourceStrStr(hMod, sWave, "WAVE")
    hmemWave = LoadResource(hMod, hWave)
    pWave = LockResource(hmemWave)
    
    ' All you have to do is figure out how to play AVI from memory
    'ImaginaryPlayWaveAPI pWave
    
    Call UnlockResource(hmemWave)
    Call FreeResource(hmemWave)
End Sub

Function GetVersionData(sExe As String, _
                        Optional ByVal cMaxChar As Long = 40) As String
    Dim version As New CVersion, s As String
    On Error GoTo NoVersionData
With version
    ' Initialize version object
    version = sExe
    ' Read and return properties
    s = s & WordWrap(.ProductName, cMaxChar) & sCrLf
    s = s & "Exe type: " & .ExeType & sCrLf
    s = s & "Internal name: " & .InternalName & sCrLf
    If .BuildString <> sEmpty Then
        s = s & "Build: " & .BuildString & sCrLf
    End If
    If .OriginalFilename <> sEmpty And _
       .OriginalFilename <> .InternalName Then
        s = s & "Original name: " & .OriginalFilename & sCrLf
    End If
    s = s & "Product version: " & .FullProductVersion & sCrLf
    s = s & "File version: " & .FullFileVersion & sCrLf
    s = s & "Company: " & WordWrap(.Company, cMaxChar) & sCrLf
    If .Comments <> sEmpty Then
        s = s & "Comments: " & WordWrap(.Comments, cMaxChar) & sCrLf
    End If
    s = s & "Copyright: " & WordWrap(.Copyright, cMaxChar) & sCrLf
    If .Trademarks <> sEmpty Then
        s = s & "Trademarks: " & WordWrap(.Trademarks, cMaxChar) & sCrLf
    End If
    's = s & "Host OS: " & .Environment & sCrLf
    s = s & "OS Version: " & .ProductVersionString & sCrLf
    If .Description <> sEmpty Then
        s = s & "Description: " & WordWrap(.Description, cMaxChar) & sCrLf
    End If
    Dim dt As Date
    dt = .TimeStamp
    If dt <> 0 Then
        s = s & "Time stamp: " & dt & sCrLf
    End If
    GetVersionData = s
End With
    Exit Function
    
NoVersionData:
    GetVersionData = "Unable to display version information"
End Function

Sub UpdateDisplay(eut As EUpdateType, hThing As Long)
    Dim idProc As Long, hWnd As Long, hTopWnd As Long, hMod As Long
    
    ' If top window is no longer valid, start from scratch
    Dim hTemp As Long
    If eut = eutTopWindow Or eut = eutWindow Then
        hTemp = ProcIDFromWnd(hThing)
    Else
        hTemp = ProcIDFromWnd(hWndCur)
    End If
    If hTemp = 0 Then
        RefreshAllLists
        Exit Sub
    End If

    Select Case eut
    Case eutTopWindow
        BugMessage "Top window update"
        idProc = ProcIDFromWnd(hThing)
        hWnd = hThing
        hTopWnd = hThing
        hMod = hModCur
    Case eutWindow
        BugMessage "Window update"
        idProc = ProcIDFromWnd(hThing)
        hWnd = hThing
        hTopWnd = TopWndFromProcID(idProc)
        hMod = hModCur
    Case eutProcess
        BugMessage "Process update"
        idProc = hThing
        hWnd = TopWndFromProcID(idProc)
        ' If process doesn't belong to a window in the
        ' top window list, don't change top window or
        ' window hierarchy displays
        If hWnd = hNull Then
            hWnd = hWndCur
            hTopWnd = hTopWndCur
        Else
            hTopWnd = hWnd
        End If
        hMod = hModCur
    Case eutModule
        BugMessage "Module update"
        hWnd = hWndCur
        hTopWnd = hTopWndCur
        idProc = idProcCur
        hMod = hThing
        sModCur = lstModule.Text
    End Select
    
    If hWnd Then BugMessage ExeNameFromWnd(hWnd)
    
    ' If window changed, update it
    If hWnd <> hWndCur Then
        hWndCur = hWnd
        lblWin.Caption = GetWndInfo(hWnd)
        tvwWin.Nodes.Item("W" & hWnd).Selected = True
    End If
    
    ' If top window changed, update it
    If hTopWnd <> hTopWndCur Then
        hTopWndCur = hTopWnd
        lstTopWin.ListIndex = LookupItemData(lstTopWin, hTopWnd)
    End If
    
    ' If process changed, update it
    If idProc <> idProcCur Then
        idProcCur = idProc
        ' Unload previous process
        If hModFree Then Call FreeLibrary(hModFree)
        sModCur = ExePathFromProcID(idProcCur)
        hMod = LoadLibraryEx(sModCur, 0, LOAD_LIBRARY_AS_DATAFILE)
        ' Save process handle for FreeLibrary
        hModFree = hMod
        ' Store module handles of new process
        RefreshModuleList idProc
        lblProc.Caption = GetProcInfo(idProc)
        lstProcess.ListIndex = LookupItemData(lstProcess, idProc)
    End If
        
    ' Update resources if module changed
    If hMod <> hModCur Then
        ' Update the resource list and the module information
        hModCur = hMod
        UpdateResources hModCur
        lblMod.Caption = "Module: " & sModCur & sCrLf & _
                         "Handle: " & Hex$(hMod)
        If lstResource.ListCount Then lstResource.ListIndex = 0
    End If
    
    If hWndCur Then hInstCur = InstFromWnd(hWndCur)
    
End Sub

Private Sub tvwWin_NodeClick(ByVal Node As Node)
    ' Get current window handle from treeview node Key property
    fInClick = True
    UpdateDisplay eutWindow, CLng(Mid$(tvwWin.SelectedItem.Key, 2))
    fInClick = False
End Sub
