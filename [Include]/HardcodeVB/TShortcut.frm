VERSION 5.00
Begin VB.Form FTestShortcuts 
   Caption         =   "Test Shortcuts"
   ClientHeight    =   5730
   ClientLeft      =   2070
   ClientTop       =   3060
   ClientWidth     =   6360
   Icon            =   "TShortcut.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5730
   ScaleWidth      =   6360
   Begin VB.TextBox txtDescription 
      Height          =   375
      Left            =   1920
      TabIndex        =   26
      Top             =   5244
      Width           =   4344
   End
   Begin VB.CheckBox chkCtl 
      Caption         =   "Ctl"
      Height          =   192
      Left            =   1212
      TabIndex        =   24
      Top             =   5175
      Value           =   1  'Checked
      Width           =   540
   End
   Begin VB.CheckBox chkAlt 
      Caption         =   "Alt"
      Height          =   255
      Left            =   1212
      TabIndex        =   23
      Top             =   5370
      Value           =   1  'Checked
      Width           =   495
   End
   Begin VB.CheckBox chkShift 
      Caption         =   "Shift"
      Height          =   168
      Left            =   1212
      TabIndex        =   22
      Top             =   5016
      Width           =   660
   End
   Begin VB.ListBox lstHotKey 
      Height          =   450
      Left            =   90
      TabIndex        =   21
      Top             =   5010
      Width           =   1050
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "&New"
      Height          =   384
      Left            =   96
      TabIndex        =   19
      Top             =   576
      Width           =   1575
   End
   Begin VB.CommandButton cmdFill 
      Caption         =   "Fill"
      Height          =   375
      Left            =   5385
      TabIndex        =   18
      Top             =   3864
      Width           =   855
   End
   Begin VB.TextBox txtPath 
      Height          =   375
      Left            =   1920
      TabIndex        =   16
      Top             =   2460
      Width           =   4332
   End
   Begin VB.TextBox txtArguments 
      Height          =   375
      Left            =   1920
      TabIndex        =   15
      Top             =   4548
      Width           =   4344
   End
   Begin VB.ListBox lstDisplay 
      Height          =   450
      ItemData        =   "TShortcut.frx":0CFA
      Left            =   84
      List            =   "TShortcut.frx":0D07
      TabIndex        =   13
      Top             =   4164
      Width           =   1560
   End
   Begin VB.ListBox lstLocation 
      Height          =   1425
      ItemData        =   "TShortcut.frx":0D30
      Left            =   84
      List            =   "TShortcut.frx":0D4C
      TabIndex        =   12
      Top             =   2304
      Width           =   1596
   End
   Begin VB.Frame fmTarget 
      Caption         =   "Target File"
      Height          =   2136
      Left            =   1920
      TabIndex        =   7
      Top             =   0
      Width           =   4332
      Begin VB.ComboBox cboPattern 
         Height          =   288
         ItemData        =   "TShortcut.frx":0DC7
         Left            =   2016
         List            =   "TShortcut.frx":0DE6
         TabIndex        =   25
         Top             =   1704
         Width           =   2148
      End
      Begin VB.DriveListBox drvLink 
         Height          =   288
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1836
      End
      Begin VB.DirListBox dirLink 
         Height          =   1368
         Left            =   120
         TabIndex        =   9
         Top             =   612
         Width           =   1824
      End
      Begin VB.FileListBox fileLink 
         Height          =   1260
         Left            =   2016
         TabIndex        =   8
         Top             =   204
         Width           =   2160
      End
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "&Reset"
      Height          =   408
      Left            =   96
      TabIndex        =   5
      Top             =   1032
      Width           =   1575
   End
   Begin VB.TextBox txtDirectory 
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   3852
      Width           =   3348
   End
   Begin VB.TextBox txtLink 
      Height          =   375
      Left            =   1935
      TabIndex        =   1
      Top             =   3144
      Width           =   4320
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   384
      Left            =   84
      TabIndex        =   0
      Top             =   108
      Width           =   1575
   End
   Begin VB.Label lbl 
      Caption         =   "Description:"
      Height          =   255
      Index           =   7
      Left            =   1920
      TabIndex        =   27
      Top             =   5010
      Width           =   1335
   End
   Begin VB.Image imgIcon 
      Height          =   504
      Left            =   96
      Top             =   1524
      Width           =   660
   End
   Begin VB.Label lbl 
      Caption         =   "Hot key:"
      Height          =   264
      Index           =   5
      Left            =   96
      TabIndex        =   20
      Top             =   4800
      Width           =   792
   End
   Begin VB.Label lbl 
      Caption         =   "Target path:"
      Height          =   255
      Index           =   6
      Left            =   1905
      TabIndex        =   17
      Top             =   2220
      Width           =   1335
   End
   Begin VB.Label lbl 
      Caption         =   "Arguments:"
      Height          =   255
      Index           =   4
      Left            =   1920
      TabIndex        =   14
      Top             =   4305
      Width           =   1335
   End
   Begin VB.Label lbl 
      Caption         =   "Set location:"
      Height          =   252
      Index           =   3
      Left            =   96
      TabIndex        =   11
      Top             =   2088
      Width           =   1212
   End
   Begin VB.Label lbl 
      Caption         =   "Display:"
      Height          =   252
      Index           =   2
      Left            =   108
      TabIndex        =   6
      Top             =   3900
      Width           =   1332
   End
   Begin VB.Label lbl 
      Caption         =   "Link file: "
      Height          =   255
      Index           =   1
      Left            =   1920
      TabIndex        =   4
      Top             =   2910
      Width           =   1215
   End
   Begin VB.Label lbl 
      Caption         =   "Working Directory:"
      Height          =   255
      Index           =   0
      Left            =   1920
      TabIndex        =   3
      Top             =   3615
      Width           =   1815
   End
End
Attribute VB_Name = "FTestShortcuts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private shortcut As New CShortcut
Private sInitDir As String
Private sFilePath As String     ' d:\dir\
Private sFileName As String     ' base.ext
Private sCurDir As String
Private fInside As Boolean

Private Sub Form_Load()
    ChDir App.Path
    sCurDir = CurDir$
    lstDisplay.ItemData(0) = edmNormal
    lstDisplay.ItemData(1) = edmMinimized
    lstDisplay.ItemData(2) = edmMaximized
    lstLocation.ListIndex = edstDesktop
    cboPattern.ListIndex = 0
    dirLink.Path = Environ("windir")
    fileLink_PathChange
    fileLink_Click
    lstDisplay.ListIndex = edmNormal - 1
    InitKeys lstHotKey
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ChDrive sCurDir
    ChDir sCurDir
End Sub

Private Sub cmdSave_Click()
With shortcut
    On Error GoTo FailSave
    .Path = txtPath
    If txtDirectory <> "" Then .WorkingDirectory = txtDirectory
    If txtArguments <> "" Then .Arguments = txtArguments
    If txtDescription <> "" Then .Description = txtDescription
    If lstHotKey.ListIndex <> 0 Then
        .Hotkey = lstHotKey.ItemData(lstHotKey.ListIndex) Or _
                  (chkShift.Value * &H100) Or _
                  (chkCtl.Value * &H200) Or _
                  (chkAlt.Value * &H400)
    Else
        .Hotkey = 0
    End If
    .DisplayMode = lstDisplay.ItemData(lstDisplay.ListIndex)
    txtLink = .Save(txtLink)
    Exit Sub
FailSave:
    MsgBox "Can't save"
End With
End Sub

Private Sub cmdNew_Click()
    Set shortcut = Nothing
    txtDirectory = sEmpty
    txtArguments = sEmpty
    txtDescription = sEmpty
    lstDisplay.ListIndex = 0
    lstHotKey.ListIndex = 0
End Sub

Private Sub cmdReset_Click()
Dim i As Integer
With shortcut
    i = fileLink.ListIndex
    fileLink.Refresh
    fInside = True
    fileLink.ListIndex = i
    fInside = False
    txtPath = .Path
    txtDirectory = .WorkingDirectory
    txtArguments = .Arguments
    txtDescription = .Description
    lstDisplay.ListIndex = LookupItemData(lstDisplay, .DisplayMode)
    lstHotKey.ListIndex = LookupItemData(lstHotKey, .Hotkey And &HFF)
    chkShift.Value = -CBool(.Hotkey And &H100)
    chkCtl.Value = -CBool(.Hotkey And &H200)
    chkAlt.Value = -CBool(.Hotkey And &H400)
    Set imgIcon.Picture = .Icon
End With
End Sub

Private Sub cmdFill_Click()
    txtDirectory = sFilePath
End Sub

Private Sub dirLink_Change()
    fileLink.Path = dirLink.Path
    If fileLink.ListCount > 0 Then
        fileLink.ListIndex = 0
    End If
    txtPath = sFilePath & sFileName
End Sub

Private Sub cboPattern_Click()
    Dim s As String, iParen As Long
    iParen = InStr(cboPattern.Text, "(")
    s = Mid$(cboPattern.Text, iParen + 1)
    s = Left$(s, Len(s) - 1)
    Select Case Left$(cboPattern.Text, iParen - 2)
    Case "Desktop"
        dirLink.Path = GetDesktop
    Case "Common Desktop"
        dirLink.Path = GetCommonDesktop
    Case "Common Programs"
        dirLink.Path = GetCommonPrograms
    Case "Programs"
        dirLink.Path = GetPrograms
    Case "Start Menu"
        dirLink.Path = GetStartMenu
    End Select
    fileLink.Pattern = s
End Sub

Private Sub drvLink_Change()
    dirLink.Path = drvLink.Drive
    txtPath = sFilePath & sFileName
End Sub

Private Sub fileLink_PathChange()
    sFilePath = NormalizePath(fileLink.Path)
    If fileLink.ListCount > 0 Then fileLink.ListIndex = 0
End Sub

Private Sub fileLink_Click()
With shortcut
    If fInside Then Exit Sub
    fInside = True
    sFileName = fileLink.filename
    txtPath = sFilePath & sFileName
    If UCase$(GetFileExt(sFileName)) = ".LNK" Then
        ' Update all fields for
        txtLink = txtPath
        .Resolve txtLink
        cmdReset_Click
    Else
        lstLocation_Click
        Dim hIcon As Long
        hIcon = ExtractIcon(App.hInstance, txtPath, 0)
        Set imgIcon.Picture = IconToPicture(hIcon)
    End If
    fInside = False
End With
End Sub

Private Sub fileLink_PatternChange()
    If fileLink.ListCount > 0 Then
        fileLink.ListIndex = 0
    End If
    txtPath = sFilePath & sFileName
End Sub

Private Sub lstDisplay_Click()
    shortcut.DisplayMode = lstDisplay.ItemData(lstDisplay.ListIndex)
End Sub

'Private Sub lstLocation_Click()
'    shortcut.Location = lstLocation.ListIndex
'    txtLink = shortcut.Location
'End Sub

'Private Sub txtLink_LostFocus()
'    shortcut.Location = txtLink
'End Sub

Sub InitKeys(lst As Control)
With lst
    
    .AddItem "(None)": .ItemData(.ListCount - 1) = 0
    
    ' Function Keys
    .AddItem "F1": .ItemData(.ListCount - 1) = vbKeyF1        ' F1 key
    .AddItem "F2": .ItemData(.ListCount - 1) = vbKeyF2        ' F2 key
    .AddItem "F3": .ItemData(.ListCount - 1) = vbKeyF3        ' F3 key
    .AddItem "F4": .ItemData(.ListCount - 1) = vbKeyF4        ' F4 key
    .AddItem "F5": .ItemData(.ListCount - 1) = vbKeyF5        ' F5 key
    .AddItem "F6": .ItemData(.ListCount - 1) = vbKeyF6        ' F6 key
    .AddItem "F7": .ItemData(.ListCount - 1) = vbKeyF7        ' F7 key
    .AddItem "F8": .ItemData(.ListCount - 1) = vbKeyF8        ' F8 key
    .AddItem "F9": .ItemData(.ListCount - 1) = vbKeyF9        ' F9 key
    .AddItem "F10": .ItemData(.ListCount - 1) = vbKeyF10      ' F10 key
    .AddItem "F11": .ItemData(.ListCount - 1) = vbKeyF11      ' F11 key
    .AddItem "F12": .ItemData(.ListCount - 1) = vbKeyF12      ' F12 key
    .AddItem "F13": .ItemData(.ListCount - 1) = vbKeyF13      ' F13 key
    .AddItem "F14": .ItemData(.ListCount - 1) = vbKeyF14      ' F14 key
    .AddItem "F15": .ItemData(.ListCount - 1) = vbKeyF15      ' F15 key
    .AddItem "F16": .ItemData(.ListCount - 1) = vbKeyF16      ' F16 key
    
    ' Miscellaneous control keys
    .AddItem "Cancel": .ItemData(.ListCount - 1) = vbKeyCancel      ' CANCEL key
    .AddItem "Back": .ItemData(.ListCount - 1) = vbKeyBack          ' BACKSPACE key
    .AddItem "Tab": .ItemData(.ListCount - 1) = vbKeyTab            ' TAB key
    .AddItem "Clear": .ItemData(.ListCount - 1) = vbKeyClear        ' CLEAR key
    .AddItem "Return": .ItemData(.ListCount - 1) = vbKeyReturn      ' ENTER key
    .AddItem "Menu": .ItemData(.ListCount - 1) = vbKeyMenu          ' MENU key
    .AddItem "Pause": .ItemData(.ListCount - 1) = vbKeyPause        ' PAUSE key
    .AddItem "Escape": .ItemData(.ListCount - 1) = vbKeyEscape      ' ESC key
    .AddItem "Space": .ItemData(.ListCount - 1) = vbKeySpace        ' SPACEBAR key
    .AddItem "PageUp": .ItemData(.ListCount - 1) = vbKeyPageUp      ' PAGE UP key
    .AddItem "PageDown": .ItemData(.ListCount - 1) = vbKeyPageDown  ' PAGE DOWN key
    .AddItem "End": .ItemData(.ListCount - 1) = vbKeyEnd            ' END key
    .AddItem "Home": .ItemData(.ListCount - 1) = vbKeyHome          ' HOME key
    .AddItem "Left": .ItemData(.ListCount - 1) = vbKeyLeft          ' LEFT ARROW key
    .AddItem "Up": .ItemData(.ListCount - 1) = vbKeyUp              ' UP ARROW key
    .AddItem "Right": .ItemData(.ListCount - 1) = vbKeyRight        ' RIGHT ARROW key
    .AddItem "Down": .ItemData(.ListCount - 1) = vbKeyDown          ' DOWN ARROW key
    .AddItem "Select": .ItemData(.ListCount - 1) = vbKeySelect      ' SELECT key
    .AddItem "Print": .ItemData(.ListCount - 1) = vbKeyPrint        ' PRINT SCREEN key
    .AddItem "Execute": .ItemData(.ListCount - 1) = vbKeyExecute    ' EXECUTE key
    .AddItem "Snapshot": .ItemData(.ListCount - 1) = vbKeySnapshot  ' SNAPSHOT key
    .AddItem "Insert": .ItemData(.ListCount - 1) = vbKeyInsert      ' INS key
    .AddItem "Delete": .ItemData(.ListCount - 1) = vbKeyDelete      ' DEL key
    .AddItem "Help": .ItemData(.ListCount - 1) = vbKeyHelp          ' HELP key
    .AddItem "Numlock": .ItemData(.ListCount - 1) = vbKeyNumlock    ' NUM LOCK key
    
    ' Keys on the Numeric Keypad
    .AddItem "Numpad 0": .ItemData(.ListCount - 1) = vbKeyNumpad0    ' 0 key
    .AddItem "Numpad 1": .ItemData(.ListCount - 1) = vbKeyNumpad1    ' 1 key
    .AddItem "Numpad 2": .ItemData(.ListCount - 1) = vbKeyNumpad2    ' 2 key
    .AddItem "Numpad 3": .ItemData(.ListCount - 1) = vbKeyNumpad3    ' 3 key
    .AddItem "Numpad 4": .ItemData(.ListCount - 1) = vbKeyNumpad4    ' 4 key
    .AddItem "Numpad 5": .ItemData(.ListCount - 1) = vbKeyNumpad5    ' 5 key
    .AddItem "Numpad 6": .ItemData(.ListCount - 1) = vbKeyNumpad6    ' 6 key
    .AddItem "Numpad 7": .ItemData(.ListCount - 1) = vbKeyNumpad7    ' 7 key
    .AddItem "Numpad 8": .ItemData(.ListCount - 1) = vbKeyNumpad8    ' 8 key
    .AddItem "Numpad 9": .ItemData(.ListCount - 1) = vbKeyNumpad9    ' 9 key
    .AddItem "Multiply": .ItemData(.ListCount - 1) = vbKeyMultiply   ' MULTIPLICATION SIGN (*) key
    .AddItem "Add": .ItemData(.ListCount - 1) = vbKeyAdd             ' PLUS SIGN (+) key
    .AddItem "Separator": .ItemData(.ListCount - 1) = vbKeySeparator ' ENTER (keypad) key
    .AddItem "Subtract": .ItemData(.ListCount - 1) = vbKeySubtract   ' MINUS SIGN (-) key
    .AddItem "Decimal": .ItemData(.ListCount - 1) = vbKeyDecimal     ' DECIMAL POINT(.) key
    .AddItem "Divide": .ItemData(.ListCount - 1) = vbKeyDivide       ' DIVISION SIGN (/) key
    
    ' KeyA Through KeyZ Are the Same as Their ASCII Equivalents
    .AddItem "A": .ItemData(.ListCount - 1) = vbKeyA         ' A key
    .AddItem "B": .ItemData(.ListCount - 1) = vbKeyB         ' B key
    .AddItem "C": .ItemData(.ListCount - 1) = vbKeyC         ' C key
    .AddItem "D": .ItemData(.ListCount - 1) = vbKeyD         ' D key
    .AddItem "E": .ItemData(.ListCount - 1) = vbKeyE         ' E key
    .AddItem "F": .ItemData(.ListCount - 1) = vbKeyF         ' F key
    .AddItem "G": .ItemData(.ListCount - 1) = vbKeyG         ' G key
    .AddItem "H": .ItemData(.ListCount - 1) = vbKeyH         ' H key
    .AddItem "I": .ItemData(.ListCount - 1) = vbKeyI         ' I key
    .AddItem "J": .ItemData(.ListCount - 1) = vbKeyJ         ' J key
    .AddItem "K": .ItemData(.ListCount - 1) = vbKeyK         ' K key
    .AddItem "L": .ItemData(.ListCount - 1) = vbKeyL         ' L key
    .AddItem "M": .ItemData(.ListCount - 1) = vbKeyM         ' M key
    .AddItem "N": .ItemData(.ListCount - 1) = vbKeyN         ' N key
    .AddItem "O": .ItemData(.ListCount - 1) = vbKeyO         ' O key
    .AddItem "P": .ItemData(.ListCount - 1) = vbKeyP         ' P key
    .AddItem "Q": .ItemData(.ListCount - 1) = vbKeyQ         ' Q key
    .AddItem "R": .ItemData(.ListCount - 1) = vbKeyR         ' R key
    .AddItem "S": .ItemData(.ListCount - 1) = vbKeyS         ' S key
    .AddItem "T": .ItemData(.ListCount - 1) = vbKeyT         ' T key
    .AddItem "U": .ItemData(.ListCount - 1) = vbKeyU         ' U key
    .AddItem "V": .ItemData(.ListCount - 1) = vbKeyV         ' V key
    .AddItem "W": .ItemData(.ListCount - 1) = vbKeyW         ' W key
    .AddItem "X": .ItemData(.ListCount - 1) = vbKeyX         ' X key
    .AddItem "Y": .ItemData(.ListCount - 1) = vbKeyY         ' Y key
    .AddItem "Z": .ItemData(.ListCount - 1) = vbKeyZ         ' Z key

    ' Key0 Through Key9 Are the Same as Their ASCII Equivalents: '0' Through '9
    .AddItem "0": .ItemData(.ListCount - 1) = vbKey0         ' 0 key
    .AddItem "1": .ItemData(.ListCount - 1) = vbKey1         ' 1 key
    .AddItem "2": .ItemData(.ListCount - 1) = vbKey2         ' 2 key
    .AddItem "3": .ItemData(.ListCount - 1) = vbKey3         ' 3 key
    .AddItem "4": .ItemData(.ListCount - 1) = vbKey4         ' 4 key
    .AddItem "5": .ItemData(.ListCount - 1) = vbKey5         ' 5 key
    .AddItem "6": .ItemData(.ListCount - 1) = vbKey6         ' 6 key
    .AddItem "7": .ItemData(.ListCount - 1) = vbKey7         ' 7 key
    .AddItem "8": .ItemData(.ListCount - 1) = vbKey8         ' 8 key
    .AddItem "9": .ItemData(.ListCount - 1) = vbKey9         ' 9 key
    
    .ListIndex = 0
End With
End Sub


Private Sub lstHotKey_Click()
    If lstHotKey.ListIndex = 0 Then
        chkAlt = vbUnchecked
        chkCtl = vbUnchecked
        chkShift = vbUnchecked
        chkAlt.Enabled = False
        chkCtl.Enabled = False
        chkShift.Enabled = False
    Else
        chkAlt.Enabled = True
        chkCtl.Enabled = True
        chkShift.Enabled = True
    End If
End Sub

Private Sub lstLocation_Click()
    Dim sFile As String
    sFile = GetFileBase(sFileName) & ".LNK"
    Select Case lstLocation.ListIndex
    Case edstDesktop
        txtLink = GetDesktop & sFile
    Case edstCommonDesktop
        txtLink = GetCommonDesktop & sFile
    Case edstPrograms
        txtLink = GetPrograms & sFile
    Case edstCommonPrograms
        txtLink = GetCommonPrograms & sFile
    Case edstStartMenu
        txtLink = GetStartMenu & sFile
    Case edstPath
        txtLink = sFilePath & sFile
    Case edstCurrent
        txtLink = NormalizePath(CurDir$) & sFile
    Case Else
        txtLink.Text = sFile
        txtLink.SelStart = 0
        txtLink.SetFocus
    End Select
End Sub
