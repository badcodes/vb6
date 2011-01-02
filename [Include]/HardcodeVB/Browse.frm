VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FBrowsePictures 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Browse Picture Files"
   ClientHeight    =   5412
   ClientLeft      =   2136
   ClientTop       =   2220
   ClientWidth     =   8076
   Icon            =   "Browse.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5412
   ScaleWidth      =   8076
   Begin MSComctlLib.Toolbar bar 
      Align           =   1  'Align Top
      Height          =   312
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   8076
      _ExtentX        =   14245
      _ExtentY        =   550
      ButtonWidth     =   487
      ButtonHeight    =   466
      ImageList       =   "imlstBar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Delete"
            Description     =   "Delete File"
            Object.ToolTipText     =   "Delete File"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Copy"
            Description     =   "Copy File"
            Object.ToolTipText     =   "Copy File"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Move"
            Description     =   "Move File"
            Object.ToolTipText     =   "Move File"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Rename"
            Description     =   "Rename File"
            Object.ToolTipText     =   "Rename File"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Connect"
            Description     =   "Map Network Drive"
            Object.ToolTipText     =   "Map Network Drive"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Disconnect"
            Description     =   "Disconnect Net Drive"
            Object.ToolTipText     =   "Disconnect Net Drive"
            ImageIndex      =   6
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picPic 
      AutoSize        =   -1  'True
      Height          =   1848
      Left            =   5508
      ScaleHeight     =   1800
      ScaleWidth      =   2256
      TabIndex        =   12
      Top             =   930
      Visible         =   0   'False
      Width           =   2304
   End
   Begin VB.PictureBox picPal 
      AutoRedraw      =   -1  'True
      Height          =   324
      Left            =   144
      ScaleHeight     =   276
      ScaleWidth      =   7704
      TabIndex        =   11
      Top             =   5016
      Width           =   7752
   End
   Begin VB.ComboBox cboPicType 
      Height          =   288
      Left            =   144
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   4200
      Width           =   2490
   End
   Begin VB.FileListBox filPic 
      Height          =   2568
      Left            =   144
      TabIndex        =   2
      Top             =   930
      Width           =   2505
   End
   Begin VB.DirListBox dirPic 
      Height          =   2790
      Left            =   2928
      TabIndex        =   1
      Top             =   930
      Width           =   2412
   End
   Begin VB.DriveListBox drvPic 
      Height          =   288
      Left            =   2925
      TabIndex        =   0
      Top             =   4200
      Width           =   2490
   End
   Begin MSComctlLib.ImageList imlstBar 
      Left            =   6984
      Top             =   2868
      _ExtentX        =   804
      _ExtentY        =   804
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Browse.frx":0CFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Browse.frx":0E0C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Browse.frx":0F1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Browse.frx":1030
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Browse.frx":1142
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Browse.frx":1254
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblDescribe 
      Height          =   1752
      Left            =   5544
      TabIndex        =   9
      Top             =   3492
      UseMnemonic     =   0   'False
      Width           =   2292
   End
   Begin VB.Image imgSIcon 
      Height          =   276
      Left            =   6096
      Top             =   2880
      Width           =   360
   End
   Begin VB.Image imgLIcon 
      Height          =   480
      Left            =   5520
      Top             =   2868
      Width           =   480
   End
   Begin VB.Image imgSound 
      Height          =   330
      Left            =   6570
      Top             =   3030
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgPic 
      Height          =   1740
      Left            =   5520
      Top             =   1020
      Width           =   2280
   End
   Begin VB.Label lbl 
      Caption         =   "File Name:"
      Height          =   228
      Index           =   1
      Left            =   192
      TabIndex        =   8
      Top             =   108
      Width           =   2436
   End
   Begin VB.Label lbl 
      Caption         =   "Directories:"
      Height          =   216
      Index           =   5
      Left            =   2916
      TabIndex        =   7
      Top             =   108
      Width           =   2436
   End
   Begin VB.Label lbl 
      Caption         =   "List Files of Type:"
      Height          =   312
      Index           =   4
      Left            =   144
      TabIndex        =   6
      Top             =   3864
      Width           =   2508
   End
   Begin VB.Label lbl 
      Caption         =   "Drives:"
      Height          =   312
      Index           =   3
      Left            =   2925
      TabIndex        =   5
      Top             =   3864
      Width           =   2436
   End
   Begin VB.Label lblPic 
      Height          =   396
      Left            =   168
      TabIndex        =   4
      Top             =   480
      Width           =   7668
   End
End
Attribute VB_Name = "FBrowsePictures"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Create an object that notifies of file changes
Private notify As CFileNotify

' Implement an interface that connects to CFileNotify
Implements IFileNotifier

Private hNotifyDir As Long
Private hNotifyFile As Long
Private hNotifyChange As Long

Private sInitDir As String
Private sFilePath As String     ' d:\path\
Private sFileName As String     ' base.ext
' Full file spec is sFilePath & sFileName
Private nsPicType As New Collection
Private dxPic As Long, dyPic As Long, yLbl As Long
Private ordMouse As Integer
Private afFilter As Long
Private fHasShell As Boolean
Private hPalBmp As Long

Private Sub Form_Load()

    dxPic = imgPic.Width
    dyPic = imgPic.Height
    yLbl = lblDescribe.Top
    BugMessage "Load"
    fHasShell = HasShell
    
    With cboPicType
        .AddItem "All Picture Files"
        nsPicType.Add "*.bmp;*.dib;*.rle;*.gif;*.jpg;*.jpeg;*.ico;*.wmf;*.emf;*.cur;*.wav"
        .AddItem "Bitmaps (.BMP;.DIB;.RLE;.GIF;.JPG;.JPEG)"
        nsPicType.Add "*.bmp;*.dib;*.gif;*.jpg;*.jpeg"
        .AddItem "Metafiles (.WMF;*.EMF)"
        nsPicType.Add "*.wmf;*.emf"
        .AddItem "Icons (.ICO)"
        nsPicType.Add "*.ico"
        .AddItem "Cursors (.CUR;.ICO)"
        nsPicType.Add "*.cur;*.ico"
        .AddItem "Waves (.WAV)"
        nsPicType.Add "*.wav"
        .AddItem "All Files (*.*)"
        nsPicType.Add "*.*"
        .ListIndex = 0
    End With
    
    ' Save mouse pointer so we can restore
    ordMouse = MousePointer
    
    ' Restore form state
    RestoreForm Me
    ' Get last directory
    sInitDir = GetSetting(App.EXEName, "Options", "LastPath", CurDir$)
    
    ' Changing path triggers notification initialization
    hNotifyDir = hInvalid
    hNotifyFile = hInvalid
    hNotifyChange = hInvalid
    Set notify = New CFileNotify
    On Error Resume Next
    ChDrive sInitDir
    dirPic.Path = sInitDir
    If Err Then dirPic.Path = App.Path
    On Error GoTo 0
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    BugMessage "Unload"
    notify.Disconnect hNotifyDir
    notify.Disconnect hNotifyFile
    notify.Disconnect hNotifyChange
    SaveSetting App.EXEName, "Options", "LastPath", dirPic.Path
    SaveForm Me
    'SaveWindow Me.hWnd, Me.Caption
End Sub

' FileTitle is read-only
Public Property Get FileTitle() As String
    FileTitle = sFileName  ' FileTitle is actually file name
End Property

Public Property Get FileName() As String
    If sFileName <> sEmpty Then
        FileName = sFilePath & sFileName
    ' Else (comments out because strings are empty by default)
    '    FileName = sEmpty
    End If
End Property

Public Property Let FileName(sFilePathA As String)
    sFilePath = sFilePathA
End Property

Public Property Get InitDir() As String
    InitDir = sInitDir
End Property

Public Property Let InitDir(sInitDirA As String)
    sInitDir = sInitDirA
End Property

Public Property Let FilterType(afFilterA As Long)
    afFilter = afFilterA
End Property

Private Sub bar_ButtonClick(ByVal Button As Button)
    Dim sFullPath As String, sDst As String, errOK As Long
    Static sLastString As String
    Const sTitle = "Destination"
    sFullPath = sFilePath & filPic.FileName
    Select Case Button.Key
    Case "Delete"
        If Not DeleteAnyFile(sFullPath, FOF_NOCONFIRMATION) Then
            MsgBox "Can't delete file"
        End If
    Case "Copy"
        sDst = InputBox("Copy destination: ", sTitle, sLastString)
        If sDst = sEmpty Then Exit Sub
        If Not CopyAnyFile(sFullPath, sDst) Then
            MsgBox "Can't copy file"
        Else
            sLastString = sDst
        End If
    Case "Move"
        sDst = InputBox("Move destination: ", sTitle, sLastString)
        If sDst = sEmpty Then Exit Sub
        If Not MoveAnyFile(sFullPath, sDst) Then
            MsgBox "Can't move file"
        Else
            sLastString = sDst
        End If
    Case "Rename"
        sDst = InputBox("New name: ", sTitle, sLastString)
        If sDst = sEmpty Then Exit Sub
        If Not RenameAnyFile(sFullPath, sFilePath & sDst) Then
            MsgBox "Can't rename file"
        Else
            sLastString = sDst
        End If
    Case "Connect"
        errOK = WNetConnectionDialog(Me.hWnd, RESOURCETYPE_DISK)
        drvPic.Refresh
    Case "Disconnect"
        errOK = WNetDisconnectDialog(Me.hWnd, RESOURCETYPE_DISK)
        drvPic.Refresh
    End Select
End Sub

Private Sub cboPicType_Click()
    filPic.Pattern = nsPicType(cboPicType.ListIndex + 1)
End Sub

Private Sub cmdCancel_Click()
    sFileName = sEmpty
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub dirPic_Change()
With notify
    ' Synchronize the file control and select the first file
    filPic.Path = dirPic.Path
    If filPic.ListCount > 0 Then filPic.ListIndex = 0

    ' Watch whole drive for directory changes
    If hNotifyDir <> -1 Then .Disconnect hNotifyDir
    hNotifyDir = .Connect(Me, dirPic.Path, _
                          FILE_NOTIFY_CHANGE_DIR_NAME, False)
    ' Watch current directory for name changes (delete, rename, create)
    If hNotifyFile <> -1 Then .Disconnect hNotifyFile
    hNotifyFile = .Connect(Me, dirPic.Path, _
                           FILE_NOTIFY_CHANGE_FILE_NAME, False)
    ' Watch current directory for modifications of file contents
    If hNotifyChange <> -1 Then notify.Disconnect hNotifyChange
    hNotifyChange = .Connect(Me, dirPic.Path, _
                             FILE_NOTIFY_CHANGE_LAST_WRITE, False)
End With
End Sub

Private Sub drvPic_Change()
    On Error GoTo NoDrive
    ChDrive drvPic.Drive
    dirPic.Path = drvPic.Drive
    Exit Sub
NoDrive:
End Sub

Private Sub filPic_PathChange()
    sFilePath = NormalizePath(filPic.Path)
    If filPic.ListCount > 0 Then filPic.ListIndex = 0
End Sub

Private Sub filPic_Click()
    sFileName = filPic.FileName
    UpdateFile sFilePath & sFileName
End Sub

Private Sub filPic_PatternChange()
    If filPic.ListCount > 0 Then
        filPic.ListIndex = 0
    End If
End Sub

Private Sub UpdateFile(sFile As String)
With imgPic
    Static fAviOpen As Boolean
    If fAviOpen Then
        mciSendString "close mywin", vbNullString, 0, hNull
        fAviOpen = False
        picPic.Visible = False
    End If
    PaletteMode = vbPaletteModeHalftone
    picPal.Visible = False
    MousePointer = ordMouse
    .Visible = False
    lblPic.Caption = sFile
    .Stretch = False
    Set imgSIcon.Picture = Nothing
    Set imgLIcon.Picture = Nothing
    Dim s As String, sExt As String
    sExt = UCase$(GetFileExt(sFile))
    imgSIcon.Picture = Nothing
    imgLIcon.Picture = Nothing
    lblDescribe.Caption = sEmpty
    lblDescribe.Top = yLbl
    hPalBmp = hNull
    
    Dim x As Long, y As Long, xHot As Long, yHot As Long
    On Error Resume Next
    Select Case sExt
    Case ".WAV"
        Set .Picture = imgSound.Picture
        s = s & "File length: " & FileLen(sFile)
        BugMessage s
        lblDescribe.Caption = s
        .Visible = True
        sndPlaySound sFile, 0
    Case ".ICO"
        s = "Type: Icon" & sCrLf
        Set imgPic = LoadPicture(sFile)
        imgPic.Width = 128 * Screen.TwipsPerPixelX
        imgPic.Height = 128 * Screen.TwipsPerPixelY
        Set imgLIcon.Picture = LoadAnyPicture(sFile, eisDefault)
        GetIconSize imgLIcon.Picture.Handle, x, y
        s = s & "Large: " & x & "x" & y & sCrLf
        If fHasShell Then
            Set imgSIcon.Picture = LoadAnyPicture(sFile, eisSmall)
            GetIconSize imgSIcon.Picture.Handle, x, y
            s = s & "Small: " & x & "x" & y & sCrLf
        End If
        
    Case ".CUR"
        Set imgPic = LoadPicture(sFile)
        s = "Type: Cursor" & sCrLf
        Me.MousePointer = vbCustom
        Set Me.MouseIcon = .Picture
        GetIconSize imgPic.Picture.Handle, x, y, xHot, yHot
        s = s & "Hot spot: " & xHot & "x" & yHot & sCrLf
        
    Case ".BMP", ".DIB", ".RLE", ".GIF", ".JPG", ".JPEG"
        lblDescribe.Top = imgLIcon.Top
        s = "Type: Bitmap" & sCrLf
        Dim hBmp As Long, hPal As Long
        ' LoadImage doesn't know about .GIF and .JPG, but LoadPicture does
        If sExt = ".GIF" Or sExt = ".JPG" Or sExt = ".JPEG" Then
            Set imgPic = LoadPicture(sFile)
        Else
            hBmp = LoadImage(App.hInstance, sFile, IMAGE_BITMAP, 0, 0, _
                             LR_LOADFROMFILE Or LR_CREATEDIBSECTION)
            hPal = GetBitmapPalette(hBmp)
            Set imgPic = BitmapToPicture(hBmp, hPal)
        End If
        If ScaleX(.Picture.Width) > dxPic Then
            .Height = (dxPic / ScaleX(.Picture.Width)) * _
                      ScaleY(.Picture.Height)
            .Width = dxPic
            .Stretch = True
        ElseIf ScaleY(.Picture.Height) > dyPic Then
            .Width = (dyPic / ScaleY(.Picture.Height)) * _
                     ScaleX(.Picture.Width)
            .Height = dyPic
            .Stretch = True
        End If
        GetBitmapSize imgPic.Picture.Handle, x, y
        s = s & "Image: " & Int(ScaleX(imgPic.Width, vbTwips, vbPixels)) & _
                "x" & Int(ScaleX(imgPic.Height, vbTwips, vbPixels)) & sCrLf
        s = s & "Picture: " & x & "x" & y & sCrLf
        s = s & "Handle: &H" & Hex(.Picture.Handle) & sCrLf
        s = s & "File length: " & FileLen(sFile) & sCrLf
        s = s & "Palette handle: &H" & Hex(.Picture.hPal) & sCrLf
        If .Picture.hPal Then
            s = s & "Palette size: " & PalSize(.Picture.hPal) & " colors" & sCrLf
            PaletteMode = vbPaletteModeCustom
            Palette = .Picture
            picPal.Cls
            picPal.Visible = True
            DrawPalette picPal, .Picture.hPal
            hPalBmp = .Picture.hPal
        End If
    Case ".WMF", ".EMF"
        s = "Type: Metafile" & sCrLf
        Set imgPic = LoadPicture(sFile)
        imgPic.Width = dxPic
        imgPic.Height = dyPic
        s = s & "Image: " & Int(ScaleX(imgPic.Width, vbTwips, vbPixels)) & _
                "x" & Int(ScaleX(imgPic.Height, vbTwips, vbPixels)) & sCrLf
        s = s & "Picture: " & Int(ScaleX(.Picture.Width, 8, vbPixels)) & _
                "x" & Int(ScaleY(.Picture.Height, 8, vbPixels)) & sCrLf
        s = s & "Handle: &H" & Hex(.Picture.Handle) & sCrLf
        s = s & "File length: " & FileLen(sFile)
    Case Else    ' Unknown extension
        Set imgPic = LoadPicture(sFile)
        If Err Then
            s = "Unknown format"
        Else
        
        End If
    End Select
    
    BugMessage s
    lblDescribe.Caption = s
    .Stretch = True
    .Visible = True
End With
End Sub

Private Sub IFileNotifier_Change(sDir As String, _
                                 efn As FileNotify.EFILE_NOTIFY, _
                                 fSubTree As Boolean)
    BugMessage "Directory: " & sDir & _
               " (" & efn & ":" & fSubTree & ")" & sCrLf
    Select Case efn
    Case FILE_NOTIFY_CHANGE_DIR_NAME, FILE_NOTIFY_CHANGE_FILE_NAME
        Dim i As Integer
        ' Refresh drive, directory, and file lists
        i = filPic.ListIndex
        filPic.Refresh
        filPic.ListIndex = IIf(i, i - 1, i)
        dirPic.Refresh
        drvPic.Refresh
    Case FILE_NOTIFY_CHANGE_LAST_WRITE
        ' Refresh current picture in case it changed
        filPic_Click
    End Select
End Sub
'
