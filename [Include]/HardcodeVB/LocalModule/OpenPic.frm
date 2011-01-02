VERSION 5.00
Begin VB.Form FOpenPictureFile 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Open Picture File"
   ClientHeight    =   3276
   ClientLeft      =   936
   ClientTop       =   2160
   ClientWidth     =   7140
   LinkTopic       =   "Form1"
   ScaleHeight     =   3276
   ScaleWidth      =   7140
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdNetwork 
      Caption         =   "Network..."
      Height          =   330
      Left            =   5565
      TabIndex        =   10
      Top             =   945
      Width           =   1380
   End
   Begin VB.ComboBox cboPicType 
      Height          =   288
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   2760
      Width           =   2484
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   330
      Left            =   5565
      TabIndex        =   4
      Top             =   525
      Width           =   1380
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   330
      Left            =   5565
      TabIndex        =   3
      Top             =   105
      Width           =   1380
   End
   Begin VB.FileListBox filPic 
      Height          =   1800
      Left            =   180
      TabIndex        =   2
      Top             =   480
      Width           =   2505
   End
   Begin VB.DirListBox dirPic 
      Height          =   1665
      Left            =   2940
      TabIndex        =   1
      Top             =   720
      Width           =   2412
   End
   Begin VB.DriveListBox drvPic 
      Height          =   288
      Left            =   2928
      TabIndex        =   0
      Top             =   2760
      Width           =   2496
   End
   Begin VB.Image imgSound 
      Height          =   264
      Left            =   6840
      Top             =   3000
      Visible         =   0   'False
      Width           =   288
   End
   Begin VB.Image imgPic 
      Height          =   1425
      Left            =   5550
      Top             =   1560
      Width           =   1380
   End
   Begin VB.Label lbl 
      Caption         =   "Directories:"
      Height          =   210
      Index           =   5
      Left            =   3000
      TabIndex        =   9
      Top             =   480
      Width           =   2430
   End
   Begin VB.Label lbl 
      Caption         =   "List Files of Type:"
      Height          =   312
      Index           =   4
      Left            =   204
      TabIndex        =   8
      Top             =   2436
      Width           =   2508
   End
   Begin VB.Label lbl 
      Caption         =   "Drives:"
      Height          =   315
      Index           =   3
      Left            =   2925
      TabIndex        =   7
      Top             =   2415
      Width           =   2430
   End
   Begin VB.Label lblPic 
      Height          =   270
      Left            =   240
      TabIndex        =   6
      Top             =   120
      Width           =   5055
   End
End
Attribute VB_Name = "FOpenPictureFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Basic provides five constants, but we need six
Private Enum EPictureType
    ' vbPicTypeNone = 0
    ' vbPicTypeBitmap = 1
    ' vbPicTypeMetafile = 2
    ' vbPicTypeIcon = 3
    ' vbPicTypeEMetafile = 4
    vbPicTypeCursor = 5
    vbPicTypeWave = 6
End Enum

Private sInitDir As String
Private sFilePath As String     ' d:\path\
Private sFileName As String     ' base.ext
' Full file spec is sFilePath & sFileName
Private nsPicType As New Collection
Private dxPic As Integer, dyPic As Integer
Private ordMouse As Integer
Private ordPicType As Integer
Private afFilter As Long

' FileTitle is read-only
Friend Property Get FileTitle() As String
    FileTitle = sFileName  ' FileTitle is actually filename
End Property

Friend Property Get FileName() As String
    If sFileName <> sEmpty Then
        FileName = sFilePath & sFileName
    ' Else (commented out because strings are empty by default)
    '    FileName = sEmpty
    End If
End Property

Friend Property Let FileName(sFilePathA As String)
    sFilePath = sFilePathA
End Property

Friend Property Get InitDir() As String
    InitDir = sInitDir
End Property

Friend Property Let InitDir(sInitDirA As String)
    sInitDir = sInitDirA
End Property

Friend Property Get PicType() As Integer
    PicType = ordPicType
End Property

Friend Property Get FilterType() As EFilterPicture
    FilterType = afFilter
End Property

Friend Property Let FilterType(afFilterA As EFilterPicture)
    afFilter = afFilterA
End Property

Private Sub cboPicType_Click()
    filPic.Pattern = nsPicType(cboPicType.ListIndex + 1)
End Sub

Private Sub cmdCancel_Click()
    sFileName = sEmpty
    Unload Me
End Sub

Private Sub cmdNetwork_Click()
    Dim errOK As Long
    errOK = WNetConnectionDialog(Me.hWnd, 1) ' WNTYPE_DRIVE
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub dirPic_Change()
    filPic.Path = dirPic.Path
    If filPic.ListCount > 0 Then
        filPic.ListIndex = 0
    End If
End Sub

Private Sub drvPic_Change()
    dirPic.Path = drvPic.Drive
End Sub

Private Sub filPic_DblClick()
    Unload Me
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

Private Sub Form_Initialize()
    BugMessage "Initialize"
End Sub

Private Sub Form_Load()
    BugMessage "Load"
    dxPic = imgPic.Width
    dyPic = imgPic.Height
    If sInitDir <> sEmpty Then
        dirPic.Path = NormalizePath(sInitDir)
    Else
        sInitDir = NormalizePath(dirPic.Path)
    End If
    With cboPicType
        If afFilter = 0 Then afFilter = efpEverything
        If afFilter And efpAllPicture Then
            .AddItem "All Picture Files"
            nsPicType.Add "*.bmp;*.dib;*.rle;*.jpg;*.jpeg;*.gif;*.ico;*.wmf;*.emf;*.cur;*.wav"
        End If
        If afFilter And efpBitmap Then
            .AddItem "Bitmaps (.BMP;.DIB;.RLE;.JPG;.JPEG;.GIF)"
            nsPicType.Add "*.bmp;*.dib;*.rle;*.jpg;*.jpeg;*.gif"
        End If
        If afFilter And efpMetafile Then
            .AddItem "Metafiles (.WMF;.EMF)"
            nsPicType.Add "*.wmf;*.emf"
        End If
        If afFilter And efpIcon Then
            .AddItem "Icons (.ICO)"
            nsPicType.Add "*.ico"
        End If
        If afFilter And efpCursor Then
            .AddItem "Cursors (.CUR;.ICO)"
            nsPicType.Add "*.cur;*.ico"
        End If
        If afFilter And efpWave Then
            .AddItem "Waves (.WAV)"
            nsPicType.Add "*.wav"
        End If
        If afFilter And efpAllFile Then
            .AddItem "All Files"
            nsPicType.Add "*.*"
        End If
        If .ListCount Then .ListIndex = 0
    End With
    ' Save mouse pointer so we can restore
    ordMouse = MousePointer
    dirPic_Change
    filPic_PathChange
    filPic_Click
End Sub

Private Sub UpdateFile(sFile As String)
    MousePointer = ordMouse
    With imgPic
        .Visible = False
        lblPic.Caption = sFile
        .Stretch = False
        If UCase$(Right$(sFile, 4)) = ".WAV" Then
            sndPlaySound sFile, 0
            .Picture = imgSound.Picture
            .Visible = True
            ordPicType = vbPicTypeWave
            Exit Sub
        End If
        On Error Resume Next
        .Picture = LoadPicture(sFile)
        If Err Then Exit Sub
        On Error GoTo 0
        ordPicType = .Picture.Type
        Select Case .Picture.Type
        Case vbPicTypeIcon
            If UCase$(Right$(sFile, 4)) = ".CUR" Then
                ordPicType = vbPicTypeCursor
                On Error Resume Next
                MousePointer = vbCustom
                MouseIcon = .Picture
                If Err = 0 Then Exit Sub
                On Error GoTo 0
            End If
        Case vbPicTypeBitmap
            .Visible = True
            If ScaleX(.Picture.Width) > dxPic Then
                imgPic.Height = (dxPic / ScaleX(.Picture.Width)) * _
                                ScaleY(.Picture.Height)
                imgPic.Width = dxPic
                .Stretch = True
            ElseIf ScaleY(.Picture.Height) > dyPic Then
                imgPic.Width = (dyPic / ScaleY(.Picture.Height)) * _
                                ScaleX(.Picture.Width)
                imgPic.Height = dyPic
                .Stretch = True
            End If
            BugMessage "Palette: " & .Picture.hPal
        Case vbPicTypeMetafile, vbPicTypeEMetafile
            imgPic.Width = dxPic
            imgPic.Height = dyPic
            .Stretch = True
        End Select
        BugMessage "Type: " & .Picture.Type
        BugMessage "Handle: " & .Picture.Handle
        .Visible = True
    End With
End Sub

Private Sub Form_Terminate()
    BugMessage "Terminate"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    BugMessage "Unload"
End Sub

