VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FAllAbout 
   Caption         =   "All About..."
   ClientHeight    =   5640
   ClientLeft      =   2016
   ClientTop       =   4248
   ClientWidth     =   8604
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   7.8
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "AllAbout.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5640
   ScaleWidth      =   8604
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab tabAbout 
      Height          =   4452
      Left            =   336
      TabIndex        =   0
      Top             =   636
      Width           =   7836
      _ExtentX        =   13822
      _ExtentY        =   7853
      _Version        =   393216
      Tabs            =   4
      Tab             =   3
      TabsPerRow      =   4
      TabHeight       =   420
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "System"
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lbl(0)"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Video"
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lbl(1)"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Drives"
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cmdDrives"
      Tab(2).Control(1)=   "lbl(2)"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Version"
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "lbl(3)"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "lblFile"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "cmdVersion"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).ControlCount=   3
      Begin VB.CommandButton cmdVersion 
         Caption         =   "New..."
         Height          =   372
         Left            =   6624
         TabIndex        =   6
         Top             =   3792
         Width           =   972
      End
      Begin VB.CommandButton cmdDrives 
         Caption         =   "Refresh"
         Height          =   372
         Left            =   -68412
         TabIndex        =   5
         Top             =   3804
         Width           =   972
      End
      Begin VB.Label lblFile 
         Height          =   192
         Left            =   84
         TabIndex        =   7
         Top             =   336
         Width           =   7608
      End
      Begin VB.Label lbl 
         Height          =   3972
         Index           =   0
         Left            =   -74880
         TabIndex        =   4
         Top             =   360
         Width           =   7620
      End
      Begin VB.Label lbl 
         Height          =   3516
         Index           =   3
         Left            =   120
         TabIndex        =   3
         Top             =   756
         Width           =   7584
      End
      Begin VB.Label lbl 
         Height          =   3972
         Index           =   2
         Left            =   -74892
         TabIndex        =   2
         Top             =   336
         Width           =   7644
      End
      Begin VB.Label lbl 
         Height          =   3972
         Index           =   1
         Left            =   -74772
         TabIndex        =   1
         Top             =   360
         Width           =   7512
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "FAllAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private sExe As String
Private fUpdate As Boolean
Private sFilterString As String

Const ordSystem = 0
Const ordVideo = 1
Const ordDrives = 2
Const ordVersion = 3

Private Sub Form_Initialize()
    BugLocalMessage "Initializing All About"
End Sub

Private Sub Form_Load()
            
    BugLocalMessage "Loading All About"
    ChDir App.Path
    
    sExe = RealExeName
    ' Tranfer to any previous instance
    If App.PrevInstance Then
        Dim sTitle As String
        ' Save my title
        sTitle = Me.Caption
        ' Change my title bar so I won't activate myself
        Me.Caption = Hex$(Me.hWnd)
        ' Activate other instance
        AppActivate sTitle
        ' Terminate myself
        End
    End If
    
    If IsExe() Then sExe = App.EXEName & ".EXE"
    
    Show
    DoEvents
    ' Try every which way to make first tab display
    tabAbout.Tab = 0
    DoEvents
    tabAbout.TabVisible(0) = True
    DoEvents
    tabAbout_Click 1
     
    sFilterString = "Executable (*.exe;*.dll;*.vbx;*.ocx;*.fon):*.exe;*.dll;*.vbx;*.ocx;*.fon"
    sFilterString = sFilterString & "Program (*.exe):*.exe" & "|"
    sFilterString = sFilterString & "DLL (*.dll):*.dll" & "|"
    sFilterString = sFilterString & "Control (*.vbx;*.ocx):*.vbx;*.ocx" & "|"
    sFilterString = sFilterString & "Font (*.fon):*.fon"

    Refresh
          
End Sub

Private Sub Form_Activate()
    BugLocalMessage "Activating All About"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    BugLocalMessage "Unloading All About"
End Sub
   
Private Sub Form_Terminate()
    BugLocalMessage "Terminating All About"
End Sub

Private Sub cmdAbout_Click()
    mnuAbout_Click
End Sub

Private Sub cmdExit_Click()
    mnuExit_Click
End Sub

Private Sub cmdDrives_Click()
    fUpdate = True
    tabAbout_Click (0)
End Sub

Private Sub cmdVersion_Click()
    Dim f As Boolean, sFile As String, fReadOnly As Boolean
    f = VBGetOpenFileName( _
            FileName:=sFile, _
            ReadOnly:=fReadOnly, _
            Filter:=sFilterString, _
            Owner:=hWnd)
    If f And sFile <> sEmpty Then
        sExe = sFile
        tabAbout_Click 0
    End If
End Sub

Private Sub mnuAbout_Click()
    Dim about As New CAbout
    With about
        On Error GoTo FailAbout
        ' Set properties
        Set .Client = App
        Set .Icon = Forms(0).Icon
        .UserInfo(2) = "Don't even think " & _
                            "about stealing this program"
        ' Load after all properties are set
        .Load
        ' Modal form will return here when finished
        Exit Sub
    End With
FailAbout:
        MsgBox "I don't know nuttin'"
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub tabAbout_Click(PreviousTab As Integer)
    Dim s As String
    Select Case tabAbout.Tab
    Case ordSystem
        lbl(ordSystem) = ShowSystem
        
    Case ordVideo
        lbl(ordVideo) = ShowVideo
    
    Case ordDrives
        lbl(ordDrives) = "Getting drive information..."
        lbl(ordDrives).Refresh
        lbl(ordDrives) = ShowDrives(fUpdate)
        fUpdate = False
        cmdDrives.Visible = True

    Case ordVersion
        lblFile.Caption = sExe
        lbl(ordVersion) = ShowVersion(sExe)
        
    End Select
    
End Sub

Private Function ShowSystem() As String
    Dim s As String
    s = "Free Physical Memory: " & System.FreePhysicalMemory & " KB" & sCrLf & _
        "Available Physical Memory: " & System.TotalPhysicalMemory & " KB" & sCrLf & _
        "Free Virtual Memory: " & System.FreeVirtualMemory & " KB" & sCrLf & _
        "Available Virtual Memory: " & System.TotalVirtualMemory & " KB" & sCrLf & _
        "Free Page File: " & System.FreePageFile & " KB" & sCrLf & _
        "Available Page File: " & System.TotalPageFile & " KB" & sCrLf & _
        "Memory Load: " & System.MemoryLoad & "%" & sCrLf & _
        "Windows Version: " & System.WinMajor & "." & System.WinMinor & sCrLf & _
        "Processor: " & System.Processor & sCrLf & _
        "Mode: " & System.Mode & sCrLf & _
        "Windows Directory: " & System.WindowsDir & sCrLf & _
        "System Directory: " & System.SystemDir & sCrLf & _
        "User Name: " & System.User & sCrLf & _
        "Machine Name: " & System.Machine & sCrLf
    ShowSystem = s
End Function

Private Function ShowDrives(fUpdate As Boolean) As String

    Static s As String
    If Not fUpdate And s <> sEmpty Then
        ShowDrives = s
        Exit Function
    End If
    
    Dim driveCur As New CDrive
    driveCur = 0       ' Initialize to current drive
    
    Debug.Print driveCur
    s = "Drive information for current drive:" & sCrLf
    Const sBFormat = "#,###,###,##0"
    With driveCur
        s = s & "Drive " & .Root & " [" & .Label & ":" & _
                .Serial & "] (" & .KindStr & ") has " & _
                Format$(.FreeBytes, sBFormat) & " free from " & _
                Format$(.TotalBytes, sBFormat) & sCrLf
    End With
    
    driveCur = "C:\"       ' Initialize to current drive
    
    s = "Drive information for drive C:" & sCrLf
    Debug.Print driveCur
    With driveCur
        s = s & "Drive " & .Root & " [" & .Label & ":" & _
                .Serial & "] (" & .KindStr & ") has " & _
                Format$(.FreeBytes, sBFormat) & " free from " & _
                Format$(.TotalBytes, sBFormat) & sCrLf
    End With
    
    s = s & sCrLf
    s = s & "Drive information for available drives:" & sCrLf
    Dim drives As New CDrives, drive As CDrive
    For Each drive In drives
        With drive
            s = s & "Drive " & .Root & " [" & .Label & ":" & _
                    .Serial & "] (" & .KindStr & ") has " & _
                    Format$(.FreeBytes, sBFormat) & " free from " & _
                    Format$(.TotalBytes, sBFormat) & sCrLf
        End With
    Next
    Debug.Print drives("C:\").Label
    ShowDrives = s
    
End Function
        
Private Function ShowVersion(sFile As String) As String
    Dim vc As New CVersion, s As String
    On Error Resume Next
    vc = sFile
    If Err Or vc.EXEName = sEmpty Then
        ShowVersion = "Can't get version"
        Exit Function
    End If
    
    If vc.Description <> sEmpty Then
        s = s & "Description: " & vc.Description & sCrLf
    End If
    If vc.InternalName <> sEmpty Then
        s = s & "Internal name: " & vc.InternalName & sCrLf
    End If
    If vc.OriginalFilename <> sEmpty Then
        s = s & "Original filename: " & vc.OriginalFilename & sCrLf
    End If
    If vc.TimeStamp <> 0& Then
        s = s & "Time stamp: " & vc.TimeStamp & sCrLf
    End If
    s = s & "File version: " & vc.FullFileVersion & sCrLf
    s = s & "Product version: " & vc.FullProductVersion & sCrLf
    If vc.FileVersionString <> sEmpty Then
        s = s & "File version string: " & vc.FileVersionString & sCrLf
    End If
    If vc.ProductVersionString <> sEmpty Then
        s = s & "Product version string: " & vc.ProductVersionString & sCrLf
    End If
    If vc.ProductName <> sEmpty Then
        s = s & "Product: " & vc.ProductName & sCrLf
    End If
    If vc.Company <> sEmpty Then
        s = s & "Company: " & vc.Company & sCrLf
    End If
    If vc.Comments <> sEmpty Then
        s = s & "Comments: " & vc.Comments & sCrLf
    End If
    If vc.BuildString <> sEmpty Then
        s = s & "Build: " & vc.BuildString & sCrLf
    End If
    If vc.Environment <> sEmpty Then
        s = s & "Environment: " & vc.Environment & sCrLf
    End If
    If vc.ExeType <> sEmpty Then
        s = s & "Executable type: " & vc.ExeType & sCrLf
    End If
    If vc.Copyright <> sEmpty Then
        s = s & "Copyright: " & vc.Copyright & sCrLf
    End If
    If vc.Trademarks <> sEmpty Then
        s = s & "Trademarks: " & vc.Trademarks & sCrLf
    End If
    Dim sT As String
    sT = vc.Custom("CustomBuild")
    If sT <> sEmpty Then
        s = s & "Custom build: " & sT & sCrLf
    End If
    sT = vc.Custom("SpecialBuild")
    If sT <> sEmpty Then
        s = s & "Special build: " & sT & sCrLf
    End If
    Dim dt As Date
    dt = vc.TimeStamp
    If dt <> 0 Then
        s = s & "Time stamp: " & dt & sCrLf
    End If
    ShowVersion = s
End Function

Private Function ShowVideo() As String
    Dim s As String
    With Video
        s = "Display type: " & _
            Choose(.TECHNOLOGY + 1, "Plotter", "Raster Display", _
                "Raster Printer", "Raster Camera", "Character Stream", _
                "Metafile", "Display File") & sCrLf
        s = s & "Screen size: " & .XPixels & "," & .YPixels & sCrLf
        s = s & "Bits per pixel: " & .BitsPerPixel
        s = s & "  Color Planes: " & .ColorPlanes
        s = s & "  Palette size: " & .PaletteSize & sCrLf
        s = s & "Brushes: " & .BrushCount
        s = s & "  Pens: " & .PenCount
        s = s & "  Fonts: " & .FontCount
        s = s & "  Colors: " & .ColorCount & sCrLf
        s = s & "Transparent blits: " & .TransparentBlt & sCrLf
        s = s & "Aspect: X=" & .XAspect & ", Y=" & .YAspect & ", XY=" & .XYAspect & sCrLf
        
        Dim af As Long
        s = s & "Raster: "
        af = .RasterCapability
        If af And RC_BITBLT Then s = s & "BitBlt "
        If af And RC_BITMAP64 Then s = s & "BigBitmaps "
        If af And RC_FLOODFILL Then s = s & "FloodFill "
        If af And RC_PALETTE Then s = s & "Palette "
        If af And RC_STRETCHBLT Then s = s & "StretchBlt "
        If .TransparentBlt Then s = s & "TransparentBlt "
        s = s & sCrLf
        
        s = s & "Curves: "
        af = .CurveCapability
        If af And CC_CIRCLES Then s = s & "Circles "
        If af And CC_PIE Then s = s & "Pie"
        If af And CC_CHORD Then s = s & "Chord "
        If af And CC_ELLIPSES Then s = s & "Ellipses "
        If af And CC_ROUNDRECT Then s = s & "RoundRect "
        s = s & sCrLf
        
        s = s & "Lines: "
        af = .LineCapability
        If af And LC_POLYLINE Then s = s & "PolyLine "
        If af And LC_MARKER Then s = s & "Marker "
        If af And LC_POLYMARKER Then s = s & "PolyMarker "
        s = s & sCrLf
        
        s = s & "Polygons: "
        af = .PolygonCapability
        If af And PC_POLYGON Then s = s & "Polygon "
        If af And PC_RECTANGLE Then s = s & "Rectangle "
        If af And PC_WINDPOLYGON Then s = s & "WindPolygon "
        If af And PC_SCANLINE Then s = s & "ScanLine "
        s = s & sCrLf
        
        s = s & "Text: "
        af = .TextCapability
        If af And TC_CR_90 Then s = s & "Rotate 90"
        If af And TC_CR_ANY Then s = s & "RotateAny "
        If af And TC_IA_ABLE Then s = s & "Italic "
        If af And TC_UA_ABLE Then s = s & "Underline "
        If af And TC_SO_ABLE Then s = s & "StrikeOut "
        If af And TC_RA_ABLE Then s = s & "Raster "
        If af And TC_VA_ABLE Then s = s & "Vector "
        s = s & sCrLf
               
    End With
    ShowVideo = s
End Function
