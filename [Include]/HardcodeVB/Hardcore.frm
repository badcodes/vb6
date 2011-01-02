VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FHardCore 
   AutoRedraw      =   -1  'True
   Caption         =   "Hardcore Samples"
   ClientHeight    =   5364
   ClientLeft      =   3432
   ClientTop       =   3768
   ClientWidth     =   8628
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   7.8
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Hardcore.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5364
   ScaleWidth      =   8628
   WhatsThisHelp   =   -1  'True
   Begin MSComDlg.CommonDialog dlgOpenPic 
      Left            =   6840
      Top             =   7080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdGuid 
      Caption         =   "GUIDs"
      Height          =   495
      Left            =   120
      TabIndex        =   11
      Top             =   3120
      Width           =   975
   End
   Begin VB.OptionButton optOpen 
      Caption         =   "Picture Form"
      Height          =   264
      Index           =   2
      Left            =   120
      TabIndex        =   10
      Top             =   4980
      Width           =   1425
   End
   Begin VB.OptionButton optOpen 
      Caption         =   "Control"
      Height          =   264
      Index           =   1
      Left            =   120
      TabIndex        =   9
      Top             =   4716
      Width           =   972
   End
   Begin VB.OptionButton optOpen 
      Caption         =   "API"
      Height          =   264
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   4440
      Value           =   -1  'True
      Width           =   972
   End
   Begin VB.TextBox txtTest 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4272
      Left            =   1560
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   7
      Top             =   108
      Width           =   5940
   End
   Begin VB.CommandButton cmdExe 
      Caption         =   "Exe Type"
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton cmdMenus 
      Caption         =   "Menus"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Open File"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   3720
      Width           =   975
   End
   Begin VB.PictureBox pbBitmap 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   1356
      ScaleHeight     =   852
      ScaleWidth      =   888
      TabIndex        =   3
      Top             =   108
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.CommandButton cmdBlob 
      Caption         =   "Blob"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   975
   End
   Begin VB.CommandButton cmdWin32 
      Caption         =   "Win32"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton cmdBits 
      Caption         =   "Bits"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open..."
      End
      Begin VB.Menu mnuDead 
         Caption         =   "&Dead"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuGone 
         Caption         =   "&Gone"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuCut 
         Caption         =   "Cu&t"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuView 
         Caption         =   "&View"
         WindowList      =   -1  'True
         Begin VB.Menu mnuSome 
            Caption         =   "&Some"
         End
         Begin VB.Menu mnuAll 
            Caption         =   "&All"
         End
      End
      Begin VB.Menu mnuCheck 
         Caption         =   "&Check"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuContents 
         Caption         =   "&Contents"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About..."
      End
   End
End
Attribute VB_Name = "FHardCore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type ChunkType
    abData(0 To 255) As Byte
End Type
    
Private Type TestType
    lp As String
    l As Long
End Type

Private aEmpty(0) As Variant

Private hWndTB As Integer
Private hWndSB As Integer

Const SB_HORZ = 0
Const SB_VERT = 1
Const SB_CTL = 2

Private Sub Form_Load()
    ChDrive App.Path
    ChDir App.Path
    Show
End Sub

Private Sub cmdBlob_Click()
    Dim s As String, i As Integer, ab() As Byte, ab2() As Byte
    Dim sMsg As String
    
    sMsg = sMsg & "Type library: " & _
                  IIf(UnicodeTypeLib, "Unicode", "ANSI") & sCrLf
    sMsg = sMsg & "Assign string to byte array and byte string" & sCrLf
    StrToBytes ab, "1234567890"
    s = ab
    sMsg = sMsg & "Byte string as byte array: " & BytesToStr(ab) & sCrLf
    sMsg = sMsg & "Length: " & LenBytes(ab) & sCrLf
    sMsg = sMsg & "Byte string as string: " & StrBToStr(s) & sCrLf
    sMsg = sMsg & "Byte length: " & LenB(s) & sCrLf
    sMsg = sMsg & "String length: " & Len(s) & sCrLfCrLf

    sMsg = sMsg & "Read and insert numbers from byte array" & sCrLf
    sMsg = sMsg & "Word (from string) at &H4: &H" & FmtHex(WordFromStrB(s, 4)) & sCrLf
    sMsg = sMsg & "Word (from bytes) at &H4: &H" & FmtHex(BytesToWord(ab, 4)) & sCrLf
    sMsg = sMsg & "DWord at &H4: &H" & FmtHex(BytesToDWord(ab, 4)) & sCrLf
    BytesFromWord &H7372, ab, 4
    sMsg = sMsg & "Insert &H7372 at &H4: " & BytesToStr(ab) & sCrLf
    BytesFromDWord &H65666768, ab, 4
    sMsg = sMsg & "Insert &H65666768 at &H4: " & BytesToStr(ab) & sCrLf
    sMsg = sMsg & "Word at &H4: &H" & FmtHex(BytesToWord(ab, 4)) & sCrLf
    sMsg = sMsg & "DWord at &H4: &H" & FmtHex(BytesToDWord(ab, 4)) & sCrLfCrLf

    sMsg = sMsg & "Extract and insert strings on byte array" & sCrLf
    ab = StrToBytesV("1234567890")
    s = ab
    sMsg = sMsg & "Left 3: " & LeftBytes(ab, 3) & sCrLf
    sMsg = sMsg & "Right 3: " & RightBytes(ab, 3) & sCrLf
    sMsg = sMsg & "From &H5: " & MidBytes(ab, 5) & sCrLf
    sMsg = sMsg & "From &H5 length 2: " & MidBytes(ab, 5, 2) & sCrLfCrLf
    InsBytes "ABC", ab, 2
    sMsg = sMsg & "Insert 'ABC' at &h2: " & BytesToStr(ab) & sCrLf
#If 0 Then
    ' This is legal, but textbox doesn't like it
    InsBytes "ABC", ab, 4, 4
    sMsg = sMsg & "Insert 'ABC' at &H4 in field of 4: " & BytesToStr(ab) & sCrLf
#End If
    sMsg = sMsg & "From &H4 length 5 to null: " & MidBytes(ab, 4, 5, True) & sCrLf
    sMsg = sMsg & "From &H4 length 5: " & MidBytes(ab, 4, 5) & sCrLfCrLf
    FillBytes ab, Asc(" "), 5, 4
    sMsg = sMsg & "Insert spaces at &H5 in field of 4: " & BytesToStr(ab) & sCrLfCrLf

    ' Test asserts (uncomment for tests)
#If 0 Then
    InsBytes "ABC", ab, 9
    sMsg = sMsg & "Insert 'ABC' at position 9: " & BytesToStr(ab) & sCrLf
    InsBytes "ABC", ab, 8, 4
    sMsg = sMsg & "Insert 'ABC' at position 8 in field of 4: " & BytesToStr(ab) & sCrLf
    FillBytes ab, Asc(" "), 7, 4
    sMsg = sMsg & "Insert spaces at position 7 in field of 4: " & BytesToStr(ab) & sCrLf
#End If

    sMsg = sMsg & "Find string in byte array" & sCrLf
    StrToBytes ab, "1234567890"
    s = ab
    ab2 = StrToBytesV("56")
    sMsg = sMsg & "56 at position: " & InStrB(ab, ab2) & sCrLf
    sMsg = sMsg & "56 at position: " & InStrB(s, ab2) & sCrLf

    StrToBytes ab, "1234567890"
    s = ab
    sMsg = sMsg & "Hex dump of byte arrays, byte strings, strings" & sCrLf
    sMsg = sMsg & "Dump byte array: " & sCrLf & HexDump(ab, False) & sCrLf
    sMsg = sMsg & "Dump byte string: " & sCrLf & HexDumpB(s, False) & sCrLf
    sMsg = sMsg & "Dump string: " & sCrLf & HexDumpS(s, False) & sCrLfCrLf
    
    sMsg = sMsg & "ANSI characters that don't match Unicode versions" & sCrLf
    For i = 0 To 255
        If AscW(Chr$(i)) <> i Then
            sMsg = sMsg & "ANSI: &H" & FmtHex(i, 2) & sTab
            sMsg = sMsg & "  Unicode: &H" & FmtHex(AscW(Chr$(i)), 4) & sTab
            sMsg = sMsg & "  Character: " & Chr$(i) & sCrLf
        End If
    Next
    sMsg = sMsg & sCrLf

    ' Open first file for processing
    Dim sBinFile As String, nBinFile As Integer
    Dim sBin As String, abBin() As Byte
    sBinFile = Dir("*.*")
    nBinFile = FreeFile
    Open sBinFile For Binary Access Read Write Lock Write As #nBinFile
    ReDim abBin(LOF(nBinFile))
    Get #nBinFile, 1, abBin
    sBin = abBin
    sMsg = sMsg & "Open file " & sBinFile & " and process as byte string or byte array" & sCrLf
    sMsg = sMsg & "Dump first 20 byte characters: " & sCrLf
    sMsg = sMsg & HexDumpB(MidB$(sBin, 1, 20)) & sCrLf
    sMsg = sMsg & "Dump first 20 bytes: " & sCrLf
    sMsg = sMsg & HexDump(MidBytes(abBin, 0, 20)) & sCrLf
    abBin = sBin
    Put #nBinFile, 1, abBin
    Close #nBinFile
    
    BugMessage sMsg
    txtTest.Text = sMsg
End Sub

Private Sub cmdBits_Click()
    txtTest.Visible = True
    pbBitmap.Visible = False
    Dim dw As Long, w As Integer, r As Single, d As Double
    Dim c As Currency, s As String, i As Integer
    Dim pl As Long, PI As Long, pr As Long, pd As Long
    Dim pc As Long, ps As Long, psz As Long
    Dim sOutput As String
    sOutput = ""

    w = &HABCD
    dw = &HFEDCBA98
    'dw = &HFFFF0000
    r = 1.23456789
    d = 9.87654321
    c = 999.99
    s = "Test"

    Dim bHi As Byte, bLo As Byte
    Dim wHi As Integer, wLo As Integer
    Dim wPack  As Integer, dwPack  As Long
    Dim wRShift As Integer, wLShift As Integer
    Dim dwRShift As Long, dwLShift As Long

#If 1 Then
    bLo = LoByte(w)
    sOutput = sOutput & "Low byte of word (" & Hex$(w) & "): " & Hex$(bLo) & sCrLf
    bHi = HiByte(w)
    sOutput = sOutput & "High byte of word (" & Hex$(w) & "): " & Hex$(bHi) & sCrLf
    wPack = MakeWord(bHi, bLo)
    sOutput = sOutput & "Packed hi/lo bytes of word: " & Hex$(wPack) & sCrLf
    wLo = LoWord(dw)
    sOutput = sOutput & "Low Word of DWord (" & Hex$(dw) & "): " & Hex$(wLo) & sCrLf
    wHi = HiWord(dw)
    sOutput = sOutput & "High Word of DWord (" & Hex$(dw) & "): " & Hex$(wHi) & sCrLf
    dwPack = MakeDWord(wHi, wLo)
    sOutput = sOutput & "Packed hi/lo Word of DWord: " & Hex$(dwPack) & sCrLf
#End If
    
#If 1 Then
    sOutput = sOutput & "Word shifted right" & sCrLf
    For i = 0 To 15
        sOutput = sOutput & Hex$(RShiftWord(w, i)) & "  "
    Next
    sOutput = sOutput & sCrLf
    sOutput = sOutput & "Word shifted left" & sCrLf
    For i = 0 To 15
        sOutput = sOutput & Hex$(LShiftWord(w, i)) & "  "
    Next
    sOutput = sOutput & sCrLf
    sOutput = sOutput & "DWord shifted right C" & sCrLf
    dw = &H70000000
    For i = 0 To 31
        sOutput = sOutput & Hex$(RShiftDWord(dw, i)) & "  "
    Next
    sOutput = sOutput & sCrLf
    sOutput = sOutput & "DWord shifted left C" & sCrLf
    dw = 1
    For i = 0 To 31
        sOutput = sOutput & Hex$(LShiftDWord(dw, i)) & "  "
    Next
    sOutput = sOutput & sCrLf
#End If
    
    w = &H1234
    dw = &H12345678
#If 1 Then
    bLo = LoByte(w)
    sOutput = sOutput & "Low byte of word (" & Hex$(w) & "): " & Hex$(bLo) & sCrLf
    bHi = HiByte(w)
    sOutput = sOutput & "High byte of word (" & Hex$(w) & "): " & Hex$(bHi) & sCrLf
    wPack = MakeWord(bHi, bLo)
    sOutput = sOutput & "Packed hi/lo bytes of word: " & Hex$(wPack) & sCrLf
    wLo = LoWord(dw)
    sOutput = sOutput & "Low Word of DWord (" & Hex$(dw) & "): " & Hex$(wLo) & sCrLf
    wHi = HiWord(dw)
    sOutput = sOutput & "High Word of DWord (" & Hex$(dw) & "): " & Hex$(wHi) & sCrLf
    dwPack = MakeDWord(wHi, wLo)
    sOutput = sOutput & "Packed hi/lo Word of DWord: " & Hex$(dwPack) & sCrLf
#End If
    
#If 1 Then
    sOutput = sOutput & "Word shifted right" & sCrLf
    For i = 0 To 15
        sOutput = sOutput & Hex$(RShiftWord(w, i)) & "  "
    Next
    sOutput = sOutput & sCrLf
    sOutput = sOutput & "Word shifted left" & sCrLf
    For i = 0 To 15
        sOutput = sOutput & Hex$(LShiftWord(w, i)) & "  "
    Next
    sOutput = sOutput & sCrLf
    sOutput = sOutput & "DWord shifted right C" & sCrLf
    dw = &H70000000
    For i = 0 To 31
        sOutput = sOutput & Hex$(RShiftDWord(dw, i)) & "  "
    Next
    sOutput = sOutput & sCrLf
    sOutput = sOutput & "DWord shifted left C" & sCrLf
    dw = 1
    For i = 0 To 31
        sOutput = sOutput & Hex$(LShiftDWord(dw, i)) & "  "
    Next
    sOutput = sOutput & sCrLf
#End If
    
    Dim secStart As Currency, sec As Currency
    ProfileStart secStart
    dw = 50
    For i = 1 To 5000
        dw = RShiftDWord(50, 7)
    Next
    ProfileStop secStart, sec
    sOutput = sOutput & "5000 shifts: " & sec & " seconds" & sCrLf
        
    
    BugMessage sOutput
    txtTest.Text = sOutput

End Sub

Private Sub cmdExe_Click()
    txtTest.Visible = True
    pbBitmap.Visible = False

    Const sFilter = "Executables (*.EXE;*.DLL;*.OCX)|*.exe;*.dll;*.ocx|" & _
                    "EXE Files|*.exe|" & _
                    "DLL Files(*.DLL;*.OCX)|*.dll;*.ocx|" & _
                    "All Files (*.*)|*.*"
    Static iFilterIndex As Long, sFile As String, sInitDir As String
    Dim f As Boolean
    If sInitDir = sEmpty Then sInitDir = WindowsDir
    f = VBGetOpenFileName( _
        FileName:=sFile, _
        Flags:=cdlOFNFileMustExist Or cdlOFNHideReadOnly, _
        InitDir:=sInitDir, _
        Filter:=sFilter, _
        FilterIndex:=iFilterIndex)
    sInitDir = GetFileDir(sFile)
    sFile = GetFileBaseExt(sFile)
    txtTest = "EXE type of " & UCase$(sFile) & ": " & ExeTypeStr(sFile)
End Sub

Private Sub cmdMenus_Click()
    Dim menu As New CMenuList, item As CMenuItem
    
    txtTest.Visible = True
    pbBitmap.Visible = False
    txtTest = "Some tests of a perfectly good class from the first " & _
              "edition" & vbCrLf & "that didn't make the grade for the " & _
              "the second edition:" & vbCrLf & vbCrLf
    Call menu.Create(Me.hWnd)
    menu.Walk
    Dim s As String
    s = InputBox("Enter menu item to find: ")
    If Not menu.Find(s, item) Then
        MsgBox "Can't find item: " & s
        Exit Sub
    End If
    With item
        s = "Name: " & .Name & sCrLf
        s = s & "Text: " & .Text & sCrLf & "State: "
        s = s & IIf(.Disabled, "Disabled ", sEmpty)
        s = s & IIf(.Checked, "Checked ", sEmpty)
        s = s & IIf(.Grayed, "Grayed ", sEmpty)
        s = s & IIf(.Popup, "Popup ", sEmpty) & sCrLf
        txtTest = txtTest & s
    End With
    Call item.Execute
    menu.Refresh
    Dim f As Boolean
    If menu.Find("Dead", item) Then
        BugMessage item.Disabled
        BugMessage item.Grayed
        item.Text = "&Live"
        BugMessage item.Disabled
        BugMessage item.Grayed
        item.Disabled = False
        BugMessage item.Disabled
        BugMessage item.Grayed
    ElseIf menu.Find("Live", item) Then
        BugMessage item.Disabled
        BugMessage item.Grayed
        item.Text = "&Dead"
        BugMessage item.Disabled
        BugMessage item.Grayed
        item.Disabled = True
        BugMessage item.Disabled
        BugMessage item.Grayed
    End If

    Dim SysMenu As New CMenuList
    f = SysMenu.Create(Me.hWnd, True)
    SysMenu.Walk
    If WindowState = vbMaximized Then
        f = SysMenu.Find("Restore", item)
    Else
        f = SysMenu.Find("Maximize", item)
    End If
    'f = SysMenu.Find("Switch To", item)
    f = item.Execute
        
End Sub

Private Sub cmdOpen_Click()
    txtTest.Visible = False
    pbBitmap.Visible = True
    
    Const sFilters = "All Picture Files|*.bmp;*.dib;*.ico;*.wmf;*.cur|" & _
                     "Bitmaps (*.BMP;*.DIB)|*.bmp;*.dib|" & _
                     "Metafiles (*.WMF)|*.wmf|" & _
                     "Icons (*.ICO)|*.ico|" & _
                     "Cursors (*.CUR;*.ICO)|*.cur;*.ico|" & _
                     "All Files (*.*)|*.*"
    Select Case GetOption(optOpen)
    Case 0
        Dim sFilter As String
        sFilter = sFilters
        Dim sFile As String, f As Boolean
        f = VBGetOpenFileName( _
            FileName:=sFile, _
            InitDir:=WindowsDir, _
            Flags:=cdlOFNFileMustExist Or cdlOFNHideReadOnly, _
            Filter:=sFilter) ' *.bmp;*.dib;*.ico;*.wmf;*.cur
        If f And sFile <> sEmpty Then
            ' Ignore any errors, but file load will fail
            On Error Resume Next
            pbBitmap.Picture = LoadPicture(sFile)
            On Error GoTo 0
        End If
    Case 1
        With dlgOpenPic
            .InitDir = WindowsDir
            .Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly
            .Filter = sFilters ' *.bmp;*.dib;*.ico;*.wmf;*.cur
            .ShowOpen
            If .FileName <> sEmpty Then
                ' Ignore any errors, but file load will fail
                On Error Resume Next
                pbBitmap.Picture = LoadPicture(.FileName)
                On Error GoTo 0
            End If
        End With
    Case 2
        Dim opfile As New COpenPictureFile
        With opfile
            .InitDir = WindowsDir
            .Load Left + (Width / 4), Top + (Height / 4)
            If .FileName <> sEmpty Then
                ' Ignore any errors, but file load will fail
                On Error Resume Next
                pbBitmap.Picture = LoadPicture(.FileName)
                On Error GoTo 0
            End If
        End With
    End Select

End Sub

Private Sub cmdGuid_Click()

    Dim sUUID As String, sCLSID As String, sIID As String
    Dim sProgID As String, sOut As String
    Dim uid As UUID, cid As UUID, id As UUID
    
    sOut = sOut & "Get CLSID from ProgID" & sCrLf
    sProgID = "COMCTL.Toolbar"
    sOut = sOut & "ProgID: " & sProgID & sCrLf
    CLSIDFromProgID sProgID, cid
    sOut = sOut & "Call CLSIDFromProgID" & sCrLf
    sOut = sOut & "First long from CLSID: " & Hex$(cid.Data1) & sCrLf
    sCLSID = CLSIDToString(cid)
    sOut = sOut & "Call CLSIDToString" & sCrLf
    sOut = sOut & "CLSID string: " & sCLSID & sCrLf & sCrLf
    
    sOut = sOut & "Get ProgID from CLSID" & sCrLf
    sCLSID = "{58DA8D8F-9D6A-101B-AFC0-4210102A8DA7}"
    sOut = sOut & "CLSID string: " & sCLSID & sCrLf
    sOut = sOut & "Call CLSIDFromString" & sCrLf
    CLSIDFromString sCLSID, cid
    sOut = sOut & "First long from CLSID: " & Hex$(cid.Data1) & sCrLf
    sOut = sOut & "Call ProgIDFromCLSID" & sCrLf
    sProgID = CLSIDToProgID(cid)
    sOut = sOut & "ProgID: " & sProgID & sCrLf & sCrLf
    
    sOut = sOut & "Get IID from IID string" & sCrLf
    With id
        .Data1 = &H12345678
        .Data2 = &H1234
        .Data3 = &H5678
        .Data4(0) = &H11: .Data4(1) = &H22: .Data4(2) = &H33: .Data4(3) = &H44
        .Data4(4) = &H55: .Data4(5) = &H66: .Data4(6) = &H77: .Data4(7) = &H88
    End With
    sOut = sOut & "First long from IID: " & Hex$(id.Data1) & sCrLf
    sOut = sOut & "Call IIDToString" & sCrLf
    sIID = IIDToString(id)
    sOut = sOut & "IID string: " & sIID & sCrLf & sCrLf
    
    sOut = sOut & "Get IID from IID string" & sCrLf
    sIID = "{87654321-8765-1234-ABCD-887766554433}"
    sOut = sOut & "IID string: " & sCLSID & sCrLf
    sOut = sOut & "Call IIDFromString" & sCrLf
    IIDFromString sCLSID, id
    sOut = sOut & "First long from IID: " & Hex$(id.Data1) & sCrLf & sCrLf
    
    sOut = sOut & "Create a new GUID with CoCreateGuid" & sCrLf
    CoCreateGuid uid
    sOut = sOut & "First long from GUID: " & Hex$(uid.Data1) & sCrLf
    sUUID = GUIDToString(uid)
    sOut = sOut & "Call GUIDToString" & sCrLf
    sOut = sOut & "GUID string: " & sUUID & sCrLf & sCrLf
    
    sOut = sOut & "Create object from CLSID string" & sCrLf
    Dim sCLSIDDrive As String, clsidDrive As UUID
    Dim sIIDDrive As String, iidDrive As UUID, obj As Object
    sCLSIDDrive = "{630B8825-E48F-11D0-B253-00AA005754FD}"
    CLSIDFromString sCLSIDDrive, clsidDrive
    sIIDDrive = "{B2142BB8-398C-11D2-81F0-00C0F030B702}"
    IIDFromString sIIDDrive, iidDrive
    sOut = sOut & "First long from IID: " & Hex$(clsidDrive.Data1) & sCrLf
    
    txtTest = sOut
End Sub

' These are in VBCore for VB6, but no public UDTs in VB5
#If iVBVer <= 5 Then
Private Function CLSIDToString(cid As UUID) As String
    Dim pStr As Long
    StringFromCLSID cid, pStr
    CLSIDToString = PointerToString(pStr)
    CoTaskMemFree pStr
End Function

Private Function IIDToString(id As UUID) As String
    Dim pStr As Long
    StringFromIID id, pStr
    IIDToString = PointerToString(pStr)
    CoTaskMemFree pStr
End Function

Private Function GUIDToString(uid As UUID) As String
    Dim s As String, c As Long
    c = 40
    s = String$(c, 0)
    c = StringFromGUID2(uid, s, c)
    GUIDToString = Left$(s, c - 1)
End Function

Private Function CLSIDToProgID(cid As UUID) As String
    Dim pStr As Long
    ProgIDFromCLSID cid, pStr
    CLSIDToProgID = PointerToString(pStr)
End Function
#End If

Private Sub cmdWin32_Click()

    Dim i As Integer, s As String, sVal As String
    Dim sName As String, sFullName As String
    Dim c As Long, f As Boolean
    Dim iDir As Long, iBase As Long, iExt As Long
    
    txtTest.Visible = True
    pbBitmap.Visible = False
    sName = GetTempDir()
    
    ' Test ExistFile
    s = "Test ExistFile" & sCrLf & sCrLf
    sName = Environ$("COMSPEC")
    s = s & "File " & sName & " exists: " & ExistFile(sName) & sCrLf
    sName = "nosuch.txt"
    s = s & "File " & sName & " exists: " & ExistFile(sName) & sCrLf
    
    ' Test GetFullPathName
    s = s & sCrLf & "Test GetFullPathName" & sCrLf & sCrLf
    Dim sBase As String, pBase As Long
    sFullName = String$(cMaxPath, 0)
    c = GetFullPathName(sName, cMaxPath, sFullName, pBase)
    sFullName = Left$(sFullName, c)
    If c Then s = s & "Full name: " & sFullName & sCrLf
    ' Can't use pBase because pointer is to temporary Unicode string
 
#If 1 Then
    s = s & sCrLf & "Test GetFullPath with invalid argument" & sCrLf & sCrLf
    sFullName = GetFullPath("", iBase, iExt, iDir)
    If sFullName = sEmpty Then
        s = s & "Failed: Error " & Err.LastDllError & sCrLf
    Else
        s = s & "File: " & sFullName & sCrLf
    End If
#End If

    s = s & sCrLf & "Test GetFullPath with all arguments" & sCrLf & sCrLf
    sFullName = GetFullPath(sName, iBase, iExt, iDir)
    If sFullName <> sEmpty Then
        s = s & "Relative file: " & sName & sCrLf
        s = s & "Full name: " & sFullName & sCrLf
        s = s & "File: " & Mid$(sFullName, iBase) & sCrLf
        s = s & "Extension: " & Mid$(sFullName, iExt) & sCrLf
        s = s & "Base name: " & Mid$(sFullName, iBase, _
                                     iExt - iBase) & sCrLf
        s = s & "Drive: " & Left$(sFullName, iDir - 1) & sCrLf
        s = s & "Directory: " & Mid$(sFullName, iDir, _
                                     iBase - iDir) & sCrLf
        s = s & "Path: " & Left$(sFullName, iBase - 1) & sCrLf
    Else
        s = s & "Invalid name: " & sName
    End If
        
    sFullName = GetFullPath(sName, iBase, iExt, iDir)
    sFullName = GetFullPath(sName, iBase, iExt)
    sFullName = GetFullPath(sName, iBase)
    sFullName = GetFullPath(sName)
    sFullName = GetFullPath(sName, , iExt)
    sFullName = GetFullPath(sName, , iExt, iDir)
    sFullName = GetFullPath(sName, , , iDir)
    sFullName = GetFullPath(sName, iBase, , iDir)
    
    Dim sPart As String
    sName = "Hardcore.frm"
    sPart = GetFullPath(sName)      ' C:\Hardcore\Hardcore.frm
    sPart = GetFileBase(sName)      ' Hardcore
    sPart = GetFileBaseExt(sName)   ' Hardcore.frm
    sPart = GetFileExt(sName)       ' .frm
    sPart = GetFileDir(sName)       ' C:\Hardcore\


    s = s & sCrLf & "Test GetFullPath with some arguments" & sCrLf & sCrLf
    sFullName = GetFullPath(sName, iBase, iExt)
    If sFullName <> sEmpty Then
        s = s & "Relative file: " & sName & sCrLf
        s = s & "Full name: " & sFullName & sCrLf
        s = s & "File: " & Mid$(sFullName, iBase) & sCrLf
        s = s & "Extension: " & Mid$(sFullName, iExt) & sCrLf
        s = s & "Base name: " & Mid$(sFullName, iBase, _
                                     iExt - iBase) & sCrLf
        s = s & "Path: " & Left$(sFullName, iBase - 1) & sCrLf
    Else
        s = s & "Invalid name: " & sName
    End If
    
    s = s & sCrLf & "Test GetFullPath with no optional arguments" & sCrLf & sCrLf
    sFullName = GetFullPath(sName)
    If sFullName <> sEmpty Then
        s = s & "Relative file: " & sName & sCrLf
        s = s & "Full name: " & sFullName & sCrLf
    Else
        s = s & "Invalid name: " & sName
    End If
    
    ' Test SearchPath
    s = s & sCrLf & "Test SearchPath" & sCrLf
    sName = "c2.exe"
    sFullName = String$(cMaxPath, 0)
    i = SearchPath(vbNullString, sName, vbNullString, cMaxPath, sFullName, pBase)
    sFullName = Left$(sFullName, i)
    If i Then
        s = s & "File " & sName & " found in: " & sFullName & sCrLf
        ' Can't use pBase because pointer is to temporary Unicode string
    Else
        s = s & "File " & sName & " not found" & sCrLf
    End If
    
    s = s & sCrLf & "Test SearchDirs(""calc"", "".exe"", , iBase, iExt, iDir)" & sCrLf
    sFullName = SearchDirs("calc", ".exe", , iBase, iExt, iDir)
    If sFullName <> sEmpty Then
        s = s & "File found in: " & sFullName & sCrLf
    Else
        s = s & "File not found" & sCrLf
    End If
    
    s = s & sCrLf & "Test SearchDirs(""calc.exe"", , , iBase, iExt)" & sCrLf
    sFullName = SearchDirs("calc.exe", , , iBase, iExt)
    If sFullName <> sEmpty Then
        s = s & "File found in: " & sFullName & sCrLf
    Else
        s = s & "File not found" & sCrLf
    End If
    
    s = s & sCrLf & "Test SearchDirs(""calc"", "".exe"", Environ(""PATH""), iBase)" & sCrLf
    sFullName = SearchDirs("calc", ".exe", Environ("PATH"), iBase)
    If sFullName <> sEmpty Then
        s = s & "File found in: " & sFullName & sCrLf
    Else
        s = s & "File not found" & sCrLf
    End If
    
    s = s & sCrLf & "Test SearchDirs(""calc.exe"")" & sCrLf
    sFullName = SearchDirs("calc.exe")
    If sFullName <> sEmpty Then
        s = s & "File found in: " & sFullName & sCrLf
    Else
        s = s & "File not found" & sCrLf
    End If
    
    s = s & sCrLf & "Test SearchDirs with different files" & sCrLf & sCrLf
       
    sName = "link.exe"
    sFullName = SearchDirs(sName, sEmpty, sEmpty, iBase, iExt, iDir)
    If sFullName <> sEmpty Then
        s = s & "Found file " & sName
        s = s & " in " & sFullName & sCrLf
        s = s & "File: " & Mid$(sFullName, iBase) & sCrLf
        s = s & "Extension: " & Mid$(sFullName, iExt) & sCrLf
        s = s & "Base name: " & Mid$(sFullName, iBase, _
                                     iExt - iBase) & sCrLf
        s = s & "Drive: " & Left$(sFullName, iDir - 1) & sCrLf
        s = s & "Directory: " & Mid$(sFullName, iDir, _
                                     iBase - iDir) & sCrLf
        s = s & "Path: " & Left$(sFullName, iBase - 1) & sCrLf
    Else
        s = s & "File " & sName & " not found" & sCrLf
    End If

    sName = "hardcore.frm"
    sFullName = SearchDirs(sName)
    If sFullName <> sEmpty Then
        s = s & "File " & sName & " found in: " & sFullName & sCrLf
    Else
        s = s & "File " & sName & " not found" & sCrLf
    End If
    
    sName = "calc.exe"
    sFullName = SearchDirs(sName)
    If sFullName <> sEmpty Then
        s = s & "File " & sName & " found in: " & sFullName & sCrLf
    Else
        s = s & "File " & sName & " not found" & sCrLf
    End If
    
    sName = "gdi32.dll"
    sFullName = SearchDirs(sName)
    If sFullName <> sEmpty Then
        s = s & "File " & sName & " found in: " & sFullName & sCrLf
    Else
        s = s & "File " & sName & " not found" & sCrLf
    End If

    sName = "find.exe"
    sFullName = SearchDirs(sName)
    If sFullName <> sEmpty Then
        s = s & "File " & sName & " found in: " & sFullName & sCrLf
    Else
        s = s & "File " & sName & " not found" & sCrLf
    End If

    sFullName = SearchDirs("WINDOWS.H", , Environ("INCLUDE"))
    sName = "WINDOWS.H"
    sFullName = SearchDirs(sName, , Environ("INCLUDE"))
    If sFullName <> sEmpty Then
        s = s & "File " & sName & " found in: " & sFullName & sCrLf
    Else
        s = s & "File " & sName & " not found" & sCrLf
    End If
    
    sFullName = SearchDirs("DEBUG.BAS", , ".")
    sName = "DEBUG.BAS"
    sFullName = SearchDirs(sName, , ".")
    If sFullName <> sEmpty Then
        s = s & "File " & sName & " found in: " & sFullName & sCrLf
    Else
        s = s & "File " & sName & " not found" & sCrLf
    End If
          
    sFullName = GetFullPath("DEBUG.BAS")
    
    sName = "EDIT"
    Dim asExts(1 To 4) As String
    asExts(1) = ".EXE": asExts(2) = ".COM"
    asExts(3) = ".BAT": asExts(4) = ".PIF"
    For i = 1 To 4
        sFullName = SearchDirs(sName, asExts(i))
        If sFullName <> sEmpty Then Exit For
    Next
    If sFullName <> sEmpty Then
        s = s & "File found in: " & sFullName & sCrLf
    Else
        s = s & "File " & sName & " not found" & sCrLf
    End If
    
    ' Test GetDiskFreeSpace and GetDriveType
    s = s & sCrLf & "Test GetDiskFreeSpace and GetDriveType" & sCrLf & sCrLf
    Dim iSectors As Long, iBytes As Long
    Dim iFree As Long, iTotal As Long
    Dim rFree As Double, rTotal As Double
    sName = "%:\"
    Dim sTab As String
    For i = 1 To 26
        sVal = Chr$(i + Asc("A") - 1)
        Mid$(sName, 1, 1) = sVal

        c = GetDriveType(sName)
        s = s & "Disk " & sVal & " type: "
        s = s & Choose(c + 1, "Unknown", "Invalid", "Floppy ", _
                              "Hard   ", "Network", "CD-ROM ", "RAM    ")

        f = GetDiskFreeSpace(sName, iSectors, iBytes, iFree, iTotal)
        rFree = iSectors * iBytes * CDbl(iFree)
        rTotal = iSectors * iBytes * CDbl(iTotal)
        If f Then
            s = s & " with " & Format$(rFree, "#,###,###,##0")
            s = s & " free from " & Format$(rTotal, "#,###,###,##0") & sCrLf
        Else
            s = s & sCrLf
        End If
    Next
    ' txtTest.Text = s

    ' Test GetTempPath and GetTempFileName
    s = s & sCrLf & "Test GetTempPath and GetTempFileName" & sCrLf & sCrLf
    c = cMaxPath
    sFullName = String$(c, 0)
    c = GetTempPath(c, sFullName)
    sFullName = Left$(sFullName, c)
    s = s & "Temp Path: " & sFullName & sCrLf
    sFullName = String$(cMaxPath, 0)
    Call GetTempFileName(".", "HC", 0, sFullName)
    sFullName = Left$(sFullName, InStr(sFullName, vbNullChar) - 1)
    s = s & "Temp File: " & sFullName & sCrLf
    
    s = s & sCrLf & "Test GetTempFile and GetTempDir" & sCrLf & sCrLf
    ' Get temp file for current directory
    sFullName = GetTempFile("VB", ".")
    s = s & "Temp file in current directory: " & sFullName & sCrLf
    ' Get temp file for TEMP directory
    sFullName = GetTempFile("VB", GetTempDir)
    ' Get temp file for TEMP directory default
    sFullName = GetTempFile("VB")
    ' Get temp file for TEMP directory with no prefix
    sFullName = GetTempFile
    s = s & "Temp file in TEMP directory: " & sFullName & sCrLf
    sFullName = GetTempFile
    s = s & "Temp file with defaults (no prefix, TEMP directory): " & sFullName & sCrLf
    sFullName = GetTempFile("HC")
    s = s & "Temp file with path default (TEMP directory): " & sFullName & sCrLf
    sFullName = GetTempFile(, ".")
    s = s & "Temp file with prefix default (no prefix directory): " & sFullName & sCrLf

   ' Test GetLogicalDrives
    s = s & sCrLf & "Test GetLogicalDrives" & sCrLf & sCrLf
    sVal = VBGetLogicalDrives()
    s = s & "Drives    ABCDEFGHIJKLMNOPQRSTUVWXYZ" & sCrLf
    s = s & "Drives    " & sVal & sCrLf

    On Error Resume Next
'    Kill "~HC*.tmp"
'    Kill "HC*.tmp"
    On Error GoTo 0
    
    BugMessage s
    txtTest.Text = s
    
End Sub

Sub ShowStr(s As String)
    Debug.Print s
End Sub

Sub ShowBytes(ab() As Byte)
    Dim i As Integer, iMin As Integer, iMax As Integer, s As String
    iMin = LBound(ab): iMax = UBound(ab)
    For i = iMin To iMax
        s = s & Chr$(ab(i))
    Next
    Debug.Print s
End Sub

Private Sub mnuCheck_Click()
    mnuCheck.Checked = Not mnuCheck.Checked
End Sub

Private Sub mnuDead_Click()
    MsgBox mnuDead.Caption
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuOpen_Click()
    MsgBox mnuOpen.Caption
End Sub

Private Sub mnuGone_Click()
    MsgBox mnuGone.Caption
End Sub

Private Sub mnuCut_Click()
    MsgBox mnuCut.Caption
End Sub

Private Sub mnuPaste_Click()
    MsgBox mnuPaste.Caption
End Sub

Private Sub mnuSome_Click()
    MsgBox mnuSome.Caption
End Sub

Private Sub mnuAll_Click()
    MsgBox mnuAll.Caption
End Sub

Private Sub mnuContents_Click()
    MsgBox mnuContents.Caption
End Sub

Private Sub mnuAbout_Click()
    MsgBox mnuAbout.Caption
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Function WordFromStrB(sBuf As String, iOffset As Long) As Integer
    BugAssert (iOffset + 2) <= LenB(sBuf) - 1
    Dim dw As Long
    dw = AscB(MidB$(sBuf, iOffset + 2, 1)) * 256&
    dw = dw + AscB(MidB$(sBuf, iOffset + 1, 1))
    If dw And &H8000& Then
        WordFromStrB = &H8000 Or (dw And &H7FFF&)
    Else
        WordFromStrB = dw And &HFFFF&
    End If
End Function
'
Private Sub pick_Picked(color As stdole.OLE_COLOR)
    BackColor = color
End Sub
