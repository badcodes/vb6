VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "BookManager"
   ClientHeight    =   5865
   ClientLeft      =   225
   ClientTop       =   825
   ClientWidth     =   11820
   LinkTopic       =   "Form1"
   ScaleHeight     =   5865
   ScaleWidth      =   11820
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame fraShow 
      BorderStyle     =   0  'None
      Caption         =   "BookInfo"
      Height          =   5385
      Left            =   5100
      TabIndex        =   9
      Top             =   120
      Width           =   6375
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   540
         Left            =   120
         TabIndex        =   10
         Top             =   4620
         Width           =   6090
      End
      Begin VB.Image imgCover 
         Height          =   4155
         Left            =   405
         Stretch         =   -1  'True
         Top             =   225
         Width           =   4110
      End
   End
   Begin BookManager.VSpliter VSplit 
      Height          =   3405
      Left            =   4980
      TabIndex        =   5
      Top             =   255
      Width           =   105
      _ExtentX        =   185
      _ExtentY        =   6006
   End
   Begin VB.Frame fraEditor 
      BorderStyle     =   0  'None
      Caption         =   "BookInfo"
      Height          =   5385
      Left            =   5205
      TabIndex        =   1
      Top             =   135
      Visible         =   0   'False
      Width           =   6375
      Begin VB.Frame fraCmd 
         BorderStyle     =   0  'None
         Height          =   540
         Left            =   120
         TabIndex        =   3
         Top             =   4620
         Width           =   6090
         Begin VB.CommandButton cmdA 
            Caption         =   "重命名"
            Height          =   360
            Index           =   3
            Left            =   765
            TabIndex        =   8
            Top             =   135
            Width           =   1110
         End
         Begin VB.CommandButton cmdA 
            Caption         =   "保存"
            Height          =   360
            Index           =   2
            Left            =   2145
            TabIndex        =   7
            Top             =   135
            Width           =   1110
         End
         Begin VB.CommandButton cmdA 
            Caption         =   "目录浏览"
            Height          =   360
            Index           =   1
            Left            =   3510
            TabIndex        =   6
            Top             =   120
            Width           =   1110
         End
         Begin VB.CommandButton cmdA 
            Caption         =   "Pdg阅读器"
            Height          =   360
            Index           =   0
            Left            =   4815
            TabIndex        =   4
            Top             =   120
            Width           =   1215
         End
      End
      Begin BookManager.KeyValueEditor KeyValueEditor 
         Height          =   3870
         Left            =   255
         TabIndex        =   2
         Top             =   405
         Width           =   5970
         _ExtentX        =   10530
         _ExtentY        =   6826
         Appearance      =   0
         BorderStyle     =   1
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin BookManager.FileBox FileBox 
      Height          =   3390
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   5980
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Drive           =   "c: [System]"
      Path            =   "c:\"
      FilePanel       =   0   'False
   End
   Begin VB.Menu mnuFile 
      Caption         =   "文件(&F)"
      Begin VB.Menu mnuOption 
         Caption         =   "设置(&P)"
      End
      Begin VB.Menu mnuNone 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnuExit 
         Caption         =   "退出(&Q)"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mMouseOnSpliter As Boolean
Private m_StrPdgProg As String
Private m_StrFolderProg As String
Private m_StrFormating As String
Private mConfigFile As String

Private Sub cmdA_Click(Index As Integer)
If Index = 0 Then
        Dim pdgDir As String
        Dim pdgFile As String
        Dim pdgOpener As String
        pdgOpener = m_StrPdgProg
        If pdgOpener = "" Then
            mnuOption_Click
            cmdA_Click (Index)
            Exit Sub
        End If
        pdgDir = BuildPath(KeyValueEditor.GetField("目录"))
        
        pdgFile = Dir$(pdgDir & "*.pdg")
        
        If pdgFile <> "" Then
            Shell pdgOpener & " " & QuoteString(pdgDir & pdgFile), vbNormalFocus
            
        Else
            MsgBox QuoteString(pdgDir) & "中不存在任何PDG文件", vbCritical + vbOKOnly
        End If
        Reset
    ElseIf Index = 1 Then
        Dim prog As String
        Dim arg As String
        arg = QuoteString(KeyValueEditor.GetField("目录"))
        prog = m_StrFolderProg
        If prog = "" Then prog = "explorer.exe"
        Shell prog & " " & arg, vbNormalFocus
    ElseIf Index = 2 Then
        SaveBook KeyValueEditor.GetField("目录")
    ElseIf Index = 3 Then
        RenameFolder KeyValueEditor.GetField("目录")
    End If
End Sub

Private Sub FileBox_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    mMouseOnSpliter = False
End Sub

Private Sub FileBox_PathChange(Path As String)
    If FolderExists(Path) Then LoadBook Path
End Sub

Private Sub FileBox_PathClick(Path As String)
    Dim bookPath As String
    bookPath = FileBox.Paths(FileBox.PathIndex)
    If FolderExists(bookPath) Then LoadBook bookPath
    'If FolderExists(FileBox.Paths(FileBox.PathIndex)) Then LoadBook Path
End Sub

Private Sub Form_Initialize()
SSLIB_Init
End Sub

Private Sub Form_Load()
On Error Resume Next
frmMain.Caption = App.ProductName & " V" & App.Major & "." & App.Minor & App.Revision

    KeyValueEditor.TwoColumnMode = False
 KeyValueEditor.AddItem "目录", VCT_Label, , False
Dim i As Long
    For i = CST_SSLIB_FIELDS_LBound To CST_SSLIB_FIELDS_IMPORTANT_UBOUND
        KeyValueEditor.AddItem SSLIB_ChnFieldName(i), VCT_NORMAL, , False     ', , bookInfo.Field(i), False
        'KeyValueEditor.FieldEnabled(SSLIB_ChnFieldName(i)) = False
    Next
   
    'KeyValueEditor.AddItem "文件目录"
    'KeyValueEditor.SetFieldStyle SSLIB_ChnFieldName(SSF_HEADER), VCT_MultiLine, False
    'KeyValueEditor.SetFieldStyle SSLIB_ChnFieldName(SSF_SAVEDIN), VCT_Combox + VCT_DIR, True


mConfigFile = App.Path & "\" & App.EXEName & ".ini"
Dim iniHnd As CLiNInI
Set iniHnd = New CLiNInI
iniHnd.Source = mConfigFile
            
            FormStateFromString Me, iniHnd.GetSetting("App", "WindowPosition")
            m_StrPdgProg = iniHnd.GetSetting("App", "PdgProg")
            m_StrFolderProg = iniHnd.GetSetting("App", "FolderProg")
            m_StrFormating = iniHnd.GetSetting("App", "FormatString")
            ControlStateFromString VSplit, iniHnd.GetSetting("App", "Spliter")
            FileBox.Path = iniHnd.GetSetting("App", "Path")
Set iniHnd = Nothing


End Sub

'Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    If X > VSplit.Left And _
'        X < VSplit.Left + VSplit.Width And _
'        Y > VSplit.Top And _
'        Y < VSplit.Top + VSplit.Height Then
'        mMouseOnSpliter = True
'        Me.MousePointer = 9
'    Else
'        mMouseOnSpliter = False
'        Me.MousePointer = 0
'    End If
'End Sub
'
'
'Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'
'If mMouseOnSpliter Then VSplit.Move X: Form_Resize: Exit Sub
'
'        If X > VSplit.Left And _
'        X < VSplit.Left + VSplit.Width And _
'        Y > VSplit.Top And _
'        Y < VSplit.Top + VSplit.Height Then
'            'Me.MousePointer = 9
'        Else
'            'Me.MousePointer = 0
'        End If
'
'
'End Sub
'
'Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    mMouseOnSpliter = False
'    Me.MousePointer = 0
'End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    Dim pH As Single
    Dim pW As Single
    pH = Me.ScaleHeight
    pW = Me.ScaleWidth
    
    VSplit.Top = 0
    VSplit.Height = pH
    
    FileBox.Move 0, 0, VSplit.Left, pH
    
    fraEditor.Move VSplit.Left + VSplit.Width, 0, pW - VSplit.Left - VSplit.Width, pH ' - 240
    
    fraCmd.Left = fraEditor.Width - fraCmd.Width
    fraCmd.Top = fraEditor.Height - fraCmd.Height '- 120 ', fraEditor.Width
    
    
    
    KeyValueEditor.Move 0, 0, fraEditor.Width, fraCmd.Top ' - 120  ' VSplit.Left + VSplit.Width, 0, Me.ScaleWidth - VSplit.Left - VSplit.Width, Me.ScaleHeight
    With fraEditor
    fraShow.Move .Left, .Top, .Width, .Height
    End With
End Sub
Private Sub SaveBook(Path As String)
    Dim infofile As String
    infofile = BuildPath(Path, "bookinfo.dat")
    Dim bookInfo As CBookInfo
    Set bookInfo = New CBookInfo
    'bookInfo.LoadFromFile infofile
    Dim i As Long
    For i = CST_SSLIB_FIELDS_LBound To CST_SSLIB_FIELDS_UBound
        bookInfo.Field(i) = KeyValueEditor.GetField(SSLIB_ChnFieldName(i))
    Next
    bookInfo.SaveToFile infofile
    Set bookInfo = Nothing
End Sub
Private Sub LoadBook(Path As String)
    Dim infofile As String
    Dim cover As String
    Dim bookInfo As CBookInfo
    infofile = BuildPath(Path, "bookinfo.dat")
    cover = BuildPath(Path, "cov001.pdg")
    If Not FileExists(cover) Then
        cover = BuildPath(Path, "cov001.jpg")
    End If
    If FileExists(cover) Then
        Set imgCover.Picture = LoadPicture(cover, , , imgCover.Width, imgCover.Height)
    End If
    If FileExists(infofile) = False Then
        Exit Sub
    End If
    Set bookInfo = New CBookInfo
    bookInfo.LoadFromFile infofile
    KeyValueEditor.Clear
    'If bookInfo.Field(SSF_Title) <> "" Then Exit Sub
    Dim i As Long
    For i = CST_SSLIB_FIELDS_LBound To CST_SSLIB_FIELDS_UBound
        KeyValueEditor.SetField SSLIB_ChnFieldName(i), bookInfo.Field(i)
    Next
    KeyValueEditor.SetField "目录", Path
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim iniHnd As CLiNInI
    Set iniHnd = New CLiNInI
    iniHnd.Source = mConfigFile
        iniHnd.SaveSetting "App", "WindowPosition", FormStateToString(Me)
        iniHnd.SaveSetting "App", "PdgProg", m_StrPdgProg
        iniHnd.SaveSetting "App", "FolderProg", m_StrFolderProg
        iniHnd.SaveSetting "App", "Spliter", ControlStateToString(VSplit)
        iniHnd.SaveSetting "App", "Path", FileBox.Path
        iniHnd.SaveSetting "App", "FormatString", m_StrFormating
    iniHnd.Save
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

'Private Sub VSplit_MouseDown()
'    Me.MousePointer = 9
'End Sub
'
''Private Sub VSplit_MouseMove()
''    Me.MousePointer = 9
''End Sub
'
'Private Sub VSplit_MouseUp()
'    Me.MousePointer = 0
'End Sub

Private Sub VSplit_Moving(x As Single, y As Single)
    VSplit.Move x
    Form_Resize
End Sub
Private Sub mnuOption_Click()
    Load frmOptions
    With frmOptions
        .PdgProg = m_StrPdgProg
        .FolderProg = m_StrFolderProg
        .FormatString = m_StrFormating
        .Show 1, Me
        m_StrPdgProg = .PdgProg
        m_StrFolderProg = .FolderProg
        m_StrFormating = .FormatString
    End With
    Unload frmOptions
End Sub

Private Sub RenameFolder(Oldname As String)


Dim srcdir As String
Dim dstdir As String

Dim bookInfo As CBookInfo
Dim namef As String

On Error Resume Next


'Set fso = New FileSystemObject
If FolderExists(Oldname) = False Then
    MsgBox "错误路径 : " & Oldname, vbCritical
    Exit Sub
End If
Set bookInfo = New CBookInfo
bookInfo(SSF_AUTHOR) = "test"
bookInfo(SSF_Title) = "test"
bookInfo(SSF_SSID) = "test"

namef = pdgformat(bookInfo, m_StrFormating)
If namef = "" Then
    MsgBox "错误的命名规则，必须至少含有%t,%a,%s,%p,%c,%d其中一项", vbCritical
    mnuOption_Click
    Exit Sub
End If
Set bookInfo = Nothing


    
    srcdir = BuildPath(Oldname)
    
    
    If FileExists(srcdir + "bookinfo.dat") Then
        
        Set bookInfo = New CBookInfo
    'ispdg = True
     'If fso.FileExists(workdir + "000001.pdg") Then ispdg = True Else ispdg = False
     'If Dir(workdir + "*.zip") = "" Then iszip = False Else iszip = True
    ' If ispdg Or iszip Then
        
        bookInfo.LoadFromFile srcdir + "bookinfo.dat"


        namef = cleanFilename(pdgformat(bookInfo, m_StrFormating))
        
        If namef <> "" Then
'            If iszip Then
'                srczip = Dir(workdir + "*.zip")
'                srczip = workdir + srczip
'                Dir (Environ("temp"))
'                dstzip = workdir + namef + ".zip"
'                If LCase(dstzip) <> LCase(srcdir) Then
'                    fso.MoveFile srczip, dstzip
'                End If
'            End If
            
                srcdir = Oldname
                If Right(srcdir, 1) = "\" Then srcdir = Left(srcdir, Len(srcdir) - 1)
                dstdir = BuildPath(GetParentFolderName(srcdir), namef)
                
                srcdir = Replace$(srcdir, "/", "\")
                dstdir = Replace$(dstdir, "/", "\")
    
                If LCase$(dstdir) <> LCase$(srcdir) Then
                    If FolderExists(dstdir) Then
                        Dim mR As VbMsgBoxResult
                        mR = MsgBox("文件夹" & Chr$(34) & dstdir & Chr$(34) & "已经存在，覆盖？", vbYesNo)
                        If mR = vbYes Then
                            DeleteFolder dstdir
                        Else
                            Exit Sub
                        End If
                    End If
                    Err.Clear
                    If MoveFile(srcdir, dstdir) Then
                        MsgBox "完成。" & srcdir & vbCrLf & "->" & vbCrLf & dstdir, vbInformation
                        UpdateFileBox dstdir
                    Else
                        MsgBox "失败，不能将" & srcdir & vbCrLf & "命名为" & vbCrLf & dstdir, vbCritical
                        Exit Sub
                    End If
                    
                    'fso.MoveFolder srcdir, dstdir
                End If
            End If
     End If
End Sub

Private Sub UpdateFileBox(Path As String)
     Dim pFolder As String
     pFolder = GetParentFolderName(Path)
     Dim fname As String
     fname = GetFileName(Path)
     
     On Error Resume Next
     FileBox.Path = "C:\"
     FileBox.Path = pFolder
     Dim i As Long
     For i = 0 To FileBox.PathCount
        If FileBox.Paths(i) = Path Then FileBox.PathIndex = i: Exit Sub
     Next
     


End Sub
