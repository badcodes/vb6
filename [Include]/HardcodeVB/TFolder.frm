VERSION 5.00
Begin VB.Form FShellFolder 
   Caption         =   "Test Shell Folders"
   ClientHeight    =   5400
   ClientLeft      =   1500
   ClientTop       =   4116
   ClientWidth     =   8496
   Icon            =   "TFolder.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5400
   ScaleWidth      =   8496
   Begin VB.ComboBox cboWalk 
      Height          =   315
      ItemData        =   "TFolder.frx":0CFA
      Left            =   120
      List            =   "TFolder.frx":0D07
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   2604
      Width           =   1200
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "...  "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1992
      TabIndex        =   11
      Top             =   3468
      Width           =   330
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop!"
      Height          =   336
      Left            =   1416
      TabIndex        =   10
      Top             =   2616
      Width           =   960
   End
   Begin VB.CommandButton cmdWalkDir 
      Caption         =   "Walk One Directory"
      Height          =   372
      Left            =   120
      TabIndex        =   9
      Top             =   1614
      Width           =   2256
   End
   Begin VB.CommandButton cmdWalkFolder 
      Caption         =   "Walk One Folder"
      Height          =   372
      Left            =   120
      TabIndex        =   8
      Top             =   1104
      Width           =   2256
   End
   Begin VB.CommandButton cmdWalkDirs 
      Caption         =   "Walk All Directories"
      Height          =   372
      Left            =   120
      TabIndex        =   7
      Top             =   594
      Width           =   2256
   End
   Begin VB.CheckBox chkPath 
      Caption         =   "Use Path"
      Height          =   240
      Left            =   120
      TabIndex        =   6
      Top             =   3156
      Value           =   1  'Checked
      Width           =   1056
   End
   Begin VB.CommandButton cmdContext 
      Caption         =   "Context Menu"
      Height          =   372
      Left            =   108
      TabIndex        =   5
      Top             =   2124
      Width           =   2256
   End
   Begin VB.ListBox lstSpecial 
      Height          =   1008
      ItemData        =   "TFolder.frx":0D2B
      Left            =   120
      List            =   "TFolder.frx":0D2D
      TabIndex        =   3
      Top             =   4056
      Width           =   2268
   End
   Begin VB.TextBox txtPath 
      Height          =   312
      Left            =   120
      TabIndex        =   2
      Top             =   3456
      Width           =   1836
   End
   Begin VB.TextBox txtOut 
      Height          =   5184
      Left            =   2460
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   108
      Width           =   5904
   End
   Begin VB.CommandButton cmdWalkFolders 
      Caption         =   "Walk All Folders"
      Height          =   372
      Left            =   120
      TabIndex        =   0
      Top             =   84
      Width           =   2256
   End
   Begin VB.Label lbl 
      Caption         =   "Special folders:"
      Height          =   192
      Index           =   1
      Left            =   108
      TabIndex        =   4
      Top             =   3804
      Width           =   2268
   End
End
Attribute VB_Name = "FShellFolder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements IUseFolder
Implements IUseFile

Private fWalkAll As Long
Private fStop As Long

Private Sub Form_Load()
    With lstSpecial
        .AddItem "Desktop"
        .ItemData(.NewIndex) = &H0
        .AddItem "Programs"
        .ItemData(.NewIndex) = &H2
        .AddItem "Controls"
        .ItemData(.NewIndex) = &H3
        .AddItem "Printers"
        .ItemData(.NewIndex) = &H4
        .AddItem "Personal"
        .ItemData(.NewIndex) = &H5
        .AddItem "Favorites"
        .ItemData(.NewIndex) = &H6
        .AddItem "Startup"
        .ItemData(.NewIndex) = &H7
        .AddItem "Recent"
        .ItemData(.NewIndex) = &H8
        .AddItem "SendTo"
        .ItemData(.NewIndex) = &H9
        .AddItem "Bitbucket"
        .ItemData(.NewIndex) = &HA
        .AddItem "StartMenu"
        .ItemData(.NewIndex) = &HB
        .AddItem "DesktopDirectory"
        .ItemData(.NewIndex) = &H10
        .AddItem "Drives"
        .ItemData(.NewIndex) = &H11
        .AddItem "Network"
        .ItemData(.NewIndex) = &H12
        .AddItem "Nethood"
        .ItemData(.NewIndex) = &H13
        .AddItem "Fonts"
        .ItemData(.NewIndex) = &H14
        .AddItem "Templates"
        .ItemData(.NewIndex) = &H15
        .AddItem "Common StartMenu"
        .ItemData(.NewIndex) = &H16
        .AddItem "Common Programs"
        .ItemData(.NewIndex) = &H17
        .AddItem "Common Startup"
        .ItemData(.NewIndex) = &H18
        .AddItem "Common DestkopDirectory"
        .ItemData(.NewIndex) = &H19
        .AddItem "AppData"
        .ItemData(.NewIndex) = &H1A
        .AddItem "Printhood"
        .ItemData(.NewIndex) = &H1B
        .ListIndex = 0
    End With
    cboWalk.ListIndex = 2
    txtPath = GetTempDir
End Sub

Private Sub Form_Activate()
    chkPath_Click
End Sub

Private Sub chkPath_Click()
    If chkPath Then
        txtPath.Enabled = True
        txtPath.SetFocus
        lstSpecial.Enabled = False
        cmdWalkDirs.Enabled = True
        cmdWalkDir.Enabled = True
    Else
        txtPath.Enabled = False
        lstSpecial.Enabled = True
        cmdWalkDirs.Enabled = False
        cmdWalkDir.Enabled = False
    End If
End Sub

Private Sub cmdBrowse_Click()
    Dim s As String, af As Long
    af = BIF_RETURNONLYFSDIRS Or BIF_STATUSTEXT
    s = BrowseForFolder(hWnd, , af, "Select directory:", , txtPath)
    If s <> sEmpty Then txtPath = s
End Sub

Private Sub cmdContext_Click()
    If chkPath Then
        With txtPath
            ContextPopMenu hWnd, .Text, .Left, .Top
        End With
    Else
        With lstSpecial
            ContextPopMenu hWnd, .ItemData(.ListIndex), .Left, .Top
        End With
    End If
End Sub

Private Sub cmdStop_Click()
    fStop = True
End Sub

Private Sub cmdWalkDir_Click()
    txtOut = "Walk one directory: " & sCrLfCrLf
    fStop = False
    fWalkAll = False
    Dim c As Long
    WalkFiles Me, WalkType(cboWalk.ListIndex), txtPath, c
    txtOut = txtOut & vbCrLf & "File count: " & c & vbCrLf
    txtOut.SelStart = Len(txtOut)
End Sub

Private Sub cmdWalkDirs_Click()
    txtOut = "Walk directories recursively: " & sCrLfCrLf
    fWalkAll = True
    WalkAllFiles Me, WalkType(cboWalk.ListIndex), txtPath
End Sub

Private Sub cmdWalkFolder_Click()
    Dim folder As IVBShellFolder
    txtOut = "Walk one folder: " & sCrLfCrLf
    fWalkAll = False
    If chkPath Then
        Set folder = FolderFromItem(txtPath)
    Else
        With lstSpecial
            Set folder = FolderFromItem(.ItemData(.ListIndex))
        End With
    End If
    Dim c As Long
    WalkFolders folder, Me, c, WalkType(cboWalk.ListIndex)
    txtOut = txtOut & vbCrLf & "File count: " & c & vbCrLf
    txtOut.SelStart = Len(txtOut)
End Sub

Function WalkType(ByVal i As Integer) As EWalkMode
    Select Case i
    Case 0
        WalkType = ewmFolders
    Case 1
        WalkType = ewmNonfolders
    Case 2
        WalkType = ewmBoth
    End Select
End Function

Private Sub cmdWalkFolders_Click()
    Dim folder As IVBShellFolder
    txtOut = "Walk folders recursively: " & sCrLfCrLf
    fStop = False
    fWalkAll = True
    If chkPath Then
        Set folder = FolderFromItem(txtPath)
    Else
        With lstSpecial
            Set folder = FolderFromItem(.ItemData(.ListIndex))
        End With
    End If
    WalkAllFolders folder, Me, 0, WalkType(cboWalk.ListIndex)
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, _
                         X As Single, Y As Single)
    If Button And 2 Then
        With lstSpecial
            ContextPopMenu hWnd, .ItemData(.ListIndex), .Left, .Top
        End With
    End If
End Sub


Private Sub txtPath_GotFocus()
    txtPath.SelStart = 0
    txtPath.SelLength = 256
End Sub

Private Function IUseFolder_UseFolder(UserData As Variant, _
                                      CurFolder As IVBShellFolder, _
                                      ByVal ItemList As Long) As Boolean
' Turn folder and pidl into CFileInfo
With FileInfoFromFolder(CurFolder, ItemList)
    Dim s As String, sSize As String
    ' Don't show size for directories
    If (.Attributes And vbDirectory) = 0 Then
        sSize = " : " & Format$(.Length, "#,##0")
    End If
    ' Display different information for single or recursive walk
    If fWalkAll Then
        s = Space$(UserData * 4) & .DisplayName & " (" & _
            .TypeName & ") " & .Modified & sSize & vbCrLf
    Else
        UserData = UserData + 1
        s = .DisplayName & " (" & .TypeName & ") " & _
            .Modified & sSize & vbCrLf
    End If
    txtOut = txtOut & s
    txtOut.SelStart = Len(txtOut)
    ' Let other windows process so we can recognize stop flag
    DoEvents
    If fStop Then IUseFolder_UseFolder = True
End With
End Function

Private Function IUseFile_UseFile(UserData As Variant, _
                                   FilePath As String, _
                                   FileInfo As CFileInfo) As Boolean
With FileInfo
    Dim s As String, sSize As String
    If (.Attributes And vbDirectory) = 0 Then
        sSize = " : " & Format$(.Length, "#,##0")
    End If
    If fWalkAll Then
        s = Space$((UserData - 1) * 4) & .DisplayName & " (" & _
            .TypeName & ") " & .Modified & sSize & vbCrLf
    Else
        UserData = UserData + 1
        s = .DisplayName & " (" & .TypeName & ") " & _
            .Modified & sSize & vbCrLf
    End If
    txtOut = txtOut & s
    txtOut.SelStart = Len(txtOut)
    DoEvents
    If fStop Then IUseFile_UseFile = True
End With
End Function

