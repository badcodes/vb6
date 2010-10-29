VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form MainFrm 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "音乐压缩档解压辅助"
   ClientHeight    =   1260
   ClientLeft      =   120
   ClientTop       =   732
   ClientWidth     =   8736
   Icon            =   "mainfrm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1260
   ScaleWidth      =   8736
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cmDlg 
      Left            =   4140
      Top             =   330
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ComboBox cmbDir 
      Height          =   315
      Index           =   3
      Left            =   6645
      TabIndex        =   6
      Top             =   225
      Width           =   1905
   End
   Begin VB.ComboBox cmbDir 
      Height          =   315
      Index           =   2
      Left            =   4455
      TabIndex        =   5
      Top             =   210
      Width           =   1905
   End
   Begin VB.ComboBox cmbDir 
      Height          =   315
      Index           =   1
      Left            =   2295
      TabIndex        =   4
      Top             =   210
      Width           =   1905
   End
   Begin VB.TextBox txtFilename 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   1185
      Locked          =   -1  'True
      OLEDropMode     =   1  'Manual
      TabIndex        =   2
      Text            =   "将文件拖放于此..."
      Top             =   810
      Width           =   5760
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "解压执行"
      Default         =   -1  'True
      Height          =   375
      Left            =   7320
      TabIndex        =   1
      Top             =   705
      Width           =   1215
   End
   Begin VB.ComboBox cmbDir 
      Height          =   315
      Index           =   0
      Left            =   135
      TabIndex        =   0
      Top             =   195
      Width           =   1905
   End
   Begin VB.Shape ShpFilename 
      Height          =   390
      Left            =   1095
      Top             =   735
      Width           =   6000
   End
   Begin VB.Label lblFilename 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "文件名:"
      Height          =   195
      Left            =   150
      TabIndex        =   3
      Top             =   795
      Width           =   585
   End
   Begin VB.Menu mnuPath 
      Caption         =   "设置路径(&S)"
      Begin VB.Menu mnuPathMusic 
         Caption         =   "音乐目录"
      End
      Begin VB.Menu mnuPathWinRar 
         Caption         =   "Winrar路径"
      End
   End
End
Attribute VB_Name = "MainFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Cmdline As String
Dim MusicPath As String
Dim RarPath As String
Dim sIniFile As String



Private Sub cmbdir_Click(Index As Integer)

If Index = cmbDir.UBound Then Exit Sub

Dim i As Integer
Dim Realpath As String
Dim fso As New FileSystemObject
Dim fsoFolder As Folder
Dim curText As String

curText = cmbDir(Index + 1).Text

For i = Index + 1 To cmbDir.UBound
    cmbDir(i).Clear
Next

Realpath = MusicPath
For i = 0 To Index
    Realpath = fso.BuildPath(Realpath, cmbDir(i).Text)
Next
    
If fso.FolderExists(Realpath) = False Then Exit Sub
    
For Each fsoFolder In fso.GetFolder(Realpath).SubFolders
    cmbDir(Index + 1).AddItem fsoFolder.Name
Next

If curText <> "" Then cmbDir(Index + 1).Text = curText
End Sub


Private Sub cmbdir_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode <> Asc(vbCr) Then Exit Sub

If Index = cmbDir.UBound Then Exit Sub

Dim i As Integer
Dim Realpath As String
Dim fso As New FileSystemObject
Dim fsoFolder As Folder



For i = Index + 1 To cmbDir.UBound
    cmbDir(i).Clear
Next

Realpath = MusicPath
For i = 0 To Index
    Realpath = fso.BuildPath(Realpath, cmbDir(i).Text)
Next
    
If fso.FolderExists(Realpath) = False Then Exit Sub
    
For Each fsoFolder In fso.GetFolder(Realpath).SubFolders
    cmbDir(Index + 1).AddItem fsoFolder.Name
Next

End Sub




Private Sub cmdsave_Click()

If Cmdline = "" Then Exit Sub
Dim ExtractToPath As String
Dim shortFilename As String
Dim longFileName As String
Dim longExtractToPath As String
Dim shortRar As String
Dim fso As New FileSystemObject
Dim fsoFolder As Folder
Dim fsoSubfolder As Folder

ExtractToPath = MusicPath
For i = 0 To cmbDir.UBound
    If fso.FolderExists(ExtractToPath) = False Then
        fso.CreateFolder ExtractToPath
    End If
ExtractToPath = fso.BuildPath(ExtractToPath, cmbDir(i).Text)
Next
If fso.FolderExists(ExtractToPath) = False Then
    fso.CreateFolder ExtractToPath
End If
Debug.Print ExtractToPath

longExtractToPath = ExtractToPath
longFileName = fso.GetFile(Cmdline).Path

ExtractToPath = fso.GetFolder(ExtractToPath).ShortPath
shortFilename = fso.GetFile(Cmdline).ShortPath
shortRar = fso.GetFile(RarPath).ShortPath


Dim theCMD As String
theCMD = shortRar + " x -o+ -r " + shortFilename + " *.* " + ExtractToPath
ShellAndClose theCMD, vbNormalFocus
Set fsoFolder = fso.GetFolder(ExtractToPath)
If fsoFolder.SubFolders.Count = 1 And fsoFolder.Files.Count = 0 Then
    For Each fsoSubfolder In fsoFolder.SubFolders
    sfolder = fsoSubfolder.Path: Exit For
    Next
Set fsoFolder = Nothing
Set fsoSubfolder = Nothing
tempfolder = fso.BuildPath(fso.GetParentFolderName(ExtractToPath), fso.GetTempName)
If fso.FolderExists(sfolder) Then
    fso.MoveFolder sfolder, tempfolder
    fso.DeleteFolder ExtractToPath, True
    fso.MoveFolder tempfolder, longExtractToPath
End If

End If

fso.DeleteFile longFileName

End Sub

Private Sub Form_DblClick()
Dim fso As New FileSystemObject

MusicPath = InputBox("Type the path where you store musics.", App.ProductName, MusicPath)
If MusicPath = "" Or fso.FolderExists(MusicPath) = False Then Exit Sub
iniSaveSetting sIniFile, "Preference", "MusicPath", MusicPath
RarPath = InputBox("Type the path where the Rar App Located in.", App.ProductName, RarPath)
If RarPath = "" Or fso.FileExists(RarPath) = False Then Exit Sub
iniSaveSetting sIniFile, "Preference", "RarPath", RarPath

End Sub


Private Sub Form_Load()


Dim fso As New FileSystemObject

sIniFile = fso.BuildPath(App.Path, "config.ini")

MusicPath = iniGetSetting(sIniFile, "Preference", "MusicPath")
If MusicPath = "" Or fso.FolderExists(MusicPath) = False Then
MusicPath = InputBox("Type the path where you store musics.", App.ProductName, MusicPath)
End If
If MusicPath = "" Or fso.FolderExists(MusicPath) = False Then End
iniSaveSetting sIniFile, "Preference", "MusicPath", MusicPath

If Command <> "" Then MusicPath = Command

mnuPathMusic.Caption = mnuPathMusic.Caption & ":" & MusicPath

RarPath = iniGetSetting(sIniFile, "Preference", "RarPath")
If RarPath = "" Or fso.FileExists(RarPath) = False Then
RarPath = InputBox("Type the path where the Rar App Located in.", App.ProductName, RarPath)
End If
If RarPath = "" Or fso.FileExists(RarPath) = False Then End
iniSaveSetting sIniFile, "Preference", "RarPath", RarPath

mnuPathWinRar.Caption = mnuPathWinRar.Caption & ":" & RarPath



Dim fsoFolder As Folder

For Each fsoFolder In fso.GetFolder(MusicPath).SubFolders

cmbDir(0).AddItem fsoFolder.Name

Next
cmbDir(0).ListIndex = cmbDir(0).ListCount - 1


End Sub


Private Sub mnuPathMusic_Click()
Dim fso As New FileSystemObject

MusicPath = InputBox("Type the path where you store musics.", App.ProductName, MusicPath)
If MusicPath = "" Or fso.FolderExists(MusicPath) = False Then Exit Sub
iniSaveSetting sIniFile, "Preference", "MusicPath", MusicPath
mnuPathMusic.Caption = "音乐目录:" & MusicPath

End Sub

Private Sub mnuPathWinRar_Click()
Dim fso As New FileSystemObject

RarPath = InputBox("Type the path where the Rar App Located in.", App.ProductName, RarPath)
If RarPath = "" Or fso.FileExists(RarPath) = False Then Exit Sub
iniSaveSetting sIniFile, "Preference", "RarPath", RarPath

mnuPathWinRar.Caption = "Winrar路径:" & RarPath
End Sub

Private Sub txtFilename_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim fso As New FileSystemObject
Dim i As Integer
Dim sFilename As String
Dim sArtist As String
Dim sAlbum As String
Dim lEnd As Integer
'lEnd = cmbDir.Count - 1
'For i = 0 To lEnd
'cmbDir(0).Text = ""
'Next
Cmdline = Data.Files(1)
sFilename = fso.GetBaseName(Cmdline)
txtFilename.Text = sFilename
sFilename = txtFilename.Text
If mMusic.getInfo_VeryCD(sFilename, sArtist, sAlbum) Then
    cmbDir(1).Text = sArtist
    cmbDir(2).Text = sAlbum
Else
    sFilename = Replace(sFilename, ".", " ")
    cmbDir(1).Text = sFilename
'    lEnd = cmbDir(0).ListCount
'    For i = 1 To lEnd
'        If InStr(sFilenameCleared, cmbDir(0).List(i - 1)) > 0 Then
'            cmbDir(0).ListIndex = i - 1
'            Exit For
'        End If
'    Next i
End If
End Sub


Sub rebuilddir()
Dim fso As New FileSystemObject
Dim fsoFolder As Folder
Dim fsoSubfolder As Folder
Dim ExtractToPath As String
Dim sfolder As String
Dim tempfolder As String


For Each fsoFolder In fso.GetFolder("E:\Music\Billboard Top Hits").SubFolders

If fsoFolder.SubFolders.Count = 1 And fsoFolder.Files.Count = 0 Then
    ExtractToPath = fsoFolder.Path
    For Each fsoSubfolder In fsoFolder.SubFolders
    sfolder = fsoSubfolder.Path: Exit For
    Next
Set fsoFolder = Nothing
Set fsoSubfolder = Nothing
tempfolder = fso.BuildPath(fso.GetParentFolderName(ExtractToPath), fso.GetTempName)
If fso.FolderExists(sfolder) Then
    fso.MoveFolder sfolder, tempfolder
    fso.DeleteFolder ExtractToPath, True
    fso.MoveFolder tempfolder, ExtractToPath
End If

End If

Next
End Sub

