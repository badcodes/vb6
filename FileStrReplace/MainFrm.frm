VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form MainFrm 
   Caption         =   "File Str Replace"
   ClientHeight    =   5460
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8190
   Icon            =   "MainFrm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   8190
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkSDir 
      Caption         =   "Include Sub Folder  ( Just 1 Level Depth )"
      Height          =   375
      Left            =   3480
      TabIndex        =   8
      Top             =   4680
      Width           =   3495
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   5
      Top             =   5145
      Width           =   8190
      _ExtentX        =   14446
      _ExtentY        =   556
      ShowTips        =   0   'False
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   11289
            MinWidth        =   11289
            Key             =   "SBFile"
            Object.Tag             =   "SBFile"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Key             =   "SBCount"
            Object.Tag             =   "SBCount"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdExec 
      Caption         =   "Execute"
      Height          =   375
      Left            =   7080
      TabIndex        =   4
      Top             =   4680
      Width           =   975
   End
   Begin VB.CommandButton cmdOdir 
      Caption         =   "Path"
      Height          =   375
      Left            =   7080
      TabIndex        =   3
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox txtDir 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   6735
   End
   Begin VB.TextBox txtReplace 
      Height          =   1815
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   2760
      Width           =   7935
   End
   Begin VB.TextBox txtSearch 
      Height          =   1575
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   840
      Width           =   7935
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7560
      Top             =   3600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      MaxFileSize     =   32767
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Text to Replace:"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   2520
      Width           =   1185
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Text to Search for:"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   600
      Width           =   1320
   End
End
Attribute VB_Name = "MainFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExec_Click()
Dim fso As New FileSystemObject
Dim fsoFolders As Folders
Dim fsoFolder As Folder
Dim fsoFiles As Files
Dim fsoFile As File
Dim strDir As String
Dim iFileCount As Integer
Dim iFileNum As Integer
strDir = txtDir.Text
If fso.FolderExists(strDir) = False Then exitfor ("Folder:" + strDir + " doesn't exist."): Exit Sub
Set fsoFolder = fso.GetFolder(strDir)
Set fsoFiles = fsoFolder.Files
Set fsoFolders = fsoFolder.SubFolders
If txtSearch.Text = "" Then exitfor ("Text to search for is Null"): Exit Sub
If txtSearch.Text = txtReplace.Text Then exitfor ("Text to replace is the same as text to search for."): Exit Sub
iFileCount = fsoFiles.Count
If chkSDir.Value = 1 Then
    For Each fsoFolder In fsoFolders
        iFileCount = iFileCount + fsoFolder.Files.Count
    Next
End If
If iFileCount = 0 Then exitfor (strDir + " is empty."): Exit Sub
For Each fsoFile In fsoFiles
    iFileNum = iFileNum + 1
    StatusBar1.Panels("SBCount").Text = iFileNum + "/" + ifilecunt
    StatusBar1.Panels("SBFile").Text = "Searching " + fsoFile.Name + " ..."
    FileStrReplace fsoFile.Path, txtSearch.Text, txtReplace.Text
Next
If chkSDir.Value = 1 Then
    For Each fsoFolder In fsoFolders
        Set fsoFiles = fsoFolder.Files
        For Each fsoFile In fsoFiles
            iFileNum = iFileNum + 1
            StatusBar1.Panels("SBCount").Text = iFileNum + "/" + ifilecunt
            StatusBar1.Panels("SBFile").Text = "Searching " + fsoFile.Name + " ..."
            FileStrReplace fsoFile.Path, txtSearch.Text, txtReplace.Text
        Next
    Next
End If


End Sub
Sub exitfor(strReason As String)
MsgBox strReason, vbCritical, App.ProductName

End Sub
