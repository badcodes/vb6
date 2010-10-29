VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SyncDirectory"
   ClientHeight    =   4455
   ClientLeft      =   30
   ClientTop       =   345
   ClientWidth     =   7350
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   7350
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstResult 
      Height          =   1230
      ItemData        =   "frmMain.frx":0442
      Left            =   120
      List            =   "frmMain.frx":0444
      TabIndex        =   8
      Top             =   2760
      Width           =   7095
   End
   Begin VB.CommandButton cmdSync 
      Caption         =   "Start"
      Height          =   375
      Left            =   5940
      TabIndex        =   7
      Top             =   2160
      Width           =   1275
   End
   Begin VB.CheckBox chkDelete 
      Caption         =   "Delete source?"
      Enabled         =   0   'False
      Height          =   435
      Left            =   120
      TabIndex        =   6
      Top             =   2100
      Width           =   5595
   End
   Begin VB.CommandButton cmdFilename 
      Caption         =   "Select..."
      Height          =   375
      Index           =   1
      Left            =   5940
      TabIndex        =   5
      Top             =   1425
      Width           =   1275
   End
   Begin VB.TextBox txtFilename 
      Height          =   375
      Index           =   1
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   1455
      Width           =   5640
   End
   Begin VB.CommandButton cmdFilename 
      Caption         =   "Select..."
      Height          =   375
      Index           =   0
      Left            =   5940
      TabIndex        =   2
      Top             =   480
      Width           =   1275
   End
   Begin VB.TextBox txtFilename 
      Height          =   375
      Index           =   0
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   510
      Width           =   5640
   End
   Begin VB.Label lblStatus 
      Caption         =   ":Ready"
      ForeColor       =   &H80000003&
      Height          =   270
      Left            =   120
      TabIndex        =   9
      Top             =   4080
      Width           =   7110
   End
   Begin VB.Label Label1 
      Caption         =   "Destionation Directory:"
      Height          =   330
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   7110
   End
   Begin VB.Label Label1 
      Caption         =   "Source Directory:"
      Height          =   330
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   135
      Width           =   7110
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_App As IApp
Private m_bUnloaded As Boolean

Private Sub cmdFilename_Click(Index As Integer)
Dim FileName As String
FileName = MDlgOpenDir.openDirDialog(Me.hWnd)
If (FileName <> "") Then
    txtFilename(Index).Text = FileName
End If
End Sub


Private Sub cmdSync_Click()
    
    Dim count As Long
    
    lstResult.Clear
    cmdSync.Enabled = False
    count = MFileSystem.SyncDirectory(txtFilename(0).Text, txtFilename(1).Text, chkDelete.Value, Me, "PushList", "Notify")
    cmdSync.Enabled = True
    lblStatus.Caption = ":OK " + CStr(count) + " files or directories copyed"
    
End Sub

Private Sub Form_Load()
    
    Set m_App = New CApp
    m_App.Initialize Me
    Me.Caption = m_App.Title
    Call m_App.OnLoad
    
    Dim argc As Integer
    Dim argv() As String
    argc = MShell.GetCommandLine(argv)
    If (argc > 0) Then txtFilename(0).Text = argv(0)
    If (argc > 1) Then txtFilename(1).Text = argv(1)
    txtFilename_Change 0
    txtFilename_Change 1
     
End Sub

Private Sub Form_Terminate()
    If Not m_bUnloaded Then
        m_bUnloaded = True
        m_App.OnUnload
        Set m_App = Nothing
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not m_bUnloaded Then
        m_bUnloaded = True
        m_App.OnUnload
        Set m_App = Nothing
    End If
End Sub

Public Sub PushList(ByRef rstr As String)
    lstResult.AddItem rstr
    lblStatus.Caption = ":Copying " & rstr & "..."
    frmMain.Refresh
    DoEvents

End Sub

Public Sub Notify(ByRef rstr As String)
    lblStatus.Caption = ":Checking " & rstr & "..."
    frmMain.Refresh
    DoEvents

End Sub

Private Sub txtFilename_Change(Index As Integer)
    If txtFilename(Index).Text <> "" And FolderExists(txtFilename(Index).Text) Then
        cmdSync.Enabled = True
    Else
        cmdSync.Enabled = False
    End If
End Sub
