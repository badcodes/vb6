VERSION 5.00
Begin VB.Form FVB6ToVB5 
   Caption         =   "VB6 To VB5"
   ClientHeight    =   5208
   ClientLeft      =   1416
   ClientTop       =   1320
   ClientWidth     =   4620
   Icon            =   "VB6ToVB5.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5208
   ScaleWidth      =   4620
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Height          =   396
      Left            =   1464
      TabIndex        =   8
      Top             =   3216
      Visible         =   0   'False
      Width           =   636
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   492
      Left            =   120
      TabIndex        =   7
      Top             =   2424
      Width           =   2052
   End
   Begin VB.FileListBox fileCur 
      Height          =   1608
      Left            =   2388
      Pattern         =   "*.frm;*.vbp;*.cls"
      TabIndex        =   5
      Top             =   2760
      Width           =   2076
   End
   Begin VB.DirListBox dirCur 
      Height          =   1584
      Left            =   2400
      TabIndex        =   4
      Top             =   1080
      Width           =   2052
   End
   Begin VB.DriveListBox drvCur 
      Height          =   288
      Left            =   2400
      TabIndex        =   3
      Top             =   600
      Width           =   2052
   End
   Begin VB.CommandButton cmdAll 
      Caption         =   "Translate Directory Tree"
      Height          =   492
      Left            =   120
      TabIndex        =   2
      Top             =   1788
      Width           =   2052
   End
   Begin VB.CommandButton cmdDir 
      Caption         =   "Translate Directory"
      Height          =   492
      Left            =   132
      TabIndex        =   1
      Top             =   1200
      Width           =   2052
   End
   Begin VB.CommandButton cmdFile 
      Caption         =   "Translate File"
      Height          =   492
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   2052
   End
   Begin VB.Label lblStatus 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   120
      TabIndex        =   9
      Top             =   4680
      Width           =   4335
   End
   Begin VB.Label lblFileCur 
      Height          =   492
      Left            =   240
      TabIndex        =   6
      Top             =   120
      Width           =   4092
   End
End
Attribute VB_Name = "FVB6ToVB5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private sFileCur As String
Private sDirCur As String

' CVBVerFilter part of object
Private WithEvents vbver As CVBVerFilter
Attribute vbver.VB_VarHelpID = -1

Private Sub Form_Load()
    dirCur.Path = CurDir$
    If fileCur.ListCount > 0 Then
        fileCur.ListIndex = 0
    End If
    sDirCur = NormalizePath(CurDir$)
    Set vbver = New CVBVerFilter
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdFile_Click()
    HourGlass Me
    vbver.ConvertFile sFileCur
    lblStatus.Caption = "Finished"
    HourGlass Me
End Sub

Private Sub cmdDir_Click()
    HourGlass Me
    cmdStop.Visible = True
    
    ' Process directory until done or user says stop
    WalkFiles vbver, ewmfFiles, sDirCur
    
    cmdStop.Visible = False
    lblStatus.Caption = "Finished"
    HourGlass Me
End Sub

Private Sub cmdAll_Click()
    HourGlass Me
    cmdStop.Visible = True
    
    ' Process all directories until done or user says stop
    WalkAllFiles vbver, ewmfFiles, sDirCur
    
    cmdStop.Visible = False
    lblStatus.Caption = "Finished"
    HourGlass Me
End Sub

Private Sub cmdStop_Click()
    vbver.StopWalk
    DoEvents
End Sub

Private Sub dirCur_Change()
    fileCur.Path = dirCur.Path
    If fileCur.ListCount > 0 Then
        fileCur.ListIndex = 0
    End If
    sDirCur = NormalizePath(dirCur.Path)
End Sub

Private Sub drvCur_Change()
    dirCur.Path = drvCur.Drive
    sDirCur = NormalizePath(dirCur.Path)
End Sub

Private Sub fileCur_Click()
    lblFileCur.Caption = fileCur.FileName
    sFileCur = NormalizePath(fileCur.Path) & fileCur.FileName
    lblFileCur.Caption = sFileCur
End Sub

Private Sub vbver_ConvertedFile(ByVal sFullPath As String)
    lblStatus.Caption = sFullPath
End Sub
