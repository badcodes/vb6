VERSION 5.00
Begin VB.Form FBugWizard 
   Caption         =   "Debug Wizard"
   ClientHeight    =   4920
   ClientLeft      =   1416
   ClientTop       =   1320
   ClientWidth     =   4620
   Icon            =   "BugWiz.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4920
   ScaleWidth      =   4620
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   492
      Left            =   120
      TabIndex        =   10
      Top             =   4200
      Width           =   2052
   End
   Begin VB.CommandButton cmdBug 
      Caption         =   "Trim Assertions"
      Height          =   492
      Index           =   5
      Left            =   120
      TabIndex        =   9
      Top             =   3600
      Width           =   2052
   End
   Begin VB.FileListBox fileCur 
      Height          =   1800
      Left            =   2400
      Pattern         =   "*.bas;*.cls;*.frm;*.ctl"
      TabIndex        =   7
      Top             =   2760
      Width           =   2052
   End
   Begin VB.DirListBox dirCur 
      Height          =   1584
      Left            =   2400
      TabIndex        =   6
      Top             =   1080
      Width           =   2052
   End
   Begin VB.DriveListBox drvCur 
      Height          =   288
      Left            =   2400
      TabIndex        =   5
      Top             =   600
      Width           =   2052
   End
   Begin VB.CommandButton cmdBug 
      Caption         =   "Disable Profile Procedures"
      Height          =   492
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   2052
   End
   Begin VB.CommandButton cmdBug 
      Caption         =   "Enable Profile Procedures"
      Height          =   492
      Index           =   3
      Left            =   120
      TabIndex        =   3
      Top             =   2400
      Width           =   2052
   End
   Begin VB.CommandButton cmdBug 
      Caption         =   "Expand Assertions"
      Height          =   492
      Index           =   4
      Left            =   120
      TabIndex        =   2
      Top             =   3012
      Width           =   2052
   End
   Begin VB.CommandButton cmdBug 
      Caption         =   "Enable Bug Procedures"
      Height          =   492
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   2052
   End
   Begin VB.CommandButton cmdBug 
      Caption         =   "Disable Bug Procedures"
      Height          =   492
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   2052
   End
   Begin VB.Label lblFileCur 
      Height          =   492
      Left            =   240
      TabIndex        =   8
      Top             =   120
      Width           =   4092
   End
End
Attribute VB_Name = "FBugWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public sFileCur As String

Private Sub Form_Load()
    If fileCur.ListCount > 0 Then
        fileCur.ListIndex = 0
    End If
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdBug_Click(Index As Integer)
    HourGlass Me
    ' CBugFilter part of object
    Dim bug As CBugFilter
    Set bug = New CBugFilter
    ' IFilter part of object
    Dim filter As IFilter
    Set filter = bug
    ' Set FilterType property on bug variable
    bug.FilterType = Index
    ' Set Source property on filter variable
    filter.Source = sFileCur
    ' Pass either variable to FilterTextFile
#If fOddDayOfTheMonth Then
    FilterTextFile bug
#Else
    FilterTextFile filter
#End If
    HourGlass Me
End Sub

Private Sub dirCur_Change()
    fileCur.Path = dirCur.Path
    If fileCur.ListCount > 0 Then
        fileCur.ListIndex = 0
    End If
End Sub

Private Sub drvCur_Change()
    dirCur.Path = drvCur.Drive
End Sub

Private Sub fileCur_Click()
    lblFileCur.Caption = fileCur.FileName
    sFileCur = NormalizePath(fileCur.Path) & fileCur.FileName
    lblFileCur.Caption = sFileCur
End Sub

