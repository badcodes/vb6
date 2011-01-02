VERSION 5.00
Begin VB.Form FTestThread 
   Caption         =   "Test Thread"
   ClientHeight    =   3345
   ClientLeft      =   1395
   ClientTop       =   2070
   ClientWidth     =   4410
   Icon            =   "TThread.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3345
   ScaleWidth      =   4410
   Begin VB.CommandButton cmdStartStop 
      Caption         =   "&Start Thread"
      Height          =   495
      Left            =   324
      TabIndex        =   9
      Top             =   216
      Width           =   1548
   End
   Begin VB.TextBox txtStartStop 
      Height          =   375
      Left            =   2508
      TabIndex        =   7
      Text            =   "0"
      Top             =   480
      Width           =   1575
   End
   Begin VB.TextBox txtAPI 
      Height          =   375
      Left            =   2508
      TabIndex        =   5
      Top             =   1950
      Width           =   1575
   End
   Begin VB.TextBox txtBasic 
      Height          =   375
      Left            =   2508
      TabIndex        =   3
      Top             =   2685
      Width           =   1575
   End
   Begin VB.TextBox txtCalc 
      Height          =   375
      Left            =   2508
      TabIndex        =   1
      Top             =   1215
      Width           =   1575
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update Data"
      Height          =   495
      Left            =   324
      TabIndex        =   0
      Top             =   828
      Width           =   1548
   End
   Begin VB.Label lblStartStop 
      Caption         =   "Argument:"
      Height          =   240
      Left            =   2508
      TabIndex        =   8
      Top             =   240
      Width           =   1440
   End
   Begin VB.Label lbl 
      Caption         =   "API count:"
      Height          =   240
      Index           =   1
      Left            =   2505
      TabIndex        =   6
      Top             =   1710
      Width           =   1440
   End
   Begin VB.Label lbl 
      Caption         =   "Basic time:"
      Height          =   240
      Index           =   2
      Left            =   2505
      TabIndex        =   4
      Top             =   2445
      Width           =   1440
   End
   Begin VB.Label lbl 
      Caption         =   "Calculated count:"
      Height          =   240
      Index           =   0
      Left            =   2505
      TabIndex        =   2
      Top             =   975
      Width           =   1440
   End
End
Attribute VB_Name = "FTestThread"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdStartStop_Click()
    If cmdStartStop.Caption = "&Start Thread" Then
        StartThread txtStartStop
        cmdStartStop.Caption = "&Stop Thread"
        lblStartStop.Caption = "Return value:"
        cmdUpdate_Click
    Else
        txtStartStop = StopThread
        cmdStartStop.Caption = "&Start Thread"
        lblStartStop.Caption = "Argument:"
        cmdUpdate_Click
    End If
End Sub

Private Sub cmdUpdate_Click()
    txtCalc = CalcCount
    txtAPI = APICount
    txtBasic = BasicTime
End Sub

Private Sub Form_Load()
    Randomize 5
    Debug.Print "Caller: " & Rnd
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If ThreadRunning Then StopThread
End Sub
