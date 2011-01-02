VERSION 5.00
Begin VB.Form FTestTimer 
   Caption         =   "Test Timer"
   ClientHeight    =   4464
   ClientLeft      =   2784
   ClientTop       =   7308
   ClientWidth     =   5448
   Icon            =   "TTimer.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4464
   ScaleWidth      =   5448
   Begin VB.CommandButton cmdDeleteTimer 
      Caption         =   "Delete Timer"
      Height          =   396
      Left            =   240
      TabIndex        =   5
      Top             =   3120
      Width           =   1200
   End
   Begin VB.ListBox lstTimers 
      Height          =   1392
      Left            =   240
      TabIndex        =   4
      Top             =   1560
      Width           =   1092
   End
   Begin VB.TextBox txtSec 
      Height          =   348
      Left            =   240
      TabIndex        =   2
      Text            =   "2"
      Top             =   480
      Width           =   1176
   End
   Begin VB.TextBox txtOut 
      Height          =   4092
      Left            =   1620
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   228
      Width           =   3648
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "New Timer"
      Height          =   396
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   1200
   End
   Begin VB.Label lbl 
      Caption         =   "Delay in seconds:"
      Height          =   204
      Left            =   168
      TabIndex        =   3
      Top             =   240
      Width           =   1320
   End
End
Attribute VB_Name = "FTestTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private cTimers As Long
Private aTimers(0 To 9) As CTimer

' Unfortunately VB doesn't support arrays of objects using WithEvents,
' so we have to put the objects in an array by hand
Private WithEvents wait0 As CTimer
Attribute wait0.VB_VarHelpID = -1
Private WithEvents wait1 As CTimer
Attribute wait1.VB_VarHelpID = -1
Private WithEvents wait2 As CTimer
Attribute wait2.VB_VarHelpID = -1
Private WithEvents wait3 As CTimer
Attribute wait3.VB_VarHelpID = -1
Private WithEvents wait4 As CTimer
Attribute wait4.VB_VarHelpID = -1
Private WithEvents wait5 As CTimer
Attribute wait5.VB_VarHelpID = -1
Private WithEvents wait6 As CTimer
Attribute wait6.VB_VarHelpID = -1
Private WithEvents wait7 As CTimer
Attribute wait7.VB_VarHelpID = -1
Private WithEvents wait8 As CTimer
Attribute wait8.VB_VarHelpID = -1
Private WithEvents wait9 As CTimer
Attribute wait9.VB_VarHelpID = -1

Private Sub cmdStart_Click()
    Select Case cTimers
    Case 0
        Set wait0 = New CTimer
        InitTimer wait0
    Case 1
        Set wait1 = New CTimer
        InitTimer wait1
    Case 2
        Set wait2 = New CTimer
        InitTimer wait2
    Case 3
        Set wait3 = New CTimer
        InitTimer wait3
    Case 4
        Set wait4 = New CTimer
        InitTimer wait4
    Case 5
        Set wait5 = New CTimer
        InitTimer wait5
    Case 6
        Set wait6 = New CTimer
        InitTimer wait6
    Case 7
        Set wait7 = New CTimer
        InitTimer wait7
    Case 8
        Set wait8 = New CTimer
        InitTimer wait8
    Case 9
        Set wait9 = New CTimer
        InitTimer wait9
    Case Else
        MsgBox "Too many timers"
    End Select
End Sub

Private Sub cmdDeleteTimer_Click()
    KillTimer
End Sub

Private Sub lstTimers_DblClick()
    KillTimer
End Sub

Sub InitTimer(wait As CTimer)
    ' Get the delay, but don't allow 0
    Dim i As Long
    i = Val(txtSec)
    If i <= 0 Then Exit Sub
    ' Set item (could be any associated data) and interval
    wait.Item = cTimers
    Set aTimers(cTimers) = wait
    cTimers = cTimers + 1
    wait.Interval = i * 1000
    ' Display in ListBox
    lstTimers.AddItem i & " seconds"
    lstTimers.ListIndex = cTimers - 1
    txtSec = i + 1
End Sub

Sub CallTimer(wait As CTimer)
    Dim s As String
    s = txtOut
    s = s & "Timer " & wait.Item & " has interval " & _
            wait.Interval / 1000 & " seconds" & vbCrLf
    txtOut = s
End Sub

Sub KillTimer()
With lstTimers
    Dim i As Long
    i = .ListIndex
    aTimers(i).Interval = 0
    Set aTimers(i) = Nothing
    .AddItem "Dead", i + 1
    .RemoveItem i
End With
End Sub

Private Sub wait0_ThatTime()
    CallTimer wait0
End Sub

Private Sub wait1_ThatTime()
    CallTimer wait1
End Sub

Private Sub wait2_ThatTime()
    CallTimer wait2
End Sub

Private Sub wait3_ThatTime()
    CallTimer wait3
End Sub

Private Sub wait4_ThatTime()
    CallTimer wait4
End Sub

Private Sub wait5_ThatTime()
    CallTimer wait5
End Sub

Private Sub wait6_ThatTime()
    CallTimer wait6
End Sub

Private Sub wait7_ThatTime()
    CallTimer wait7
End Sub

Private Sub wait8_ThatTime()
    CallTimer wait8
End Sub

Private Sub wait9_ThatTime()
    CallTimer wait9
End Sub



