VERSION 5.00
Begin VB.Form MainForm 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TimeCMD"
   ClientHeight    =   1512
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   5784
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1512
   ScaleWidth      =   5784
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtCMD 
      Appearance      =   0  'Flat
      Height          =   320
      Left            =   240
      TabIndex        =   4
      Top             =   1080
      Width           =   5295
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   320
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmdSetTime 
      Caption         =   "Set Time"
      Height          =   320
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   1335
   End
   Begin VB.Timer MainTimer 
      Interval        =   300
      Left            =   3240
      Top             =   120
   End
   Begin VB.Label lblRemainTimeHead 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "At"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   1800
      TabIndex        =   6
      Top             =   216
      Width           =   180
   End
   Begin VB.Label lblStartTimeHead 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "At"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   1800
      TabIndex        =   5
      Top             =   720
      Width           =   180
   End
   Begin VB.Label lblStartTime 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   2028
      TabIndex        =   2
      Top             =   720
      Width           =   96
   End
   Begin VB.Label lblRemainTime 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   2064
      TabIndex        =   0
      Top             =   216
      Width           =   96
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public timerstart As Boolean

Private Type localStrTab
    cpStartTime As String
    cpRemainTime As String
    cpStart As String
    cpStop As String
    cpSetTime As String
    cpHour As String
    cpMinute As String
    cpSecond As String
    cpInputCmd As String
    cpTimeFormat As String
End Type
Private stringTable As localStrTab
Private Sub loadStrTable()
    With stringTable
        .cpStartTime = LoadResString(101)
        .cpRemainTime = LoadResString(102)
        .cpStart = LoadResString(103)
        .cpStop = LoadResString(104)
        .cpSetTime = LoadResString(105)
        .cpHour = LoadResString(106)
        .cpMinute = LoadResString(107)
        .cpSecond = LoadResString(108)
        .cpInputCmd = LoadResString(109)
        .cpTimeFormat = LoadResString(110)
    End With
End Sub

Private Sub cmdSetTime_Click()
Dim sdtime As Date
Dim pos As Integer
On Error GoTo subEndOfCmdSetTime
sdtime = CDate(InputBox(stringTable.cpTimeFormat, App.ProductName, lblStartTime.Caption))
lblStartTime.Caption = CStr(sdtime)
lblRemainTime.Caption = ""
timerstart = False
cmdStart.Caption = stringTable.cpStart
subEndOfCmdSetTime:
End Sub


Private Sub cmdStart_Click()
If timerstart Then
    timerstart = False
    cmdStart.Caption = stringTable.cpStart  ' "Start"
Else
    timerstart = True
    cmdStart.Caption = stringTable.cpStop '  '"Stop"
End If
End Sub

Private Sub Form_Load()
Dim timecmd As String
Call loadStrTable
timecmd = VBA.Interaction.GetSetting(App.ProductName, "Setting", "TimeCMD")
If timecmd = "" Then timecmd = InputBox(stringTable.cpInputCmd, App.ProductName)
txtCMD.Text = timecmd
lblStartTime.Caption = VBA.Interaction.GetSetting(App.ProductName, "Setting", "TheTime")
lblStartTimeHead.Caption = stringTable.cpStartTime
lblStartTime.Left = lblStartTimeHead.Left + lblStartTimeHead.Width
lblRemainTimeHead.Caption = stringTable.cpRemainTime
lblRemainTime.Left = lblRemainTimeHead.Left + lblRemainTimeHead.Width
cmdSetTime.Caption = stringTable.cpSetTime
cmdStart.Caption = stringTable.cpStart

End Sub

Private Sub Form_Unload(Cancel As Integer)
VBA.Interaction.SaveSetting App.ProductName, "Setting", "TimeCMD", txtCMD.Text
VBA.Interaction.SaveSetting App.ProductName, "Setting", "TheTime", lblStartTime.Caption
End Sub



Private Sub MainTimer_Timer()
Dim theTime As Date
Dim dh As Integer
Dim dm As Integer
Dim ds As Integer
theTime = Time$
MainForm.Caption = App.ProductName + " - " + CStr(theTime)
sh = Hour(theTime)
sm = Minute(theTime)
ss = Second(theTime)
If lblStartTime.Caption <> "" And timerstart Then
sdtime = CDate(lblStartTime.Caption)
dh = Hour(sdtime)
dm = Minute(sdtime)
ds = Second(sdtime)
If dh = sh And dm = sm And ds >= ss Then
    MainTimer.Enabled = False
    Shell txtCMD.Text, vbNormalFocus
    Unload MainForm
    End
End If

ds = ds - ss
If ds < 0 Then ds = ds + 60: dm = dm - 1
dm = dm - sm
If dm < 0 Then dm = dm + 60: dh = dh - 1
dh = dh - sh
If dh < 0 Then dh = dh + 24
lblRemainTime.Caption = Strnum(dh, 2) & stringTable.cpHour & _
                                       Strnum(dm, 2) & stringTable.cpMinute & _
                                       Strnum(ds, 2) & stringTable.cpSecond
End If
MainForm.Refresh
End Sub
Function Strnum(num As Integer, numnum As Integer) As String
Strnum = LTrim(Str(num))
If Len(Strnum) >= numnum Then Exit Function
Strnum = String(numnum - Len(Strnum), 48) + Strnum
End Function
