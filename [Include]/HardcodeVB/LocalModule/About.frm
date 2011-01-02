VERSION 5.00
Begin VB.Form FAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About"
   ClientHeight    =   5328
   ClientLeft      =   3756
   ClientTop       =   3108
   ClientWidth     =   6480
   FillStyle       =   3  'Vertical Line
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   7.8
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5328
   ScaleWidth      =   6480
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox pbAnimate 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1308
      Left            =   30
      ScaleHeight     =   1308
      ScaleWidth      =   1176
      TabIndex        =   10
      Top             =   1395
      Width           =   1176
   End
   Begin VB.Frame fmUserInfo 
      BackColor       =   &H00C0C0C0&
      Height          =   975
      Left            =   1290
      TabIndex        =   5
      Top             =   2730
      Width           =   5055
      Begin VB.Label lblUserInfo 
         BackColor       =   &H00C0C0C0&
         Caption         =   "UserInfo"
         Height          =   210
         Index           =   2
         Left            =   150
         TabIndex        =   9
         Top             =   630
         Width           =   4800
      End
      Begin VB.Label lblUserInfo 
         BackColor       =   &H00C0C0C0&
         Caption         =   "UserInfo"
         Height          =   210
         Index           =   1
         Left            =   150
         TabIndex        =   8
         Top             =   420
         Width           =   4860
      End
      Begin VB.Label lblUserInfo 
         BackColor       =   &H00C0C0C0&
         Caption         =   "UserInfo"
         Height          =   210
         Index           =   0
         Left            =   150
         TabIndex        =   6
         Top             =   210
         Width           =   4830
      End
   End
   Begin VB.CommandButton cmdInfo 
      Caption         =   "&System Info..."
      Height          =   330
      Left            =   4575
      TabIndex        =   1
      Top             =   4365
      Width           =   1680
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4575
      TabIndex        =   0
      Top             =   3945
      Width           =   1695
   End
   Begin VB.Label lblMode 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Mode"
      Height          =   255
      Left            =   1350
      TabIndex        =   21
      Top             =   1335
      Width           =   3870
   End
   Begin VB.Label lbl 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Virtual Memory:"
      Height          =   225
      Index           =   1
      Left            =   1350
      TabIndex        =   20
      Top             =   1785
      Width           =   1815
   End
   Begin VB.Label lbl 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Memory Load:"
      Height          =   225
      Index           =   2
      Left            =   1350
      TabIndex        =   19
      Top             =   2010
      Width           =   1815
   End
   Begin VB.Label lbl 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Physical Memory:"
      Height          =   225
      Index           =   0
      Left            =   1350
      TabIndex        =   18
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label lbl 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Machine:"
      Height          =   225
      Index           =   3
      Left            =   1350
      TabIndex        =   17
      Top             =   2460
      Width           =   1815
   End
   Begin VB.Label lbl 
      BackColor       =   &H00C0C0C0&
      Caption         =   "User:"
      Height          =   225
      Index           =   4
      Left            =   1350
      TabIndex        =   16
      Top             =   2235
      Width           =   1815
   End
   Begin VB.Label lblVirtualMemory 
      BackColor       =   &H00C0C0C0&
      Caption         =   "xxx"
      Height          =   216
      Left            =   3240
      TabIndex        =   15
      Top             =   1785
      Width           =   2412
   End
   Begin VB.Label lblPhysicalMemory 
      BackColor       =   &H00C0C0C0&
      Caption         =   "xxx"
      Height          =   216
      Left            =   3240
      TabIndex        =   14
      Top             =   1560
      Width           =   2412
   End
   Begin VB.Label lblMemoryLoad 
      BackColor       =   &H00C0C0C0&
      Caption         =   "xxx"
      Height          =   216
      Left            =   3240
      TabIndex        =   13
      Top             =   2010
      Width           =   2412
   End
   Begin VB.Label lblUser 
      BackColor       =   &H00C0C0C0&
      Caption         =   "xxx"
      Height          =   216
      Left            =   3240
      TabIndex        =   12
      Top             =   2235
      Width           =   2412
   End
   Begin VB.Label lblMachine 
      BackColor       =   &H00C0C0C0&
      Caption         =   "xxx"
      Height          =   216
      Left            =   3240
      TabIndex        =   11
      Top             =   2460
      Width           =   2412
   End
   Begin VB.Label lblVersion 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Version"
      Height          =   255
      Left            =   135
      TabIndex        =   7
      Top             =   4995
      Width           =   5745
   End
   Begin VB.Line Line1 
      X1              =   165
      X2              =   6360
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Label lblComment 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Warning:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Left            =   150
      TabIndex        =   4
      Top             =   3960
      Width           =   4185
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblRights 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Copyright "
      Height          =   390
      Left            =   1350
      TabIndex        =   3
      Top             =   840
      Width           =   3975
   End
   Begin VB.Label lblApp 
      BackColor       =   &H00C0C0C0&
      Caption         =   "My Application"
      Height          =   240
      Left            =   1350
      TabIndex        =   2
      Top             =   240
      Width           =   3975
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   405
      Top             =   285
      Width           =   480
   End
End
Attribute VB_Name = "FAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'$ UTILITY.BAS

Public Client As App
Public ClientIcon As StdPicture
Public InfoProg As String
Public Copyright As String
Public Comments As String
Public SecretButton As Integer
Public SecretKey As Integer
Public SecretShift As Integer

Private fError As Boolean
Private asUserInfo(1 To 3) As String

Private anim As IAnimation

Sub Form_Load()
With Client
    BugMessage "Loading About"
    If anim Is Nothing Then Set anim = New CButterFly
    Set anim.Canvas = pbAnimate
    Me.Caption = "About " & .ProductName
    lblApp.Caption = .Title
    Dim sInfo As String
    On Error Resume Next
    sInfo = GetRegStr("SOFTWARE\Microsoft\Shared Tools\MSInfo", _
                      "Path", HKEY_LOCAL_MACHINE)
    ' Allow override because some customers might not have MSINFO
    If InfoProg = sEmpty Then InfoProg = sInfo
    If ExistFile(InfoProg) = False Then cmdInfo.Visible = False
        
    ' Icon from first form is application icon
    If Not ClientIcon Is Nothing Then
        If ClientIcon.Type = vbPicTypeIcon Then
            Set Me.Icon = ClientIcon
        End If
        Set imgIcon.Picture = ClientIcon
    End If
    lblMode.Caption = System.Mode & " on " & System.Processor
    lblPhysicalMemory.Caption = System.FreePhysicalMemory & _
        " KB of " & System.TotalPhysicalMemory & " KB"
    lblVirtualMemory.Caption = System.FreeVirtualMemory & _
        " KB of " & System.TotalVirtualMemory & " KB"
    lblMemoryLoad.Caption = System.MemoryLoad & "%"
    lblUser.Caption = System.User
    lblMachine.Caption = System.Machine
    If UserInfo(1) = sEmpty And UserInfo(2) = sEmpty And _
                                UserInfo(3) = sEmpty Then
        fmUserInfo.Visible = False
    Else
        fmUserInfo.Visible = True
        lblUserInfo(0).Caption = UserInfo(1)
        lblUserInfo(1).Caption = UserInfo(2)
        lblUserInfo(2).Caption = UserInfo(3)
    End If
    If Copyright = sEmpty Then Copyright = .LegalCopyright
    lblRights.Caption = Copyright
    If Comments = sEmpty Then Comments = .Comments
    lblComment.Caption = Comments
    lblVersion.Caption = "Version " & .Major & "." & .Minor
End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' Clean up UserInfo--it must be set before each load
    asUserInfo(1) = sEmpty
    asUserInfo(2) = sEmpty
    asUserInfo(3) = sEmpty
    Set anim = Nothing
End Sub

Private Sub Form_Initialize()
    BugMessage "Initializing About"
End Sub

Private Sub Form_Terminate()
    BugMessage "Terminating About"
End Sub

Private Sub Form_Activate()
    BugMessage "Activating About"
    ' Default for undocumented secret feature!
    If SecretButton = 0 Then SecretButton = vbRightButton
    If SecretKey = 0 Then SecretKey = vbKeySubtract
    If SecretShift = 0 Then SecretShift = vbShiftMask Or vbAltMask
End Sub

Private Sub cmdInfo_Click()
    On Error GoTo InfoFail
    Shell InfoProg, vbNormalFocus
    Exit Sub
InfoFail:
    MsgBox "Can't find information program"
End Sub

Private Sub cmdOK_Click()
    ' Make sure animation is off so About form can die
    anim.Running = False
    DoEvents
    Me.Hide
End Sub

Property Get Animator() As IAnimation
    Set Animator = anim
End Property

Property Set Animator(ia As IAnimation)
    Set anim = ia
End Property

Property Get UserInfo(i As Integer) As String
    Select Case i
    Case 1 To 3
        UserInfo = asUserInfo(i)
    Case Else
        UserInfo = sEmpty
    End Select
End Property

Property Let UserInfo(i As Integer, asUserInfoA As String)
    Select Case i
    Case 1 To 3
        asUserInfo(i) = asUserInfoA
    End Select
End Property

Property Get Error() As Boolean
    Error = fError
End Property

' Undocumented secret feature!
Private Sub pbAnimate_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = SecretButton Then
        anim.Running = Not anim.Running
    End If
End Sub

' Undocumented secret feature!
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = SecretKey Then
        If Shift And SecretShift Then
            anim.Running = Not anim.Running
        End If
    End If
End Sub

