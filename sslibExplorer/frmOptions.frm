VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   3720
   ClientLeft      =   2565
   ClientTop       =   1500
   ClientWidth     =   7320
   Icon            =   "frmOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   7320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraOptions 
      Height          =   2370
      Index           =   0
      Left            =   210
      TabIndex        =   9
      Top             =   525
      Width           =   6870
      Begin VB.CommandButton cmdSelect 
         Caption         =   "选择..."
         Height          =   375
         Left            =   5550
         TabIndex        =   18
         Top             =   1815
         Width           =   1095
      End
      Begin VB.TextBox txtListFile 
         Height          =   345
         Left            =   135
         TabIndex        =   16
         Top             =   1845
         Width           =   5055
      End
      Begin VB.TextBox txtHomePage 
         Height          =   345
         Left            =   150
         TabIndex        =   12
         Text            =   "http://pds.sslibrary.com"
         Top             =   540
         Width           =   6495
      End
      Begin VB.TextBox txtUserName 
         Height          =   345
         Left            =   900
         TabIndex        =   11
         Top             =   1050
         Width           =   2235
      End
      Begin VB.TextBox txtPassword 
         Height          =   345
         Left            =   4395
         TabIndex        =   10
         Top             =   1050
         Width           =   2235
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "任务输出文件名："
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   1545
         Width           =   1440
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "主页："
         Height          =   195
         Left            =   135
         TabIndex        =   15
         Top             =   240
         Width           =   540
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "用戶名："
         Height          =   195
         Left            =   135
         TabIndex        =   14
         Top             =   1125
         Width           =   720
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "密码："
         Height          =   195
         Left            =   3630
         TabIndex        =   13
         Top             =   1125
         Width           =   540
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   3
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample4 
         Caption         =   "Sample 4"
         Height          =   1785
         Left            =   2100
         TabIndex        =   8
         Top             =   840
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   2
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample3 
         Caption         =   "Sample 3"
         Height          =   1785
         Left            =   1545
         TabIndex        =   7
         Top             =   675
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   1
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample2 
         Caption         =   "Sample 2"
         Height          =   1785
         Left            =   645
         TabIndex        =   6
         Top             =   300
         Width           =   2055
      End
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "应用"
      Height          =   375
      Left            =   6120
      TabIndex        =   2
      Top             =   3150
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消"
      Height          =   375
      Left            =   4935
      TabIndex        =   1
      Top             =   3150
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定"
      Height          =   375
      Left            =   3690
      TabIndex        =   0
      Top             =   3150
      Width           =   1095
   End
   Begin MSComctlLib.TabStrip tbsOptions 
      Height          =   2910
      Left            =   105
      TabIndex        =   19
      Top             =   150
      Width           =   7110
      _ExtentX        =   12541
      _ExtentY        =   5133
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "浏览器"
            Key             =   "explorer"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mHomePage As String
Private mUsername As String
Private mPassword As String
Private mTaskListFile As String

Public Property Get TaskListFile() As String
    TaskListFile = mTaskListFile
End Property

Public Property Let TaskListFile(ByVal StrValue As String)
    mTaskListFile = StrValue
    txtListFile.text = StrValue
End Property

Public Property Get Password() As String
    Password = mPassword
End Property

Public Property Let Password(ByVal StrValue As String)
    mPassword = StrValue
    txtPassword.text = StrValue
End Property

Public Property Get UserName() As String
    UserName = mUsername
End Property

Public Property Let UserName(ByVal StrValue As String)
    mUsername = StrValue
    txtUserName.text = StrValue
End Property

Public Property Get HomePage() As String
    HomePage = mHomePage
End Property

Public Property Let HomePage(ByVal StrValue As String)
    mHomePage = StrValue
    txtHomePage.text = StrValue
End Property



Private Sub cmdApply_Click()
    mHomePage = txtHomePage.text
    mUsername = txtUserName.text
    mPassword = txtPassword.text
    mTaskListFile = txtListFile.text
End Sub

Private Sub cmdCancel_Click()
    Me.Hide
    'Unload Me
End Sub

Private Sub cmdOK_Click()
    cmdApply_Click
    Me.Hide
    'MsgBox "Place code here to set options and close dialog!"
    'Unload Me
End Sub


Private Sub cmdSelect_Click()
    Dim FileName As String
    Dim dlg As CCommonDialogLite
    Set dlg = New CCommonDialogLite
    If dlg.VBGetSaveFileName(FileName) Then

        txtListFile.text = FileName

        'LoadConfigFrom FileName
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    'handle ctrl+tab to move to the next tab
    If Shift = vbCtrlMask And KeyCode = vbKeyTab Then
        i = tbsOptions.SelectedItem.Index
        If i = tbsOptions.Tabs.Count Then
            'last tab so we need to wrap to tab 1
            Set tbsOptions.SelectedItem = tbsOptions.Tabs(1)
        Else
            'increment the tab
            Set tbsOptions.SelectedItem = tbsOptions.Tabs(i + 1)
        End If
    End If
End Sub

Private Sub Form_Load()
    Me.Icon = frmMain.Icon     'center the form
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
End Sub


Private Sub tbsOptions_Click()
    Dim idx As Long
    idx = tbsOptions.SelectedItem.Index - 1
    Dim i As Long
    For i = 0 To fraOptions.UBound
        If i = idx Then fraOptions(i).Visible = True Else fraOptions(i).Visible = False
    Next
    
    'fraOptions(tbsOptions.SelectedItem.Index + 1).Visible = True
End Sub
