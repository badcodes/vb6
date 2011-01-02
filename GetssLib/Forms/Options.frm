VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   3945
   ClientLeft      =   2565
   ClientTop       =   1500
   ClientWidth     =   7320
   Icon            =   "Options.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   7320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraOptions 
      Height          =   2550
      Index           =   2
      Left            =   195
      TabIndex        =   22
      Top             =   540
      Visible         =   0   'False
      Width           =   6870
      Begin VB.CheckBox chkRenameJPG 
         Caption         =   "将 .jpg 改为 .pdg"
         Height          =   375
         Left            =   2340
         TabIndex        =   33
         Top             =   2040
         Width           =   2490
      End
      Begin VB.TextBox txtJPGBookQuality 
         Height          =   345
         Left            =   1170
         TabIndex        =   31
         Text            =   "1"
         Top             =   2055
         Width           =   765
      End
      Begin VB.TextBox txtPassword 
         Height          =   345
         Left            =   1170
         TabIndex        =   29
         Top             =   1560
         Width           =   5505
      End
      Begin VB.TextBox txtUserName 
         Height          =   345
         Left            =   1170
         TabIndex        =   27
         Top             =   1095
         Width           =   5505
      End
      Begin VB.TextBox txtPostUrl 
         Height          =   345
         Left            =   1185
         TabIndex        =   25
         Top             =   615
         Width           =   5505
      End
      Begin VB.TextBox txtLoginUrl 
         Height          =   345
         Left            =   1185
         TabIndex        =   23
         Text            =   "http://pds.sslibrary.com"
         Top             =   180
         Width           =   5505
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "大图质量："
         Height          =   195
         Left            =   120
         TabIndex        =   32
         Top             =   2115
         Width           =   960
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "密码："
         Height          =   195
         Left            =   120
         TabIndex        =   30
         Top             =   1620
         Width           =   960
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "用户名："
         Height          =   195
         Left            =   120
         TabIndex        =   28
         Top             =   1155
         Width           =   960
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "POSTURL："
         Height          =   195
         Left            =   120
         TabIndex        =   26
         Top             =   675
         Width           =   960
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "登录URL："
         Height          =   195
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   960
      End
   End
   Begin VB.Frame fraOptions 
      Height          =   2070
      Index           =   1
      Left            =   210
      TabIndex        =   17
      Top             =   510
      Visible         =   0   'False
      Width           =   6870
      Begin MSComctlLib.Slider sldTasks 
         Height          =   525
         Left            =   2880
         TabIndex        =   20
         Top             =   435
         Width           =   3150
         _ExtentX        =   5556
         _ExtentY        =   926
         _Version        =   393216
         LargeChange     =   2
         Min             =   1
         SelStart        =   1
         Value           =   1
      End
      Begin MSComctlLib.Slider sldThreads 
         Height          =   525
         Left            =   2880
         TabIndex        =   21
         Top             =   1125
         Width           =   3150
         _ExtentX        =   5556
         _ExtentY        =   926
         _Version        =   393216
         LargeChange     =   2
         Min             =   1
         SelStart        =   1
         Value           =   1
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "同时处理任务数："
         Height          =   195
         Left            =   615
         TabIndex        =   19
         Top             =   630
         Width           =   1440
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "单个任务下载进程数："
         Height          =   195
         Left            =   615
         TabIndex        =   18
         Top             =   1275
         Width           =   1800
      End
   End
   Begin VB.Frame fraOptions 
      Height          =   2070
      Index           =   0
      Left            =   210
      TabIndex        =   10
      Top             =   510
      Width           =   6870
      Begin VB.CommandButton cmdSelect 
         Caption         =   "选择..."
         Height          =   375
         Index           =   1
         Left            =   5655
         TabIndex        =   16
         Top             =   1380
         Width           =   1095
      End
      Begin VB.TextBox txtFolderProg 
         Height          =   375
         Left            =   150
         TabIndex        =   15
         Top             =   1395
         Width           =   5280
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "选择..."
         Height          =   375
         Index           =   0
         Left            =   5640
         TabIndex        =   13
         Top             =   540
         Width           =   1095
      End
      Begin VB.TextBox txtPDGProg 
         Height          =   375
         Left            =   135
         TabIndex        =   12
         Top             =   555
         Width           =   5280
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "文件夹浏览程序："
         Height          =   195
         Left            =   150
         TabIndex        =   14
         Top             =   1095
         Width           =   1440
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "PDG阅读程序："
         Height          =   195
         Left            =   150
         TabIndex        =   11
         Top             =   240
         Width           =   1245
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   3
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample4 
         Caption         =   "Sample 4"
         Height          =   1785
         Left            =   2100
         TabIndex        =   9
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
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample3 
         Caption         =   "Sample 3"
         Height          =   1785
         Left            =   1545
         TabIndex        =   8
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
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample2 
         Caption         =   "Sample 2"
         Height          =   1785
         Left            =   645
         TabIndex        =   7
         Top             =   300
         Width           =   2055
      End
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "应用"
      Height          =   375
      Left            =   6105
      TabIndex        =   3
      Top             =   3390
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消"
      Height          =   375
      Left            =   4920
      TabIndex        =   2
      Top             =   3390
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定"
      Height          =   375
      Left            =   3675
      TabIndex        =   1
      Top             =   3390
      Width           =   1095
   End
   Begin MSComctlLib.TabStrip tbsOptions 
      Height          =   3075
      Left            =   105
      TabIndex        =   0
      Top             =   150
      Width           =   7110
      _ExtentX        =   12541
      _ExtentY        =   5424
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "路径"
            Key             =   "path"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "下载"
            Key             =   "download"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "JPG大图"
            Key             =   "jpgbook"
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
Private m_StrPdgProg As String
Private m_StrFolderProg As String
Private mTasksProcessing As Long
Private mThreadsDownloading As Long
'Private mJpgBookStartURL As String
'Private mJpgBookCookie As String
Private mLoginUrl As String
Private mPostUrl As String
Private mUserName As String
Private mPassword As String

Public Property Get Password() As String
    Password = mPassword
End Property

Public Property Let Password(ByVal StrValue As String)
    mPassword = StrValue
    txtPassword.Text = StrValue
End Property

Public Property Get UserName() As String
    UserName = mUserName
End Property

Public Property Let UserName(ByVal StrValue As String)
    mUserName = StrValue
    txtUserName.Text = StrValue
End Property

Public Property Get PostUrl() As String
    PostUrl = mPostUrl
End Property

Public Property Let PostUrl(ByVal StrValue As String)
    mPostUrl = StrValue
    txtPostUrl.Text = StrValue
End Property

Public Property Get LoginUrl() As String
    LoginUrl = mLoginUrl
End Property

Public Property Let LoginUrl(ByVal StrValue As String)
    mLoginUrl = StrValue
    txtLoginUrl.Text = StrValue
End Property

'Public Property Get JpgBookCookie() As String
'    JpgBookCookie = mJpgBookCookie
'End Property
'
'Public Property Let JpgBookCookie(ByVal StrValue As String)
'    mJpgBookCookie = StrValue
'    txtJPGBookCookie.text = StrValue
'End Property



'Public Property Get JpgBookStartURL() As String
'    JpgBookStartURL = mJpgBookStartURL
'End Property
'
'Public Property Let JpgBookStartURL(ByVal StrValue As String)
'    mJpgBookStartURL = StrValue
'    txtLoginUrl.text = StrValue
'End Property

Public Property Get ThreadsDownloading() As Long
    ThreadsDownloading = mThreadsDownloading
End Property

Public Property Let ThreadsDownloading(ByVal LonValue As Long)
    mThreadsDownloading = LonValue
    On Error Resume Next
    sldThreads.Value = LonValue
End Property

Public Property Get TasksProcessing() As Long
    TasksProcessing = mTasksProcessing
End Property

Public Property Let TasksProcessing(ByVal LonValue As Long)
    mTasksProcessing = LonValue
    On Error Resume Next
    sldTasks.Value = LonValue
End Property

Public Property Get FolderProg() As String
    FolderProg = m_StrFolderProg
End Property

Public Property Let FolderProg(ByVal StrValue As String)
    m_StrFolderProg = StrValue
    txtFolderProg.Text = StrValue
End Property

Public Property Get PdgProg() As String
    PdgProg = m_StrPdgProg
End Property

Public Property Let PdgProg(ByVal StrValue As String)
    m_StrPdgProg = StrValue
    txtPDGProg.Text = StrValue
End Property


Private Sub Form_Load()
    With frmMain
        Me.Icon = frmMain.Icon     'center the form
        txtJPGBookQuality.Text = .JPGBookQuality
        chkRenameJPG.Value = IIf(.RenameJPGToPdg, 1, 0)
    End With
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
End Sub

Private Sub cmdApply_Click()
    m_StrPdgProg = txtPDGProg.Text
    m_StrFolderProg = txtFolderProg.Text
    mTasksProcessing = sldTasks.Value
    mThreadsDownloading = sldThreads.Value
    mPassword = txtPassword.Text
    mUserName = txtUserName.Text
    mPostUrl = txtPostUrl.Text
    mLoginUrl = txtLoginUrl.Text
    
    With frmMain
        .JPGBookQuality = StringToInteger(txtJPGBookQuality.Text)
        .RenameJPGToPdg = IIf(chkRenameJPG.Value = 1, True, False)
    End With
'    mJpgBookStartURL = txtLoginUrl.text
'    mJpgBookCookie = txtJPGBookCookie.text
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


Private Sub cmdSelect_Click(Index As Integer)
    Dim filename As String
    Dim dlg As CCommonDialogLite
    Set dlg = New CCommonDialogLite
    If dlg.VBGetOpenFileName(filename, , , , , , "程序文件(*.exe)|*.exe|所有文件(*.*)|*.*", , , , "exe", Me.hwnd) Then
        
        If Index = 0 Then txtPDGProg.Text = filename
        If Index = 1 Then txtFolderProg.Text = filename
        
        'LoadConfigFrom FileName
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    'handle ctrl+tab to move to the next tab
    If Shift = vbCtrlMask And KeyCode = vbKeyTab Then
        i = tbsOptions.SelectedItem.Index
        If i = tbsOptions.Tabs.count Then
            'last tab so we need to wrap to tab 1
            Set tbsOptions.SelectedItem = tbsOptions.Tabs(1)
        Else
            'increment the tab
            Set tbsOptions.SelectedItem = tbsOptions.Tabs(i + 1)
        End If
    End If
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
