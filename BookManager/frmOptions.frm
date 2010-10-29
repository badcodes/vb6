VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   4200
   ClientLeft      =   2565
   ClientTop       =   1500
   ClientWidth     =   7215
   Icon            =   "frmOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   7215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmOption 
      Caption         =   "2、命名规则："
      Height          =   960
      Left            =   135
      TabIndex        =   16
      Top             =   2520
      Width           =   6855
      Begin VB.TextBox txtFormat 
         Height          =   288
         Left            =   168
         TabIndex        =   17
         Top             =   240
         Width           =   6495
      End
      Begin VB.Label lblFormat 
         Height          =   276
         Left            =   168
         TabIndex        =   18
         Top             =   636
         Width           =   5112
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "1、路径："
      Height          =   2220
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   6870
      Begin VB.CommandButton cmdSelect 
         Caption         =   "选择..."
         Height          =   375
         Index           =   1
         Left            =   5655
         TabIndex        =   15
         Top             =   1470
         Width           =   1095
      End
      Begin VB.TextBox txtFolderProg 
         Height          =   375
         Left            =   150
         TabIndex        =   14
         Top             =   1485
         Width           =   5280
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "选择..."
         Height          =   375
         Index           =   0
         Left            =   5640
         TabIndex        =   12
         Top             =   630
         Width           =   1095
      End
      Begin VB.TextBox txtPDGProg 
         Height          =   375
         Left            =   135
         TabIndex        =   11
         Top             =   645
         Width           =   5280
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "文件夹浏览程序："
         Height          =   195
         Left            =   150
         TabIndex        =   13
         Top             =   1185
         Width           =   1440
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "PDG阅读程序："
         Height          =   195
         Left            =   150
         TabIndex        =   10
         Top             =   315
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
      Caption         =   "Apply"
      Height          =   375
      Left            =   5880
      TabIndex        =   2
      Top             =   3660
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   3660
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   3450
      TabIndex        =   0
      Top             =   3660
      Width           =   1095
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
Private mFormatString As String

Public Property Get FormatString() As String
    FormatString = mFormatString
End Property

Public Property Let FormatString(ByVal StrValue As String)
    mFormatString = StrValue
    txtFormat.Text = StrValue
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


Private Sub cmdApply_Click()
    m_StrPdgProg = txtPDGProg.Text
    m_StrFolderProg = txtFolderProg.Text
    mFormatString = txtFormat.Text
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
    Dim FileName As String
    Dim dlg As CCommonDialogLite
    Set dlg = New CCommonDialogLite
    If dlg.VBGetOpenFileName(FileName, , , , , , "程序文件(*.exe)|*.exe|所有文件(*.*)|*.*", , , , "exe", Me.hwnd) Then
        
        If Index = 0 Then txtPDGProg.Text = FileName
        If Index = 1 Then txtFolderProg.Text = FileName
        
        'LoadConfigFrom FileName
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'    Dim i As Integer
'    'handle ctrl+tab to move to the next tab
'    If Shift = vbCtrlMask And KeyCode = vbKeyTab Then
'        i = tbsOptions.SelectedItem.Index
'        If i = tbsOptions.Tabs.Count Then
'            'last tab so we need to wrap to tab 1
'            Set tbsOptions.SelectedItem = tbsOptions.Tabs(1)
'        Else
'            'increment the tab
'            Set tbsOptions.SelectedItem = tbsOptions.Tabs(i + 1)
'        End If
'    End If
End Sub

Private Sub Form_Load()
    Me.Icon = frmMain.Icon     'center the form
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
    lblFormat.Caption = "%t = 书名, %a = 作者, %p = 页数, %d = 日期, %s = ss号, %c = 出版社"
End Sub


Private Sub tbsOptions_Click()

End Sub
