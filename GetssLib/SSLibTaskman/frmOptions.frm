VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   3435
   ClientLeft      =   2565
   ClientTop       =   1500
   ClientWidth     =   7320
   Icon            =   "frmOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   7320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2070
      Left            =   195
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
      Caption         =   "Apply"
      Height          =   375
      Left            =   6135
      TabIndex        =   3
      Top             =   2910
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4935
      TabIndex        =   2
      Top             =   2910
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   3705
      TabIndex        =   1
      Top             =   2910
      Width           =   1095
   End
   Begin MSComctlLib.TabStrip tbsOptions 
      Height          =   2640
      Left            =   105
      TabIndex        =   0
      Top             =   150
      Width           =   7110
      _ExtentX        =   12541
      _ExtentY        =   4657
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Path"
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

Public Property Get FolderProg() As String
    FolderProg = m_StrFolderProg
End Property

Public Property Let FolderProg(ByVal StrValue As String)
    m_StrFolderProg = StrValue
    txtFolderProg.text = StrValue
End Property

Public Property Get PdgProg() As String
    PdgProg = m_StrPdgProg
End Property

Public Property Let PdgProg(ByVal StrValue As String)
    m_StrPdgProg = StrValue
    txtPDGProg.text = StrValue
End Property


Private Sub cmdApply_Click()
    m_StrPdgProg = txtPDGProg.text
    m_StrFolderProg = txtFolderProg.text
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
        
        If Index = 0 Then txtPDGProg.text = FileName
        If Index = 1 Then txtFolderProg.text = FileName
        
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


