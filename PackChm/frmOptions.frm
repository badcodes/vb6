VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   4920
   ClientLeft      =   2565
   ClientTop       =   1500
   ClientWidth     =   6150
   Icon            =   "frmOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   6150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   3660
      TabIndex        =   16
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4890
      TabIndex        =   15
      Top             =   4440
      Width           =   1095
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   3
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample4 
         Caption         =   "Sample 4"
         Height          =   1785
         Left            =   2100
         TabIndex        =   7
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
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample3 
         Caption         =   "Sample 3"
         Height          =   1785
         Left            =   1545
         TabIndex        =   6
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
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample2 
         Caption         =   "Sample 2"
         Height          =   1785
         Left            =   645
         TabIndex        =   5
         Top             =   300
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   0
      Left            =   210
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame frmHolder 
         BorderStyle     =   0  'None
         Height          =   3615
         Left            =   60
         TabIndex        =   8
         Top             =   60
         Width           =   5535
         Begin VB.CommandButton cmdSelectPath 
            Caption         =   "Select..."
            Height          =   375
            Index           =   1
            Left            =   4620
            TabIndex        =   14
            Top             =   1200
            Width           =   915
         End
         Begin VB.TextBox txtPath 
            Height          =   315
            Index           =   1
            Left            =   0
            TabIndex        =   13
            Text            =   "Text1"
            Top             =   1260
            Width           =   4455
         End
         Begin VB.CommandButton cmdSelectPath 
            Caption         =   "Select..."
            Height          =   375
            Index           =   0
            Left            =   4620
            TabIndex        =   11
            Top             =   360
            Width           =   915
         End
         Begin VB.TextBox txtPath 
            Height          =   315
            Index           =   0
            Left            =   0
            TabIndex        =   10
            Text            =   "Text1"
            Top             =   420
            Width           =   4455
         End
         Begin VB.Label lblSelect 
            Caption         =   "HTML Help Workshop Path:"
            Height          =   315
            Index           =   1
            Left            =   0
            TabIndex        =   12
            Top             =   960
            Width           =   5475
         End
         Begin VB.Label lblSelect 
            Caption         =   "HTML Help Compiler Path:"
            Height          =   315
            Index           =   0
            Left            =   0
            TabIndex        =   9
            Top             =   120
            Width           =   5475
         End
      End
   End
   Begin MSComctlLib.TabStrip tbsOptions 
      Height          =   4245
      Left            =   105
      TabIndex        =   0
      Top             =   120
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   7488
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Tools"
            Key             =   "Tools"
            Object.ToolTipText     =   "Set Options for Tools"
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

Private Sub cmdApply_Click()
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    With frmMain
        .sHHCPath = txtPath(0).Text
        .sHHWPath = txtPath(1).Text
    End With
    Unload Me
End Sub

Private Sub cmdSelectPath_Click(Index As Integer)
Dim ret As String
Dim dlg As CCommonDialogLite
Set dlg = New CCommonDialogLite
ret = txtPath(Index).Text
If (dlg.VBGetOpenFileName(ret, , , , , , "Exe Files(*.exe)|*.exe", , , "Select " & lblSelect(Index))) Then
    txtPath(Index).Text = ret
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
    'center the form
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
    txtPath(0).Text = frmMain.sHHCPath
    txtPath(1).Text = frmMain.sHHWPath
    Me.Icon = frmMain.Icon
End Sub

Private Sub tbsOptions_Click()
    
    Dim i As Integer
    'show and enable the selected tab's controls
    'and hide and disable all others
    For i = 0 To tbsOptions.Tabs.Count - 1
        If i = tbsOptions.SelectedItem.Index - 1 Then
            picOptions(i).Left = 210
            picOptions(i).Enabled = True
        Else
            picOptions(i).Left = -20000
            picOptions(i).Enabled = False
        End If
    Next
    
End Sub
