VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Netcat"
   ClientHeight    =   7695
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11175
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7695
   ScaleWidth      =   11175
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtContent 
      Height          =   3630
      Left            =   15
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   5
      Text            =   "frmMain.frx":0000
      Top             =   3855
      Width           =   11040
   End
   Begin VB.TextBox txtRespone 
      Height          =   1500
      Left            =   -15
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   4
      Text            =   "frmMain.frx":0008
      Top             =   2160
      Width           =   11040
   End
   Begin VB.TextBox txtHeader 
      Height          =   1500
      Left            =   15
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Text            =   "frmMain.frx":0010
      Top             =   480
      Width           =   11040
   End
   Begin VB.CommandButton cmdGet 
      Caption         =   "Get"
      Height          =   360
      Left            =   9915
      TabIndex        =   2
      Top             =   30
      Width           =   1125
   End
   Begin VB.TextBox txtUrl 
      Height          =   345
      Left            =   345
      TabIndex        =   1
      Top             =   30
      Width           =   9375
   End
   Begin VB.Label lblUrl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "URL:"
      Height          =   195
      Left            =   30
      TabIndex        =   0
      Top             =   105
      Width           =   345
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents cNet As CNetConnection
Attribute cNet.VB_VarHelpID = -1

Private Sub cmdGet_Click()
On Error GoTo error_cmdGet
    Dim cHeader As CHttpHeader
    Set cHeader = New CHttpHeader
    cHeader.Init txtHeader.Text
    
    Set cNet = New CNetConnection
    cNet.URL = txtUrl.Text
    cNet.header = cHeader.HeaderString
    cNet.Connect
    
    txtRespone.Text = cNet.Respone.HeaderString
    
    txtContent.Text = ""
    cNet.Retrieve
    Exit Sub
error_cmdGet:
    MsgBox Err.Description, vbCritical
    Exit Sub
End Sub

Private Sub cNet_DataArrived(Data() As Byte, ByVal Size As Long)
    txtContent.Text = txtContent.Text & StrConv(Data, vbUnicode)
End Sub

