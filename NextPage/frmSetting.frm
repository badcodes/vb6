VERSION 5.00
Begin VB.Form frmSetting 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "P"
   ClientHeight    =   2484
   ClientLeft      =   36
   ClientTop       =   300
   ClientWidth     =   5832
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2484
   ScaleWidth      =   5832
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtFilenameToDownload 
      Height          =   456
      Left            =   108
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   1464
      Width           =   5508
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "SaveSetting"
      Height          =   312
      Left            =   4128
      TabIndex        =   1
      Top             =   2064
      Width           =   1476
   End
   Begin VB.TextBox txtExtToDownload 
      Height          =   504
      Left            =   132
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   516
      Width           =   5508
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000012&
      Caption         =   "注：以 | 隔开各项"
      ForeColor       =   &H8000000F&
      Height          =   252
      Left            =   108
      TabIndex        =   5
      Top             =   2088
      Width           =   3636
   End
   Begin VB.Label Label2 
      Caption         =   "识别为下载文件的文件名（不带扩展名)："
      Height          =   300
      Left            =   120
      TabIndex        =   4
      Top             =   1104
      Width           =   5520
   End
   Begin VB.Label Label1 
      Caption         =   "识别为下载文件的扩展名："
      Height          =   300
      Left            =   144
      TabIndex        =   3
      Top             =   132
      Width           =   5436
   End
End
Attribute VB_Name = "frmSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sAppIni As String


Private Sub cmdSave_Click()

iniSaveSetting sAppIni, "Download", "Ext", txtExtToDownload.Text
iniSaveSetting sAppIni, "Download", "Filename", txtFilenameToDownload.Text

End Sub

Private Sub Form_Load()

sAppIni = frmProgress.iniFile
txtExtToDownload.Text = iniGetSetting(sAppIni, "Download", "Ext")
txtFilenameToDownload.Text = iniGetSetting(sAppIni, "Download", "Filename")

End Sub


