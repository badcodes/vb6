VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "开心OL注册辅助 － 元宵快乐，陌大小姐"
   ClientHeight    =   6600
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10320
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   10320
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame 
      Height          =   5655
      Left            =   4440
      TabIndex        =   19
      Top             =   720
      Width           =   5655
      Begin SHDocVwCtl.WebBrowser WebBrowser 
         Height          =   5295
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   5415
         ExtentX         =   9551
         ExtentY         =   9340
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   ""
      End
   End
   Begin VB.CommandButton Command 
      Caption         =   "注册"
      Default         =   -1  'True
      Height          =   495
      Left            =   2400
      TabIndex        =   18
      Top             =   5880
      Width           =   1815
   End
   Begin VB.TextBox Text 
      Height          =   375
      Index           =   7
      Left            =   240
      TabIndex        =   17
      Text            =   "Text"
      Top             =   6000
      Width           =   1815
   End
   Begin VB.TextBox Text 
      Height          =   375
      Index           =   6
      Left            =   240
      TabIndex        =   14
      Text            =   "Text"
      Top             =   5280
      Width           =   3975
   End
   Begin VB.TextBox Text 
      Height          =   375
      Index           =   5
      Left            =   240
      TabIndex        =   12
      Text            =   "Text"
      Top             =   4560
      Width           =   3975
   End
   Begin VB.TextBox Text 
      Height          =   375
      Index           =   4
      Left            =   240
      TabIndex        =   11
      Text            =   "Text"
      Top             =   3840
      Width           =   3975
   End
   Begin VB.TextBox Text 
      Enabled         =   0   'False
      Height          =   375
      Index           =   3
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   9
      Text            =   "Text"
      Top             =   3120
      Width           =   3975
   End
   Begin VB.TextBox Text 
      Height          =   375
      Index           =   2
      Left            =   240
      TabIndex        =   7
      Text            =   "Text"
      Top             =   2400
      Width           =   3975
   End
   Begin VB.TextBox Text 
      Height          =   375
      Index           =   1
      Left            =   240
      TabIndex        =   5
      Text            =   "Text"
      Top             =   1680
      Width           =   3975
   End
   Begin VB.TextBox txtUrl 
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Text            =   "txtUrl"
      Top             =   120
      Width           =   9255
   End
   Begin VB.TextBox Text 
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Text            =   "Text"
      Top             =   960
      Width           =   3975
   End
   Begin VB.Label Label 
      Caption         =   "验证码："
      Height          =   375
      Index           =   8
      Left            =   240
      TabIndex        =   16
      Top             =   5760
      Width           =   3855
   End
   Begin VB.Label Label 
      Caption         =   "真实名："
      Height          =   375
      Index           =   6
      Left            =   240
      TabIndex        =   15
      Top             =   4320
      Width           =   3855
   End
   Begin VB.Label Label 
      Caption         =   "身份证："
      Height          =   375
      Index           =   7
      Left            =   240
      TabIndex        =   13
      Top             =   5040
      Width           =   3855
   End
   Begin VB.Label Label 
      Caption         =   "密码："
      Height          =   375
      Index           =   5
      Left            =   240
      TabIndex        =   10
      Top             =   3600
      Width           =   3855
   End
   Begin VB.Label Label 
      Caption         =   "用戶名："
      Height          =   375
      Index           =   4
      Left            =   240
      TabIndex        =   8
      Top             =   2880
      Width           =   3855
   End
   Begin VB.Label Label 
      Caption         =   "当前数值："
      Height          =   375
      Index           =   3
      Left            =   240
      TabIndex        =   6
      Top             =   2160
      Width           =   3855
   End
   Begin VB.Label Label 
      Caption         =   "数字长度："
      Height          =   375
      Index           =   2
      Left            =   240
      TabIndex        =   4
      Top             =   1440
      Width           =   3855
   End
   Begin VB.Label Label 
      Caption         =   "地址："
      Height          =   375
      Index           =   1
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label 
      Caption         =   "基础ID:"
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   3855
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Enum xrzRegData
    xrzRDBasicName = 0
    xrzRDNumLength = 1
    xrzRDNumCurrent = 2
    xrzRDID = 3
    xrzRDPassword = 4
    xrzRDRealName = 5
    xrzRDIDCard = 6
    xrzRDWords = 7
End Enum

Private Function SaveData(ByVal iKey As xrzRegData) As String
    SaveSetting App.EXEName, "Save", "Data" & iKey, Text(iKey).Text
End Function

Private Function LoadData(ByVal iKey As xrzRegData) As String
    Text(iKey).Text = GetSetting(App.EXEName, "Save", "Data" & iKey, "")
End Function


Private Sub Form_Load()
    Dim i As Integer
    For i = 0 To 7
        LoadData i
    Next
    txtUrl.Text = GetSetting(App.EXEName, "Save", "URL", "")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer
    For i = 0 To 7
        SaveData i
    Next
    SaveSetting App.EXEName, "Save", "URL", txtUrl.Text
End Sub

Private Sub Text_Change(Index As Integer)

End Sub
