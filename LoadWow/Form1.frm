VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "LoadWow"
   ClientHeight    =   1392
   ClientLeft      =   36
   ClientTop       =   360
   ClientWidth     =   2724
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1392
   ScaleWidth      =   2724
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOpenDir 
      Caption         =   "打开魔兽目录"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   984
      TabIndex        =   6
      Top             =   984
      Width           =   1620
   End
   Begin VB.CommandButton cmdZone 
      Caption         =   "五区"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   5
      Left            =   1020
      TabIndex        =   5
      Top             =   564
      Width           =   696
   End
   Begin VB.CommandButton cmdZone 
      Caption         =   "六区"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   6
      Left            =   1896
      TabIndex        =   4
      Top             =   564
      Width           =   696
   End
   Begin VB.CommandButton cmdZone 
      Caption         =   "四区"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   4
      Left            =   120
      TabIndex        =   3
      Top             =   576
      Width           =   696
   End
   Begin VB.CommandButton cmdZone 
      Caption         =   "三区"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   3
      Left            =   1884
      TabIndex        =   2
      Top             =   84
      Width           =   696
   End
   Begin VB.CommandButton cmdZone 
      Caption         =   "二区"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   2
      Left            =   1008
      TabIndex        =   1
      Top             =   96
      Width           =   696
   End
   Begin VB.CommandButton cmdZone 
      Caption         =   "一区"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   1
      Left            =   144
      TabIndex        =   0
      Top             =   96
      Width           =   696
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const ZONE1 = "cn1.grunt.wowchina.com"
Const ZONE2 = "cn2.grunt.wowchina.com"
Const ZONE3 = "cn3.grunt.wowchina.com"
Const ZONE4 = "cn4.grunt.wowchina.com"
Const ZONE5 = "cn5.grunt.wowchina.com"
Const ZONE6 = "cn6.grunt.wowchina.com"
Const CONFIG = "realmlist.wtf"
Const WOWBIN = "wow.exe"

Sub LoadZone(iArg As Integer)
Dim sZone As String
Dim iFile As Integer

Select Case iArg
Case 1
    sZone = ZONE1
Case 2
    sZone = ZONE2
Case 3
    sZone = ZONE3
Case 4
    sZone = ZONE4
Case 5
    sZone = ZONE5
Case 6
    sZone = ZONE6
End Select
iFile = FreeFile
Open CONFIG For Output As #iFile
Print #iFile, "SET realmlist " & Chr$(34) & sZone & Chr$(34)
Close #iFile
Shell WOWBIN, vbNormalFocus
End Sub

Private Sub cmdOpenDir_Click()
Shell "explorer.exe " & CurDir$, vbNormalFocus
End Sub

Private Sub cmdZone_Click(Index As Integer)
LoadZone Index
Unload Me
End
End Sub

Private Sub Form_Load()
If FileExists(WOWBIN) = False Then
    MsgBox "程序必须在魔兽根目录下运行"
    End
End If
Dim iArg As Integer
iArg = Val(Left$(Command$, 1))
If iArg > 0 And iArg < 7 Then LoadZone iArg: End
End Sub
Function FileExists(ByVal FileName As String) As Boolean
Dim Temp$
    'Set Default
    FileExists = True
     'Set up error handler
On Error Resume Next

    'Attempt to grab date and time
    Temp$ = FileDateTime(FileName)

    'Process errors
    Select Case Err
        Case 53, 76, 68   'File Does Not Exist
            FileExists = False
            Err = 0
        Case Else
            If Err <> 0 Then
                MsgBox "Error Number: " & Err & Chr$(10) & Chr$(13) & " " & Error, vbOKOnly, "Error"
                End
            End If
    End Select
End Function
