VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form MainFrm 
   Appearance      =   0  'Flat
   BackColor       =   &H006C5332&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BlogWrite"
   ClientHeight    =   5172
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   9144
   Icon            =   "BlogWrite.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5172
   ScaleWidth      =   9144
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNew 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "New"
      Height          =   375
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   240
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog Dlog1 
      Left            =   3840
      Top             =   2400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "txt"
      DialogTitle     =   "SaveBlog"
      Filter          =   "Text File|*.txt|All|*.*"
   End
   Begin RichTextLib.RichTextBox rtxtContent 
      Height          =   3975
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   8655
      _ExtentX        =   15261
      _ExtentY        =   7006
      _Version        =   393217
      BackColor       =   13742484
      ScrollBars      =   2
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"BlogWrite.frx":0442
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ËÎÌו"
         Size            =   10.8
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton CmdSave 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Save "
      Height          =   375
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   240
      Width           =   975
   End
   Begin VB.TextBox TxtTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00D1B194&
      BeginProperty Font 
         Name            =   "ËÎÌו"
         Size            =   10.8
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   5895
   End
End
Attribute VB_Name = "MainFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim blogdir As String
Dim time_begin As String
Dim time_end As String
Dim iniApp As String
Dim fso As New gCFileSystem



Private Sub cmdNew_Click()
Dim m As VbMsgBoxResult
m = MsgBox("To start New Blog ?", vbYesNo, "BlogWrite")
If m = vbYes Then
TxtTitle.Text = ""
rtxtContent.Text = ""
time_begin = Date$ + " " + Time$
End If
End Sub

Private Sub CmdSave_Click()
If blogdir <> "" Then
If fso.PathExists(blogdir) = False Then MkDir (blogdir)
Dlog1.InitDir = blogdir
End If
Dlog1.FileName = "[" + Date$ + "] " + TxtTitle.Text + ".txt"
Dlog1.CancelError = False
time_end = Date$ + " " + Time$
Dlog1.ShowSave
On Error Resume Next
Dim a As MSComDlg.ErrorConstants
a = Err.Number
If a = cdlCancel Then Exit Sub
Dim thefile As String
thefile = Dlog1.FileName

If thefile = "" Then Exit Sub
blogdir = fso.GetParentFolderName(thefile)
iniSaveSetting App.ProductName & ".ini", "BlogDir", "Path", blogdir

Dim fNum As Integer
fNum = FreeFile
Open thefile For Output As #fNum
Print #fNum, rtxtContent.Text
Print #fNum, ""
Print #fNum, "     Began:  [" + time_begin + "]"
Print #fNum, "  Finished:  [" + time_end + "]"
Close #fNum

End Sub

Private Sub Form_Load()

iniApp = fso.BuildPath(App.Path, "config.ini")
blogdir = iniGetSetting(iniApp, "BlogDir", "Path")
time_begin = Date$ + " " + Time$
End Sub


