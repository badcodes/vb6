VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form MainFrm 
   Appearance      =   0  'Flat
   BackColor       =   &H006C5332&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BlogSave"
   ClientHeight    =   4944
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   7668
   Icon            =   "BlogSave.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4944
   ScaleWidth      =   7668
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGetClip 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "GetClipboard"
      Height          =   375
      Left            =   4812
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   204
      Width           =   1272
   End
   Begin MSComDlg.CommonDialog DLG 
      Left            =   3804
      Top             =   1704
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "txt"
      DialogTitle     =   "SaveBlog"
      Filter          =   "Text File|*.txt|All|*.*"
   End
   Begin RichTextLib.RichTextBox txtContent 
      Height          =   3972
      Left            =   216
      TabIndex        =   2
      Top             =   768
      Width           =   7116
      _ExtentX        =   12552
      _ExtentY        =   7006
      _Version        =   393217
      BackColor       =   13742484
      ScrollBars      =   2
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"BlogSave.frx":1272
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
      Left            =   6336
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   192
      Width           =   975
   End
   Begin VB.TextBox txtTitle 
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
      Top             =   216
      Width           =   4260
   End
End
Attribute VB_Name = "MainFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim blogtext As String
Dim iniApp As String
Dim fso As New gCFileSystem



Private Sub cmdGetClip_Click()


blogtext = Clipboard.GetText
txtContent.Text = blogtext
Dim i As Integer
Dim linecount As Integer
Dim linenum As Integer
Dim blogtitle As String
Dim titleline As Integer
strline = Split(blogtext, vbCrLf)
linecount = UBound(strline)
For i = 0 To linecount
tempstr = RTrim(LTrim(strline(i)))
    If tempstr <> "" Then
    If IsDate(tempstr) Then
    tempstr = Year(tempstr) + "-" + Strnum(Month(tempstr), 2) + "-" + Strnum(Day(tempstr), 2)
    End If
    blogtitle = "[" + tempstr + "]"
    titleline = i
    Exit For
    End If
Next
For i = titleline + 1 To linecount
tempstr = RTrim(LTrim(strline(i)))
    If tempstr <> "" Then
    blogtitle = blogtitle + " " + tempstr
    Exit For
    End If
Next
blogtitle = blogtitle + ".txt"

txtTitle.Text = blogtitle

End Sub

Private Sub CmdSave_Click()
Dim blogDir As String
Dim sFile As String
blogDir = iniGetSetting(iniApp, "Blogdir", "Path")
DLG.InitDir = blogDir
DLG.FileName = txtTitle.Text
DLG.CancelError = False
DLG.ShowSave
sFile = DLG.FileName
If sFile = "" Then Exit Sub
blogDir = fso.GetParentFolderName(sFile)
iniSaveSetting iniApp, "Blogdir", "Path", blogDir

Dim fNum As Integer
fNum = FreeFile
Open sFile For Output As fNum
Print #fNum, txtContent.Text
Close fNum
End Sub

Private Sub Form_Load()
iniApp = fso.BuildPath(App.Path, "config.ini")
End Sub


