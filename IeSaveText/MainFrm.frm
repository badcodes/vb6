VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form MainFrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SaveText"
   ClientHeight    =   5295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6015
   Icon            =   "MainFrm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   6015
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox txtSave 
      Height          =   2535
      Left            =   375
      TabIndex        =   7
      Top             =   675
      Width           =   5265
      _ExtentX        =   9287
      _ExtentY        =   4471
      _Version        =   393217
      ScrollBars      =   3
      Appearance      =   0
      TextRTF         =   $"MainFrm.frx":0442
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   3330
      Left            =   180
      TabIndex        =   6
      Top             =   150
      Width           =   5640
      _ExtentX        =   9948
      _ExtentY        =   5874
      MultiRow        =   -1  'True
      ShowTips        =   0   'False
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Text"
            Key             =   "Text"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Html"
            Key             =   "Html"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   300
      Left            =   4845
      TabIndex        =   5
      Top             =   4800
      Width           =   855
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "Save "
      Default         =   -1  'True
      Height          =   300
      Left            =   3705
      TabIndex        =   3
      Top             =   4800
      Width           =   855
   End
   Begin VB.TextBox txtFilename 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000007&
      Height          =   300
      Left            =   1365
      TabIndex        =   2
      Top             =   4320
      Width           =   4335
   End
   Begin VB.TextBox txtSaveIn 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1365
      TabIndex        =   0
      Top             =   3840
      Width           =   4335
   End
   Begin VB.Label Label2 
      Caption         =   "Filename:"
      Height          =   375
      Left            =   405
      TabIndex        =   4
      Top             =   4320
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Save in:"
      Height          =   180
      Left            =   405
      TabIndex        =   1
      Top             =   3840
      Width           =   720
   End
End
Attribute VB_Name = "MainFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim ftmpText As String
Dim ftmpHtml As String
Const MaxLenFN = 100



Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub CmdSave_Click()

'If txtSave.Text = "" Then Exit Sub
If txtSaveIn.Text = "" Then Exit Sub
If txtFilename.Text = "" Then Exit Sub
Dim fso As New FileSystemObject
Dim dstfile As String
If fso.FolderExists(txtSaveIn.Text) = False Then fso.CreateFolder (txtSaveIn.Text)
dstfile = fso.BuildPath(txtSaveIn.Text, txtFilename.Text)
txtSave.SaveFile dstfile, rtfText

End Sub

Private Sub Form_Load()

Dim ftmp As String
ftmp = Command$
If ftmp = "" Then Unload Me: End
ftmpText = ftmp & ".txt"
ftmpHtml = ftmp & ".htm"
Dim fso As New FileSystemObject
If fso.FileExists(ftmpText) = False Then Unload Me: End
If fso.FileExists(ftmpHtml) = False Then Unload Me: End
txtSaveIn.Text = GetSetting(App.ProductName, "SaveIn", "Path")
If txtSaveIn.Text = "" Then txtSaveIn.Text = App.Path
TabStrip1.Tabs("Text").Selected = True
End Sub


Private Sub Form_Unload(Cancel As Integer)
Dim fso As New FileSystemObject
If fso.FileExists(ftmpText) Then fso.DeleteFile ftmpText, True
If fso.FileExists(ftmpHtml) Then fso.DeleteFile ftmpHtml, True
SaveSetting App.ProductName, "SaveIn", "Path", txtSaveIn
End Sub


Private Sub TabStrip1_Click()

Dim fso As New FileSystemObject
Dim ts As TextStream
Dim ftmp As String
If TabStrip1.SelectedItem.Key = "Text" Then
    ftmp = ftmpText
Else
    ftmp = ftmpHtml
End If
Set ts = fso.OpenTextFile(ftmp, ForReading, False, TristateTrue)
Dim strTmp As String
Dim fl As Integer

txtSave.Text = ""
Do Until ts.AtEndOfStream
strTmp = ts.ReadLine
If ts.AtEndOfStream = False Then
    txtSave.Text = txtSave.Text + strTmp + vbCrLf
Else
    txtSave.Text = txtSave.Text + strTmp
End If
If LTrim(RTrim(strTmp)) <> "" Then
    strTmp = LTrim(RTrim(strTmp))
    fl = Len(strTmp)
    If fl > MaxLenFN Then fl = MaxLenFN
    strTmp = Left(strTmp, fl)
    If ftmp = ftmpText Then
        txtFilename = strTmp + ".txt"
    Else
        txtFilename = strTmp + ".htm"
    End If
    Exit Do
End If
Loop

Do Until ts.AtEndOfStream
strTmp = ts.ReadLine
If ts.AtEndOfStream = False Then
    txtSave.Text = txtSave.Text + strTmp + vbCrLf
Else
    txtSave.Text = txtSave.Text + strTmp
    Exit Do
End If
Loop
ts.Close

End Sub

Private Sub txtFilename_DblClick()

txtFilename.SelLength = Len(txtFilename.Text)
txtFilename.SelStart = 0

End Sub
