VERSION 5.00
Object = "{831FDD16-0C5C-11d2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmTreeviewListviewTemplate 
   Caption         =   "ffff"
   ClientHeight    =   5220
   ClientLeft      =   -9996
   ClientTop       =   2292
   ClientWidth     =   5676
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5220
   ScaleWidth      =   5676
   Begin VB.PictureBox picTitles 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   0
      ScaleHeight     =   462.08
      ScaleMode       =   0  'User
      ScaleWidth      =   5797.147
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   5676
      Begin VB.Label lblTitle 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Treeview:"
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   3
         Top             =   15
         Width           =   1995
      End
      Begin VB.Label lblTitle 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Listview:"
         Height          =   255
         Index           =   1
         Left            =   2010
         TabIndex        =   2
         Top             =   15
         Width           =   3180
      End
   End
   Begin VB.PictureBox picSplitter 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      FillColor       =   &H00808080&
      Height          =   4800
      Left            =   5400
      ScaleHeight     =   5012.637
      ScaleMode       =   0  'User
      ScaleWidth      =   260
      TabIndex        =   0
      Top             =   315
      Visible         =   0   'False
      Width           =   156
   End
   Begin MSComctlLib.TreeView tvTreeView 
      Height          =   4800
      Left            =   -15
      TabIndex        =   5
      Top             =   315
      Width           =   2010
      _ExtentX        =   3535
      _ExtentY        =   8467
      _Version        =   393217
      Style           =   7
      Appearance      =   1
   End
   Begin MSComctlLib.ListView lvListView 
      Height          =   4800
      Left            =   2040
      TabIndex        =   4
      Top             =   315
      Width           =   3210
      _ExtentX        =   5673
      _ExtentY        =   8467
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Image imgSplitter 
      Height          =   4785
      Left            =   1965
      MouseIcon       =   "Treeview Listview Splitter.frx":0000
      MousePointer    =   99  'Custom
      Top             =   315
      Width           =   150
   End
End
Attribute VB_Name = "frmTreeviewListviewTemplate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Const NAME_COLUMN = 0
Const TYPE_COLUMN = 1
Const SIZE_COLUMN = 2
Const DATE_COLUMN = 3

Dim mbMoving As Boolean
Const sglSplitLimit = 500


Private Sub Form_Resize()
  If Me.Width < 3000 Then Me.Width = 3000
  SizeControls imgSplitter.Left
End Sub

Private Sub imgSplitter_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  With imgSplitter
    picSplitter.Move .Left, .Top, .Width - 20, .Height - 20
  End With
  picSplitter.Visible = True
  mbMoving = True
End Sub

Private Sub imgSplitter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim sglPos As Single
  
  If mbMoving Then
    sglPos = X + imgSplitter.Left
    If sglPos < sglSplitLimit Then
      picSplitter.Left = sglSplitLimit
    ElseIf sglPos > Me.Width - sglSplitLimit Then
      picSplitter.Left = Me.Width - sglSplitLimit
    Else
      picSplitter.Left = sglPos
    End If
  End If
End Sub
Private Sub imgSplitter_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  SizeControls picSplitter.Left
  picSplitter.Visible = False
  mbMoving = False
End Sub

Sub SizeControls(X As Single)
  On Error Resume Next
  
  'set the width
  If X < 1500 Then X = 1500
  If X > (Me.Width - 1500) Then X = Me.Width - 1500
  tvTreeView.Width = X
  imgSplitter.Left = X
  lvListView.Left = X + 40
  lvListView.Width = Me.Width - (tvTreeView.Width + 140)
  lblTitle(0).Width = tvTreeView.Width
  lblTitle(1).Left = lvListView.Left + 20
  lblTitle(1).Width = lvListView.Width - 40

  'set the top
'  If tbToolBar.Visible Then
'    tvTreeView.Top = tbToolBar.Height + picTitles.Height
'  Else
    tvTreeView.Top = picTitles.Height
'  End If
  lvListView.Top = tvTreeView.Top
  
  'set the height
'  If sbStatusBar.Visible Then
'    tvTreeView.Height = Me.ScaleHeight - (picTitles.Top + picTitles.Height + sbStatusBar.Height)
'  Else
    tvTreeView.Height = Me.ScaleHeight - (picTitles.Top + picTitles.Height) ' + h)
'  End If
  
  lvListView.Height = tvTreeView.Height
  imgSplitter.Top = tvTreeView.Top
  imgSplitter.Height = tvTreeView.Height
End Sub

Private Sub tvTreeView_DragDrop(Source As Control, X As Single, Y As Single)
  If Source = imgSplitter Then
    SizeControls X
  End If
End Sub
