VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmList 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Content List"
   ClientHeight    =   4965
   ClientLeft      =   270
   ClientTop       =   1710
   ClientWidth     =   6615
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmList.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   6615
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList ImageList3 
      Left            =   4440
      Top             =   2520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmList.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmList.frx":0894
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmList.frx":0CE6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView List 
      Height          =   5055
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   8916
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   423
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "ImageList3"
      BorderStyle     =   1
      Appearance      =   0
   End
End
Attribute VB_Name = "frmList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Dim a As Integer
a = KeyCode
KeyCode = 0
MainFrm.WherevertheKeyHit a
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
KeyCode = 0
End Sub

Private Sub Form_Load()

CopytoIfont theRS.ListFont, List.Font

With theRS.ListPos
Me.Top = .Top
Me.Left = .Left
Me.Width = .Width
Me.Height = .Height
SetWindowPos Me.hWnd, hwnd_topmost, .Left, .Top, .Width, .Height, swp_nosize Or swp_nomove
End With
End Sub


Private Sub Form_Resize()

With List
.Top = 0
.Left = 0
.Height = frmList.ScaleHeight - .Top
.Width = frmList.ScaleWidth
End With
End Sub



Private Sub list_NodeClick(ByVal Node As MSComctlLib.Node)
Dim n As Integer
n = Node.Index
If Right(ztmContent(n, 1), 1) = "\" Then Exit Sub
MainFrm.GetView ztmContent(n, 1)



End Sub
