VERSION 5.00
Begin VB.Form frmListPickerTemplate 
   Caption         =   "Form1"
   ClientHeight    =   1872
   ClientLeft      =   1068
   ClientTop       =   1668
   ClientWidth     =   6420
   LinkTopic       =   "Form1"
   ScaleHeight     =   1872
   ScaleWidth      =   6420
   Begin VB.CommandButton cmdDown 
      Height          =   435
      Left            =   5835
      Picture         =   "Mover ListBox.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1110
      Width           =   435
   End
   Begin VB.CommandButton cmdUp 
      Height          =   435
      Left            =   5820
      Picture         =   "Mover ListBox.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   525
      Width           =   435
   End
   Begin VB.CommandButton cmdLeftAll 
      Caption         =   "<<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   336
      Left            =   2550
      MaskColor       =   &H00000000&
      TabIndex        =   5
      Top             =   1440
      Width           =   576
   End
   Begin VB.CommandButton cmdLeftOne 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   336
      Left            =   2550
      MaskColor       =   &H00000000&
      TabIndex        =   4
      Top             =   1065
      Width           =   576
   End
   Begin VB.CommandButton cmdRightAll 
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   336
      Left            =   2550
      MaskColor       =   &H00000000&
      TabIndex        =   3
      Top             =   690
      Width           =   576
   End
   Begin VB.CommandButton cmdRightOne 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   336
      Left            =   2550
      MaskColor       =   &H00000000&
      TabIndex        =   2
      Top             =   315
      Width           =   576
   End
   Begin VB.ListBox lstSelected 
      Height          =   1392
      Left            =   3435
      TabIndex        =   1
      Top             =   315
      Width           =   2220
   End
   Begin VB.ListBox lstAll 
      Height          =   1392
      Left            =   75
      TabIndex        =   0
      Top             =   315
      Width           =   2220
   End
   Begin VB.Label lblSelected 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Selected Items:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3435
      TabIndex        =   9
      Tag             =   "2407"
      Top             =   60
      Width           =   1095
   End
   Begin VB.Label lblAll 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "All Items:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   75
      TabIndex        =   8
      Tag             =   "2406"
      Top             =   60
      Width           =   630
   End
End
Attribute VB_Name = "frmListPickerTemplate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdUp_Click()
  On Error Resume Next
  Dim nItem As Integer
  
  With lstSelected
    If .ListIndex < 0 Then Exit Sub
    nItem = .ListIndex
    If nItem = 0 Then Exit Sub  'can't move 1st item up
    'move item up
    .AddItem .Text, nItem - 1
    'remove old item
    .RemoveItem nItem + 1
    'select the item that was just moved
    .Selected(nItem - 1) = True
  End With
End Sub

Private Sub cmdDown_Click()
  On Error Resume Next
  Dim nItem As Integer
  
  With lstSelected
    If .ListIndex < 0 Then Exit Sub
    nItem = .ListIndex
    If nItem = .ListCount - 1 Then Exit Sub 'can't move last item down
    'move item down
    .AddItem .Text, nItem + 2
    'remove old item
    .RemoveItem nItem
    'select the item that was just moved
    .Selected(nItem + 1) = True
  End With
End Sub

Private Sub cmdRightOne_Click()
  On Error Resume Next
  Dim i As Integer
  
  If lstAll.ListCount = 0 Then Exit Sub
  
  lstSelected.AddItem lstAll.Text
  i = lstAll.ListIndex
  lstAll.RemoveItem lstAll.ListIndex
  If lstAll.ListCount > 0 Then
    If i > lstAll.ListCount - 1 Then
      lstAll.ListIndex = i - 1
    Else
      lstAll.ListIndex = i
    End If
  End If
  lstSelected.ListIndex = lstSelected.NewIndex
End Sub

Private Sub cmdRightAll_Click()
  On Error Resume Next
  Dim i As Integer
  For i = 0 To lstAll.ListCount - 1
    lstSelected.AddItem lstAll.List(i)
  Next
  lstAll.Clear
  lstSelected.ListIndex = 0
End Sub

Private Sub cmdLeftOne_Click()
  On Error Resume Next
  Dim i As Integer
  
  If lstSelected.ListCount = 0 Then Exit Sub
  
  lstAll.AddItem lstSelected.Text
  i = lstSelected.ListIndex
  lstSelected.RemoveItem i
  
  lstAll.ListIndex = lstAll.NewIndex
  If lstSelected.ListCount > 0 Then
    If i > lstSelected.ListCount - 1 Then
      lstSelected.ListIndex = i - 1
    Else
      lstSelected.ListIndex = i
    End If
  End If
End Sub

Private Sub cmdLeftAll_Click()
  On Error Resume Next
  Dim i As Integer
  For i = 0 To lstSelected.ListCount - 1
    lstAll.AddItem lstSelected.List(i)
  Next
  lstSelected.Clear
  lstAll.ListIndex = lstAll.NewIndex

End Sub

Private Sub Form_Load()
  lstAll.AddItem "aaa"
  lstAll.AddItem "bbb"
  lstAll.AddItem "ccc"
  lstAll.ListIndex = 0
End Sub

Private Sub lstAll_DblClick()
  cmdRightOne_Click
End Sub

Private Sub lstSelected_DblClick()
  cmdLeftOne_Click
End Sub
