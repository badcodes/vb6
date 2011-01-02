VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FTestSplitter 
   AutoRedraw      =   -1  'True
   Caption         =   "Test Splitters"
   ClientHeight    =   5532
   ClientLeft      =   1092
   ClientTop       =   3516
   ClientWidth     =   6876
   ClipControls    =   0   'False
   DrawStyle       =   6  'Inside Solid
   Icon            =   "TSPLIT2.frx":0000
   LinkTopic       =   "Form1"
   MouseIcon       =   "TSPLIT2.frx":0CFA
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5532
   ScaleWidth      =   6876
   Begin MSComctlLib.StatusBar stat 
      Align           =   2  'Align Bottom
      Height          =   372
      Left            =   0
      TabIndex        =   7
      Top             =   5160
      Width           =   6876
      _ExtentX        =   12129
      _ExtentY        =   656
      SimpleText      =   "Status Bar Text"
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "Status Bar Text"
            TextSave        =   "Status Bar Text"
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtB 
      Height          =   1260
      Left            =   3588
      MultiLine       =   -1  'True
      TabIndex        =   4
      Text            =   "TSPLIT2.frx":0F0C
      Top             =   2268
      Visible         =   0   'False
      Width           =   2820
   End
   Begin VB.TextBox txtA 
      Height          =   1260
      Left            =   225
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "TSPLIT2.frx":0F28
      Top             =   1995
      Visible         =   0   'False
      Width           =   3420
   End
   Begin VB.ListBox lstB 
      Height          =   816
      Left            =   4176
      MousePointer    =   1  'Arrow
      TabIndex        =   1
      Top             =   720
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.ListBox lstA 
      Height          =   624
      Left            =   240
      MousePointer    =   1  'Arrow
      TabIndex        =   0
      Top             =   765
      Visible         =   0   'False
      Width           =   3225
   End
   Begin VB.PictureBox pbBack 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   2928
      Left            =   2640
      Picture         =   "TSPLIT2.frx":0F56
      ScaleHeight     =   2880
      ScaleWidth      =   3840
      TabIndex        =   2
      Top             =   1152
      Visible         =   0   'False
      Width           =   3888
   End
   Begin VB.PictureBox pbB 
      AutoRedraw      =   -1  'True
      Height          =   3120
      Left            =   3255
      Picture         =   "TSPLIT2.frx":A5D8
      ScaleHeight     =   3072
      ScaleWidth      =   3072
      TabIndex        =   6
      Top             =   1425
      Visible         =   0   'False
      Width           =   3120
   End
   Begin VB.PictureBox pbA 
      Height          =   3870
      Left            =   195
      Picture         =   "TSPLIT2.frx":1265A
      ScaleHeight     =   3828
      ScaleWidth      =   3828
      TabIndex        =   5
      Top             =   870
      Visible         =   0   'False
      Width           =   3870
   End
   Begin VB.Menu mnuOption 
      Caption         =   "&Options"
      Begin VB.Menu mnuVertical 
         Caption         =   "&Vertical"
      End
      Begin VB.Menu mnuPicture 
         Caption         =   "&Background Picture"
      End
      Begin VB.Menu mnuShowDrag 
         Caption         =   "&Show Drag"
      End
      Begin VB.Menu mnuAutoSize 
         Caption         =   "&Auto Size"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuControl 
         Caption         =   "&Control"
         Begin VB.Menu mnuControls 
            Caption         =   "&List Boxes"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu mnuControls 
            Caption         =   "&Text Boxes"
            Index           =   1
         End
         Begin VB.Menu mnuControls 
            Caption         =   "&Picture Box"
            Index           =   2
         End
      End
   End
End
Attribute VB_Name = "FTestSplitter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private split As New CSplitter
Private ctlNW  As Object, ctlSE As Object
Private ordControl As Integer

Private Sub cboSplit_Click()
    NewSplit
End Sub

Private Sub mnuControls_Click(Index As Integer)
    Dim i As Integer
    ordControl = Index
    For i = 0 To 2
        mnuControls(i).Checked = IIf(i = Index, vbChecked, vbUnchecked)
    Next
    NewSplit
End Sub

Private Sub mnuPicture_Click()
    mnuPicture.Checked = Not mnuPicture.Checked
    NewSplit
End Sub

Private Sub mnuAutoSize_Click()
    mnuAutoSize.Checked = Not mnuAutoSize.Checked
    NewSplit
End Sub

Private Sub mnuShowDrag_Click()
    mnuShowDrag.Checked = Not mnuShowDrag.Checked
    NewSplit
End Sub

Private Sub mnuVertical_Click()
    mnuVertical.Checked = Not mnuVertical.Checked
    NewSplit
End Sub

Private Sub Form_Load()
    Dim iX As Integer
    For iX = 1 To 200
        lstA.AddItem "List Item Number " & iX
        lstB.AddItem "List Item Number " & iX
    Next iX
    Show
    With pbBack
        .Visible = True
        .Left = (ScaleWidth / 2) - (.Width / 2)
        .Top = (ScaleHeight / 2) - (.Height / 2)
    End With
    NewSplit
End Sub

Sub NewSplit()
    Set split = Nothing
    
    ' This is unnecessarily complicated because of multiple surfaces
    ' and options; in real programs you normally choose options at
    ' design time
    
    If Not ctlNW Is Nothing Then ctlNW.Visible = False
    If Not ctlSE Is Nothing Then ctlSE.Visible = False
    
    Select Case ordControl
    Case 0 ' List boxes
        ' List boxes control their own borders, so bottom
        ' border may look too large
        Set ctlNW = lstA
        Set ctlSE = lstB
    Case 1 ' Text boxes
        Set ctlNW = txtA
        Set ctlSE = txtB
    Case 2 ' Picture boxes
        Set ctlNW = pbA
        Set ctlSE = pbB
    End Select
    If mnuPicture.Checked Then
        pbBack.Visible = True
        Set ctlNW.Container = pbBack
        Set ctlSE.Container = pbBack
        ctlNW.Visible = True
        ctlSE.Visible = True
        ctlNW.Left = pbBack.Width * 0.05
        ctlNW.Top = pbBack.Height * 0.05
        ctlNW.Width = pbBack.Width * 0.6
        ctlNW.Height = pbBack.Height * 0.6
        ctlSE.Left = ctlNW.Width + 1
        ctlSE.Top = ctlNW.Height + 1
    Else
        pbBack.Visible = False
        Set ctlNW.Container = Me
        Set ctlSE.Container = Me
        ctlNW.Left = Width * 0.05
        ctlNW.Top = Height * 0.05
        ctlNW.Width = Width * 0.6
        ctlNW.Height = Height * 0.6
        ctlSE.Left = ctlNW.Width + 1
        ctlSE.Top = ctlNW.Height + 1
        ctlNW.Visible = True
        ctlSE.Visible = True
    End If
    
    Dim fAuto As Boolean, fVertical As Boolean, fShowDrag As Boolean
    Dim picDrag As StdPicture
    fAuto = mnuAutoSize.Checked
    fVertical = mnuVertical.Checked
    fShowDrag = mnuShowDrag.Checked
    
    On Error Resume Next
    split.Create LeftControl:=ctlNW, _
                 RightControl:=ctlSE, _
                 Vertical:=fVertical, _
                 BorderPixels:=4, _
                 AutoBorder:=fAuto, _
                 Resizeable:=True, _
                 Percent:=50, _
                 ShowDrag:=fShowDrag
    If Err Then MsgBox "Can't create splitter"
End Sub

' Normally events simply pass through four simple event handlers,
' but multiple controls and options complicate things here
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not mnuPicture.Checked Then
        split.Splitter_MouseDown Button, Shift, X, Y
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not mnuPicture.Checked Then
        split.Splitter_MouseMove Button, Shift, X, Y
    End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not mnuPicture.Checked Then
        split.Splitter_MouseUp Button, Shift, X, Y
    End If
End Sub

Private Sub Form_Resize()
    If Not mnuPicture.Checked Then split.Splitter_Resize
End Sub

Private Sub pbBack_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mnuPicture.Checked Then
        split.Splitter_MouseDown Button, Shift, X, Y
    End If
End Sub

Private Sub pbBack_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mnuPicture.Checked Then
        split.Splitter_MouseMove Button, Shift, X, Y
    End If
End Sub

Private Sub pbBack_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mnuPicture.Checked Then
        split.Splitter_MouseUp Button, Shift, X, Y
    End If
End Sub

Private Sub pbBack_Resize()
    If mnuPicture.Checked Then split.Splitter_Resize
End Sub



