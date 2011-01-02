VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FTestSplitter 
   AutoRedraw      =   -1  'True
   Caption         =   "Test Splitters"
   ClientHeight    =   5535
   ClientLeft      =   2055
   ClientTop       =   3480
   ClientWidth     =   6870
   ClipControls    =   0   'False
   DrawStyle       =   6  'Inside Solid
   Icon            =   "TSplit.frx":0000
   LinkTopic       =   "Form1"
   MouseIcon       =   "TSplit.frx":0CFA
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5535
   ScaleWidth      =   6870
   Begin MSComctlLib.StatusBar stat 
      Align           =   2  'Align Bottom
      Height          =   372
      Left            =   0
      TabIndex        =   8
      Top             =   5160
      Width           =   6876
      _ExtentX        =   12118
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1764
            MinWidth        =   1764
            Text            =   "Vertical"
            TextSave        =   "Vertical"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2469
            MinWidth        =   2469
            Text            =   "Full form"
            TextSave        =   "Full form"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2469
            MinWidth        =   2469
            Text            =   "Size by position"
            TextSave        =   "Size by position"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "Move split lines"
            TextSave        =   "Move split lines"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1940
            MinWidth        =   1940
            Text            =   "List Box"
            TextSave        =   "List Box"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar bar 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   6870
      _ExtentX        =   12118
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Split Vertical"
            Object.ToolTipText     =   "Split Vertical"
            ImageKey        =   "Split Vertical"
            Style           =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Background"
            Object.ToolTipText     =   "Background Picture"
            ImageKey        =   "Background Picture"
            Style           =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Auto Size"
            Object.ToolTipText     =   "Auto Size"
            ImageKey        =   "Auto Size"
            Style           =   1
            Value           =   1
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Show Drag"
            Description     =   "Show splitter when dragging"
            Object.ToolTipText     =   "Show Drag"
            ImageKey        =   "Show Drag"
            Style           =   1
         EndProperty
      EndProperty
      Begin VB.ComboBox cboControl 
         Height          =   315
         ItemData        =   "TSplit.frx":0F0C
         Left            =   1440
         List            =   "TSplit.frx":0F19
         TabIndex        =   9
         Text            =   "cboControl"
         ToolTipText     =   "Control"
         Top             =   0
         Width           =   1212
      End
   End
   Begin VB.TextBox txtB 
      Height          =   1260
      Left            =   3588
      MultiLine       =   -1  'True
      TabIndex        =   4
      Text            =   "TSplit.frx":0F3E
      Top             =   2268
      Visible         =   0   'False
      Width           =   2820
   End
   Begin VB.TextBox txtA 
      Height          =   1260
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "TSplit.frx":0F5A
      Top             =   1995
      Visible         =   0   'False
      Width           =   3420
   End
   Begin VB.ListBox lstB 
      Height          =   645
      Left            =   4176
      MousePointer    =   1  'Arrow
      TabIndex        =   1
      Top             =   720
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.ListBox lstA 
      Height          =   450
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
      Height          =   3660
      Left            =   2640
      Picture         =   "TSplit.frx":0F88
      ScaleHeight     =   3600
      ScaleWidth      =   4800
      TabIndex        =   2
      Top             =   1152
      Visible         =   0   'False
      Width           =   4860
   End
   Begin VB.PictureBox pbB 
      AutoRedraw      =   -1  'True
      Height          =   3120
      Left            =   3255
      Picture         =   "TSplit.frx":A60A
      ScaleHeight     =   3060
      ScaleWidth      =   3060
      TabIndex        =   6
      Top             =   1425
      Visible         =   0   'False
      Width           =   3120
   End
   Begin VB.PictureBox pbA 
      Height          =   3870
      Left            =   195
      Picture         =   "TSplit.frx":1268C
      ScaleHeight     =   3810
      ScaleWidth      =   3810
      TabIndex        =   5
      Top             =   870
      Visible         =   0   'False
      Width           =   3870
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   6360
      Top             =   4560
      _ExtentX        =   794
      _ExtentY        =   794
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TSplit.frx":1A70E
            Key             =   "Split Vertical"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TSplit.frx":1A820
            Key             =   "Background Picture"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TSplit.frx":1A932
            Key             =   "Auto Size"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TSplit.frx":1AA44
            Key             =   "Show Drag"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuOption 
      Caption         =   "&Options"
      Begin VB.Menu mnuVertical 
         Caption         =   "Split &Vertical"
      End
      Begin VB.Menu mnuPicture 
         Caption         =   "&Background Picture"
      End
      Begin VB.Menu mnuAutoSize 
         Caption         =   "&Auto Size"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuShowDrag 
         Caption         =   "Show &Drag"
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

Enum ECommand
    ecSplitVertical = 1
    ecBackgroundPicture
    ecAutoSize
    ecShowDrag
    ecControl
End Enum

Private split As New CSplitter
Private ctlNW  As Object, ctlSE As Object
Private ordControl As Integer

Private Sub bar_ButtonClick(ByVal Button As MSComctlLib.Button)
With Button
    Select Case Button.Index
    Case ecSplitVertical
        mnuVertical_Click
    Case ecBackgroundPicture
        mnuPicture_Click
    Case ecAutoSize
        mnuAutoSize_Click
    Case ecShowDrag
        mnuShowDrag_Click
    Case ecControl
    End Select
End With
End Sub

Private Sub cboControl_Click()
    mnuControls_Click cboControl.ListIndex
    bar.Refresh
End Sub

Private Sub mnuControls_Click(Index As Integer)
    Dim i As Integer
    ordControl = Index
    For i = 0 To 2
        mnuControls(i).Checked = IIf(i = Index, vbChecked, vbUnchecked)
    Next
    Select Case Index
    Case 0
        stat.Panels(ecControl).Text = "List Box"
    Case 1
        stat.Panels(ecControl).Text = "Text Box"
    Case 2
        stat.Panels(ecControl).Text = "Picture Box"
    End Select
    NewSplit
End Sub

Private Sub mnuPicture_Click()
    mnuPicture.Checked = Not mnuPicture.Checked
    With bar.Buttons(ecBackgroundPicture)
        .Value = -mnuPicture.Checked
    End With
    If mnuPicture.Checked Then
        stat.Panels(ecBackgroundPicture).Text = "Background picture"
    Else
        stat.Panels(ecBackgroundPicture).Text = "Full form"
    End If
    NewSplit
End Sub

Private Sub mnuAutoSize_Click()
    mnuAutoSize.Checked = Not mnuAutoSize.Checked
    With bar.Buttons(ecAutoSize)
        .Value = -mnuAutoSize.Checked
    End With
    NewSplit
End Sub

Private Sub mnuShowDrag_Click()
    mnuShowDrag.Checked = Not mnuShowDrag.Checked
    With bar.Buttons(ecShowDrag)
        .Value = -mnuShowDrag.Checked
    End With
    If mnuShowDrag.Checked Then
        stat.Panels(ecShowDrag).Text = "Show drag divider"
    Else
        stat.Panels(ecShowDrag).Text = "Show split lines"
    End If
    NewSplit
End Sub

Private Sub mnuVertical_Click()
    mnuVertical.Checked = Not mnuVertical.Checked
    With bar.Buttons(ecSplitVertical)
        .Value = -mnuVertical.Checked
    End With
    If mnuVertical.Checked Then
        stat.Panels(ecSplitVertical).Text = "Vertical"
    Else
        stat.Panels(ecSplitVertical).Text = "Horizontal"
    End If
    NewSplit
End Sub

Private Sub Form_Load()
    cboControl.ListIndex = 0
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
    
    ' This is artificially complicated because of multiple surfaces
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
        ctlSE.Left = pbBack.Width * 0.4
        ctlSE.Height = ctlNW.Height
        ctlSE.Top = pbBack.Height - ctlNW.Height - ctlNW.Left
    Else
        pbBack.Visible = False
        Set ctlNW.Container = Me
        Set ctlSE.Container = Me
        ctlNW.Visible = True
        ctlSE.Visible = True
        ctlNW.Left = ScaleWidth * 0.05
        ctlNW.Top = (ScaleHeight * 0.05) + bar.Height
        ctlNW.Width = ScaleWidth * 0.4
        ctlNW.Height = ScaleHeight * 0.4
        ctlSE.Left = ScaleWidth * 0.4
        ctlSE.Height = ctlNW.Height
        ctlSE.Top = ScaleHeight - bar.Height - (ScaleHeight * 0.05) - ctlSE.Height
    End If
    
    If mnuAutoSize.Checked = False And mnuPicture.Checked = False Then
        bar.Visible = False
        stat.Visible = False
    Else
        bar.Visible = True
        stat.Visible = True
    End If
    
    If mnuAutoSize.Checked Then
        stat.Panels(ecAutoSize).Text = "Automatic sizing"
    Else
        stat.Panels(ecAutoSize).Text = "Size by position"
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
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Not mnuPicture.Checked Then
        split.Splitter_MouseDown Button, Shift, x, y
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Not mnuPicture.Checked Then
        split.Splitter_MouseMove Button, Shift, x, y
    End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Not mnuPicture.Checked Then
        split.Splitter_MouseUp Button, Shift, x, y
    End If
End Sub

Private Sub Form_Resize()
    If Not mnuPicture.Checked Then split.Splitter_Resize
End Sub

Private Sub pbBack_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If mnuPicture.Checked Then
        split.Splitter_MouseDown Button, Shift, x, y
    End If
End Sub

Private Sub pbBack_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If mnuPicture.Checked Then
        split.Splitter_MouseMove Button, Shift, x, y
    End If
End Sub

Private Sub pbBack_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If mnuPicture.Checked Then
        split.Splitter_MouseUp Button, Shift, x, y
    End If
End Sub

Private Sub pbBack_Resize()
    If mnuPicture.Checked Then split.Splitter_Resize
End Sub



