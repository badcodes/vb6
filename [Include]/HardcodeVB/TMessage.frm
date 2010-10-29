VERSION 5.00
Begin VB.Form FTestMessages 
   Caption         =   "Test Message Capture"
   ClientHeight    =   4356
   ClientLeft      =   7476
   ClientTop       =   3336
   ClientWidth     =   3996
   Icon            =   "TMessage.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4356
   ScaleWidth      =   3996
   Begin VB.Frame fmCapture 
      Caption         =   "Message Capture"
      Height          =   3984
      Left            =   168
      TabIndex        =   9
      Top             =   204
      Width           =   3636
      Begin VB.TextBox txtMinMax 
         Height          =   288
         Index           =   0
         Left            =   180
         TabIndex        =   1
         Top             =   1524
         Width           =   1296
      End
      Begin VB.TextBox txtMinMax 
         Height          =   288
         Index           =   1
         Left            =   2136
         TabIndex        =   2
         Top             =   1512
         Width           =   1296
      End
      Begin VB.TextBox txtMinMax 
         Height          =   288
         Index           =   2
         Left            =   180
         TabIndex        =   3
         Top             =   2172
         Width           =   1296
      End
      Begin VB.TextBox txtMinMax 
         Height          =   288
         Index           =   3
         Left            =   2136
         TabIndex        =   4
         Top             =   2172
         Width           =   1296
      End
      Begin VB.TextBox txtMinMax 
         Height          =   288
         Index           =   4
         Left            =   180
         TabIndex        =   5
         Top             =   2820
         Width           =   1296
      End
      Begin VB.TextBox txtMinMax 
         Height          =   288
         Index           =   5
         Left            =   2136
         TabIndex        =   6
         Top             =   2820
         Width           =   1296
      End
      Begin VB.TextBox txtMinMax 
         Height          =   288
         Index           =   6
         Left            =   180
         TabIndex        =   7
         Top             =   3468
         Width           =   1296
      End
      Begin VB.TextBox txtMinMax 
         Height          =   288
         Index           =   7
         Left            =   2136
         TabIndex        =   8
         Top             =   3480
         Width           =   1296
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "New Form"
         Height          =   396
         Left            =   168
         TabIndex        =   0
         Top             =   660
         Width           =   3288
      End
      Begin VB.Label lbl 
         Caption         =   "Check the system menu"
         Height          =   324
         Index           =   0
         Left            =   888
         TabIndex        =   18
         Top             =   264
         Width           =   1764
      End
      Begin VB.Label lbl 
         Caption         =   "Minimum width: "
         Height          =   228
         Index           =   1
         Left            =   180
         TabIndex        =   17
         Top             =   1224
         Width           =   1296
      End
      Begin VB.Label lbl 
         Caption         =   "Maximized left:"
         Height          =   228
         Index           =   4
         Left            =   180
         TabIndex        =   16
         Top             =   2520
         Width           =   1296
      End
      Begin VB.Label lbl 
         Caption         =   "Maximimum height:"
         Height          =   228
         Index           =   5
         Left            =   2136
         TabIndex        =   15
         Top             =   1872
         Width           =   1428
      End
      Begin VB.Label lbl 
         Caption         =   "Maximimum width:"
         Height          =   228
         Index           =   6
         Left            =   180
         TabIndex        =   14
         Top             =   1872
         Width           =   1296
      End
      Begin VB.Label lbl 
         Caption         =   "Minimum height:"
         Height          =   228
         Index           =   7
         Left            =   2136
         TabIndex        =   13
         Top             =   1212
         Width           =   1296
      End
      Begin VB.Label lbl 
         Caption         =   "Maximized width:"
         Height          =   228
         Index           =   8
         Left            =   2136
         TabIndex        =   12
         Top             =   3180
         Width           =   1296
      End
      Begin VB.Label lbl 
         Caption         =   "Maximized width:"
         Height          =   228
         Index           =   9
         Left            =   180
         TabIndex        =   11
         Top             =   3168
         Width           =   1296
      End
      Begin VB.Label lbl 
         Caption         =   "Maximized top:"
         Height          =   228
         Index           =   10
         Left            =   2136
         TabIndex        =   10
         Top             =   2532
         Width           =   1296
      End
   End
End
Attribute VB_Name = "FTestMessages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RepId" ,"4FB51841-CEAF-11CF-A15E-00AA00A74D48-0050"
Option Explicit

Private WithEvents sysmenu As CSysMenu
Attribute sysmenu.VB_VarHelpID = -1
Private idAbout As Long
Private idAround As Long

Private minmax As New CMinMax
Attribute minmax.VB_VarHelpID = -1
Enum EMinMax
    emmMinWidth
    emmMinHeight
    emmMaxWidth
    emmMaxHeight
    emmMaximizedLeft
    emmMaximizedTop
    emmMaximizedWidth
    emmMaximizedHeight
End Enum

Private Sub cmdNew_Click()
    Dim frm As New FTestMessages
    frm.Left = Left + Width * 0.1
    frm.Top = Top + Height * 0.1
    frm.Top = Top + Height * 0.1
    frm.Caption = Caption & "A"
    frm.Show
End Sub

Private Sub Form_Load()
    ' Initialize system menu
    Set sysmenu = New CSysMenu
    sysmenu.Create hWnd
    Call sysmenu.AddItem("-")   ' Separator
    idAbout = sysmenu.AddItem("About...")
    idAround = sysmenu.AddItem("Around...")
    'Show
    
    ' Initialize minimums and maximums
    With minmax
        .Create hWnd
        .MinWidth = fmCapture.Width
        txtMinMax(emmMinWidth) = .MinWidth
        .MinHeight = fmCapture.Height
        txtMinMax(emmMinHeight) = .MinHeight
        .MaxWidth = Screen.Width * 0.8
        txtMinMax(emmMaxWidth) = .MaxWidth
        .MaxHeight = Screen.Height * 0.8
        txtMinMax(emmMaxHeight) = .MaxHeight
        .MaximizedLeft = Screen.Width * 0.1
        txtMinMax(emmMaximizedLeft) = .MaximizedLeft
        .MaximizedTop = Screen.Height * 0.1
        txtMinMax(emmMaximizedTop) = .MaximizedTop
        .MaximizedWidth = Screen.Width * 0.8
        txtMinMax(emmMaximizedWidth) = .MaximizedWidth
        .MaximizedHeight = Screen.Height * 0.8
        txtMinMax(emmMaximizedHeight) = .MaximizedHeight
    End With
End Sub

Private Sub Form_Activate()
    txtMinMax(0).SetFocus
End Sub

Private Sub Form_Resize()
    fmCapture.Left = (ScaleWidth / 2) - (fmCapture.Width / 2)
    fmCapture.Top = (ScaleHeight / 2) - (fmCapture.Height / 2)
End Sub

Private Sub sysmenu_MenuClick(sItem As String, ByVal ID As Long)
    Select Case ID
    Case idAbout
        MsgBox "About time"
    Case idAround
        MsgBox "Around and around"
    End Select
End Sub

Private Sub txtMinMax_GotFocus(Index As Integer)
With txtMinMax(Index)
    .SelLength = Len(.Text)
End With
End Sub

Private Sub txtMinMax_LostFocus(Index As Integer)
With minmax
    Select Case Index
    Case emmMaximizedLeft
        .MaximizedLeft = txtMinMax(emmMaximizedLeft)
    Case emmMaximizedTop
        .MaximizedTop = txtMinMax(emmMaximizedTop)
    Case emmMaximizedWidth
        .MaximizedWidth = txtMinMax(emmMaximizedWidth)
    Case emmMaximizedHeight
        .MaximizedHeight = txtMinMax(emmMaximizedHeight)
    Case emmMinWidth
        .MinWidth = txtMinMax(emmMinWidth)
    Case emmMinHeight
        .MinHeight = txtMinMax(emmMinHeight)
    Case emmMaxWidth
        .MaxWidth = txtMinMax(emmMaxWidth)
    Case emmMaxHeight
        .MaxHeight = txtMinMax(emmMaxHeight)
    End Select
    ' Update all because some changes affect others
    txtMinMax(emmMinWidth) = .MinWidth
    txtMinMax(emmMinHeight) = .MinHeight
    txtMinMax(emmMaxWidth) = .MaxWidth
    txtMinMax(emmMaxHeight) = .MaxHeight
    txtMinMax(emmMaximizedLeft) = .MaximizedLeft
    txtMinMax(emmMaximizedTop) = .MaximizedTop
    txtMinMax(emmMaximizedWidth) = .MaximizedWidth
    txtMinMax(emmMaximizedHeight) = .MaximizedHeight
End With
End Sub
