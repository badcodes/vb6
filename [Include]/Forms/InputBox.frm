VERSION 5.00
Begin VB.Form frmInputBox 
   Caption         =   "InputBox"
   ClientHeight    =   4200
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8085
   LinkTopic       =   "Form1"
   ScaleHeight     =   4200
   ScaleWidth      =   8085
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消"
      Height          =   375
      Left            =   6705
      TabIndex        =   2
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确认"
      Height          =   375
      Left            =   5280
      TabIndex        =   1
      Top             =   3705
      Width           =   1215
   End
   Begin VB.TextBox Text 
      Height          =   2340
      Left            =   105
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   660
      Width           =   7785
   End
   Begin VB.Label Info 
      BackStyle       =   0  'Transparent
      Caption         =   "Info"
      Height          =   465
      Left            =   105
      TabIndex        =   3
      Top             =   120
      Width           =   7815
   End
End
Attribute VB_Name = "frmInputBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const cstLeft As Single = 120
Private Const cstTop As Single = 120

Private Sub cmdCancel_Click()
    Text.Text = ""
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    Me.Hide
End Sub

Private Sub Form_Activate()
 Text.SetFocus
         Text.SelStart = 0
        Text.SelLength = Len(Text.Text)
End Sub

Private Sub Form_Load()
    Form_Resize
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Dim pH As Single
    Dim pW As Single
    pH = Me.ScaleHeight - 3 * cstTop - cmdOK.Height
    pW = Me.ScaleWidth - 2 * cstLeft ' - Info.Width
    'Info.Move cstLeft, cstTop, Info.Width, pH
    Text.Move cstLeft, cstTop + Info.Top + Info.Height '
    Text.Height = pH - Text.Top '- cstTop
    Text.Width = pW
    ' Left + Info.Width + cstLeft, cstTop, pW, pH
    cmdCancel.Move Me.ScaleWidth - cmdCancel.Width - cstLeft, Text.Top + Text.Height + cstTop
    cmdOK.Move cmdCancel.Left - cstLeft - cmdOK.Width, cmdCancel.Top
End Sub

Public Function Popup(Optional vText As String, Optional vTitle As String, Optional vInfo As String, Optional x As Single, Optional y As Single)
    If x > 0 Or y > 0 Then Me.Move x, y
    
    
    If vTitle = "" Then vTitle = App.EXEName
    Me.Caption = vTitle
    Info.Caption = vInfo
    Text.Text = vText
    Me.Show 1
    Popup = Text.Text
End Function



Private Sub Text_KeyDown(KeyCode As Integer, Shift As Integer)
Debug.Print "Shift:" & Shift & vbTab & "Code:" & KeyCode & "(" & Chr$(KeyCode) & ")"
    If Shift = 1 And KeyCode = 97 Or KeyCode = 65 Then
        Text.SelStart = 0
        Text.SelLength = Len(Text.Text)
        Exit Sub
    End If
   
End Sub

