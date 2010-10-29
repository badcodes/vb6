VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form ProgressDialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Loading..."
   ClientHeight    =   855
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5385
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   855
   ScaleWidth      =   5385
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ProgressBar ProgressBar 
      Height          =   330
      Left            =   75
      TabIndex        =   0
      ToolTipText     =   "Any key to cancel"
      Top             =   390
      Width           =   5205
      _ExtentX        =   9181
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   150
      TabIndex        =   1
      Top             =   120
      Width           =   90
   End
End
Attribute VB_Name = "ProgressDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public MSTimeOut As Integer
Public Title As String
Public Text As String

Public Event Canceled()
Public Event Progressed()

Private fStart As Long
Private fClicked As Boolean
Public Sub Run()
    Dim fEnd As Long
    Dim dStep As Double
    If MSTimeOut < 0 Then MSTimeOut = 500
    If Title = "" Then Title = "Loading"
    If Text = "" Then Text = ""
    fClicked = False
    Me.Caption = Title
    Me.lblInfo.Caption = Text
    
    fStart = GetTime
    
    ProgressBar.Min = 0
    ProgressBar.Max = MSTimeOut / 10
    
    Dim fNow As Long
    
    Do
        DoEvents
        fNow = GetTime - fStart
        If fNow > ProgressBar.Max Then
            ProgressBar.Value = ProgressBar.Max
            Exit Do
        Else
            ProgressBar.Value = fNow
        End If
        If fClicked Then Exit Do
    Loop
    
    If fClicked Then
        RaiseEvent Canceled
    Else
        RaiseEvent Progressed
    End If
    
End Sub

Private Function GetTime() As Long
    GetTime = DateTime.Timer * 100
End Function

Private Sub Form_Activate()
    Me.Run
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    fClicked = True
End Sub

Private Sub lblInfo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    fClicked = True
End Sub

Private Sub ProgressBar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    fClicked = True
End Sub
