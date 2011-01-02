VERSION 5.00
Begin VB.Form frmDirSelect 
   Caption         =   "Direcotry"
   ClientHeight    =   1032
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   3744
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1032
   ScaleWidth      =   3744
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   300
      Left            =   2712
      TabIndex        =   2
      Top             =   600
      Width           =   900
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "Select"
      Height          =   300
      Left            =   1548
      TabIndex        =   1
      Top             =   612
      Width           =   960
   End
   Begin VB.TextBox txtPath 
      Height          =   288
      Left            =   108
      TabIndex        =   0
      Top             =   144
      Width           =   3504
   End
End
Attribute VB_Name = "frmDirSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Property Get dirPath() As String
    dirPath = txtPath.Text
End Property

Public Property Let dirPath(ByVal Value As String)
    txtPath.Text = Value
End Property

Private Sub cmdOK_Click()
Me.Hide
End Sub

Private Sub cmdSelect_Click()
Dim strPath As String
strPath = MDlgOpenDir.openDirDialog(Me.hwnd)
If strPath <> "" Then Me.dirPath = strPath
End Sub

Private Sub Form_Load()
    Dim iniHdl As CSetting
    Dim sINI As String
    
    sINI = App.ProductName
    If sINI = "" Then sINI = App.Title
    sINI = sINI & ".ini"
    sINI = App.Path & "\" & sINI
    Set iniHdl = New CSetting
    
    With iniHdl
        .iniFile = sINI
        .Load Me.txtPath, SF_TEXT
    End With
    
    Set iniHdl = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim iniHdl As CSetting
    Dim sINI As String
    
    sINI = App.ProductName
    If sINI = "" Then sINI = App.Title
    sINI = sINI & ".ini"
    
    Set iniHdl = New CSetting
    
    With iniHdl
        .iniFile = sINI
        .Save Me.txtPath, SF_TEXT
    End With
    Set iniHdl = Nothing
End Sub
