VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "textPDGMerger"
   ClientHeight    =   3540
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3540
   ScaleWidth      =   9000
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdProcess 
      Caption         =   "运行"
      Height          =   495
      Left            =   7560
      TabIndex        =   7
      Top             =   2790
      Width           =   1215
   End
   Begin VB.TextBox txtInfo 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   1470
      Left            =   225
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   1815
      Width           =   7125
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "选择..."
      Height          =   345
      Index           =   1
      Left            =   7650
      TabIndex        =   5
      Top             =   1290
      Width           =   1000
   End
   Begin VB.TextBox txtPath 
      Height          =   375
      Index           =   1
      Left            =   240
      TabIndex        =   4
      Top             =   1290
      Width           =   7140
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "选择..."
      Height          =   345
      Index           =   0
      Left            =   7650
      TabIndex        =   2
      Top             =   450
      Width           =   1000
   End
   Begin VB.TextBox txtPath 
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   450
      Width           =   7110
   End
   Begin VB.Label Label1 
      Caption         =   "目标PDF文件:"
      Height          =   270
      Index           =   1
      Left            =   240
      TabIndex        =   3
      Top             =   1020
      Width           =   1560
   End
   Begin VB.Label Label1 
      Caption         =   "文本PDG文件夹:"
      Height          =   270
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   180
      Width           =   1560
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdProcess_Click()
    If CheckApp = False Then Exit Sub
    'MShell.ShellExecute Me.hWnd, "open", "mbtPdfAsm.exe", "-b" & " -m00.*\.pdf" & " -o" & QuoteString(txtPath(1).Text), App.Path, SW_SHOWNORMAL
     MShell.ShellExecute Me.hWnd, "open", App.Path & "\mbtPdfAsm.exe", " -m00.*\.pdf" & " -d" & QuoteString(txtPath(1).Text), txtPath(0).Text, SW_SHOWNORMAL
End Sub
Private Function QuoteString(ByRef vStr As String) As String
    QuoteString = Chr$(34) & vStr & Chr$(34)
End Function
Private Function CheckApp() As Boolean
    CheckApp = False
    
End Function
Private Sub cmdSelect_Click(Index As Integer)
    Select Case Index
    
        Case 0
            txtPath(0).Text = FolderBrowser_Select(txtPath(0).Text)
            txtPath(1).Text = txtPath(0).Text & ".PDF"
        Case 1
            txtPath(1).Text = selectFile(txtPath(1).Text)
    End Select
End Sub

Private Sub Form_Load()
    Dim configFile As String
    configFile = App.Path & "\" & App.EXEName & ".ini"
    Dim iniHnd As CLiNInI
    Set iniHnd = New CLiNInI
    iniHnd.source = configFile
    AppSetting_LoadAllTextBox iniHnd, Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim configFile As String
    configFile = App.Path & "\" & App.EXEName & ".ini"
    Dim iniHnd As CLiNInI
    Set iniHnd = New CLiNInI
    iniHnd.source = configFile
    AppSetting_SaveAllTextBox iniHnd, Me
    iniHnd.Save
End Sub

Private Function selectFile(ByRef vIntFile As String) As String
    Dim dlg As CCommonDialogLite
    Set dlg = New CCommonDialogLite
    Dim ret As String
    ret = vIntFile
    If dlg.VBGetSaveFileName(ret, "合并PDF目标文件", True, "PDF 文件(*.pdf)| *.pdf|All(*.*)| *.*", , , , ".pdf", Me.hWnd) Then
        selectFile = ret
    Else
        selectFile = vIntFile
    End If
End Function


