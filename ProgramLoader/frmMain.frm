VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Program Loader"
   ClientHeight    =   2190
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6180
   LinkTopic       =   "Form1"
   ScaleHeight     =   2190
   ScaleWidth      =   6180
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消"
      Height          =   360
      Left            =   4995
      TabIndex        =   6
      Top             =   1710
      Width           =   1065
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "确定"
      Default         =   -1  'True
      Height          =   360
      Left            =   3645
      TabIndex        =   5
      Top             =   1710
      Width           =   1065
   End
   Begin VB.TextBox txtArg 
      Height          =   345
      Left            =   90
      TabIndex        =   4
      Top             =   1185
      Width           =   6000
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "选择..."
      Height          =   360
      Left            =   5025
      TabIndex        =   2
      Top             =   435
      Width           =   1065
   End
   Begin VB.TextBox txtPath 
      Height          =   345
      Left            =   105
      TabIndex        =   1
      Top             =   435
      Width           =   4680
   End
   Begin VB.Label Label1 
      Caption         =   "参数："
      Height          =   225
      Left            =   105
      TabIndex        =   3
      Top             =   930
      Width           =   1215
   End
   Begin VB.Label lblApp 
      Caption         =   "程序路径："
      Height          =   225
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   1215
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents Progress As ProgressDialog
Attribute Progress.VB_VarHelpID = -1


Dim mArg As String

Private fShow As Boolean

Private Sub cmdCancel_Click()
Unload Me
End Sub

'CSEH: ErrMsgBox
Private Sub cmdOk_Click()
        '<EhHeader>
        On Error GoTo cmdOk_Click_Err
        cmdOk.Caption = "运行中..."
        cmdOk.Enabled = False
        '</EhHeader>
        Dim sArg As String
        Dim pApp As String
100     If txtArg.Text <> "" Then
102         sArg = txtArg.Text
        Else
104         sArg = mArg
        End If
106     pApp = txtPath.Text
    
108     If pApp = "" Then
110         MsgBox "程序路径不能为空", vbInformation + vbOKOnly, "警告"
        Else
112         MShell32.ShellExecute Me.hWnd, "Open", txtPath.Text, sArg, CurDir$, SW_SHOWNORMAL
        End If
        '<EhFooter>
        cmdOk.Caption = "确定"
        cmdOk.Enabled = True
        Unload Me
        Exit Sub

cmdOk_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in ProgramLoader.frmMain.cmdOk_Click " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdSelect_Click()
    Dim ret As String
    Dim dlg As CCommonDialogLite
    Set dlg = New CCommonDialogLite
    ret = txtPath.Text
    If dlg.VBGetOpenFileName(ret, , , , , , "Exe 文件 (*.exe)|*.exe|Cmd 文件(*.cmd)|*.cmd|Bat 文件(*.bat)|*.bat｜All (*.*)| *.*") Then
        txtPath.Text = ret
    End If
End Sub

Private Sub Form_Load()
    Dim pApp As String
    Dim pCmd As String
    Dim pPos As String
    Dim pChar As String
    pCmd = Trim$(Command$)
    pChar = Left$(pCmd, 1)
    If pChar = """" Or pChar = "'" Then
        pCmd = Mid$(pCmd, 2)
    Else
        pChar = " "
    End If
    pPos = InStr(1, pCmd, pChar)
    If pPos > 1 Then
        pApp = Mid$(pCmd, 1, pPos - 1)
        mArg = Mid$(pCmd, pPos + 1)
    Else
        pApp = pCmd
        mArg = ""
    End If
    If InStr(pApp, "/") < 1 And InStr(pApp, "\") < 1 Then
        pApp = CurDir$ & "\" & pApp
        pApp = Replace$(pApp, "/", "\")
        pApp = Replace$(pApp, "\\", "\")
    End If
    txtPath.Text = pApp
    
    fShow = True
    Me.Hide
    Set Progress = New ProgressDialog
    Progress.Title = Me.Caption
    Progress.Text = pApp
    Progress.MSTimeOut = 1000
    Progress.Show 1, Me
    Set Progress = Nothing

    If fShow Then
        Me.Show
    Else
        cmdOk_Click
        Unload Me
    End If
    
End Sub



Private Sub Progress_Canceled()
    Progress.Hide
    fShow = True

End Sub

Private Sub Progress_Progressed()
    Progress.Hide
    fShow = False
End Sub

