VERSION 5.00
Begin VB.Form FTestExecute 
   AutoRedraw      =   -1  'True
   Caption         =   "Test Execute"
   ClientHeight    =   3960
   ClientLeft      =   1020
   ClientTop       =   2652
   ClientWidth     =   7056
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   7.8
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "texecute.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3960
   ScaleWidth      =   7056
   Begin VB.TextBox txtPipeIn 
      Height          =   960
      Left            =   132
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   45
      Top             =   4224
      Width           =   6828
   End
   Begin VB.TextBox txtPipeOut 
      Height          =   1040
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   44
      Top             =   5472
      Width           =   6828
   End
   Begin VB.Frame fmPlus 
      Caption         =   "Extra Execute Options"
      Height          =   1695
      Left            =   4110
      TabIndex        =   39
      Top             =   45
      Width           =   2790
      Begin VB.TextBox txtLeft 
         Height          =   330
         Left            =   930
         TabIndex        =   13
         Text            =   "-1"
         Top             =   870
         Width           =   510
      End
      Begin VB.TextBox txtTop 
         Height          =   330
         Left            =   2175
         TabIndex        =   14
         Text            =   "-1"
         Top             =   870
         Width           =   510
      End
      Begin VB.TextBox txtDir 
         Height          =   315
         Left            =   105
         TabIndex        =   12
         Top             =   480
         Width           =   2580
      End
      Begin VB.TextBox txtWidth 
         Height          =   315
         Left            =   924
         TabIndex        =   15
         Text            =   "-1"
         Top             =   1275
         Width           =   510
      End
      Begin VB.TextBox txtHeight 
         Height          =   315
         Left            =   2160
         TabIndex        =   16
         Text            =   "-1"
         Top             =   1272
         Width           =   510
      End
      Begin VB.Label lbl 
         Caption         =   "Working directory:"
         Height          =   270
         Index           =   3
         Left            =   150
         TabIndex        =   26
         Top             =   255
         Width           =   1755
      End
      Begin VB.Label lbl 
         Caption         =   "Left:"
         Height          =   270
         Index           =   9
         Left            =   135
         TabIndex        =   25
         Top             =   915
         Width           =   570
      End
      Begin VB.Label lbl 
         Caption         =   "Top:"
         Height          =   270
         Index           =   10
         Left            =   1485
         TabIndex        =   23
         Top             =   915
         Width           =   570
      End
      Begin VB.Label lbl 
         Caption         =   "Width:"
         Height          =   270
         Index           =   11
         Left            =   135
         TabIndex        =   24
         Top             =   1305
         Width           =   615
      End
      Begin VB.Label lbl 
         Caption         =   "Height:"
         Height          =   270
         Index           =   12
         Left            =   1485
         TabIndex        =   40
         Top             =   1305
         Width           =   675
      End
   End
   Begin VB.Frame fmChar 
      Caption         =   "Character Mode Options"
      Height          =   1995
      Left            =   4110
      TabIndex        =   33
      Top             =   1800
      Width           =   2790
      Begin VB.CheckBox chkPipe 
         Caption         =   "Use Pipes"
         Height          =   195
         Left            =   1485
         MaskColor       =   &H00000000&
         TabIndex        =   43
         Top             =   1680
         Width           =   1260
      End
      Begin VB.ComboBox cboFront 
         Height          =   315
         ItemData        =   "texecute.frx":0CFA
         Left            =   105
         List            =   "texecute.frx":0D31
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   1290
         Width           =   1275
      End
      Begin VB.CheckBox chkFull 
         Caption         =   "Full Screen"
         Height          =   195
         Left            =   120
         MaskColor       =   &H00000000&
         TabIndex        =   22
         Top             =   1680
         Width           =   1296
      End
      Begin VB.ComboBox cboBack 
         Height          =   315
         ItemData        =   "texecute.frx":0DBE
         Left            =   1485
         List            =   "texecute.frx":0DF5
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   1290
         Width           =   1200
      End
      Begin VB.TextBox txtRow 
         Height          =   360
         Left            =   2175
         TabIndex        =   19
         Text            =   "-1"
         Top             =   660
         Width           =   510
      End
      Begin VB.TextBox txtCol 
         Height          =   345
         Left            =   930
         TabIndex        =   18
         Text            =   "-1"
         Top             =   660
         Width           =   510
      End
      Begin VB.TextBox txtTitle 
         Height          =   315
         Left            =   645
         TabIndex        =   17
         Text            =   "Test Execute"
         Top             =   240
         Width           =   2040
      End
      Begin VB.Label lbl 
         Caption         =   "Back Color:"
         Height          =   225
         Index           =   5
         Left            =   1485
         TabIndex        =   38
         Top             =   1050
         Width           =   1260
         WordWrap        =   -1  'True
      End
      Begin VB.Label lbl 
         Caption         =   "Front Color:"
         Height          =   210
         Index           =   15
         Left            =   105
         TabIndex        =   37
         Top             =   1050
         Width           =   1140
         WordWrap        =   -1  'True
      End
      Begin VB.Label lbl 
         Caption         =   "Rows:"
         Height          =   270
         Index           =   14
         Left            =   1560
         TabIndex        =   36
         Top             =   705
         Width           =   570
      End
      Begin VB.Label lbl 
         Caption         =   "Columns:"
         Height          =   270
         Index           =   13
         Left            =   120
         TabIndex        =   35
         Top             =   705
         Width           =   795
      End
      Begin VB.Label lbl 
         Caption         =   "Title:"
         Height          =   270
         Index           =   4
         Left            =   150
         TabIndex        =   34
         Top             =   285
         Width           =   600
      End
   End
   Begin VB.Timer tmrWait 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   6732
      Top             =   4032
   End
   Begin VB.TextBox txtArgs 
      Height          =   315
      Left            =   1755
      TabIndex        =   2
      Top             =   960
      Width           =   2235
   End
   Begin VB.TextBox txtProgram 
      Height          =   315
      Left            =   132
      TabIndex        =   1
      Top             =   960
      Width           =   1575
   End
   Begin VB.Frame frame 
      Caption         =   "Display Options"
      Height          =   1695
      Index           =   1
      Left            =   132
      TabIndex        =   28
      Top             =   1395
      Width           =   1680
      Begin VB.CheckBox chkMax 
         Caption         =   "Maximized"
         Height          =   255
         Left            =   120
         MaskColor       =   &H00000000&
         TabIndex        =   7
         Top             =   1320
         Width           =   1380
      End
      Begin VB.CheckBox chkMin 
         Caption         =   "Minimized"
         Height          =   255
         Left            =   120
         MaskColor       =   &H00000000&
         TabIndex        =   6
         Top             =   990
         Width           =   1485
      End
      Begin VB.CheckBox chkFocus 
         Caption         =   "Has Focus"
         Height          =   255
         Left            =   120
         MaskColor       =   &H00000000&
         TabIndex        =   5
         Top             =   660
         Value           =   1  'Checked
         Width           =   1500
      End
      Begin VB.CheckBox chkHidden 
         Caption         =   "Hidden"
         Height          =   255
         Left            =   120
         MaskColor       =   &H00000000&
         TabIndex        =   4
         Top             =   330
         Width           =   1515
      End
   End
   Begin VB.CommandButton cmdRun 
      Caption         =   "Run"
      Default         =   -1  'True
      Height          =   375
      Left            =   132
      MaskColor       =   &H00000000&
      TabIndex        =   3
      Top             =   3204
      Width           =   1215
   End
   Begin VB.Frame frame 
      Caption         =   "Execute Options"
      Height          =   1710
      Index           =   0
      Left            =   1875
      TabIndex        =   27
      Top             =   1395
      Width           =   2115
      Begin VB.CheckBox chkDead 
         Caption         =   "Dead"
         Enabled         =   0   'False
         Height          =   255
         Left            =   1068
         MaskColor       =   &H00000000&
         TabIndex        =   42
         Top             =   324
         Width           =   888
      End
      Begin VB.OptionButton optRun 
         Caption         =   "VBShellExecute"
         Height          =   270
         Index           =   2
         Left            =   120
         MaskColor       =   &H00000000&
         TabIndex        =   11
         Top             =   1320
         Width           =   1725
      End
      Begin VB.OptionButton optRun 
         Caption         =   "Executive"
         Height          =   270
         Index           =   1
         Left            =   120
         MaskColor       =   &H00000000&
         TabIndex        =   10
         Top             =   990
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.OptionButton optRun 
         Caption         =   "Shell"
         Height          =   270
         Index           =   0
         Left            =   120
         MaskColor       =   &H00000000&
         TabIndex        =   9
         Top             =   675
         Width           =   1515
      End
      Begin VB.CheckBox chkWait 
         Caption         =   "Wait"
         Height          =   255
         Left            =   120
         MaskColor       =   &H00000000&
         TabIndex        =   8
         Top             =   330
         Width           =   1095
      End
   End
   Begin VB.TextBox txtCmd 
      Height          =   300
      Left            =   132
      TabIndex        =   0
      Top             =   375
      Width           =   3900
   End
   Begin VB.Label lbl 
      Caption         =   "Piped output text:"
      Height          =   216
      Index           =   7
      Left            =   144
      TabIndex        =   47
      Top             =   5232
      Width           =   1596
   End
   Begin VB.Label lbl 
      Caption         =   "Piped input text:"
      Height          =   252
      Index           =   6
      Left            =   132
      TabIndex        =   46
      Top             =   3936
      Width           =   1572
   End
   Begin VB.Label lblExeType 
      Height          =   240
      Left            =   1692
      TabIndex        =   41
      Top             =   72
      Width           =   2436
   End
   Begin VB.Label lbl 
      Caption         =   "Arguments:"
      Height          =   255
      Index           =   2
      Left            =   1785
      TabIndex        =   29
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label lbl 
      Caption         =   "Program:"
      Height          =   255
      Index           =   1
      Left            =   132
      TabIndex        =   30
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label lblStatus 
      Height          =   636
      Left            =   1632
      TabIndex        =   32
      Top             =   3168
      Width           =   2340
   End
   Begin VB.Label lbl 
      Caption         =   "Command line:"
      Height          =   255
      Index           =   0
      Left            =   75
      TabIndex        =   31
      Top             =   75
      Width           =   1575
   End
End
Attribute VB_Name = "FTestExecute"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private idProg As Long
Private fCmdFocus As Boolean
Private sRunMode As String, sProgram As String, sArgs As String
Private dx As Long, dxBorder As Long, dy As Long, dyBorder As Long
 
Private Sub Form_Load()
    txtProgram.Text = "Notepad"
    dx = Width
    dy = Height
    dxBorder = Width - ScaleWidth
    dyBorder = Height - ScaleHeight
    optRun_Click 1
    'optRun(1).Value = True
    cboFront.ListIndex = 0
    cboBack.ListIndex = 0
    
#Const fTesting = 0
#If fTesting Then
    Dim exec As New CExecutive
    exec.Run "Notepad"

    Dim idProg As Long, iExit As Long
    idProg = Shell("mktyplib shelllnk.odl", vbHide)
    iExit = WaitOnProgram(idProg)
    If iExit Then MsgBox "Compile failed"

    With exec
        .WaitMode = ewmWaitIdle
        .Show = vbHide
        .Run "mktyplib shelllnk.odl"
        If .ExitCode Then MsgBox "Compile failed"
    End With
    With exec
        ' Notepad half the screen size 20 percent in from left and top
        .Left = Screen.Width / Screen.TwipsPerPixelX * 0.2
        .Top = Screen.Height / Screen.TwipsPerPixelY * 0.2
        .Width = Screen.Width / Screen.TwipsPerPixelX * 0.5
        .Height = Screen.Height / Screen.TwipsPerPixelY * 0.5
        .Show = vbNormalFocus
        .InitDir = Left$(CurDir$, 3)
        .WaitMode = ewmNoWait
        .Run "Notepad colors.txt"
        .Title = "The Meaning of Life"
        .Left = Screen.Width / Screen.TwipsPerPixelX * 0.1
        .Top = Screen.Height / Screen.TwipsPerPixelY * 0.1
        ' Start a red on cyan command session 70 columns by 64 rows
        .Columns = 70
        .Rows = 64
        .BackColor = qbGreen
        .ForeColor = qbLightYellow
        .Run "%COMSPEC% /k dir"
    End With
    Dim sUnsortedText As String, sSortedText As String
    sUnsortedText = "5" & sCrLf & "3" & sCrLf & "2" & sCrLf & "7" & sCrLf
    With exec
        .PipedInText = sUnsortedText
        .Show = vbHide
        .WaitMode = ewmWaitDead
        .Run "sort"
        sSortedText = .PipedOutText
    End With
#End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If chkWait = False And IsRunning(idProg) Then
        MsgBox "Can't quit until program finishes or you turn off waiting"
        Cancel = True
    End If
End Sub

Private Sub cmdRun_Click()
    Dim ewmWait As EWaitMode, hPipeRead As Long, sError As String
    ewmWait = -chkWait
    If chkDead Then ewmWait = -ewmWait
    cmdRun.SetFocus
    txtPipeOut = sEmpty
    Dim res As VbMsgBoxResult, ept As EProgramType
    Const sMsg = "You could wait forever for a hidden Windows program!"
    If txtProgram <> sEmpty Then
        ept = ExeType(SearchForExe(ExpandEnvStr(txtProgram)))
        If (chkHidden = vbChecked) And (ept <> eptMSDOS) And _
           (ept <> eptWin32Console) Then
            If MsgBox(sMsg, vbOKCancel) = vbCancel Then Exit Sub
        End If
    End If
    lblStatus.Caption = "Executing..."
    lblStatus.Refresh
    On Error GoTo RunFail
    Select Case sRunMode
    Case "Shell"
        idProg = Shell(txtCmd.Text, GetDisplay)
        If chkWait Then WaitOnProgram idProg, chkDead
        If chkWait Then
            lblStatus.Caption = sRunMode & " returned"
        Else
            lblStatus.Caption = sEmpty
            tmrWait.Enabled = True
        End If
    Case "Executive"
        Dim exec As New CExecutive, sChunk As String
        With exec
            .Show = GetDisplay
            .WaitMode = ewmWait
            If txtDir <> sEmpty Then .InitDir = txtDir
            .Left = txtLeft
            .Top = txtTop
            .Width = txtWidth
            .Height = txtHeight
            .Title = txtTitle
            .Columns = txtCol
            .Rows = txtRow
            .BackColor = cboFront.ListIndex - 1
            .ForeColor = cboBack.ListIndex - 1
            .FullScreen = (chkFull = vbChecked)
            .PipedInText = txtPipeIn
            .Run txtCmd.Text
            If .WaitMode = ewmNoWait Then
                lblStatus.Caption = sEmpty
                Do
                    Dim c As Long
                    ' A real program would do work here
                    DoWaitEvents 50
                    lblStatus.Caption = "Working: " & c
                    c = c + 1
                Loop Until .ReadPipeChunk And .Completed
            End If
            txtPipeOut = .PipedOutText
        End With
    Case "VBShellExecute"
        Dim sSpec As String, sArgs As String, sProg As String
        sArgs = txtArgs.Text
        sProg = txtProgram.Text
        If sProg = sEmpty Then
            sSpec = sArgs
        Else
            If HasShell Then
                sSpec = sProg
            Else
                sSpec = SearchForExe(sProg)
            End If
        End If
        Call VBShellExecute(sSpec, sArgs, GetDisplay)
    End Select
    Exit Sub
RunFail:
    lblStatus.Caption = sRunMode & " failed: " & _
                        Err.Description

End Sub

Private Sub chkWait_Click()
    chkDead.Enabled = chkWait
    chkDead.Value = vbUnchecked
    Debug.Print chkWait
End Sub

Private Sub chkFocus_Click()
    If chkFocus.Value = vbUnchecked And chkMax.Value = vbChecked Then
        chkMax.Value = vbUnchecked
    End If
    If chkFocus.Value = vbChecked Then chkHidden.Value = vbUnchecked
End Sub

Private Sub chkHidden_Click()
    If chkHidden.Value = vbChecked Then
        chkFocus.Value = vbUnchecked
        chkMin.Value = vbUnchecked
        chkMax.Value = vbUnchecked
    End If
    'proc.Visible = (chkHidden.Value = 0)
End Sub

Private Sub chkMax_Click()
    If chkMax.Value = vbChecked Then
        chkMin.Value = vbUnchecked
        chkFocus.Value = vbChecked
        chkHidden.Value = vbUnchecked
    End If
End Sub

Private Sub chkMin_Click()
    If chkMin.Value = vbChecked Then
        chkMax.Value = vbUnchecked
        chkHidden.Value = vbUnchecked
    End If
End Sub

Private Sub chkPipe_Click()
    Static ordLast(1 To 4) As Long
    If chkPipe Then
        Height = txtPipeOut.Top + txtPipeOut.Height + dyBorder + dxBorder
        ordLast(1) = chkHidden
        ordLast(2) = chkFocus
        ordLast(3) = chkMin
        ordLast(4) = chkMax
        chkHidden = vbChecked
    Else
        Height = dy
        chkHidden = ordLast(1)
        chkFocus = ordLast(2)
        chkMin = ordLast(3)
        chkMax = ordLast(4)
    End If
End Sub

Private Sub optRun_Click(Index As Integer)
    sRunMode = optRun(Index).Caption
    Select Case sRunMode
    Case "Shell"
        chkWait.Enabled = True
        Width = fmChar.Left + dxBorder
        Height = fmChar.Top + fmChar.Height + dyBorder
        chkWait_Click
    Case "Executive"
        chkWait.Enabled = True
        Width = dx
        Height = dy
        chkPipe_Click
        chkWait_Click
    Case "VBShellExecute"
        Width = fmChar.Left + dxBorder
        Height = fmChar.Top + fmChar.Height + dyBorder
        chkWait.Enabled = False
        chkWait.Value = vbUnchecked
        chkDead.Enabled = False
        chkDead.Value = vbUnchecked
    End Select
End Sub

Private Sub tmrWait_Timer()
    Static c As Long
    If chkWait = False And IsRunning(idProg) Then
        lblStatus.Caption = "Working: " & c
        c = c + 1
    Else
        lblStatus.Caption = "Program returned"
        tmrWait.Enabled = False
        c = 0
    End If
End Sub

Private Sub txtCmd_Change()
    If fCmdFocus Then sCmdLine = txtCmd.Text
    txtProgram.Text = sProgram
    txtArgs.Text = sArgs
End Sub

Private Sub txtArgs_Change()
    sArgs = txtArgs.Text
    txtCmd.Text = sCmdLine
End Sub

Private Sub txtPipeIn_GotFocus()
    cmdRun.Default = False
End Sub

Private Sub txtPipeIn_LostFocus()
    cmdRun.Default = True
End Sub

Private Sub txtProgram_Change()
    sProgram = txtProgram.Text
    txtCmd.Text = sCmdLine
End Sub

Private Sub txtProgram_GotFocus()
With txtProgram
    .SelStart = 0
    .SelLength = Len(.Text)
    fCmdFocus = False
    If .Text <> sEmpty Then lblExeType.Caption = ExeTypeStr(SearchForExe(.Text))
End With
End Sub

Private Sub txtCmd_GotFocus()
With txtCmd
    .SelStart = 0
    .SelLength = Len(.Text)
    fCmdFocus = True
    If .Text <> sEmpty Then lblExeType.Caption = ExeTypeStr(SearchForExe(.Text))
End With
End Sub

Private Sub txtArgs_GotFocus()
With txtArgs
    .SelStart = 0
    .SelLength = Len(.Text)
    fCmdFocus = False
    If .Text <> sEmpty Then
        lblExeType.Caption = ExeTypeStr(SearchForExe(.Text))
    End If
End With
End Sub

Private Sub InitControls()
    txtProgram.Text = GetFileBase(Environ$("COMSPEC"))
    txtArgs.Text = sEmpty
    chkHidden.Value = vbUnchecked
    chkFocus.Value = vbChecked
    chkMin.Value = vbUnchecked
    chkMax.Value = vbUnchecked
    chkWait.Value = vbChecked
End Sub

' Convert Process properties to constants expected by Shell
Private Function GetDisplay() As Long
    If chkHidden.Value = vbChecked Then
        GetDisplay = vbHide
    ElseIf chkFocus.Value = vbChecked Then         ' (with focus)
        If chkMin.Value = vbChecked Then
            GetDisplay = vbMinimizedFocus
        ElseIf chkMax.Value = vbChecked Then
            GetDisplay = vbMaximizedFocus
        Else
            GetDisplay = vbNormalFocus
        End If
    Else                        ' Disabled (without focus)
        If chkMin.Value = vbChecked Then
            GetDisplay = vbNormalNoFocus
        ElseIf chkMax.Value = vbChecked Then
            ' No such thing as maximized without focus
            GetDisplay = vbMaximizedFocus
        Else
            GetDisplay = vbNormalNoFocus
        End If
    End If
End Function

Private Property Get sCmdLine() As String
    If sProgram = sEmpty Then
        sCmdLine = sArgs
    ElseIf sArgs = sEmpty Then
        sCmdLine = sProgram
    Else
        sCmdLine = sProgram & " " & sArgs
    End If
End Property

Private Property Let sCmdLine(sCmdLineA As String)
    sCmdLineA = LTrim$(sCmdLineA)
    Dim i As Integer, iT As Integer
    ' Check for quoted program name
    If Left$(sCmdLineA, 1) = """" Then
        i = InStr(sCmdLineA, """")
        If i Then i = i + 1
    Else
        ' Find first space or tab
        i = InStr(sCmdLineA, " ")
        iT = InStr(sCmdLineA, sTab)
        If iT And i > iT Then i = iT
    End If
    If i = 0 Then
        ' No tab or space, no argument
        sProgram = sCmdLineA
        sArgs = sEmpty
    Else
        ' Split out the parts
        sProgram = Left$(sCmdLineA, i - 1)
        sArgs = Mid$(sCmdLineA, i + 1)
    End If
End Property
