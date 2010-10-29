VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "Get SSLib"
   ClientHeight    =   7995
   ClientLeft      =   180
   ClientTop       =   855
   ClientWidth     =   11400
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "MainV2.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7995
   ScaleWidth      =   11400
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraContent 
      Height          =   6975
      Left            =   240
      TabIndex        =   3
      Top             =   -60
      Width           =   11385
      Begin VB.Frame fraTabs 
         BorderStyle     =   0  'None
         Height          =   4500
         Index           =   2
         Left            =   6375
         TabIndex        =   13
         Top             =   390
         Visible         =   0   'False
         Width           =   8265
         Begin VB.Frame fraAddition 
            BorderStyle     =   0  'None
            Height          =   1140
            Left            =   315
            TabIndex        =   15
            Top             =   3600
            Width           =   7830
            Begin VB.CommandButton cmdOk 
               Caption         =   "确认"
               Height          =   345
               Left            =   5385
               TabIndex        =   21
               Top             =   585
               Width           =   1095
            End
            Begin VB.CommandButton cmdCancel 
               Cancel          =   -1  'True
               Caption         =   "取消"
               Height          =   345
               Left            =   6735
               TabIndex        =   20
               Top             =   585
               Width           =   1095
            End
            Begin VB.CommandButton cmdReset 
               Caption         =   "重置"
               Height          =   345
               Left            =   3975
               TabIndex        =   19
               Top             =   585
               Width           =   1095
            End
            Begin VB.CheckBox chkDetectClipboard 
               Caption         =   "自动检测剪贴板"
               Height          =   345
               Left            =   1980
               TabIndex        =   18
               Tag             =   "NoReseting"
               Top             =   105
               Width           =   1785
            End
            Begin VB.CommandButton cmdCheckClipboard 
               Caption         =   "检测剪贴板"
               Height          =   360
               Left            =   375
               TabIndex        =   17
               Top             =   570
               Width           =   2895
            End
            Begin VB.CheckBox chkStartDownload 
               Caption         =   "立即开始下载"
               Height          =   345
               Left            =   15
               TabIndex        =   16
               Tag             =   "NoReseting"
               Top             =   105
               Width           =   1740
            End
         End
         Begin SSLibTaskman.KeyValueEditor BookEditor 
            Height          =   4320
            Left            =   1335
            TabIndex        =   14
            Top             =   195
            Width           =   8190
            _ExtentX        =   14446
            _ExtentY        =   7620
            Appearance      =   0
            BorderStyle     =   1
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TwoColumnMode   =   0   'False
         End
      End
      Begin VB.Frame fraTabs 
         BorderStyle     =   0  'None
         Height          =   5850
         Index           =   1
         Left            =   735
         TabIndex        =   5
         Top             =   705
         Width           =   9015
         Begin VB.Frame FraButtons 
            BorderStyle     =   0  'None
            Height          =   3060
            Left            =   4035
            TabIndex        =   22
            Top             =   225
            Width           =   5535
            Begin VB.CommandButton cmd 
               Caption         =   "添加"
               Height          =   360
               Index           =   0
               Left            =   795
               TabIndex        =   28
               Top             =   555
               Width           =   1125
            End
            Begin VB.CommandButton cmd 
               Caption         =   "删除"
               Height          =   360
               Index           =   1
               Left            =   1530
               TabIndex        =   27
               Top             =   1545
               Width           =   1125
            End
            Begin VB.CommandButton cmd 
               Caption         =   "开始"
               Height          =   360
               Index           =   2
               Left            =   2580
               TabIndex        =   26
               Top             =   330
               Width           =   1125
            End
            Begin VB.CommandButton cmd 
               Caption         =   "停止"
               Height          =   360
               Index           =   3
               Left            =   3240
               TabIndex        =   25
               Top             =   885
               Width           =   1125
            End
            Begin VB.CommandButton cmd 
               Caption         =   "重新开始"
               Height          =   360
               Index           =   4
               Left            =   3645
               TabIndex        =   24
               Top             =   1935
               Width           =   1125
            End
            Begin VB.CommandButton cmd 
               Caption         =   "属性"
               Height          =   360
               Index           =   5
               Left            =   4215
               TabIndex        =   23
               Top             =   2505
               Width           =   1125
            End
         End
         Begin VB.Frame FrameTaskInfo 
            BorderStyle     =   0  'None
            Caption         =   "Task Info"
            Height          =   3120
            Left            =   240
            TabIndex        =   7
            Top             =   2370
            Width           =   4000
            Begin MSComctlLib.ListView lstTaskInfo 
               Height          =   855
               Left            =   60
               TabIndex        =   8
               Top             =   1425
               Width           =   3000
               _ExtentX        =   5292
               _ExtentY        =   1508
               View            =   3
               LabelEdit       =   1
               LabelWrap       =   0   'False
               HideSelection   =   0   'False
               FullRowSelect   =   -1  'True
               GridLines       =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BorderStyle     =   1
               Appearance      =   0
               NumItems        =   4
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Object.Width           =   452
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   1
                  Text            =   "正在下载"
                  Object.Width           =   2646
               EndProperty
               BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   2
                  Text            =   "进度"
                  Object.Width           =   2822
               EndProperty
               BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   3
                  Text            =   "状态"
                  Object.Width           =   5080
               EndProperty
            End
            Begin VB.Label lblTaskInfo 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "URL"
               ForeColor       =   &H00FF0000&
               Height          =   195
               Index           =   3
               Left            =   160
               TabIndex        =   12
               Top             =   1125
               UseMnemonic     =   0   'False
               Width           =   285
            End
            Begin VB.Label lblTaskInfo 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "下载位置"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   195
               Index           =   2
               Left            =   160
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   11
               ToolTipText     =   "点击打开文件夹"
               Top             =   840
               UseMnemonic     =   0   'False
               Width           =   720
            End
            Begin VB.Label lblTaskInfo 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "任务名"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   195
               Index           =   1
               Left            =   160
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   10
               ToolTipText     =   "用Pdg阅读程序打开"
               Top             =   525
               UseMnemonic     =   0   'False
               Width           =   540
            End
            Begin VB.Label lblTaskInfo 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "ID"
               ForeColor       =   &H00FF0000&
               Height          =   195
               Index           =   0
               Left            =   160
               TabIndex        =   9
               Top             =   240
               UseMnemonic     =   0   'False
               Width           =   165
            End
         End
         Begin MSComctlLib.ListView LstTasks 
            Height          =   1605
            Left            =   60
            TabIndex        =   6
            Top             =   630
            Width           =   4005
            _ExtentX        =   7064
            _ExtentY        =   2831
            View            =   3
            LabelEdit       =   1
            MultiSelect     =   -1  'True
            LabelWrap       =   0   'False
            HideSelection   =   0   'False
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            Icons           =   "ImgListTask"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            NumItems        =   7
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "ID"
               Object.Width           =   903
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "状态"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "SS号"
               Object.Width           =   1806
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Key             =   "title"
               Text            =   "书名"
               Object.Width           =   5644
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "作者"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Key             =   "pages"
               Text            =   "已下载/页数"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "错误"
               Object.Width           =   1411
            EndProperty
         End
         Begin VB.Line lineOne 
            BorderColor     =   &H80000015&
            BorderStyle     =   6  'Inside Solid
            X1              =   0
            X2              =   0
            Y1              =   0
            Y2              =   4785
         End
      End
      Begin MSComctlLib.TabStrip TabContent 
         Height          =   330
         Left            =   60
         TabIndex        =   4
         Top             =   195
         Width           =   8820
         _ExtentX        =   15558
         _ExtentY        =   582
         Style           =   2
         Separators      =   -1  'True
         TabMinWidth     =   1806
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   3
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "下载"
               Key             =   "main"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "任务属性"
               Key             =   "taskproperty"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "添加任务"
               Key             =   "task"
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame fraStatus 
      BorderStyle     =   0  'None
      Height          =   420
      Left            =   15
      TabIndex        =   0
      Top             =   7410
      Width           =   5415
      Begin MSComctlLib.ProgressBar pbDownload 
         Height          =   360
         Left            =   3390
         TabIndex        =   1
         Top             =   30
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   105
         TabIndex        =   2
         Top             =   105
         Width           =   45
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "文件(&F)"
      Begin VB.Menu mnuSave 
         Caption         =   "保存(&S)"
      End
      Begin VB.Menu mnuImport 
         Caption         =   "导入(&I)"
      End
      Begin VB.Menu mnuImportDir 
         Caption         =   "导入文件夹(&D)"
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnuExit 
         Caption         =   "退出程序(&X)"
      End
   End
   Begin VB.Menu mnuTask 
      Caption         =   "任务(&T)"
      Begin VB.Menu mnuSaveAs 
         Caption         =   "保存列表为...(&A)"
      End
      Begin VB.Menu mnuImportList 
         Caption         =   "导入列表...(&I)"
      End
      Begin VB.Menu mnuTaskClear 
         Caption         =   "清除列表(&C)"
      End
   End
   Begin VB.Menu mnuPreference 
      Caption         =   "设置(&P)"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Implements ITaskNotify

Private mTasks As Collection

Private mTaskCount As Long

Private mLastID As Long
Private Const cstMaxRuning As Long = 1
Private mMaxRuning As Long
Private mCurrentRuning As Long
Private mConfigFile As String
Private fDoUpdating As Boolean
Private fTaskStatusChanged As Boolean
Private m_StrPdgProg As String
Private m_StrFolderProg As String

Private Const CSTUpdateStatusInterval As Single = 1#
'Private Const CSTTaskDirectoryName As String = "Tasks"
Private Const CSTTaskConfigFilename As String = "GetSSLib.ini"
Private Const cstUseFrmProgress As Boolean = False
Private Const cstTaskListFilter As String = "Tasks list (*.lst *.ini)|*.lst;*.ini|All (*.*)|*.*"
Private Const cstTaskListExt As String = "lst"
Private mDefaultCaption As String
Private WithEvents Timer As CTimer
Attribute Timer.VB_VarHelpID = -1
Private Const cst_Timer_Interval As Long = 1000
Private mUnloading As Boolean

Private mTask As CTask

Private cboSavedIn As Control

'Private mSavedInIndex As Long





Private Sub cmdAdd_Click()

frmTask.EditTaskMode = False
frmTask.Init Nothing
frmTask.UpdateFromClipboard
frmTask.Show 0, Me

End Sub

Private Function ListItemToTask(ByRef vListItem As ListItem) As CTask
    On Error Resume Next
    Dim sKey As String
    sKey = vListItem.Key
    Set ListItemToTask = mTasks(vListItem.Key)
    
End Function

Private Function TaskToTaskListItem(ByRef vtask As CTask) As ListItem
    On Error Resume Next
    Set TaskToTaskListItem = LstTasks.ListItems.Item(vtask.taskId)
End Function

Private Function TaskToInfoListItem(ByRef vtask As CTask) As ListItem
    
    On Error Resume Next
    Set TaskToInfoListItem = lstTaskInfo.ListItems.Item(vtask.taskId)
End Function

Private Sub cmdEdit_Click()
    Dim selectListItem As ListItem
    Set selectListItem = LstTasks.SelectedItem
    If selectListItem Is Nothing Then
        MsgBox "No tasks selected.", vbInformation
        Exit Sub
    End If
        
    Dim Task As CTask
    Set Task = ListItemToTask(selectListItem)
    If Task Is Nothing Then Exit Sub
                                
    frmTask.Show 0, Me

    If Task.Status = STS_START Then
        frmTask.Init Task, True
    Else
        frmTask.Init Task, False
    End If
    frmTask.EditTaskMode = True
    'frmTask.chkDetectClipboard = False
    
    'UpdateTaskItem Task, , True
    'frmTask.EditTaskMode = False
    'selectListItem.Text = task.Title & "_" & task.SSID
    
End Sub


Public Sub ActionOnTaskItem(ByVal vAction As String)
On Error Resume Next
    
    
    
    If LstTasks.SelectedItem Is Nothing Then
        MsgBox "No tasks selected.", vbInformation, vAction
        Exit Sub
    End If
    
    Timer.Interval = 0 ' False
    vAction = UCase$(vAction)
    Dim selectListItem As ListItem
    Dim selected() As ListItem
    Dim Count As Long
    For Each selectListItem In LstTasks.ListItems
        If selectListItem.selected Then
            Count = Count + 1
            ReDim Preserve selected(1 To Count)
            Set selected(Count) = selectListItem
        End If
    Next
    
    Dim fDeleteFolder As Boolean
    If vAction = "REMOVE" Then
        
    End If
    
    Dim i As Long
    Dim Task As CTask
    For i = 1 To Count
            Set selectListItem = selected(i)
            Set Task = ListItemToTask(selectListItem)
            If Not Task Is Nothing Then
                Select Case vAction
                    Case "REMOVE"
                        LstTasks.ListItems.Remove (selectListItem.Key)
                        If (MsgBox("删除任务文件夹？" & vbCrLf & Task.Directory, vbQuestion + vbYesNo) = vbYes) Then
                            DeleteFolder Task.Directory
                        End If
                        mTasks.Remove Task.taskId
                    Case "START"
                        If Not Task.Status = STS_START Then
                            Task.Status = STS_PENDING
                            UpdateTaskItem Task, selectListItem, True
                            Task.Changed = True
                        End If
                    Case "STOP"
                        If Not Task.Status = STS_PAUSE Then
                            Task.StopNow = True
                            Task.Status = STS_PAUSE
                            UpdateTaskItem Task, selectListItem, True
                            Task.Changed = True
                        End If
                    Case "RESTART"
                        If Not Task.Status = STS_START Then
                            Task.Restart
                            Task.Status = STS_PENDING
                            UpdateTaskItem Task, selectListItem, True
                            Task.Changed = True
                        End If
                End Select
            End If
    Next
    If vAction = "START" Or vAction = "RESTART" Or vAction = "STOP" Then ProcessQueue
    Timer.Interval = cst_Timer_Interval
    
End Sub


Private Sub cmdRemove_Click()
    ActionOnTaskItem "Remove"
    
End Sub

Private Sub cmdRestart_Click()
    ActionOnTaskItem "Restart"
End Sub

Private Sub cmdSaveConfig_Click()
    SaveConfig
End Sub

Private Sub cmdStart_Click()
ActionOnTaskItem "Start"
End Sub

Private Sub cmdStop_Click()
   ActionOnTaskItem "Stop"
End Sub





Private Sub cmd_Click(Index As Integer)
    Select Case Index
    
        Case 0
            cmdAdd_Click
        Case 1
            cmdRemove_Click
        Case 2
            cmdStart_Click
        Case 3
            cmdStop_Click
        Case 4
            cmdRestart_Click
        Case 5
            cmdEdit_Click
        
    End Select
End Sub

Private Sub cmdCheckClipboard_Click()

    UpdateFromClipboard

End Sub

Private Sub UpdateFromClipboard(Optional vText As String)
Dim vData() As String ' As SSLIB_BOOKINFO
vData() = SSLIB_ParseInfoText(vText)
If SafeUBound(vData()) < CST_SSLIB_FIELDS_LBound Then Exit Sub
On Error Resume Next
Dim i As Long
For i = CST_SSLIB_FIELDS_LBound To CST_SSLIB_FIELDS_UBound
    If vData(i) <> "" Then BookEditor.SetField N_(i), vData(i)
Next
End Sub
Private Sub Form_Load()
        
    Set Timer = New CTimer
    Timer.Interval = 0


    mDefaultCaption = App.ProductName & " " & App.Major & "." & App.Minor & " Build " & App.Revision
    Me.Caption = mDefaultCaption
    
    'Timer.Enabled = False
    mMaxRuning = cstMaxRuning
    mCurrentRuning = 0
    fDoUpdating = True
    Set mTasks = New Collection
    



    
    SSLIB_Init
    BookEditor.TwoColumnMode = True
    
    Dim i As Long
    For i = CST_SSLIB_FIELDS_LBound To CST_SSLIB_FIELDS_TASKS_UBOUND
        BookEditor.AddItem SSLIB_ChnFieldName(i), VCT_NORMAL, "", False
        'Map(i - CST_SSLIB_FIELDS_LBound, 0) = SSLIB_ChnFieldName(i)
    Next
    'Process Map()

    BookEditor.SetFieldStyle SSLIB_ChnFieldName(SSF_SAVEDIN), VCT_Combox + VCT_DIR, False
    BookEditor.SetFieldStyle SSLIB_ChnFieldName(SSF_HEADER), VCT_MultiLine, True
    
    Set cboSavedIn = BookEditor.GetValueControlByName(SSLIB_ChnFieldName(SSF_SAVEDIN))
    
    If cboSavedIn Is Nothing Then MsgBox "Error when initlizing KeyValueEditor", vbCritical: End
    
    
    mConfigFile = App.Path & "\" & App.EXEName & ".ini"
    LoadConfig
    
    
    'mSavedInIndex = BookEditor.SearchIndex(SSLIB_ChnFieldName(SSF_SAVEDIN))
    


    'cboValue(mSavedInIndex).Tag = CST_FORM_FLAGS_NORESETING
    'chkStartDownload.Tag = CST_FORM_FLAGS_NORESETING
    'chkDetectClipboard.Tag = CST_FORM_FLAGS_NORESETING
    'Timer.Interval = 0 ' False
'    Dim configHnd As CLiNInI
'    Set configHnd = New CLiNInI
'    With configHnd
'        .Source = App.Path & "\" & configFile
'        ComboxItemsFromString BookEditor.GetValueControl(mSavedInIndex), .GetSetting("SavedIn", "Path")
'        FormStateFromString Me, .GetSetting("Form", "State")
'    End With
'    Set configHnd = Nothing
    
    
    
    
    If LstTasks.ListItems.Count > 0 Then
        LstTasks.ListItems(1).selected = True
        UpdateTaskInfo ListItemToTask(LstTasks.ListItems(1))
    End If
    
    'Timer.Enabled = True
    
    Timer.Interval = cst_Timer_Interval
End Sub

Private Sub Form_Resize()

On Error Resume Next

    If frmMain.ScaleHeight < 1 Then Exit Sub
    'frmMain.Enabled = False
    
    Dim cstLeft As Single
    Dim cstTop As Single
    cstLeft = 60
    cstTop = 60
    
        
    fraStatus.Move cstLeft, frmMain.ScaleHeight - fraStatus.Height, frmMain.ScaleWidth - 2 * cstLeft
    With fraStatus
        pbDownload.Move fraStatus.Width - pbDownload.Width
        'lblStatus.Move lblStatus.Left, , fraStatus.Width - pbDownload.Width - cstLeft
    End With
    
        
    
    
    
    fraContent.Move -cstLeft, 0, frmMain.ScaleWidth + 2 * cstLeft, fraStatus.Top
    
    Dim i As Long
    
    With fraContent
        'TabContent.Move cstLeft, 160, .Width - 2 * cstLeft
        For i = 1 To fraTabs.Count
            If (fraTabs(i).Visible = True) Then
                fraTabs(i).Move cstLeft, TabContent.Top + TabContent.Height + cstTop, .Width - 2 * cstLeft, .Height - TabContent.Top - TabContent.Height - 2 * cstTop
                'fraTabs(i).Move cstLeft, cstTop, .Width - 2 * cstLeft, .Height - 2 * cstTop
            End If
        Next
    
    End With
           
    If fraTabs(1).Visible Then
        
        FraButtons.Top = cstTop
        FraButtons.Left = cstLeft
        FraButtons.Width = cmd(0).Width + 2 * cstLeft
        FraButtons.Height = fraTabs(1).Height - FraButtons.Top - cstTop
    
        
        cmd(0).Move cstLeft, cstTop
        For i = 1 To cmd.UBound
            cmd(i).Move cmd(i - 1).Left, cmd(i - 1).Top + cmd(i - 1).Height + 2 * cstTop
        Next
        
        
        FrameTaskInfo.Left = FraButtons.Left + FraButtons.Width + 2 * cstLeft
        FrameTaskInfo.Width = fraTabs(1).Width - FrameTaskInfo.Left - 2 * cstLeft
        

        
    '    Dim lastLabel As Label
    '    Set lastLabel = lblTaskInfo(lblTaskInfo.count - 1)
        'lblTaskInfo(0).Width = FrameTaskInfo.Width - 2 * lblTaskInfo(0).Left
        'lbltaskinfo(0).Top =
        For i = 1 To lblTaskInfo.UBound
            'lbltaskinfo(i).Move
            lblTaskInfo(i).Top = lblTaskInfo(i - 1).Top + lblTaskInfo(i - 1).Height + 120
         '   lblTaskInfo(i).Width = lblTaskInfo(0).Width
        Next
        i = lblTaskInfo.UBound
        lstTaskInfo.Move lstTaskInfo.Left, lblTaskInfo(i).Top + lblTaskInfo(i).Height + 120, _
            FrameTaskInfo.Width - lstTaskInfo.Left - 2 * cstLeft ',   FrameTaskInfo.Height -lblTaskInfo(i).Top - lblTaskInfo(i).Height - 180
            

        FrameTaskInfo.Height = lstTaskInfo.Top + lstTaskInfo.Height + cstTop
        FrameTaskInfo.Top = fraTabs(1).Height - FrameTaskInfo.Height
        
            
            
            LstTasks.Left = FrameTaskInfo.Left
            LstTasks.Top = cstTop ' cmdAdd.Top + cmdAdd.Height + 120
            LstTasks.Width = fraTabs(1).Width - LstTasks.Left - 2 * cstLeft
            LstTasks.Height = FrameTaskInfo.Top - LstTasks.Top
            
            lineOne.X1 = LstTasks.Left - 1.5 * cstLeft '/ 2
            lineOne.X2 = lineOne.X1
            lineOne.Y1 = FraButtons.Top '+ 2 * cstTop
            lineOne.Y2 = lineOne.Y1 + FraButtons.Height ' - 2 * cstTop
        
    End If
       
    If fraTabs(2).Visible Then
        fraAddition.Top = fraTabs(2).Height - 120 - fraAddition.Height ' txtBox.Top + txtBox.Height + 120
        fraAddition.Left = fraTabs(2).Width - 120 - fraAddition.Width


        BookEditor.Move 120, 120, fraTabs(2).Width - 240, fraAddition.Top - 120
    End If
       
    
'    lsttasks.Top = cmdAdd.Top + cmdAdd.Height + 120
'    lsttasks.Width = frmMain.ScaleWidth - 2 * lsttasks.Left
'    lsttasks.Height = frmMain.ScaleHeight - StatusBar.Height - lsttasks.Top
    
'    StatusBar.Top = frmMain.ScaleHeight - StatusBar.Height
    
    
        
   
    
    'With lstTaskInfo
        
        
    
    'End With
'
'    For i = 1 To lblTaskInfo.count - 1
'        lblTaskInfo(i).Top = lblTaskInfo(i - 1).Top
'    Next
    
    
    'frmMain.Enabled = True
End Sub

Public Function AddTaskItem(ByRef vtask As CTask, Optional vPending As Boolean = False, Optional vSave As Boolean = False) As ListItem
    If vPending Then vtask.Status = STS_PENDING
    vtask.taskId = NewTaskID()
    vtask.Init Me
       
    mTasks.Add vtask, vtask.taskId
    Dim taskListItem As ListItem
    Set taskListItem = LstTasks.ListItems.Add(, vtask.taskId)
'    If vSave Then vtask.AutoSave
    UpdateTaskItem vtask, taskListItem, True
    
    Set AddTaskItem = taskListItem
    
End Function

Private Function AddTask(vBookInfoArray() As String, Optional vPending As Boolean) As CTask
        Dim newTask As CTask
        Set newTask = New CTask
       ' Dim taskId As String
        'taskId = NewTaskID()
        newTask.bookInfo.LoadFromArray vBookInfoArray
        AddTaskItem newTask, vPending, False
                
        Set AddTask = newTask
        ', "TASK" & index, newTask.newTask.Title & "_" & newTask.SSID
        'LstTasks.ListItems.Add "TASK" & index, tvwChild, "INFO" & index, "Paused"
        'UpdateTaskInfo newTask
End Function

Public Sub CallBack_EditTask(ByRef vtask As CTask)
    UpdateTaskItem vtask, , True
    vtask.Changed = True
    vtask.AutoSave
End Sub
Public Sub CallBack_AddTask(vBookInfoArray() As String, Optional vPending As Boolean)

        Dim vtask As CTask
        Set vtask = AddTask(vBookInfoArray, vPending)
        If Not vtask Is Nothing Then vtask.Changed = True
        vtask.Changed = True
        vtask.AutoSave

End Sub
'
'Public Sub CallBack_AddTask(bookInfo As CBookInfo)
'    Dim vtask As CTask
'    Set vtask = New CTask
'    Set vtask.bookInfo = bookInfo
'End Sub

Private Function TextOfStatus(ByRef Status As SSLIBTaskStatus) As String
    Dim text As String
Select Case Status
    Case SSLIBTaskStatus.STS_COMPLETE
        text = "完成"
    Case SSLIBTaskStatus.STS_PAUSE
        text = "停止"
    Case SSLIBTaskStatus.STS_PENDING
        text = "排队中..."
    Case SSLIBTaskStatus.STS_START
        text = "正在下载..."
    Case Else
        text = "停止"
End Select

TextOfStatus = text
End Function

Private Function TaskItemFrom(ByRef vtask As CTask) As ListItem
    On Error Resume Next
    If vtask Is Nothing Then Exit Function
    Dim Item As ListItem
    Set Item = LstTasks.ListItems(vtask.taskId)
    Set TaskItemFrom = Item
End Function

Private Sub UpdateTaskItem(Optional ByRef vtask As CTask, Optional ByRef vItem As ListItem, Optional vForce As Boolean = False)
    On Error Resume Next
    'Dim i As Integer
    'Me.Enabled = False
    DoEvents
    
    If Not vtask Is Nothing Then
            If vtask.FilesCount > vtask.PagesCount Then
                pbDownload.value = pbDownload.Max
            Else
                pbDownload.Max = vtask.PagesCount
                pbDownload.value = vtask.FilesCount
            End If
    End If
    
    
    Static LastUpdateTime As Single
    Dim CurrentTime As Single
    CurrentTime = DateTime.Timer
    If Not vForce Then
        If CurrentTime - LastUpdateTime < CSTUpdateStatusInterval Then Exit Sub
    End If
    LastUpdateTime = CurrentTime
    
    Dim pBookInfo As CBookInfo
    
    If Not vtask Is Nothing Then
        Set pBookInfo = vtask.bookInfo
        If vItem Is Nothing Then Set vItem = TaskItemFrom(vtask)
        If vItem Is Nothing Then Exit Sub
        If vItem.selected = True Then
            
            lblTaskInfo(0).Caption = "ID: " & vtask.taskId & " " & pBookInfo(SSF_SSID)
            lblTaskInfo(1).Caption = "任务名: " & pBookInfo(SSF_Title) & " " & pBookInfo(SSF_Author)
            lblTaskInfo(1).Tag = vtask.Directory
            
            lblTaskInfo(2).Tag = lblTaskInfo(1).Tag
            lblTaskInfo(2).Caption = "下载位置: " & lblTaskInfo(2).Tag
            
            
            lblTaskInfo(3).Caption = "URL: " & pBookInfo(SSF_URL)
            '            lblOpenPdg.Tag = lblTaskInfo(2).Tag
            
        End If
        
        vItem.text = vItem.Index
            vItem.SubItems(1) = TextOfStatus(vtask.Status)
            vItem.SubItems(2) = pBookInfo(SSF_SSID)
            vItem.SubItems(3) = pBookInfo(SSF_Title)
            vItem.SubItems(4) = pBookInfo(SSF_Author)
            vItem.SubItems(5) = vtask.FilesCount & "/" & vtask.PagesCount
            

            
            
            'pbDownload.Min = 0
            
            
            'vItem.SubItems(6) = vTask.FilesCount
            vItem.SubItems(6) = vtask.Downloader.ErrorsCount
        'Me.Enabled = True
            'vItem.text = GetFileName(vTask.Directory)
            
            UpdateTaskInfo vtask, vItem
        Exit Sub
    End If
    
    For Each vtask In mTasks
        Set vItem = TaskItemFrom(vtask)
        If Not vItem Is Nothing Then
            Set pBookInfo = vtask.bookInfo
            vItem.text = vItem.Index
            vItem.SubItems(1) = TextOfStatus(vtask.Status)
            vItem.SubItems(2) = pBookInfo(SSF_SSID)
            vItem.SubItems(3) = pBookInfo(SSF_Title)
            vItem.SubItems(4) = pBookInfo(SSF_Author)
            vItem.SubItems(5) = vtask.FilesCount & "/" & vtask.PagesCount
            'vItem.SubItems(6) = vTask.FilesCount
            vItem.SubItems(6) = vtask.Downloader.ErrorsCount
        End If
    Next
    'Me.Enabled = True
End Sub
Private Sub UpdateTaskInfo(ByRef vtask As CTask, Optional ByRef vItem As ListItem)

'Exit Sub

'Static timeLastUpdate As Single
'Dim CurrentTime As Single
'CurrentTime = DateTime.Timer
'If CurrentTime - timeLastUpdate > 0.8 Then
'    UpdateTaskItem vTask, vItem
'End If
'timeLastUpdate = CurrentTime

DoEvents

If vtask Is Nothing Then Exit Sub
If vItem Is Nothing Then Set vItem = TaskItemFrom(vtask)
If vItem Is Nothing Then Exit Sub
If vItem.selected = False Then Exit Sub

Dim i As Long
Dim c As Long
c = vtask.ConnectionsCount

Me.Enabled = False
lstTaskInfo.ListItems.Clear

For i = 1 To c

Dim Progress As CDownloadProgress
Set Progress = vtask.Connections(i)
If Progress Is Nothing Then GoTo ContinueFor
'If Not fDoUpdating Then Exit Sub
    
Dim itemTask As ListItem
If lstTaskInfo.ListItems.Count < i Then
    Set itemTask = lstTaskInfo.ListItems.Add()
Else
    Set itemTask = lstTaskInfo.ListItems.Item(i)
    itemTask.text = ""
    itemTask.ListSubItems.Clear
End If
If itemTask Is Nothing Then GoTo ContinueFor

itemTask.text = i
itemTask.SubItems(1) = GetFileName(Progress.saveAs) ' vTask.CurrentFile
If Progress.CurrentBytes > 0 Then
    itemTask.SubItems(2) = "[" & CStr(Progress.CurrentBytes) & "/" & CStr(Progress.TotalBytes) & "]"
Else
    itemTask.SubItems(2) = ""
End If
 itemTask.SubItems(3) = Progress.TextInfo
'Text = Text & " FileCount:" & CStr(vTask.FilesCount) & " LastFile:" & vTask.CurrentFile
 'itemTask.SubItems(4) = Progress.URL
ContinueFor:
Next



Me.Enabled = True

DoEvents
'ListItem.Text = Text


End Sub


Private Function SelectTaskListItem(ByRef taskId As String) As ListItem
On Error Resume Next
Set SelectTaskListItem = LstTasks.ListItems.Item(taskId)
End Function

Private Function SelectInfoListItem(ByRef taskId As String) As ListItem
On Error Resume Next
Set SelectInfoListItem = lstTaskInfo.ListItems.Item(taskId)
End Function



Private Sub Form_Terminate()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'If frmTask Is Nothing Then MsgBox "Nothing"
    'Set frmTask = Nothing
    
    'Unload frmTask
    
    'Exit Sub
    'MsgBox "1"
            
    Dim Task As CTask
    SetStatusText "退出程序..."
    If mUnloading = True Then
        For Each Task In mTasks
            If Task.Status = STS_START Then
                'Task.Status = STS_PENDING
                SetStatusText "正在停止任务：[" & Task.taskId & "]" & Task.bookInfo(SSF_Title) & "..."
                Cancel = True
                Exit Sub
            End If
        Next
    Else
        For Each Task In mTasks
            If Task.Status = STS_START Then Task.Status = STS_PENDING
        Next
        SetStatusText "正在保存程序和任务配置..."
        SaveConfig
        For Each Task In mTasks
            Task.StopNow = True
        Next
        mUnloading = True
        Cancel = True
        Exit Sub
    End If
                
    mUnloading = False
    Timer.Interval = 0 ' False
    
    'MsgBox "3"
    'Unload frmTask
    'MsgBox "4"
    Set mTasks = Nothing
    Set Timer = Nothing
    'MsgBox "5"
    End
End Sub





Private Sub SetStatusText(ByRef vText As String)
    lblStatus.Caption = vText
End Sub



Private Sub ITaskNotify_DownloadStatusChange(Task As CTask)
    UpdateTaskInfo Task
End Sub

Private Sub ITaskNotify_TaskComplete(Task As CTask)
    If Task Is Nothing Then
        mCurrentRuning = mCurrentRuning - 1
        
        SetStatusText ""
        pbDownload.value = pbDownload.Min
        
        'Me.Caption = mDefaultCaption
    Else
        SetStatusText ""
        UpdateTaskItem Task, , True
    End If
End Sub

Private Sub ITaskNotify_TaskStatusChange(Task As CTask)
    UpdateTaskItem Task
    
End Sub

Private Function QuoteString(ByRef vString As String) As String
    QuoteString = Chr$(34) & vString & Chr$(34)
End Function

Private Sub lblTaskInfo_Click(Index As Integer)
   
    
    If Index = 1 Then
        Dim pdgDir As String
        Dim pdgFile As String
        Dim pdgOpener As String
        pdgOpener = m_StrPdgProg
        If pdgOpener = "" Then pdgOpener = "start"
        pdgDir = BuildPath(lblTaskInfo(1).Tag)
        
        pdgFile = Dir$(pdgDir & "*.pdg")
        
        If pdgFile <> "" Then
            Shell pdgOpener & " " & QuoteString(pdgDir & pdgFile), vbNormalFocus
            
        Else
            MsgBox QuoteString(pdgDir) & "中不存在任何PDG文件", vbCritical + vbOKOnly
        End If
    ElseIf Index = 2 Then
        Dim prog As String
        Dim arg As String
        arg = QuoteString(lblTaskInfo(2).Tag)
        prog = m_StrFolderProg
        If prog = "" Then prog = "explorer.exe"
        Shell prog & " " & arg, vbNormalFocus

    End If
End Sub

Private Sub LstTasks_DblClick()
    If LstTasks.SelectedItem Is Nothing Then Exit Sub
    cmdEdit_Click
End Sub

Private Sub LstTasks_ItemClick(ByVal Item As MSComctlLib.ListItem)
    On Error Resume Next
    Static lastKey As String
    If Item.Key = lastKey Then Exit Sub
    lastKey = Item.Key
    'If Item Is LstTasks.SelectedItem Then Exit Sub
    UpdateTaskItem mTasks(lastKey), Item, True
    UpdateTaskInfo mTasks(lastKey) 'ListItemToTask(item)
End Sub

Private Sub LstTasks_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 1 And KeyCode = 97 Or KeyCode = 65 Then
        Dim Item As ListItem
        For Each Item In LstTasks.ListItems
            Item.selected = True
        Next
        'LstTasks.se
    End If
End Sub


Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuImport_Click()
    Dim FileName As String
    Dim dlg As CCommonDialogLite
    Set dlg = New CCommonDialogLite
    If dlg.VBGetOpenFileName(FileName, , , , , , cstTaskListFilter, , , , cstTaskListExt, Me.hwnd) Then
        LoadConfigFrom FileName
    End If
End Sub

Private Sub mnuImportDir_Click()
    Static Lastfolder As String
    Dim folder As String
    Dim dlg As CFolderBrowser
    Set dlg = New CFolderBrowser
    dlg.Owner = Me.hwnd
    If Lastfolder <> "" Then dlg.InitDirectory = Lastfolder
    folder = dlg.Browse
    If folder <> "" Then
        Lastfolder = folder
        Dim subdirs() As String
        Dim Count As Long
        Count = subFolders(BuildPath(folder), subdirs())
        Dim i As Long
        Dim fn As String
        If Count > 0 Then
            For i = 1 To Count
                fn = BuildPath(subdirs(i), "GetssLib.ini")
                If FileExists(fn) Then LoadConfigFrom (fn)
            Next
        End If
    End If
    
End Sub

Private Sub mnuImportList_Click()
    mnuImport_Click
End Sub

Private Sub mnuPreference_Click()
    Load frmOptions
    With frmOptions
        .PdgProg = m_StrPdgProg
        .FolderProg = m_StrFolderProg
        .Show 1, Me
        m_StrPdgProg = .PdgProg
        m_StrFolderProg = .FolderProg
    End With
    Unload frmOptions
End Sub

Private Sub mnuSave_Click()
    SaveConfig
End Sub

Private Sub mnuSaveAs_Click()
    Dim FileName As String
    Dim dlg As CCommonDialogLite
    Set dlg = New CCommonDialogLite
    If dlg.VBGetSaveFileName(FileName, , , cstTaskListFilter, , , , cstTaskListExt, Me.hwnd) Then
        SaveConfigTo FileName
    End If
    
End Sub

Private Sub mnuTaskClear_Click()
    Dim Task As CTask
    Dim tasks() As CTask
    Dim Item As ListItem
    Dim Count As Long
    On Error Resume Next
    For Each Task In mTasks
    If Not Task.Status = STS_START Then
        Count = Count + 1
        ReDim Preserve tasks(1 To Count)
        Set tasks(Count) = Task
    End If
    Next
   
    Dim i As Long
    For i = 1 To Count
         If tasks(i).Changed = True Then tasks(i).AutoSave
         LstTasks.ListItems.Remove tasks(i).taskId
         mTasks.Remove tasks(i).taskId
         Set tasks(i) = Nothing
    Next
   
End Sub

Private Sub TabContent_Click()

    On Error Resume Next

    Dim i As Long
    For i = 1 To fraTabs.Count
        fraTabs(i).Visible = False
    Next
    
    Select Case TabContent.SelectedItem.Index
    
        Case 1
            fraTabs(1).Visible = True
        Case 2
            fraTabs(2).Visible = True
            cmdOK.Caption = "确认"
            
            Call EditTask
        Case 3
            fraTabs(2).Visible = True
            cmdOK.Caption = "添加任务"
            
            Call AddNewTask 'Nothing
            
    
    End Select
    
    'fraTabs(TabContent.SelectedItem.Index).Visible = True
    Form_Resize

'    If TabContent.SelectedItem.Index = 2 Then
'        frmTask.SolidMode = True
'        frmTask.Show 0, Me
'        frmTask.Move fraTabs(2).Left + (fraTabs(2).Width - frmTask.Width) / 2, fraTabs(2).Top + (fraTabs(2).Height - frmTask.Height) / 2
'    End If
    
End Sub

Private Sub TabContent_Validate(Cancel As Boolean)
'    Dim i As Long
'    For i = 1 To fraTabs.count
'        fraTabs(i).Visible = False
'    Next
'    fraTabs(TabContent.SelectedItem.Index).Visible = True
End Sub

Private Sub Timer_Timer()

    'UpdateTaskItem
    ProcessQueue
    

End Sub

Private Sub ProcessQueue()
    If mUnloading Then Exit Sub
    If mCurrentRuning >= mMaxRuning Then Exit Sub
    Dim Task As CTask
    For Each Task In mTasks
        If Task.Status = STS_PENDING Then
            If mUnloading Then Exit Sub
            Task.Status = STS_START
            mCurrentRuning = mCurrentRuning + 1
            UpdateTaskItem Task, , True
            SetStatusText QuoteString(Task.bookInfo(SSF_Title)) & " 正在下载..."
            If cstUseFrmProgress Then
                    Dim newDownload As frmProgress
                    Set newDownload = New frmProgress
                    Set newDownload.MainApp = Me
                    Set newDownload.Task = Task
                    newDownload.Show 0, Me
        '            newDownload.StartTask task
                    
                    'Load frmProgress
            Else
                    Task.StartDownload
                    mCurrentRuning = mCurrentRuning - 1
                    Me.Caption = App.ProductName
            End If
            
            'Set newDownload = Nothing
            Exit Sub
        End If
    Next
    

End Sub

Public Function NewTaskID() As String
    Static lastNum As Long
    If lastNum < 1 Then
        lastNum = 1
    Else
        lastNum = lastNum + 1
    End If
    NewTaskID = "SSBOOK" & CStr(lastNum)
End Function

Public Sub SaveConfig()
    SaveConfigTo mConfigFile, True
End Sub


Public Sub SaveConfigTo(ByRef vFilename As String, Optional withAppSetting As Boolean = False)

On Error GoTo ErrorSaveTaskTo

    Dim iniHnd As CLiNInI
    Set iniHnd = New CLiNInI
    On Error Resume Next
    If withAppSetting Then
        iniHnd.SaveSetting "App", "WindowPosition", FormStateToString(Me)
        iniHnd.SaveSetting "App", "PdgProg", m_StrPdgProg
        iniHnd.SaveSetting "App", "FolderProg", m_StrFolderProg
        iniHnd.SaveSetting "App", "SavedIn", ComboxItemsToString(cboSavedIn)
         iniHnd.SaveSetting "App", "DetectClipboard", chkDetectClipboard.value
         iniHnd.SaveSetting "App", "StartDownload", chkStartDownload.value
    End If
    iniHnd.SaveSetting "GetSSLib", "TaskDirectoryCount", CStr(mTasks.Count)
    Dim Task As CTask
    Dim i As Long
    Dim taskconfig As CLiNInI
    Dim taskDir As String
    For Each Task In mTasks
        i = i + 1
        taskDir = Task.Directory
        If Not FolderExists(taskDir) Then xMkdir taskDir
        
        iniHnd.SaveSetting "GetSSLib", "TaskDirectory" & CStr(i), taskDir
        If withAppSetting And Task.Changed = True Then
            Set taskconfig = New CLiNInI
            Task.PersistTo taskconfig, "TaskInfo"
            taskconfig.WriteTo BuildPath(taskDir, CSTTaskConfigFilename)
            Set taskconfig = Nothing
            Task.Changed = False
        End If
    Next
    iniHnd.WriteTo vFilename
    Set iniHnd = Nothing
    
    Exit Sub
ErrorSaveTaskTo:
    MsgBox Err.Description, vbCritical, "SaveConfigTo"
End Sub

'CSEH: ErrAsk
Public Sub LoadConfigFrom(ByRef vFilename As String, Optional withAppSetting As Boolean = False)
        '<EhHeader>
        On Error GoTo LoadConfigFrom_Err
        '</EhHeader>
        Dim Task As CTask
        Dim iniHnd As CLiNInI
        Dim i As Long
        Dim taskDir As String
        Dim taskconfig As CLiNInI
        
100     Set iniHnd = New CLiNInI
105     iniHnd.File = vFilename
    
        If withAppSetting Then
            FormStateFromString Me, iniHnd.GetSetting("App", "WindowPosition")
            m_StrPdgProg = iniHnd.GetSetting("App", "PdgProg")
            m_StrFolderProg = iniHnd.GetSetting("App", "FolderProg")
            ComboxItemsFromString cboSavedIn, iniHnd.GetSetting("App", "SavedIn")
            chkDetectClipboard.value = StringToLong(iniHnd.GetSetting("App", "DetectClipboard"))
            chkStartDownload.value = StringToLong(iniHnd.GetSetting("App", "StartDownload"))
        End If
        Dim iTaskCount As Long
    
110     iTaskCount = StringToLong(iniHnd.GetSetting("GetSSLib", "TaskCount"))

115     For i = 1 To iTaskCount
120         Set Task = New CTask
125         Task.LoadFrom iniHnd, "Task" & CStr(i)
            Task.Changed = False
130         AddTaskItem Task, False
        Next
    
135     iTaskCount = 0
140     iTaskCount = StringToLong(iniHnd.GetSetting("GetSSLib", "TaskDirectoryCount"))
145     For i = 1 To iTaskCount
150         taskDir = iniHnd.GetSetting("GetSSLib", "TaskDirectory" & CStr(i))
155         If FolderExists(taskDir) Then
160             Set Task = New CTask
165             Set taskconfig = New CLiNInI
170             taskconfig.File = BuildPath(taskDir, CSTTaskConfigFilename)
175             Task.LoadFrom taskconfig, "TaskInfo"
180             AddTaskItem Task, False
                Task.Changed = False
185             Set taskconfig = Nothing
190             Set Task = Nothing
            End If
        Next
    


200     If iniHnd.SectionExists("TaskInfo") Then
205         Set Task = New CTask
210         Task.LoadFrom iniHnd, "TaskInfo"
            Task.Changed = False
215         AddTaskItem Task, False
        End If

        '<EhFooter>
        Exit Sub

LoadConfigFrom_Err:
    If MsgBox(Err.Description & vbCrLf & _
               "in GetSSLib.frmMain.LoadConfigFrom " & _
               "at line " & Erl & vbCrLf & _
               "Continue (Click No will terminate GetSSLib)?", _
               vbExclamation + vbYesNo, "Application Error") = vbYes Then
        Resume Next
    Else
        End
    End If

        '</EhFooter>
End Sub

Public Sub LoadConfig()

    LoadConfigFrom mConfigFile, True
    

End Sub


Private Sub Timer_ThatTime()
    'Timer.Interval = 0
    If mUnloading Then
        Unload Me
    Else
        ProcessQueue
    End If
    'Timer.Interval = CST_Timer_Interval
End Sub
