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
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7995
   ScaleWidth      =   11400
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox fraTabs 
      Height          =   495
      Index           =   2
      Left            =   8355
      ScaleHeight     =   435
      ScaleWidth      =   1155
      TabIndex        =   30
      Top             =   6180
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox fraTabs 
      BorderStyle     =   0  'None
      Height          =   5325
      Index           =   1
      Left            =   690
      ScaleHeight     =   5325
      ScaleWidth      =   9810
      TabIndex        =   5
      Top             =   435
      Width           =   9810
      Begin VB.Frame FrameTaskInfo 
         BorderStyle     =   0  'None
         Caption         =   "Task Info"
         Height          =   3120
         Left            =   120
         TabIndex        =   23
         Top             =   1905
         Width           =   4000
         Begin MSComctlLib.ListView lstTaskInfo 
            Height          =   1545
            Left            =   60
            TabIndex        =   24
            Top             =   1425
            Width           =   3000
            _ExtentX        =   5292
            _ExtentY        =   2725
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
               Object.Width           =   6456
            EndProperty
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
            TabIndex        =   28
            Top             =   240
            UseMnemonic     =   0   'False
            Width           =   165
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
            TabIndex        =   27
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
            TabIndex        =   26
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
            Caption         =   "URL"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   3
            Left            =   160
            TabIndex        =   25
            Top             =   1125
            UseMnemonic     =   0   'False
            Width           =   285
         End
      End
      Begin VB.PictureBox FraButtons 
         BorderStyle     =   0  'None
         Height          =   3810
         Left            =   4020
         ScaleHeight     =   3810
         ScaleWidth      =   5685
         TabIndex        =   6
         Top             =   285
         Width           =   5685
         Begin VB.CommandButton cmd 
            Caption         =   "开始运行"
            Height          =   360
            Index           =   19
            Left            =   480
            TabIndex        =   34
            Top             =   2355
            Width           =   1125
         End
         Begin VB.CommandButton cmd 
            Appearance      =   0  'Flat
            Caption         =   "---------------"
            Enabled         =   0   'False
            Height          =   360
            Index           =   18
            Left            =   4140
            TabIndex        =   33
            TabStop         =   0   'False
            Top             =   1875
            Width           =   1125
         End
         Begin VB.CommandButton cmd 
            Caption         =   "打开链接"
            Height          =   360
            Index           =   12
            Left            =   2580
            TabIndex        =   32
            Top             =   1200
            Width           =   1125
         End
         Begin VB.CommandButton cmd 
            Caption         =   "---------------"
            Enabled         =   0   'False
            Height          =   360
            Index           =   3
            Left            =   3870
            TabIndex        =   31
            TabStop         =   0   'False
            Top             =   225
            Width           =   1125
         End
         Begin VB.CommandButton cmd 
            Caption         =   "打开浏览器"
            Height          =   360
            Index           =   17
            Left            =   2775
            TabIndex        =   22
            Top             =   1860
            Width           =   1125
         End
         Begin VB.CommandButton cmd 
            Caption         =   "添加"
            Height          =   360
            Index           =   0
            Left            =   195
            TabIndex        =   21
            Top             =   255
            Width           =   1125
         End
         Begin VB.CommandButton cmd 
            Caption         =   "删除"
            Height          =   360
            Index           =   2
            Left            =   2625
            TabIndex        =   20
            Top             =   225
            Width           =   1125
         End
         Begin VB.CommandButton cmd 
            Caption         =   "开始"
            Height          =   360
            Index           =   4
            Left            =   195
            TabIndex        =   19
            Top             =   675
            Width           =   1125
         End
         Begin VB.CommandButton cmd 
            Caption         =   "停止"
            Height          =   360
            Index           =   6
            Left            =   1770
            TabIndex        =   18
            Top             =   690
            Width           =   1125
         End
         Begin VB.CommandButton cmd 
            Caption         =   "重新下载"
            Height          =   360
            Index           =   8
            Left            =   3210
            TabIndex        =   17
            Top             =   705
            Width           =   1125
         End
         Begin VB.CommandButton cmd 
            Caption         =   "属性"
            Height          =   360
            Index           =   13
            Left            =   3645
            TabIndex        =   16
            Top             =   1140
            Width           =   1125
         End
         Begin VB.CommandButton cmd 
            Caption         =   "全部开始"
            Height          =   360
            Index           =   5
            Left            =   1005
            TabIndex        =   15
            Top             =   690
            Width           =   1125
         End
         Begin VB.CommandButton cmd 
            Caption         =   "全部停止"
            Height          =   360
            Index           =   7
            Left            =   2505
            TabIndex        =   14
            Top             =   675
            Width           =   1125
         End
         Begin VB.CommandButton cmd 
            Caption         =   "登录sslib"
            Height          =   360
            Index           =   16
            Left            =   180
            TabIndex        =   13
            Top             =   1785
            Width           =   1125
         End
         Begin VB.CommandButton cmd 
            Appearance      =   0  'Flat
            Caption         =   "---------------"
            Enabled         =   0   'False
            Height          =   360
            Index           =   9
            Left            =   3990
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   690
            Width           =   1125
         End
         Begin VB.CommandButton cmd 
            Caption         =   "设置Cookie"
            Height          =   360
            Index           =   15
            Left            =   1470
            TabIndex        =   11
            Top             =   1725
            Width           =   1125
         End
         Begin VB.CommandButton cmd 
            Caption         =   "批量添加"
            Enabled         =   0   'False
            Height          =   360
            Index           =   1
            Left            =   1395
            TabIndex        =   10
            Top             =   240
            Width           =   1125
         End
         Begin VB.CommandButton cmd 
            Appearance      =   0  'Flat
            Caption         =   "---------------"
            Enabled         =   0   'False
            Height          =   360
            Index           =   14
            Left            =   4425
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   1305
            Width           =   1125
         End
         Begin VB.CommandButton cmd 
            Caption         =   "下载JPG"
            Height          =   360
            Index           =   11
            Left            =   1470
            TabIndex        =   8
            Top             =   1155
            Width           =   1125
         End
         Begin VB.CommandButton cmd 
            Caption         =   "下载PDG"
            Height          =   360
            Index           =   10
            Left            =   180
            TabIndex        =   7
            Top             =   1125
            Width           =   1125
         End
      End
      Begin MSComctlLib.ListView LstTasks 
         Height          =   1605
         Left            =   -15
         TabIndex        =   29
         Top             =   0
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
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Key             =   "id"
            Text            =   "ID"
            Object.Width           =   903
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Key             =   "status"
            Text            =   "状态"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Key             =   "ssid"
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
            Key             =   "author"
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
            Key             =   "checked"
            Object.Width           =   903
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Key             =   "error"
            Text            =   "格式"
            Object.Width           =   1411
         EndProperty
      End
   End
   Begin VB.Frame fraContent 
      Height          =   6975
      Left            =   -15
      TabIndex        =   3
      Top             =   15
      Width           =   11385
      Begin MSComctlLib.TabStrip TabContent 
         Height          =   330
         Left            =   60
         TabIndex        =   4
         Top             =   195
         Visible         =   0   'False
         Width           =   8820
         _ExtentX        =   15558
         _ExtentY        =   582
         Style           =   2
         Separators      =   -1  'True
         TabMinWidth     =   1806
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   2
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "下载"
               Key             =   "main"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "添加任务"
               Key             =   "task"
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
      Begin VB.Line lineOne 
         X1              =   10680
         X2              =   10680
         Y1              =   465
         Y2              =   5250
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
         Caption         =   "导入任务(&I)"
      End
      Begin VB.Menu mnuImportDir 
         Caption         =   "导入多个任务(&D)"
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnuExit 
         Caption         =   "退出程序(&X)"
      End
   End
   Begin VB.Menu mnuList 
      Caption         =   "列表(&L)"
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
   Begin VB.Menu mnuTask 
      Caption         =   "任务(&T)"
      Begin VB.Menu mnuTaskID 
         Caption         =   ""
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuTaskSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTaskAction 
         Caption         =   "下载"
         Index           =   0
      End
      Begin VB.Menu mnuTaskAction 
         Caption         =   "停止"
         Index           =   1
      End
      Begin VB.Menu mnuTaskAction 
         Caption         =   "重新下载"
         Index           =   2
      End
      Begin VB.Menu mnuTaskAction 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuTaskAction 
         Caption         =   "删除"
         Index           =   4
      End
      Begin VB.Menu mnuTaskAction 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuTaskAction 
         Caption         =   "打开链接"
         Index           =   6
      End
      Begin VB.Menu mnuTaskAction 
         Caption         =   "更新信息"
         Index           =   7
      End
      Begin VB.Menu mnuTaskAction 
         Caption         =   "属性"
         Index           =   8
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
Private Const cstMaxTasksDownloading As Long = 1
Private Const cstTaskDownloadThreads As Long = 3
Private mMaxTasksDownloading As Long
Private mCurrentRuning As Long
Private mConfigFile As String
Private fDoUpdating As Boolean
Private fTaskStatusChanged As Boolean
Private m_StrPdgProg As String
Private m_StrFolderProg As String

Private Const CSTUpdateStatusInterval As Single = 1#
'Private Const CSTTaskDirectoryName As String = "Tasks"
Private Const CSTAppConfigname As String = "GetSSLib.ini"
Private Const CSTTasksFilename As String = "Tasks.lst"
Private Const CSTTaskConfigFilename As String = "GetSSLib.ini"
Private Const CSTTaskIncomeFilename As String = "Incoming.lst"
Private Const CSTTaskIncomeFilename2 As String = "Incoming.tmp"
Private Const cstUseFrmProgress As Boolean = False
Private Const cstTaskListFilter As String = "Tasks list (*.lst *.ini)|*.lst;*.ini|All (*.*)|*.*"
Private Const cstTaskListExt As String = "lst"
Private mDefaultCaption As String
'Private WithEvents mainTimer As CTimer
Private Const cst_mainTimer_Interval As Long = 1500
Private Const cstTaskFormModel As Long = 0
Private mUnloading As Boolean
Private mIncomingFile As String
Private mIncomingFile2 As String
Private mTaskDownloadThreads As Long
Private mTaskRunner() As CTaskRunner
Private mLoginUrl As String
Private mPostUrl As String
Private mUserName As String
Private mPassword As String
Private mJpgBookCookie As String
Public JPGBookQuality As Integer
Public RenameJPGToPdg As Boolean
Public LastAddress As String
Private mTimerHit As Boolean
Private mTaskListFile As String

Private mTimesCount As Long

#If Not afNoCTimer = 1 Then
    Private WithEvents mainTimer As CTimer
Attribute mainTimer.VB_VarHelpID = -1



Private Sub fraStatus_DblClick()
    Dim task As CTask
    For Each task In mTasks
        If task.Status = STS_START Then
            Dim tItem As ListItem
            Set tItem = TaskItemFrom(task)
            If Not tItem Is Nothing Then
                On Error Resume Next
                LstTasks.SetFocus
                LstTasks.SelectedItem.selected = False
                tItem.selected = True
                tItem.EnsureVisible
                Exit For
            End If
        End If
    Next
End Sub

    Private Sub mainTimer_ThatTime()
#Else
    Private Sub mainTimer_Timer()
#End If
        'mTimerHit = True
        Static ImBusy As Boolean
        If ImBusy Then Exit Sub
        ImBusy = True
        main_loop
        ImBusy = False
    End Sub
    


Public Sub main_loop()
    
    
        If mUnloading Then
            Unload Me
            Exit Sub
        Else
            CheckTaskIncoming
            ProcessQueue
            mTimesCount = mTimesCount + 1
            If mTimesCount > 40 Then
                mTimesCount = 0
                ConfigSave
            End If
        End If
'        Do Until mTimerHit
'            DoEvents
'        Loop
'        mTimerHit = False
    'Loop
End Sub
Private Sub ITaskNotify_UpdateJpgBook(vTask As CTask)
    UpdateJpgBook vTask
End Sub
Public Property Let TimerEnabled(ByVal vEnable As Boolean)
'    If mainTimer Is Nothing Then Exit Property
'    If vEnable Then
'        mainTimer.Interval = cst_mainTimer_Interval
'    Else
'        mainTimer.Interval = 0
'    End If
End Property

Public Property Get TimerEnabled() As Boolean
    TimerEnabled = (mainTimer.Interval > 0)
End Property
Private Sub cmdAdd_Click()

frmTask.EditTaskMode = False
frmTask.Init Nothing
frmTask.UpdateFromClipboard
frmTask.Show cstTaskFormModel, Me

End Sub

Private Function ListItemToTask(ByRef vListItem As ListItem) As CTask
    On Error Resume Next
    Dim sKey As String
    sKey = vListItem.Key
    Set ListItemToTask = mTasks(vListItem.Key)
    
End Function

Private Function TaskToTaskListItem(ByRef vTask As CTask) As ListItem
    On Error Resume Next
    Set TaskToTaskListItem = LstTasks.ListItems.Item(vTask.taskId)
End Function

Private Function TaskToInfoListItem(ByRef vTask As CTask) As ListItem
    
    On Error Resume Next
    Set TaskToInfoListItem = lstTaskInfo.ListItems.Item(vTask.taskId)
End Function

Private Sub cmdEdit_Click()
    Dim selectListItem As ListItem
    Set selectListItem = LstTasks.SelectedItem
    If selectListItem Is Nothing Then
        MsgBox "No tasks selected.", vbInformation
        Exit Sub
    End If
        
    Dim task As CTask
    Set task = ListItemToTask(selectListItem)
    If task Is Nothing Then Exit Sub
                                
    

    If task.Status = STS_START Then
        frmTask.Init task, True
    Else
        frmTask.Init task, False
    End If
    frmTask.EditTaskMode = True
    
    frmTask.chkDetectClipboard.Value = False
    frmTask.Show cstTaskFormModel, Me
    
    
    'frmTask.chkDetectClipboard = False
    
    'UpdateTaskItem Task, , True
    'frmTask.EditTaskMode = False
    'selectListItem.Text = task.Title & "_" & task.SSID
    
End Sub


Public Sub ActionOnAllTask(ByVal vAction As String)
    vAction = UCase$(vAction)
    Dim pTask As CTask
    Select Case vAction
        Case "开始"
            For Each pTask In mTasks
                If pTask.Status <> STS_PENDING And pTask.Status <> STS_START Then
                    pTask.Status = STS_PENDING
                    pTask.Changed = True
                End If
            Next
        Case "停止"
            For Each pTask In mTasks
                If pTask.Status = STS_START Then
                    If FreeTaskRunner(pTask) Then
                        pTask.Status = STS_PAUSE
                        pTask.Changed = True
                    End If
                ElseIf pTask.Status = STS_PENDING Then
                    pTask.Status = STS_PAUSE
                    pTask.Changed = True
                End If
            Next
        Case "停止运行任务"
            For Each pTask In mTasks
                If pTask.Status = STS_START Then
                    FreeTaskRunner pTask
                    pTask.Status = STS_PENDING
                End If
            Next
    End Select
    TimerEnabled = False
    Windows_LockWindow LstTasks.hWnd
    UpdateTaskItem Nothing, , True
    TimerEnabled = True
    Windows_LockWindow 0
End Sub

'CSEH: ErrMsgBox
Public Sub ActionOnSingleItem(ByVal vAction As String, Optional ByRef vTask As CTask, Optional ByRef vItem As ListItem)
        '<EhHeader>
        On Error GoTo ActionOnSingleItem_Err
        '</EhHeader>
    
100     If vAction = "" Then Exit Sub
    
102     If vTask Is Nothing And vItem Is Nothing Then
104         Set vItem = LstTasks.SelectedItem
106         Set vTask = ListItemToTask(vItem)
108     ElseIf vTask Is Nothing Then
110         Set vTask = ListItemToTask(vItem)
112     ElseIf vItem Is Nothing Then
114         Set vItem = TaskToTaskListItem(vTask)
        End If
    
116     If vTask Is Nothing Then Exit Sub
118     If vItem Is Nothing Then Exit Sub
    
120     Select Case vAction
    
            Case "开始"
122             If Not vTask.Status = STS_START Then
124                 vTask.Status = STS_PENDING
126                 UpdateTaskItem vTask, vItem, True
128                 vTask.Changed = True
                End If
130         Case "停止"
132             If Not vTask.Status = STS_PAUSE Then
134                 FreeTaskRunner vTask
                    'mTaskRunner(task.DownloadId).Abort
136                 vTask.Status = STS_PAUSE
138                 UpdateTaskItem vTask, vItem, True
140                 vTask.Changed = True
                End If
142         Case "删除"
144                 FreeTaskRunner vTask
146                 LstTasks.ListItems.Remove (vItem.Key)
148                 If (MsgBox("删除任务文件夹？" & vbCrLf & vTask.Directory, vbQuestion + vbYesNo) = vbYes) Then
                        On Error Resume Next
150                     DeleteFolder vTask.Directory
                    End If
152                 mTasks.Remove vTask.taskId
154         Case "属性"
156             If vTask.Status = STS_START Then
158                 frmTask.Init vTask, True
                Else
160                 frmTask.Init vTask, False
                End If
162             frmTask.EditTaskMode = True
164             frmTask.chkDetectClipboard.Value = False
166             frmTask.Show cstTaskFormModel, Me
168         Case "更新信息"
170             If vTask.IsJpgBook Then
172                 UpdateJpgBook vTask
                End If
174         Case "重新下载"
176             If Not vTask.Status = STS_START Then
178                 If (MsgBox("重新下载任务？" & vbCrLf & vTask.Directory, vbQuestion + vbYesNo) = vbYes) Then
                        On Error Resume Next
180                     Kill BuildPath(vTask.Directory, "*.pdg")
                    End If
182                 vTask.Status = STS_PENDING
    '                vTask.Changed = True
    '                vTask.InitForDownload
    '                vTask.AutoSave
                End If
184         Case "下载JPG"
                If vTask.IsJpgBook = False Then
186                 vTask.IsJpgBook = True
188                 UpdateJpgBook vTask
                    vTask.Changed = True
                End If
190         Case "下载PDG"
                If vTask.IsJpgBook = True Then
192                 vTask.IsJpgBook = False
                    vTask.Changed = True
                End If
194         Case "打开链接"
196             OpenInSSReader vTask
                'vTask.GetJPGPage = True
        End Select
        '<EhFooter>
        Exit Sub

ActionOnSingleItem_Err:
        MsgBox Err.Description & vbCrLf & _
               "in GetSSLibX.frmMain.ActionOnSingleItem " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        '</EhFooter>
End Sub
Public Sub OpenInSSReader(ByRef vTask As CTask)
    Dim mainSite As String
    mainSite = SubStringBetween(mLoginUrl, "http://", "/")
    If mainSite = "" Then mainSite = SubStringUntilMatch(mLoginUrl, 1, "/")
    If mainSite = "" Then mainSite = mLoginUrl ' SubStringBetween(mLoginUrl, "http://", "")
    If mainSite <> "" Then mainSite = "http://" & mainSite

    Dim pBook As CBookInfo
    Set pBook = vTask.bookInfo
    If pBook Is Nothing Then Exit Sub
    If pBook(SSF_SSURL) = "" Then
        If pBook(SSF_PAGEURL) <> "" Then
            pBook(SSF_SSURL) = SSLIB_GetReadBookURL(pBook(SSF_PAGEURL), mainSite, mJpgBookCookie)
            vTask.Changed = True
        End If
    End If
    If pBook(SSF_SSURL) = "" Then
        MsgBox "不能打开" & vTask.ToString & vbCrLf & "链接无效", vbInformation
    Else
        MShell32.ShellExecute -1, "open", pBook(SSF_SSURL), "", "", SW_SHOWNA
    End If
End Sub
Public Sub StartProcess()
    Dim I As Long
    For I = 0 To cmd.UBound
        If cmd(I).Caption = "开始运行" Then cmd(I).Caption = "停止运行"
    Next
    'ActionOnTaskItem "全部停止"
    mainTimer.Interval = cst_mainTimer_Interval
End Sub

Public Sub StopProcess()
    Dim I As Long
    For I = 0 To cmd.UBound
        If cmd(I).Caption = "停止运行" Then cmd(I).Caption = "开始运行": Exit For
    Next
    'ActionOnTaskItem "全部停止"
    mainTimer.Interval = 0
End Sub

Public Sub ActionOnTaskItem(ByVal vAction As String)
On Error Resume Next
    
    
    Select Case vAction
    
        Case "添加"
            cmdAdd_Click
            Exit Sub
        Case "全部开始"
            ActionOnAllTask "开始"
            Exit Sub
        Case "全部停止"
            ActionOnAllTask "停止"
            Exit Sub
        Case "打开浏览器"
            OpenExplorer
            Exit Sub
        Case "登录sslib"
            LoginSSlib
            Exit Sub
        Case "设置Cookie"
            SetCookie
            Exit Sub
        Case "批量添加"
            AddBatchTasks
            Exit Sub
        Case "开始运行"
            StartProcess
            Exit Sub
        Case "停止运行"
            StopProcess
            ActionOnAllTask "停止运行任务"
            Exit Sub
    End Select
    
    
    If LstTasks.SelectedItem Is Nothing Then
        MsgBox "No tasks selected.", vbInformation, vAction
        Exit Sub
    End If
    
    Select Case vAction
        Case "属性"
            ActionOnSingleItem vAction, , LstTasks.SelectedItem
            Exit Sub
        Case "打开链接"
            ActionOnSingleItem vAction, , LstTasks.SelectedItem
            Exit Sub
    End Select
    
    'mainTimer.Interval = 0 ' False
    Windows_LockWindow LstTasks.hWnd
    Dim selectListItem As ListItem
    Dim selected() As ListItem
    Dim count As Long
    For Each selectListItem In LstTasks.ListItems
        If selectListItem.selected Then
            count = count + 1
            ReDim Preserve selected(1 To count)
            Set selected(count) = selectListItem
        End If
    Next
        
    Dim I As Long
    Dim task As CTask
    For I = 1 To count
        Set selectListItem = selected(I)
        Set task = ListItemToTask(selectListItem)
        ActionOnSingleItem vAction, task, selectListItem
    Next
    UpdateTaskItem
    Windows_LockWindow 0
    'mainTimer.Interval = cst_mainTimer_Interval
    
End Sub

Private Function FreeTaskRunner(ByRef vTask As CTask) As Boolean
    If vTask.DownloadId > 0 Then
        mTaskRunner(vTask.DownloadId).Abort
        vTask.DownloadId = 0
        FreeTaskRunner = True
    End If
End Function





Private Sub cmdConfigSave_Click()
    ConfigSave
End Sub

Private Sub cmd_Click(Index As Integer)
    If cmd(Index).Enabled = True Then
        ActionOnTaskItem cmd(Index).Caption
    End If

End Sub
Public Sub AddBatchTasks()
    frmBatchTasksAdd.Show 1, Me
End Sub
Public Sub SetCookie()
    Dim ret As String
    ret = frmInputBox.Popup(mJpgBookCookie, App.ProductName, "设置Http Cookie:")
    If ret <> "" Then mJpgBookCookie = ret
End Sub
Public Sub LoginSSlib()
    SetStatusText "登录" & mLoginUrl & "..."
    Dim ret As String
    ret = SSLibraryLogin(mUserName, mPassword, mLoginUrl, mPostUrl, True)
    If ret <> "" Then
        MsgBox "登录成功", vbOKOnly
        mJpgBookCookie = ret
    Else
        MsgBox "登录失败", vbCritical
    End If
    SetStatusText ""
End Sub

Private Sub OpenExplorer()
    On Error Resume Next
    Shell "sslibExplorer.exe", vbNormalFocus
    'Load frmBrowser
    'frmBrowser.StartingAddress = LastAddress
    'frmBrowser.Init mJpgBookStartURL, mUsername, mJpgBookCookie
    'frmBrowser.Show 0, Me
End Sub
'Private Sub Form_KeyPress(KeyAscii As Integer)
'Debug.Print Chr$(KeyAscii) & vbTab & KeyAscii
'End Sub
Private Sub InitTaskRunner()
    Dim pU As Long
    On Error Resume Next
    pU = UBound(mTaskRunner)
    If mMaxTasksDownloading < 1 Then mMaxTasksDownloading = 1
    If mMaxTasksDownloading > pU Then pU = mMaxTasksDownloading
    ReDim Preserve mTaskRunner(1 To pU)
    Dim I As Long
    For I = 1 To pU
        If mTaskRunner(I) Is Nothing Then Set mTaskRunner(I) = New CTaskRunner
        'mTaskRunner(i).Client = Me
    Next
End Sub


Private Sub Form_Load()

    InitCommonControlsVB
    MDebug.BugInit
    
    #If Not afNoCTimer = 1 Then
        Set mainTimer = New CTimer
    #End If
    mainTimer.Interval = 0


    
    
    'mainTimer.Enabled = False
    
    
    
    
    fDoUpdating = True
    Set mTasks = New Collection
    
    
    
    If Command$ <> "" Then
        Dim pName As String
        pName = Command$
        mDefaultCaption = App.EXEName & " " & pName & " " & App.Major & "." & App.Minor & " Build " & App.Revision
        mTaskListFile = App.Path & "\" & pName & "Tasks.lst"
        mConfigFile = App.Path & "\" & pName & "Config.ini"
        mIncomingFile = App.Path & "\" & pName & "Incoming.lst"
        mIncomingFile2 = App.Path & "\" & pName & "Incoming2.lst"
    Else
        mDefaultCaption = App.EXEName & " " & App.Major & "." & App.Minor & " Build " & App.Revision
        mConfigFile = App.Path & "\" & CSTTaskConfigFilename
        mIncomingFile = App.Path & "\" & CSTTaskIncomeFilename
        mIncomingFile2 = App.Path & "\" & CSTTaskIncomeFilename2
        mTaskListFile = App.Path & "\" & App.EXEName & "Tasks.lst"
    End If
   
   
    Me.Caption = mDefaultCaption
    ConfigLoad
    
    If mMaxTasksDownloading < 1 Then mMaxTasksDownloading = cstMaxTasksDownloading
    If mTaskDownloadThreads < 1 Then mTaskDownloadThreads = cstTaskDownloadThreads
    mCurrentRuning = 0
    
    InitTaskRunner
    
    If LstTasks.ListItems.count > 0 Then
        LstTasks.ListItems(1).selected = True
        'UpdateDownloadInfo ListItemToTask(LstTasks.ListItems(1))
    End If
    
    'LoginSSlib
    'mainTimer.Interval = cst_mainTimer_Interval
    'main_loop
    
    'mainTimer.Enabled = True
    
    
    
    
   
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
    
    Dim I As Long
    
    With fraContent
        'TabContent.Move cstLeft, 160, .Width - 2 * cstLeft
        For I = 1 To fraTabs.count
            'If (fraTabs(i).Visible = True) Then
                'fraTabs(i).Move cstLeft, TabContent.Top + TabContent.Height + cstTop, .Width - 2 * cstLeft, .Height - TabContent.Top - TabContent.Height - 2 * cstTop
                fraTabs(I).Move cstLeft, cstTop, .Width - 2 * cstLeft, .Height - 2 * cstTop
            'End If
        Next
    
    End With
           
    If fraTabs(1).Visible Then
        
        FraButtons.Top = cstTop
        FraButtons.Left = cstLeft
        FraButtons.Width = cmd(0).Width + 2 * cstLeft
        FraButtons.Height = fraTabs(1).Height - FraButtons.Top - cstTop
    
        
        cmd(0).Move cstLeft, cstTop
        For I = 1 To cmd.UBound
            cmd(I).Move cmd(I - 1).Left, cmd(I - 1).Top + cmd(I - 1).Height + 2 * cstTop
        Next
        
        
        FrameTaskInfo.Left = FraButtons.Left + FraButtons.Width + 2 * cstLeft
        FrameTaskInfo.Width = fraTabs(1).Width - FrameTaskInfo.Left - 2 * cstLeft
        

        
    '    Dim lastLabel As Label
    '    Set lastLabel = lblTaskInfo(lblTaskInfo.count - 1)
        'lblTaskInfo(0).Width = FrameTaskInfo.Width - 2 * lblTaskInfo(0).Left
        'lbltaskinfo(0).Top =
        For I = 1 To lblTaskInfo.UBound
            'lbltaskinfo(i).Move
            lblTaskInfo(I).Top = lblTaskInfo(I - 1).Top + lblTaskInfo(I - 1).Height + 120
         '   lblTaskInfo(i).Width = lblTaskInfo(0).Width
        Next
        I = lblTaskInfo.UBound
        lstTaskInfo.Move lstTaskInfo.Left, lblTaskInfo(I).Top + lblTaskInfo(I).Height + 120, _
            FrameTaskInfo.Width - lstTaskInfo.Left - 2 * cstLeft ',   FrameTaskInfo.Height -lblTaskInfo(i).Top - lblTaskInfo(i).Height - 180
            

        FrameTaskInfo.Height = lstTaskInfo.Top + lstTaskInfo.Height + cstTop
        FrameTaskInfo.Top = fraTabs(1).Height - FrameTaskInfo.Height
        
            
            
            LstTasks.Left = FrameTaskInfo.Left
            LstTasks.Top = cstTop ' cmdAdd.Top + cmdAdd.Height + 120
            LstTasks.Width = fraTabs(1).Width - LstTasks.Left - 2 * cstLeft
            LstTasks.Height = FrameTaskInfo.Top - LstTasks.Top
            
            lineOne.X1 = LstTasks.Left - cstLeft / 2
            lineOne.X2 = lineOne.X1
            lineOne.Y1 = FraButtons.Top + 2 * cstTop
            lineOne.Y2 = lineOne.Y1 + FraButtons.Height - 2 * cstTop
        
    End If
       
    If fraTabs(2).Visible Then
        
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

Public Function AddTaskItem(ByRef vTask As CTask, Optional vpending As Boolean = False, Optional vSave As Boolean = False) As ListItem
    If vpending Then vTask.Status = STS_PENDING
    vTask.taskId = NewTaskID()
    'vtask.Init Me
       
    mTasks.Add vTask, vTask.taskId
    Dim taskListItem As ListItem
    Set taskListItem = LstTasks.ListItems.Add(, vTask.taskId)
'    If vSave Then vtask.AutoSave
    UpdateTaskItem vTask, taskListItem, True
    SetStatusText vTask.ToString & " 添加到任务列表"
    Set AddTaskItem = taskListItem
    
End Function

Public Sub UpdateJpgBook(ByRef vTask As CTask)
    If vTask Is Nothing Then Exit Sub
    
    SetStatusText vTask.bookInfo(SSF_SSID) & vTask.bookInfo(SSF_Title) & "：更新JPG大图信息..."
    vTask.UpdateFolder
    Dim mainSite As String
    mainSite = SubStringBetween(mLoginUrl, "http://", "/")
    If mainSite = "" Then mainSite = SubStringUntilMatch(mLoginUrl, 1, "/")
    If mainSite = "" Then mainSite = mLoginUrl ' SubStringBetween(mLoginUrl, "http://", "")
    If mainSite <> "" Then mainSite = "http://" & mainSite
    If (SSLIB_UpdateTaskOnNeed(vTask.bookInfo, vTask.Directory, mJpgBookCookie, False, mainSite) = True) Then
        vTask.UpdateFolder
        vTask.Changed = True
        UpdateTaskItem vTask, , True
    End If
    SetStatusText ""
End Sub

Private Function AddTaskByRef(ByRef vTask As CTask, Optional vpending As Boolean) As CTask
    If vTask Is Nothing Then Exit Function
    AddTaskItem vTask, vpending
    Set AddTaskByRef = vTask
End Function

Private Function AddTaskByBookInfo(ByRef vBookInfo As CBookInfo, Optional vpending As Boolean) As CTask
    Dim newTask As CTask
    Set newTask = New CTask
    Set newTask.bookInfo = vBookInfo
    Set AddTaskByBookInfo = AddTaskByRef(newTask, vpending)
End Function

Private Function AddTask(vBookInfoArray() As String, Optional vpending As Boolean) As CTask
        Dim bookInfo As CBookInfo
        Set bookInfo = New CBookInfo
        bookInfo.LoadFromArray vBookInfoArray, False
        Set AddTask = AddTaskByBookInfo(bookInfo, vpending)
End Function

Public Sub CallBack_EditTask(ByRef vTask As CTask)
    If vTask.IsJpgBook Then UpdateJpgBook vTask
    'SSLIB_UpdateTaskOnNeed vTask, mJpgBookCookie
    vTask.UpdateFolder
    UpdateTaskItem vTask, , True
    vTask.Changed = True
    vTask.AutoSave
End Sub

Public Sub CallBack_AddTask(ByVal vName As String, vBookInfoArray() As String, Optional vpending As Boolean)
        TimerEnabled = False
        Dim vTask As CTask
        Set vTask = AddTask(vBookInfoArray, vpending)
        vTask.Name = vName
        'If Not vtask Is Nothing Then vtask.Changed = True
        vTask.Changed = True
        If vTask.IsJpgBook Then UpdateJpgBook vTask
        vTask.AutoSave
        TimerEnabled = True
End Sub
'
'Public Sub CallBack_AddTask(bookInfo As CBookInfo)
'    Dim vtask As CTask
'    Set vtask = New CTask
'    Set vtask.bookInfo = bookInfo
'End Sub

Private Function TextOfStatus(ByRef Status As SSLIBTaskStatus) As String
    Dim Text As String
Select Case Status
    Case SSLIBTaskStatus.STS_COMPLETE
        Text = "完成"
    Case SSLIBTaskStatus.STS_PAUSE
        Text = "停止"
    Case SSLIBTaskStatus.STS_PENDING
        Text = "排队中..."
    Case SSLIBTaskStatus.STS_START
        Text = "正在下载..."
    Case SSLIBTaskStatus.STS_ERRORS
        Text = "出错了"
    Case Else
        Text = "停止"
End Select

TextOfStatus = Text
End Function

Private Function TaskItemFrom(ByRef vTask As CTask) As ListItem
    On Error Resume Next
    If vTask Is Nothing Then Exit Function
    Dim Item As ListItem
    Set Item = LstTasks.ListItems(vTask.taskId)
    Set TaskItemFrom = Item
End Function

Private Sub UpdateTaskItem(Optional ByRef vTask As CTask, Optional ByRef vItem As ListItem, Optional vForce As Boolean = False)
    On Error Resume Next
    'Dim i As Integer
    'Me.Enabled = False
   ' DoEvents
'    Dim vTask As CTask
'    Set vTask = vTaskRunner.task
    
    If Not vTask Is Nothing Then
            If vTask.FilesCount > vTask.PagesCount Then
                pbDownload.Value = pbDownload.Max
            Else
                pbDownload.Max = vTask.PagesCount
                pbDownload.Value = vTask.FilesCount
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
    
    If Not vTask Is Nothing Then
        Set pBookInfo = vTask.bookInfo
        If vItem Is Nothing Then Set vItem = TaskItemFrom(vTask)
        If vItem Is Nothing Then Exit Sub
        If vItem.selected = True Then
            
            lblTaskInfo(0).Caption = "ID: " & vTask.taskId & " " & pBookInfo(SSF_SSID)
            lblTaskInfo(1).Caption = "任务名: " & pBookInfo(SSF_Title) & " " & pBookInfo(SSF_AUTHOR)
            lblTaskInfo(1).Tag = vTask.Directory
            
            lblTaskInfo(2).Tag = lblTaskInfo(1).Tag
            lblTaskInfo(2).Caption = "下载位置: " & lblTaskInfo(2).Tag
            
            
            lblTaskInfo(3).Caption = "URL: " & IIf(vTask.IsJpgBook, pBookInfo(SSF_JPGURL), pBookInfo(SSF_URL))
            '            lblOpenPdg.Tag = lblTaskInfo(2).Tag
            
        End If
        
        vItem.Text = Mid$(vTask.taskId, 7)
            vItem.SubItems(1) = TextOfStatus(vTask.Status)
            vItem.SubItems(2) = pBookInfo(SSF_SSID)
            vItem.SubItems(3) = pBookInfo(SSF_Title)
            vItem.SubItems(4) = pBookInfo(SSF_AUTHOR)
            vItem.SubItems(5) = vTask.FilesCount & "/" & vTask.PagesCount
            
            If vTask.FilesCount = vTask.PagesCount Then
                vItem.SubItems(6) = "√"
            ElseIf vTask.FilesCount > vTask.PagesCount Then
                vItem.SubItems(6) = ">>"
            Else
                vItem.SubItems(6) = "<<"
            End If
           If vTask.IsJpgBook Then vItem.SubItems(7) = "JPG" Else vItem.SubItems(7) = "PDG"

            
            
            'pbDownload.Min = 0
            
            
            'vItem.SubItems(6) = vTask.FilesCount
            'vItem.SubItems(6) = vTask.Downloader.ErrorsCount
        'Me.Enabled = True
            'vItem.text = GetFileName(vTask.Directory)
            
        '    UpdateDownloadInfo vtask.Downloader, vItem
        Exit Sub
    End If
    
    For Each vTask In mTasks
        Set vItem = TaskItemFrom(vTask)
        If Not vItem Is Nothing Then
            Set pBookInfo = vTask.bookInfo
            vItem.Text = Mid$(vTask.taskId, 7)
            vItem.SubItems(1) = TextOfStatus(vTask.Status)
            vItem.SubItems(2) = pBookInfo(SSF_SSID)
            vItem.SubItems(3) = pBookInfo(SSF_Title)
            vItem.SubItems(4) = pBookInfo(SSF_AUTHOR)
            vItem.SubItems(5) = vTask.FilesCount & "/" & vTask.PagesCount
            If vTask.FilesCount = vTask.PagesCount Then
                vItem.SubItems(6) = "√"
            ElseIf vTask.FilesCount > vTask.PagesCount Then
                vItem.SubItems(6) = ">>"
            Else
                vItem.SubItems(6) = "<<"
            End If
            If vTask.IsJpgBook Then vItem.SubItems(7) = "JPG" Else vItem.SubItems(7) = "PDG"
            'vItem.SubItems(6) = vTask.FilesCount
            'vItem.SubItems(6) = vTask.Downloader.ErrorsCount
        End If
    Next
    

    'Me.Enabled = True
End Sub
Private Sub UpdateDownloadInfo(ByRef vTaskRunner As CTaskRunner, Optional ByRef vItem As ListItem)

'Exit Sub

'Static timeLastUpdate As Single
'Dim CurrentTime As Single
'CurrentTime = DateTime.mainTimer
'If CurrentTime - timeLastUpdate > 0.8 Then
'    UpdateTaskItem vTask, vItem
'End If
'timeLastUpdate = CurrentTime

DoEvents
Dim vTask As CTask
Set vTask = vTaskRunner.task

If vTask Is Nothing Then Exit Sub
If vItem Is Nothing Then Set vItem = TaskItemFrom(vTask)
If vItem Is Nothing Then Exit Sub
If vItem.selected = False Then Exit Sub

Dim I As Long
Dim C As Long
C = vTaskRunner.ThreadProgressCount '  .ThreadCount ' .ConnectionsCount

'Me.Enabled = False
'lstTaskInfo.ListItems.Clear

For I = 0 To C - 1

Dim Progress As CDownloadProgress
Set Progress = vTaskRunner.ThreadProgress(I)
If Progress Is Nothing Then GoTo ContinueFor
'If Not fDoUpdating Then Exit Sub
    
Dim itemTask As ListItem
If lstTaskInfo.ListItems.count < I + 1 Then
    Set itemTask = lstTaskInfo.ListItems.Add()
Else
    Set itemTask = lstTaskInfo.ListItems.Item(I + 1)
    itemTask.Text = ""
    itemTask.ListSubItems.Clear
End If
If itemTask Is Nothing Then GoTo ContinueFor

itemTask.Text = I
itemTask.SubItems(1) = Progress.URL  ' vTask.CurrentFile
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



'Me.Enabled = True

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
    mainTimer.Interval = 0
    #If Not afNoCTimer = 1 Then
        Set mainTimer = Nothing
    #End If
    
    Dim task As CTask
    SetStatusText "退出程序..."
'    If mUnloading = True Then
'        For Each task In mTasks
'            If task.Status = STS_START Then
'                'Task.Status = STS_PENDING
'                SetStatusText "正在停止任务：[" & task.taskId & "]" & task.bookInfo(SSF_Title) & "..."
'                Cancel = True
'                Exit Sub
'            End If
'        Next
'    Else
'        mUnloading = True
'        Cancel = True
        SetStatusText "停止正在下载的任务..."
        For Each task In mTasks
            If task.Status = STS_START Then
                FreeTaskRunner task
                task.Status = STS_PENDING
            End If
        Next
        Dim I As Long
        For I = 1 To mMaxTasksDownloading
            mTaskRunner(I).Abort
            Set mTaskRunner(I) = Nothing
        Next
    
        SetStatusText "正在保存程序和任务配置..."
        ConfigSave
'        For Each task In mTasks
'            If task.Status = STS_START Then FreeTaskRunner task
'        Next
'        Exit Sub
'    End If
                

    
    mUnloading = False
   ' mainTimer.Interval = 0 ' False
    
    'MsgBox "3"
    'Unload frmTask
    'MsgBox "4"
    Set mTasks = Nothing
    'Set mainTimer = Nothing
    'MsgBox "5"
    
    MDebug.BugTerm
    
    End
End Sub





Private Sub SetStatusText(ByRef vText As String)
    lblStatus.Caption = vText
End Sub

Private Sub ITaskNotify_DownloadStatusChange(vTaskRunner As CTaskRunner)
    UpdateDownloadInfo vTaskRunner
End Sub

Private Sub ITaskNotify_StatusChange(vText As String)
    SetStatusText vText
End Sub

Private Sub ITaskNotify_TaskComplete(vTaskRunner As CTaskRunner)
    On Error Resume Next
    'If vTaskRunner Is Nothing Then
        mCurrentRuning = mCurrentRuning - 1
        
        SetStatusText ""
        pbDownload.Value = pbDownload.Min
        
        'Me.Caption = mDefaultCaption
       UpdateTaskItem vTaskRunner.task, , True
       UpdateDownloadInfo vTaskRunner
   ' End If
        FreeTaskRunner vTaskRunner.task
       'vTaskRunner.task.DownloadId = 0
End Sub

Private Sub ITaskNotify_TaskStatusChange(vTaskRunner As CTaskRunner)
    On Error Resume Next
    UpdateTaskItem vTaskRunner.task
    UpdateDownloadInfo vTaskRunner
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
        If pdgOpener = "" Then pdgOpener = "开始"
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



Private Sub LstTasks_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Dim Index As Long
    Index = ColumnHeader.Index - 1
    LstTasks.Sorted = False
    If Index = LstTasks.SortKey Then
        If LstTasks.SortOrder = lvwAscending Then
            LstTasks.SortOrder = lvwDescending
        Else
            LstTasks.SortOrder = lvwAscending
        End If
    Else
        LstTasks.SortOrder = lvwAscending
    End If
    LstTasks.SortKey = Index
    LstTasks.Sorted = True
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
    
    'lstTaskInfo.ListItems.Clear
    If mTasks(lastKey).DownloadId > 0 Then
        UpdateDownloadInfo mTaskRunner(mTasks(lastKey).DownloadId)  'ListItemToTask(item)
    End If
End Sub

Private Sub LstTasks_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Debug.Print "Shift:" & Shift & vbTab & "Code:" & KeyCode & "(" & Chr$(KeyCode) & ")"
    If Shift = 1 And KeyCode = 97 Or KeyCode = 65 Then
        Windows_LockWindow LstTasks.hWnd
        Dim Item As ListItem
        For Each Item In LstTasks.ListItems
            Item.selected = True
        Next
        Windows_LockWindow 0
        'LstTasks.se
        Exit Sub
    End If
    Select Case KeyCode
    
        Case 46
            ActionOnTaskItem "删除"
            'cmdRemove_Click
    
    End Select
    
End Sub


Private Sub LstTasks_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        mnuTask_Click
        PopupMenu mnuTask, , _
             fraContent.Left + fraTabs(1).Left + LstTasks.Left + x, _
            fraContent.Top + fraTabs(1).Top + LstTasks.Top + y
    End If
End Sub





Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuImport_Click()
    Dim filename As String
    Dim dlg As CCommonDialogLite
    Set dlg = New CCommonDialogLite
    If dlg.VBGetOpenFileName(filename, , , , , , cstTaskListFilter, , , , cstTaskListExt, Me.hWnd) Then
            TaskNewFromFile filename
    End If
End Sub

Private Sub mnuImportDir_Click()
    Static Lastfolder As String
    Dim folder As String
    Dim dlg As CFolderBrowser
    Set dlg = New CFolderBrowser
    dlg.Owner = Me.hWnd
    If Lastfolder <> "" Then dlg.InitDirectory = Lastfolder
    folder = dlg.Browse
    If folder <> "" Then
        Lastfolder = folder
        Dim subdirs() As String
        Dim count As Long
        count = subFolders(BuildPath(folder), subdirs())
        Dim I As Long
        Dim fn As String
        If count > 0 Then
            Me.TimerEnabled = False
            Me.Enabled = False
            Windows_LockWindow LstTasks.hWnd
            For I = 0 To count - 1
                If FolderExists(subdirs(I)) Then TaskNewFromDirectory subdirs(I)
            Next
            Windows_LockWindow 0
            Me.Enabled = True
            Me.TimerEnabled = True
        End If
    End If
    
End Sub

Private Sub mnuImportList_Click()
    Dim filename As String
    Dim dlg As CCommonDialogLite
    Set dlg = New CCommonDialogLite
    If dlg.VBGetOpenFileName(filename, , , , , , cstTaskListFilter, , , , cstTaskListExt, Me.hWnd) Then
            TaskListReadFrom filename
    End If
End Sub

Private Sub mnuPreference_Click()
   EditConfig
End Sub

Private Sub mnuSave_Click()
    ConfigSave
End Sub

Private Sub mnuSaveAs_Click()
    Dim filename As String
    Dim dlg As CCommonDialogLite
    Set dlg = New CCommonDialogLite
    If dlg.VBGetSaveFileName(filename, , , cstTaskListFilter, , , , cstTaskListExt, Me.hWnd) Then
        TaskListWriteTo filename
    End If
    
End Sub

Private Function GetSelectedTask() As CTask
    Set GetSelectedTask = ListItemToTask(LstTasks.SelectedItem)
End Function

'Private Function GetSelectedItem() As ListItem
'    Set GetSelectedItem = LstTasks.SelectedItem
'End Function
Private Sub mnuTask_Click()
    Dim task As CTask
    Set task = GetSelectedTask()
    If task Is Nothing Then
        mnuTaskID.Caption = "目标：（无）"
    Else
        mnuTaskID.Caption = "目标：[" & task.bookInfo(SSF_SSID) & "]" & task.bookInfo(SSF_Title)
    End If
End Sub

Private Sub mnuTaskAction_Click(Index As Integer)
        If mnuTaskAction(Index).Caption <> "-" Then
            ActionOnSingleItem mnuTaskAction(Index).Caption
        End If
        'lsttasks.
End Sub

Private Sub mnuTaskClear_Click()

    Dim task As CTask
    Dim tasks() As CTask
    Dim Item As ListItem
    Dim count As Long
    On Error Resume Next
    For Each task In mTasks
    If Not task.Status = STS_START Then
        count = count + 1
        ReDim Preserve tasks(1 To count)
        Set tasks(count) = task
    End If
    Next
    Me.Enabled = False
    Windows_LockWindow LstTasks.hWnd
    Dim I As Long
    For I = 1 To count
         If tasks(I).Changed = True Then tasks(I).AutoSave
         LstTasks.ListItems.Remove tasks(I).taskId
         mTasks.Remove tasks(I).taskId
         Set tasks(I) = Nothing
    Next
   Windows_LockWindow 0
   Me.Enabled = True
End Sub



Private Sub TabContent_Click()
    Dim I As Long
    For I = 1 To fraTabs.count
        fraTabs(I).Visible = False
    Next
    fraTabs(TabContent.SelectedItem.Index).Visible = True
    On Error Resume Next
    If TabContent.SelectedItem.Index = 2 Then
        frmTask.SolidMode = True
        frmTask.Show 0, Me
        frmTask.Move fraTabs(2).Left + (fraTabs(2).Width - frmTask.Width) / 2, fraTabs(2).Top + (fraTabs(2).Height - frmTask.Height) / 2
    End If
    
End Sub

Private Sub TabContent_Validate(Cancel As Boolean)
'    Dim i As Long
'    For i = 1 To fraTabs.count
'        fraTabs(i).Visible = False
'    Next
'    fraTabs(TabContent.SelectedItem.Index).Visible = True
End Sub

Private Function GetFreeTaskRunner() As Long
    Dim I As Long
    For I = 1 To mMaxTasksDownloading
        If mTaskRunner(I).IsFree Then
            mTaskRunner(I).JPGBookQuality = JPGBookQuality
            mTaskRunner(I).RenameJPGBook = RenameJPGToPdg
            GetFreeTaskRunner = I
            Exit Function
        End If
    Next
End Function

Private Sub ProcessQueue()
'    Static sProcessing As Boolean
'    If sProcessing Then Exit Sub
    If mUnloading Then Exit Sub
    
    'sProcessing = True
    On Error Resume Next
    
    Dim DownloadId As Long
    DownloadId = GetFreeTaskRunner()
    If DownloadId > 0 Then
    
        Dim I As Long
        Dim e As Long
        e = LstTasks.ListItems.count ' - 1 '   mTasks.Count
        Dim task As CTask
        For I = 1 To e
         'For Each task In mTasks
            Set task = Nothing
            Set task = mTasks(LstTasks.ListItems.Item(I).Key)
            If Not (task Is Nothing) Then
               If task.Status = STS_PENDING Then
                    task.Status = STS_START
                    If task.IsJpgBook Then UpdateJpgBook task
                    UpdateTaskItem task, , True
                    SetStatusText QuoteString(task.bookInfo(SSF_Title)) & " 正在下载..."
                    task.DownloadId = DownloadId
                    'Set mTaskRunner(DownloadId).Client = Me
                    'If task.IsJpgBook Then task.bookInfo(SSF_HEADER) = "Cookie: " & mJpgBookCookie
                    mTaskRunner(DownloadId).Execute Me, task, mTaskDownloadThreads
                    Exit For
                End If
            End If
        Next
    End If
'    sProcessing = False
'    For i = 1 To mMaxTasksDownloading
'        If mTaskRunner(i).IsFree = False Then mTaskRunner(i).Runonce
'    Next
    
    
End Sub
'Private Sub ProcessQueue_OLD()
'    DoEvents
'    If mUnloading Then Exit Sub
'    If mCurrentRuning >= mMaxTasksDownloading Then Exit Sub
'    Dim task As CTask
'    For Each task In mTasks
'        If task.Status = STS_PENDING Then
'            If mUnloading Then Exit Sub
'            task.Status = STS_START
'
'            mCurrentRuning = mCurrentRuning + 1
'            UpdateTaskItem task, , True
'            SetStatusText QuoteString(task.bookInfo(SSF_Title)) & " 正在下载..."
'            task.StartDownload
'            'taskRunner.Execute Me, task
''            If cstUseFrmProgress Then
''                    Dim newDownload As frmProgress
''                    Set newDownload = New frmProgress
''                    Set newDownload.MainApp = Me
''                    Set newDownload.Task = Task
''                    newDownload.Show 0, Me
''        '            newDownload.StartTask task
''
''                    'Load frmProgress
''            Else
''                    Task.StartDownload
''                    mCurrentRuning = mCurrentRuning - 1
''                    Me.Caption = App.ProductName
''            End If
'
'            'Set newDownload = Nothing
'            Exit Sub
'        End If
'    Next
'
'
'End Sub

Public Function NewTaskID() As String
    Static lastNum As Long
    If lastNum < 1 Then
        lastNum = 1
    Else
        lastNum = lastNum + 1
    End If
    NewTaskID = "SSBOOK" & StrNum(lastNum, 4)
    'NewTaskID = "SSBOOK" & CStr(lastNum)
End Function
Public Sub ConfigLoad()
        Dim iniHnd As CLiNInI
100     Set iniHnd = New CLiNInI
105     iniHnd.File = BuildPath(App.Path) & CSTAppConfigname

            FormStateFromString Me, iniHnd.GetSetting("App", "WindowPosition")
            m_StrPdgProg = iniHnd.GetSetting("App", "PdgProg")
            m_StrFolderProg = iniHnd.GetSetting("App", "FolderProg")
            mMaxTasksDownloading = StringToLong(iniHnd.GetSetting("App", "MaxTasksDownloading"))
            mTaskDownloadThreads = StringToLong(iniHnd.GetSetting("App", "TaskDownloadThreads"))
            mPassword = iniHnd.GetSetting("Browser", "Password")
            mUserName = iniHnd.GetSetting("Browser", "Username")
            mLoginUrl = iniHnd.GetSetting("Browser", "LoginUrl")
            mPostUrl = iniHnd.GetSetting("Browser", "PostUrl")
            mJpgBookCookie = iniHnd.GetSetting("Browser", "JpgBookCookie")
            JPGBookQuality = StringToInteger(iniHnd.GetSetting("JPGBook", "Quality"))
            RenameJPGToPdg = iniHnd.GetSetting("JPGBook", "RenameToPDG") = "1"

    TaskListReadFrom mTaskListFile
End Sub
Public Sub ConfigSave()
    TimerEnabled = False
    Dim iniHnd As CLiNInI
    Set iniHnd = New CLiNInI
    On Error Resume Next
    SetStatusText "正在保存配置和任务信息..."

        iniHnd.SaveSetting "App", "WindowPosition", FormStateToString(Me)
        iniHnd.SaveSetting "App", "PdgProg", m_StrPdgProg
        iniHnd.SaveSetting "App", "FolderProg", m_StrFolderProg
        iniHnd.SaveSetting "App", "MaxTasksDownloading", mMaxTasksDownloading
        iniHnd.SaveSetting "App", "TaskDownloadThreads", mTaskDownloadThreads
        iniHnd.SaveSetting "Browser", "Password", mPassword
        iniHnd.SaveSetting "Browser", "Username", mUserName
        iniHnd.SaveSetting "Browser", "LoginUrl", mLoginUrl '
        iniHnd.SaveSetting "Browser", "PostUrl", mPostUrl '
        iniHnd.SaveSetting "Browser", "JpgBookCookie", mJpgBookCookie
        iniHnd.SaveSetting "JPGBook", "Quality", JPGBookQuality
        iniHnd.SaveSetting "JPGBook", "RenameToPDG", IIf(RenameJPGToPdg, "1", "")

    iniHnd.WriteTo BuildPath(App.Path) & CSTAppConfigname
    TaskListWriteTo mTaskListFile
    
    Dim task As CTask
    For Each task In mTasks
        If task.Changed = True Then task.AutoSave
    Next
    TimerEnabled = True
    SetStatusText ""
    'ConfigSaveTo mConfigFile, True
End Sub

Public Sub EditConfig()
 Load frmOptions
    With frmOptions
        .PdgProg = m_StrPdgProg
        .FolderProg = m_StrFolderProg
        .TasksProcessing = mMaxTasksDownloading
        .ThreadsDownloading = mTaskDownloadThreads
        .Password = mPassword
        .UserName = mUserName
        .LoginUrl = mLoginUrl
        .PostUrl = mPostUrl
        .Show 1, Me
        m_StrPdgProg = .PdgProg
        m_StrFolderProg = .FolderProg
        mTaskDownloadThreads = .ThreadsDownloading
        mMaxTasksDownloading = .TasksProcessing
        mPassword = .Password
        mUserName = .UserName
        mLoginUrl = .LoginUrl
        mPostUrl = .PostUrl
        InitTaskRunner
    End With
    Unload frmOptions
End Sub
'Public Sub ConfigSaveTo(ByRef vFilename As String, Optional withAppSetting As Boolean = False)
'
'On Error GoTo ErrorSaveTaskTo
'
'    Dim iniHnd As CLiNInI
'    Set iniHnd = New CLiNInI
'    On Error Resume Next
'    If withAppSetting Then
'        iniHnd.SaveSetting "App", "WindowPosition", FormStateToString(Me)
'        iniHnd.SaveSetting "App", "PdgProg", m_StrPdgProg
'        iniHnd.SaveSetting "App", "FolderProg", m_StrFolderProg
'        iniHnd.SaveSetting "App", "MaxTasksDownloading", mMaxTasksDownloading
'        iniHnd.SaveSetting "App", "TaskDownloadThreads", mTaskDownloadThreads
'        iniHnd.SaveSetting "Browser", "Password", mPassword
'        iniHnd.SaveSetting "Browser", "Username", mUserName
'        iniHnd.SaveSetting "Browser", "LoginUrl", mLoginUrl '
'        iniHnd.SaveSetting "Browser", "PostUrl", mPostUrl '
'        iniHnd.SaveSetting "Browser", "JpgBookCookie", mJpgBookCookie
'        iniHnd.SaveSetting "JPGBook", "Quality", JPGBookQuality
'        iniHnd.SaveSetting "JPGBook", "RenameToPDG", IIf(RenameJPGToPdg, "1", "")
'     End If
'    iniHnd.SaveSetting "GetSSLib", "TaskDirectoryCount", CStr(mTasks.Count)
'    Dim task As CTask
'    Dim i As Long
'    Dim taskconfig As CLiNInI
'    Dim taskDir As String
'    For Each task In mTasks
'        i = i + 1
'        taskDir = task.Directory
'        If Not FolderExists(taskDir) Then xMkdir taskDir
'
'        iniHnd.SaveSetting "GetSSLib", "TaskDirectory" & CStr(i), taskDir
'        If withAppSetting And task.Changed = True Then
'            Set taskconfig = New CLiNInI
'            task.PersistTo taskconfig, "TaskInfo"
'            taskconfig.WriteTo BuildPath(taskDir, CSTTaskConfigFilename)
'            Set taskconfig = Nothing
'            task.Changed = False
'        End If
'    Next
'    iniHnd.WriteTo vFilename
'    Set iniHnd = Nothing
'
'    Exit Sub
'ErrorSaveTaskTo:
'    MsgBox Err.Description, vbCritical, "ConfigSaveTo"
'End Sub

''CSEH: ErrAsk
'Public Sub ConfigLoadFrom(ByRef vFilename As String)
'        '<EhHeader>
'        On Error GoTo ConfigLoadFrom_Err
'        '</EhHeader>
'        Dim task As CTask
'        Dim iniHnd As CLiNInI
'        Dim i As Long
'        Dim taskDir As String
'        Dim taskconfig As CLiNInI
'
'100     Set iniHnd = New CLiNInI
'105     iniHnd.File = vFilename
'
'        If withAppSetting Then
'            FormStateFromString Me, iniHnd.GetSetting("App", "WindowPosition")
'            m_StrPdgProg = iniHnd.GetSetting("App", "PdgProg")
'            m_StrFolderProg = iniHnd.GetSetting("App", "FolderProg")
'            mMaxTasksDownloading = StringToLong(iniHnd.GetSetting("App", "MaxTasksDownloading"))
'            mTaskDownloadThreads = StringToLong(iniHnd.GetSetting("App", "TaskDownloadThreads"))
'            mPassword = iniHnd.GetSetting("Browser", "Password")
'            mUserName = iniHnd.GetSetting("Browser", "Username")
'            mLoginUrl = iniHnd.GetSetting("Browser", "LoginUrl")
'            mPostUrl = iniHnd.GetSetting("Browser", "PostUrl")
'            mJpgBookCookie = iniHnd.GetSetting("Browser", "JpgBookCookie")
'            JPGBookQuality = StringToInteger(iniHnd.GetSetting("JPGBook", "Quality"))
'            RenameJPGToPdg = iniHnd.GetSetting("JPGBook", "RenameToPDG") = "1"
'        End If
'        Dim iTaskCount As Long
'
'110     iTaskCount = StringToLong(iniHnd.GetSetting("GetSSLib", "TaskCount"))
'
'115     For i = 1 To iTaskCount
'120         Set task = New CTask
'125         task.LoadFrom iniHnd, "Task" & CStr(i)
'            task.Changed = False
'            If task.IsJpgBook Then UpdateJpgBook task ' SSLIB_UpdateTaskOnNeed task, mJpgBookCookie
'130         AddTaskItem task, False
'        Next
'
'135     iTaskCount = 0
'140     iTaskCount = StringToLong(iniHnd.GetSetting("GetSSLib", "TaskDirectoryCount"))
'145     For i = 1 To iTaskCount
'150         taskDir = iniHnd.GetSetting("GetSSLib", "TaskDirectory" & CStr(i))
'155         If FolderExists(taskDir) Then
'160             Set task = New CTask
'
'165             Set taskconfig = New CLiNInI
'170             taskconfig.File = BuildPath(taskDir, CSTTaskConfigFilename)
'175             task.LoadFrom taskconfig, "TaskInfo"
'                task.Directory = taskDir
'
'180             AddTaskItem task, False
'                task.Changed = False
'                If task.IsJpgBook Then UpdateJpgBook task
'185             Set taskconfig = Nothing
'190             Set task = Nothing
'            End If
'        Next
'
'
'
'200     If iniHnd.SectionExists("TaskInfo") Then
'205         Set task = New CTask
'210         task.LoadFrom iniHnd, "TaskInfo"
'            task.Directory = GetParentFolderName(vFilename)
'            task.Changed = False
'            If task.IsJpgBook Then UpdateJpgBook task
'215         AddTaskItem task, False
'        End If
'
'        '<EhFooter>
'        Exit Sub
'
'ConfigLoadFrom_Err:
'    If MsgBox(Err.Description & vbCrLf & _
'               "in GetSSLib.frmMain.ConfigLoadFrom " & _
'               "at line " & Erl & vbCrLf & _
'               "Continue (Click No will terminate GetSSLib)?", _
'               vbExclamation + vbYesNo, "Application Error") = vbYes Then
'        Resume Next
'    Else
'        End
'    End If
'
'        '</EhFooter>
'End Sub






Public Sub TaskListWriteTo(ByVal vFilename As String)
    On Error GoTo TaskListWriteTo_Error
    
    TimerEnabled = False
    Dim fNum As Integer
    fNum = FreeFile
    If FileExists(vFilename) Then Kill vFilename
    
    Dim task As CTask
    Dim taskDir() As Byte
    
    Open vFilename For Binary Access Write As #fNum
    ReDim taskDir(0 To 1)
    taskDir(0) = &HFF
    taskDir(1) = &HFE
    Put #fNum, , taskDir
    'Dim taskDir As String

    For Each task In mTasks
        taskDir = task.Directory & vbCrLf
        Put #fNum, , taskDir
    Next
    Close #fNum
    TimerEnabled = True
    '<EhFooter>
    Exit Sub

TaskListWriteTo_Error:
    Debug.Print "TaskListWriteTo:" & Err.Description
    On Error Resume Next
    Close #fNum
    TimerEnabled = True
    Err.Clear
    
End Sub
'CSEH: ErrExit

Public Sub TaskNewFromDirectory(ByVal vFolder As String)
        Dim task As CTask
        Set task = New CTask
        task.Directory = vFolder
        task.AutoLoad
        If task.bookInfo(SSF_SSID) = "" And _
            task.bookInfo(SSF_HEADER) = "" And _
            task.bookInfo(SSF_URL) = "" And _
            task.bookInfo(SSF_Title) = "" Then
            Exit Sub
        End If
        task.Directory = vFolder
        task.Changed = False
        AddTaskByRef task
End Sub

Public Sub TaskNewFromFile(ByVal vFilename As String)
    Dim task As CTask
    Set task = New CTask
    task.bookInfo.LoadFromFile vFilename
    task.Directory = GetParentFolderName(vFilename)
    task.Changed = True
    AddTaskByRef task
End Sub

Public Sub TaskListReadFrom(ByVal vFilename As String)
    On Error GoTo TaskListReadFrom_Error

    TimerEnabled = False
    Windows_LockWindow LstTasks.hWnd
    Dim fNum As Integer
    fNum = FreeFile
    Open vFilename For Binary Access Read As #fNum
    If LOF(fNum) < 2 Then GoTo TaskListReadFrom_Error
    Dim a() As Byte
    ReDim a(0 To 1)
    Get #fNum, , a()
    Dim taskDir As String
    Dim vLines() As String
    
    If a(0) = &HFF And a(1) = &HFE Then
        ReDim a(0 To LOF(fNum) - 3)
        Get #fNum, , a()
        vLines = Split(a(), vbCrLf)
        Dim I As Long
        Dim u As Long
        u = UBound(vLines)
        For I = 0 To u
            taskDir = vLines(I)
            If FolderExists(taskDir) Then
                TaskNewFromDirectory taskDir
            Else
                taskDir = SubStringBetween(taskDir, "=", "", True)
                If FolderExists(taskDir) Then TaskNewFromDirectory taskDir
            End If
            taskDir = ""
        Next
    Else
        Close #fNum
        fNum = FreeFile
        Open vFilename For Input As #fNum
        Do Until EOF(fNum)
            taskDir = ""
            Line Input #fNum, taskDir
            If FolderExists(taskDir) Then
                TaskNewFromDirectory taskDir
            Else
                taskDir = SubStringBetween(taskDir, "=", "", True)
                If FolderExists(taskDir) Then TaskNewFromDirectory taskDir
            End If
        Loop
    End If
    Close #fNum
    TimerEnabled = True
    Windows_LockWindow 0
    '<EhFooter>
    Exit Sub

TaskListReadFrom_Error:
    Debug.Print "TaskListReadFrom:" & Err.Description
    On Error Resume Next
    Close #fNum
    TimerEnabled = True
    Windows_LockWindow 0
    Err.Clear

    '</EhFooter>
End Sub

'CSEH: ErrExit
Public Sub CheckTaskIncoming()
    '<EhHeader>
    On Error GoTo CheckTaskIncoming_Err
    '</EhHeader>
    If FileExists(mIncomingFile2) Then
    ElseIf FileExists(mIncomingFile) Then
        Name mIncomingFile As mIncomingFile2
    Else
        Exit Sub
    End If
    TaskListReadFrom mIncomingFile2
    Kill mIncomingFile2
    TimerEnabled = True
    '<EhFooter>
    Exit Sub

CheckTaskIncoming_Err:
    Debug.Print "CheckTaskIncoming:" & Err.Description
    On Error Resume Next
    Kill mIncomingFile2
    TimerEnabled = True
    Err.Clear

    '</EhFooter>
End Sub



