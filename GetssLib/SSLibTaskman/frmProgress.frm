VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProgress 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Downloading"
   ClientHeight    =   2160
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   11175
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   11175
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin MSComctlLib.ListView lstTaskInfo 
      Height          =   2190
      Left            =   0
      TabIndex        =   0
      Top             =   -15
      Width           =   11355
      _ExtentX        =   20029
      _ExtentY        =   3863
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "线程"
         Object.Width           =   1806
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "正在下载"
         Object.Width           =   5644
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
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "URL"
         Object.Width           =   10583
      EndProperty
   End
End
Attribute VB_Name = "frmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements ITaskNotify
Public MainApp As ITaskNotify
Public Task As CTask

Private Sub Form_Activate()
    If Not Task Is Nothing Then
        Me.Hide
        Me.Enabled = False
        Me.StartTask Task
        If Not MainApp Is Nothing Then MainApp.TaskComplete Nothing
        Set Task = Nothing
    End If
End Sub

'public property Set

Private Sub Form_Resize()
On Error Resume Next

    If Me.ScaleHeight < 1 Then Exit Sub
    'ME.Enabled = False
    
'    lsttasks.Top = cmdAdd.Top + cmdAdd.Height + 120
'    lsttasks.Width = ME.ScaleWidth - 2 * lsttasks.Left
'    lsttasks.Height = ME.ScaleHeight - StatusBar.Height - lsttasks.Top
    
    'StatusBar.Top = Me.ScaleHeight - StatusBar.Height
    
    lstTaskInfo.Move 0, 0, Me.ScaleWidth, Me.ScaleWidth
    
'    lstTaskInfo.Left = 0
'    lstTaskInfo.Top = 0 '
'
'    'StatusBar.Top -lstTaskInfo.Height
'    lstTaskInfo.Width = Me.ScaleWidth - 2 * lstTaskInfo.Left
'
'    LstTasks.Left = 0
'    LstTasks.Top = cmdAdd.Top + cmdAdd.Height + 120
'    LstTasks.Width = Me.ScaleWidth - 2 * LstTasks.Left
'    LstTasks.Height = lstTaskInfo.Top - LstTasks.Top
End Sub

Private Sub ITaskNotify_DownloadStatusChange(vtask As CTask)
If Not MainApp Is Nothing Then MainApp.TaskStatusChange vtask
If Me.Enabled = False Then Exit Sub
'Static timeLastUpdate As Single
'Dim CurrentTime As Single
'CurrentTime = DateTime.Timer
'If CurrentTime - timeLastUpdate > 0.8 Then
'    UpdateTaskItem vTask, vItem
'End If
'timeLastUpdate = CurrentTime

If vtask Is Nothing Then Exit Sub

Dim i As Long
Dim c As Long
c = vtask.ConnectionsCount

'Me.Enabled = False
For i = 1 To c

Dim Progress As CDownloadProgress
Set Progress = vtask.Connections(i)
If Progress Is Nothing Then GoTo ContinueFor
'If Not fDoUpdating Then Exit Sub
    DoEvents

Dim itemTask As ListItem
If lstTaskInfo.ListItems.count < i Then
    Set itemTask = lstTaskInfo.ListItems.Add()
Else
    Set itemTask = lstTaskInfo.ListItems.item(i)
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
 itemTask.SubItems(4) = Progress.URL
ContinueFor:
Next

DoEvents

'Me.Enabled = True


End Sub

Private Sub ITaskNotify_TaskComplete(Task As CTask)

End Sub

Private Sub ITaskNotify_TaskStatusChange(Task As CTask)
    If Not MainApp Is Nothing Then MainApp.TaskStatusChange Task
    If Me.Enabled = False Then Exit Sub
End Sub

Private Sub lblDirectory_Click()
End Sub

Public Sub StartTask(Task As CTask)
    Me.Caption = "Downloading " & Task.Directory
    If Me.Enabled = False Then Task.Init MainApp Else Task.Init Me
    'Task.Init Me
    Task.Status = STS_START
    Task.StartDownload
    Set Task = Nothing
    Unload Me
End Sub
