Attribute VB_Name = "MTaskman"
Option Explicit

Private Const CSTTaskIncomeFilename As String = "Incoming.lst"
Private mIncomingFile As String

Public Sub CallBack_EditTask(ByRef vtask As CTask)
'    vtask.UpdateFolder
'    UpdateTaskItem vtask, , True
'    vtask.Changed = True
'    vtask.AutoSave
End Sub
Public Sub CallBack_AddTask(ByVal vName As String, vBookInfoArray() As String, Optional vPending As Boolean)
        'TimerEnabled = False
        Dim vtask As CTask
        Set vtask = AddTask(vBookInfoArray, vPending)
        vtask.Name = vName
        'If Not vtask Is Nothing Then vtask.Changed = True
        vtask.UpdateFolder
        vtask.Changed = True
        vtask.AutoSave
        
        Dim fNum As Integer
        On Error Resume Next
        
        'fExists = FileExists(mIncomingFile)
        If vtask.Directory <> "" Then
            If Not FolderExists(vtask.Directory) Then Exit Sub
            fNum = FreeFile
            Open mIncomingFile For Append As #fNum
            Print #fNum, vtask.Directory
            Close #fNum
        End If
        'TimerEnabled = True
End Sub

Private Function AddTask(vBookInfoArray() As String, Optional vPending As Boolean) As CTask
        Dim newTask As CTask
        Set newTask = New CTask
       ' Dim taskId As String
        'taskId = NewTaskID()
        newTask.bookInfo.LoadFromArray vBookInfoArray
        'AddTaskItem newTask, vPending, False
                
        Set AddTask = newTask
        ', "TASK" & index, newTask.newTask.Title & "_" & newTask.SSID
        'LstTasks.ListItems.Add "TASK" & index, tvwChild, "INFO" & index, "Paused"
        'UpdateTaskInfo newTask
End Function


Public Sub Main()
mIncomingFile = App.Path & "\" & CSTTaskIncomeFilename
Dim MainForm As frmTask
Set MainForm = New frmTask

With MainForm
    .EditTaskMode = False
    .Init Nothing
    .UpdateFromClipboard
    .Show 0
End With
Do Until MainForm.Visible = False
    DoEvents
Loop
Set MainForm = Nothing
End
End Sub
