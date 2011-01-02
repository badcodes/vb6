VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ITaskNotify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public Sub TaskStatusChange(ByRef Runner As CTaskRunner)

End Sub


Public Sub DownloadStatusChange(ByRef Runner As CTaskRunner)

End Sub

Public Sub TaskComplete(ByRef Runner As CTaskRunner)

End Sub

Public Sub StatusChange(ByRef vText As String)

End Sub

Public Sub UpdateJpgBook(ByRef vTask As CTask)

End Sub
