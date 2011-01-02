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
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public Sub TaskStatusChange(ByRef Task As CTask)

End Sub


Public Sub DownloadStatusChange(ByRef Task As CTask)

End Sub

Public Sub TaskComplete(ByRef Task As CTask)

End Sub
