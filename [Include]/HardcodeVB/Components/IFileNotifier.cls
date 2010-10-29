VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IFileNotifier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Enum EFILE_NOTIFY
    FILE_NOTIFY_CHANGE_FILE_NAME = &H1
    FILE_NOTIFY_CHANGE_DIR_NAME = &H2
    FILE_NOTIFY_CHANGE_ATTRIBUTES = &H4
    FILE_NOTIFY_CHANGE_SIZE = &H8
    FILE_NOTIFY_CHANGE_LAST_WRITE = &H10
    FILE_NOTIFY_CHANGE_SECURITY = &H100
End Enum

' Interface for CFileNotify to communicate with its clients
Sub Change(sDir As String, efn As EFILE_NOTIFY, fSubTree As Boolean)

End Sub
'

