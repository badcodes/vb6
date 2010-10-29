VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CArchive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Property Get title() As String

End Property

Public Property Get author() As String

End Property

Public Property Get publiser() As String

End Property

Public Property Get createdDate() As String

End Property

Public Property Get catalog() As String

End Property

Public Property Get defaultFile() As String

End Property

Public zvShowLeft As zhtmVisiablity
Public zvShowMenu As zhtmVisiablity
Public zvShowStatusBar As zhtmVisiablity

Public Property Get ShowLeft() As ZZZState

End Property

Public Property Get ShowMenu() As ZZZState
End Property

Public Property Get ShowStatusBar() As ZZZState
End Property

Public Property Get entries() As String()
End Property

