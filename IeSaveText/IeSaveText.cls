VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "loader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Public Sub saveit(fTmp As String)
Shell App.Path & "/IeSaveText.exe " & fTmp, vbNormalFocus
End Sub

Public Sub LoadHtm(fHtm As String)
Shell "explorer.exe " & fHtm, vbNormalFocus
End Sub
