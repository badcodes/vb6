VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CApp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Const CONST_APPNAME As String = "LProject"

Private mfrmMain As frmMain

Public Sub Initialize()
End Sub

Public Sub Show()
    mfrmMain.Show
End Sub

Private Sub Class_Initialize()
    Set mfrmMain = New frmMain
    mfrmMain.Caption = CONST_APPNAME
End Sub

Private Sub Class_Terminate()
    Set mfrmMain = Nothing
End Sub
