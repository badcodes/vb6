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

Private mMainWindow As Form


Private Sub Class_Terminate()
    Set mMainWindow = Nothing
End Sub

Public Sub Run()
    Debug.Print CONST_APPNAME & " Start Runing"
    If Not mMainWindow Is Nothing Then mMainWindow.Show
    Debug.Print CONST_APPNAME & " Stop Runing"
End Sub


Public Sub SetWindow(Optional MainWindow As Form = Nothing)
    Set mMainWindow = MainWindow
    If Not mMainWindow Is Nothing Then
        mMainWindow.Caption = CONST_APPNAME
    End If
End Sub
