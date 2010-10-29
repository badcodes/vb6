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

Const CONST_APPNAME As String = "Test LiNVBLiB"

#If Not fNoGui = 1 Then
Private mMainWindow As Form
#End If

Private Sub Class_Terminate()
    #If Not fNoGui = 1 Then
        Set mMainWindow = Nothing
    #End If
End Sub

Public Sub Run()
    Debug.Print CONST_APPNAME & " Start Runing"
    
    #If Not fNoGui = 1 Then
        If Not mMainWindow Is Nothing Then mMainWindow.Show
    #End If
    
    Dim a As CStringBuilder
    Set a = New CStringBuilder
    a.ChunkSize = 200
    Dim b As String
    Open "X:\Download\ÂÞËØ\1950ÂÞËØ£ºÐÒ¸£Ö®Â·.txt" For Input As #5
    Do While Not EOF(5)
    Line Input #5, b
    a.Appendline b
    Loop
    Close #5
    Debug.Print a
    Debug.Print a.Length
    
    
    
    Debug.Print CONST_APPNAME & " Stop Runing"
End Sub

#If Not fNoGui = 1 Then
Public Sub SetWindow(Optional MainWindow As Form = Nothing)
    Set mMainWindow = MainWindow
    If Not mMainWindow Is Nothing Then
        mMainWindow.Caption = CONST_APPNAME
    End If
End Sub
#End If
