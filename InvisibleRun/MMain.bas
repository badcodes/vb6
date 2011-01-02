Attribute VB_Name = "MMain"
Option Explicit
Private Const cstLogFile As String = "Log.txt"
Private LogFile As String
Public Sub Main()
    LogFile = App.Path & "\" & App.ProductName & ".txt"
    On Error GoTo ErrorMain
    Dim cmd As String
    cmd = Command$
    If cmd <> "" Then
        LogAndRun cmd
    End If
    Exit Sub
ErrorMain:
    MsgBox Err.Description, vbOKOnly, App.ProductName
    Err.Clear
End Sub

Public Sub LogAndRun(cmd As String)
    On Error Resume Next
    Dim fNum As Integer
    fNum = FreeFile()
    Open LogFile For Append As #fNum
    Print #fNum, "["; Date$ & " " & Time$ & "] " & cmd
    Close #fNum
    Shell cmd, vbHide
End Sub
