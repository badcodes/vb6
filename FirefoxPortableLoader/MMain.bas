Attribute VB_Name = "MMain"
Option Explicit

Sub Main()
    Dim profile As String
    profile = Command$
    
    If profile = "-h" Then
        MsgBox "Usage:" & vbCrLf & vbCrLf & App.EXEName & ".exe {Profile Directory}" & vbCrLf & vbCrLf & _
            "Copyright 2008 xiaoranzzz@myplace", vbInformation
        Exit Sub
    End If
    
        Dim directory As String
        directory = BuildPath(App.Path)
        If FileExists(directory & "FirefoxPortable.exe") = False Then
            MsgBox "文件:" & directory & "FirefoxPortable.exe" & vbCrLf & "不存在!", vbCritical
            Exit Sub
        End If
        
    If profile <> "" Then

        
        Dim iniHnd As CLiNInI
        Set iniHnd = New CLiNInI
        iniHnd.source = directory & "FirefoxPortable.ini"
        iniHnd.SaveSetting "FirefoxPortable", "ProfileDirectory", profile
        iniHnd.Save
    End If
    Shell directory & "FirefoxPortable.exe"
    
End Sub
