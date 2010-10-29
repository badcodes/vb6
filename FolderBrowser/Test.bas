Attribute VB_Name = "MTest"
Option Explicit

Public Sub Test()
    Dim a As CFolderBrowser
    Set a = New CFolderBrowser
    With a
    .InitDirectory = ""
    .Owner = 0
    End With
    Debug.Print a.Browse
End Sub
