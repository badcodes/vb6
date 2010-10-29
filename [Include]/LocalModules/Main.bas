Attribute VB_Name = "MMain"
Option Explicit

Public AppName As String
Public AppPath As String
Public AppConfigPath As String

Public Sub Main()
    AppName = App.ProductName
    AppPath = App.Path
    AppConfigPath = App.Path
    
    'Load frmMain
    frmMain.Show 1
    'Set MainForm = Nothing
    'Unload MainForm
    'Unload frmMain
End Sub

