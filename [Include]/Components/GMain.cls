VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GMain"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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


