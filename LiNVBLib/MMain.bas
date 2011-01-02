Attribute VB_Name = "MMain"

Public Sub Main()
    Dim theApp As CApp
    Set theApp = New CApp
    
#If Not fNoGui = 1 Then
    Dim theWindow As frmMain
    Set theWindow = New frmMain
    theApp.SetWindow theWindow
#End If
   
    theApp.Run
    
    Set theApp = Nothing
#If Not fNoGui = 1 Then
    Set theWindow = Nothing
#End If
End Sub
