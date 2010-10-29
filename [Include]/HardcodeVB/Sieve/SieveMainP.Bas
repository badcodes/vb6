Attribute VB_Name = "MSieveMainP"
Option Explicit

Sub Main()
    If App.StartMode = vbSModeStandalone Then
        Dim frmSieve As New FSieveBasExeP
        frmSieve.Show
    End If
End Sub

