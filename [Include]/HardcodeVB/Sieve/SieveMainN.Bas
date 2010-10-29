Attribute VB_Name = "MSieveMainN"
Option Explicit

Sub Main()
    If App.StartMode = vbSModeStandalone Then
        Dim frmSieve As New FSieveBasExeN
        frmSieve.Show
    End If
End Sub

