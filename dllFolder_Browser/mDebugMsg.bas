Attribute VB_Name = "mDebugMsg"
Option Explicit

#Const DebugMsgLevel = 1

Public Sub DebugMsg(ByVal sMsg As String)
   #If DebugMsgLevel = 0 Then
   #ElseIf DebugMsgLevel = 1 Then
      Debug.Print sMsg
   #Else
      MsgBox sMsg, vbInformation
   #End If
End Sub

