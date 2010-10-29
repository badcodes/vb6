Attribute VB_Name = "MSubclass"
Option Explicit

Sub AttachMessage(iwp As ISubclass, ByVal hWnd As Long, _
                  ByVal iMsg As Long)
    MSubclass.AttachMessage iwp, hWnd, iMsg
End Sub

Sub DetachMessage(iwp As ISubclass, ByVal hWnd As Long, _
                  ByVal iMsg As Long)
    MSubclass.DetachMessage iwp, hWnd, iMsg
End Sub


