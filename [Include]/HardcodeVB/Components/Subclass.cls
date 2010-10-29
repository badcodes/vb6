VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GSubclass"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Sub AttachMessage(iwp As ISubclass, ByVal hWnd As Long, _
                  ByVal iMsg As Long)
    MSubclass.AttachMessage iwp, hWnd, iMsg
End Sub

Sub DetachMessage(iwp As ISubclass, ByVal hWnd As Long, _
                  ByVal iMsg As Long)
    MSubclass.DetachMessage iwp, hWnd, iMsg
End Sub


