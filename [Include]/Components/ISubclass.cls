VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ISubclass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum EMsgResponse
    emrConsume      ' Process instead of original WindowProc
    emrPostProcess  ' Process after original WindowProc
    emrPreprocess   ' Process before original WindowProc
End Enum

Public MsgResponse As EMsgResponse

Function WindowProc(ByVal hWnd As Long, _
                    ByVal iMsg As Long, _
                    ByVal wParam As Long, _
                    ByVal lParam As Long) As Long
End Function
'

