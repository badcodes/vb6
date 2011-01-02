VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CWindowToFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IWindowsHelper

Public FileNumber As Integer

Private Function IWindowsHelper_DoWindow(ByVal iLevel As Integer, _
                                         ByVal hWnd As Long) As Long

    BugAssert hWnd <> hNull
    ' Ignore desktop window (level -1)
    If iLevel >= 0 Then
        ' Write data to log file
        Print #FileNumber, GetWndInfo(hWnd, iLevel)
    End If
    ' Always successful
    IWindowsHelper_DoWindow = hNull
End Function
'

