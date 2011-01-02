Attribute VB_Name = "MTestSysMenu"
Option Explicit

Public procOld As Long
Public Const IDM_ABOUT As Long = 1010

#If 0 Then  ' Enable if you don't have the type library
Public Declare Function CallWindowProc Lib "user32" _
    Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, _
    ByVal hwnd As Long, ByVal Msg As Long, _
    ByVal wParam As Long, lParam As Any) As Long
#End If

Public Function SysMenuProc(ByVal hwnd As Long, ByVal iMsg As Long, _
                            ByVal wParam As Long, lParam As Long) As Long
    ' Ignore everything but system commands
    If iMsg = WM_SYSCOMMAND Then
        ' Check for one special menu item
        If wParam = IDM_ABOUT Then
            MsgBox "What are you talking about?"
            Exit Function
        End If
    End If
    ' Let old window procedure handle other messages
    SysMenuProc = CallWindowProc(procOld, hwnd, iMsg, wParam, lParam)
End Function
'


