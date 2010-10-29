Attribute VB_Name = "modSubclass"
Option Explicit

Public VBInstance                   As VBIDE.VBE 'this has the instantiated application object

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const IDX_WINDOWPROC        As Long = -4

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hwnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Private Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Const PropName              As String = "Hooked"

Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As Point) As Long
Private Type Point
    X As Long
    Y As Long
End Type

'Private Enum ArrayIndexes
'    idxCompoName = 0
'    idxMemberName = 1
'    idxMemberScope = 2
'End Enum

Public CursorPos                    As Point

'Private KeyWords                    As Collection
'Private CodeMembers                 As Collection

'the classes to work with
Private Const VBA_WINDOW          As String = "VbaWindow"           ' hide LnkWnds if mouse over
'Private Const THUNDER_FORM        As String = "ThunderForm"         ' show LnkWnds if mouse over
'Private Const WNDCLASS_DESKED_GSK As String = "wndclass_desked_gsk" ' show LnkWnds if mouse over
'Private Const DESIGNER_WINDOW     As String = "DesignerWindow"      ' show LnkWnds if mouse over
'Private Const MSO_COMMANDBAR      As String = "MsoCommandBar"
'Private Const VBA_IMMEDIATE       As String = "Immediate"
Private Const VBA_COMBOBOX          As String = "ComboBox"

'the events to retrive by hook

Private Const WM_DEVICECHANGE = &H219
Private Const WM_VSCROLL = &H115
Private Const WM_NOTIFY = &H4E
Private Const WM_SETTEXT = &HC


Private Const WM_PARENTNOTIFY = &H210
Private Const WM_NCMOUSEMOVE       As Long = &HA0
Public Const WM_MOUSEMOVE          As Long = &H200
'Public Const WM_LBUTTONDOWN        As Long = &H201
Public Const WM_LBUTTONUP          As Long = &H202
'Public Const WM_RBUTTONDOWN        As Long = &H204
'Public Const WM_RBUTTONUP          As Long = &H205
'Public Const WM_MBUTTONDOWN        As Long = &H207
'Public Const WM_MBUTTONUP          As Long = &H208

Private Const WM_SETFOCUS           As Long = 7
Private Const WM_KILLFOCUS          As Long = 8
'Private Const WM_KEYDOWN            As Long = &H100
Private Const WM_CHAR               As Long = &H102
'Private Const WM_MOUSEWHEEL         As Long = &H20A
Private Const WM_MDIACTIVATE        As Long = &H222

Private hWndMDIClient               As Long
Private hWndCodePane                As Long
Private UserTypedCode               As Boolean
Public IDEhwnd                      As Long

'»'Send Mail
'»Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'»Private Const SW_SHOWNORMAL         As Long = 1
'»Private Const SE_NO_ERROR           As Long = 33 'Values below 33 are error returns
'»
'»'Registry
'»Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
'»Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType, lpData As Any, lpcbData As Long) As Long
'»Private Declare Sub RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long)
'»
'»Private Const KEY_QUERY_VALUE       As Long = 1
'»Private Const REG_OPTION_RESERVED   As Long = 0
'»Private Const ERROR_NONE            As Long = 0
'»
'»'Registry scroll setting access keys...
'»Private Const HKEY_CURRENT_USER     As Long = &H80000001
'»Private Const DesktopSettings       As String = "Control Panel\Desktop"
'»Private Const SmoothScroll          As String = "SmoothScroll"
'»Private Const WheelScrollLines      As String = "WheelScrollLines"
'»
'»'...and our own settings...
'»Public Const sOptions               As String = "Options"
'»Public Const sLines                 As String = "Lines"
'»Public Const sMode                  As String = "Mode"
'»Public Const sSmooth                As String = "Smooth"
'»Public Const sInstant               As String = "Instant"
'»Public Const sAutoComplete          As String = "AutoComplete"
'»Public Const sNoisy                 As String = "Noisy"
'»Public Const sTriggerLength         As String = "TriggerLength"
'»Public Const sOn                    As String = "On"
'»Public Const sOff                   As String = "Off"
'»
'»'..and finally what we got (or didn't get) from the Registry or from our own options
'»Public LinesToScroll                As String
'»Public Smooth                       As Long
'»Public AutoComplete                 As Boolean
'»Public Noisy                        As Boolean
'»Public TriggerLength                As Long 'the minimum length of word fragment to trigger autocomplete
'»
'»'- - - - - - - - - - - - - - - - - - - - - - - - modify both values to correspond- - - - -
'»Public Const opHpCapt               As String = "Half a &Page"
'»Private Const ScrollFraction        As Single = 1 / 2 'fraction of page to scroll
'»'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Public MacrosEnabled                As Boolean
Public bSpaceEnhancer               As Boolean
Public bParentNotify                As Boolean
Public bClosingSession              As Boolean
Public bInRunMode                   As Boolean
Public Function ClassNameFromWnd(ByVal hwnd As Long) As String

'a well known wrapper
Dim sName As String, cName As Long

sName = String$(80, 0)
cName = GetClassName(hwnd, sName, 80)
ClassNameFromWnd = Left$(sName, cName)
End Function

Public Function IDEMode(vbInst As VBIDE.VBE) As Long
'==================================
'Returns the mode the IDE is in:
'   vbext_vm_Run = "Run mode"
'   vbext_vm_Break = "Break Mode"
'   vbext_vm_Design = "Design Mode"
'==================================
Dim lMode As Long
Const cRun = "Run"
Const cEnd = "End"
Const cBreak = "Break"

lMode = vbext_vm_Design

On Error GoTo ErrH
If vbInst.CommandBars(cRun).Controls(cEnd).Enabled = True Then
' The IDE is at least in run mode
    lMode = vbext_vm_Run
    If vbInst.CommandBars(cRun).Controls(cBreak).Enabled = False Then
' The IDE is in Break mode
        lMode = vbext_vm_Break
    End If
End If
IDEMode = lMode
Exit Function
ErrH:

End Function

Public Function IsDesignMode(vbInst As VBIDE.VBE) As Boolean
'==================================
'Returns the mode the IDE is in:
'   vbext_vm_Run = "Run mode"
'   vbext_vm_Break = "Break Mode"
'   vbext_vm_Design = "Design Mode"
'==================================
Dim DesignMode As Boolean
Const cRun = "Run"
Const cEnd = "End"


DesignMode = True

On Error GoTo ErrH
If vbInst.CommandBars(cRun).Controls(cEnd).Enabled = True Then DesignMode = False
IsDesignMode = DesignMode


Exit Function
ErrH:

End Function
Private Function CodePaneProc(ByVal hwnd As Long, ByVal nMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long


'Window Procedure for the active IDE Codepane
CodePaneProc = CallWindowProc(GetProp(hwnd, PropName), hwnd, nMsg, wParam, lParam)

If bInRunMode = False Then
    Select Case nMsg
    Case WM_CHAR 'the user has typed a char
        If MacrosEnabled Then
            Select Case wParam
            Case vbKeySpace
                InsertMacro
            End Select
        End If
    Case WM_MOUSEMOVE
        If bSpaceEnhancer Then
            If LinkedWindowsVisible Then
                HideLinkedWindows
            End If
        End If
'    Case WM_PARENTNOTIFY
'        If wParam = 513 Then
'
'            bParentNotify = True
'        End If
    End Select
End If

End Function

Private Sub HookCodePane()

On Error GoTo eH
With VBInstance
    If .ActiveWindow Is .ActiveCodePane.Window Then
        hWndCodePane = FindWindowEx(hWndMDIClient, 0, VBA_WINDOW, .ActiveWindow.Caption)
        If hWndCodePane Then
            If GetProp(hWndCodePane, PropName) = 0 Then
                SetProp hWndCodePane, PropName, GetWindowLong(hWndCodePane, IDX_WINDOWPROC)
                SetWindowLong hWndCodePane, IDX_WINDOWPROC, AddressOf CodePaneProc
            End If
        End If
    End If
End With 'VBINSTANCE

Exit Sub
eH:
'MsgBox "HookCodePane"
End Sub

Public Sub HookMDIClient()

On Error Resume Next
hWndMDIClient = FindWindowEx(VBInstance.MainWindow.hwnd, 0, "MDIClient", vbNullString)
If hWndMDIClient Then
    If GetProp(hWndMDIClient, PropName) = 0 Then
        SetProp hWndMDIClient, PropName, GetWindowLong(hWndMDIClient, IDX_WINDOWPROC)
        SetWindowLong hWndMDIClient, IDX_WINDOWPROC, AddressOf MDIClientProc
    End If
    HookCodePane
End If
On Error GoTo 0

End Sub
Public Sub HookMainWindow()

On Error Resume Next

If IDEhwnd Then
    If GetProp(IDEhwnd, PropName) = 0 Then
        SetProp IDEhwnd, PropName, GetWindowLong(IDEhwnd, IDX_WINDOWPROC)
        SetWindowLong IDEhwnd, IDX_WINDOWPROC, AddressOf MainWindowProc
    End If
End If

HookCodePane
On Error GoTo 0

End Sub
Private Function MDIClientProc(ByVal hwnd As Long, ByVal nMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long


'Window procedure for VB's MDIClient window

MDIClientProc = CallWindowProc(GetProp(hwnd, PropName), hwnd, nMsg, wParam, lParam)  'call the original winproc to do what has to be done

If bInRunMode = False Then
    With VBInstance
        Select Case nMsg 'and now split on message type
        Case WM_KILLFOCUS 'this codepane just lost the focus (remember - the original procedure has already been performed)
            UnhookCodePane
        Case WM_MDIACTIVATE, WM_SETFOCUS 'another codepane has been (re)activated by the user
            HookCodePane
        End Select
    End With 'VBINSTANCE
End If

End Function
Private Function MainWindowProc(ByVal hwnd As Long, ByVal nMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

'Window procedure for VB's MainWindow window

MainWindowProc = CallWindowProc(GetProp(hwnd, PropName), hwnd, nMsg, wParam, lParam)  'call the original winproc to do what has to be done

If bInRunMode = False Then
    Select Case nMsg
    Case WM_NCMOUSEMOVE
        If bSpaceEnhancer Then
            If LinkedWindowsVisible = False Then
                ShowLinkedWindows
            End If
        End If
'    Case WM_LBUTTONUP
'        Dim sTmp As String
'        'If ClassNameFromWnd(hWnd) = VBA_COMBOBOX Then
'        sTmp = ClassNameFromWnd(hWnd)
'        MsgBox sTmp
'        'End If
'        'AddProcedureToStack
    End Select
End If

End Function




Private Sub UnhookCodePane()

If (bSpaceEnhancer = False) And (MacrosEnabled = False) Then
    If hWndCodePane Then
        SetWindowLong hWndCodePane, IDX_WINDOWPROC, GetProp(hWndCodePane, PropName)
        RemoveProp hWndCodePane, PropName
        hWndCodePane = 0
    End If
End If
End Sub

Public Sub UnhookMainWindow()

If IDEhwnd Then
    UnhookCodePane
    SetWindowLong IDEhwnd, IDX_WINDOWPROC, GetProp(IDEhwnd, PropName)
    RemoveProp IDEhwnd, PropName 'remove the property
End If
End Sub
Public Sub UnhookMDIClient()

If hWndMDIClient Then
    UnhookCodePane
    SetWindowLong hWndMDIClient, IDX_WINDOWPROC, GetProp(hWndMDIClient, PropName)
    RemoveProp hWndMDIClient, PropName 'remove the property
    hWndMDIClient = 0
End If
End Sub

