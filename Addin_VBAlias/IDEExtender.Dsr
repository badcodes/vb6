VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} Connect 
   ClientHeight    =   8340
   ClientLeft      =   1740
   ClientTop       =   1545
   ClientWidth     =   15270
   _ExtentX        =   26935
   _ExtentY        =   14711
   _Version        =   393216
   Description     =   "Enhancements for fast coding"
   DisplayName     =   "Vb IDE Extender"
   AppName         =   "Visual Basic"
   AppVer          =   "Visual Basic 6.0"
   LoadName        =   "Startup"
   LoadBehavior    =   1
End
Attribute VB_Name = "Connect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Close Process
'Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
'Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
'Private Const PROCESS_ALL_ACCESS = &H1F0FFF
'Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long

'Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'Private Const WM_CLOSE = &H10
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long

'Private AddinProcessID As Long
'''''''''''''

Public FormDisplayed                 As Boolean
Public VBInstance                    As VBIDE.VBE

'the main form
Dim mfrmAddIn                        As New frmAddIn

'the submenu in the command bar
Dim CommandBarMenu                   As Office.CommandBarControl
Public WithEvents CommandBarMenu_ev  As CommandBarEvents          'command bar event handler
Attribute CommandBarMenu_ev.VB_VarHelpID = -1

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'the toolbar
Dim Toolbar1 As CommandBar

'toolbar button clear immediate window
Private ButtonDebug         As Office.CommandBarButton
Private WithEvents ButtonDebug_ev As CommandBarEvents
Attribute ButtonDebug_ev.VB_VarHelpID = -1

'toolbar button wipe minus designer
Private ButWipeND         As Office.CommandBarButton
Private WithEvents ButWipeND_ev As CommandBarEvents
Attribute ButWipeND_ev.VB_VarHelpID = -1

'toolbar button wipe minus active
Private ButWipeNA         As Office.CommandBarButton
Private WithEvents ButWipeNA_ev As CommandBarEvents
Attribute ButWipeNA_ev.VB_VarHelpID = -1

'toolbar button wipe all windows
Private ButWipeA         As Office.CommandBarButton
Private WithEvents ButWipeA_ev As CommandBarEvents
Attribute ButWipeA_ev.VB_VarHelpID = -1

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'popup menu in code panes, used by macros
Dim ButtonMacro As Office.CommandBarControl
Public WithEvents ButtonMacro_ev As CommandBarEvents 'event handler
Attribute ButtonMacro_ev.VB_VarHelpID = -1

'2 popup menus in code panes, used for line comments in and out
Dim MenuComment As Office.CommandBarControl
Public WithEvents MenuComment_ev As CommandBarEvents 'event handler
Attribute MenuComment_ev.VB_VarHelpID = -1

Dim MenuUncomment As Office.CommandBarControl
Public WithEvents MenuUncomment_ev As CommandBarEvents 'event handler
Attribute MenuUncomment_ev.VB_VarHelpID = -1

'popup menu in code panes, used to swap 2 parts of a line around an equal sign
Dim SwapByEqual As Office.CommandBarControl
Public WithEvents SwapByEqual_ev As CommandBarEvents 'event handler
Attribute SwapByEqual_ev.VB_VarHelpID = -1

'Get IDE status
Private WithEvents eVBBuildEvents As VBBuildEvents
Attribute eVBBuildEvents.VB_VarHelpID = -1

'working variables
Private bClosingSession As Boolean
Public StartingInst As Boolean
Private mPosSizes As Collection

Private Const mstrGuid As String = "{4C929F58-C633-4b3a-9810-F613490AC250}"

Private bHostShutdown As Boolean
Private bUserClosed As Boolean
Private hModule As Long
Public Sub SwapAroundEqual()
    Dim sl As Long
    Dim el As Long
    Dim sc As Long
    Dim ec As Long
    Dim i As Long
    Dim sLine As String
    Dim pos As Long
    Dim item As String, item1 As String

    'get the current selection
    VBInstance.ActiveCodePane.GetSelection sl, sc, el, ec

    If sl > 0 Then ' make sure we have at least one line selected
        For i = sl To el
            sLine = VBInstance.ActiveCodePane.CodeModule.Lines(i, 1)
            pos = InStr(sLine, Chr(61))

            If pos Then
                item = Mid(sLine, 1, pos - 1)
                item1 = Mid(sLine, pos + 1, Len(sLine))
                sLine = Trim$(item1) & Chr(32) & Chr(61) & Chr(32) & Trim$(item)
            End If

            VBInstance.ActiveCodePane.CodeModule.ReplaceLine i, sLine
        Next i
    End If
End Sub
Public Sub BulkComment(Optional bRemove As Boolean = False)

    On Error Resume Next

    Dim sl As Long
    Dim el As Long
    Dim sc As Long
    Dim ec As Long
    Dim i As Long
    Dim sLine As String

    'get the current selection
    VBInstance.ActiveCodePane.GetSelection sl, sc, el, ec

    If Not bRemove Then
        'add comment
        If sl > 0 Then ' make sure we have at least one line selected
            For i = sl To el
                sLine = VBInstance.ActiveCodePane.CodeModule.Lines(i, 1)
                sLine = "'" & sLine
                VBInstance.ActiveCodePane.CodeModule.ReplaceLine i, sLine
            Next i
        End If
    Else
        For i = sl To el  ' make sure we have at least one line selected
            sLine = VBInstance.ActiveCodePane.CodeModule.Lines(i, 1)
            If Left(sLine, 1) = "'" Then
                sLine = Right$(sLine, Len(sLine) - 1)
                VBInstance.ActiveCodePane.CodeModule.ReplaceLine i, sLine
            End If
        Next i
    End If

    VBInstance.ActiveCodePane.SetSelection el, 1, el, 1
End Sub
Sub LoadFrmAddin()
    'On Error Resume Next

    If mfrmAddIn Is Nothing Then
        Set mfrmAddIn = New frmAddIn
        If (mfrmAddIn.VBInstance Is Nothing) Then Set mfrmAddIn.VBInstance = VBInstance
        If (mfrmAddIn.IDEExt Is Nothing) Then Set mfrmAddIn.IDEExt = Me
    End If

    mfrmAddIn.Hide
    FormDisplayed = False
End Sub
Sub AddMenusAndButtons()

    Dim oMenu As CommandBar                          'popup fast macros
    Dim cbMenuCommandBar As Office.CommandBarControl 'submenu
    Dim cbMenu As Object

    'save and restore window positions ad size
    'Dim nCmdBarControl As CommandBarControl
    'Dim nAddInCmdBar As CommandBar

    'macros, popup codepane menu
    Set oMenu = VBInstance.CommandBars("Code Window")
    Set ButtonMacro = oMenu.Controls.Add(msoControlButton, , , 1)
    Clipboard.Clear
    Clipboard.SetData LoadResPicture(10001, 0)
    ButtonMacro.PasteFace
    ButtonMacro.Caption = "&Fast Macro"
    ButtonMacro.ToolTipText = "Selected text as macro "
    Set ButtonMacro_ev = VBInstance.Events.CommandBarEvents(ButtonMacro)

    'Menu Comment
    'Set oMenu = VBInstance.CommandBars("Code Window")
    Set MenuComment = oMenu.Controls.Add(msoControlButton, , , 2)
    MenuComment.Caption = "&Comment Selection"
    MenuComment.ToolTipText = "Comment selected lines "
    Set MenuComment_ev = VBInstance.Events.CommandBarEvents(MenuComment)

    'Menu Uncomment
    'Set oMenu = VBInstance.CommandBars("Code Window")
    Set MenuUncomment = oMenu.Controls.Add(msoControlButton, , , 3)
    MenuUncomment.Caption = "&Uncomment Selection"
    MenuUncomment.ToolTipText = "Uncomment selected lines "
    Set MenuUncomment_ev = VBInstance.Events.CommandBarEvents(MenuUncomment)

    'Menu SwapByEqual
    'Set oMenu = VBInstance.CommandBars("Code Window")
    Set SwapByEqual = oMenu.Controls.Add(msoControlButton, , , 4)
    SwapByEqual.Caption = "Swap &Member Sides"
    SwapByEqual.ToolTipText = "Swap member sides around sign equal"
    Set SwapByEqual_ev = VBInstance.Events.CommandBarEvents(SwapByEqual)
    
End Sub
Sub PrepareShutdown()
    bSpaceEnhancer = False
    ShowLinkedWindows
    UnRefLinkedWindows
    k = 0
    'UnhookMDIClient
    'UnhookMainWindow
    frmAddIn.ReadyToClose = True
End Sub

Sub SetObjectsAndVariables()

    'set the objects
    Set modMacros.VBInstance = VBInstance
    Set modSubclass.VBInstance = VBInstance
    Set mfrmAddIn.VBInstance = VBInstance

    'start preparing the hook on the linked windows
    IDEhwnd = VBInstance.MainWindow.hwnd 'Not listed but endorsed!

End Sub

Sub UnloadMenusAndButtons()

    On Error Resume Next

    CommandBarMenu.Delete
    Set CommandBarMenu = Nothing
    Set CommandBarMenu_ev = Nothing

    ButtonMacro.Delete
    Set ButtonMacro = Nothing
    Set ButtonMacro_ev = Nothing

    MenuComment.Delete
    Set MenuComment = Nothing
    Set MenuComment_ev = Nothing

    MenuUncomment.Delete
    Set MenuUncomment = Nothing
    Set MenuUncomment_ev = Nothing

    SwapByEqual.Delete
    Set SwapByEqual = Nothing
    Set SwapByEqual_ev = Nothing

End Sub

Sub UnloadVariablesAndClasses()

    'On Error Resume Next

    Dim Form As Form

    If FormDisplayed Then
        SaveSetting App.Title, "Settings", "DisplayOnConnect", "1"
        FormDisplayed = False
    Else
        SaveSetting App.Title, "Settings", "DisplayOnConnect", "0"
    End If


    Set mfrmAddIn.VBInstance = Nothing
    Set mfrmAddIn.IDEExt = Nothing

    Set eVBBuildEvents = Nothing

    Set Ini = Nothing

    For Each Form In Forms
        If Not Form Is Nothing Then
            Unload Form
            Set Form = Nothing
        End If
    Next Form

    Set mPosSizes = Nothing
    Set modMacros.VBInstance = Nothing
    Set modSubclass.VBInstance = Nothing

End Sub

Private Sub SaveToolbarPosition()
    ' routine to save current position of toolbar
    On Error Resume Next
    With VBInstance.CommandBars("exToolbar")
        SaveSetting App.Title, "Settings", "exToolbarPosition", CStr(.Position)
        SaveSetting App.Title, "Settings", "exToolbarIndex", CStr(.RowIndex)
        SaveSetting App.Title, "Settings", "exToolbarLeft", CStr(.Left)
        SaveSetting App.Title, "Settings", "exToolbarTop", CStr(.Top)

        If .Visible = True Then
            SaveSetting App.Title, "Settings", "exToolbarVisible", "True"
        Else
            SaveSetting App.Title, "Settings", "exToolbarVisible", "False"
        End If
    End With
End Sub 'SaveToolbarPosition()

Private Sub SetToolbarPosition()
    ' routine to get the last saved position of toolbar and put the toolbar there
    Dim ToolbarPosition As Long
    Dim ToolbarIndex As Long
    Dim ToolbarTop As Long
    Dim ToolbarLeft As Long

    On Error Resume Next
    ToolbarPosition = CLng(GetSetting(App.Title, "Settings", "exToolbarPosition", CStr(msoBarTop)))
    ToolbarIndex = CLng(GetSetting(App.Title, "Settings", "exToolbarIndex", "0"))
    ToolbarLeft = CLng(GetSetting(App.Title, "Settings", "exToolbarLeft", "0"))
    ToolbarTop = CLng(GetSetting(App.Title, "Settings", "exToolbarTop", "0"))

    With VBInstance.CommandBars("exToolbar")
        .Position = ToolbarPosition
        Select Case ToolbarPosition
            Case msoBarTop, msoBarBottom
                .RowIndex = ToolbarIndex
                .Left = ToolbarLeft
            Case msoBarLeft, msoBarRight
                .RowIndex = ToolbarIndex
                .Top = ToolbarTop
            Case Else
                .Top = ToolbarTop
                .Left = ToolbarLeft
        End Select

        If GetSetting(App.Title, "Settings", "exToolbarVisible", "True") = "True" Then
            .Visible = True
        Else
            .Visible = False
        End If
    End With

End Sub 'GetToolbarPosition()
Function AddToolbar()
    'Disable Tool Bar by xiaoranzzz
    Exit Function
    

    Set Toolbar1 = VBInstance.CommandBars.Add("exToolbar", msoBarFloating, , True)

    With Toolbar1.Controls
        'debug window
        Set ButtonDebug = .Add(msoControlButton)
        With ButtonDebug
            .Caption = "ClearDebug"
            .ToolTipText = "Clear Immediate window"
            .Style = msoButtonIcon
            .BeginGroup = True
            .State = msoButtonUp
            Clipboard.Clear
            Clipboard.SetData LoadResPicture(10006, 0)
            .PasteFace
        End With 'ButtonDebug
        Set ButtonDebug_ev = VBInstance.Events.CommandBarEvents(ButtonDebug)

        'ButWipeNoDesigner
        Set ButWipeND = .Add(msoControlButton)
        With ButWipeND
            .Caption = "ButWipeND"
            .ToolTipText = "Close all windows, Designer excluded"
            .Style = msoButtonIcon
            .BeginGroup = True
            .State = msoButtonUp
            Clipboard.Clear
            Clipboard.SetData LoadResPicture(10007, 0)
            .PasteFace
        End With 'ButWipeND
        Set ButWipeND_ev = VBInstance.Events.CommandBarEvents(ButWipeND)

        'ButWipeNoActive
        Set ButWipeNA = .Add(msoControlButton)
        With ButWipeNA
            .Caption = "ButWipeNA"
            .ToolTipText = "Close all windows, active pane excluded"
            .Style = msoButtonIcon
            .BeginGroup = False
            .State = msoButtonUp
            Clipboard.Clear
            Clipboard.SetData LoadResPicture(10008, 0)
            .PasteFace
        End With 'ButWipeNA
        Set ButWipeNA_ev = VBInstance.Events.CommandBarEvents(ButWipeNA)

        'ButWipeA
        Set ButWipeA = .Add(msoControlButton)

        With ButWipeA
            .Caption = "ButWipeA"
            .ToolTipText = "Close all windows"
            .Style = msoButtonIcon
            .BeginGroup = False
            .State = msoButtonUp
            Clipboard.Clear
            Clipboard.SetData LoadResPicture(10009, 0)
            .PasteFace
        End With 'ButWipeA
        Set ButWipeA_ev = VBInstance.Events.CommandBarEvents(ButWipeA)

        SetToolbarPosition
    End With



End Function


Sub Hide()

    On Error Resume Next

    FormDisplayed = False
    mfrmAddIn.Hide

End Sub

Sub Show()

    On Error Resume Next

    If mfrmAddIn Is Nothing Then
        Set mfrmAddIn = New frmAddIn
    End If

    Set mfrmAddIn.VBInstance = VBInstance
    Set mfrmAddIn.IDEExt = Me
    FormDisplayed = True
    mfrmAddIn.Show

End Sub

Private Sub UnloadToolbar()
    'Disable toolbar by xiaoranzzz
    Exit Sub
    
    On Error Resume Next
    
    SaveToolbarPosition

    ButWipeA.Delete
    ButWipeNA.Delete
    ButWipeND.Delete
    ButtonDebug.Delete

    Toolbar1.Delete
    
    
    Set ButtonDebug_ev = Nothing
    Set ButWipeND_ev = Nothing
    Set ButWipeNA_ev = Nothing
    Set ButWipeA_ev = Nothing

    Set Toolbar1 = Nothing
End Sub

Private Sub AddinInstance_OnBeginShutdown(custom() As Variant)
    'if the user closes the ide or sends Alt F4, what is the same thing

    On Error Resume Next

    bSpaceEnhancer = False
    ShowLinkedWindows
    UnRefLinkedWindows
    UnhookMDIClient
    UnhookMainWindow

    k = 0
    bClosingSession = True
End Sub


'------------------------------------------------------
'this method adds the Add-In to VB
'------------------------------------------------------
Private Sub AddinInstance_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)
    On Error GoTo error_handler
    Dim nEvents2 As Events2

    Set VBInstance = Application
    SetObjectsAndVariables

    If ConnectMode = ext_cm_External Then
        'Used by the wizard toolbar to start this wizard
        'Me.Show
        MsgBox "This Add-In must be activated from the Add-Ins menu."
    Else
        'submenu on commandbar
        Set CommandBarMenu = AddToAddInCommandBar("VB Alias")
        Set CommandBarMenu_ev = VBInstance.Events.CommandBarEvents(CommandBarMenu)

        'detect if IDE in run, compiling or design mode
        Set nEvents2 = Application.Events
        Set eVBBuildEvents = nEvents2.VBBuildEvents

        LoadFrmAddin
        AddToolbar
        AddMenusAndButtons

    End If

    If ConnectMode = ext_cm_AfterStartup Then
        'if we open a project in a already existing IDE
        'start the hooks if not already working
        HookMDIClient 'now used by macros and stack, must be always started
        If bSpaceEnhancer = True Then HookMainWindow 'used only by space enhancer

    End If

    'Call GetWindowThreadProcessId(mfrmAddIn.hwnd, AddinProcessID)
    hModule = GetModuleHandle("VbAlias.dll")

    Exit Sub

error_handler:

    MsgBox Err.Description

End Sub

'------------------------------------------------------
'this method removes the Add-In from VB
'------------------------------------------------------
Private Sub AddinInstance_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant)
    'Dim ProcessID As Long, hProg As Long, cExit As Long
    On Error Resume Next

    PrepareShutdown
    UnloadToolbar
    UnloadMenusAndButtons
    UnloadVariablesAndClasses

    If RemoveMode = vbext_dm_HostShutdown Then
        bHostShutdown = True
    ElseIf RemoveMode = vbext_dm_UserClosed Then
        bUserClosed = True
    End If

    '    Call GetWindowThreadProcessId(IDEhwnd, ProcessID)
    '    hProg = OpenProcess(PROCESS_ALL_ACCESS, False, ProcessID)
    '    Call TerminateProcess(hProg, cExit)
    'End If
    'Dim ProcessID As Long, hProg As Long, cExit As Long
    'On Error Resume Next
    'If RemoveMode = vbext_dm_UserClosed Then
    'If RemoveMode = vbext_dm_HostShutdown Then
    '    hProg = OpenProcess(PROCESS_ALL_ACCESS, False, AddinProcessID)
    '    Call TerminateProcess(hProg, cExit)
    'End If


    Unload Me


End Sub


Private Sub AddinInstance_OnStartupComplete(custom() As Variant)
    DoEvents
    If Not Len(VBInstance.ActiveVBProject.FileName) = 0 Then _
            mfrmAddIn.MaximiseOnStartup
End Sub

Private Sub AddinInstance_Terminate()

    On Error Resume Next

    Set VBInstance = Nothing

    If bHostShutdown Then
        FreeLibrary hModule
    ElseIf bUserClosed Then

    End If

End Sub

Private Sub ButtonDebug_ev_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)

    Dim oDebugWindow As Window

    For Each oDebugWindow In VBInstance.ActiveVBProject.VBE.Windows
        If oDebugWindow.Type = vbext_wt_Immediate Then
            If oDebugWindow.Visible = False Then
                oDebugWindow.Visible = True
            End If

            oDebugWindow.SetFocus
            SendKeys "^{Home}", True
            SendKeys "^+{End}", True
            SendKeys "{Del}", True
            Exit For
        End If
    Next oDebugWindow

    DoEvents 'not necessary, but solid

    For Each oDebugWindow In VBInstance.ActiveVBProject.VBE.Windows
        If oDebugWindow.Type = vbext_wt_Immediate Then
            oDebugWindow.Close
            Exit For
        End If
    Next oDebugWindow

End Sub

Private Sub ButtonMacro_ev_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)

    On Error GoTo SomeError

    If VBInstance.ActiveVBProject Is Nothing Then Exit Sub
    If VBInstance.ActiveCodePane Is Nothing Then Exit Sub
    If VBInstance.ActiveCodePane.CodeModule Is Nothing Then Exit Sub
    If VBInstance.ActiveCodePane.CodeModule.CodePane Is Nothing Then Exit Sub

    If mfrmAddIn.WLCheck1.Value = wlChecked Then
        mfrmAddIn.PutSelectedText
        IsFastMacro = True
    End If

    Exit Sub
SomeError:
    MsgBox "ButtonMacro_ev_Click"

End Sub





Private Sub ButWipeA_ev_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)

    Dim pWindow As VBIDE.Window

    For Each pWindow In VBInstance.Windows
        If (pWindow.Type = vbext_wt_CodeWindow) Or _
                (pWindow.Type = vbext_wt_Designer) Then
            pWindow.Close
        End If
    Next pWindow

End Sub


Private Sub ButWipeNA_ev_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)

    Dim pWindow As VBIDE.Window
    Dim strCaption As String

    strCaption = VBInstance.ActiveCodePane.Window.Caption

    For Each pWindow In VBInstance.Windows
        If (pWindow.Type = vbext_wt_CodeWindow) Or _
                (pWindow.Type = vbext_wt_Designer) Then
            If pWindow.Caption <> strCaption Then pWindow.Close
        End If
    Next pWindow

End Sub


Private Sub ButWipeND_ev_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)

    Dim pWindow As VBIDE.Window

    For Each pWindow In VBInstance.Windows
        If pWindow.Type = vbext_wt_CodeWindow Then
            pWindow.Close
        End If
    Next pWindow
End Sub


'this event fires when the menu is clicked in the IDE
Private Sub CommandBarMenu_ev_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)

    Me.Show
End Sub

Function AddToAddInCommandBar(sCaption As String) As Office.CommandBarControl
    Dim cbMenuCommandBar As Office.CommandBarControl  'command bar object
    Dim cbMenu As Object

    On Error GoTo AddToAddInCommandBarErr

    'see if we can find the Add-Ins menu
    Set cbMenu = VBInstance.CommandBars("Add-Ins")
    If cbMenu Is Nothing Then
        'not available so we fail
        Exit Function
    End If

    'add it to the command bar
    Set cbMenuCommandBar = cbMenu.Controls.Add(1)
    'set the caption
    cbMenuCommandBar.Caption = sCaption

    Set AddToAddInCommandBar = cbMenuCommandBar

    Exit Function

AddToAddInCommandBarErr:

End Function








Private Sub eVBBuildEvents_EnterDesignMode()
    bInRunMode = False
End Sub

Private Sub eVBBuildEvents_EnterRunMode()
    bInRunMode = True
End Sub


Private Sub MenuComment_ev_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    BulkComment False
End Sub


Private Sub MenuUncomment_ev_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    BulkComment True
End Sub


Private Sub SwapByEqual_ev_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
SwapAroundEqual
End Sub


