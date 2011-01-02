VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} Connect 
   ClientHeight    =   6240
   ClientLeft      =   1740
   ClientTop       =   1548
   ClientWidth     =   8244
   _ExtentX        =   14542
   _ExtentY        =   11007
   _Version        =   393216
   Description     =   "openDir of curfile"
   DisplayName     =   "openDir"
   AppName         =   "Visual Basic"
   AppVer          =   "Visual Basic 98 (ver 6.0)"
   LoadName        =   "Command Line / Startup"
   LoadBehavior    =   5
   RegLocation     =   "HKEY_CURRENT_USER\Software\Microsoft\Visual Basic\6.0"
   CmdLineSupport  =   -1  'True
End
Attribute VB_Name = "Connect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public VBInstance             As VBIDE.VBE
Dim mcbMenuCommandBar         As Office.CommandBarControl
'Dim mfrmAddIn                 As New frmAddIn
Public WithEvents MenuHandler As CommandBarEvents          'command bar event handler
Attribute MenuHandler.VB_VarHelpID = -1
Public WithEvents ProjectHandler As VBProjectsEvents
Attribute ProjectHandler.VB_VarHelpID = -1
'Public WithEvents addinMenu As CommandBarEvents


'------------------------------------------------------
'this method adds the Add-In to VB
'------------------------------------------------------
Private Sub AddinInstance_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)
    On Error GoTo error_handler
    
    'save the vb instance
    Set VBInstance = Application
    
    'this is a good place to set a breakpoint and
    'test various addin objects, properties and methods


        Set Me.ProjectHandler = VBInstance.Events.VBProjectsEvents
        Set mcbMenuCommandBar = AddToAddInCommandBar("打开工程目录")
        'sink the event
        Set Me.MenuHandler = VBInstance.Events.CommandBarEvents(mcbMenuCommandBar)
        'Set Me.addinMenu = VBInstance.Events.CommandBarEvents(VBInstance.CommandBars("Add-Ins"))

  
    Exit Sub
    
error_handler:
    
    MsgBox Err.Description
    
End Sub

'------------------------------------------------------
'this method removes the Add-In from VB
'------------------------------------------------------
Private Sub AddinInstance_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant)
    On Error Resume Next
    
    'delete the command bar entry
    mcbMenuCommandBar.Delete

End Sub


'this event fires when the menu is clicked in the IDE
Private Sub MenuHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
On Error Resume Next
    Dim curFile As String
    curFile = filePath
    If curFile = "" Then Exit Sub
    Shell "explorer.exe " & curFile, vbMaximizedFocus
End Sub

Function AddToAddInCommandBar(sCaption As String) As Office.CommandBarControl
    Dim cbMenuCommandBar As Office.CommandBarControl  'command bar object
    Dim cbMenu As Object
  
    On Error GoTo AddToAddInCommandBarErr
    
    'see if we can find the Add-Ins menu
    Set cbMenu = VBInstance.CommandBars("Tools")
    If cbMenu Is Nothing Then
        'not available so we fail
        Exit Function
    End If
    
    'add it to the command bar
    Set cbMenuCommandBar = cbMenu.Controls.Add(Office.MsoBarType.msoBarTypeMenuBar)
    'set the caption
    With cbMenuCommandBar
    cbMenuCommandBar.Caption = sCaption
    cbMenuCommandBar.Move cbMenu, cbMenu.Controls.Count
    cbMenuCommandBar.BeginGroup = True
    End With
    
    Set AddToAddInCommandBar = cbMenuCommandBar
    
    Exit Function
    
AddToAddInCommandBarErr:

End Function

Private Function filePath() As String
    filePath = GetParentFolderName(VBInstance.ActiveVBProject.FileName)
End Function

Private Sub ProjectHandler_ItemAdded(ByVal VBProject As VBIDE.VBProject)
    Dim projectPath As String
    projectPath = filePath()
    If Len(projectPath) > 2 Then
        ChDrive Left$(projectPath, 2)
        ChDir projectPath
    End If
End Sub
