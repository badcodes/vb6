VERSION 5.00
Begin VB.UserControl FileBox 
   ClientHeight    =   5910
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4770
   PropertyPages   =   "FileBox.ctx":0000
   ScaleHeight     =   5910
   ScaleWidth      =   4770
   ToolboxBitmap   =   "FileBox.ctx":0023
   Begin VB.DriveListBox MyDrive 
      Height          =   315
      Left            =   -15
      TabIndex        =   2
      Top             =   15
      Width           =   4740
   End
   Begin VB.DirListBox MyPath 
      Height          =   2565
      Left            =   0
      TabIndex        =   1
      Top             =   345
      Width           =   4725
   End
   Begin VB.FileListBox MyFile 
      Height          =   2610
      Left            =   15
      TabIndex        =   0
      Top             =   2910
      Width           =   4725
   End
End
Attribute VB_Name = "FileBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Event Declarations:
Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyDown
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Event KeyPress(KeyAscii As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyPress
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyUp
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
Event PathChange(Path As String)
Event FilenameChange(FileName As String)
Event DriveChange(Drive As String)
Event Change()
Attribute Change.VB_Description = "Occurs when the contents of a control have changed."
Event FileClick(FileName As String)
Event FileDoubleClick(FileName As String)
Event PathClick(Path As String)

'Event PathChange()
'Event FilenameChange()
'Event DriveChange()
'Default Property Values:
'Const m_def_FilePanel = 0
'Const m_def_FilePanel = 1
'Const m_def_Hidden = 0
'Property Variables:
'Dim m_FilePanel As Boolean
'Dim m_FilePanel As Boolean
'Dim m_Hidden As Boolean



Public Property Get FileSelected(ByVal vIndex As Long) As Boolean
    FileSelected = MyFile.Selected(vIndex)
End Property

Public Property Let FileSelected(ByVal vIndex As Long, ByVal vYes As Boolean)
    On Error Resume Next
    MyFile.Selected(vIndex) = vYes
    RaiseEvent FilenameChange(MyFile.List(vIndex))
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    MyDrive.BackColor() = New_BackColor
    MyPath.BackColor() = New_BackColor
    MyFile.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = UserControl.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    UserControl.ForeColor() = New_ForeColor
    MyDrive.ForeColor = New_ForeColor
    MyPath.ForeColor = New_ForeColor
    MyFile.ForeColor = New_ForeColor
    
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    MyDrive.Enabled = New_Enabled
    MyPath.Enabled = New_Enabled
    MyFile.Enabled = New_Enabled
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.Font = New_Font
    Set MyDrive.Font = New_Font
    Set MyFile.Font = New_Font
    Set MyPath.Font = New_Font
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackStyle
'CSEH: ErrExit
Public Property Get BackStyle() As Integer
Attribute BackStyle.VB_Description = "Indicates whether a Label or the background of a Shape is transparent or opaque."
    '<EhHeader>
    On Error GoTo BackStyle_Err
    '</EhHeader>
    BackStyle = UserControl.BackStyle
    '<EhFooter>
    Exit Property

BackStyle_Err:
    Err.Clear

    '</EhFooter>
End Property

Public Property Let BackStyle(ByVal New_BackStyle As Integer)
    '<EhHeader>
    On Error GoTo BackStyle_Err
    '</EhHeader>
    UserControl.BackStyle() = New_BackStyle
    PropertyChanged "BackStyle"
    '<EhFooter>
    Exit Property

BackStyle_Err:
    Err.Clear

    '</EhFooter>
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BorderStyle
Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    On Error Resume Next
    UserControl.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
    UserControl.Refresh
    MyDrive.Refresh
    MyPath.Refresh
    MyFile.Refresh
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Appearance
Public Property Get Appearance() As Integer
Attribute Appearance.VB_Description = "Returns/sets whether or not an object is painted at run time with 3-D effects."
    Appearance = UserControl.Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As Integer)
    UserControl.Appearance() = New_Appearance
    MyDrive.Appearance = UserControl.Appearance
    MyPath.Appearance = UserControl.Appearance
    MyFile.Appearance = UserControl.Appearance
    PropertyChanged "Appearance"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=MyDrive,MyDrive,-1,Drive
Public Property Get Drive() As String
Attribute Drive.VB_Description = "Returns/sets the selected drive at run time."
    Drive = MyDrive.Drive
End Property

Public Property Let Drive(ByVal New_Drive As String)
    On Error Resume Next
    MyDrive.Drive = New_Drive
    RaiseEvent DriveChange(MyDrive.Drive)
    'MyPath.Path = MyDrive.Drive
    PropertyChanged "Drive"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=MyFile,MyFile,-1,FileName
Public Property Get FileName() As String
Attribute FileName.VB_Description = "Returns/sets the path and filename of a selected file."
    FileName = MyFile.FileName
End Property

Public Property Let FileName(ByVal New_FileName As String)
    MyFile.FileName() = New_FileName
    RaiseEvent FilenameChange(MyFile.FileName)
    PropertyChanged "FileName"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,FillColor
Public Property Get FillColor() As OLE_COLOR
Attribute FillColor.VB_Description = "Returns/sets the color used to fill in shapes, circles, and boxes."
    FillColor = UserControl.FillColor
End Property

Public Property Let FillColor(ByVal New_FillColor As OLE_COLOR)
    UserControl.FillColor() = New_FillColor
    PropertyChanged "FillColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,FillStyle
Public Property Get FillStyle() As Integer
Attribute FillStyle.VB_Description = "Returns/sets the fill style of a shape."
    FillStyle = UserControl.FillStyle
End Property

Public Property Let FillStyle(ByVal New_FillStyle As Integer)
    UserControl.FillStyle() = New_FillStyle
    PropertyChanged "FillStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hDC
Public Property Get hdc() As Long
Attribute hdc.VB_Description = "Returns a handle (from Microsoft Windows) to the object's device context."
    hdc = UserControl.hdc
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hWnd
Public Property Get hwnd() As Long
Attribute hwnd.VB_Description = "Returns a handle (from Microsoft Windows) to an object's window."
    hwnd = UserControl.hwnd
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=MyFile,MyFile,-1,MultiSelect
Public Property Get MultiSelect() As Integer
Attribute MultiSelect.VB_Description = "Returns/sets a value that determines whether a user can make multiple selections in a control."
    MultiSelect = MyFile.MultiSelect
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=MyPath,MyPath,-1,Path
Public Property Get Path() As String
Attribute Path.VB_Description = "Returns/sets the current path."
    Path = MyPath.Path
End Property

Public Property Let Path(ByVal New_Path As String)
    MyPath.Path() = New_Path
    RaiseEvent PathChange(MyPath.Path)
    RaiseEvent Change
    Drive = New_Path
    PropertyChanged "Path"
End Property

Private Sub MyDrive_Change()
    RaiseEvent DriveChange(MyDrive.Drive)
    MyPath.Path = MyDrive.Drive
End Sub

Private Sub MyFile_Click()
    RaiseEvent FilenameChange(MyFile.FileName)
    RaiseEvent FileClick(MyFile.FileName)
    RaiseEvent Change
End Sub

Private Sub MyFile_DblClick()
    RaiseEvent FileDoubleClick(MyFile.FileName)
    RaiseEvent FilenameChange(MyFile.FileName)
    RaiseEvent Change
End Sub

Private Sub MyPath_Change()
    MyFile.Path = MyPath.Path
    MyFile.Refresh
    'RaiseEvent FilenameChange(MyFile.Filename)
    RaiseEvent PathChange(MyPath.Path)
    RaiseEvent Change
End Sub

Private Sub MyPath_Click()
    RaiseEvent PathClick(MyPath.Path)
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    Set UserControl.Font = Ambient.Font
    Set MyDrive.Font = Ambient.Font
    Set MyPath.Font = Ambient.Font
    Set MyFile.Font = Ambient.Font
'    m_Hidden = m_def_Hidden
'    m_FilePanel = m_def_FilePanel
    'm_FilePanel = m_def_FilePanel
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    Enabled = PropBag.ReadProperty("Enabled", True)
    Font = PropBag.ReadProperty("Font", Ambient.Font)
    BackStyle = PropBag.ReadProperty("BackStyle", 1)
    BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
    Appearance = PropBag.ReadProperty("Appearance", 1)
    Drive = PropBag.ReadProperty("Drive", "")
    FileName = PropBag.ReadProperty("FileName", "")
    FillColor = PropBag.ReadProperty("FillColor", &H0&)
    FillStyle = PropBag.ReadProperty("FillStyle", 1)
    Path = PropBag.ReadProperty("Path", "")
    MyFile.Archive = PropBag.ReadProperty("ShowArchive", True)
'    m_Hidden = PropBag.ReadProperty("Hidden", m_def_Hidden)
    'MyFile.ListIndex = PropBag.ReadProperty("ListIndex", 0)
    MyFile.Normal = PropBag.ReadProperty("ShowNormal", True)
    MyFile.System = PropBag.ReadProperty("ShowSystem", False)
    MyFile.ReadOnly = PropBag.ReadProperty("ShowReadOnly", True)
    MyFile.Hidden = PropBag.ReadProperty("ShowHidden", False)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", 0)
'    m_FilePanel = PropBag.ReadProperty("FilePanel", m_def_FilePanel)
    MyFile.Visible = PropBag.ReadProperty("FilePanel", True)
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    Dim vTop As Long
    Dim vHeight As Long
    vTop = 0
    vHeight = UserControl.Height
    If MyDrive.Visible Then
        MyDrive.Move 0, 0, UserControl.Width
        vTop = MyDrive.Top + MyDrive.Height
        vHeight = vHeight - vTop
    End If
    If MyPath.Visible And MyFile.Visible Then vHeight = vHeight / 2
    If MyPath.Visible Then
        MyPath.Move 0, vTop, UserControl.Width, vHeight
        vTop = MyPath.Top + MyPath.Height
    End If
    If MyFile.Visible Then
        MyFile.Move 0, vTop, UserControl.Width, vHeight
    End If
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("ForeColor", UserControl.ForeColor, &H80000012)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("BackStyle", UserControl.BackStyle, 1)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 0)
    Call PropBag.WriteProperty("Appearance", UserControl.Appearance, 1)
    Call PropBag.WriteProperty("Drive", MyDrive.Drive, "")
    Call PropBag.WriteProperty("FileName", MyFile.FileName, "")
    Call PropBag.WriteProperty("FillColor", UserControl.FillColor, &H0&)
    Call PropBag.WriteProperty("FillStyle", UserControl.FillStyle, 1)
    Call PropBag.WriteProperty("Path", MyPath.Path, "")
    Call PropBag.WriteProperty("Archive", MyFile.Archive, True)
'    Call PropBag.WriteProperty("Hidden", m_Hidden, m_def_Hidden)
   ' Call PropBag.WriteProperty("ListIndex", MyFile.ListIndex, 0)
    Call PropBag.WriteProperty("ShowNormal", MyFile.Normal, True)
    Call PropBag.WriteProperty("ShowSystem", MyFile.System, False)
    Call PropBag.WriteProperty("ShowReadOnly", MyFile.ReadOnly, True)
    Call PropBag.WriteProperty("ShowHidden", MyFile.Hidden, False)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)
'    Call PropBag.WriteProperty("FilePanel", m_FilePanel, m_def_FilePanel)
    Call PropBag.WriteProperty("FilePanel", MyFile.Visible, True)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=MyFile,MyFile,-1,Archive
Public Property Get ShowArchive() As Boolean
Attribute ShowArchive.VB_Description = "Determines whether a FileListBox control displays files with Archive attributes."
    ShowArchive = MyFile.Archive
End Property

Public Property Let ShowArchive(ByVal New_Archive As Boolean)
    MyFile.Archive() = New_Archive
    PropertyChanged "ShowArchive"
End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MemberInfo=0,0,0,0
'Public Property Get Hidden() As Boolean
'    Hidden = m_Hidden
'End Property
'
'Public Property Let Hidden(ByVal New_Hidden As Boolean)
'    m_Hidden = New_Hidden
'    PropertyChanged "Hidden"
'End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=MyFile,MyFile,-1,List
Public Property Get Files(ByVal Index As Integer) As String
Attribute Files.VB_Description = "Returns/sets the items contained in a control's list portion."
    Files = MyFile.List(Index)
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=MyFile,MyFile,-1,ListCount
Public Property Get FileCount() As Integer
Attribute FileCount.VB_Description = "Returns the number of items in the list portion of a control."
    FileCount = MyFile.ListCount
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=MyFile,MyFile,-1,ListIndex
Public Property Get FileIndex() As Integer
Attribute FileIndex.VB_Description = "Returns/sets the index of the currently selected item in the control."
    FileIndex = MyFile.ListIndex
End Property

Public Property Let FileIndex(ByVal New_ListIndex As Integer)
    MyFile.ListIndex() = New_ListIndex
    PropertyChanged "FileIndex"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=MyDrive,MyDrive,-1,List
Public Property Get Drives(ByVal Index As Integer) As String
    Drives = MyDrive.List(Index)
End Property
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=MyPath,MyPath,-1,List
Public Property Get Paths(ByVal Index As Integer) As String
    Paths = MyPath.List(Index)
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=MyPath,MyPath,-1,ListCount
Public Property Get PathCount() As Integer
    PathCount = MyPath.ListCount
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=MyPath,MyPath,-1,ListIndex
Public Property Get PathIndex() As Integer
    PathIndex = MyPath.ListIndex
End Property

Public Property Let PathIndex(ByVal New_ListIndex As Integer)
    MyPath.ListIndex() = New_ListIndex
    PropertyChanged "PathIndex"
End Property
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=MyDrive,MyDrive,-1,ListCount
Public Property Get DriveCount() As Integer
    DriveCount = MyDrive.ListCount
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=MyDrive,MyDrive,-1,ListIndex
Public Property Get DriveIndex() As Integer
    DriveIndex = MyDrive.ListIndex
End Property

Public Property Let DriveIndex(ByVal New_ListIndex As Integer)
    MyDrive.ListIndex() = New_ListIndex
    PropertyChanged "DriveIndex"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=MyFile,MyFile,-1,Normal
Public Property Get Normal() As Boolean
Attribute Normal.VB_Description = "Determines whether a FileListBox control displays files with Normal attributes."
    ShowNormal = MyFile.Normal
End Property

Public Property Let ShowNormal(ByVal New_Normal As Boolean)
    MyFile.Normal() = New_Normal
    PropertyChanged "ShowNormal"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=MyFile,MyFile,-1,System
Public Property Get ShowSystem() As Boolean
Attribute ShowSystem.VB_Description = "Determines whether a FileListBox control displays files with System attributes."
    ShowSystem = MyFile.System
End Property

Public Property Let ShowSystem(ByVal New_System As Boolean)
    MyFile.System() = New_System
    PropertyChanged "ShowSystem"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=MyFile,MyFile,-1,ReadOnly
Public Property Get ShowReadOnly() As Boolean
Attribute ShowReadOnly.VB_Description = "Returns/sets a value that determines whether files with read-only attributes are displayed in the file list or not."
    ShowReadOnly = MyFile.ReadOnly
End Property

Public Property Let ShowReadOnly(ByVal New_ReadOnly As Boolean)
    MyFile.ReadOnly() = New_ReadOnly
    PropertyChanged "ShowReadOnly"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=MyFile,MyFile,-1,Hidden
Public Property Get ShowHidden() As Boolean
Attribute ShowHidden.VB_Description = "Determines whether a FileListBox control displays files with Hidden attributes."
    ShowHidden = MyFile.Hidden
End Property

Public Property Let ShowHidden(ByVal New_Hidden As Boolean)
    MyFile.Hidden() = New_Hidden
    PropertyChanged "ShowHidden"
End Property

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,MouseIcon
Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "Sets a custom mouse icon."
    Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set UserControl.MouseIcon = New_MouseIcon
    Set MyDrive.MouseIcon = New_MouseIcon
    Set MyFile.MouseIcon = New_MouseIcon
    Set MyPath.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseMove(Button, Shift, x, y)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,MousePointer
Public Property Get MousePointer() As Integer
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object."
    MousePointer = UserControl.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As Integer)
    UserControl.MousePointer() = New_MousePointer
    MyDrive.MousePointer = New_MousePointer
    MyPath.MousePointer = New_MousePointer
    MyFile.MousePointer = New_MousePointer
    PropertyChanged "MousePointer"
End Property

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseUp(Button, Shift, x, y)
End Sub
 
'
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get FilePanel() As Boolean
    FilePanel = MyFile.Visible
End Property

Public Property Let FilePanel(ByVal New_FilePanel As Boolean)
    MyFile.Visible = New_FilePanel
    UserControl_Resize
    PropertyChanged "FilePanel"
End Property

