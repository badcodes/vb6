VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CShortcut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum EErrorShortcut
    eeBaseShortcut = 13210  ' CShortcut
End Enum

Private link As New CShellLink

 ' Destination constants
Enum EDestination
    edstDesktop
    edstCommonDesktop
    edstPrograms
    edstCommonPrograms
    edstStartMenu
    edstPath
    edstCurrent
End Enum

' Show command constants (a subset of constants for SetWindowPos API)
Enum EDisplayMode
    edmNormal = SW_NORMAL
    edmMinimized = SW_SHOWMINNOACTIVE
    edmMaximized = SW_MAXIMIZE
End Enum

'' Properties

'' Path of file represented by shortcut
Property Get Path() As String
    Dim fd As WIN32_FIND_DATA, s As String
    s = String$(cMaxPath, 0)
    link.GetPath s, cMaxPath, fd, SLGP_UNCPRIORITY
    Path = MUtility.StrZToStr(s)
End Property

Property Let Path(sPathA As String)
    ' Make sure file exists
    If Not MUtility.ExistFile(sPathA) Then ErrRaise eeFileNotFound
    link.SetPath sPathA
End Property

'' Startup directory for shortcut target
Property Get WorkingDirectory() As String
    Dim s As String
    s = String$(cMaxPath, 0)
    link.GetWorkingDirectory s, cMaxPath
    WorkingDirectory = MUtility.StrZToStr(s)
End Property

Property Let WorkingDirectory(sWorkingA As String)
    link.SetWorkingDirectory sWorkingA
End Property

' Shortcut dialog ignores description, but we can save and restore it
Property Get Description() As String
    Dim s As String
    s = String$(cMaxPath, 0)
    link.GetDescription s, cMaxPath
    Description = MUtility.StrZToStr(s)
End Property

Property Let Description(sDescription As String)
    link.SetDescription sDescription
End Property

'' Arguments for shortcut target
Property Get Arguments() As String
    Dim s As String
    s = String$(cMaxPath, 0)
    link.GetArguments s, cMaxPath
    Arguments = MUtility.StrZToStr(s)
End Property

Property Let Arguments(sArgumentsA As String)
    link.SetArguments sArgumentsA
End Property

'' Display command can be Normal, Minimized, or Maximized
Property Get DisplayMode() As EDisplayMode
    DisplayMode = link.showCmd
End Property

Property Let DisplayMode(edm As EDisplayMode)
    ' IShellLink doesn't handle all SW_ constants, but we do
    Select Case edm
    Case SW_HIDE, SW_NORMAL, SW_SHOWNOACTIVATE, SW_SHOW, _
         SW_SHOWNA, SW_RESTORE, SW_SHOWDEFAULT
        ' Convert all these to normal: 0, 1, 4, 5, 8, 9, 10
        edm = edmNormal
    Case SW_SHOWMINIMIZED, SW_MINIMIZE, SW_SHOWMINNOACTIVE
        ' Convert all these to minimized: 2, 6, 7
        edm = edmMinimized
    Case SW_MAXIMIZE
        ' Pass maximize through: 3
        edm = edmMaximized
    Case Else
        ' Convert anything else to normal
        edm = edmNormal
    End Select
    link.showCmd = edm
End Property

Property Get HotKey() As KeyCodeConstants
    HotKey = link.HotKey
End Property

Property Let HotKey(kcc As KeyCodeConstants)
    link.HotKey = kcc
End Property

Property Get Icon() As Variant
    Dim s As String, i As Long, hIcon As Long
    s = String$(cMaxPath, 0)
    link.GetIconLocation s, cMaxPath, i
    hIcon = ExtractIcon(App.hInstance, s, i)
    Set Icon = MPicTool.IconToPicture(hIcon)
End Property

Property Let Icon(vIcon As Variant)
    If VarType(vIcon) = vbString Then
        ' Assume icon file (index 0)
        link.SetIconLocation CStr(vIcon), 0
    Else
        ' Assume index into embedded EXE
        link.SetIconLocation Path, CLng(vIcon)
    End If
End Property

' Link file parameter is Variant to accept any of these:
'      edstDesktop          - Put on desktop
'      edstCommonDesktop    - Put on shared desktop
'      edstPrograms         - Put on programs menu
'      edstCommonPrograms   - Put on shared programs menu
'      edstStartMenu        - Put on start menu
'      edstCurrent          - Put in current directory
'      edstPath             - Put in same directory as target file
'      [directory]          - Put in hardcoded path
'      [file.LNK]           - Put in hardcoded file
Function Save(vLinkFile As Variant) As String
    Dim sLink As String
    ' Convert constant or directory to full path
    sLink = FixLocation(vLinkFile)
    If sLink = sEmpty Then ErrRaise eeFileNotFound

    ' Save the object to disk
    MCasts.IVBPersistFile(link).Save sLink, APITRUE
    Save = sLink
End Function

' Flags control behavior if LNK file reference can't be resolved:
'    SLR_ANY_MATCH - Display a dialog (with hWnd parameter as parent
'                    window) asking user whether to search for reference
'    SLR_NO_UI     - Search the disk for the time period specified by
'                    TimeOut parameter
Sub Resolve(sFileA As String, _
            Optional Flags As ESLR = SLR_ANY_MATCH, _
            Optional hWnd As Long = hNull, _
            Optional TimeOut As Integer = 0)
    ' Load from LNK file and resolve
    MCasts.IVBPersistFile(link).Load sFileA, STGM_DIRECT
    If Flags = SLR_NO_UI And TimeOut > 0 Then
        Flags = Flags Or MBytes.LShiftDWord(TimeOut, 16)
    End If
    link.Resolve hWnd, Flags
End Sub


' Location of .LNK file. Output is always full path of .LNK file
Private Function FixLocation(vLocationA As Variant) As String
    Dim s As String, sPath As String

    ' If user passes in string, save it internally
    If VarType(vLocationA) = vbString Then
        ' Convert to a full path
        s = MUtility.GetFullPath(CStr(vLocationA))
        ' If this is already a link file, return it
        If UCase$(MUtility.GetFileExt(s)) = ".LNK" Then
            FixLocation = s
            Exit Function
        Else
            ' Can't use a location that doesn't exist
            If Not MUtility.ExistFile(s) Then Exit Function

            ' Make sure directory ends with backslash
            s = MUtility.NormalizePath(s)
        End If
    Else

        ' If location hasn't been set, we can't do anything
        sPath = Path
        If sPath = sEmpty Then Exit Function
    
        ' Create a directory from setting
        Select Case vLocationA
        Case edstCurrent:
            ' Current directory
            s = CurDir$
        Case edstPath:
            ' Directory of shortcut target
            s = MUtility.GetFileDir(Path)
        Case edstPath:
        
        Case edstCommonDesktop:
            s = MRegTool.GetCommonDesktop
        
        Case edstPrograms:
            s = MRegTool.GetPrograms
        
        Case edstCommonPrograms:
            s = MRegTool.GetCommonPrograms
        
        Case edstStartMenu:
            s = MRegTool.GetStartMenu
        
        Case Else ' Includes edstDesktop and any invalid arguments
            ' Desktop directory
            s = MRegTool.GetDesktop
        End Select
    End If
    
    ' Combine directory, path name, and LNK extension
    FixLocation = s & MUtility.GetFileBase(sPath) & ".LNK"
    
End Function
'

#If fComponent = 0 Then
Private Sub ErrRaise(e As Long)
    Dim sText As String, sSource As String
    If e > 1000 Then
        sSource = App.ExeName & ".Shortcut"
        Select Case e
        Case eeBaseShortcut
            BugAssert True
       ' Case ee...
       '     Add additional errors
        End Select
        Err.Raise COMError(e), sSource, sText
    Else
        ' Raise standard Visual Basic error
        sSource = App.ExeName & ".VBError"
        Err.Raise e, sSource
    End If
End Sub
#End If

