VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GDlgOpenDir"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Brought to you by:
'  Brad Martinez
'  btmtz@aol.com
'  http://members.aol.com/btmtz/vb

'///////////////////////////////////////////////////////////////////////////////////////////////////////////

' A little info...
' Objects in the shell’s namespace are assigned item identifiers and item
' identifier lists. An item identifier uniquely identifies an item within its parent
' folder. An item identifier list uniquely identifies an item within the shell’s
' namespace by tracing a path to the item from the desktop.

'///////////////////////////////////////////////////////////////////////////////////////////////////////////

' An item identifier is defined by the variable-length SHITEMID structure.
' The first two bytes of this structure specify its size, and the format of
' the remaining bytes depends on the parent folder, or more precisely
' on the software that implements the parent folder’s IShellFolder interface.
' Except for the first two bytes, item identifiers are not strictly defined, and
' applications should make no assumptions about their format.
Type SHITEMID   ' mkid
    cb As Long       ' Size of the ID (including cb itself)
    abID() As Byte  ' The item ID (variable length)
End Type
Public Const MAX_PATH = 260

' The ITEMIDLIST structure defines an element in an item identifier list
' (the only member of this structure is an SHITEMID structure). An item
' identifier list consists of one or more consecutive ITEMIDLIST structures
' packed on byte boundaries, followed by a 16-bit zero value. An application
' can walk a list of item identifiers by examining the size specified in each
' SHITEMID structure and stopping when it finds a size of zero. A pointer
' to an item identifier list, is sometimes called a PIDL (pronounced piddle)
Type ITEMIDLIST   ' idl
    mkid As SHITEMID
End Type

' Converts an item identifier list to a file system path.
' Returns TRUE if successful or FALSE if an error occurs, for example,
' if the location specified by the pidl parameter is not part of the file system.
Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" _
                              (ByVal pIdl As Long, ByVal pszPath As String) As Long

' Retrieves the location of a special (system) folder.
' Returns NOERROR if successful or an OLE-defined error result otherwise.
Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" _
                              (ByVal hwndOwner As Long, ByVal nFolder As Long, _
                              pIdl As ITEMIDLIST) As Long

' SHGetSpecialFolderLocation successful rtn val
Public Const NOERROR = 0

' SHGetSpecialFolderLocation nFolder params:
' Most folder locations are stored in:
' [HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders]
' Value specifying the types of folders to be listed in the dialog box as well as other options.
' This member can be 0 or one of the following values:

' Windows desktop, virtual folder at the root of the name space.
Public Const CSIDL_DESKTOP = &H0

' File system directory that contains the user's program groups
' (which are also file system directories).
Public Const CSIDL_PROGRAMS = &H2

' Control Panel, virtual folder containing icons for the control panel applications.
Public Const CSIDL_CONTROLS = &H3

' Printers folder, virtual folder containing installed printers.
Public Const CSIDL_PRINTERS = &H4

' File system directory that serves as a common respository for documents.
Public Const CSIDL_PERSONAL = &H5   ' (Documents folder)

' File system directory that contains the user's favorite Internet Explorer URLs.
Public Const CSIDL_FAVORITES = &H6

' File system directory that corresponds to the user's Startup program group.
Public Const CSIDL_STARTUP = &H7

' File system directory that contains the user's most recently used documents.
Public Const CSIDL_RECENT = &H8   ' (Recent folder)

' File system directory that contains Send To menu items.
Public Const CSIDL_SENDTO = &H9

' Recycle bin, file system directory containing file objects in the user's recycle bin.
' The location of this directory is not in the registry; it is marked with the hidden and
' system attributes to prevent the user from moving or deleting it.
Public Const CSIDL_BITBUCKET = &HA

' File system directory containing Start menu items.
Public Const CSIDL_STARTMENU = &HB

' File system directory used to physically store file objects on the desktop
' (not to be confused with the desktop folder itself).
Public Const CSIDL_DESKTOPDIRECTORY = &H10

' My Computer, virtual folder containing everything on the local computer: storage
' devices, printers, and Control Panel. The folder may also contain mapped network drives.
Public Const CSIDL_DRIVES = &H11

' Network Neighborhood, virtual folder representing the top level of the network hierarchy.
Public Const CSIDL_NETWORK = &H12

' File system directory containing objects that appear in the network neighborhood.
Public Const CSIDL_NETHOOD = &H13

' Virtual folder containing fonts.
Public Const CSIDL_FONTS = &H14

' File system directory that serves as a common repository for document templates.
Public Const CSIDL_TEMPLATES = &H15   ' (ShellNew folder)

'========================================================

' Frees memory allocated by SHBrowseForFolder()
Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)

' Displays a dialog box that enables the user to select a shell folder.
' Returns a pointer to an item identifier list that specifies the location
' of the selected folder relative to the root of the name space. If the user
' chooses the Cancel button in the dialog box, the return value is NULL.
Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" _
                              (lpBrowseInfo As BROWSEINFO) As Long ' ITEMIDLIST

' Contains parameters for the the SHBrowseForFolder function and receives
' information about the folder selected by the user.
Public Type BROWSEINFO   ' bi
    
    ' Handle of the owner window for the dialog box.
    hOwner As Long
    
    ' Pointer to an item identifier list (an ITEMIDLIST structure) specifying the location
    ' of the "root" folder to browse from. Only the specified folder and its subfolders
    ' appear in the dialog box. This member can be NULL, and in that case, the
    ' name space root (the desktop folder) is used.
    pidlRoot As Long
    
    ' Pointer to a buffer that receives the display name of the folder selected by the
    ' user. The size of this buffer is assumed to be MAX_PATH bytes.
    pszDisplayName As String
    
    ' Pointer to a null-terminated string that is displayed above the tree view control
    ' in the dialog box. This string can be used to specify instructions to the user.
    lpszTitle As String
    
    ' Value specifying the types of folders to be listed in the dialog box as well as
    ' other options. This member can include zero or more of the following values below.
    ulFlags As Long
    
    ' Address an application-defined function that the dialog box calls when events
    ' occur. For more information, see the description of the BrowseCallbackProc
    ' function. This member can be NULL.
    lpfn As Long
    
    ' Application-defined value that the dialog box passes to the callback function
    ' (if one is specified).
    lParam As Long
    
    ' Variable that receives the image associated with the selected folder. The image
    ' is specified as an index to the system image list.
    iImage As Long

End Type

' BROWSEINFO ulFlags values:
' Value specifying the types of folders to be listed in the dialog box as well as
' other options. This member can include zero or more of the following values:

' Only returns file system directories. If the user selects folders
' that are not part of the file system, the OK button is grayed.
Public Const BIF_RETURNONLYFSDIRS = &H1

' Does not include network folders below the domain level in the tree view control.
' For starting the Find Computer
Public Const BIF_DONTGOBELOWDOMAIN = &H2

' Includes a status area in the dialog box. The callback function can set
' the status text by sending messages to the dialog box.
Public Const BIF_STATUSTEXT = &H4

' Only returns file system ancestors. If the user selects anything other
' than a file system ancestor, the OK button is grayed.
Public Const BIF_RETURNFSANCESTORS = &H8

' Only returns computers. If the user selects anything other
' than a computer, the OK button is grayed.
Public Const BIF_BROWSEFORCOMPUTER = &H1000

' Only returns (network) printers. If the user selects anything other
' than a printer, the OK button is grayed.
Public Const BIF_BROWSEFORPRINTER = &H2000

Public Function openDirDialog(hOwner As Long) As String

    Dim BI As BROWSEINFO
    'Dim nFolder As Long
    'Dim IDL As ITEMIDLIST
    Dim pIdl As Long
    'Dim sPath As String

    With BI
        ' The dialog's owner window...
        .hOwner = hOwner
        ' Set the Browse dialog root folder
        '    nFolder = GetFolderValue(m_wCurOptIdx)
        ' Fill the item id list with the pointer of the selected folder item, rtns 0 on success
        ' ==================================================
        ' If this function fails because the selected folder doesn't exist,
        ' .pidlRoot will be uninitialized & will equal 0 (CSIDL_DESKTOP)
        ' and the root will be the Desktop.
        ' DO NOT specify the CSIDL_ constants for .pidlRoot !!!!
        ' The SHBrowseForFolder() call below will generate a fatal exception
        ' (GPF) if the folder indicated by the CSIDL_ constant does not exist!!
        ' ==================================================
        '    If SHGetSpecialFolderLocation(ByVal Me.hWnd, ByVal nFolder, IDL) = NOERROR Then
        '      .pidlRoot = IDL.mkid.cb
        '    End If
        .pidlRoot = 0
        ' Initialize the buffer that rtns the display name of the selected folder
        .pszDisplayName = String$(MAX_PATH, 0)
        ' Set the dialog's banner text
        .lpszTitle = "Select Directory:"
        ' Set the type of folders to display & return
        ' -play with these option constants to see what can be returned
        .ulFlags = BIF_RETURNONLYFSDIRS
    End With

    ' Show the Browse dialog
    pIdl = SHBrowseForFolder(BI)
    ' If the dialog was cancelled...

    If pIdl = 0 Then Exit Function
    ' Fill sPath w/ the selected path from the id list
    ' (will rtn False if the id list can't be converted)
    openDirDialog = String$(MAX_PATH, 0)
    SHGetPathFromIDList ByVal pIdl, ByVal openDirDialog
    ' Frees the memory SHBrowseForFolder()
    ' allocated for the pointer to the item id list
    CoTaskMemFree pIdl
    openDirDialog = Left$(openDirDialog, InStr(openDirDialog, Chr$(0)) - 1)

End Function

   

