Attribute VB_Name = "mBrowseForFolder"
Option Explicit


Private alloc As IMalloc

Public Const MAX_PATH = 260
Private Const WM_USER = &H400
Private Const BFFM_INITIALIZED As Long = 1
Private Const BFFM_SELCHANGED  As Long = 2
Public Const BFFM_SETSTATUSTEXTA As Long = (WM_USER + 100)
Public Const BFFM_SETSELECTIONA As Long = (WM_USER + 102)
Private Const BFFM_VALIDATEFAILEDA = 3      '// lParam:szPath ret:1(cont),0(EndDialog)
'// message from browser
'#define BFFM_INITIALIZED        1
'#define BFFM_VALIDATEFAILEDA    3   // lParam:szPath ret:1(cont),0(EndDialog)
'#define BFFM_VALIDATEFAILEDW    4   // lParam:wzPath ret:1(cont),0(EndDialog)
'
'// messages to browser
Private Const BFFM_SETSTATUSTEXT = (WM_USER + 100)
Private Const BFFM_ENABLEOK = (WM_USER + 101)
'#define BFFM_SETSELECTIONA      (WM_USER + 102)
'#define BFFM_SETSELECTIONW      (WM_USER + 103)
'#define BFFM_SETSTATUSTEXTW     (WM_USER + 104)

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SendMessageStr Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Private Declare Function SHGetMalloc Lib "shell32.dll" (ppMalloc As IMalloc) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Private Declare Sub CopyMemoryLpToStr Lib "kernel32" Alias "RtlMoveMemory" ( _
    ByVal lpvDest As String, lpvSource As Long, ByVal cbCopy As Long)
Private Declare Function lstrlenptr Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Long) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long

Property Get Allocator() As IMalloc
    If alloc Is Nothing Then SHGetMalloc alloc
    Set Allocator = alloc
End Property

Public Property Get ObjectFromPtr(ByVal lPtr As Long) As cBrowseForFolder
Dim oThis As cBrowseForFolder
    CopyMemory oThis, lPtr, 4
    Set ObjectFromPtr = oThis
    CopyMemory oThis, 0&, 4
End Property

' This function for standard module only--global module version
' must be in separate file
Public Function BrowseCallbackProc(ByVal hwnd As Long, _
                            ByVal uMsg As Long, _
                            ByVal lParam As Long, _
                            ByVal lpData As Long) As Long

Dim sPath As String
Dim lR As Long
Dim pidl As Long
Dim cBF As cBrowseForFolder

   Select Case uMsg
   ' Browse dialog box has finished initializing (lParam is NULL)
   Case BFFM_INITIALIZED
      Debug.Print "BFFM_INITIALIZED"
      ' Set the selection
      If lpData <> 0 Then
         Set cBF = ObjectFromPtr(lpData)
         If Not cBF Is Nothing Then
            pidl = cBF.pidlInitial
            If pidl > 0 Then
               lR = SendMessage(hwnd, BFFM_SETSELECTIONA, 0, ByVal pidl)
            End If
            cBF.Initialized hwnd
         End If
         
      End If
      BrowseCallbackProc = 0
      
   ' Selection has changed (lParam contains pidl of selected folder)
   Case BFFM_SELCHANGED
      Debug.Print "BFFM_SELCHANGED"
      ' Display full path if status area if enabled
      sPath = PathFromPidl(lParam)
      lR = SendMessageStr(hwnd, BFFM_SETSTATUSTEXTA, 0&, sPath)
      If lpData <> 0 Then
         ObjectFromPtr(lpData).SelectionChange hwnd, sPath, lParam
      End If
      BrowseCallbackProc = 0
   ' Invalid name in edit box (lParam parameter has invalid name string)
   Case BFFM_VALIDATEFAILEDA
      Debug.Print "BFFM_VALIDATEFAILED"
      ' Return zero to dismiss dialog or nonzero to keep it displayed
      ' Disable the OK button
      lR = SendMessage(hwnd, BFFM_ENABLEOK, ByVal 0&, ByVal 0&)
      sPath = PointerToString(lParam)
      sPath = "Path invalid: " & sPath
      lR = SendMessageStr(hwnd, BFFM_SETSTATUSTEXT, ByVal 0&, sPath)
      If lpData <> 0 Then
         BrowseCallbackProc = ObjectFromPtr(lpData).ValidateFailed(hwnd, sPath)
      Else
         BrowseCallbackProc = 0
      End If
   End Select

End Function
Public Function PointerToString(lPtr As Long) As String
Dim lLen As Long
Dim sR As String
    ' Get length of Unicode string to first null
    lLen = lstrlenptr(lPtr)
    ' Allocate a string of that length
    sR = String$(lLen, 0)
    ' Copy the pointer data to the string
    CopyMemoryLpToStr sR, ByVal lPtr, lLen
    PointerToString = sR
End Function

Public Function PathFromPidl(ByVal pidl As Long) As String
Dim sPath As String
Dim lR As Long
   sPath = String$(MAX_PATH, 0)
   lR = SHGetPathFromIDList(pidl, sPath)
   If lR <> 0 Then
      PathFromPidl = left$(sPath, lstrlen(sPath))
   End If
End Function



