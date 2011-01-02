Attribute VB_Name = "MFolderBrowser"
Option Explicit




''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Copyright ©1996-2008 VBnet, Randy Birch, All Rights Reserved.
' Some pages may also contain other copyrights by the author.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Distribution: You can freely use this code in your own
'               applications, but you may not reproduce
'               or publish this code on any web site,
'               online service, or distribute as source
'               on any media without express permission.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'common to both methods
Public Type BROWSEINFO
  hOwner As Long
  pidlRoot As Long
  pszDisplayName As String
  lpszTitle As String
  ulFlags As Long
  lpfn As Long
  lParam As Long
  iImage As Long
End Type

Public Declare Function SHBrowseForFolder Lib "shell32" _
   Alias "SHBrowseForFolderA" _
   (lpBrowseInfo As BROWSEINFO) As Long

Public Declare Function SHGetPathFromIDList Lib "shell32" _
   Alias "SHGetPathFromIDListA" _
   (ByVal pidl As Long, _
   ByVal pszPath As String) As Long

Public Declare Sub CoTaskMemFree Lib "ole32" (ByVal pv As Long)

Public Declare Function SendMessage Lib "user32" _
   Alias "SendMessageA" _
   (ByVal hWnd As Long, _
   ByVal wMsg As Long, _
   ByVal wParam As Long, _
   lParam As Any) As Long
   
Public Declare Sub CopyMemory Lib "kernel32" _
   Alias "RtlMoveMemory" _
   (pDest As Any, _
    pSource As Any, _
    ByVal dwLength As Long)
    
Public Const MAX_PATH As Long = 260
Public Const WM_USER = &H400
Public Const BFFM_INITIALIZED = 1

'Constants ending in 'A' are for Win95 ANSI
'calls; those ending in 'W' are the wide Unicode
'calls for NT.

'Sets the status text to the null-terminated
'string specified by the lParam parameter.
'wParam is ignored and should be set to 0.
Public Const BFFM_SETSTATUSTEXTA As Long = (WM_USER + 100)
Public Const BFFM_SETSTATUSTEXTW As Long = (WM_USER + 104)

'If the lParam  parameter is non-zero, enables the
'OK button, or disables it if lParam is zero.
'(docs erroneously said wParam!)
'wParam is ignored and should be set to 0.
Public Const BFFM_ENABLEOK As Long = (WM_USER + 101)

'Selects the specified folder. If the wParam
'parameter is FALSE, the lParam parameter is the
'PIDL of the folder to select , or it is the path
'of the folder if wParam is the C value TRUE (or 1).
'Note that after this message is sent, the browse
'dialog receives a subsequent BFFM_SELECTIONCHANGED
'message.
Public Const BFFM_SETSELECTIONA As Long = (WM_USER + 102)
Public Const BFFM_SETSELECTIONW As Long = (WM_USER + 103)
   

'specific to the PIDL method
'Undocumented call for the example. IShellFolder's
'ParseDisplayName member function should be used instead.
Public Declare Function SHSimpleIDListFromPath Lib _
   "shell32" Alias "#162" _
   (ByVal szPath As String) As Long


'specific to the STRING method
Public Declare Function LocalAlloc Lib "kernel32" _
   (ByVal uFlags As Long, _
    ByVal uBytes As Long) As Long
    
Public Declare Function LocalFree Lib "kernel32" _
   (ByVal hMem As Long) As Long

Public Declare Function lstrcpyA Lib "kernel32" _
   (lpString1 As Any, lpString2 As Any) As Long

Public Declare Function lstrlenA Lib "kernel32" _
   (lpString As Any) As Long

Public Const LMEM_FIXED = &H0
Public Const LMEM_ZEROINIT = &H40
Public Const LPTR = (LMEM_FIXED Or LMEM_ZEROINIT)

'windows-defined type OSVERSIONINFO
Public Type OSVERSIONINFO
  OSVSize         As Long
  dwVerMajor      As Long
  dwVerMinor      As Long
  dwBuildNumber   As Long
  PlatformID      As Long
  szCSDVersion    As String * 128
End Type
        
Public Const VER_PLATFORM_WIN32_NT = 2

Public Declare Function GetVersionEx Lib "kernel32" _
   Alias "GetVersionExA" _
  (lpVersionInformation As OSVERSIONINFO) As Long
  
Public Declare Function GetLogicalDriveStrings Lib "kernel32" _
   Alias "GetLogicalDriveStringsA" _
  (ByVal nBufferLength As Long, _
   ByVal lpBuffer As String) As Long



Public Function BrowseCallbackProcStr(ByVal hWnd As Long, _
                                      ByVal uMsg As Long, _
                                      ByVal lParam As Long, _
                                      ByVal lpData As Long) As Long
                                       
  'Callback for the Browse STRING method.
 
  'On initialization, set the dialog's
  'pre-selected folder from the pointer
  'to the path allocated as bi.lParam,
  'passed back to the callback as lpData param.
 
   Select Case uMsg
      Case BFFM_INITIALIZED
      
         Call SendMessage(hWnd, BFFM_SETSELECTIONA, _
                          1&, ByVal lpData)
                          
         Case Else:
         
   End Select
          
End Function
          

Public Function BrowseCallbackProc(ByVal hWnd As Long, _
                                   ByVal uMsg As Long, _
                                   ByVal lParam As Long, _
                                   ByVal lpData As Long) As Long
 
  'Callback for the Browse PIDL method.
 
  'On initialization, set the dialog's
  'pre-selected folder using the pidl
  'set as the bi.lParam, and passed back
  'to the callback as lpData param.
   Select Case uMsg
      Case BFFM_INITIALIZED
      
         Call SendMessage(hWnd, BFFM_SETSELECTIONA, _
                          0&, ByVal lpData)
                          
         Case Else:
         
   End Select

End Function


Public Function FARPROC(pfn As Long) As Long
  
  'A dummy procedure that receives and returns
  'the value of the AddressOf operator.
 
  'This workaround is needed as you can't
  'assign AddressOf directly to a member of a
  'user-defined type, but you can assign it
  'to another long and use that instead!
  FARPROC = pfn

End Function



Public Function BrowseForFolderByPath(sSelPath As String, Optional hWnd As Long = 0) As String

   Dim BI As BROWSEINFO
   Dim pidl As Long
   Dim lpSelPath As Long
   Dim sPath As String * MAX_PATH
   
   With BI
      .hOwner = hWnd
      .pidlRoot = 0
      .lpszTitle = "Pre-selecting folder using the folder's string."
      .lpfn = FARPROC(AddressOf BrowseCallbackProcStr)
    
      lpSelPath = LocalAlloc(LPTR, Len(sSelPath) + 1)
      CopyMemory ByVal lpSelPath, ByVal sSelPath, Len(sSelPath) + 1
      .lParam = lpSelPath
    
   End With
    
   pidl = SHBrowseForFolder(BI)
   
   If pidl Then
     
      If SHGetPathFromIDList(pidl, sPath) Then
         BrowseForFolderByPath = Left$(sPath, InStr(sPath, vbNullChar) - 1)
      Else
         BrowseForFolderByPath = ""
      End If
      
      Call CoTaskMemFree(pidl)
   
   Else
      BrowseForFolderByPath = ""
   End If
   
  Call LocalFree(lpSelPath)

End Function


Public Function BrowseForFolderByPIDL(sSelPath As String, Optional hWnd As Long = 0) As String

   Dim BI As BROWSEINFO
   Dim pidl As Long
   Dim sPath As String * MAX_PATH
     
   With BI
      .hOwner = hWnd
      .pidlRoot = 0
      .lpszTitle = "Pre-selecting a folder using the folder's pidl."
      .lpfn = FARPROC(AddressOf BrowseCallbackProc)
      .lParam = GetPIDLFromPath(sSelPath)
   End With
  
   pidl = SHBrowseForFolder(BI)
  
   If pidl Then
      If SHGetPathFromIDList(pidl, sPath) Then
         BrowseForFolderByPIDL = Left$(sPath, InStr(sPath, vbNullChar) - 1)
      Else
         BrowseForFolderByPIDL = ""
      End If
     
     'free the pidl from SHBrowseForFolder call
      Call CoTaskMemFree(pidl)
   Else
      BrowseForFolderByPIDL = ""
   End If
  
 'free the pidl (lparam) from GetPIDLFromPath call
   Call CoTaskMemFree(BI.lParam)
  
End Function


Public Function GetPIDLFromPath(sPath As String) As Long

  'return the pidl to the path supplied by calling the
  'undocumented API #162 (our name for this undocumented
  'function is "SHSimpleIDListFromPath").
  'This function is necessary as, unlike documented APIs,
  'the API is not implemented in 'A' or 'W' versions.

   If IsWinNT() Then
      GetPIDLFromPath = SHSimpleIDListFromPath(StrConv(sPath, vbUnicode))
   Else
      GetPIDLFromPath = SHSimpleIDListFromPath(sPath)
   End If

End Function


Public Function IsWinNT() As Boolean

   #If Win32 Then
  
      Dim OSV As OSVERSIONINFO
   
      OSV.OSVSize = Len(OSV)
   
     'API returns 1 if a successful call
      If GetVersionEx(OSV) = 1 Then
   
        'PlatformId contains a value representing
        'the OS; if VER_PLATFORM_WIN32_NT,
        'return true
         IsWinNT = OSV.PlatformID = VER_PLATFORM_WIN32_NT
      End If

   #End If

End Function


Public Function IsValidDrive(sPath As String) As Boolean

   Dim buff As String
   Dim nBuffsize As Long
   
  'Call the API with a buffer size of 0.
  'The call fails, and the required size
  'is returned as the result.
   nBuffsize = GetLogicalDriveStrings(0&, buff)

  'pad a buffer to hold the results
   buff = Space$(nBuffsize)
   nBuffsize = Len(buff)
   
  'and call again
   If GetLogicalDriveStrings(nBuffsize, buff) Then
   
     'if the drive letter passed is in
     'the returned logical drive string,
     'return True.
      IsValidDrive = InStr(1, buff, sPath, vbTextCompare) > 0
   
   End If

End Function


Public Function FixPath(sPath As String) As String

  'The Browse callback requires the path string
  'in a specific format - trailing slash if a
  'drive only, or minus a trailing slash if a
  'file system path. This routine assures the
  'string is formatted correctly.
  '
  'In addition, because the calls to LocalAlloc
  'requires a valid path for the call to succeed,
  'the path defaults to C:\ if the passed string
  'is empty.
  
  'Test 1: check for empty string. Since
  'we're setting it we can assure it is
  'formatted correctly, so can bail.
   If Len(sPath) = 0 Then
      FixPath = "C:\"
      Exit Function
   End If
   
  'Test 2: is path a valid drive?
  'If this far we did not set the path,
  'so need further tests. Here we ensure
  'the path is properly terminated with
  'a trailing slash as needed.
  '
  'Drives alone require the trailing slash;
  'file system paths must have it removed.
   If IsValidDrive(sPath) Then
      
     'IsValidDrive only determines if the
     'path provided is contained in
     'GetLogicalDriveStrings. Since
     'IsValidDrive() will return True
     'if either C: or C:\ is passed, we
     'need to ensure the string is formatted
     'with the trailing slash.
      FixPath = QualifyPath(sPath)
   Else
     'The string passed was not a drive, so
     'assume it's a path and ensure it does
     'not have a trailing space.
      FixPath = UnqualifyPath(sPath)
   End If
   
End Function


Public Function QualifyPath(sPath As String) As String
 
   If Len(sPath) > 0 Then
 
      If Right$(sPath, 1) <> "\" Then
         QualifyPath = sPath & "\"
      Else
         QualifyPath = sPath
      End If
      
   Else
      QualifyPath = ""
   End If
   
End Function


Public Function UnqualifyPath(sPath As String) As String

  'Qualifying a path involves assuring that its format
  'is valid, including a trailing slash, ready for a
  'filename. Since SHBrowseForFolder will not pre-select
  'the path if it contains the trailing slash, it must be
  'removed, hence 'unqualifying' the path.
   If Len(sPath) > 0 Then
   
      If Right$(sPath, 1) = "\" Then
      
         UnqualifyPath = Left$(sPath, Len(sPath) - 1)
         Exit Function
      
      End If
   
   End If
   
   UnqualifyPath = sPath
   
End Function
