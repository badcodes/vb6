Attribute VB_Name = "MRegistry"
'=========================================================================================
'  modRegistry
'  registry functions and routines
'=========================================================================================
'  Adapted and Modified By: Marc Cramer
'  Published Date: 04/18/2001
'  Copyright Datr: Marc Cramer ?04/18/2001
'  WebSite: www.mkccomputers.com
'=========================================================================================
'  Based On: API description and examples from Windows API Guide
'  WebSite: Windows API Guide  @ www.vbapi.com
'  Based On: API description and examples from The AllAPI Network
'  WebSite: The AllAPI Network @ www.allapi.net
'=========================================================================================
Option Explicit
'=========================================================================================
' Enums/Constants used for Registry Access
'=========================================================================================
Public Enum KeyRoot
  [HKEY_CLASSES_ROOT] = &H80000000  'stores OLE class information and file associations
  [HKEY_CURRENT_CONFIG] = &H80000005 'stores computer configuration information
  [HKEY_CURRENT_USER] = &H80000001 'stores program information for the current user.
  [HKEY_LOCAL_MACHINE] = &H80000002 'stores program information for all users
  [HKEY_USERS] = &H80000003 'has all the information for any user (not just the one provided by HKEY_CURRENT_USER)
End Enum
Public Enum KeyType
  [REG_BINARY] = 3 'A non-text sequence of bytes
  [REG_DWORD] = 4 'A 32-bit integer...visual basic data type of Long
  [REG_SZ] = 1 'A string terminated by a null character
End Enum

Private Const KEY_ALL_ACCESS = &HF003F 'Permission for all types of access.
Private Const KEY_ENUMERATE_SUB_KEYS = &H8 'Permission to enumerate subkeys.
Private Const KEY_READ = &H20019 'Permission for general read access.
Private Const KEY_WRITE = &H20006 'Permission for general write access.
Private Const KEY_QUERY_VALUE = &H1 'Permission to query subkey data.
' used for import/export registry key
Private Const REG_FORCE_RESTORE As Long = 8& 'Permission to overwrite a registry key
Private Const TOKEN_QUERY As Long = &H8&
Private Const TOKEN_ADJUST_PRIVILEGES As Long = &H20&
Private Const SE_PRIVILEGE_ENABLED As Long = &H2
Private Const SE_RESTORE_NAME = "SeRestorePrivilege" 'Important for what we're trying to accomplish
Private Const SE_BACKUP_NAME = "SeBackupPrivilege"
'=========================================================================================
' Type used for Registry
'=========================================================================================
' used for writing registry keys
Private Type SECURITY_ATTRIBUTES
  nLength As Long
  lpSecurityDescriptor As Long
  bInheritHandle As Boolean
End Type
' used for enumerating registrykeys
Private Type FILETIME
  dwLowDateTime As Long
  dwHighDateTime As Long
End Type
' used for import/export registry key
Private Type LUID
  lowpart As Long
  highpart As Long
End Type
Private Type LUID_AND_ATTRIBUTES
  pLuid As LUID
  Attributes As Long
End Type
Private Type TOKEN_PRIVILEGES
  PrivilegeCount As Long
  Privileges As LUID_AND_ATTRIBUTES
End Type
'=========================================================================================
' API Function Declarations used for Registry
'=========================================================================================
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegSaveKey Lib "advapi32.dll" Alias "RegSaveKeyA" (ByVal hKey As Long, ByVal lpFile As String, lpSecurityAttributes As Any) As Long
Private Declare Function RegRestoreKey Lib "advapi32.dll" Alias "RegRestoreKeyA" (ByVal hKey As Long, ByVal lpFile As String, ByVal dwFlags As Long) As Long
Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As FILETIME) As Long
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long
' used for export/import registry keys
Private Declare Function AdjustTokenPrivileges Lib "advapi32.dll" (ByVal TokenHandle As Long, ByVal DisableAllPriv As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long                'Used to adjust your program's security privileges, can't restore without it!
Private Declare Function LookupPrivilegeValue Lib "advapi32.dll" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As Any, ByVal lpName As String, lpLuid As LUID) As Long          'Returns a valid LUID which is important when making security changes in NT.
Private Declare Function OpenProcessToken Lib "advapi32.dll" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
'=========================================================================================
' Routines used to Access Registry
'=========================================================================================
Public Function ExportRegKey(KeyRoot As KeyRoot, KeyPath As String, FileName As String) As Boolean
  ' routine to export registry keys
  On Error Resume Next
  Dim hKey As Long
  Dim ReturnValue As Long

  ' check to see if allowed to do this
  If EnablePrivilege(SE_BACKUP_NAME) = False Then
    ExportRegKey = False
    Exit Function
  End If
  ' open the registry key
  ReturnValue = RegOpenKeyEx(KeyRoot, KeyPath, 0&, KEY_ALL_ACCESS, hKey)
  If ReturnValue <> 0 Then
    ' error encountered
    ExportRegKey = False
    ReturnValue = RegCloseKey(hKey)
    Exit Function
  End If
  ' check for a copy of the export and delete old one if applicable
  If Dir(FileName) <> "" Then Kill FileName
  ' export the registry key
  ReturnValue = RegSaveKey(hKey, FileName, ByVal 0&)
  If ReturnValue = 0 Then
    ' no error encountered
    ExportRegKey = True
  Else
    ' error encountered
    ExportRegKey = False
  End If
  ' close the registry key
  ReturnValue = RegCloseKey(hKey)
End Function 'ExportRegKey(KeyRoot As KeyRoot, KeyPath As String, FileName As String) As Boolean
'=========================================================================================
Public Function ImportRegKey(KeyRoot As KeyRoot, KeyPath As String, FileName As String) As Boolean
  ' routine to import registry keys
  ' will overwrite current settings, but will not create keys
  On Error Resume Next
  Dim hKey As Long
  Dim ReturnValue As Long

  ' check to see if allowed to do this
  If EnablePrivilege(SE_RESTORE_NAME) = False Then
    ImportRegKey = False
    Exit Function
  End If
  ' open the registry key
  ReturnValue = RegOpenKeyEx(KeyRoot, KeyPath, 0&, KEY_ALL_ACCESS, hKey)
  If ReturnValue <> 0 Then
    ' error encountered
    ImportRegKey = False
    ReturnValue = RegCloseKey(hKey)
    Exit Function
  End If
  ' import the registry key
  ReturnValue = RegRestoreKey(hKey, FileName, REG_FORCE_RESTORE)
  If ReturnValue = 0 Then
    ' no error encountered
    ImportRegKey = True
  Else
    ' error encountered
    ImportRegKey = False
  End If
  ' close the registry key
  ReturnValue = RegCloseKey(hKey)
End Function 'ImportRegKey(KeyRoot As KeyRoot, KeyPath As String, FileName As String) As Boolean
'=========================================================================================
Public Function ReadRegKey(KeyRoot As KeyRoot, KeyPath As String, SubKey As String, Optional NoKeyFoundValue As String = "") As String
  ' routine to read entry from registry
  On Error Resume Next
  Dim hKey As Long  ' receives a handle to the opened registry key
  Dim ReturnValue As Long  ' return value

  ' open the registry key
  ReturnValue = RegOpenKeyEx(KeyRoot, KeyPath, 0, KEY_READ, hKey)
  If ReturnValue <> 0 Then
    ' key doesn't exist so return default value
    ReadRegKey = NoKeyFoundValue
    ReturnValue = RegCloseKey(hKey)
    Exit Function
  End If
  ' get the keys value
  ReadRegKey = GetSubKeyValue(hKey, SubKey)
  ' close the registry key
  ReturnValue = RegCloseKey(hKey)
End Function 'ReadRegKey(KeyRoot As KeyRoot, KeyPath As String, SubKey As String, Optional NoKeyFoundValue As String = "") As String
'=========================================================================================
Public Function WriteRegKey(KeyType As KeyType, KeyRoot As KeyRoot, KeyPath As String, SubKey As String, SubKeyValue As String) As Boolean
  ' routine to write entry to registry
  On Error Resume Next
  Dim hKey As Long  ' receives handle to the newly created or opened registry key
  Dim SecurityAttribute As SECURITY_ATTRIBUTES  ' security settings of the key
  Dim NewKey As Long  ' receives 1 if new key was created or 2 if an existing key was opened
  Dim ReturnValue As Long  ' return value

  ' Set the name of the new key and the default security settings
  SecurityAttribute.nLength = Len(SecurityAttribute)  ' size of the structure
  SecurityAttribute.lpSecurityDescriptor = 0  ' default security level
  SecurityAttribute.bInheritHandle = True  ' the default value for this setting

  ' create or open the registry key
  ReturnValue = RegCreateKeyEx(KeyRoot, KeyPath, 0, "", 0, KEY_WRITE, SecurityAttribute, hKey, NewKey)
  If ReturnValue <> 0 Then
    ' error encountered
    WriteRegKey = False
    ReturnValue = RegCloseKey(hKey)
    Exit Function
  End If

  ' determine type of key and write it to the registry
  Select Case KeyType
    Case REG_SZ
      ReturnValue = RegSetValueEx(hKey, SubKey, 0, KeyType, ByVal SubKeyValue, Len(SubKeyValue))
    Case REG_DWORD
      ReturnValue = RegSetValueEx(hKey, SubKey, 0, KeyType, CLng(SubKeyValue), 4)
    Case REG_BINARY
      ReturnValue = RegSetValueEx(hKey, SubKey, 0, KeyType, CByte(SubKeyValue), 4)
  End Select

  If ReturnValue = 0 Then
    ' no error encountered
    WriteRegKey = True
  Else
    ' error encountered
    WriteRegKey = False
  End If

  ' close the registry key
  ReturnValue = RegCloseKey(hKey)
End Function 'WriteRegKey(KeyType As KeyType, KeyRoot As KeyRoot, KeyPath As String, SubKey As String, SubKeyValue As String) As Boolean
'=========================================================================================
Public Function EnumerateRegKeys(KeyRoot As KeyRoot, KeyPath As String) As String
  ' routine to enumerate all subkeys under a registry key
  On Error Resume Next
  Dim hKey As Long  ' receives a handle to the opened registry key
  Dim ReturnValue As Long  ' return value
  Dim Counter As Long
  Dim MyBuffer As String
  Dim MyBufferSize As Long
  Dim ClassNameBuffer As String
  Dim ClassNameBufferSize As Long
  Dim LastWrite As FILETIME

  ' open the registry key
  ReturnValue = RegOpenKeyEx(KeyRoot, KeyPath, 0, KEY_ENUMERATE_SUB_KEYS, hKey)
  If ReturnValue <> 0 Then
    ' key doesn't exist so return default value
    EnumerateRegKeys = ""
    ReturnValue = RegCloseKey(hKey)
    Exit Function
  End If
  Counter = 0
  ' loop until no more registry keys
  Do Until ReturnValue <> 0
    MyBuffer = Space(255)
    ClassNameBuffer = Space(255)
    MyBufferSize = 255
    ClassNameBufferSize = 255
    ReturnValue = RegEnumKeyEx(hKey, Counter, MyBuffer, MyBufferSize, ByVal 0, ClassNameBuffer, ClassNameBufferSize, LastWrite)
    If ReturnValue = 0 Then
      MyBuffer = Left$(MyBuffer, MyBufferSize)
      ClassNameBuffer = Left$(ClassNameBuffer, ClassNameBufferSize)
      EnumerateRegKeys = EnumerateRegKeys & MyBuffer & ","
    End If
    Counter = Counter + 1
  Loop
  ' trim off the last delimiter
  If EnumerateRegKeys <> "" Then EnumerateRegKeys = Left$(EnumerateRegKeys, Len(EnumerateRegKeys) - 1)
  ' close the registry key
  ReturnValue = RegCloseKey(hKey)
End Function 'EnumerateRegKeys(KeyRoot As KeyRoot, KeyPath As String) As String
'=========================================================================================
Public Function EnumerateRegKeyValues(KeyRoot As KeyRoot, KeyPath As String) As String
  ' routine to enumerate all the values under a key in the registry
  On Error Resume Next
  Dim hKey As Long  ' receives a handle to the opened registry key
  Dim ReturnValue As Long  ' return value
  Dim Counter As Long
  Dim MyBuffer As String
  Dim MyBufferSize As Long
  Dim KeyType As KeyType

  ' open the registry key to enumerate the values of.
  ReturnValue = RegOpenKeyEx(KeyRoot, KeyPath, 0, KEY_QUERY_VALUE, hKey)
  ' check to see if an error occured.
  If ReturnValue <> 0 Then
    EnumerateRegKeyValues = ""
    ReturnValue = RegCloseKey(hKey)
    Exit Function
  End If
  Counter = 0
  ' loop until no more registry keys value
  Do Until ReturnValue <> 0
    MyBuffer = Space(255)
    MyBufferSize = 255
    ReturnValue = RegEnumValue(hKey, Counter, MyBuffer, MyBufferSize, 0, KeyType, ByVal 0&, ByVal 0&) 'ByteData(0), ByteDataSize)
    If ReturnValue = 0 Then
      MyBuffer = Left$(MyBuffer, MyBufferSize)
      EnumerateRegKeyValues = EnumerateRegKeyValues & MyBuffer & "*"
      EnumerateRegKeyValues = EnumerateRegKeyValues & GetSubKeyValue(hKey, MyBuffer) & ","
    End If
    Counter = Counter + 1
  Loop
  ' trim off the last delimiter
  If EnumerateRegKeyValues <> "" Then EnumerateRegKeyValues = Left$(EnumerateRegKeyValues, Len(EnumerateRegKeyValues) - 1)
  ' close the registry key
  ReturnValue = RegCloseKey(hKey)
End Function 'EnumerateRegKeyValues(KeyRoot As KeyRoot, KeyPath As String) As String
'=========================================================================================
Public Function DeleteRegKey(KeyRoot As KeyRoot, KeyPath As String, SubKey As String) As Boolean
  ' routine to delete a registry key
  ' under Win NT/2000 all subkeys must be deleted first
  ' under Win 9x all subkeys are deleted
  On Error Resume Next
  Dim ReturnValue As Long  ' return value

  ' Attempt to delete the desired registry key.
  ReturnValue = RegDeleteKey(KeyRoot, KeyPath & "\" & SubKey)
  If ReturnValue = 0 Then
    ' no error encountered
    DeleteRegKey = True
  Else
    ' error encountered
    DeleteRegKey = False
  End If
End Function 'DeleteRegKey(KeyRoot As KeyRoot, KeyPath As String, SubKey As String) As Boolean
'=========================================================================================
Public Function DeleteRegKeyValue(KeyRoot As KeyRoot, KeyPath As String, Optional SubKey As String = "") As Boolean
  ' routine to delete a value from a key (but not the key) in the registry
  On Error Resume Next
  Dim hKey As Long  ' handle to the open registry key
  Dim ReturnValue As Long  ' return value

  ' First, open up the registry key which holds the value to delete.
  ReturnValue = RegOpenKeyEx(KeyRoot, KeyPath, 0, KEY_ALL_ACCESS, hKey)
  If ReturnValue <> 0 Then
    ' error encountered
    DeleteRegKeyValue = False
    ReturnValue = RegCloseKey(hKey)
    Exit Function
  End If
  ' check to see if we are deleting a subkey or primary key
  If SubKey = "" Then SubKey = KeyPath
  ' successfully opened registry key so delete the desired value from the key.
  ReturnValue = RegDeleteValue(hKey, SubKey)
  If ReturnValue = 0 Then
    ' no error encountered
    DeleteRegKeyValue = True
  Else
    ' error encountered
    DeleteRegKeyValue = False
  End If
  ' close the registry key
  ReturnValue = RegCloseKey(hKey)
End Function 'DeleteRegKeyValue(KeyRoot As KeyRoot, KeyPath As String, Optional SubKey As String = "") As Boolean
'=========================================================================================
Private Function GetSubKeyValue(ByVal hKey As Long, ByVal SubKey As String) As String
  ' routine to get the registry key value and convert to a string
  On Error Resume Next
  Dim ReturnValue As Long
  Dim KeyType As KeyType
  Dim MyBuffer As String
  Dim MyBufferSize As Long

  'get registry key information
  ReturnValue = RegQueryValueEx(hKey, SubKey, 0, KeyType, ByVal 0, MyBufferSize)
  If ReturnValue = 0 Then ' no error encountered
    ' determine what the KeyType is
    Select Case KeyType
      Case REG_SZ
        ' create a buffer
        MyBuffer = String(MyBufferSize, Chr$(0))
        ' retrieve the key's content
        ReturnValue = RegQueryValueEx(hKey, SubKey, 0, 0, ByVal MyBuffer, MyBufferSize)
        If ReturnValue = 0 Then
          ' remove the unnecessary chr$(0)'s
          GetSubKeyValue = Left$(MyBuffer, InStr(1, MyBuffer, Chr$(0)) - 1)
        End If
      Case Else 'REG_DWORD or REG_BINARY
        Dim MyNewBuffer As Long
        ' retrieve the key's value
        ReturnValue = RegQueryValueEx(hKey, SubKey, 0, 0, MyNewBuffer, MyBufferSize)
        If ReturnValue = 0 Then ' no error encountered
          GetSubKeyValue = MyNewBuffer
        End If
    End Select
  End If
End Function 'GetSubKeyValue(ByVal hKey As Long, ByVal SubKey As String) As String
'=========================================================================================
Private Function EnablePrivilege(seName As String) As Boolean
  ' routine to enable inport/export of registry settings
  On Error Resume Next
  Dim p_lngRtn As Long
  Dim p_lngToken As Long
  Dim p_lngBufferLen As Long
  Dim p_typLUID As LUID
  Dim p_typTokenPriv As TOKEN_PRIVILEGES
  Dim p_typPrevTokenPriv As TOKEN_PRIVILEGES

  ' open the current process token
  p_lngRtn = OpenProcessToken(GetCurrentProcess(), TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY, p_lngToken)
  If p_lngRtn = 0 Then
    ' error encountered
    EnablePrivilege = False
    Exit Function
  End If
  If Err.LastDllError <> 0 Then
    ' error encountered
    EnablePrivilege = False
    Exit Function
  End If
  ' look up the privileges LUID
  p_lngRtn = LookupPrivilegeValue(0&, seName, p_typLUID)
  If p_lngRtn = 0 Then
    ' error encountered
    EnablePrivilege = False
    Exit Function
  End If
  ' adjust the program's security privilege.
  p_typTokenPriv.PrivilegeCount = 1
  p_typTokenPriv.Privileges.Attributes = SE_PRIVILEGE_ENABLED
  p_typTokenPriv.Privileges.pLuid = p_typLUID
  ' try to adjust privileges and return success or failure
  EnablePrivilege = (AdjustTokenPrivileges(p_lngToken, False, p_typTokenPriv, Len(p_typPrevTokenPriv), p_typPrevTokenPriv, p_lngBufferLen) <> 0)
End Function 'EnablePrivilege(seName As String) As Boolean
'=========================================================================================

