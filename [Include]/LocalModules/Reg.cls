VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public m_key As String
Public m_root As Long

Public Enum RegRoot
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
    HKEY_PERFORMANCE_DATA = &H80000004
    HKEY_CURRENT_CONFIG = &H80000005
    HKEY_DYN_DATA = &H80000006
End Enum

Public Enum RegData
    REG_NONE = 0&
    REG_EXPAND_SZ = 2&
    REG_SZ = 1&
    REG_BINARY = 3&
    REG_DWORD = 4&
    REG_DWORD_BIG_ENDIAN = 5&
    REG_LINK = 6&
    REG_MULTI_SZ = 7&
    REG_RESOURCE_LIST = 8&
    REG_FULL_RESOURCE_DESCRIPTOR = 9&
End Enum

Public Enum ErrorList
    ERROR_SUCCESS = 0&
    ERROR_FILE_NOT_FOUND = 2&
    ERROR_ACCESS_DENIED = 5&
    ERROR_INVALID_HANDLE = 6&
    ERROR_OUTOFMEMORY = 14&
    ERROR_INVALID_PARAMETER = 87&
    ERROR_BAD_PATHNAME = 161&
    ERROR_MORE_DATA = 234&
    ERROR_NO_MORE_ITEMS = 259&
    ERROR_BADKEY = 1010&
    ERROR_CANTOPEN = 1011&
    ERROR_CANTREAD = 1012&
    ERROR_CANTWRITE = 1013&
    ERROR_KEYDELETED = 1018&
    ERROR_EMPTY_ARRAY = vbObjectError + 512 + 5
End Enum

Private Const SYNCHRONIZE = &H100000
Private Const STANDARD_RIGHTS_ALL As Long = &H1F0000
Private Const READ_CONTROL As Long = &H20000
Private Const STANDARD_RIGHTS_READ As Long = (READ_CONTROL)
Private Const STANDARD_RIGHTS_WRITE As Long = (READ_CONTROL)

Private Const KEY_QUERY_VALUE As Long = &H1
Private Const KEY_SET_VALUE As Long = &H2
Private Const KEY_CREATE_SUB_KEY As Long = &H4
Private Const KEY_ENUMERATE_SUB_KEYS As Long = &H8
Private Const KEY_NOTIFY As Long = &H10
Private Const KEY_CREATE_LINK As Long = &H20
Private Const KEY_READ As Long = _
   ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or _
   KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) _
   And (Not SYNCHRONIZE))
Private Const KEY_EXECUTE As Long = _
   ((KEY_READ) And (Not SYNCHRONIZE))
Private Const KEY_WRITE As Long = _
   ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE _
   Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))
Private Const KEY_ALL_ACCESS As Long = _
   ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or _
   KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or _
   KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or _
   KEY_CREATE_LINK) And (Not SYNCHRONIZE))


Private Const REG_OPTION_NON_VOLATILE As Long = 0

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Declare Function RegCloseKey Lib "advapi32.dll" _
    (ByVal hKey As Long) As Long

Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias _
    "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, _
    ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As _
    Long) As Long

Private Declare Function RegQueryInfoKey Lib "advapi32.dll" Alias _
    "RegQueryInfoKeyA" (ByVal hKey As Long, ByVal lpClass As String, _
    lpcbClass As Long, ByVal lpReserved As Long, lpcSubKeys As Long, _
    lpcbMaxSubKeyLen As Long, lpcbMaxClassLen As Long, lpcValues As Long, _
    lpcbMaxValueNameLen As Long, lpcbMaxValueLen As Long, _
    ByVal lpcbSecurityDescriptor As Long, lpftLastWriteTime As FILETIME) As Long

Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias _
    "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, _
    ByVal lpName As String, lpcbName As Long, ByVal lpReserved As Long, _
    ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As FILETIME) As Long
    
Private Declare Function RegEnumValueNull Lib "advapi32.dll" Alias _
    "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, _
    ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, _
    lpType As Long, lpData As Byte, lpcbData As Long) As Long
    
Private Declare Function RegEnumValueString Lib "advapi32.dll" Alias _
    "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, _
    ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, _
    lpType As Long, ByVal lpValue As String, lpcbData As Long) As Long
    
Private Declare Function RegEnumValueLong Lib "advapi32.dll" Alias _
    "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, _
    ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, _
    lpType As Long, lpValue As Long, lpcbData As Long) As Long
    
Private Declare Function RegEnumValueBinary Lib "advapi32.dll" Alias _
    "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, _
    ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, _
    lpType As Long, lpValue As Byte, lpcbData As Long) As Long
    
Private Sub Class_Initialize()
    m_root = HKEY_CURRENT_USER
    m_key = "Software\VB and VBA Program Settings"
End Sub

Public Function GetAllSettings(vArray() As Variant) As ErrorList
On Error GoTo Handler
Dim lRetVal As Long, dwIndex As Long

Dim hKey As Long, lpcSubKeys As Long

Dim lpcbMaxSubKeyLen As Long, lpcbMaxClassLen As Long
Dim lpcValues As Long, lpcbMaxValueNameLen As Long
Dim lpcbMaxValueLen As Long

Dim lpcbValueName As Long, lpcbData As Long
Dim lValue As Long, lpType As Long

Dim lpValueName As String, sClass As String, lpData As String

Dim bValue() As Byte

Dim FTime As FILETIME

If OpenKey(hKey) <> ERROR_SUCCESS Then Exit Function
lRetVal = RegQueryInfoKey(hKey, 0&, 0&, 0&, _
   lpcSubKeys, lpcbMaxSubKeyLen, _
   lpcbMaxClassLen, lpcValues, lpcbMaxValueNameLen, _
   lpcbMaxValueLen, 0&, FTime)
   
If lpcValues = 0 Then _
   Err.Raise ERROR_EMPTY_ARRAY

ReDim vArray(lpcValues - 1, 1)

    dwIndex = 0
    Do
        lpcbValueName = lpcbMaxValueNameLen + 1
        lpcbData = lpcbMaxValueLen + 1
        lpType = 0
        lpValueName = space(lpcbValueName)
        If RegEnumValueNull(hKey, dwIndex, 0&, 0&, 0&, lpType, 0&, 0&) _
           = ERROR_NO_MORE_ITEMS Then Exit Do
        Select Case lpType
        Case REG_SZ:
            lpData = space(lpcbData)
            lpcbValueName = lpcbMaxValueNameLen + 1
            lRetVal = RegEnumValueString(hKey, dwIndex, _
               lpValueName, lpcbValueName, 0&, lpType, lpData, lpcbData)
            If Left(lpValueName, lpcbValueName) = "" Then
                vArray(dwIndex, 0) = "Default"
            Else
                vArray(dwIndex, 0) = Left(lpValueName, lpcbValueName)
            End If
            If lRetVal = ERROR_SUCCESS Then
                vArray(dwIndex, 1) = Left$(lpData, lpcbData - 1)
            Else
                vArray(dwIndex, 1) = Empty
            End If
        Case REG_EXPAND_SZ:
            lpData = space(lpcbData)
            lpcbValueName = lpcbMaxValueNameLen + 1
            lRetVal = RegEnumValueString(hKey, dwIndex, _
               lpValueName, lpcbValueName, 0&, lpType, lpData, lpcbData)
            If Left(lpValueName, lpcbValueName) = "" Then
                vArray(dwIndex, 0) = "Default"
            Else
                vArray(dwIndex, 0) = Left(lpValueName, lpcbValueName)
            End If
            If lRetVal = ERROR_SUCCESS Then
                vArray(dwIndex, 1) = Left$(lpData, lpcbData - 1)
            Else
                vArray(dwIndex, 1) = Empty
            End If
        Case REG_BINARY:
            ReDim bValue(lpcbData)
            lpcbValueName = lpcbMaxValueNameLen
            lpcbData = lpcbMaxValueLen
            lRetVal = RegEnumValueBinary(hKey, dwIndex, lpValueName, lpcbValueName, 0&, lpType, bValue(0), lpcbData)
            If Left(lpValueName, lpcbValueName) = "" Then
                vArray(dwIndex, 0) = "Default"
            Else
                vArray(dwIndex, 0) = Left(lpValueName, lpcbValueName)
            End If
            If lRetVal = ERROR_SUCCESS Then vArray(dwIndex, 1) = bValue
            
        Case REG_DWORD:
            lpcbValueName = lpcbMaxValueNameLen
            lpcbData = lpcbMaxValueLen
            lRetVal = RegEnumValueLong(hKey, dwIndex, lpValueName, lpcbValueName, 0&, lpType, lValue, lpcbData)
            If Left(lpValueName, lpcbValueName) = "" Then
                vArray(dwIndex, 0) = "Default"
            Else
                vArray(dwIndex, 0) = Left(lpValueName, lpcbValueName)
            End If
            If lRetVal = ERROR_SUCCESS Then vArray(dwIndex, 1) = lValue
            
        Case Else
            vArray(dwIndex, 0) = "Unsupported Data Type"
            vArray(dwIndex, 1) = "Unsupported Data Type"
        End Select
        dwIndex = dwIndex + 1
    Loop

RegCloseKey (hKey)
GetAllSettings = lRetVal
Exit Function
Handler:
If Err.Number = ERROR_EMPTY_ARRAY Then
Err.clear
GetAllSettings = ERROR_EMPTY_ARRAY
End If
End Function

Public Function GetAllKeys(ByRef sArray() As String) As ErrorList
On Error GoTo Handler
Dim lRetVal As Long

Dim hKey As Long, lpcSubKeys As Long, lpcMaxSubKeyLen As Long
Dim lpcbMaxClassLen As Long, lpcValues As Long
Dim lpcbMaxValueNameLen As Long, lpcbMaxValueLen As Long

Dim dwIndex As Long, lpcbName As Long
Dim lpName As String, lpClass As String

Dim FTime As FILETIME

lRetVal = OpenKey(hKey)
If lRetVal <> ERROR_SUCCESS Then
    GetAllKeys = lRetVal
    Exit Function
End If

lRetVal = RegQueryInfoKey(hKey, 0&, 0&, 0&, lpcSubKeys, _
  lpcMaxSubKeyLen, lpcbMaxClassLen, lpcValues, _
  lpcbMaxValueNameLen, lpcbMaxValueLen, 0&, FTime)
  
If lpcSubKeys = 0 Then _
   Err.Raise ERROR_EMPTY_ARRAY


ReDim sArray(lpcSubKeys - 1) ', 2)
    
For dwIndex = 0 To lpcSubKeys - 1
    lpName = space(lpcMaxSubKeyLen + 1)
    lpcbName = lpcMaxSubKeyLen + 1
    lRetVal = RegEnumKeyEx(hKey, dwIndex, lpName, _
       lpcbName, 0&, lpClass, Len(lpClass), FTime)
    sArray(dwIndex) = Left(lpName, lpcbName)
Next dwIndex

RegCloseKey (hKey)
GetAllKeys = lRetVal
Exit Function
Handler:
If Err.Number = ERROR_EMPTY_ARRAY Then
   Err.clear
   GetAllKeys = ERROR_EMPTY_ARRAY
Else
  GetAllKeys = lRetVal
End If
RegCloseKey (hKey)
End Function

Private Function OpenKey(ByRef hKey As Long) As ErrorList

If Len(m_key) = 0 Then
    hKey = m_root
    OpenKey = ERROR_SUCCESS
Else
    OpenKey = RegOpenKeyEx(m_root, m_key, 0, KEY_ALL_ACCESS, hKey)
End If

End Function

Public Property Get Root() As RegRoot
    Root = m_root
End Property

Public Property Let Root(ByVal idx As RegRoot)
    If idx < HKEY_CLASSES_ROOT Or idx > HKEY_DYN_DATA Then Exit Property
    m_root = idx
End Property

Public Property Get Key() As String
    Key = m_key
End Property
Public Property Let Key(ByVal sKey As String)
    Key = sKey
End Property

Public Function ModifyKey(strKeySuffix As String, flag As Boolean) As ErrorList
Dim strOldKey As String
strOldKey = m_key


If flag Then
    If Len(m_key) = 0 And Len(strKeySuffix) = 0 Then
        m_key = ""
    ElseIf Len(m_key) = 0 Then
        m_key = strKeySuffix
    ElseIf Len(strKeySuffix) = 0 Then
        m_key = m_key
    ElseIf Right(strKeySuffix, 1) = "\" Then
        m_key = m_key & strKeySuffix
    Else
        m_key = m_key & "\" & strKeySuffix
    End If
Else
    'm_key = Left(m_key, Len(m_key) - Len(strKeySuffix))
    If Len(m_key) >= Len(strKeySuffix) Then _
       m_key = Left(m_key, Len(m_key) - Len(strKeySuffix))
End If
Dim hKey As Long, lRetVal As Long
lRetVal = RegOpenKeyEx(m_root, m_key, 0, KEY_ALL_ACCESS, hKey)
If OpenKey(hKey) = ERROR_SUCCESS Then
    RegCloseKey (hKey)
Else
    m_key = strOldKey
End If
ModifyKey = lRetVal
End Function
