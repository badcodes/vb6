VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CRegistry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' =========================================================
' Class:    cRegistry
' Author:   Steve McMahon
' Date  :   21 Feb 1997
'
' A nice class wrapper around the registry functions
' Allows searching,deletion,modification and addition
' of Keys or Values.
'
' Updated 29 April 1998 for VB5.
'   * Fixed GPF in EnumerateValues
'   * Added support for all registry types, not just strings
'   * Put all declares in local class
'   * Added VB5 Enums
'   * Added CreateKey and DeleteKey methods
'
' Updated 2 January 1999
'   * The CreateExeAssociation method failed to set up the
'     association correctly if the optional document icon
'     was not provided.
'   * Added new parameters to CreateExeAssociation to set up
'     other standard handlers: Print, Add, New
'   * Provided the CreateAdditionalEXEAssociations method
'     to allow non-standard menu items to be added (for example,
'     right click on a .VBP file.  VB installs Run and Make
'     menu items).
'
' Updated 8 February 2000
'   * Ensure CreateExeAssociation and related items sets up the
'     registry keys in the
'           HKEY_LOCAL_MACHINE\SOFTWARE\Classes
'     branch as well as the HKEY_CLASSES_ROOT branch.
'
' ---------------------------------------------------------------------------
' vbAccelerator - free, advanced source code for VB programmers.
'     http://vbaccelerator.com
' =========================================================

'Registry Specific Access Rights
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SET_VALUE = &H2
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const KEY_CREATE_LINK = &H20
Private Const KEY_ALL_ACCESS = &H3F

'Open/Create Options
Private Const REG_OPTION_NON_VOLATILE = 0&
Private Const REG_OPTION_VOLATILE = &H1

'Key creation/open disposition
Private Const REG_CREATED_NEW_KEY = &H1
Private Const REG_OPENED_EXISTING_KEY = &H2

'masks for the predefined standard access types
Private Const STANDARD_RIGHTS_ALL = &H1F0000
Private Const SPECIFIC_RIGHTS_ALL = &HFFFF

'Define severity codes
Private Const ERROR_SUCCESS = 0&
Private Const ERROR_ACCESS_DENIED = 5
Private Const ERROR_INVALID_DATA = 13&
Private Const ERROR_MORE_DATA = 234 '  dderror
Private Const ERROR_NO_MORE_ITEMS = 259


'Structures Needed For Registry Prototypes
Private Type SECURITY_ATTRIBUTES
  nLength As Long
  lpSecurityDescriptor As Long
  bInheritHandle As Boolean
End Type

Private Type FILETIME
  dwLowDateTime As Long
  dwHighDateTime As Long
End Type

'Registry Function Prototypes
Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" _
  (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, _
  ByVal samDesired As Long, phkResult As Long) As Long

Private Declare Function RegSetValueExStr Lib "advapi32" Alias "RegSetValueExA" _
  (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, _
   ByVal dwType As Long, ByVal szData As String, ByVal cbData As Long) As Long
Private Declare Function RegSetValueExLong Lib "advapi32" Alias "RegSetValueExA" _
  (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, _
   ByVal dwType As Long, szData As Long, ByVal cbData As Long) As Long
Private Declare Function RegSetValueExByte Lib "advapi32" Alias "RegSetValueExA" _
  (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, _
   ByVal dwType As Long, szData As Byte, ByVal cbData As Long) As Long

Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long

Private Declare Function RegQueryValueExStr Lib "advapi32" Alias "RegQueryValueExA" _
  (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, _
   ByRef lpType As Long, ByVal szData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegQueryValueExLong Lib "advapi32" Alias "RegQueryValueExA" _
  (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, _
   ByRef lpType As Long, szData As Long, ByRef lpcbData As Long) As Long
Private Declare Function RegQueryValueExByte Lib "advapi32" Alias "RegQueryValueExA" _
  (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, _
   ByRef lpType As Long, szData As Byte, ByRef lpcbData As Long) As Long
   
Private Declare Function RegCreateKeyEx Lib "advapi32" Alias "RegCreateKeyExA" _
  (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, _
   ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, _
   lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, _
   lpdwDisposition As Long) As Long

Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" _
  (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, _
   lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, _
   lpcbClass As Long, lpftLastWriteTime As FILETIME) As Long

Private Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" ( _
    ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, _
    ByVal cbName As Long) As Long

Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" _
  (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, _
   lpcbValueName As Long, ByVal lpReserved As Long, ByVal lpType As Long, _
   ByVal lpData As Long, ByVal lpcbData As Long) As Long
   
Private Declare Function RegEnumValueLong Lib "advapi32.dll" Alias "RegEnumValueA" _
  (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, _
   lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, _
   lpData As Long, lpcbData As Long) As Long
Private Declare Function RegEnumValueStr Lib "advapi32.dll" Alias "RegEnumValueA" _
  (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, _
   lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, _
   ByVal lpData As String, lpcbData As Long) As Long
Private Declare Function RegEnumValueByte Lib "advapi32.dll" Alias "RegEnumValueA" _
  (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, _
   lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, _
   lpData As Byte, lpcbData As Long) As Long

Private Declare Function RegQueryInfoKey Lib "advapi32.dll" Alias "RegQueryInfoKeyA" _
   (ByVal hKey As Long, ByVal lpClass As String, _
   lpcbClass As Long, ByVal lpReserved As Long, lpcSubKeys As Long, _
   lpcbMaxSubKeyLen As Long, lpcbMaxClassLen As Long, lpcValues As Long, _
   lpcbMaxValueNameLen As Long, lpcbMaxValueLen As Long, lpcbSecurityDescriptor As Long, _
   lpftLastWriteTime As Any) As Long

Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" _
  (ByVal hKey As Long, ByVal lpSubKey As String) As Long

Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" _
  (ByVal hKey As Long, ByVal lpValueName As String) As Long

' Other declares:
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Private Declare Function ExpandEnvironmentStrings Lib "kernel32" Alias "ExpandEnvironmentStringsA" (ByVal lpSrc As String, ByVal lpDst As String, ByVal nSize As Long) As Long


Public Enum ERegistryClassConstants
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
End Enum

Public Enum ERegistryValueTypes
'Predefined Value Types
    REG_NONE = (0)                         'No value type
    REG_SZ = (1)                           'Unicode nul terminated string
    REG_EXPAND_SZ = (2)                    'Unicode nul terminated string w/enviornment var
    REG_BINARY = (3)                       'Free form binary
    REG_DWORD = (4)                        '32-bit number
    REG_DWORD_LITTLE_ENDIAN = (4)          '32-bit number (same as REG_DWORD)
    REG_DWORD_BIG_ENDIAN = (5)             '32-bit number
    REG_LINK = (6)                         'Symbolic Link (unicode)
    REG_MULTI_SZ = (7)                     'Multiple Unicode strings
    REG_RESOURCE_LIST = (8)                'Resource list in the resource map
    REG_FULL_RESOURCE_DESCRIPTOR = (9)     'Resource list in the hardware description
    REG_RESOURCE_REQUIREMENTS_LIST = (10)
End Enum

Private m_hClassKey As Long
Private m_sSectionKey As String
Private m_sValueKey As String
Private m_vValue As Variant
Private m_sSetValue As String
Private m_vDefault As Variant
Private m_eValueType As ERegistryValueTypes

Public Property Get KeyExists() As Boolean
    'KeyExists = bCheckKeyExists( _
    '                m_hClassKey, _
    '                m_sSectionKey _
    '            )
Dim hKey As Long
    If RegOpenKeyEx(m_hClassKey, m_sSectionKey, 0, 1, hKey) = ERROR_SUCCESS Then
        KeyExists = True
        RegCloseKey hKey
    Else
        KeyExists = False
    End If
    
End Property
Public Function CreateKey() As Boolean
Dim tSA As SECURITY_ATTRIBUTES
Dim hKey As Long
Dim lCreate As Long
Dim e As Long

    'Open or Create the key
    e = RegCreateKeyEx(m_hClassKey, m_sSectionKey, 0, "", REG_OPTION_NON_VOLATILE, _
                 KEY_ALL_ACCESS, tSA, hKey, lCreate)
    If e Then
        Err.Raise 26001, App.EXEName & ".cRegistry", "Failed to create registry Key: '" & m_sSectionKey
    Else
        CreateKey = (e = ERROR_SUCCESS)
        'Close the key
        RegCloseKey hKey
    End If
End Function
Public Function DeleteKey() As Boolean
Dim e As Long
    e = RegDeleteKey(m_hClassKey, m_sSectionKey)
    If e Then
        Err.Raise 26001, App.EXEName & ".cRegistry", "Failed to delete registry Key: '" & m_hClassKey & "',Section: '" & m_sSectionKey
    Else
        DeleteKey = (e = ERROR_SUCCESS)
    End If
    
End Function
Public Function DeleteValue() As Boolean
Dim e As Long
Dim hKey As Long

    e = RegOpenKeyEx(m_hClassKey, m_sSectionKey, 0, KEY_ALL_ACCESS, hKey)
    If e Then
        Err.Raise 26001, App.EXEName & ".cRegistry", "Failed to open key '" & m_hClassKey & "',Section: '" & m_sSectionKey & "' for delete access"
    Else
        e = RegDeleteValue(hKey, m_sValueKey)
        If e Then
            Err.Raise 26001, App.EXEName & ".cRegistry", "Failed to delete registry Key: '" & m_hClassKey & "',Section: '" & m_sSectionKey & "',Key: '" & m_sValueKey
        Else
            DeleteValue = (e = ERROR_SUCCESS)
        End If
    End If

End Function
Public Property Get Value() As Variant
Dim vValue As Variant
Dim cData As Long, sData As String, ordType As Long, e As Long
Dim hKey As Long

    e = RegOpenKeyEx(m_hClassKey, m_sSectionKey, 0, KEY_QUERY_VALUE, hKey)
    'ApiRaiseIf e

    e = RegQueryValueExLong(hKey, m_sValueKey, 0&, ordType, 0&, cData)
    If e And e <> ERROR_MORE_DATA Then
        Value = m_vDefault
        Exit Property
    End If
    
    m_eValueType = ordType
    Select Case ordType
    Case REG_DWORD, REG_DWORD_LITTLE_ENDIAN
        Dim iData As Long
        e = RegQueryValueExLong(hKey, m_sValueKey, 0&, _
                               ordType, iData, cData)
        vValue = CLng(iData)
        
    Case REG_DWORD_BIG_ENDIAN  ' Unlikely, but you never know
        Dim dwData As Long
        e = RegQueryValueExLong(hKey, m_sValueKey, 0&, _
                               ordType, dwData, cData)
        vValue = SwapEndian(dwData)
        
    Case REG_SZ, REG_MULTI_SZ ' Same thing to Visual Basic
        sData = String$(cData - 1, 0)
        e = RegQueryValueExStr(hKey, m_sValueKey, 0&, _
                               ordType, sData, cData)
        vValue = sData
        
    Case REG_EXPAND_SZ
        sData = String$(cData - 1, 0)
        e = RegQueryValueExStr(hKey, m_sValueKey, 0&, _
                               ordType, sData, cData)
        vValue = ExpandEnvStr(sData)
        
    ' Catch REG_BINARY and anything else
    Case Else
        Dim abData() As Byte
        ReDim abData(cData)
        e = RegQueryValueExByte(hKey, m_sValueKey, 0&, _
                                ordType, abData(0), cData)
        vValue = abData
        
    End Select
    Value = vValue
    
End Property
Public Property Let Value( _
        ByVal vValue As Variant _
    )
Dim ordType As Long
Dim c As Long
Dim hKey As Long
Dim e As Long
Dim lCreate As Long
Dim tSA As SECURITY_ATTRIBUTES

    'Open or Create the key
    e = RegCreateKeyEx(m_hClassKey, m_sSectionKey, 0, "", REG_OPTION_NON_VOLATILE, _
                 KEY_ALL_ACCESS, tSA, hKey, lCreate)
    
    If e Then
        Err.Raise 26001, App.EXEName & ".cRegistry", "Failed to set registry value Key: '" & m_hClassKey & "',Section: '" & m_sSectionKey & "',Key: '" & m_sValueKey & "' to value: '" & m_vValue & "'"
    Else

        Select Case m_eValueType
        Case REG_BINARY
            If (VarType(vValue) = vbArray + vbByte) Then
                Dim ab() As Byte
                ab = vValue
                ordType = REG_BINARY
                c = UBound(ab) - LBound(ab) - 1
                e = RegSetValueExByte(hKey, m_sValueKey, 0&, ordType, ab(0), c)
            Else
                Err.Raise 26001
            End If
        Case REG_DWORD, REG_DWORD_BIG_ENDIAN, REG_DWORD_LITTLE_ENDIAN
            If (VarType(vValue) = vbInteger) Or (VarType(vValue) = vbLong) Then
                Dim i As Long
                i = vValue
                ordType = REG_DWORD
                e = RegSetValueExLong(hKey, m_sValueKey, 0&, ordType, i, 4)
            End If
        Case REG_SZ, REG_EXPAND_SZ
            Dim s As String, iPos As Long
            s = vValue
            ordType = REG_SZ
            ' Assume anything with two non-adjacent percents is expanded string
            iPos = InStr(s, "%")
            If iPos Then
                If InStr(iPos + 2, s, "%") Then ordType = REG_EXPAND_SZ
            End If
            c = Len(s) + 1
            e = RegSetValueExStr(hKey, m_sValueKey, 0&, ordType, s, c)
            
        ' User should convert to a compatible type before calling
        Case Else
            e = ERROR_INVALID_DATA
            
        End Select
        
        If Not e Then
            m_vValue = vValue
        Else
            Err.Raise vbObjectError + 1048 + 26001, App.EXEName & ".cRegistry", "Failed to set registry value Key: '" & m_hClassKey & "',Section: '" & m_sSectionKey & "',Key: '" & m_sValueKey & "' to value: '" & m_vValue & "'"
        End If
        
        'Close the key
        RegCloseKey hKey
    
    End If
    
End Property
Public Function EnumerateValues( _
        ByRef sKeyNames() As String, _
        ByRef iKeyCount As Long _
    ) As Boolean
Dim lResult As Long
Dim hKey As Long
Dim sName As String
Dim lNameSize As Long
Dim sData As String
Dim lIndex As Long
Dim cJunk As Long
Dim cNameMax As Long
Dim ft As Currency
   
   ' Log "EnterEnumerateValues"

   iKeyCount = 0
   Erase sKeyNames()
    
   lIndex = 0
   lResult = RegOpenKeyEx(m_hClassKey, m_sSectionKey, 0, KEY_QUERY_VALUE, hKey)
   If (lResult = ERROR_SUCCESS) Then
      ' Log "OpenedKey:" & m_hClassKey & "," & m_sSectionKey
      lResult = RegQueryInfoKey(hKey, "", cJunk, 0, _
                               cJunk, cJunk, cJunk, cJunk, _
                               cNameMax, cJunk, cJunk, ft)
       Do While lResult = ERROR_SUCCESS
   
           'Set buffer space
           lNameSize = cNameMax + 1
           sName = String$(lNameSize, 0)
           If (lNameSize = 0) Then lNameSize = 1
           
           ' Log "Requesting Next Value"
         
           'Get value name:
           lResult = RegEnumValue(hKey, lIndex, sName, lNameSize, _
                                  0&, 0&, 0&, 0&)
           ' Log "RegEnumValue returned:" & lResult
           If (lResult = ERROR_SUCCESS) Then
       
                ' Although in theory you can also retrieve the actual
                ' value and type here, I found it always (ultimately) resulted in
                ' a GPF, on Win95 and NT.  Why?  Can anyone help?
       
               sName = Left$(sName, lNameSize)
               ' Log "Enumerated value:" & sName
                 
               iKeyCount = iKeyCount + 1
               ReDim Preserve sKeyNames(1 To iKeyCount) As String
               sKeyNames(iKeyCount) = sName
           End If
           lIndex = lIndex + 1
       Loop
   End If
   If (hKey <> 0) Then
      RegCloseKey hKey
   End If

   ' Log "Exit Enumerate Values"
   EnumerateValues = True
   Exit Function
   
EnumerateValuesError:
   If (hKey <> 0) Then
      RegCloseKey hKey
   End If
   Err.Raise vbObjectError + 1048 + 26003, App.EXEName & ".cRegistry", Err.Description
   Exit Function

End Function
Public Function EnumerateSections( _
        ByRef sSect() As String, _
        ByRef iSectCount As Long _
    ) As Boolean
Dim lResult As Long
Dim hKey As Long
Dim dwReserved As Long
Dim szBuffer As String
Dim lBuffSize As Long
Dim lIndex As Long
Dim lType As Long
Dim sCompKey As String
Dim iPos As Long

On Error GoTo EnumerateSectionsError

   iSectCount = 0
   Erase sSect
'
   lIndex = 0

   lResult = RegOpenKeyEx(m_hClassKey, m_sSectionKey, 0, KEY_ENUMERATE_SUB_KEYS, hKey)
   Do While lResult = ERROR_SUCCESS
       'Set buffer space
       szBuffer = String$(255, 0)
       lBuffSize = Len(szBuffer)
      
      'Get next value
       lResult = RegEnumKey(hKey, lIndex, szBuffer, lBuffSize)
                             
       If (lResult = ERROR_SUCCESS) Then
           iSectCount = iSectCount + 1
           ReDim Preserve sSect(1 To iSectCount) As String
           iPos = InStr(szBuffer, Chr$(0))
           If (iPos > 0) Then
              sSect(iSectCount) = Left(szBuffer, iPos - 1)
           Else
              sSect(iSectCount) = Left(szBuffer, lBuffSize)
           End If
       End If
       
       lIndex = lIndex + 1
   Loop
   If (hKey <> 0) Then
      RegCloseKey hKey
   End If
   EnumerateSections = True
   Exit Function

EnumerateSectionsError:
   If (hKey <> 0) Then
      RegCloseKey hKey
   End If
   Err.Raise vbObjectError + 1048 + 26002, App.EXEName & ".cRegistry", Err.Description
   Exit Function
End Function
Private Sub pSetClassValue(ByVal sValue As String)
Dim sSection As String
   ClassKey = HKEY_CLASSES_ROOT
   Value = sValue
   sSection = SectionKey
   ClassKey = HKEY_LOCAL_MACHINE
   SectionKey = "SOFTWARE\Classes\" & sSection
   Value = sValue
   SectionKey = sSection
End Sub
Public Sub CreateEXEAssociation( _
        ByVal sExePath As String, _
        ByVal sClassName As String, _
        ByVal sClassDescription As String, _
        ByVal sAssociation As String, _
        Optional ByVal sOpenMenuText As String = "&Open", _
        Optional ByVal bSupportPrint As Boolean = False, _
        Optional ByVal sPrintMenuText As String = "&Print", _
        Optional ByVal bSupportNew As Boolean = False, _
        Optional ByVal sNewMenuText As String = "&New", _
        Optional ByVal bSupportInstall As Boolean = False, _
        Optional ByVal sInstallMenuText As String = "", _
        Optional ByVal lDefaultIconIndex As Long = -1 _
    )
   ' Check if path is wrapped in quotes:
   sExePath = Trim$(sExePath)
   If (Left$(sExePath, 1) <> """") Then
      sExePath = """" & sExePath
   End If
   If (Right$(sExePath, 1) <> """") Then
      sExePath = sExePath & """"
   End If
    
    ' Create the .File to Class association:
   SectionKey = "." & sAssociation
   ValueType = REG_SZ
   ValueKey = ""
   pSetClassValue sClassName
   
   ' Create the Class shell open command:
   SectionKey = sClassName
   pSetClassValue sClassDescription
   
   SectionKey = sClassName & "\shell\open"
   If (sOpenMenuText = "") Then sOpenMenuText = "&Open"
   ValueKey = ""
   pSetClassValue sOpenMenuText
   SectionKey = sClassName & "\shell\open\command"
   ValueKey = ""
   pSetClassValue sExePath & " ""%1"""
   
   If (bSupportPrint) Then
      SectionKey = sClassName & "\shell\print"
      If (sPrintMenuText = "") Then sPrintMenuText = "&Print"
      ValueKey = ""
      pSetClassValue sPrintMenuText
      SectionKey = sClassName & "\shell\print\command"
      ValueKey = ""
      pSetClassValue sExePath & " /p ""%1"""
   End If
   
   If (bSupportInstall) Then
      If (sInstallMenuText = "") Then
         sInstallMenuText = "&Install " & sAssociation
      End If
      SectionKey = sClassName & "\shell\add"
      ValueKey = ""
      pSetClassValue sInstallMenuText
      SectionKey = sClassName & "\shell\add\command"
      ValueKey = ""
      pSetClassValue sExePath & " /a ""%1"""
   End If
   
   If (bSupportNew) Then
      SectionKey = sClassName & "\shell\new"
      ValueKey = ""
      If (sNewMenuText = "") Then sNewMenuText = "&New"
      pSetClassValue sNewMenuText
      SectionKey = sClassName & "\shell\new\command"
      ValueKey = ""
      pSetClassValue sExePath & " /n ""%1"""
   End If
   
   If lDefaultIconIndex > -1 Then
      SectionKey = sClassName & "\DefaultIcon"
      ValueKey = ""
      pSetClassValue sExePath & "," & CStr(lDefaultIconIndex)
   End If
    
End Sub
Public Sub CreateAdditionalEXEAssociations( _
      ByVal sClassName As String, _
      ParamArray vItems() As Variant _
   )
Dim iItems As Long
Dim iItem As Long
   
   On Error Resume Next
   iItems = UBound(vItems) + 1
   If (iItems Mod 3) <> 0 Or (Err.Number <> 0) Then
      Err.Raise vbObjectError + 1048 + 26004, App.EXEName & ".cRegistry", "Invalid parameter list passed to CreateAdditionalEXEAssociations - expected Name/Text/Command"
   Else
      ' Check if it exists:
      SectionKey = sClassName
      If Not (KeyExists) Then
         Err.Raise vbObjectError + 1048 + 26005, App.EXEName & ".cRegistry", "Error - attempt to create additional associations before class defined."
      Else
         For iItem = 0 To iItems - 1 Step 3
            ValueType = REG_SZ
            SectionKey = sClassName & "\shell\" & vItems(iItem)
            ValueKey = ""
            pSetClassValue vItems(iItem + 1)
            SectionKey = sClassName & "\shell\" & vItems(iItem) & "\command"
            ValueKey = ""
            pSetClassValue vItems(iItem + 2)
         Next iItem
      End If
   End If
   
End Sub
Public Property Get ValueType() As ERegistryValueTypes
    ValueType = m_eValueType
End Property
Public Property Let ValueType(ByVal eValueType As ERegistryValueTypes)
    m_eValueType = eValueType
End Property
Public Property Get ClassKey() As ERegistryClassConstants
    ClassKey = m_hClassKey
End Property
Public Property Let ClassKey( _
        ByVal eKey As ERegistryClassConstants _
    )
    m_hClassKey = eKey
End Property
Public Property Get SectionKey() As String
    SectionKey = m_sSectionKey
End Property
Public Property Let SectionKey( _
        ByVal sSectionKey As String _
    )
    m_sSectionKey = sSectionKey
End Property
Public Property Get ValueKey() As String
    ValueKey = m_sValueKey
End Property
Public Property Let ValueKey( _
        ByVal sValueKey As String _
    )
    m_sValueKey = sValueKey
End Property
Public Property Get Default() As Variant
    Default = m_vDefault
End Property
Public Property Let Default( _
        ByVal vDefault As Variant _
    )
    m_vDefault = vDefault
End Property
Private Function SwapEndian(ByVal dw As Long) As Long
    CopyMemory ByVal VarPtr(SwapEndian) + 3, dw, 1
    CopyMemory ByVal VarPtr(SwapEndian) + 2, ByVal VarPtr(dw) + 1, 1
    CopyMemory ByVal VarPtr(SwapEndian) + 1, ByVal VarPtr(dw) + 2, 1
    CopyMemory SwapEndian, ByVal VarPtr(dw) + 3, 1
End Function
Private Function ExpandEnvStr(sData As String) As String
    Dim c As Long, s As String
    ' Get the length
    s = "" ' Needed to get around Windows 95 limitation
    c = ExpandEnvironmentStrings(sData, s, c)
    ' Expand the string
    s = String$(c - 1, 0)
    c = ExpandEnvironmentStrings(sData, s, c)
    ExpandEnvStr = s
End Function
