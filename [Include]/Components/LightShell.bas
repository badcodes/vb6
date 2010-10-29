Attribute VB_Name = "MLightShell"
Option Explicit

Enum SHREGDEL_FLAGS
    SHREGDEL_DEFAULT = &H0              ' Delete's HKCU, or HKLM if HKCU is not found.
    SHREGDEL_HKCU = &H1                 ' Delete HKCU only
    SHREGDEL_HKLM = &H10                ' Delete HKLM only.
    SHREGDEL_BOTH = &H11                ' Delete both HKCU and HKLM.
End Enum

Enum SHREGENUM_FLAGS
    SHREGENUM_DEFAULT = &H0              ' Enumerates HKCU or HKLM if not found.
    SHREGENUM_HKCU = &H1                 ' Enumerates HKCU only
    SHREGENUM_HKLM = &H10                ' Enumerates HKLM only.
    SHREGENUM_BOTH = &H11                ' Enumerates both HKCU and HKLM without duplicates.
End Enum

'
'=============== String Routines ===================================
'
Declare Function StrChr Lib "SHLWAPI" Alias "StrChrW" ( _
    ByVal lStart As Long, ByVal wMatch As Integer) As Long

Declare Function StrChrI Lib "SHLWAPI" Alias "StrChrIW" ( _
    ByVal lpStart As Long, ByVal wMatch As Integer) As Long

Declare Function StrCmpN Lib "SHLWAPI" Alias "StrCmpNW" ( _
    ByVal lpStr1 As Long, ByVal lpStr2 As Long, _
    ByVal nChar As Long) As Long

Declare Function StrNCmp Lib "SHLWAPI" Alias "StrCmpNW" ( _
    ByVal lpStr1 As Long, ByVal lpStr2 As Long, _
    ByVal nChar As Long) As Long

Declare Function StrCmpNI Lib "SHLWAPI" Alias "StrCmpNIW" ( _
    ByVal lpStr1 As Long, ByVal lpStr2 As Long, _
    ByVal nChar As Long) As Long

Declare Function StrNCmpI Lib "SHLWAPI" Alias "StrCmpNIW" ( _
    ByVal lpStr1 As Long, ByVal lpStr2 As Long, ByVal nChar As Long) As Long

Declare Function StrCSpn Lib "SHLWAPI" Alias "StrCSpnW" ( _
    ByVal lpStr As Long, ByVal lpSet As Long) As Long

Declare Function StrCSpnI Lib "SHLWAPI" Alias "StrCSpnIW" ( _
    ByVal lpStr As Long, ByVal lpSet As Long) As Long

' Not a good idea for Visual Basic!!!
'Declare Function StrDup Lib "SHLWAPI" Alias "StrDupW" ( _
    ByVal lpSrch As Long) As Long

Declare Function StrFormatByteSize Lib "SHLWAPI" _
    Alias "StrFormatByteSizeW" (ByVal qdw As Currency, _
    ByVal szBuf As Long, ByVal uiBufSize As Long) As Long

Declare Function StrFromTimeInterval Lib "SHLWAPI" _
    Alias "StrFromTimeIntervalW" ( _
    ByVal pszOut As Long, ByVal cchMax As Long, _
    ByVal dwTimeMS As Long, ByVal digits As Long) As Long

Declare Function StrIsIntlEqual Lib "SHLWAPI" Alias "StrIsIntlEqualW" ( _
    ByVal fCaseSens As Long, ByVal lpString1 As Long, _
    ByVal lpString2 As Long, ByVal nChar As Long) As Long

Declare Function StrNCat Lib "SHLWAPI" Alias "StrNCatW" ( _
    ByVal psz1 As Long, ByVal psz2 As Long, ByVal cchMax As Long) As Long

Declare Function StrCatN Lib "SHLWAPI" Alias "StrNCatW" ( _
    ByVal psz1 As Long, ByVal psz2 As Long, ByVal cchMax As Long) As Long

Declare Function StrPBrk Lib "SHLWAPI" Alias "StrPBrkW" ( _
    ByVal psz As Long, ByVal pszSet As Long) As Long

Declare Function StrRChr Lib "SHLWAPI" Alias "StrRChrW" ( _
    ByVal lpStart As Long, ByVal lpEnd As Long, _
    ByVal wMatch As Integer) As Long

Declare Function StrRChrI Lib "SHLWAPI" Alias "StrRChrIW" ( _
    ByVal lpStart As Long, ByVal lpEnd As Long, _
    ByVal wMatch As Integer) As Long

Declare Function StrRStrI Lib "SHLWAPI" Alias "StrRStrIW" ( _
    ByVal lpSource As Long, ByVal lpLast As Long, _
    ByVal lpSrch As Long) As Long

Declare Function StrSpn Lib "SHLWAPI" Alias "StrSpnW" ( _
    ByVal psz As Long, ByVal pszSet As Long) As Long

Declare Function StrStr Lib "SHLWAPI" Alias "StrStrW" ( _
    ByVal lpFirst As Long, ByVal lpSrch As Long) As Long

Declare Function StrStrI Lib "SHLWAPI" Alias "StrStrIW" ( _
    ByVal lpFirst As Long, ByVal lpSrch As Long) As Long

Declare Function StrToInt Lib "SHLWAPI" Alias "StrToIntW" ( _
    ByVal lpSrc As Long) As Long

Declare Function StrToLong Lib "SHLWAPI" Alias "StrToIntW" ( _
    ByVal lpSrc As Long) As Long

Declare Function StrToIntEx Lib "SHLWAPI" Alias "StrToIntExW" ( _
    ByVal pszString As Long, ByVal dwFlags As Long, _
    piRet As Long) As Long

Declare Function StrTrim Lib "SHLWAPI" Alias "StrTrimW" ( _
    ByVal psz As Long, ByVal pszTrimChars As Long) As Long

Declare Function StrCat Lib "SHLWAPI" Alias "StrCatW" ( _
    ByVal psz1 As Long, ByVal psz2 As Long) As Long

Declare Function StrCmp Lib "SHLWAPI" Alias "StrCmpW" ( _
    ByVal psz1 As Long, ByVal psz2 As Long) As Long

Declare Function StrCmpI Lib "SHLWAPI" Alias "StrCmpIW" ( _
    ByVal psz1 As Long, ByVal psz2 As Long) As Long

Declare Function StrCpy Lib "SHLWAPI" Alias "StrCpyW" ( _
    ByVal psz1 As Long, ByVal psz2 As Long) As Long

Declare Function StrCpyN Lib "SHLWAPI" Alias "StrCpyNW" ( _
    ByVal psz1 As Long, ByVal psz2 As Long, _
    ByVal cchMax As Long) As Long

Declare Function StrNCpy Lib "SHLWAPI" Alias "StrCpyNW" ( _
    ByVal psz1 As Long, ByVal psz2 As Long, _
    ByVal cchMax As Long) As Long

Declare Function ChrCmpI Lib "SHLWAPI" Alias "ChrCmpIW" ( _
    ByVal w1 As Integer, ByVal w2 As Integer) As Long

' Flags for StrToIntEx
Const STIF_DEFAULT = &H0
Const STIF_SUPPORT_HEX = &H1

'
'=============== Path Routines ===================================
'

Declare Function PathAddBackslash Lib "SHLWAPI" Alias "PathAddBackslashW" ( _
    ByVal pszPath As Long) As Long

Declare Function PathAddExtension Lib "SHLWAPI" Alias "PathAddExtensionW" ( _
    ByVal pszPath As Long, ByVal pszExt As Long) As Long

Declare Function PathAppend Lib "SHLWAPI" Alias "PathAppendW" ( _
    ByVal pszPath As Long, ByVal pMore As Long) As Long

Declare Function PathBuildRoot Lib "SHLWAPI" Alias "PathBuildRootW" ( _
    ByVal szRoot As Long, ByVal iDrive As Long) As Long

Declare Function PathCanonicalize Lib "SHLWAPI" Alias "PathCanonicalizeW" ( _
    ByVal pszBuf As Long, ByVal pszPath As Long) As Long

Declare Function PathCombine Lib "SHLWAPI" Alias "PathCombineW" ( _
    ByVal szDest As Long, ByVal lpszDir As Long, _
    ByVal lpszFile As Long) As Long

Declare Function PathCompactPath Lib "SHLWAPI" Alias "PathCompactPathW" ( _
    ByVal hDC As Long, ByVal pszPath As Long, _
    ByVal dx As Long) As Long

Declare Function PathCompactPathEx Lib "SHLWAPI" Alias "PathCompactPathExW" ( _
    ByVal pszOut As Long, ByVal pszSrc As Long, _
    ByVal cchMax As Long, ByVal dwFlags As Long) As Long

Declare Function PathCommonPrefix Lib "SHLWAPI" Alias "PathCommonPrefixW" ( _
    ByVal pszFile1 As Long, ByVal pszFile2 As Long, _
    ByVal achPath As Long) As Long

Declare Function PathFileExists Lib "SHLWAPI" Alias "PathFileExistsW" ( _
    ByVal pszPath As Long) As Long

Declare Function PathFindExtension Lib "SHLWAPI" Alias "PathFindExtensionW" ( _
    ByVal pszPath As Long) As Long

Declare Function PathFindFileName Lib "SHLWAPI" Alias "PathFindFileNameW" ( _
    ByVal pszPath As Long) As Long

Declare Function PathFindNextComponent Lib "SHLWAPI" _
    Alias "PathFindNextComponentW" (ByVal pszPath As Long) As Long

Declare Function PathFindOnPath Lib "SHLWAPI" Alias "PathFindOnPathW" ( _
    ByVal pszPath As Long, ByVal ppszOtherDirs As Long) As Long

Declare Function PathGetArgs Lib "SHLWAPI" Alias "PathGetArgsW" ( _
    ByVal pszPath As Long) As Long

Declare Function PathGetCharType Lib "SHLWAPI" Alias "PathGetCharTypeW" ( _
    ByVal ch As Integer) As Long

' Return flags for PathGetCharType
Const GCT_INVALID = &H0
Const GCT_LFNCHAR = &H1
Const GCT_SHORTCHAR = &H2
Const GCT_WILD = &H4
Const GCT_SEPARATOR = &H8

Declare Function PathGetDriveNumber Lib "SHLWAPI" _
    Alias "PathGetDriveNumberW" (ByVal pszPath As Long) As Long

Declare Function PathIsDirectory Lib "SHLWAPI" Alias "PathIsDirectoryW" ( _
    ByVal pszPath As Long) As Long

Declare Function PathIsFileSpec Lib "SHLWAPI" Alias "PathIsFileSpecW" ( _
    ByVal pszPath As Long) As Long

Declare Function PathIsPrefix Lib "SHLWAPI" Alias "PathIsPrefixW" ( _
    ByVal pszPrefix As Long, ByVal pszPath As Long) As Long

Declare Function PathIsRelative Lib "SHLWAPI" Alias "PathIsRelativeW" ( _
    ByVal pszPath As Long) As Long

Declare Function PathIsRoot Lib "SHLWAPI" Alias "PathIsRootW" ( _
    ByVal pszPath As Long) As Long

Declare Function PathIsSameRoot Lib "SHLWAPI" Alias "PathIsSameRootW" ( _
    ByVal pszPath1 As Long, ByVal pszPath2 As Long) As Long

Declare Function PathIsUNC Lib "SHLWAPI" Alias "PathIsUNCW" ( _
    ByVal pszPath As Long) As Long

Declare Function PathIsUNCServer Lib "SHLWAPI" Alias "PathIsUNCServerW" ( _
    ByVal pszPath As Long) As Long

Declare Function PathIsUNCServerShare Lib "SHLWAPI" _
    Alias "PathIsUNCServerShareW" (ByVal pszPath As Long) As Long

Declare Function PathIsContentType Lib "SHLWAPI" Alias "PathIsContentTypeW" ( _
    ByVal pszPath As Long, ByVal pszContentType As Long) As Long

Declare Function PathIsURL Lib "SHLWAPI" Alias "PathIsURLW" ( _
    ByVal pszPath As Long) As Long

Declare Function PathMakePretty Lib "SHLWAPI" Alias "PathMakePrettyW" ( _
    ByVal pszPath As Long) As Long

Declare Function PathMatchSpec Lib "SHLWAPI" Alias "PathMatchSpecW" ( _
    ByVal pszFile As Long, ByVal pszSpec As Long) As Long

Declare Function PathParseIconLocation Lib "SHLWAPI" _
    Alias "PathParseIconLocationW" (ByVal pszIconFile As Long) As Long

Declare Function PathQuoteSpaces Lib "SHLWAPI" Alias "PathQuoteSpacesW" ( _
    ByVal lpsz As Long) As Long

Declare Function PathRelativePathTo Lib "SHLWAPI" _
    Alias "PathRelativePathToW" (ByVal pszPath As Long, _
    ByVal pszFrom As Long, ByVal dwAttrFrom As Long, _
    ByVal pszTo As Long, ByVal dwAttrTo As Long) As Long

Declare Function PathRemoveArgs Lib "SHLWAPI" Alias "PathRemoveArgsW" ( _
ByVal pszPath As Long)

Declare Function PathRemoveBackslash Lib "SHLWAPI" _
    Alias "PathRemoveBackslashW" (ByVal pszPath As Long) As Long

Declare Function PathRemoveBlanks Lib "SHLWAPI" _
    Alias "PathRemoveBlanksW" (ByVal pszPath As Long) As Long

Declare Function PathRemoveExtension Lib "SHLWAPI" _
    Alias "PathRemoveExtensionW" (ByVal pszPath As Long) As Long

Declare Function PathRemoveFileSpec Lib "SHLWAPI" _
    Alias "PathRemoveFileSpecW" (ByVal pszPath As Long) As Long

Declare Function PathRenameExtension Lib "SHLWAPI" _
    Alias "PathRenameExtensionW" (ByVal pszPath As Long, _
    ByVal pszExt As Long) As Long

Declare Function PathSearchAndQualify Lib "SHLWAPI" _
    Alias "PathSearchAndQualifyW" (ByVal pszPath As Long, _
    ByVal pszBuf As Long, ByVal cchBuf As Long) As Long

Declare Function PathSetDlgItemPath Lib "SHLWAPI" _
    Alias "PathSetDlgItemPathW" (ByVal hDlg As Long, _
    ByVal id As Long, ByVal pszPath As Long) As Long

Declare Function PathSkipRoot Lib "SHLWAPI" Alias "PathSkipRootW" ( _
    ByVal pszPath As Long) As Long

Declare Function PathStripPath Lib "SHLWAPI" Alias "PathStripPathW" ( _
    ByVal pszPath As Long) As Long

Declare Function PathStripToRoot Lib "SHLWAPI" Alias "PathStripToRootW" ( _
    ByVal pszPath As Long) As Long

Declare Function PathUnquoteSpaces Lib "SHLWAPI" _
    Alias "PathUnquoteSpacesW" (ByVal lpsz As Long) As Long

Declare Function PathMakeSystemFolder Lib "SHLWAPI" _
    Alias "PathMakeSystemFolderW" (ByVal pszPath As Long) As Long

Declare Function PathUnmakeSystemFolder Lib "SHLWAPI" _
    Alias "PathUnmakeSystemFolderW" (ByVal pszPath As Long) As Long

Declare Function PathIsSystemFolder Lib "SHLWAPI" _
    Alias "PathIsSystemFolderW" ( _
    ByVal pszPath As Long, ByVal dwAttrib As Long) As Long
'
'=============== Registry Routines ===================================
'

' SHDeleteEmptyKey mimics RegDeleteKey as it behaves on NT.
' SHDeleteKey mimics RegDeleteKey as it behaves on Win95.

Declare Function SHDeleteEmptyKey Lib "SHLWAPI" Alias "SHDeleteEmptyKeyW" ( _
    ByVal hKey As Long, ByVal pszSubKey As Long) As Long

Declare Function SHDeleteKey Lib "SHLWAPI" Alias "SHDeleteKeyW" ( _
    ByVal hKey As Long, ByVal pszSubKey As Long) As Long

' These functions open the key, get/set/delete the value, then close
' the key.

Declare Function SHDeleteValue Lib "SHLWAPI" Alias "SHDeleteValueW" ( _
    ByVal hKey As Long, ByVal pszSubKey As Long, _
    ByVal pszValue As Long) As Long

Declare Function SHGetValue Lib "SHLWAPI" Alias "SHGetValueW" ( _
    ByVal hKey As Long, ByVal pszSubKey As Long, ByVal pszValue As Long, _
    pdwType As Long, pvData As Any, pcbData As Long) As Long

Declare Function SHGetValueStr Lib "SHLWAPI" Alias "SHGetValueW" ( _
    ByVal hKey As Long, ByVal pszSubKey As Long, ByVal pszValue As Long, _
    pdwType As Long, ByVal pvData As Long, pcbData As Long) As Long

Declare Function SHSetValue Lib "SHLWAPI" Alias "SHSetValueW" ( _
    ByVal hKey As Long, ByVal pszSubKey As Long, ByVal pszValue As Long, _
    ByVal dwType As Long, pvData As Any, ByVal cbData As Long) As Long

Declare Function SHSetValueStr Lib "SHLWAPI" Alias "SHSetValueW" ( _
    ByVal hKey As Long, ByVal pszSubKey As Long, _
    ByVal pszValue As Long, ByVal dwType As Long, _
    ByVal pvData As Long, ByVal cbData As Long) As Long

' These functions work just like RegQueryValueEx, except if the
' data type is REG_EXPAND_SZ, then these will go ahead and expand
' out the string.  *pdwType will always be massaged to REG_SZ
' if this happens.  REG_SZ values are also guaranteed to be null
' terminated.

Declare Function SHQueryValueEx Lib "SHLWAPI" Alias "SHQueryValueExW" ( _
    ByVal hKey As Long, ByVal pszValue As Long, _
    ByVal pdwReserved As Long, pdwType As Long, _
    pvData As Any, pcbData As Long) As Long

Declare Function SHQueryValueExStr Lib "SHLWAPI" Alias "SHQueryValueExW" ( _
    ByVal hKey As Long, ByVal pszValue As Long, _
    ByVal pdwReserved As Long, pdwType As Long, _
    ByVal pvData As Long, pcbData As Long) As Long

' Enumeration functions support.

Declare Function SHEnumKeyEx Lib "SHLWAPI" Alias "SHEnumKeyExW" ( _
    ByVal hKey As Long, ByVal dwIndex As Long, _
    ByVal pszName As Long, pcchName As Long) As Long

Declare Function SHEnumValue Lib "SHLWAPI" Alias "SHEnumValueW" ( _
    ByVal hKey As Long, ByVal dwIndex As Long, _
    ByVal pszValueName As Long, pcchValueName As Long, _
    pdwType As Long, pvData As Any, pcbData As Long) As Long

Declare Function SHEnumValueStr Lib "SHLWAPI" Alias "SHEnumValueW" ( _
    ByVal hKey As Long, ByVal dwIndex As Long, _
    ByVal pszValueName As Long, pcchValueName As Long, _
    pdwType As Long, ByVal pvData As Long, pcbData As Long) As Long

Declare Function SHQueryInfoKey Lib "SHLWAPI" Alias "SHQueryInfoKeyW" ( _
    ByVal hKey As Long, pcSubKeys As Long, _
    pcchMaxSubKeyLen As Long, pcValues As Long, _
pcchMaxValueNameLen As Long) As Long

'''''''''''''''''''''''
' User Specific Registry Access Functions
'''''''''''''''''''''''

' Flags for StrToIntEx
Const SHREGSET_HKCU = &H1                        ' Write to HKCU if empty
Const SHREGSET_FORCE_HKCU = &H2                  ' Write to HKCU
Const SHREGSET_HKLM = &H4                        ' Write to HKLM if empty
Const SHREGSET_FORCE_HKLM = &H8                  ' Write to HKLM
Const SHREGSET_DEFAULT = &H6                     ' Default is SHREGSET_FORCE_HKCU | SHREGSET_HKLM.

Declare Function SHRegCreateUSKey Lib "SHLWAPI" Alias "SHRegCreateUSKeyW" ( _
    ByVal pwzPath As Long, ByVal samDesired As Long, _
    ByVal hRelativeUSKey As Long, phNewUSKey As Long, _
    ByVal dwFlags As Long) As Long

Declare Function SHRegOpenUSKey Lib "SHLWAPI" Alias "SHRegOpenUSKeyW" ( _
    ByVal pwzPath As Long, ByVal samDesired As Long, _
    ByVal hRelativeUSKey As Long, phNewUSKey As Long, _
    ByVal fIgnoreHKCU As Long) As Long

Declare Function SHRegQueryUSValue Lib "SHLWAPI" Alias "SHRegQueryUSValueW" ( _
    ByVal hUSKey As Long, ByVal pwzValue As Long, _
    pdwType As Long, pvData As Any, _
    pcbData As Long, ByVal fIgnoreHKCU As Long, _
    pvDefaultData As Any, ByVal dwDefaultDataSize As Long) As Long

Declare Function SHRegQueryUSValueStr Lib "SHLWAPI" _
    Alias "SHRegQueryUSValueW" ( _
    ByVal hUSKey As Long, ByVal pwzValue As Long, _
    pdwType As Long, ByVal pvData As Long, _
    pcbData As Long, ByVal fIgnoreHKCU As Long, _
    ByVal pvDefaultData As Long, ByVal dwDefaultDataSize As Long) As Long

Declare Function SHRegWriteUSValue Lib "SHLWAPI" _
    Alias "SHRegWriteUSValueW" ( _
    ByVal hUSKey As Long, ByVal pwzValue As Long, ByVal dwType As Long, _
    pvData As Any, ByVal cbData As Long, ByVal dwFlags As Long) As Long

Declare Function SHRegWriteUSValueStr Lib "SHLWAPI" _
    Alias "SHRegWriteUSValueW" ( _
    ByVal hUSKey As Long, ByVal pwzValue As Long, ByVal dwType As Long, _
    ByVal pvData As Long, ByVal cbData As Long, ByVal dwFlags As Long) As Long

Declare Function SHRegDeleteEmptyUSKey Lib "SHLWAPI" _
    Alias "SHRegDeleteEmptyUSKeyW" ( _
    ByVal hUSKey As Long, ByVal pwzSubKey As Long, _
    ByVal delRegFlags As SHREGDEL_FLAGS) As Long

Declare Function SHRegDeleteUSValue Lib "SHLWAPI" _
    Alias "SHRegDeleteUSValueW" ( _
    ByVal hUSKey As Long, ByVal pwzValue As Long, _
    ByVal delRegFlags As SHREGDEL_FLAGS) As Long

Declare Function SHRegEnumUSKey Lib "SHLWAPI" Alias "SHRegEnumUSKeyW" ( _
    ByVal hUSKey As Long, ByVal dwIndex As Long, ByVal pwzName As Long, _
    pcchName As Long, ByVal enumRegFlags As SHREGENUM_FLAGS) As Long

Declare Function SHRegEnumUSValue Lib "SHLWAPI" Alias "SHRegEnumUSValueW" ( _
    ByVal hUSKey As Long, ByVal dwIndex As Long, _
    ByVal pszValueName As Long, pcchValueName As Long, _
    pdwType As Long, pvData As Any, _
    pcbData As Long, ByVal enumRegFlags As SHREGENUM_FLAGS) As Long

Declare Function SHRegEnumUSValueStr Lib "SHLWAPI" _
    Alias "SHRegEnumUSValueW" ( _
    ByVal hUSKey As Long, ByVal dwIndex As Long, _
    ByVal pszValueName As Long, pcchValueName As Long, _
    pdwType As Long, ByVal pvData As Long, _
    pcbData As Long, ByVal enumRegFlags As SHREGENUM_FLAGS) As Long

Declare Function SHRegQueryInfoUSKey Lib "SHLWAPI" _
    Alias "SHRegQueryInfoUSKeyW" (ByVal hUSKey As Long, _
    pcSubKeys As Long, pcchMaxSubKeyLen As Long, pcValues As Long, _
    pcchMaxValueNameLen As Long, ByVal enumRegFlags As SHREGENUM_FLAGS) As Long

Declare Function SHRegCloseUSKey Lib "SHLWAPI" (ByVal hUSKey As Long) As Long

' These calls are equal to an SHRegOpenUSKey, SHRegQueryUSValue, and then a SHRegCloseUSKey

Declare Function SHRegGetUSValue Lib "SHLWAPI" Alias "SHRegGetUSValueW" ( _
    ByVal pwzSubKey As Long, ByVal pwzValue As Long, _
    pdwType As Long, pvData As Any, pcbData As Long, _
    ByVal fIgnoreHKCU As Long, pvDefaultData As Any, _
    ByVal dwDefaultDataSize As Long) As Long

Declare Function SHRegGetUSValueStr Lib "SHLWAPI" Alias "SHRegGetUSValueW" ( _
    ByVal pwzSubKey As Long, ByVal pwzValue As Long, _
    pdwType As Long, ByVal pvData As Long, pcbData As Long, _
    ByVal fIgnoreHKCU As Long, ByVal pvDefaultData As Long, _
    ByVal dwDefaultDataSize As Long) As Long

Declare Function SHRegSetUSValue Lib "SHLWAPI" Alias "SHRegSetUSValueW" ( _
    ByVal pwzSubKey As Long, ByVal pwzValue As Long, _
    ByVal dwType As Long, pvData As Any, ByVal cbData As Long, _
    ByVal dwFlags As Long) As Long

Declare Function SHRegSetUSValueStr Lib "SHLWAPI" Alias "SHRegSetUSValueW" ( _
    ByVal pwzSubKey As Long, ByVal pwzValue As Long, _
    ByVal dwType As Long, ByVal pvData As Long, _
    ByVal cbData As Long, ByVal dwFlags As Long) As Long

'
'====== GDI helper functions  ================================================
'

Declare Function SHCreateShellPalette Lib "SHLWAPI" ( _
        ByVal hDC As Long) As Long

' Call this as run-time to see if
Function HasShellWapi() As Boolean
    On Error Resume Next
    Call StrToLong(StrPtr("1"))
    If Err = 0 Then HasShellWapi = True
End Function





