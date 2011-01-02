Attribute VB_Name = "MDeclare"
Option Explicit

' These are types and declares described in the book (especially Chapter 2).
' They are in a false conditionals so that the ones in the Windows API type
' library will override them. You can enable these versions to confirm that
' the Declare statements and type library entries are equivalent.
#Const fUseDeclares = 0
#If fUseDeclares Then

Type LARGE_INTEGER
    LowPart As Long
    HighPart As Long
End Type

Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

Type POINTL
    x As Long
    y As Long
End Type

Type WINDOWPLACEMENT
    length As Long
    Flags As Long
    showCmd As Long
    ptMinPosition As POINTL
    ptMaxPosition As POINTL
    rcNormalPosition As RECT
End Type

Declare Function Polygon Lib "GDI32" (ByVal hDC As Long, _
    lpPoints As POINTL, ByVal nCount As Long) As Long

Declare Function SetWindowPlacement Lib "USER32" ( _
    ByVal hWnd As Long, lpwndpl As WINDOWPLACEMENT) As Long
Declare Function GetWindowPlacement Lib "USER32" ( _
    ByVal hWnd As Long, lpwndpl As WINDOWPLACEMENT) As Long

Public Declare Sub CopyMemory Lib "KERNEL32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

Declare Function GetScrollRange Lib "USER32" (ByVal hWnd As Long, _
    ByVal nBar As Long, lpMin As Long, lpMax As Long) As Long

Declare Function FindWindow Lib "USER32" Alias "FindWindowA" ( _
    Optional ByVal Class As String, _
    Optional ByVal Title As String) As Long

Declare Function WindowFromPointXY Lib "USER32" _
    Alias "WindowFromPoint" (ByVal xPoint As Long, _
    ByVal yPoint As Long) As Long
Declare Function ChildWindowFromPointXY Lib "USER32" _
    Alias "WindowFromPoint" (ByVal hWnd As Long, _
    ByVal xPoint As Long, ByVal yPoint As Long) As Long
Declare Function WindowFromPoint Lib "USER32" _
    (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Declare Function VBGetObject Lib "GDI32" Alias "GetObjectA" ( _
    ByVal hObject As Long, ByVal cbBuffer As Long, _
    lpvObject As Any) As Long

Declare Function GetObjectBrush Lib "GDI32" Alias "GetObjectA" ( _
    ByVal hBrush As Long, ByVal cbBuffer As Long, _
    lpBrush As LOGBRUSH) As Long
 Declare Function GetObjectBitmap Lib "GDI32" Alias "GetObjectA" ( _
    ByVal hBitmap As Long, ByVal cbBuffer As Long, _
    lpBitmap As BITMAP) As Long

Declare Function SendMessageVal Lib "USER32" Alias "SendMessageA" ( _
    ByVal hWnd As Long, ByVal wMsg As Long, _
    ByVal wParam As Long, ByVal lParam As Any) As Long
Declare Function SendMessageStr Lib "USER32" Alias "SendMessageA" ( _
    ByVal hWnd As Long, ByVal wMsg As Long, _
    ByVal wParam As Long, ByVal lParam As String) As Long
Declare Function SendMessage Lib "USER32" Alias "SendMessageA" ( _
    ByVal hWnd As Long, ByVal wMsg As Long, _
    wParam As Any, lParam As Any) As Long

Declare Function GetWindowText Lib "USER32" Alias "GetWindowTextA" ( _
    ByVal hWnd As Long, ByVal lpString As String, _
    ByVal nMaxCount As Long) As Long

Declare Function GetWindowRect Lib "USER32" (ByVal hWnd As Long, _
    lprc As RECT) As Long
Declare Sub ClientToScreen Lib "USER32" (ByVal hWnd As Long, _
    lpPoint As POINTL)

Declare Function FloodFill Lib "GDI32" (ByVal hDC As Long, _
    ByVal nXStart As Long, ByVal nYStart As Long, _
    ByVal crFill As Long) As Long
    
Public Declare Function EnumWindows Lib "USER32" ( _
    ByVal lpEnumFunc As Long, lParam As Any) As Long

Private Declare Function SearchPath Lib "kernel32.dll" ( _
    ByVal lpPath As String, ByVal lpFileName As String, _
    ByVal lpExtension As String, ByVal nBufferLenght As Long, _
    ByVal lpBuffer As String, lpFilePart As Long) As Long

Declare Function QueryPerformanceCounter Lib "KERNEL32" ( _
    lpPerformanceCount As Currency) As Long

Declare Function QueryPerformanceFrequency Lib "KERNEL32" ( _
    lpFrequency As Currency) As Long
    
Declare Function GetTickCount Lib "KERNEL32" () As Long
    
Declare Function CharLower Lib "USER32" Alias "CharLowerA" ( _
    ByVal lpsz As String) As String
Declare Function CharLowerBuff Lib "USER32" Alias "CharLowerBuffA" ( _
    ByVal lpsz As String, ByVal cchLength As Long) As Long
Declare Function CharLowerBuffPtr Lib "USER32" Alias "CharLowerBuffA" ( _
    ByVal lpsz As Long, ByVal cchLength As Long) As Long
Declare Function CharUpper Lib "USER32" Alias "CharUpperA" ( _
    ByVal lpsz As String) As String
Declare Function CharUpperBuff Lib "USER32" Alias "CharUpperBuffA" ( _
    ByVal lpsz As String, ByVal cchLength As Long) As Long
Declare Function CharUpperBuffPtr Lib "USER32" Alias "CharUpperBuffA" ( _
    ByVal lpsz As Long, ByVal cchLength As Long) As Long
Declare Function CharNext Lib "USER32" Alias "CharNextA" ( _
    ByVal lpsz As String) As String
Declare Function CharPrev Lib "USER32" Alias "CharPrevA" ( _
    ByVal lpszStart As String, ByVal lpszCurrent As String) As String
Declare Function CharToOem Lib "USER32" Alias "CharToOemA" ( _
    ByVal lpszSrc As String, ByVal lpszDst As String) As Long
Declare Function OemToChar Lib "USER32" Alias "OemToCharA" ( _
    ByVal lpszSrc As String, ByVal lpszDst As String) As Long
Declare Function CharToOemBuff Lib "USER32" Alias "CharToOemBuffA" ( _
    ByVal lpszSrc As String, ByVal lpszDst As String, _
    ByVal cchDstLength As Long) As Long
Declare Function OemToCharBuff Lib "USER32" Alias "OemToCharBuffA" ( _
    ByVal lpszSrc As String, ByVal lpszDst As String, _
    ByVal cchDstLength As Long) As Long

Declare Function MultiByteToWideChar Lib "KERNEL32" ( _
    ByVal CodePage As Long, ByVal dwFlags As Long, _
    lpMultiByteStr As Any, ByVal cchMultiByte As Long, _
    lpWideCharStr As Any, ByVal cchWideChar As Long) As Long
Declare Function WideCharToMultiByte Lib "KERNEL32" ( _
    ByVal CodePage As Long, ByVal dwFlags As Long, _
    lpWideCharStr As Any, ByVal cchWideChar As Long, _
    lpMultiByteStr As Any, ByVal cchMultiByte As Long, _
    ByVal lpDefaultChar As String, ByVal lpUsedDefaultChar As Long) As Long

Declare Function CompareString Lib "KERNEL32" Alias "CompareStringA" ( _
    ByVal Locale As Long, ByVal dwCmpFlags As Long, _
    ByVal lpString1 As String, ByVal cchCount1 As Long, _
    ByVal lpString2 As String, ByVal cchCount2 As Long) As Long
Declare Function LCMapString Lib "KERNEL32" Alias "LCMapStringA" ( _
    ByVal Locale As Long, ByVal dwMapFlags As Long, _
    ByVal lpSrcStr As String, ByVal cchSrc As Long, _
    ByVal lpDestStr As String, ByVal cchDest As Long) As Long
Declare Function GetLocaleInfo Lib "KERNEL32" Alias "GetLocaleInfoA" ( _
    ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, _
    ByVal cchData As Long) As Long
Declare Function GetSystemDefaultLangID Lib "KERNEL32" () As Integer
Declare Function GetUserDefaultLangID Lib "KERNEL32" () As Integer
Declare Function GetSystemDefaultLCID Lib "KERNEL32" () As Long
Declare Function GetUserDefaultLCID Lib "KERNEL32" () As Long

' Locale Independent APIs

Declare Function GetStringTypeA Lib "KERNEL32" ( _
    ByVal lcid As Long, ByVal dwInfoType As Long, _
    ByVal lpSrcStr As String, ByVal cchSrc As Long, _
    lpCharType As Long) As Long
Declare Function FoldString Lib "KERNEL32" Alias "FoldStringA" ( _
    ByVal dwMapFlags As Long, ByVal lpSrcStr As String, _
    ByVal cchSrc As Long, ByVal lpDestStr As String, _
    ByVal cchDest As Long) As Long

' Language dependent Routines
Declare Function IsCharAlpha Lib "USER32" Alias "IsCharAlphaA" ( _
    ByVal cChar As Byte) As Long
Declare Function IsCharAlphaNumeric Lib "USER32" Alias "IsCharAlphaNumericA" ( _
    ByVal cChar As Byte) As Long
Declare Function IsCharUpper Lib "USER32" Alias "IsCharUpperA" ( _
    ByVal cChar As Byte) As Long
Declare Function IsCharLower Lib "USER32" Alias "IsCharLowerA" ( _
    ByVal cChar As Byte) As Long

Declare Function lstrcat Lib "KERNEL32" Alias "lstrcatA" ( _
    ByVal lpString1 As String, ByVal lpString2 As String) As Long
Declare Function lstrcpyn Lib "KERNEL32" Alias "lstrcpynA" ( _
    ByVal lpString1 As String, ByVal lpString2 As String, _
    ByVal iMaxLength As Long) As Long
Declare Function lstrcpynP Lib "KERNEL32" Alias "lstrcpynA" ( _
    lpString1 As Any, lpString2 As Any, _
    ByVal iMaxLength As Long) As Long
Declare Function lstrcpy Lib "KERNEL32" Alias "lstrcpyA" ( _
    ByVal lpString1 As String, ByVal lpString2 As String) As Long
Declare Function lstrcmp Lib "KERNEL32" Alias "lstrcmpA" ( _
    ByVal lpString1 As String, ByVal lpString2 As String) As Long
Declare Function lstrcmpi Lib "KERNEL32" Alias "lstrcmpiA" ( _
    ByVal lpString1 As String, ByVal lpString2 As String) As Long
Declare Function lstrlen Lib "KERNEL32" Alias "lstrlenA" ( _
    ByVal lpString As String) As Long

Public Const CP_ACP = 0         ' Default to ANSI code page
Public Const CP_OEMCP = 1       ' Default to OEM code page
Public Const CP_MACCP = 2       ' Default to MAC code page

#End If



