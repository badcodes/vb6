VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GApiFileSystem"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Depends
'MApiTime

Public Declare Function API_FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function API_FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function API_FindClose Lib "kernel32" Alias "FindClose" (ByVal hFindFile As Long) As Long

Public Const MAX_PATH = 260
Public Type WIN32_FIND_DATA
        dwFileAttributes As enumFileAttribute
        ftCreationTime As FILETIME
        ftLastAccessTime As FILETIME
        ftLastWriteTime As FILETIME
        nFileSizeHigh As Long
        nFileSizeLow As Long
        dwReserved0 As Long
        dwReserved1 As Long
        cFileName As String * MAX_PATH
        cAlternate As String * 14
End Type

Public Enum enumFileAttribute
 File_attribute_archive = &H20
 FILE_ATTRIBUTE_COMPRESSED = &H800
 File_attribute_directory = &H10
 FILE_ATTRIBUTE_HIDDEN = &H2
 FILE_ATTRIBUTE_NORMAL = &H80
 FILE_ATTRIBUTE_READONLY = &H1
 FILE_ATTRIBUTE_SYSTEM = &H4
 FILE_ATTRIBUTE_TEMPORARY = &H100
End Enum


