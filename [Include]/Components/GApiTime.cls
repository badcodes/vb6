VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GApiTime"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Declare Function API_FileTimeToLocalFileTime Lib "kernel32" Alias "FileTimeToLocalFileTime" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long
Public Declare Function API_FileTimeToSystemTime Lib "kernel32" Alias "FileTimeToSystemTime" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long

Public Type SYSTEMTIME
        wYear As Integer
        wMonth As Integer
        wDayOfWeek As Integer
        wDay As Integer
        wHour As Integer
        wMinute As Integer
        wSecond As Integer
        wMilliseconds As Integer
End Type

Public Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type


' Convert a SYSTEMTIME into a Date.
Public Function FileTimeToDate(ftSrc As FILETIME) As Date
    Dim system_time As SYSTEMTIME
    API_FileTimeToSystemTime ftSrc, system_time
    With system_time
        FileTimeToDate = DateSerial(.wYear, .wMonth, .wDay) + _
                   TimeSerial(.wHour, .wMinute, .wSecond)
    End With
End Function

Public Function SystemTimeToDate(stSrc As SYSTEMTIME) As Date
    With stSrc
        SystemTimeToDate = DateSerial(.wYear, .wMonth, .wDay) + _
                   TimeSerial(.wHour, .wMinute, .wSecond)
    End With
End Function


