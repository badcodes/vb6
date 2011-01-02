VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GgCIni"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' Ini File Functions Class
' Copyright (C) 1996, Jens Balchen
'
' Uses
'
' Exposes
'     Function GetSetting
'     Function SaveSetting
'     Function GetSection
'
' Comments

Option Explicit

' --------
'  Public
' --------
'
' Property for file to read

' ---------
'  Private
' ---------
'
' API to read/write ini's

#If Win32 Then
   Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
   Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal Appname As String, ByVal KeyName As Any, ByVal NewString As Any, ByVal FileName As String) As Integer
#Else
   Private Declare Function GetPrivateProfileString Lib "Kernel" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
   Private Declare Function WritePrivateProfileString Lib "Kernel" (ByVal Appname As String, ByVal KeyName As Any, ByVal NewString As Any, ByVal FileName As String) As Integer
#End If
Private gString As gCString

Sub iniDeleteSection(ByVal File As String, ByVal Section As String)

    Dim retval As Integer
    retval = WritePrivateProfileString(Section, 0&, "", File)

End Sub

Public Function iniSaveSetting(ByVal File As String, ByVal Section$, ByVal Key$, ByVal value$) As Integer

    'Dim retval As Integer
    iniSaveSetting = WritePrivateProfileString(Section$, Key$, value$, File)

End Function

Public Function iniGetSetting(ByVal File As String, ByVal Section As String, ByVal KeyName As String) As String

    Dim retval As Integer
    Dim t As String * 255
    ' Get the value
    retval = GetPrivateProfileString(Section, KeyName, "", t, Len(t), File)
    ' If there is one, return it

    If retval > 0 Then
        iniGetSetting = gString.rdel(Left$(t, retval))
    Else
        iniGetSetting = ""
    End If

End Function

Public Function iniGetSection(ByVal File As String, ByVal Section As String, KeyArray() As String) As Integer

    Dim retval As Integer
    ' Allocate space for return value
    Dim t As String * 2500
    Dim lastpointer As Integer
    Dim nullpointer As Integer
    Dim ArrayCount As Integer
    Dim keystring As String
    ReDim KeyArray(0)
    ' Get the value
    retval = GetPrivateProfileString(Section, 0&, "", t, Len(t), File)
    ' If there is one, return it

    If retval > 0 Then
        '
        ' Separate the keys and store them in the array
        nullpointer = InStr(t, Chr$(0))
        lastpointer = 1

        Do While (nullpointer <> 0 And nullpointer > lastpointer + 1)
            '
            ' Extract key string
            keystring = Mid$(t, lastpointer, nullpointer - lastpointer)
            '
            ' Now add to array
            ArrayCount = ArrayCount + 1
            ReDim Preserve KeyArray(ArrayCount)
            KeyArray(ArrayCount) = keystring
            '
            ' Find next null
            lastpointer = nullpointer + 1
            nullpointer = InStr(nullpointer + 1, t, Chr$(0))
        Loop

    End If

    '
    ' Return the number of array elements
    iniGetSection = ArrayCount

End Function

Private Sub Class_Initialize()

    Set gString = New gCString

End Sub

Private Sub Class_Terminate()

    Set gString = Nothing

End Sub


