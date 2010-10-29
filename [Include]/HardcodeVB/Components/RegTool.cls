VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GRegTool"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum EErrorRegTool
    eeBaseRegTool = 13590   ' RegTool
End Enum

Const sWin = "Software\Microsoft\Windows\"
Const sExp = "CurrentVersion\Explorer\Shell Folders"
Const sWinExp = sWin & sExp
Const sBack = "\"
Const sAppPath = "SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths"


Function GetRegValue(ByVal hKey As Long, sName As String, _
                     vValue As Variant) As Long
    Dim cData As Long, sData As String, ordType As Long, e As Long
    e = RegQueryValueEx(hKey, sName, pNull, ordType, 0&, cData)
    If e And e <> ERROR_MORE_DATA Then Exit Function
    Select Case ordType
    Case REG_DWORD, REG_DWORD_LITTLE_ENDIAN
        Dim iData As Long
        e = RegQueryValueExInt(hKey, sName, pNull, _
                               ordType, iData, cData)
        vValue = iData
        
    Case REG_DWORD_BIG_ENDIAN  ' Unlikely, but you never know
        Dim dwData As Long
        e = RegQueryValueExInt(hKey, sName, pNull, _
                               ordType, dwData, cData)
        vValue = MBytes.SwapEndian(dwData)
        
    Case REG_SZ, REG_MULTI_SZ ' Same thing to Visual Basic
        sData = String$(cData - 1, 0)
        e = RegQueryValueExStr(hKey, sName, pNull, _
                               ordType, sData, cData)
        vValue = sData
        
    Case REG_EXPAND_SZ
        sData = String$(cData - 1, 0)
        e = RegQueryValueExStr(hKey, sName, pNull, _
                               ordType, sData, cData)
        vValue = MUtility.ExpandEnvStr(sData)
        
    ' Catch REG_BINARY and anything else
    Case Else
        Dim abData() As Byte
        ReDim abData(cData)
        e = RegQueryValueExByte(hKey, sName, pNull, _
                                ordType, abData(0), cData)
        vValue = abData
        
    End Select
    GetRegValue = e
End Function

Function CreateRegValue(vValueA As Variant, ByVal hKeyA As Long, _
                        Optional sNameA As String) As Long
    Dim c As Long, e As Long, ordType As Long
    Select Case VarType(vValueA)
    Case vbArray + vbByte
        Dim ab() As Byte
        ab = vValueA
        ordType = REG_BINARY
        c = UBound(ab) - LBound(ab) - 1
        e = RegSetValueExByte(hKeyA, sNameA, pNull, ordType, ab(0), c)
        
    Case vbLong, vbInteger
        Dim i As Long
        i = vValueA
        ordType = REG_DWORD
        e = RegSetValueExInt(hKeyA, sNameA, pNull, ordType, i, 4)
        
    Case vbString
        Dim s As String, iPos As Long
        s = vValueA
        ordType = REG_SZ
        ' Assume anything with two non-adjacent percents is expanded string
        iPos = InStr(s, "%")
        If iPos Then
            If InStr(iPos + 2, s, "%") Then ordType = REG_EXPAND_SZ
        End If
        c = Len(s) + 1
        e = RegSetValueExStr(hKeyA, sNameA, pNull, ordType, s, c)
        
    ' User should convert to a compatible type before calling
    Case Else
        e = ERROR_INVALID_DATA
        
    End Select
    CreateRegValue = e
End Function

Function GetRegValueNext(ByVal hKey As Long, _
                         i As Long, _
                         sName As String, _
                         vValue As Variant) As Long
    Dim cName As Long, cData As Long, sData As String
    Dim ordType As Long, cJunk As Long, ft As FILETIME
    ' Get the value name and type in the first call
    vValue = Empty
    ' Name buffer of 256 should be big enough for most
    Do
        cName = cName + 256
        sName = String$(cName, 0)
        GetRegValueNext = RegEnumValue(hKey, i, sName, cName, _
                                       pNull, ordType, pNull, cData)
        ' Repeat with bigger buffer in unlikely event name is too long
    Loop While GetRegValueNext = ERROR_MORE_DATA
    ' Fail for other errors
    If GetRegValueNext Then Exit Function
    sName = Left$(sName, cName)

    ' Handle each type separately
    Select Case ordType
    Case REG_DWORD, REG_DWORD_LITTLE_ENDIAN
        Dim iData As Long
        GetRegValueNext = _
            RegEnumValueInt(hKey, i, sName, cName + 1, _
                            pNull, ordType, iData, cData)
        vValue = iData
        
    Case REG_DWORD_BIG_ENDIAN  ' Unlikely, but you never know
        Dim dwData As Long
        GetRegValueNext = _
            RegEnumValueInt(hKey, i, sName, cName + 1, _
                            pNull, ordType, dwData, cData)
        vValue = MBytes.SwapEndian(dwData)
        
    Case REG_SZ, REG_MULTI_SZ ' Same thing to Visual Basic
        sData = String$(cData - 1, 0)
        GetRegValueNext = _
            RegEnumValueStr(hKey, i, sName, cName + 1, _
                            pNull, ordType, sData, cData)
        vValue = sData
        
    Case REG_EXPAND_SZ         ' Expand environment variables
        sData = String$(cData - 1, 0)
        GetRegValueNext = _
            RegEnumValueStr(hKey, i, sName, cName + 1, _
                            pNull, ordType, sData, cData)
        vValue = MUtility.ExpandEnvStr(sData)
    
    Case Else       ' Catch REG_BINARY and anything else
        Dim abData() As Byte
        ReDim abData(cData)
        GetRegValueNext = _
            RegEnumValueByte(hKey, i, sName, cName + 1, _
                             pNull, ordType, _
                             abData(0), cData)
        vValue = abData
        
    End Select
    
End Function

Function GetRegNodeNext(ByVal hKey As Long, i As Long, sName As String) As Long
    Dim cName As Long, cJunk As Long, ft As FILETIME
    Do
        cName = cName + 256
        sName = String$(cName, 0)
        GetRegNodeNext = RegEnumKeyEx(hKey, i, sName, cName, _
                                      pNull, sNullStr, cJunk, ft)
        ' Repeat with bigger buffer in unlikely event name is too long
    Loop While GetRegNodeNext = ERROR_MORE_DATA
    ' Fail for other errors
    If GetRegNodeNext Then Exit Function
    sName = Left$(sName, cName)
   
End Function

Function CreateRegNode(ByVal hKey As Long, sKeyNew As String, _
                       hKeyNew As Long, Optional fExisted As Boolean, _
                       Optional ByVal afAccess As Long = KEY_ALL_ACCESS _
                       ) As Long
    Dim ordResult As Long
    CreateRegNode = RegCreateKeyEx(hKey, sKeyNew, 0&, sEmpty, _
                                   REG_OPTION_NON_VOLATILE, _
                                   afAccess, ByVal pNull, _
                                   hKeyNew, ordResult)
    fExisted = (ordResult = REG_OPENED_EXISTING_KEY)
End Function

' Delete node, but only if it has no subnodes (emulate WinNT RegDeleteKey)
Function DeleteOneRegNode(ByVal hKeyRoot As Long, sKey As String) As Long
    If MUtility.IsNT Then
        DeleteOneRegNode = RegDeleteKey(hKeyRoot, sKey)
    Else
        ' Check to see if there are subnodes
        Dim cJunk As Long, e As Long, cNode As Long, ft As FILETIME
        e = RegQueryInfoKey(hKeyRoot, sNullStr, cJunk, _
                            pNull, cNode, cJunk, cJunk, _
                            cJunk, cJunk, cJunk, cJunk, ft)
        ' Delete only if no nodes
        If cNode = 0 Then
            DeleteOneRegNode = RegDeleteKey(hKeyRoot, sKey)
        Else
            DeleteOneRegNode = ERROR_ACCESS_DENIED
        End If
    End If
End Function

' Delete node and all its subnodes (emulate Win95 RegDeleteKey)
Function DeleteRegNodes(ByVal hKeyRoot As Long, sKey As String) As Long
    Dim sKeyT As String, hSubKey As Long, ft As FILETIME

    ' Try to delete whole thing--always works for Win95, but fails on
    ' nodes with subnodes in WinNT
    DeleteRegNodes = RegDeleteKey(hKeyRoot, sKey)
    If DeleteRegNodes = ERROR_SUCCESS Then Exit Function
    DeleteRegNodes = RegOpenKeyEx(hKeyRoot, sKey, 0, _
                                  KEY_ALL_ACCESS, hSubKey)
    ' Delete each subnode
    Do While DeleteRegNodes = ERROR_SUCCESS
        sKeyT = String$(cMaxPath, 0)
        DeleteRegNodes = RegEnumKeyEx(hSubKey, 0, sKeyT, cMaxPath, _
                                      pNull, sNullStr, 0, ft)
        sKeyT = MUtility.StrZToStr(sKeyT)
        ' Recursive call to remove node and any subnodes
        If DeleteRegNodes = ERROR_SUCCESS Then
            DeleteRegNodes = DeleteRegNodes(hSubKey, sKeyT)
        End If
    Loop
    Call RegCloseKey(hSubKey)
    ' Try to delete root again
    DeleteRegNodes = RegDeleteKey(hKeyRoot, sKey)
    
End Function

Function GetRegStr(sKey As String, sItem As String, _
                   Optional ByVal hRoot As EROOTKEY _
                       = HKEY_CURRENT_USER) As String
    Dim e As Long, hKey As Long, s As String
    ' Open a subkey
    e = RegOpenKeyEx(hRoot, sKey, 0, KEY_QUERY_VALUE, hKey)
    ApiRaiseIf e
    Dim ert As EREGTYPE, c As Long
    ' Get the length and make sure it's a string
    e = RegQueryValueEx(hKey, sItem, 0&, ert, 0&, c)
    ' This error means no such entry, so return empty
    If e = ERROR_FILE_NOT_FOUND Then Exit Function
    ApiRaiseIf e
    BugAssert ert = REG_SZ
    If c <> 0 Then
        s = String$(c - 1, 0)
        ' Read the string
        e = RegQueryValueExStr(hKey, sItem, 0&, ert, s, c)
        ApiRaiseIf e
    End If
    RegCloseKey hKey
    GetRegStr = s
End Function

Function GetRegInt(sKey As String, sItem As String, _
                   Optional ByVal hRoot As EROOTKEY = HKEY_CURRENT_USER _
                   ) As Long
    Dim e As Long, hKey As Long
    ' Open a subkey
    e = RegOpenKeyEx(hRoot, sKey, 0, KEY_QUERY_VALUE, hKey)
    ApiRaiseIf e
    Dim ert As EREGTYPE, iVal As Long, c As Long
    ' Get the length and make sure it's an integer
    e = RegQueryValueEx(hKey, sItem, 0&, ert, 0&, c)
    ' This error means no such entry, so return empty
    If e = ERROR_FILE_NOT_FOUND Then Exit Function
    ApiRaiseIf e
    BugAssert ert = REG_DWORD
    If c <> 0 Then
        ' Read the integer
        e = RegQueryValueExInt(hKey, sItem, 0&, ert, iVal, c)
        ApiRaiseIf e
    End If
    RegCloseKey hKey
    GetRegInt = iVal
End Function

Sub SetRegStr(sItem As String, sValue As String, Optional sKey As String, _
              Optional ByVal hKey As EROOTKEY = HKEY_CURRENT_USER)
    If sKey <> sEmpty Then
        Dim hKeyNew As Long
        On Error Resume Next
        CreateRegNode hKey, sKey, hKeyNew
        If Err Then Exit Sub
        hKey = hKeyNew
    End If
    RegSetValueExStr hKeyNew, sItem, 0&, REG_SZ, sValue, Len(sValue) + 1
End Sub

Sub SetRegInt(sItem As String, iValue As Long, Optional sKey As String, _
              Optional ByVal hKey As EROOTKEY = HKEY_CURRENT_USER)
    ' If sKey not given, use given key handle, otherwise open key
    If sKey <> sEmpty Then
        Dim hKeyNew As Long
        On Error Resume Next
        CreateRegNode hKey, sKey, hKeyNew
        If Err Then Exit Sub
        hKey = hKeyNew
    End If
    RegSetValueExInt hKey, sItem, 0&, REG_DWORD, iValue, 4&
End Sub

' Get key locations in registry

Function GetDesktop() As String
    GetDesktop = GetRegStr(sWinExp, "Desktop") & sBack
End Function

Function GetFavorites() As String
    GetFavorites = GetRegStr(sWinExp, "Favorites") & sBack
End Function

Function GetStartMenu() As String
    GetStartMenu = GetRegStr(sWinExp, "Start Menu") & sBack
End Function

Function GetStartup() As String
    GetStartup = GetRegStr(sWinExp, "Startup") & sBack
End Function

Function GetPrograms() As String
    GetPrograms = GetRegStr(sWinExp, "Programs") & sBack
End Function

Function GetAppData() As String
    GetAppData = GetRegStr(sWinExp, "AppData") & sBack
End Function

Function GetCommonDesktop() As String
    GetCommonDesktop = GetRegStr(sWinExp, "Common Desktop") & sBack
End Function

Function GetCommonStartMenu() As String
    GetCommonStartMenu = GetRegStr(sWinExp, "Common Start Menu") & sBack
End Function

Function GetCommonStartup() As String
    GetCommonStartup = GetRegStr(sWinExp, "Common Startup") & sBack
End Function

Function GetCommonPrograms() As String
    GetCommonPrograms = GetRegStr(sWinExp, "Common Programs") & sBack
End Function

' AppPath procedures

Function GetAppPath(sExe As String) As String
    Dim s As String, f As Boolean
    On Error Resume Next
    GetAppPath = GetRegStr(sAppPath & "\" & sExe, sEmpty, HKEY_LOCAL_MACHINE)
End Function

Function SetAppPath(sExeSpec As String, _
                    fSetPath As Boolean) As Boolean
    Dim sExeName As String, hKeyNew As Long
    On Error GoTo FailSetAppPath
    sExeName = MUtility.GetFileBaseExt(sExeSpec)
    SetRegStr sEmpty, sExeSpec, sAppPath & "\" & sExeName, HKEY_LOCAL_MACHINE
    
    ' Success if name set regardless of path success
    SetAppPath = True
    On Error Resume Next
    
    ' Set path subkey if requested
    If fSetPath Then
        Dim sExePath As String
        sExePath = MUtility.GetFileDir(sExeSpec)
        sExePath = Left(sExePath, Len(sExePath) - 1) & ";"
        SetRegStr "Path", sExePath, sAppPath & "\" & sExeName, HKEY_LOCAL_MACHINE
    End If
    RegCloseKey hKeyNew

FailSetAppPath:
    
End Function

Function RemoveAppPath(sExe As String) As Boolean
    Dim e As Long
    e = DeleteOneRegNode(HKEY_LOCAL_MACHINE, sAppPath & "\" & sExe)
    RemoveAppPath = (e = 0)
End Function


#If fComponent = 0 Then
Private Sub ErrRaise(e As Long)
    Dim sText As String, sSource As String
    If e > 1000 Then
        sSource = App.ExeName & ".RegTool"
        Select Case e
        Case eeBaseRegTool
            BugAssert True
       ' Case ee...
       '     Add additional errors
        End Select
        Err.Raise COMError(e), sSource, sText
    Else
        ' Raise standard Visual Basic error
        sSource = App.ExeName & ".VBError"
        Err.Raise e, sSource
    End If
End Sub
#End If

