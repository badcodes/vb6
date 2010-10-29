Attribute VB_Name = "MEnumProc"
Option Explicit

Public lstEnumRef As ListBox

Function EnumWndProc(ByVal hWnd As Long, lParam As Long) As Long
    ' Increment count
    lParam = lParam + 1
    ' Get window title and insert into ListBox
    Dim s As String
    s = Left$(WindowTextFromWnd(hWnd), 20)
    If s <> sEmpty Then
        lstEnumRef.AddItem s
        'lstEnumRef.ItemData(lstEnumRef.NewIndex) = hWnd
    End If
    ' Return True to keep enumerating
    EnumWndProc = True
End Function

Function EnumFontFamProc(elf As ENUMLOGFONT, tm As NEWTEXTMETRIC, dwType As Long, lpData As Long) As Long
    lpData = lpData + 1
    Dim s As String
    s = StrZToStr(BytesToStr(elf.elfLogFont.lfFaceName))
    If s <> sEmpty Then lstEnumRef.AddItem s
    EnumFontFamProc = True
End Function

Function EnumResTypeProc(ByVal hModule As Long, ByVal lpszType As Long, _
                         lParam As Long) As Long
    If lpszType < 65535 Then
        ' Enumerate resources by ID
        Call EnumResourceNamesID(hNull, lpszType, _
                                 AddressOf EnumResNameProc, lParam)
    Else
        ' Enumerate resources by string name
        Call EnumResourceNamesStr(hNull, PointerToString(lpszType), _
                                  AddressOf EnumResNameProc, lParam)
    End If
    EnumResTypeProc = True
End Function

Function EnumResNameProc(ByVal hModule As Long, ByVal lpszType As Long, ByVal lpszName As Long, lParam As Long) As Long
    Dim sType As String, sName As String, sHandle As String
    If lpszName < 65535 Then
        sName = lpszName
    Else
        sName = PointerToString(lpszName)
    End If
    If lpszType < 65535 Then
        sType = ResourceIdToStr(lpszType)
    Else
        sType = PointerToString(lpszType)
    End If
    If sType <> sEmpty Then
        lstEnumRef.AddItem sName & Chr$(9) & sType
    End If
    lParam = lParam + 1
    EnumResNameProc = True
End Function



