Attribute VB_Name = "MResProc"
Option Explicit

Function ResTypeProc(ByVal hModule As Long, ByVal lpszType As Long, _
                     frm As Form) As Long
    ResTypeProc = True      ' Always return True
    If lpszType <= 65535 Then
        ' Enumerate resources by ID
        Call EnumResourceNamesID(hModule, lpszType, _
                                 AddressOf ResNameProc, frm)
    Else
        ' Enumerate resources by string name
        Call EnumResourceNamesStr(hModule, PointerToString(lpszType), _
                                  AddressOf ResNameProc, frm)
    End If
End Function

Function ResNameProc(ByVal hModule As Long, ByVal lpszType As Long, _
                     ByVal lpszName As Long, frm As Form) As Long
    Dim sType As String, sName As String
    ResNameProc = True      ' Always return True
    If lpszName <= 65535 Then
        sName = Format$(lpszName, "00000")
    Else
        sName = PointerToString(lpszName)
    End If
    If lpszType <= 65535 Then
        sType = ResourceIdToStr(lpszType)
    Else
        sType = PointerToString(lpszType)
    End If
    If frm.chkFilter = vbChecked Then
        If Not ValidateResource(hModule, sName, sType) Then Exit Function
    End If
    frm.lstResource.AddItem sName & "   " & sType
End Function


Function ValidateResource(hMod As Long, ByVal sName As String, _
                          ByVal sType As String) As Boolean

    Dim i As Integer, hRes As Long

    ' Extract resource ID and type
    If Left$(sName, 1) = "0" Then sName = "#" & Left$(sName, 5)
    
    Select Case UCase$(sType)
    Case "CURSOR", "GROUP_CURSOR", "GROUP CURSOR"
        hRes = LoadImage(hMod, sName, IMAGE_CURSOR, 0, 0, 0)
        If hRes Then ValidateResource = True
        Call DeleteObject(hRes)
    Case "BITMAP"
        hRes = LoadBitmap(hMod, sName)
        If hRes Then ValidateResource = True
        Call DeleteObject(hRes)
    Case "ICON", "GROUP_ICON", "GROUP ICON"
        hRes = LoadImage(hMod, sName, IMAGE_ICON, 0, 0, 0)
        If hRes Then ValidateResource = True
        Call DeleteObject(hRes)
    Case "STRING", "STRINGTABLE"
        hRes = FindResourceStrId(hMod, sName, RT_STRING)
        If hRes Then ValidateResource = True
        Call FreeResource(hRes)
    Case "WAVE", "FONTDIR", "FONT", "DIALOG", "ACCELERATOR", _
         "VERSION", "MENU", "AVI"
        ' Always accept these
        ValidateResource = True
    Case Else
        hRes = FindResourceStrStr(hMod, sName, sType)
        If hRes Then ValidateResource = True
        Call FreeResource(hRes)
    End Select
    
End Function



