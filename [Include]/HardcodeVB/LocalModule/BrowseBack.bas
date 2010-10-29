Attribute VB_Name = "MBrowseBack"
Option Explicit

#If fComponent Then
' This function in separate standard module for global module version
' only--standard module version in same file with BrowseForFolders
Function BrowseCallbackProc(ByVal hWnd As Long, _
                            ByVal uMsg As Long, _
                            ByVal lParam As Long, _
                            ByVal lpData As Long) As Long
    Dim s As String, iRet As Long
    Select Case uMsg
    ' Browse dialog box has finished initializing (lParam is NULL)
    Case BFFM_INITIALIZED
        Debug.Print "BFFM_INITIALIZED"
        s = MUtility.PointerToString(lpData)
        MUtility.DenormalizePath s
        ' Set the selection
        iRet = SendMessageStr(hWnd, BFFM_SETSELECTION, ByVal APITRUE, s)
        
    ' Selection has changed (lParam contains pidl of selected folder)
    Case BFFM_SELCHANGED
        Debug.Print "BFFM_SELCHANGED"
        ' Display full path if status area if enabled
        s = MFoldTool.PathFromPidl(lParam)
        iRet = SendMessageStr(hWnd, BFFM_SETSTATUSTEXT, ByVal 0&, s)
        
    ' Invalid name in edit box (lParam parameter has invalid name string)
    Case BFFM_VALIDATEFAILED
        Debug.Print "BFFM_VALIDATEFAILED"
        ' Return zero to dismiss dialog or nonzero to keep it displayed
        ' Disable the OK button
        iRet = SendMessage(hWnd, BFFM_ENABLEOK, ByVal 0&, ByVal APIFALSE)
        s = MUtility.PointerToString(lParam)
        s = "Path invalid: " & s
        iRet = SendMessageStr(hWnd, BFFM_SETSTATUSTEXT, ByVal 0&, s)

    Case Else
        Debug.Print uMsg
    End Select
    BrowseCallbackProc = 0
End Function
#End If



