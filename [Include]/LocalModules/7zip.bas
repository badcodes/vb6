Attribute VB_Name = "M7zip"
Option Explicit
Private Const dq As String = """"
Private Function QuoteString(ByRef strNake As String) As String
    QuoteString = dq & strNake & dq
End Function
Public Function ShellExtract(ByVal v7zip As String, ByVal vArg As String, ByVal vSrc As String, ByVal pDstName As String) As String
        '<EhHeader>
        On Error GoTo RunOn_Err
        '</EhHeader>

        Dim pCmdLine As String
        pCmdLine = QuoteString(v7zip) & " " & vArg & " x " & QuoteString(vSrc) & " -o" & QuoteString(pDstName)
        Debug.Print pCmdLine
        'MsgBox pCmdLine
        'Exit Function
        Dim pExit As Long
        pExit = MShell32.ShellAndClose(pCmdLine, vbMinimizedNoFocus)
        Select Case pExit
            Case 0
                ShellExtract = " [OK]"
            Case 1
                ShellExtract = " [Warning]"
            Case 2
                ShellExtract = " [Fatal error]"
            Case 7
                ShellExtract = " [Command line error]"
            Case 8
                ShellExtract = " [Not enough memory for operation"
            Case 255
                ShellExtract = " [User stopped the process]"
            Case Else
                ShellExtract = " [Failed]"
        End Select
        Exit Function

RunOn_Err:
        MsgBox Err.Description & vbCrLf & "in M7zip.ShellExtract ", vbExclamation And vbOKOnly, "Application Error"
        On Error Resume Next
'    ChDrive pCurDrive
'     ChDir pCurDir
     ShellExtract = " [Error]"
        '</EhFooter>
End Function

