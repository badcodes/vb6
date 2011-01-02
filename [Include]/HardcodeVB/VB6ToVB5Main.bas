Attribute VB_Name = "MBugWizMain"
Sub Main()
    Dim iExit As Long ' = 0
    If Command$ = sEmpty Then
        Dim frm As FVB6ToVB5
        Set frm = New FVB6ToVB5
        frm.Show
    Else
        Dim sSwitch As String, sPath As String
        Dim vbver As New CVBVerFilter
        On Error GoTo Failure
        sSwitch = GetToken(Command$, " " & sTab)
        sFile = GetToken(sEmpty, " ")
        Select Case UCase$(sSwitch)
        Case "/F", "-F"
            If sFile = sEmpty Then
                iExit = 1
                SyntaxBox
            Else
                vbver.ConvertFile sPath
            End If
        Case "/D", "-D"
            If sPath = sEmpty Then sPath = CurDir$
            WalkFiles vbver, ewmfFiles, sPath
        Case "/A", "-A"
            If sPath = sEmpty Then sPath = CurDir$
            WalkAllFiles vbver, ewmfFiles, sPath
        Case Else
            SyntaxBox
            iExit = 1
        End Select
        Beep
        ' ExitProcess must be at the end after all cleanup is finished
        ' If you call ExitProcess in the IDE, say goodbye
        If IsExe Then ExitProcess iExit
    End If
    Exit Sub
Failure:
    If IsExe Then ExitProcess 1
End Sub

Sub SyntaxBox()
    Dim s As String
    s = "Syntax: VB6ToVB5 <[/F file | /D [dir] | /A [dir]]" & sCrLf & _
        "    /F  -  Convert specified file" & sCrLf & _
        "    /D  -  Convert files in specified or current directory" & sCrLf & _
        "    /A  -  Convert files recursively starting at specified or current directory"
    MsgBox s
End Sub
