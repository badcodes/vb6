Attribute VB_Name = "MBugWizMain"
Sub Main()
    Dim sWhat As String, sHow As String, sWhere As String
    sWhat = GetQToken(Command$, " " & sTab)
    sHow = GetQToken(sEmpty, " " & sTab)
    sWhere = GetQToken(sEmpty, " " & sTab)
    If sWhat = sEmpty Then
        Dim frm As FBugWizard
        Set frm = New FBugWizard
        frm.sHow
    Else
        Dim bug As CBugFilter
        Set bug = New CBugFilter
        Select Case UCase$(sWhat)
        Case "/B"
            'B  - Enable Bug statements
            bug.FilterType = eftEnableBug
        Case "/RB"
            'RB - Remove Bug statements
            bug.FilterType = eftDisableBug
        Case "/P"
            'P  - Enable Profile statements
            bug.FilterType = eftEnableProfile
        Case "/RP"
            'RP - Remove Profile statements
            bug.FilterType = eftDisableProfile
        Case "/X"
            'X  - EXpand BugAsserts
            bug.FilterType = eftExpandAsserts
        Case "/RX"
            'RX - Trim BugAsserts
            bug.FilterType = eftTrimAsserts
        Case Else
            SyntaxMessage
            Exit Sub
        End Select
        
        Select Case UCase$(sHow)
        Case "/F"
            ' Process one file
            If sWhere = sEmpty Then
                SyntaxMessage
                Exit Sub
            End If
            IFilter(bug).Source = sWhere
            FilterTextFile bug
        Case "/D"
            ' Process one directory
            If sWhere = sEmpty Then sWhere = CurDir$
            WalkFiles bug, ewmfFiles, sWhere
        Case "/A", "-A"
            ' Process all directories
            If sPath = sEmpty Then sPath = CurDir$
            WalkAllFiles bug, ewmfFiles, sWhere
        Case Else
            SyntaxMessage
            Exit Sub
        End Select
        Beep
    End If
End Sub

Sub SyntaxMessage()
    MsgBox "Syntax: BugWiz <actionopt> <scopeopt> [<path>]" & sCrLf & _
           "    Action options:" & sCrLf & _
           "       /B  - Enable Bug statements" & sCrLf & _
           "       /RB - Remove Bug statements" & sCrLf & _
           "       /P  - Enable Profile statements" & sCrLf & _
           "       /RP - Remove Profile statements" & sCrLf & _
           "       /X  - EXpand BugAsserts" & sCrLf & _
           "       /RP - Trim BugAsserts" & sCrLf & _
           "    Scope options:" & sCrLf & _
           "       /F  - One file given by <path>" & sCrLf & _
           "       /D  - Files in <path> or current directory" & sCrLf & _
           "       /A  - Files recursively starting in <path> or current directory"
End Sub

