Attribute VB_Name = "MAppPath"
Option Explicit

Sub Main()
    If Not HasShell Then
        MsgBox "This environment does not support App Paths"
        Exit Sub
    End If
    Dim sExeSpec As String, sExe As String, f As Boolean
    Dim s As String, ret As Integer
    Dim fOverRide As Boolean, fSetPath As Boolean, fRemove As Boolean
    If Command$ = sEmpty Then
        ' Query for EXE files
        f = VBGetOpenFileName(FileName:=sExeSpec, _
                              Flags:=OFN_FILEMUSTEXIST Or OFN_HIDEREADONLY, _
                              InitDir:="c:\", _
                              filter:="EXE Files|*.exe")
        sExe = GetFileBaseExt(sExeSpec)
        s = GetAppPath(sExe)
        If s <> sEmpty Then
            ret = MsgBox("App Path set to: " & s & ".  Remove?", vbYesNo)
            If ret = vbYes Then fRemove = True
        End If
        f = True
        fSetPath = True
        If f = False Or sExeSpec = sEmpty Then Exit Sub
    Else
        ' Parse and handle command line
        Dim sToken As String
        Const sSep = sSpace & sTab
        sToken = GetQToken(Command$, sSep)
        Do While sToken <> sEmpty
            Select Case UCase$(Left$(sToken, 2))
            Case "/O"
                fOverRide = True
            Case "/R"
                fRemove = True
            Case "/P"
                fSetPath = True
            Case Else
                If Left$(sToken, 1) = "/" Or sExeSpec <> sEmpty Then
                    MsgBox "Invalid command line" & sCrLfCrLf & _
                           "Syntax: AppPath [/O[verride]] [/R[emove]] [/P[athSet]] <fullpathspec>"
		    Exit Sub
                Else
                    sExeSpec = GetFileFullSpec(sToken)
                End If
            End Select
            sToken = GetQToken(sEmpty, sSep)
        Loop
        sExe = GetFileBaseExt(sExeSpec)
    End If
    If fRemove Then
        If Not RemoveAppPath(sExe) Then
            MsgBox "Couldn't delete App Path for: " & sExe
        End If
        Exit Sub
    End If
    ' See if path already exists
    s = GetAppPath(sExe)
    If s <> sEmpty And fOverRide = False Then
        ret = MsgBox("App Path set to: " & s & ".  Override?", vbYesNo)
        If ret = vbNo Then Exit Sub
    End If
    SetAppPath sExeSpec, fSetPath
End Sub





