Attribute VB_Name = "MShort"
Option Explicit

Public Enum EErrorShort
    eeBaseShort = 13610     ' Short
End Enum


Sub CreateShortcut(LinkFile As Variant, _
                   Path As String, _
                   Optional DisplayMode As Long = edmNormal, _
                   Optional WorkingDirectory As String, _
                   Optional Arguments As String, _
                   Optional Description As String, _
                   Optional IconIndex As Long = 0)
    Dim shortcut As New CShortcut
    With shortcut
        .Path = Path
        .WorkingDirectory = WorkingDirectory
        .DisplayMode = DisplayMode
        .Arguments = Arguments
        .Description = Description
        .Icon = IconIndex
        .Save LinkFile
    End With
End Sub

Sub UpdateShortcut(LinkFile As String, _
                   Optional Path As String, _
                   Optional WorkingDirectory As String, _
                   Optional DisplayMode As Long, _
                   Optional Arguments As String, _
                   Optional Description As String, _
                   Optional Icon As Variant)
    Dim shortcut As New CShortcut
    With shortcut
        .Resolve LinkFile
        Path = .Path
        WorkingDirectory = .WorkingDirectory
        DisplayMode = .DisplayMode
        Arguments = .Arguments
        Description = .Description
        Icon = .Icon
    End With
End Sub

#If fComponent = 0 Then
Private Sub ErrRaise(e As Long)
    Dim sText As String, sSource As String
    If e > 1000 Then
        sSource = App.ExeName & ".Short"
        Select Case e
        Case eeBaseShort
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

