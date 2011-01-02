Attribute VB_Name = "MErrors"
Option Explicit

#If fComponent Then
Sub ErrRaise(e As Long)
    Dim sText As String, sSource As String
    If e > 1000 Then
        sSource = App.ExeName & "." & LoadResString((e \ 10) * 10)
        sText = LoadResString(e)
        Err.Raise e, sSource, sText
    Else
        ' Raise standard Visual Basic error
        sSource = App.ExeName & ".VBError"
        Err.Raise e, sSource
    End If
    ' Challenge: Enhance to use help files
End Sub
#End If

Sub ApiRaiseIf(ByVal e As Long)
    If e Then MErrors.ApiRaise e
End Sub

Sub ApiRaise(ByVal e As Long)
    Err.Raise vbObjectError + 29000 + e, _
              App.ExeName & ".Windows", ApiError(e)
End Sub

Function ComToApi(ByVal e As Long) As Long
    ComToApi = (e And &HFFFF&) - 29000
End Function

Function ApiToCom(ByVal e As Long) As Long
    ApiToCom = (e Or &H80040000) + 29000
End Function

Function ComToApiStr(ByVal e As Long) As String
    ComToApiStr = ApiError((e And &HFFFF&) - 29000)
End Function

Function ApiError(ByVal e As Long) As String
    Dim s As String, c As Long
    s = String(256, 0)
    c = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM Or _
                      FORMAT_MESSAGE_IGNORE_INSERTS, _
                      pNull, e, 0&, s, Len(s), ByVal pNull)
    If c Then ApiError = Left$(s, c)
End Function

Function LastApiError() As String
    LastApiError = ApiError(Err.LastDllError)
End Function

Function BasicError(ByVal e As Long) As Long
    BasicError = e And &HFFFF&
End Function

Function COMError(e As Long) As Long
    COMError = e Or vbObjectError
End Function
'
