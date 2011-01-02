VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GErrors"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

#If fComponent Then
Sub ErrRaise(e As Long)
    MErrors.ErrRaise e
End Sub
#End If

Sub ApiRaiseIf(ByVal e As Long)
    MErrors.ApiRaiseIf e
End Sub

Sub ApiRaise(ByVal e As Long)
    MErrors.ApiRaise e
End Sub

Function ComToApi(ByVal e As Long) As Long
    ComToApi = MErrors.ComToApi(e)
End Function

Function ApiToCom(ByVal e As Long) As Long
    ApiToCom = MErrors.ApiToCom(e)
End Function

Function ComToApiStr(ByVal e As Long) As String
    ComToApiStr = MErrors.ComToApiStr(e)
End Function

Function ApiError(ByVal e As Long) As String
    ApiError = MErrors.ApiError(e)
End Function

Function LastApiError() As String
    LastApiError = MErrors.LastApiError
End Function

Function BasicError(ByVal e As Long) As Long
    BasicError = MErrors.BasicError(e)
End Function

Function COMError(e As Long) As Long
    COMError = MErrors.COMError(e)
End Function
'


