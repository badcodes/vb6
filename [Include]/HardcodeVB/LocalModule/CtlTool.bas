Attribute VB_Name = "MCtlTool"
Option Explicit

Function UniqueControlName(sPrefix As String, Ext As Object) As String
    Dim v As Variant, s As String, c As Long, fFound As Boolean
    On Error GoTo UniqueControlNameFail
    s = sPrefix
    Do
        fFound = False
        ' Search for a control with the proposed prefix name
        For Each v In Ext.Container.Controls
            If v.Name = s Then
                ' Nope, try another name
                fFound = True
                c = c + 1
                s = sPrefix & c
                Exit For
            End If
        Next
    Loop Until fFound = False
    ' Use this name
    UniqueControlName = s
    Exit Function
    
UniqueControlNameFail:
    ' Failure probably means no Extender.Container.Controls
    UniqueControlName = sPrefix
End Function

' This function returns 0/1 rather than True/False for
' easier use with check boxes.
Private Function CheckBit(iValue As Integer, iBitPos As Integer) As Integer
    If iValue And (2 ^ iBitPos) Then
        CheckBit = 1
    Else
        CheckBit = 0
    End If
End Function

' This is a renamed version of MWinTool.ChangeStyleBit. It is
' duplicated here to avoid bringing MWinTool and many of its
' dependencies into control projects. This compromises my code
' reuse principals, but complete purity would have compromised
' my efficiency goals.
Sub ModifyStyleBit(hWnd As Long, f As Boolean, afNew As Long)
    Dim af As Long, hParent As Long
    af = GetWindowLong(hWnd, GWL_STYLE)
    If f Then
        af = af Or afNew
    Else
        af = af And (Not afNew)
    End If
    Call SetWindowLong(hWnd, GWL_STYLE, af)
    ' Reset the parent so that change will "take"
    hParent = GetParent(hWnd)
    SetParent hWnd, hParent
    ' Redraw for added insurance
    Call SetWindowPos(hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, _
                      SWP_NOZORDER Or SWP_NOSIZE Or _
                      SWP_NOMOVE Or SWP_DRAWFRAME)
End Sub




