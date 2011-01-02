VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GWinIter"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum EErrorWinIter
    eeBaseWinIter = 13630   ' WinIter
End Enum

Function IterateChildWindows(ByVal iLevel As Integer, _
                             ByVal hWnd As Long, _
                             helper As IWindowsHelper) As Long
    BugAssert hWnd <> hNull

    ' Handle current window, allowing user to fail
    
    IterateChildWindows = helper.DoWindow(iLevel, hWnd)
    If IterateChildWindows <> hNull Then Exit Function
    ' Get its child (if any)
    hWnd = GetWindow(hWnd, GW_CHILD)
    ' Iterate through each child window
    Do While hWnd <> hNull
        IterateChildWindows = _
            IterateChildWindows(iLevel + 1, hWnd, helper)
        If IterateChildWindows <> hNull Then Exit Function
        ' Get next child
        hWnd = GetWindow(hWnd, GW_HWNDNEXT)
        ' Give other processes some cycles
        DoEvents
    Loop
    ' Nothing found
    IterateChildWindows = hNull

End Function

#If fComponent = 0 Then
Private Sub ErrRaise(e As Long)
    Dim sText As String, sSource As String
    If e > 1000 Then
        sSource = App.ExeName & ".WinIter"
        Select Case e
        Case eeBaseWinIter
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

