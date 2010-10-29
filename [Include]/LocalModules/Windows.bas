Attribute VB_Name = "MWindows"
Option Explicit
Public Const SWP_FRAMECHANGED = &H20
Public Const SWP_HIDEWINDOW = &H80
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_NOCOPYBITS = &H100
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOOWNERZORDER = &H200
Public Const SWP_NOREDRAW = &H8
Public Const SWP_NOREPOSITION = SWP_NOOWNERZORDER
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOZORDER = &H4
Public Const SWP_SHOWWINDOW = &H40
Public Const SWP_DRAWFRAME = SWP_FRAMECHANGED
Public Enum WindowBorderStyle
    bsNone = 100663296
    bsFixedSingle = 113770496
    bsResizable = 114229248
    bsFixedDialog = 113770624
End Enum
Public Enum WindowPlacementOrder
    HWND_BOTTOM = 1
    'HWND_BROADCAST = &HFFFF&
    ' HWND_DESKTOP = 0
    HWND_NOTOPMOST = -2
    HWND_TOP = 0
    HWND_TOPMOST = -1
End Enum
Private Const GWL_STYLE = -16
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Private Type tagInitCommonControlsEx
   lngSize As Long
   lngICC As Long
End Type
Private Declare Function InitCommonControlsEx Lib "comctl32.dll" _
   (iccex As tagInitCommonControlsEx) As Boolean
Private Const ICC_USEREX_CLASSES = &H200

Public Function InitCommonControlsVB() As Boolean
   On Error Resume Next
   Dim iccex As tagInitCommonControlsEx
   ' Ensure CC available:
   With iccex
       .lngSize = LenB(iccex)
       .lngICC = ICC_USEREX_CLASSES
   End With
   InitCommonControlsEx iccex
   InitCommonControlsVB = (Err.Number = 0)
   On Error GoTo 0
End Function

Public Function Windows_LockWindow(ByVal vHwnd As Long) As Boolean
    Windows_LockWindow = (LockWindowUpdate(vHwnd) = 0)
End Function

Public Sub setPosition(objForm As Form, wpo As WindowPlacementOrder)

'FIXIT: 'hwnd' is not a property of the generic 'Form' object in Visual Basic .NET. To access 'hwnd' declare 'objForm' using its actual type instead of 'Form'     FixIT90210ae-R1460-RCFE85
    SetWindowPos objForm.hwnd, wpo, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE

End Sub

' 依據指定形式，在執行階段重設表單的 BorderStyle
Public Sub setBorderStyle(objForm As Form, ByVal wndStyle As WindowBorderStyle)

    With objForm
'FIXIT: 'hwnd' is not a property of the generic 'Form' object in Visual Basic .NET. To access 'hwnd' declare 'objForm' using its actual type instead of 'Form'     FixIT90210ae-R1460-RCFE85
        SetWindowLong .hwnd, GWL_STYLE, wndStyle
        .Hide
        ' 必須使表單 Resize 與重繪才能夠實現新邊框樣式
        .Width = .Width - Screen.TwipsPerPixelX
        .Width = .Width + Screen.TwipsPerPixelX
        .Show
    End With

End Sub


'' 滑鼠雙擊表單時，即可切換全螢幕顯示
'Public Sub switch_FullScreen(objForm As Form)
'  Static isBsNone As Boolean
'  Static preLeft As Long, preTop As Long
'  Static preWidth As Long, preHeight As Long
'
'  isBsNone = Not isBsNone
'  SetBorderStyle objForm, bsResizable
'  SetBorderStyle objForm, IIf(isBsNone, bsNone, bsResizable)
'  If isBsNone Then
'    ' 在實現全螢幕之前，先紀錄表單的位置大小
'    preLeft = objForm.Left: preTop = objForm.Top
'    preWidth = objForm.Width: preHeight = objForm.Height
'    objForm.Move 0, 0, Screen.Width, Screen.Height
'  Else
'    ' 恢復全螢幕之前的表單位置大小
'    objForm.Move preLeft, preTop, preWidth, preHeight
'  End If
'End Sub
'Public Function switch_FullScreen(objForm As Form) As Boolean
'
'    Static formPos As MYPoS
'    Static bFullScreen As Boolean
'
'    On Error Resume Next
'
'    If bFullScreen = True Then
'
'        With formPos
'            objForm.Move .Left, .Top, .Width, .Height
'        End With
'
'        setBorderStyle objForm, bsResizable
'        bFullScreen = False
'    Else
'
'        With formPos
'            .Left = objForm.Left
'            .Top = objForm.Top
'            .Height = objForm.Height
'            .Width = objForm.Width
'        End With
'
'        objForm.Move 0, 0, Screen.Width, Screen.Height
'        setBorderStyle objForm, bsNone
'        bFullScreen = True
'    End If
'
'    switch_FullScreen = bFullScreen
'
'End Function
'Public Function switch_BorderStyle(objForm As Form) As Boolean
'
'    Static bChanged As Boolean
'
'    On Error Resume Next
'
'    If bChanged Then
'        setBorderStyle objForm, bsResizable
'        bChanged = False
'    Else
'
'        setBorderStyle objForm, bsNone
'        bChanged = True
'    End If
'
'    switch_BorderStyle = bChanged
'
'End Function

