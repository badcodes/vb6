Attribute VB_Name = "MWinTool"
Option Explicit

Public Enum EErrorWinTool
    eeBaseWinTool = 13640   ' WinTool
End Enum

#If fComponent Then
Sub SetRedraw(ctl As Object, f As Boolean)
#Else
Sub SetRedraw(ctl As Control, f As Boolean)
#End If
    Call SendMessageVal(ctl.hWnd, WM_SETREDRAW, -CLng(f), 0&)
End Sub

#If fComponent Then
Function LookupItemData(ctl As Object, data As Long) As Integer
#Else
Function LookupItemData(ctl As Control, data As Long) As Integer
#End If
    Dim i As Integer
    LookupItemData = -1
    For i = 0 To ctl.ListCount - 1
        If data = ctl.ItemData(i) Then
            LookupItemData = i
            Exit Function
        End If
    Next
End Function

#If fComponent Then
Function LookupItem(ctl As Object, sItem As String) As Long
#Else
Function LookupItem(ctl As Control, sItem As String) As Long
#End If
    If TypeName(ctl) = "ComboBox" Then
        LookupItem = SendMessageStr(ctl.hWnd, CB_FINDSTRING, -1&, sItem)
    Else
        LookupItem = SendMessageStr(ctl.hWnd, LB_FINDSTRING, -1&, sItem)
    End If
End Function

Function ClassNameFromWnd(ByVal hWnd As Long) As String
    Dim sName As String, cName As Long
    BugAssert hWnd <> hNull
    sName = String$(80, 0)
    cName = GetClassName(hWnd, sName, 80)
    ClassNameFromWnd = Left$(sName, cName)
End Function

Function InstFromWnd(ByVal hWnd As Long) As Long
    BugAssert hWnd <> hNull
    InstFromWnd = GetWindowLong(hWnd, GWL_HINSTANCE)
End Function

Function ProcIDFromWnd(ByVal hWnd As Long) As Long
    Dim idProc As Long
    Call GetWindowThreadProcessId(hWnd, idProc)
    ProcIDFromWnd = idProc
End Function

Function ProcFromWnd(ByVal hWnd As Long) As Long
    BugAssert hWnd <> hNull
    ProcFromWnd = MModTool.ProcFromProcID(ProcIDFromWnd(hWnd))
End Function

Function ThreadIDFromWnd(ByVal hWnd As Long) As Long
    Dim idProc As Long
    BugAssert hWnd <> hNull
    ThreadIDFromWnd = GetWindowThreadProcessId(hWnd, idProc)
End Function

Function GetWndOwner(ByVal hWnd As Long) As String
    Dim hwndOwner As Long
    BugAssert hWnd <> hNull
    hwndOwner = GetWindow(hWnd, GW_OWNER)
    If hwndOwner <> hNull Then
        GetWndOwner = WindowTextLineFromWnd(hwndOwner)
    Else
        GetWndOwner = sEmpty
    End If
End Function

Function IsWindowLocal(ByVal hWnd As Long) As Boolean
    Dim idWnd As Long
    Call GetWindowThreadProcessId(hWnd, idWnd)
    IsWindowLocal = (idWnd = GetCurrentProcessId())
End Function

Function IsVisibleTopWnd(hWnd As Long, _
                Optional IgnoreEmpty As Boolean = False, _
                Optional IgnoreVisible As Boolean = False, _
                Optional IgnoreOwned As Boolean = False) _
                As Boolean
    If IgnoreEmpty Or WindowTextFromWnd(hWnd) <> sEmpty Then
        If IgnoreVisible Or IsWindowVisible(hWnd) Then
            If IgnoreOwned Or GetWindow(hWnd, GW_OWNER) = hNull Then
                IsVisibleTopWnd = True
            End If
        End If
    End If
End Function

Function VBFindWindow(Optional Class As String, _
                      Optional Title As String) As Long
    VBFindWindow = FindWindow(Class, Title)
End Function

Function WindowTextFromWnd(ByVal hWnd As Long) As String
    Dim c As Long, s As String
    c = GetWindowTextLength(hWnd)
    ' Some windows return huge length--ignore 0 <= c < 4096
    If c And &HFFFFF000 Then Exit Function
    s = String$(c, 0)
    c = GetWindowText(hWnd, s, c + 1)
    WindowTextFromWnd = s
End Function

Function WindowTextLineFromWnd(ByVal hWnd As Long) As String
    Dim sTitle As String, cTitle As Long
    sTitle = WindowTextFromWnd(hWnd)
    ' Chop off end of multiline captions
    cTitle = InStr(sTitle, sCr)
    WindowTextLineFromWnd = IIf(cTitle, Left$(sTitle, cTitle), sTitle)
End Function

Function VBFindTopWindow(sClass As String, sTitle As String) As Long
                       
    ' Assume fail for easy exit
    VBFindTopWindow = hNull
    If sClass = sEmpty And sTitle = sEmpty Then Exit Function

    ' Get first sibling to start iterating top level windows
    Dim fClass As Boolean, fTitle As Boolean
    Dim hWnd As Long
    hWnd = GetWindow(GetDesktopWindow(), GW_CHILD)
    Do While hWnd <> hNull
        
        ' Check class
        fClass = True
        If sClass <> sEmpty Then
            fClass = (StrComp(sClass, ClassNameFromWnd(hWnd)) = 0)
        End If
        
        ' Check title
        fTitle = True
        If sTitle <> sEmpty Then
            fTitle = (WindowTextFromWnd(hWnd) Like sTitle)
        End If

        ' Check success
        If fClass And fTitle Then
            VBFindTopWindow = hWnd
            Exit Function
        End If

        ' Get next sibling
        hWnd = GetWindow(hWnd, GW_HWNDNEXT)
    Loop

End Function

Sub ChangeStyleBit(hWnd As Long, f As Boolean, afNew As Long)
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

Function GetStyleBits(hWnd As Long) As Long
    GetStyleBits = GetWindowLong(hWnd, GWL_STYLE)
End Function

Sub ChangeExtStyleBit(hWnd As Long, f As Boolean, afNew As Long)
    Dim af As Long, hParent As Long
    af = GetWindowLong(hWnd, GWL_EXSTYLE)
    If f Then
        af = af Or afNew
    Else
        af = af And (Not afNew)
    End If
    Call SetWindowLong(hWnd, GWL_EXSTYLE, af)
    ' Reset the parent so that change will "take"
    hParent = GetParent(hWnd)
    SetParent hWnd, hParent
    ' Redraw for added insurance
    Call SetWindowPos(hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, _
                      SWP_NOZORDER Or SWP_NOSIZE Or _
                      SWP_NOMOVE Or SWP_DRAWFRAME)
End Sub

Function GetExtStyleBits(hWnd As Long) As Long
    GetExtStyleBits = GetWindowLong(hWnd, GWL_EXSTYLE)
End Function

' Something that uses ChangeStyleBit

Sub SetClipControls(hWnd As Long, f As Boolean)
    ' You want to do this:
    'Me.ClipControls = f
    ' But Visual Basic won't let you; do this instead:
    ChangeStyleBit hWnd, f, WS_CLIPCHILDREN
End Sub
    
Sub ClientToScreenXY(ByVal hWnd As Long, x As Long, y As Long)
    Dim pt As POINTL
    pt.x = x \ Screen.TwipsPerPixelX
    pt.y = y \ Screen.TwipsPerPixelY
    ClientToScreen hWnd, pt
    x = pt.x
    y = pt.y
End Sub
        
Function GetWndStyle(hWnd) As String
    Dim af As Long, s As String
    BugAssert hWnd <> hNull

    ' Get normal style
    af = GetWindowLong(hWnd, GWL_STYLE)
    If af And WS_BORDER Then s = s & "Border "
    If af And WS_CAPTION Then s = s & "Caption "
    If af And WS_CHILD Then s = s & "Child "
    If af And WS_CLIPCHILDREN Then s = s & "ClipChildren "
    If af And WS_CLIPSIBLINGS Then s = s & "ClipSiblings "
    If af And WS_DLGFRAME Then s = s & "DlgFrame "
    If af And WS_GROUP Then s = s & "Group "
    If af And WS_HSCROLL Then s = s & "HScroll "
    If af And WS_MAXIMIZEBOX Then s = s & "MaximizeBox "
    If af And WS_MINIMIZEBOX Then s = s & "MinimizeBox "
    If af And WS_POPUP Then s = s & "Popup "
    If af And WS_SYSMENU Then s = s & "SysMenu "
    If af And WS_TABSTOP Then s = s & "TabStop "
    If af And WS_THICKFRAME Then s = s & "ThickFrame "
    If af And WS_VSCROLL Then s = s & "VScroll "

    ' Get extended style
    af = GetWindowLong(hWnd, GWL_EXSTYLE)
    If af And WS_EX_DLGMODALFRAME Then s = s & "DlgModalFrame "
    If af And WS_EX_NOPARENTNOTIFY Then s = s & "NoParentNotify "
    If af And WS_EX_TOPMOST Then s = s & "Topmost "
    If af And WS_EX_ACCEPTFILES Then s = s & "AcceptFiles "
    If af And WS_EX_TRANSPARENT Then s = s & "Transparent "

    GetWndStyle = s

End Function

Public Function GetWndInfo(ByVal hWnd As Long, Optional TabStop As Integer = 0) As String
    Dim sStart As String, s As String, sTemp As String
    BugAssert hWnd <> hNull
    
    ' Nested starting position
    sStart = Space$(TabStop * 4)
    ' Window information
    sTemp = WindowTextLineFromWnd(hWnd)
    'sTemp = WindowTextFromWnd(hWnd)
    If sTemp = sEmpty Then sTemp = "[none]"
    s = sStart & "Title: " & sTemp & sCrLf
    s = s & sStart & "Class: " & ClassNameFromWnd(hWnd) & sCrLf
    s = s & sStart & "Style: " & GetWndStyle(hWnd) & sCrLf
    sTemp = GetWndOwner(hWnd)
    If sTemp <> sEmpty Then
        s = s & sStart & "Owner: " & sTemp & sCrLf
    End If

    GetWndInfo = s

End Function

Public Function GetWndView(hWnd) As String
    Dim s As String
    BugAssert hWnd <> hNull
    s = IIf(IsWindowVisible(hWnd), "Visible ", "Invisible ")
    s = s & IIf(IsWindowEnabled(hWnd), "Enabled", "Disabled ")
    s = s & IIf(IsZoomed(hWnd), "Zoomed ", sEmpty)
    s = s & IIf(IsIconic(hWnd), "Iconic ", sEmpty)
    GetWndView = s
End Function

Function GetTextExtentWnd(ByVal hWnd As Long, s As String, _
                     Optional dy As Single) As Single
                     
    Dim hDC As Long, sz As SIZEL, f As Long
    hDC = GetDC(hWnd)
    f = GetTextExtentPoint32(hDC, s, Len(s), sz)
    If f Then
        ' Most common x value in return
        GetTextExtentWnd = sz.cx
        ' Optional y value through reference variable
        dy = sz.cy
    End If
    Call ReleaseDC(hWnd, hDC)
                     
End Function

Property Get WindowsDir() As String
    Dim s As String, c As Long
    s = String$(cMaxPath, 0)
    c = GetWindowsDirectory(s, cMaxPath)
    WindowsDir = Left(s, c)
End Property

Property Get SystemDir() As String
    Dim s As String, c As Long
    s = String$(cMaxPath, 0)
    c = GetSystemDirectory(s, cMaxPath)
    SystemDir = Left(s, c)
End Property
'

#If fComponent = 0 Then
Private Sub ErrRaise(e As Long)
    Dim sText As String, sSource As String
    If e > 1000 Then
        sSource = App.ExeName & ".WinTool"
        Select Case e
        Case eeBaseWinTool
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

