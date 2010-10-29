VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMenuItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' This is a VB4 class from Hardcore Visual Basic, First Edition. It showed some
' interesting techniques, so I left it in the sample. It could probably be improved
' using Friend and other new features.

' Don't call this class directly (access only through CMenuList)

Private hWnd As Long, hMenu As Long
Private idID As Long, iPos As Long, fSys As Boolean
Private rmenuChild As CMenuList

Sub Class_Initialize()
    Set rmenuChild = Nothing
End Sub

Function Create(iPosA As Long, hMenuA As Long, _
                rmenuA As CMenuList) As Boolean
    ' Store properties
    BugAssert IsMenu(hMenuA)
    hMenu = hMenuA: iPos = iPosA
    hWnd = rmenuA.WinHandle: fSys = rmenuA.SysMenu
    idID = GetMenuItemID(hMenu, iPos)
    
    ' Create new menu list for any submenu
    If idID = -1 Then
        Dim menu As New CMenuList, f As Boolean
        ' Must set parent before creating--yuck!
        Set menu.Parent = rmenuA
        f = menu.Create(GetSubMenu(hMenu, iPos))
        BugAssert f     ' Should never fail
        Set rmenuChild = menu
    End If
    Create = True
End Function

' Read-only properties to get state
Property Get Separator() As Boolean
    Checked = MF_SEPARATOR And GetMenuState(hMenu, iPos, MF_BYPOSITION)
End Property

Property Get id() As Boolean
    id = idID
End Property

Property Get Popup() As Boolean
    Popup = (idID = -1)
End Property

Property Get WinHandle() As Long
    WinHandle = hWnd
End Property

Property Get Child() As CMenuList
    Set Child = rmenuChild
End Property

' Convert text to recognizable name by stripping unnecessary parts
Property Get Name() As String
    Dim ch As String, sText As String
    Dim s As String, i As Integer, c As Integer
    sText = Text
    i = 1: c = Len(sText)

    ' Skip any leading spaces and tabs
    Do While i <= c
        ch = Mid$(sText, i, 1)
        If ch <> sSpace And ch <> sTab Then Exit Do
        i = i + 1
    Loop

    ' Strip any ampersand (&) and chop off anything after tab or ...
    Do While i <= c
        Select Case ch
        Case "&"
            ' Continue, skipping ampersand
        Case sTab
            ' Truncate at any tab after leading tab
            Exit Do
        Case "."
            If Mid$(sText, i, 3) = "..." Then
                Exit Do
            Else
                s = s & ch
            End If
        Case Else
            ' Append normal letters
            s = s & ch
        End Select
        ' Next letter
        i = i + 1
        ch = Mid$(sText, i, 1)
    Loop
    Name = s
End Property

' Read/write properties to get or set state
Property Get Text() As String
    Dim s As String, c As Integer
    Const cMaxStr = 80
    s = String$(cMaxStr, 0)
    c = GetMenuString(hMenu, iPos, s, cMaxStr, MF_BYPOSITION)
    Text = Left$(s, c)
End Property

Property Let Text(sTextA As String)
    Dim afState As Long
    afState = GetMenuState(hMenu, iPos, MF_BYPOSITION)
    afState = afState Or MF_BYPOSITION Or MF_STRING
    Call ModifyMenu(hMenu, iPos, afState, idID, sTextA)
End Property

Property Get Disabled() As Boolean
    Disabled = MF_DISABLED And GetMenuState(hMenu, iPos, MF_BYPOSITION)
End Property

' Windows allows Disabled Ungrayed menus, but we don't
Property Let Disabled(fDisabledA As Boolean)
    Dim afState As Long
    afState = GetMenuState(hMenu, iPos, MF_BYPOSITION)
    If fDisabledA Then
        afState = afState Or MF_DISABLED Or MF_GRAYED
    Else
        afState = afState And Not (MF_DISABLED Or MF_GRAYED)
    End If
    Call EnableMenuItem(hMenu, iPos, afState Or MF_BYPOSITION)
    DrawMenuBar hWnd
End Property

Property Get Grayed() As Boolean
    Grayed = MF_GRAYED And GetMenuState(hMenu, iPos, MF_BYPOSITION)
End Property

' Windows allows Grayed Enabled menus, but we don't
Property Let Grayed(fGrayedA As Boolean)
    Disabled = fGrayedA
End Property

Property Get Checked() As Boolean
    Checked = MF_CHECKED And GetMenuState(hMenu, iPos, MF_BYPOSITION)
End Property

Property Let Checked(fCheckedA As Boolean)
    Dim afState As Long
    afState = GetMenuState(hMenu, iPos, MF_BYPOSITION)
    If fCheckedA Then
        afState = afState Or MF_CHECKED
    Else
        afState = afState And Not (MF_CHECKED)
    End If
    Call CheckMenuItem(hMenu, idID, afState)
    DrawMenuBar hWnd
End Property

' Methods
Sub Hilite()
    Call HiliteMenuItem(hWnd, hMenu, iPos, MF_BYPOSITION Or MF_HILITE)
End Sub

Sub UnHilite()
    Call HiliteMenuItem(hWnd, hMenu, iPos, MF_BYPOSITION Or MF_UNHILITE)
End Sub

Sub Remove()
    Call RemoveMenu(hMenu, iPos, MF_BYPOSITION)
End Sub

' Execute current item and return results
Function Execute() As Boolean
    Dim iMsg As Long
    iMsg = IIf(fSys, WM_SYSCOMMAND, WM_COMMAND)
    Execute = (SendMessage(hWnd, iMsg, ByVal idID, ByVal 0&) = 0)
End Function
'
