VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMenuList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' This is a VB4 class from Hardcore Visual Basic, First Edition. It showed some
' interesting techniques, so I left it in the sample. It could probably be improved
' using Friend and other new features.

Private nItems As New Collection
Private hMenu As Long, idMax As Long
' Warning! These members should be static--that is, one variable
' available to all objects of the given type. The entire menu tree
' has only one window handle. But Basic doesn't have static
' variables, so you must fake them.
Private hWnd As Long, fSys As Boolean
' Warning! Violation of encapsulation standards! This property
' should not be public, but its partner class won't work otherwise.
' Don't use this property directly.
Public Parent As CMenuList

Sub Class_Initialize()
    ' Make sure Parent property is initialized
    Set Parent = Nothing
End Sub

Sub Class_Terminate()
    DestroyMenus
End Sub

' Call this by passing window handle (it can take menu handle, but
' should never get one through caller)
Function Create(hA As Long, Optional fSysA As Boolean = False) As Boolean
    fSys = fSysA
    If IsWindow(hA) Then
        ' Create system or normal menu from hWnd
        If fSys Then
            hMenu = GetSystemMenu(hA, False)
        Else
            hMenu = GetMenu(hA)
        End If
        hWnd = hA
    Else
        ' Don't accept menu handle from top node
        If IsMenu(hA) And Parent Is Nothing Then Exit Function
        hMenu = hA: hWnd = WinHandle
    End If

    ' Create each item in list and add to collection
    Dim item As CMenuItem, i As Long, f As Boolean
    DestroyMenus
    For i = 0 To Count - 1
        Set item = New CMenuItem
        ' Create will also create new submenus
        f = item.Create(i, hMenu, Me)
        BugAssert f     ' Should never fail
        nItems.Add item
        ' Needed by InsertNew method
        If item.id > idMax Then idMax = item.id
    Next
    Create = True
End Function

' Redraw after each change
Sub ReDraw()
    DrawMenuBar hWnd
End Sub

' Insert new item, moving everything down
Function InsertNew(sItem As String, Optional iPos As Long = 0, _
                   Optional afFlags As Long = MF_STRING Or MF_DISABLED) As Boolean
    If iPos <= 0 Then iPos = Count   ' Append to end
    InsertNew = InsertMenu(hMenu, iPos, MF_BYPOSITION Or afFlags, _
                           idMax + 10, sItem)
    Dim f As Boolean, item As CMenuItem
    f = item.Create(CInt(iPos), hMenu, Me)
    BugAssert f     ' Should never fail
    nItems.Add item
    If item.id > idMax Then idMax = item.id
End Function

' Move up through recursive levels and recreate window from top
Function Refresh()
    If Parent Is Nothing Then
        Refresh = Create(hWnd, fSys)
    Else
        Refresh = Parent.Refresh
    End If
End Function

' Number of items in menu
Property Get Count() As Integer
    Count = GetMenuItemCount(hMenu)
    If Count = -1 Then Count = 0
End Property

' Move up through recursive levels to find top level window handle
' (static member would be much easier)
Property Get WinHandle() As Long
    If Parent Is Nothing Then
        WinHandle = hWnd
    Else
        WinHandle = Parent.WinHandle
    End If
End Property

Property Get SysMenu() As Boolean
    If Parent Is Nothing Then
        SysMenu = fSys
    Else
        SysMenu = Parent.SysMenu
    End If
End Property

' Clear everything from here on down
Public Sub DestroyMenus()
    Dim item As CMenuItem
    For Each item In nItems
        If item.Popup Then item.Child.DestroyMenus
        Set item = Nothing
        nItems.Remove 1
    Next
End Sub

' Find menu item by its string name
Function Find(sName As String, item As CMenuItem) As Boolean
    Find = True
    Dim i As Integer
    ' Step through each item, searching for match
    For Each item In nItems
        ' Test against name (stripped version of text)
        If item.Name Like sName Then Exit Function
        If item.Popup Then
            ' Recurse through any submenus
            If item.Child.Find(sName, item) Then Exit Function
        End If
    Next
    ' If we got all way through, it's not there
    Find = False
End Function

Sub Walk(Optional iLevel As Integer = 0)
    Dim item As CMenuItem
    For Each item In nItems
        ' Walk through current list until user says stop
        If Not MenuWalker(item, iLevel) Then Exit Sub
        ' Recurse through submenus
        If item.Popup Then item.Child.Walk iLevel + 1
    Next
End Sub
    
#If fMenuWalker = 0 Then
Function MenuWalker(item As CMenuItem, iLevel As Integer) As Boolean
    Dim s As String
    s = "Name: " & item.Name & " ( "
    s = s & IIf(item.Disabled, "Disabled ", "")
    s = s & IIf(item.Checked, "Checked ", "")
    s = s & IIf(item.Grayed, "Grayed ", "")
    s = s & IIf(item.Popup, "Popup ", "") & ")"
    Debug.Print String$(iLevel, sTab) & s
    ' Your MenuWalker can return False to stop walk
    MenuWalker = True
End Function
#End If
'
