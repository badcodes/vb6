VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSysMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ISubclass
Public Event MenuClick(sItem As String, ByVal ID As Long)

Public Enum EErrorSysMenu
    eeBaseSysMenu = 13240   ' CSysMenu
    eeHandleNotInit         ' Handle not initialized
    eeTooManyItems          ' Too many menu items
    eeNoSysWindow           ' Can't get system window
End Enum

Private hWnd As Long, hSysMenu As Long
Private idCur As Long
Private emr As EMsgResponse, pOldProc As Long
Private Const idFirst As Long = 2000
Private Const cMaxItem = 10

Private Type TItem
    sText As String
    ID As Long
End Type

' Anyone putting too many items on system menu deserves rude errors

Private aItem(1 To cMaxItem + 1) As TItem

#If fComponent = 0 Then
Private Sub ErrRaise(e As Long)
    Dim sText As String, sSource As String
    If e > 1000 Then
    
        sSource = App.ExeName & ".CSysMenu"
        Select Case e
        Case eeHandleNotInit
            sText = "Add/RemoveItem: Handle not initialized"
        Case eeTooManyItems
            sText = "AddItem: Too many menu items"
        Case eeNoSysWindow
            sText = "Create: Can't get system window"
        End Select
        Err.Raise e Or vbObjectError, sSource, sText
    Else
        ' Raise standard Visual Basic error
        Err.Raise e, sSource
    End If
End Sub
#End If

Private Sub Class_Initialize()
    idCur = idFirst
End Sub

Private Sub Class_Terminate()
    Destroy
End Sub

Sub Create(hWndA As Long)
    ' Get handle of system menu
    hSysMenu = GetSystemMenu(hWndA, 0&)
    If hSysMenu = hNull Then ErrRaise eeNoSysWindow
    hWnd = hWndA
    AttachMessage Me, hWndA, WM_SYSCOMMAND
End Sub

Sub Destroy()
    Dim i As Integer
    ' Remove all the menu items
    i = 1
    Do While aItem(i).ID
        Call RemoveMenu(hSysMenu, aItem(i).ID, MF_BYCOMMAND)
        i = i + 1
    Loop
    DetachMessage Me, hWnd, WM_SYSCOMMAND
    hWnd = hNull
End Sub

Property Get Identifier(i As Integer) As Long
    Identifier = aItem(i).ID
End Property

Property Get Text(i As Integer) As String
    Text = aItem(i).sText
End Property

Function AddItem(sItem As String) As Long
    If hWnd = hNull Then
        ErrRaise eeHandleNotInit
    End If
    ' Append the new menu item or separator
    idCur = idCur + 10
    If sItem = sEmpty Or sItem = "-" Then
        Call AppendMenu(hSysMenu, MF_SEPARATOR Or MF_BYCOMMAND, idCur, 0&)
    Else
        Call AppendMenu(hSysMenu, MF_BYCOMMAND, idCur, sItem)
    End If
    ' Insert item
    Dim i As Integer
    For i = 1 To cMaxItem
        If aItem(i).ID = 0 Then
            aItem(i).ID = idCur
            aItem(i).sText = sItem
            AddItem = idCur
            Exit Function
        End If
    Next
    ErrRaise eeTooManyItems
End Function

Sub RemoveItem(ID As Long)
    If hWnd = hNull Then
        ErrRaise eeHandleNotInit
    End If
    ' Find item, remove it, overwrite it
    Dim f As Boolean, i As Integer, fDeleting As Boolean
    i = 1
    For i = 1 To cMaxItem
        ' Remove reference from vector
        If aItem(i).ID = ID Then
            f = RemoveMenu(hSysMenu, ID, MF_BYCOMMAND)
            fDeleting = True
        End If
        ' Overwrite deleted item
        If fDeleting Then
            aItem(i).ID = aItem(i + 1).ID
            aItem(i).sText = aItem(i + 1).sText
        End If
    Next
End Sub

' Interface window procedure method
Private Function ISubclass_WindowProc(ByVal hWnd As Long, _
                                      ByVal iMsg As Long, _
                                      ByVal wParam As Long, _
                                      ByVal lParam As Long) As Long
    ' Assume original WindowProc will handle
    emr = emrPostProcess
    ' Subclasser should never call unless it's our message
    BugAssert iMsg = WM_SYSCOMMAND
    ' Ignore everything except system commands
    If wParam <= 3000 Then
        ' Check IDs and raise event if found
        Dim i As Long
        For i = 1 To cMaxItem
            If aItem(i).ID = 0 Then Exit For
            If wParam = aItem(i).ID Then
                RaiseEvent MenuClick(aItem(i).sText, aItem(i).ID)
                ' We've finished so original WindowProc not needed
                emr = emrConsume
                Exit Function
            End If
        Next
    End If
End Function

' Interface properties
Private Property Get ISubclass_MsgResponse() As EMsgResponse
    ISubclass_MsgResponse = emr
End Property
Private Property Let ISubclass_MsgResponse(ByVal emrA As EMsgResponse)
    emr = emrA
End Property
'

