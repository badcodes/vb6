VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTrayIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ISubclass
Event TrayMessage(ByVal Message As Long, ByVal ID As Long)

Public Enum EErrorTrayIcon
    eeBaseTrayIcon = 13260  ' CTrayIcon
End Enum

Private picIcon As StdPicture
Private icodat As NOTIFYICONDATA
Private iMsg As Long, hWnd As Long, procOld As Long
Private emr As EMsgResponse

Sub Create(ByVal hWndA As Long, picIconA As StdPicture, sTip As String)
    ' Verify that we have valid handle and icon
    If IsWindow(hWndA) = False Then ApiRaise ERROR_INVALID_HANDLE
    If picIconA.Type <> vbPicTypeIcon Then ApiRaise ERROR_INVALID_DATA
    ' Register a unique message based on object pointer
    hWnd = hWndA
    iMsg = RegisterWindowMessage(Hex$(ObjPtr(Me)))
    If iMsg = 0 Then ApiRaise Err.LastDllError
    ' Fill the structure and pass to Shell_NotifyIcon
    Set picIcon = picIconA
    icodat.cbSize = Len(icodat)
    icodat.hWnd = hWnd
    ' ID is same as message
    icodat.uid = iMsg
    ' Handle message, icon, and tip
    icodat.uFlags = NIF_MESSAGE Or NIF_ICON Or NIF_TIP
    icodat.uCallbackMessage = iMsg
    icodat.hIcon = picIcon.Handle
    MBytes.StrToBytes icodat.szTip, sTip
    If Shell_NotifyIcon(NIM_ADD, icodat) = False Then
        ApiRaise Err.LastDllError
    End If
    Debug.Print "Create notify icon"
    ' Subclass this message
    AttachMessage Me, hWnd, iMsg
End Sub

Sub Destroy()
    If iMsg = 0 Then Exit Sub
    Debug.Print "Destroy notify icon"
    ' Unsubclass the message and destroy the icon
    DetachMessage Me, icodat.hWnd, icodat.uCallbackMessage
    Shell_NotifyIcon NIM_DELETE, icodat
    hWnd = hNull
End Sub

Property Get Message() As Long
    Message = icodat.uCallbackMessage
End Property

Property Get ID() As Long
    ID = icodat.uid
End Property

Property Get OldProc() As Long
    OldProc = procOld
End Property

Property Get Tip() As String
    Tip = MBytes.BytesToStr(icodat.szTip)
End Property

Property Let Tip(sTip As String)
    If hWnd = hNull Then Exit Property
    MBytes.StrToBytes icodat.szTip, sTip & vbNullChar
    icodat.uFlags = NIF_TIP
    Shell_NotifyIcon NIM_MODIFY, icodat
End Property

Property Get Icon() As Picture
    Set Icon = picIcon
End Property

Property Set Icon(picIconA As Picture)
    If hWnd = hNull Then Exit Property
    If picIconA.Type <> vbPicTypeIcon Then Exit Property
    Set picIcon = picIconA
    icodat.hIcon = picIcon.Handle
    icodat.uFlags = NIF_ICON
    Shell_NotifyIcon NIM_MODIFY, icodat
End Property

Private Sub Class_Terminate()
    Destroy
End Sub

' Implement ISubclass

Private Property Let ISubclass_MsgResponse(ByVal emrA As EMsgResponse)
    emr = emrA
End Property

Private Property Get ISubclass_MsgResponse() As EMsgResponse
    ISubclass_MsgResponse = emr
End Property

Private Function ISubclass_WindowProc(ByVal hWnd As Long, _
                                      ByVal iMsg As Long, _
                                      ByVal wParam As Long, _
                                      ByVal lParam As Long) As Long
    ' Subclasser should never call unless it's our message
    BugAssert iMsg = icodat.uCallbackMessage
    ' Pass event back to client (message is lParam, ID is wParam
    RaiseEvent TrayMessage(lParam, wParam)
    ' We've finished so original WindowProc not needed
    emr = emrConsume
End Function
    
#If fComponent = 0 Then
Private Sub ErrRaise(e As Long)
    Dim sText As String, sSource As String
    If e > 1000 Then
        sSource = App.ExeName & ".TrayIcon"
        Select Case e
        Case eeBaseTrayIcon
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

