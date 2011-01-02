VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GSubclass"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' SubTimer is independent of VBCore, so it hard codes error handling

Public Enum EErrorWindowProc
    eeBaseWindowProc = 13080 ' WindowProc
    eeCantSubclass           ' Can't subclass window
    eeAlreadyAttached        ' Message already handled by another class
    eeInvalidWindow          ' Invalid window
    eeNoExternalWindow       ' Can't modify external window
End Enum

Private Sub ErrRaise(e As Long)
    Dim sText As String, sSource As String
    If e > 1000 Then
        sSource = App.EXEName & ".WindowProc"
        Select Case e
        Case eeCantSubclass
            sText = "Can't subclass window"
        Case eeAlreadyAttached
            sText = "Message already handled by another class"
        Case eeInvalidWindow
            sText = "Invalid window"
        Case eeNoExternalWindow
            sText = "Can't modify external window"
        End Select
        Err.Raise e Or vbObjectError, sSource, sText
    Else
        ' Raise standard Visual Basic error
        Err.Raise e, sSource
    End If
End Sub

Sub AttachMessage(iwp As ISubclass, ByVal hWnd As Long, _
                  ByVal iMsg As Long)
    Dim procOld As Long, f As Long, c As Long
    ' Validate window
    If IsWindow(hWnd) = False Then ErrRaise eeInvalidWindow
    If IsWindowLocal(hWnd) = False Then ErrRaise eeNoExternalWindow

    ' Get the message count
    c = GetProp(hWnd, "C" & hWnd)
    If c = 0 Then
        ' Subclass window by installing window procecure
        procOld = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf WindowProc)
        If procOld = 0 Then ErrRaise eeCantSubclass
        ' Associate old procedure with handle
        f = SetProp(hWnd, hWnd, procOld)
        BugAssert f <> 0
        ' Count this message
        c = 1
        f = SetProp(hWnd, "C" & hWnd, c)
    Else
        ' Count this message
        c = c + 1
        f = SetProp(hWnd, "C" & hWnd, c)
    End If
    BugAssert f <> 0
    ' This message had better not be already attached
    If GetProp(hWnd, hWnd & "#" & iMsg) <> pNull Then
        ErrRaise eeAlreadyAttached
    End If
    ' Associate object with message (one per handle)
    f = SetProp(hWnd, hWnd & "#" & iMsg, ObjPtr(iwp))
    BugAssert f <> 0
End Sub

Sub DetachMessage(iwp As ISubclass, ByVal hWnd As Long, _
                  ByVal iMsg As Long)
    Dim procOld As Long, f As Long, c As Long
    ' Get the message count
    c = GetProp(hWnd, "C" & hWnd)
    If c = 1 Then
        ' This is the last message, so unsubclass
        procOld = GetProp(hWnd, hWnd)
        BugAssert procOld <> pNull
        ' Unsubclass by reassigning old window procedure
        Call SetWindowLong(hWnd, GWL_WNDPROC, procOld)
        ' Remove unneeded handle (oldProc)
        RemoveProp hWnd, hWnd
        ' Remove unneeded count
        RemoveProp hWnd, "C" & hWnd
    Else
        ' Uncount this message
        c = GetProp(hWnd, "C" & hWnd)
        c = c - 1
        f = SetProp(hWnd, "C" & hWnd, c)
    End If
    ' Remove unneeded message (subclass object pointer)
    RemoveProp hWnd, hWnd & "#" & iMsg
End Sub

Private Function WindowProc(ByVal hWnd As Long, ByVal iMsg As Long, _
                            ByVal wParam As Long, ByVal lParam As Long) _
                            As Long
    Dim procOld As Long, pSubclass As Long, f As Long
    Dim iwp As ISubclass, iwpT As ISubclass
    ' Get the old procedure from the window
    procOld = GetProp(hWnd, hWnd)
    BugAssert procOld <> pNull
    ' Get the object pointer from the message
    pSubclass = GetProp(hWnd, hWnd & "#" & iMsg)
    If pSubclass = pNull Then
        ' This message not handled, so pass on to old procedure
        WindowProc = CallWindowProc(procOld, hWnd, iMsg, _
                                    wParam, ByVal lParam)
        Exit Function
    End If
    
    ' Turn the pointer into an illegal, uncounted interface
    CopyMemory iwpT, pSubclass, 4
    ' Do NOT hit the End button here! You will crash!
    BugMessage "Got object"
    ' Assign to legal reference
    Set iwp = iwpT
    ' Still do NOT hit the End button here! You will still crash!
    ' Destroy the illegal reference
    CopyMemory iwpT, 0&, 4
    ' OK, hit the End button if you must--you'll probably still crash,
    ' but it will be because of the subclass, not the uncounted reference
    
    ' Use the interface to call back to the class
    With iwp
        ' Preprocess
        If .MsgResponse = emrPreprocess Then
            WindowProc = CallWindowProc(procOld, hWnd, iMsg, _
                                        wParam, ByVal lParam)
        End If
        ' Consume
        WindowProc = .WindowProc(hWnd, iMsg, wParam, ByVal lParam)
        ' PostProcess
        If .MsgResponse = emrPostProcess Then
            WindowProc = CallWindowProc(procOld, hWnd, iMsg, _
                                        wParam, ByVal lParam)
        End If
    End With
    
End Function

' Cheat! Cut and paste from MWinTool rather than reusing
' file because reusing file would cause many unneeded dependencies
Function IsWindowLocal(ByVal hWnd As Long) As Boolean
    Dim idWnd As Long
    Call GetWindowThreadProcessId(hWnd, idWnd)
    IsWindowLocal = (idWnd = GetCurrentProcessId())
End Function
'


