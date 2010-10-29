VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CKeyboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum EErrorKeyboard
    eeBaseKeyboard = 13080  ' CKeyboard
End Enum

Private abKeys(0 To 255) As Byte

' Get or set toggle state of Caps Lock key
Property Get CapsState() As Boolean
    ' Get toggled and pressed state of all keys
    GetKeyboardState abKeys(0)
    ' Check low bit for state
    CapsState = abKeys(VK_CAPITAL) And 1
End Property

Property Let CapsState(fCapsStateA As Boolean)
    ' Get toggled and pressed state of all keys
    GetKeyboardState abKeys(0)
    ' Set low bit to new pressed state
    If fCapsStateA Then
        abKeys(VK_CAPITAL) = abKeys(VK_CAPITAL) Or 1
    Else
        abKeys(VK_CAPITAL) = abKeys(VK_CAPITAL) And &HFE
    End If
    ' Store changed array
    SetKeyboardState abKeys(0)
End Property

' Get or set toggle state of Insert key
Property Get InsState() As Boolean
    ' Get toggle and pressed state of all keys
    GetKeyboardState abKeys(0)
    ' Check low bit for state
    InsState = abKeys(VK_INSERT) And 1
End Property

Property Let InsState(fCapsStateA As Boolean)
    ' Get toggle and pressed state of all keys
    GetKeyboardState abKeys(0)
    ' Set low bit to new pressed state
    If fCapsStateA Then
        abKeys(VK_INSERT) = abKeys(VK_INSERT) Or 1
    Else
        abKeys(VK_INSERT) = abKeys(VK_INSERT) And &HFE
    End If
    ' Store changed array
    SetKeyboardState abKeys(0)
End Property

' Get or set toggle state of Num Lock key
Property Get NumState() As Boolean
    ' Get toggle and pressed state of all keys
    GetKeyboardState abKeys(0)
    ' Check low bit for state
    NumState = abKeys(VK_NUMLOCK) And 1
End Property

Property Let NumState(fCapsStateA As Boolean)
    ' Get toggle and pressed state of all keys
    GetKeyboardState abKeys(0)
    ' Set low bit to new pressed state
    If fCapsStateA Then
        abKeys(VK_NUMLOCK) = abKeys(VK_NUMLOCK) Or 1
    Else
        abKeys(VK_NUMLOCK) = abKeys(VK_NUMLOCK) And &HFE
    End If
    ' Store changed array
    SetKeyboardState abKeys(0)
End Property

' Get or set toggle state of Scroll Lock key
Property Get ScrollState() As Boolean
    ' Get toggle and pressed state of all keys
    GetKeyboardState abKeys(0)
    ' Check low bit for state
    ScrollState = abKeys(VK_SCROLL) And 1
End Property

Property Let ScrollState(fCapsStateA As Boolean)
    ' Get toggle and pressed state of all keys
    GetKeyboardState abKeys(0)
    ' Set low bit to new pressed state
    If fCapsStateA Then
        abKeys(VK_SCROLL) = abKeys(VK_SCROLL) Or 1
    Else
        abKeys(VK_SCROLL) = abKeys(VK_SCROLL) And &HFE
    End If
    ' Store changed array
    SetKeyboardState abKeys(0)
End Property

' Get or set toggle state of any toggle key
Property Get KeyState(iKey As Integer) As Boolean
    ' Get toggle and pressed state of all keys
    GetKeyboardState abKeys(0)
    ' Check low bit for state
    KeyState = abKeys(iKey) And 1
End Property

Property Let KeyState(iKey As Integer, fKeyStateA As Boolean)
    ' Get toggle and pressed state of all keys
    GetKeyboardState abKeys(0)
    ' Set low bit to new pressed state
    If fKeyStateA Then
        abKeys(iKey) = abKeys(iKey) Or 1
    Else
        abKeys(iKey) = abKeys(iKey) And &HFE
    End If
    ' Store changed array
    SetKeyboardState abKeys(0)
End Property

' Get or set pressed state of any key
Property Get KeyPressed(iKey As Integer) As Boolean
    ' Get toggled and pressed state of all keys
    GetKeyboardState abKeys(0)
    ' Check high bit for state
    KeyPressed = abKeys(iKey) And &H80
End Property

Property Let KeyPressed(iKey As Integer, fKeyPressedA As Boolean)
    ' Get toggle and pressed state of all keys
    GetKeyboardState abKeys(0)
    ' Set high bit to new pressed state
    If fKeyPressedA Then
        abKeys(iKey) = abKeys(iKey) Or &H80
    Else
        abKeys(iKey) = abKeys(iKey) And &H7F
    End If
    ' Store changed array
    SetKeyboardState abKeys(0)
End Property

#If fComponent = 0 Then
Private Sub ErrRaise(e As Long)
    Dim sText As String, sSource As String
    If e > 1000 Then
        sSource = App.ExeName & ".Keyboard"
        Select Case e
        Case eeBaseKeyboard
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

