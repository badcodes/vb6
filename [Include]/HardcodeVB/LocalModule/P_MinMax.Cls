VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMinMax"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum EErrorMinMax
    eeBaseMinMax = 13120    ' CMinMax
End Enum

Implements ISubclass

Private tx As Long, ty As Long
Private emr As EMsgResponse
Private mmi As MINMAXINFO
Private hWnd As Long

Private Sub Class_Initialize()
    ' Do object access only once
    tx = Screen.TwipsPerPixelX
    ty = Screen.TwipsPerPixelY
    ' Signal default
    mmi.ptMaxSize.x = -1
    mmi.ptMaxSize.y = -1
    mmi.ptMaxPosition.x = -1
    mmi.ptMaxPosition.y = -1
    mmi.ptMinTrackSize.x = -1
    mmi.ptMinTrackSize.y = -1
    mmi.ptMaxTrackSize.x = -1
    mmi.ptMaxTrackSize.y = -1
End Sub

Private Sub Class_Terminate()
    Destroy
End Sub

Sub Create(hWndA As Long)
    ' Get handle of system menu
    hWnd = hWndA
    AttachMessage Me, hWndA, WM_GETMINMAXINFO
End Sub

Sub Destroy()
    DetachMessage Me, hWnd, WM_GETMINMAXINFO
    hWnd = hNull
End Sub

' Interface window procedure method
Private Function ISubclass_WindowProc(ByVal hWnd As Long, _
                                      ByVal iMsg As Long, _
                                      ByVal wParam As Long, _
                                      ByVal lParam As Long) As Long
    Dim mmiT As MINMAXINFO
    ' Copy parameter to local variable for processing
    CopyMemory mmiT, ByVal lParam, LenB(mmiT)
    ' Subclasser should never call unless it's our message
    BugAssert iMsg = WM_GETMINMAXINFO
    
    ' Maximized width and height
    With mmi.ptMaxSize
        If .x <> -1 Then mmiT.ptMaxSize.x = .x
        If .y <> -1 Then mmiT.ptMaxSize.y = .y
    End With
    
    ' Maximized position of top left
    With mmi.ptMaxPosition
        If .x <> -1 Then mmiT.ptMaxPosition.x = .x
        If .y <> -1 Then mmiT.ptMaxPosition.y = .y
    End With
    
    ' Minimium width and height for sizing
    With mmi.ptMinTrackSize
        If .x <> -1 Then mmiT.ptMinTrackSize.x = .x
        If .y <> -1 Then mmiT.ptMinTrackSize.y = .y
    End With
    
    ' Maximium width and height for sizing
    With mmi.ptMaxTrackSize
        If .x <> -1 Then mmiT.ptMaxTrackSize.x = .x
        If .y <> -1 Then mmiT.ptMaxTrackSize.y = .y
    End With
    
    ' Copy modified results back to parameter
    CopyMemory ByVal lParam, mmiT, LenB(mmiT)
    
    ' Don't pass back to original WindowProc
    emr = emrConsume
End Function

' Interface properties
Private Property Get ISubclass_MsgResponse() As EMsgResponse
    ISubclass_MsgResponse = emr
End Property
Private Property Let ISubclass_MsgResponse(ByVal emrA As EMsgResponse)
    emr = emrA
End Property

Property Get MinWidth() As Long
    MinWidth = mmi.ptMinTrackSize.x * tx
End Property
Property Let MinWidth(ByVal dxMinA As Long)
    ' Must be positive, less than screen, and less than maximum
    If dxMinA <= 0 Or dxMinA > Screen.Width Then Exit Property
    If (mmi.ptMaxTrackSize.x > 0) And _
       (dxMinA > (mmi.ptMaxTrackSize.x * tx)) Then Exit Property
    mmi.ptMinTrackSize.x = dxMinA / tx
End Property

Property Get MinHeight() As Long
    MinHeight = mmi.ptMinTrackSize.y * ty
End Property
Property Let MinHeight(ByVal dyMinA As Long)
    ' Must be positive, less than screen, and less than maximum
    If dyMinA <= 0 Or dyMinA > Screen.Height Then Exit Property
    If (mmi.ptMaxTrackSize.y > 0) And _
       (dyMinA > (mmi.ptMaxTrackSize.y * ty)) Then Exit Property
    mmi.ptMinTrackSize.y = dyMinA / ty
End Property

Property Get MaxWidth() As Long
    MaxWidth = mmi.ptMaxTrackSize.x * tx
End Property
Property Let MaxWidth(ByVal dxMaxA As Long)
With mmi
    ' Must be less than screen and greater than minimimum
    If dxMaxA > Screen.Width Or dxMaxA < (.ptMinTrackSize.x * tx) Then
        Exit Property
    End If
    .ptMaxTrackSize.x = dxMaxA / tx
    ' Maximized size can't be greater than maximimum size
    If .ptMaxTrackSize.x > .ptMaxSize.x Then .ptMaxSize.x = .ptMaxTrackSize.x
End With
End Property

Property Get MaxHeight() As Long
    MaxHeight = mmi.ptMaxTrackSize.y * ty
End Property
Property Let MaxHeight(ByVal dyMaxA As Long)
With mmi
    ' Must be less than screen and greater than minimimum
    If dyMaxA > Screen.Width Or dyMaxA < (.ptMinTrackSize.y * ty) Then
        Exit Property
    End If
    .ptMaxTrackSize.y = dyMaxA / ty
    ' Maximized size can't be greater than maximimum size
    If .ptMaxTrackSize.y > .ptMaxSize.y Then .ptMaxSize.y = .ptMaxTrackSize.y
End With
End Property

Property Get MaximizedWidth() As Long
    MaximizedWidth = mmi.ptMaxSize.x * tx
End Property
Property Let MaximizedWidth(ByVal dxMaximizedA As Long)
With mmi
    ' Must be less than screen and greater than minimimum
    If dxMaximizedA > Screen.Width Or _
       dxMaximizedA < (.ptMinTrackSize.x * tx) Then Exit Property
    .ptMaxSize.x = dxMaximizedA / tx
    ' Maximized size can't be greater than maximimum size
    If .ptMaxSize.x > .ptMaxTrackSize.x Then .ptMaxTrackSize.x = .ptMaxSize.x
End With
End Property

Property Get MaximizedHeight() As Long
    MaximizedHeight = mmi.ptMaxSize.y * ty
End Property
Property Let MaximizedHeight(ByVal dyMaximizedA As Long)
With mmi
    ' Must be less than screen and greater than minimimum
    If dyMaximizedA > Screen.Height Or _
       dyMaximizedA < (.ptMinTrackSize.y * ty) Then Exit Property
    .ptMaxSize.y = dyMaximizedA / ty
    ' Maximized size can't be greater than maximimum size
    If .ptMaxSize.y > .ptMaxTrackSize.y Then .ptMaxTrackSize.y = .ptMaxSize.y
End With
End Property

Property Get MaximizedLeft() As Long
    MaximizedLeft = mmi.ptMaxPosition.x * tx
End Property
Property Let MaximizedLeft(ByVal xMaximizedA As Long)
    ' Must be positive and less than screen (but we won't enforce on
    ' screen because we don't know property assignment order)
    If xMaximizedA < 0 Or xMaximizedA >= Screen.Width Then Exit Property
    mmi.ptMaxPosition.x = xMaximizedA / tx
End Property

Property Get MaximizedTop() As Long
    MaximizedTop = mmi.ptMaxPosition.y * ty
End Property
Property Let MaximizedTop(ByVal yMaximizedA As Long)
    ' Must be positive and less than screen (but we won't enforce on
    ' screen because we don't know property assignment order)
    If yMaximizedA < 0 Or yMaximizedA >= Screen.Height Then Exit Property
    mmi.ptMaxPosition.y = yMaximizedA / ty
End Property
'

#If fComponent = 0 Then
Private Sub ErrRaise(e As Long)
    Dim sText As String, sSource As String
    If e > 1000 Then
        sSource = App.ExeName & ".MinMax"
        Select Case e
        Case eeBaseMinMax
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

