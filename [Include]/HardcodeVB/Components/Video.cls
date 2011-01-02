VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CVideo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum EErrorVideo
    eeBaseVideo = 13420     ' CVideo
End Enum

Private hdcScreen As Long

Private Sub Class_Initialize()
    hdcScreen = GetDC(hNull)
End Sub

Private Sub Class_Terminate()
    Call ReleaseDC(hNull, hdcScreen)
End Sub

Private Function GetCaps(iCode As Long) As Long
    GetCaps = GetDeviceCaps(hdcScreen, iCode)
End Function

Property Get BitsPerPixel() As Long
    BitsPerPixel = GetCaps(BITSPIXEL)
End Property

Property Get ColorPlanes() As Long
    ColorPlanes = GetCaps(PLANES)
End Property

Property Get Technology() As Long
    ' Windows constant TECHNOLOGY conflicts with name of property
    Technology = GetCaps(Win.Technology)
End Property

Property Get XPixels() As Long
    XPixels = GetCaps(HORZRES)
End Property

Property Get YPixels() As Long
    YPixels = GetCaps(VERTRES)
End Property

Property Get BrushCount() As Long
    BrushCount = GetCaps(NUMBRUSHES)
End Property

Property Get PenCount() As Long
    PenCount = GetCaps(NUMPENS)
End Property

Property Get FontCount() As Long
    FontCount = GetCaps(NUMFONTS)
End Property

Property Get ColorCount() As Long
    ColorCount = GetCaps(NUMCOLORS)
End Property

Property Get XAspect() As Long
    XAspect = GetCaps(ASPECTX)
End Property

Property Get YAspect() As Long
    YAspect = GetCaps(ASPECTY)
End Property

Property Get XYAspect() As Long
    XYAspect = GetCaps(ASPECTXY)
End Property

Property Get PaletteSize() As Long
    PaletteSize = 0&
    If RasterCapability And RC_PALETTE Then
        PaletteSize = GetCaps(SIZEPALETTE)
    End If
End Property

Property Get RasterCapability() As Long
    RasterCapability = GetCaps(RASTERCAPS)
End Property

Property Get CurveCapability() As Long
    CurveCapability = GetCaps(CURVECAPS)
End Property

Property Get LineCapability() As Long
    LineCapability = GetCaps(LINECAPS)
End Property

Property Get PolygonCapability() As Long
    PolygonCapability = GetCaps(POLYGONALCAPS)
End Property

Property Get TextCapability() As Long
    TextCapability = GetCaps(TEXTCAPS)
End Property

Property Get TransparentBlt() As Boolean
    TransparentBlt = GetCaps(CAPS1) And C1_TRANSPARENT
End Property

#If fComponent = 0 Then
Private Sub ErrRaise(e As Long)
    Dim sText As String, sSource As String
    If e > 1000 Then
        sSource = App.ExeName & ".Video"
        Select Case e
        Case eeBaseVideo
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

