VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GGDITool"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum EErrorGDITool
    eeBaseGDITool = 13510   ' GDITool
End Enum

Function VBPolygon(ByVal hDC As Long, aPoint() As Long) As Boolean
    Dim apt() As POINTL, i As Long, iMin As Long, c As Long
    iMin = LBound(aPoint)
    c = UBound(aPoint) - iMin + 1
    BugAssert 0 = (c Mod 2)     ' Even number of elements
    c = c / 2
    ' Create array of pixel-adjusted points
    ReDim apt(0 To c - 1) As POINTL
    Do While i < c
        apt(i).x = aPoint(iMin) / Screen.TwipsPerPixelX
        iMin = iMin + 1
        apt(i).y = aPoint(iMin) / Screen.TwipsPerPixelY
        iMin = iMin + 1
        i = i + 1
    Loop
    ' Pass first element and count to Polygon
    VBPolygon = Polygon(hDC, apt(0), c)
End Function

Function VBFloodFill(ByVal hDC As Long, ByVal x As Long, _
                     ByVal y As Long, ByVal clr As Long) As Boolean
    VBFloodFill = FloodFill(hDC, x / Screen.TwipsPerPixelX, _
                                 y / Screen.TwipsPerPixelY, clr)
End Function

' Create combined ROP for MaskBlt
Function MakeRop4(ropFore As Long, ropBack As Long) As Long
    ' MakeRop4 = ((ropBack SHL 8) And &HFF000000) Or ropFore
#If 1 Then
    MakeRop4 = (MBytes.LShiftDWord(ropBack, 8) And &HFF000000) Or ropFore
#Else
    ' Hack to do same shift in Basic
    If ropBack And &H800000 Then
        Dim ropTmp As Long
        ' Remove high bit
        ropTmp = (ropBack And &HFF7FFFFF)
        ' Do calculation
        ropTmp = ((ropTmp * 256) And &HFF000000) Or ropFore
        ' Put high bit back in
        MakeRop4 = ropTmp Or &H80000000
    Else
        MakeRop4 = ((ropBack * 256) And &HFF000000) Or ropFore
    End If
#End If
End Function

#If fComponent = 0 Then
Private Sub ErrRaise(e As Long)
    Dim sText As String, sSource As String
    If e > 1000 Then
        sSource = App.ExeName & ".GDITool"
        Select Case e
        Case eeBaseGDITool
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

