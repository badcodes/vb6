VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GDraw"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum EErrorDraw
    eeBaseDraw = 13460  ' Draw
End Enum

Const PI = 3.1415

Sub BmpSpiral(cvsDst As Object, picSrc As Picture)
With cvsDst
    ' Calculate sizes
    Dim dxSrc As Long, dySrc As Long, dxDst As Long, dyDst As Long
    dxSrc = .ScaleX(picSrc.Width): dySrc = .ScaleY(picSrc.Height)
    dxDst = .ScaleWidth: dyDst = .ScaleHeight
    ' Set defaults (play with these numbers for different effects)
    Dim xInc As Long, yInc As Long, xSize As Long, ySize As Long
    Dim x As Long, y As Long
    xInc = CInt(dxSrc * 0.01): yInc = CInt(dySrc * 0.01)
    xSize = CInt(dxSrc * 0.1): ySize = CInt(dySrc * 0.1)
    Dim radCur As Single, degCur As Integer, angInc As Integer
    degCur = 0: angInc = 55
    ' Start in center
    x = (dxDst \ 2) - (dxSrc \ 2): y = (dyDst \ 2) - (dySrc \ 2)
    
    ' Spiral until off destination
    Do
        ' Draw at current position
        .PaintPicture picSrc, x, y, , , , , , , vbSrcAnd
        ' Calculate angle in radians
        radCur = (degCur - 90) * (PI / 180)
        ' Calculate next x and y
        x = x + (xSize * Cos(radCur))
        y = y + (ySize * Sin(radCur))
        ' Widen spiral
        xSize = xSize + xInc: ySize = ySize + yInc + 1
        ' Turn angle
        degCur = (degCur + angInc) Mod 360
    Loop While (x > 0) And (x + dxSrc < dxDst - dxSrc) And _
               (y > 0) And (y + dySrc < dyDst)
End With
End Sub

Sub SpiralBmp(cvsDst As Object, picSrc As Picture, _
              ByVal xOff As Long, ByVal yOff As Long)
With cvsDst
    Dim xLeft As Long, xRight As Long, yTop As Long, yBottom As Long
    Dim dxSrc As Long, dySrc As Long, xSrc As Long, ySrc As Long
    Dim xDst As Long, yDst As Long, xInc As Long, yInc As Long
    Dim x As Long, y As Long
    ' Initialize
    dxSrc = .ScaleX(picSrc.Width): dySrc = .ScaleY(picSrc.Height)
    xInc = dxSrc / 20: yInc = dySrc / 20
    xLeft = 0: yTop = 0:
    xRight = dxSrc - xInc: yBottom = dySrc - yInc

    ' Draw each side
    Do While (xLeft <= xRight) And (yTop <= yBottom)
        ' Top
        For x = xLeft To xRight Step xInc
            .PaintPicture picSrc, x + xOff, y + yOff, xInc, yInc, _
                x, y, xInc, yInc, vbSrcCopy
        Next
        x = x - xInc: yTop = yTop + yInc
        ' Right
        For y = yTop To yBottom Step yInc
            .PaintPicture picSrc, x + xOff, y + yOff, xInc, yInc, _
                x, y, xInc, yInc, vbSrcCopy
        Next
        y = y - yInc: xRight = x - xInc
        ' Bottom
        For x = xRight To xLeft Step -xInc
            .PaintPicture picSrc, x + xOff, y + yOff, xInc, yInc, _
                x, y, xInc, yInc, vbSrcCopy
        Next
        x = x + xInc: yBottom = y - yInc
        ' Left
        For y = yBottom To yTop Step -yInc
            .PaintPicture picSrc, x + xOff, y + yOff, xInc, yInc, _
                x, y, xInc, yInc, vbSrcCopy
        Next
        y = y + yInc: xLeft = xLeft + xInc
    Loop
End With
End Sub

Sub BmpTile(cvsDst As Object, picSrc As Picture)
With cvsDst
    ' Calculate sizes
    Dim dxSrc As Long, dySrc As Long, dxDst As Long, dyDst As Long
    dxSrc = .ScaleX(picSrc.Width): dySrc = .ScaleY(picSrc.Height)
    dxDst = .ScaleWidth: dyDst = .ScaleHeight
    ' Tile until off destination
    Dim x As Long, y As Long, fAutoRedraw As Boolean
    fAutoRedraw = .AutoRedraw
    .AutoRedraw = False
    Do While y < dyDst
        Do While x < dxDst
            ' Draw at current position
            .PaintPicture picSrc, x, y
            x = x + dxSrc
        Loop
        y = y + dySrc
        x = 0
    Loop
    .AutoRedraw = fAutoRedraw
End With
End Sub

Sub Star(cvsDst As Object, ByVal x As Long, ByVal y As Long, _
         ByVal dxyRadius As Long, clrBorder As Long, _
         Optional clrOut As Long = -1, Optional clrIn As Long = -1)
With cvsDst
    ' Handle optional arguments
    If clrOut = -1 Then clrOut = clrBorder
    If clrIn = -1 Then clrIn = clrOut
    
    ' Start is 144 degrees (converted to radians)
    Const radStar As Double = 144 * PI / 180
    
    ' Calculate each point
    Dim ptPoly(1 To 10) As Long, i As Long
    For i = 1 To 10 Step 2
        ptPoly(i) = x + (Cos((i \ 2 + 1) * radStar) * dxyRadius)
        ptPoly(i + 1) = y + (Sin((i \ 2 + 1) * radStar) * dxyRadius)
    Next
    
    ' Set colors and style for star
    .ForeColor = clrBorder    ' SetTextColor
    .FillColor = clrOut       ' CreateSolidBrush
    .FillStyle = vbSolid      ' More CreateSolidBrush
    
    Call MGDITool.VBPolygon(.hDC, ptPoly)
    
    ' Set color for center
    .FillColor = clrIn        ' CreateSolidBrush
    Call MGDITool.VBFloodFill(.hDC, x, y, .ForeColor)
End With
End Sub

Sub Fade(cvsDst As Object, _
         Optional Red As Boolean = False, _
         Optional Green As Boolean = False, _
         Optional Blue As Boolean = True, _
         Optional Vertical As Boolean = True, _
         Optional Horizontal As Boolean = False, _
         Optional LightToDark As Boolean = True)
With cvsDst
    ' Trap errors
    On Error Resume Next
    
    ' Save properties
    Dim fAutoRedraw As Boolean, ordDrawStyle As Integer
    Dim ordDrawMode As Integer, iDrawWidth As Integer
    Dim ordScaleMode As Integer
    Dim rScaleWidth As Single, rScaleHeight As Single
    fAutoRedraw = .AutoRedraw: iDrawWidth = .DrawWidth
    ordDrawStyle = .DrawStyle: ordDrawMode = .DrawMode
    rScaleWidth = .ScaleWidth: rScaleHeight = .ScaleHeight
    ordScaleMode = .ScaleMode
    ' Err set if object lacks one of previous properties
    If Err Then Exit Sub
    ' If you get here, object is OK (Printer lacks AutoRedraw)
    fAutoRedraw = .AutoRedraw
    
    ' Set properties required for fade
    .AutoRedraw = True
    .DrawWidth = 3              ' Must be greater than 1 for dithering
    .DrawStyle = vbInsideSolid  ' vbInvisible gives an interesting effect
    .DrawMode = vbCopyPen       ' Try vbXorPen or vbMaskNotPen
    .ScaleMode = vbPixels
    .ScaleWidth = 256 * 2: .ScaleHeight = 256 * 2
    
    Dim clr As Long, i As Integer, x As Integer, y As Integer
    Dim iRed As Integer, iGreen As Integer, iBlue As Integer
    For i = 0 To 255
        ' Set line color
        If LightToDark Then
            If Red Then iRed = 255 - i
            If Blue Then iBlue = 255 - i
            If Green Then iGreen = 255 - i
        Else
            If Red Then iRed = i
            If Blue Then iBlue = i
            If Green Then iGreen = i
        End If
        clr = RGB(iRed, iGreen, iBlue)
        ' Draw each line of fade
        If Vertical Then
            cvsDst.Line (0, y)-(.ScaleWidth, y + 2), clr, BF
            y = y + 2
        End If
        If Horizontal Then
            cvsDst.Line (x, 0)-(x + 2, .ScaleHeight), clr, BF
            x = x + 2
        End If
    Next
    ' Put things back the way you found them
    .AutoRedraw = fAutoRedraw: .DrawWidth = iDrawWidth
    .DrawStyle = ordDrawStyle: .DrawMode = ordDrawMode
    .ScaleMode = ordScaleMode
    .ScaleWidth = rScaleWidth: .ScaleHeight = rScaleHeight
End With
End Sub
'

#If fComponent = 0 Then
Private Sub ErrRaise(e As Long)
    Dim sText As String, sSource As String
    If e > 1000 Then
        sSource = App.ExeName & ".Draw"
        Select Case e
        Case eeBaseDraw
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

