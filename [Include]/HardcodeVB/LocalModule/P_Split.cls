VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSplitter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Thanks to Elliott Whitticar and Vadim Katsman who fixed bugs and suggested
' improvements to my CHSplitter and CVSplitter classes. I integrated some of
' their code, but didn't solve all possible splitter problems. The primary
' improvement (implemented separately by both Elliot and Vadim) is combining
' vertical and horizontal splitters into a single class. Minh Nguyen identified
' the problem with cursors not working in the private version of the class.
' Jon Evans suggested provided code for several bug fixes and new features.
' The ShowDrag parameter and having AutoBorder calculate the bottom edge are
' his ideas.

' Thanks to all.

' Internal variables for forms and controls
Private ctlNW As Object             ' Left/Top control
Private ctlSE As Object             ' Right/Bottom control
Private objContainer As Object

' Sizes of borders and pixels
Private xySplit As Single           ' Position of splitter bar
Private dxySplit As Single          ' Width/height of splitter bar in scale units
Private dxyMin As Single            ' Minimum control width/height
Private xPixel As Single
Private yPixel As Single
Private dxBorder As Single
Private dyBorder As Single
Private dyBottom As Single          ' Bottom border (useful if form has status bar)
Private cBorderPixels As Long

' Flags
Private fResize As Boolean          ' True => move left and right control
Private fAutoBorder As Boolean
Private fDragging As Boolean
Private fCreated As Boolean
Private fVertical As Boolean        ' True => Vertical, False => Horizontal
Private iPercent As Integer         ' 1 to 99 gives initial split Percent
Private fShowDrag As Boolean        ' True => Move NW and SE interactively

' Old mouse pointer, draw style, and draw mode
Private mpOld As Integer
Private dsOld As Integer
Private dmOld As Integer

Private mpResize As Integer         ' MousePointer to use when resizing

' AutoRedraw
Private arOld As Boolean

Public Enum EErrorSplitter
    eeBaseSplitter = 13690  ' CSplitter
    eeInvalidControl        ' Invalid controls or container
    eeSplitNotCreated       ' Create splitter before using members
End Enum

#If fComponent = 0 Then

' Arrays for vertical and horizontal cursors used only for private class
' (VBCore uses cursors in resources)
Private abXorHorz(0 To 127) As Byte
Private abAndHorz(0 To 127) As Byte

Private abXorVert(0 To 127) As Byte
Private abAndVert(0 To 127) As Byte

Private hcurVert As Long, hcurHorz As Long
Private picVert As StdPicture, picHorz As StdPicture

Sub InitVert()
    ' This weird code initialization wouldn't be necessary if VB
    ' had data initialization
    If abAndVert(0) = &HFF Then Exit Sub    ' Already initialized
    ' Initialize And mask to &HFF (Xor mask is already &H00)
    Dim i As Long
    For i = 0 To 127
        abAndVert(i) = &HFF
    Next
    ' Insert non-zero Xor values, non-FF And values
    abXorVert(37) = &H7: abAndVert(37) = &HF8
    abXorVert(38) = &HE0: abAndVert(38) = &H1F
    abXorVert(41) = &H5: abAndVert(41) = &HF8
    abXorVert(42) = &HA0: abAndVert(42) = &H1F
    abXorVert(45) = &H5: abAndVert(45) = &HF8
    abXorVert(46) = &HA0: abAndVert(46) = &H1F
    abXorVert(49) = &H5: abAndVert(49) = &HF8
    abXorVert(50) = &HA0: abAndVert(50) = &H1F
    abXorVert(53) = &H25: abAndVert(53) = &HD8
    abXorVert(54) = &HA4: abAndVert(54) = &H1B
    abXorVert(57) = &H55: abAndVert(57) = &H88
    abXorVert(58) = &HAA: abAndVert(58) = &H11
    abXorVert(61) = &H9D: abAndVert(61) = &H0
    abXorVert(62) = &HB9: abAndVert(62) = &H0
    abXorVert(64) = &H1: abAndVert(64) = &HFE
    abXorVert(65) = &H1: abAndVert(65) = &H0
    abXorVert(66) = &H80: abAndVert(66) = &H0
    abXorVert(67) = &H80: abAndVert(67) = &H7F
    abXorVert(69) = &H9D: abAndVert(69) = &H0
    abXorVert(70) = &HB9: abAndVert(70) = &H0
    abXorVert(73) = &H55: abAndVert(73) = &H88
    abXorVert(74) = &HAA: abAndVert(74) = &H11
    abXorVert(77) = &H25: abAndVert(77) = &HD8
    abXorVert(78) = &HA4: abAndVert(78) = &H1B
    abXorVert(81) = &H5: abAndVert(81) = &HF8
    abXorVert(82) = &HA0: abAndVert(82) = &H1F
    abXorVert(85) = &H5: abAndVert(85) = &HF8
    abXorVert(86) = &HA0: abAndVert(86) = &H1F
    abXorVert(89) = &H5: abAndVert(89) = &HF8
    abXorVert(90) = &HA0: abAndVert(90) = &H1F
    abXorVert(93) = &H7: abAndVert(93) = &HF8
    abXorVert(94) = &HE0: abAndVert(94) = &H1F
    
End Sub

Sub InitHorz()
    ' Already initialized
    If abAndHorz(0) = &HFF Then Exit Sub
    ' Initialize And mask to &HFF (Xor mask is already &H00)
    Dim i As Long
    For i = 0 To 127
        abAndHorz(i) = &HFF
    Next
    ' Insert non-zero Xor values, non-FF And values
    abXorHorz(37) = &H1:        abAndHorz(37) = &HFE
    abXorHorz(38) = &H0:        abAndHorz(38) = &HFF
    abXorHorz(39) = &H0:        abAndHorz(39) = &HFF
    abXorHorz(40) = &H0:        abAndHorz(40) = &HFF
    abXorHorz(41) = &H2:        abAndHorz(41) = &HFC
    abXorHorz(42) = &H80:       abAndHorz(42) = &H7F
    abXorHorz(43) = &H0:        abAndHorz(43) = &HFF
    abXorHorz(44) = &H0:        abAndHorz(44) = &HFF
    abXorHorz(45) = &H4:        abAndHorz(45) = &HF8
    abXorHorz(46) = &H40:       abAndHorz(46) = &H3F
    abXorHorz(47) = &H0:        abAndHorz(47) = &HFF
    abXorHorz(48) = &H0:        abAndHorz(48) = &HFF
    abXorHorz(49) = &H8:        abAndHorz(49) = &HF0
    abXorHorz(50) = &H20:       abAndHorz(50) = &H1F
    abXorHorz(51) = &H0:        abAndHorz(51) = &HFF
    abXorHorz(52) = &H0:        abAndHorz(52) = &HFF
    abXorHorz(53) = &H6:        abAndHorz(53) = &HF8
    abXorHorz(54) = &HC0:       abAndHorz(54) = &H3F
    abXorHorz(55) = &H0:        abAndHorz(55) = &HFF
    abXorHorz(56) = &H1:        abAndHorz(56) = &HFE
    abXorHorz(57) = &HFE:       abAndHorz(57) = &H0
    abXorHorz(58) = &HFF:       abAndHorz(58) = &H0
    abXorHorz(59) = &H0:        abAndHorz(59) = &HFF
    abXorHorz(60) = &H1:        abAndHorz(60) = &HFE
    abXorHorz(61) = &H0:        abAndHorz(61) = &H0
    abXorHorz(62) = &H1:        abAndHorz(62) = &H0
    abXorHorz(63) = &H0:        abAndHorz(63) = &HFF
    abXorHorz(64) = &H1:        abAndHorz(64) = &HFE
    abXorHorz(65) = &HFF:       abAndHorz(65) = &H0
    abXorHorz(66) = &HFF:       abAndHorz(66) = &H0
    abXorHorz(67) = &H0:        abAndHorz(67) = &HFF
    abXorHorz(68) = &H1:        abAndHorz(68) = &HFE
    abXorHorz(69) = &HFF:       abAndHorz(69) = &H0
    abXorHorz(70) = &HFF:       abAndHorz(70) = &H0
    abXorHorz(71) = &H0:        abAndHorz(71) = &HFF
    abXorHorz(72) = &H1:        abAndHorz(72) = &HFE
    abXorHorz(73) = &H0:        abAndHorz(73) = &H0
    abXorHorz(74) = &H1:        abAndHorz(74) = &H0
    abXorHorz(75) = &H0:        abAndHorz(75) = &HFF
    abXorHorz(76) = &H1:        abAndHorz(76) = &HFE
    abXorHorz(77) = &HFE:       abAndHorz(77) = &H0
    abXorHorz(78) = &HFF:       abAndHorz(78) = &H0
    abXorHorz(79) = &H0:        abAndHorz(79) = &HFF
    abXorHorz(80) = &H0:        abAndHorz(80) = &HFF
    abXorHorz(81) = &H6:        abAndHorz(81) = &HF8
    abXorHorz(82) = &HC0:       abAndHorz(82) = &H3F
    abXorHorz(83) = &H0:        abAndHorz(83) = &HFF
    abXorHorz(84) = &H0:        abAndHorz(84) = &HFF
    abXorHorz(85) = &H8:        abAndHorz(85) = &HF0
    abXorHorz(86) = &H20:       abAndHorz(86) = &H1F
    abXorHorz(87) = &H0:        abAndHorz(87) = &HFF
    abXorHorz(88) = &H0:        abAndHorz(88) = &HFF
    abXorHorz(89) = &H4:        abAndHorz(89) = &HF8
    abXorHorz(90) = &H40:       abAndHorz(90) = &H3F
    abXorHorz(91) = &H0:        abAndHorz(91) = &HFF
    abXorHorz(92) = &H0:        abAndHorz(92) = &HFF
    abXorHorz(93) = &H2:        abAndHorz(93) = &HFC
    abXorHorz(94) = &H80:       abAndHorz(94) = &H7F
    abXorHorz(95) = &H0:        abAndHorz(95) = &HFF
    abXorHorz(96) = &H0:        abAndHorz(96) = &HFF
    abXorHorz(97) = &H1:        abAndHorz(97) = &HFE
End Sub

Private Sub Class_Initialize()
    ' Initialize cursor arrays (only in private class)
    InitVert
    InitHorz
End Sub

Private Sub Class_Terminate()
    If hcurVert Then
        Set picVert = Nothing
        DestroyCursor hcurVert
    End If
    If hcurHorz Then
        Set picHorz = Nothing
        DestroyCursor hcurHorz
    End If
End Sub

Function VerticalCursor() As StdPicture
    ' Create cursor from array, picture from cursor handle
    If hcurVert = 0 Then
        hcurVert = CreateCursor(App.hInstance, 16, 16, 32, 32, _
                                abAndVert(0), abXorVert(0))
        Set picVert = CursorToPicture(hcurVert)
    End If
    Set VerticalCursor = picVert
End Function

Function HorizontalCursor() As StdPicture
    ' Create cursor from array, picture from cursor handle
    If hcurHorz = 0 Then
        hcurHorz = CreateCursor(App.hInstance, 16, 16, 32, 32, _
                                abAndHorz(0), abXorHorz(0))
        Set picHorz = CursorToPicture(hcurHorz)
    End If
    Set HorizontalCursor = picHorz
End Function

Private Sub ErrRaise(e As Long)
    Dim sText As String, sSource As String
    If e > 1000 Then
        sSource = App.EXEName & ".CSplitter"
        Select Case e
        Case eeBaseSplitter
            BugAssert True
        Case eeInvalidControl
            sText = "Create: Invalid controls or container"
        Case eeSplitNotCreated
            sText = "Create splitter before using members"
        End Select
        Err.Raise COMError(e), sSource, sText
    Else
        ' Raise standard Visual Basic error
        Err.Raise e, sSource
    End If
End Sub
#End If

' Create a splitter window using the following paramaters:
'    Vertical - Vertical or horizontal
'    BorderPixels - Width of border (AutoBorder:=False only)
'    AutoBorder - Calculates border from initial top and left of NW pane
'                 and from bottom of SE pane
'    Resizeable - Recalculates size window resized (AutoBorder:=False only)
'    Percent - Initial relative position of splitter
'    Cursor - Custom cursor to be used as splitter
'    ShowDrag - Show dragging windows instead of splitter lines
'
' Design notes: The AutoBorder property was enhanced so that CSplitter could
' work well on forms that have toolbars and status bars. The old algorithm
' didn't provide a way to make the bottom border appear above the status bar.
' In the new algorithm you must specify the bottom border by the position of
' the bottom right (southeast) control. This might break some old client code.
' The BorderPixels parameter will not work well with forms that have a toolbar
' or status bar. Use AutoBorder for these forms. The new ShowDrag parameter
' usually causes too much flicker when used with controls that have a lot of
' text such as TextBoxes and ListBoxes. It can give an interesting effect with
' graphic controls such as PictureBoxes.

Sub Create(LeftControl As Object, RightControl As Object, _
           Vertical As Boolean, _
           Optional BorderPixels As Long = 0, _
           Optional AutoBorder As Boolean = True, _
           Optional Resizeable As Boolean = True, _
           Optional Percent As Integer = 50, _
           Optional Cursor As Picture, _
           Optional ShowDrag As Boolean = False)

    fCreated = False
    
    On Error GoTo CreateError
    
    ' Set internal controls
    Set ctlNW = LeftControl
    Set ctlSE = RightControl
    Set objContainer = ctlNW.Container
    BugAssert objContainer Is ctlSE.Container
    
    ' Splitters work--sort of--with ClipControls True, but the splitter line
    ' isn't drawn correctly, so we disallow it in debug version (go ahead
    ' and ignore in release if you really want)
    BugAssert objContainer.ClipControls = False
    
    ' Save flags
    fVertical = Vertical
    fShowDrag = ShowDrag
    fAutoBorder = AutoBorder
    fResize = Resizeable
    ' Can't resize if borders are auto
    If fAutoBorder Then fResize = False
    
    ' Handle split percent
    iPercent = Percent
    If iPercent > 99 Then iPercent = 99
    If iPercent < 1 Then iPercent = 1
    
    ' Set splitter size
    cBorderPixels = BorderPixels
    If cBorderPixels = 0 Then
        fAutoBorder = True
        cBorderPixels = 4
    End If
    
    With objContainer
        ' Size of one in pixel in current scale
        xPixel = .ScaleX(1, vbPixels, .ScaleMode)
        yPixel = .ScaleY(1, vbPixels, .ScaleMode)
    
        ' Set cursor
        If Cursor Is Nothing Then
#If fComponent Then
            If fVertical Then
                Set .MouseIcon = LoadResPicture("VSplit", vbResCursor)
            Else
                Set .MouseIcon = LoadResPicture("HSplit", vbResCursor)
            End If
#Else
            If fVertical Then
                Set .MouseIcon = VerticalCursor
            Else
                Set .MouseIcon = HorizontalCursor
            End If
#End If
        Else
            Set .MouseIcon = Cursor
        End If
        ' Get the .MousePointer value to use when resizing
        If .MouseIcon.Type = vbPicTypeIcon Then
            mpResize = vbCustom
        ElseIf fVertical Then
            mpResize = vbSizeWE
        Else
            mpResize = vbSizeNS
        End If
        
        ' Set border size
        If fAutoBorder Then
            dxBorder = ctlNW.Left
            dyBorder = ctlNW.Top
            dyBottom = objContainer.ScaleHeight - yBottom(ctlSE)
        Else
            dxBorder = cBorderPixels * xPixel
            dyBorder = cBorderPixels * yPixel
            dyBottom = cBorderPixels * yPixel
        End If

        ' Set the splitter bar and minimum width/height in scale units
        ' (Ideally we'd use control properties for minimum width/height)
        If fVertical Then
            dxySplit = cBorderPixels * xPixel
            dxyMin = 20 * xPixel + 2 * dyBorder
        Else
            dxySplit = cBorderPixels * yPixel
            dxyMin = 20 * yPixel + 2 * dyBorder
        End If
        SplitPercent = Percent

    End With
    
    fCreated = True
    Exit Sub
CreateError:
    ErrRaise eeInvalidControl
End Sub

Property Get Capture() As Boolean
    If Not fCreated Then ErrRaise eeSplitNotCreated
    
    ' See if the container form or control has captured mouse events
    Capture = (GetCapture = objContainer.hWnd)
End Property


Property Let Capture(fCapture As Boolean)
    If Not fCreated Then ErrRaise eeSplitNotCreated
    
    ' Capture or release mouse events
    If fCapture Then
        SetCapture objContainer.hWnd
    Else
        ReleaseCapture
    End If
End Property

Property Get ShowDrag() As Boolean
    ShowDrag = fShowDrag
End Property
Property Let ShowDrag(fShowDragA As Boolean)
    fShowDrag = fShowDragA
End Property

Private Sub Draw(ByVal xyDraw As Single, _
                 Optional xyNew As Single = -1#)
With objContainer
    ' Draw the splitter bar at the specified position
    
    ' The second argument is provided so MouseMove can invert the
    ' old drawing at xyDraw and draw a new splitter bar at xyNew

    If fShowDrag Then Exit Sub

    .DrawStyle = vbInsideSolid
    .DrawMode = vbInvert
    If fVertical Then
        ' Erase old line
        objContainer.Line (xyDraw - xPixel, ctlNW.Top)-(xyDraw + xPixel, yBottom(ctlNW)), , B
        If Not xyNew = -1 Then
            ' Draw new line
            xyDraw = xyNew
            objContainer.Line (xyDraw - xPixel, ctlNW.Top)-(xyDraw + xPixel, yBottom(ctlNW)), , B
        End If
    Else
        ' Erase old line
        objContainer.Line (ctlNW.Left, xyDraw - yPixel)-(xRight(ctlNW), xyDraw + yPixel), , B
        If Not xyNew = -1 Then
            ' Draw new line
            xyDraw = xyNew
            objContainer.Line (ctlNW.Left, xyDraw - yPixel)-(xRight(ctlNW), xyDraw + yPixel), , B
        End If
    End If
End With
End Sub

Public Property Get SplitterSize() As Long
    ' Width/Height of splitter bar in scale units
    SplitterSize = dxySplit
End Property

Public Property Get SplitPercent() As Long
With objContainer
    If fVertical Then
        SplitPercent = 100 / (.ScaleWidth / (xySplit - (dxySplit / 2)))
    Else
        SplitPercent = 100 / (.ScaleHeight / (xySplit - (dxySplit / 2)))
    End If
End With
End Property

Public Property Let SplitPercent(ByVal iPercentCur As Long)
With objContainer
    Dim iMinPercent As Long
    If fVertical Then
        iMinPercent = 100 / (.ScaleWidth / (dxyMin - (dxySplit / 2)))
        If iPercentCur < iMinPercent Then iPercentCur = iMinPercent
        If iPercentCur > 100 - iMinPercent Then iPercentCur = 100 - iMinPercent
        xySplit = ((iPercentCur / 100) * .ScaleWidth) - (dxySplit / 2)
        ctlNW.Move dxBorder, dyBorder, _
                   xySplit - dxBorder, _
                   .ScaleHeight - dyBorder - dyBottom
        ctlSE.Move xRight(ctlNW) + dxySplit, dyBorder, _
                   .ScaleWidth - ctlNW.Width - dxySplit - (dxBorder * 2), _
                   ctlNW.Height
    Else
        iMinPercent = 100 / ((.ScaleHeight - dyBottom) / (dxyMin - (dxySplit / 2)))
        If iPercentCur < iMinPercent Then iPercentCur = iMinPercent
        If iPercentCur > 100 - iMinPercent Then iPercentCur = 100 - iMinPercent
        dxySplit = cBorderPixels * yPixel
        dxyMin = 20 * yPixel + 2 * dyBorder
        xySplit = ((iPercentCur / 100) * (.ScaleHeight - dyBottom + dyBorder)) - (dxySplit / 2)
        ctlNW.Move dxBorder, dyBorder, _
                   .ScaleWidth - (dxBorder * 2), _
                   xySplit - dyBorder
        ctlSE.Move dxBorder, yBottom(ctlNW) + dxySplit, _
                   ctlNW.Width, _
                   .ScaleHeight - ctlNW.Height - dxySplit - dyBorder - dyBottom
    End If
End With
End Property

Private Sub ResizePanes()
With objContainer
    If fDragging And Not fShowDrag Then Exit Sub
    
    If fVertical Then
        Dim xNew As Long, dxNew As Long
        ctlNW.Width = xySplit - ctlNW.Left - (dxySplit / 2)
        xNew = xRight(ctlNW) + dxySplit
        dxNew = .ScaleWidth - xNew - dxBorder
        ctlSE.Move xNew, ctlSE.Top, dxNew, ctlSE.Height
    Else
        Dim yNew As Long, dyNew As Long
        ctlNW.Height = xySplit - ctlNW.Top - (dxySplit / 2)
        yNew = yBottom(ctlNW) + dxySplit
        dyNew = .ScaleTop + .ScaleHeight - yNew - dyBottom
        ctlSE.Move ctlSE.Left, yNew, ctlSE.Width, dyNew
    End If
End With
End Sub

Sub Resize()
With objContainer

    Dim rScaleFac As Single
    Dim xyStart As Single, dxyFarEdge As Single
    If fVertical Then
        rScaleFac = .ScaleWidth / (dxBorder + ctlNW.Width + dxySplit + _
                                   ctlSE.Width + dxBorder)
        ' Move everything in border size from the edge
        dxyFarEdge = .ScaleHeight - (dyBorder * 2)
        ctlNW.Move dxBorder, .ScaleTop + dyBorder, _
                   ctlNW.Width * rScaleFac, dxyFarEdge
    
        xyStart = xRight(ctlNW) + dxySplit
        ctlSE.Move xyStart, dyBorder, _
                   .ScaleWidth - xyStart - dxBorder, dxyFarEdge
    Else    ' Resize Horizontal Splitter

        rScaleFac = (.ScaleHeight - dyBottom) / (dyBorder + ctlNW.Height + dxySplit + _
                                    ctlSE.Height + dyBorder)

        ' Move everything in border size from the edge
        dxyFarEdge = .ScaleWidth - (dxBorder * 2)
        ctlNW.Move dxBorder, .ScaleTop + dyBorder, _
                   dxyFarEdge, ctlNW.Height * rScaleFac
    
        xyStart = yBottom(ctlNW) + dxySplit
        ctlSE.Move dxBorder, xyStart, dxyFarEdge, _
                   .ScaleHeight - xyStart ' - dyBottom
    End If
End With
End Sub

Sub Splitter_MouseMove(Button As Integer, Shift As Integer, _
                       x As Single, y As Single)
With objContainer
    If Not fCreated Then ErrRaise eeSplitNotCreated
    Dim xyNew As Single, xyMinPos As Single, xyMaxPos As Single
    Dim fZone As Boolean    ' Over Splitter Bar flag

    ' Change the cursor to splitter or back
    ' Are we in the container's client area?
    If x >= 0 And x <= .ScaleWidth Then
        If y >= .ScaleTop And _
            y <= .ScaleTop + .ScaleHeight Then
            ' Are we in the splitter bar zone?
            If fVertical Then
                If x < ctlSE.Left And x > xRight(ctlNW) Then fZone = True
            Else
                If y < ctlSE.Top And y > yBottom(ctlNW) Then fZone = True
            End If
        End If
    End If
    
    If fZone Then   ' We're over the splitter bar
        If .MousePointer <> mpResize Then
            mpOld = .MousePointer
            .MousePointer = mpResize
            Me.Capture = True
        End If
    ElseIf (.MousePointer = mpResize) And Not fDragging Then
        .MousePointer = mpOld
        Me.Capture = False
    End If
    
    ' Move the splitter line if within range
    If fDragging Then
        If fVertical Then
            xyNew = x
            xyMinPos = dxyMin
            xyMaxPos = .ScaleWidth - dxyMin
        Else    ' We're moving the horizontal line
            xyNew = y
            xyMinPos = .ScaleTop + dxyMin
            xyMaxPos = .ScaleTop + .ScaleHeight - dxyMin
        End If
        If (xySplit <> xyNew) Then
            If xyNew < xyMinPos Then
                xyNew = xyMinPos
            ElseIf xyNew > xyMaxPos Then
                xyNew = xyMaxPos
            End If
            ' Erase the old line at xySplit and draw the new line
            Draw xySplit, xyNew
            xySplit = xyNew
            ResizePanes
        End If
    End If
End With
End Sub

Sub Splitter_MouseDown(Button As Integer, Shift As Integer, _
                       x As Single, y As Single)
With objContainer
    If Not fCreated Then ErrRaise eeSplitNotCreated
    Dim fZone As Boolean
    
    If fVertical Then
        fZone = x > xRight(ctlNW) And x < ctlSE.Left
    Else
        fZone = y > yBottom(ctlNW) And y < ctlSE.Top
    End If
    
    ' If over splitter start a drag
    If fZone Then
        If Button = vbLeftButton Then
            ' Save and restore state
            fDragging = True
            dsOld = .DrawStyle
            dmOld = .DrawMode
            arOld = .AutoRedraw
            .AutoRedraw = False
            ' Determine splitter line position
            If fVertical Then
                xySplit = xRight(ctlNW) + (dxBorder / 3)
            Else
                xySplit = yBottom(ctlNW) + (dyBorder / 3)
            End If
            ' Draw the splitter line
            Draw xySplit
        End If
    Else
        If .MousePointer = mpResize Then .MousePointer = mpOld
    End If
End With
End Sub

Sub Splitter_MouseUp(Button As Integer, Shift As Integer, _
                     x As Single, y As Single)
With objContainer
    If Not fCreated Then ErrRaise eeSplitNotCreated
    
    If fDragging Then
        ' Erase old line
        Draw xySplit
        fDragging = False

        ResizePanes
        .DrawStyle = dsOld
        .DrawMode = dmOld
        .AutoRedraw = arOld
    End If
    
    ' Restore the pointer
    If .MousePointer = mpResize Then
         .MousePointer = mpOld
         Me.Capture = False
    End If
End With
End Sub
 
Sub Splitter_Resize()
    If objContainer Is Nothing Then Exit Sub
    If Not fCreated Then ErrRaise eeSplitNotCreated
    
    If fResize Then
        On Error Resume Next
        ' Only forms have WindowState
        If objContainer.WindowState <> vbMinimized Then Resize
    End If
End Sub

Private Function xRight(obj As Object) As Single
    xRight = obj.Left + obj.Width
End Function

Private Function yBottom(obj As Object) As Single
    yBottom = obj.Top + obj.Height
End Function




