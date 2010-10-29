VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CButterFly"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IAnimation

Private fRunning As Boolean
Private objCanvas As Object

Public Enum eeErrorButterFly
    eeBaseButterFly = 13010     ' CButterfly
End Enum

Private Sub Class_Terminate()
    Set objCanvas = Nothing
End Sub

Private Property Let IAnimation_Running(fRunningA As Boolean)
    fRunning = fRunningA
    If fRunning Then Draw
End Property
Private Property Get IAnimation_Running() As Boolean
    IAnimation_Running = fRunning
End Property

Private Property Get IAnimation_Canvas() As Object
    Set IAnimation_Canvas = objCanvas
End Property

Private Property Set IAnimation_Canvas(objCanvasA As Object)
    Set objCanvas = objCanvasA
End Property

Private Sub Draw()
    
    fRunning = True
    Const PI = 3.1415

    Dim x As Single, y As Single, theta As Single
    Dim xx As Single, yy As Single, R As Single
    Dim mulxy As Single, addxy As Single
    Dim powr As Single, divt As Single
    Dim mult As Single, mulc As Single
    Dim incr As Single, muloop As Single
    Dim f As Boolean, clr As Long
      
    ' Initialize variables
    mulxy = 24#     ' Controls size
    powr = 5#       ' Controls some aspects of shape
    divt = 40#      ' Controls density
    mult = 4#       ' Helps make butterfly shape
    mulc = 2#       ' Helps make butterfly shape
    incr = 0.05     ' Controls roundness
    muloop = 50#    ' Controls iterations and density

    ' Make it square
    If objCanvas.Width < objCanvas.Height Then
        objCanvas.Width = objCanvas.Height
    Else
        objCanvas.Height = objCanvas.Width
    End If
    ' Set coordinate system
    objCanvas.ScaleLeft = -100#
    objCanvas.ScaleTop = -110#
    objCanvas.ScaleWidth = 200#
    objCanvas.ScaleHeight = 200#

    Do While fRunning
        ' Draw in random color or erase with background
        f = Not f
        If f Then
            clr = QBColor(MRandom.Random(1, 15))
        Else
            clr = objCanvas.BackColor
        End If
        ' Draw shape
        For theta = 0# To muloop * PI Step incr
            
            R = Exp(Cos(theta)) - _
                mulc * Cos(mult * theta) + _
                Sin(theta / divt) ^ powr
            x = R * Sin(theta)
            y = R * Cos(theta)
            xx = -((x * mulxy) + addxy)
            yy = -((y * mulxy) + addxy)
            If theta = 0# Then
               ' Move to center without drawing
               objCanvas.CurrentX = xx
               objCanvas.CurrentY = yy
            Else
                objCanvas.Line -(xx, yy), clr
            End If
            theta = theta + incr
            DoEvents
            If Not fRunning Then Exit Sub
        Next
    Loop
        
End Sub

#If fComponent = 0 Then
Private Sub ErrRaise(e As Long)
    Dim sText As String, sSource As String
    If e > 1000 Then
        sSource = App.ExeName & ".ButterFly"
        Select Case e
        Case eeBaseButterFly
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


