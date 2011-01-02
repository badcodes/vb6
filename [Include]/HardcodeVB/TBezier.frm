VERSION 5.00
Begin VB.Form FTestBezier 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   Caption         =   "Bezier Curves"
   ClientHeight    =   5388
   ClientLeft      =   1092
   ClientTop       =   1512
   ClientWidth     =   4416
   DrawStyle       =   2  'Dot
   Icon            =   "TBEZIER.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5388
   ScaleWidth      =   4416
End
Attribute VB_Name = "FTestBezier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private apt(0 To 3) As POINTL

Private Sub Form_Load()
    Show
    InitBezier ScaleWidth, ScaleHeight
    DrawBezier
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button Then MoveBezier Button, x, y
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    MoveBezier Button, x, y
End Sub


Sub InitBezier(cxClient As Long, cyClient As Long)
    apt(0).x = ScaleX(cxClient / 2, vbTwips, vbPixels)
    apt(0).y = ScaleY(cyClient / 10, vbTwips, vbPixels)
    apt(1).x = ScaleX(cxClient / 4, vbTwips, vbPixels)
    apt(1).y = ScaleY(cyClient / 2, vbTwips, vbPixels)
    apt(2).x = ScaleX(3 * cxClient / 4, vbTwips, vbPixels)
    apt(2).y = ScaleY(cyClient / 2, vbTwips, vbPixels)
    apt(3).x = ScaleX(cxClient / 2, vbTwips, vbPixels)
    apt(3).y = ScaleY(9 * cyClient / 10, vbTwips, vbPixels)
    ForeColor = vbRed
End Sub

Sub DrawBezier()
    DrawStyle = vbSolid
    PolyBezier hDC, apt(0), 4
    DrawStyle = vbDot
    MoveTo hDC, apt(0).x, apt(0).y
    LineTo hDC, apt(1).x, apt(1).y
    MoveTo hDC, apt(2).x, apt(2).y
    LineTo hDC, apt(3).x, apt(3).y
    ' This line required in VB6 because of change in painting model
    Refresh
End Sub

Sub MoveBezier(ordButton As Integer, cx As Single, cy As Single)
    ForeColor = BackColor
    DrawBezier
    If ordButton = vbLeftButton Then
        apt(1).x = ScaleX(cx, vbTwips, vbPixels)
        apt(1).y = ScaleY(cy, vbTwips, vbPixels)
    End If
    If ordButton = vbRightButton Then
        apt(2).x = ScaleX(cx, vbTwips, vbPixels)
        apt(2).y = ScaleY(cy, vbTwips, vbPixels)
    End If
    ForeColor = vbRed
    DrawBezier
End Sub
