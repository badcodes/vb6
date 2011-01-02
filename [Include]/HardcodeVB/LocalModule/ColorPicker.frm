VERSION 5.00
Begin VB.Form FColorPicker 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1128
   ClientLeft      =   4788
   ClientTop       =   6072
   ClientWidth     =   4260
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   94
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   355
   Begin VB.Image imgPalette 
      Height          =   636
      Left            =   672
      Picture         =   "ColorPicker.frx":0000
      Top             =   1932
      Width           =   3276
   End
End
Attribute VB_Name = "FColorPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private aColor() As OLE_COLOR
Private clrCur As OLE_COLOR, clrLast As OLE_COLOR
Private fWide As Boolean, fDragging As Boolean
Private ixCur As Long, iyCur As Long, ixMax As Long, iyMax As Long
Private ix As Long, iy As Long, ixStart As Long, iyStart As Long

Private Sub Form_Initialize()
    ' Defaults if no one assigns different
    Wide = False
    Color = vbWhite
End Sub

Private Sub Form_Load()
    ' Set the form width and height exactly
    clrLast = clrCur
End Sub

Private Sub Form_Resize()
    ' Set the form width and height exactly
    Move Left, Top, ScaleX((ixMax * 17) + 3, vbPixels, vbTwips), _
         ScaleY((iyMax * 17) + 3, vbPixels, vbTwips)
    Refresh
End Sub

Private Sub Form_Paint()
    Dim ix As Long, iy As Long
    ' Draw colors in their boxes
    FillStyle = vbSolid
    For ix = 1 To ixMax
        For iy = 1 To iyMax
            FillColor = aColor(ix, iy)
            Line (((ix - 1) * 17) + 1, _
                  ((iy - 1) * 17) + 1)-Step(15, 15), , B
        Next
    Next
    DrawSelection ixCur, iyCur, True
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, _
                           X As Single, Y As Single)
    DrawSelection ixCur, iyCur, False
    ' Calculate the current position
    ixCur = ((X + 1) \ 17) + 1
    iyCur = ((Y + 1) \ 17) + 1
    If ixCur > ixMax Then ixCur = ixMax
    If iyCur > iyMax Then iyCur = iyMax
    fDragging = True
    DrawSelection ixCur, iyCur, True
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, _
                           X As Single, Y As Single)
    ' Calculate the current position
    Dim ix As Long, iy As Long
    ix = ((X + 1) \ 17) + 1
    iy = ((Y + 1) \ 17) + 1
    If ix > ixMax Then ix = ixMax
    If iy > iyMax Then iy = iyMax
    If fDragging Then
        DrawSelection ixCur, iyCur, False
        ixCur = ix: iyCur = iy
        DrawSelection ixCur, iyCur, True
    End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, _
                         X As Single, Y As Single)
    If Button = 2 Then
        Color = clrLast
    Else
        clrCur = aColor(ixCur, iyCur)
        FillColor = clrCur
        fDragging = False
    End If
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Color = clrLast
        Unload Me
    End If
End Sub

Friend Property Get Color() As OLE_COLOR
    Color = clrCur
End Property

Friend Property Let Color(clrCurA As OLE_COLOR)
    Dim ix As Long, iy As Long
    For ix = 1 To ixMax
        For iy = 1 To iyMax
            If aColor(ix, iy) = clrCurA Then
                ixCur = ix: iyCur = iy
                clrCur = clrCurA
                If ixCur Then Form_Paint
                Exit Property
            End If
        Next
    Next
End Property

Friend Property Get Wide() As Boolean
    Wide = fWide
End Property

Friend Property Let Wide(fWideA As Boolean)
    Dim clr As OLE_COLOR
    fWide = fWideA
    If fWide Then
        ixMax = 16
        iyMax = 3
    Else
        ixMax = 8
        iyMax = 6
    End If
    clr = Color
    InitArray
    Color = clr
    Form_Resize
End Property

Sub DrawSelection(ByVal ix As Long, ByVal iy As Long, fSelect As Boolean)
    ' Box the selection
    If ix = 0 And iy = 0 Then Exit Sub
    Dim ordFillStyle As FillStyleConstants
    ordFillStyle = FillStyle
    FillStyle = vbTransparent
    FillColor = aColor(ix, iy)
    If fSelect Then
        Line (((ix - 1) * 17) + 1, _
              ((iy - 1) * 17) + 1)-Step(15, 15), vbBlack, B
        Line (((ix - 1) * 17), _
              ((iy - 1) * 17))-Step(16, 16), vbWhite, B
        Line (((ix - 1) * 17) + 1, _
              ((iy - 1) * 17) + 1)-Step(14, 14), vbBlack, B
    Else
        Line (((ix - 1) * 17), _
              ((iy - 1) * 17))-Step(16, 16), vbButtonFace, B
        Line (((ix - 1) * 17) + 1, _
              ((iy - 1) * 17) + 1)-Step(15, 15), , B
    End If
    FillStyle = ordFillStyle
End Sub

Sub InitArray()
    ReDim aColor(1 To ixMax, 1 To iyMax) As Long
    If fWide Then
        aColor(1, 1) = &HFFFFFF
        aColor(1, 2) = &HC0C0C0
        aColor(1, 3) = &H808080
        aColor(2, 1) = &HE0E0E0
        aColor(2, 2) = &H404040
        aColor(2, 3) = &H0
        aColor(3, 1) = &HC0C0FF
        aColor(3, 2) = &H8080FF
        aColor(3, 3) = &HFF&
        aColor(4, 1) = &HC0E0FF
        aColor(4, 2) = &H80C0FF
        aColor(4, 3) = &H80FF&
        aColor(5, 1) = &HC0FFFF
        aColor(5, 2) = &H80FFFF
        aColor(5, 3) = &HFFFF&
        aColor(6, 1) = &HC0FFC0
        aColor(6, 2) = &H80FF80
        aColor(6, 3) = &HFF00&
        aColor(7, 1) = &HFFFFC0
        aColor(7, 2) = &HFFFF80
        aColor(7, 3) = &HFFFF00
        aColor(8, 1) = &HFFC0C0
        aColor(8, 2) = &HFF8080
        aColor(8, 3) = &HFF0000
        aColor(9, 1) = &HFFC0FF
        aColor(9, 2) = &HFF80FF
        aColor(9, 3) = &HFF00FF
        aColor(10, 1) = &HC0&
        aColor(10, 2) = &H80&
        aColor(10, 3) = &H40&
        aColor(11, 1) = &H40C0&
        aColor(11, 2) = &H4080&
        aColor(11, 3) = &H404080
        aColor(12, 1) = &HC0C0&
        aColor(12, 2) = &H8080&
        aColor(12, 3) = &H4040&
        aColor(13, 1) = &HC000&
        aColor(13, 2) = &H8000&
        aColor(13, 3) = &H4000&
        aColor(14, 1) = &HC0C000
        aColor(14, 2) = &H808000
        aColor(14, 3) = &H404000
        aColor(15, 1) = &HC00000
        aColor(15, 2) = &H800000
        aColor(15, 3) = &H400000
        aColor(16, 1) = &HC000C0
        aColor(16, 2) = &H800080
        aColor(16, 3) = &H400040
    Else
        ' Initialize color array
        aColor(1, 1) = &HFFFFFF
        aColor(1, 2) = &HE0E0E0
        aColor(1, 3) = &HC0C0C0
        aColor(1, 4) = &H808080
        aColor(1, 5) = &H404040
        aColor(1, 6) = &H0&
        aColor(2, 1) = &HC0C0FF
        aColor(2, 2) = &H8080FF
        aColor(2, 3) = &HFF&
        aColor(2, 4) = &HC0&
        aColor(2, 5) = &H80
        aColor(2, 6) = &H40
        aColor(3, 1) = &HC0E0FF
        aColor(3, 2) = &H80C0FF
        aColor(3, 3) = &H80FF&
        aColor(3, 4) = &H40C0&
        aColor(3, 5) = &H4080&
        aColor(3, 6) = &H404080
        aColor(4, 1) = &HC0FFFF
        aColor(4, 2) = &H80FFFF
        aColor(4, 3) = &HFFFF&
        aColor(4, 4) = &HC0C0&
        aColor(4, 5) = &H8080&
        aColor(4, 6) = &H4040&
        aColor(5, 1) = &HC0FFC0
        aColor(5, 2) = &H80FF80
        aColor(5, 3) = &HFF00&
        aColor(5, 4) = &HC000&
        aColor(5, 5) = &H8000&
        aColor(5, 6) = &H4000&
        aColor(6, 1) = &HFFFFC0
        aColor(6, 2) = &HFFFF80
        aColor(6, 3) = &HFFFF00
        aColor(6, 4) = &HC0C000
        aColor(6, 5) = &H808000
        aColor(6, 6) = &H404000
        aColor(7, 1) = &HFFC0C0
        aColor(7, 2) = &HFF8080
        aColor(7, 3) = &HFF0000
        aColor(7, 4) = &HC00000
        aColor(7, 5) = &H800000
        aColor(7, 6) = &H400000
        aColor(8, 1) = &HFFC0FF
        aColor(8, 2) = &HFF80FF
        aColor(8, 3) = &HFF00FF
        aColor(8, 4) = &HC000C0
        aColor(8, 5) = &H800080
        aColor(8, 6) = &H400040
    End If
End Sub





