VERSION 5.00
Begin VB.UserControl Colorful 
   BackColor       =   &H000080FF&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   EditAtDesignTime=   -1  'True
   PaletteMode     =   0  'Halftone
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Begin VB.Shape Shape1 
      BorderWidth     =   5
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   1575
      Left            =   0
      Shape           =   2  'Oval
      Top             =   0
      Width           =   1575
   End
End
Attribute VB_Name = "Colorful"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Property Get Color() As OLE_COLOR
  Color = Shape1.FillColor
End Property

Public Property Let Color(ByVal c As OLE_COLOR)
  Shape1.FillColor = c
End Property

Private Sub UserControl_Click()
  Shape1.FillColor = &HFF0000
End Sub


Private Sub UserControl_EnterFocus()
  Shape1.BorderColor = &HFFFFFF
End Sub

Private Sub UserControl_ExitFocus()
  Shape1.BorderColor = &H0
End Sub

Private Sub UserControl_Resize()
  d = Shape1.BorderWidth / 2
  Shape1.Left = d
  Shape1.Top = d
  Shape1.Width = ScaleWidth - d * 2
  Shape1.Height = ScaleHeight - d * 2
End Sub
