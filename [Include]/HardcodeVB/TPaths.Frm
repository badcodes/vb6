VERSION 5.00
Begin VB.Form FTestPaths 
   AutoRedraw      =   -1  'True
   Caption         =   "Test Paths"
   ClientHeight    =   3168
   ClientLeft      =   2172
   ClientTop       =   1752
   ClientWidth     =   5856
   DrawWidth       =   2
   FillColor       =   &H00000080&
   FillStyle       =   4  'Upward Diagonal
   BeginProperty Font 
      Name            =   "Lucida Console"
      Size            =   8.4
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "TPaths.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3168
   ScaleWidth      =   5856
   ShowInTaskbar   =   0   'False
End
Attribute VB_Name = "FTestPaths"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Show
    ' Start recording path
    BeginPath hDC
    ' Do some draw operations
    ForeColor = vbYellow
    FillColor = vbRed
    FillStyle = vbSolid
    Line (ScaleWidth * 0.2, ScaleHeight * 0.3)- _
         (ScaleWidth * 0.8, ScaleHeight * 0.7)
    Circle (ScaleWidth * 0.5, ScaleHeight * 0.5), _
           ScaleHeight * 0.3
    ' Stop recording path
    EndPath hDC
    ' Do something to finished path
    StrokeAndFillPath hDC
    ' Update device context
    Refresh
    
    BeginPath hDC
    Circle (1000, 1000), 500
    EndPath hDC
    'WidenPath hDC
    FlattenPath hDC
    'StrokePath hDC
    FillPath hDC
    'StrokeAndFillPath hDC
    Refresh
    
    ' Set a large, TrueType font
    With Font
        .Name = "Lucida Console"
        .Bold = True
        .Italic = True
        .Size = 48
    End With
    ForeColor = vbGreen
    FillColor = vbMagenta
    FillStyle = vbDiagonalCross
    ' Make Print statement into a path, and then fill
    BeginPath hDC
    Print "Hello"
    EndPath hDC
    StrokeAndFillPath hDC
    Refresh
    
End Sub
