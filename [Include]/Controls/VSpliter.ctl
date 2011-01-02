VERSION 5.00
Begin VB.UserControl VSpliter 
   BackStyle       =   0  'Transparent
   ClientHeight    =   3405
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   135
   FillStyle       =   0  'Solid
   MousePointer    =   9  'Size W E
   ScaleHeight     =   3405
   ScaleWidth      =   135
   Begin VB.Shape VSplit 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H00E0E0E0&
      FillStyle       =   7  'Diagonal Cross
      Height          =   3420
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   105
   End
End
Attribute VB_Name = "VSpliter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mMouseOnSpliter As Boolean
Event Moving(X As Single, Y As Single)
Event MouseDown()
Event MouseMove()
Event MouseUp()

Private Sub UserControl_LostFocus()
    RaiseEvent MouseUp
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mMouseOnSpliter = True
        RaiseEvent MouseDown
End Sub


Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If mMouseOnSpliter Then RaiseEvent Moving(X + UserControl.Extender.Left, Y + UserControl.Extender.Top): Exit Sub


            RaiseEvent MouseMove
       
        
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mMouseOnSpliter = False
    RaiseEvent MouseUp
End Sub

Private Sub UserControl_Resize()
    VSplit.Move 0, 0, UserControl.Width, UserControl.Height
End Sub
