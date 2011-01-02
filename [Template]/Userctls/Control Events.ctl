VERSION 5.00
Begin VB.UserControl Events 
   BackColor       =   &H80000001&
   ClientHeight    =   3180
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3855
   FillColor       =   &H80000005&
   ScaleHeight     =   3180
   ScaleWidth      =   3855
End
Attribute VB_Name = "Events"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
    Debug.Print "AccessKeyPress: " & KeyAscii
End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)
    If PropertyName = Null Or PropertyName = "" Then
        Debug.Print "AmbientChanged: all"
    Else
        Debug.Print "AmbientChanged: " & PropertyName
    End If
End Sub

Private Sub UserControl_AsyncReadComplete(PropertyValue As AsyncProperty)
    Debug.Print "AsyncReadComplete: " & PropertyValue.AsyncType & ", " & PropertyValue.PropertyName
End Sub

Private Sub UserControl_Click()
    Debug.Print "Click"
End Sub

Private Sub UserControl_DblClick()
    Debug.Print "DblClick"
End Sub

Private Sub UserControl_DragDrop(Source As Control, x As Single, y As Single)
    Debug.Print "DragDrop: " & Source.Name & ", " & x & ", " & y
End Sub

Private Sub UserControl_DragOver(Source As Control, x As Single, y As Single, State As Integer)
    Debug.Print "DragOver: " & Source.Name & ", " & x & ", " & y & ", " & State
End Sub

Private Sub UserControl_EnterFocus()
    Debug.Print "EnterFocus"
End Sub

Private Sub UserControl_ExitFocus()
    Debug.Print "ExitFocus"
End Sub

Private Sub UserControl_GotFocus()
    Debug.Print "GotFocus"
End Sub

Private Sub UserControl_Initialize()
    Debug.Print "Initialize"
End Sub

Private Sub UserControl_InitProperties()
    Debug.Print "InitProperties"
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    Debug.Print "KeyDown: " & KeyCode & ", " & Shift
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    Debug.Print "KeyPress: " & KeyAscii
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    Debug.Print "KeyUp: " & KeyCode & ", " & Shift
End Sub

Private Sub UserControl_LostFocus()
    Debug.Print "LostFocus"
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Debug.Print "MouseDown: " & Button & ", " & Shift & ", " & x & ", " & y
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Debug.Print "MouseMove: " & Button & ", " & Shift & ", " & x & ", " & y
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Debug.Print "MouseUp: " & Button & ", " & Shift & ", " & x & ", " & y
End Sub

Private Sub UserControl_OLECompleteDrag(Effect As Long)
    Debug.Print "OLECompleteDrag: " & Effect
End Sub

Private Sub UserControl_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    Debug.Print "OLEDragDrop: " & Effect & ", " & Button & ", " & x & ", " & y
End Sub

Private Sub UserControl_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
    Debug.Print "OLEDragOver: " & Effect & ", " & Button & ", " & x & ", " & y
End Sub

Private Sub UserControl_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
    Debug.Print "OLEGiveFeedback: " & Effect & ", " & DefaultCursors
End Sub

Private Sub UserControl_OLESetData(Data As DataObject, DataFormat As Integer)
    Debug.Print "OLESetData: " & DataFormat
End Sub

Private Sub UserControl_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
    Debug.Print "OLEStartDrag: " & AllowedEffects
End Sub

Private Sub UserControl_Paint()
    Debug.Print "Paint"
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Debug.Print "ReadProperties"
End Sub

Private Sub UserControl_Resize()
    Debug.Print "Resize: " & Width & ", " & Height
End Sub

Private Sub UserControl_Terminate()
    Debug.Print "Terminate"
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Debug.Print "WriteProperties"
End Sub
