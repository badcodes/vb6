VERSION 5.00
Begin VB.UserControl XPictureGlass 
   BackColor       =   &H80000014&
   BackStyle       =   0  'Transparent
   ClientHeight    =   1128
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1152
   ScaleHeight     =   1128
   ScaleWidth      =   1152
   ToolboxBitmap   =   "PicGlass.ctx":0000
End
Attribute VB_Name = "XPictureGlass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Enum EErrorPictureGlass
    eeBasePictureGlass = 13740  ' XPictureGlass
End Enum

'Event Declarations:
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyDown
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Event KeyPress(KeyAscii As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyPress
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyUp
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
Event WriteProperties(PropBag As PropertyBag) 'MappingInfo=UserControl,UserControl,-1,WriteProperties
Attribute WriteProperties.VB_Description = "Occurs when a user control or user document is asked to write its data to a file."
Event Paint() 'MappingInfo=UserControl,UserControl,-1,Paint
Attribute Paint.VB_Description = "Occurs when any part of a form or PictureBox control is moved, enlarged, or exposed."
Event OLEStartDrag(Data As DataObject, AllowedEffects As Long) 'MappingInfo=UserControl,UserControl,-1,OLEStartDrag
Attribute OLEStartDrag.VB_Description = "Occurs when an OLE drag/drop operation is initiated either manually or automatically."
Event OLESetData(Data As DataObject, DataFormat As Integer) 'MappingInfo=UserControl,UserControl,-1,OLESetData
Attribute OLESetData.VB_Description = "Occurs at the OLE drag/drop source control when the drop target requests data that was not provided to the DataObject during the OLEDragStart event."
Event OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean) 'MappingInfo=UserControl,UserControl,-1,OLEGiveFeedback
Attribute OLEGiveFeedback.VB_Description = "Occurs at the source control of an OLE drag/drop operation when the mouse cursor needs to be changed."
Event OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer) 'MappingInfo=UserControl,UserControl,-1,OLEDragOver
Attribute OLEDragOver.VB_Description = "Occurs when the mouse is moved over the control during an OLE drag/drop operation, if its OLEDropMode property is set to manual."
Event OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=UserControl,UserControl,-1,OLEDragDrop
Attribute OLEDragDrop.VB_Description = "Occurs when data is dropped onto the control via an OLE drag/drop operation, and OLEDropMode is set to manual."
Event OLECompleteDrag(Effect As Long) 'MappingInfo=UserControl,UserControl,-1,OLECompleteDrag
Attribute OLECompleteDrag.VB_Description = "Occurs at the OLE drag/drop source control after a manual or automatic drag/drop has been completed or canceled."
Event ReadProperties(PropBag As PropertyBag) 'MappingInfo=UserControl,UserControl,-1,ReadProperties
Attribute ReadProperties.VB_Description = "Occurs when a user control or user document is asked to read its data from a file."
Event Resize() 'MappingInfo=UserControl,UserControl,-1,Resize
Attribute Resize.VB_Description = "Occurs when a form is first displayed or the size of an object changes."

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = UserControl.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    UserControl.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.Font = New_Font
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BorderStyle
Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    UserControl.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
    UserControl.Refresh
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseMove(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    BugLocalMessage "UserControl_WriteProperties"
    RaiseEvent WriteProperties(PropBag)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("ForeColor", UserControl.ForeColor, &H80000012)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Font", Font, Ambient.Font)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 0)
    Call PropBag.WriteProperty("ScaleWidth", UserControl.ScaleWidth, 2235)
    Call PropBag.WriteProperty("ScaleTop", UserControl.ScaleTop, 0)
    Call PropBag.WriteProperty("ScaleMode", UserControl.ScaleMode, 1)
    Call PropBag.WriteProperty("ScaleLeft", UserControl.ScaleLeft, 0)
    Call PropBag.WriteProperty("ScaleHeight", UserControl.ScaleHeight, 2145)
    Call PropBag.WriteProperty("Picture", Picture, Nothing)
    Call PropBag.WriteProperty("OLEDropMode", UserControl.OLEDropMode, 0)
    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MaskColor", UserControl.MaskColor, -2147483633)
    Call PropBag.WriteProperty("FillStyle", UserControl.FillStyle, 1)
    Call PropBag.WriteProperty("FillColor", UserControl.FillColor, &H0&)
    Call PropBag.WriteProperty("DrawWidth", UserControl.DrawWidth, 1)
    Call PropBag.WriteProperty("DrawStyle", UserControl.DrawStyle, 0)
    Call PropBag.WriteProperty("DrawMode", UserControl.DrawMode, 13)
    Call PropBag.WriteProperty("CurrentY", UserControl.CurrentY, 0)
    Call PropBag.WriteProperty("CurrentX", UserControl.CurrentX, 0)
    Call PropBag.WriteProperty("AutoRedraw", UserControl.AutoRedraw, False)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,TextWidth
Public Function TextWidth(Str As String) As Single
Attribute TextWidth.VB_Description = "Returns the width of a text string as it would be printed in the current font."
    TextWidth = UserControl.TextWidth(Str)
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,TextHeight
Public Function TextHeight(Str As String) As Single
Attribute TextHeight.VB_Description = "Returns the height of a text string as it would be printed in the current font."
    TextHeight = UserControl.TextHeight(Str)
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ScaleY
Public Function ScaleY(Height As Single, Optional FromScale As Variant, Optional ToScale As Variant) As Single
Attribute ScaleY.VB_Description = "Converts the value for the height of a Form, PictureBox, or Printer from one unit of measure to another."
    ScaleY = UserControl.ScaleY(Height, FromScale, ToScale)
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ScaleX
Public Function ScaleX(Width As Single, Optional FromScale As Variant, Optional ToScale As Variant) As Single
Attribute ScaleX.VB_Description = "Converts the value for the width of a Form, PictureBox, or Printer from one unit of measure to another."
    ScaleX = UserControl.ScaleX(Width, FromScale, ToScale)
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ScaleWidth
Public Property Get ScaleWidth() As Single
Attribute ScaleWidth.VB_Description = "Returns/sets the number of units for the horizontal measurement of an object's interior."
    ScaleWidth = UserControl.ScaleWidth
End Property

Public Property Let ScaleWidth(ByVal New_ScaleWidth As Single)
    UserControl.ScaleWidth() = New_ScaleWidth
    PropertyChanged "ScaleWidth"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ScaleTop
Public Property Get ScaleTop() As Single
Attribute ScaleTop.VB_Description = "Returns/sets the vertical coordinates for the top edges of an object."
    ScaleTop = UserControl.ScaleTop
End Property

Public Property Let ScaleTop(ByVal New_ScaleTop As Single)
    UserControl.ScaleTop() = New_ScaleTop
    PropertyChanged "ScaleTop"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ScaleMode
Public Property Get ScaleMode() As Integer
Attribute ScaleMode.VB_Description = "Returns/sets a value indicating measurement units for object coordinates when using graphics methods or positioning controls."
    ScaleMode = UserControl.ScaleMode
End Property

Public Property Let ScaleMode(ByVal New_ScaleMode As Integer)
    UserControl.ScaleMode() = New_ScaleMode
    PropertyChanged "ScaleMode"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ScaleLeft
Public Property Get ScaleLeft() As Single
Attribute ScaleLeft.VB_Description = "Returns/sets the horizontal coordinates for the left edges of an object."
    ScaleLeft = UserControl.ScaleLeft
End Property

Public Property Let ScaleLeft(ByVal New_ScaleLeft As Single)
    UserControl.ScaleLeft() = New_ScaleLeft
    PropertyChanged "ScaleLeft"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ScaleHeight
Public Property Get ScaleHeight() As Single
Attribute ScaleHeight.VB_Description = "Returns/sets the number of units for the vertical measurement of an object's interior."
    ScaleHeight = UserControl.ScaleHeight
End Property

Public Property Let ScaleHeight(ByVal New_ScaleHeight As Single)
    UserControl.ScaleHeight() = New_ScaleHeight
    PropertyChanged "ScaleHeight"
End Property

'The Underscore following "Scale" is necessary because it
'is a Reserved Word in VBA.
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Scale
Public Sub Scale_(Optional X1 As Variant, Optional Y1 As Variant, Optional X2 As Variant, Optional Y2 As Variant)
    UserControl.Scale (X1, Y1)-(X2, Y2)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,PSet
Public Sub PSetRel(x As Single, y As Single, _
                  Optional Color As Long = vbWindowBackground)
    If Color = vbvbWindowBackground Then Color = UserControl.ForeColor
    UserControl.PSet Step(x, y), Color
End Sub

Public Sub PSetAbs(x As Single, y As Single, _
                  Optional Color As Long = vbWindowBackground)
    If Color = vbvbWindowBackground Then Color = UserControl.ForeColor
    UserControl.PSet (x, y), Color
End Sub

'The Underscore following "Point" is necessary because it
'is a Reserved Word in VBA.
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Point
Public Function Point(x As Single, y As Single) As Long
Attribute Point.VB_Description = "Returns, as an integer of type Long, the RGB color of the specified point on a Form or PictureBox object."
    Point = UserControl.Point(x, y)
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Picture
Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Returns/sets a graphic to be displayed in a control."
    Set Picture = UserControl.Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
    Set UserControl.Picture = New_Picture
    ' Begin new code
    Set UserControl.MaskPicture = New_Picture
    If Not New_Picture Is Nothing Then
        UserControl.Width = New_Picture.Width
        UserControl.Height = New_Picture.Height
    End If
    ' End new code
    PropertyChanged "Picture"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,PaintPicture
Public Sub PaintPicture(Picture As Picture, X1 As Single, Y1 As Single, Optional Width1 As Variant, Optional Height1 As Variant, Optional X2 As Variant, Optional Y2 As Variant, Optional Width2 As Variant, Optional Height2 As Variant, Optional Opcode As Variant)
Attribute PaintPicture.VB_Description = "Draws the contents of a graphics file on a Form, PictureBox, or Printer object."
    UserControl.PaintPicture Picture, X1, Y1, Width1, Height1, X2, Y2, Width2, Height2, Opcode
End Sub

Private Sub UserControl_Paint()
    BugLocalMessage "UserControl_Paint"
    RaiseEvent Paint
End Sub

Private Sub UserControl_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
    RaiseEvent OLEStartDrag(Data, AllowedEffects)
End Sub

Private Sub UserControl_OLESetData(Data As DataObject, DataFormat As Integer)
    RaiseEvent OLESetData(Data, DataFormat)
End Sub

Private Sub UserControl_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
    RaiseEvent OLEGiveFeedback(Effect, DefaultCursors)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,OLEDropMode
Public Property Get OLEDropMode() As Integer
Attribute OLEDropMode.VB_Description = "Returns/Sets whether this object can act as an OLE drop target."
    OLEDropMode = UserControl.OLEDropMode
End Property

Public Property Let OLEDropMode(ByVal New_OLEDropMode As Integer)
    UserControl.OLEDropMode() = New_OLEDropMode
    PropertyChanged "OLEDropMode"
End Property

Private Sub UserControl_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
    RaiseEvent OLEDragOver(Data, Effect, Button, Shift, x, y, State)
End Sub

Private Sub UserControl_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, x, y)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,OLEDrag
Public Sub OLEDrag()
Attribute OLEDrag.VB_Description = "Starts an OLE drag/drop event with the given control as the source."
    UserControl.OLEDrag
End Sub

Private Sub UserControl_OLECompleteDrag(Effect As Long)
    RaiseEvent OLECompleteDrag(Effect)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,MousePointer
Public Property Get MousePointer() As Integer
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object."
    MousePointer = UserControl.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As Integer)
    UserControl.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,MouseIcon
Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "Sets a custom mouse icon."
    Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set UserControl.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,MaskColor
Public Property Get MaskColor() As Long
Attribute MaskColor.VB_Description = "Returns/sets the color that specifies transparent areas in the MaskPicture."
    MaskColor = UserControl.MaskColor
End Property

Public Property Let MaskColor(ByVal New_MaskColor As Long)
    UserControl.MaskColor() = New_MaskColor
    PropertyChanged "MaskColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Line
Public Sub LineAbsAbs(X1 As Single, Y1 As Single, X2 As Single, Y2 As Single, _
                      Optional Color As Long = vbWindowBackground, _
                      Optional fBox As Boolean = False, _
                      Optional fFill As Boolean = False)
    If Color = vbvbWindowBackground Then Color = UserControl.ForeColor
    If fBox Then
        If fFill Then
            UserControl.Line (X1, Y1)-(X2, Y2), Color, BF
        Else
            UserControl.Line (X1, Y1)-(X2, Y2), Color, B
        End If
    Else
        ' Test not necessary because F without B illegal
        ' Could raise an error, but allow illegal combination
        UserControl.Line (X1, Y1)-(X2, Y2), Color
    End If
End Sub

Public Sub LineRelRel(X1 As Single, Y1 As Single, X2 As Single, Y2 As Single, _
                      Optional Color As Long = vbWindowBackground, _
                      Optional fBox As Boolean = False, _
                      Optional fFill As Boolean = False)
    If Color = vbvbWindowBackground Then Color = UserControl.ForeColor
    If fBox Then
        If fFill Then
            UserControl.Line Step(X1, Y1)-Step(X2, Y2), Color, BF
        Else
            UserControl.Line Step(X1, Y1)-Step(X2, Y2), Color, B
        End If
    Else
        ' Test not necessary because F without B illegal
        ' Could raise an error, but allow illegal combination
        UserControl.Line Step(X1, Y1)-Step(X2, Y2), Color
    End If
End Sub

Public Sub LineAbsRel(X1 As Single, Y1 As Single, X2 As Single, Y2 As Single, _
                      Optional Color As Long = vbWindowBackground, _
                      Optional fBox As Boolean = False, _
                      Optional fFill As Boolean = False)
    If Color = vbvbWindowBackground Then Color = UserControl.ForeColor
    If fBox Then
        If fFill Then
            UserControl.Line (X1, Y1)-Step(X2, Y2), Color, BF
        Else
            UserControl.Line (X1, Y1)-Step(X2, Y2), Color, B
        End If
    Else
        ' Test not necessary because F without B illegal
        ' Could raise an error, but allow illegal combination
        UserControl.Line (X1, Y1)-Step(X2, Y2), Color
    End If
End Sub

Public Sub LineRelAbs(X1 As Single, Y1 As Single, X2 As Single, Y2 As Single, _
                      Optional Color As Long = vbWindowBackground, _
                      Optional fBox As Boolean = False, _
                      Optional fFill As Boolean = False)
    If Color = vbvbWindowBackground Then Color = UserControl.ForeColor
    If fBox Then
        If fFill Then
            UserControl.Line Step(X1, Y1)-(X2, Y2), Color, BF
        Else
            UserControl.Line Step(X1, Y1)-(X2, Y2), Color, B
        End If
    Else
        ' Test not necessary because F without B illegal
        ' Could raise an error, but allow illegal combination
        UserControl.Line Step(X1, Y1)-(X2, Y2), Color
    End If
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hWnd
Public Property Get hWnd() As Long
Attribute hWnd.VB_Description = "Returns a handle (from Microsoft Windows) to an object's window."
    hWnd = UserControl.hWnd
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hDC
Public Property Get hDC() As Long
Attribute hDC.VB_Description = "Returns a handle (from Microsoft Windows) to the object's device context."
    hDC = UserControl.hDC
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,FillStyle
Public Property Get FillStyle() As Integer
Attribute FillStyle.VB_Description = "Returns/sets the fill style of a shape."
    FillStyle = UserControl.FillStyle
End Property

Public Property Let FillStyle(ByVal New_FillStyle As Integer)
    UserControl.FillStyle() = New_FillStyle
    PropertyChanged "FillStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,FillColor
Public Property Get FillColor() As OLE_COLOR
Attribute FillColor.VB_Description = "Returns/sets the color used to fill in shapes, circles, and boxes."
    FillColor = UserControl.FillColor
End Property

Public Property Let FillColor(ByVal New_FillColor As OLE_COLOR)
    UserControl.FillColor() = New_FillColor
    PropertyChanged "FillColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,DrawWidth
Public Property Get DrawWidth() As Integer
Attribute DrawWidth.VB_Description = "Returns/sets the line width for output from graphics methods."
    DrawWidth = UserControl.DrawWidth
End Property

Public Property Let DrawWidth(ByVal New_DrawWidth As Integer)
    UserControl.DrawWidth() = New_DrawWidth
    PropertyChanged "DrawWidth"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,DrawStyle
Public Property Get DrawStyle() As Integer
Attribute DrawStyle.VB_Description = "Determines the line style for output from graphics methods."
    DrawStyle = UserControl.DrawStyle
End Property

Public Property Let DrawStyle(ByVal New_DrawStyle As Integer)
    UserControl.DrawStyle() = New_DrawStyle
    PropertyChanged "DrawStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,DrawMode
Public Property Get DrawMode() As Integer
Attribute DrawMode.VB_Description = "Sets the appearance of output from graphics methods or of a Shape or Line control."
    DrawMode = UserControl.DrawMode
End Property

Public Property Let DrawMode(ByVal New_DrawMode As Integer)
    UserControl.DrawMode() = New_DrawMode
    PropertyChanged "DrawMode"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,CurrentY
Public Property Get CurrentY() As Single
Attribute CurrentY.VB_Description = "Returns/sets the vertical coordinates for next print or draw method."
    CurrentY = UserControl.CurrentY
End Property

Public Property Let CurrentY(ByVal New_CurrentY As Single)
    UserControl.CurrentY() = New_CurrentY
    PropertyChanged "CurrentY"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,CurrentX
Public Property Get CurrentX() As Single
Attribute CurrentX.VB_Description = "Returns/sets the horizontal coordinates for next print or draw method."
    CurrentX = UserControl.CurrentX
End Property

Public Property Let CurrentX(ByVal New_CurrentX As Single)
    UserControl.CurrentX() = New_CurrentX
    PropertyChanged "CurrentX"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Cls
Public Sub Cls()
Attribute Cls.VB_Description = "Clears graphics and text generated at run time from a Form, Image, or PictureBox."
    UserControl.Cls
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ClipControls
Public Property Get ClipControls() As Boolean
Attribute ClipControls.VB_Description = "Determines whether graphics methods in Paint events repaint an entire object or newly exposed areas."
    ClipControls = UserControl.ClipControls
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Circle
Public Sub CircleRel(x As Single, y As Single, Radius As Single, _
                     Optional Color As Long = vbWindowBackground, _
                     Optional StartPos As Single = 0, _
                     Optional EndPos As Variant = 2 * 3.14159, _
                     Optional Aspect As Single = 1#)
    If Color = vbvbWindowBackground Then Color = UserControl.ForeColor
    UserControl.Circle Step(x, y), Radius, Color, StartPos, EndPos, Aspect
End Sub

Public Sub CircleAbs(x As Single, y As Single, Radius As Single, _
                     Optional Color As Long = vbWindowBackground, _
                     Optional StartPos As Single = 0, _
                     Optional EndPos As Variant = 2 * 3.14159, _
                     Optional Aspect As Single = 1#)
    If Color = vbvbWindowBackground Then Color = UserControl.ForeColor
    UserControl.Circle (x, y), Radius, Color, StartPos, EndPos, Aspect
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,AutoRedraw
Public Property Get AutoRedraw() As Boolean
Attribute AutoRedraw.VB_Description = "Returns/sets the output from a graphics method to a persistent bitmap."
    AutoRedraw = UserControl.AutoRedraw
End Property

Public Property Let AutoRedraw(ByVal New_AutoRedraw As Boolean)
    UserControl.AutoRedraw() = New_AutoRedraw
    PropertyChanged "AutoRedraw"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Appearance
Public Property Get Appearance() As Integer
Attribute Appearance.VB_Description = "Returns/sets whether or not an object is painted at run time with 3-D effects."
    Appearance = UserControl.Appearance
End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    BugLocalMessage "UserControl_ReadProperties"
    RaiseEvent ReadProperties(PropBag)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    UserControl.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set Font = PropBag.ReadProperty("Font", Ambient.Font)
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
    UserControl.ScaleWidth = PropBag.ReadProperty("ScaleWidth", 2235)
    UserControl.ScaleTop = PropBag.ReadProperty("ScaleTop", 0)
    UserControl.ScaleMode = PropBag.ReadProperty("ScaleMode", 1)
    UserControl.ScaleLeft = PropBag.ReadProperty("ScaleLeft", 0)
    UserControl.ScaleHeight = PropBag.ReadProperty("ScaleHeight", 2145)
    Set Picture = PropBag.ReadProperty("Picture", Nothing)
    UserControl.OLEDropMode = PropBag.ReadProperty("OLEDropMode", 0)
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    UserControl.MaskColor = PropBag.ReadProperty("MaskColor", -2147483633)
    UserControl.FillStyle = PropBag.ReadProperty("FillStyle", 1)
    UserControl.FillColor = PropBag.ReadProperty("FillColor", &H0&)
    UserControl.DrawWidth = PropBag.ReadProperty("DrawWidth", 1)
    UserControl.DrawStyle = PropBag.ReadProperty("DrawStyle", 0)
    UserControl.DrawMode = PropBag.ReadProperty("DrawMode", 13)
    UserControl.CurrentY = PropBag.ReadProperty("CurrentY", 0)
    UserControl.CurrentX = PropBag.ReadProperty("CurrentX", 0)
    UserControl.AutoRedraw = PropBag.ReadProperty("AutoRedraw", False)
End Sub

Private Sub UserControl_InitProperties()
    BugLocalMessage "UserControl_InitProperties"
    Set Font = Ambient.Font
    Extender.Name = UniqueControlName("pg", Extender)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Image
Public Property Get Image() As Picture
Attribute Image.VB_Description = "Returns a handle, provided by Microsoft Windows, to a persistent bitmap."
    Set Image = UserControl.Image
End Property

Private Sub UserControl_Resize()
    BugLocalMessage "UserControl_Resize"
    RaiseEvent Resize
End Sub



