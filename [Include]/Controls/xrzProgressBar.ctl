VERSION 5.00
Begin VB.UserControl xrzProgressBar 
   BackColor       =   &H8000000D&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   360
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3390
   ClipBehavior    =   0  'None
   FillStyle       =   6  'Cross
   FontTransparent =   0   'False
   ForeColor       =   &H8000000C&
   MousePointer    =   11  'Hourglass
   ScaleHeight     =   360
   ScaleWidth      =   3390
End
Attribute VB_Name = "xrzProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Default Property Values:
Const m_def_MinWidth = 0
Const m_def_MaxWidth = 1
'Property Variables:
Dim m_MinWidth As Long
Dim m_MaxWidth As Long
'Event Declarations:
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."



'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Shape,Shape,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = Shape.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    Shape.BackColor() = New_BackColor
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
'MappingInfo=Shape,Shape,-1,BackStyle
Public Property Get BackStyle() As Integer
Attribute BackStyle.VB_Description = "Indicates whether a Label or the background of a Shape is transparent or opaque."
    BackStyle = Shape.BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As Integer)
    Shape.BackStyle() = New_BackStyle
    PropertyChanged "BackStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Shape,Shape,-1,BorderStyle
Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = Shape.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    Shape.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Shape,Shape,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
    Shape.Refresh
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Shape,Shape,-1,BorderColor
Public Property Get BorderColor() As Long
Attribute BorderColor.VB_Description = "Returns/sets the color of an object's border."
    BorderColor = Shape.BorderColor
End Property

Public Property Let BorderColor(ByVal New_BorderColor As Long)
    Shape.BorderColor() = New_BorderColor
    PropertyChanged "BorderColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Shape,Shape,-1,AutoRedraw
Public Property Get AutoRedraw() As Boolean
Attribute AutoRedraw.VB_Description = "Returns/sets the output from a graphics method to a persistent bitmap."
    AutoRedraw = Shape.AutoRedraw
End Property

Public Property Let AutoRedraw(ByVal New_AutoRedraw As Boolean)
    Shape.AutoRedraw() = New_AutoRedraw
    PropertyChanged "AutoRedraw"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Shape,Shape,-1,Shape
Public Property Get Shape() As Integer
Attribute Shape.VB_Description = "Returns/sets a value indicating the appearance of a control."
    Shape = Shape.Shape
End Property

Public Property Let Shape(ByVal New_Shape As Integer)
    Shape.Shape() = New_Shape
    PropertyChanged "Shape"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get MinWidth() As Long
    MinWidth = m_MinWidth
End Property

Public Property Let MinWidth(ByVal New_MinWidth As Long)
    m_MinWidth = New_MinWidth
    PropertyChanged "MinWidth"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,1
Public Property Get MaxWidth() As Long
    MaxWidth = m_MaxWidth
End Property

Public Property Let MaxWidth(ByVal New_MaxWidth As Long)
    m_MaxWidth = New_MaxWidth
    PropertyChanged "MaxWidth"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_MinWidth = m_def_MinWidth
    m_MaxWidth = m_def_MaxWidth
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    Shape.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    UserControl.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Shape.BackStyle = PropBag.ReadProperty("BackStyle", 1)
    Shape.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
    Shape.BorderColor = PropBag.ReadProperty("BorderColor", -2147483640)
    Shape.AutoRedraw = PropBag.ReadProperty("AutoRedraw", False)
    Shape.Shape = PropBag.ReadProperty("Shape", 4)
    m_MinWidth = PropBag.ReadProperty("MinWidth", m_def_MinWidth)
    m_MaxWidth = PropBag.ReadProperty("MaxWidth", m_def_MaxWidth)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", Shape.BackColor, &H8000000F)
    Call PropBag.WriteProperty("ForeColor", UserControl.ForeColor, &H80000012)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("BackStyle", Shape.BackStyle, 1)
    Call PropBag.WriteProperty("BorderStyle", Shape.BorderStyle, 1)
    Call PropBag.WriteProperty("BorderColor", Shape.BorderColor, -2147483640)
    Call PropBag.WriteProperty("AutoRedraw", Shape.AutoRedraw, False)
    Call PropBag.WriteProperty("Shape", Shape.Shape, 4)
    Call PropBag.WriteProperty("MinWidth", m_MinWidth, m_def_MinWidth)
    Call PropBag.WriteProperty("MaxWidth", m_MaxWidth, m_def_MaxWidth)
End Sub

