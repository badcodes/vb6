VERSION 5.00
Begin VB.UserControl XDropStack 
   ClientHeight    =   1170
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2145
   ScaleHeight     =   1170
   ScaleWidth      =   2145
   ToolboxBitmap   =   "dropstack.ctx":0000
   Begin VB.ComboBox cbo 
      Height          =   288
      ItemData        =   "dropstack.ctx":00FA
      Left            =   120
      List            =   "dropstack.ctx":00FC
      TabIndex        =   0
      Top             =   240
      Width           =   804
   End
End
Attribute VB_Name = "XDropStack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum EErrorDropStack
    eeBaseDropStack = 13710     ' XDropStack
End Enum

Private cMaxCount As Integer
Private fInUpdate As Boolean
Private fCompleted As Boolean

'Event Declarations:
Event Change()
Attribute Change.VB_Description = "Occurs when the contents of a control have changed."
Event Completed(Text As String)

Private Sub UserControl_Resize()
    BugLocalMessage "XDropStack UserControl_Resize"
    cbo.Left = 0
    cbo.Top = 0
    cbo.Width = Width
    Height = cbo.Height
End Sub

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = cbo.ForeColor
End Property

Public Property Let ForeColor(ByVal clrForeColor As OLE_COLOR)
    cbo.ForeColor() = clrForeColor
    PropertyChanged "ForeColor"
End Property

Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = cbo.Font
End Property

Public Property Set Font(ByVal fntFont As Font)
    Set cbo.Font = fntFont
    PropertyChanged "Font"
End Property

Public Property Get List() As Collection
Attribute List.VB_Description = "Returns/sets the items contained in a control's list portion."
    Dim n As Collection, i As Integer
    Set n = New Collection
    For i = 0 To cbo.ListCount - 1
        n.Add cbo.List(i)
    Next
    Set List = n
End Property

Public Property Set List(n As Collection)
With cbo
    ' Remove any old items and add new list
    If .ListCount Then .Clear
    Dim v As Variant
    For Each v In n
        If VarType(v) = vbString Then
            If v <> sEmpty Then .AddItem v
        End If
    Next
    ' Select first item
    .Refresh
    fInUpdate = True
    If .ListCount Then .ListIndex = 0
    fInUpdate = False
End With
End Property

Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
    cbo.Refresh
End Sub

Public Property Get Appearance() As Integer
Attribute Appearance.VB_Description = "Returns/sets whether or not an object is painted at run time with 3-D effects."
    Appearance = cbo.Appearance
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = cbo.BackColor
End Property

Public Property Let BackColor(ByVal clrBackColor As OLE_COLOR)
    cbo.BackColor() = clrBackColor
    PropertyChanged "BackColor"
End Property

Public Property Get hWnd() As Long
    hWnd = cbo.hWnd
End Property

Public Property Get Count() As Integer
Attribute Count.VB_Description = "Returns the number of items in the list portion of a control."
Attribute Count.VB_MemberFlags = "400"
    Count = cbo.ListCount
End Property

Public Property Get MaxCount() As Long
    MaxCount = cMaxCount
End Property

Public Property Let MaxCount(ByVal cMaxCountA As Long)
    cMaxCount = cMaxCountA
End Property

Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "Sets a custom mouse icon."
    Set MouseIcon = cbo.MouseIcon
End Property

Public Property Set MouseIcon(ByVal picMouseIcon As Picture)
    Set cbo.MouseIcon = picMouseIcon
End Property

Public Property Get MousePointer() As MousePointerConstants
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object."
    MousePointer = cbo.MousePointer
End Property

Public Property Let MousePointer(ByVal ordMousePointer As MousePointerConstants)
    cbo.MousePointer() = ordMousePointer
    PropertyChanged "MousePointer"
End Property

Public Property Get SelLength() As Long
Attribute SelLength.VB_Description = "Returns/sets the number of characters selected."
Attribute SelLength.VB_MemberFlags = "400"
    SelLength = cbo.SelLength
End Property

Public Property Let SelLength(ByVal cSelLength As Long)
    cbo.SelLength() = cSelLength
End Property

Public Property Get SelStart() As Long
Attribute SelStart.VB_Description = "Returns/sets the starting point of text selected."
Attribute SelStart.VB_MemberFlags = "400"
    SelStart = cbo.SelStart
End Property

Public Property Let SelStart(ByVal iSelStart As Long)
    cbo.SelStart() = iSelStart
End Property

Public Property Get SelText() As String
Attribute SelText.VB_Description = "Returns/sets the string containing the currently selected text."
Attribute SelText.VB_MemberFlags = "400"
    SelText = cbo.SelText
End Property

Public Property Let SelText(ByVal sSelText As String)
    cbo.SelText() = sSelText
End Property

Public Property Get Text() As String
Attribute Text.VB_Description = "Returns/sets the text contained in the control."
Attribute Text.VB_UserMemId = 0
Attribute Text.VB_MemberFlags = "400"
    Text = cbo.Text
End Property

Public Property Let Text(ByVal sText As String)
    If sText = sEmpty Then Exit Property
    UpdateItem sText
    PropertyChanged "Text"
End Property

Public Property Get ItemData() As Long
    If cbo.ListCount Then ItemData = cbo.ItemData(0)
End Property

Public Property Let ItemData(ByVal iItemData As Long)
    If cbo.ListCount Then cbo.ItemData(0) = iItemData
End Property

Public Sub Clear()
Attribute Clear.VB_Description = "Clears the contents of a control or the system Clipboard."
    cbo.Clear
End Sub

Private Sub cbo_Click()
    If fInUpdate Then Exit Sub
    UpdateItem cbo.Text
End Sub

Private Sub cbo_Change()
    fCompleted = False
    RaiseEvent Change
End Sub

Private Sub cbo_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyReturn, vbKeyEscape ' , vbKeyTab
        Text = cbo.Text
    End Select
End Sub

Private Sub cbo_LostFocus()
    BugLocalMessage "XDropStack cbo_LostFocus"
    If fCompleted = False Then Text = cbo.Text
End Sub

' Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    BugLocalMessage "XDropStack UserControl_InitProperties"
    Extender.Name = UniqueControlName("drop", Extender)
End Sub

' Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    BugLocalMessage "XDropStack UserControl_ReadProperties"
    cbo.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
    cbo.ForeColor = PropBag.ReadProperty("ForeColor", &H80000008)
    Set Font = PropBag.ReadProperty("Font", Ambient.Font)
    cMaxCount = PropBag.ReadProperty("MaxCount", 0)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    cbo.MousePointer = PropBag.ReadProperty("MousePointer", 0)
End Sub

' Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    BugLocalMessage "XDropStack UserControl_WriteProperties"
    Call PropBag.WriteProperty("BackColor", cbo.BackColor, &H80000005)
    Call PropBag.WriteProperty("ForeColor", cbo.ForeColor, &H80000008)
    Call PropBag.WriteProperty("Font", Font, Ambient.Font)
    Call PropBag.WriteProperty("MaxCount", cMaxCount, 0)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", cbo.MousePointer, 0)
End Sub

Private Sub UpdateItem(sText As String)
    BugAssert sText <> sEmpty
With cbo
    Dim i As Integer, f As Boolean
    For i = 0 To .ListCount - 1
        ' If item is in list, remove in order to move to top
        If .List(i) = sText Then .RemoveItem i
    Next
    ' Add item to top of list
    .AddItem sText, 0
    ' Remove any extra
    If cMaxCount And .ListCount > cMaxCount Then
        .RemoveItem cMaxCount
    End If
    ' Disable cbo_Click procedure
    fInUpdate = True
    ' Select new item
    .ListIndex = 0
    fInUpdate = False
    fCompleted = True
    RaiseEvent Completed(sText)
End With
End Sub
