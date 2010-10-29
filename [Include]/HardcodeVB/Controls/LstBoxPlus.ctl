VERSION 5.00
Begin VB.UserControl XListBoxPlus 
   ClientHeight    =   1896
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3420
   PropertyPages   =   "LstBoxPlus.ctx":0000
   ScaleHeight     =   1896
   ScaleWidth      =   3420
   ToolboxBitmap   =   "LstBoxPlus.ctx":002A
   Begin VB.ListBox lst 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   600
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "XListBoxPlus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit

Public Enum EErrorListBoxPlus
    eeBaseListBoxPlus = 13730   ' XListBoxPlus
End Enum

Private myWidth As Integer
Private myHeight As Integer

Private esmlMode As ESortModeList
Private fHiToLo As Boolean
Private eaAppearance As EAppearance

Private fCompletion As Boolean
Private sPartial As String
Private iPrevKey As Long

'Event Declarations:
Event Click()
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event DblClick()
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event KeyDown(KeyCode As Integer, Shift As Integer)
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Event KeyPress(KeyAscii As Integer)
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Event KeyUp(KeyCode As Integer, Shift As Integer)
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
Event ItemCheck(item As Integer)
Attribute ItemCheck.VB_Description = "Occurs when a ListBox control's Style property is set to 1 (checkboxes) and an item's checkbox in the ListBox control is selected or cleared."
Event OLECompleteDrag(Effect As Long)
Attribute OLECompleteDrag.VB_Description = "Occurs at the OLE drag/drop source control after a manual or automatic drag/drop has been completed or canceled."
Event OLEDragDrop(data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
Attribute OLEDragDrop.VB_Description = "Occurs when data is dropped onto the control via an OLE drag/drop operation, and OLEDropMode is set to manual."
Event OLEDragOver(data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
Attribute OLEDragOver.VB_Description = "Occurs when the mouse is moved over the control during an OLE drag/drop operation, if its OLEDropMode property is set to manual."
Event OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
Attribute OLEGiveFeedback.VB_Description = "Occurs at the source control of an OLE drag/drop operation when the mouse cursor needs to be changed."
Event OLESetData(data As DataObject, DataFormat As Integer)
Attribute OLESetData.VB_Description = "Occurs at the OLE drag/drop source control when the drop target requests data that was not provided to the DataObject during the OLEDragStart event."
Event OLEStartDrag(data As DataObject, AllowedEffects As Long)
Attribute OLEStartDrag.VB_Description = "Occurs when an OLE drag/drop operation is initiated either manually or automatically."
Event Scroll()
Attribute Scroll.VB_Description = "Occurs when you reposition the scroll box on a control."

' Friend properties to make data structure accessible to walker
Friend Property Get ListItems(i As Long) As String
    ListItems = item(i)
End Property

' NewEnum must have the procedure ID -4 in Procedure Attributes dialog
' Create a new data walker object and connect to it
Public Function NewEnum() As IEnumVARIANT
Attribute NewEnum.VB_UserMemId = -4
    ' Create a new iterator object
    Dim ListItemwalker As CListItemWalker
    Set ListItemwalker = New CListItemWalker
    ' Connect it with collection data
    ListItemwalker.Attach Me
    ' Return it
    Set NewEnum = ListItemwalker.NewEnum
End Function

Private Sub UserControl_Initialize()
    Debug.Print "Initialize"
End Sub

' Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    esmlMode = esmlUnsorted
    fHiToLo = False
    Extender.Name = UniqueControlName("list", Extender)
End Sub

Private Sub UserControl_Paint()
    DrawAppearance lst
End Sub

' Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
With lst
    Appearance = PropBag.ReadProperty("Appearance", ea3D)
    'Current = PropBag.ReadProperty("Current", 1)
    '.Columns = PropBag.ReadProperty("Columns", 0)
    .BackColor = PropBag.ReadProperty("BackColor", vbWindowBackground)
    .ForeColor = PropBag.ReadProperty("ForeColor", vbButtonText)
    '.DataField = PropBag.ReadProperty("DataField", 0)
    '.DataSource = PropBag.ReadProperty("DataSource", 0)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set .Font = PropBag.ReadProperty("Font", lst.Font)
    HiToLo = PropBag.ReadProperty("HiToLo", True)
    SortMode = PropBag.ReadProperty("SortMode", esmlSortVal)
    IntegralHeight = PropBag.ReadProperty("IntegralHeight", True)
    Dim i As Integer, iListCount As Integer
    iListCount = PropBag.ReadProperty("ListCount", 0)
    If iListCount Then
        Clear
        For i = 0 To iListCount - 1
            Add PropBag.ReadProperty("List" & i, sEmpty), , _
                PropBag.ReadProperty("ItemData" & i, 0)
        Next
    End If
    Completion = PropBag.ReadProperty("Completion", False)
    Set .MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    .MousePointer = PropBag.ReadProperty("MousePointer", 0)
    '.MultiSelect = PropBag.ReadProperty("MultiSelect", vbMultiSelectNone)
    .OLEDragMode = PropBag.ReadProperty("OLEDragMode", 0)
    .OLEDropMode = PropBag.ReadProperty("OLEDropMode", 0)
    .RightToLeft = PropBag.ReadProperty("RightToLeft", False)
'TO DO: The member you have mapped to contains an array of data.
'   You must supply the code to persist the array.  A prototype
'   line is shown next:
'   .Selected(Index) = PropBag.ReadProperty("Selected" & Index, 0)
    .Text = PropBag.ReadProperty("Text", "")
    '.TopIndex = PropBag.ReadProperty("TopItem", 1)
End With
End Sub

' Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
With lst
    Debug.Print "WriteProperties: "
    PropBag.WriteProperty "Appearance", Appearance, ea3D
    'PropBag.WriteProperty "Current", Current, 1
    'PropBag.WriteProperty "Columns", Columns, 1
    PropBag.WriteProperty "BackColor", .BackColor, vbWindowBackground
    PropBag.WriteProperty "ForeColor", .ForeColor, vbButtonText
    'PropBag.WriteProperty "DataField", .DataField, 0
    'PropBag.WriteProperty "DataSource", .DataSource, 0
    PropBag.WriteProperty "Enabled", UserControl.Enabled, True
    PropBag.WriteProperty "Font", .Font, lst.Font
    PropBag.WriteProperty "HiToLo", HiToLo, True
    PropBag.WriteProperty "IntegralHeight", IntegralHeight, True
    PropBag.WriteProperty "SortMode", SortMode, esmlSortVal
    PropBag.WriteProperty "ListCount", ListCount
    Dim i As Integer
    For i = 0 To ListCount - 1
        PropBag.WriteProperty "List" & i, List(i)
        PropBag.WriteProperty "ItemData" & i, ItemData(i)
    Next
    PropBag.WriteProperty "ListIndex", .ListIndex
    PropBag.WriteProperty "Completion", Completion
    PropBag.WriteProperty "MouseIcon", .MouseIcon, Nothing
    PropBag.WriteProperty "MousePointer", .MousePointer, 0
    PropBag.WriteProperty "MultiSelect", .MultiSelect, vbMultiSelectNone
    PropBag.WriteProperty "OLEDragMode", .OLEDragMode, 0
    PropBag.WriteProperty "OLEDropMode", .OLEDropMode, 0
    PropBag.WriteProperty "RightToLeft", .RightToLeft, False
'TO DO: The member you have mapped to contains an array of data.
'   You must supply the code to persist the array.  A prototype
'   line is shown next:
'   PropBag.WriteProperty "Selected" & Index, .Selected(Index), 0
    PropBag.WriteProperty "Text", .Text, ""
    'PropBag.WriteProperty "TopItem", .TopIndex, 1
End With
End Sub

Private Sub UserControl_Resize()
    Static fInside As Boolean
    If fInside Then Exit Sub
    fInside = True
    ' Adjust control to ListBox
    lst.Move 0, 0, Width, Height
    ' But ListBox height is in item increments, so adjust again
    Height = lst.Height
    Width = lst.Width
    myHeight = Height
    myWidth = Width
    fInside = False
End Sub

Private Sub UserControl_Show()
    ' Handle List?
    Randomize
End Sub

''' Public Methods Unique to This Class '''
    
' Collection Methods

' Add ignores the optional iPos argument except in unsorted mode.
' You cannot specify the insert position with a sorted list as you
' can with an unsorted list. Also, you cannot insert an item that
' already exists into a sorted list. A request to do so will generate
' an error.
Sub Add(sItem As String, Optional iPos As Integer = 1, _
        Optional iItemData As Long)
With lst
    ' Adding differs depending on the mode
    Select Case esmlMode
    Case esmlUnsorted
        ' Add where directed (start is default)
        .AddItem sItem, iPos - 1
    Case esmlShuffle
        ' Add at random position
        iPos = GetRandom(0, .ListCount - 1)
        If .ListCount Then
            .AddItem sItem, iPos
        Else
            .AddItem sItem
        End If
        
    Case Else   ' Some kind of sorting
        ' Binary search for the item
        If BSearch(sItem, iPos) Then
            ErrRaise eseDuplicateNotAllowed
        Else
            ' Insert at sorted position
            If .ListCount Then
                .AddItem sItem, iPos
            Else
                .AddItem sItem
            End If
        End If
    End Select
    .ItemData(.NewIndex) = iItemData
    PropertyChanged "List"
End With
End Sub

' Same as RemoveItem but has collection name and is 1-based
Sub Remove(ByVal vIndex As Variant)
    If VarType(vIndex) = vbString Then
        vIndex = Match(vIndex)
        If vIndex = 0 Then ErrRaise eseItemNotFound
    Else
        If vIndex > Count Or vIndex < 1 Then ErrRaise eseOutOfRange
    End If
    lst.RemoveItem vIndex - 1
End Sub

' AddItem and RemoveItem are 0-based for compatibility
Sub AddItem(sItem As String, Optional iPos As Integer, Optional iItemData As Long)
    Add sItem, iPos + 1, iItemData
End Sub

Sub RemoveItem(vIndex As Variant)
    If VarType(vIndex) = vbString Then
        Remove vIndex
    Else
        Remove vIndex + 1
    End If
End Sub

' Similar to List property
Property Get item(ByVal vIndex As Variant) As String
Attribute item.VB_UserMemId = 0
    If VarType(vIndex) <> vbString Then
        ' For numeric index, return string value
        item = lst.List(vIndex - 1)
    Else
        ' For string index, return matching index or 0 for none
        item = Match(vIndex)
        If item = 0 Then ErrRaise eseItemNotFound
    End If
End Property

Property Let item(ByVal vIndex As Variant, sItemA As String)
    ' For string index, look up matching index
    If VarType(vIndex) = vbString Then
        vIndex = Match(vIndex)
        ' Fail if old item isn't found or if new item is found
        If vIndex = 0 Then ErrRaise eseItemNotFound
    End If
    If Match(sItemA) Then ErrRaise eseDuplicateNotAllowed
    ' Assign value by removing old and inserting new
    Remove vIndex
    Add sItemA
    PropertyChanged "List"
End Property

''' Public Properties Unique to This Class '''

Property Let HiToLo(fHiToLoA As Boolean)
    fHiToLo = fHiToLoA
    Select Case esmlMode
    Case esmlUnsorted, esmlShuffle
        ' Leave as is
    Case Else   ' Some kind of sorting
        Sort 0, lst.ListCount - 1
    End Select
End Property

Property Get HiToLo() As Boolean
    HiToLo = fHiToLo
End Property

Property Let SortMode(esmlModeA As ESortModeList)
    esmlMode = esmlModeA
    Select Case esmlMode
    Case esmlUnsorted
        ' Leave everything as is
    Case esmlShuffle
        Shuffle
    Case Else   ' Some kind of sorting
        Sort 0, lst.ListCount - 1
    End Select
End Property

Property Get SortMode() As ESortModeList
    SortMode = esmlMode
End Property

' Gives away the store for iteration
Property Get Items() As Collection
    Set Items = lst
End Property

' Collection name
Property Get Count() As Integer
    Count = lst.ListCount
End Property

Property Get Current() As Variant
Attribute Current.VB_MemberFlags = "400"
    Current = lst.ListIndex + 1
End Property

Property Let Current(vIndexA As Variant)
    If lst.ListCount = 0 Then Exit Property
    If VarType(vIndexA) <> vbString Then
        lst.ListIndex = vIndexA - 1
    Else
        lst.ListIndex = Match(vIndexA) - 1
    End If
    If lst.ListIndex = -1 Then ErrRaise eseItemNotFound
End Property

Property Get IndexItem() As Variant
    IndexItem = lst.List(lst.ListIndex)
End Property

' 1-based versions of ItemData
Property Get data(i As Integer) As Variant
    data = lst.ItemData(i - 1)
End Property

Property Let data(i As Integer, vData As Variant)
    lst.ItemData(i - 1) = vData
End Property

''' Public Methods From Contained Class '''
Public Sub Clear()
Attribute Clear.VB_Description = "Clears the contents of a control or the system Clipboard."
    lst.Clear
End Sub

Sub Drag(Optional vAction As Variant)
    If IsMissing(vAction) Then
        lst.Drag
    Else
        lst.Drag vAction
    End If
End Sub

Sub Move(x As Variant, Optional y As Variant, Optional dx As Variant, Optional dy As Variant)
    If IsMissing(y) Then
        lst.Move x
    ElseIf IsMissing(dx) Then
        lst.Move x, y
    ElseIf IsMissing(dy) Then
        lst.Move x, y, dx
    Else
        lst.Move x, y, dx, dy
    End If
End Sub

Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
    lst.Refresh
End Sub

Sub SetFocus()
    lst.SetFocus
End Sub

Sub ZOrder(Optional vPosition As Variant)
    If IsMissing(vPosition) Then
        lst.ZOrder
    Else
        lst.ZOrder vPosition
    End If
End Sub

Public Property Let Completion(fCompletionA As Boolean)
Attribute Completion.VB_Description = "Enables/disables word completion."
    If fCompletion = False And fCompletionA = True Then
        sPartial = sEmpty
    End If
    fCompletion = fCompletionA
    PropertyChanged "Completion"
End Property
Public Property Get Completion() As Boolean
    Completion = fCompletion
End Property

Public Property Let PartialWord(sPartialA As String)
    If fCompletion Then
        sPartial = sPartialA
        If (SortMode = esmlSortText) Or _
           (SortMode = esmlSortBin) Then
            If sPartial <> sEmpty Then
                CompleteWord
            Else
                If ListIndex <> -1 Then lst.Selected(ListIndex) = False
            End If
        End If
    End If
End Property
Public Property Get PartialWord() As String
Attribute PartialWord.VB_Description = "Returns/sets the string to be completed."
Attribute PartialWord.VB_MemberFlags = "400"
    PartialWord = sPartial
End Property

''' Private Procedures Used by Class '''

Private Sub CompleteWord()
    Dim iPos As Integer, cPartial As Integer, sItem As String
    If BSearch(sPartial, iPos) Then
        lst.Selected(iPos) = True
    Else
        ' Item not found. Look for possible completion
        If Compare(Left$(List(iPos), Len(sPartial)), sPartial) = 0 Then
            ' Found a completion
            lst.Selected(iPos) = True
        Else
            ' Didn't find a completion
            If lst.ListIndex <> -1 Then
                lst.Selected(lst.ListIndex) = False
            End If
        End If
    End If
End Sub

Private Function Match(ByVal sItem As String) As Integer
    Dim iPos As Integer
    Select Case esmlMode
    Case esmlUnsorted, esmlShuffle
        Match = LookupItem(lst, sItem) + 1
    Case Else   ' Some kind of sorting
        If BSearch(sItem, iPos) Then Match = iPos + 1 Else Match = 0
    End Select
End Function

Private Sub Sort(iFirst As Integer, iLast As Integer)
    Dim vSplit As Variant

    If iFirst < iLast Then

        ' Only two elements in this subdivision. Exchange if
        ' they are out of order, and end recursive calls.
        If iLast - iFirst = 1 Then
            If Compare(lst.List(iFirst), lst.List(iLast)) > 0 Then
                Swap iFirst, iLast
            End If
        Else

            Dim i As Integer, j As Integer, iRand As Integer

            ' Pick pivot element at random and move to end
            ' (consider calling Randomize before sorting)
            iRand = GetRandom(iFirst, iLast)
            Swap iLast, iRand
            vSplit = lst.List(iLast)
            Do

                ' Move in from both sides towards the pivot element
                i = iFirst: j = iLast
                Do While (i < j) And _
                    Compare(lst.List(i), vSplit) <= 0
                    i = i + 1
                Loop
                Do While (j > i) And _
                    Compare(lst.List(j), vSplit) >= 0
                    j = j - 1
                Loop

                ' If we haven't reached the pivot element, it means
                ' that two elements on either side are out of order,
                ' so swap them
                If i < j Then
                    Swap i, j
                End If
            Loop While i < j

            ' Move the pivot element back to its proper place
            Swap i, iLast

            ' Recursively call Sort (pass the smaller
            ' subdivision first to use less stack space)
            If (i - iFirst) < (iLast - i) Then
                Sort iFirst, i - 1
                Sort i + 1, iLast
            Else
                Sort i + 1, iLast
                Sort iFirst, i - 1
            End If
        End If
    End If

End Sub

Private Function BSearch(sKey As String, iPos As Integer) As Boolean
    Dim iLo As Integer, iHi As Integer, iComp As Integer, iMid As Integer
    iLo = 0: iHi = lst.ListCount - 1
    Do
        iMid = iLo + ((iHi - iLo) \ 2)
        iComp = Compare(lst.List(iMid), sKey)
        Select Case iComp
        Case 0
            ' Item found
            iPos = iMid
            BSearch = True
            Exit Function
        Case Is > 0
            ' Item is in upper half
            iHi = iMid
            If iLo = iHi Then Exit Do
        Case Is < 0
            ' Item is in lower half
            iLo = iMid + 1
            If iLo > iHi Then Exit Do
        End Select
    Loop
    ' Item not found, but return position to insert
    iPos = iMid - (iComp < 0)
    BSearch = False

End Function

Sub Shuffle()
    Dim iFirst As Integer, iLast As Integer
    iFirst = 0: iLast = lst.ListCount - 1
    ' Randomize list
    Dim i As Integer, v As Variant, iRnd As Integer
    For i = iLast To iFirst + 1 Step -1
        ' Swap random element with last element
        iRnd = GetRandom(iFirst, i)
        Swap i, iRnd
    Next
End Sub

Private Function Compare(v1 As Variant, v2 As Variant) As Integer
    Dim i As Integer
    If IsNumeric(v1) And IsNumeric(v2) Then
        v1 = Val(v1)
        v2 = Val(v2)
    End If
    
    Select Case esmlMode
    ' Sort by value (same as esmlSortBin for strings)
    Case esmlSortVal
        If v1 < v2 Then
            i = -1
        ElseIf v1 = v2 Then
            i = 0
        Else
            i = 1
        End If
    ' Sort case-insensitive
    Case esmlSortText
        i = StrComp(v1, v2, 1)
    ' Sort case-sensitive
    Case esmlSortBin
        i = StrComp(v1, v2, 0)
    ' Sort by string length
    Case esmlSortLen
        If Len(v1) = Len(v2) Then
            If v1 = v2 Then
                i = 0
            ElseIf v1 < v2 Then
                i = -1
            Else
                i = 1
            End If
        ElseIf Len(v1) < Len(v2) Then
            i = -1
        Else
            i = 1
        End If
    End Select
    If fHiToLo Then i = -i
    Compare = i
End Function

Sub Swap(i1 As Integer, i2 As Integer)
    Dim s As String
    s = lst.List(i1)
    lst.List(i1) = lst.List(i2)
    lst.List(i2) = s
End Sub

' Delegated properties
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = lst.BackColor
End Property

Public Property Let BackColor(ByVal clrBackColor As OLE_COLOR)
    lst.BackColor() = clrBackColor
    PropertyChanged "BackColor"
End Property

Property Get Columns() As Integer
    Columns = lst.Columns
End Property

Property Let Columns(iColumnsA As Integer)
    lst.Columns = iColumnsA
    PropertyChanged "Columns"
End Property

Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal fEnabled As Boolean)
    UserControl.Enabled() = fEnabled
    PropertyChanged "Enabled"
End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = lst.ForeColor
End Property

Public Property Let ForeColor(ByVal clrForeColor As OLE_COLOR)
    lst.ForeColor() = clrForeColor
    PropertyChanged "ForeColor"
End Property

Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = lst.Font
End Property

Public Property Set Font(ByVal fntFont As Font)
    Set lst.Font = fntFont
    PropertyChanged "Font"
End Property

Property Get Appearance() As EAppearance
    Appearance = eaAppearance
End Property

Property Let Appearance(eaAppearanceA As EAppearance)
With lst
    ' Can't do this:
    'lst.Appearance = eaAppearance
    ' Do this instead
    eaAppearance = eaAppearanceA
    DrawAppearance lst
    PropertyChanged "Appearance"
End With
End Property
    
Private Sub DrawAppearance(lst As ListBox)
With lst
    Dim rc As RECT
    
'    UserControl_Resize
'    rc.Left = 0
'    rc.Top = 0
'    rc.Right = myWidth
'    rc.bottom = myHeight
'    Dim iScaleOld As Integer
'    iScaleOld = .Parent.ScaleMode
'    .Parent.ScaleMode = vbPixels
'    If eaAppearance = eaFlat Then
'        'DrawEdge hDC, rc, EDGE_RAISED, BF_ADJUST
'        DrawEdge .Parent.hDC, rc, 0, BF_RECT
'    Else
'    Dim bdrFlags As Long, stlFlags As Long
'    bdrFlags = bdrFlags Or BDR_RAISEDOUTER
'    bdrFlags = bdrFlags Or BDR_RAISEDINNER
'    bdrFlags = bdrFlags Or BDR_RAISED
'    bdrFlags = bdrFlags Or BDR_SUNKEN
'    bdrFlags = bdrFlags Or BDR_SUNKENOUTER
'    bdrFlags = bdrFlags Or BDR_SUNKENINNER
'    DrawEdge .Parent.hDC, rc, EDGE_SUNKEN, BF_RECT
'    End If
'    .Parent.ScaleMode = iScaleOld
    Debug.Print "DrawAppearance"
End With
End Sub

Public Property Get hWnd() As Long
Attribute hWnd.VB_Description = "Returns a handle (from Microsoft Windows) to an object's window."
    hWnd = lst.hWnd
End Property

Public Property Get ItemData(Index As Integer) As Long
Attribute ItemData.VB_Description = "Returns/sets a specific number for each item in a ComboBox or ListBox control."
Attribute ItemData.VB_ProcData.VB_Invoke_Property = "List"
    ItemData = lst.ItemData(Index)
End Property

Public Property Let ItemData(Index As Integer, ByVal iItemData As Long)
    lst.ItemData(Index) = iItemData
    PropertyChanged "ItemData"
End Property

Public Property Get IntegralHeight() As Boolean
    'IntegralHeight = lst.IntegralHeight
    IntegralHeight = GetStyleBits(lst.hWnd) And LBS_NOINTEGRALHEIGHT
End Property

Public Property Let IntegralHeight(ByVal fIntegralHeight As Boolean)
    ' Can't do this:
    'lst.IntegralHeight = fIntegralHeight
    ' Do this instead
    ChangeStyleBit lst.hWnd, fIntegralHeight, LBS_NOINTEGRALHEIGHT
    PropertyChanged "IntegralHeight"
End Property

' For compatibility
Public Property Get List(Index As Integer) As String
Attribute List.VB_Description = "Returns/sets the items contained in a control's list portion."
Attribute List.VB_ProcData.VB_Invoke_Property = "List;List"
    List = lst.List(Index)
End Property

Public Property Let List(Index As Integer, ByVal sList As String)
    lst.List(Index) = sList
    PropertyChanged "List"
End Property

' For compatibility
Public Property Get ListCount() As Integer
Attribute ListCount.VB_Description = "Returns the number of items in the list portion of a control."
    ListCount = lst.ListCount
End Property

' For compatibility
Public Property Get ListIndex() As Integer
Attribute ListIndex.VB_Description = "Returns/sets the index of the currently selected item in the control."
Attribute ListIndex.VB_MemberFlags = "400"
    ListIndex = lst.ListIndex
End Property

Public Property Let ListIndex(ByVal iListIndex As Integer)
    lst.ListIndex() = iListIndex
    PropertyChanged "ListIndex"
End Property

Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "Sets a custom mouse icon."
    Set MouseIcon = lst.MouseIcon
End Property

Public Property Set MouseIcon(ByVal picMouseIcon As Picture)
    Set lst.MouseIcon = picMouseIcon
    PropertyChanged "MouseIcon"
End Property

Public Property Get MousePointer() As MousePointerConstants
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object."
    MousePointer = lst.MousePointer
End Property

Public Property Let MousePointer(ByVal ordMousePointer As MousePointerConstants)
    lst.MousePointer() = ordMousePointer
    PropertyChanged "MousePointer"
End Property

Public Property Get MultiSelect() As MultiSelectConstants
Attribute MultiSelect.VB_Description = "Returns/sets a value that determines whether a user can make multiple selections in a control."
    Dim af As Long
    af = GetStyleBits(lst.hWnd)
    If af And LBS_MULTIPLESEL Then
        MultiSelect = vbMultiSelectSimple
    ElseIf af And LBS_EXTENDEDSEL Then
        MultiSelect = vbMultiSelectExtended
    Else
        MultiSelect = vbMultiSelectNone
    End If
    'MultiSelect = lst.MultiSelect
End Property

Property Let MultiSelect(ordMultiSelectA As MultiSelectConstants)
'    lst.MultiSelect = ordMultiSelectA
    Select Case ordMultiSelectA
    Case vbMultiSelectNone
        ChangeStyleBit lst.hWnd, False, LBS_MULTIPLESEL
        ChangeStyleBit lst.hWnd, False, LBS_EXTENDEDSEL
    Case vbMultiSelectSimple
        ChangeStyleBit lst.hWnd, True, LBS_MULTIPLESEL
        ChangeStyleBit lst.hWnd, False, LBS_EXTENDEDSEL
    Case vbMultiSelectExtended
        ChangeStyleBit lst.hWnd, False, LBS_MULTIPLESEL
        ChangeStyleBit lst.hWnd, True, LBS_EXTENDEDSEL
    End Select
    lst.Refresh
    PropertyChanged "IntegralHeight"
End Property

Public Property Get NewIndex() As Integer
Attribute NewIndex.VB_Description = "Returns the index of the item most recently added to a control."
    NewIndex = lst.NewIndex
End Property

Public Property Get OLEDragMode() As Integer
Attribute OLEDragMode.VB_Description = "Returns/Sets whether this object can act as an OLE drag/drop source, and whether this process is started automatically or under programmatic control."
    OLEDragMode = lst.OLEDragMode
End Property

Public Property Let OLEDragMode(ByVal iOLEDragMode As Integer)
    lst.OLEDragMode() = iOLEDragMode
    PropertyChanged "OLEDragMode"
End Property

Public Property Get OLEDropMode() As Integer
Attribute OLEDropMode.VB_Description = "Returns/Sets whether this object can act as an OLE drop target."
    OLEDropMode = lst.OLEDropMode
End Property

Public Property Let OLEDropMode(ByVal iOLEDropMode As Integer)
    lst.OLEDropMode() = iOLEDropMode
    PropertyChanged "OLEDropMode"
End Property

Public Property Get RightToLeft() As Boolean
Attribute RightToLeft.VB_Description = "Determines text display direction and control visual appearance on a bidirectional system."
    RightToLeft = lst.RightToLeft
End Property

Public Property Let RightToLeft(ByVal fRightToLeft As Boolean)
    lst.RightToLeft() = fRightToLeft
    PropertyChanged "RightToLeft"
End Property

Public Property Get SelCount() As Integer
Attribute SelCount.VB_Description = "Returns the number of selected items in a ListBox control."
    SelCount = lst.SelCount
End Property

Public Property Get Selected(Index As Integer) As Boolean
Attribute Selected.VB_Description = "Returns/sets the selection status of an item in a control."
    Selected = lst.Selected(Index)
End Property

Public Property Let Selected(Index As Integer, ByVal fSelected As Boolean)
    lst.Selected(Index) = fSelected
    PropertyChanged "Selected"
End Property

Public Property Get Text() As String
Attribute Text.VB_Description = "Returns/sets the text contained in the control."
Attribute Text.VB_MemberFlags = "424"
    Text = lst.Text
End Property

Public Property Let Text(ByVal sText As String)
    lst.Text() = sText
    PropertyChanged "Text"
End Property

Public Property Get TopItem() As Integer
Attribute TopItem.VB_Description = "Returns/sets which item in a control is displayed in the topmost position."
Attribute TopItem.VB_MemberFlags = "400"
    TopItem = lst.TopIndex
End Property

Public Property Let TopItem(ByVal iTopItem As Integer)
    lst.TopIndex() = iTopItem
    PropertyChanged "TopItem"
End Property

' Delegated methods

Public Sub OLEDrag()
Attribute OLEDrag.VB_Description = "Starts an OLE drag/drop event with the given control as the source."
    lst.OLEDrag
End Sub

' Event delegators
Private Sub lst_Click()
    RaiseEvent Click
End Sub

Private Sub lst_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub lst_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub lst_KeyPress(KeyAscii As Integer)
    Static msPrevKeyPress As Long
    Dim msCurKeyPress As Long, iKey As Integer
    iKey = KeyAscii
    If fCompletion Then
        msCurKeyPress = GetTickCount
        ' Check time between keypresses
        If (msCurKeyPress - msPrevKeyPress) >= 1000 Then
            sPartial = sEmpty
        End If
        ' Handle special case keys
        Select Case iKey
        Case vbKeyBack
            ' Handle backspace
            If Len(sPartial) Then
                PartialWord = Left$(sPartial, Len(sPartial) - 1)
            End If
        Case Is >= vbKeySpace
            ' For ASCII keys add keystroke to current partial word
            PartialWord = sPartial & Chr$(iKey)
        Case Else
            ' Ignore other control keys
        End Select
        msPrevKeyPress = msCurKeyPress
        ' Default text box behavior interferes with
        ' word completion, so throw away all keypresses
        KeyAscii = 0
    End If
    RaiseEvent KeyPress(iKey)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    'RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub lst_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub lst_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub lst_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseMove(Button, Shift, x, y)
End Sub

Private Sub lst_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

Private Sub lst_ItemCheck(item As Integer)
    RaiseEvent ItemCheck(item)
End Sub

Private Sub lst_OLECompleteDrag(Effect As Long)
    RaiseEvent OLECompleteDrag(Effect)
End Sub

Private Sub lst_OLEDragDrop(data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent OLEDragDrop(data, Effect, Button, Shift, x, y)
End Sub

Private Sub lst_OLEDragOver(data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
    RaiseEvent OLEDragOver(data, Effect, Button, Shift, x, y, State)
End Sub

Private Sub lst_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
    RaiseEvent OLEGiveFeedback(Effect, DefaultCursors)
End Sub

Private Sub lst_OLESetData(data As DataObject, DataFormat As Integer)
    RaiseEvent OLESetData(data, DataFormat)
End Sub

Private Sub lst_OLEStartDrag(data As DataObject, AllowedEffects As Long)
    RaiseEvent OLEStartDrag(data, AllowedEffects)
End Sub

Private Sub lst_Scroll()
    RaiseEvent Scroll
End Sub

Private Sub ErrRaise(e As Long)
    Dim sText As String, sSource As String
    If e > 1000 Then
        sSource = App.ExeName
        Select Case e
        Case eseNone
            BugAssert True
        Case eseItemNotFound
            sText = "Item not in list"
        Case eseOutOfRange
            sText = "Index out of range"
        Case eseDuplicateNotAllowed
            sText = "Duplicate entries not allowed"
        End Select
        Err.Raise COMError(e), sSource, sText
    Else
        ' Raise standard Visual Basic error
        sSource = App.ExeName & ".VBError"
        Err.Raise e, sSource
    End If
End Sub

'Private Function GetRandom(ByVal iLo As Long, ByVal iHi As Long) As Long
'    GetRandom = Int(iLo + (Rnd * (iHi - iLo + 1)))
'End Function

