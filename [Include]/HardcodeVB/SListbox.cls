VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSortedListBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'$ Uses SLISTBOX.BAS UTILITY.BAS

Public Enum ESortError
    eseNone = 0
    eseItemNotFound = 1
    eseOutOfRange = 2
    eseDuplicateNotAllowed = 3
    eseNoSelection = 4
    eseNotListBox = 5
    eseAlreadySorted = 6
    eseUnknown = 7
End Enum

Private lst As Control
Private esmMode As ESortMode
Private fHiToLo As Boolean

Private asError(0 To 7) As String

Public LastError As Integer
  
''' Public Methods Unique to This Class '''
    
' Create sorted list box from unsorted list box or similar control
Function Create(lstA As Control, Optional fHiToLoA As Boolean = False, _
                Optional esmModeA As ESortMode = esmUnsorted) As Boolean
    
    ' Must have ListCount, List, ListIndex, AddItem, and RemoveItem
    On Error GoTo CreateFail
    Dim v As Variant
    LastError = eseNotListBox
    v = lstA.ListCount
    lstA.AddItem v, 0
    v = lstA.list(0)
    lstA.list(0) = v
    v = lstA.ListIndex
    lstA.RemoveItem 0
    ' Make sure list isn't already sorted
    LastError = eseAlreadySorted
    If lstA.Sorted Then GoTo CreateFail
        
    ' Initialize internal data
    LastError = eseUnknown
    fHiToLo = fHiToLoA
    esmMode = esmSortVal
    If esmModeA <> -1 Then esmMode = esmModeA
    Set lst = lstA
    ' Sort it
    Sort 0, lst.ListCount - 1
    Create = True
    LastError = eseNone
    Exit Function
    
CreateFail:
    Create = False
End Function

' Note that our AddItem only takes an item. You cannot specify the
' insert position with a sorted list as you can with an unsorted list.
' Also, you cannot insert an item that already exists into a sorted
' list. A request to do so will be ignored.
Sub AddItem(sItem As String)
    LastError = eseNone
    ' Binary search for the item
    Dim iPos As Integer
    If BSearch(sItem, iPos) Then
        LastError = eseDuplicateNotAllowed
    Else
        lst.AddItem sItem, iPos
    End If
End Sub

' RemoveItem takes numeric 0-based item index
Sub RemoveItem(iItem As Integer)
    lst.RemoveItem iItem
End Sub

' Collection Methods
' Same as AddItem but with collection name (but not extra arguments)
Sub Add(vItem As Variant)
    LastError = eseNone
    ' Binary search for the item
    Dim iPos As Integer
    If BSearch(vItem, iPos) Then
        LastError = eseDuplicateNotAllowed
    Else
        lst.AddItem vItem, iPos - 1
    End If
End Sub

' Same as RemoveItem but has collection name and is 1-based
Sub Remove(vItem As Variant)
    LastError = eseNone
    If IsNumeric(vItem) Then
        If vItem > Count Or vItem < 1 Then LastError = eseOutOfRange: Exit Sub
    Else
        vItem = Match(vItem)
        If vItem = 0 Then LastError = eseItemNotFound: Exit Sub
    End If
    lst.RemoveItem vItem - 1
End Sub

' Similar to List property
Property Get Item(vIndex As Variant) As Variant
    LastError = eseNone
    If IsNumeric(vIndex) Then
        ' For numeric index, return string value
        Item = lst.list(vIndex - 1)
    Else
        ' For string index, return matching index or 0 for none
        Item = Match(vIndex)
        If Item = 0 Then LastError = eseItemNotFound
    End If
End Property

Property Let Item(vIndex As Variant, vItemA As Variant)
    LastError = eseNone
    ' For string index, look up matching index
    If Not IsNumeric(vIndex) Then
        vIndex = Match(vIndex)
        ' Quit if old item isn't found or if new item is found
        If vIndex = 0 Then
            LastError = eseItemNotFound
            Exit Property
        End If
        If Match(vItemA) Then
            LastError = eseDuplicateNotAllowed
            Exit Property
        End If
    End If
    ' Assign value by removing old and inserting new
    Remove vIndex
    Add vItemA
End Property

''' Public Properties Unique to This Class '''

Property Let HiToLo(fHiToLoA As Boolean)
    fHiToLo = fHiToLoA
    Sort 0, lst.ListCount - 1
End Property

Property Get HiToLo() As Boolean
    HiToLo = fHiToLo
End Property

Property Let SortMode(esmModeA As ESortMode)
    esmMode = esmModeA
    Sort 0, lst.ListCount - 1
End Property

Property Get SortMode() As ESortMode
    SortMode = esmMode
End Property

' Gives away the store for iteration
Property Get Items() As Collection
    Set Items = lst
End Property

' Collection name
Property Get Count() As Integer
    Count = lst.ListCount
End Property

' Index replaces Index property of a control array--if you really need
' to manage control-array at run-time, use Index of external list box
Property Get Index() As Variant
    LastError = eseNone
    Index = lst.ListIndex + 1
End Property

Property Let Index(vIndexA As Variant)
    LastError = eseNone
    If IsNumeric(vIndexA) Then
        lst.ListIndex = vIndexA - 1
    Else
        lst.ListIndex = Match(vIndexA) - 1
    End If
    If lst.ListIndex = -1 Then LastError = eseItemNotFound
End Property

Property Get IndexItem() As Variant
    LastError = eseNone
    IndexItem = lst.list(lst.ListIndex)
End Property

' 1-based versions of ItemData
Property Get Data(i As Integer) As Variant
    Data = lst.ItemData(i - 1)
End Property

Property Let Data(i As Integer, vDataA As Variant)
    lst.Data(i - 1) = vDataA
End Property

Property Get LastErrorStr() As String
    LastErrorStr = asError(LastError)
End Property


''' Public Methods From Contained Class '''
    
Sub Clear()
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

Sub Refresh()
    Sort 0, lst.ListCount - 1
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

''' Public Properties From Contained Class '''

Property Get BackColor() As Long
    BackColor = lst.BackColor
End Property

Property Let BackColor(iBackColorA As Long)
    lst.BackColor = iBackColorA
End Property

Property Get Columns() As Integer
    Columns = lst.Columns
End Property

Property Let Columns(iColumnsA As Integer)
    lst.Columns = iColumnsA
End Property

Property Get Enabled() As Boolean
    Enabled = lst.Enabled
End Property

Property Let Enabled(fEnabledA As Boolean)
    lst.Enabled = fEnabledA
End Property

Property Get ForeColor() As Long
    ForeColor = lst.ForeColor
End Property

Property Let ForeColor(iForeColorA As Long)
    lst.ForeColor = iForeColorA
End Property

Property Get Height() As Single
    Height = lst.Height
End Property

Property Let Height(rHeightA As Single)
    lst.Height = rHeightA
End Property

Property Get HelpContextID() As Integer
    HelpContextID = lst.HelpContextID
End Property

Property Let HelpContextID(iHelpContextIDA As Integer)
    lst.HelpContextID = iHelpContextIDA
End Property

Property Get hWnd() As Integer
    hWnd = lst.hWnd
End Property

Property Let hWnd(hWndA As Integer)
    lst.hWnd = hWndA
End Property

Property Get ItemData(i As Integer) As Variant
    ItemData = lst.ItemData(i)
End Property

Property Let ItemData(i As Integer, vItemDataA As Variant)
    lst.ItemData(i) = vItemDataA
End Property

Property Get Left() As Single
    Left = lst.Left
End Property

Property Let Left(rLeftA As Single)
    lst.Left = rLeftA
End Property

Property Get list(iIndex As Integer) As String
    list = lst.list(iIndex)
End Property

Property Let list(iIndex As Integer, sListA As String)
    lst.list(iIndex) = sListA
End Property

Property Get ListCount() As Integer
    ListCount = lst.ListCount
End Property

Property Get ListIndex() As Integer
    ListIndex = lst.ListIndex
End Property

Property Let ListIndex(iListIndexA As Integer)
    lst.ListIndex = iListIndexA
End Property

Property Get MousePointer() As Integer
    MousePointer = lst.MousePointer
End Property

Property Let MousePointer(iMousePointerA As Integer)
    lst.MousePointer = iMousePointerA
End Property

Property Get MultiSelect() As Integer
    MultiSelect = lst.MultiSelect
End Property

Property Let MultiSelect(iMultiSelectA As Integer)
    lst.MultiSelect = iMultiSelectA
End Property

Property Get NewIndex() As Integer
    NewIndex = lst.NewIndex
End Property

Property Get Parent() As Form
    Set Parent = lst.Parent
End Property

Property Get Selected(i As Integer) As Boolean
    Selected = lst.Selected(i)
End Property

Property Let Selected(i As Integer, fSelectedA As Boolean)
    lst.Selected(i) = fSelectedA
End Property

Property Get TabIndex() As Integer
    TabIndex = lst.TabIndex
End Property

Property Let TabIndex(iTabIndexA As Integer)
    lst.TabIndex = iTabIndexA
End Property

Property Get TabStop() As Boolean
    TabStop = lst.TabStop
End Property

Property Let TabStop(fTabStopA As Boolean)
    lst.TabStop = fTabStopA
End Property

Property Get Tag() As String
    Tag = lst.Tag
End Property

Property Let Tag(sTagA As String)
    lst.Tag = sTagA
End Property

Property Get Text() As String
    Text = lst.Text
End Property

Property Let Text(sTextA As String)
    lst.Text = sTextA
End Property

Property Get Top() As Single
    Top = lst.Top
End Property

Property Let Top(rTopA As Single)
    lst.Top = rTopA
End Property

Property Get TopIndex() As Integer
    TopIndex = lst.TopIndex
End Property

Property Let TopIndex(iTopIndexA As Integer)
    lst.TopIndex = iTopIndexA
End Property

Property Get Visible() As Boolean
    Visible = lst.Visible
End Property

Property Let Visible(fVisibleA As Boolean)
    lst.Visible = fVisibleA
End Property

Property Get Width() As Single
    Width = lst.Width
End Property

Property Let Width(rWidthA As Single)
    lst.Width = rWidthA
End Property

''' Private Procedures Used by Class '''

Private Function Match(ByVal sItem As String) As Long
    Dim iPos As Integer
    If BSearch(sItem, iPos) Then Match = iPos + 1 Else Match = 0
End Function

Private Sub Sort(iFirst As Integer, iLast As Integer)
    Dim vSplit As Variant
    Static fRand As Integer
    If fRand = False Then
        Randomize
        fRand = True
    End If

    If iFirst < iLast Then

        ' Only two elements in this subdivision. Exchange if
        ' they are out of order, and end recursive calls.
        If iLast - iFirst = 1 Then
            If Compare(lst.list(iFirst), lst.list(iLast)) > 0 Then
                Swap iFirst, iLast
            End If
        Else

            Dim i As Integer, j As Integer, iRand As Integer

            ' Pick pivot element at random and move to end
            ' (consider calling Randomize before sorting)
            iRand = GetRandom(iFirst, iLast)
            Swap iLast, iRand
            vSplit = lst.list(iLast)
            Do

                ' Move in from both sides towards the pivot element
                i = iFirst: j = iLast
                Do While (i < j) And _
                    Compare(lst.list(i), vSplit) <= 0
                    i = i + 1
                Loop
                Do While (j > i) And _
                    Compare(lst.list(j), vSplit) >= 0
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

Private Function BSearch(vKey As Variant, iPos As Integer) As Boolean
    Dim iLo As Integer, iHi As Integer, iComp As Integer, iMid As Integer
    iLo = 0: iHi = lst.ListCount - 1
    Do
        iMid = iLo + ((iHi - iLo) \ 2)
        iComp = Compare(lst.list(iMid), vKey)
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

Private Function Compare(v1 As Variant, v2 As Variant) As Integer
    Dim i As Integer
    If IsNumeric(v1) And IsNumeric(v2) Then v1 = Val(v1): v2 = Val(v2)
    
    Select Case esmMode
    ' Sort by value (same as esmSortBin for strings)
    Case esmSortVal
        If v1 < v2 Then
            i = -1
        ElseIf v1 = v2 Then
            i = 0
        Else
            i = 1
        End If
    ' Sort case-insensitive
    Case esmSortText
        i = StrComp(v1, v2, 1)
    ' Sort case-sensitive
    Case esmSortbin
        i = StrComp(v1, v2, 0)
    ' Sort by string length
    Case esmSortLen
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
    Dim vT As Variant
    vT = lst.list(i1)
    lst.list(i1) = lst.list(i2)
    lst.list(i2) = vT
End Sub

Private Sub Class_Initialize()
    Const sNone = ""
    Const sItemNotFound = "Item not found"
    Const sOutOfRange = "Out of range"
    Const sDuplicateNotAllowed = "Duplicate item not allowed"
    Const sNoSelection = "Nothing selected"
    Const sNotListBox = "Lacks list box methods or properties"
    Const sAlreadySorted = "List box already sorted"
    Const sUnknown = "Unknown error"
    
    asError(eseNone) = sNone
    asError(eseItemNotFound) = sItemNotFound
    asError(eseOutOfRange) = sOutOfRange
    asError(eseDuplicateNotAllowed) = sDuplicateNotAllowed
    asError(eseNoSelection) = sNoSelection
    asError(eseNotListBox) = sNotListBox
    asError(eseAlreadySorted) = sAlreadySorted
    asError(eseUnknown) = sUnknown
End Sub

Private Sub Class_Terminate()
    Set lst = Nothing
End Sub
