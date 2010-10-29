VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CRegNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
   
Public Enum EErrorRegNode
    eeBaseRegNode = 13180   ' CRegNode
End Enum

Private hKey As Long
Private sName As String
Private afAccess As Long
Private cItem As Long, cNode As Long

' Create
'RegOpenKeyEx
'RegCloseKey

' Connect
'RegConnectRegistry

' Enumerate
'RegEnumKeyEx

Private Sub Class_Initialize()
    afAccess = KEY_ALL_ACCESS
    hKey = HKEY_CURRENT_USER
    cItem = -1
    cNode = -1
End Sub

Sub Create(vIndex As Variant, _
           Optional RootKey As Long = HKEY_CURRENT_USER, _
           Optional AccessRights As Long = KEY_ALL_ACCESS)
    Dim e As Long, hKeyRoot As Long
    hKeyRoot = RootKey
    afAccess = AccessRights
    Select Case VarType(vIndex)
    Case vbString
        sName = vIndex
        ' Key is a key name
        e = RegOpenKeyEx(hKeyRoot, sName, 0&, afAccess, hKey)
        ApiRaiseIf e
        
    Case vbInteger, vbLong
        ' Index is a handle
        e = MRegTool.GetRegNodeNext(hKeyRoot, CLng(vIndex), sName)
        ApiRaiseIf e
    
    Case Else
        ApiRaise ERROR_INVALID_DATA
        
    End Select

End Sub

' Destroy
'RegCloseKey

Private Sub Class_Terminate()
    If hKey = 0 Then Exit Sub
    ApiRaiseIf RegCloseKey(hKey)
End Sub

' NewEnum must have the procedure ID -4 in Procedure Attributes dialog
' Create a new data walker object and connect to it
Public Function NewEnum() As IEnumVARIANT
Attribute NewEnum.VB_UserMemId = -4
    ' Create a new iterator object
    Dim walker As CRegNodeWalker
    Set walker = New CRegNodeWalker
    ' Connect it with collection data
    walker.Attach Me
    ' Return it
    Set NewEnum = walker.NewEnum
End Function

'!Public Property Get Item(vIndex As Variant) '! As DataType
    '!Set Item = data(vIndex)
'!End Property

Property Get Name() As String
    Name = sName
End Property

Friend Property Let Name(sNameA As String)
    sName = sNameA
End Property

Property Get Key(Optional RootKey As Long = -1) As Variant
    ' Ignore RootKey
    Key = hKey
End Property

' Ways of initializing a node
' Open HKEY_CLASSES_ROOT key
'   node.Key = HKEY_CLASSES_ROOT
' Open Software key
'   node.Key(HKEY_LOCAL_MACHINE) = "Software"
' Open Software key in default root HKEY_CURRENT_USER
'   node.Key = "Software"
' Open next key of current key
'   node.Key(hKey) = iNode

Property Let Key(Optional RootKey As Long = -1, vKey As Variant)
    Dim e As Long
    If VarType(vKey) = vbString Then
        If RootKey = -1 Then RootKey = hKey
        sName = vKey
        ' Key is a key name
        e = RegOpenKeyEx(RootKey, sName, pNull, afAccess, hKey)
    Else
        If RootKey = -1 Then
            ' Ignore any hRootKey, key is the root handle
            hKey = CLng(vKey)
            e = RegOpenKeyEx(hKey, sNullStr, pNull, afAccess, hKey)
            Select Case hKey
            Case HKEY_CLASSES_ROOT
                sName = "HKEY_CLASSES_ROOT"
            Case HKEY_CURRENT_USER
                sName = "HKEY_CURRENT_USER"
            Case HKEY_LOCAL_MACHINE
                sName = "HKEY_LOCAL_MACHINE"
            Case HKEY_USERS
                sName = "HKEY_USERS"
            Case HKEY_CURRENT_CONFIG
                sName = "HKEY_CURRENT_CONFIG"
            Case HKEY_PERFORMANCE_DATA
                sName = "HKEY_PERFORMANCE_DATA"
            Case HKEY_DYN_DATA
                sName = "HKEY_DYN_DATA"
            Case Else
                ' Query for name
            End Select
        Else
            ' Key is the iteration index
            e = MRegTool.GetRegNodeNext(RootKey, CLng(vKey), sName)
            ApiRaiseIf e
            e = RegOpenKeyEx(RootKey, sName, pNull, afAccess, hKey)
        End If
    End If
    ApiRaiseIf e
End Property

' Only for use by AddNode
Friend Function SetKey(hKeyA As Long)
    hKey = hKeyA
End Function

Property Get Access() As Long
    Access = afAccess
End Property

Property Let Access(afAccessA As Long)
    ' Mask out any invalid flags
    afAccess = afAccessA And KEY_ALL_ACCESS
End Property

' Get/Set Property
'RegQueryValue
'RegQueryValueEx
'RegSetValue
'RegSetValueEx

Property Get Value() As Variant
    Value = Items(sEmpty).Value
End Property

Property Let Value(vValueA As Variant)
    Items(sEmpty).Value = vValueA
End Property

' Summary
'RegQueryInfoKey

Property Get ItemCount() As Long
    If cItem <> -1 Then
        ItemCount = cItem
    Else
        Dim e As Long, sJunk As String, cJunk As Long, cClsNM As Long
        Dim cNodeNM As Long, cItemNM As Long, cItemM As Long
        Dim ft As FILETIME
        e = RegQueryInfoKey(hKey, sNullStr, cJunk, pNull, cNode, _
                            cNodeNM, cJunk, cItem, cItemNM, cItemM, _
                            cJunk, ft)
        ApiRaiseIf e
        ItemCount = cItem
    End If
End Property

Property Get NodeCount() As Long
    If cNode <> -1 Then
        NodeCount = cNode
    Else
        Dim e As Long, sJunk As String, cJunk As Long, cClsNM As Long
        Dim cNodeNM As Long, cItemNM As Long, cItemM As Long
        Dim ft As FILETIME
        e = RegQueryInfoKey(hKey, sNullStr, cJunk, pNull, _
                            cNode, cNodeNM, cJunk, _
                            cItem, cItemNM, cItemM, cJunk, ft)
        ApiRaiseIf e
        NodeCount = cNode
    End If
End Property

' Nodes and items

Function Nodes(vIndexA As Variant) As CRegNode
Attribute Nodes.VB_UserMemId = 0
    Dim node As CRegNode
    Set node = New CRegNode
    ' Pass parent access to child
    node.Access = afAccess
    node.Key(hKey) = vIndexA
    Set Nodes = node
End Function

Function Items(Optional Index As Variant) As CRegItem
    Dim item As CRegItem
    Set item = New CRegItem
    If IsMissing(Index) Then Index = sEmpty
    item.Create Me, Index
    Set Items = item
End Function

Friend Function BaseNode() As CRegNode
    Set BaseNode = Me
End Function

' Add/Delete Node
'RegCreateKey
'RegCreateKeyEx
'RegDeleteKey
Function AddNode(sNameA As String) As CRegNode
    Dim e As Long, ordResult As Long, hKeyRes As Long
    e = RegCreateKeyEx(hKey, sNameA, 0&, sEmpty, REG_OPTION_NON_VOLATILE, _
                       afAccess, ByVal pNull, hKeyRes, ordResult)
    ApiRaiseIf e
    Dim nodeT As CRegNode
    Set nodeT = New CRegNode
    nodeT.Access = afAccess
    nodeT.SetKey hKeyRes
    nodeT.Name = sNameA
    Set AddNode = nodeT
End Function

Function RemoveNode(vNode As Variant, _
                    Optional AllChild = True) As Boolean
    Dim e As Long
    If AllChild Then
        e = MRegTool.DeleteRegNodes(hKey, Nodes(vNode).Name)
    Else
        e = MRegTool.DeleteOneRegNode(hKey, Nodes(vNode).Name)
    End If
    ' Translate failure to delete node with children to false return
    Select Case e
    Case 0
        RemoveNode = True
    Case ERROR_ACCESS_DENIED
        RemoveNode = False
    Case Else
        ApiRaiseIf e
    End Select
End Function

' Add/Delete Item
'RegSetValueEx
'RegDeleteValue
Sub AddItem(vValue As Variant, Optional ItemName As String)
    Dim e As Long
    e = MRegTool.CreateRegValue(vValue, hKey, ItemName)
    ApiRaiseIf e
End Sub

Sub RemoveItem(Optional Index As Variant)
    Dim e As Long, sName As String
    sName = Items(Index).Name
    e = RegDeleteValue(hKey, sName)
    ApiRaiseIf e
End Sub

Function WalkNodes(use As IUseRegItems, _
                   ByVal iLevel As Long) As CRegNode
#If 1 Then
    ' This way is faster, but the other also works
    Dim i As Long
    For i = 0 To NodeCount - 1
        If use.UseNode(Nodes(i), iLevel + 1) Then
            Set WalkNodes = Nodes(i)
            Exit Function
        End If
    Next
#Else
    Dim node As CRegNode
    For Each node In Me
        If use.UseNode(node, iLevel + 1) Then
            Set WalkNodes = node
            Exit Function
        End If
    Next
#End If
End Function

Function WalkItems(use As IUseRegItems, _
                   ByVal iLevel As Long) As CRegItem
    Dim i As Long
    For i = 0 To ItemCount - 1
        If use.UseItem(Items(i), iLevel + 1) Then
            Set WalkItems = Items(i)
            Exit Function
        End If
    Next
End Function

Function WalkAllNodes(use As IUseRegItems, nodeStart As CRegNode, _
                      ByVal iLevel As Long) As CRegNode
    If use.UseNode(nodeStart, iLevel) Then
        Set WalkAllNodes = nodeStart
        Exit Function
    End If
    Dim i As Long, nodeT As CRegNode
    ' Iterate by index for greater speed
    For i = 0 To nodeStart.NodeCount - 1
        Set WalkAllNodes = WalkAllNodes(use, nodeStart.Nodes(i), _
                                        iLevel + 1)
    Next
End Function

' Write to and from file
'RegSaveKey
'RegRestoreKey
'RegReplaceKey
'RegLoadKey
'RegUnLoadKey

#If fComponent = 0 Then
Private Sub ErrRaise(e As Long)
    Dim sText As String, sSource As String
    If e > 1000 Then
        sSource = App.ExeName & ".RegNode"
        Select Case e
        Case eeBaseRegNode
            BugAssert True
       ' Case ee...
       '     Add additional errors
        End Select
        Err.Raise COMError(e), sSource, sText
    Else
        ' Raise standard Visual Basic error
        sSource = App.ExeName & ".VBError"
        Err.Raise e, sSource
    End If
End Sub
#End If

