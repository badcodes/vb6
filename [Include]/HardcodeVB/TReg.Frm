VERSION 5.00
Begin VB.Form FTestRegistry 
   Caption         =   "Test Registry"
   ClientHeight    =   5952
   ClientLeft      =   996
   ClientTop       =   2520
   ClientWidth     =   8280
   Icon            =   "TReg.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5952
   ScaleWidth      =   8280
   Begin VB.ListBox lstItem 
      Height          =   816
      ItemData        =   "TReg.frx":0CFA
      Left            =   144
      List            =   "TReg.frx":0CFC
      TabIndex        =   10
      Top             =   4224
      Width           =   1596
   End
   Begin VB.ListBox lstNode 
      Height          =   816
      ItemData        =   "TReg.frx":0CFE
      Left            =   144
      List            =   "TReg.frx":0D05
      TabIndex        =   9
      Top             =   2976
      Width           =   1596
   End
   Begin VB.CommandButton cmdTestClass 
      Caption         =   "Test All Classes"
      Height          =   396
      Left            =   156
      TabIndex        =   7
      Top             =   1080
      Width           =   1596
   End
   Begin VB.CommandButton cmdTestFunc 
      Caption         =   "Test All Functions"
      Height          =   396
      Left            =   156
      TabIndex        =   6
      Top             =   564
      Width           =   1596
   End
   Begin VB.ListBox lstRoot 
      Height          =   816
      ItemData        =   "TReg.frx":0D12
      Left            =   144
      List            =   "TReg.frx":0D19
      TabIndex        =   4
      Top             =   1812
      Width           =   1590
   End
   Begin VB.TextBox txtValue 
      Height          =   285
      Left            =   144
      TabIndex        =   2
      Top             =   5448
      Width           =   1596
   End
   Begin VB.TextBox txtOut 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5688
      Left            =   1932
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   96
      Width           =   6168
   End
   Begin VB.CommandButton cmdIterate 
      Caption         =   "Iterate Node"
      Height          =   396
      Left            =   156
      TabIndex        =   0
      Top             =   96
      Width           =   1584
   End
   Begin VB.Label lbl 
      Caption         =   "Item:"
      Height          =   252
      Index           =   4
      Left            =   144
      TabIndex        =   11
      Top             =   3996
      Width           =   972
   End
   Begin VB.Label lbl 
      Caption         =   "Node:"
      Height          =   252
      Index           =   3
      Left            =   144
      TabIndex        =   8
      Top             =   2736
      Width           =   972
   End
   Begin VB.Label lbl 
      Caption         =   "Root:"
      Height          =   252
      Index           =   2
      Left            =   144
      TabIndex        =   5
      Top             =   1572
      Width           =   972
   End
   Begin VB.Label lbl 
      Caption         =   "Value:"
      Height          =   252
      Index           =   1
      Left            =   144
      TabIndex        =   3
      Top             =   5184
      Width           =   972
   End
End
Attribute VB_Name = "FTestRegistry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements IUseRegItems
Private sOut As String
Private hRootCur As Long
Private nodeCur As New CRegNode, nodeRoot As New CRegNode
Private itemCur As New CRegItem
Private valCur As Variant

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdIterate_Click()
    Dim node As New CRegNode, f As Boolean
    txtOut.Text = sEmpty
    sOut = sEmpty
    HourGlass Me
    On Error GoTo IterateFail
    node.Key(hRootCur) = nodeCur.Name
    node.WalkAllNodes Me, node, 0
    txtOut = sOut
    HourGlass Me
    Exit Sub
IterateFail:
    txtOut = "Can't iterate item: " & Err.Description
    HourGlass Me
End Sub

Private Sub cmdTestClass_Click()
    Dim hKey As Long, hSubKey As Long, hSubSubKey As Long
    Dim ordDispose As Long, s As String, e As Long
    Dim node As CRegNode
    Dim nodesTop As New CRegNode
    ' Connect to first-level node by name
    nodesTop.Create "Software\VB and VBA Program Settings"
    ' Connect HKEY_CLASSES_ROOT node
    nodesTop.Key = HKEY_CLASSES_ROOT
    ' Connect VBCore.CAbout node in current node (HKEY_CLASSES_ROOT)
    nodesTop.Key = "VisualCore.CAbout"
    ' Connect Software node in specified root HKEY_LOCAL_MACHINE
    nodesTop.Key(HKEY_LOCAL_MACHINE) = "Software"
    ' Open first node of current node
    nodesTop.Key(nodesTop.Key) = 1
    ' Connect to first-level node by name
    nodesTop.Create "Software\VB and VBA Program Settings"

    s = s & "Opened VB and VBA node" & sCrLf
    ' Add a node
    Set node = nodesTop.AddNode("FirstLevel")
    s = s & "Created new FirstLevel key" & sCrLf
    ' Add two node to that node
    node.AddNode "SecondLevel1"
    s = s & "Created new SecondLevel1 key" & sCrLf
    node.AddNode "SecondLevel2"
    s = s & "Created new SecondLevel2 key" & sCrLf
    ' Add a default item (must be a string)
    node.AddItem "Default"
    s = s & "Created value: default" & sCrLf
    Dim ab() As Byte
    ' Add bytes item
    ab = "The bytes"
    node.AddItem ab, "Bytes"
    s = s & "Created value: Bytes" & sCrLf
    ' Add string item
    node.AddItem "A String", "String"
    s = s & "Created value: String" & sCrLf
    ' Add numeric item
    node.AddItem 5&, "Number"
    s = s & "Created value: Number" & sCrLf
    ' Add string item containing embedded environment variable
    node.AddItem "A %TEMP% string", "ExpandString"
    s = s & "Created value: ExpandString" & sCrLf
    
    Dim v As Variant
    ' Get default item
    v = node.Items(sEmpty)
    s = s & "Get default: " & v & sCrLf
    ' Get Bytes item
    v = node.Items("Bytes")
    ab = v
    s = s & "Get Bytes: " & HexDump(ab, ehdSample8) & sCrLf
    ' Get String item
    v = node.Items("String")
    s = s & "Get String: " & v & sCrLf
    ' Get Number item
    v = node.Items("Number")
    s = s & "Get Number: " & v & sCrLf
    v = node.Items("ExpandString")
    ' Get item with environment variable in string
    s = s & "Get ExpandString: " & v & sCrLf
    
    v = node.Items(1)
    s = s & "Get unknown item: " & VarToStr(v) & sCrLf
    
    ' Add some more values by string index
    node("SecondLevel1").AddItem "DefaultString"
    node("SecondLevel1").AddItem "Stuff", "Value1"
    node("SecondLevel2").AddItem 689, "Value1"
    
    ' Iterate node by numeric index
    Dim i As Long, sName As String
    For i = 0 To node.NodeCount - 1
        sName = node.Nodes(i).Name
        s = s & "Node(" & i & "): " & sName & sCrLf
    Next
    
    ' Iterate items by numeric index
    For i = 0 To node.ItemCount - 1
        With node.Items(i)
            s = s & .Name & "(" & i & ") = " & VarToStr(.Value) & sCrLf
        End With
    Next
    
    ' Iterate node with For Each
    s = s & node.Name & sCrLf
    Dim item As CRegItem
    For Each item In node.Items
        s = s & item.Name & " : " & VarToStr(item.Value) & sCrLf
    Next
    Dim nodesSub As CRegNode
    ' Iterate subnodes with For Each
    For Each nodesSub In node
        s = s & nodesSub.Name & sCrLf
        ' Iterate items with For Each
        For Each item In nodesSub.Items
            s = s & item.Name & " = " & VarToStr(item.Value) & sCrLf
        Next
    Next
    
    node.WalkNodes Me, 0
    s = s & sOut
    sOut = sEmpty
    node.WalkItems Me, 0
    s = s & sOut
    sOut = sEmpty
    Call node.WalkAllNodes(Me, node, 0)
    s = s & sOut
    sOut = sEmpty
    
    node.RemoveItem 1
    node.RemoveItem "String"
    
    Dim f As Boolean
    f = nodesTop.RemoveNode("FirstLevel", AllChild:=False)
    s = s & "Delete one node succeeded: " & f & sCrLf
                         
    f = nodesTop.RemoveNode("FirstLevel")
    s = s & "Delete all nodes succeeded: " & f & sCrLf
    
    BugMessage s
    txtOut = s

End Sub

Private Sub cmdTestFunc_Click()
    Dim hKey As Long, hSubKey As Long, hSubSubKey As Long
    Dim ordDispose As Long, s As String, e As Long
    e = RegOpenKeyEx(HKEY_CURRENT_USER, _
                     "Software\VB and VBA Program Settings", _
                     0&, KEY_ALL_ACCESS, hKey)
    If e Then GoTo RegFail
    s = s & "Opened VB and VBA key" & sCrLf
    e = RegCreateKeyEx(hKey, "FirstLevel", 0&, sEmpty, _
                       REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, _
                       ByVal pNull, hSubKey, ordDispose)
    If e Then GoTo RegFail
    e = RegCloseKey(hKey)
    If ordDispose = REG_CREATED_NEW_KEY Then
        s = s & "Created new FirstLevel key" & sCrLf
    Else
        s = s & "Found existing FirstLevel key" & sCrLf
    End If
    Dim fExisted As Boolean
    e = CreateRegNode(hSubKey, "SecondLevel1", hSubSubKey, fExisted)
    If e Then GoTo RegFail
    If fExisted Then
        s = s & "Found existing SecondLevel1 key" & sCrLf
    Else
        s = s & "Created new SecondLevel1 key" & sCrLf
    End If
    e = RegCloseKey(hSubSubKey)
    If e Then GoTo RegFail
    
    e = CreateRegNode(hSubKey, "SecondLevel2", hSubSubKey)
    If e Then GoTo RegFail
    If fExisted Then
        s = s & "Found existing SecondLevel2 key" & sCrLf
    Else
        s = s & "Created new SecondLevel2 key" & sCrLf
    End If
    e = RegCloseKey(hSubSubKey)
    
    e = CreateRegValue("Default", hSubKey)
    If e Then GoTo RegFail
    s = s & "Created value: default" & sCrLf
    Dim ab() As Byte
    ab = "The bytes"
    e = CreateRegValue(ab, hSubKey, "Bytes")
    If e Then GoTo RegFail
    s = s & "Created value: Bytes" & sCrLf
    e = CreateRegValue("A String", hSubKey, "String")
    If e Then GoTo RegFail
    s = s & "Created value: String" & sCrLf
    e = CreateRegValue(5&, hSubKey, "Number")
    If e Then GoTo RegFail
    s = s & "Created value: Number" & sCrLf
    e = CreateRegValue("A %TEMP% string", hSubKey, "ExpandString")
    If e Then GoTo RegFail
    s = s & "Created value: ExpandString" & sCrLf
    
    Dim v As Variant
    e = GetRegValue(hSubKey, sEmpty, v)
    If e Then GoTo RegFail
    s = s & "Get default: " & v & sCrLf
    e = GetRegValue(hSubKey, "Bytes", v)
    If e Then GoTo RegFail
    ab = v
    s = s & "Get Bytes: " & HexDump(ab, ehdSample8) & sCrLf
    e = GetRegValue(hSubKey, "String", v)
    If e Then GoTo RegFail
    s = s & "Get String: " & v & sCrLf
    e = GetRegValue(hSubKey, "Number", v)
    If e Then GoTo RegFail
    s = s & "Get Number: " & v & sCrLf
    e = GetRegValue(hSubKey, "ExpandString", v)
    If e Then GoTo RegFail
    s = s & "Get ExpandString: " & v & sCrLf
    
    Dim i As Long, sName As String
    Do
        e = GetRegValueNext(hSubKey, i, sName, v)
        If e = 0 Then
            s = s & "Get item " & i & ": " & VarToStr(v) & sCrLf
        End If
        i = i + 1
    Loop While e = 0
    
    e = RegCloseKey(hSubKey)
    e = DeleteOneRegNode(HKEY_CURRENT_USER, _
                         "Software\VB and VBA Program Settings\FirstLevel")
    s = s & "Delete one node succeeded: " & (e = 0) & sCrLf
                         
    e = DeleteRegNodes(HKEY_CURRENT_USER, _
                       "Software\VB and VBA Program Settings\FirstLevel")
    s = s & "Delete all nodes succeeded: " & (e = 0) & sCrLf
   
    Do
        e = GetRegValueNext(hSubKey, i, sName, v)
        If e = 0 Then
            s = s & "Get item " & i & ": " & VarToStr(v) & sCrLf
        End If
        i = i + 1
    Loop While e = 0
    
    BugMessage s
    txtOut = s
    Exit Sub
RegFail:
    s = s & sCrLf & "Last API call failed: " & ApiError(e)
    txtOut = s
    
End Sub

Private Sub Form_Activate()
    lstRoot.AddItem "Classes Root"
    lstRoot.ItemData(0) = HKEY_CLASSES_ROOT
    lstRoot.AddItem "Current User"
    lstRoot.ItemData(1) = HKEY_CURRENT_USER
    lstRoot.AddItem "Local Machine"
    lstRoot.ItemData(2) = HKEY_LOCAL_MACHINE
    lstRoot.AddItem "Users"
    lstRoot.ItemData(3) = HKEY_USERS
    lstRoot.AddItem "Current Config"
    lstRoot.ItemData(4) = HKEY_CURRENT_CONFIG
    If Not IsNT Then
        lstRoot.AddItem "Dynamic Data"
        lstRoot.ItemData(5) = HKEY_DYN_DATA
    End If
    lstRoot.ListIndex = 1
End Sub

Private Sub lstItem_Click()
    Set itemCur = nodeCur.Items(lstItem.Text)
    txtValue = VarToStr(itemCur.Value)
End Sub

Private Sub lstNode_Click()
With lstItem
    txtValue = sEmpty
    Set nodeCur = nodeRoot.Nodes(lstNode.Text)
    Dim item As CRegItem
    .Clear
    For Each item In nodeCur.Items
        .AddItem item.Name
    Next
    If .ListCount Then .ListIndex = 0
End With
End Sub

Private Sub lstRoot_Click()
With lstNode
    hRootCur = lstRoot.ItemData(lstRoot.ListIndex)
    nodeRoot.Create sEmpty, hRootCur
    Dim node As CRegNode
    .Clear
    For Each node In nodeRoot
        .AddItem node.Name
    Next
    If .ListCount Then .ListIndex = 0
End With
End Sub

Private Function IUseRegItems_UseItem(item As CRegItem, _
                                      ByVal iLevel As Long) As Boolean
With item
    sOut = sOut & Space$((iLevel - 1) * 4) & " > " & _
                  .Name & " : " & VarToStr(.Value) & sCrLf
End With
End Function

Private Function IUseRegItems_UseNode(node As CRegNode, _
                                      ByVal iLevel As Long) As Boolean
With node
    sOut = sOut & Space$((iLevel) * 4) & .Name & " : " & sCrLf
    .WalkItems Me, iLevel
    DoEvents
End With
End Function

Function VarToStr(ByVal v As Variant) As String
    Select Case VarType(v)
    Case vbArray Or vbByte
        Dim ab() As Byte
        ab = v
        VarToStr = HexDump(ab, ehdSample8)
    Case vbLong
        VarToStr = "&H" & FmtHex(v, 8) & " (" & CStr(v) & ")"
    Case Else
        VarToStr = CStr(v)
    End Select
End Function



