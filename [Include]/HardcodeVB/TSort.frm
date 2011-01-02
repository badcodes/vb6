VERSION 5.00
Object = "{2DD06898-E157-11D0-8C51-00C04FC29CEC}#1.1#0"; "LISTBOXPLUS.OCX"
Begin VB.Form FTestSort 
   Caption         =   "Test Sort"
   ClientHeight    =   6600
   ClientLeft      =   1215
   ClientTop       =   2115
   ClientWidth     =   5205
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "TSort.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6600
   ScaleWidth      =   5205
   Begin ListBoxPlus.XListBoxPlus list 
      Height          =   2565
      Left            =   3465
      TabIndex        =   17
      Top             =   1980
      Width           =   1470
      _ExtentX        =   2593
      _ExtentY        =   4233
      BackColor       =   16777215
      ListCount       =   0
      ListIndex       =   -1
      Completion      =   0   'False
   End
   Begin VB.ListBox lstSort 
      Height          =   1230
      ItemData        =   "TSort.frx":0CFA
      Left            =   3444
      List            =   "TSort.frx":0D10
      TabIndex        =   16
      Top             =   396
      Width           =   1488
   End
   Begin VB.CommandButton cmdReplaceSList 
      Caption         =   "Replace..."
      Height          =   375
      Left            =   3465
      TabIndex        =   15
      Top             =   6108
      Width           =   1416
   End
   Begin VB.CommandButton cmdFindSList 
      Caption         =   "Find..."
      Height          =   375
      Left            =   3465
      TabIndex        =   14
      Top             =   4692
      Width           =   1416
   End
   Begin VB.CommandButton cmdInsertSList 
      Caption         =   "Insert..."
      Height          =   375
      Left            =   3465
      TabIndex        =   13
      Top             =   5172
      Width           =   1416
   End
   Begin VB.CommandButton cmdRemoveSList 
      Caption         =   "Remove..."
      Height          =   375
      Left            =   3465
      TabIndex        =   12
      Top             =   5652
      Width           =   1416
   End
   Begin VB.TextBox txtArray 
      Height          =   2568
      Left            =   276
      MultiLine       =   -1  'True
      TabIndex        =   11
      Top             =   1968
      Width           =   1476
   End
   Begin VB.CommandButton cmdFindArray 
      Caption         =   "Find..."
      Height          =   375
      Left            =   276
      TabIndex        =   10
      Top             =   4680
      Width           =   1476
   End
   Begin VB.CommandButton cmdRemoveCollect 
      Caption         =   "Remove..."
      Height          =   375
      Left            =   1872
      TabIndex        =   9
      Top             =   5652
      Width           =   1464
   End
   Begin VB.CommandButton cmdInsertCollect 
      Caption         =   "Insert..."
      Height          =   375
      Left            =   1872
      TabIndex        =   8
      Top             =   5172
      Width           =   1464
   End
   Begin VB.CommandButton cmdFindCollect 
      Caption         =   "Find..."
      Height          =   375
      Left            =   1872
      TabIndex        =   7
      Top             =   4692
      Width           =   1464
   End
   Begin VB.TextBox txtCollect 
      Height          =   2595
      Left            =   1872
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   1968
      Width           =   1464
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   375
      Left            =   276
      TabIndex        =   1
      Top             =   6096
      Width           =   1476
   End
   Begin VB.CheckBox chkDirection 
      Caption         =   "High to Low"
      Height          =   288
      Left            =   3447
      TabIndex        =   0
      Top             =   36
      Width           =   1416
   End
   Begin VB.Label lbl 
      Caption         =   "Array"
      Height          =   216
      Index           =   2
      Left            =   276
      TabIndex        =   6
      Top             =   1728
      Width           =   1476
   End
   Begin VB.Label lbl 
      Caption         =   "Collection"
      Height          =   240
      Index           =   1
      Left            =   1872
      TabIndex        =   5
      Top             =   1728
      Width           =   1464
   End
   Begin VB.Label lbl 
      Caption         =   "List Box Plus"
      Height          =   372
      Index           =   0
      Left            =   3465
      TabIndex        =   4
      Top             =   1728
      Width           =   1416
   End
   Begin VB.Label lblOut 
      Height          =   375
      Left            =   276
      TabIndex        =   2
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "FTestSort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private aInts(1 To 10) As Variant
Private aStrs(1 To 10) As Variant
Private aConst(1 To 10) As String
Private nStrs As Collection
Private esmlMode As ESortModeList
Private helper As CSortHelper

Sub Form_Load()
    
    Set helper = New CSortHelper
    Set nStrs = New Collection
    
    aConst(1) = "One"
    aConst(2) = "two"
    aConst(3) = "Three"
    aConst(4) = "four"
    aConst(5) = "Five"
    aConst(6) = "six"
    aConst(7) = "Seven"
    aConst(8) = "Eight"
    aConst(9) = "Nine"
    aConst(10) = "ten"
    
    aInts(1) = 5
    aInts(2) = 4
    aInts(3) = 9
    aInts(4) = 1
    aInts(5) = 7
    aInts(6) = 6
    aInts(7) = 3
    aInts(8) = 2
    aInts(9) = 10
    aInts(10) = 8
    
    aStrs(1) = "Five"
    aStrs(2) = "four"
    aStrs(3) = "Nine"
    aStrs(4) = "One"
    aStrs(5) = "Seven"
    aStrs(6) = "six"
    aStrs(7) = "Three"
    aStrs(8) = "two"
    aStrs(9) = "ten"
    aStrs(10) = "Eight"
    
    nStrs.Add "Apple"
    nStrs.Add "bean"
    nStrs.Add "Pear"
    nStrs.Add "banana"
    nStrs.Add "peach"
    nStrs.Add "CarRot"
    nStrs.Add "appleberry"
    nStrs.Add "Tangerine"
    nStrs.Add "wine"
    nStrs.Add "Beer"
    
    ' Put some items in a list box
    List.Clear
    List.Add "BEAR"
    List.Add "Lion"
    List.Add "tiger"
    List.Add "dog"
    List.Add "ZebrA"
    List.Add "kangaroo"
    List.Add "ELK"
    List.Add "WartHog"
    List.Add "Elephant"
    List.Add "stoat"
    
    Show
    
    lstSort.ListIndex = 0
   
    SortAll
    
End Sub

Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFindArray_Click()
    Dim iPos As Long, vKey As Variant, f As Boolean
    vKey = InputBox("Array item to find")
    If esmlMode = esmSortVal Then
        vKey = LookupString(aConst, vKey)
        f = BSearchArray(aInts(), vKey, iPos, helper)
    Else
        f = BSearchArray(aStrs(), vKey, iPos, helper)
    End If
    If f Then
        lblOut.Caption = "Found at position: " & iPos
    Else
        lblOut.Caption = "Insert at position: " & iPos
    End If
 End Sub

Private Sub cmdFindCollect_Click()
    Dim iPos As Long
    If BSearchCollection(nStrs, InputBox("Collection item to find: "), iPos, helper) Then
        lblOut.Caption = "Found at position: " & iPos
    Else
        lblOut.Caption = "Insert at position: " & iPos
    End If
End Sub

Private Sub cmdFindSList_Click()
    Dim v As Variant
    v = InputBox("List item to find: ")
    On Error Resume Next
    List.Current = v
    If Err Then lblOut.Caption = Err.Description
End Sub

Private Sub cmdInsertCollect_Click()
    Dim v As Variant, iPos As Long
    v = InputBox("Collection item to insert: ")
    If BSearchCollection(nStrs, v, iPos, helper) Then
        lblOut.Caption = "Can't insert duplicate item: " & v
    Else
        lblOut.Caption = sEmpty
        On Error GoTo IndexError
        nStrs.Add v, , iPos
        ShowCollect
    End If
    
    Exit Sub
IndexError:
    ' Item needs to be inserted at end of collection
    nStrs.Add v
    ShowCollect
End Sub

Private Sub cmdInsertSList_Click()
    Dim s As String, iPos As Long
    s = InputBox("List item to insert: ")
    On Error Resume Next
    List.Add s
    If Err Then lblOut.Caption = Err.Description
End Sub

Private Sub cmdRemoveCollect_Click()
    Dim v As Variant, iPos As Long
    v = InputBox("Collection item to remove: ")
    If IsNumeric(v) Then
        iPos = Val(v)
        If iPos > nStrs.Count Or iPos < 0 Then
            lblOut.Caption = "Invalid index: " & iPos
            Exit Sub
        End If
    ElseIf BSearchCollection(nStrs, v, iPos, helper) Then
        lblOut.Caption = sEmpty
    Else
        lblOut.Caption = "Item not in collection: " & v
        Exit Sub
    End If
    nStrs.Remove iPos
    ShowCollect
End Sub

Private Sub cmdRemoveSList_Click()
    Dim v As Variant, iPos As Long
    v = InputBox("List item to remove: ")
    On Error Resume Next
    List.Remove v
    If Err Then lblOut.Caption = Err.Description
End Sub

Private Sub cmdReplaceSList_Click()
    Dim vGet As Variant, vPut As Variant, iPos As Long
    vGet = InputBox("List item to replace: ")
    vPut = InputBox("New List item: ")
    On Error Resume Next
    List(vGet) = vPut
    If Err Then lblOut.Caption = Err.Description
End Sub

Sub chkDirection_Click()
    helper.HiToLo = (chkDirection.Value = vbChecked)
    List.HiToLo = (chkDirection.Value = vbChecked)
    If esmlMode = esmlShuffle Or esmlMode = esmlUnsorted Then Exit Sub
    SortAll
End Sub

Private Sub lstSort_Click()
    esmlMode = lstSort.ListIndex
    helper.SortMode = esmlMode
    SortAll
End Sub

Sub ShowArray()
    Dim i As Integer, s As String
    Static fInitialized As Boolean
    If esmlMode = esmlUnsorted Then
        If fInitialized = False Then
            For i = 1 To 10
                s = s & aConst(aInts(i)) & sCrLf
            Next
            fInitialized = True
        Else
            Exit Sub
        End If
    ElseIf esmlMode = esmlSortVal Then
        For i = 1 To 10
            s = s & aConst(aInts(i)) & sCrLf
        Next
    Else
        For i = 1 To 10
            s = s & aStrs(i) & sCrLf
        Next
    End If
    txtArray.Text = s
End Sub

Sub ShowCollect()
    Dim i As Integer, s As String, v As Variant
    s = sEmpty
    For Each v In nStrs
        s = s & v & sCrLf
    Next
    txtCollect.Text = s
End Sub

Sub SortAll()
    Select Case esmlMode
    Case esmlUnsorted
        ' Exit Sub
    Case esmlSortVal
        SortArray aInts(), , , helper
        SortCollection nStrs, , , helper
    Case esmlShuffle
        ShuffleArray aStrs(), helper
        ShuffleCollection nStrs, helper
    Case Else
        SortArray aStrs(), , , helper
        SortCollection nStrs, , , helper
    End Select
    ShowArray
    ShowCollect
    List.SortMode = esmlMode
End Sub

Function LookupString(A() As String, vKey As Variant) As Integer
    Dim i As Integer
    For i = 1 To 10
        If A(i) = vKey Then
            LookupString = i
            Exit Function
        End If
    Next
    LookupString = -1    ' Not found
End Function

' Uncomment to do tests when right click on Exit button
#Const fTestListPlus = 1

#If fTestListPlus Then
Private Sub cmdExit_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' Right click
    If Button = 2 Then TestList
End Sub

' Test of some features
Private Sub TestList()
    Dim s As String, i As Long
    Stop
    Debug.Print List(20)
    Debug.Print List(3)
    Debug.Print List("Lion")
    On Error Resume Next
    i = List("Giraffe")
    Debug.Print IIf(Err, "No Giraffe", "Giraffe")
    
    List(3) = "Deer"
    List("Lion") = "Big Cat"
    ShowAll
    With List
        Debug.Print .Item(20)
        Debug.Print .Item(3)
        Debug.Print .Item("Lion")
        ShowAll
        Debug.Print .Item("Giraffe")
        .Item(3) = "Deer"
        .Item("Lion") = "Big Cat"
        
        .Current = "dog"
        .Current = 1
        .Current = 4
        Debug.Print .Current
        Debug.Print .Item(.Current)
        Debug.Print .IndexItem
        Debug.Print .Text
        .Add "Dog"
        .Add "Tigger"
        .Add "dog"
        .Remove "dog"
        .Remove "Marten"
        .Remove 5
        .Remove 20
    End With
    ShowAll
End Sub

Private Sub ShowAll()
    Dim v As Variant, s As String
    For Each v In List
        s = s & v & " "
    Next
    Debug.Print "List: " & s
End Sub
#End If


