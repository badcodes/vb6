VERSION 5.00
Begin VB.Form FTestCollect 
   Caption         =   "Test Collections"
   ClientHeight    =   6336
   ClientLeft      =   2040
   ClientTop       =   1536
   ClientWidth     =   6804
   Icon            =   "TCollect.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6336
   ScaleWidth      =   6804
   WhatsThisHelp   =   -1  'True
   Begin VB.CheckBox chkOld 
      Caption         =   "Old version"
      Height          =   252
      Left            =   120
      TabIndex        =   15
      Top             =   5040
      Width           =   1332
   End
   Begin VB.CommandButton cmdDrives 
      Caption         =   "Drives"
      Height          =   492
      Left            =   120
      MaskColor       =   &H00000000&
      TabIndex        =   14
      Top             =   4440
      Width           =   1332
   End
   Begin VB.Frame fm 
      Caption         =   "Stack Type"
      Height          =   1512
      Left            =   120
      TabIndex        =   8
      Top             =   2508
      Width           =   1356
      Begin VB.TextBox txtCount 
         Height          =   288
         Left            =   720
         TabIndex        =   12
         Text            =   "2000"
         Top             =   1080
         Width           =   492
      End
      Begin VB.OptionButton optStack 
         Caption         =   "List"
         Height          =   255
         Index           =   0
         Left            =   84
         MaskColor       =   &H00000000&
         TabIndex        =   11
         Top             =   240
         Value           =   -1  'True
         Width           =   960
      End
      Begin VB.OptionButton optStack 
         Caption         =   "Vector"
         Height          =   255
         Index           =   1
         Left            =   84
         MaskColor       =   &H00000000&
         TabIndex        =   10
         Top             =   480
         Width           =   972
      End
      Begin VB.OptionButton optStack 
         Caption         =   "Collection"
         Height          =   255
         Index           =   2
         Left            =   84
         MaskColor       =   &H00000000&
         TabIndex        =   9
         Top             =   720
         Width           =   1032
      End
      Begin VB.Label lbl 
         Caption         =   "Count:"
         Height          =   252
         Left            =   120
         TabIndex        =   13
         Top             =   1080
         Width           =   612
      End
   End
   Begin VB.CommandButton cmdVector 
      Caption         =   "Vector"
      Height          =   504
      Left            =   120
      MaskColor       =   &H00000000&
      TabIndex        =   7
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton cmdList 
      Caption         =   "Linked List"
      Height          =   504
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   1332
   End
   Begin VB.ListBox lstStuff 
      Height          =   432
      Left            =   5520
      TabIndex        =   5
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   120
      MaskColor       =   &H00000000&
      TabIndex        =   4
      Top             =   5760
      Width           =   1335
   End
   Begin VB.CommandButton cmdInternal 
      Caption         =   "Internal"
      Height          =   384
      Left            =   5520
      MaskColor       =   &H00000000&
      TabIndex        =   3
      Top             =   5880
      Width           =   1212
   End
   Begin VB.CommandButton cmdStack 
      Caption         =   "Stack"
      Height          =   495
      Left            =   96
      MaskColor       =   &H00000000&
      TabIndex        =   2
      Top             =   1908
      Width           =   1335
   End
   Begin VB.CommandButton cmdCollect 
      Caption         =   "Collection"
      Height          =   495
      Left            =   120
      MaskColor       =   &H00000000&
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox txtOut 
      Height          =   4695
      Left            =   1644
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   5052
   End
End
Attribute VB_Name = "FTestCollect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdCollect_Click()
    Dim s As String
    ' Declare collection
    Dim animals As New Collection
        
    s = "Add items to collection: Lion, Tiger, Bear, Shrew before 1, Weasel after 1" & sCrLf
    ' Create collection
    animals.Add "Lion"
    animals.Add "Tiger"
    animals.Add "Bear"
    animals.Add "Shrew", , 1
    animals.Add "Weasel", , , 1
        
    ' Access collection items
    Debug.Print animals(3) & " " & animals.item(3)
    
    ' Iterate through collection
    s = s & sCrLf & "Iterate with For Each: " & sCrLf
    Dim vAnimal As Variant
    For Each vAnimal In animals
        s = s & Space$(4) & vAnimal & sCrLf
        Debug.Print vAnimal
    Next
    
    ' Replace collection item
    s = s & "Access item 2: " & animals.item(2) & sCrLf
    s = s & "Replace " & animals.item(2) & " with Wolverine" & sCrLf
    'animals(2) = "Wolverine"
    'Set animals(2) = "Wolverine"
    animals.Add "Wolverine", , 2
    animals.Remove 3
        
    s = s & sCrLf & "Iterate with For I: " & sCrLf
    Dim i As Integer
    For i = 1 To animals.Count
        s = s & Space$(4) & animals(i) & sCrLf
        Debug.Print animals(i)
    Next
    
    Dim vAnimal2 As Variant
    s = s & sCrLf & "Nested iteration loops with For Each: " & sCrLf
    For Each vAnimal In animals
        s = s & Space$(4) & vAnimal & sCrLf
        If vAnimal = "Lion" Then
            For Each vAnimal2 In animals
                s = s & Space$(8) & vAnimal2 & sCrLf
            Next
        End If
    Next
    
    BugMessage s
    txtOut.Text = s
   
End Sub

Private Sub cmdDrives_Click()
    txtOut.Text = sEmpty
    txtOut.Refresh
    Dim s As String
   
    Dim driveCur As New CDrive
    driveCur = 0       ' Initialize to current drive
    
    Debug.Print driveCur
    s = "Drive information for current drive:" & sCrLf
    Const sBFormat = "#,###,###,##0"
    With driveCur
        s = s & "Drive " & .Root & " [" & .Label & ":" & _
                .Serial & "] (" & .KindStr & ") has " & _
                Format$(.FreeBytes, sBFormat) & " free from " & _
                Format$(.TotalBytes, sBFormat) & sCrLf
    End With
    
    driveCur = "C:\"       ' Initialize to current drive
    
    s = "Drive information for drive C:" & sCrLf
    Debug.Print driveCur
    With driveCur
        s = s & "Drive " & .Root & " [" & .Label & ":" & _
                .Serial & "] (" & .KindStr & ") has " & _
                Format$(.FreeBytes, sBFormat) & " free from " & _
                Format$(.TotalBytes, sBFormat) & sCrLf
    End With
    
    s = s & sCrLf
    s = s & "Drive information for available drives:" & sCrLf
    Dim drives As Object, drive As CDrive
    If chkOld Then
        Set drives = New CDrivesO
    Else
        Set drives = New CDrives
    End If
    For Each drive In drives
        With drive
            s = s & "Drive " & .Root & " [" & .Label & ":" & _
                    .Serial & "] (" & .KindStr & ") has " & _
                    Format$(.FreeBytes, sBFormat) & " free from " & _
                    Format$(.TotalBytes, sBFormat) & sCrLf
        End With
    Next
    Debug.Print drives("C:\").Label
    
    BugMessage s
    txtOut.Text = s
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdInternal_Click()
    Dim s As String, i As Integer
    s = "Forms collection:" & sCrLf
    Dim obj As Object
    For Each obj In Forms
        s = s & obj.Name & sCrLf
        Debug.Print obj.Name
    Next
    
    s = s & sCrLf & "Controls collection:" & sCrLf
    Dim ctl As Control
    For Each ctl In Controls
        s = s & ctl.Name & sCrLf
        Debug.Print ctl.Name
    Next
    
    s = s & sCrLf & "Printers collection:" & sCrLf
    Dim prt As Printer
    For Each prt In Printers
        s = s & prt.DriverName & sCrLf
        Debug.Print prt.CurrentX ' .DriverName
        ' Do something with each prt
    Next
    
    s = s & sCrLf & "Control array:" & sCrLf
    For Each ctl In optStack
        s = s & ctl.Caption & Space$(4) & ctl.Value & sCrLf
        Debug.Print ctl.Caption & Space$(4) & ctl.Value
    Next
    
    For i = optStack.LBound To optStack.UBound
        s = s & optStack(i).Caption & Space$(4) & optStack(i).Value & sCrLf
        Debug.Print optStack(i).Caption & Space$(4) & optStack(i).Value
    Next

    lstStuff.AddItem "Lions"
    lstStuff.AddItem "Tigers"
    lstStuff.AddItem "Bears"
    
    Debug.Print lstStuff.list(1)

    ' You can't do that!
    'For Each obj In lstStuff
    '    Debug.Print obj.Name
    'Next
    
    s = s & sCrLf & "ListBox:" & sCrLf
    For i = 0 To lstStuff.ListCount - 1
        s = s & lstStuff.list(i) & sCrLf
        Debug.Print lstStuff.list(i)
    Next
    
    BugMessage s
    txtOut.Text = s

End Sub

Private Sub cmdList_Click()
    Dim s As String
    
    s = "Add items:" & sCrLf & _
        Space$(4) & "Add Bear, Tiger, Lion, Elephant, Horse, Dog" & sCrLf
    ' Insert item into list
    Dim list As New CList
    list.Add "Bear"
    list.Add "Tiger"
    list.Add "Lion"
    list.Add "Elephant"
    list.Add "Horse"
    list.Add "Dog"
    s = s & "Count: " & list.Count & sCrLf
    s = s & "Head: " & list & sCrLf
    s = s & "Item 2: " & list(2) & sCrLf
    s = s & "Item Tiger: " & list("Tiger") & sCrLf
    
    s = s & "Iterate:" & sCrLf
    Dim walker As New CListWalker
    walker.Attach list
    Do While walker.More
        s = s & Space$(4) & walker & sCrLf
    Loop
    
    s = s & "Replace Elephant with Pig" & sCrLf
    list("Elephant") = "Pig"
    
    s = s & "Remove head: " & list & sCrLf
    list.Remove
    s = s & "Remove Bear" & sCrLf
    list.Remove "Bear"
    s = s & "Remove 3: " & list(3) & sCrLf
    list.Remove 3
    
    Dim walker2 As New CListWalker
    s = s & "Nesting iterate:" & sCrLf
    walker.Attach list
    Do While walker.More
        s = s & Space$(4) & walker & sCrLf
        If walker = "Pig" Then
            walker2.Attach list
            s = s & Space$(4) & "Nested iterate:" & sCrLf
            Do While walker2.More
                s = s & Space$(8) & walker2 & sCrLf
            Loop
        End If
    Loop

    s = s & "Iterate with For Each:" & sCrLf
    Dim v As Variant
    For Each v In list
        s = s & Space$(4) & "V: " & v & sCrLf
    Next
        
    s = s & "Clear and then iterate:" & sCrLf
    list.Clear
    For Each v In list
        s = s & Space$(4) & "V: " & v & sCrLf
    Next
    
    BugMessage s
    txtOut.Text = s
End Sub

Private Sub cmdStack_Click()
    Dim s As String
    s = "Push animals onto stack: " & sCrLf
    txtOut.Text = s
    txtOut.Refresh
    Dim beasts As IStack
    Select Case GetOption(optStack)
    Case 0
        Set beasts = New CStackLst
    Case 1
        Set beasts = New CStackVec
    Case 2
        Set beasts = New CStackCol
    End Select
    s = s & Space$(4) & "Push Lion" & sCrLf
    beasts.Push "Lion"
    s = s & Space$(4) & "Push Tiger" & sCrLf
    beasts.Push "Tiger"
    s = s & Space$(4) & "Push Bear" & sCrLf
    beasts.Push "Bear"
    s = s & Space$(4) & "Push Shrew" & sCrLf
    beasts.Push "Shrew"
    s = s & Space$(4) & "Push Weasel" & sCrLf
    beasts.Push "Weasel"
    s = s & Space$(4) & "Push Yetti" & sCrLf
    beasts.Push "Yetti"
    
    s = s & "Pop animals off stack: " & sCrLf
    Do While beasts.Count
        s = s & Space$(4) & "Pop " & beasts.Pop & sCrLf
    Loop
           
    Dim numbers As IStack
    Select Case GetOption(optStack)
    Case 0
        Set numbers = New CStackLst
    Case 1
        Set numbers = New CStackVec
    Case 2
        Set numbers = New CStackCol
    End Select
    Dim i As Integer, sec As Currency, secDone As Currency
    ProfileStart sec
    For i = 1 To txtCount
        numbers.Push i
    Next
    Do
        i = numbers.Pop
    Loop While numbers.Count
    ProfileStop sec, secDone
    s = s & sCrLf & "Push/Pop Timing: " & secDone & sCrLf
    
    Dim langs As New CStack
    s = s & sCrLf & "Push languages onto real stack..." & sCrLf
    
    s = s & Space$(4) & "Push Basic" & sCrLf
    langs.Push "Basic"
    s = s & Space$(4) & "Push Pascal" & sCrLf
    langs.Push "Pascal"
    s = s & Space$(4) & "Push C++" & sCrLf
    langs.Push "C++"
    s = s & Space$(4) & "Push Java" & sCrLf
    langs.Push "Java"
    s = s & Space$(4) & "Push REXX" & sCrLf
    langs.Push "REXX"
    s = s & Space$(4) & "Push Forth" & sCrLf
    langs.Push "Forth"
        
    s = s & "Pop languages off stack: " & sCrLf
    Do While langs.Count
        s = s & Space$(4) & "Pop " & langs.Pop & sCrLf
    Loop
    
    BugMessage s
    txtOut.Text = s
    
End Sub

Private Function CertifyCollection(obj As Object) As Boolean
    Dim v As Variant
    With obj
        On Error Resume Next
        .Add .Count         ' Test Add and Count by adding
        v = .item(.Count)   ' Test Item by accessing
        For Each v In obj   ' Test iteration
        Next
        .Remove .Count      ' Test Remove by removing
        CertifyCollection = (Err = 0)
    End With
End Function
        
Private Sub cmdVector_Click()

    Dim vector As New CVector, i As Long, s As String
    s = "Insert numbers in vector: " & sCrLf
    For i = 1 To 15
        vector(i) = i * i
        s = s & Space$(4) & i * i & ": vector(" & i & ")" & sCrLf
    Next
    s = s & "Read numbers from vector: " & sCrLf
    For i = 1 To vector.Last
        s = s & Space$(4) & "vector(" & i & ") = " & vector(i) & sCrLf
    Next
    s = s & "Shrink vector to 5 and read numbers: " & sCrLf
    vector.Last = 5
    For i = 1 To vector.Last
        s = s & Space$(4) & "vector(" & i & ") = " & vector(i) & sCrLf
    Next
    
    s = s & "Read numbers with For Each: " & sCrLf
    Dim v As Variant
    For Each v In vector
        s = s & Space$(4) & "v = " & v & sCrLf
    Next
    
    BugMessage s
    txtOut = s
End Sub

Private Sub optStack_Click(Index As Integer)
    txtOut = sEmpty
End Sub



