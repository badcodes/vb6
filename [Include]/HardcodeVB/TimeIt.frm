VERSION 5.00
Begin VB.Form FTimeIt 
   BackColor       =   &H8000000A&
   Caption         =   "Time It"
   ClientHeight    =   5748
   ClientLeft      =   1260
   ClientTop       =   1668
   ClientWidth     =   9120
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   7.8
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "timeit.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5748
   ScaleWidth      =   9120
   Begin VB.TextBox txtIteration 
      Height          =   495
      Left            =   105
      TabIndex        =   5
      Top             =   1665
      Width           =   1575
   End
   Begin VB.ListBox lstProblems 
      Height          =   2544
      Left            =   108
      TabIndex        =   3
      Top             =   2595
      Width           =   3135
   End
   Begin VB.CommandButton cmdTimeIt 
      Caption         =   "&Time It"
      Default         =   -1  'True
      Height          =   495
      Left            =   150
      TabIndex        =   2
      Top             =   105
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   495
      Left            =   150
      TabIndex        =   0
      Top             =   705
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Problems:"
      Height          =   255
      Left            =   105
      TabIndex        =   7
      Top             =   2265
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Iterations:"
      Height          =   255
      Left            =   105
      TabIndex        =   6
      Top             =   1425
      Width           =   1215
   End
   Begin VB.Label lblDescription 
      Height          =   648
      Left            =   3408
      TabIndex        =   4
      Top             =   120
      UseMnemonic     =   0   'False
      Width           =   5484
   End
   Begin VB.Label lblOutput 
      Height          =   4812
      Left            =   3384
      TabIndex        =   1
      Top             =   840
      Width           =   5604
   End
End
Attribute VB_Name = "FTimeIt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private problems As New Collection
Private problem As CProblem

Const sLogicalAndVsNestedIf = "Logical And Versus Nested If"
Const sByValVsByRef = "By Value Versus By Reference"
Const sCompareTypeProcessing = "Processing Different Types"
Const sInlineVsFunction = "Inline Versus Procedure"
Const sDeclareVsTypeLib = "Declare Versus Type Library"
Const sFixedVsVariableString = "Fixed Versus Variable String"
Const sCompareLoWords = "LoWord Variations"
Const sCShiftVsBShift = "Shift Versus Multiply or Divide"
Const sIIfVsIfThen = "IIf Versus If/Then/Else"
Const sDollarVsNone = "String Versus Variant Functions"
Const sEmptyVsQuotes = "Empty Versus Blank String"
Const sWithWithout = "Object Access Using With"
Const sMethodVsProc = "Methods Versus Procedures"
Const sForEachVsForI = "For Each Versus For Index"
Const sSortCollectVsArray = "Shuffle, Sort, and Search"
Const sAddCollect = "Add To Collection"
Const sSortRecurseVsIterate = "Recursive Versus Iterative Sorting"
Const sSortNameVsSortPoly = "Compare Sorting Methods"
Const sCompareFindFiles = "Compare File Finding Methods"
Const sCompareExistFile = "Compare File Existence Tests"
Const sFriendVsPublic = "Friend Properties Versus Public Properties"
Const sXorVsTmpSwap = "Temp Swap Versus XOR Swap"
#If iVBVer > 5 Then
Const sArrayVsVariant = "Array Return Versus Variant Return"
#End If

Private Sub Form_Load()
   
    Dim f As Boolean
    ChDir CurDir$
    txtIteration.Locked = False
          
    Set problem = New CProblem
    problem.Title = sLogicalAndVsNestedIf
    problem.Description = _
        "Does Basic short-circuit logical expressions? " & _
        "Or does a logical And operation take the " & _
        "same time as an equivalent nested If/Then?"
    problem.Iterations = 300000
    problems.Add problem, problem.Title

    Set problem = New CProblem
    problem.Title = sByValVsByRef
    problem.Description = _
        "How does the timing of ByVal arguments compare " & _
        "to the timing of the default By Reference arguments?"
    problem.Iterations = 60000
    problems.Add problem, problem.Title

    Set problem = New CProblem
    problem.Description = _
        "Compare counting with different types: Integer, " & _
        "Long, Single, Double, Currency, and Variant"
    problem.Title = sCompareTypeProcessing
    problem.Iterations = 500
    problems.Add problem, problem.Title

    Set problem = New CProblem
    problem.Title = sInlineVsFunction
    problem.Description = _
        "Compare an operation inline with the same " & _
        "operation in a function."
    problem.Iterations = 30000
    problems.Add problem, problem.Title

    Set problem = New CProblem
    problem.Title = sDeclareVsTypeLib
    problem.Description = _
        "Compare calling an API function through a Declare " & _
        "statement to calling through a type library."
    problem.Iterations = 30000
    problems.Add problem, problem.Title

    Set problem = New CProblem
    problem.Title = sFixedVsVariableString
    problem.Description = _
        "Compare various operations on variable-length strings with the same " & _
        "operations on fixed-length strings."
    problem.Iterations = 30000
    problems.Add problem, problem.Title

    Set problem = New CProblem
    problem.Title = sCompareLoWords
    problem.Description = _
        "Compare different methods of getting the low " & _
        "word of a long without getting overflow errors"
    problem.Iterations = 5000
    problems.Add problem, problem.Title

    Set problem = New CProblem
    problem.Title = sIIfVsIfThen
    problem.Description = _
        "Is IIf faster than an equivalent If/Then/Else block?"
    problem.Iterations = 30000
    problems.Add problem, problem.Title

    Set problem = New CProblem
    problem.Title = sDollarVsNone
    problem.Description = _
        "Do Mid and Mid$ do exactly the same thing?"
    problem.Iterations = 300
    problems.Add problem, problem.Title

    Set problem = New CProblem
    problem.Title = sEmptyVsQuotes
    problem.Description = _
        "Compare various forms of empty strings including " & _
        "double quotes, sEmpty, Empty, and vbNullString."
    problem.Iterations = 10000
    problems.Add problem, problem.Title

    Set problem = New CProblem
    problem.Title = sWithWithout
    problem.Description = _
        "Compare object access using With to equivalent " & _
        "fully-specified object access when used both with " & _
        "object variables and reference object variables"
    problem.Iterations = 3000
    problems.Add problem, problem.Title

    Set problem = New CProblem
    problem.Title = sMethodVsProc
    problem.Description = _
        "Compare direct and indirect (using Set) object operations " & _
        "to comparable operations using procedures and variables."
    problem.Iterations = 5000
    problems.Add problem, problem.Title

    Set problem = New CProblem
    problem.Title = sForEachVsForI
    problem.Description = _
        "Compare iterating through various collections and arrays " & _
        "with For Each to iterating through with For using an index " & _
        "variable."
    problem.Iterations = 60
    problems.Add problem, problem.Title

    Set problem = New CProblem
    problem.Title = sSortCollectVsArray
    problem.Description = _
        "Compare filling, shuffling, sorting, and searching an " & _
        "array with the same operations on a collection."
    problem.Iterations = 300
    problems.Add problem, problem.Title

    Set problem = New CProblem
    problem.Title = sAddCollect
    problem.Description = _
        "Compare adding new items to a collection at the " & _
        "beginning, middle, and end"
    problem.Iterations = 8000
    problems.Add problem, problem.Title
    
    Set problem = New CProblem
    problem.Title = sSortRecurseVsIterate
    problem.Description = _
        "Compare sorting recursively with sorting iteratively."
    problem.Iterations = 200
    problems.Add problem, problem.Title
    
    Set problem = New CProblem
    problem.Title = sSortNameVsSortPoly
    problem.Description = _
        "Compare faking procedure variables with a name-space " & _
        "hack to faking them with a polymorphic class."
    problem.Iterations = 200
    problems.Add problem, problem.Title
    
    Set problem = New CProblem
    problem.Title = sCompareFindFiles
    problem.Description = _
        "Compare finding files with Dir$ and with the " & _
        "FindFirstFiles API function."
    problem.Iterations = 1
    problems.Add problem, problem.Title

    Set problem = New CProblem
    problem.Title = sCompareExistFile
    problem.Description = _
        "Compare different ways to test for the existence of a file."
    problem.Iterations = 500
    problems.Add problem, problem.Title
    
    Set problem = New CProblem
    problem.Description = _
        "Compare calling Friend properties on classes and forms with " & _
        "calling equivalent Public properties on classes and forms."
    problem.Title = sFriendVsPublic
    problem.Iterations = 100000
    problems.Add problem, problem.Title

    Set problem = New CProblem
    problem.Description = _
        "Compare swapping with a temporary variable to swapping with " & _
        "XOR operations."
    problem.Title = sXorVsTmpSwap
    problem.Iterations = 100000
    problems.Add problem, problem.Title
    
#If iVBVer > 5 Then
    Set problem = New CProblem
    problem.Description = _
        "Compare receiving and returning arrays directly to " & _
        "returning arrays through Variants."
    problem.Title = sArrayVsVariant
    problem.Iterations = 1000
    problems.Add problem, problem.Title
#End If

    For Each problem In problems
        lstProblems.AddItem problem.Title
    Next
    lstProblems.ListIndex = 0

End Sub

Private Sub cmdTimeIt_Click()
    HourGlass Me
    lblOutput.Caption = "Working..."
    DoEvents
    Dim s As String, cIter As Long
    cIter = problem.Iterations
    Select Case problem.Title
    Case sLogicalAndVsNestedIf
        s = LogicalAndVsNestedIf(cIter)
    Case sByValVsByRef
        s = ByValVsByRef(cIter)
    Case sCompareTypeProcessing
        s = CompareTypeProcessing(cIter)
    Case sInlineVsFunction
        s = InlineVsFunction(cIter)
    Case sDeclareVsTypeLib
        s = DeclareVsTypeLib(cIter)
    Case sFixedVsVariableString
        s = FixedVsVariableString(cIter)
    Case sCompareLoWords
        s = CompareLoWords(cIter)
    Case sIIfVsIfThen
        s = IIfVsIfThen(cIter)
    Case sDollarVsNone
        s = DollarVsNone(cIter)
    Case sEmptyVsQuotes
        s = EmptyVsQuotes(cIter)
    Case sWithWithout
        s = WithWithout(cIter)
    Case sMethodVsProc
        s = MethodVsProc(cIter)
    Case sForEachVsForI
        s = ForEachVsForI(cIter)
    Case sSortCollectVsArray
        s = SortCollectVsArray(cIter)
    Case sAddCollect
        s = AddCollect(cIter)
    Case sSortRecurseVsIterate
        s = SortRecurseVsIterate(cIter)
    Case sSortNameVsSortPoly
        s = SortNameVsSortPoly(cIter)
    Case sCompareFindFiles
        s = CompareFindFiles(cIter)
    Case sCompareExistFile
        s = CompareExistFile(cIter)
    Case sFriendVsPublic
        s = CompareFriendVsPublic(cIter)
    Case sXorVsTmpSwap
        s = XorVsTmpSwap(cIter)
#If iVBVer > 5 Then
    Case sArrayVsVariant
        s = ArrayVsVariant(cIter)
#End If
    End Select
    HourGlass Me
    BugMessage problem.Description
    BugMessage "Iterations: " & problem.Iterations
    BugMessage s
    lblOutput.Caption = s
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub lstProblems_DblClick()
    cmdTimeIt_Click
End Sub


Private Sub txtIteration_Change()
    Dim i As Long
    i = Val(txtIteration.Text)
    If i Then
        problem.Iterations = i
    Else
        Beep
        txtIteration.Text = problem.Iterations
    End If
End Sub

Private Sub lstProblems_Click()
    Dim s As String
    lblOutput = sEmpty
    s = lstProblems.List(lstProblems.ListIndex)
    Set problem = problems(s)
    lblDescription.Caption = problem.Description
    txtIteration.Text = problem.Iterations
    'txtIteration.SetFocus
    txtIteration.SelStart = 0
    txtIteration.SelLength = Len(txtIteration.Text)
End Sub

Public Property Get ProcProp() As Long
    ProcProp = 1
End Property

Friend Property Get FriendProp() As Long
    FriendProp = 1
End Property



