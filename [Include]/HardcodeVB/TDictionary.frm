VERSION 5.00
Begin VB.Form FTestDictionary 
   Caption         =   "Test Dictionary"
   ClientHeight    =   5016
   ClientLeft      =   1236
   ClientTop       =   3336
   ClientWidth     =   7092
   Icon            =   "TDictionary.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5016
   ScaleWidth      =   7092
   Begin VB.CommandButton cmdDictionary 
      Caption         =   "Dictionary"
      Height          =   495
      Left            =   156
      TabIndex        =   6
      Top             =   228
      Width           =   1215
   End
   Begin VB.CommandButton cmdFiles 
      Caption         =   "Files"
      Height          =   495
      Left            =   144
      TabIndex        =   5
      Top             =   1416
      Width           =   1215
   End
   Begin VB.CheckBox chkDescend 
      Caption         =   "Descending"
      Height          =   372
      Left            =   156
      TabIndex        =   4
      Top             =   3828
      Width           =   1260
   End
   Begin VB.CheckBox chkText 
      Caption         =   "Text Compare"
      Height          =   372
      Left            =   156
      TabIndex        =   3
      Top             =   3468
      Width           =   1428
   End
   Begin VB.TextBox txtOut 
      Height          =   4572
      Left            =   1620
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Top             =   240
      Width           =   5304
   End
   Begin VB.CommandButton cmdFilters 
      Caption         =   "Filters"
      Height          =   495
      Left            =   156
      TabIndex        =   1
      Top             =   2028
      Width           =   1215
   End
   Begin VB.CommandButton cmdWordCount 
      Caption         =   "Word Count"
      Height          =   492
      Left            =   156
      TabIndex        =   0
      Top             =   828
      Width           =   1212
   End
End
Attribute VB_Name = "FTestDictionary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private cmp As VbCompareMethod
Private fDescend As Boolean

Private Type TPerson
    Name As String
    Age As Integer
End Type


Private Sub Form_Load()
    ChDir App.Path
    cmp = IIf(chkText.Value = vbChecked, vbTextCompare, vbBinaryCompare)
    fDescend = -chkDescend.Value
    'TestArrays
    'TestUDTs
    
End Sub

Private Sub chkDescend_Click()
    fDescend = -chkDescend.Value
End Sub

Private Sub chkText_Click()
    cmp = IIf(chkText.Value = vbChecked, vbTextCompare, vbBinaryCompare)
End Sub

Private Sub cmdDictionary_Click()

    txtOut = sEmpty
    
    Dim sOut As String
    sOut = sOut & "Add animals and characteristics" & sCrLf
    Dim animals As New Dictionary
    With animals
        .Add "Lion", "fierce"
        .Add "Tiger", "cunning"
        .Add "Bear", "savage"
        .Add "Shrew", "bloodthirsty"
        .Add "Weasel", "ruthless"
    End With
    
    sOut = sOut & "Print keys: "
    
    Dim animal As Variant
    For Each animal In animals
        Debug.Print animal
        sOut = sOut & animal & " "
    Next
    
    sOut = sOut & sCrLf & "Print items: "
    
    Dim character As Variant
    For Each character In animals.Items
        Debug.Print character
        sOut = sOut & character & " "
    Next
    
    sOut = sOut & sCrLf & "Try to print keys with index: "
    
    Dim i As Long
    For i = 1 To animals.Count
        Debug.Print animals(i)
        sOut = sOut & i & ":" & animals(i) & " "
    Next
    
    sOut = sOut & sCrLf & "Remove junk created by bogus indexes "
    For i = 1 To animals.Count
        If animals.Exists(i) Then animals.Remove i
    Next
    
    sOut = sOut & sCrLf & "Print hidden hash values: "
    For Each animal In animals.Keys
        sOut = sOut & animals.HashVal(animal) & " "
    Next
    
    sOut = sOut & sCrLf & "Add Cow and Mouse by assignment: " & sCrLf
    animals("Cow") = "docile"
    animals("Mouse") = "timid"
    
    sOut = sOut & "Show new items: " & sCrLf
    For Each animal In animals.Keys
        sOut = sOut & sTab & animal & ":" & animals(animal) & sCrLf
    Next
    
    sOut = sOut & "Assign items and keys to arrays: " & sCrLf
    Dim aKeys As Variant, aItems As Variant
    aKeys = animals.Keys
    aItems = animals.Items
    sOut = sOut & "Index key array: "
    
    For i = 0 To animals.Count - 1
        Debug.Print aKeys(i)
        sOut = sOut & aKeys(i) & " "
    Next
    
    sOut = sOut & sCrLf & "Index item array: "
    For i = 0 To animals.Count - 1
        Debug.Print aKeys(i)
        sOut = sOut & aItems(i) & " "
    Next
    
    sOut = sOut & sCrLf & "Remove everything and prove it:"
    animals.RemoveAll
    For Each animal In animals
        Debug.Print animal
        sOut = sOut & animal & " "
    Next
        
    txtOut = sOut
    
End Sub

Private Sub cmdWordCount_Click()
    
    txtOut = sEmpty
    Dim dicWordCount As New Dictionary, sFile As String
    Dim vasWords As Variant, i As Long, sKey As String, sSep As String
    sFile = "TDictionary.frm"
    ' Fails if you move the compiled file, but what the heck
    If IsExe Then sFile = "..\" & sFile
    ' Separate with all the symbols used in VB programming
    sSep = " ,.!?()#<>+-*/=&\:;" & sTab & sCr & sLf & sQuote2 & sQuote1
    ' Read the words from a file into an array
    On Error Resume Next
    vasWords = GetFileTokens(sFile, sSep)
    If IsEmpty(vasWords) Then
        MsgBox "Can't load file"
        Exit Sub
    End If
    On Error GoTo 0
    ' Put all the words into a table, counting each
    For i = 0 To UBound(vasWords)
        sKey = vasWords(i)
        ' Increment count when a word is encountered
        dicWordCount(sKey) = dicWordCount(sKey) + 1
    Next
    
    ' Sort the keys
    vasWords = SortKeys(dicWordCount.Keys, fDescend, cmp)
    ' Filter all numbers out of the list using a Like pattern
    vasWords = FilterLike(vasWords, "#*", False)
    
    ' Could apply this alternate filter for a different effect
    'vas = Filter(vas, "as", True, vbTextCompare)
    ' More efficent (and more confusing) to chain array tools
    'vasWords = FilterLike(SortKeys(dicWordCount.Keys, fDescend, cmp), _
                          "#*", False)
    
    ' Output the table contents
    Dim sOut As String
    For i = 0 To UBound(vasWords)
        sKey = vasWords(i)
        sOut = sOut & sKey & " : " & dicWordCount(sKey) & sCrLf
    Next
    txtOut = sOut
    
End Sub

Private Sub cmdFiles_Click()
#If iVBVer > 5 Then
    Dim dicFiles As Dictionary, s As String
    Set dicFiles = FileDictionary("*.*")
    
    Dim vaSort As Variant, i As Long, sName As String, vfi As Variant
    vaSort = SortKeys(dicFiles.Keys, fDescend, cmp)
    For i = 0 To UBound(vaSort)
        sName = vaSort(i)
        vfi = dicFiles(sName)
        With vfi
            s = s & sName & " : " & Format$(.Length, "#,###")
            s = s & " : " & .LastWrite & " : " & AttrString(.Attribs) & sCrLf
        End With
    Next
    txtOut = s
#Else
    txtOut = "This techniques depends on Public UDTs," & _
             "which aren't allowed in VB5"
#End If
End Sub

Private Sub cmdFilters_Click()
    
    Dim vaFiles As Variant, s As String, i As Long
    
    s = "List of text files in Windows directory from FileArray: " & sCrLf
    vaFiles = FileArray(WindowsDir & "\*.txt")
    s = s & LineWrap(Join(vaFiles, ", "), 65)
    
    Dim vaTokens As Variant, vaLines As Variant
    s = s & sCrLf & sCrLf & "File tokens from first file: " & sCrLf
    vaTokens = GetFileTokens(WindowsDir & "\" & vaFiles(0))
    s = s & Left$(Join(vaTokens), 65) & "..." & sCrLf
    
    s = s & sCrLf & "First three lines from first file: " & sCrLf
    vaLines = GetFileLines(WindowsDir & "\" & vaFiles(0))
    For i = 0 To UBound(vaLines)
        s = s & vaLines(i) & sCrLf
        If i = 2 Then Exit For
    Next
    
    Dim vaFruit As Variant, v As Variant, sFruit As String
    sFruit = "berry, Apple, pomegranate, banana, Orange, pear, Date, fig"
    s = s & sCrLf & "Fail to separate a string with Split: " & sCrLf
    vaFruit = Split(sFruit, " ,")
    s = s & Join(vaFruit, "#") & sCrLf

    s = s & sCrLf & "Separate a string with Splits: " & sCrLf
    vaFruit = Splits(sFruit, " ,")
    s = s & Join(vaFruit, "#") & sCrLf
    
    Dim vaSort As Variant, vaShuffle As Variant, vaReverse As Variant
    s = s & sCrLf & "Sort fruit: " & sCrLf
    vaSort = SortKeys(vaFruit, fDescend, cmp)
    s = s & Join(vaSort, "#") & sCrLf
    
    s = s & sCrLf & "Shuffle fruit: " & sCrLf
    vaShuffle = ShuffleKeys(vaFruit)
    s = s & Join(vaShuffle, "#") & sCrLf
    
    s = s & sCrLf & "Reverse shuffled fruit: " & sCrLf
    vaReverse = ReverseArray(vaShuffle)
    s = s & Join(vaReverse, "#") & sCrLf
    
    
    Dim vaFilter As Variant
'#If iVBVer > 5 Then
    s = s & sCrLf & "Filter in sorted fruit containing 'te': " & sCrLf
    vaFilter = Filter(vaSort, "te", True, cmp)
    s = s & Join(vaFilter, "#") & sCrLf
'#End If
    
    s = s & sCrLf & "Filter out sorted fruit starting with vowels: " & sCrLf
    vaFilter = FilterLike(vaSort, "[AEIOUaeiou]*", False)
    s = s & Join(vaFilter, "#") & sCrLf
    txtOut = s
    
End Sub

#If iVBVer > 5 Then
Sub TestArrays()

    ' Some code to test assignment combinations
    Dim avsT() As Variant, asT() As String
    Dim vasT As Variant, vavT As Variant
    ReDim avsT(0 To 2) As Variant   ' Array of variants containing strings
    ReDim asT(0 To 2) As String     ' Array of strings
    ReDim vasT(0 To 2) As String    ' Variant containing array of strings
    ReDim vavT(0 To 2) As Variant   ' Variant containing array of variants
    
    avsT(0) = "Zero": asT(0) = "Zero": vasT(0) = "Zero": vavT(0) = "Zero"
    avsT(1) = "One": asT(1) = "One": vasT(1) = "One": vavT(1) = "One"
    avsT(2) = "Two": asT(2) = "Two": vasT(2) = "Two": vavT(2) = "Two"
    
    'avsT = asT  ' Can't assign string array to variant array
    'avsT = vasT ' Can't variant with string array to variant array
    avsT = vavT ' Assign variant with variant array to variant array
    
    'asT = avsT  ' Can't assign variant array to string array
    asT = vasT  ' Assign variant with string array to string array
    'asT = vavT  ' Can't assign variant with variant array to string array
    
    vavT = avsT ' Assign array of variants containing strings to variant
    vavT = asT  ' Assign string array to variant
    vavT = vasT ' Assign variant with string array to variant
    
    vasT = avsT ' Assign array of variants containing strings to variant
    vasT = asT  ' Assign string array to variant
    vasT = vavT ' Assign variant with variant array to variant

End Sub

Function AttrString(attr As VbFileAttribute) As String
    Dim s As String
    If attr And vbDirectory Then s = s & "d"
    If attr And vbArchive Then s = s & "a"
    If attr And vbReadOnly Then s = s & "r"
    If attr And vbHidden Then s = s & "h"
    If attr And vbSystem Then s = s & "s"
    If attr And vbVolume Then s = s & "v"
    AttrString = s
End Function

Sub TestUDTs()

    Dim fd As New CFileData
    
    ' Wrong way to use UDT class properties
    With fd
        ' Assign FileInfo members
        .FileInfo.Length = 22
        .FileInfo.Attribs = vbReadOnly
        .FileInfo.LastWrite = Now
        ' Use FileInfo members
        Debug.Print .FileInfo.LastWrite
    End With
    
    ' Right way to use UDT class properties
    Dim fi As TFileInfo
    With fi
        ' Assign to a temporary UDT variable
        .Length = 22
        .Attribs = vbReadOnly
        .LastWrite = Now
    End With
    ' Assign the variable to the property
    fd.FileInfo = fi
    ' Now you can access the property
    Debug.Print fd.FileInfo.LastWrite
    
End Sub
#Else
' Define the mighty Join function for VB5
Function Join(vaSourceArray As Variant, _
              Optional vsDelimiter As Variant) As String
    Dim i As Long
    If IsMissing(vsDelimiter) Then vsDelimiter = " "
    For i = LBound(vaSourceArray) To UBound(vaSourceArray) - 1
        Join = Join & vaSourceArray(i) & vsDelimiter
    Next
    Join = Join & vaSourceArray(i)
End Function
#End If
