VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GSort"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum EErrorSort
    eeBaseSort = 13620  ' Sort
End Enum

' Iterative QuickSort algorithm
Sub SortArray(aTarget() As Variant, Optional vFirst As Variant, _
              Optional vLast As Variant, Optional helper As ISortHelper)
    Dim iFirst As Long, iLast As Long
    If IsMissing(vFirst) Then iFirst = LBound(aTarget) Else iFirst = vFirst
    If IsMissing(vLast) Then iLast = UBound(aTarget) Else iLast = vLast
    If helper Is Nothing Then Set helper = New CSortHelper
    
With helper
    Dim iLo As Long, iHi As Long, iRand As Long, stack As New CStack
    Do
        Do
            ' Swap from ends until first and last meet in the middle
            If iFirst < iLast Then
                ' If we're in the middle and out of order, swap
                If iLast - iFirst = 1 Then
                    If .Compare(aTarget(iFirst), aTarget(iLast)) > 0 Then
                        .Swap aTarget(iFirst), aTarget(iLast)
                    End If
                Else
                    ' Split at some random point
                    .Swap aTarget(iLast), _
                          aTarget(MRandom.Random(iFirst, iLast))
                    ' Swap high values below the split for low values above
                    iLo = iFirst: iHi = iLast
                    Do
                        ' Find any low value larger than split
                        Do While (iLo < iHi) And _
                                 (.Compare(aTarget(iLo), aTarget(iLast)) <= 0)
                            iLo = iLo + 1
                        Loop
                        ' Find any high value smaller than split
                        Do While (iHi > iLo) And _
                                 (.Compare(aTarget(iHi), aTarget(iLast)) >= 0)
                            iHi = iHi - 1
                        Loop
                        ' Swap too high low value for too low high value
                        If iLo < iHi Then .Swap aTarget(iLo), aTarget(iHi)
                    Loop While iLo < iHi
                    ' Current (iLo) is larger than split (iLast), so swap
                    .Swap aTarget(iLo), aTarget(iLast)
                    ' Push range markers of larger part for later sorting
                    If (iLo - iFirst) < (iLast - iLo) Then
                        stack.Push iLo + 1
                        stack.Push iLast
                        iLast = iLo - 1
                    Else
                        stack.Push iFirst
                        stack.Push iLo - 1
                        iFirst = iLo + 1
                    End If
                    ' Exit from inner loop to process smaller part
                    Exit Do
                End If
            End If
            
            ' If stack empty, Exit outer loop
            If stack.Count = 0 Then Exit Sub
            ' Else pop first and last from last deferred section
            iLast = stack.Pop
            iFirst = stack.Pop
        Loop
    Loop
End With
End Sub

' Argument is array in Variant rathern than in an array.
' Useful for writing filters in VB5 with no array assignment.
Sub SortVArray(aTarget As Variant, Optional vFirst As Variant, _
              Optional vLast As Variant, Optional helper As ISortHelper)
    Dim iFirst As Long, iLast As Long
    If IsMissing(vFirst) Then iFirst = LBound(aTarget) Else iFirst = vFirst
    If IsMissing(vLast) Then iLast = UBound(aTarget) Else iLast = vLast
    If helper Is Nothing Then Set helper = New CSortHelper
    
With helper
    Dim iLo As Long, iHi As Long, iRand As Long, stack As New CStack
    Do
        Do
            ' Swap from ends until first and last meet in the middle
            If iFirst < iLast Then
                ' If we're in the middle and out of order, swap
                If iLast - iFirst = 1 Then
                    If .Compare(aTarget(iFirst), aTarget(iLast)) > 0 Then
                        .Swap aTarget(iFirst), aTarget(iLast)
                    End If
                Else
                    ' Split at some random point
                    .Swap aTarget(iLast), _
                          aTarget(MRandom.Random(iFirst, iLast))
                    ' Swap high values below the split for low values above
                    iLo = iFirst: iHi = iLast
                    Do
                        ' Find any low value larger than split
                        Do While (iLo < iHi) And _
                                 (.Compare(aTarget(iLo), aTarget(iLast)) <= 0)
                            iLo = iLo + 1
                        Loop
                        ' Find any high value smaller than split
                        Do While (iHi > iLo) And _
                                 (.Compare(aTarget(iHi), aTarget(iLast)) >= 0)
                            iHi = iHi - 1
                        Loop
                        ' Swap too high low value for too low high value
                        If iLo < iHi Then .Swap aTarget(iLo), aTarget(iHi)
                    Loop While iLo < iHi
                    ' Current (iLo) is larger than split (iLast), so swap
                    .Swap aTarget(iLo), aTarget(iLast)
                    ' Push range markers of larger part for later sorting
                    If (iLo - iFirst) < (iLast - iLo) Then
                        stack.Push iLo + 1
                        stack.Push iLast
                        iLast = iLo - 1
                    Else
                        stack.Push iFirst
                        stack.Push iLo - 1
                        iFirst = iLo + 1
                    End If
                    ' Exit from inner loop to process smaller part
                    Exit Do
                End If
            End If
            
            ' If stack empty, Exit outer loop
            If stack.Count = 0 Then Exit Sub
            ' Else pop first and last from last deferred section
            iLast = stack.Pop
            iFirst = stack.Pop
        Loop
    Loop
End With
End Sub

' QuickSort algorithm
Sub SortCollection(nTarget As Collection, Optional vFirst As Variant, _
                   Optional vLast As Variant, _
                   Optional helper As ISortHelper)
    Dim iFirst As Long, iLast As Long
    If IsMissing(vFirst) Then iFirst = 1 Else iFirst = vFirst
    If IsMissing(vLast) Then iLast = nTarget.Count Else iLast = vLast
    If helper Is Nothing Then Set helper = New CSortHelper
    
With helper
    Dim iLo As Long, iHi As Long, stack As New CStack
    Do
        Do
            ' Swap from ends until first and last meet in the middle
            If iFirst < iLast Then
                ' If we're in the middle and out of order, swap
                If iLast - iFirst = 1 Then
                    If .Compare(nTarget(iFirst), nTarget(iLast)) > 0 Then
                        .CollectionSwap nTarget, iFirst, iLast
                    End If
                Else
                    ' Split at some random point
                    .CollectionSwap nTarget, iLast, _
                                    MRandom.Random(iFirst, iLast)
                    ' Swap high values below the split for low values above
                    iLo = iFirst: iHi = iLast
                    Do
                        ' Find find any low value larger than split
                        Do While (iLo < iHi) And _
                                 (.Compare(nTarget(iLo), nTarget(iLast)) <= 0)
                            iLo = iLo + 1
                        Loop
                        ' Find any high value smaller than split
                        Do While (iHi > iLo) And _
                                 (.Compare(nTarget(iHi), nTarget(iLast)) >= 0)
                            iHi = iHi - 1
                        Loop
                        ' Swap too high low value for too low high value
                        If iLo < iHi Then .CollectionSwap nTarget, iLo, iHi
                    Loop While iLo < iHi
                    ' Current (iLo) is larger than split (iLast), so swap
                    .CollectionSwap nTarget, iLo, iLast
                    ' Push range markers of larger part for later sorting
                    If (iLo - iFirst) < (iLast - iLo) Then
                        stack.Push iLo + 1
                        stack.Push iLast
                        iLast = iLo - 1
                    Else
                        stack.Push iFirst
                        stack.Push iLo - 1
                        iFirst = iLo + 1
                    End If
                    ' Exit from inner loop to process smaller part
                    Exit Do
                End If
            End If
            
            ' If stack empty, Exit outer loop
            If stack.Count = 0 Then Exit Sub
            ' Else pop first and last from last deferred section
            iLast = stack.Pop
            iFirst = stack.Pop
        Loop
    Loop
End With
End Sub

Sub SortStringArray(asTarget() As String, _
                    Optional fHiToLo As Boolean = False, _
                    Optional cmp As VbCompareMethod = vbBinaryCompare)
    Dim iFirst As Long, iLast As Long, iSign As Long
    iFirst = LBound(asTarget)
    iLast = UBound(asTarget)
    iSign = IIf(fHiToLo, -1, 1)
    
    Dim iLo As Long, iHi As Long, iRand As Long, stack As New CStack
    Do
        Do
            ' Swap from ends until first and last meet in the middle
            If iFirst < iLast Then
                ' If we're in the middle and out of order, swap
                If iLast - iFirst = 1 Then
                    If (iSign * StrComp(asTarget(iFirst), asTarget(iLast), cmp)) > 0 Then
                        MUtility.SwapStrings asTarget(iFirst), asTarget(iLast)
                    End If
                Else
                    ' Split at some random point
                    MUtility.SwapStrings asTarget(iLast), _
                          asTarget(MRandom.Random(iFirst, iLast))
                    ' Swap high values below the split for low values above
                    iLo = iFirst: iHi = iLast
                    Do
                        ' Find any low value larger than split
                        Do While (iLo < iHi) And _
                                 ((iSign * StrComp(asTarget(iLo), asTarget(iLast), cmp)) <= 0)
                            iLo = iLo + 1
                        Loop
                        ' Find any high value smaller than split
                        Do While (iHi > iLo) And _
                                 (iSign * (StrComp(asTarget(iHi), asTarget(iLast), cmp)) >= 0)
                            iHi = iHi - 1
                        Loop
                        ' Swap too high low value for too low high value
                        If iLo < iHi Then MUtility.SwapStrings asTarget(iLo), asTarget(iHi)
                    Loop While iLo < iHi
                    ' Current (iLo) is larger than split (iLast), so swap
                    MUtility.SwapStrings asTarget(iLo), asTarget(iLast)
                    ' Push range markers of larger part for later sorting
                    If (iLo - iFirst) < (iLast - iLo) Then
                        stack.Push iLo + 1
                        stack.Push iLast
                        iLast = iLo - 1
                    Else
                        stack.Push iFirst
                        stack.Push iLo - 1
                        iFirst = iLo + 1
                    End If
                    ' Exit from inner loop to process smaller part
                    Exit Do
                End If
            End If
            
            ' If stack empty, Exit outer loop
            If stack.Count = 0 Then Exit Sub
            ' Else pop first and last from last deferred section
            iLast = stack.Pop
            iFirst = stack.Pop
        Loop
    Loop
End Sub

#If iVBVer > 5 Then
Function SortStrings(asIn() As String, _
                     Optional fHiToLo As Boolean = False, _
                     Optional cmp As VbCompareMethod = vbBinaryCompare _
                     ) As String()
    ' Assign the input array to the return value
    SortStrings = asIn
    ' Sort the return value
    SortStringArray SortStrings, fHiToLo, cmp
End Function
#End If

' Alas, Dictionary.Keys is an array of Variants rather than an array of Strings
Function SortKeys(vas As Variant, _
                  Optional fHiToLo As Boolean = False, _
                  Optional cmp As VbCompareMethod = vbBinaryCompare _
                  ) As Variant
    If (VarType(vas) And vbArray) = 0 Then ErrRaise eeTypeMismatch
    ' Assign the input array to the return value
#If iVBVer > 5 Then
    Dim avsRet() As Variant
    avsRet = vas
    ' Sort the return value
    Dim helper As New CSortHelper
    helper.SortMode = IIf(cmp = vbTextCompare, esmSortText, esmSortBin)
    helper.HiToLo = fHiToLo
    SortArray avsRet, , , helper
    SortKeys = avsRet
#Else
    ' Hack with Variants since VB5 doesn't support array assignment
    Dim avsRet As Variant
    avsRet = vas
    ' Sort the return value
    Dim helper As New CSortHelper
    helper.SortMode = IIf(cmp = vbTextCompare, esmSortText, esmSortBin)
    helper.HiToLo = fHiToLo
    SortVArray avsRet, , , helper
    SortKeys = avsRet
#End If
End Function

Function BSearchArray(av() As Variant, ByVal vKey As Variant, _
                      iPos As Long, _
                      Optional helper As ISortHelper) As Boolean
    Dim iLo As Long, iHi As Long
    Dim iComp As Long, iMid As Long
    If helper Is Nothing Then Set helper = New CSortHelper
    
    iLo = LBound(av): iHi = UBound(av)
    Do
        iMid = iLo + ((iHi - iLo) \ 2)
        iComp = helper.Compare(av(iMid), vKey)
        Select Case iComp
        Case 0
            ' Item found
            iPos = iMid
            BSearchArray = True
            Exit Function
        Case Is > 0
            ' Item is in lower half
            iHi = iMid - 1
            If iHi < iLo Then Exit Do
        Case Is < 0
            ' Item is in upper half
            iLo = iMid + 1
            If iLo > iHi Then Exit Do
        End Select
    Loop
    ' Item not found, but return position to insert
    iPos = iMid - (iComp < 0)
        
End Function

Function BSearchStrings(asIn() As String, sKey As String, _
                        iPos As Long, _
                        Optional cmp As VbCompareMethod = vbBinaryCompare _
                        ) As Boolean
    Dim iLo As Long, iHi As Long
    Dim iComp As Long, iMid As Long
    
    iLo = LBound(asIn): iHi = UBound(asIn)
    Do
        iMid = iLo + ((iHi - iLo) \ 2)
        iComp = StrComp(asIn(iMid), sKey)
        Select Case iComp
        Case 0
            ' Item found
            iPos = iMid
            BSearchStrings = True
            Exit Function
        Case Is > 0
            ' Item is in lower half
            iHi = iMid - 1
            If iHi < iLo Then Exit Do
        Case Is < 0
            ' Item is in upper half
            iLo = iMid + 1
            If iLo > iHi Then Exit Do
        End Select
    Loop
    ' Item not found, but return position to insert
    iPos = iMid - (iComp < 0)
        
End Function

' BSearchCollection performs a binary search on a collection and
' returns True or False depending on whether the search item is
' found. BSearchCollection also returns the index of the search
' item in iPos. If the item isn't found, iPos will contain the
' index that the item should occupy in the collection. Note that
' iPos will equal 1 if the collection is empty, and will equal
' n.Count + 1 if the search item should be inserted at the end
' of the collection.
'
' The following example uses BSearchCollection to insert an item
' in sorted order:
'
'    Dim n as new Collection, v As Variant, iPos As Long
'
'    v = InputBox("Collection item to insert: ")
'    ' Insert item in collection if item doesn't already exist
'    If Not BSearchCollection(n, v, iPos) Then
'        On Error GoTo IndexError
'        ' The following line of code generates an error if the
'        ' collection is empty or iPos > n.Count. In either case,
'        ' the error handler adds the item to the end of the collection
'        n.Add v, , iPos
'    End If
'
'    Exit Sub
'IndexError:
'    ' Item needs to be inserted at end of collection
'    n.Add v

Function BSearchCollection(n As Collection, ByVal vKey As Variant, _
                           iPos As Long, _
                           Optional helper As ISortHelper) As Boolean
    Dim iLo As Long, iHi As Long
    Dim iComp As Long, iMid As Long
    If helper Is Nothing Then Set helper = New CSortHelper
    
    ' Special case if empty collection
    If n.Count = 0 Then
        iPos = 1
        Exit Function
    End If
    
    iLo = 1: iHi = n.Count
    Do
        iMid = iLo + ((iHi - iLo) \ 2)
        iComp = helper.Compare(n(iMid), vKey)
        Select Case iComp
        Case 0
            ' Item found
            iPos = iMid
            BSearchCollection = True
            Exit Function
        Case Is > 0
            ' Item is in lower half
            iHi = iMid - 1
            If iHi < iLo Then Exit Do
        Case Is < 0
            ' Item is in upper half
            iLo = iMid + 1
            If iLo > iHi Then Exit Do
        End Select
    Loop
    ' Item not found, but return position to insert
    iPos = iMid - (iComp < 0)
    
End Function

Sub ShuffleArray(av() As Variant, Optional helper As ISortHelper)
    Dim iFirst As Long, iLast As Long
    If helper Is Nothing Then Set helper = New CSortHelper
    
    iFirst = LBound(av): iLast = UBound(av)
    ' Randomize array
    Dim i As Long, v As Variant, iRnd As Long
    For i = iLast To iFirst + 1 Step -1
        ' Swap random element with last element
        iRnd = MRandom.Random(iFirst, i)
        helper.Swap av(i), av(iRnd)
    Next
End Sub

Sub ShuffleVArray(av As Variant, Optional helper As ISortHelper)
    Dim iFirst As Long, iLast As Long
    If helper Is Nothing Then Set helper = New CSortHelper
    
    iFirst = LBound(av): iLast = UBound(av)
    ' Randomize array
    Dim i As Long, v As Variant, iRnd As Long
    For i = iLast To iFirst + 1 Step -1
        ' Swap random element with last element
        iRnd = MRandom.Random(iFirst, i)
        helper.Swap av(i), av(iRnd)
    Next
End Sub

Sub ShuffleCollection(n As Collection, Optional helper As ISortHelper)
    Dim iFirst As Long, iLast As Long
    If helper Is Nothing Then Set helper = New CSortHelper
    
    iFirst = 1: iLast = n.Count
    ' Randomize collection
    Dim i As Long, v As Variant, iRnd As Long
    For i = iLast To iFirst + 1 Step -1
        ' Swap random element with last element
        iRnd = MRandom.Random(iFirst, i)
        helper.CollectionSwap n, i, iRnd
    Next
End Sub

Sub ShuffleStringArray(asIn() As String)
    Dim iFirst As Long, iLast As Long
    
    iFirst = LBound(asIn): iLast = UBound(asIn)
    ' Randomize array
    Dim i As Long, s As String, iRnd As Long
    For i = iLast To iFirst + 1 Step -1
        ' Swap random element with last element
        iRnd = MRandom.Random(iFirst, i)
        MUtility.SwapStrings asIn(i), asIn(iRnd)
    Next
End Sub

#If iVBVer > 5 Then
Function ShuffleStrings(asIn() As String) As String()
    ' Assign the input array to the return value
    ShuffleStrings = asIn
    ' Sort the return value
    ShuffleStringArray ShuffleStrings
End Function
#End If

Function ShuffleKeys(va As Variant) As Variant
#If iVBVer > 5 Then
    If (VarType(va) And vbArray) = 0 Then ErrRaise eeTypeMismatch
    ' Assign the input array to the return value
    Dim avsRet() As Variant
    avsRet = va
    ' Shuffle the return value
    ShuffleArray avsRet
    ' Return it
    ShuffleKeys = avsRet
#Else
    ' Hack with Variants since VB5 doesn't support array assignment
    If (VarType(va) And vbArray) = 0 Then ErrRaise eeTypeMismatch
    ' Assign the input array to the return value
    Dim avsRet As Variant
    avsRet = va
    ' Shuffle the return value
    ShuffleVArray avsRet
    ' Return it
    ShuffleKeys = avsRet
#End If
End Function

#If fComponent = 0 Then
Private Sub ErrRaise(e As Long)
    Dim sText As String, sSource As String
    If e > 1000 Then
        sSource = App.ExeName & ".Sort"
        Select Case e
        Case eeBaseSort
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


