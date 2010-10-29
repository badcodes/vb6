Attribute VB_Name = "MSortOld"
Option Explicit

'$ Uses UTILITY.BAS

' Old iterative QuickSort algorithm
Public Sub SortArrayO(aTarget() As Variant, _
                      Optional vFirst As Variant, Optional vLast As Variant)
    Dim iFirst As Long, iLast As Long
    If IsMissing(vFirst) Then iFirst = LBound(aTarget) Else iFirst = vFirst
    If IsMissing(vFirst) Then iLast = UBound(aTarget) Else iLast = vLast
    
    Dim iLo As Long, iHi As Long, stack As New CStack
    Do
        Do
            ' Keep swapping from ends until first and last meet in the middle
            If iFirst < iLast Then
                ' If we're in the middle and out of order, swap
                If iLast - iFirst = 1 Then
                    If SortCompare(aTarget(iFirst), aTarget(iLast)) > 0 Then
                        SortSwap aTarget(iFirst), aTarget(iLast)
                    End If
                Else
                    ' Split at some random point
                    SortSwap aTarget(iLast), aTarget(Random(iFirst, iLast))
                    ' Swap high values below the split for high values above
                    iLo = iFirst: iHi = iLast
                    Do
                        ' Find find any low value larger than split
                        Do While (iLo < iHi) And SortCompare(aTarget(iLo), aTarget(iLast)) <= 0
                            iLo = iLo + 1
                        Loop
                        ' Find any high value smaller than split
                        Do While (iHi > iLo) And SortCompare(aTarget(iHi), aTarget(iLast)) >= 0
                            iHi = iHi - 1
                        Loop
                        ' Swap the too high low value for the too low high value
                        If iLo < iHi Then SortSwap aTarget(iLo), aTarget(iHi)
                    Loop While iLo < iHi
                    ' Current item (iLo) is always larger than split (iLast), so swap
                    SortSwap aTarget(iLo), aTarget(iLast)
                    ' Push range markers of smaller part for later processing
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

' Old recursive QuickSort algorithm
Sub SortArrayRecO(aTarget() As Variant, _
              iFirst As Long, iLast As Long)
    If iFirst < iLast Then

        ' Only two elements in this subdivision; exchange if
        ' they are out of order, and end recursive calls
        If iLast - iFirst = 1 Then
            If SortCompare(aTarget(iFirst), aTarget(iLast)) > 0 Then
                SortSwap aTarget(iFirst), aTarget(iLast)
            End If
        Else

            Dim i As Long, j As Long

            ' Pick pivot element at random and move to end
            SortSwap aTarget(iLast), aTarget(Random(iFirst, iLast))
            i = iFirst: j = iLast
            Do

                ' Move in from both sides toward pivot element
                Do While (i < j) And _
                         SortCompare(aTarget(i), aTarget(iLast)) <= 0
                    i = i + 1
                Loop
                Do While (j > i) And _
                         SortCompare(aTarget(j), aTarget(iLast)) >= 0
                    j = j - 1
                Loop

                ' If you haven't reached pivot element, it means
                ' that the two elements on either side are out of
                ' order, so swap them
                If i < j Then
                    SortSwap aTarget(i), aTarget(j)
                End If
            Loop While i < j

            ' Move pivot element back to its proper place
            SortSwap aTarget(i), aTarget(iLast)

            ' Recursively call SortArrayO (pass smaller
            ' subdivision first to use less stack space)
            If (i - iFirst) < (iLast - i) Then
                SortArrayRecO aTarget(), iFirst, i - 1
                SortArrayRecO aTarget(), i + 1, iLast
            Else
                SortArrayRecO aTarget(), i + 1, iLast
                SortArrayRecO aTarget(), iFirst, i - 1
            End If
        End If
    End If

End Sub

' QuickSort algorithm
Sub SortCollectionO(nTarget As Collection, iFirst As Long, iLast As Long)
    If iFirst < iLast Then

        ' Only two elements in this subdivision; exchange if
        ' they are out of order, and end recursive calls
        If iLast - iFirst = 1 Then
            If SortCompare(nTarget(iFirst), nTarget(iLast)) > 0 Then
                CollectionSwap nTarget, iFirst, iLast
            End If
        Else

            Dim i As Long, j As Long

            ' Pick pivot element at random and move to end
            CollectionSwap nTarget, iLast, Random(iFirst, iLast)
            i = iFirst: j = iLast
            Do

                ' Move in from both sides toward pivot element
                Do While (i < j) And _
                    SortCompare(nTarget(i), nTarget(iLast)) <= 0
                    i = i + 1
                Loop
                Do While (j > i) And _
                    SortCompare(nTarget(j), nTarget(iLast)) >= 0
                    j = j - 1
                Loop

                ' If you haven't reached pivot element, it means
                ' that the two elements on either side are out of
                ' order, so swap them
                If i < j Then
                    CollectionSwap nTarget, i, j
                End If
            Loop While i < j

            ' Move pivot element back to its proper place
            CollectionSwap nTarget, i, iLast

            ' Recursively call SortCollectionO (pass smaller
            ' subdivision first to use less stack space)
            If (i - iFirst) < (iLast - i) Then
                SortCollectionO nTarget, iFirst, i - 1
                SortCollectionO nTarget, i + 1, iLast
            Else
                SortCollectionO nTarget, i + 1, iLast
                SortCollectionO nTarget, iFirst, i - 1
            End If
        End If
    End If

End Sub

Function BSearchArrayO(av() As Variant, vKey As Variant, _
                      iPos As Long) As Boolean
    Dim iLo As Long, iHi As Long
    Dim iComp As Long, iMid As Long
    iLo = LBound(av): iHi = UBound(av)
    Do
        iMid = iLo + ((iHi - iLo) \ 2)
        iComp = SortCompare(av(iMid), vKey)
        Select Case iComp
        Case 0
            ' Item found
            iPos = iMid
            BSearchArrayO = True
            Exit Function
        Case Is > 0
            ' Item is in lower half
            iHi = iMid
            If iLo = iHi Then Exit Do
        Case Is < 0
            ' Item is in upper half
            iLo = iMid + 1
            If iLo > iHi Then Exit Do
        End Select
    Loop
    ' Item not found, but return position to insert
    iPos = iMid - (iComp < 0)
    BSearchArrayO = False
        
End Function

Function BSearchCollectionO(n As Collection, vKey As Variant, _
                           iPos As Long) As Boolean
    Dim iLo As Long, iHi As Long
    Dim iComp As Long, iMid As Long
    iLo = 1: iHi = n.Count
    Do
        iMid = iLo + ((iHi - iLo) \ 2)
        iComp = SortCompare(n(iMid), vKey)
        Select Case iComp
        Case 0
            ' Item found
            iPos = iMid
            BSearchCollectionO = True
            Exit Function
        Case Is > 0
            ' Item is in lower half
            iHi = iMid
            If iLo = iHi Then Exit Do
        Case Is < 0
            ' Item is in upper half
            iLo = iMid + 1
            If iLo > iHi Then Exit Do
        End Select
    Loop
    ' Item not found, but return position to insert
    iPos = iMid - (iComp < 0)
    BSearchCollectionO = False
        
End Function

Sub ShuffleArrayO(av() As Variant)
    Dim iFirst As Long, iLast As Long
    iFirst = LBound(av): iLast = UBound(av)
        
    ' Randomize array
    Dim i As Long, v As Variant, iRnd As Long
    For i = iLast To iFirst + 1 Step -1
        ' Swap random element with last element
        iRnd = Random(iFirst, i)
        SortSwap av(i), av(iRnd)
    Next
End Sub

Sub ShuffleCollectionO(n As Collection)
    Dim iFirst As Long, iLast As Long
    iFirst = 1: iLast = n.Count
    
    ' Randomize collection
    Dim i As Long, v As Variant, iRnd As Long
    For i = iLast To iFirst + 1 Step -1
        ' Swap random element with last element
        iRnd = Random(iFirst, i)
        CollectionSwap n, i, iRnd
    Next
End Sub

' Define fSortCompareDef to use default SortCompare
#If fSortCompareDef Then
Private Function SortCompare(ByVal v1 As Variant, _
                             ByVal v2 As Variant) As Long
    If v1 < v2 Then
        SortCompare = -1
    ElseIf v1 = v2 Then
        SortCompare = 0
    Else
        SortCompare = 1
    End If
End Function
#End If

' Define fSortSwapNoDef if you provide your own swap routine
#If fSortSwapNoDef = 0 Then
Sub SortSwap(v1 As Variant, v2 As Variant)
    Dim vT As Variant
    vT = v1
    v1 = v2
    v2 = vT
End Sub
#End If

Sub CollectionSwap(n As Collection, i1 As Long, i2 As Long)
    Dim vT As Variant
    vT = n(i1)
    n.Add n(i2), , , i1
    n.Remove i1
    n.Add vT, , , i2
    n.Remove i2
End Sub
