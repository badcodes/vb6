Attribute VB_Name = "MSortRecursive"
Option Explicit

'$ Uses UTILITY.BAS

' Recursive QuickSort algorithm
Sub SortArrayRec(aTarget() As Variant, _
                 Optional vFirst As Variant, Optional vLast As Variant, _
                 Optional helper As ISortHelper)
    Dim iFirst As Long, iLast As Long
    If IsMissing(vFirst) Then iFirst = LBound(aTarget) Else iFirst = vFirst
    If IsMissing(vLast) Then iLast = UBound(aTarget) Else iLast = vLast
    If helper Is Nothing Then Set helper = New CSortHelper
    
With helper
    If iFirst < iLast Then

        ' Only two elements in this subdivision; exchange if
        ' they are out of order, and end recursive calls
        If iLast - iFirst = 1 Then
            If .Compare(aTarget(iFirst), aTarget(iLast)) > 0 Then
                .Swap aTarget(iFirst), aTarget(iLast)
            End If
        Else

            Dim iLo As Long, iHi As Long
            ' Pick pivot element at random and move to end
            .Swap aTarget(iLast), aTarget(Random(iFirst, iLast))
            iLo = iFirst: iHi = iLast
            Do

                ' Move in from both sides toward pivot element
                Do While (iLo < iHi) And _
                         .Compare(aTarget(iLo), aTarget(iLast)) <= 0
                    iLo = iLo + 1
                Loop
                Do While (iHi > iLo) And _
                         .Compare(aTarget(iHi), aTarget(iLast)) >= 0
                    iHi = iHi - 1
                Loop

                ' If you haven't reached pivot element, it means
                ' that two elements on either side are out of
                ' order, so swap them
                If iLo < iHi Then .Swap aTarget(iLo), aTarget(iHi)
            Loop While iLo < iHi

            ' Move pivot element back to its proper place
            .Swap aTarget(iLo), aTarget(iLast)

            ' Recursively call SortArrayRec (pass smaller
            ' subdivision first to use less stack space)
            If (iLo - iFirst) < (iLast - iLo) Then
                SortArrayRec aTarget(), iFirst, iLo - 1, helper
                SortArrayRec aTarget(), iLo + 1, iLast, helper
            Else
                SortArrayRec aTarget(), iLo + 1, iLast, helper
                SortArrayRec aTarget(), iFirst, iLo - 1, helper
            End If
        End If
    End If
End With
End Sub

' Recursive QuickSort algorithm
Sub SortCollectionRec(nTarget As Collection, _
                      Optional vFirst As Variant, _
                      Optional vLast As Variant, _
                      Optional helper As ISortHelper)
    Dim iFirst As Long, iLast As Long
    If IsMissing(vFirst) Then iFirst = 1 Else iFirst = vFirst
    If IsMissing(vLast) Then iLast = nTarget.Count Else iLast = vLast
    If helper Is Nothing Then Set helper = New CSortHelper

With helper
    If iFirst < iLast Then

        ' Only two elements in this subdivision; exchange if
        ' they are out of order, and end recursive calls
        If iLast - iFirst = 1 Then
            If .Compare(nTarget(iFirst), nTarget(iLast)) > 0 Then
                .CollectionSwap nTarget, iFirst, iLast
            End If
        Else

            Dim iLo As Long, iHi As Long
            ' Pick pivot element at random and move to end
            .CollectionSwap nTarget, iLast, Random(iFirst, iLast)
            iLo = iFirst: iHi = iLast
            Do

                ' Move in from both sides toward pivot element
                Do While (iLo < iHi) And _
                    .Compare(nTarget(iLo), nTarget(iLast)) <= 0
                    iLo = iLo + 1
                Loop
                Do While (iHi > iLo) And _
                    .Compare(nTarget(iHi), nTarget(iLast)) >= 0
                    iHi = iHi - 1
                Loop

                ' If you haven't reached pivot element, it means
                ' that the two elements on either side are out of
                ' order, so swap them
                If iLo < iHi Then
                    .CollectionSwap nTarget, iLo, iHi
                End If
            Loop While iLo < iHi

            ' Move pivot element back to its proper place
            .CollectionSwap nTarget, iLo, iLast

            ' Recursively call SortCollection (pass smaller
            ' subdivision first to use less stack space)
            If (iLo - iFirst) < (iLast - iLo) Then
                SortCollectionRec nTarget, iFirst, iLo - 1, helper
                SortCollectionRec nTarget, iLo + 1, iLast, helper
            Else
                SortCollectionRec nTarget, iLo + 1, iLast, helper
                SortCollectionRec nTarget, iFirst, iLo - 1, helper
            End If
        End If
    End If
End With
End Sub

