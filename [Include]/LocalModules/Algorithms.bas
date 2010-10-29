Attribute VB_Name = "MAlgorithms"

Option Explicit
'Under CVS Controls


'打乱数组
'FIXIT: Declare 'v1' and 'v2' with an early-bound data type                                FixIT90210ae-R1672-R1B8ZE
Public Sub SwapV(ByRef v1, ByRef v2)
'FIXIT: Declare 'vtmp' with an early-bound data type                                       FixIT90210ae-R1672-R1B8ZE
Dim vtmp As Variant
vtmp = v1
v1 = v2
v2 = vtmp
End Sub

Public Function getRndnum(iFrom As Long, iTo As Long) As Long
If iFrom > iTo Then SwapV iFrom, iTo
getRndnum = Int((iTo - iFrom + 1) * Rnd + iFrom)
End Function

'打乱数组
'只支持一维数组
'low 和 high 为打乱数组的范围
'low 和 high 默认为-1,会将low转为数组的下限,high转为数组的上限
'low>high 的情况将low和high的值调换
'FIXIT: Declare 'arrV' with an early-bound data type                                       FixIT90210ae-R1672-R1B8ZE
Public Sub BedlamArr(ByRef arrV As Variant, Optional ByVal low As Long = -1, Optional ByVal high As Long = -1)
Dim lLb As Long
Dim lUB As Long
Dim iRnd As Long
'判断是否为数组
If IsArray(arrV) = False Then Exit Sub
'修正参数
If low > high Then SwapV low, high
lLb = LBound(arrV)
lUB = UBound(arrV)
If low = -1 Or low < lLb Then low = lLb
If high = -1 Or high > lUB Then high = lUB
If low = high Then Exit Sub
low = low + 1
Randomize
Do While (low <= high)
    iRnd = getRndnum(low, high)
    SwapV arrV(low - 1), arrV(iRnd)
    low = low + 1
Loop

End Sub


'FIXIT: Declare 'vData' with an early-bound data type                                      FixIT90210ae-R1672-R1B8ZE
Public Sub QuickSort(vData As Variant, low As Long, Hi As Long)

    ' ---------------------------------------------------------
    '
    ' Syntax:     QuickSort TmpAray(), Low, Hi
    '
    ' Parameters:
    '     vData - A variant pointing to an array to be sorted.
    '       Low - LBounds(vData) low number of elements in the array
    '        Hi - UBounds(vData) high number of elements in the array
    '
    ' NOTE:       I start my arrays with one and not zero.
    '             Make the appropriate changes to suit your code.
    ' ---------------------------------------------------------
    ' ---------------------------------------------------------
    ' Test to see if an array was passed
    ' ---------------------------------------------------------

    If Not IsArray(vData) Then Exit Sub
    ' ---------------------------------------------------------
    ' Define local variables
    ' ---------------------------------------------------------
    Dim lTmpLow As Long
    Dim lTmpHi As Long
    Dim lTmpMid As Long
'FIXIT: Declare 'vTempVal' with an early-bound data type                                   FixIT90210ae-R1672-R1B8ZE
    Dim vTempVal As Variant
'FIXIT: Declare 'vTmpHold' with an early-bound data type                                   FixIT90210ae-R1672-R1B8ZE
    Dim vTmpHold As Variant
    ' ---------------------------------------------------------
    ' Initialize local variables
    ' ---------------------------------------------------------
    lTmpLow = low
    lTmpHi = Hi
    ' ---------------------------------------------------------
    ' Leave if there is nothing to sort
    ' ---------------------------------------------------------

    If Hi <= low Then Exit Sub
    ' ---------------------------------------------------------
    ' Find the middle to start comparing values
    ' ---------------------------------------------------------
    lTmpMid = (low + Hi) \ 2
    ' ---------------------------------------------------------
    ' Move the item in the middle of the array to the
    ' temporary holding area as a point of reference while
    ' sorting.  This will change each time we make a recursive
    ' call to this routine.
    ' ---------------------------------------------------------
    vTempVal = vData(lTmpMid)
    ' ---------------------------------------------------------
    ' Loop until we eventually meet in the middle
    ' ---------------------------------------------------------

    Do While (lTmpLow <= lTmpHi)
        ' Always process the low end first.  Loop as long
        ' the array data element is less than the data in
        ' the temporary holding area and the temporary low
        ' value is less than the maximum number of array
        ' elements.

        Do While (vData(lTmpLow) < vTempVal And lTmpLow < Hi)
            lTmpLow = lTmpLow + 1
        Loop

        ' Now, we will process the high end.  Loop as long
        ' the data in the temporary holding area is less
        ' than the array data element and the temporary high
        ' value is greater than the minimum number of array
        ' elements.

        Do While (vTempVal < vData(lTmpHi) And lTmpHi > low)
            lTmpHi = lTmpHi - 1
        Loop

        ' if the temp low end is less than or equal
        ' to the temp high end, then swap places

        If (lTmpLow <= lTmpHi) Then
            vTmpHold = vData(lTmpLow)          ' Move the Low value to Temp Hold
            vData(lTmpLow) = vData(lTmpHi)     ' Move the high value to the low
            vData(lTmpHi) = vTmpHold           ' move the Temp Hod to the High
            lTmpLow = lTmpLow + 1              ' Increment the temp low counter
            lTmpHi = lTmpHi - 1                ' Dcrement the temp high counter
        End If

    Loop

    ' ---------------------------------------------------------
    ' If the minimum number of elements in the array is
    ' less than the temp high end, then make a recursive
    ' call to this routine.  I always sort the low end
    ' of the array first.
    ' ---------------------------------------------------------

    If (low < lTmpHi) Then
        QuickSort vData, low, lTmpHi
    End If

    ' ---------------------------------------------------------
    ' If the temp low end is less than the maximum number
    ' of elements in the array, then make a recursive call
    ' to this routine.  The high end is always sorted last.
    ' ---------------------------------------------------------

    If (lTmpLow < Hi) Then
        QuickSort vData, lTmpLow, Hi
    End If

End Sub

Public Sub QuickSortFiles(vData() As String, low As Long, Hi As Long)

    ' ---------------------------------------------------------
    '
    ' Syntax:     QuickSort TmpAray(), Low, Hi
    '
    ' Parameters:
    '     vData - A variant pointing to an array to be sorted.
    '       Low - LBounds(vData) low number of elements in the array
    '        Hi - UBounds(vData) high number of elements in the array
    '
    ' NOTE:       I start my arrays with one and not zero.
    '             Make the appropriate changes to suit your code.
    ' ---------------------------------------------------------
    ' ---------------------------------------------------------
    ' Test to see if an array was passed
    ' ---------------------------------------------------------

    If Not IsArray(vData) Then Exit Sub
    ' ---------------------------------------------------------
    ' Define local variables
    ' ---------------------------------------------------------
    Dim lTmpLow As Long
    Dim lTmpHi As Long
    Dim lTmpMid As Long
    Dim vTempVal As String
    Dim vTmpHold As String
    ' ---------------------------------------------------------
    ' Initialize local variables
    ' ---------------------------------------------------------
    lTmpLow = low
    lTmpHi = Hi
    ' ---------------------------------------------------------
    ' Leave if there is nothing to sort
    ' ---------------------------------------------------------

    If Hi <= low Then Exit Sub
    ' ---------------------------------------------------------
    ' Find the middle to start comparing values
    ' ---------------------------------------------------------
    lTmpMid = (low + Hi) \ 2
    ' ---------------------------------------------------------
    ' Move the item in the middle of the array to the
    ' temporary holding area as a point of reference while
    ' sorting.  This will change each time we make a recursive
    ' call to this routine.
    ' ---------------------------------------------------------
    vTempVal = vData(lTmpMid)
    ' ---------------------------------------------------------
    ' Loop until we eventually meet in the middle
    ' ---------------------------------------------------------

    Do While (lTmpLow <= lTmpHi)
        ' Always process the low end first.  Loop as long
        ' the array data element is less than the data in
        ' the temporary holding area and the temporary low
        ' value is less than the maximum number of array
        ' elements.

        Do While (slashCountInstr(vData(lTmpLow)) < slashCountInstr(vTempVal) And lTmpLow < Hi)
            lTmpLow = lTmpLow + 1
        Loop

        ' Now, we will process the high end.  Loop as long
        ' the data in the temporary holding area is less
        ' than the array data element and the temporary high
        ' value is greater than the minimum number of array
        ' elements.

        Do While (slashCountInstr(vTempVal) < slashCountInstr(vData(lTmpHi)) And lTmpHi > low)
            lTmpHi = lTmpHi - 1
        Loop

        ' if the temp low end is less than or equal
        ' to the temp high end, then swap places

        If (lTmpLow <= lTmpHi) Then
            vTmpHold = vData(lTmpLow)          ' Move the Low value to Temp Hold
            vData(lTmpLow) = vData(lTmpHi)     ' Move the high value to the low
            vData(lTmpHi) = vTmpHold           ' move the Temp Hod to the High
            lTmpLow = lTmpLow + 1              ' Increment the temp low counter
            lTmpHi = lTmpHi - 1                ' Dcrement the temp high counter
        End If

    Loop

    ' ---------------------------------------------------------
    ' If the minimum number of elements in the array is
    ' less than the temp high end, then make a recursive
    ' call to this routine.  I always sort the low end
    ' of the array first.
    ' ---------------------------------------------------------

    If (low < lTmpHi) Then
        QuickSortFiles vData, low, lTmpHi
    End If

    ' ---------------------------------------------------------
    ' If the temp low end is less than the maximum number
    ' of elements in the array, then make a recursive call
    ' to this routine.  The high end is always sorted last.
    ' ---------------------------------------------------------

    If (lTmpLow < Hi) Then
        QuickSortFiles vData, lTmpLow, Hi
    End If

End Sub

