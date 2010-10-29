Attribute VB_Name = "modSort"
Option Explicit

'
' This code is based on the execelent array sorting and searching algorithm
' module: mdArray.bas by Philippe Lord // Marton, Email: StromgaldMarton@Hotmail.com
'
' James Richardson
' JamesRichardson7@Compuserve.com
'

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef lpDest As Any, ByRef lpSource As Any, ByVal iLen As Long)


Public Sub sTriQuickSortString(ByRef sArray() As String)
   Dim iLBound As Long
   Dim iUBound As Long
   On Error GoTo sTriQuickSortString_ERROR
   
   iLBound = LBound(sArray, 2)
   iUBound = UBound(sArray, 2)
   
   sTriQuickSortString2 sArray, 4, iLBound, iUBound
   sInsertionSortString sArray, iLBound, iUBound
   Exit Sub
   
sTriQuickSortString_ERROR:
   Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub sCompactString(ByRef sArray() As String)
Dim i As Long, j As Long, k As Long

' Compact array.
For i = 0 To UBound(sArray, 2) - 1
    If Len(Trim$(sArray(1, i))) = 0 Then
        For j = i To UBound(sArray, 2) - 1
            sArray(1, j) = sArray(1, j + 1)
            sArray(1, j + 1) = vbNullString
            sArray(2, j) = sArray(2, j + 1)
            sArray(2, j + 1) = vbNullString
        Next j
    End If
Next i

For i = 0 To UBound(sArray, 2) - 1
    If Len(Trim$(sArray(1, i))) > 0 Then k = k + 1
Next

If UBound(sArray, 2) > k Then
    ReDim Preserve sArray(1 To 2, 0 To k)
End If

Exit Sub

End Sub

Private Sub sTriQuickSortString2(ByRef sArray() As String, ByVal iSplit As Long, ByVal iMin As Long, ByVal iMax As Long)
   Dim i      As Long
   Dim j      As Long
   Dim sTemp1 As String
   Dim sTemp2 As String
   
   If (iMax - iMin) > iSplit Then
      i = (iMax + iMin) / 2
      
      If sArray(1, iMin) > sArray(1, i) Then
         sSwapStrings sArray(1, iMin), sArray(1, i)
         sSwapStrings sArray(2, iMin), sArray(2, i)
      End If
      
      If sArray(1, iMin) > sArray(1, iMax) Then
         sSwapStrings sArray(1, iMin), sArray(1, iMax)
         sSwapStrings sArray(2, iMin), sArray(2, iMax)
      End If
      
      If sArray(1, i) > sArray(1, iMax) Then
         sSwapStrings sArray(1, i), sArray(1, iMax)
         sSwapStrings sArray(2, i), sArray(2, iMax)
      End If
      
      j = iMax - 1
      sSwapStrings sArray(1, i), sArray(1, j)
      sSwapStrings sArray(2, i), sArray(2, j)
      i = iMin
      CopyMemory ByVal VarPtr(sTemp1), ByVal VarPtr(sArray(1, j)), 4
      CopyMemory ByVal VarPtr(sTemp2), ByVal VarPtr(sArray(2, j)), 4
      
      Do
         Do
            i = i + 1
         Loop While sArray(1, i) < sTemp1
         
         Do
            j = j - 1
         Loop While sArray(1, j) > sTemp1
         
         If j < i Then Exit Do
         sSwapStrings sArray(1, i), sArray(1, j)
         sSwapStrings sArray(2, i), sArray(2, j)
      Loop
      
      sSwapStrings sArray(1, i), sArray(1, iMax - 1)
      sSwapStrings sArray(2, i), sArray(2, iMax - 1)
      
      sTriQuickSortString2 sArray, iSplit, iMin, j
      sTriQuickSortString2 sArray, iSplit, i + 1, iMax
   End If
   
   i = 0
   CopyMemory ByVal VarPtr(sTemp1), ByVal VarPtr(i), 4
   CopyMemory ByVal VarPtr(sTemp2), ByVal VarPtr(i), 4
End Sub

Private Sub sInsertionSortString(ByRef sArray() As String, ByVal iMin As Long, ByVal iMax As Long)
   Dim i      As Long
   Dim j      As Long
   Dim sTemp1 As String
   Dim sTemp2 As String
   
   For i = iMin + 1 To iMax
      CopyMemory ByVal VarPtr(sTemp1), ByVal VarPtr(sArray(1, i)), 4
      CopyMemory ByVal VarPtr(sTemp2), ByVal VarPtr(sArray(2, i)), 4
      j = i
      
      Do While j > iMin
         If sArray(1, j - 1) <= sTemp1 Then Exit Do

         CopyMemory ByVal VarPtr(sArray(1, j)), ByVal VarPtr(sArray(1, j - 1)), 4
         CopyMemory ByVal VarPtr(sArray(2, j)), ByVal VarPtr(sArray(2, j - 1)), 4
         j = j - 1
      Loop
      
      CopyMemory ByVal VarPtr(sArray(1, j)), ByVal VarPtr(sTemp1), 4
      CopyMemory ByVal VarPtr(sArray(2, j)), ByVal VarPtr(sTemp2), 4
   Next i
   
   i = 0
   CopyMemory ByVal VarPtr(sTemp1), ByVal VarPtr(i), 4
   CopyMemory ByVal VarPtr(sTemp2), ByVal VarPtr(i), 4
End Sub

Private Sub sSwapStrings(ByRef s1 As String, ByRef s2 As String)
   Dim i As Long

   i = StrPtr(s1)
   If i = 0 Then CopyMemory ByVal VarPtr(i), ByVal VarPtr(s1), 4

   CopyMemory ByVal VarPtr(s1), ByVal VarPtr(s2), 4
   CopyMemory ByVal VarPtr(s2), i, 4
End Sub

