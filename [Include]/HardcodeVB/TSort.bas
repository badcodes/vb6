Attribute VB_Name = "MTestSort"
Option Explicit


Public esmMode As Integer
Public fSortHiToLo As Integer

' Modify to add more sort modes
Public Function SortCompare(v1 As Variant, v2 As Variant) As Integer
    ' Use string comparisons only on strings
    If TypeName(v1) <> "String" Then esmMode = esmSortVal
    
    Dim i As Integer
    Select Case esmMode
    ' Sort by value (same as esmSortBin for strings)
    Case esmSortVal
        If v1 < v2 Then
            i = -1
        ElseIf v1 = v2 Then
            i = 0
        Else
            i = 1
        End If
    ' Sort case-insensitive
    Case esmSortText
        i = StrComp(v1, v2, 1)
    ' Sort case-sensitive
    Case esmSortbin
        i = StrComp(v1, v2, 0)
    ' Sort by string length
    Case esmSortLen
        If Len(v1) = Len(v2) Then
            If v1 = v2 Then
                i = 0
            ElseIf v1 < v2 Then
                i = -1
            Else
                i = 1
            End If
        ElseIf Len(v1) < Len(v2) Then
            i = -1
        Else
            i = 1
        End If
    End Select
    If fSortHiToLo Then i = -i
    SortCompare = i
End Function
'
