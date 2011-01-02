VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSortHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ISortHelper

Public Enum EErrorSortHelper
    eeBaseSortHelper = 13220    ' CSortHelper
    eeMissingKey                ' Key is missing
End Enum

Enum ESortMode
    esmUnsorted = -1
    esmSortVal
    esmSortText
    esmSortBin
    esmSortLen
End Enum

Private esmMode As Integer
Private fHiToLo As Boolean

' Modify to add more sort modes
Private Function ISortHelper_Compare(v1 As Variant, _
                                     v2 As Variant) As Integer
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
    Case esmSortBin
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
    If fHiToLo Then i = -i
    ISortHelper_Compare = i
End Function

Private Sub ISortHelper_Swap(v1 As Variant, v2 As Variant)
    Dim vT As Variant
    vT = v1
    v1 = v2
    v2 = vT
End Sub

Private Sub ISortHelper_CollectionSwap(n As Collection, _
                                       i1 As Variant, _
                                       i2 As Variant, _
                                       Optional key1 As Variant, _
                                       Optional key2 As Variant)
                                       
    ' Be sure both keys are used or neither key is used
    If IsMissing(key1) Xor IsMissing(key2) Then
        ErrRaise eeMissingKey
    End If
    
    Dim v1 As Variant, v2 As Variant, vT As Variant
    If IsMissing(key1) Then     ' Swap without keys
        v1 = n(i1)
        n.Add n(i2), , , i1
        n.Remove i1
        n.Add v1, , , i2
        n.Remove i2
    Else                        ' Swap with keys
        v1 = n(i1)
        v2 = n(i2)
        n.Add vT, , , i1        ' Add placeholder after i1
        n.Remove i1
        n.Add vT, , , i2        ' Add placeholder after i2
        n.Remove i2
        n.Add v2, key2, , i1
        n.Remove i1             ' Remove first placeholder
        n.Add v1, key1, , i2
        n.Remove i2             ' Remove second placeholder
   End If
    
End Sub

Property Get SortMode() As Integer
    SortMode = esmMode
End Property

Property Let SortMode(esmModeA As Integer)
    Select Case esmModeA
    Case esmSortVal, esmSortText, esmSortBin, esmSortLen
        esmMode = esmModeA
    Case Else
        esmMode = esmSortVal
    End Select
End Property

Property Get HiToLo() As Boolean
    HiToLo = fHiToLo
End Property

Property Let HiToLo(fHiToLoA As Boolean)
    fHiToLo = fHiToLoA
End Property

#If fComponent = 0 Then
Private Sub ErrRaise(e As Long)
    Dim sText As String, sSource As String
    If e > 1000 Then
        sSource = App.ExeName & ".SortHelper"
        Select Case e
        Case eeBaseSortHelper
            BugAssert True
        Case eeMissingKey
            sText = "CollectionSwap: Key is missing"
        End Select
        Err.Raise COMError(e), sSource, sText
    Else
        ' Raise standard Visual Basic error
        sSource = App.ExeName & ".VBError"
        Err.Raise e, sSource
    End If
End Sub
#End If

