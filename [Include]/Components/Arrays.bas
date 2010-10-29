Attribute VB_Name = "MArrays"
Option Explicit
Public Const CST_ARRAYS_INVALID_UBOUND As Long = -32768
Public Const CST_ARRAYS_INVALID_LBOUND As Long = CST_ARRAYS_INVALID_UBOUND + 11

Public Function SafeUBound(mArray As Variant, Optional w As Long = 1) As Long
    On Error GoTo ErrorSafeUbound
    SafeUBound = UBound(mArray, w)
    Exit Function

ErrorSafeUbound:
    SafeUBound = CST_ARRAYS_INVALID_UBOUND
End Function

Public Function SafeLBound(mArray As Variant, Optional w As Long = 1) As Long
    On Error GoTo ErrorSafeUbound
    SafeUBound = LBound(mArray, w)
    Exit Function

ErrorSafeUbound:
    SafeUBound = CST_ARRAYS_INVALID_LBOUND
End Function

Public Function ArrayIsEmpty(mArray As Variant) As Boolean
    On Error GoTo ErrorArrayIsEmpty
    If SafeLBound(mArray) <> CST_ARRAYS_INVALID_LBOUND Then
        ArrayIsEmpty = True
        Exit Sub
    End If
ErrorArrayIsEmpty:
    ArrayIsEmpty = False
End Function
