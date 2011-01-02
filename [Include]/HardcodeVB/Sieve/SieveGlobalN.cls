VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GSieveGlobalN"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Function SieveGlobalN(ai() As Integer) As Integer
    Dim iLast As Integer, cPrime As Integer, iCur As Integer, i As Integer
    Dim af() As Boolean
    ' Parameter should have dynamic array for maximum number of primes
    If LBound(ai) <> 0 Then Exit Function
    iLast = UBound(ai)
    ' Create array large enough for maximum prime (initializing to zero)
    ReDim af(0 To iLast + 1) As Boolean
    For iCur = 2 To iLast
        ' Anything still zero is a prime
        If Not af(iCur) Then
            ' Cancel its multiples because they can't be prime
            For i = iCur + iCur To iLast Step iCur
                af(i) = True
            Next
            ' Count this prime
            ai(cPrime) = iCur
            cPrime = cPrime + 1
        End If
    Next
    ' Resize array to the number of primes found
    ReDim Preserve ai(0 To cPrime) As Integer
    SieveGlobalN = cPrime
End Function


