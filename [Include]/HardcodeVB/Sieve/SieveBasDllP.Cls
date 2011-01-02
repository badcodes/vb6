VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSieveBasDllP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Sieve of Eratosthenes"
Option Explicit

Private af() As Boolean, iCur As Integer
Private iMaxPrime As Integer, cPrime As Integer

Private Sub Class_Initialize()
    ' Default size is largest integer
    iMaxPrime = 32766
    ReInitialize
End Sub

Sub ReInitialize()
    ReDim af(0 To iMaxPrime)
    iCur = 1: cPrime = 0
End Sub

Property Get NextPrime() As Integer
    NextPrime = 0
    ' Loop until we find a prime or overflow array
    iCur = iCur + 1
    On Error GoTo OverMaxPrime
    Do While af(iCur)
        iCur = iCur + 1
    Loop
    ' Cancel multiples of this prime
    Dim i As Long
    For i = iCur + iCur To iMaxPrime Step iCur
        af(i) = True
    Next
    ' Count and return it
    cPrime = cPrime + 1
    NextPrime = iCur
OverMaxPrime:       ' Array overflow comes here
End Property

Property Get MaxPrime() As Integer
    MaxPrime = iMaxPrime
End Property

Property Let MaxPrime(iMaxPrimeA As Integer)
    iMaxPrime = iMaxPrimeA
    ReInitialize
End Property

Property Get Primes() As Integer
    Primes = cPrime
End Property

Sub AllPrimes(ai() As Integer)
    If LBound(ai) <> 0 Then Exit Sub
    iMaxPrime = UBound(ai)
    cPrime = 0
    Dim i As Integer
    For iCur = 2 To iMaxPrime
        If Not af(iCur) Then    ' Found a prime
            For i = iCur + iCur To iMaxPrime Step iCur
                af(i) = True    ' Cancel its multiples
            Next
            ai(cPrime) = iCur
            cPrime = cPrime + 1
        End If
    Next
    ReDim Preserve ai(0 To cPrime) As Integer
    iCur = 1
End Sub
