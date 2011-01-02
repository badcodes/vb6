VERSION 5.00
Begin VB.UserControl XSieveN 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   636
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   648
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   7.8
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   648
   ScaleMode       =   0  'User
   ScaleWidth      =   648
   ToolboxBitmap   =   "SieveN.ctx":0000
   Begin VB.Label lbl 
      Caption         =   "Sieve"
      Height          =   204
      Left            =   36
      TabIndex        =   0
      Top             =   24
      Width           =   540
   End
   Begin VB.Image img 
      Height          =   180
      Left            =   48
      Picture         =   "SieveN.ctx":00FA
      Top             =   252
      Width           =   192
   End
End
Attribute VB_Name = "XSieveN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private af() As Boolean, iCur As Integer
Private iMaxPrime As Integer, cPrime As Integer

Private Sub UserControl_Initialize()
    Debug.Print "UserControl_Initialize"
End Sub

' Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    ' Default size is largest integer
    iMaxPrime = 32766
End Sub

' Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    iMaxPrime = PropBag.ReadProperty("MaxPrime", 32766)
End Sub

' Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("MaxPrime", iMaxPrime, 32766)
End Sub

Private Sub UserControl_Show()
    ReInitialize
    If Ambient.UserMode Then Extender.Visible = False
End Sub

Private Sub UserControl_Resize()
    Width = lbl.Width
    Height = lbl.Width
End Sub

Sub ReInitialize()
    ReDim af(0 To iMaxPrime)
    iCur = 1: cPrime = 0
End Sub

Property Get NextPrime() As Integer
Attribute NextPrime.VB_MemberFlags = "400"
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
    PropertyChanged "MaxPrime"
End Property

Property Get Primes() As Integer
Attribute Primes.VB_MemberFlags = "400"
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

