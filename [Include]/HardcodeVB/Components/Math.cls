VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GMath"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum EErrorMath
    eeBaseMath = 13520      ' Math
End Enum

' Derived math functions from language reference Appendix D

' Secant
Function Sec(x As Double) As Double
    Sec = 1 / Cos(x)
End Function

' Cosecant
Function CoSec(x As Double) As Double
    CoSec = 1 / Sin(x)
End Function

' Cotangent
Function CoTan(x As Double) As Double
    CoTan = 1 / Tan(x)
End Function

' Inverse Sine
Function ArcSin(x As Double) As Double
    ArcSin = Atn(x / Sqr(-x * x + 1))
End Function

' Inverse Cosine
Function ArcCos(x As Double) As Double
    ArcCos = Atn(-x / Sqr(-x * x + 1)) + 2 * Atn(1)
End Function

' Inverse Secant
Function ArcSec(x As Double) As Double
    ArcSec = Atn(x / Sqr(x * x - 1)) + Sgn(x - 1) * (2 * Atn(1))
End Function

' Inverse Cosecant
Function ArcCoSec(x As Double) As Double
    ArcCoSec = Atn(x / Sqr(x * x - 1)) + (Sgn(x) - 1) * (2 * Atn(1))
End Function

' Inverse Cotangent
Function ArcCoTan(x As Double) As Double
    ArcCoTan = Atn(x) + 2 * Atn(1)
End Function

' Hyperbolic Sine
Function HSin(x As Double) As Double
    HSin = (Exp(x) - Exp(-x)) / 2
End Function

' Hyperbolic Cosine
Function HCos(x As Double) As Double
    HCos = (Exp(x) + Exp(-x)) / 2
End Function

' Hyperbolic Tangent
Function HTan(x As Double) As Double
    HTan = (Exp(x) - Exp(-x)) / (Exp(x) + Exp(-x))
End Function

' Hyperbolic Secant
Function HSec(x As Double) As Double
    HSec = 2 / (Exp(x) + Exp(-x))
End Function

' Hyperbolic Cosecant
Function HCoSec(x As Double) As Double
    HCoSec = 2 / (Exp(x) - Exp(-x))
End Function

' Hyperbolic Cotangent
Function HCotan(x As Double) As Double
    HCotan = (Exp(x) + Exp(-x)) / (Exp(x) - Exp(-x))
End Function

' Inverse Hyperbolic Sine
Function HArcSin(x As Double) As Double
    HArcSin = Log(x + Sqr(x * x + 1))
End Function

' Inverse Hyperbolic Cosine
Function HArcCos(x As Double) As Double
    HArcCos = Log(x + Sqr(x * x - 1))
End Function

' Inverse Hyperbolic Tangent
Function HArcTan(x As Double) As Double
    HArcTan = Log((1 + x) / (1 - x)) / 2
End Function

' Inverse Hyperbolic Secant
Function HArcSec(x As Double) As Double
    HArcSec = Log((Sqr(-x * x + 1) + 1) / x)
End Function

' Inverse Hyperbolic Cosecant
Function HArcCoSec(x As Double) As Double
    HArcCoSec = Log((Sgn(x) * Sqr(x * x + 1) + 1) / x)
End Function

' Inverse Hyperbolic Cotangent
Function HArcCoTan(x As Double) As Double
    HArcCoTan = Log((x + 1) / (x - 1)) / 2
End Function

' Logarithm to base N
Function LogN(x As Double, n As Double) As Double
    LogN = Log(x) / Log(n)
End Function

#If fComponent = 0 Then
Private Sub ErrRaise(e As Long)
    Dim sText As String, sSource As String
    If e > 1000 Then
        sSource = App.ExeName & ".Math"
        Select Case e
        Case eeBaseMath
            BugAssert True
       ' Case ee...
       '     Add additional errors
        End Select
        Err.Raise COMError(e), sSource, sText
    Else
        ' Raise standard Visual Basic error
        sSource = App.ExeName & ".VBError"
        Err.Raise e, sSource
    End If
End Sub
#End If

