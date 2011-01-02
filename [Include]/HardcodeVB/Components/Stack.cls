VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CStack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum EErrorStack
    eeBaseStack = 13230     ' CStack
End Enum

Private av() As Variant
Private Const cChunk = 10
Private iLast As Long, iCur As Long

Sub Push(vArg As Variant)
    iCur = iCur + 1
    On Error GoTo FailPush
    If IsObject(vArg) Then
        Set av(iCur) = vArg
    Else
        av(iCur) = vArg
    End If
    Exit Sub
FailPush:
    iLast = iLast + cChunk  ' Grow
    ReDim Preserve av(1 To iLast) As Variant
    Resume                  ' Try again
End Sub

Function Pop() As Variant
    If iCur Then
        If IsObject(av(iCur)) Then
            Set Pop = av(iCur)
        Else
            Pop = av(iCur)
        End If
        iCur = iCur - 1
        If iCur < (iLast - cChunk) Then
            iLast = iLast - cChunk      ' Shrink
            ReDim Preserve av(1 To iLast) As Variant
        End If
    End If
End Function

Property Get Count() As Long
    Count = iCur
End Property
'

#If fComponent = 0 Then
Private Sub ErrRaise(e As Long)
    Dim sText As String, sSource As String
    If e > 1000 Then
        sSource = App.ExeName & ".Stack"
        Select Case e
        Case eeBaseStack
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

