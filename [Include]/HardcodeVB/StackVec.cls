VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CStackVec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IStack

Private av() As Variant
Private Const cChunk = 10
Private iLast As Long, iCur As Long

Private Sub IStack_Push(vArg As Variant)
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

Private Function IStack_Pop() As Variant
    If iCur Then
        If IsObject(av(iCur)) Then
            Set IStack_Pop = av(iCur)
        Else
            IStack_Pop = av(iCur)
        End If
        iCur = iCur - 1
        If iCur < (iLast - cChunk) Then
            iLast = iLast - cChunk      ' Shrink
            ReDim Preserve av(1 To iLast) As Variant
        End If
    End If
End Function

Private Property Get IStack_Count() As Long
    IStack_Count = iCur
End Property
'
