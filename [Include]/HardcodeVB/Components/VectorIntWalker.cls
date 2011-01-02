VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CVectorIntWalker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Implement Basic-friendly version of IEnumVARIANT
Implements IVariantWalker
' Connect back to parent collection
Private connect As CVectorInt

Public Enum EErrorVectorIntWalker
    eeBaseVectorIntWalker = 13330   ' CVectorIntWalker
End Enum

' Private state data
Private iCur As Long

Private Sub Class_Initialize()
    ' Initialize position in collection
    iCur = 0
End Sub

' Receive connection from CVectorInt
Friend Sub Attach(connectA As CVectorInt)
    Set connect = connectA
End Sub

' Return IEnumVARIANT (indirectly) to client collection
Function NewEnum() As stdole.IEnumVARIANT
    ' Delegate to class that implements real IEnumVARIANT
    Dim vars As CEnumVariant
    ' Connect walker to CEnumVariant so it can call methods
    Set vars = New CEnumVariant
    vars.Attach Me
    ' Return walker to collection data
    Set NewEnum = vars
End Function

' Implement IVariantWalker methods
Private Function IVariantWalker_More(v As Variant) As Boolean
    ' Move to next element
    iCur = iCur + 1
    ' Return False if no more data
    If iCur > connect.Last Then Exit Function
    ' Return element through reference
    v = connect.Vector(iCur)
    IVariantWalker_More = True
End Function

Private Sub IVariantWalker_Reset()
    ' Move to first element
    iCur = 0
End Sub

Private Sub IVariantWalker_Skip(c As Long)
    ' Skip a given number of elements
    iCur = iCur + c
End Sub

#If fComponent = 0 Then
Private Sub ErrRaise(e As Long)
    Dim sText As String, sSource As String
    If e > 1000 Then
        sSource = App.ExeName & ".VectorIntWalker"
        Select Case e
        Case eeBaseVectorIntWalker
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

