VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CRegItemWalker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum EErrorRegItemWalker
    eeBaseRegItemWalker = 13170 ' CRegItemWalker
End Enum

' Implement Basic-friendly version of IEnumVARIANT
Implements IVariantWalker
' Connect back to parent collection
Private connect As CRegItem

' Private state data
Private iCur As Long

Private Sub Class_Initialize()
    ' Initialize position in collection
    iCur = -1
End Sub

' Receive connection from CRegItem
Friend Sub Attach(connectA As CRegItem)
    Set connect = connectA
End Sub

' Return IEnumVARIANT (indirectly) to client collection
Friend Function NewEnum() As stdole.IEnumVARIANT
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
    ' We can't fail in a walker, and yet the registry will return errors
    On Error Resume Next
    iCur = iCur + 1
    Do While iCur < connect.Count
        ' Get next node from registry
        Set v = connect.RegItems(iCur)
        Select Case Err.Number
        Case 0
            ' If more data, return True
            IVariantWalker_More = True
            Exit Function
        Case ApiToCom(ERROR_ACCESS_DENIED)
            ' WinNT returns access error if you don't have permission
            Err.Clear
            iCur = iCur + 1
        Case ApiToCom(ERROR_NO_MORE_ITEMS)
            ' Sometimes happens under WinNT
            iCur = connect.Count
            Err.Clear
            Exit Function
        Case Else
            ' Failure terminates loop, so do only for catastrophies
            Err.Raise Err.Number, Err.Source, Err.Description
        End Select
    Loop
#If 0 Then
    ' Move to next element
    iCur = iCur + 1
    ' If more data, return True and update data
    If iCur < connect.Count Then
        IVariantWalker_More = True
        Set v = connect.RegItems(iCur)
    End If
#End If
End Function

Private Sub IVariantWalker_Reset()
    ' Move to first element
    iCur = -1
End Sub

Private Sub IVariantWalker_Skip(c As Long)
    ' Skip a given number of elements
    iCur = iCur + c
End Sub

#If fComponent = 0 Then
Private Sub ErrRaise(e As Long)
    Dim sText As String, sSource As String
    If e > 1000 Then
        sSource = App.ExeName & ".RegItemWalker"
        Select Case e
        Case eeBaseRegItemWalker
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

