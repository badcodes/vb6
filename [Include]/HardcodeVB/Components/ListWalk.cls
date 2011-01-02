VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CListWalker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum EErrorListWalker
    eeBaseListWalker = 13110    ' CListWalker
End Enum

' Implement Basic-friendly version of IEnumVARIANT
Implements IVariantWalker
' Connect back to parent collection
Private connect As CList
' Current link
Private lnkCur As CLink

Private Sub Class_Initialize()
    BugLocalMessage "CListWalker:Initialize"
End Sub

Private Sub Class_Terminate()
    BugLocalMessage "CListWalker:Terminate"
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

' Attach a list to iterator
Sub Attach(connectA As CList, Optional fEnumerate As Boolean = False)
    ' Initialize position in collection
    Set connect = connectA
    ' The fEnumerate parameter was a mistake, but can't remove it
    ' without breaking compatibility
End Sub

' Report whether there are more links to iterate
Function More() As Boolean
    If lnkCur Is Nothing Then
        ' Don't skip the first time through
        Set lnkCur = connect.Head
    Else
        ' Skip to the next item
        Set lnkCur = lnkCur.NextLink
    End If
    ' When the next link is nothing, we're done (handles empty list)
    If Not lnkCur Is Nothing Then More = True
End Function

' Default member
Property Get item() As Variant
Attribute item.VB_UserMemId = 0
    If IsObject(lnkCur.item) Then
        Set item = lnkCur.item
    Else
        item = lnkCur.item
    End If
End Property

' Expose current link to friends
Friend Property Get CurLink() As CLink
    Set CurLink = lnkCur
End Property
Friend Property Set CurLink(lnkCurA As CLink)
    Set lnkCur = lnkCurA
End Property

' Implement IVariantWalker methods
Private Function IVariantWalker_More(v As Variant) As Boolean
    ' Move to next element
    IVariantWalker_More = More
    If IVariantWalker_More = False Then Exit Function
    ' Return element through reference
    If IsObject(lnkCur.item) Then
        Set v = lnkCur.item
    Else
        v = lnkCur.item
    End If
End Function

Private Sub IVariantWalker_Reset()
    ' Move to first element
    If connect.Count Then Set lnkCur = connect.Head
End Sub

Private Sub IVariantWalker_Skip(c As Long)
    ' Skip a given number of elements
    Dim i As Long, v As Variant
    For i = 1 To c
        If IVariantWalker_More(v) = False Then Exit For
    Next
End Sub
'

#If fComponent = 0 Then
Private Sub ErrRaise(e As Long)
    Dim sText As String, sSource As String
    If e > 1000 Then
        sSource = App.ExeName & ".ListWalker"
        Select Case e
        Case eeBaseListWalker
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

