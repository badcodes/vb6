VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CListItemWalker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Implement Basic-friendly version of IEnumVARIANT
Implements IVariantWalker
' Connect back to parent collection
Private connect As XListBoxPlus

' Private state data
Private iCur As Long

Private Sub Class_Initialize()
    ' Initialize position in collection
    iCur = 0
End Sub

' Receive connection from XListBoxPlus
Friend Sub Attach(connectA As XListBoxPlus)
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
    ' Move to next element
    iCur = iCur + 1
    ' If more data, return True and update data
    If iCur <= connect.Count Then
        IVariantWalker_More = True
        v = connect.ListItems(iCur)
    End If
End Function

Private Sub IVariantWalker_Reset()
    ' Move to first element
    iCur = 0
End Sub

Private Sub IVariantWalker_Skip(c As Long)
    ' Skip a given number of elements
    iCur = iCur + c
End Sub

