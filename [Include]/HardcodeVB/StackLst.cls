VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CStackLst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IStack

Private lnkHead As CLink
Private c As Long

Private Sub IStack_Push(vArg As Variant)
    ' Create a temporary link with new value
    Dim lnkTmp As New CLink
    Set lnkTmp = New CLink
    If IsObject(vArg) Then
        Set lnkTmp.Item = vArg
    Else
        lnkTmp.Item = vArg
    End If
    ' Insert it at head
    Set lnkTmp.NextLink = lnkHead
    Set lnkHead = lnkTmp
    c = c + 1
End Sub

Private Function IStack_Pop() As Variant
    ' Can't pop from empty list
    If lnkHead Is Nothing Then Exit Function
    ' Copy head value and remove the link
    If IsObject(lnkHead.Item) Then
        Set IStack_Pop = lnkHead.Item
    Else
        IStack_Pop = lnkHead.Item
    End If
    Set lnkHead = lnkHead.NextLink
    c = c - 1
End Function

Private Property Get IStack_Count() As Long
    IStack_Count = c
End Property

