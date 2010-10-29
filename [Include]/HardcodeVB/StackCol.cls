VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CStackCol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' CStackCol implements IStack
Implements IStack
Private stack As Collection

Private Sub Class_Initialize()
    Set stack = New Collection
End Sub

Private Sub IStack_Push(vArg As Variant)
    stack.Add vArg
End Sub

Private Function IStack_Pop() As Variant
    If stack.Count Then
        If IsObject(stack(stack.Count)) Then
            Set IStack_Pop = stack(stack.Count)
        Else
            IStack_Pop = stack(stack.Count)
        End If
        stack.Remove stack.Count
    End If
End Function

Private Property Get IStack_Count() As Long
    IStack_Count = stack.Count
End Property
'
