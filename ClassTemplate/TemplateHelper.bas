Attribute VB_Name = "MTemplateHelper"
Option Explicit

Public Sub Assign(vTypeName, vLeft, vRight)
    If IsObject(vLeft) Then
            Set vLeft = vRight
    Else
            vLeft = vRight
    End If
End Sub
    
Public Function Compare(vTypeName, vLeft, vRight) As Boolean
        If IsObject(vLeft) Then
            If vLeft Is vRight Then Compare = True
        Else
            If vLeft = vRight Then Compare = True
        End If
End Function

Public Sub Delete(vTypeName, vWhat)
    If IsObject(vWhat) Then
        Set vWhat = Nothing
    Else
    End If
End Sub

Public Sub DeleteArray(vTypeName, vArray, vStart, vCount)
    Dim i As Long
    For i = vStart To vStart + vCount - 1
        Delete vTypeName, vArray(i)
    Next
End Sub
