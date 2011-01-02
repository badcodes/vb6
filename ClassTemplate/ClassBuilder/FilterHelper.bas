Attribute VB_Name = "MFilterHelper"
Option Explicit

Public Const SpaceChar = " "
Public Const TOKEN_SEPARATOR As String = vbTab & SpaceChar

Public Sub AssignString(ByRef sTarget As String, ByRef sNew As String)
    Dim nPos As Integer
    Dim i As Long
    Dim nLen As Long
    nPos = 0
    nLen = Len(sTarget)
    Dim c As String
    For i = 1 To nLen
        c = Mid$(sTarget, i, 1)
        If (c = SpaceChar Or c = vbTab) Then nPos = i Else Exit For
    Next
    sTarget = Left$(sTarget, nPos) & sNew
End Sub

