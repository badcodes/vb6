Attribute VB_Name = "mShareFunction"
Option Explicit
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

Public Function CBytesToStr(ByRef CBytes() As Byte, Optional localID As Long = 0) As String

    Dim lUB As Long, lLb As Long
    Dim iPos As Long
    Dim bTemp() As Byte
    Dim l As Long
    
    lUB = UBound(CBytes)
    lLb = LBound(CBytes)
    
    For iPos = lLb To lUB
        If CBytes(iPos) = 0 Then Exit For
    Next
    
    If iPos = 0 Then
        CBytesToStr = StrConv(CBytes, vbUnicode, localID)
    ElseIf iPos = lLb Then
        CBytesToStr = ""
    Else
        ReDim bTemp(lLb To iPos - 1)
        CopyMemory bTemp(lLb), CBytes(lLb), iPos - lLb
        CBytesToStr = StrConv(bTemp, vbUnicode, localID)
    End If
    Debug.Print CBytesToStr
    
End Function

Public Sub StrToCBytes(ByVal strUnicode As String, ByRef CBytes() As Byte, Optional localID As Long = 0)

    Dim lUB As Long, lLb As Long
    Dim bTemp() As Byte
    Dim lSize As Long
    
    lUB = UBound(CBytes)
    lLb = LBound(CBytes)
    
    bTemp = StrConv(strUnicode, vbFromUnicode, localID)
    
    lSize = UBound(bTemp) + 1
    ReDim Preserve bTemp(lSize)
    bTemp(lSize) = 0
    
    If lSize > lUB - lLb Then
        lSize = lUB - lLb
        bTemp(lSize) = 0
    End If
    
    CopyMemory CBytes(lLb), bTemp(0), lSize + 1
       
End Sub

