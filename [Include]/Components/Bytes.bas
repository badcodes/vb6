Attribute VB_Name = "MBytes"
Option Explicit

Public Function Bytes_Merge(ByRef vFirst() As Byte, ByRef vSecond() As Byte) As Byte()
    Dim first As String
    first = vFirst()
    Dim second As String
    second = vSecond
    Dim result As String
    result = first & second
    Bytes_Merge = result
End Function

Public Function Bytes_Extract(ByRef vData() As Byte, Optional ByVal vStart As Long = 0, Optional vLength As Long = -1) As Byte()
    '<EhHeader>
    On Error GoTo Bytes_Extract_Err
    '</EhHeader>
    Dim result As String
    Dim data As String
    data = vData()
    If vStart < 0 Then vStart = 0
    If vLength < 1 Then
        result = MidB$(data, vStart + 1)
    Else
    result = MidB$(data, vStart + 1, vLength)
    End If
    Bytes_Extract = result
    '<EhFooter>
    Exit Function

Bytes_Extract_Err:
    Debug.Print "GetSSLib.MBytes.Bytes_Extract:Error " & Err.Description
    Err.Clear

    '</EhFooter>
End Function


