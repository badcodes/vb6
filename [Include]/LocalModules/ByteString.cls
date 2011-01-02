VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CByteString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mData() As Byte

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

Public Property Get StringType() As String
Attribute StringType.VB_UserMemId = 0
    StringType = mData()
End Property

Public Property Let StringType(ByVal StrValue As String)
    mData() = StrValue
End Property

Public Property Get ByteType() As Byte()
    ByteType = mData
End Property

Public Property Let ByteType(ByRef vData() As Byte)
    mData = vData
End Property

Public Sub AddBytes(ByRef vData() As Byte)
    mData = Bytes_Merge(mData, vData)
End Sub

Public Sub AddString(ByRef vString As String)
    Dim data() As Byte
    data() = vString
    AddBytes data()
End Sub

Public Function ExtractBytes(Optional vStart As Long = 0, Optional vLength As Long = -1) As Byte()
    ExtractBytes = Bytes_Extract(mData(), vStart, vLength)
End Function

Public Function ExtractString(Optional vStart As Long = 1, Optional vLength As Long = -1) As String
    vStart = vStart - 1
    ExtractString = ExtractBytes(vStart, vLength)
End Function

