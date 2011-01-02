VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CKeysetsStrInt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Template header:
'================================================================================
#Const ObjectType = 1
#Const NormalType = 2
#Const VariantType = (ObjectType Or NormalType)
#Const DefaultType = VariantType

#Const fTPLAType = DefaultType
#Const fTPLBType = DefaultType
'================================================================================


Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
      (ByVal pDst As Long, ByVal pSrc As Long, ByVal ByteLen As Long)
      
Public Type ValueArray
    Count As Long
    Data() As Integer 
End Type

Private Const CONST_TABLE_SIZE = 1024
Private mTable() As ValueArray
Private mSize As Long

Public Sub Add(ByRef key As String, ByRef value As Integer) 
    Dim ikey As Long
    ComputeIndex key, ikey
    AddValueTo mTable(ikey).Data, mTable(ikey).Count, value
End Sub

Public Function GetValueArray(ByRef key As String) As TKeysets.ValueArray 
    Dim ikey As Long
    ComputeIndex key, ikey
    GetValueArray = mTable(ikey)
End Function

Public Function IsExists(ByRef key As String) As Boolean 
    Dim ikey As Long
    ComputeIndex key, ikey
    If (mTable(ikey).Count > 0) Then IsExists = True Else IsExists = False
End Function

Public Function GetValue(ByRef key As String, index As Long) As Integer 
    Dim ikey As Long
    ComputeIndex key, ikey
    If index < 0 Then Exit Function
    If index > mTable(ikey).Count Then Exit Function
    If IsObject(GetValue)  then GetValue = mTable(ikey).Data(index) else Set GetValue = mTable(ikey).Data(index) 
End Function
Private Sub ComputeIndex(ByRef key As String, ByRef ikey As Long) 

    Dim strKey As String
    Dim bKey() As Byte
    
    strKey = "0"
    
    On Error Resume Next
    
    #If TypeOfTPLATYPE = VariantType Then
        If IsObject(key) Then
            strKey = key.ToString
        Else
            strKey = CStr(key)
        End If
    #ElseIf TypeOfTPLATYPE = ObjectType Then
        strKey = key.ToString
    #ElseIf TypeOfTPLATYPE = StringType Then
        strKey = key
    #Else
        strKey = CStr(key)
    #End If
    
    ikey = 0
    bKey = StrConv(strKey, vbFromUnicode)
    Dim i As Long
    For i = LBound(bKey) To UBound(bKey)
        ikey = ikey + bKey(i)
    Next
    ikey = ikey Mod mSize + 1

End Sub

Private Sub AddValueTo(ByRef aValue() As Integer, ByRef cValue As Long, ByRef value As Integer) 
    ReDim aValue(0 To cValue)
    #If TypeOfTPLBTYPE = VariantType Then
        If IsObject(value) Then
            Set aValue(cValue) = value
        Else
            aValue(cValue) = value
        End If
    #ElseIf TypeOfTPLBTYPE = ObjectType Then
        Set aValue(cValue) = value
    #Else
        aValue(cValue) = value
    #End If
End Sub





