VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "THash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
#Const VariantType = 0
#Const ObjectType = 1
#Const StringType = 2
#Const IntegerType = 3
#Const LongType = 4
#Const DoubleType = 5
#Const DefaultType = 0

#Const TypeOfTPLATYPE = DefaultType
#Const TypeOfTPLBTYPE = DefaultType

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
      (ByVal pDst As Long, ByVal pSrc As Long, ByVal ByteLen As Long)
      
Private Type ValueArray
    Count As Long
    Data() As TPLBType
End Type

Private Const CONST_TABLE_SIZE = 1024
Private mTable() As ValueArray
Private mSize As Long

Public Sub Add(ByRef key As TPLAType, ByRef value As TPLBType)
    Dim iKey As Long
    ComputeIndex key, iKey
    AddValueTo mTable(iKey).Data, mTable(iKey).Count, value
End Sub


Private Sub ComputeIndex(ByRef key As TPLAType, ByRef iKey As Long)

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
    
    iKey = 0
    bKey = StrConv(strKey, vbFromUnicode)
    Dim i As Long
    For i = LBound(bKey) To UBound(bKey)
        iKey = iKey + bKey(i)
    Next
    iKey = iKey Mod mSize + 1

End Sub

Private Sub AddValueTo(ByRef aValue() As TPLBType, ByRef cValue As Long, ByRef value As TPLBType)
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



