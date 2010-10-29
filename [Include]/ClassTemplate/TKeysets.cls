VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TKeysets"
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
    Data() As TPLBType
End Type

Private Const CONST_TABLE_SIZE = 1024
Private mtable() As ValueArray
Private mSize As Long

Public Sub Add(ByRef key As TPLAType, ByRef value As TPLBType)
    Dim ikey As Long
    ComputeIndex key, ikey
    AddValueTo mtable(ikey).Data, mtable(ikey).Count, value
End Sub

Public Function GetValueArray(ByRef key As TPLAType) As TKeysets.ValueArray
    Dim ikey As Long
    ComputeIndex key, ikey
    GetValueArray = mtable(ikey)
End Function

Public Function IsExists(ByRef key As TPLAType) As Boolean
    Dim ikey As Long
    ComputeIndex key, ikey
    If (mtable(ikey).Count > 0) Then IsExists = True Else IsExists = False
End Function

Public Function GetValue(ByRef key As TPLAType, Index As Long) As TPLBType
    Dim ikey As Long
    ComputeIndex key, ikey
    If Index < 0 Then Exit Function
    If Index > mtable(ikey).Count Then Exit Function
    MTemplateHelper.Assign TPLBType, GetValue, mtable(ikey).Data(Index)
End Function

Public Function ClearKey(ByRef key As TPLAType) As Boolean
    Dim ikey As Long
    ComputeIndex key, ikey
    mtable(ikey).Count = 0
End Function

Public Sub Reset()
    
End Sub
Private Sub ComputeIndex(ByRef key As TPLAType, ByRef ikey As Long)

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
    #ElseIf TypeOfTPLATYPE = NormalType Then
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

Private Sub AddValueTo(ByRef aValue() As TPLBType, ByRef cValue As Long, ByRef value As TPLBType)
    ReDim aValue(0 To cValue)
    MTemplateHelper.Assign TPLBType, aValue(cValue), value
End Sub





Private Sub Class_Initialize()
    mSize = CONST_TABLE_SIZE
    ReDim mtable(0 To mSize) As TKeysets.ValueArray
End Sub
