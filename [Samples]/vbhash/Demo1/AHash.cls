VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CAHash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' hash table, array Method

Private Node As CNode                  ' class for allocating nodes

Private NextNode() As Long             ' next node
Private key() As Variant               ' keys
Private rec() As Variant               ' record

Private HashTableSize As Long          ' Size of HashTable
Private HashTable() As Long            ' HashTable(0..HashTableSize-1)
Private GrowthFactor As Single                ' growth factor

Private Function Hash(ByVal KeyVal As Variant) As Long
'   inputs:
'       KeyVal                key
'   returns:
'       hashed value of key
'   action:
'       Compute hash value based on KeyVal.
'
    Hash = KeyVal Mod HashTableSize
End Function

Public Sub Insert(ByVal KeyVal As Variant, ByRef RecVal As Variant)
'   inputs:
'       KeyVal                key of node to insert
'       RecVal                record associated with key
'   action:
'       Inserts record RecVal with key KeyVal.
'
    Dim p As Long
    Dim p0 As Long
    Dim bucket As Long

    ' allocate node and insert in table

    ' insert node at beginning of list
    bucket = Hash(KeyVal)
    p = Node.Alloc()
    If p > UBound(key) Then
        ReDim Preserve NextNode(1 To UBound(NextNode) * GrowthFactor)
        ReDim Preserve key(1 To UBound(key) * GrowthFactor)
        ReDim Preserve rec(1 To UBound(rec) * GrowthFactor)
    End If
    p0 = HashTable(bucket)
    HashTable(bucket) = p
    NextNode(p) = p0
    key(p) = KeyVal
    rec(p) = RecVal
End Sub


Public Sub Delete(ByVal KeyVal As Variant)
'   inputs:
'       KeyVal                key of node to delete
'   action:
'       Deletes record with key KeyVal.
'   error:
'       errKeyNotFound
'
    Dim p0 As Long
    Dim p As Long
    Dim bucket As Long

   ' delete node containing key from table

    ' find node
    p0 = 0
    bucket = Hash(KeyVal)
    p = HashTable(bucket)
    Do While p <> 0
        If key(p) = KeyVal Then Exit Do
        p0 = p
        p = NextNode(p)
    Loop
    If p = 0 Then Raise errKeyNotFound, "CHash.Delete"

    ' p designates node to delete, remove it from list
    If p0 <> 0 Then
        ' not first node, p0 points to previous node
        NextNode(p0) = NextNode(p)
    Else
        ' first node on chain
        HashTable(bucket) = NextNode(p)
    End If

    Node.Free (p)
    Set rec(p) = Nothing
End Sub

Public Function Find(ByVal KeyVal As Variant) As Variant
'   inputs:
'       KeyVal                key of node to delete
'   returns:
'       record associated with key
'   action:
'       Finds record with key KeyVal
'   error:
'       errKeyNotFound
'
    Dim p As Long

    '  find node containing key

    p = HashTable(Hash(KeyVal))
    Do While p <> 0
        If key(p) = KeyVal Then Exit Do
        p = NextNode(p)
    Loop

    If p = 0 Then Raise errKeyNotFound, "CHash.Find"

    Find = rec(p)
End Function

Public Sub Init( _
        ByVal TableSizeVal As Long, _
        ByVal InitialAllocVal As Long, _
        ByVal GrowthFactorVal As Single)

    ' save values
    HashTableSize = TableSizeVal
    GrowthFactor = GrowthFactorVal

    ' initialize hash table
    ReDim HashTable(0 To TableSizeVal - 1)

    ' initialize nodes
    ReDim key(1 To InitialAllocVal)
    ReDim NextNode(1 To InitialAllocVal)
    ReDim rec(1 To InitialAllocVal)
    Set Node = New CNode
    Node.Init InitialAllocVal, GrowthFactorVal
End Sub

Public Sub Class_Terminate()

    ' terminate hash table
    ' chained nodes are deleted automatically,
    ' as they're no longer referenced
    Set Node = Nothing
End Sub
