Attribute VB_Name = "OHash"
Option Explicit

' hash table algorithm, object method

Private hashTableSize As Long       ' size of hashTable
Private hashTable() As COHash        ' hashTable(0..hashTableSize-1)

Public Function Hash(ByVal KeyVal As Variant) As Long
'   inputs:
'       KeyVal                key
'   returns:
'       hashed value of key
'   action:
'       Compute hash value based on KeyVal.
'
    Hash = KeyVal Mod hashTableSize
End Function

Public Sub Insert(ByVal KeyVal As Variant, ByRef RecVal As Variant)
'   inputs:
'       KeyVal                key of node to insert
'       RecVal                record associated with key
'   action:
'       Inserts record RecVal with key KeyVal.
'
    Dim p As COHash
    Dim p0 As COHash
    Dim bucket As Long

    ' allocate node and insert in table

    ' insert node at beginning of list
    bucket = Hash(KeyVal)
    Set p = New COHash
    Set p0 = hashTable(bucket)
    Set hashTable(bucket) = p
    Set p.Nxt = p0
    p.Key = KeyVal
    p.Rec = RecVal
End Sub


Public Sub Delete(ByVal KeyVal As Variant)
'   inputs:
'       KeyVal                key of node to delete
'   action:
'       Deletes record with key KeyVal.
'   error:
'       errKeyNotFound
'
    Dim p0 As COHash
    Dim p As COHash
    Dim bucket As Long

   ' delete node containing key from table

    ' find node
    Set p0 = Nothing
    bucket = Hash(KeyVal)
    Set p = hashTable(bucket)
    Do While Not p Is Nothing
        If p.Key = KeyVal Then Exit Do
        Set p0 = p
        Set p = p.Nxt
    Loop
    If p Is Nothing Then Raise errKeyNotFound, "Hash.Delete"

    ' p designates node to delete, remove it from list
    If Not p0 Is Nothing Then
        ' not first node, p0 points to previous node
        Set p0.Nxt = p.Nxt
    Else
        ' first node on chain
        Set hashTable(bucket) = p.Nxt
    End If

    ' p will be automatically freed on return, as it's no longer referenced

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
    Dim p As COHash

    '  find node containing key

    Set p = hashTable(Hash(KeyVal))
    Do While Not p Is Nothing
        If p.Key = KeyVal Then Exit Do
        Set p = p.Nxt
    Loop

    If p Is Nothing Then Raise errKeyNotFound, "Hash.Find"

    ' copy data fields to user
    Find = p.Rec

End Function

Public Sub Init(ByVal tableSize As Long)
'   inputs:
'       tableSize               size of hashtable
'   action:
'       initialize hash table
'
    hashTableSize = tableSize
    ReDim hashTable(0 To tableSize - 1)
End Sub

Public Sub Term()
'   action
'       terminate hash table
'       chained nodes are deleted automatically,
'       as they're no longer referenced
'
    Erase hashTable

End Sub
