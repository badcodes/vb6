VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CStringMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const cst_hash_table_size As Long = 256



Private mHashTable() As CStringLink
Private mHashTableSize As Long

Private Sub Class_Initialize()
    Init cst_hash_table_size
End Sub

Public Sub Init(ByRef HashTableSize As Long)
    If HashTableSize < 0 Then Exit Sub
    ReDim mHashTable(0 To HashTableSize)
    mHashTableSize = HashTableSize
End Sub

Private Function ComputeHash(ByRef vString As String) As Long
    Dim iKey As Long
    Dim bKey() As Byte
    iKey = 0
    bKey = StrConv(vString, vbFromUnicode)
    Dim i As Long
    For i = LBound(bKey) To UBound(bKey)
        iKey = iKey + bKey(i)
    Next
    iKey = iKey Mod mHashTableSize
    ComputeHash = iKey
    
End Function

Private Function SearchForLink(ByRef vKey As String) As CStringLink
    Dim idx As Long
    idx = ComputeHash(vKey)
    Dim currentLink As CStringLink
    Set currentLink = mHashTable(idx)
    Do Until currentLink Is Nothing
        If currentLink.Data = vKey Then
            Exit Do
        Else
            Set currentLink = currentLink.NextLink
        End If
    Loop
    Set SearchForLink = currentLink
End Function

'Private Sub SearchForKey(ByRef vKeyIndex As Long, ByRef KeyLink As CStringLink, ByRef PrevLink As CStringLink)
'
'    Dim pLink As CStringLink
'    Dim cLink As CStringLink
'
'    Set pLink = mHashTable(vKeyIndex)
'    If pLink Is Nothing Then
'        Set KeyLink = Nothing
'        Set PrevLink = Nothing
'    End If
'    Set cLink = pLink.NextLink
'    Do Until cLink Is Nothing
'        If (cLink.Data = vKey) Then Exit Do
'        Set pLink = cLink
'        Set cLink = pLink.NextLink
'    Loop
'    Set PrevLink = pLink
'    Set KeyLink = cLink
'End Sub

Public Property Get Map(ByRef vKey As String) As String
    Dim index As Long
    index = ComputeHash(vKey)
    'If mHashTable(index) Is Nothing Then Exit Property
    Dim cLink As CStringLink
    Set cLink = mHashTable(index)
    Do Until cLink Is Nothing
        If cLink.Data = vKey Then
            Map = GetValue(cLink)
            Exit Property
        Else
            Set cLink = GetNextLink(cLink)
        End If
    Loop
    
End Property

Public Function Exists(ByRef vKey As String) As Boolean
    Dim index As Long
    index = ComputeHash(vKey)
    Dim cLink As CStringLink
    Set cLink = mHashTable(index)
    Do Until cLink Is Nothing
        If cLink.Data = vKey Then
            Exists = True
            Exit Function
        Else
            Set cLink = GetNextLink(cLink)
        End If
    Loop
    
End Function

Public Sub Delete(ByRef vKey As String)
    Dim index As Long
    index = ComputeHash(vKey)
    
    If mHashTable(index) Is Nothing Then Exit Sub
    
    If mHashTable(index).Data = vKey Then
        Set mHashTable(index) = GetNextLink(mHashTable(index))
        Exit Sub
    End If
    
    Dim cLink As CStringLink
    Dim pLink As CStringLink
    
    Set pLink = mHashTable(index)
    Set cLink = GetNextLink(pLink)
    Do Until cLink Is Nothing
        If cLink.Data = vKey Then
            Set cLink = GetNextLink(cLink)
            Exit Sub
        Else
            Set pLink = cLink
            Set cLink = GetNextLink(cLink)
        End If
    Loop



End Sub

Public Property Let Map(ByRef vKey As String, ByRef vValue As String)
    Dim index As Long
    index = ComputeHash(vKey)
    
    If mHashTable(index) Is Nothing Then
        Set mHashTable(index) = New CStringLink
        SetKey mHashTable(index), vKey
        SetValue mHashTable(index), vValue
        Exit Property
    End If
    
    If mHashTable(index).Data = vKey Then
        SetValue mHashTable(index), vValue
        Exit Property
    End If
    
    Dim cLink As CStringLink
    Dim pLink As CStringLink
    
    Set pLink = mHashTable(index)
    Set cLink = GetNextLink(pLink)
    Do Until cLink Is Nothing
        If cLink.Data = vKey Then
            SetValue cLink, vValue
            Exit Property
        Else
            Set pLink = cLink
            Set cLink = GetNextLink(cLink)
        End If
    Loop

    Set cLink = New CStringLink
    SetUpLink cLink, vKey, vValue
    AppendLink pLink, cLink

End Property

Private Sub AppendLink(ByVal pLink As CStringLink, ByVal cLink As CStringLink)
    If pLink Is Nothing Then Exit Sub
    If pLink.NextLink Is Nothing Then Set pLink.NextLink = New CStringLink
    Set pLink.NextLink.NextLink = cLink
    
End Sub

Private Function GetKey(ByVal vLink As CStringLink) As String
    On Error Resume Next
    GetKey = vLink.Data
End Function

Private Function GetValue(ByVal vLink As CStringLink) As String
    On Error Resume Next
    GetValue = vLink.NextLink.Data
End Function
Private Sub SetKey(ByVal vLink As CStringLink, ByRef vKey As String)
    If vLink Is Nothing Then Set vLink = New CStringLink
    vLink.Data = vKey
End Sub
Private Sub SetValue(ByVal vLink As CStringLink, ByRef vValue As String)
    On Error Resume Next
    If vLink Is Nothing Then Exit Sub
    Dim cLink As CStringLink
    Set cLink = vLink.NextLink
    If cLink Is Nothing Then Set cLink = New CStringLink
    cLink.Data = vValue
    Set vLink.NextLink = cLink
End Sub

Private Function GetNextLink(ByVal vLink As CStringLink) As CStringLink
    On Error Resume Next
    Set GetNextLink = vLink.NextLink.NextLink
End Function

Private Sub SetUpLink(ByVal vLink As CStringLink, ByRef vKey As String, ByRef vValue As String)
    If vLink Is Nothing Then Set vLink = New CStringLink
    vLink.Data = vKey
    If vLink.NextLink Is Nothing Then Set vLink.NextLink = New CStringLink
    vLink.NextLink.Data = vValue
End Sub

Public Property Get Keys() As String()
    Dim i As Long
    Dim u As Long
    Dim c As Long
    u = mHashTableSize
    Dim result() As String
    Dim cLink As CStringLink
    For i = 0 To u
        Set cLink = mHashTable(i)
        Do Until cLink Is Nothing
            ReDim Preserve result(0 To c)
            result(c) = GetKey(cLink)
            c = c + 1
            Set cLink = GetNextLink(cLink)
        Loop
    Next
    Keys = result
End Property

Public Property Get Values() As String()
    Dim i As Long
    Dim u As Long
    Dim c As Long
    u = mHashTableSize
    Dim result() As String
    Dim cLink As CStringLink
    For i = 0 To u
        Set cLink = mHashTable(i)
        Do Until cLink Is Nothing
            ReDim Preserve result(0 To c)
            result(c) = GetValue(cLink)
            c = c + 1
            Set cLink = GetNextLink(cLink)
        Loop
    Next
    Values = result
End Property

Public Sub Reset()
    Init mHashTableSize
End Sub

Public Property Get KeyLinks() As CStringLink
     Dim i As Long
    Dim u As Long
    Dim c As Long
    u = mHashTableSize
    Dim cLink As CStringLink
    Dim result As CStringLink
    For i = 0 To u
        Set cLink = mHashTable(i)
        Do Until cLink Is Nothing
            If result Is Nothing Then
                Set KeyLinks = New CStringLink
                Set result = KeyLinks
                result.Data = GetKey(cLink)
            Else
                Set result = result.AppendNext(GetKey(cLink))
            End If
            Set cLink = GetNextLink(cLink)
        Loop
    Next
    
End Property

Public Property Get ValueLinks() As CStringLink
     Dim i As Long
    Dim u As Long
    Dim c As Long
    u = mHashTableSize
    Dim cLink As CStringLink
    Dim result As CStringLink
    For i = 0 To u
        Set cLink = mHashTable(i)
        Do Until cLink Is Nothing
            If result Is Nothing Then
                Set ValueLinks = New CStringLink
                Set result = ValueLinks
                'Set result = New CStringLink
                result.Data = GetValue(cLink)
            Else
                Set result = result.AppendNext(GetValue(cLink))
            End If
            Set cLink = GetNextLink(cLink)
        Loop
    Next
End Property

Public Property Get ToStringLink() As CStringLink
    Dim i As Long
    Dim u As Long
    Dim c As Long
    u = mHashTableSize
    Dim cLink As CStringLink
    Dim result As CStringLink
    For i = 0 To u
        Set cLink = mHashTable(i)
        Do Until cLink Is Nothing
            If result Is Nothing Then
                Set result = New CStringLink
                Set ToStringLink = result
                result.Data = GetKey(cLink)
                Set result = result.AppendNext(GetValue(cLink))
            Else
                Set result = result.AppendNext(GetKey(cLink))
                Set result = result.AppendNext(GetValue(cLink))
            End If
            Set cLink = GetNextLink(cLink)
        Loop
    Next
    'Set ToStringLink = result
End Property

