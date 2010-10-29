VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CHashTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private hashTable_SplitFlag   As String
Const MAXTAB = 200
Private mTable(1 To MAXTAB) As String

Private Function hashName(Name As String) As Long
Dim v As Long
Dim i As Long
Dim iEnd As Long
iEnd = Len(Name)
For i = 1 To iEnd
    v = v + AscW(Mid$(Name, i, 1))
Next
If v < 0 Then v = Abs(v)
hashName = v Mod (MAXTAB - 1) + 1
End Function

Public Sub Insert(Name As String, Value As String)
Dim index As Long
index = hashName(Name)
If mTable(index) = "" Then
    mTable(index) = Value
Else
    mTable(index) = mTable(index) & hashTable_SplitFlag & Value
End If
End Sub

Public Function Value(Name As String, ByRef valueOut() As String) As Long
Dim index As Long
Value = -1
index = hashName(Name)
If mTable(index) = "" Then Exit Function
valueOut = Split(mTable(index), hashTable_SplitFlag)
Value = UBound(valueOut)
End Function

Public Sub reset()
Erase mTable
End Sub

Private Sub Class_Initialize()
    hashTable_SplitFlag = Chr$(0) & "," & Chr$(0)
End Sub
