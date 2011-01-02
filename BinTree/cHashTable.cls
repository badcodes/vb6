VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cHashTable"
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

Private Function hashName(name As String) As Long
Dim v As Long
Dim i As Long
Dim iEnd As Long
iEnd = Len(name)
For i = 1 To iEnd
    v = v + AscW(Mid$(name, i, 1))
Next
hashName = v Mod (MAXTAB - 1) + 1
End Function

Public Sub Insert(name As String, value As String)
Dim index As Long
index = hashName(name)
If mTable(index) = "" Then
    mTable(index) = value
Else
    mTable(index) = mTable(index) & hashTable_SplitFlag & value
End If
End Sub

Public Sub value(name As String, ByRef valueOut() As String)
Dim index As Long
index = hashName(name)
valueOut = Split(mTable(index), hashTable_SplitFlag)
End Sub

Public Sub reset()
Erase mTable
Dim mTable(1 To MAXTAB) As String
End Sub

Private Sub Class_Initialize()
hashTable_SplitFlag = Chr$(0) & "," & Chr$(0)
End Sub

Private Sub Class_Terminate()
Erase mTable
End Sub
