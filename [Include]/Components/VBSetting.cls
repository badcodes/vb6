VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CVBSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Implements ISetting
Private mSource As String
Private mSection As String
Private Property Let ISetting_CompareMethod(ByVal RHS As VbCompareMethod)
End Property

Private Property Get ISetting_CompareMethod() As VbCompareMethod
End Property

Private Function ISetting_ReadKey(ByVal key As String, Optional ByVal DefaultValue As String) As String
    ISetting_ReadKey = ReadKey(key, DefaultValue)
End Function

Private Property Let ISetting_Section(ByVal RHS As String)
    mSection = RHS
End Property

Private Property Get ISetting_Section() As String
    ISetting_Section = mSection
End Property

Private Property Let ISetting_Source(ByVal RHS As String)
    mSource = RHS
End Property

Private Property Get ISetting_Source() As String
    ISetting_Source = mSource
End Property

Private Sub ISetting_WriteKey(ByVal key As String, ByVal Value As String)
    WriteKey key, Value
End Sub

Public Property Get Section() As String
    Section = mSection
End Property

Public Property Let Section(ByVal Value As String)
    mSection = Value
End Property

Public Property Get Appname() As String
    Appname = mSource
End Property

Public Property Let Appname(ByVal Value As String)
    mSource = Value
End Property

Private Sub Class_Initialize()
'    mCompareMethod = vbTextCompare
End Sub

Public Function ReadKey(ByVal key As String, Optional ByVal DefaultValue As String) As String
   ReadKey = GetSetting(mSource, mSection, key, DefaultValue)
End Function

Public Sub WriteKey(ByVal key As String, ByVal Value As String)
    SaveSetting mSource, mSection, key, Value

End Sub

Public Sub ReadPropText(ByRef dstObject As Object, key As String, Optional DefaultValue As String)
    dstObject.Text = ReadKey(key, DefaultValue)
End Sub
Public Sub ReadPropValue(ByRef dstObject As Object, key As String)
    On Error Resume Next
    dstObject.Value = ReadKey(key, "")
End Sub
Public Sub WritePropValue(ByRef dstObject As Object, key As String)
    On Error Resume Next
    WriteKey key, dstObject.Value
End Sub
Public Sub ReadPropList(ByRef dstObject As Object, key As String)
    Dim idx As Long
    Dim count As Long
    count = CLng(ReadKey(key & ".count", "0"))
    idx = CLng(ReadKey(key & ".index", "-1"))
    If count <= 0 Then Exit Sub
    Dim i As Integer
    For i = 1 To count
        dstObject.AddItem ReadKey(key & ".item" & CStr(i))
    Next
    If idx > 0 Then dstObject.ListIndex = idx
End Sub

Public Sub ReadPropListItems(ByRef dstObject As Object, key As String)
    Dim count As Long
    count = CLng(ReadKey(key & ".count", "0"))
    If (count < 1) Then Exit Sub
    Dim i As Long
    Dim item As Object
    Dim subItem As Object
    Dim cSub As Long
    Dim j As Long
    For i = 1 To count
        Set item = dstObject.ListItems.Add(, , ReadKey(key & ".item" & CStr(i), ""))
        cSub = CLng(ReadKey(key & ".item" & CStr(i) & ".count", "0"))
        For j = 1 To cSub
            item.ListSubItems.Add , , ReadKey(key & ".item" & CStr(i) & ".item" & CStr(j), "")
        Next
    Next
End Sub
Public Sub WritePropText(ByRef dstObject As Object, key As String)
    WriteKey key, dstObject.Text
End Sub

Public Sub WritePropList(ByRef dstObject As Object, key As String)
    Dim idx As Long
    Dim count As Long
    idx = dstObject.ListIndex
    count = dstObject.ListCount
    WriteKey key & ".count", count
    WriteKey key & ".index", idx
    If count <= 0 Then Exit Sub
    Dim i As Integer
    For i = 1 To count
        WriteKey key & ".item" & CStr(i), dstObject.List(i)
    Next
End Sub

Public Sub WritePropListItems(ByRef dstObject As Object, key As String)
    Dim i As Long, j As Long, count As Long, cSub As Long
    count = dstObject.ListItems.count
    WriteKey key & ".count", CStr(count)
    For i = 1 To count
        WriteKey key & ".item" & CStr(i), dstObject.ListItems(i).Text
        cSub = dstObject.ListItems(i).ListSubItems.count
        WriteKey key & ".item" & CStr(i) & ".count", CStr(cSub)
        For j = 1 To cSub
            WriteKey key & ".item" & CStr(i) & ".item" & CStr(j), dstObject.ListItems(i).ListSubItems(j).Text
        Next
    Next
End Sub



