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
Attribute VB_Exposed = False
Option Explicit

'Implements ISetting
Private mSource As String
Private mSection As String
Public Enum EVBSetting_ControlType
    evbsctText
    evbsctValue
End Enum

Private Property Let ISetting_CompareMethod(ByVal RHS As VbCompareMethod)
End Property

Private Property Get ISetting_CompareMethod() As VbCompareMethod
End Property

Private Function ISetting_ReadKey(ByVal Key As String, Optional ByVal DefaultValue As String) As String
    ISetting_ReadKey = ReadKey(Key, DefaultValue)
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

Private Sub ISetting_WriteKey(ByVal Key As String, ByVal Value As String)
    WriteKey Key, Value
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

Public Function ReadKey(ByVal Key As String, Optional ByVal DefaultValue As String) As String
   ReadKey = GetSetting(mSource, mSection, Key, DefaultValue)
End Function

Public Sub WriteKey(ByVal Key As String, ByVal Value As String)
    SaveSetting mSource, mSection, Key, Value

End Sub

Public Sub ReadObjectArray(ByRef vObject As Object, Optional vType As EVBSetting_ControlType, Optional vKey As String, Optional vDefault As String)
    Dim iStart As Long
    Dim iEnd As Long
    iStart = -1
    iEnd = -1
    On Error Resume Next
    iStart = vObject.lbound
    iEnd = vObject.ubound
    If iStart < 0 Or iEnd < 0 Then Exit Sub
    If vKey = "" Then vKey = TypeName(vObject(iStart)) & "_" & vObject(iStart).Name
    Dim i As Long
    Select Case vType
        Case EVBSetting_ControlType.evbsctText
            For i = iStart To iEnd
                ReadPropText vObject(i), vKey & i, ""
            Next
        Case EVBSetting_ControlType.evbsctValue
            For i = iStart To iEnd
                ReadPropValue vObject(i), vKey & i
            Next
    End Select
End Sub

Public Sub WriteObjectArray(ByRef vObject As Object, Optional vType As EVBSetting_ControlType, Optional vKey As String, Optional vDefault As String)
    Dim iStart As Long
    Dim iEnd As Long
    iStart = -1
    iEnd = -1
    On Error Resume Next
    iStart = vObject.lbound
    iEnd = vObject.ubound
    If iStart < 0 Or iEnd < 0 Then Exit Sub
    If vKey = "" Then vKey = TypeName(vObject(iStart)) & "_" & vObject(iStart).Name
    Dim i As Long
    Select Case vType
        Case EVBSetting_ControlType.evbsctText
            For i = iStart To iEnd
                WritePropText vObject(i), vKey & i
            Next
        Case EVBSetting_ControlType.evbsctValue
            For i = iStart To iEnd
                WritePropValue vObject(i), vKey & i
            Next
    End Select
End Sub

Public Sub ReadPropTexts(ByRef vObject As Object, Optional vKey As String, Optional vDefault As String)
    ReadObjectArray vObject, evbsctText, vKey, vDefault
End Sub

Public Sub WritePropTexts(ByRef vObject As Object, Optional vKey As String)
    WriteObjectArray vObject, evbsctText, vKey
End Sub

Public Sub ReadPropValues(ByRef vObject As Object, Optional vKey As String, Optional vDefault As String)
    ReadObjectArray vObject, evbsctValue, vKey, vDefault
End Sub
Public Sub WritePropValues(ByRef vObject As Object, Optional vKey As String)
    WriteObjectArray vObject, evbsctValue, vKey
End Sub
Public Sub ReadPropText(ByRef dstObject As Object, Key As String, Optional DefaultValue As String)
    dstObject.Text = ReadKey(Key, DefaultValue)
End Sub
Public Sub ReadPropValue(ByRef dstObject As Object, Key As String)
    On Error Resume Next
    dstObject.Value = ReadKey(Key, "")
End Sub

'Public Sub ReadProp(ByRef vObject As Object, ByVal vPropName As String, Optional Key As String, Optional vDefault As String)
'    On Error Resume Next
'
'End Sub


Public Sub WritePropValue(ByRef dstObject As Object, Key As String)
    On Error Resume Next
    WriteKey Key, dstObject.Value
End Sub
Public Sub ReadPropList(ByRef dstObject As Object, Key As String)
    Dim idx As Long
    Dim count As Long
    count = CLng(ReadKey(Key & ".count", "0"))
    idx = CLng(ReadKey(Key & ".index", "-1"))
    If count <= 0 Then Exit Sub
    Dim i As Integer
    For i = 1 To count
        dstObject.AddItem ReadKey(Key & ".item" & CStr(i))
    Next
    If idx > 0 Then dstObject.ListIndex = idx
End Sub

Public Sub ReadPropListItems(ByRef dstObject As Object, Key As String)
    Dim count As Long
    count = CLng(ReadKey(Key & ".count", "0"))
    If (count < 1) Then Exit Sub
    Dim i As Long
    Dim item As Object
    Dim subItem As Object
    Dim cSub As Long
    Dim j As Long
    For i = 1 To count
        Set item = dstObject.ListItems.Add(, , ReadKey(Key & ".item" & CStr(i), ""))
        cSub = CLng(ReadKey(Key & ".item" & CStr(i) & ".count", "0"))
        For j = 1 To cSub
            item.ListSubItems.Add , , ReadKey(Key & ".item" & CStr(i) & ".item" & CStr(j), "")
        Next
    Next
End Sub
Public Sub WritePropText(ByRef dstObject As Object, Key As String)
    WriteKey Key, dstObject.Text
End Sub

Public Sub WritePropList(ByRef dstObject As Object, Key As String)
    Dim idx As Long
    Dim count As Long
    idx = dstObject.ListIndex
    count = dstObject.ListCount
    WriteKey Key & ".count", count
    WriteKey Key & ".index", idx
    If count <= 0 Then Exit Sub
    Dim i As Integer
    For i = 1 To count
        WriteKey Key & ".item" & CStr(i), dstObject.List(i)
    Next
End Sub

Public Sub WritePropListItems(ByRef dstObject As Object, Key As String)
    Dim i As Long, j As Long, count As Long, cSub As Long
    count = dstObject.ListItems.count
    WriteKey Key & ".count", CStr(count)
    For i = 1 To count
        WriteKey Key & ".item" & CStr(i), dstObject.ListItems(i).Text
        cSub = dstObject.ListItems(i).ListSubItems.count
        WriteKey Key & ".item" & CStr(i) & ".count", CStr(cSub)
        For j = 1 To cSub
            WriteKey Key & ".item" & CStr(i) & ".item" & CStr(j), dstObject.ListItems(i).ListSubItems(j).Text
        Next
    Next
End Sub


