VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CAddOnlyStringArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private iItemCount As Long 'local copy
Private iCurIndex As Long
Private strMagic() As String
Private Const perMount = 300
Public Function getString(Index As Long) As String
Attribute getString.VB_UserMemId = 0
    If Index < 1 Then Exit Function
    If Index > iItemCount Then Exit Function
    getString = strMagic(Index)
End Function

Public Function getIndex(ByVal StrSearch As String, Optional StartIndex As Long = 1, Optional cmp As VbCompareMethod = vbTextCompare) As Long
    If StartIndex < 1 Then StartIndex = 1
    If StartIndex > iItemCount Then Exit Function
    Dim i As Long
    For i = StartIndex To iItemCount
    If StrComp(strMagic(i), StrSearch, cmp) = 0 Then getIndex = i: Exit Function
    Next
End Function

Public Property Get count() As Long
    count = iItemCount
End Property

Public Sub add(ByVal strComing As String)
    iItemCount = iItemCount + 1
    If iItemCount > UBound(strMagic) Then ReDim Preserve strMagic(1 To iItemCount + perMount) As String
    strMagic(iItemCount) = strComing
End Sub

Private Sub Class_Initialize()
    ReDim strMagic(1 To perMount) As String
End Sub

Public Sub clear()
    iItemCount = 0
    iCurIndex = 0
    'ReDim strMagic(1 To perMount) As String
End Sub

Public Function getLength(Index As Long) As Long
    If Index < 1 Then Exit Function
    If Index > iItemCount Then Exit Function
    getLength = Len(getString(Index))
End Function


Public Property Get HaveNextItem() As Boolean
    If iCurIndex < iItemCount Then HaveNextItem = True
End Property

Public Property Get NextItem() As String
    If iCurIndex >= iItemCount Then Exit Property
    iCurIndex = iCurIndex + 1
    NextItem = getString(iCurIndex)
End Property

Public Sub StartEnum()
    iCurIndex = 0
End Sub



