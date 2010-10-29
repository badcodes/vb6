VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CStringCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Collection(Fake) for Strings"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Public Enum CStringCollectionType
    sctUnique = 0
    sctSimple = 1
End Enum
Private iItemCount As Long 'local copy
Private iCurIndex As Long
'Private sLargeString As New CAppendString
Private strMagic As String
Private iEndPos() As Long
Private sctMe As CStringCollectionType
Private Const perMount = 300
Private Const scEmptyString = vbNullChar & "Em" & vbNullChar & "PtySt" & vbNullChar & "RinG" & vbNullChar

Public Property Get CollectionType() As CStringCollectionType

    CollectionType = sctMe

End Property

Public Property Let CollectionType(sctLet As CStringCollectionType)

    sctMe = sctLet

End Property

Public Property Get ToString() As String

    ToString = Replace$(strMagic, scEmptyString, "")

End Property

Public Function getString(Index As Long) As String
Attribute getString.VB_UserMemId = 0

    If Index < 1 Then Exit Function

    If Index > iItemCount Then Exit Function
    'Dim strMagic As String
    'strMagic = sLargeString.Value

    If strMagic = "" Then Exit Function
    getString = Mid$(strMagic, iEndPos(Index - 1) + 1, iEndPos(Index) - iEndPos(Index - 1))
    getString = filterOutputString(getString)

End Function

Public Function getIndex(ByVal StrSearch As String, Optional cmp As VbCompareMethod = vbTextCompare) As Long

    StrSearch = filterInputString(StrSearch) '"" Then Exit Function
    'Dim strMagic As String
    Dim pos As Long
    Dim nextPos As Long
    Dim bGetit As Boolean
    Dim iPos2 As Long
    Dim l As Long
    'strMagic = sLargeString.Value
    'If strMagic = "" Then Exit Function

    Do Until getIndex > 0
        pos = InStr(nextPos + 1, strMagic, StrSearch, cmp)

        If pos < 1 Then Exit Do
        nextPos = pos
        pos = pos - 1
        iPos2 = pos + Len(StrSearch)

        For l = 0 To iItemCount

            If pos = iEndPos(l) And iPos2 = iEndPos(l + 1) Then getIndex = l + 1: Exit For

            If pos < iEndPos(l) Then Exit For
        Next

    Loop

End Function

Public Property Get count() As Long

    count = iItemCount

End Property

Public Sub add(ByVal strComing As String)

    strComing = filterInputString(strComing)

    If sctMe = sctUnique Then Remove strComing
    iItemCount = iItemCount + 1

    If iItemCount > UBound(iEndPos) Then ReDim Preserve iEndPos(iItemCount + perMount) As Long
    iEndPos(iItemCount) = iEndPos(iItemCount - 1) + Len(strComing)
    'sLargeString.Append strcoming
    strMagic = strMagic & strComing

End Sub

Private Sub Class_Initialize()

    ReDim iEndPos(perMount) As Long
    sctMe = sctSimple

End Sub

Public Sub clear()

    iItemCount = 0
    iCurIndex = 0
    strMagic = ""
    'sLargeString.clear
    'ReDim iEndPos(perMount) As Long

End Sub

Public Function getLength(Index As Long) As Long

    If Index < 1 Then Exit Function

    If Index > iItemCount Then Exit Function
    getLength = Len(getString(Index)) ' iEndPos(Index) - iEndPos(Index - 1)

End Function

Public Property Get Length() As Long

    Length = Len(Replace$(strMagic, scEmptyString, ""))

End Property

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

Public Function Remove(IndexOrString As Variant) As Boolean

    Dim iIndex As Long
    Dim sTemp As String
    Dim iLoop As Long
    Dim iEndLoop As Long
    Dim iRemoveLength As Long

    If VarType(IndexOrString) = vbString Then
        iIndex = getIndex(CStr(IndexOrString))
    Else
        iIndex = CLng(IndexOrString)
    End If

    If iIndex < 1 Or iIndex > iItemCount Then Exit Function
    iRemoveLength = iEndPos(iIndex) - iEndPos(iIndex - 1)
    sTemp = strMagic
    strMagic = Left$(sTemp, iEndPos(iIndex - 1))
    strMagic = strMagic & Mid$(sTemp, iEndPos(iIndex) + 1)
    iEndLoop = iItemCount - 1

    For iLoop = iIndex To iEndLoop
        iEndPos(iLoop) = iEndPos(iLoop + 1) - iRemoveLength
    Next

    iItemCount = iItemCount - 1
    Remove = True

End Function

'Public Sub Insert(ByVal strComing As String, Optional iBefore As Long = -1)
'Dim strTemp As String
'Dim iStrLen As Long
'
'If iBefore < 1 Then iBefore = 1
'If iBefore > iItemCount Then iBefore = -1
'If iBefore = -1 Then
'    add strComing
'    Exit Sub
'End If
'strComing = filterInputString(strComing)
'
'iStrLen = Len(strComing)
'strTemp = strMagic
'strMagic = Left$(strTemp, iEndPos(iBefore - 1))
'strMagic = strMagic & strComing
'strMagic = strMagic & Mid$(strTemp, iEndPos(iBefore - 1) + 1)
'
'End Sub
Private Function filterInputString(ByRef strInput As String) As String

    If strInput = "" Then
        filterInputString = scEmptyString
    Else
        filterInputString = strInput
    End If

End Function

Private Function filterOutputString(ByRef strOutput As String) As String

    If strOutput = scEmptyString Then
        filterOutputString = ""
    Else
        filterOutputString = strOutput
    End If

End Function

