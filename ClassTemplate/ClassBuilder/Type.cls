VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Implements ITypeInfo



Private Type TypeInfo
    Template As String
    Name As String
    Style As CTTypeStyles
End Type


Private mType() As TypeInfo
Private mTypeStyle As CTypeStyle
Private mPrefix As String
Private mSuffix As String

Private mCount As Long
Private mSize As Long
Private Const DEFAULT_SIZE As Long = 26
Private mStabled As Boolean

Private Sub Class_Initialize()
    ReDim mType(0 To DEFAULT_SIZE - 1) As TypeInfo
    mSize = DEFAULT_SIZE
    Set mTypeStyle = New CTypeStyle
End Sub


Private Property Get ITypeInfo_ConstTypePrefix() As String
    ITypeInfo_ConstTypePrefix = mPrefix
End Property

Private Property Get ITypeInfo_ConstTypeSuffix() As String
    ITypeInfo_ConstTypeSuffix = mSuffix
End Property

Private Function ITypeInfo_ConstVarOf(ByVal ctStyle As CTTypeStyles) As String
    ITypeInfo_ConstVarOf = mTypeStyle.ConstVarOf(ctStyle)
End Function

Private Property Get ITypeInfo_Count() As Long
    ITypeInfo_Count = mCount
End Property

Private Function ITypeInfo_GetAllName() As String()
    ReDim Result(0 To mCount - 1, 0 To 1) As String
    Dim i As Long
    For i = 0 To mCount - 1
        Result(i, 0) = mType(i).Template
        Result(i, 1) = mType(i).Name
    Next
    ITypeInfo_GetAllName = Result
End Function

Private Function ITypeInfo_IdOf(sTemplate As String) As Integer
    ITypeInfo_IdOf = IndexOf(sTemplate)
End Function

Private Function ITypeInfo_NameOf(ByVal nId As Integer) As String
    If nId >= 0 And nId < mCount Then ITypeInfo_NameOf = mType(nId).Name
End Function


Public Sub Register(sTemplate As String, sName As String, ttStyle As CTTypeStyles)
    Dim nIndex As Long
    nIndex = IndexOf(sTemplate)
    If (nIndex >= 0) Then
        mType(nIndex).Name = sName
        mType(nIndex).Style = ttStyle
    Else
        If mCount >= mSize - 1 Then
            mSize = mSize + DEFAULT_SIZE
            ReDim Preserve mType(0 To mSize - 1) As TypeInfo
        End If
        mType(mCount).Template = sTemplate
        mType(mCount).Name = sName
        mType(mCount).Style = ttStyle
        mCount = mCount + 1
    End If
End Sub

Private Property Let ITypeInfo_Stabled(RHS As Boolean)
    mStabled = RHS
End Property

Private Property Get ITypeInfo_Stabled() As Boolean
    ITypeInfo_Stabled = mStabled
End Property

Private Function ITypeInfo_StyleOf(ByVal nId As Integer) As CTTypeStyles
    If nId >= 0 And nId < mCount Then ITypeInfo_StyleOf = mType(nId).Style
End Function

Private Function IndexOf(sTemplate As String) As Long
    Static lastTemplate As String
    Static lastIndex As String
    
    IndexOf = -1
    If mCount < 1 Then Exit Function
    If sTemplate = "" Then Exit Function
    
    If mStabled Then
        If (lastTemplate = sTemplate) Then IndexOf = lastIndex: Exit Function
    End If
    Dim i As Long
    For i = 0 To mCount - 1
        If mType(i).Template = sTemplate Or mType(i).Name = sTemplate Then IndexOf = i: Exit For
    Next
    If mStabled Then
        lastTemplate = sTemplate
        lastIndex = IndexOf
    End If
End Function

Public Property Set typeStyle(ByRef oTypeStyle As CTypeStyle)
    If oTypeStyle Is Nothing Then Exit Property
    Set mTypeStyle = oTypeStyle
End Property

Public Property Let ConstTypePrefix(ByVal sPrefix As String)
    mPrefix = sPrefix
End Property

Public Property Let ConstTypeSuffix(ByVal sSuffix As String)
    mSuffix = sSuffix
End Property
