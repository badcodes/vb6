VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBookInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private mProperties(CST_SSLIB_FIELDS_LBound To CST_SSLIB_FIELDS_UBound) As String
'As CStringMap
'Private Const cst_Properties_Max_Size As Integer = 256
Private Const cst_Section_Name As String = "General Information"


Public Property Get Field(ByRef vFieldId As SSLIBFields) As String
Attribute Field.VB_UserMemId = 0
On Error Resume Next
    Field = mProperties(vFieldId)
End Property
Public Property Let Field(ByRef vFieldId As SSLIBFields, ByVal vValue As String)
    If vFieldId = SSF_URL Then
        If Right$(vValue, 5) = ".pdg/" Then vValue = Mid$(vValue, 1, Len(vValue) - 1)
        If Right$(vValue, 1) <> "/" Then
            Dim pos As Long
            pos = InStrRev(vValue, "/")
            If pos > 0 Then
                vValue = Mid$(vValue, 1, pos)
            End If
            'vValue = vValue & "/"
        End If
    ElseIf vFieldId = SSF_SAVEDIN Then
        If Right$(vValue, 1) <> "\" Then vValue = vValue & "\"
    End If
    mProperties(vFieldId) = vValue
End Property

Public Sub Reset()
On Error Resume Next
    Dim i As Long
    For i = CST_SSLIB_FIELDS_LBound To CST_SSLIB_FIELDS_UBound
        mProperties(i) = ""
    Next
End Sub

Public Sub LoadFromFile(ByRef pFilename As String, Optional vIniHnd As CLiNInI = Nothing, Optional vSection As String = cst_Section_Name)
    Dim ini As CLiNInI
    If vIniHnd Is Nothing Then
        Set ini = New CLiNInI
        ini.Source = pFilename
    Else
        Set ini = vIniHnd
        If (ini.Source = "") Then ini.Source = pFilename
    End If
    
    Dim i As Long
    Dim tmp As String
    For i = CST_SSLIB_FIELDS_LBound To CST_SSLIB_FIELDS_UBound
        tmp = ini.GetSetting(vSection, SSLIB_ChnFieldName(i))
'        If tmp <> "" Then
        Field(i) = tmp
    Next
    
End Sub


Public Sub LoadFromArray(ByRef vArray() As String, Optional vAcceptEmpty As Boolean = True)
    On Error Resume Next
    Dim i As Long
        For i = CST_SSLIB_FIELDS_LBound To CST_SSLIB_FIELDS_UBound
        If vAcceptEmpty Or vArray(i) <> "" Then Field(i) = vArray(i)
    Next
End Sub

Public Sub SaveToFile(ByRef pFilename As String, Optional vIniHnd As CLiNInI = Nothing, Optional vSection As String = cst_Section_Name, Optional pBriefMode As Boolean = False)
    On Error GoTo ErrorSaveToFile

    Dim ini As CLiNInI
    If vIniHnd Is Nothing Then
        Set ini = New CLiNInI
        ini.Source = pFilename
    Else
        Set ini = vIniHnd
        If (ini.Source = "") Then ini.Source = pFilename
    End If
    
    Dim u As Long
    
    If pBriefMode Then
        u = CST_SSLIB_FIELDS_IMPORTANT_UBOUND
    Else
        u = CST_SSLIB_FIELDS_UBound
    End If
    
    Dim i As Long
    For i = CST_SSLIB_FIELDS_LBound To u
        ini.SaveSetting vSection, SSLIB_ChnFieldName(i), mProperties(i)
    Next
    If vIniHnd Is Nothing Then ini.Save
    Exit Sub
ErrorSaveToFile:
    Exit Sub
End Sub

Private Sub Class_Initialize()
    SSLIB_Init
End Sub
