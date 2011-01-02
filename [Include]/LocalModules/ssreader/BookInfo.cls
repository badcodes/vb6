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
Private Const CST_END_OF_PROPERTY As String = vbNullString & vbNullString & "EOF" & vbNullString & vbNullString

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
        If vValue <> vbNullString And Right$(vValue, 1) <> "\" Then vValue = vValue & "\"
    End If
    mProperties(vFieldId) = vValue
End Property

Public Sub Reset()
On Error Resume Next
    Dim i As Long
    For i = CST_SSLIB_FIELDS_LBound To CST_SSLIB_FIELDS_UBound
        mProperties(i) = vbNullString
    Next
End Sub

'CSEH: ErrResumeNext
Public Sub SerializeTo(ByVal vFilename As String)
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>
    Dim i As Long
    Dim fNum As Integer
    Open vFilename For Output As #fNum
    For i = CST_SSLIB_FIELDS_LBound To CST_SSLIB_FIELDS_UBound
        Print #fNum, SSLIB_EngFieldName(i)
        Print #fNum, mProperties(i)
        Print #fNum, CST_END_OF_PROPERTY
    Next
    Close #fNum
End Sub

'CSEH: ErrResumeNext
Public Sub SerializeFrom(ByVal vFilename As String)
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>
    Dim i As Long
    Dim fNum As Integer
    fNum = FreeFile
    i = -1
    Open vFilename For Input As #fNum
    Dim tmp As String
    Do Until EOF(fNum)
        Line Input #fNum, tmp
        If i < 0 Then
            i = MSSReader.SSLIB_GetFiledID(tmp)
        ElseIf tmp = CST_END_OF_PROPERTY Then
            If Right$(mProperties(i), 2) = vbCrLf Then
                mProperties(i) = Left$(mProperties(i), Len(mProperties(i)) - 2)
            End If
        Else
            mProperties(i) = mProperties(i) & tmp & vbCrLf
        End If
    Loop
    Close #fNum
End Sub

Public Sub LoadFromFile(ByRef pFilename As String, Optional vIniHnd As CLiNInI = Nothing, Optional vSection As String = cst_Section_Name)
    Dim ini As CLiNInI
    If vIniHnd Is Nothing Then
        Set ini = New CLiNInI
        ini.Source = pFilename
    Else
        Set ini = vIniHnd
        If (ini.Source = vbNullString) Then ini.Source = pFilename
    End If
    
    Dim pSection As String
    If vSection = cst_Section_Name Then
        If ini.SectionExists(vSection) Then
            pSection = vSection
        Else
            pSection = "TaskInfo"
        End If
    Else
        pSection = vSection
    End If
    If ini.SectionExists(pSection) = False Then Exit Sub
    
    Dim i As Long
    Dim tmp As String
    For i = CST_SSLIB_FIELDS_LBound To CST_SSLIB_FIELDS_UBound
        tmp = ini.GetSetting(pSection, SSLIB_ChnFieldName(i))
'        If tmp <> vbNULLSTRING Then
        Field(i) = tmp
    Next
    
End Sub


Public Sub LoadFromArray(ByRef vArray() As String, Optional vAcceptEmpty As Boolean = True)
    On Error Resume Next
    Dim i As Long
        For i = CST_SSLIB_FIELDS_LBound To CST_SSLIB_FIELDS_UBound
        If vAcceptEmpty Or vArray(i) <> vbNullString Then Field(i) = vArray(i)
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
        If (ini.Source = vbNullString) Then ini.Source = pFilename
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
