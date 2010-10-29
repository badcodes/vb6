VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLiNInI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'--------------------------------------------------------------------------------
'    Component  : CLiNInI
'
'    Description: ini File Class
'
'    Author   : xrLin
'
'    Date     : 2005-2006
'--------------------------------------------------------------------------------
Option Explicit

    Private m_iniFile As String
    Private hIni As CIniText
    'Private m_Text As CString
'local variable(s) to hold property value(s)

Public Property Let CompareMethod(ByVal vData As VbCompareMethod)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.CompareMethod = Form1
    hIni.CompareMethod = vData
    
End Property


Public Property Get CompareMethod() As VbCompareMethod
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CompareMethod
    CompareMethod = hIni.CompareMethod
End Property



    
    'Private Const cstExcapeBrace = "L" & vbNullChar & "eFT" & vbNullChar & "B" & vbNullChar & "raCe" & vbNullChar

Public Function ToString() As String
    ToString = hIni.iniString
End Function

Public Sub Clear()
    hIni.Clear
End Sub

Public Sub DeleteSection(ByRef Section As String)

    hIni.DeleteSection Section
End Sub
Public Sub DeleteSetting(ByRef sSection As String, ByVal sKeyName As String)
    
    hIni.DeleteSetting sSection, sKeyName
End Sub
Public Sub SaveSetting(ByRef sSection As String, ByVal sKeyName As String, ByVal sValue As String)
    hIni.SaveSetting sSection, sKeyName, sValue
    
End Sub
Public Function GetSetting(ByRef sSection As String, ByVal sKeyName As String) As String

    GetSetting = hIni.GetSetting(sSection, sKeyName)

End Function

Public Function GetSectionText(ByRef Section As String) As String

    GetSectionText = hIni.GetSectionText(Section)
End Function

Public Sub SaveSection(ByRef Section As String, ByRef sSecText As String)
    hIni.SaveSection Section, sSecText
    
End Sub

Public Sub Create(ByRef inifileName As String)

    File = inifileName

End Sub

Private Function FileExists(ByRef FileName As String) As Boolean
    On Error GoTo errorFileExists
    FileLen FileName
    FileExists = True
    Exit Function
errorFileExists:
    FileExists = False
    Err.Clear
End Function


Public Sub SaveTo(ByRef FileName As String)
Dim fNUM As Long
    'Dim l As Long
    Dim bUnicode As Boolean
    Dim sText As String
    
    bUnicode = True 'isUnicode(m_iniFile)
    sText = Me.ToString
    
    fNUM = FreeFile
    On Error GoTo FileWriteErr:
    If FileExists(FileName) Then Kill FileName
    If bUnicode Then
        Dim c_B(1) As Byte
        ReDim bText(LenB(sText)) As Byte
        c_B(0) = 255
        c_B(1) = 254
        bText = sText
        'Open fileName For Output As #fNum
        Open FileName For Binary Access Write As #fNUM
        Put #fNUM, , c_B()
        Put #fNUM, , bText
    Else
        Open FileName For Output As #fNUM
        Print #fNUM, ToString;
    End If
    
    Close #fNUM
    
    Exit Sub
FileWriteErr:
    Err.Raise Err.Number, "CLini.SaveTo", Err.Description
    
End Sub

Public Sub Save()

    Dim fNUM As Long
    'Dim l As Long
    Dim bUnicode As Boolean
    Dim sText As String
    
    bUnicode = True 'isUnicode(m_iniFile)
    sText = Me.ToString
    
    fNUM = FreeFile
    On Error GoTo FileWriteErr:
    
    If bUnicode Then
        Dim c_B(1) As Byte
        ReDim bText(LenB(sText)) As Byte
        c_B(0) = 255
        c_B(1) = 254
        bText = sText
        Open m_iniFile For Binary Access Write As #fNUM
        Put #fNUM, , c_B()
        Put #fNUM, , bText
    Else
        Open m_iniFile For Output As #fNUM
        Print #fNUM, ToString;
    End If
    
    Close #fNUM
    
    Exit Sub
FileWriteErr:
    Err.Raise Err.Number, "CLini.Save", Err.Description
    '       MsgBox Err.Description, vbExclamation, App.ProductName

End Sub

Public Property Get File() As String

    File = m_iniFile

End Property

Private Function isUnicode(ByRef FileName As String) As Boolean
        '<EhHeader>
        On Error GoTo isUnicode_Err
        '</EhHeader>
        
        Dim fNUM As Integer
        Dim b As Byte
        
        isUnicode = False
        fNUM = FreeFile
        Open FileName For Binary Access Read Shared As fNUM
        Get #fNUM, , b
        If b <> 255 Then GoTo isUnicode_Err
        Get #fNUM, , b
        If b <> 254 Then GoTo isUnicode_Err
        isUnicode = True
        Close fNUM
            
        '<EhFooter>
        Exit Function

isUnicode_Err:
        On Error Resume Next
        Close fNUM
End Function
'使用open语句打开文件
Public Property Let File(ByRef Value As String)

    m_iniFile = Value
    On Error GoTo FileReadErr
    Dim fNUM As Long
    Dim bText() As Byte
    Dim sText As String
    Dim bUnicode As Boolean
    
    bUnicode = isUnicode(m_iniFile)
                
    fNUM = FreeFile
    Open m_iniFile For Binary Access Read As #fNUM
    If LOF(fNUM) < 1 Then
        Close fNUM
        GoTo FileReadErr
    End If
    
    If bUnicode Then
        ReDim bText(LOF(fNUM) - 2) As Byte
        Seek fNUM, 3
        Get #fNUM, , bText()
        sText = bText
    Else
        ReDim bText(LOF(fNUM)) As Byte
        Get #fNUM, , bText()
        sText = StrConv(bText, vbUnicode)
    End If
    
    Close #fNUM
    
    
    hIni.iniString = sText


    Exit Property
FileReadErr:
    
    'm_Text = ""
    Debug.Print Err.Description

End Property


''使用File System Obj 打开文件
'Public Property Let File(ByRef Value As String)
'
'    m_iniFile = Value
'    On Error GoTo FileReadErr
'    Dim fso As New FileSystemObject
'    Dim ts As TextStream
'    Dim strLine As String
'    Dim sName As String
'
'    If fso.FileExists(Value) = False Then GoTo FileReadErr
'
'    Set ts = fso.OpenTextFile(Value, ForReading, False, TristateUseDefault)
'
'    iSecCount = 0
'    Erase tSection
'
'    Do Until ts.AtEndOfStream
'        strLine = ts.ReadLine
'        sName = Trim$(strLine)
'        'strLine = Trim(strLine)
'        If Left$(sName, 1) = "[" And Right$(sName, 1) = "]" Then
'            iSecCount = iSecCount + 1
'            ReDim Preserve tSection(1 To iSecCount) As iniSection
'            sName = Mid$(sName, 2, Len(sName) - 2)
'            If sName = "" Then sName = "NoNameSection" & iSecCount
'            tSection(iSecCount).Name = sName ' Mid$(strLine, 2, Len(strLine) - 2)
'        Else
'
'            If iSecCount < 1 Then
'                iSecCount = iSecCount + 1
'                ReDim Preserve tSection(1 To iSecCount) As iniSection
'                tSection(iSecCount).Name = "DefaultSection-NoName"
'                tSection(iSecCount).Text = strLine
'            Else
'                tSection(iSecCount).Text = tSection(iSecCount).Text & strLine & vbCrLf
'            End If
'
'        End If
'
'    Loop
'
'    ts.Close
'    Set ts = Nothing
'    Set fso = Nothing
'
'
'    '            Open m_iniFile For Binary Access Read As #fNum
'    '            mText = String$(LOF(fNum), " ")
'    '            Get #fNum, , mText
'    '            Close #fNum
'    '            'mText = Trim(mText)
'    '            m_Text = mText
'    'If m_Text.StartsWith(vbCrLf) = False Then m_Text = vbCrLf & m_Text.Value
'    Exit Property
'FileReadErr:
'    'm_Text = ""
'    Err.Raise Err.Number, "CLini.File", Err.Description
'
'End Property
Private Sub Class_Initialize()

    Set hIni = New CIniText
    CompareMethod = vbTextCompare
    'Set m_Text = New CString
    'm_Text.CompareMethod = m_CompareMethod

End Sub

'Private Sub Class_Terminate()
'
'    'Set m_Text = Nothing
'
'End Sub

Public Function GetArrayText(ByRef vSection As String, ByRef vKey As String, ByRef vText() As String) As Long
    On Error Resume Next
    Dim c As Long
    c = CLng(GetSetting(vSection, vKey & "ArrayCount"))
    If (c > 0) Then
        ReDim vText(0 To c - 1)
        Dim i As Long
        For i = 1 To c
            vText(i - 1) = GetSetting(vSection, vKey & "Item" & CStr(i))
        Next
    End If
    GetArrayText = c
End Function
Public Sub SaveArrayText(ByRef vSection As String, ByRef vKey As String, ByRef vText() As String)
    On Error Resume Next
    Dim c As Long
    c = UBound(vText) + 1
    SaveSetting vSection, vKey & "ArrayCount", CStr(c)
    'c = CLng(GetSetting(vSection, vKey & "ArrayCount"))
    If (c > 0) Then
        Dim i As Long
        For i = 1 To c
            SaveSetting vSection, vKey & "Item" & CStr(i), vText(i - 1)
        Next
    End If
End Sub
Public Sub LoadToList(ByRef AddItemTo As Object, Optional ByRef vSection As String, Optional ByRef vKey As String)
    On Error Resume Next
    If AddItemTo Is Nothing Then Exit Sub
    Dim text() As String
    If vSection = "" Then vSection = AddItemTo.name
    If vSection = "" Then vSection = "ObjectList"
    Dim c As Long
    c = GetArrayText(vSection, vKey, text)
    AddItemTo.Clear
    AddItemTo.Reset
    Dim i As Long
    For i = 0 To c - 1
        AddItemTo.AddItem text(i)
    Next
    AddItemTo.ListIndex = 0
    
End Sub

Public Sub SaveFromList(ByRef AddItemTo As Object, Optional ByRef vSection As String, Optional ByRef vKey As String)
    On Error Resume Next
    If AddItemTo Is Nothing Then Exit Sub
    Dim text() As String
    If vSection = "" Then vSection = AddItemTo.name
    If vSection = "" Then vSection = "ObjectList"
    Dim c As Long
    c = AddItemTo.ListCount
    If (c > 0) Then
        ReDim text(0 To c - 1) As String
        Dim i As Long
        For i = 0 To c - 1
            text(i) = AddItemTo.List(i)
        Next
        SaveArrayText vSection, vKey, text
    End If
End Sub

              

