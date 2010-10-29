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


Private Type KeyValuePair
    Key As String
    Value As String
End Type

Private Type Section
    name As String
    KeyValues() As KeyValuePair
    Count As Long
    Size As Long
End Type

Private Type IniStructure
    Sections() As Section
    Count As Long
    Size As Long
End Type


Private Const CST_INISTRUCTURE_INIT_SIZE As Long = 16
Private Const CST_SECTION_INIT_SIZE As Long = 32
Private Const CST_INVALID_ARRAY_LBOUND As Long = -32765
Private Const CST_INVALID_ARRAY_UBOUND As Long = CST_INVALID_ARRAY_LBOUND - 1

Private Const cstEscapeCRLF As String = vbBack & "CRLF" & vbBack
Private Const CST_NAMELESS_SECTION_NAME As String = "NAMELESS_SECTION"

Private Const CST_COMMENT_PREFIX As String = ";"
Private Const CST_COMMENT_PREFIX_LEN As Integer = 1


Private Const CST_KEYVALUE_SEPARATOR As String = "="
Private Const CST_KEYVALUE_SEPARATOR_LEN As Integer = 1
Private Const CST_KEYVALUE_SEPARATOR_ESCAPED As String = vbNullChar & vbBack & "EQUAL" & vbBack & vbNullChar

Private Const CST_SECTION_PREFIX As String = "["
Private Const CST_SECTION_PREFIX_LEN As Integer = 1
Private Const CST_SECTION_SUFFIX As String = "]"
Private Const CST_SECTION_SUFFIX_LEN As Integer = 1
Private Const CST_DEFAULT_COMPARE_METHOD As Integer = vbTextCompare

Private mCompareMethod As VbCompareMethod

Private mIniStruct As IniStructure
Private mSource As String
Private mUnicodeOutput As Boolean

Public Property Get UnicodeOutput() As Boolean
    UnicodeOutput = mUnicodeOutput
End Property

Public Property Let UnicodeOutput(ByVal bValue As Boolean)
    mUnicodeOutput = bValue
End Property

Private Function ConvertKey(ByVal vKey As String, Optional vEscaped As Boolean = False) As String
    If vEscaped Then
        ConvertKey = Replace$(vKey, CST_KEYVALUE_SEPARATOR_ESCAPED, CST_KEYVALUE_SEPARATOR)
    Else
        ConvertKey = Replace$(vKey, CST_KEYVALUE_SEPARATOR, CST_KEYVALUE_SEPARATOR_ESCAPED)
    End If
End Function

Private Function ConvertValue(ByVal vValue As String, Optional vEscaped As Boolean = False) As String
    If vEscaped Then
        ConvertValue = Replace$(vValue, cstEscapeCRLF, vbCrLf)
    Else
        ConvertValue = Replace$(vValue, vbCrLf, cstEscapeCRLF)
    End If
End Function

Public Property Get Source() As String
    Source = mSource
End Property

Public Property Let Source(ByVal StrValue As String)

    If FileExists(StrValue) Then
        mIniStruct = ParseFile(StrValue)
        mSource = StrValue
        Exit Property
    End If

    If InStr(StrValue, vbLf) > 0 Then
        mIniStruct = ParseText(StrValue)
    Else
        mSource = StrValue
    End If
    
End Property

Public Property Get File() As String
    File = mSource
End Property

Public Property Let File(ByRef vFilename As String)
    mIniStruct = ParseFile(vFilename)
    mSource = vFilename
End Property

Public Property Get CompareMethod() As VbCompareMethod
    CompareMethod = mCompareMethod
End Property

Public Property Let CompareMethod(ByVal VBCValue As VbCompareMethod)
    mCompareMethod = VBCValue
End Property

Private Function NewIniStructure(Optional vSize As Long = -1) As IniStructure
    If vSize < 1 Then vSize = CST_INISTRUCTURE_INIT_SIZE
    ReDim Sections(1 To vSize) As Section
    NewIniStructure.Sections = Sections()
    NewIniStructure.Count = 0
    NewIniStructure.Size = vSize
End Function

Private Function NewSection(Optional vSize As Long = -1) As Section
    If vSize < 1 Then vSize = CST_SECTION_INIT_SIZE
    ReDim KeyValues(1 To vSize) As KeyValuePair
    NewSection.KeyValues = KeyValues()
    NewSection.Count = 0
    NewSection.Size = vSize
End Function

Private Function NewKeyValue(Optional vKey As String, Optional vValue As String) As KeyValuePair
    NewKeyValue.Key = vKey
    NewKeyValue.Value = vValue
End Function
'CSEH: ErrExit
Private Function ArrayBound(ByRef vArray() As String, Optional GetUbound As Boolean = False) As Long
    '<EhHeader>
    On Error GoTo ArrayBound_Err
    '</EhHeader>
    If GetUbound Then
        ArrayBound = UBound(vArray())
    Else
        ArrayBound = LBound(vArray())
    End If
    '<EhFooter>
    Exit Function

ArrayBound_Err:
    If GetUbound Then
        ArrayBound = CST_INVALID_ARRAY_UBOUND
    Else
        ArrayBound = CST_INVALID_ARRAY_LBOUND
    End If
    Err.Clear
    
    '</EhFooter>
End Function

Private Function ParseText(ByVal vText As String) As IniStructure
    Dim vLines() As String
    If InStr(vText, vbCrLf) >= 0 Then
        vLines = Split(vText, vbCrLf)
    Else
        vLines = Split(vText, vbLf)
    End If
    ParseText = ParseLines(vLines)
End Function

Private Function BuildSection(ByRef vText As String) As String
    If vText = "" Then Exit Function
    BuildSection = CST_SECTION_PREFIX & vText & CST_SECTION_SUFFIX
End Function
Private Function ExtractSection(ByRef vText As String) As String
    On Error Resume Next
    ExtractSection = Mid$(vText, CST_SECTION_PREFIX_LEN + 1, Len(vText) - CST_SECTION_PREFIX_LEN - CST_SECTION_SUFFIX_LEN)
End Function

Private Function ExtractKeyValue(ByRef vText As String, Optional vMidPos As Long = -1) As KeyValuePair
    If vMidPos < 0 Then vMidPos = InStr(vText, CST_KEYVALUE_SEPARATOR)
    If vMidPos < 0 Then Exit Function
    
    ExtractKeyValue.Key = ConvertKey(RTrim$(Mid$(vText, 1, vMidPos - 1)), True)
    ExtractKeyValue.Value = ConvertValue(LTrim$(Mid$(vText, vMidPos + CST_KEYVALUE_SEPARATOR_LEN)), True)
    
End Function

Private Function BuildKeyValue(ByRef vPair As KeyValuePair) As String
    If vPair.Key = "" Then
        BuildKeyValue = ConvertValue(vPair.Value)
    Else
        BuildKeyValue = ConvertKey(vPair.Key, False) & " " & CST_KEYVALUE_SEPARATOR & " " & ConvertValue(vPair.Value, False)
    End If
End Function
Private Function SearchSection(ByRef vIniStruct As IniStructure, ByRef vName As String) As Long
    SearchSection = -1
    If vName = "" Then Exit Function
    Dim i As Long
    For i = 1 To vIniStruct.Count
        If StrComp(vIniStruct.Sections(i).name, vName, mCompareMethod) = 0 Then
            SearchSection = i
            Exit Function
        End If
    Next

End Function

Private Function SearchKey(ByRef vSection As Section, ByVal vKey As String) As Long
    SearchKey = -1
    If vKey = "" Then Exit Function
    Dim i As Long
    For i = 1 To vSection.Count
        If StrComp(vSection.KeyValues(i).Key, vKey, mCompareMethod) = 0 Then
            SearchKey = i
            Exit Function
        End If
    Next
End Function
Private Function AddSection(ByRef vIniStruct As IniStructure, ByRef vName As String) As Long
    Dim idx As Long
    idx = SearchSection(vIniStruct, vName)
    If idx > 0 Then
        AddSection = idx
        Exit Function
    End If

    idx = vIniStruct.Count + 1
    If idx > vIniStruct.Size Then
        vIniStruct.Size = vIniStruct.Size + CST_INISTRUCTURE_INIT_SIZE
        ReDim Preserve vIniStruct.Sections(1 To vIniStruct.Size)
        'vIniStruct.Size = idx
    End If
    vIniStruct.Count = vIniStruct.Count + 1
    vIniStruct.Sections(idx) = NewSection()
    vIniStruct.Sections(idx).name = vName
    AddSection = idx
End Function

Private Function AddKeyValuePair(ByRef vSection As Section, vPair As KeyValuePair) As Long
    AddKeyValuePair = AddKeyValue(vSection, vPair.Key, vPair.Value)
End Function

Private Function AddKeyValue(ByRef vSection As Section, ByVal vKey As String, Optional ByVal vValue As String) As Long
    Dim idx As Long
    idx = SearchKey(vSection, vKey)
    If idx > 0 Then
        vSection.KeyValues(idx).Value = vValue
        AddKeyValue = idx
        Exit Function
    End If
    
    idx = vSection.Count + 1
    If idx > vSection.Size Then
        vSection.Size = vSection.Size + CST_INISTRUCTURE_INIT_SIZE
        ReDim Preserve vSection.KeyValues(1 To vSection.Size)
    End If
    vSection.KeyValues(idx).Key = vKey
    vSection.KeyValues(idx).Value = vValue
    vSection.Count = vSection.Count + 1
    AddKeyValue = idx
End Function
Private Function ParseLines(ByRef vLines() As String, Optional vStart As Long = CST_INVALID_ARRAY_LBOUND, Optional vSize As Long = -1) As IniStructure

    If vStart = CST_INVALID_ARRAY_LBOUND Then vStart = ArrayBound(vLines(), False)
    If vStart = CST_INVALID_ARRAY_LBOUND Then Exit Function
    If vSize = -1 Then vSize = ArrayBound(vLines(), True) - vStart + 1
    If vSize < 1 Then Exit Function

    'ParseLines = NewIniStructure()
    
    Dim i As Long
    Dim pEnd As Long
    Dim pSecIdx As Long
    Dim pText As String
    

    
    pEnd = vStart + vSize - 1
    For i = vStart To pEnd
        pText = Trim$(vLines(i))
        If pText = "" Then GoTo NextLine
        If Mid$(pText, 1, CST_COMMENT_PREFIX_LEN) = CST_COMMENT_PREFIX Then
            If pSecIdx < 1 Then pSecIdx = AddSection(ParseLines, "")
            'Debug.Print "Comment||"; pText; "||"
            AddKeyValue ParseLines.Sections(pSecIdx), "", vLines(i)
            GoTo NextLine
        End If
        
        If Left$(pText, CST_SECTION_PREFIX_LEN) = CST_SECTION_PREFIX And _
            Right$(pText, CST_SECTION_SUFFIX_LEN) = CST_SECTION_SUFFIX Then
                        'Debug.Print "Section||"; pText; "||"
                pSecIdx = AddSection(ParseLines, ExtractSection(pText))
                GoTo NextLine
        Else
            If pSecIdx < 1 Then pSecIdx = AddSection(ParseLines, "")
            
            Dim ppPos As Long
            ppPos = InStr(pText, CST_KEYVALUE_SEPARATOR)
            If ppPos > 0 Then
                'Debug.Print "KeyValue||"; pText; "||"
                AddKeyValuePair ParseLines.Sections(pSecIdx), ExtractKeyValue(pText, ppPos)
            Else
                'ebug.Print "PlainText||"; pText; "||"
                AddKeyValue ParseLines.Sections(pSecIdx), "", vLines(i)
            End If
        End If
        
NextLine:
    Next
    
End Function

Public Function ToString() As String
    Dim i As Long
    Dim j As Long
    For i = 1 To mIniStruct.Count
        If mIniStruct.Sections(i).name <> "" Then
            ToString = ToString & vbCrLf & BuildSection(mIniStruct.Sections(i).name)
        End If
        For j = 1 To mIniStruct.Sections(i).Count
            ToString = ToString & vbCrLf & BuildKeyValue(mIniStruct.Sections(i).KeyValues(j))
        Next
    Next
    ToString = Mid$(ToString, Len(vbCrLf) + 1)
End Function

Private Function FileExists(ByRef Filename As String) As Boolean
    On Error GoTo errorFileExists
    FileLen Filename
    FileExists = True
    Exit Function
errorFileExists:
    FileExists = False
    Err.Clear
End Function
Public Sub Save()
    If mSource = "" Then Exit Sub
    WriteTo mSource
End Sub
Public Sub WriteTo(ByRef vFilename As String)
    Dim fNum As Long
    Dim l As Long
    Dim bUnicode As Boolean
    Dim sText As String

    bUnicode = True 'isUnicode(m_iniFile)
    sText = Me.ToString

    fNum = FreeFile
    On Error GoTo FileWriteErr:
    If FileExists(vFilename) Then Kill vFilename
    If mUnicodeOutput Then
        Dim c_B(1) As Byte
        ReDim bText(LenB(sText)) As Byte
        c_B(0) = 255
        c_B(1) = 254
        bText = sText
        'Open fileName For Output As #fNum
        Open vFilename For Binary Access Write As #fNum
        Put #fNum, , c_B()
        Put #fNum, , bText
    Else
        Open vFilename For Output As #fNum
        Print #fNum, sText;
    End If

    Close #fNum

    Exit Sub
FileWriteErr:
    Err.Raise Err.Number, "CLiNIni.WriteTo", Err.Description
    On Error Resume Next
    Close #fNum
End Sub
'
Private Function OpenUnicodeFile(ByRef Filename As String, ByRef fNum As Integer) As Boolean
        '<EhHeader>
        On Error GoTo isUnicode_Err
        '</EhHeader>

        
        Dim b As Byte

        OpenUnicodeFile = False
        fNum = FreeFile
        
        Open Filename For Binary Access Read Shared As fNum
        If LOF(fNum) < 1 Then GoTo isUnicode_Err
        Get #fNum, , b
        If b <> 255 Then GoTo noUnicode
        Get #fNum, , b
        If b <> 254 Then GoTo noUnicode
        OpenUnicodeFile = True
        

        '<EhFooter>
        Exit Function

noUnicode:
    Seek fNum, 1
    Exit Function

isUnicode_Err:
        On Error Resume Next
        Close fNum
        fNum = -1
End Function
''使用open语句打开文件
Private Function ParseFile(ByRef Value As String) As IniStructure

    On Error GoTo FileReadErr
    Dim fNum As Integer
    Dim bText() As Byte
    Dim sText As String
    Dim bUnicode As Boolean

    bUnicode = OpenUnicodeFile(Value, fNum)
    If (fNum < 1) Then Exit Function
    ReDim bText(LOF(fNum) - Loc(fNum))
    Get fNum, , bText()
    Close #fNum
    
    If bUnicode Then
        ParseFile = ParseText(bText)
    Else
        ParseFile = ParseText(StrConv(bText, vbUnicode))
    End If
    

    Exit Function
FileReadErr:

    'm_Text = ""
    Debug.Print Err.Description
    On Error Resume Next
    Close #fNum
End Function

Public Function GetArrayText(ByVal vSection As String, ByVal vKey As String, ByRef vText() As String) As Long
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
Public Sub SaveArrayText(ByVal vSection As String, ByVal vKey As String, ByRef vText() As String)
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

'
Public Sub Reset()
    Dim tmp As IniStructure
    mIniStruct = tmp
End Sub

Public Sub DeleteSection(ByVal vSection As String)

    Dim iPoint As Long
    Dim iDelPoint As Long
    Dim iEndPoint As Long

    'Section = Trim(Section)

    iDelPoint = SearchSection(mIniStruct, vSection)
    If iDelPoint < 1 Then Exit Sub
    iEndPoint = mIniStruct.Count - 1

    For iPoint = iDelPoint To iEndPoint
        mIniStruct.Sections(iPoint) = mIniStruct.Sections(iPoint + 1)
    Next
    
    mIniStruct.Sections(iEndPoint + 1) = NewSection()
    mIniStruct.Count = mIniStruct.Count - 1
   

End Sub
Public Sub DeleteSetting(ByVal vSection As String, ByVal vKey As String)
    Dim iSec As Long
    iSec = SearchSection(mIniStruct, vSection)
    If iSec < 1 Then Exit Sub
    
    Dim iKey As Long
    iKey = SearchKey(mIniStruct.Sections(iSec), vKey)
    If iKey < 1 Then Exit Sub
    
    Dim iEnd As Long
    iEnd = mIniStruct.Sections(iSec).Count - 1
    
    Dim i As Long
    
    For i = iKey To iEnd
        mIniStruct.Sections(iSec).KeyValues(i) = mIniStruct.Sections(iSec).KeyValues(i + 1)
    Next
    
    mIniStruct.Sections(iSec).Count = iEnd

End Sub
Public Sub SaveSetting(ByVal vSection As String, ByVal vKey As String, ByVal vValue As String)
    Dim iSec As Long
    iSec = AddSection(mIniStruct, vSection)
    AddKeyValue mIniStruct.Sections(iSec), vKey, vValue
End Sub

Public Function GetSetting(ByVal vSection As String, ByVal vKey As String) As String
    Dim iSec As Long
    iSec = SearchSection(mIniStruct, vSection)
    If iSec < 1 Then Exit Function
    
    Dim iKey As Long
    iKey = SearchKey(mIniStruct.Sections(iSec), vKey)
    If iKey < 1 Then Exit Function
    
    GetSetting = mIniStruct.Sections(iSec).KeyValues(iKey).Value
End Function
Public Sub ChangeSectionName(ByVal vSection As String, ByVal vNewSection As String)
    Dim iSec As Long
    iSec = SearchSection(mIniStruct, vSection)
    If iSec < 1 Then Exit Sub
    mIniStruct.Sections(iSec).name = vNewSection
    
End Sub
Public Sub ChangeKeyName(ByVal vSection As String, ByVal vKey As String, ByVal vNewKey As String)
    Dim iSec As Long
    iSec = SearchSection(mIniStruct, vSection)
    If iSec < 1 Then Exit Sub
    
    Dim iKey As Long
    iKey = SearchKey(mIniStruct.Sections(iSec), vKey)
    If iKey < 1 Then Exit Sub
    
    mIniStruct.Sections(iSec).KeyValues(iKey).Key = vNewKey
End Sub

Public Function GetSection(ByVal vSection As String, ByRef vKeyValue() As String) As Long
    Dim iSec As Long
    iSec = SearchSection(mIniStruct, vSection)
    If iSec < 1 Then Exit Function
    ReDim vKeys(0 To mIniStruct.Sections(iSec).Count - 1, 0 To 1)
    Dim i As Long
    For i = 1 To mIniStruct.Sections(iSec).Count
        vKeys(i - 1, 0) = mIniStruct.Sections(iSec).KeyValues(i).Key
        vKeys(i - 1, 1) = mIniStruct.Sections(iSec).KeyValues(i).Value
    Next
    GetSection = mIniStruct.Sections(iSec).Count
End Function

Public Sub SaveSection(ByVal vSection As String, ByRef vKeyValue() As String)

    Dim iSec As Long
    iSec = AddSection(mIniStruct, vSection)
    Dim i As Long
    Dim iStart As Long
    Dim iEnd As Long
    iStart = ArrayBound(vKeyValue(), False)
    iEnd = ArrayBound(vKeyValue(), True)
    For i = iStart To iEnd
        AddKeyValue mIniStruct.Sections(iSec), vKeyValue(i, 0), vKeyValue(i, 1)
    Next
    
End Sub

Private Sub Class_Initialize()
    mCompareMethod = CST_DEFAULT_COMPARE_METHOD
    mUnicodeOutput = True
End Sub


Public Function SectionExists(ByVal vSection As String) As Boolean
    SectionExists = SearchSection(mIniStruct, vSection) > 0
End Function

Public Function KeyExists(ByVal vSection As String, ByVal vKey As String) As Boolean
    Dim iSec As Long
    iSec = SearchSection(mIniStruct, vSection)
    If iSec > 0 Then
        KeyExists = SearchKey(mIniStruct.Sections(iSec), vKey) > 0
    End If
End Function


