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
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'--------------------------------------------------------------------------------
'    Component  : CLiNInI
'    Project    : SMDH
'
'    Description: ini File Class
'
'    Author   : xrLin
'
'    Date     : 2005-2006
'--------------------------------------------------------------------------------
Option Explicit

    Private m_iniFile As String
    Private hIni As CiniText
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

Public Sub Save()

    Dim fNum As Long
    'Dim l As Long
    Dim bUnicode As Boolean
    Dim sText As String
    
    bUnicode = True 'isUnicode(m_iniFile)
    sText = Me.ToString
    
    fNum = FreeFile
    On Error GoTo FileWriteErr:
    
    If bUnicode Then
        Dim c_B(1) As Byte
        ReDim bText(LenB(sText)) As Byte
        c_B(0) = 255
        c_B(1) = 254
        bText = sText
        Open m_iniFile For Binary Access Write As #fNum
        Put #fNum, , c_B()
        Put #fNum, , bText
    Else
        Open m_iniFile For Output As #fNum
        Print #fNum, ToString;
    End If
    
    Close #fNum
    
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
        
        Dim fNum As Integer
        Dim b As Byte
        
        isUnicode = False
        fNum = FreeFile
        Open FileName For Binary Access Read Shared As fNum
        Get #fNum, , b
        If b <> 255 Then GoTo isUnicode_Err
        Get #fNum, , b
        If b <> 254 Then GoTo isUnicode_Err
        isUnicode = True
        Close fNum
            
        '<EhFooter>
        Exit Function

isUnicode_Err:
        On Error Resume Next
        Close fNum
End Function
'使用open语句打开文件
Public Property Let File(ByRef Value As String)

    m_iniFile = Value
    On Error GoTo FileReadErr
    Dim fNum As Long
    Dim bText() As Byte
    Dim sText As String
    Dim bUnicode As Boolean
    
    bUnicode = isUnicode(m_iniFile)
                
    fNum = FreeFile
    Open m_iniFile For Binary Access Read As #fNum
    If LOF(fNum) < 1 Then
        Close fNum
        GoTo FileReadErr
    End If
    
    If bUnicode Then
        ReDim bText(LOF(fNum) - 2) As Byte
        Seek fNum, 3
        Get #fNum, , bText()
        sText = bText
    Else
        ReDim bText(LOF(fNum)) As Byte
        Get #fNum, , bText()
        sText = StrConv(bText, vbUnicode)
    End If
    
    Close #fNum
    
    
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

    Set hIni = New CiniText
    CompareMethod = vbTextCompare
    'Set m_Text = New CString
    'm_Text.CompareMethod = m_CompareMethod

End Sub

'Private Sub Class_Terminate()
'
'    'Set m_Text = Nothing
'
'End Sub

