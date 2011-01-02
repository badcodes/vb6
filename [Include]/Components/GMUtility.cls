VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GUtility"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Const EMPTY_ARRAY_UBOUND As Long = "-2"
Private Const EMPTY_ARRAY_LBOUND As Long = "-1"


Public Function IsArrayEmpty(arrTarget As Variant) As Boolean
    IsArrayEmpty = True
    If Not IsArray(arrTarget) Then Exit Function
    
End Function

Public Function SafeUbound(arrTarget As Variant) As Long
    SafeUbound = EMPTY_ARRAY_UBOUND
    On Error GoTo Exit_ME
    SafeUbound = UBound(arrTarget)
    Exit Function
Exit_ME:
    Err.Clear
    Exit Function
End Function

Public Function SafeLbound(arrTarget As Variant) As Long
    SafeLbound = EMPTY_ARRAY_LBOUND
    On Error GoTo Exit_ME
    SafeLbound = LBound(arrTarget)
    Exit Function
Exit_ME:
    Err.Clear
    Exit Function
End Function

Public Function GetTempName(Optional sPrefix As String = "XRTmp", Optional sSuffix As String = ".tmp") As String
    Static idx As Long
    If idx = 0 Then idx = &HAA9
    Dim fdTemp As String
    fdTemp = Environ$("TEMP") & "\"
    Dim fnTemp As String
    Do
        idx = idx + 1
        fnTemp = fdTemp & sPrefix & Hex(idx) & sSuffix
    Loop While IsFileExists(fnTemp)
    GetTempName = fnTemp
End Function

'CSEH: ErrorAbort
Public Function IsFileExists(ByRef sFilename As String) As Boolean
    '<EhHeader>
    On Error GoTo IsFileExists_Abort

    '</EhHeader>
    
    FileLen sFilename
    IsFileExists = True
    
    '<EhFooter>
    Exit Function

IsFileExists_Abort:
    Debug.Print Err.Number; ":" & Err.Description
    '</EhFooter>
End Function

'CSEH: ErrorAbort
Public Function IsFolderExists(ByRef sFilename As String) As Boolean
    '<EhHeader>
    On Error GoTo IsFolderExists_Abort

    '</EhHeader>
    If GetAttr(sFilename) And vbDirectory Then IsFolderExists = True
    '<EhFooter>
    Exit Function

IsFolderExists_Abort:
    Debug.Print Err.Number; ":" & Err.Description
    '</EhFooter>
End Function

'CSEH: ErrorAbort
Public Function FileToText(ByRef sFilename As String) As String
    '<EhHeader>
    On Error GoTo FileToText_Abort

    '</EhHeader>
    Dim n As Integer
    n = FreeFile
    Open sFilename For Input As #n
    FileToText = Input(LOF(n), n)
    Close #n
    '<EhFooter>
    Exit Function

FileToText_Abort:
    On Error Resume Next
    Dim nErr As Long
    nErr = Err.Number
    Debug.Print Err.Number; ":" & Err.Description
    Close #n
    Err.Raise nErr
    '</EhFooter>
End Function

'CSEH: ErrorAbort
Public Sub TextToFile(ByRef sFilename As String, sText As String)
    '<EhHeader>
    On Error GoTo TextToFile_Abort

    '</EhHeader>
    Dim n As Integer
    n = FreeFile
    Open sFilename For Output As #n
    Print #n, sText;
    Close #n
    '<EhFooter>
    Exit Sub

TextToFile_Abort:
    On Error Resume Next
    Dim nErr As Long
    nErr = Err.Number
    Debug.Print Err.Number; ":" & Err.Description
    Close #n
    Err.Raise nErr
    '</EhFooter>
End Sub

