VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GVBSource"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Enum EVBSourceFileType
    evbsftStandard
    evbsftClass
    evbsftUnknown
End Enum

Enum EVBSourceModuleType
    evbsmtStandard
    evbsmtClassPublic
    evbsmtClassGlobal
    evbsmtClassPrivate
    evbsmtInvalid
End Enum

Public Type TVBSourceModuleInfo
'    Fullname As String
'    Basename As String
    FileText As String
    ModuleName As String
    ModuleType As EVBSourceModuleType
    FileType As EVBSourceFileType
End Type

Const cst_VBSource_TargetName As String = "Attribute VB_Name = """
Const cst_VBSource_TargetPublic As String = "VB_Exposed = "
Const cst_VBSource_TargetGlobal As String = "Attribute VB_GlobalNameSpace = "
Const cst_VBSource_Quote2 As String = """"
'

Public Function VBSource_ExistFile(sSpec As String) As Boolean
    On Error Resume Next
    Call FileLen(sSpec)
    VBSource_ExistFile = (Err = 0)
End Function
Public Function VBSource_GetFileText(sFileName As String) As String
    Dim nFile As Integer, sText As String
    nFile = FreeFile
    'Open sFileName For Input As nFile ' Don't do this!!!
    If Not VBSource_ExistFile(sFileName) Then Exit Function
    ' Let others read but not write
    Open sFileName For Binary Access Read Lock Write As nFile
    ' sText = Input$(LOF(nFile), nFile) ! Don't do this!!!
    ' This is much faster
    sText = String$(LOF(nFile), 0)
    Get nFile, 1, sText
    Close nFile
    VBSource_GetFileText = sText
End Function

Public Function VBSource_GetModuleInfo(ByVal vFilename As String) As TVBSourceModuleInfo
Dim s As String, iStart As Long, iEnd As Long, sTmp As String
    On Error GoTo VBSource_GetModuleInfo_Error
    ' Get text of file regardless
    s = VBSource_GetFileText(vFilename)
     VBSource_GetModuleInfo.FileText = s
    ' Find module name
    iStart = InStr(s, cst_VBSource_TargetName)
    
    If iStart = 0 Then GoTo VBSource_GetModuleInfo_Error
    
    iStart = iStart + Len(cst_VBSource_TargetName)
    iEnd = InStr(iStart, s, cst_VBSource_Quote2)
    If iEnd = 0 Then GoTo VBSource_GetModuleInfo_Error
    
  
    
    VBSource_GetModuleInfo.ModuleName = Mid$(s, iStart, iEnd - iStart)
    ' Find module type
    
    Dim pExt As String
    pExt = UCase$(Right$(vFilename, 4))
    
    If pExt = ".BAS" Then
        VBSource_GetModuleInfo.FileType = evbsftStandard
        VBSource_GetModuleInfo.ModuleType = evbsmtStandard
    ElseIf pExt = ".CLS" Then
        VBSource_GetModuleInfo.FileType = evbsftClass
        ' Find public attribute
        iStart = InStr(s, cst_VBSource_TargetPublic)
        If iStart = 0 Then GoTo VBSource_GetModuleInfo_Error
        iStart = iStart + Len(cst_VBSource_TargetPublic)
        sTmp = Mid$(s, iStart, 1)
        Select Case sTmp
        Case "F"
            VBSource_GetModuleInfo.ModuleType = evbsmtClassPrivate
        Case "T"
            ' Find global attribute
            iStart = InStr(s, cst_VBSource_TargetGlobal)
            If iStart = 0 Then GoTo VBSource_GetModuleInfo_Error
            iStart = iStart + Len(cst_VBSource_TargetGlobal)
            sTmp = Mid$(s, iStart, 1)
            Select Case sTmp
            Case "F"
                VBSource_GetModuleInfo.ModuleType = evbsmtClassPublic
            Case "T"
                VBSource_GetModuleInfo.ModuleType = evbsmtClassGlobal
            Case Else
                GoTo VBSource_GetModuleInfo_Error
            End Select
        Case Else
            GoTo VBSource_GetModuleInfo_Error
        End Select
    Else
        VBSource_GetModuleInfo.FileType = evbsftUnknown
    End If
    'GetModuleInfo = s
    Exit Function
    
VBSource_GetModuleInfo_Error:
    ' Any number of reasons why module might be invalid
    VBSource_GetModuleInfo.ModuleType = evbsmtInvalid
End Function


