Attribute VB_Name = "MClassTemplate"
Private Const cstConstPrefix As String = "f"
Private Const cstTypePrefix As String = "TPL"
Private Const cstTypeSuffix As String = "Type"
Private Const cstTypeDef As String = "    ToString As String"
Private Const cstTypeIdStart As Integer = 64
Private Const cstMaxType As Integer = 26
Private Const cstDefaultType As String = "DefaultType"


    
Public Sub ValidCount(ByRef nCount As Integer)
    If nCount > cstMaxType Then nCount = cstMaxType
    If nCount < 1 Then nCount = 1
End Sub
Public Sub GenerateConstType()

PrintLine "#Const ObjectType = 1"
PrintLine "#Const NormalType = 2"
PrintLine "#Const VariantType = (ObjectType Or NormalType)"
PrintLine "#Const DefaultType = VariantType"
    
End Sub

Public Function GetTypeName(ByVal idx As Integer)
   GetTypeName = cstTypePrefix & Chr$(cstTypeIdStart + idx) & cstTypeSuffix
End Function
Public Sub GenerateConstDefault(Optional ByVal nCount As Integer = cstMaxType)
    ValidCount nCount
    For i = 1 To nCount
        PrintLine "#Const " & cstConstPrefix & GetTypeName(i) & " = " & cstDefaultType
    Next
End Sub

Public Sub GenerateTemplateHeader(Optional ByVal nCount As Integer = cstMaxType)

    PrintLine "'Template header:"
    PrintLine "'" & String$(80, "=")
    GenerateConstType
    PrintLine ""
    GenerateConstDefault nCount
    PrintLine "'" & String$(80, "=")

End Sub

Public Sub GenerateTemplateType(Optional ByVal nCount As Integer = cstMaxType)
    
    ValidCount nCount
    PrintLine "'Template type definition:"
    PrintLine "'" & String$(80, "=")
    For i = 1 To nCount
        PrintLine "Public Type " & GetTypeName(i)
        PrintLine cstTypeDef
        PrintLine "End Type"
    Next
    PrintLine "'" & String$(80, "=")

End Sub

Public Sub GenerateTemplate(Optional ByVal nCount As Integer = cstMaxType)
    GenerateTemplateHeader nCount
    PrintLine ""
    PrintLine ""
    GenerateTemplateType nCount
End Sub

Public Sub GenerateTemplateFunction()

End Sub

Public Sub PrintLine(ByRef sLine As String)
    Debug.Print sLine
End Sub


Public Sub CreateTemplateTypeClass(Optional ByVal nCount As Integer = cstMaxType)


    Dim sTemplateText As String

    Dim nFile As Integer
    nFile = FreeFile
    Open App.Path & "\CTemplateType.cls" For Input As #nFile
    sTemplateText = Input(LOF(nFile), nFile)
    Close #nFile

    
    Dim i As Integer
    Dim sTypeName As String
    For i = 1 To nCount
        sTypeName = GetTypeName(i)
        nFile = FreeFile
        Open App.Path & "\" & sTypeName & ".cls" For Binary As #nFile
        Put #nFile, , Replace$(sTemplateText, "CTemplateType", GetTypeName(i))
        Close #nFile
    Next

End Sub

Public Sub CreateTemplateTypeProject(Optional ByVal nCount As Integer = cstMaxType)
    ValidCount nCount
    Dim nFile As Integer
    Dim q As String
    q = Chr$(34)
    nFile = FreeFile
    Open App.Path & "\TemplateType.vbp" For Output As #nFile
    
Print #nFile, "Type=OleDll"
Print #nFile, "Reference=*\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\WINDOWS\system32\stdole2.tlb#OLE Automation"
    
    Dim i As Integer
    For i = 1 To nCount
    Print #nFile, "Class=" & GetTypeName(i) & "; " & GetTypeName(i) & ".cls"
    Next
Print #nFile, "Startup=" & q & "(None)" & q
Print #nFile, "Command32=" & q & q
Print #nFile, "Name=" & q & "TemplateType" & q
Print #nFile, "HelpContextID=" & q & "0" & q
Print #nFile, "CompatibleMode=" & q & "1" & q
Print #nFile, "MajorVer=1"
Print #nFile, "MinorVer=0"
Print #nFile, "RevisionVer=0"
Print #nFile, "AutoIncrementVer=0"
Print #nFile, "ServerSupportFiles=0"
Print #nFile, "VersionCompanyName=" & q & "MYPLACE" & q
Print #nFile, "CompilationType=0"
Print #nFile, "OptimizationType=0"
Print #nFile, "FavorPentiumPro(tm)=0"
Print #nFile, "CodeViewDebugInfo=0"
Print #nFile, "NoAliasing=0"
Print #nFile, "BoundsCheck=0"
Print #nFile, "OverflowCheck=0"
Print #nFile, "FlPointCheck=0"
Print #nFile, "FDIVCheck=0"
Print #nFile, "UnroundedFP=0"
Print #nFile, "StartMode=1"
Print #nFile, "Unattended=0"
Print #nFile, "Retained=0"
Print #nFile, "ThreadPerObject=0"
Print #nFile, "MaxNumberOfThreads=1"
Print #nFile, "ThreadingModel=1"
Print #nFile, ""
Print #nFile, "[MS Transaction Server]"
Print #nFile, "AutoRefresh=1"

    Close #nFile
End Sub

Public Sub GenerateNewTypeFunction(Optional ByVal nCount As Integer = cstMaxType)
    ValidCount nCount
    Dim i As Integer
    Dim sTypeName As String
    For i = 1 To nCount
        sTypeName = GetTypeName(i)
        PrintLine "Public Function New" & sTypeName & "(ByVal sInit as String) as " & sTypeName
        PrintLine "    Set New" & sTypeName & " = New " & sTypeName
        PrintLine "    New" & sTypeName & ".toString = sInit"
        PrintLine "End Function"
    Next
End Sub
