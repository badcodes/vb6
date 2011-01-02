Attribute VB_Name = "MTest"
Option Explicit

Sub testStringMap()
    Dim a As CStringMap
    Set a = New CStringMap
    a("a") = "AAAAAAAAAAAAAAAAAAAAA"
    a("b") = "BBBBBBBBBBBBBBBBBBBBBB"
    a("c") = "CCCCCCCCCCCCCCCCCCCCCCCC"
    a.Remove "b"
    Debug.Print a("a")
    Debug.Print a("b")
    Debug.Print a("c")
End Sub

Sub TestStack()
    Dim a As CStack
    Set a = New CStack
    a.Push ("AAAAAAAAAAAAAAAAAAAAA")
    a.Push ("BBBBBBBBBBBBBBBBBBBBB")
    a.Push ("CCCCCCCCCCCCCCCCCCCCC")
    Debug.Print a.Pop
    a.Push ("DDDDDDDDDDDDDDDDDDDDD")
    Debug.Print a.Peek()
    Debug.Print a.Pop
    Debug.Print a.Count
End Sub
Sub TestStringArray()
    Dim a As CStringArray
    Set a = New CStringArray
    a.Add "AAAAAAAAAAAAAAAA"
    a.Add "BBBBBBBBBBBBBBBB"
    a.Shink
    a.Insert 13, "CCCCCCCCCCCCCCCCCCCC"
    a.Remove 12, -3
    a.Remove 2, 8
    Debug.Print a.Find("CCCCCCCCCCCCCCCCCCCC")
    Dim i As Long
    Dim aa() As String
    aa = a.ToArray
    For i = LBound(aa) To UBound(aa)
        Debug.Print aa(i)
    Next
End Sub


Public Sub TestRegExp(vSrc As String)

    Dim expType As RegExp
    Dim vDest As String
    Set expType = New RegExp
    expType.Global = True
    expType.IgnoreCase = True
    expType.MultiLine = False
    expType.Pattern = "\s+as\s+(\w+)"
    
    Dim regMatches As MatchCollection
    Dim regMatch As Match
    Set regMatches = expType.Execute(vSrc)
    
    Dim sTypeName As String
    Dim aTarget(1 To 26, 1 To 2) As String
    Dim cTarget As Long
    
    Dim nStart As Long
    Dim nStop As Long
    
    nStart = 1
    vDest = ""
    For Each regMatch In regMatches
        sTypeName = regMatch.SubMatches.Item(0)
        'If (mInfo.Exists(sTypeName)) Then
            nStop = regMatch.FirstIndex
            vDest = vDest & Mid$(vSrc, nStart, nStop - nStart + 1) & " AS " & "HAHAHAHA"
            nStart = nStop + regMatch.Length + 1
        'End If
    Next
    nStop = Len(vSrc)
    If nStop >= nStart Then
        vDest = vDest & Mid$(vSrc, nStart, nStop - nStart + 1)
    End If
    Debug.Print vDest
   
End Sub

Public Sub TestClassBuilder()
    Dim builder As CTemplateBuilder
    Set builder = New CTemplateBuilder
    
    
    Dim typeStyle As CTypeStyle
    Set typeStyle = New CTypeStyle
    With typeStyle
        .ConstVarOf(CTTypeNormal) = "NormalType"
        .ConstVarOf(CTTypeObject) = "ObjectType"
        .ConstVarOf(CTTypeVariant) = "VariantType"
    End With
       
    Dim templateType As CType
    Set templateType = New CType
    Set templateType.typeStyle = typeStyle
    
    templateType.ConstTypePrefix = "f"
    templateType.ConstTypeSuffix = ""

    
    builder.InitType templateType
    builder.AddFilter New CFilterModule
    builder.AddFilter New CFilterTypeName
    builder.AddFilter New CFilterConstVar
    builder.AddFilter New CFilterTypeOP
'    Dim tnFilter As ITemplateFilter
'    Set tnFilter = New CTypeNameFilter
'    builder.AddFilter tnFilter
'    Set tnFilter = New CModuleFilter
'    builder.AddFilter tnFilter
    
    builder.AddType "TArray", "CStringArray", CTTypeObject
    builder.AddType "TPLAType", "String", CTTypeNormal
    builder.Process "X:\Workspace\VisualBasic6\[Template]\ClassTemplate\TArray.cls", "X:\Workspace\VisualBasic6\[Template]\ClassTemplate\CStringArray.cls"
    
End Sub

Public Sub BuildClass(ByRef vSrcName As String, ByRef vDstName As String)
    
Const cst_ClassDirectory As String = "X:\Workspace\VB6\[Include]\ClassTemplate\"
Dim pSrc As String
pSrc = cst_ClassDirectory & vSrcName
Dim pDst As String
pDst = cst_ClassDirectory & vDstName
'If FileExists(pDst) Then
'    Kill pDst
'End If

    Dim builder As CTemplateBuilder
    Set builder = New CTemplateBuilder
    
    
    Dim typeStyle As CTypeStyle
    Set typeStyle = New CTypeStyle
    With typeStyle
        .ConstVarOf(CTTypeNormal) = "NormalType"
        .ConstVarOf(CTTypeObject) = "ObjectType"
        .ConstVarOf(CTTypeVariant) = "VariantType"
    End With
       
    Dim templateType As CType
    Set templateType = New CType
    Set templateType.typeStyle = typeStyle
    
    templateType.ConstTypePrefix = "f"
    templateType.ConstTypeSuffix = ""

    
    builder.InitType templateType
    builder.AddFilter New CFilterModule
    builder.AddFilter New CFilterConstVar
    builder.AddFilter New CFilterTypeName
    builder.AddFilter New CFilterTypeOP
'    Dim tnFilter As ITemplateFilter
'    Set tnFilter = New CTypeNameFilter
'    builder.AddFilter tnFilter
'    Set tnFilter = New CModuleFilter
'    builder.AddFilter tnFilter
    
    'builder.AddType "TArray", "CStringArray", CTTypeObject
    builder.AddType Left$(vSrcName, Len(vSrcName) - 4), _
                    "C" & Left$(vDstName, Len(vDstName) - 4), _
                    CTTypeObject
    If vTypeA <> "" Then
        builder.AddType "TPLAType", vTypeA, vTypeAType
    End If
    If vTypeB <> "" Then
        builder.AddType "TPLBType", vTypeB, vTypeBType
    End If
    builder.Process pSrc, pDst
    
End Sub
