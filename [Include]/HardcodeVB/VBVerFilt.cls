VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CVBVerFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' CVBVerFilter implements IFilter and IUseFile
Implements IFilter
Implements IUseFile

Enum EFilterType
    eftMinFilter = 0
    eftProject = 0
    eftClass
    eftForm
    eftMaxFilter
End Enum

Private eftFilterType As EFilterType
Private fStop As Boolean

' Class lines
Const sPersistable = "  Persistable ="
Const sDataBind = "  DataBindingBehavior ="
Const sDataSource = "  DataSourceBehavior  ="
Const sMTSTransact = "  MTSTransactionMode  ="
Private cPersistable As Long
Private cDataBind As Long
Private cDataSource As Long
Private cMTSTransact As Long

' VBP lines

Const sCondition = "CondComp="
Private cCondition As Long
Const sVB6 = "iVBVer = 6"
Private cVB6 As Long
Const sRetained = "Retained="
Private cRetained As Long
Const sDebugStart = "DebugStartupOption="
Private cDebugStart As Long
Const sUseBrowser = "UsingExistingBrowser="
Private cUseBrowser As Long
Const sProjComCtl61 = "Object={831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0; mscomctl.ocx"
Private cProjComCtl61 As Long
Const sProjComCtl62 = "Object={86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0; MSCOMCT2.OCX"
Private cProjComCtl62 As Long
Const sProjComCtl51 = "Object={6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.1#0; COMCTL32.OCX"
Private cProjComCtl51 As Long
Const sProjComCtl52 = "Object={FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.0#0; COMCT232.OCX"
Private cProjComCtl52 As Long

' FRM lines

Const sFormComCtl61 = "Object = ""{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0""; ""mscomctl.ocx"""
Private cFormComCtl61 As Long
Const sFormComCtl62 = "Object = ""{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0""; ""MSCOMCT2.OCX"""
Private cFormComCtl62 As Long
Const sFormComCtl51 = "Object = ""{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.1#0""; ""COMCTL32.OCX"""
Private cFormComCtl51 As Long
Const sFormComCtl52 = """Object = ""{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.0#0""; ""COMCT232.OCX"""
Private cFormComCtl52 As Long

Const sComCtl6 = "MSComCtlLib."
Private cComCtl6 As Long
Const sComCtl5 = "ComctlLib."
Private cComCtl5 As Long

' Implementation of IFilter interface
Private sSource As String, sTarget As String

Public Event ConvertedFile(ByVal sFullPath As String)

Private Sub Class_Initialize()
    
    ' Initialize all lengths
    cPersistable = Len(sPersistable)
    cDataBind = Len(sDataBind)
    cDataSource = Len(sDataSource)
    cMTSTransact = Len(sMTSTransact)

    cCondition = Len(sCondition)
    cVB6 = Len(sVB6)

    cRetained = Len(sRetained)
    cDebugStart = Len(sDebugStart)
    cUseBrowser = Len(sUseBrowser)
    
    cProjComCtl61 = Len(sProjComCtl61)
    cProjComCtl62 = Len(sProjComCtl62)
    cProjComCtl51 = Len(sProjComCtl51)
    cProjComCtl52 = Len(sProjComCtl52)
    
    cFormComCtl61 = Len(sFormComCtl61)
    cFormComCtl62 = Len(sFormComCtl62)
    cFormComCtl51 = Len(sFormComCtl51)
    cFormComCtl52 = Len(sFormComCtl52)
    cComCtl6 = Len(sComCtl6)
End Sub

Private Property Get IFilter_Source() As String
    IFilter_Source = sSource
End Property
Private Property Let IFilter_Source(sSourceA As String)
    sSource = sSourceA
End Property

Private Property Get IFilter_Target() As String
    IFilter_Target = sTarget
End Property
Private Property Let IFilter_Target(sTargetA As String)
    sTarget = sTargetA
End Property

Private Function IFilter_Translate(sLine As String, _
                                   ByVal iLine As Long) As EChunkAction
    Select Case eftFilterType
    Case eftProject
        IFilter_Translate = TranslateProject(sLine)
    Case eftClass
        ' The illegal lines are only at the start of the file
        If iLine > 10 Then
            IFilter_Translate = ecaTranslateAll
        Else
            IFilter_Translate = TranslateClass(sLine)
        End If
    Case eftForm
        IFilter_Translate = TranslateForm(sLine)
    End Select
End Function

Property Get FilterType() As EFilterType
    FilterType = eftFilterType
End Property

Property Let FilterType(eftFilterTypeA As EFilterType)
    If eftFilterTypeA >= eftMinFilter And _
       eftFilterTypeA <= eftMaxFilter Then
        eftFilterType = eftFilterTypeA
    Else
        eftFilterType = eftMaxFilter
    End If
End Property

Sub StopWalk()
    fStop = True
End Sub

Sub ConvertFile(sFile As String)

    ' IFilter part of object
    Dim filter As IFilter
    Set filter = Me
    ' Set FilterType property
    Select Case UCase$(GetFileExt(sFile))
    Case ".VBP"
        FilterType = eftProject
    Case ".CLS"
        FilterType = eftClass
    Case ".FRM"
        FilterType = eftForm
    Case Else
        ' Ignore other files
        Exit Sub
    End Select
    ' Set Source property on filter variable
    filter.Source = sFile
    ' Pass filter to FilterTextFile
    FilterTextFile filter
    RaiseEvent ConvertedFile(sFile)

End Sub

Private Function TranslateProject(sLine As String) As EChunkAction
    ' Delete line starting with unrecognized options
    If Left$(sLine, cRetained) = sRetained Then
        TranslateProject = ecaSkip
    ElseIf Left$(sLine, cDebugStart) = sDebugStart Then
        TranslateProject = ecaSkip
    ElseIf Left$(sLine, cUseBrowser) = sUseBrowser Then
        TranslateProject = ecaSkip
    ' Change VB6 to VB5
    ElseIf Left$(sLine, cCondition) = sCondition Then
        Dim iPos As Long
        iPos = InStr(sLine, sVB6)
        If iPos Then
            Mid$(sLine, iPos + cVB6 - 1) = "5"
        End If
        TranslateProject = ecaTranslate
    ' Translate changed common control lines
    ElseIf StrComp(sLine, sProjComCtl61, vbTextCompare) = 0 Then
        sLine = sProjComCtl51
        TranslateProject = ecaTranslate
    ElseIf StrComp(sLine, sProjComCtl62, vbTextCompare) = 0 Then
        sLine = sProjComCtl52
        TranslateProject = ecaTranslate
    Else
        TranslateProject = ecaTranslate
    End If
End Function

Private Function TranslateClass(sLine As String) As EChunkAction
    ' Delete lines starting with any of the unknown strings
    If Left$(sLine, cPersistable) = sPersistable Then
        TranslateClass = ecaSkip
    ElseIf Left$(sLine, cDataBind) = sDataBind Then
        TranslateClass = ecaSkip
    ElseIf Left$(sLine, cDataSource) = sDataSource Then
        TranslateClass = ecaSkip
    ElseIf Left$(sLine, cMTSTransact) = sMTSTransact Then
        TranslateClass = ecaSkip
    Else
        TranslateClass = ecaTranslate
    End If
End Function

Private Function TranslateForm(sLine As String) As EChunkAction
    ' Translate changed common control lines
    If StrComp(sLine, sFormComCtl61, vbTextCompare) = 0 Then
        ' Change to new line
        sLine = sFormComCtl51
        TranslateForm = ecaTranslate
    ElseIf StrComp(sLine, sFormComCtl62, vbTextCompare) = 0 Then
        ' Change to new line
        sLine = sFormComCtl52
        TranslateForm = ecaTranslate
    Else
        Dim i As Long, s As String
        i = InStr(1, sLine, sComCtl6, vbTextCompare)
        If i Then
            ' Change to new line
            s = Left$(sLine, i - 1) & sComCtl5
            sLine = s & Mid$(sLine, i + cComCtl6)
        End If
        TranslateForm = ecaTranslate
    End If
End Function

Private Function IUseFile_UseFile(UserData As Variant, FilePath As String, _
                                  FileInfo As CFileInfo) As Boolean
With FileInfo
    Dim sFile As String
    ConvertFile .item
    DoEvents
    If fStop Then
        IUseFile_UseFile = True
        fStop = False
    End If
End With
End Function




