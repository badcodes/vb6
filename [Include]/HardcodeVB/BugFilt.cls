VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBugFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' CBugFilter implements IFilter
Implements IFilter
Implements IUseFile

Enum EFilterType
    eftMinFilter = 0
    eftDisableBug = 0
    eftEnableBug
    eftDisableProfile
    eftEnableProfile
    eftExpandAsserts
    eftTrimAsserts
    eftMaxFilter
End Enum
Private eftFilterType As EFilterType

Const sBug = "Bug"
Const sProfile = "Profile"
Const sComment = "'"
Const sBugAssert = "BugAssert "

' Implementation of IFilter interface
Private sSource As String, sTarget As String

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
    IFilter_Translate = ecaTranslate  ' Always translate with this filter
    
    Select Case eftFilterType
    Case eftDisableBug
        CommentOut sLine, sBug
    Case eftEnableBug
        CommentIn sLine, sBug
    Case eftDisableProfile
        CommentOut sLine, sProfile
    Case eftEnableProfile
        CommentIn sLine, sProfile
    Case eftExpandAsserts
        ExpandAsserts sLine, iLine
    Case eftTrimAsserts
        TrimAsserts sLine
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

Private Sub CommentOut(sLine As String, sTarget As String)
    
    ' Check to see if line contains target
    Dim iPos As Integer
    iPos = InStr(sLine, sTarget)
    If iPos Then
    
        ' If text is first nonblank, comment it out
        Dim s As String
        s = Space$(iPos - 1)
        If Left$(sLine, iPos - 1) = s Then
            sLine = s & sComment & Mid$(sLine, iPos)
        End If
    End If
            
End Sub

Private Sub CommentIn(sLine As String, sTarget As String)
    
    ' Check to see if line contains string
    Dim iPos As Integer
    iPos = InStr(sLine, sComment & sTarget)
    If iPos Then
    
        ' If text is first nonblank, comment it in
        Dim s As String
        s = Space$(iPos - 1)
        If Left$(sLine, iPos - 1) = s Then
            sLine = s & Mid$(sLine, iPos + 1)
        End If
    End If
            
End Sub

Private Sub ExpandAsserts(sLine As String, iLine As Long)
    
    Dim iPos As Integer, i As Integer, sComment As String
    ' Check to see if line contains Bug string
    iPos = InStr(sLine, sBugAssert)
    If iPos Then
    
        ' Save comment so it won't be processed
        i = InStr(sLine, "'")
        If i Then
            ' Ignore commented out Bug strings
            If iPos > i Then Exit Sub
            ' Remove comment
            sComment = Mid$(sLine, i)
            sLine = Left$(sLine, i - 1)
        End If
        ' Move to first argument
        iPos = iPos + Len(sBugAssert)
        
        ' If it already has a second argument, replace
        i = InStr(iPos, sLine, ",")
        If i Then sLine = Left$(sLine, i - 1)
            
        ' Add second argument
        Dim s As String
        s = Mid$(sLine, iPos)
        sLine = sLine & ", """ & Trim(s) & ", " & _
                GetFileBaseExt(sSource) & ", line " & iLine & """"
        If sComment <> sEmpty Then sLine = sLine & " " & sComment
    End If
            
End Sub

Private Sub TrimAsserts(sLine As String)
    
    ' Check to see if line contains string
    Dim iPos As Integer, i As Integer, sComment As String
    iPos = InStr(sLine, sBugAssert)
    If iPos Then
    
        ' Save comment so it won't be processed
        i = InStr(sLine, "'")
        If i Then
            ' Ignore commented out Bug strings
            If iPos > i Then Exit Sub
            ' Remove comment
            sComment = Mid$(sLine, i)
            sLine = Left$(sLine, i - 1)
        End If
        ' Move to first argument
        iPos = iPos + Len(sBugAssert)
        
        ' Remove any second argument
        i = InStr(iPos, sLine, ",")
        If i Then sLine = Left$(sLine, i - 1)
        If sComment <> sEmpty Then sLine = sLine & sComment
            
    End If
            
End Sub

Private Function IUseFile_UseFile(UserData As Variant, _
                                  FilePath As String, _
                                  FileInfo As CFileInfo) As Boolean
With FileInfo
    Dim filter As IFilter
    Set filter = Me
    ' Set FilterType property
    Select Case UCase$(GetFileExt(.Item))
    ' All the code file formats I know about
    Case ".CLS", "FRM", "CTL", "BAS", "PAG", "DOB", "DSR"
        filter.Source = .Item
        ' Pass filter to FilterTextFile
        FilterTextFile filter
    Case Else
        ' Ignore other files
    End Select
    DoEvents
End With
End Function
