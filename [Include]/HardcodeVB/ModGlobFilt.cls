VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CModGlobFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IFilter

Private sSource As String, sTarget As String
Private sName As String

' CModGlobFilter-specific methods and properties
Public Property Let Name(sNameA As String)
    sName = sNameA
End Property
Public Property Get Name() As String
    Name = sName
End Property

' IFilter implementation
Private Property Let IFilter_Source(sSourceA As String)
    sSource = sSourceA
End Property
Private Property Get IFilter_Source() As String
    IFilter_Source = sSource
End Property

Private Property Let IFilter_Target(sTargetA As String)
    sTarget = sTargetA
End Property
Private Property Get IFilter_Target() As String
    IFilter_Target = sTarget
End Property

Private Function IFilter_Translate(sLineA As String, ByVal iLineA As Long) As EChunkAction
    ' Translate every line
    IFilter_Translate = ecaTranslate
    ' Change only the first line of module
    If iLineA = 1 Then CreateGlobHeader sLineA
End Function

Private Sub CreateGlobHeader(sLine As String)
    Dim sTok As String, sSep As String

    ' Parse module header
    sSep = " " & sTab
    sTok = GetQToken(sLine, sSep)
    BugAssert sTok = "Attribute"
    sTok = GetQToken(sEmpty, sSep)
    BugAssert sTok = "VB_Name"
    sTok = GetQToken(sEmpty, sSep)
    BugAssert sTok = "="
    
    ' Use default global name if global name isn't already set
    If sName = sEmpty Then
        sName = GetQToken(sEmpty, sSep)
        ' Remove this block if you don't use M as a tag on standard modules
        If Left$(sName, 1) = "M" Then
            sName = "G" & Right$(sName, Len(sName) - 1)
        Else
            sName = "G" & sName
        End If
    End If
    
    ' Generate global class header
    sLine = "VERSION 1.0 CLASS" & sCrLf & _
            "BEGIN" & sCrLf & _
            "  MultiUse = -1  'True" & sCrLf & _
            "END" & sCrLf & _
            "Attribute VB_Name = " & sQuote2 & sName & sQuote2 & sCrLf & _
            "Attribute VB_GlobalNameSpace = True" & sCrLf & _
            "Attribute VB_Creatable = True" & sCrLf & _
            "Attribute VB_PredeclaredId = False" & sCrLf & _
            "Attribute VB_Exposed = True"
End Sub


