VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CGlobModFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IFilter

Private sSource As String, sTarget As String
Private fAttribute As Boolean, fName As Boolean
Private sName As String

' CPubPrivFilter-specific methods and properties
Public Property Let Name(sNameA As String)
    sName = sNameA
End Property
Public Property Get Name() As String
    Name = sName
End Property

Private Sub Class_Initialize()
    fAttribute = True
End Sub

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
    ' Translate most lines
    IFilter_Translate = ecaTranslate
    
    If Not fName Then
        ' Skip everything until we find VB_Name
        fName = IsNameFound(sLineA)
        If Not fName Then IFilter_Translate = ecaSkip
    ElseIf fAttribute Then
        ' Skip the remaining attributes
        fAttribute = IsAttribute(sLineA)
        If fAttribute Then IFilter_Translate = ecaSkip
    End If
End Function

Private Function IsNameFound(sLine As String) As Boolean
    If sLine = sEmpty Then Exit Function
    
    ' Find VB_Name attribute and change it
    Dim sSep As String
    sSep = " " & sTab
    If GetQToken(sLine, sSep) = "Attribute" Then
        If GetQToken(sEmpty, sSep) = "VB_Name" Then
            ' Use default standard name if standard name isn't already set
            If sName = sEmpty Then
                ' Skip "="
                Call GetQToken(sEmpty, sSep)
                sName = GetQToken(sEmpty, sSep)
                If Left$(sName, 1) = "G" Then
                    sName = "M" & Right$(sName, Len(sName) - 1)
                Else
                    sName = "M" & sName
                End If
            End If
            sLine = "Attribute VB_Name = " & sQuote2 & sName & sQuote2
            IsNameFound = True
        End If
    End If
End Function

Private Function IsAttribute(sLine As String) As Boolean
    If sLine = sEmpty Then Exit Function
    
    Dim sSep As String
    sSep = "_"
    IsAttribute = (GetQToken(sLine, sSep) = "Attribute VB")
End Function
