VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CModGlobDelFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' CModGlobDelFilter implements IFilter
Implements IFilter

Enum EProcType
    eptNone
    eptMethodSub
    eptMethodFunc
    eptPropertyGet
    eptPropertyLet
    eptPropertySet
    eptUnknown
End Enum

Private sSource As String, sTarget As String
Private sModule As String, sCallLine As String, sArgList As String
Private fLineContinue As Boolean, eptType As EProcType
Private sName As String

' CModGlobDelFilter-specific methods and properties
Public Property Let Name(sNameA As String)
    sName = sNameA
End Property
Public Property Get Name() As String
    Name = sName
End Property

' Implementation of IFilter interface
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

' Great big, long, complex state machine all in one ugly chunk
Private Function IFilter_Translate(sLine As String, _
                                   ByVal iLine As Long) As EChunkAction
    Dim sTok As String, sSep As String
    sSep = " (" & sTab
    IFilter_Translate = ecaSkip ' We'll skip most lines
    
    ' Handle first line of module
    If iLine = 1 Then
        sTok = GetQToken(sLine, sSep)
        BugAssert sTok = "Attribute"
        sTok = GetQToken(sEmpty, sSep)
        BugAssert sTok = "VB_Name"
        sTok = GetQToken(sEmpty, sSep)
        BugAssert sTok = "="
        sModule = GetQToken(sEmpty, sSep)
        ' Use default global name if global name isn't already set
        If sName = sEmpty Then
            ' Remove this block if you don't use M as a tag on standard modules
            If Left$(sModule, 1) = "M" Then
                sName = "G" & Right$(sModule, Len(sModule) - 1)
            Else
                sName = "G" & sModule
            End If
        End If
        sLine = "VERSION 1.0 CLASS" & sCrLf & _
                "BEGIN" & sCrLf & _
                "  MultiUse = -1  'True" & sCrLf & _
                "END" & sCrLf & _
                "Attribute VB_Name = " & sQuote2 & sName & sQuote2 & sCrLf & _
                "Attribute VB_GlobalNameSpace = True" & sCrLf & _
                "Attribute VB_Creatable = True" & sCrLf & _
                "Attribute VB_PredeclaredId = False" & sCrLf & _
                "Attribute VB_Exposed = True"
        IFilter_Translate = ecaTranslate
        Exit Function
    End If
    
    ' Don't skip empty lines
    If sLine = sEmpty Then
        IFilter_Translate = ecaTranslate
        Exit Function
    End If
    
    ' Special case for line continuation on procedure definitions
    If Not fLineContinue Then
    
        ' Get first token
        sTok = GetQToken(sLine, sSep)
    
        ' Skip Public modifier
        If sTok = "Public" Then sTok = GetQToken(sEmpty, sSep)
    
        ' Look for Procedures
        Select Case sTok
        ' Create the delegated procedure
        Case "Attribute", "Option"
            ' Pass Attribute and Option lines through unchanged regardless of position
            IFilter_Translate = ecaTranslate
            sTok = GetQToken(sEmpty, sSep)
            Exit Function
        Case "Sub"
            ' Make ending line
            eptType = eptMethodSub
            sTok = GetQToken(sEmpty, sSep)
            sCallLine = "   " & sModule & "." & sTok & " "
        Case "Function"
            ' Make ending line
            eptType = eptMethodFunc
            sTok = GetQToken(sEmpty, sSep)
            sCallLine = "    " & sTok & " = " & sModule & "." & sTok & "("
        Case "Property"
            ' Make ending line
            sTok = GetQToken(sEmpty, sSep)
            ' Handle different property types
            Select Case sTok
            Case "Get"
                eptType = eptPropertyGet
                sTok = GetQToken(sEmpty, sSep)
                sCallLine = "   " & sTok & " = " & sModule & "." & sTok
            Case "Let"
                eptType = eptPropertyLet
                sTok = GetQToken(sEmpty, sSep)
                sCallLine = "   " & sModule & "." & sTok & " = " & sTok
            Case "Set"
                eptType = eptPropertySet
                sTok = GetQToken(sEmpty, sSep)
                sCallLine = "   " & "Set " & sModule & "." & sTok & " = " & sTok
            End Select
        Case sEmpty
            IFilter_Translate = ecaTranslate
            Exit Function
        Case Else
            ' Skip all other lines
            IFilter_Translate = ecaSkip
            Exit Function
        End Select
        ' Get the first parameter token (space separated only)
        sTok = GetQToken(sEmpty, " ")
    Else
        sTok = GetQToken(sLine, " ")
    End If
    
    'Debug.Print sTok
    ' Process arguments
    Dim cParams As Long
    Do While sTok <> sEmpty
        If Left$(sTok, 1) = "(" Then sTok = Mid$(sTok, 2)
        
        'If sTok = "" Then GoTo NextCase2
        Select Case sTok
        Case "ByVal", "ByRef", "Optional", "ParamArray"
            ' Ignore modifiers
            GoTo NextCase2
        Case "_"
            ' Line continuation
            fLineContinue = True
            IFilter_Translate = ecaTranslate
            Exit Function
        Case ")"
            ' Empty argument list
            If Right$(sCallLine, 1) = "(" Then
                sCallLine = Left$(sCallLine, Len(sCallLine) - 1)
            End If
            Exit Do
        End Select
        
'       If Right$(sTok, 1) = ")" Then
'            sArgList = sArgList & sTok
'            Exit Do
'        'sTok = Left$(sTok, Len(sTok) - 1)
'        End If
       'If Right$(sTok, 1) = "," Then sTok = Left$(sTok, Len(sTok) - 1)
        cParams = cParams + 1
        sArgList = sArgList & sTok
        ' Get As
        sTok = GetQToken(sEmpty, " ")
            If sTok <> "As" Then
                sArgList = sEmpty
                sCallLine = "    ' Can't translate"
                Exit Do
            End If
'        If sTok = ")" Then
'            Debug.Print sTok
'        ElseIf sTok <> "As" Then
'
'        End If
        
        ' Get type with ending , or )

        sTok = GetQToken(sEmpty, sSep)
        Dim sNext As String
NextCase:
        sNext = Right$(sTok, 1)
        Select Case sNext
        Case ","
            sArgList = sArgList & sNext & " "
        Case ")"
            If eptType = eptMethodFunc Then sArgList = sArgList & sNext
            Exit Do
        Case Else
            ' If no terminating , or ), throw away optional argument
            sTok = GetQToken(sEmpty, sSep)
'            If sTok = "," Or sTok = ")" Then GoTo NextCase
            If sTok <> "=" Then
                sArgList = sEmpty
                sCallLine = "    ' Can't translate"
                Exit Do
            End If
'            Debug.Print sTok
            sTok = GetQToken(sEmpty, sSep)
'            Debug.Print sTok
'            Debug.Assert sTok <> "lTmp"
'            'Do Until sTok <> "," And sTok <> ")"
'            If Right$(sTok, 1) <> "," And Right$(sTok, 1) <> ")" Then
'                sTok = GetQToken(sEmpty, " ")
'                Debug.Assert sTok <> "lTmp"
'                Debug.Print sTok
'            End If
            GoTo NextCase
        End Select
        ' Next parameter name
NextCase2:
        sTok = GetQToken(sEmpty, " ")
    Loop
    
    ' Add the delegated line
    Select Case eptType
    Case eptMethodSub
        sLine = sLine & sCrLf & sCallLine & sArgList & sCrLf & "End Sub"
    Case eptMethodFunc
        sLine = sLine & sCrLf & sCallLine & sArgList & sCrLf & "End Function"
    Case eptPropertyLet, eptPropertySet
        If cParams > 1 Then
            sArgList = sEmpty
            sCallLine = "    ' Can't translate"
        End If
        sLine = sLine & sCrLf & sCallLine & sArgList & sCrLf & "End Property"
    Case eptPropertyGet
        If cParams Then
            sArgList = sEmpty
            sCallLine = "    ' Can't translate"
        End If
        sLine = sLine & sCrLf & sCallLine & sArgList & sCrLf & "End Property"
    End Select
    ' Reset defaults
    sArgList = sEmpty
    eptType = eptNone
    sCallLine = sEmpty
    fLineContinue = False
    IFilter_Translate = ecaTranslate
    Exit Function

End Function

