VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFilterTypeOP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ITemplateFilter
Private mHeader As String

Private Enum FTOPType
    FTOPNotSupport = -1
    FTOPAssign = 0
    FTOPDelete = 1
    FTOPCompare = 2
    FTOPDeleteArray = 3
End Enum

Private mOperator() As String
Private mOPCount As Long

'

Private Sub Class_Initialize()
    mHeader = "MTemplateHelper."
    ReDim mOperator(0 To 3) As String
    mOperator(0) = "Assign"
    mOperator(1) = "Delete"
    mOperator(2) = "Compare"
    mOperator(3) = "DeleteArray"
    mOPCount = 4
End Sub

Private Function ITemplateFilter_Process(Lines() As String, ByVal Count As Long, Types As ITypeInfo, Reportor As IFilterReporter) As Boolean

    Dim i As Long
    Dim Token As String
    Dim fChanged As Boolean
    Dim fEspectTypename As Boolean


    Dim sPadToken As String
    Dim nId As Long



    Dim sSuffix As String
    Dim sPrefix As String
    Dim sOld As String
    Dim sNew As String
    Dim fModified As Boolean
    Dim nIdx As Long
    
    ITemplateFilter_Process = False
    If Count < 1 Then Exit Function
    If Types Is Nothing Then Exit Function
    If Types.Count < 1 Then Exit Function
    
    Dim fInOperator As Boolean
    Dim fFunction As Boolean
    Dim fEspectPara As Boolean
    Dim sParams() As String
    Dim sOp As String
    Dim tOP As FTOPType
    Dim sTypeName As String
    Dim tTypeName As CTTypeStyles
    Dim sSave As String
    Dim sPara As String
    Dim nPara As Long
    Dim ch As String
    
    For i = 0 To Count - 1
Continue_Line:
        Token = GetToken(Lines(i), TOKEN_SEPARATOR)
        If Left$(Token, 1) = "'" Then GoTo Next_Line
        
        fFunction = False
        fModified = False
        fEspectTypename = False
        fEspectPara = False
        sNew = ""
        sPrefix = ""
        sSuffix = ""
        fInOperator = False
        sOp = ""
        tOP = FTOPNotSupport
        tTypeName = CTTypeUnknown
        sTypeName = ""
        nPara = 0
        Erase sParams()
        Do While Token <> ""
Continue_Token:
            fChanged = False
            If Not fInOperator Then
                sOp = ExtractOperator(Token, fFunction, sPrefix, sSuffix)
                If (sOp <> "") Then
                    tOP = CheckOperator(sOp)
                    If (tOP <> FTOPNotSupport) Then
                        sSave = Token
                        fInOperator = True
                        fEspectTypename = True
                        If sSuffix = "" Then GoTo Next_Token
                        Token = sSuffix
                        sSuffix = ""
                        GoTo Continue_Token
                    End If
                End If
                sPrefix = ""
                sSuffix = ""
                fFunction = False
                GoTo Save_And_Next_Token
            Else
                If fEspectTypename Then
                    sTypeName = Token
                    sTypeName = Replace$(sTypeName, """", "")
                    'If Right$(sTypeName, 1) = "," Then sTypeName = Left$(sTypeName, Len(sTypeName) - 1)
                    'If Left$(sTypeName, 1) = """" Then sTypeName = Mid$(sTypeName, 2, Len(sTypeName) - 2)
                    If (Right$(sTypeName, 1) = ",") Then
                        sTypeName = Mid$(sTypeName, 1, Len(sTypeName) - 1)
                        'sTypeName = Mid$(sTypeName, 2, Len(sTypeName) - 3)
                        nId = Types.IdOf(sTypeName)
                        If (nIdx >= 0) Then
                            tTypeName = Types.StyleOf(nId)
                            fEspectPara = True
                            fEspectTypename = False
                            GoTo Next_Token
                        End If

                    Else
                        'sTypeName = Mid$(sTypeName, 2, Len(sTypeName) - 2)
                        nId = Types.IdOf(sTypeName)
                        If (nIdx >= 0) Then
                            tTypeName = Types.StyleOf(nId)
                            fEspectTypename = False
                            fInOperator = False
                            Token = WriteOperator(tOP, tTypeName, sParams, fFunction)
                            fChanged = True
                            GoTo Save_And_Next_Token
                            GoTo Next_Token
                        End If
                    End If
                    If nIdx < 1 Then
                        fInOperator = False
                        fEspectTypename = False
                        fChanged = True
                        sPrefix = ""
                        sSuffix = ""
                        Token = sSave & " " & Token
                        GoTo Save_And_Next_Token
                    End If
                ElseIf fEspectPara Then
                    fChanged = True
                    sPara = Token
                    ch = Right$(sPara, 1)
                    If (ch = ",") Then
                        sPara = Left$(sPara, Len(sPara) - 1)
                        AddToArray sParams, nPara, sPara
                        fEspectPara = True
                    ElseIf fFunction And ch = ")" Then
                        sPara = Left$(sPara, Len(sPara) - 1)
                        AddToArray sParams, nPara, sPara
                        fInOperator = False
                        Token = WriteOperator(tOP, tTypeName, sParams, fFunction)
                        fChanged = True
                        GoTo Save_And_Next_Token
                    Else
                        AddToArray sParams, nPara, sPara
                        fInOperator = False
                        Token = WriteOperator(tOP, tTypeName, sParams, fFunction)
                        fChanged = True
                        GoTo Save_And_Next_Token
                    End If
                    GoTo Next_Token
                Else
                    GoTo Save_And_Next_Token
                End If
            End If
    
Save_And_Next_Token:
            sNew = sNew & sPrefix & Token & sSuffix & " "
            If (fChanged) Then
                fModified = True
                sPrefix = ""
                sSuffix = ""
            End If
Next_Token:
            Token = GetToken("", TOKEN_SEPARATOR)
        Loop
Save_And_Next_Line:
        If (fModified) Then
            sOld = Lines(i)
            MFilterHelper.AssignString Lines(i), sNew
            Reportor.NotifyChange sOld, Lines(i), i
        End If
Next_Line:
    Next
End Function

Public Property Let OpHeader(ByRef sHeader As String)
    mHeader = sHeader
End Property

Public Property Get OpHeader() As String
    OpHeader = mHeader
End Property

Private Function CheckOperator(ByRef sOperator As String) As FTOPType
    Dim i As Integer
    For i = 0 To mOPCount - 1
        If sOperator = mOperator(i) Then CheckOperator = i: Exit Function
    Next
    CheckOperator = FTOPNotSupport
End Function

Private Function CheckType(ByRef sTypeName As String, ByRef Types As ITypeInfo) As CTTypeStyles
    Dim idx As Long
    idx = Types.IdOf(sTypeName)
    If idx >= 0 Then CheckType = Types.StyleOf(idx) Else CheckType = CTTypeUnknown
End Function

Private Function WriteOperator(ByVal OPType As FTOPType, ByVal typeStyle As CTTypeStyles, ByRef Param() As String, fFunction As Boolean) As String
    If OPType = FTOPNotSupport Then Exit Function
    Select Case OPType
        Case FTOPType.FTOPAssign
            If typeStyle = CTTypeNormal Then
                WriteOperator = Param(0) & " = " & Param(1)
            ElseIf typeStyle = CTTypeObject Then
                WriteOperator = "Set " & Param(0) & " = " & Param(1)
            Else
                WriteOperator = "If IsObject(" & Param(1) & ")  then " & "Set " & Param(0) & " = " & Param(1) & " Else " & Param(0) & " = " & Param(1)
            End If
        Case FTOPType.FTOPDelete
            If typeStyle = CTTypeNormal Then
                WriteOperator = "" ' "'Nothing to do for deleting " & Param(0)
            ElseIf typeStyle = CTTypeObject Then
                WriteOperator = "Set " & Param(0) & " = Nothing"
            Else
                WriteOperator = "If IsObject(" & Param(0) & ")  then Set " & Param(0) & " = Nothing"
            End If
        Case FTOPType.FTOPDeleteArray
            If typeStyle = CTTypeNormal Then
                WriteOperator = "" ' "'Nothing to do for deleting " & Param(0)
            Else
                Dim rPara(0 To 0) As String
                rPara(0) = Left$(Param(0), Len(Param(0)) - 1) & "idel)"
                WriteOperator = "Dim iDel as long" & vbCrLf
                WriteOperator = WriteOperator & "For idel = " & Param(1) & " to " & Param(2) & " - 1 + " & Param(1) & vbCrLf
                WriteOperator = WriteOperator & vbTab & WriteOperator(FTOPDelete, typeStyle, rPara, fFunction) & vbCrLf
                WriteOperator = WriteOperator & "Next"
            End If
        Case FTOPType.FTOPCompare
             If typeStyle = CTTypeNormal Then
                WriteOperator = Param(0) & " = " & Param(1)
            ElseIf typeStyle = CTTypeObject Then
                WriteOperator = Param(0) & " Is " & Param(1)
            Else
                WriteOperator = "Iif(IsObject(" & Param(0) & "), " & Param(0) & " = " & Param(1) & ", " & Param(0) & " is " & Param(1) & ")"
            End If
            If fFunction Then WriteOperator = "(" & WriteOperator & ")"
        Case Else
            Debug.Print "Operator is not support"
            Exit Function
    End Select
    
End Function

Private Sub AddToArray(ByRef aString() As String, ByRef aCount As Long, ByRef sWhat As String)
    ReDim Preserve aString(0 To aCount) As String
    aString(aCount) = sWhat
    aCount = aCount + 1
End Sub

Private Function ExtractOperator(ByRef sToken As String, ByRef fFunction As Boolean, ByRef sLeft As String, ByRef sRight As String) As String
    Dim nIdx As Long
    
    nIdx = InStr(1, sToken, mHeader)
    If nIdx < 1 Then Exit Function
    sLeft = Left$(sToken, nIdx - 1)
    nIdx = Len(mHeader) + nIdx - 1
    sRight = Right$(sToken, Len(sToken) - nIdx)
    nIdx = InStr(sRight, "(")
    If nIdx > 0 Then
        fFunction = True
        ExtractOperator = Left(sRight, nIdx - 1)
        sRight = Right$(sRight, Len(sRight) - nIdx)
    Else
        fFunction = False
        ExtractOperator = sRight
        sRight = ""
    End If
End Function

