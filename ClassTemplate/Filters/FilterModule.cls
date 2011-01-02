VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFilterModule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ITemplateFilter
Private Const VBQuote As String = """"
Private Const TEST_ATtribute_VB As String = "Attribute VB_Name = " & VBQuote

'Private Function ITemplateFilter_Process(sLine As String, vDest As String, Types As ITypeInfo, ByRef Reportor As IFilterReporter) As Boolean
'    Dim nPosStart As Long
'    Dim nPosEnd As Long
'    Dim sTemplate As String
'    Dim nId As Integer
'    nPosStart = InStr(1, sLine, TEST_ATtribute_VB)
'    If (nPosStart > 0) Then
'        nPosStart = nPosStart + Len(TEST_ATtribute_VB) - 1
'        nPosEnd = InStr(nPosStart + 1, sLine, VBQuote)
'        If (nPosEnd > 0) Then
'            sTemplate = Mid$(sLine, nPosStart + 1, nPosEnd - nPosStart - 1)
'            nId = Types.IdOf(sTemplate)
'            If (nId >= 0) Then
'                vDest = Left$(sLine, nPosStart) & Types.NameOf(nId) & Mid$(sLine, nPosEnd)
'                Reportor.NotifyChange "VB_Name = " & sTemplate, "VB_Name = " & Types.NameOf(nId)
'                ITemplateFilter_Process = True
'            End If
'        End If
'    End If
'End Function
Private Function ITemplateFilter_Process(Lines() As String, ByVal Count As Long, Types As ITypeInfo, Reportor As IFilterReporter) As Boolean
    If Count < 1 Then Exit Function
    Dim i As Long
    Dim nPosStart As Long
    Dim nPosEnd As Long
    Dim nId As Integer
    Dim sTemplate As String
    Dim sTypeName As String
    Dim sLine As String
    For i = 0 To Count - 1
        sLine = Lines(i)
        nPosStart = InStr(1, sLine, TEST_ATtribute_VB)
        nPosEnd = 0
        nId = -1
        If (nPosStart > 0) Then
            nPosStart = nPosStart + Len(TEST_ATtribute_VB) - 1
            nPosEnd = InStr(nPosStart + 1, sLine, VBQuote)
            If (nPosEnd > 0) Then
                sTemplate = Mid$(sLine, nPosStart + 1, nPosEnd - nPosStart - 1)
                nId = Types.IdOf(sTemplate)
                If (nId >= 0) Then
                    sTypeName = Types.NameOf(nId)
                    Lines(i) = Left$(sLine, nPosStart) & sTypeName & Mid$(sLine, nPosEnd)
                    Reportor.NotifyChange sLine, Lines(i), i
                    ITemplateFilter_Process = True
                End If
            End If
        End If
    Next
End Function
