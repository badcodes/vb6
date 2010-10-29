VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cHandleForIElOG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements IObjectWithSiteTLB.IObjectWithSite
Dim WithEvents IE As SHDocVw.InternetExplorer
Attribute IE.VB_VarHelpID = -1
Dim Site As IObjectWithSiteTLB.IUnknownVB
'Const URLFORVIEWLOG = "ViewLog:"
'Const URLFORCLEARLOG = "ClearLog:"
Private sLogFile As String

Private Sub Class_Initialize()
sLogFile = App.Path & "\" & Date$ & ".htm"
End Sub

'Private Sub IE_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
'On Error Resume Next
'Dim sUrl As String
'sUrl = CStr(URL)
'If StrComp(sUrl, URLFORVIEWLOG, vbTextCompare) = 0 Then
'    Cancel = True
'    IE.Navigate2 sLogFile
'ElseIf StrComp(sUrl, URLFORCLEARLOG, vbTextCompare) = 0 Then
'    Cancel = True
'    Kill sLogFile
'    IE.Document.write "log cleared"
''ElseIf Not fliterUrl(sUrl) Then
''    Dim fnum As Integer
''    fnum = FreeFile
''    Open sLogFile For Append As fnum
''    Print #fnum, "[" & Date$ & " " & Time$ & "]" & sUrl
''    Close fnum
'End If
'End Sub

Private Sub IE_DocumentComplete(ByVal pDisp As Object, URL As Variant)
On Error Resume Next
Dim sUrl As String
sUrl = IE.LocationURL
'If StrComp(sUrl, URLFORVIEWLOG, vbTextCompare) = 0 Then
'    Cancel = True
'    IE.Navigate2 sLogFile
'ElseIf StrComp(sUrl, URLFORCLEARLOG, vbTextCompare) = 0 Then
'    Cancel = True
'    Kill sLogFile
'    IE.Document.write "log cleared"
'Else
If Not fliterUrl(sUrl) Then
    Dim fnum As Integer
    Dim sTitle As String
    sTitle = IE.LocationName
    fnum = FreeFile
    Open sLogFile For Append As fnum
    Print #fnum, "[" & Time$ & "]" & _
            "<a href='" & sUrl & "'>" & sTitle & "</a><br>"
    Close fnum
End If
End Sub

Private Sub IObjectWithSite_GetSite(ByVal priid As IObjectWithSiteTLB.GUIDPtr, ppvObj As IObjectWithSiteTLB.VOIDPtr)
    Site.QueryInterface priid, ppvObj
End Sub

Private Sub IObjectWithSite_SetSite(ByVal pSite As IObjectWithSiteTLB.IUnknownVB)
    Set Site = pSite
    Set IE = pSite
End Sub
Private Function fliterUrl(sUrl As String) As Boolean
    If StrComp(sUrl, "about:blank", vbTextCompare) = 0 Then
       fliterUrl = True
    End If
End Function

