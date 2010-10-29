VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CToolKit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IObjectWithSiteTLB.IObjectWithSite
Dim WithEvents IE As SHDocVw.InternetExplorer
Attribute IE.VB_VarHelpID = -1
Dim Site As IObjectWithSiteTLB.IUnknownVB
Const ssReader As String = "book:/"

Private Sub IE_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)

#If DebugMode = 1 Then
    MDebug.DebugFPrint "IE_BeforeNavigate2:URL_I=" & URL
#End If



'For ssReader
Dim urlHander As bookUrl.CHandler
Set urlHander = New bookUrl.CHandler
'If urlHander.wakeUp(CStr(URL)) Then
'    Cancel = True
'Else
'    Cancel = False
'End If
Cancel = urlHander.wakeUp(CStr(URL))
Set urlHander = Nothing

    
'    If InStr(1, URL, ss_param_replace, vbTextCompare) > 0 Then
'        Exit Sub
'    ElseIf InStr(1, URL, ss_param_find, vbTextCompare) > 0 Then
'        URL = Replace$(URL, ss_param_find, ss_param_replace, , , vbTextCompare)
'    Else
'        URL = URL & "&" & ss_param_replace
'    End If
'    Cancel = True




End Sub


Private Sub IObjectWithSite_GetSite(ByVal priid As IObjectWithSiteTLB.GUIDPtr, ppvObj As IObjectWithSiteTLB.VOIDPtr)
    Site.QueryInterface priid, ppvObj
End Sub

Private Sub IObjectWithSite_SetSite(ByVal pSite As IObjectWithSiteTLB.IUnknownVB)
    Set Site = pSite
    Set IE = pSite
End Sub




