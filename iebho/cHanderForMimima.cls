VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cHanderForMimina"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IObjectWithSiteTLB.IObjectWithSite
Dim WithEvents IE As SHDocVw.InternetExplorer
Attribute IE.VB_VarHelpID = -1
Dim Site As IObjectWithSiteTLB.IUnknownVB
Const MIMIMA As String = "/link.php"


Private Sub IE_DocumentComplete(ByVal pDisp As Object, URL As Variant)
Dim sUrl As String
Dim pObj As Object
sUrl = LCase$(IE.LocationURL)
On Error Resume Next
If InStr(1, sUrl, MIMIMA) > 0 Then
    Set pObj = IE.Document.getElementsByTagName("form")
    pObj(0).setAttribute "onSubmit", ""
    pObj(0).onsubmit = ""
End If
End Sub


Private Sub IObjectWithSite_GetSite(ByVal priid As IObjectWithSiteTLB.GUIDPtr, ppvObj As IObjectWithSiteTLB.VOIDPtr)
    Site.QueryInterface priid, ppvObj
End Sub

Private Sub IObjectWithSite_SetSite(ByVal pSite As IObjectWithSiteTLB.IUnknownVB)
    Set Site = pSite
    Set IE = pSite
End Sub




