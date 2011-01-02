VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "forIe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Public Sub addFavorite(url, info)
Dim docXML As New DOMDocument40
Dim tagrs As tagREADYSTATE
docXML.Load url
Do
    DoEvents
    tagrs = docXML.readyState
Loop Until tagrs = READYSTATE_COMPLETE
Dim nl As IXMLDOMNodeList
Dim n As IXMLDOMNode
Dim ctitle
Set nl = docXML.getElementsByTagName("channel")
If nl.length > 0 Then
    Set n = nl(0).selectSingleNode("title")
    ctitle = n.Text
End If
    
If ctitle <> "" Then info = ctitle
Dim ss As New ShellUIHelper
ss.addFavorite url + "|rssfeed", info
End Sub

