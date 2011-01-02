VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Sub eBeforeNavigate(ByRef URL As Variant, ByRef Cancel As Boolean, ByRef IEView As WebBrowser, Optional targetFrameName As String)
End Sub
Sub eNavigateComplete(ByRef URL As Variant, ByRef IEView As WebBrowser)
End Sub
Sub eStatusTextChange(ByVal text As String, ByRef IEView As WebBrowser)
End Sub
Sub MGetView(shortfile As String, ByRef IEView As WebBrowser, Optional ByRef targetFrameName As String)
End Sub
Sub startUP()
End Sub
Sub endUP()
End Sub
Sub MNavigate(sUrl As String, IE As WebBrowser, Optional targetFrameName As String)
End Sub

