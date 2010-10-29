VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GInternet"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFilename As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long

Public Function netDownloadFile(ByVal url As String, ByVal sSaveAs As String) As Boolean
    
Dim hR As Long
hR = URLDownloadToFile(0, url, sSaveAs, 0, 0)
If hR = 0 Then netDownloadFile = True

End Function

