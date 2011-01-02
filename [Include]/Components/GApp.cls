VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GApp"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function aboutApp() As String
aboutApp = App.ProductName & " (Build" & Str$(App.Major) + "." + Str$(App.Minor) & "." & Str$(App.Revision) & ")"
aboutApp = aboutApp & vbCrLf & App.LegalCopyright & " " & App.CompanyName & " " & CStr(Year(Date))
End Function

