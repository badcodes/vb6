Attribute VB_Name = "MApp"
Option Explicit

Public Function aboutApp() As String
aboutApp = App.ProductName & " (Build" & Str$(App.Major) + "." + Str$(App.Minor) & "." & Str$(App.Revision) & ")"
aboutApp = aboutApp & vbCrLf & App.LegalCopyright & " " & App.CompanyName & " " & CStr(Year(Date))
End Function
