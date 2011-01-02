Attribute VB_Name = "MTEST"
Option Explicit
Public Sub TEST()
Dim A As New CHtmLDocument
Dim r() As String
Dim l As Long
Dim i As Long
A.CreateFromFile ("c:\1.htm")
l = A.GetTagProperty(r(), "href", "a")
For i = 1 To l
Debug.Print r(i)
Next

End Sub
