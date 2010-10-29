Attribute VB_Name = "Module1"
Option Explicit
Public Type searchEngine
name As String
alias As String
href As String
End Type
Public sLookoutInI As String

Public Function loadIni() As Boolean
Dim fso As New FileSystemObject
Dim ts As TextStream
Dim sTmp As String
Dim sETmp As String
Dim sHtmp As String
Dim sEngine() As String
Dim sHistory() As String
Dim lECount As Long
Dim lHCount As Long
Dim tseTMP As searchEngine
Dim l As Long
Dim t As Long
Set ts = fso.OpenTextFile(sLookoutInI, ForReading, True)
If ts.AtEndOfStream Then
    ts.Close
    Exit Function
End If
sTmp = ts.ReadAll

sETmp = LeftRange(sTmp, "[Engine]" & vbCrLf, vbCrLf & "[", vbTextCompare)
sHtmp = LeftRight(sTmp, "[History]" & vbCrLf)
sEngine = Split(sETmp, vbCrLf)
lECount = UBound(sEngine) + 1
sHistory = Split(sHtmp, vbCrLf)
lHCount = UBound(sHistory) + 1

For l = 0 To lECount - 1
tseTMP.name = LeftLeft(sEngine(l), ",")
tseTMP.href = LeftRight(sEngine(l), ",")
tseTMP.alias = RightRight(tseTMP.name, "|")
If tseTMP.alias <> "" Then tseTMP.name = RightLeft(tseTMP.name, "|")
If tseTMP.name <> "" And tseTMP.href <> "" Then
 ReDim tSE(t) As searchEngine
 tSE(t) = tseTMP
 t = t + 1
End If

Next




End Function
