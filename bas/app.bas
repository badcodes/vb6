Attribute VB_Name = "MYVBAPP"
Sub Creatini(ininame() As String, Inifile As String)
Dim fso As New FileSystemObject
Dim ft As TextStream
Set ft = fso.CreateTextFile(Inifile, True)
For i = 0 To UBound(ininame)
ft.WriteLine ininame(0, i) + "=" + ininame(1, i)
Next
ft.Close
End Sub

Function readini(ininame() As String, Inifile As String)
Dim fso As New FileSystemObject

Dim ft As TextStream
Dim tmpstr As String
Dim pos As Integer
Set ft = fso.OpenTextFile(Inifile)

'tmpstr = ft.ReadAll
'For i = 1 To ininum
'pos = InStr(1, tempstr, ininame(i, 1), vbTextCompare)
'If pos > 0 Then
'    pos2 = InStr(pos + Len(ininame(i, 1)), tempstr, vbCrLf, vbTextCompare)
'    If pos2 > 0 Then ininame(i, 2) = Mid(tempstr, pos + Len(ininame(i, 1)), pos2 - pos - Len(ininame(i, 1)) + 1)
'End If
'Next
i = -1
Do Until ft.AtEndOfStream
tmpstr = ft.ReadLine
pos = InStr(tmpstr, "=")

If pos > 0 Then
    i = i + 1
    ReDim Preserve ininame(1, i) As String
    ininame(0, i) = Left(tmpstr, pos - 1)
    ininame(1, i) = Right(tmpstr, Len(tmpstr) - pos)
End If

Loop

ft.Close

End Function
