Attribute VB_Name = "htmlLIN"

Function gethref(marka As String) As String

Dim srcstr As String
Dim tempchar As String
srcstr = LCase(marka)

n = InStr(srcstr, "href=" + Chr(34))

If n > 0 Then
    n = n + 6
    tempchar = Mid(marka, n, 1)
    Do Until tempchar = Chr(34) Or n > Len(marka)
    gethref = gethref + tempchar
    n = n + 1
    tempchar = Mid(marka, n, 1)
    Loop
    
End If

'gethref = Chr(34) + gethref + Chr(34)
End Function

Function getlinks(htmlfile As String, links() As String, linknum As Integer) As Boolean

Dim HtmDoc As New HTMLDocument
Dim theHtm As IHTMLDocument2
Dim fso As New FileSystemObject
If fso.FileExists(htmlfile) = False Then Exit Function
Set theHtm = HtmDoc.createDocumentFromUrl(htmlfile, "")
Do Until theHtm.readyState = "complete"
DoEvents
Loop

Dim localLinks(10240, 1) As String


Dim A_C As IHTMLElementCollection
Set A_C = theHtm.All.tags("A")
For i = 0 To A_C.length - 1
tempstr = A_C(i).href
If Left(tempstr, 8) = "file:///" Then
    linknum = linknum + 1
    localLinks(linknum, 0) = gethref(A_C(i).outerHTML)
    If Left(localLinks(linknum, 0), 1) = "#" Then localLinks(linknum, 0) = fso.GetFileName(htmlfile) + localLinks(linknum, 0)
    localLinks(linknum, 1) = A_C(i).innerText
End If

Next
ReDim links(linknum, 1) As String
For i = 1 To linknum
links(i, 0) = localLinks(i, 0)
links(i, 1) = localLinks(i, 1)
Next


End Function

Sub htm2txt(htmlfile As String, textfile As String)
Dim HtmDoc As New HTMLDocument
Dim theHtm As IHTMLDocument2
Dim fso As New FileSystemObject
If fso.FileExists(htmlfile) = False Then Exit Sub
Set theHtm = HtmDoc.createDocumentFromUrl(htmlfile, "")
Do Until theHtm.readyState = "complete"
DoEvents
Loop
Dim ts As TextStream
Set ts = fso.OpenTextFile(textfile, ForWriting, True)
ts.Write theHtm.body.innerText
ts.Close
End Sub
Sub hhc2ztc(hhcfile As String, ztcfile As String)

Dim fso As New FileSystemObject
If fso.FileExists(hhcfile) = False Then Exit Sub
Dim tempstr() As String
Dim wstr As String
Dim strC As Integer
FilestrBetween hhcfile, "<param", ">", tempstr, strC
Dim ts As TextStream
Set ts = fso.OpenTextFile(ztcfile, ForWriting, True)

For i = 1 To strC
wstr = strBetween(tempstr(i), "value=" + Chr(34), Chr(34))
If LCase(strBetween(tempstr(i), "name=" + Chr(34), Chr(34))) = "name" Then
ts.Write wstr + ","
ElseIf LCase(strBetween(tempstr(i), "name=" + Chr(34), Chr(34))) = "local" Then
ts.Write wstr + vbCrLf
End If
Next
ts.Close
End Sub

Public Sub MoveAL(htmfile As String, moveTo As String)
Dim links() As String
Dim linknum As Integer
Dim fso As New FileSystemObject
pdir = fso.GetParentFolderName(htmfile)
getlinks htmfile, links(), linknum
For i = 1 To linknum
    srcfile = pdir + "\" + links(i, 0)
    dstfile = moveTo + "\" + links(i, 1) + ".htm"
    If fso.FileExists(srcfile) Then
        If fso.FileExists(dstfile) Then fso.DeleteFile dstfile
        fso.MoveFile srcfile, dstfile
    End If
Next
End Sub
Public Sub saveAl(htmlfile As String)
Dim HtmDoc As New HTMLDocument
Dim theHtm As IHTMLDocument2
Dim fso As New FileSystemObject
If fso.FileExists(htmlfile) = False Then Exit Sub
Set theHtm = HtmDoc.createDocumentFromUrl(htmlfile, "")
Do Until theHtm.readyState = "complete"
DoEvents
Loop
Dim ts As TextStream
If theHtm.Title <> "" Then
Set ts = fso.OpenTextFile(fso.GetParentFolderName(htmlfile) + "\" + theHtm.Title + ".txt", ForWriting, True)
Else
Set ts = fso.OpenTextFile(fso.GetParentFolderName(htmlfile) + "\" + fso.GetBaseName(htmlfile) + ".txt", ForWriting, True)
End If
Dim A_C As IHTMLElementCollection
Set A_C = theHtm.All.tags("A")
For i = 0 To A_C.length - 1
ts.WriteLine A_C(i).href + "|" + A_C(i).innerText
Next
ts.Close
fso.DeleteFile htmlfile
End Sub
