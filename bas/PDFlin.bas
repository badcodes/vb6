Attribute VB_Name = "PDFlin"
Function pdfReadline(filenum As Integer, Optional pos As Long) As String
If pos > 0 Then Seek filenum, pos
Dim thebyte As String
thebyte = Input(filenum, 1)
Do Until Asc(thebyte) = &HD Or Loc(filenum) = LOF(filenum)
If Asc(thebyte) <> &HA Then pdfReadline = pdfReadline + thebyte
thebyte = Input(filenum, 1)
Loop
End Function

Function pdfreadobj(filenum As Integer, beginpos As Long) As String
Seek filenum, beginpos
Dim tempstr As String
tempstr = pdfReadline(filenum)
Do Until Trim(tempstr) = "endobj"
pdfreadobj = pdfreadobj + tempstr + Chr(13) + Chr(10)
tempstr = pdfReadline(filenum)
Loop
pdfreadobj = pdfreadobj + "endobj"
End Function

Function getREFtable(filenum As Integer, reftable() As Long, refnum As Integer)
Seek filenum, LOF(filenum) - 7
Dim refposstr As String
Dim thebyte As Byte
Get #filenum, , thebyte

Do Until thebyte = &HD
refposstr = Chr(thebyte) + refposstr
Seek filenum, Loc(filenum) - 1
Get #filenum, , thebyte
Loop
Dim refpos As Long
refpos = Val(refposstr)
Seek filenum, refpos
pdfReadline filenum
pdfReadline filenum
tempstr = Trim(pdfReadline(filenum))
refnum = Val(Right(tempstr, Len(tempstr) - InStr(tempstr, " ")))
pdfReadline filenum
refnum = refnum - 1
ReDim reftable(refnum) As Long
For i = 1 To refnum
tempstr = Trim(pdfRadline(filenum))
reftable(i) = Val(Left(tempstr, InStr(tempstr, " ") - 1))
Next


End Function
