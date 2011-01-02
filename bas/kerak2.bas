Attribute VB_Name = "kerak"
Sub insertline(srcfile As String, txtLINE As String, insTYPE As Integer)
Dim fso As FileSystemObject
Dim srcf As file, dstf As file
Dim srcfs As TextStream, dstfs As TextStream
Dim tmpstr As String
Set fso = New FileSystemObject
Set srcf = fso.GetFile(srcfile)
fso.CreateTextFile "$1$2$3.$$$", True
Set dstf = fso.GetFile("$1$2$3.$$$")
Set srcfs = srcf.OpenAsTextStream(ForReading)
Set dstfs = dstf.OpenAsTextStream(ForWriting)
Select Case insTYPE
Case 1
    Do Until srcfs.AtEndOfStream
    tmpstr = srcfs.ReadLine
    dstfs.WriteLine (txtLINE + tmpstr)
    Loop
Case Else
    Do Until srcfs.AtEndOfStream
    tmpstr = srcfs.ReadLine
    dstfs.WriteLine (tmpstr + txtLINE)
    Loop
End Select
srcfs.Close
dstfs.Close
fso.DeleteFile srcfile, True
fso.MoveFile "$1$2$3.$$$", srcfile

End Sub

Function rndarray(num As Integer, numarray() As Integer)
num = Int(num)
ReDim numarray(num) As Integer
ReDim recnum(num) As Integer
rec = 0
Randomize Time
For i = 1 To num Step 1
numarray(i) = Int(Rnd(Time) * num) + 1
    For j = 1 To rec Step 1
    If numarray(i) = recnum(j) Then
        numarray(i) = Int(Rnd(Time) * num) + 1
        j = 1
    End If
    Next
rec = rec + 1
recnum(rec) = numarray(i)

Next
    

End Function


Sub Creatini(ininame() As String, ininum As Integer, inifile As String)
Dim fso As New FileSystemObject
Dim ff As file
Dim ft As TextStream
Set ft = fso.CreateTextFile(inifile)
For i = 1 To ininum
ft.WriteLine ininame(i, 1) + "=" + ininame(i, 2)
Next
ft.Close
End Sub

Function readini(ininame() As String, ininum As Integer, inifile As String)
Dim fso As New FileSystemObject

Dim ft As TextStream
Dim tmpstr As String
Set ft = fso.OpenTextFile(inifile)

Do Until ft.AtEndOfStream
tmpstr = ft.ReadLine

For i = 1 To ininum
If LCase(ininame(i, 1)) = LCase(Left(tmpstr, InStr(tmpstr, "=") - 1)) Then
    ininame(i, 2) = Right(tmpstr, Len(tmpstr) - InStr(tmpstr, "="))
    Exit For
End If
Next

Loop
ft.Close



End Function
Public Function bddir(dirname As String) As String
bddir = dirname
If Right(bddir, 1) <> "\" Then bddir = bddir + "\"
End Function


Public Sub getexfile()


Dim dbe As New DBEngine
Dim dbword As Database
Dim recword As Recordset
Dim recex As Recordset
Dim fieldword As Field
Dim fieldex As Field
Dim fieldtmp As Field
Dim wrdwrd(10000, 5) As String
tdbfile = "f:\personal\document\words.mdb"
Set dbword = dbe.OpenDatabase(tdbfile, , True, True)

Set recword = dbword.OpenRecordset("研究生入学考试词汇")
Set recex = dbword.OpenRecordset("句子")

For Each fieldtmp In recword.Fields
If fieldtmp.Name = "Word" Then Set fieldword = fieldtmp
Next
For Each fieldtmp In recex.Fields
If fieldtmp.Name = "E" Then Set fieldex = fieldtmp
Next
Dim fso As New FileSystemObject
Dim ft As TextStream
Set ft = fso.CreateTextFile(bddir(fso.GetParentFolderName(tdbfile)) + "result.txt")
ft.WriteLine "word|ex1|ex2|ex3"
Dim writedata(5) As String

Dim match As Boolean
Dim tmpstr As String

Do Until recword.EOF

writedata(1) = fieldword.Value
writedata(2) = ""
writedata(3) = ""
writedata(4) = ""

 km = 0
 match = False
 recex.MoveFirst
 
 Do Until match Or recex.EOF
 
 

 If InStr(1, fieldex.Value, " " + writedata(1) + " ") Then
 km = km + 1
 writedata(km + 1) = fieldex.Value
    If km = 3 Then
    match = True
     Exit Do
     End If
 End If
 
 recex.MoveNext
 
 Loop
 
 ft.WriteLine writedata(1) + "|" + writedata(2) + "|" + writedata(3) + "|" + writedata(4)
 Debug.Print writedata(1)
 recword.MoveNext
 
Loop






Set recword = dbword.OpenRecordset("研究生入学考试词组")
For Each fieldtmp In recword.Fields
If fieldtmp.Name = "Word" Then Set fieldword = fieldtmp
Next

Do Until recword.EOF

writedata(1) = fieldword.Value
writedata(2) = ""
writedata(3) = ""
writedata(4) = ""

 km = 0
 match = False
 recex.MoveFirst
 
 Do Until match Or recex.EOF
 
 

 If InStr(1, fieldex.Value, " " + writedata(1) + " ") Then
 km = km + 1
 writedata(km + 1) = fieldex.Value
    If km = 3 Then
    match = True
     Exit Do
     End If
 End If
 
 recex.MoveNext
 
 Loop
 
 ft.WriteLine writedata(1) + "|" + writedata(2) + "|" + writedata(3) + "|" + writedata(4)
  Debug.Print writedata(1)
 recword.MoveNext



Loop




  ft.Close
  MsgBox "ok"
End Sub

