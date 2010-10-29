Attribute VB_Name = "kerak"
Public DBWord As String
Public Wordname As String
Public DBBook As String
Public Bookname As String
Public DBEX As String
Public Exname As String
Public DBRps As String
Public Rpsname As String
Public Rps As Boolean
Public Rpsdir As String
Public Firstword As String
Public Lastword As String
Public Example As Integer
Public Readsecond As Integer
Public Repeattime As Integer
Public Readex As Boolean
Public Readrnd As Boolean
Public Readword As Boolean
Public OkDBWord As Boolean
Public OkDBBook As Boolean
Public OkDBEX As Boolean
Public OkDBRps As Boolean





Sub Creatini(ininame() As String, ininum As Integer, Inifile As String)
Dim fso As New FileSystemObject
Dim ft As TextStream
Set ft = fso.CreateTextFile(Inifile)
For i = 1 To ininum
ft.WriteLine ininame(i, 1) + "=" + ininame(i, 2)
Next
ft.Close
End Sub

Function readini(ininame() As String, ininum As Integer, Inifile As String)
Dim fso As New FileSystemObject

Dim ft As TextStream
Dim tmpstr As String
Set ft = fso.OpenTextFile(Inifile)

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



Public Sub getexfile()


Dim dbe As New DBEngine
Dim DBWord As Database
Dim recword As Recordset
Dim recex As Recordset
Dim fieldword As Field
Dim fieldex As Field
Dim fieldtmp As Field
Dim wrdwrd(10000, 5) As String
tDBword = "f:\personal\document\words.mdb"
Set DBWord = dbe.OpenDatabase(tDBword, , True, True)

Set recword = DBWord.OpenRecordset("研究生入学考试词汇")
Set recex = DBWord.OpenRecordset("句子")

For Each fieldtmp In recword.Fields
If fieldtmp.Name = "Word" Then Set fieldword = fieldtmp
Next
For Each fieldtmp In recex.Fields
If fieldtmp.Name = "E" Then Set fieldex = fieldtmp
Next
Dim fso As New FileSystemObject
Dim ft As TextStream
Set ft = fso.CreateTextFile(bddir(fso.GetParentFolderName(tDBword)) + "result.txt")
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






Set recword = DBWord.OpenRecordset("研究生入学考试词组")
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


Function chkdb(theDB As String, theTable As String, theField As String, chkinfo As String) As Integer

chkdb = 1
chkinfo = ""
Dim fso As New FileSystemObject
If fso.FileExists(theDB) = False Then
    chkinfo = "数据库" + theDB + "路径错误"
    Exit Function
End If

Dim dbe As New DBEngine
Dim chkingDB As Database
Dim chkingTable As TableDef
Dim chkingField As Field

Set chkingDB = dbe.OpenDatabase(theDB, , True, True)

chkdb = 2

For Each chkingTable In chkingDB.TableDefs
    If chkingTable.Name = theTable Then
    chkdb = 0
    Exit For
    End If
Next

If chkdb = 2 Then
    chkinfo = "数据库" + theDB + "不存在名为" + theTable + "的数据表"
    Exit Function
End If

chkdb = 3

For Each chkingField In chkingDB.TableDefs(theTable).Fields
    If chkingField.Name = theField Then
    chkdb = 0
    Exit For
    End If
Next

If chkdb = 3 Then chkinfo = "数据库" + theDB + "的" + theTable + "表中不存在名为" + theField + "的纪录"
    
End Function

