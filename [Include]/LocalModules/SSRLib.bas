Attribute VB_Name = "MSSRLib"
Option Explicit
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Const libList = "D:\Read\SSREADER39\remote112\liblist.dat"
'
'Public Function queryPdgLibExample(Optional ByRef rootTree As String = "", Optional ByRef strQuery As String = "") As String
'
'    Dim Lib() As String
'    Dim libCount As Long
'    Dim Cata() As String
'    Dim cataCount As Long
'    Dim Book() As String
'    Dim bookCount As String
'
'    If rootTree = "" Then rootTree = libList
'    If strQuery = "" Then Exit Function
'
'    libCount = pdgLibList(rootTree, Lib())
'
'    Dim i As Long
'    For i = 1 To libCount
'        cataCount = pdgCatalist(Lib(2, i), Cata())
'        Dim j As Long
'        For j = 1 To cataCount
'            bookCount = pdgBookList(Cata(2, j), Book())
'        Next
'    Next
'
'End Function

Public Function pdgLibList(ByRef rootTree As String, ByRef Lib() As String) As Long
    Dim fso As FileSystemObject
    Dim ts As TextStream
    Dim tmp() As String
    
    Set fso = New FileSystemObject
    Set ts = fso.OpenTextFile(rootTree, ForReading, False)
    Erase Lib()
    Do Until ts.AtEndOfStream
        tmp = Split(ts.ReadLine, "|")
        If UBound(tmp) > 1 Then
            pdgLibList = pdgLibList + 1
            ReDim Preserve Lib(1 To 2, 1 To pdgLibList) As String
            Lib(1, pdgLibList) = tmp(0)
            If InStr(tmp(1), ":") = 2 Then
                Lib(2, pdgLibList) = fso.BuildPath(tmp(1), "bktree.dat")
            Else
                Lib(2, pdgLibList) = fso.BuildPath(fso.BuildPath(fso.GetParentFolderName(rootTree), tmp(1)), "bktree.dat")
            End If
            'Debug.Print Lib(1, pdgLibList) & " - " & Lib(2, pdgLibList)
        End If
    Loop
    ts.Close
    Set ts = Nothing
    Set fso = Nothing
End Function

Public Function pdgCatalist(ByRef cataTree As String, ByRef Cata() As String) As Long
    Dim fso As FileSystemObject
    Dim ts As TextStream
    Dim tmp() As String
    Dim tmpStr As String
    Erase Cata()
    
    Set fso = New FileSystemObject
    If fso.FileExists(cataTree) = False Then Exit Function
    Erase Cata()
    Set ts = fso.OpenTextFile(cataTree, ForReading, False)
    Do Until ts.AtEndOfStream
        tmp = Split(ts.ReadLine, "|")
        If UBound(tmp) > 1 Then
            pdgCatalist = pdgCatalist + 1
            ReDim Preserve Cata(1 To 4, 1 To pdgCatalist) As String
            Cata(4, pdgCatalist) = tmp(0)
            Cata(3, pdgCatalist) = getCataID(tmp(1))
            Cata(2, pdgCatalist) = fso.BuildPath(fso.GetParentFolderName(cataTree), tmp(2))
            Dim pCata As String
            tmpStr = tmp(1)
            If tmpStr = "" Then tmpStr = tmp(2)
            pCata = parentCata(parentCataId(getCataID(tmpStr)), Cata(), pdgCatalist - 1)

            If pCata <> "" Then
                'Debug.Print pCata
                Cata(1, pdgCatalist) = pCata & "\" & tmp(0)
            Else
                Cata(1, pdgCatalist) = tmp(0)
            End If
            'Debug.Print Cata(1, pdgCatalist) & " - " & Cata(2, pdgCatalist)
        End If
    Loop
    ts.Close
    Set ts = Nothing
    Set fso = Nothing
End Function


Public Function pdgBookList(ByRef bookTree As String, ByRef Book() As String) As Long
    Dim fso As FileSystemObject
    Dim ts As TextStream
    Dim tmp() As String
    
    Set fso = New FileSystemObject
    If fso.FileExists(bookTree) = False Then Exit Function
    Erase Book()
    Set ts = fso.OpenTextFile(bookTree, ForReading, False)
    Do Until ts.AtEndOfStream
        tmp = Split(ts.ReadLine, "|")
        If UBound(tmp) > 1 Then
            pdgBookList = pdgBookList + 1
            ReDim Preserve Book(1 To 6, 1 To pdgBookList) As String
            Book(1, pdgBookList) = tmp(4)
            Book(2, pdgBookList) = tmp(0)
            Book(3, pdgBookList) = tmp(3)
            Book(4, pdgBookList) = tmp(2)
            Book(5, pdgBookList) = tmp(5)
            Book(6, pdgBookList) = tmp(1)
            'Debug.Print Book(1, pdgBookList) & " - " & Book(2, pdgBookList)
        End If
    Loop
    ts.Close
    Set ts = Nothing
    Set fso = Nothing
End Function

Public Function parentCataId(ByVal cataID As String) As String
    If cataID = "" Then Exit Function
    parentCataId = Left$(cataID, Len(cataID) - 2)
End Function

Public Function getCataID(ByVal listName As String) As String
    If Len(listName) < 12 Then Exit Function
    getCataID = Mid$(listName, 9, Len(listName) - 12)
End Function



Public Function parentCata(ByVal cataID As String, ByRef catalist() As String, cataCount As Long) As String
Dim i As Long
'Dim fso As New FileSystemObject
'Dim cataName As String
'Dim treeName As String
'Dim tmpStr As String
Dim pCataID As String
Dim idx As Long

If cataID = "" Then Exit Function

For i = 1 To cataCount
    If StrComp(cataID, catalist(3, i), vbTextCompare) = 0 Then
        idx = i
        Exit For
    End If
Next

If idx = 0 Then Exit Function
pCataID = parentCataId(cataID)
parentCata = catalist(4, idx)

If pCataID <> "" Then
    parentCata = parentCata(pCataID, catalist, cataCount) & "\" & parentCata
End If

End Function


Public Function MyInstr(strBig As String, strList As String, Optional strListSep As String = ",", Optional cmp As VbCompareMethod = vbBinaryCompare) As Boolean

    Dim i As Long
    Dim strcount As Integer
    Dim strSmallOne() As String

    If strList = "" Then MyInstr = True: Exit Function
    strSmallOne = Split(strList, strListSep)
    strcount = UBound(strSmallOne)

    For i = 0 To strcount
        If InStr(1, strBig, strSmallOne(i), cmp) > 0 Then MyInstr = True: Exit Function
    Next

End Function

Public Function newMdb(ByRef mdbPath As String) As Database
    Dim db As DBEngine
    Dim dbase As Database
    On Error GoTo haveError:
    
    Set db = New DBEngine
    'Set newMdb = Nothing
    Set dbase = db.CreateDatabase(mdbPath, dbLangChineseSimplified, dbVersion40)
    Set newMdb = dbase
    Exit Function
haveError:
    Set newMdb = Nothing
    Set dbase = Nothing
    Set db = Nothing
End Function
Public Function newTable(ByRef dbase As Database, ByRef tableName As String) As Recordset
    'Dim rs As Recordset
    On Error GoTo Do_Nothing
    Dim tdef As TableDef
'    Dim idx As Index
'    Dim f As Field
    
    
    Set tdef = dbase.CreateTableDef(tableName, dbAttachExclusive)
'    Set idx = tdef.CreateIndex("ssid")
'    Set f = idx.CreateField("ssid", dbLong)
'
'
'    idx.Fields.Append f
'
'    tdef.Indexes.Append idx
    
    Call newField(tdef, "ssid", dbLong)
    Call newField(tdef, "title", dbText)
    Call newField(tdef, "author", dbText)
    Call newField(tdef, "pages", dbInteger)
    Call newField(tdef, "date", dbText)
    Call newField(tdef, "catalog", dbText)
        '.Fields.Append newField(tdef, "lib", dbText)
        '.Fields.Append newField(tdef, "link", dbText)
    
   dbase.TableDefs.Append tdef
      
   
   Set newTable = dbase.TableDefs(tableName).OpenRecordset
   
Do_Nothing:
End Function

Public Function newField(ByRef tdef As TableDef, ByRef Name As String, ByRef fType As DataTypeEnum) As Boolean
        newField = True
        On Error GoTo Error_newField
        Dim f As Field
        Set f = tdef.CreateField(Name, fType)
        tdef.Fields.Append f
        f.AllowZeroLength = True
        Exit Function
        
Error_newField:
    newField = False
End Function

Public Function l_CLong(ByRef str As String) As Long
    On Error GoTo Not_Long_Numeric
        l_CLong = 0
        l_CLong = CLng(str)
        Exit Function
Not_Long_Numeric:
End Function
Public Function l_CInt(ByRef str As String) As Integer
    On Error GoTo Not_int_Numeric
        l_CInt = 0
        l_CInt = CInt(str)
        Exit Function
Not_int_Numeric:
End Function
