Attribute VB_Name = "MDAO"
'CSEH: ErrMsgBox-xrlin
Option Explicit

Public Function getTabledefs(ByRef mdbFile As String, ByRef tables() As String) As Integer
    '<EhHeader>
    On Error GoTo getTabledefs_Err
    '</EhHeader>
    Dim db As DBEngine
    Dim dbase As Database
    Dim ts As TableDefs
    Dim t As TableDef
    
    Set db = New DBEngine
    Set dbase = db.openDatabase(mdbFile)
    Set ts = dbase.TableDefs
    'Erase tables()
    
    For Each t In ts
        If t.Attributes = 65536 Then
            getTabledefs = getTabledefs + 1
            ReDim Preserve tables(1 To getTabledefs) As String
            tables(getTabledefs) = t.Name
        End If
        
    Next

    Set ts = Nothing
    dbase.Close
    Set dbase = Nothing
    Set db = Nothing
    '<EhFooter>
    Exit Function

getTabledefs_Err:
    MsgBox Err.Description & vbCrLf & _
           "Reported by MDAO.getTabledefs "
    '</EhFooter>
End Function

Public Function newDatabase(ByRef mdbFile As String) As Database
    '<EhHeader>
    On Error GoTo newDatabase_Err
    '</EhHeader>

    Dim db As DBEngine
    Set db = New DBEngine
    
    Set newDatabase = db.CreateDatabase(mdbFile, dbLangChineseSimplified, dbVersion40)
    '<EhFooter>
    Exit Function

newDatabase_Err:
    MsgBox Err.Description & vbCrLf & _
           "Reported by MDAO.newDatabase "
    '</EhFooter>
End Function

'CSEH: ErrMsgBox-xrlin
Public Function openDatabase(ByRef mdbFile As String) As Database
    '<EhHeader>
    On Error GoTo openDatabase_Err
    '</EhHeader>
    Dim db As DBEngine
    Set db = New DBEngine
    Set openDatabase = db.openDatabase(mdbFile, , False)
    '<EhFooter>
    Exit Function

openDatabase_Err:
    MsgBox Err.Description & vbCrLf & _
           "Reported by MDAO.openDatabase "
    '</EhFooter>
End Function

Public Function getTable(ByRef dbase As Database, ByRef sName As String) As TableDef
    '<EhHeader>
    On Error GoTo getTable_Err
    '</EhHeader>
    Set getTable = dbase.TableDefs(sName)
    '<EhFooter>
    Exit Function

getTable_Err:
    MsgBox Err.Description & vbCrLf & _
           "Reported by MDAO.getTable "
    '</EhFooter>
End Function
Public Function newTable(ByRef dbase As Database, ByRef sName As String, Optional attr As TableDefAttributeEnum = dbAttachExclusive) As TableDef
    'Dim tdef As TableDef
    '<EhHeader>
    On Error GoTo newTable_Err
    '</EhHeader>
    Set newTable = dbase.CreateTableDef(sName, attr)

    'dbase.TableDefs.Append newTable
    '<EhFooter>
    Exit Function

newTable_Err:
    MsgBox Err.Description & vbCrLf & _
           "Reported by MDAO.newTable "
    '</EhFooter>
End Function

Public Function getRecord(ByRef tdef As TableDef) As Recordset
    '<EhHeader>
    On Error GoTo getRecord_Err
    '</EhHeader>
    Set getRecord = tdef.OpenRecordset()
    '<EhFooter>
    Exit Function

getRecord_Err:
    MsgBox Err.Description & vbCrLf & _
           "Reported by MDAO.getRecord "
    '</EhFooter>
End Function

Public Function getField(ByRef rc As Recordset, ByRef sName As String) As Field
    '<EhHeader>
    On Error GoTo getField_Err
    '</EhHeader>
    Set getField = rc.Fields(sName)
    '<EhFooter>
    Exit Function

getField_Err:
    MsgBox Err.Description & vbCrLf & _
           "Reported by MDAO.getField "
    '</EhFooter>
End Function

Public Function addField(ByRef tdef As TableDef, ByRef sName As String, Optional ByRef fType As DataTypeEnum = dbText) As Field
    '<EhHeader>
    On Error GoTo addField_Err
    '</EhHeader>
    Set addField = tdef.CreateField(sName, fType)
    tdef.Fields.Append addField
    addField.AllowZeroLength = True
    '<EhFooter>
    Exit Function

addField_Err:
    MsgBox Err.Description & vbCrLf & _
           "Reported by MDAO.addField "
    '</EhFooter>
End Function
Public Function addTable(ByRef dbase As Database, ByRef tdef As TableDef) As TableDef
    dbase.TableDefs.Append tdef
    Set addTable = tdef
End Function
Public Function test()
    Dim dbase As Database
    Dim rc As Recordset
    Dim tdef As TableDef
    Set dbase = newDatabase("c:\test.mdb")
    Set tdef = newTable(dbase, "TEST")
    addField tdef, "test"
    addField tdef, "Something"
    dbase.TableDefs.Append tdef
    
End Function
