VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPathCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private mvarSCFolders As CStringCollection
Private mvarSCFiles As CStringCollection
Private mvarAll As CStringCollection

Public Function getString(Index As Long) As String

    getString = mvarAll.getString(Index)

End Function

Public Function getFolder(Index As Long) As String

    getFolder = mvarSCFolders.getString(Index)

End Function

Public Function getFile(Index As Long) As String

    getFile = mvarSCFiles.getString(Index)

End Function

Public Function getIndex(StrSearch As String, Optional cmp As VbCompareMethod = vbTextCompare) As Long

    getIndex = mvarAll.getIndex(StrSearch, cmp)

End Function

Public Function getFolderIndex(StrSearch As String, Optional cmp As VbCompareMethod = vbTextCompare) As Long

    getFolderIndex = mvarSCFolders.getIndex(StrSearch, cmp)

End Function

Public Function getFileIndex(StrSearch As String, Optional cmp As VbCompareMethod = vbTextCompare) As Long

    getFileIndex = mvarSCFiles.getIndex(StrSearch, cmp)

End Function

Public Property Get Count() As Long

    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.count
    Count = mvarAll.Count

End Property

Public Property Get FolderCount() As Long

    FolderCount = mvarSCFolders.Count

End Property

Public Property Get FileCount() As Long

    FileCount = mvarSCFiles.Count

End Property

Public Sub Add(strcoming As String)

    mvarAll.Add strcoming

    If Right$(strcoming, 1) = "\" Or Right$(strcoming, 1) = "/" Then
        mvarSCFolders.Add strcoming
    Else
        mvarSCFiles.Add strcoming
    End If

End Sub

Public Sub AddFolder(strcoming As String)

    mvarAll.Add strcoming
    mvarSCFolders.Add strcoming

End Sub

Public Sub AddFile(strcoming As String)

    mvarAll.Add strcoming
    mvarSCFiles.Add strcoming

End Sub

Public Sub clear()

    Set mvarAll = New CStringCollection
    Set mvarSCFolders = New CStringCollection
    Set mvarSCFiles = New CStringCollection

End Sub

Public Property Get HaveNextItem() As Boolean

    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.HaveNextItem
    HaveNextItem = mvarAll.HaveNextItem

End Property

Public Property Get HaveNextFolder() As Boolean

    HaveNextFolder = mvarSCFolders.HaveNextItem

End Property

Public Property Get HaveNextFile() As Boolean

    HaveNextFile = mvarSCFiles.HaveNextItem

End Property

Public Property Get NextFolder() As String

    NextFolder = mvarSCFolders.NextItem

End Property

Public Property Get NextFile() As String

    NextFile = mvarSCFiles.NextItem

End Property

Public Property Get NextItem() As String

    'used when retrieving value of a property, on the right side of an assignment.
    'Syntax: Debug.Print X.NextItem
    NextItem = mvarAll.NextItem

End Property

Public Sub StartEnum()

    mvarAll.StartEnum

End Sub

Public Sub startEnumFolder()

    mvarSCFolders.StartEnum

End Sub

Public Sub startEnumFile()

    mvarSCFiles.StartEnum

End Sub

Private Sub Class_Initialize()

    clear

End Sub

