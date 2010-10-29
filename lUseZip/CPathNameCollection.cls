VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPathNameCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'local variable(s) to hold property value(s)
Private mvarcount As Long 'local copy
Private mvarFolderCount As Long
Private mvarFileCount As Long
Private mvarFolderIndexs() As Long
Private mvarFileIndexs() As Long
Public Function getString(Index As Long) As String
End Function
Public Function getIndex(StrSearch As String, Optional cmp As VbCompareMethod = vbTextCompare) As Long
End Function
Public Property Get count() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.count
    count = mvarcount
End Property
Public Sub add(strComing As String)
End Sub
Public Sub clear()
End Sub
Public Function getLength(Index As Long) As Long
End Function

