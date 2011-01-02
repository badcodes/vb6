VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit



Private mFilename As String
Private mInfo As TInfo
Private mEntryList() As TListNode
Private mContentList() As TListNode
Private mHandler As IDocumentHandler

Public Property Set Handler(ByRef vHnd As IDocumentHandler)
    Set mHandler = vHnd
End Property

Public Function OpenEntry(ByRef vEntry As String) As String

End Function


