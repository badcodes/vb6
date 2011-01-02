VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IArchiveItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' The private collection object
'Private m_IArchiveItems As Collection

' This sub adds a new IArchiveItem item to the collection.
Friend Sub Add(Item As IArchiveItem, Optional Key As String)
End Sub

' This function returns a IArchiveItem item from the collection. It's the default method.
Function ItembyIndex(Index As Long) As IArchiveItem
Attribute ItembyIndex.VB_UserMemId = 0
End Function

Function ItembyKey(ByVal bValue As String) As IArchiveItem
End Function


' This property returns the number of items in the collection.
Property Get Count() As Long
End Property


' This sub remove all items from the collection.
Friend Sub Clear()
End Sub


' This function adds "For Each" enumeration support. Must have a -4 DispID.
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
End Function



