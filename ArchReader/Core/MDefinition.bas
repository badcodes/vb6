Attribute VB_Name = "MDefinition"
Option Explicit

Public Type TListNode
    text As String
    Entry As String
End Type

Public Type TInfo
    Author As String
    Title As String
    Catalog As String
    Publisher As String
    Entry As String
End Type

Public Const PLUGIN_TYPE_READER As Long = 0
Public Const PLUGIN_TYPE_DOCUMENT As Long = 1
Public Const PLUGIN_TYPE_VIEWER As Long = 2


