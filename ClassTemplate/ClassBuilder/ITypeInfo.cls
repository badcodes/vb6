VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ITypeInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Function StyleOf(ByVal nId As Integer) As CTTypeStyles
End Function

Public Function NameOf(ByVal nId As Integer) As String
End Function



'Public Property Get Stabled() As Boolean
'End Property
'
'Public Property Let Stabled(f As Boolean)
'End Property

Public Function IdOf(ByRef sTemplate As String) As Integer
End Function

Public Function ConstVarOf(ByVal ctStyle As CTTypeStyles) As String
End Function

Public Property Get ConstTypePrefix() As String
End Property

Public Property Get ConstTypeSuffix() As String
End Property

Public Function GetAllName() As String()
End Function

Public Property Get Count() As Long
End Property
