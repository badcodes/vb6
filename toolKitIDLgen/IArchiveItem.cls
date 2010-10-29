VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IArchiveItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Property Get Crc() As Long
End Property

Public Property Get CreatedDate() As Date
End Property

Public Property Get Encrypted() As Boolean
End Property

Public Property Get Factor() As Long
End Property

Public Property Get Filename() As String
End Property

Public Property Get FileType() As VbFileAttribute
End Property

Public Property Get Method() As String
End Property

Public Property Get PackedSize() As Long
End Property

Public Property Get Size() As Long
End Property

Friend Property Let Crc(ByVal lValue As Long)
End Property

Friend Property Let CreatedDate(ByVal dValue As Date)
End Property

Friend Property Let Encrypted(ByVal bValue As Boolean)
End Property

Friend Property Let Factor(ByVal lValue As Long)
End Property

Friend Property Let Filename(ByVal sValue As String)
End Property

Friend Property Let FileType(ByVal vbfaValue As VbFileAttribute)
End Property

Friend Property Let Method(ByVal sValue As String)
End Property

Friend Property Let PackedSize(ByVal lValue As Long)
End Property

Friend Property Let Size(ByVal lValue As Long)
End Property

