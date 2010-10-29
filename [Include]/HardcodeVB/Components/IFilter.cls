VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Enum EChunkAction
    ecaAbort
    ecaSkip
    ecaTranslate
    ecaTranslateAll
End Enum

Property Get Source() As String
End Property
Property Let Source(sSourceA As String)
End Property

Property Get Target() As String
End Property
Property Let Target(sTargetA As String)
End Property

Function Translate(sChunkA As String, ByVal iChunkA As Long) As EChunkAction
End Function
'
