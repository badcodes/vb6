VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TLink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Template header:
'================================================================================

#Const ObjectType = 1
#Const NormalType = 2
#Const VariantType = (ObjectType Or NormalType)
#Const DefaultType = VariantType

#Const fTPLAType = DefaultType


'================================================================================

Public Data As TPLAType
Public NextLink As TLink

Public Function AppendNext(ByRef vData As TPLAType) As TLink
    Dim newLink As TLink
    Set newLink = New TLink
    Set NextLink = newLink
    MTemplateHelper.Assign "TPLAType", newLink.Data, vData
    Set AppendNext = newLink
End Function

