VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CIntegerLink"
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

#Const fInteger = NormalType


'================================================================================

Public Data As Integer 
Public NextLink As CIntegerLink 

Public Function AppendNext(ByRef vData As Integer) As CIntegerLink 
    Dim newLink As CIntegerLink 
    Set newLink = New CIntegerLink 
    Set NextLink = newLink
    newLink.Data = vData 
    Set AppendNext = newLink
End Function

