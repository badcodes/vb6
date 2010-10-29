VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLink"
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

#Const fVariant = VariantType


'================================================================================

Public Data As Variant 
Public NextLink As CLink 

Public Function AppendNext(ByRef vData As Variant) As CLink 
    Dim newLink As CLink 
    Set newLink = New CLink 
    Set NextLink = newLink
    If IsObject(vData)  then Set newLink.Data = vData Else newLink.Data = vData 
    Set AppendNext = newLink
End Function

