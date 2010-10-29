VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CObjectLink"
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

#Const fObject = ObjectType


'================================================================================

Public Data As Object 
Public NextLink As CObjectLink 

Public Function AppendNext(ByRef vData As Object) As CObjectLink 
    Dim newLink As CObjectLink 
    Set newLink = New CObjectLink 
    Set NextLink = newLink
    Set newLink.Data = vData 
    Set AppendNext = newLink
End Function

