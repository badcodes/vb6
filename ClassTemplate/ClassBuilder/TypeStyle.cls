VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTypeStyle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum CTTypeStyles
    CTTypeUnknown = 0
    CTTypeNormal = 1
    CTTypeObject = 2
    CTTypeVariant = 4
End Enum

Private Const CONST_NAME_TypeUnknown As String = "VariantType"
Private Const CONST_NAME_TypeNormal As String = "NormalType"
Private Const CONST_NAME_TypeObject As String = "ObjectType"
Private Const CONST_NAME_TypeVariant As String = "VariantType"

Private mNormal As String
Private mObject As String
Private mVariant As String

Public Function IsObject(ByVal ctStyle As CTTypeStyles) As Boolean
    IsObject = (ctStyle = CTTypeObject)
End Function

Public Function IsNormal(ByVal ctStyle As CTTypeStyles) As Boolean
    IsNormal = (ctStyle = CTTypeNormal)
End Function

Public Function IsVariant(ByVal ctStyle As CTTypeStyles) As Boolean
    IsVariant = (ctStyle <> CTTypeNormal Or ctStyle <> CTTypeObject)
End Function

Public Property Let ConstVarOf(ByVal ctStyle As CTTypeStyles, ByRef sName As String)
    If ctStyle = CTTypeNormal Then
        mNormal = sName
    ElseIf ctStyle = CTTypeObject Then
        mObject = sName
    Else
        mVariant = sName
    End If
End Property

Private Sub Class_Initialize()
    mNormal = CONST_NAME_TypeNormal
    mObject = CONST_NAME_TypeObject
    mVariant = CONST_NAME_TypeVariant
End Sub

Public Property Get ConstVarOf(ByVal ctStyle As CTTypeStyles) As String
    If ctStyle = CTTypeNormal Then
        ConstVarOf = mNormal
    ElseIf ctStyle = CTTypeObject Then
        ConstVarOf = mObject
    Else
        ConstVarOf = mVariant
    End If
End Property
