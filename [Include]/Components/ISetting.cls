VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ISetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'--------------------------------------------------------------------------------
'    Component  : CLiNInI
'
'    Description: ini File Class
'
'    Author   : xrLin
'
'    Date     : 2005-2006
'--------------------------------------------------------------------------------
Public Property Get Source() As String
End Property

Public Property Let Source(ByVal StrValue As String)
End Property

Public Property Get CompareMethod() As VbCompareMethod
    CompareMethod = mCompareMethod
End Property

Public Property Let CompareMethod(ByVal VBCValue As VbCompareMethod)
    mCompareMethod = VBCValue
End Property



Public Function ToString() As String

End Function


Public Sub Save()
End Sub
Public Sub WriteTo(ByRef vTarget As String)
   
End Sub


Public Function GetArrayText(ByRef vSection As String, ByRef vKey As String, ByRef vText() As String) As Long

End Function
Public Sub SaveArrayText(ByRef vSection As String, ByRef vKey As String, ByRef vText() As String)
   
End Sub

'
Public Sub Reset()

End Sub

Public Sub DeleteSection(ByRef vSection As String)

    

End Sub
Public Sub DeleteSetting(ByRef vSection As String, ByVal vKey As String)
    
End Sub
Public Sub SaveSetting(ByRef vSection As String, ByRef vKey As String, ByRef vValue As String)
    
End Sub

Public Function GetSetting(ByRef vSection As String, ByRef vKey As String) As String
    
End Function
Public Sub ChangeSectionName(ByRef vSection As String, ByRef vNewSection As String)
    
    
End Sub
Public Sub ChangeKeyName(ByRef vSection As String, ByRef vKey As String, ByRef vNewKey As String)
    
End Sub

Public Function GetSection(ByRef vSection As String, ByRef vKeyValue() As String) As Long
End Function

Public Sub SaveSection(ByRef vSection As String, ByRef vKeyValue() As String)
End Sub





'Private Sub ISetting_ChangeKeyName(vSection As String, vKey As String, vNewKey As String)
'    ChangeKeyName vSection, vKey, vNewKey
'End Sub
'
'Private Sub ISetting_ChangeSectionName(vSection As String, vNewSection As String)
'    ChangeSectionName vSection, vNewSection
'End Sub
'
'Private Property Get ISetting_CompareMethod() As VbCompareMethod
'    ISetting_CompareMethod = CompareMethod
'End Property
'
'Private Property Let ISetting_CompareMethod(ByVal RHS As VbCompareMethod)
'    CompareMethod = RHS
'End Property
'
'Private Sub ISetting_DeleteSection(vSection As String)
'    DeleteSection vSection
'End Sub
'
'Private Sub ISetting_DeleteSetting(vSection As String, ByVal vKey As String)
'    DeleteSetting vSection, vKey
'End Sub
'
'Private Function ISetting_GetArrayText(vSection As String, vKey As String, vText() As String) As Long
'    ISetting_GetArrayText = GetArrayText(vSection, vKey, vText())
'End Function
'
'Private Function ISetting_GetSection(vSection As String, vKeyValue() As String) As Long
' ISetting_GetSection = GetSection(vSection, vKeyValue)
'End Function
'
'Private Function ISetting_GetSetting(vSection As String, vKey As String) As String
' ISetting_GetSetting = GetSetting(vSection, vKey)
'End Function
'
'Private Sub ISetting_Reset()
'    Reset
'End Sub
'
'Private Sub ISetting_Save()
'    Save
'End Sub
'
'Private Sub ISetting_SaveArrayText(vSection As String, vKey As String, vText() As String)
' SaveArrayText vSection, vKey, vText
'End Sub
'
'Private Sub ISetting_SaveSection(vSection As String, vKeyValue() As String)
' SaveSection vSection, vKeyValue
'End Sub
'
'Private Sub ISetting_SaveSetting(vSection As String, vKey As String, vValue As String)
'    SaveSetting vSection, vKey, vValue
'End Sub
'
'Private Property Let ISetting_Source(ByVal RHS As String)
'    Source = RHS
'End Property
'
'Private Property Get ISetting_Source() As String
'    ISetting_Source = Source
'End Property
'
'Private Function ISetting_ToString() As String
'    ISetting_ToString = ToString
'End Function
'
'Private Sub ISetting_WriteTo(vTarget As String)
'    WriteTo vTarget
'End Sub


