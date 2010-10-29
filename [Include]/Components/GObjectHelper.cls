VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GObjectHelper"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Function GetBase(ByRef thisObject As IObject) As IObject
    
End Function

Public Function obNewObject(ByRef thisObject As IObject) As IObject
        
    Set obNewObject = CreateObject(thisObject.ClassName)
    
End Function

Public Function CObject(ByRef thisObject As IObject) As IObject
    Set CObject = thisObject
End Function


Public Function GetClassName(thisObject, baseObject) As String

End Function

Public Sub ContructObject(thisObject As IObject, baseObject As IObject)
    If IsObject(baseObject) Then baseObject.Initialize
    If IsObject(thisObject) Then thisObject.Initialize
End Sub

Public Sub DestructObject(thisObject, baseObject)

End Sub

Public Sub InitializeObject(ByRef thisObject As IObjectFirend)
    thisObject.Initialize
    Dim strBase As String
    Dim tempClass As CObject
    strBase = thisObject.getBaseClassName
    If (strBase <> "") Then
        Set CObject = CreateObject(strBase)
        InitializeObject (CObject)
        thisObject.setBaseClass CObject
    End If
    Set tempClass = thisObject
    tempClass
    thisObject.Initialize
End Sub

