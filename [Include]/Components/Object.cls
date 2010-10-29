VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Implements IObjectFirend
Implements IObject

Private mClassName As String
Private mBaseObject As IObject


Private Property Get IObject_ClassName() As String
    IObject_ClassName = mClassName
End Property

Private Sub IObject_Destory()

End Sub

Private Sub IObject_DestructorVoid()
End Sub

Private Sub IObject_Initialize()
    mClassName = "CObject"
End Sub

Private Function IObjectFirend_getBaseClassName() As String

End Function

Private Sub IObjectFirend_Initialize()

End Sub

Private Sub IObjectFirend_setBaseClass(baseClass As CObject)

End Sub

Private Sub IObjectFirend_setClassName(Str As String)
    mClassName = Str
End Sub

