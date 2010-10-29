VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CAutoSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'--------------------------------------------------------------------------------
'    Component  : CAutoSetting
'
'    Description: AutoSetting Class
'
'    Author   : xrLin
'
'    Date     : 2005-2006
'--------------------------------------------------------------------------------
Option Explicit

Private hSetting As CSetting
Private Type settingItem
obj As Object
flag As csSettingFlag
secName As String
ReadOnly As Boolean
End Type
Private itemCol() As settingItem
Private itemCount As Long
Private Const autoIni = "autoini.ini"
'local variable(s) to hold property value(s)
Private mvarfileNameSaveTo As String 'local copy


Public Property Let fileNameSaveTo(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.fileNameSaveTo = 5
    mvarfileNameSaveTo = vData

End Property

Public Property Get fileNameSaveTo() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.fileNameSaveTo
    fileNameSaveTo = mvarfileNameSaveTo
End Property

Public Function Add(ByRef obj As Object, ByRef flag As csSettingFlag, Optional secName As String = "", Optional ReadOnly As Boolean = False)
On Error Resume Next
If hSetting.iniFile = "" Then hSetting.iniFile = Me.fileNameSaveTo
itemCount = itemCount + 1
ReDim Preserve itemCol(1 To itemCount) As settingItem
Set itemCol(itemCount).obj = obj
itemCol(itemCount).flag = flag
itemCol(itemCount).secName = secName
itemCol(itemCount).ReadOnly = ReadOnly
hSetting.Load obj, flag, secName
End Function
Private Sub Class_Initialize()
Set hSetting = New CSetting
Me.fileNameSaveTo = App.Path & "\" & autoIni
End Sub
Private Sub Class_Terminate()
Dim i As Long
On Error Resume Next
For i = 1 To itemCount
    If itemCol(i).ReadOnly = False Then
        hSetting.Save itemCol(i).obj, itemCol(i).flag, itemCol(i).secName
    End If
    Set itemCol(i).obj = Nothing
Next
Set hSetting = Nothing
End Sub

