VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CRegItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum EErrorRegItem
    eeBaseRegItem = 13160   ' CRegItem
End Enum

Private node As CRegNode
Private vValue As Variant
Private sName As String

' Create/Destroy Value

Sub Create(nodeA As CRegNode, vIndex As Variant)
    Set node = nodeA
    Dim e As Long
    Select Case VarType(vIndex)
    Case vbString
        sName = vIndex
        e = MRegTool.GetRegValue(node.Key, CStr(vIndex), vValue)
                
    Case vbInteger, vbLong
        e = MRegTool.GetRegValueNext(node.Key, CLng(vIndex), sName, vValue)
        
    Case Else
        ApiRaiseIf ERROR_INVALID_DATA
        
    End Select
    ApiRaiseIf e
End Sub

Sub Destroy()
    ' Nothing to do
End Sub

' Friend properties to make data structure accessible to walker
Friend Property Get RegItems(i As Long) As CRegItem
    Set RegItems = node.Items(i)
End Property

' NewEnum must have the procedure ID -4 in Procedure Attributes dialog
' Create a new data walker object and connect to it
Public Function NewEnum() As IEnumVARIANT
Attribute NewEnum.VB_UserMemId = -4
    ' Create a new iterator object
    Dim walker As CRegItemWalker
    Set walker = New CRegItemWalker
    ' Connect it with collection data
    walker.Attach Me
    ' Return it
    Set NewEnum = walker.NewEnum
End Function

Property Get Count() As Long
    Count = node.ItemCount
End Property

' Default property
'!Public Property Get Item(vIndex As Variant) '! As DataType
    '!Set Item = data(vIndex)
'!End Property
' Get/Set Value
'RegQueryValue
'RegQueryValueEx
'RegSetValue
'RegSetValueEx

Property Get Name() As String
    Name = sName
End Property

Property Get Value() As Variant
Attribute Value.VB_UserMemId = 0
    Value = vValue
End Property

Property Let Value(vValueA As Variant)
    Dim ordType As Long, e As Long
    Select Case VarType(vValueA)
    Case vbString
        ordType = REG_SZ
    Case vbInteger, vbLong
        ordType = REG_DWORD
    Case Else
    End Select
    e = MRegTool.CreateRegValue(vValueA, node.Key, sName)
    ApiRaiseIf e
    e = RegFlushKey(node.Key)
    ApiRaiseIf e
End Property

#If fComponent = 0 Then
Private Sub ErrRaise(e As Long)
    Dim sText As String, sSource As String
    If e > 1000 Then
        sSource = App.ExeName & ".RegItem"
        Select Case e
        Case eeBaseRegItem
            BugAssert True
       ' Case ee...
       '     Add additional errors
        End Select
        Err.Raise COMError(e), sSource, sText
    Else
        ' Raise standard Visual Basic error
        sSource = App.ExeName & ".VBError"
        Err.Raise e, sSource
    End If
End Sub
#End If




