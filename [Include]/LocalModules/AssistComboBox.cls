VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CAssistComboBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private mCmbBox As ComboBox
Private secName As String
Public SaveText As Boolean
Private Const className = "AssistComboBox"

Public Sub Attach(cmbBox As ComboBox)

    Set mCmbBox = cmbBox
    Dim iListCount As Long
    Dim iStart As Long
    Dim sText As String
    secName = App.ProductName & "->" & mCmbBox.Parent.name & "->" & mCmbBox.name
    iListCount = CLng(Val(GetSetting(className, secName, "Count")))

    For iStart = 1 To iListCount
        sText = GetSetting(className, secName, "Text" & Str$(iStart))

        If sText <> "" Then mCmbBox.AddItem sText
    Next

End Sub

Public Sub Clean()

    On Error Resume Next

    If secName = "" Then Exit Sub
    DeleteSetting className, secName

End Sub

Public Sub Unistall()

    On Error Resume Next
    DeleteSetting className

End Sub

Private Sub Class_Initialize()

    SaveText = True

End Sub

Private Sub Class_Terminate()

    If SaveText Then
        Clean
        Dim iIndex As Long
        Dim iEnd As Long
        iEnd = mCmbBox.ListCount

        If iEnd > 0 Then SaveSetting className, secName, "Count", Str$(iEnd)
        'iEnd = iEnd - 1

        For iIndex = 1 To iEnd
            SaveSetting className, secName, "Text" & Str$(iIndex), mCmbBox.List(iIndex - 1)
        Next

    End If

    Set mCmbBox = Nothing

End Sub



