VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'--------------------------------------------------------------------------------
'    Component  : CSetting
'
'    Description: Setting Class
'
'    Author   : xrLin
'
'    Date     : 2005-2006
'--------------------------------------------------------------------------------
Option Explicit

Public Enum csSettingFlag
    SF_LISTTEXT = 1
    SF_POSITION = 2
    SF_MENUARRAY = 3
    SF_FORM = 4
    SF_FONT = 5
    SF_COLOR = 6
    SF_Tag = 7
    SF_CHECKED = 8
    SF_CAPTION = 9
    SF_TEXT = 10
    SF_WIDTH = 11
    SF_HEIGHT = 12
    SF_VALUE = 13
End Enum

Private hIni As CLiNInI
Public Property Get iniFile() As String
iniFile = hIni.File
End Property
Public Property Let iniFile(Value As String)
hIni.File = Value
End Property
Public Sub Load(obj As Object, af As csSettingFlag, Optional secName As String = "")

    'Dim secName As String
    Dim iKeyCount As Long
    Dim iIndex As Long
    Dim sKeyValue As String
    Dim iKeyValue As String
    'Dim fTestCIni As Boolean
    
    If iniFile = "" Then Exit Sub
    
    On Error Resume Next
    
    If secName = "" Then
    
        secName = obj.Name
        If secName = "" Then secName = obj.Caption
        If secName = "" Then secName = obj.Text
        If secName = "" Then secName = obj.Tag
        If secName = "" Then secName = TypeName(obj)
        secName = secName & CStr(obj.Index)
        secName = obj.Parent.Name & "." & secName
        secName = secName & "_" & CStr(af)
        
    End If
    
    Select Case af
    Case SF_LISTTEXT
                iKeyCount = hIni.GetSetting(secName, "Count")
                For iIndex = 1 To iKeyCount
                    sKeyValue = hIni.GetSetting(secName, "LISTTEXT" & str$(iIndex))
                    If sKeyValue <> "" Then obj.AddItem sKeyValue
                Next
    Case SF_POSITION
                For iIndex = 1 To 4
                iKeyValue = CLng(Val(hIni.GetSetting(secName, "POSITION" & str$(iIndex))))
                If iKeyValue > 0 Then
                    If iIndex = 1 Then obj.Left = iKeyValue
                    If iIndex = 2 Then obj.Top = iKeyValue
                    If iIndex = 3 Then obj.Height = iKeyValue
                    If iIndex = 4 Then obj.Width = iKeyValue
                End If
                Next
    Case SF_FORM
                For iIndex = 1 To 4
                iKeyValue = CLng(Val(hIni.GetSetting(secName, "POSITION" & str$(iIndex))))
                If iKeyValue <> 0 Then
                    If iIndex = 1 Then obj.Left = iKeyValue
                    If iIndex = 2 Then obj.Top = iKeyValue
                    If iIndex = 3 Then obj.Height = iKeyValue
                    If iIndex = 4 Then obj.Width = iKeyValue
                End If
                Next
                obj.WindowState = Val(hIni.GetSetting(secName, "WINDOWSTATE"))
    Case SF_MENUARRAY
                iKeyCount = hIni.GetSetting(secName, "Count")
                For iIndex = obj.Count To iKeyCount - 1
                    VB.Load obj(iIndex)
                    obj(iIndex).Visible = True
                Next
                For iIndex = 1 To iKeyCount
                    obj(iIndex - 1).Caption = hIni.GetSetting(secName, "CAPTION" & str$(iIndex))
                    obj(iIndex - 1).Tag = hIni.GetSetting(secName, "TAG" & str$(iIndex))
                    'obj(iIndex - 1).Checked = CBoolStr(hINi.GetSetting(secName, "CHECKED" & Str$(iIndex)))
               Next
    Case SF_CHECKED
           obj.Checked = Val(hIni.GetSetting(secName, "CHECKED")) <> 0
           'obj.Tag = hINI.GetSetting(secName, "TAG")
    Case SF_Tag
            obj.Tag = hIni.GetSetting(secName, "TAG")
    Case SF_COLOR
            obj.ForeColor = Val(hIni.GetSetting(secName, "FORECOLOR"))
            obj.BackColor = Val(hIni.GetSetting(secName, "BACKCOLOR"))
    Case SF_FONT
            With obj.Font
                .Bold = (Val(hIni.GetSetting(secName, "Bold")) <> 0)
                .Italic = (Val(hIni.GetSetting(secName, "Italic")) <> 0)
                .Underline = (Val(hIni.GetSetting(secName, "Underline")) <> 0)
                .Strikethrough = (Val(hIni.GetSetting(secName, "Strikethrough")) <> 0)
                sKeyValue = hIni.GetSetting(secName, "Name")
                If sKeyValue <> "" Then .Name = sKeyValue
                iKeyValue = Val(hIni.GetSetting(secName, "Size"))
                If iKeyValue > 0 Then .Size = iKeyValue
            End With
    Case SF_TEXT
        sKeyValue = hIni.GetSetting(secName, "Text")
        If sKeyValue <> "" Then obj.Text = sKeyValue
        
    Case SF_WIDTH
        obj.Width = hIni.GetSetting(secName, "Width")
    Case SF_HEIGHT
        obj.Height = hIni.GetSetting(secName, "Height")
    Case SF_VALUE
        obj.Value = hIni.GetSetting(secName, "Value")
    Case SF_CAPTION
        sKeyValue = hIni.GetSetting(secName, "Caption")
        If sKeyValue <> "" Then obj.Caption = sKeyValue
    End Select

End Sub

Private Sub Class_Initialize()
Set hIni = New CLiNInI
End Sub

Private Sub Class_Terminate()
On Error Resume Next
If iniFile <> "" Then hIni.Save
Set hIni = Nothing
End Sub

Public Sub Save(obj As Object, af As csSettingFlag, Optional secName As String = "")

    'Dim secName As String
    Dim iKeyCount As Long
    Dim iIndex As Long
    Dim sKeyValue As String
    'Dim iKeyValue As String
    
    If iniFile = "" Then Exit Sub
    
    On Error Resume Next
    If secName = "" Then
        'secName = ""
        secName = obj.Name
        If secName = "" Then secName = obj.Caption
        If secName = "" Then secName = obj.Text
        If secName = "" Then secName = obj.Tag
        If secName = "" Then secName = TypeName(obj)
        secName = secName & CStr(obj.Index)
        secName = obj.Parent.Name & "." & secName
        secName = secName & "_" & CStr(af)
    End If
        hIni.DeleteSection secName
        
        Debug.Print secName
        Select Case af
        Case SF_LISTTEXT
                    iKeyCount = obj.ListCount
                    hIni.SaveSetting secName, "Count", iKeyCount
                    For iIndex = 1 To iKeyCount
                        sKeyValue = obj.List(iIndex - 1)
                        hIni.SaveSetting secName, "LISTTEXT" & str$(iIndex), sKeyValue
                    Next
        Case SF_POSITION, SF_FORM
                    hIni.SaveSetting secName, "POSITION" & str$(1), obj.Left
                    hIni.SaveSetting secName, "POSITION" & str$(2), obj.Top
                    hIni.SaveSetting secName, "POSITION" & str$(3), obj.Height
                    hIni.SaveSetting secName, "POSITION" & str$(4), obj.Width
                    hIni.SaveSetting secName, "WINDOWSTATE", obj.WindowState
        Case SF_MENUARRAY
                    iKeyCount = obj.Count
                    hIni.SaveSetting secName, "Count", iKeyCount
                    For iIndex = 1 To iKeyCount
                    hIni.SaveSetting secName, "CAPTION" & str$(iIndex), obj(iIndex - 1).Caption
                    hIni.SaveSetting secName, "TAG" & str$(iIndex), obj(iIndex - 1).Tag
                    hIni.SaveSetting secName, "CHECKED" & str$(iIndex), CLng(obj(iIndex - 1).Checked)
                   Next
        Case SF_CHECKED
                    hIni.SaveSetting secName, "CHECKED", CLng(obj.Checked)
        Case SF_Tag
                    hIni.SaveSetting secName, "TAG", obj.Tag
        Case SF_COLOR
                hIni.SaveSetting secName, "FORECOLOR", obj.ForeColor
                hIni.SaveSetting secName, "BACKCOLOR", obj.BackColor
        Case SF_FONT
                With obj.Font
                    hIni.SaveSetting secName, "Bold", CInt(.Bold)
                    hIni.SaveSetting secName, "Italic", CInt(.Italic)
                    hIni.SaveSetting secName, "Underline", CInt(.Underline)
                    hIni.SaveSetting secName, "Strikethrough", CInt(.Strikethrough)
                    hIni.SaveSetting secName, "Size", CInt(.Size)
                    hIni.SaveSetting secName, "Name", .Name
                End With
        Case SF_TEXT
             hIni.SaveSetting secName, "Text", obj.Text
        Case SF_WIDTH
            hIni.SaveSetting secName, "Width", obj.Width
        Case SF_HEIGHT
            hIni.SaveSetting secName, "Height", obj.Height
        Case SF_VALUE
            hIni.SaveSetting secName, "Value", obj.Value
        Case SF_CAPTION
            Debug.Print obj.Caption
            hIni.SaveSetting secName, "Caption", obj.Caption
        End Select
   
    'Set obj = Nothing

End Sub

