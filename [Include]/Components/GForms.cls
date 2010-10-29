VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GForms"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Const c_w_space As Long = 120
Public Const c_h_space As Long = 240
Public Enum E_AlignPos
    E_AP_LEFT
    E_AP_RIGHT
    E_AP_TOP
    E_AP_BOTTOM
End Enum

Public Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExW" ( _
    ByVal dwExStyle As Long, _
    ByVal lpClassName As String, _
    ByVal lpWindowName As String, _
    ByVal dwStyle As Long, _
    ByVal x As Long, _
    ByVal y As Long, _
    ByVal nWidth As Long, _
    ByVal nHeight As Long, _
    ByVal hWndParent As Long, _
    ByVal hMenu As Long, _
    ByVal hInstance As Long, _
    lpParam As Any) As Long


Public Const CST_FORM_FLAGS_NORESETING As String = "NoReseting"

Private Function testObj(ByRef obj As Object) As Boolean
    testObj = False
    If obj Is Nothing Then Exit Function
    If obj.Parent Is Nothing Then Exit Function
    testObj = True
End Function

Public Sub move_TopLeft(ByRef obj As Object)
    If testObj(obj) = False Then Exit Sub
    On Error Resume Next
    obj.Top = c_w_space
    obj.Left = c_w_space
End Sub
Public Sub move_Right(ByRef obj As Object)
    If testObj(obj) = False Then Exit Sub
    On Error Resume Next
    obj.Left = obj.Parent.ScaleWidth - obj.Width - c_w_space
End Sub
Public Sub move_LeftTo(ByRef objRight As Object, ByRef objLeft As Object)
    If testObj(objLeft) = False Then Exit Sub
    If testObj(objRight) = False Then Exit Sub
        With objRight
            objLeft.Top = .Top
            objLeft.Left = .Left - c_w_space - objLeft.Width
        End With
End Sub
Public Sub move_RightTo(ByRef objLeft As Object, ByRef objRight As Object)
    If testObj(objLeft) = False Then Exit Sub
    If testObj(objRight) = False Then Exit Sub
    On Error Resume Next
    With objRight
        .Top = objLeft.Top
        .Left = objLeft.Left + objLeft.Width + c_w_space
    End With
End Sub
Public Sub move_Below(ByRef objUpper As Object, ByRef objLower As Object)
    If testObj(objUpper) = False Then Exit Sub
    If testObj(objLower) = False Then Exit Sub
    'On Error Resume Next
    With objUpper
        objLower.Top = .Top + .Height + c_h_space
        'objLower.Left = .Left
    End With
End Sub

Public Sub move_Align(ByRef objStand As Object, ByRef objMove As Object, Optional ByRef tAS As E_AlignPos = E_AP_TOP)
    If testObj(objStand) = False Then Exit Sub
    If testObj(objMove) = False Then Exit Sub
    On Error Resume Next
    Select Case tAS
    Case E_AP_LEFT
        objMove.Left = objStand.Left
    Case E_AP_RIGHT
        objMove.Left = objStand.Left + objStand.Width - objMove.Width
    Case E_AP_TOP
        objMove.Top = objStand.Top
    Case E_AP_BOTTOM
        objMove.Top = objStand.Top + objStand.Height - objMove.Height
    End Select
End Sub

Public Sub ResetForm(ByRef vForm As Form)
    On Error Resume Next
    Dim ctl As Control
    For Each ctl In vForm.Controls
        If Not ctl.Tag = CST_FORM_FLAGS_NORESETING Then
            Select Case TypeName(ctl)
                Case "TextBox"
                    ctl.Text = ""
                Case "CheckBox"
                    ctl.value = 0
                    ctl.Checked = False
                Case "ComboBox"
                    ctl.Clear
                    ctl.Text = ""
            End Select
        End If
    Next
End Sub

Public Function ControlStateToString(ByRef vControl As Control) As String
    On Error Resume Next
    ControlStateToString = ControlStateToString & vControl.Left
    ControlStateToString = ControlStateToString & "|"
    ControlStateToString = ControlStateToString & vControl.Top
    ControlStateToString = ControlStateToString & "|"
    ControlStateToString = ControlStateToString & vControl.Width
    ControlStateToString = ControlStateToString & "|"
    ControlStateToString = ControlStateToString & vControl.Height
End Function

Public Sub ControlStateFromString(ByRef vControl As Control, ByRef vString As String)
    On Error Resume Next
    Dim u As Long
    Dim strs() As String
    strs = Split(vString, "|")
    u = UBound(strs)
    If u < 3 Then Exit Sub
    vControl.Move CSng(strs(0)), CSng(strs(1)), CSng(strs(2)), CSng(strs(3))


    
End Sub


Public Function FormStateToString(ByRef vForm As Form) As String
    On Error Resume Next
    FormStateToString = vForm.WindowState
    FormStateToString = FormStateToString & "|"
    FormStateToString = FormStateToString & vForm.Left
    FormStateToString = FormStateToString & "|"
    FormStateToString = FormStateToString & vForm.Top
    FormStateToString = FormStateToString & "|"
    FormStateToString = FormStateToString & vForm.Width
    FormStateToString = FormStateToString & "|"
    FormStateToString = FormStateToString & vForm.Height
End Function

Public Sub FormStateFromString(ByRef vForm As Form, ByRef vString As String)
    On Error Resume Next
    Dim u As Long
    Dim strs() As String
    strs = Split(vString, "|")
    u = UBound(strs)
    If u < 4 Then Exit Sub
    vForm.WindowState = CInt(strs(0))
    If strs(0) = CStr(vbNormal) Then
        vForm.Move CSng(strs(1)), CSng(strs(2)), CSng(strs(3)), CSng(strs(4))
    Else
    End If
    
End Sub

Public Function ComboxItemsToString(vCombox As Control) As String
    Dim i As Long
    Dim c As Long
    c = vCombox.ListCount
    Do While i < c
        ComboxItemsToString = ComboxItemsToString & vbBack & vCombox.List(i)
        i = i + 1
    Loop
    If ComboxItemsToString <> "" Then ComboxItemsToString = Mid$(ComboxItemsToString, Len(vbBack) + 1)
End Function

Public Sub ComboxItemsFromString(vCombox As Control, ByRef vString As String)
    On Error Resume Next
    Dim pStrings() As String
    pStrings = Split(vString, vbBack)
    Dim i As Long, u As Long
    u = UBound(pStrings())
    For i = 0 To u
        vCombox.AddItem pStrings(i)
    Next
    vCombox.ListIndex = u
End Sub
Public Sub CopyPosition( _
    ByRef vDest As Control, _
    ByRef vSource As Control, _
    Optional vLeft As Boolean = True, _
    Optional vTop As Boolean = True, _
    Optional vWidth As Boolean = True, _
    Optional vHeight As Boolean = True)
    If vDest Is Nothing Then Exit Sub
    If vSource Is Nothing Then Exit Sub
    On Error Resume Next
    If vLeft Then vDest.Left = vSource.Left
    If vTop Then vDest.Top = vSource.Top
    If vHeight Then vDest.Height = vSource.Height
    If vWidth Then vDest.Width = vSource.Width
End Sub

