Attribute VB_Name = "MAppSetting"
Option Explicit

Public Sub AppSetting_SaveAllTextBox(ByRef vIniHnd As CLiNInI, ByRef vForm As Form)
    If vForm Is Nothing Then Exit Sub
    If vIniHnd Is Nothing Then Exit Sub
    Dim ctl As Control
    For Each ctl In vForm
        If LCase$(TypeName(ctl)) = "textbox" Then
            AppSetting_SaveTextBox vIniHnd, ctl
        End If
    Next
End Sub
Public Sub AppSetting_LoadAllTextBox(ByRef vIniHnd As CLiNInI, ByRef vForm As Form)
    If vForm Is Nothing Then Exit Sub
    If vIniHnd Is Nothing Then Exit Sub
    Dim ctl As Control
    For Each ctl In vForm
        If LCase$(TypeName(ctl)) = "textbox" Then
            AppSetting_LoadTextBox vIniHnd, ctl
        End If
    Next
End Sub
Public Sub AppSetting_SaveTextBox(ByRef vIniHnd As CLiNInI, ByRef vTextBox As TextBox)

    If vTextBox Is Nothing Then Exit Sub
    If vIniHnd Is Nothing Then Exit Sub
    Dim vKey As String
    On Error Resume Next
    vKey = vTextBox.Index
    vKey = vTextBox.Name & vKey
    vIniHnd.SaveSetting "AppSettingTextBox", vKey, vTextBox.Text
End Sub

Public Sub AppSetting_LoadTextBox(ByRef vIniHnd As CLiNInI, ByRef vTextBox As TextBox)
    If vTextBox Is Nothing Then Exit Sub
    If vIniHnd Is Nothing Then Exit Sub
    Dim vKey As String
    On Error Resume Next
    vKey = vTextBox.Index
    vKey = vTextBox.Name & vKey
    vTextBox.Text = vIniHnd.GetSetting("AppSettingTextBox", vKey)
End Sub
