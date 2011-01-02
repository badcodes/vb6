VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GLanguage"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub saveFormStr(frm As Form, Optional ByRef iniLanguage As String)

 On Error Resume Next


    Dim ctl As Control
    Dim obj As Object
    Dim sCtlType As String


    If iniLanguage = "" Then iniLanguage = bddir(App.Path) & "language.ini"
    
    saveLangStr frm.Caption, iniLanguage
    
     For Each ctl In frm.Controls
        sCtlType = TypeName(ctl)
        If sCtlType = "TabStrip" Then
            For Each obj In ctl.Tabs
                saveLangStr obj.Caption, iniLanguage
                saveLangStr obj.ToolTipText, iniLanguage
            Next
        ElseIf sCtlType = "Toolbar" Then
            For Each obj In ctl.Buttons
                saveLangStr obj.ToolTipText, iniLanguage
            Next
        Else
            saveLangStr ctl.Caption, iniLanguage
            saveLangStr ctl.ToolTipText, iniLanguage
        End If
    Next

End Sub
Public Sub loadFormStr(frm As Form, Optional ByRef iniLanguage As String)

 On Error Resume Next


    Dim ctl As Control
    Dim obj As Object
    Dim sCtlType As String

    If iniLanguage = "" Then iniLanguage = bddir(App.Path) & "language.ini"
    
    frm.Caption = loadLangStr(frm.Caption, iniLanguage)
    
     For Each ctl In frm.Controls
        sCtlType = TypeName(ctl)
        If sCtlType = "TabStrip" Then
            For Each obj In ctl.Tabs
                obj.Caption = loadLangStr(obj.Caption, iniLanguage)
                obj.ToolTipText = loadLangStr(obj.ToolTipText, iniLanguage)
            Next
        ElseIf sCtlType = "Toolbar" Then
            For Each obj In ctl.Buttons
                obj.ToolTipText = loadLangStr(obj.Caption, iniLanguage)
            Next
        Else
            ctl.Caption = loadLangStr(ctl.Caption, iniLanguage)
            ctl.ToolTipText = loadLangStr(ctl.ToolTipText, iniLanguage)
        End If
    Next

End Sub

Public Function loadLangStr(sEnglish As String, ByRef iniLanguage As String) As String

Dim sTransfered As String
If sEnglish = "" Then Exit Function
sTransfered = iniGetSetting(iniLanguage, "Language", sEnglish)

If sTransfered <> "" Then
    loadLangStr = sTransfered
Else
    loadLangStr = sEnglish
End If

End Function

Public Sub saveLangStr(sEnglish As String, ByRef iniLanguage As String)

If sEnglish = "" Then Exit Sub

iniSaveSetting iniLanguage, "Language", sEnglish, sEnglish

End Sub

