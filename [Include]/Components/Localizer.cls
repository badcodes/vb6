VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLocalizer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Dim hIni As CLiNInI
Private m_ApplyTo As Object 'local copy
Private fInstalled As Boolean

'Public Property Set ApplyTo(ByVal vData As Form)
'    Set m_ApplyTo = vData
'    Install
'End Property
'
'Public Property Let LanguageInifile(ByVal vData As String)
'   m_LanguageInifile = vData
'   Install
'End Property
'
'Public Property Get LanguageInifile() As String
'   LanguageInifile = m_LanguageInifile
'End Property
'Public Sub saveFormStr()
'
'    If fInstalled = False Then Exit Sub
'    On Error Resume Next
'    Dim ctl As Control
'    Dim obj As Object
'    Dim sCtlType As String
'    saveLangStr m_ApplyTo.Caption
'
'    For Each ctl In m_ApplyTo.Controls
'        sCtlType = TypeName(ctl)
'
'        If sCtlType = "TabStrip" Then
'
'            For Each obj In ctl.Tabs
'                saveLangStr obj.Caption
'                saveLangStr obj.ToolTipText
'            Next
'
'        ElseIf sCtlType = "Toolbar" Then
'
'            For Each obj In ctl.Buttons
'                saveLangStr obj.ToolTipText
'            Next
'
'        Else
'            saveLangStr ctl.Caption
'            saveLangStr ctl.ToolTipText
'        End If
'
'    Next
'
'    hIni.Save
'
'End Sub

Public Sub loadFormStr()

    If fInstalled = False Then Exit Sub
    On Error Resume Next
    Dim ctl As Control
    Dim obj As Object
    Dim sCtlType As String
    m_ApplyTo.Caption = loadLangStr(m_ApplyTo.Caption)

    For Each ctl In m_ApplyTo.Controls
        sCtlType = TypeName(ctl)

        If sCtlType = "TabStrip" Then

            For Each obj In ctl.Tabs
                obj.Caption = loadLangStr(obj.Caption)
                obj.ToolTipText = loadLangStr(obj.ToolTipText)
            Next

        ElseIf sCtlType = "Toolbar" Then

            For Each obj In ctl.Buttons
                obj.ToolTipText = loadLangStr(obj.Caption)
            Next

        Else
            ctl.Caption = loadLangStr(ctl.Caption)
            ctl.ToolTipText = loadLangStr(ctl.ToolTipText)
        End If

    Next

End Sub

Public Function loadLangStr(sEnglish As String) As String

    If fInstalled = False Then Exit Function
    Dim sTransfered As String

    If sEnglish = "" Then Exit Function
    sTransfered = hIni.GetSetting("Language", sEnglish)

    If sTransfered <> "" Then
        loadLangStr = sTransfered
    Else
        loadLangStr = sEnglish
    End If

End Function

Public Sub saveLangStr(sEnglish As String)

    If fInstalled = False Then Exit Sub

    If sEnglish = "" Then Exit Sub
    hIni.SaveSetting "Language", sEnglish, sEnglish

End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    Set hIni = Nothing

End Sub

Public Sub Install(applyTo As Form, LanguageIni As String, Optional cmp As VbCompareMethod = vbTextCompare)

    If IsObject(applyTo) = False Then Exit Sub

    If LanguageIni = "" Then Exit Sub
    Set hIni = New CLiNInI
    Set m_ApplyTo = applyTo
    hIni.Create LanguageIni
    hIni.CompareMethod = cmp
    fInstalled = True

End Sub


