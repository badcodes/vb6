VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GSettings"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum EErrorSettings
    eeBaseSettings = 13600  ' Settings
End Enum

Const sSep = ","
Const sWinSection = "Window List"
Const sCtlSection = "Control List"

Sub RestoreWindow(ByVal hWnd As Long, sTitle As String)
    Dim sValue As String

    sValue = GetSetting(App.ExeName, sWinSection, sTitle, sEmpty)
    ' Quit if no entry
    If sValue = sEmpty Then Exit Sub

    Dim wp As WINDOWPLACEMENT
    wp.showCmd = Val(MParse.GetToken(sValue, sSep))
    If IsWindowVisible(hWnd) = False Then
        wp.showCmd = SW_HIDE
    End If
    ' wp.Flags = 0 (no flags required)
    wp.ptMinPosition.x = Val(MParse.GetToken(sEmpty, sSep))
    wp.ptMinPosition.y = Val(MParse.GetToken(sEmpty, sSep))
    wp.ptMaxPosition.x = Val(MParse.GetToken(sEmpty, sSep))
    wp.ptMaxPosition.y = Val(MParse.GetToken(sEmpty, sSep))
    wp.rcNormalPosition.Left = Val(MParse.GetToken(sEmpty, sSep))
    wp.rcNormalPosition.Top = Val(MParse.GetToken(sEmpty, sSep))
    wp.rcNormalPosition.Right = Val(MParse.GetToken(sEmpty, sSep))
    wp.rcNormalPosition.bottom = Val(MParse.GetToken(sEmpty, sSep))
    ' Remember to set length
    wp.length = Len(wp)
    ' Send all your settings to the system
    Call SetWindowPlacement(hWnd, wp)

End Sub

Sub SaveWindow(ByVal hWnd As Long, sTitle As String)
    
    Dim sValue As String, wp As WINDOWPLACEMENT
    ' First set type length for Windows
    wp.length = Len(wp)
    ' Get coordinates and other data about the window
    Call GetWindowPlacement(hWnd, wp)
    ' Read and use the data
    sValue = wp.showCmd & ","
    sValue = sValue & wp.ptMinPosition.x & ","
    sValue = sValue & wp.ptMinPosition.y & ","
    sValue = sValue & wp.ptMaxPosition.x & ","
    sValue = sValue & wp.ptMaxPosition.y & ","
    sValue = sValue & wp.rcNormalPosition.Left & ","
    sValue = sValue & wp.rcNormalPosition.Top & ","
    sValue = sValue & wp.rcNormalPosition.Right & ","
    sValue = sValue & wp.rcNormalPosition.bottom
    
    SaveSetting App.ExeName, sWinSection, sTitle, sValue

End Sub

Sub ClearWindowSetting(sTitle As String)
    DeleteSetting App.ExeName, sWinSection, sTitle
End Sub

' These could have been implemented with Form properties, but
' API does a better job of handling minimized and maximized windows.

' Call only in Form_Load before Show
#If fComponent Then
Sub RestoreForm(frm As Object, Optional sTitle As String)
#Else
Sub RestoreForm(frm As Form, Optional sTitle As String)
#End If
    If sTitle = sEmpty Then sTitle = frm.Caption
    RestoreWindow frm.hWnd, sTitle
End Sub

' Call only in Form_Unload when form is closing
#If fComponent Then
Sub SaveForm(frm As Object, Optional sTitle As String)
#Else
Sub SaveForm(frm As Form, Optional sTitle As String)
#End If
    If sTitle = sEmpty Then sTitle = frm.Caption
    SaveWindow frm.hWnd, sTitle
End Sub

#If fComponent Then
Sub ClearFormSetting(frm As Object, Optional sTitle As String)
#Else
Sub ClearFormSetting(frm As Form, Optional sTitle As String)
#End If
    If sTitle = sEmpty Then sTitle = frm.Caption
    DeleteSetting App.ExeName, sWinSection, sTitle
End Sub

#If fComponent Then
Sub RestoreCtl(ctl As Object, sTitle As String)
#Else
Sub RestoreCtl(ctl As Control, sTitle As String)
#End If
With ctl
    Dim sValue As String, s As String, i As Long
    sValue = GetSetting(App.ExeName, sCtlSection, sTitle, sEmpty)
    ' Quit if no entry
    If sValue = sEmpty Then Exit Sub

    ' Set left and top while in normal mode
    On Error Resume Next
    .Left = MParse.GetToken(sValue, sSep)
    .Top = MParse.GetToken(sEmpty, sSep)
    .Width = MParse.GetToken(sEmpty, sSep)
    .Height = MParse.GetToken(sEmpty, sSep)
    .Enabled = MParse.GetToken(sEmpty, sSep)
    .Visible = MParse.GetToken(sEmpty, sSep)
    s = .Caption
    If Err = 0 Then .Caption = MParse.GetToken(sEmpty, sSep)
    i = .BackColor
    If Err = 0 Then .BackColor = MParse.GetToken(sEmpty, sSep)
    i = .ForeColor
    If Err = 0 Then .ForeColor = MParse.GetToken(sEmpty, sSep)
    ' Enhance to restore anything else you saved

End With
End Sub

#If fComponent Then
Sub SaveCtl(ctl As Object, sTitle As String)
#Else
Sub SaveCtl(ctl As Control, sTitle As String)
#End If
With ctl
    Dim sValue As String, s As String, i As Long

    On Error Resume Next
    sValue = sValue & .Left & "," & .Top & ","
    sValue = sValue & .Width & "," & .Height & ","
    sValue = sValue & .Enabled & ","
    sValue = sValue & .Visible & ","
    s = .Caption
    If Err = 0 Then sValue = sValue & .Caption & ","
    i = .BackColor
    If Err = 0 Then sValue = sValue & .BackColor & ","
    i = .ForeColor
    If Err = 0 Then sValue = sValue & .ForeColor & ","
    
    ' Enhance to save anything else you need

    SaveSetting App.ExeName, sCtlSection, sTitle, sValue

End With
End Sub

' Add more save and restore functions:
'     SaveFont/RestoreFont
'     SaveTextBox/RestoreTextBox

#If fComponent = 0 Then
Private Sub ErrRaise(e As Long)
    Dim sText As String, sSource As String
    If e > 1000 Then
        sSource = App.ExeName & ".Settings"
        Select Case e
        Case eeBaseSettings
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

