VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TextBoxReportor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IFilterReporter
Private mTextBox As TextBox

Private Sub IFilterReporter_NotifyChange(sFrom As String, sTo As String, Optional Index As Long = -1&)
    If mTextBox Is Nothing Then Exit Sub
    mTextBox.Text = mTextBox.Text & "#" & CStr(Index) & ":" & "[" & sFrom & "] TO [" & sTo & "]" & vbCrLf
End Sub

Private Sub IFilterReporter_RequestChange(nLineIndex As Long, sText As String)

End Sub

Public Property Set TextBox(ByRef vTextBox As TextBox)
    Set mTextBox = vTextBox
    mTextBox.Text = ""
End Property
