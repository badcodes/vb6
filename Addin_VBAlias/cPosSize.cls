VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPosSize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const kVBEWindowControl = "VBEWindowControl"

Private pHeight As Long
Private pLeft As Long
Private pTop As Long
Private pWidth As Long
Private pID As String

Public Sub Init(iID As String)
pID = iID
pHeight = CLng(GetSetting(kVBEWindowControl, pID, _
        "Height", 0))
pWidth = CLng(GetSetting(kVBEWindowControl, pID, _
        "Width", 0))
pLeft = CLng(GetSetting(kVBEWindowControl, pID, _
        "Left", 0))
pTop = CLng(GetSetting(kVBEWindowControl, pID, _
        "Top", 0))
End Sub

Public Sub Move(iWindow As Window)
If CBool(pWidth) And CBool(pHeight) Then
    With iWindow
        .Left = pLeft
        .Top = pTop
        .Width = pWidth
        .Height = pHeight
    End With
End If
End Sub

Public Sub Store(iWindow As Window)
Dim nSection As String

With iWindow
    pLeft = .Left
    pTop = .Top
    pWidth = .Width
    pHeight = .Height
End With
SaveSetting kVBEWindowControl, pID, "Width", pWidth
SaveSetting kVBEWindowControl, pID, "Height", pHeight
SaveSetting kVBEWindowControl, pID, "Left", pLeft
SaveSetting kVBEWindowControl, pID, "Top", pTop
End Sub


