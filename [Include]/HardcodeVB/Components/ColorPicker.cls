VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CColorPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum EErrorColorPicker
    eeBaseColorPicker = 13670   ' CColorPicker
End Enum

Private frm As New FColorPicker

Sub Load(Optional ByVal Left As Single = -1#, _
         Optional ByVal Top As Single = -1#, _
         Optional ByVal Wide As Boolean = True)
    frm.Wide = Wide
    If Left <> -1# Then frm.Left = Left
    If Top <> -1# Then frm.Top = Top
    frm.Show vbModal
End Sub

Property Get Color() As Long
    Color = frm.Color
End Property

Property Let Color(clr As Long)
    frm.Color = clr
End Property

Property Get Wide() As Boolean
    Wide = frm.Wide
End Property

Property Let Wide(fWide As Boolean)
    frm.Wide = fWide
End Property


