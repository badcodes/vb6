VERSION 5.00
Object = "{35DFF35F-DEB3-11D0-8C50-00C04FC29CEC}#1.1#0"; "ColorPicker.ocx"
Begin VB.Form FTestColorPick 
   AutoRedraw      =   -1  'True
   Caption         =   "Test Color Pickers"
   ClientHeight    =   3708
   ClientLeft      =   1188
   ClientTop       =   2928
   ClientWidth     =   5352
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   7.8
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "TColorPick.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3708
   ScaleWidth      =   5352
   WhatsThisHelp   =   -1  'True
   Begin ColorPicker.XColorPicker pick 
      Height          =   1284
      Left            =   216
      TabIndex        =   2
      Top             =   120
      Width           =   1692
      _ExtentX        =   2985
      _ExtentY        =   2265
   End
   Begin VB.CheckBox chkWideControl 
      Caption         =   "Wide Control"
      Height          =   372
      Left            =   216
      TabIndex        =   1
      Top             =   2250
      UseMaskColor    =   -1  'True
      Width           =   1572
   End
   Begin VB.CheckBox chkWideForm 
      Caption         =   "Wide Form"
      Height          =   372
      Left            =   216
      TabIndex        =   0
      Top             =   1905
      UseMaskColor    =   -1  'True
      Width           =   1695
   End
   Begin VB.Label lbl 
      Caption         =   "Right-click form to display FColorPicker"
      Height          =   255
      Left            =   216
      TabIndex        =   3
      Top             =   2880
      Width           =   3495
   End
End
Attribute VB_Name = "FTestColorPick"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkWideControl_Click()
    pick.Wide = -chkWideControl
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, _
                         X As Single, Y As Single)
    If Button = 2 Then
        Dim getclr As New CColorPicker
        Static clrLast As Long
        ' Load last color used
        If clrLast <> 0& Then getclr.Color = clrLast
        ' Load dialog at given position and shape
        getclr.Load Left + X, Top + Y, -chkWideForm
        ' Save chosen color for next time
        clrLast = getclr.Color
        ' Change color of form and check boxes
        AllColors clrLast
    End If
End Sub

Private Sub pick_Picked(Color As stdole.OLE_COLOR)
    AllColors Color
End Sub

Sub AllColors(ByVal clr As Long)
    BackColor = clr
    chkWideForm.BackColor = clr
    chkWideControl.BackColor = clr
    chkWideForm.MaskColor = clr
    chkWideControl.MaskColor = clr
End Sub
