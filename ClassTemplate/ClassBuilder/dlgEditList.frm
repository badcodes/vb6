VERSION 5.00
Begin VB.Form dlgEditList 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit List"
   ClientHeight    =   2310
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2310
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cboField 
      Height          =   315
      Index           =   2
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1800
      Width           =   4395
   End
   Begin VB.ComboBox cboField 
      Height          =   315
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   1140
      Width           =   4395
   End
   Begin VB.ComboBox cboField 
      Height          =   315
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   420
      Width           =   4395
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   4
      Top             =   1740
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4680
      TabIndex        =   3
      Top             =   1260
      Width           =   1215
   End
   Begin VB.Label lblAny 
      Caption         =   "Style"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   7
      Top             =   1560
      Width           =   4335
   End
   Begin VB.Label lblAny 
      Caption         =   "To:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   4335
   End
   Begin VB.Label lblAny 
      Caption         =   "From:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "dlgEditList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private mOK As Integer

Private Sub CancelButton_Click()
    mOK = -1
    Me.Hide
End Sub

Private Sub Form_Load()
    mOK = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mOK = 0 Then
        Dim answer As VbMsgBoxResult
        answer = MsgBox("Not saved yet,Quit?", vbYesNo, "Confirm...")
        If (answer = vbOK) Then Cancel = 1
    End If
End Sub

Private Sub OKButton_Click()
    mOK = 1
    Me.Hide
End Sub

Public Property Get IsOK() As Long
    IsOK = mOK
End Property

Public Property Let IsOK(ByVal idx As Long)
    mOK = idx
End Property
Public Property Get TextField(ByVal idx As Long) As ComboBox
    If idx < 0 Or idx > cboField.count Then Exit Property
    Set TextField = cboField(idx)
End Property

Public Property Get TextFieldCount() As Long
    TextFieldCount = cboField.count
End Property
