VERSION 5.00
Begin VB.Form FDemo 
   Caption         =   "Hash Table Demo"
   ClientHeight    =   3660
   ClientLeft      =   3195
   ClientTop       =   3615
   ClientWidth     =   4185
   LinkTopic       =   "Form1"
   ScaleHeight     =   3660
   ScaleWidth      =   4185
   Begin VB.TextBox txtSize 
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Top             =   1560
      Width           =   1335
   End
   Begin VB.ComboBox cmbMethod 
      Height          =   315
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   600
      Width           =   1815
   End
   Begin VB.TextBox txtCount 
      Height          =   285
      Left            =   240
      TabIndex        =   2
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "Test"
      Height          =   495
      Left            =   2640
      TabIndex        =   3
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label lblSize 
      Caption         =   "Table Size:"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label lblMethod 
      Caption         =   "Method:"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label lblCount 
      Caption         =   "Count:"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   2280
      Width           =   1215
   End
End
Attribute VB_Name = "FDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public ItemCount As Long
Public Method As Long
Public Size As Long

Private Sub cmbMethod_Click()
    Method = cmbMethod.ListIndex
    If Method = 2 Then
        txtSize.Enabled = False
        lblSize.Enabled = False
    Else
        txtSize.Enabled = True
        lblSize.Enabled = True
    End If
End Sub

Private Sub cmdTest_Click()
    Test.Test
End Sub

Private Sub Form_Load()
    ItemCount = 1000
    txtCount.Text = ItemCount
    cmbMethod.AddItem "Objects"
    cmbMethod.AddItem "Arrays"
    cmbMethod.AddItem "Collections"
    Method = 0
    cmbMethod.ListIndex = Method
    Size = 100
    txtSize.Text = Size
End Sub

Private Sub txtCount_Validate(Cancel As Boolean)
    If IsNumeric(txtCount.Text) Then
        ItemCount = txtCount.Text
    Else
        MsgBox "Count must be numeric"
        Cancel = True
    End If
End Sub

Private Sub txtSize_Validate(Cancel As Boolean)
    If IsNumeric(txtSize.Text) Then
        Size = txtSize.Text
    Else
        MsgBox "Size must be numeric"
        Cancel = True
    End If
End Sub
