VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Packing Results"
   ClientHeight    =   5460
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6795
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   6795
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton Command1 
      Caption         =   "Return"
      Height          =   495
      Left            =   5400
      TabIndex        =   1
      Top             =   4920
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000004&
      Height          =   4575
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   240
      Width           =   6255
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Text1.Text = ""
Form2.Hide
Form1.Show
End Sub

