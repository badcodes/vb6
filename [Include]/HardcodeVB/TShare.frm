VERSION 5.00
Begin VB.Form FTestSharedMemory 
   Caption         =   "Test Shared Memory"
   ClientHeight    =   1530
   ClientLeft      =   1560
   ClientTop       =   2070
   ClientWidth     =   3840
   FillColor       =   &H00000080&
   FillStyle       =   7  'Diagonal Cross
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   -1  'True
   EndProperty
   Icon            =   "TShare.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1530
   ScaleWidth      =   3840
   Begin VB.CommandButton cmdGet 
      Caption         =   "&Get String"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2364
      TabIndex        =   3
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton cmdSet 
      Caption         =   "&Set String"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   210
      TabIndex        =   2
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox txtShare 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   210
      TabIndex        =   0
      Top             =   480
      Width           =   3375
   End
   Begin VB.Label lbl 
      Caption         =   "Shared String:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   228
      TabIndex        =   1
      Top             =   240
      Width           =   1692
   End
End
Attribute VB_Name = "FTestSharedMemory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private ss As New CSharedString

Private Sub Form_Load()
    ss.Create "MyShare"
    If ss = sEmpty Then
        ss = "Hello from the Creator"
    End If
    txtShare = ss
End Sub

Private Sub cmdSet_Click()
    ss = txtShare
End Sub

Private Sub cmdGet_Click()
    txtShare = ss
End Sub


