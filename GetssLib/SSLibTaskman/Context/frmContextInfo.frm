VERSION 5.00
Begin VB.Form frmContextInfo 
   Caption         =   "ContextInfo"
   ClientHeight    =   5715
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9345
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   9345
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtCookie 
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   1500
      Width           =   9300
   End
   Begin VB.TextBox txtInfo 
      Height          =   375
      Left            =   -45
      TabIndex        =   3
      Top             =   1005
      Width           =   9300
   End
   Begin VB.TextBox txtText 
      Height          =   3570
      Left            =   -15
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Top             =   2025
      Width           =   9300
   End
   Begin VB.TextBox txtURL 
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   510
      Width           =   9300
   End
   Begin VB.TextBox txtRefer 
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   30
      Width           =   9300
   End
End
Attribute VB_Name = "frmContextInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

