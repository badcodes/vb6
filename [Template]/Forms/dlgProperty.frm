VERSION 5.00
Begin VB.Form dlgProperty 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Ù–‘"
   ClientHeight    =   5460
   ClientLeft      =   3585
   ClientTop       =   2085
   ClientWidth     =   4455
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frmProperty 
      Caption         =   "≥£πÊ"
      ClipControls    =   0   'False
      Height          =   4710
      Left            =   105
      TabIndex        =   1
      Top             =   165
      Width           =   4230
      Begin VB.Label lblProperty 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   4185
         Left            =   180
         TabIndex        =   2
         Top             =   285
         Width           =   3900
      End
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   330
      Left            =   3495
      TabIndex        =   0
      Top             =   5010
      Width           =   825
   End
End
Attribute VB_Name = "dlgProperty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub OKButton_Click()
Unload Me
End Sub
