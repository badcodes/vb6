VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   6225
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   12840
   LinkTopic       =   "Form1"
   ScaleHeight     =   6225
   ScaleWidth      =   12840
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView typeList 
      Height          =   3855
      Left            =   300
      TabIndex        =   7
      Top             =   1920
      Width           =   12255
      _ExtentX        =   21616
      _ExtentY        =   6800
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FlatScrollBar   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "from"
         Text            =   "From"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "to"
         Text            =   "To"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "style"
         Text            =   "Style"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame frmOption 
      Caption         =   "Type Table"
      Height          =   4215
      Left            =   180
      TabIndex        =   6
      Top             =   1680
      Width           =   12495
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "Command1"
      Height          =   375
      Index           =   1
      Left            =   11280
      TabIndex        =   4
      Top             =   1200
      Width           =   1395
   End
   Begin VB.TextBox txtFilename 
      Height          =   315
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   1260
      Width           =   10995
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "Command1"
      Height          =   375
      Index           =   0
      Left            =   11280
      TabIndex        =   1
      Top             =   420
      Width           =   1395
   End
   Begin VB.TextBox txtFilename 
      Height          =   315
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   480
      Width           =   10995
   End
   Begin VB.Label lblAnything 
      Caption         =   "Label1"
      Height          =   255
      Index           =   1
      Left            =   180
      TabIndex        =   5
      Top             =   900
      Width           =   12495
   End
   Begin VB.Label lblAnything 
      Caption         =   "Label1"
      Height          =   255
      Index           =   0
      Left            =   180
      TabIndex        =   2
      Top             =   120
      Width           =   12495
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_Load()
    Form_Resize
End Sub

Private Sub Form_Resize()
    With typeList
        Dim nCol As Integer
        Dim lCol As Double
        nCol = .ColumnHeaders.Count
        lCol = .Width / nCol
        Do Until nCol < 1
            .ColumnHeaders(nCol).Width = lCol
            nCol = nCol - 1
        Loop
    End With
End Sub
