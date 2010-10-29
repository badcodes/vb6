VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ExecLine"
   ClientHeight    =   5055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11730
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   11730
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraTaskContainer 
      Height          =   3855
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   11415
      Begin VB.VScrollBar vsTasks 
         Enabled         =   0   'False
         Height          =   3735
         Left            =   11040
         TabIndex        =   25
         Top             =   120
         Width           =   375
      End
      Begin VB.Frame frmTask 
         BorderStyle     =   0  'None
         Caption         =   "Tasks"
         Height          =   3555
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   10815
         Begin VB.CommandButton cmdDelete 
            Caption         =   "&Delete"
            Enabled         =   0   'False
            Height          =   375
            Index           =   1
            Left            =   9840
            TabIndex        =   18
            Top             =   120
            Width           =   855
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "&Edit"
            Enabled         =   0   'False
            Height          =   375
            Index           =   1
            Left            =   8760
            TabIndex        =   17
            Top             =   120
            Width           =   855
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "&Delete"
            Enabled         =   0   'False
            Height          =   375
            Index           =   2
            Left            =   9840
            TabIndex        =   16
            Top             =   720
            Width           =   855
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "&Edit"
            Enabled         =   0   'False
            Height          =   375
            Index           =   2
            Left            =   8760
            TabIndex        =   15
            Top             =   720
            Width           =   855
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "&Delete"
            Enabled         =   0   'False
            Height          =   375
            Index           =   3
            Left            =   9840
            TabIndex        =   14
            Top             =   1320
            Width           =   855
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "&Edit"
            Enabled         =   0   'False
            Height          =   375
            Index           =   3
            Left            =   8760
            TabIndex        =   13
            Top             =   1320
            Width           =   855
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "&Delete"
            Enabled         =   0   'False
            Height          =   375
            Index           =   4
            Left            =   9840
            TabIndex        =   12
            Top             =   1920
            Width           =   855
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "&Edit"
            Enabled         =   0   'False
            Height          =   375
            Index           =   4
            Left            =   8760
            TabIndex        =   11
            Top             =   1920
            Width           =   855
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "&Delete"
            Enabled         =   0   'False
            Height          =   375
            Index           =   5
            Left            =   9840
            TabIndex        =   10
            Top             =   2520
            Width           =   855
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "&Edit"
            Enabled         =   0   'False
            Height          =   375
            Index           =   5
            Left            =   8760
            TabIndex        =   9
            Top             =   2520
            Width           =   855
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "&Delete"
            Enabled         =   0   'False
            Height          =   375
            Index           =   6
            Left            =   9840
            TabIndex        =   8
            Top             =   3120
            Width           =   855
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "&Edit"
            Enabled         =   0   'False
            Height          =   375
            Index           =   6
            Left            =   8760
            TabIndex        =   7
            Top             =   3120
            Width           =   855
         End
         Begin VB.Label lblTask 
            Caption         =   "Task"
            ForeColor       =   &H8000000B&
            Height          =   360
            Index           =   1
            Left            =   60
            TabIndex        =   24
            Top             =   120
            Width           =   8520
         End
         Begin VB.Label lblTask 
            Caption         =   "Task"
            ForeColor       =   &H8000000B&
            Height          =   360
            Index           =   2
            Left            =   60
            TabIndex        =   23
            Top             =   720
            Width           =   8520
         End
         Begin VB.Label lblTask 
            Caption         =   "Task"
            ForeColor       =   &H8000000B&
            Height          =   360
            Index           =   3
            Left            =   60
            TabIndex        =   22
            Top             =   1320
            Width           =   8520
         End
         Begin VB.Label lblTask 
            Caption         =   "Task"
            ForeColor       =   &H8000000B&
            Height          =   360
            Index           =   4
            Left            =   60
            TabIndex        =   21
            Top             =   1920
            Width           =   8520
         End
         Begin VB.Label lblTask 
            Caption         =   "Task"
            ForeColor       =   &H8000000B&
            Height          =   360
            Index           =   5
            Left            =   60
            TabIndex        =   20
            Top             =   2520
            Width           =   8520
         End
         Begin VB.Label lblTask 
            Caption         =   "Task"
            ForeColor       =   &H8000000B&
            Height          =   360
            Index           =   6
            Left            =   60
            TabIndex        =   19
            Top             =   3120
            Width           =   8520
         End
      End
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit"
      Enabled         =   0   'False
      Height          =   375
      Index           =   0
      Left            =   10560
      TabIndex        =   4
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton cmdProcess 
      Caption         =   "&Process"
      Enabled         =   0   'False
      Height          =   375
      Left            =   9720
      TabIndex        =   2
      Top             =   4560
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Clear"
      Enabled         =   0   'False
      Height          =   375
      Left            =   8640
      TabIndex        =   1
      Top             =   4560
      Width           =   855
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Enabled         =   0   'False
      Height          =   375
      Left            =   7440
      TabIndex        =   0
      Top             =   4560
      Width           =   855
   End
   Begin VB.Label lblTask 
      Caption         =   "Task"
      ForeColor       =   &H8000000B&
      Height          =   360
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   10200
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
