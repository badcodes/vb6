VERSION 5.00
Begin VB.Form FTestWindow 
   Caption         =   "Test Window Class"
   ClientHeight    =   4908
   ClientLeft      =   1176
   ClientTop       =   1572
   ClientWidth     =   5784
   Icon            =   "TWindow.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4908
   ScaleWidth      =   5784
   Begin VB.TextBox txtOut 
      Height          =   4668
      Left            =   1500
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   120
      Width           =   4188
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   492
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   1248
   End
   Begin VB.CommandButton cmdHandle 
      Caption         =   "Window Data"
      Height          =   492
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1248
   End
End
Attribute VB_Name = "FTestWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdHandle_Click()
    Dim s As String
    Dim wndCur As New CWindow
    wndCur = Me.hWnd
    s = s & "Title: " & wndCur.Caption & sCrLf
    s = s & "Class Name: " & wndCur.ClassName & sCrLf
    s = s & "Instance: " & Hex$(wndCur.Instance) & sCrLf
    s = s & "Module: " & Hex$(wndCur.Module) & sCrLf
    s = s & "Process: " & Hex$(wndCur.Process) & sCrLf
    s = s & "Exe Name: " & wndCur.ExeName & sCrLf
    s = s & "Exe Path: " & wndCur.ExePath & sCrLf
    s = s & "Style: " & Hex$(wndCur.Style) & sCrLf
    s = s & "Extended Style: " & Hex$(wndCur.ExStyle) & sCrLf
    s = s & "Class Style: " & Hex$(wndCur.ClassStyle) & sCrLf
    s = s & "Parent: " & Hex$(wndCur.Parent) & sCrLf
    s = s & "Owner: " & Hex$(wndCur.Owner) & sCrLf
    s = s & "Owner Name: " & wndCur.OwnerName & sCrLf
    s = s & "Child: " & Hex$(wndCur.Child) & sCrLf
    s = s & "Handle: " & Hex$(wndCur.Handle) & sCrLf

    txtOut.Text = s
End Sub

