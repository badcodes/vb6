VERSION 5.00
Begin VB.Form frmEditMenu 
   Caption         =   "Form1"
   ClientHeight    =   2868
   ClientLeft      =   -9996
   ClientTop       =   2880
   ClientWidth     =   4332
   LinkTopic       =   "Form1"
   ScaleHeight     =   2868
   ScaleWidth      =   4332
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditUndo 
         Caption         =   "&Undo"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuEditBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditCut 
         Caption         =   "Cu&t"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuEditPasteSpecial 
         Caption         =   "Paste &Special..."
      End
      Begin VB.Menu mnuEditBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditDSelectAll 
         Caption         =   "Select &All"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEditInvertSelection 
         Caption         =   "&Invert Selection"
      End
   End
End
Attribute VB_Name = "frmEditMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub mnuEditCopy_Click()
  MsgBox "Place Copy Code here!"
End Sub

Private Sub mnuEditCut_Click()
  MsgBox "Place Cut Code here!"
End Sub

Private Sub mnuEditDSelectAll_Click()
  MsgBox "Place Select All Code here!"
End Sub

Private Sub mnuEditInvertSelection_Click()
  MsgBox "Place Invert Selection Code here!"
End Sub

Private Sub mnuEditPaste_Click()
  MsgBox "Place Paste Code here!"
End Sub

Private Sub mnuEditPasteSpecial_Click()
  MsgBox "Place Paste Special Code here!"
End Sub

Private Sub mnuEditUndo_Click()
  MsgBox "Place Undo Code here!"
End Sub
