VERSION 5.00
Begin VB.Form frmExpFileMenu 
   Caption         =   "Form1"
   ClientHeight    =   8280
   ClientLeft      =   -9996
   ClientTop       =   1980
   ClientWidth     =   6684
   LinkTopic       =   "Form1"
   ScaleHeight     =   8280
   ScaleWidth      =   6684
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu mnuFileFind 
         Caption         =   "&Find"
      End
      Begin VB.Menu mnuFileBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSend 
         Caption         =   "Sen&d to"
      End
      Begin VB.Menu mnuFileBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
      End
      Begin VB.Menu mnuFileBar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileDelete 
         Caption         =   "&Delete"
      End
      Begin VB.Menu mnuFileRename 
         Caption         =   "Rena&me"
      End
      Begin VB.Menu mnuFileProperties 
         Caption         =   "Propert&ies"
      End
      Begin VB.Menu mnuFileBar4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "&Close"
      End
   End
End
Attribute VB_Name = "frmExpFileMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub mnuFileClose_Click()
  'unload the form
  Unload Me
End Sub

Private Sub mnuFileDelete_Click()
  MsgBox "Delete Code goes here!"
End Sub

Private Sub mnuFileNew_Click()
  MsgBox "New File Code goes here!"
End Sub

Private Sub mnuFileOpen_Click()
  MsgBox "Open Code goes here!"
End Sub

Private Sub mnuFileProperties_Click()
  MsgBox "Properties Code goes here!"
End Sub

Private Sub mnuFileRename_Click()
  MsgBox "Rename Code goes here!"
End Sub

Private Sub mnuFileSend_Click()
  MsgBox "Send Code goes here!"
End Sub
