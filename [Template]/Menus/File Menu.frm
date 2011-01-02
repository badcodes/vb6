VERSION 5.00
Begin VB.Form frmFileMenu 
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
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "&Close"
      End
      Begin VB.Menu mnuFileBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save &As..."
      End
      Begin VB.Menu mnuFileSaveAll 
         Caption         =   "Save A&ll"
      End
      Begin VB.Menu mnuFileBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileProperties 
         Caption         =   "Propert&ies"
      End
      Begin VB.Menu mnuFileBar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePrintSetup 
         Caption         =   "Print Set&up..."
      End
      Begin VB.Menu mnuFilePrintPreview 
         Caption         =   "Print Pre&view"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "&Print..."
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileBar4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSend 
         Caption         =   "Sen&d..."
      End
      Begin VB.Menu mnuFileBar5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmFileMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub mnuFileClose_Click()
  MsgBox "Close Code goes here!"
End Sub

Private Sub mnuFileExit_Click()
  'unload the form
  Unload Me
End Sub

Private Sub mnuFileNew_Click()
  MsgBox "New File Code goes here!"
End Sub

Private Sub mnuFileOpen_Click()
  MsgBox "Open Code goes here!"
End Sub

Private Sub mnuFilePrint_Click()
  MsgBox "Print Code goes here!"
End Sub

Private Sub mnuFilePrintPreview_Click()
  MsgBox "Print Preview Code goes here!"
End Sub

Private Sub mnuFilePrintSetup_Click()
  MsgBox "Print Setup Code goes here!"
End Sub

Private Sub mnuFileProperties_Click()
  MsgBox "Properties Code goes here!"
End Sub

Private Sub mnuFileSave_Click()
  MsgBox "Save File Code goes here!"
End Sub

Private Sub mnuFileSaveAll_Click()
  MsgBox "Save All Code goes here!"
End Sub

Private Sub mnuFileSaveAs_Click()
  MsgBox "Save As Code goes here!"
End Sub

Private Sub mnuFileSend_Click()
  MsgBox "Send Code goes here!"
End Sub
