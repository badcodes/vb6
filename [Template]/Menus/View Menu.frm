VERSION 5.00
Begin VB.Form frmViewMenu 
   Caption         =   "Form1"
   ClientHeight    =   8280
   ClientLeft      =   -9996
   ClientTop       =   2052
   ClientWidth     =   6684
   LinkTopic       =   "Form1"
   ScaleHeight     =   8280
   ScaleWidth      =   6684
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewToolbar 
         Caption         =   "&Toolbar"
      End
      Begin VB.Menu mnuViewStatusBar 
         Caption         =   "Status &Bar"
      End
      Begin VB.Menu mnuViewBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewLargeIcons 
         Caption         =   "Lar&ge Icons"
      End
      Begin VB.Menu mnuViewSmallIcons 
         Caption         =   "S&mall Icons"
      End
      Begin VB.Menu mnuViewList 
         Caption         =   "&List"
      End
      Begin VB.Menu mnuViewDetails 
         Caption         =   "&Details"
      End
      Begin VB.Menu mnuViewBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewArrangeIcons 
         Caption         =   "Arrange &Icons"
         Begin VB.Menu mnuVAIByName 
            Caption         =   "by &Name"
         End
         Begin VB.Menu mnuVAIByType 
            Caption         =   "by &Type"
         End
         Begin VB.Menu mnuVAIBySize 
            Caption         =   "by Si&ze"
         End
         Begin VB.Menu mnuVAIByDate 
            Caption         =   "by &Date"
         End
      End
      Begin VB.Menu mnuViewLineUpIcons 
         Caption         =   "Li&ne Up Icons"
      End
      Begin VB.Menu mnuViewBar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewRefresh 
         Caption         =   "&Refresh"
      End
      Begin VB.Menu mnuViewOptions 
         Caption         =   "&Options..."
      End
   End
End
Attribute VB_Name = "frmViewMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Const NAME_COLUMN = 0
Const TYPE_COLUMN = 1
Const SIZE_COLUMN = 2
Const DATE_COLUMN = 3

Private Sub mnuVAIByDate_Click()
'  lvListView.SortKey = DATE_COLUMN
End Sub

Private Sub mnuVAIByName_Click()
'  lvListView.SortKey = NAME_COLUMN
End Sub

Private Sub mnuVAIBySize_Click()
'  lvListView.SortKey = SIZE_COLUMN
End Sub

Private Sub mnuVAIByType_Click()
'  lvListView.SortKey = TYPE_COLUMN
End Sub

Private Sub mnuViewDetails_Click()
'  lvListView.View = lvwReport
End Sub

Private Sub mnuViewLargeIcons_Click()
'  lvListView.View = lvwIcon
End Sub

Private Sub mnuViewLineUpIcons_Click()
'  lvListView.Arrange = lvwAutoLeft
End Sub

Private Sub mnuViewList_Click()
'  lvListView.View = lvwList
End Sub

Private Sub mnuViewOptions_Click()
'  frmOptions.Show vbModal
End Sub

Private Sub mnuViewRefresh_Click()
  MsgBox "Place Refresh Code here!"
End Sub

Private Sub mnuViewSmallIcons_Click()
'  lvListView.View = lvwSmallIcon
End Sub

Private Sub mnuViewStatusBar_Click()
  If mnuViewStatusBar.Checked Then
'    sbStatusBar.Visible = False
    mnuViewStatusBar.Checked = False
  Else
'    sbStatusBar.Visible = True
    mnuViewStatusBar.Checked = True
  End If
End Sub

Private Sub mnuViewToolbar_Click()
  If mnuViewToolbar.Checked Then
'    tbToolBar.Visible = False
    mnuViewToolbar.Checked = False
  Else
'    tbToolBar.Visible = True
    mnuViewToolbar.Checked = True
  End If
End Sub
