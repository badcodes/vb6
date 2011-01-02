VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form MainFrm 
   AutoRedraw      =   -1  'True
   Caption         =   "ssLibQuery"
   ClientHeight    =   6420
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   9252
   Icon            =   "MainFrm2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6420
   ScaleWidth      =   9252
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView lvResult 
      Height          =   3492
      Left            =   204
      TabIndex        =   0
      Top             =   1740
      Width           =   8772
      _ExtentX        =   15473
      _ExtentY        =   6160
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      TextBackground  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "no"
         Text            =   "NO."
         Object.Width           =   706
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "title"
         Text            =   "Title"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "author"
         Text            =   "Author"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Key             =   "catalog"
         Text            =   "Catalog"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Key             =   "library"
         Text            =   "Library"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Key             =   "page"
         Text            =   "Pages"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Key             =   "ss"
         Text            =   "SSID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Key             =   "date"
         Text            =   "Date"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame frmProcess 
      Height          =   1200
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   8775
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear"
         Height          =   315
         Left            =   6312
         TabIndex        =   13
         Top             =   240
         Width           =   1020
      End
      Begin VB.CommandButton cmdHelp 
         Caption         =   "Help"
         Height          =   315
         Left            =   7680
         TabIndex        =   12
         Top             =   240
         Width           =   855
      End
      Begin VB.ComboBox txtDIRSearch 
         Appearance      =   0  'Flat
         Height          =   288
         Left            =   2040
         TabIndex        =   11
         Top             =   240
         Width           =   3912
      End
      Begin VB.TextBox txtKeyword 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   972
         TabIndex        =   5
         Top             =   720
         Width           =   2124
      End
      Begin VB.CommandButton cmdCTSearch 
         Caption         =   "Search"
         Default         =   -1  'True
         Height          =   315
         Left            =   6360
         TabIndex        =   4
         Top             =   696
         Width           =   975
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "Stop"
         Enabled         =   0   'False
         Height          =   315
         Left            =   7692
         TabIndex        =   3
         Top             =   696
         Width           =   855
      End
      Begin VB.ComboBox cboLib 
         Appearance      =   0  'Flat
         Height          =   288
         Left            =   3420
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   720
         Width           =   2535
      End
      Begin VB.Label lblKeyword 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "KeyWord:"
         Height          =   192
         Left            =   132
         TabIndex        =   7
         Top             =   756
         Width           =   708
      End
      Begin VB.Label lblSSReader 
         AutoSize        =   -1  'True
         Caption         =   "SSReader Library Path:"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1665
      End
   End
   Begin MSComctlLib.ProgressBar stsbar 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   8
      Top             =   6105
      Width           =   9255
      _ExtentX        =   16320
      _ExtentY        =   550
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label lblAbout 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "CopyRight"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   192
      Left            =   8172
      TabIndex        =   10
      Top             =   48
      Width           =   792
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      Caption         =   "lblStatus"
      Height          =   192
      Left            =   204
      TabIndex        =   9
      Top             =   1500
      Visible         =   0   'False
      Width           =   600
   End
End
Attribute VB_Name = "MainFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const iniFile = "ssLibquery.ini"
Public hSet As CAutoSetting
Dim Force_Stop As Boolean
Dim infoDialog As Dialog

    Dim Lib() As String
    Dim libCount As Long


Private Sub cmdClear_Click()
txtDIRSearch.Clear
End Sub

Private Sub cmdctsearch_Click()


lvResult.ListItems.Clear
'Dim fso As New FileSystemObject
'Dim root As String
'root = fso.BuildPath(txtDIRSearch.Text, "liblist.dat")
'Set fso = Nothing

cmdCTSearch.Enabled = False
cmdStop.Enabled = True
Force_Stop = False

'Dim lastFor As String
'lastFor = Time$
queryPdgLib
'MsgBox "Start at " & lastFor & vbCrLf & "End at " + Time$, vbOKCancel

End Sub



Private Sub cmdHelp_Click()
Dim helpFile As String
helpFile = App.Path & "\" & "help.htm"
ShellExecute Me.hWnd, "open", helpFile, "", "", 1
End Sub

Private Sub cmdStop_Click()

cmdStop.Enabled = False
cmdCTSearch.Enabled = True

Force_Stop = True

End Sub

Private Sub Form_Load()

Set infoDialog = Dialog

'Load infoDialog

Set hSet = New CAutoSetting
hSet.fileNameSaveTo = App.Path & "\" & iniFile
Dim i As Long


With hSet
.Add MainFrm, SF_FORM
.Add txtKeyword, SF_TEXT
.Add txtDIRSearch, SF_LISTTEXT
.Add txtDIRSearch, SF_TEXT
.Add lblSSReader, SF_CAPTION, , True
.Add lblKeyword, SF_CAPTION, , True
.Add cmdClear, SF_CAPTION, , True
.Add cmdHelp, SF_CAPTION, , True
.Add cmdCTSearch, SF_CAPTION, , True
.Add cmdStop, SF_CAPTION, , True

For i = 1 To lvResult.ColumnHeaders.Count
    .Add MainFrm.lvResult.ColumnHeaders(i), SF_WIDTH, "MainFrm.lvResult.ColumnHeaders" & CStr(i) & ".Width"
    .Add MainFrm.lvResult.ColumnHeaders(i), SF_TEXT, "Mainfrm.lvResult.ColumnHeaders" & CStr(i) & ".Text", True
Next


.Add infoDialog, SF_FORM
For i = 1 To infoDialog.lblInfo.Count
    .Add infoDialog.lblInfo(i - 1), SF_CAPTION, , True
Next

.Add infoDialog.cmdOpenInIE, SF_CAPTION, , True
.Add infoDialog.cmdOpenSS, SF_CAPTION, , True
.Add infoDialog.OKButton, SF_CAPTION, , True


'For i = 0 To chkKeyIN.Count - 1
'    .Add MainFrm.chkKeyIN(i), SF_VALUE ', "MainFrm.chkKeyIN" & CStr(i)
'Next
End With

lblAbout.Caption = "  " & App.ProductName & " V" & App.Major & "." & App.Minor & "." & App.Revision & " by " & App.LegalCopyright & "@" & App.CompanyName

'Dim tmp
'tmp = lblSSReader.ForeColor
'lblSSReader.ForeColor = lblSSReader.BackColor
'lblSSReader.BackColor = tmp


txtDIRSearch_Change
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    Dim w As Long
    
    lblAbout.Left = MainFrm.ScaleWidth - lblAbout.Width - 240
    With frmProcess
        .Width = MainFrm.ScaleWidth - .Left * 2
        txtDIRSearch.Left = lblSSReader.Left + lblSSReader.Width + 120
        txtDIRSearch.Width = .Width - txtDIRSearch.Left - cmdHelp.Width - cmdClear.Width - 3 * 120
        
        
        cmdClear.Left = txtDIRSearch.Left + txtDIRSearch.Width + 120
        cmdHelp.Left = cmdClear.Left + cmdClear.Width + 120
        
        
        lblKeyword.Left = lblSSReader.Left
        txtKeyword.Left = lblKeyword.Left + lblKeyword.Width + 120
        
'        w = .Width - lblKeyword.Width - lblAuthor.Width - cboLib.Width - 6 * 120
'        txtKeyword.Width = w / 2 '.Width - txtKeyword.Left - cboLib.Width - cmdCTSearch.Width - cmdStop.Width - 4 * 120
'
'        lblAuthor.Left = txtKeyword.Left + txtKeyword.Width + 120
'
'        txtAuthor.Left = lblAuthor.Left + lblAuthor.Width + 120
'
'        txtAuthor.Width = w / 2
'

        txtKeyword.Width = .Width - lblKeyword.Width - cboLib.Width - cmdCTSearch.Width - cmdStop.Width - 6 * 120
       
        cboLib.Left = txtKeyword.Left + txtKeyword.Width + 120
                
        cmdCTSearch.Left = cboLib.Left + cboLib.Width + 120
        cmdStop.Left = cmdCTSearch.Left + cmdCTSearch.Width + 120
        
        
    End With
        
        
    With lvResult
    .Left = frmProcess.Left
    .Width = frmProcess.Width
    .Height = stsbar.Top - .Top
    lblStatus.Top = .Top - 100
    lblStatus.Left = .Left
    End With
    
'
'    With stsbar
'    .Left = frmProcess.Left
'    .Width = frmProcess.Width
'    .Top = lvResult.Top
'    End With
'    With lvResult.ColumnHeaders
'        Dim lall As Single
'        lall = lvResult.Width - 400
'        .Item(1).Width = 0.4 * lall
'        .Item(2).Width = 0.3 * lall
'        .Item(3).Width = 0.15 * lall
'        .Item(4).Width = 0.15 * lall
'
'    End With
End Sub

Private Sub Form_Terminate()
    Force_Stop = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Force_Stop = True
    Set hSet = Nothing
    Set infoDialog = Nothing
    'Unload infoDialog
    Unload Dialog
    Unload Me
End Sub



Private Sub lvResult_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lvResult
    .SortKey = ColumnHeader.Index - 1
    If .SortOrder = lvwAscending Then .SortOrder = lvwDescending Else .SortOrder = lvwAscending
    .Sorted = True
    End With
    
End Sub



Public Function queryPdgLib() As String


    Dim Cata() As String
    Dim cataCount As Long
    Dim Book() As String
    Dim bookCount As String
    Dim LI As ListItem
    Dim c As Long
    
    
    'If MainFrm.WindowState = 1 Then MainFrm.WindowState = 0
    
    MainFrm.Height = MainFrm.Height - lvResult.Height + 120
    lvResult.Visible = False
    lblStatus.Visible = True
    
    
    Dim i As Long
    For i = 1 To libCount
        If i = cboLib.ListIndex Or cboLib.ListIndex = 0 Then
            cataCount = pdgCatalist(Lib(2, i), Cata())
            Dim j As Long
            stsbar.Value = 0
            stsbar.Min = 0
            If cataCount > 0 Then stsbar.Max = cataCount


            For j = 1 To cataCount
                lblStatus.Caption = "looking in [" & Lib(1, i) & "] - " & Cata(1, j)
                bookCount = pdgBookList(Cata(2, j), Book())
                Dim K As Long
                For K = 1 To bookCount
                
                    If match(Book(), K) Then  'InStr(Book(1, k), strQuery) > 0 Or _
                                            'InStr(Book(2, k), strQuery) > 0 Then
                        c = c + 1
'                        Set LI = lvResult.ListItems.Add(, uniqueStr(CStr(c)), StrNum(c, 5))
'                        LI.ListSubItems.Add , uniqueStr(Book(1, K)), Book(1, K)
'                        LI.ListSubItems.Add , uniqueStr(Book(2, K)), Book(2, K)
'                        LI.ListSubItems.Add , uniqueStr(Cata(1, j)), Cata(1, j)
'                        LI.ListSubItems.Add , uniqueStr(Lib(1, i)), Lib(1, i)
'                        LI.ListSubItems.Add , uniqueStr(Book(4, K)), Book(4, K)
'                        LI.ListSubItems.Add , uniqueStr(Book(5, K)), Book(5, K)
'                        LI.ListSubItems.Add , uniqueStr(Book(3, K)), Book(3, K)
'                        LI.ListSubItems.Add , uniqueStr(Book(6, K)), Book(6, K)
                        Set LI = lvResult.ListItems.Add(, , StrNum(c, 5))
                        LI.ListSubItems.Add , , Book(1, K)
                        LI.ListSubItems.Add , , Book(2, K)
                        LI.ListSubItems.Add , , Cata(1, j)
                        LI.ListSubItems.Add , , Lib(1, i)
                        LI.ListSubItems.Add , , Book(4, K)
                        LI.ListSubItems.Add , , Book(5, K)
'                        LI.ListSubItems.Add , , Book(3, K)
                        LI.ListSubItems.Add , , Book(6, K)
                    End If
                    If Force_Stop Then
                        GoTo Exit_query
                        Exit Function
                    End If
                Next
                DoEvents
                stsbar.Value = stsbar.Value + 1
            Next
            
            't = t + 1: If t > 3 Then t = 1
            'stsBar.SimpleText = " " & tt(t) & FormatPercent(i / libCount) & " | " & lib(1, i)
        End If
    Next

Exit_query:
    
    lblStatus.Visible = False
    lvResult.Visible = True
  '  If MainFrm.WindowState = 1 Then MainFrm.WindowState = 0
    MainFrm.Height = MainFrm.Height + lvResult.Height
    Form_Resize
    
    stsbar.Value = 0
    cmdCTSearch.Enabled = True
    cmdStop.Enabled = False
    Force_Stop = False
        
End Function

Private Sub lvResult_DblClick()
    On Error Resume Next
    Dim it As ListItem
    Set it = lvResult.SelectedItem
    
    If it.Selected = False Then Exit Sub
    lvResultItemClick it
End Sub

Private Sub lvResultItemClick(ByVal Item As MSComctlLib.ListItem)

'Dim infoDialog As New Dialog

'Load infoDialog

With infoDialog
.Icon = MainFrm.Icon
If .lblInfo(0).Caption = "" Then .lblInfo(0) = "Title:"
.txtInfo(0) = Item.ListSubItems(1).Text
If .lblInfo(1).Caption = "" Then .lblInfo(1) = "Author:"
.txtInfo(1) = Item.ListSubItems(2).Text
If .lblInfo(2).Caption = "" Then .lblInfo(2) = "Catalog:"
.txtInfo(2) = Item.ListSubItems(3).Text
If .lblInfo(3).Caption = "" Then .lblInfo(3) = "Library:"
.txtInfo(3) = Item.ListSubItems(4).Text
If .lblInfo(4).Caption = "" Then .lblInfo(4) = "Pages:"
.txtInfo(4) = Item.ListSubItems(5).Text
If .lblInfo(5).Caption = "" Then .lblInfo(5) = "SS Number:"
.txtInfo(5) = Item.ListSubItems(6).Text
If .lblInfo(6).Caption = "" Then .lblInfo(6) = "Link:"
'.txtInfo(6) = Item.ListSubItems(7).Text
If .lblInfo(7).Caption = "" Then .lblInfo(7) = "Date:"
.txtInfo(7) = Item.ListSubItems(7).Text

If .txtInfo(5).Text = "" Then .cmdOpenSS.Enabled = False
If InStr(.txtInfo(6).Text, ":") = 2 Or .txtInfo(6).Text = "" Then .cmdOpenInIE.Enabled = False
End With
'
'msg = "Title:" & Item.Text
'msg = msg & vbCrLf & "Author:" & Item.ListSubItems(1).Text
'msg = msg & vbCrLf & "Catalog:" & Item.ListSubItems(2).Text
'msg = msg & vbCrLf & "Library:" & Item.ListSubItems(3).Text
'msg = msg & vbCrLf & "Pages:" & Item.ListSubItems(4).Text
'msg = msg & vbCrLf & "SS Number:" & Item.ListSubItems(5).Text
'msg = msg & vbCrLf & "Link:" & Item.ListSubItems(6).Text
'msg = msg & vbCrLf & "Date:" & Item.ListSubItems(7).Text
''msg = msg & vbCrLf & "Author:" & Item.ListSubItems(8).Text
''msg = msg & vbCrLf & "Author:" & Item.ListSubItems(9).Text


infoDialog.Show 1

'MainFrm.lvResult.SetFocus
End Sub

Private Sub txtDIRSearch_Change()
    Dim fso As New FileSystemObject
    Dim root As String
    Dim i As Long
    root = fso.BuildPath(txtDIRSearch.Text, "liblist.dat")
    If fso.FileExists(root) = False Then
        lblSSReader.Enabled = False
        lblKeyword.Enabled = False
        cmdCTSearch.Enabled = False
        txtKeyword.Enabled = False
        cboLib.Enabled = False
        Exit Sub
    End If
    
    Set fso = Nothing
    
'    Dim tmp
'    tmp = lblSSReader.ForeColor
'    lblSSReader.ForeColor = lblSSReader.BackColor
'    lblSSReader.BackColor = tmp
    
    cmdCTSearch.Enabled = True
    txtKeyword.Enabled = True
    cboLib.Enabled = True
    lblSSReader.Enabled = True
    lblKeyword.Enabled = True
       
    
    loadDataBase root
    
    cboLib.Clear
    cboLib.AddItem "All", 0
    For i = 1 To libCount
        cboLib.AddItem Lib(1, i), i
    Next
    cboLib.ListIndex = 0
    
    For i = 0 To txtDIRSearch.ListCount
    If txtDIRSearch.List(i) = txtDIRSearch.Text Then Exit Sub
    Next
    
    txtDIRSearch.AddItem txtDIRSearch.Text
End Sub

Public Sub loadDataBase(Optional ByRef rootTree As String = "")

    
    If rootTree = "" Then rootTree = libList
    libCount = pdgLibList(rootTree, Lib())
    
End Sub

'Public Function uniqueStr(ByRef baseStr As String) As String
'    Static i As Long
'    i = i + 1
'    uniqueStr = "A" & baseStr & Hex$(i)
'End Function

Private Function match(ByRef Book() As String, ByRef idx As Long) As Boolean
    
    match = True
    
    If MyInstr(Book(1, idx), txtKeyword.Text, , vbTextCompare) = False And _
       MyInstr(Book(2, idx), txtKeyword.Text, , vbTextCompare) = False Then
       match = False
    End If
'
'    If txtKeyword.Text <> "" And InStr(1, Book(1, idx), txtKeyword.Text, vbTextCompare) < 1 Then Exit Function
'    If txtAuthor.Text <> "" And InStr(1, Book(2, idx), txtAuthor.Text, vbTextCompare) < 1 Then Exit Function
'    If txtSS.Text <> "" And InStr(1, Book(5, idx), txtSS.Text, vbTextCompare) < 1 Then Exit Function
'    If txtDate.Text <> "" And InStr(1, Book(6, idx), txtDate.Text, vbTextCompare) < 1 Then Exit Function
'

End Function

Private Sub txtDIRSearch_Click()
    txtDIRSearch_Change
End Sub

Function StrNum(Num As Long, lenNum As Integer) As String

    StrNum = CStr(Num)

    If Len(StrNum) >= lenNum Then
        StrNum = Left$(StrNum, lenNum)
    Else
        StrNum = String$(lenNum - Len(StrNum), "0") + StrNum
    End If

End Function
