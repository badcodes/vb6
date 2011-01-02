VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form MainFrm 
   AutoRedraw      =   -1  'True
   Caption         =   "pdgLibQuery"
   ClientHeight    =   7485
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9255
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7485
   ScaleWidth      =   9255
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frmProcess 
      Height          =   1575
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   8775
      Begin VB.CheckBox chkKeyIN 
         Caption         =   "Date"
         Height          =   255
         Index           =   3
         Left            =   6120
         TabIndex        =   14
         Top             =   1200
         Width           =   1300
      End
      Begin VB.CheckBox chkKeyIN 
         Caption         =   "SS Number"
         Height          =   255
         Index           =   2
         Left            =   4440
         TabIndex        =   13
         Top             =   1200
         Width           =   1300
      End
      Begin VB.CheckBox chkKeyIN 
         Caption         =   "Author"
         Height          =   255
         Index           =   1
         Left            =   3000
         TabIndex        =   12
         Top             =   1200
         Width           =   1300
      End
      Begin VB.CheckBox chkKeyIN 
         Caption         =   "Title"
         Height          =   255
         Index           =   0
         Left            =   1680
         TabIndex        =   11
         Top             =   1200
         Width           =   1300
      End
      Begin VB.TextBox txtDIRSearch 
         Height          =   315
         Left            =   2040
         TabIndex        =   6
         Text            =   "D:\Read\SSREADER39\remote112"
         Top             =   240
         Width           =   6495
      End
      Begin VB.TextBox txtCT 
         Height          =   315
         Left            =   1080
         TabIndex        =   5
         Top             =   720
         Width           =   2415
      End
      Begin VB.CommandButton cmdCTSearch 
         Caption         =   "Search"
         Default         =   -1  'True
         Height          =   315
         Left            =   6480
         TabIndex        =   4
         Top             =   720
         Width           =   975
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "Stop"
         Height          =   315
         Left            =   7680
         TabIndex        =   3
         Top             =   720
         Width           =   855
      End
      Begin VB.ComboBox cboLib 
         Height          =   315
         Left            =   3720
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   720
         Width           =   2535
      End
      Begin VB.Label Label1 
         Caption         =   "Searching In Fields :"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label lblKeyword 
         AutoSize        =   -1  'True
         Caption         =   "KeyWord:"
         Height          =   315
         Left            =   120
         TabIndex        =   8
         Top             =   750
         Width           =   705
      End
      Begin VB.Label lblSSReader 
         AutoSize        =   -1  'True
         Caption         =   "SSReader Library Path:"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1665
      End
   End
   Begin MSComctlLib.ListView lvResult 
      Height          =   4935
      Left            =   240
      TabIndex        =   0
      Top             =   2040
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   8705
      View            =   3
      LabelEdit       =   1
      SortOrder       =   -1  'True
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      GridLines       =   -1  'True
      TextBackground  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "title"
         Text            =   "Title"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "author"
         Text            =   "Author"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "catalog"
         Text            =   "Catalog"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Key             =   "library"
         Text            =   "Library"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "pageCount"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "ss number"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "link"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "date"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ProgressBar stsbar 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   9
      Top             =   7170
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   556
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      Caption         =   "lblStatus"
      Height          =   195
      Left            =   240
      TabIndex        =   10
      Top             =   1800
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
Const iniFile = "pdgLibquery.ini"
Dim hset As CAutoSetting
Dim Force_Stop As Boolean

    Dim Lib() As String
    Dim libCount As Long


Private Sub cmdctsearch_Click()

If chkKeyIN(0) + chkKeyIN(1) + chkKeyIN(2) + chkKeyIN(3) = 0 Then
    MsgBox "No fields checked. Stop querying.", vbOKOnly
    Exit Sub
End If

lvResult.ListItems.Clear
Dim fso As New FileSystemObject
Dim root As String
root = fso.BuildPath(txtDIRSearch.Text, "liblist.dat")
Set fso = Nothing

cmdCTSearch.Enabled = False
cmdStop.Enabled = True
Force_Stop = False

queryPdgLib root, txtCT.Text




End Sub



Private Sub cmdStop_Click()

cmdStop.Enabled = False
cmdCTSearch.Enabled = True

Force_Stop = True

End Sub

Private Sub Form_Load()
Set hset = New CAutoSetting
hset.fileNameSaveTo = App.Path & "\" & iniFile
Dim i As Long

With hset
.Add MainFrm, SF_FORM
.Add txtCT, SF_TEXT
.Add txtDIRSearch, SF_TEXT
For i = 1 To lvResult.ColumnHeaders.Count
    .Add MainFrm.lvResult.ColumnHeaders(i), SF_WIDTH, "MainFrm.lvResult.ColumnHeaders" & CStr(i)
Next
For i = 0 To chkKeyIN.Count - 1
    .Add MainFrm.chkKeyIN(i), SF_VALUE ', "MainFrm.chkKeyIN" & CStr(i)
Next
End With

txtDIRSearch_Change
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    With frmProcess
        .Width = MainFrm.ScaleWidth - .Left * 2
        txtDIRSearch.Left = lblSSReader.Left + lblSSReader.Width + 120
        txtDIRSearch.Width = .Width - txtDIRSearch.Left - 120
        lblKeyword.Left = lblSSReader.Left
        txtCT.Left = lblKeyword.Left + lblKeyword.Width + 120
        txtCT.Width = .Width - txtCT.Left - cboLib.Width - cmdCTSearch.Width - cmdStop.Width - 4 * 120
        cboLib.Left = txtCT.Left + txtCT.Width + 120
        cmdCTSearch.Left = cboLib.Left + cboLib.Width + 120
        cmdStop.Left = cmdCTSearch.Left + cmdCTSearch.Width + 120
        
        
    End With
        
    With lvResult
    .Left = frmProcess.Left
    .Width = frmProcess.Width
    .Height = stsbar.Top - .Top - 120
    lblStatus.Top = .Top - 100
    lblStatus.Left = .Left
    End With
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
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Force_Stop = True
    Set hset = Nothing
    Unload MainFrm
End Sub

Private Sub lblProgress_Click()

End Sub

Private Sub lvResult_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lvResult
    .SortKey = ColumnHeader.Index - 1
    If .SortOrder = lvwAscending Then .SortOrder = lvwDescending Else .SortOrder = lvwAscending
    .Sorted = True
    End With
    
End Sub



Public Function queryPdgLib(Optional ByRef rootTree As String = "", Optional ByRef strQuery As String = "") As String


    Dim Cata() As String
    Dim cataCount As Long
    Dim Book() As String
    Dim bookCount As String
    Dim LI As ListItem
    Dim per As Single
    
    
    If rootTree = "" Then rootTree = libList
    If strQuery = "" Then Exit Function
    
    
    MainFrm.Height = MainFrm.Height - lvResult.Height
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
                Dim k As Long
                For k = 1 To bookCount
                
                    If match(strQuery, Book(), k) Then 'InStr(Book(1, k), strQuery) > 0 Or _
                       'InStr(Book(2, k), strQuery) > 0 Then
                        Set LI = lvResult.ListItems.Add(, uniqueStr(Book(1, k)), Book(1, k))
                        LI.ListSubItems.Add , uniqueStr(Book(2, k)), Book(2, k)
                        LI.ListSubItems.Add , uniqueStr(Cata(1, j)), Cata(1, j)
                        LI.ListSubItems.Add , uniqueStr(Lib(1, i)), Lib(1, i)
                        LI.ListSubItems.Add , uniqueStr(Book(4, k)), Book(4, k)
                        LI.ListSubItems.Add , uniqueStr(Book(5, k)), Book(5, k)
                        LI.ListSubItems.Add , uniqueStr(Book(3, k)), Book(3, k)
                        LI.ListSubItems.Add , uniqueStr(Book(6, k)), Book(6, k)
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
    MainFrm.Height = MainFrm.Height + lvResult.Height
    Form_Resize
    
    stsbar.Value = 0
    cmdCTSearch.Enabled = True
    cmdStop.Enabled = False
    Force_Stop = False
        
End Function

Private Sub lvResult_ItemClick(ByVal Item As MSComctlLib.ListItem)

Dim infoDialog As New Dialog

Load infoDialog

With infoDialog
.lblInfo(0) = "Title:"
.txtInfo(0) = Item.Text
.lblInfo(1) = "Author:"
.txtInfo(1) = Item.ListSubItems(1).Text
.lblInfo(2) = "Catalog:"
.txtInfo(2) = Item.ListSubItems(2).Text
.lblInfo(3) = "Library:"
.txtInfo(3) = Item.ListSubItems(3).Text
.lblInfo(4) = "Pages:"
.txtInfo(4) = Item.ListSubItems(4).Text
.lblInfo(5) = "SS Number:"
.txtInfo(5) = Item.ListSubItems(5).Text
.lblInfo(6) = "Link:"
.txtInfo(6) = Item.ListSubItems(6).Text
.lblInfo(7) = "Date:"
.txtInfo(7) = Item.ListSubItems(7).Text
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

infoDialog.Show 1, Me

End Sub

Private Sub txtDIRSearch_Change()
    Dim fso As New FileSystemObject
    Dim root As String
    Dim i As Long
    root = fso.BuildPath(txtDIRSearch.Text, "liblist.dat")
    If fso.FileExists(root) = False Then Exit Sub
    Set fso = Nothing
    
    
    loadDataBase root
    
    cboLib.Clear
    cboLib.AddItem "All", 0
    For i = 1 To libCount
        cboLib.AddItem Lib(1, i), i
    Next
    cboLib.ListIndex = 0
    
End Sub

Public Sub loadDataBase(Optional ByRef rootTree As String)

    
    If rootTree = "" Then rootTree = libList
    libCount = pdgLibList(rootTree, Lib())
    
End Sub

Public Function uniqueStr(ByRef baseStr) As String
    Static i As Long
    i = i + 1
    uniqueStr = baseStr & "A" & Hex(i)
End Function

Private Function match(ByRef strQuery As String, ByRef Book() As String, ByRef idx As Long) As Boolean
    If chkKeyIN(0).Value = 1 And InStr(Book(1, idx), strQuery) Then
        match = True
        Exit Function
    End If
    If chkKeyIN(1).Value = 1 And InStr(Book(2, idx), strQuery) Then
        match = True
        Exit Function
    End If
    If chkKeyIN(2).Value = 1 And InStr(Book(5, idx), strQuery) Then
        match = True
        Exit Function
    End If
    If chkKeyIN(3).Value = 1 And InStr(Book(6, idx), strQuery) Then
        match = True
        Exit Function
    End If
End Function
