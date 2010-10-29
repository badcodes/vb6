VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form MainFrm 
   AutoRedraw      =   -1  'True
   Caption         =   "ssMDBQuery"
   ClientHeight    =   3732
   ClientLeft      =   132
   ClientTop       =   720
   ClientWidth     =   7176
   Icon            =   "MainFrm2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3732
   ScaleWidth      =   7176
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView lvResult 
      Height          =   1152
      Left            =   108
      TabIndex        =   0
      Top             =   2148
      Width           =   6816
      _ExtentX        =   12023
      _ExtentY        =   2032
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
      Height          =   1512
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   6660
      Begin VB.ComboBox txtSSID 
         Enabled         =   0   'False
         Height          =   288
         Left            =   3960
         TabIndex        =   12
         Top             =   552
         Width           =   2028
      End
      Begin VB.ComboBox txtDate 
         Enabled         =   0   'False
         Height          =   288
         Left            =   888
         TabIndex        =   10
         Top             =   588
         Width           =   2028
      End
      Begin VB.ComboBox txtAuthor 
         Enabled         =   0   'False
         Height          =   288
         Left            =   3960
         TabIndex        =   8
         Top             =   192
         Width           =   2028
      End
      Begin VB.CheckBox chkWildcards 
         Caption         =   "Using Wildcards"
         Enabled         =   0   'False
         Height          =   288
         Left            =   3792
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1008
         Value           =   1  'Checked
         Width           =   1524
      End
      Begin VB.ComboBox txtTitle 
         Enabled         =   0   'False
         Height          =   288
         Left            =   900
         TabIndex        =   6
         Top             =   192
         Width           =   2028
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "Search"
         Default         =   -1  'True
         Enabled         =   0   'False
         Height          =   288
         Left            =   5412
         TabIndex        =   3
         Top             =   1020
         Width           =   1188
      End
      Begin VB.ComboBox cboLib 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   288
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1008
         Width           =   2535
      End
      Begin VB.Label lblLib 
         Alignment       =   1  'Right Justify
         Caption         =   "Lib:"
         Enabled         =   0   'False
         Height          =   192
         Left            =   132
         TabIndex        =   14
         Top             =   1056
         Width           =   648
      End
      Begin VB.Label lblSSID 
         Alignment       =   1  'Right Justify
         Caption         =   "SSID:"
         Enabled         =   0   'False
         Height          =   192
         Left            =   3180
         TabIndex        =   13
         Top             =   576
         Width           =   648
      End
      Begin VB.Label lblDate 
         Alignment       =   1  'Right Justify
         Caption         =   "Date:"
         Enabled         =   0   'False
         Height          =   192
         Left            =   144
         TabIndex        =   11
         Top             =   624
         Width           =   648
      End
      Begin VB.Label lblAuthor 
         Alignment       =   1  'Right Justify
         Caption         =   "Author:"
         Enabled         =   0   'False
         Height          =   192
         Left            =   3060
         TabIndex        =   9
         Top             =   228
         Width           =   648
      End
      Begin VB.Label lblTitle 
         Alignment       =   1  'Right Justify
         Caption         =   "Title:"
         Enabled         =   0   'False
         Height          =   192
         Left            =   144
         TabIndex        =   4
         Top             =   228
         Width           =   642
      End
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      Caption         =   "lblStatus"
      Height          =   192
      Left            =   96
      TabIndex        =   5
      Top             =   1956
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFile_open 
         Caption         =   "&Open"
      End
      Begin VB.Menu mnuFile_recent 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnuFile_sep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFile_exit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelp_help 
         Caption         =   "&Help"
      End
      Begin VB.Menu mnuHelp_about 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "MainFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text


Const iniFile = "ssMDBquery.ini"
Public hSet As CAutoSetting
Dim Force_Stop As Boolean
Dim infoDialog As Dialog
Dim sMdbfile As String
    Dim Lib() As String
    Dim libCount As Long




Private Sub cmdSearch_Click()

If txtTitle.Text = "" And _
  txtAuthor.Text = "" And _
  txtSSID.Text = "" And _
  txtDate.Text = "" Then
  Exit Sub
End If

lvResult.ListItems.Clear
'Dim fso As New FileSystemObject
'Dim root As String
'root = fso.BuildPath(txtMDBFile.Text, "liblist.dat")
'Set fso = Nothing

cmdSearch.Enabled = False
'cmdStop.Enabled = True
Force_Stop = False

'Dim lastFor As String
'lastFor = Time$
'queryPdgLib
MComboboxHelper.Add_UniqueItem txtTitle, txtTitle.Text

queryMDB
'MsgBox "Start at " & lastFor & vbCrLf & "End at " + Time$, vbOKCancel

End Sub





'Private Sub cmdStop_Click()
'
''cmdStop.Enabled = False
'cmdSearch.Enabled = True
'
'Force_Stop = True
'
'End Sub

Private Sub Form_Load()

Set infoDialog = Dialog

'Load infoDialog

Set hSet = New CAutoSetting
hSet.fileNameSaveTo = App.Path & "\" & iniFile
Dim i As Long


With hSet
.Add MainFrm, SF_FORM

'Menu
.Add mnuFile, SF_CAPTION
.Add mnuFile_open, SF_CAPTION
.Add mnuFile_open, SF_Tag
.Add mnuFile_exit, SF_CAPTION
.Add mnuHelp, SF_CAPTION
.Add mnuHelp_help, SF_CAPTION
.Add mnuHelp_about, SF_CAPTION
'Recently Used Menu

.Add mnuFile_recent, SF_MENUARRAY, "mnuFile_recent"



.Add txtTitle, SF_LISTTEXT
.Add mnuFile_open, SF_Tag
.Add txtTitle, SF_TEXT
.Add lblTitle, SF_CAPTION ', , True
.Add cmdSearch, SF_CAPTION ', , True

.Add lblAuthor, SF_CAPTION
.Add txtAuthor, SF_LISTTEXT
.Add txtAuthor, SF_TEXT

.Add lblSSID, SF_CAPTION
.Add txtSSID, SF_LISTTEXT
.Add txtSSID, SF_TEXT

.Add lblDate, SF_CAPTION
.Add txtDate, SF_LISTTEXT
.Add txtDate, SF_TEXT

.Add chkWildcards, SF_VALUE
.Add chkWildcards, SF_CAPTION

.Add lblLib, SF_CAPTION

For i = 1 To lvResult.ColumnHeaders.Count
    .Add MainFrm.lvResult.ColumnHeaders(i), SF_WIDTH, "MainFrm.lvResult.ColumnHeaders" & CStr(i) & ".Width"
    .Add MainFrm.lvResult.ColumnHeaders(i), SF_TEXT, "Mainfrm.lvResult.ColumnHeaders" & CStr(i) & ".Text" ', True
Next


.Add infoDialog, SF_FORM
For i = 1 To infoDialog.lblInfo.Count
    .Add infoDialog.lblInfo(i - 1), SF_CAPTION ', , True
Next

.Add infoDialog.cmdOpenInIE, SF_CAPTION ', , True
.Add infoDialog.cmdOpenSS, SF_CAPTION ', , True
.Add infoDialog.OKButton, SF_CAPTION ', , True


'For i = 0 To chkKeyIN.Count - 1
'    .Add MainFrm.chkKeyIN(i), SF_VALUE ', "MainFrm.chkKeyIN" & CStr(i)
'Next
End With


'Dim tmp
'tmp = lblMDBPath.ForeColor
'lblMDBPath.ForeColor = lblMDBPath.BackColor
'lblMDBPath.BackColor = tmp


'txtMDBFile_Change

sMdbfile = mnuFile_open.Tag
Dim mnuHandler As CMenuArrHandle
Set mnuHandler = New CMenuArrHandle
    mnuHandler.Menus = mnuFile_recent
    mnuHandler.maxItem = 10
    mnuHandler.AddUnique sMdbfile
Set mnuHandler = Nothing

loadMDB sMdbfile
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    Dim w As Long
    
    'lblAbout.Left = MainFrm.ScaleWidth - lblAbout.Width - 240
    

    
    With frmProcess
        .Top = 0
        .Left = c_w_space
        .Width = MainFrm.ScaleWidth - .Left * 2
        w = .Width - lblTitle.Width - lblAuthor.Width - 5 * c_w_space

        txtTitle.Width = w / 2
        txtAuthor.Width = w / 2
        txtSSID.Width = w / 2
        txtDate.Width = w / 2
        cboLib.Width = w / 2
        
        'cboLib.Width = .Width - chkWildcards.Width - cmdSearch.Width - 4 * 120
   End With
  
    
    'Line 1
    'move_TopLeft lblTitle
    move_RightTo lblTitle, txtTitle
    move_RightTo txtTitle, lblAuthor
    move_RightTo lblAuthor, txtAuthor
    'Line 2
    move_Below lblTitle, lblDate
    move_Align lblTitle, lblDate, E_AP_LEFT
    move_RightTo lblDate, txtDate
    move_RightTo txtDate, lblSSID
    move_RightTo lblSSID, txtSSID
    'line 3
    move_Below lblDate, lblLib
    move_Align lblDate, lblLib, E_AP_LEFT
    move_RightTo lblLib, cboLib
    

    move_Below txtSSID, cmdSearch
    move_Align lblLib, cmdSearch, E_AP_TOP
    move_Align txtSSID, cmdSearch, E_AP_RIGHT
    move_LeftTo cmdSearch, chkWildcards
    
           
    With lblStatus
        .Left = frmProcess.Left
        .Width = frmProcess.Width
        .Top = frmProcess.Top + frmProcess.Height + 120
    End With
        
    With lvResult
    .Top = lblStatus.Top + 120
    .Left = frmProcess.Left
    .Width = frmProcess.Width
    .Height = MainFrm.ScaleHeight - .Top - 240
'    lblStatus.Top = .Top - 100
'    lblStatus.Left = .Left
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
    Unload Me
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

Public Sub queryMDB()


'    Dim Cata() As String
'    Dim cataCount As Long
'    Dim Book() As String
'    Dim bookCount As String
    Dim LI As ListItem
    Dim c As Long
    Dim mHeight As Long
    
    On Error Resume Next
    
    If MainFrm.WindowState = 1 Then MainFrm.WindowState = 0
    
    mHeight = MainFrm.Height
    
    MainFrm.Height = MainFrm.Height - lvResult.Height + 120
    
    lvResult.Visible = False
    lblStatus.Visible = True
    
    Dim mdbFile As String
    Dim db As DBEngine
    Dim dbase As Database
    Dim rc As Recordset
    Dim fs As Fields
    Dim strSearch As String
    'Dim bMatch As Boolean
    
    
    mdbFile = sMdbfile '.Text
    
    
    
    Set db = New DBEngine
    Set dbase = db.openDatabase(mdbFile)
    
    'MDebug.DebugFPrint "Test " & mdbfile
    'MDebug.DebugFPrint "Search for " & strSearch
    
    Dim i As Long
    For i = 1 To libCount
        
        'MDebug.DebugFPrint "At : " & Lib(i)
        
        If i = cboLib.ListIndex Or cboLib.ListIndex = 0 Then
            Dim strQuery As String
            strQuery = buildQuery(Lib(i))
            'Set rc = dbase.OpenRecordset(Lib(i), dbOpenForwardOnly)
            Set rc = dbase.OpenRecordset(strQuery, dbOpenForwardOnly)
            lblStatus.Caption = "looking in " & Lib(i)
            If rc Is Nothing Then GoTo For_continue
            If Not rc Is Nothing And Not rc.BOF Then
                
'                stsbar.Value = 0
'                stsbar.Min = 0
'                rc.MoveLast
'                If rc.RecordCount > 0 Then stsbar.Max = 100
'                rc.MoveFirst
                
                'MDebug.DebugFPrint "Enter " & Lib(i)
                'MDebug.DebugFPrint "RecordCount: " & rc.RecordCount
                'MDebug.DebugFPrint "FieldCount: " & rc.Fields.Count
                
                'If Not rc.EOF Then rc.MoveFirst
                Do Until rc.EOF
                    'stsbar.Value = (rc.AbsolutePosition / rc.RecordCount) * 100
                    Set fs = rc.Fields
'                    If fs("title") Like strSearch Or _
'                       fs("author") Like strSearch Or _
'                       fs("ssid") Like strSearch Or _
'                       fs("date") Like strSearch Then
                    'match(fs, strSearch) Then
                        c = c + 1
                        Set LI = lvResult.ListItems.Add(, , StrNum(c, 5))
                       ' MDebug.DebugFPrint "Match At " & rc.AbsolutePosition & "¡¶" & fs("title") & "¡·"
                        LI.ListSubItems.Add , , fs("title")
                        LI.ListSubItems.Add , , fs("author")
                        LI.ListSubItems.Add , , fs("catalog")
                        LI.ListSubItems.Add , , Lib(i)
                        LI.ListSubItems.Add , , fs("pages")
                        LI.ListSubItems.Add , , fs("ssid")
                        LI.ListSubItems.Add , , fs("date")
'                    End If
                    rc.MoveNext
'                    If Err.Number <> 0 Then
'                        MsgBox rc.Name & "Position:" & rc.AbsolutePosition & " ÓÐ´íÎó¡£" & vbCrLf & Err.Description, vbCritical
'                        Err.Clear
'                        GoTo Out_Of_Loop
'                    End If
'                    'DoEvents
'                    If Force_Stop Then
'                        GoTo Exit_query
'                        Exit Sub
'                    End If
                Loop
            End If
            'MDebug.DebugFPrint "RecordCount: " & rc.RecordCount
            rc.Close
            Set rc = Nothing
        End If
        'MDebug.DebugFPrint "Leave " & Lib(i)
For_continue:
    Next

Exit_query:
    
    lblStatus.Visible = False
    lvResult.Visible = True
    If MainFrm.WindowState = 1 Then MainFrm.WindowState = 0
    MainFrm.Height = mHeight 'MainFrm.Height + lvResult.Height

    
    'stsbar.Value = 0
    cmdSearch.Enabled = True
    'cmdStop.Enabled = False
    Force_Stop = False
    
    rc.Clone
    Set rc = Nothing
    dbase.Close
    Set dbase = Nothing
    Set db = Nothing
    
    Form_Resize
End Sub

'Public Sub queryPdgLib()
'
'
''    Dim Cata() As String
''    Dim cataCount As Long
''    Dim Book() As String
''    Dim bookCount As String
'    Dim LI As ListItem
'    Dim c As Long
'
'    On Error Resume Next
'
'    'If MainFrm.WindowState = 1 Then MainFrm.WindowState = 0
'
'    MainFrm.Height = MainFrm.Height - lvResult.Height + 120
'    lvResult.Visible = False
'    lblStatus.Visible = True
'
'    Dim mdbfile As String
'    Dim db As DBEngine
'    Dim dbase As Database
'    Dim rc As Recordset
'    Dim fs As Fields
'    Dim strSearch As String
'    Dim bMatch As Boolean
'
'
'    mdbfile = txtMDBFile.Text
'    strSearch = "title LIKE '" & txtTitle.Text & "'" & _
'                " OR author LIKE '" & txtTitle.Text & "'" & _
'                " OR ssid LIKE '" & txtTitle.Text & "'" & _
'                " OR date LIKE '" & txtTitle.Text & "'"
'
'    Set db = New DBEngine
'    Set dbase = db.OpenDatabase(mdbfile)
'
'    MDebug.DebugFPrint "Test " & mdbfile
'    MDebug.DebugFPrint "Search for " & strSearch
'
'    Dim i As Long
'    For i = 1 To libCount
'
'        MDebug.DebugFPrint "At : " & Lib(i)
'
'        If i = cboLib.ListIndex Or cboLib.ListIndex = 0 Then
'
'            Set rc = dbase.OpenRecordset(Lib(i), dbOpenSnapshot)
'            lblStatus.Caption = "looking in " & Lib(i)
'            If rc.AbsolutePosition > -1 Then
'
'                'stsbar.Value = 0
'                'stsbar.Min = 0
'                'rc.MoveLast
'                'If rc.RecordCount > 0 Then stsbar.Max = 100
'                'rc.MoveFirst
'
'                MDebug.DebugFPrint "Enter " & Lib(i)
'                'MDebug.DebugFPrint "RecordCount: " & rc.RecordCount
'                MDebug.DebugFPrint "FieldCount: " & rc.Fields.Count
'
'                rc.FindFirst strSearch
'
'                If rc.NoMatch Then
'                    MDebug.DebugFPrint "Keyword NoMatch."
'                Else
'                    MDebug.DebugFPrint "Keyword Match Some"
'                End If
'
'                'bMatch = Not rc.NoMatch
'
'                Do Until rc.NoMatch
'                   ' stsbar.Value = (rc.AbsolutePosition / rc.RecordCount) * 100
'                    Set fs = rc.Fields
'                    c = c + 1
'                    Set LI = lvResult.ListItems.Add(, , StrNum(c, 5))
'                    MDebug.DebugFPrint "Match At " & rc.AbsolutePosition & "¡¶" & fs("title") & "¡·"
'                    LI.ListSubItems.Add , , fs("title")
'                    LI.ListSubItems.Add , , fs("author")
'                    LI.ListSubItems.Add , , fs("catalog")
'                    LI.ListSubItems.Add , , rc.Name
'                    LI.ListSubItems.Add , , fs("pages")
'                    LI.ListSubItems.Add , , fs("ssid")
'                    LI.ListSubItems.Add , , fs("date")
'                    If Force_Stop Then
'                        GoTo Exit_query
'                        Exit Sub
'                    End If
'                    rc.FindNext strSearch
'                    DoEvents
'                    If Err.Number <> 0 Then
'                        MsgBox rc.Name & "Position:" & rc.AbsolutePosition & " ÓÐ´íÎó¡£" & vbCrLf & Err.Description, vbCritical
'                        Err.Clear
'                        GoTo Out_Of_Loop
'                    End If
'                Loop
'Out_Of_Loop:
'
'                'stsbar.Value = stsbar.Max
'              End If
'            MDebug.DebugFPrint "RecordCount: " & rc.RecordCount
'            rc.Close
'            Set rc = Nothing
'        End If
'        MDebug.DebugFPrint "Leave " & Lib(i)
'    Next
'
'Exit_query:
'
'    lblStatus.Visible = False
'    lvResult.Visible = True
'  '  If MainFrm.WindowState = 1 Then MainFrm.WindowState = 0
'    MainFrm.Height = MainFrm.Height + lvResult.Height
'
'
'    'stsbar.Value = 0
'    cmdSearch.Enabled = True
'    cmdStop.Enabled = False
'    Force_Stop = False
'
'    rc.Clone
'    Set rc = Nothing
'    dbase.Close
'    Set dbase = Nothing
'    Set db = Nothing
'
'    Form_Resize
'End Sub

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
If .lblInfo(5).Caption = "" Then .lblInfo(5) = "SSID:"
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


Private Sub loadMDB(ByRef sFile As String)
    Dim fso As New FileSystemObject
    Dim i As Long
    Dim o As Object

'    On Error Resume Next
    
'    Dim tmp
'    tmp = lblMDBPath.ForeColor
'    lblMDBPath.ForeColor = lblMDBPath.BackColor
'    lblMDBPath.BackColor = tmp
    
    
    
    If loadDataBase(sFile) = False Then Exit Sub
    
    sMdbfile = sFile
    mnuFile_open.Tag = sMdbfile
    MainFrm.Caption = sMdbfile & " - " & App.ProductName
    
        'lblMDBPath.Enabled = False
'        For Each o In MainFrm.Controls
'            Debug.Print o.Name
'        If TypeName(o) <> "Menu" Then o.Enabled = False
'        Next
'        lblTitle.Enabled = False
'        cmdSearch.Enabled = False
'        txtTitle.Enabled = False
'        cboLib.Enabled = False
'        chkWildcards.Enabled = False
'       Exit Sub
'    End If
    
    For Each o In MainFrm.Controls
        o.Enabled = True
    Next
    cmdSearch.Enabled = True
    txtTitle.Enabled = True
    cboLib.Enabled = True
    'lblMDBPath.Enabled = True
    lblTitle.Enabled = True
    chkWildcards.Enabled = True
    

    
    cboLib.Clear
    cboLib.AddItem "All", 0
    For i = 1 To libCount
        cboLib.AddItem Lib(i), i
    Next
    cboLib.ListIndex = 0
    
    'MComboboxHelper.Add_UniqueItem txtMDBFile, txtMDBFile.Text
    
End Sub

Public Function loadDataBase(ByRef rootTree As String) As Boolean

    'If rootTree = "" Then rootTree = libList
    Dim iCount As Long
    Dim sLib() As String
    'Dim i As Long
    If rootTree = "" Then Exit Function
    iCount = MDAO.getTabledefs(rootTree, sLib())
    If iCount < 1 Then Exit Function
    
    
    libCount = iCount
    ReDim Lib(1 To libCount) As String
     For iCount = 1 To libCount
        Lib(iCount) = sLib(iCount)
    Next
    
    loadDataBase = True
End Function

'Public Function uniqueStr(ByRef baseStr As String) As String
'    Static i As Long
'    i = i + 1
'    uniqueStr = "A" & baseStr & Hex$(i)
'End Function

'Private Function match(ByRef fs As Fields, ByRef keyword As String) As Boolean
'
'    match = True
'
'    On Error Resume Next
'
'    If fs("title") Like keyword Then Exit Function
'    If fs("author") Like keyword Then Exit Function
'    If fs("ssid") Like keyword Then Exit Function
'    If fs("date") Like keyword Then Exit Function
'
'
'    match = False
'
''    If MyInstr(Book(1, idx), txtTitle.Text, , vbTextCompare) = False And _
''       MyInstr(Book(2, idx), txtTitle.Text, , vbTextCompare) = False Then
''       match = False
''    End If
''
''    If txtTitle.Text <> "" And InStr(1, Book(1, idx), txtTitle.Text, vbTextCompare) < 1 Then Exit Function
''    If txtAuthor.Text <> "" And InStr(1, Book(2, idx), txtAuthor.Text, vbTextCompare) < 1 Then Exit Function
''    If txtSS.Text <> "" And InStr(1, Book(5, idx), txtSS.Text, vbTextCompare) < 1 Then Exit Function
''    If txtDate.Text <> "" And InStr(1, Book(6, idx), txtDate.Text, vbTextCompare) < 1 Then Exit Function
''
'
'End Function



Function StrNum(Num As Long, lenNum As Integer) As String

    StrNum = CStr(Num)

    If Len(StrNum) >= lenNum Then
        StrNum = Left$(StrNum, lenNum)
    Else
        StrNum = String$(lenNum - Len(StrNum), "0") + StrNum
    End If

End Function

Private Sub mnuFile_Exit_Click()
    Force_Stop = True
    Unload Me
End Sub

Private Sub mnuFile_open_Click()
    Dim mdbFile As String
    Dim dlg As CCommonDialogLite
    Dim mnuHandler As CMenuArrHandle
    
    Set dlg = New CCommonDialogLite
    
    If dlg.VBGetOpenFileName(filename:=mdbFile, Filter:="mdb File (*.mdb)|*.mdb") Then
    
        loadMDB mdbFile
        
        Set mnuHandler = New CMenuArrHandle
        mnuHandler.Menus = mnuFile_recent
        mnuHandler.maxItem = 10
        mnuHandler.AddUnique mdbFile
        Set mnuHandler = Nothing
        
        'txtMDBFile.Text = mdbFile
        'Call txtMDBFile_Change
    End If
End Sub

Private Sub mnuFile_recent_Click(Index As Integer)
    loadMDB mnuFile_recent(Index).Tag
End Sub

Private Sub mnuHelp_About_Click()
 MsgBox App.ProductName & " V" & App.Major & "." & App.Minor & "." & App.Revision & _
            " by " & App.LegalCopyright & "@" & App.CompanyName
End Sub

Private Sub mnuHelp_Help_Click()
    Dim helpFile As String
    helpFile = App.Path & "\" & "help.htm"
    ShellExecute Me.hWnd, "open", helpFile, "", "", 1
End Sub

Private Function buildQuery(sLib As String) As String
Dim strQuery As String
    strQuery = "SELECT * FROM " & sLib
    If txtTitle.Text <> "" Then strQuery = strQuery & " WHERE title LIKE " & getStrSearch(txtTitle.Text)
    If txtAuthor.Text <> "" Then strQuery = strQuery & " AND WHERE author LIKE " & getStrSearch(txtAuthor.Text)
    If txtSSID.Text <> "" Then strQuery = strQuery & " AND WHERE ssid LIKE " & getStrSearch(txtSSID.Text)
    If txtDate.Text <> "" Then strQuery = strQuery & " AND WHERE date LIKE " & getStrSearch(txtDate.Text)
    
    buildQuery = Replace$(strQuery, sLib & " AND ", sLib & " ")
    buildQuery = Replace$(buildQuery, "AND WHERE", "AND")



End Function
Private Function getStrSearch(ByRef sText As String) As String
    If chkWildcards.Value = 1 Then
        getStrSearch = Chr(34) & sText & Chr(34)
    Else
        getStrSearch = Chr(34) & "*" & sText & "*" & Chr(34)
    End If
End Function
