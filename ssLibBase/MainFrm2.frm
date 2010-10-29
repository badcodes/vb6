VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form MainFrm 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ssLibBase"
   ClientHeight    =   2085
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   6525
   Icon            =   "MainFrm2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2085
   ScaleWidth      =   6525
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frmProcess 
      Height          =   1212
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6132
      Begin VB.ComboBox txtDIRSearch 
         Appearance      =   0  'Flat
         Height          =   288
         Left            =   2040
         TabIndex        =   6
         Top             =   240
         Width           =   3912
      End
      Begin VB.CommandButton cmdMakeMDB 
         Caption         =   "Make MDB"
         Default         =   -1  'True
         Height          =   315
         Left            =   3840
         TabIndex        =   2
         Top             =   720
         Width           =   975
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "Stop"
         Enabled         =   0   'False
         Height          =   315
         Left            =   5040
         TabIndex        =   1
         Top             =   720
         Width           =   855
      End
      Begin VB.Label lblSSReader 
         AutoSize        =   -1  'True
         Caption         =   "SSReader Library Path:"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1665
      End
   End
   Begin MSComctlLib.ProgressBar stsbar 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   4
      Top             =   1785
      Width           =   6525
      _ExtentX        =   11509
      _ExtentY        =   529
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label lblstatus 
      Height          =   192
      Left            =   240
      TabIndex        =   7
      Top             =   1560
      Width           =   6132
   End
   Begin VB.Label lblAbout 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "CopyRight"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   192
      Left            =   240
      TabIndex        =   5
      Top             =   48
      Width           =   6108
   End
End
Attribute VB_Name = "MainFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const iniFile = "ssLibBase.ini"
Dim hSet As CAutoSetting
Dim Force_Stop As Boolean


    Dim Lib() As String
    Dim libCount As Long




Private Sub cmdMakeMDB_Click()


'lvResult.ListItems.Clear
'Dim fso As New FileSystemObject
'Dim root As String
'root = fso.BuildPath(txtDIRSearch.Text, "liblist.dat")
'Set fso = Nothing
Dim dlg As CCommonDialogLite
Dim mdbFile As String

Set dlg = New CCommonDialogLite

If dlg.VBGetSaveFileName(mdbFile) Then
    cmdMakeMDB.Enabled = False
    cmdStop.Enabled = True
    Force_Stop = False
    makeMDB mdbFile
End If


End Sub



'Private Sub cmdHelp_Click()
'Dim helpFile As String
'helpFile = App.Path & "\" & "help.htm"
'ShellExecute Me.hWnd, "open", helpFile, "", "", 1
'End Sub

Private Sub cmdStop_Click()

cmdStop.Enabled = False
cmdMakeMDB.Enabled = True

Force_Stop = True

End Sub

Private Sub Form_Load()



'Load infoDialog

Set hSet = New CAutoSetting
hSet.fileNameSaveTo = App.Path & "\" & iniFile
Dim i As Long


With hSet
'.Add MainFrm, SF_FORM
'.Add txtKeyword, SF_TEXT
.Add txtDIRSearch, SF_LISTTEXT
.Add txtDIRSearch, SF_TEXT
.Add lblSSReader, SF_CAPTION, , True
'.Add lblKeyword, SF_CAPTION, , True
'.Add cmdClear, SF_CAPTION, , True
'.Add cmdHelp, SF_CAPTION, , True
.Add cmdMakeMDB, SF_CAPTION, , True
.Add cmdStop, SF_CAPTION, , True

'For i = 1 To lvResult.ColumnHeaders.Count
'    .Add MainFrm.lvResult.ColumnHeaders(i), SF_WIDTH, "MainFrm.lvResult.ColumnHeaders" & CStr(i) & ".Width"
'    .Add MainFrm.lvResult.ColumnHeaders(i), SF_TEXT, "Mainfrm.lvResult.ColumnHeaders" & CStr(i) & ".Text", True
'Next


'.Add infoDialog, SF_FORM
'For i = 1 To infoDialog.lblInfo.Count
'    .Add infoDialog.lblInfo(i - 1), SF_CAPTION, , True
'Next

'.Add infoDialog.cmdOpenInIE, SF_CAPTION, , True
'.Add infoDialog.cmdOpenSS, SF_CAPTION, , True
'.Add infoDialog.OKButton, SF_CAPTION, , True


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

'Private Sub Form_Resize()
'    On Error Resume Next
'
'    Dim w As Long
'
'    lblAbout.Left = MainFrm.ScaleWidth - lblAbout.Width - 240
'    With frmProcess
'        .Width = MainFrm.ScaleWidth - .Left * 2
'        txtDIRSearch.Left = lblSSReader.Left + lblSSReader.Width + 120
'        txtDIRSearch.Width = .Width - txtDIRSearch.Left - cmdHelp.Width - cmdClear.Width - 3 * 120
'
'
'        cmdClear.Left = txtDIRSearch.Left + txtDIRSearch.Width + 120
'        cmdHelp.Left = cmdClear.Left + cmdClear.Width + 120
'
'
'        lblKeyword.Left = lblSSReader.Left
'        txtKeyword.Left = lblKeyword.Left + lblKeyword.Width + 120
'
''        w = .Width - lblKeyword.Width - lblAuthor.Width - cboLib.Width - 6 * 120
''        txtKeyword.Width = w / 2 '.Width - txtKeyword.Left - cboLib.Width - cmdMakeMDB.Width - cmdStop.Width - 4 * 120
''
''        lblAuthor.Left = txtKeyword.Left + txtKeyword.Width + 120
''
''        txtAuthor.Left = lblAuthor.Left + lblAuthor.Width + 120
''
''        txtAuthor.Width = w / 2
''
'
'        txtKeyword.Width = .Width - lblKeyword.Width - cboLib.Width - cmdMakeMDB.Width - cmdStop.Width - 6 * 120
'
'        cboLib.Left = txtKeyword.Left + txtKeyword.Width + 120
'
'        cmdMakeMDB.Left = cboLib.Left + cboLib.Width + 120
'        cmdStop.Left = cmdMakeMDB.Left + cmdMakeMDB.Width + 120
'
'
'    End With
'
'
'    With lvResult
'    .Left = frmProcess.Left
'    .Width = frmProcess.Width
'    .Height = stsbar.Top - .Top
'    lblStatus.Top = .Top - 100
'    lblStatus.Left = .Left
'    End With
'
''
''    With stsbar
''    .Left = frmProcess.Left
''    .Width = frmProcess.Width
''    .Top = lvResult.Top
''    End With
''    With lvResult.ColumnHeaders
''        Dim lall As Single
''        lall = lvResult.Width - 400
''        .Item(1).Width = 0.4 * lall
''        .Item(2).Width = 0.3 * lall
''        .Item(3).Width = 0.15 * lall
''        .Item(4).Width = 0.15 * lall
''
''    End With
'End Sub

Private Sub Form_Terminate()
    Force_Stop = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Force_Stop = True
    Set hSet = Nothing

    'Unload infoDialog
    'Unload Dialog
    Unload Me
End Sub



'Private Sub lvResult_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'    With lvResult
'    .SortKey = ColumnHeader.Index - 1
'    If .SortOrder = lvwAscending Then .SortOrder = lvwDescending Else .SortOrder = lvwAscending
'    .Sorted = True
'    End With
'
'End Sub



Public Sub makeMDB(ByRef mdbFile As String)


    Dim Cata() As String
    Dim cataCount As Long
    Dim Book() As String
    Dim bookCount As String
    'Dim LI As ListItem
    'Dim c As Long
    
    
    'If MainFrm.WindowState = 1 Then MainFrm.WindowState = 0
    
'    MainFrm.Height = MainFrm.Height - lvResult.Height + 120
'    lvResult.Visible = False
'    lblStatus.Visible = True
'
    Dim db As DBEngine
    Dim dbase As Database
    Dim rs As Recordset
    Dim fs As Fields
    'Dim f As Field
    Dim fso As FileSystemObject
    
    On Error Resume Next
    
    
    Set fso = New FileSystemObject
    
    If fso.FileExists(mdbFile) Then fso.DeleteFile mdbFile, True
    
    
    Set fso = Nothing
    Set db = New DBEngine
    
    Set dbase = db.CreateDatabase(mdbFile, dbLangChineseSimplified, dbVersion40)
        
    If Err.Number <> 0 Then
        MsgBox "创建数据库文件" & mdbFile & " 时发生以下错误:" & _
              vbCrLf & Err.Description, _
              vbCritical
        Err.Clear
        GoTo Exit_query
    End If
        
    Dim i As Long
    For i = 1 To libCount
    
        Set rs = newTable(dbase, Lib(1, i))
        If rs Is Nothing Then
            MsgBox "创建表" & Lib(1, i) & "时,产生错误。" & vbCrLf & Err.Description, vbCritical
            Err.Clear
            GoTo Exit_query
        End If
        
        cataCount = pdgCatalist(Lib(2, i), Cata())
        Dim j As Long
        stsbar.Value = 0
        stsbar.Min = 0
        If cataCount > 0 Then stsbar.Max = cataCount
        
        lblstatus.Caption = "[" & i & "/" & libCount & "] working at [" & Lib(1, i) & "]"

        For j = 1 To cataCount
            bookCount = pdgBookList(Cata(2, j), Book())
            Dim K As Long
            For K = 1 To bookCount
                    'rs.a
                    rs.AddNew
                    Set fs = rs.Fields
                    fs(0).Value = l_CLong(Book(1, K))
                    fs(1).Value = Book(2, K)
                    fs(2).Value = Book(3, K)
                    fs(3).Value = l_CInt(Book(4, K))
                    fs(4).Value = Book(5, K)
                    fs(5).Value = Cata(1, j)
                    'fs(6).Value = Lib(1, i)
                    'fs(7).Value = Book(6, K)
                    rs.Update
                If Force_Stop Then
                    GoTo Exit_query
                    Exit Sub
                End If
            Next
            DoEvents
            stsbar.Value = stsbar.Value + 1
        Next
        
        Set rs = Nothing
        't = t + 1: If t > 3 Then t = 1
        'stsBar.SimpleText = " " & tt(t) & FormatPercent(i / libCount) & " | " & lib(1, i)

    Next

Exit_query:

    'Set f = Nothing
    Set fs = Nothing
    Set rs = Nothing
    Set dbase = Nothing
    
    stsbar.Value = 0
    lblstatus.Caption = ""
    cmdMakeMDB.Enabled = True
    cmdStop.Enabled = False
    Force_Stop = False
        
End Sub

'Private Sub lvResult_DblClick()
'    On Error Resume Next
'    Dim it As ListItem
'    Set it = lvResult.SelectedItem
'
'    If it.Selected = False Then Exit Sub
'    lvResultItemClick it
'End Sub

'Private Sub lvResultItemClick(ByVal Item As MSComctlLib.ListItem)
'
''Dim infoDialog As New Dialog
'
''Load infoDialog
'
'With infoDialog
'.Icon = MainFrm.Icon
'If .lblInfo(0).Caption = "" Then .lblInfo(0) = "Title:"
'.txtInfo(0) = Item.ListSubItems(1).Text
'If .lblInfo(1).Caption = "" Then .lblInfo(1) = "Author:"
'.txtInfo(1) = Item.ListSubItems(2).Text
'If .lblInfo(2).Caption = "" Then .lblInfo(2) = "Catalog:"
'.txtInfo(2) = Item.ListSubItems(3).Text
'If .lblInfo(3).Caption = "" Then .lblInfo(3) = "Library:"
'.txtInfo(3) = Item.ListSubItems(4).Text
'If .lblInfo(4).Caption = "" Then .lblInfo(4) = "Pages:"
'.txtInfo(4) = Item.ListSubItems(5).Text
'If .lblInfo(5).Caption = "" Then .lblInfo(5) = "SS Number:"
'.txtInfo(5) = Item.ListSubItems(6).Text
'If .lblInfo(6).Caption = "" Then .lblInfo(6) = "Link:"
''.txtInfo(6) = Item.ListSubItems(7).Text
'If .lblInfo(7).Caption = "" Then .lblInfo(7) = "Date:"
'.txtInfo(7) = Item.ListSubItems(7).Text
'
'If .txtInfo(5).Text = "" Then .cmdOpenSS.Enabled = False
'If InStr(.txtInfo(6).Text, ":") = 2 Or .txtInfo(6).Text = "" Then .cmdOpenInIE.Enabled = False
'End With
''
''msg = "Title:" & Item.Text
''msg = msg & vbCrLf & "Author:" & Item.ListSubItems(1).Text
''msg = msg & vbCrLf & "Catalog:" & Item.ListSubItems(2).Text
''msg = msg & vbCrLf & "Library:" & Item.ListSubItems(3).Text
''msg = msg & vbCrLf & "Pages:" & Item.ListSubItems(4).Text
''msg = msg & vbCrLf & "SS Number:" & Item.ListSubItems(5).Text
''msg = msg & vbCrLf & "Link:" & Item.ListSubItems(6).Text
''msg = msg & vbCrLf & "Date:" & Item.ListSubItems(7).Text
'''msg = msg & vbCrLf & "Author:" & Item.ListSubItems(8).Text
'''msg = msg & vbCrLf & "Author:" & Item.ListSubItems(9).Text
'
'
'infoDialog.Show 1
'
''MainFrm.lvResult.SetFocus
'End Sub

Private Sub txtDIRSearch_Change()
    Dim fso As New FileSystemObject
    Dim root As String
    Dim i As Long
    root = fso.BuildPath(txtDIRSearch.Text, "liblist.dat")
    If fso.FileExists(root) = False Then
        lblSSReader.Enabled = False
 '       lblKeyword.Enabled = False
        cmdMakeMDB.Enabled = False
  '      txtKeyword.Enabled = False
'        cboLib.Enabled = False
        Exit Sub
    End If
    
    Set fso = Nothing
    
'    Dim tmp
'    tmp = lblSSReader.ForeColor
'    lblSSReader.ForeColor = lblSSReader.BackColor
'    lblSSReader.BackColor = tmp
    
    cmdMakeMDB.Enabled = True
 '   txtKeyword.Enabled = True
'    cboLib.Enabled = True
    lblSSReader.Enabled = True
 '   lblKeyword.Enabled = True
       
    
    loadDataBase root
    
'    cboLib.Clear
'    cboLib.AddItem "All", 0
'    For i = 1 To libCount
'        cboLib.AddItem Lib(1, i), i
'    Next
'    cboLib.ListIndex = 0
'
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

'Private Function match(ByRef Book() As String, ByRef idx As Long) As Boolean
'
'    match = True
'
'    If MyInstr(Book(1, idx), txtKeyword.Text, , vbTextCompare) = False And _
'       MyInstr(Book(2, idx), txtKeyword.Text, , vbTextCompare) = False Then
'       match = False
'    End If
''
''    If txtKeyword.Text <> "" And InStr(1, Book(1, idx), txtKeyword.Text, vbTextCompare) < 1 Then Exit Function
''    If txtAuthor.Text <> "" And InStr(1, Book(2, idx), txtAuthor.Text, vbTextCompare) < 1 Then Exit Function
''    If txtSS.Text <> "" And InStr(1, Book(5, idx), txtSS.Text, vbTextCompare) < 1 Then Exit Function
''    If txtDate.Text <> "" And InStr(1, Book(6, idx), txtDate.Text, vbTextCompare) < 1 Then Exit Function
''
'
'End Function

Private Sub txtDIRSearch_Click()
    txtDIRSearch_Change
End Sub

'Function StrNum(Num As Long, lenNum As Integer) As String
'
'    StrNum = CStr(Num)
'
'    If Len(StrNum) >= lenNum Then
'        StrNum = Left$(StrNum, lenNum)
'    Else
'        StrNum = String$(lenNum - Len(StrNum), "0") + StrNum
'    End If
'
'End Function




