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
MComboboxHelper.Add_UniqueItem txtKeyword, txtKeyword.Text

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
.Add mnuFile_Exit, SF_CAPTION
.Add mnuHelp, SF_CAPTION
.Add mnuHelp_Help, SF_CAPTION
.Add mnuHelp_About, SF_CAPTION

.Add txtKeyword, SF_LISTTEXT
.Add mnuFile_open, SF_Tag
.Add txtKeyword, SF_TEXT
.Add lblKeyword, SF_CAPTION ', , True
.Add cmdSearch, SF_CAPTION ', , True

.Add chkWildcards, SF_VALUE
.Add chkWildcards, SF_CAPTION

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
loadMDB sMdbfile
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    Dim w As Long
    
    'lblAbout.Left = MainFrm.ScaleWidth - lblAbout.Width - 240
    With frmProcess
        .Width = MainFrm.ScaleWidth - .Left * 2
'        txtMDBFile.Left = lblMDBPath.Left + lblMDBPath.Width + 120
'        txtMDBFile.Width = .Width - txtMDBFile.Left - cmdHelp.Width - cmdOpen.Width - 3 * 120
'
'
'        cmdOpen.Left = txtMDBFile.Left + txtMDBFile.Width + 120
'        cmdHelp.Left = cmdOpen.Left + cmdOpen.Width + 120
'
'
       lblKeyword.Left = 120
        txtKeyword.Left = lblKeyword.Left + lblKeyword.Width + 120
        
'        w = .Width - lblKeyword.Width - lblAuthor.Width - cboLib.Width - 6 * 120
'        txtKeyword.Width = w / 2 '.Width - txtKeyword.Left - cboLib.Width - cmdSearch.Width - cmdStop.Width - 4 * 120
'
'        lblAuthor.Left = txtKeyword.Left + txtKeyword.Width + 120
'
'        txtAuthor.Left = lblAuthor.Left + lblAuthor.Width + 120
'
'        txtAuthor.Width = w / 2
'

        txtKeyword.Width = .Width - lblKeyword.Width - cboLib.Width - 4 * 120
       
        cboLib.Left = txtKeyword.Left + txtKeyword.Width + 120
                
        
        cmdSearch.Left = cboLib.Left + cboLib.Width - cmdSearch.Width
        chkWildcards.Left = cmdSearch.Left - chkWildcards.Width - 120 ' txtKeyword.Left + txtKeyword.Width - chkWildcards.Width
        'cmdStop.Left = cmdSearch.Left + cmdSearch.Width + 120
        
        
    End With
        
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
    
    'If MainFrm.WindowState = 1 Then MainFrm.WindowState = 0
    
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
    strSearch = txtKeyword.Text
    
    If InStr(strSearch, "*") = 0 And _
       InStr(strSearch, "?") = 0 And _
       InStr(strSearch, "#") = 0 And _
       InStr(strSearch, "[") = 0 Then
       
       strSearch = "*" & strSearch & "*"
    End If
       
    
    
    Set db = New DBEngine
    Set dbase = db.openDatabase(mdbFile)
    
    'MDebug.DebugFPrint "Test " & mdbfile
    'MDebug.DebugFPrint "Search for " & strSearch
    
    Dim i As Long
    For i = 1 To libCount
        
        'MDebug.DebugFPrint "At : " & Lib(i)
        
        If i = cboLib.ListIndex Or cboLib.ListIndex = 0 Then
            
            Set rc = dbase.OpenRecordset(Lib(i), dbOpenForwardOnly)
            lblStatus.Caption = "looking in " & Lib(i)
            If Not rc.BOF Then
                
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
                    If fs("title") Like strSearch Or _
                       fs("author") Like strSearch Or _
                       fs("ssid") Like strSearch Or _
                       fs("date") Like strSearch Then
                    'match(fs, strSearch) Then
                        c = c + 1
                        Set LI = lvResult.ListItems.Add(, , StrNum(c, 5))
                       ' MDebug.DebugFPrint "Match At " & rc.AbsolutePosition & "¡¶" & fs("title") & "¡·"
                        LI.ListSubItems.Add , , fs("title")
                        LI.ListSubItems.Add , , fs("author")
                        LI.ListSubItems.Add , , fs("catalog")
                        LI.ListSubItems.Add , , rc.Name
                        LI.ListSubItems.Add , , fs("pages")
                        LI.ListSubItems.Add , , fs("ssid")
                        LI.ListSubItems.Add , , fs("date")
                    End If
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
    Next

Exit_query:
    
    lblStatus.Visible = False
    lvResult.Visible = True
  '  If MainFrm.WindowState = 1 Then MainFrm.WindowState = 0
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
'    strSearch = "title LIKE '" & txtKeyword.Text & "'" & _
'                " OR author LIKE '" & txtKeyword.Text & "'" & _
'                " OR ssid LIKE '" & txtKeyword.Text & "'" & _
'                " OR date LIKE '" & txtKeyword.Text & "'"
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
    Dim root As String
    Dim i As Long
    root = sFile

    
'    Dim tmp
'    tmp = lblMDBPath.ForeColor
'    lblMDBPath.ForeColor = lblMDBPath.BackColor
'    lblMDBPath.BackColor = tmp
    
    loadDataBase root
    
    If libCount < 1 Then
        'lblMDBPath.Enabled = False
        lblKeyword.Enabled = False
        cmdSearch.Enabled = False
        txtKeyword.Enabled = False
        cboLib.Enabled = False
        chkWildcards.Enabled = False
        Exit Sub
    End If
    
    
    cmdSearch.Enabled = True
    txtKeyword.Enabled = True
    cboLib.Enabled = True
    'lblMDBPath.Enabled = True
    lblKeyword.Enabled = True
    chkWildcards.Enabled = True
    

    
    cboLib.Clear
    cboLib.AddItem "All", 0
    For i = 1 To libCount
        cboLib.AddItem Lib(i), i
    Next
    cboLib.ListIndex = 0
    
    'MComboboxHelper.Add_UniqueItem txtMDBFile, txtMDBFile.Text
    
End Sub

Public Sub loadDataBase(ByRef rootTree As String)

    'If rootTree = "" Then rootTree = libList
        
    libCount = 0
    Erase Lib()
    If rootTree = "" Then Exit Sub
    libCount = MDAO.getTabledefs(rootTree, Lib())
    
End Sub

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
''    If MyInstr(Book(1, idx), txtKeyword.Text, , vbTextCompare) = False And _
''       MyInstr(Book(2, idx), txtKeyword.Text, , vbTextCompare) = False Then
''       match = False
''    End If
''
''    If txtKeyword.Text <> "" And InStr(1, Book(1, idx), txtKeyword.Text, vbTextCompare) < 1 Then Exit Function
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
    
    Set dlg = New CCommonDialogLite
    
    If dlg.VBGetOpenFileName(filename:=mdbFile, Filter:="mdb File (*.mdb)|*.mdb") Then
        sMdbfile = mdbFile
        loadMDB sMdbfile
        mnuFile_open.Tag = sMdbfile
        'txtMDBFile.Text = mdbFile
        'Call txtMDBFile_Change
    End If
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
