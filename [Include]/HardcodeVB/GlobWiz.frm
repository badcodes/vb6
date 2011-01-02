VERSION 5.00
Begin VB.Form FGlobalWizard 
   Caption         =   "VisualBasic Source Wizard"
   ClientHeight    =   9675
   ClientLeft      =   780
   ClientTop       =   2190
   ClientWidth     =   8685
   Icon            =   "GlobWiz.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9675
   ScaleWidth      =   8685
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClear 
      Caption         =   "清空记录"
      Height          =   375
      Left            =   3630
      TabIndex        =   17
      Top             =   3330
      Width           =   2256
   End
   Begin VB.TextBox txtReport 
      Height          =   4710
      Left            =   225
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   16
      Top             =   4680
      Width           =   8205
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "开始"
      Height          =   375
      Left            =   6180
      TabIndex        =   9
      Top             =   3345
      Width           =   2256
   End
   Begin VB.Frame fraOptions 
      Caption         =   "选项"
      Height          =   3075
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   8205
      Begin VB.CommandButton cmdSelect 
         Caption         =   "选择..."
         Height          =   345
         Index           =   1
         Left            =   6630
         TabIndex        =   15
         Top             =   1230
         Width           =   1380
      End
      Begin VB.TextBox txtPath 
         Height          =   315
         Index           =   1
         Left            =   240
         TabIndex        =   14
         Top             =   1245
         Width           =   6210
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "选择..."
         Height          =   345
         Index           =   0
         Left            =   6630
         TabIndex        =   12
         Top             =   540
         Width           =   1380
      End
      Begin VB.TextBox txtPath 
         Height          =   315
         Index           =   0
         Left            =   240
         TabIndex        =   11
         Top             =   555
         Width           =   6210
      End
      Begin VB.CheckBox chkDelegate 
         Caption         =   "托管模式(Delegate Mode)"
         Height          =   324
         Left            =   195
         TabIndex        =   8
         Top             =   2565
         Width           =   7785
      End
      Begin VB.TextBox txtDeclFileName 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1140
         TabIndex        =   6
         Top             =   2130
         Width           =   6030
      End
      Begin VB.CheckBox chkDeclView 
         Caption         =   "View"
         Enabled         =   0   'False
         Height          =   285
         Left            =   7305
         TabIndex        =   5
         Top             =   2130
         Width           =   780
      End
      Begin VB.OptionButton optConvertMode 
         Caption         =   "从Components到Modules"
         Height          =   285
         Index           =   1
         Left            =   4230
         TabIndex        =   4
         Top             =   1755
         Width           =   3660
      End
      Begin VB.OptionButton optConvertMode 
         Caption         =   "从Modules到Components"
         Height          =   285
         Index           =   0
         Left            =   195
         TabIndex        =   3
         Top             =   1755
         Value           =   -1  'True
         Width           =   3660
      End
      Begin VB.Label lblPath 
         AutoSize        =   -1  'True
         Caption         =   "目标文件夹："
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   13
         Top             =   975
         Width           =   1080
      End
      Begin VB.Label lblPath 
         AutoSize        =   -1  'True
         Caption         =   "源文件夹："
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   10
         Top             =   285
         Width           =   900
      End
      Begin VB.Label lbl 
         Caption         =   "声明文件:"
         Height          =   255
         Index           =   6
         Left            =   195
         TabIndex        =   7
         Top             =   2190
         Width           =   915
      End
   End
   Begin VB.Label lblSource 
      Height          =   255
      Left            =   315
      TabIndex        =   1
      Top             =   3810
      Width           =   8100
   End
   Begin VB.Label lblTarget 
      Height          =   255
      Left            =   285
      TabIndex        =   0
      Top             =   4200
      Width           =   8205
   End
End
Attribute VB_Name = "FGlobalWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private sFileSrc As String
Private sFileDst As String
Private sFileDecl As String

Enum EFileType
    eftBoth
    eftStandard
    eftClass
End Enum

Enum EModuleType
    emtStandard
    emtClassPublic
    emtClassGlobal
    emtClassPrivate
    emtInvalid
End Enum

Private emtCur As EModuleType
Private fDeclChanged As Boolean
Public mConfig As String
Private mStop As Boolean

Private Sub cmdClear_Click()
    txtReport.Text = ""
End Sub

Private Sub cmdSelect_Click(Index As Integer)
    Dim ret As String
    ret = vbcore.BrowseForFolder(Me.hWnd, "", , lblPath(Index).Caption, , txtPath(Index).Text)
    If ret <> "" Then txtPath(Index).Text = ret
End Sub

Private Sub cmdStart_Click()
    If cmdStart.Caption = "停止" Then
        mStop = True
        cmdStart.Caption = "开始"
        Exit Sub
    Else
        cmdStart.Caption = "停止"
        mStop = False
    End If
    Dim pSrcFolder As String
    Dim pDstFolder As String
    Dim pPubToPri As Boolean
    pPubToPri = optConvertMode(1).Value
    
    pSrcFolder = NormalizePath(txtPath(0).Text)
    pDstFolder = NormalizePath(txtPath(1).Text)
    
    Dim pSrcFile As String
    Dim pDstFile As String
    Dim pSrcInfo As TVBSourceModuleInfo
    Dim pSrcFileName As String
    Dim pSrcBasename As String
    
    pSrcFileName = Dir$(pSrcFolder & "*.*", vbNormal)
    
    
    Dim filterobj As Object
    Do Until pSrcFileName = ""
        pSrcFile = pSrcFolder & pSrcFileName
        pSrcInfo = MVBSource.VBSource_GetModuleInfo(pSrcFile)
        If pSrcInfo.FileType = evbsftUnknown Then GoTo Continue
        pSrcBasename = Left$(pSrcFileName, Len(pSrcFileName) - 4)
        If pPubToPri Then
            If pSrcInfo.FileType = evbsftStandard Then
                GoTo CopySource
            ElseIf pSrcInfo.ModuleType = evbsmtClassGlobal Then
                Set filterobj = New CGlobModFilter
                pDstFile = pDstFolder & pSrcBasename & ".bas"
            ElseIf pSrcInfo.ModuleType = evbsmtClassPrivate Then
                GoTo CopySource
            ElseIf pSrcInfo.ModuleType = evbsmtClassPublic Then
                Set filterobj = New CPubPrivFilter
                pDstFile = pDstFolder & pSrcFileName
            End If
        Else
            If pSrcInfo.FileType = evbsftStandard Then
                If chkDelegate.Value = 1 Then
                    CopyFile pSrcFile, pDstFolder & pSrcFileName
                    Set filterobj = New CModGlobDelFilter
                    filterobj.Name = "G" & pSrcBasename
                    pDstFile = pDstFolder & "G" & pSrcBasename & ".cls"
                Else
                    Set filterobj = New CModGlobFilter
                    pDstFile = pDstFolder & "G" & pSrcBasename & ".cls"
                End If
            ElseIf pSrcInfo.ModuleType = evbsmtClassGlobal Then
                GoTo CopySource
            ElseIf pSrcInfo.ModuleType = evbsmtClassPrivate Then
                Set filterobj = New CPrivPubFilter
                pDstFile = pDstFolder & pSrcFileName
            ElseIf pSrcInfo.ModuleType = evbsmtClassPublic Then
                GoTo CopySource
            End If
        End If
        TranslateFile pSrcInfo.FileText, pSrcFile, pDstFile, filterobj
        GoTo Continue
CopySource:
    pDstFile = pDstFolder & pSrcFileName
    CopyFile pSrcFile, pDstFile
Continue:
    lblSource.Caption = pSrcFile
    lblTarget.Caption = pDstFile
    If mStop Then Exit Sub
    pSrcFileName = Dir$()
    DoEvents
    Loop

    
    cmdStart.Caption = "开始"
    mStop = False
    
    
    
End Sub
Private Sub TranslateFile(ByVal vSource As String, ByVal vSrc As String, ByVal vDest As String, ByVal vFilter As IFilter)
    vFilter.Source = vSource
    FilterText vFilter
    txtReport.Text = txtReport.Text & vSrc & vbCrLf & vbTab & "转换为" & vDest & vbCrLf
    Dim fNum As Integer
    fNum = FreeFile
    Open vDest For Output As #fNum
    Print #fNum, vFilter.Target
    Close #fNum
End Sub
Private Sub CopyFile(ByVal vSrc As String, ByVal vDst As String)
    txtReport.Text = txtReport.Text & vSrc & vbCrLf & vbTab & "复制到" & vDst & vbCrLf
    FileCopy vSrc, vDst
    DoEvents
End Sub
Private Sub Form_Load()
'    If fileCur.ListCount > 0 Then
'        fileCur.ListIndex = 0
'    End If
    
    mConfig = App.Path & "\" & App.EXEName & ".ini"
    
    txtPath(0).Text = iniGetSetting(mConfig, "Path", "Source")
    txtPath(1).Text = iniGetSetting(mConfig, "Path", "Destination")
    txtDeclFileName.Text = iniGetSetting(mConfig, "Filename", "DeclFile")
    
'    FolderBox.AddItem "目标文件夹", VCT_NORMAL + VCT_DIR, iniGetSetting(mConfig, "Path", "Destination"), True
'
'    SourceChange FolderBox.GetField("源文件夹")
'    txtDirectory.Text = FolderBox.GetField("目标文件夹")
    
    'optConvertMode(0).Value = True
    
End Sub

'Sub GetDecl()
'    sFileDecl = NormalizePath(txtfileCur.Path)
'    sFileDecl = sFileDecl & txtDeclFileName
'
'    If Not ExistFileDir(sFileDecl) Then
'        Dim result As VbMsgBoxResult
'
'        result = MsgBox(sFileDecl & " doesn't exist. Create? ", vbYesNoCancel, "Global Wizard")
'        If result = vbYes Then
'            On Error GoTo CreateError
'            CreateDeclFile
'        Else
'            chkDeclView = vbUnchecked
'            Exit Sub
'        End If
'    End If
'
'    On Error GoTo AccessError
'    ' Get text of file regardless
'    txtDecl = GetFileText(sFileDecl)
'    fDeclChanged = False
'    Exit Sub
'
'CreateError:
'    MsgBox Err.Description & sCrLf & sCrLf & _
'           "Unable to create file " & sFileDecl, vbOKOnly + vbExclamation, "Global Wizard"
'    chkDeclView = vbUnchecked
'    Exit Sub
'
'AccessError:
'    MsgBox Err.Description & sCrLf & sCrLf & _
'           "Unable to open file " & sFileDecl, vbOKOnly + vbExclamation, "Global Wizard"
'    chkDeclView = vbUnchecked
'End Sub

'Private Sub SaveDecl()
'    On Error GoTo SaveError
'
'    fDeclChanged = False
'    SaveFileStr sFileDecl, txtDecl
'    Exit Sub
'
'SaveError:
'    MsgBox Err.Description & sCrLf & sCrLf & _
'           "Unable to save changes to global object declarations.", vbOKOnly + vbExclamation, "Global Wizard"
'End Sub
'
'Private Sub cboType_Click()
'    Select Case cboType.ListIndex
'    Case eftBoth
'        fileCur.Pattern = "*.cls;*.bas"
'    Case eftStandard
'        fileCur.Pattern = "*.bas"
'    Case eftClass
'        fileCur.Pattern = "*.cls"
'    End Select
'    fileCur.Refresh
'    If fileCur.ListCount > 0 Then
'        fileCur.ListIndex = 0
'    Else
'        DisplayNothing
'    End If
'End Sub
'
'Private Sub chkDeclView_Click()
'    On Error GoTo FileError
'
'    ' Save changes to current declarations file
'    If fDeclChanged Then SaveDecl
'    ' Load new declarations file
'    If chkDeclView = vbChecked Then GetDecl
'
'    ' Update the display
'    cmdCreate.Enabled = (chkDeclView = vbUnchecked)
'    chkDelegate.Enabled = (chkDeclView = vbUnchecked)
'    txtDstFileName.Enabled = (chkDeclView = vbUnchecked)
'    txtDstModName.Enabled = (chkDeclView = vbUnchecked)
'    txtSrc.Visible = (chkDeclView = vbUnchecked)
'    txtDst.Visible = (chkDeclView = vbUnchecked)
'    txtDecl.Visible = (chkDeclView = vbChecked)
'    Exit Sub
'
'FileError:
'    MsgBox Err.Description & sCrLf & sCrLf & _
'           "Unable to open file " & sFileDecl, vbOKOnly + vbExclamation, "Global Wizard"
'    chkDeclView = vbUnchecked
'    Resume Next
'End Sub

'Private Sub chkDelegate_Click()
'    If chkDelegate Then
'        txtDeclFileName.Text = "N/A"
'        txtDeclFileName.Enabled = False
'        chkDeclView.Enabled = False
'    Else
'        txtDeclFileName.Text = "Objects.Bas"
'        txtDeclFileName.Enabled = True
'        chkDeclView.Enabled = True
'    End If
'    UpdateTargetFileDisplay
'End Sub

'Private Sub cmdClose_Click()
'    Unload Me
'End Sub
'
'Private Sub cmdCreate_Click()
'    On Error GoTo FileError
'    If ExistFile(sFileDst) Then
'        If CreateBackupFile = vbCancel Then Exit Sub
'    End If
'    SaveFileStr sFileDst, txtDst
'
'    If (emtCur = emtStandard) And (chkDelegate = vbUnchecked) Then
'        sFileDecl = CurDir$ & txtDeclFileName
'        If Not ExistFile(sFileDecl) Then CreateDeclFile
'        UpdateDeclFile
'    End If
'    Exit Sub
'
'FileError:
'    MsgBox Err.Description & sCrLf & sCrLf & _
'           "Unable to create file " & sFileDst, vbOKOnly + vbExclamation, "Global Wizard"
'End Sub
'
'Private Sub SourceChange(ByVal vPath As String)
''    fileCur.Path = vPath
''    If fileCur.ListCount > 0 Then
''        fileCur.ListIndex = 0
''    Else
''        DisplayNothing
''    End If
'End Sub
''
''Private Sub drvCur_Change()
''    dirCur.Path = drvCur.Drive
''End Sub
'
'Private Sub fileCur_Click()
'    sFileSrc = NormalizePath(fileCur.Path)
'    sFileSrc = sFileSrc & fileCur.FileName
'    lblSource = sFileSrc
'
'    Dim sModName As String
'    txtSrc = GetModuleInfo(sFileSrc, sModName)
'    txtSrcModName = sModName
'    txtSrcFileName = GetFileBaseExt(sFileSrc)
'
'    Select Case emtCur
'    Case emtInvalid
'        DisplayInvalid
'    Case emtStandard
'        DisplayStandard
'    Case emtClassPublic
'        DisplayPublic
'    Case emtClassGlobal
'        DisplayGlobal
'    Case emtClassPrivate
'        DisplayPrivate
'    End Select
'    If emtCur <> emtInvalid Then
'        txtDstFileName.Enabled = True
'        cmdCreate.Enabled = True
'        UpdateTargetFileDisplay
'    End If
'End Sub
'
'Function GetModuleInfo(sFileSrc As String, sModName As String) As String
'    Dim s As String, iStart As Long, iEnd As Long, sTmp As String
'    Const sTargetName As String = "Attribute VB_Name = """
'    Const sTargetPublic As String = "VB_Exposed = "
'    Const sTargetGlobal As String = "Attribute VB_GlobalNameSpace = "
'
'    On Error GoTo FailGetModuleInfo
'    ' Get text of file regardless
'    s = GetFileText(sFileSrc)
'    ' Find module name
'    iStart = InStr(s, sTargetName)
'    If iStart = 0 Then GoTo FailGetModuleInfo
'    iStart = iStart + Len(sTargetName)
'    iEnd = InStr(iStart, s, sQuote2)
'    If iEnd = 0 Then GoTo FailGetModuleInfo
'    sModName = Mid$(s, iStart, iEnd - iStart)
'    ' Find module type
'    If UCase$(GetFileExt(sFileSrc)) = ".BAS" Then
'        emtCur = emtStandard
'    Else
'        ' Find public attribute
'        iStart = InStr(s, sTargetPublic)
'        If iStart = 0 Then GoTo FailGetModuleInfo
'        iStart = iStart + Len(sTargetPublic)
'        sTmp = Mid$(s, iStart, 1)
'        Select Case sTmp
'        Case "F"
'            emtCur = emtClassPrivate
'        Case "T"
'            ' Find global attribute
'            iStart = InStr(s, sTargetGlobal)
'            If iStart = 0 Then GoTo FailGetModuleInfo
'            iStart = iStart + Len(sTargetGlobal)
'            sTmp = Mid$(s, iStart, 1)
'            Select Case sTmp
'            Case "F"
'                emtCur = emtClassPublic
'            Case "T"
'                emtCur = emtClassGlobal
'            Case Else
'                GoTo FailGetModuleInfo
'            End Select
'        Case Else
'            GoTo FailGetModuleInfo
'        End Select
'    End If
'    GetModuleInfo = s
'    Exit Function
'
'FailGetModuleInfo:
'    ' Any number of reasons why module might be invalid
'    emtCur = emtInvalid
'    GetModuleInfo = s
'End Function

'Sub UpdateTargetFileDisplay()
'    HourGlass Me
'
'    ' Select the appropriate filter and assign to any old object
'    Dim filterobj As Object
'    Select Case emtCur
'    Case emtStandard
'        If chkDelegate Then
'            ' Translates standard module to global class with delegation
'            Set filterobj = New CModGlobDelFilter
'        Else
'            ' Translates standard module to global class w/o delegation
'            Set filterobj = New CModGlobFilter
'        End If
'    Case emtClassPublic
'        ' Translates public class to private class
'        Set filterobj = New CPubPrivFilter
'    Case emtClassGlobal
'        ' Translates global class to standard module
'        Set filterobj = New CGlobModFilter
'    Case emtClassPrivate
'        ' Translates private class to public class
'        Set filterobj = New CPrivPubFilter
'    Case Else
'        txtDst = ""
'        Exit Sub
'    End Select
'    ' Setting name isn't performance sensitive, so do it late bound
'    filterobj.Name = txtDstModName
'
'    ' Use early-bound variable for performance sensitive filter
'    Dim filter As IFilter
'    Set filter = filterobj
'    filter.Source = txtSrc
'    FilterText filter
'    txtDst = filter.Target
'    HourGlass Me
'End Sub

Private Sub Form_Unload(Cancel As Integer)
    iniSaveSetting mConfig, "Path", "Source", txtPath(0).Text
    iniSaveSetting mConfig, "Path", "Destination", txtPath(1).Text
    iniSaveSetting mConfig, "Filename", "DeclFile", txtDeclFileName.Text
End Sub



'Private Sub txtDecl_Change()
'    fDeclChanged = True
'End Sub
'
'Private Sub txtDeclFileName_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        KeyAscii = 0
'        ' Update the same as if we've lost focus
'        txtDeclFileName_LostFocus
'    End If
'End Sub
'
'Private Sub txtDeclFileName_LostFocus()
'    Dim sExt As String, sPath As String
'
'    If txtDeclFileName = sEmpty Then
'        txtDeclFileName = "Objects.Bas"
'    Else
'        sPath = NormalizePath(fileCur.Path)
'        sExt = GetFileExt(sPath & txtDeclFileName)
'        If sExt = sEmpty Then
'            txtDeclFileName = txtDeclFileName & ".Bas"
'        ElseIf UCase$(sExt) <> ".BAS" Then
'            MsgBox "Invalid filename", vbOKOnly + vbExclamation, "Global Wizard"
'            txtDeclFileName = "Objects.Bas"
'        End If
'    End If
'    If chkDeclView = vbChecked Then chkDeclView_Click
'End Sub
'
'Private Sub txtDirectory_LostFocus()
'    fileCur_Click
'End Sub
'
'Private Sub txtDstFileName_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        KeyAscii = 0
'        ' Update the same as if we've lost focus
'        txtDstFileName_LostFocus
'    End If
'End Sub
'
'Private Sub txtDstFileName_LostFocus()
'    Dim sExt As String, sPath As String
'
'    If txtDstFileName <> sEmpty Then
'        sPath = NormalizePath(fileCur.Path)
'        sExt = GetFileExt(sPath & txtDstFileName)
'        ' If no extension, tack on the correct one
'        If sExt = sEmpty Then
'            sExt = IIf(emtCur = emtClassGlobal, ".Bas", ".Cls")
'            txtDstFileName = txtDstFileName & sExt
'            Exit Sub
'        Else
'            ' Normalize the extension
'            sExt = UCase$(sExt)
'            ' Check for a correct extension
'            If sExt = IIf(emtCur = emtClassGlobal, ".BAS", ".CLS") Then
'                Exit Sub
'            Else
'                MsgBox "Invalid extension", vbOKOnly + vbExclamation, "Global Wizard"
'            End If
'        End If
'    End If
'    ' Target filename invalid. Display the default
'    txtDstFileName = DefaultDstFileName
'End Sub
'
'Private Sub txtDstModName_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        KeyAscii = 0
'        ' Update the same as if we've lost focus
'        txtDstModName_LostFocus
'    End If
'End Sub
'
'Private Sub txtDstModName_LostFocus()
'    If txtDstModName = sEmpty Then txtDstModName = DefaultDstModName
'    UpdateTargetFileDisplay
'End Sub
'
'Private Sub DisplayInvalid()
'    txtSrcModType = "Invalid Module"
'    chkDelegate = vbUnchecked
'    chkDelegate.Enabled = False
'    txtDst = sEmpty
'    txtDstFileName = sEmpty
'    txtDstFileName.Enabled = False
'    txtDstModName = sEmpty
'    txtDstModName.Enabled = False
'    txtDstModType = sEmpty
'    txtDeclFileName.Text = sEmpty
'    txtDeclFileName.Enabled = False
'    chkDeclView = vbUnchecked
'    chkDeclView.Enabled = False
'    cmdCreate.Enabled = False
'End Sub
'
'Private Sub DisplayStandard()
'    txtSrcModType = "Standard Module"
'    chkDelegate = vbUnchecked
'    chkDelegate.Enabled = True
'    txtDstFileName = DefaultDstFileName
'    txtDstFileName.Enabled = True
'    txtDstModName = DefaultDstModName
'    txtDstModName.Enabled = True
'    txtDstModType = "Global Class"
'    txtDeclFileName.Text = "Objects.Bas"
'    txtDeclFileName.Enabled = True
'    chkDeclView = vbUnchecked
'    chkDeclView.Enabled = True
'End Sub
'
'Private Sub DisplayPublic()
'    txtSrcModType = "Public Class"
'    chkDelegate = vbUnchecked
'    chkDelegate.Enabled = False
'    txtDstFileName = DefaultDstFileName
'    txtDstFileName.Enabled = True
'    txtDstModName = DefaultDstModName
'    txtDstModName.Enabled = True
'    txtDstModType = "Private Class"
'    txtDeclFileName.Text = "N/A"
'    txtDeclFileName.Enabled = False
'    chkDeclView = vbUnchecked
'    chkDeclView.Enabled = False
'End Sub
'
'Private Sub DisplayGlobal()
'    txtSrcModType = "Global Class"
'    chkDelegate = vbUnchecked
'    chkDelegate.Enabled = False
'    txtDstFileName = DefaultDstFileName
'    txtDstFileName.Enabled = True
'    txtDstModName = DefaultDstModName
'    txtDstModName.Enabled = True
'    txtDstModType = "Standard Module"
'    txtDeclFileName.Text = "N/A"
'    txtDeclFileName.Enabled = False
'    chkDeclView = vbUnchecked
'    chkDeclView.Enabled = False
'End Sub
'
'Private Sub DisplayPrivate()
'    txtSrcModType = "Private Class"
'    chkDelegate = vbUnchecked
'    chkDelegate.Enabled = False
'    txtDstFileName = DefaultDstFileName
'    txtDstFileName.Enabled = True
'    txtDstModName = DefaultDstModName
'    txtDstModName.Enabled = True
'    txtDstModType = "Public Class"
'    txtDeclFileName.Text = "N/A"
'    txtDeclFileName.Enabled = False
'    chkDeclView = vbUnchecked
'    chkDeclView.Enabled = False
'End Sub
'
'Private Sub DisplayNothing()
'    txtSrc = sEmpty
'    txtSrcFileName = sEmpty
'    txtSrcModName = sEmpty
'    txtSrcModType = sEmpty
'    txtDst = sEmpty
'    chkDelegate = vbUnchecked
'    chkDelegate.Enabled = False
'    txtDstFileName = sEmpty
'    txtDstFileName.Enabled = False
'    txtDstModName = sEmpty
'    txtDstModName.Enabled = False
'    txtDstModType = sEmpty
'    txtDeclFileName = sEmpty
'    txtDeclFileName.Enabled = False
'    chkDeclView = vbUnchecked
'    chkDeclView.Enabled = False
'    cmdCreate.Enabled = False
'End Sub
'
'Private Function DefaultDstFileName() As String
'    DefaultDstFileName = GetFileBase(sFileSrc)
'    Select Case emtCur
'    Case emtStandard
'        DefaultDstFileName = DefaultDstFileName & ".Cls"
'    Case emtClassGlobal
'        DefaultDstFileName = DefaultDstFileName & ".Bas"
'    Case emtClassPublic
'        DefaultDstFileName = "P_" & DefaultDstFileName & ".Cls"
'    Case emtClassPrivate
'        DefaultDstFileName = DefaultDstFileName & ".Cls"
'        If Left$(DefaultDstFileName, 2) = "P_" Then
'            DefaultDstFileName = Mid$(DefaultDstFileName, 3)
'        End If
'    End Select
'
'    Dim sPath As String
'    If txtDirectory = sEmpty Then
'        sPath = fileCur.Path
'    Else
'        sPath = GetFullPath(txtDirectory)
'    End If
'    sPath = NormalizePath(sPath)
'
'    sFileDst = sPath & DefaultDstFileName
'    lblTarget = sFileDst
'
'End Function
'
'Private Function DefaultDstModName() As String
'    Select Case emtCur
'    Case emtStandard
'        If Left$(txtSrcModName, 1) = "M" Then
'            DefaultDstModName = "G" & Right$(txtSrcModName, Len(txtSrcModName) - 1)
'        Else
'            DefaultDstModName = "G" & txtSrcModName
'        End If
'    Case emtClassGlobal
'        If Left$(txtSrcModName, 1) = "G" Then
'            DefaultDstModName = "M" & Right$(txtSrcModName, Len(txtSrcModName) - 1)
'        Else
'            DefaultDstModName = "M" & txtSrcModName
'        End If
'    Case emtClassPublic
'        DefaultDstModName = txtSrcModName
'    Case emtClassPrivate
'        DefaultDstModName = txtSrcModName
'    End Select
'End Function
'
'Private Function CreateBackupFile() As VbMsgBoxResult
'    Dim result As VbMsgBoxResult, sTemp As String
'
'    result = MsgBox(sFileDst & " exists. Make backup? ", _
'                    vbYesNoCancel, "Global Wizard")
'    If result = vbYes Then
'        sTemp = sFileDst
'        Mid$(sFileDst, Len(sFileDst)) = "K"
'        If ExistFile(sFileDst) Then Kill sFileDst
'        Name sTemp As sFileDst
'        sFileDst = sTemp
'    End If
'    CreateBackupFile = result
'End Function
'
'Private Sub CreateDeclFile()
'    Dim sHeader As String
'
'    ' Header for global objects declarations module
'    sHeader = "Attribute VB_Name = " & sQuote2 & "M" & GetFileBase(sFileDecl) & sQuote2 & sCrLf & _
'              "Option Explicit" & sCrLf & sCrLf & _
'              "' Global Wizard-generated declarations. DO NOT EDIT THIS COMMENT!" & sCrLf
'
'    SaveFileStr sFileDecl, sHeader
'End Sub
'
'Private Sub UpdateDeclFile()
'    On Error GoTo FileError
'    Dim sDeclaration As String, sComment As String
'    Dim sSrc As String, sDst As String, sLine As String
'    Dim iCommentStart As Long, iCommentEnd As Long
'    Dim iDeclStart As Long, iDeclEnd As Long
'
'    ' Read in declarations file
'    sSrc = GetFileText(sFileDecl)
'
'    ' Look for Global Wizard comment
'    sComment = "' Global Wizard-generated declarations. DO NOT EDIT THIS COMMENT!"
'    iCommentStart = InStr(1, sSrc, sComment, vbTextCompare)
'    If iCommentStart = 0 Then
'        MsgBox "File " & sFileDecl & _
'               " is not a Global Wizard-generated file. " & _
'               "Unable to update global object declarations.", vbOKOnly + vbExclamation, "Global Wizard"
'        Exit Sub
'    End If
'    iCommentEnd = iCommentStart + Len(sComment) + 1
'
'    ' Look for previous declaration
'    sDeclaration = "Public " & txtSrcModName & " As New " & txtDstModName
'    iDeclStart = InStr(1, sSrc, sDeclaration, vbTextCompare)
'    If iDeclStart = 0 Then
'        ' No previous declaration. Insert in sorted order
'        sDst = Left$(sSrc, iCommentEnd)
'        sLine = GetNextLine(Mid$(sSrc, iCommentEnd + 1))
'        Do While (sLine <> sEmpty) And (UCase$(sDeclaration & sCrLf) > UCase$(sLine))
'            sDst = sDst & sLine
'            sLine = GetNextLine
'        Loop
'
'        If sLine = sEmpty Then
'            ' Reached EOF. Insert new declaration at end
'            sDst = sDst & sDeclaration & sCrLf
'        Else
'            Dim iRemainder As Integer
'            iRemainder = Len(sSrc) - (Len(sDst) + Len(sLine))
'            ' Insert new declaration before current line
'            sDst = sDst & sDeclaration & sCrLf & sLine
'            ' Append the remaining declarations
'            sDst = sDst & Right$(sSrc, iRemainder)
'        End If
'    Else
'        ' Previous declaration. Replace with new one
'        sDst = sSrc
'        Mid$(sDst, iDeclStart, Len(sDeclaration)) = sDeclaration
'    End If
'
'    SaveFileStr sFileDecl, sDst
'    Exit Sub
'
'FileError:
'    MsgBox Err.Description & sCrLf & sCrLf & _
'           "Unable to update global object declarations.", vbOKOnly + vbExclamation, "Global Wizard"
'End Sub

