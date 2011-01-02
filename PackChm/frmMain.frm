VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PackChm"
   ClientHeight    =   5475
   ClientLeft      =   150
   ClientTop       =   855
   ClientWidth     =   8040
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5475
   ScaleWidth      =   8040
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReset 
      Caption         =   "&Reset"
      Height          =   375
      Left            =   4920
      TabIndex        =   11
      Top             =   4920
      Width           =   1395
   End
   Begin VB.ComboBox cboFileList 
      Height          =   315
      Left            =   0
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   5280
      Visible         =   0   'False
      Width           =   7695
   End
   Begin VB.TextBox txtDefaultFile 
      Height          =   315
      Left            =   180
      TabIndex        =   7
      Top             =   3600
      Width           =   7695
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   375
      Left            =   3360
      TabIndex        =   10
      Top             =   4920
      Width           =   1395
   End
   Begin VB.CommandButton cmdCompile 
      Caption         =   "&Compile..."
      Default         =   -1  'True
      Height          =   375
      Left            =   6480
      TabIndex        =   12
      Top             =   4920
      Width           =   1395
   End
   Begin VB.TextBox txtExt 
      Height          =   315
      Left            =   180
      TabIndex        =   8
      Top             =   4380
      Width           =   7695
   End
   Begin VB.CommandButton cmdSelectFile 
      Caption         =   "Select..."
      Height          =   375
      Index           =   1
      Left            =   6480
      TabIndex        =   6
      Top             =   2820
      Width           =   1395
   End
   Begin VB.TextBox txtIndexFile 
      Height          =   315
      Left            =   180
      TabIndex        =   5
      Top             =   2820
      Width           =   6135
   End
   Begin VB.CommandButton cmdSelectFile 
      Caption         =   "Select..."
      Height          =   375
      Index           =   0
      Left            =   6480
      TabIndex        =   4
      Top             =   2040
      Width           =   1395
   End
   Begin VB.TextBox txtContentFile 
      Height          =   315
      Left            =   180
      TabIndex        =   3
      Top             =   2040
      Width           =   6135
   End
   Begin VB.CommandButton cmdSelectFolder 
      Caption         =   "Select..."
      Height          =   375
      Left            =   6480
      TabIndex        =   2
      Top             =   1260
      Width           =   1395
   End
   Begin VB.TextBox txtBaseFolder 
      Height          =   315
      Left            =   180
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   1260
      Width           =   6135
   End
   Begin VB.TextBox txtProject 
      Height          =   315
      Left            =   180
      TabIndex        =   0
      Top             =   480
      Width           =   7695
   End
   Begin VB.Label Label1 
      Caption         =   ":Options marked with * is required"
      ForeColor       =   &H000000FF&
      Height          =   435
      Index           =   6
      Left            =   240
      TabIndex        =   19
      Top             =   4920
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Filename extenstions rejected(Separated by comma):"
      Height          =   255
      Index           =   5
      Left            =   180
      TabIndex        =   18
      Top             =   4080
      Width           =   7635
   End
   Begin VB.Label Label1 
      Caption         =   "Default Filename:*"
      Height          =   255
      Index           =   4
      Left            =   180
      TabIndex        =   17
      Top             =   3300
      Width           =   7635
   End
   Begin VB.Label Label1 
      Caption         =   "IndexFile(*.idx):"
      Height          =   255
      Index           =   3
      Left            =   180
      TabIndex        =   16
      Top             =   2520
      Width           =   7635
   End
   Begin VB.Label Label1 
      Caption         =   "ContentFile(*.hhc):"
      Height          =   255
      Index           =   2
      Left            =   180
      TabIndex        =   15
      Top             =   1740
      Width           =   7635
   End
   Begin VB.Label Label1 
      Caption         =   "BaseFolder:*"
      Height          =   255
      Index           =   1
      Left            =   180
      TabIndex        =   14
      Top             =   960
      Width           =   7635
   End
   Begin VB.Label Label1 
      Caption         =   "ProjectName:*"
      Height          =   255
      Index           =   0
      Left            =   180
      TabIndex        =   13
      Top             =   180
      Width           =   7635
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save &As..."
      End
      Begin VB.Menu mnuLoad 
         Caption         =   "&Load..."
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuPreference 
         Caption         =   "&Preference"
      End
      Begin VB.Menu mnuHHC 
         Caption         =   "HTML Help &Compiler"
      End
      Begin VB.Menu mnuHHW 
         Caption         =   "HTML Help &Workshop"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const regAppName As String = "PackCHM"
Const regSection As String = "LastTime"
Public sHHWPath As String
Public sHHCPath As String

Private mFilename As String

Private Sub cboFileList_LostFocus()
    EditTextBox Nothing, True
End Sub

Public Function CheckTextbox(ByRef txtTarget As TextBox, Optional ByRef sPrefix As String = "This textbox") As Boolean
    CheckTextbox = True
    If (txtTarget.Text = "") Then
        MsgBox sPrefix & " could not left empty", vbCritical
        txtTarget.SetFocus
        CheckTextbox = False
    End If
End Function
Private Sub cmdCompile_Click()
    Dim sFileList() As String
    Dim count As Long
    Dim i As Long
    If Not CheckTextbox(txtProject, Quote("Project")) Then Exit Sub
    If Not CheckTextbox(txtBaseFolder, Quote("Base Folder")) Then Exit Sub
    If Not CheckTextbox(txtDefaultFile, Quote("Default Filename")) Then Exit Sub
    If Not FileExists(sHHCPath) Then
        MsgBox "Full path of " & Quote("HTML Help Compiler") & "is not set!", vbCritical
        mnuHHC_Click
    End If
    If txtExt.Text <> "" Then
        Dim sExts() As String
        
        TestExtension "", txtExt.Text, True
        For i = 0 To cboFileList.ListCount - 1
            If TestExtension(cboFileList.List(i), "") Then
                ReDim Preserve sFileList(0 To count)
                sFileList(count) = cboFileList.List(i)
                count = count + 1
            End If
        Next
    Else
        For i = 0 To cboFileList.ListCount - 1
            ReDim Preserve sFileList(0 To count)
            sFileList(count) = cboFileList.List(i)
            count = count + 1
        Next
    End If
    
    'Debug_DumpArray (sFileList)
    Dim hhpText As String
    hhpText = MCHM.newProject(txtProject.Text, txtContentFile.Text, txtDefaultFile.Text, sFileList)
    'Debug.Print hhpText
    Dim sHHPFile As String
    sHHPFile = BuildPath(txtBaseFolder.Text, txtProject.Text & ".hhp")
    Dim nFileNum As Integer
    nFileNum = FreeFile()
    Open sHHPFile For Output As #nFileNum
    Print #nFileNum, hhpText
    Close #nFileNum
    LoadConsolePrograme Quote(sHHCPath) & " " & Quote(sHHPFile), vbNormalFocus
    
End Sub

Private Function TestExtension(ByRef sTarget As String, Optional ByVal sRule As String = "", Optional fChanged As Boolean = False) As Boolean
    Static sExts() As String
    Static naLen() As Integer
    Static fCalledBefore As Boolean
    Dim i As Integer
    Static nRuleSize As Integer
    TestExtension = True
    
    If (Not fCalledBefore) Or fChanged Then
        fCalledBefore = True
        sExts = Split(sRule, ",")
        nRuleSize = UBound(sExts)
        ReDim naLen(0 To nRuleSize)
        For i = 0 To nRuleSize
            naLen(i) = Len(sExts(i))
        Next
    End If
    
    For i = 0 To nRuleSize
        If (Right$(sTarget, naLen(i))) = sExts(i) Then TestExtension = False
    Next
    


End Function

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdReset_Click()
    Dim ctlAny As Control
    Dim sType As String
    For Each ctlAny In Me.Controls
        sType = TypeName(ctlAny)
        If sType = "TextBox" Then
            ctlAny.Text = ""
        ElseIf sType = "ComboBox" Then
            ctlAny.Clear
        End If
    Next
End Sub

Private Sub cmdSelectFile_Click(Index As Integer)
Dim ret As String
Dim dlg As CCommonDialogLite
Set dlg = New CCommonDialogLite
Select Case Index
Case 0
    ret = txtContentFile.Text
    If (dlg.VBGetOpenFileName(ret, , , , , , "Help Content File(*.hhc)|*.hhc", , txtBaseFolder.Text, "Select Content File(*.hhc)")) Then
        txtContentFile.Text = ret
    End If
Case 1
    ret = txtIndexFile.Text
    If (dlg.VBGetOpenFileName(ret, , , , , , "Help Index File(*.idx)|*.idx", , txtBaseFolder.Text, "Select Index File(*.idx)")) Then
        txtIndexFile.Text = ret
    End If
End Select
Set dlg = Nothing

End Sub

Private Sub cmdSelectFolder_Click()
Dim ret As String
Dim dlg As CFolderBrowser
Set dlg = New CFolderBrowser
dlg.InitDirectory = txtBaseFolder.Text
dlg.Owner = Me.hWnd
ret = dlg.Browse()
If (ret <> "") Then
    txtBaseFolder.Text = ret
End If
End Sub

Private Sub Form_Load()
    txtProject.Text = GetSetting(regAppName, regSection, "Project")
    txtBaseFolder.Text = GetSetting(regAppName, regSection, "BaseFolder")
    txtContentFile.Text = GetSetting(regAppName, regSection, "ContentFile")
    txtIndexFile.Text = GetSetting(regAppName, regSection, "IndexFile")
    txtDefaultFile.Text = GetSetting(regAppName, regSection, "DefaultFile")
    txtExt.Text = GetSetting(regAppName, regSection, "Exts")
    sHHCPath = GetSetting(regAppName, regSection, "hhcPath")
    sHHWPath = GetSetting(regAppName, regSection, "hhwPath")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting regAppName, regSection, "Project", txtProject.Text
    SaveSetting regAppName, regSection, "BaseFolder", txtBaseFolder.Text
    SaveSetting regAppName, regSection, "ContentFile", txtContentFile.Text
    SaveSetting regAppName, regSection, "IndexFile", txtIndexFile.Text
    SaveSetting regAppName, regSection, "DefaultFile", txtDefaultFile.Text
    SaveSetting regAppName, regSection, "Exts", txtExt.Text
    SaveSetting regAppName, regSection, "hhcPath", sHHCPath
    SaveSetting regAppName, regSection, "hhwPath", sHHWPath
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show 1, Me
End Sub

Private Sub mnuHHC_Click()
    LoadConsolePrograme Quote(sHHCPath), vbNormalFocus
End Sub

Private Sub mnuHHW_Click()
    LoadGuiPrograme Quote(sHHWPath), vbNormalFocus
End Sub

Private Sub mnuLoad_Click()
    Dim ret As String
    Dim dlg As CCommonDialogLite
    Set dlg = New CCommonDialogLite
    ret = mFilename
    If (dlg.VBGetOpenFileName(ret, , , , , , "PackCHM Project File(*.pcp)|*.pcp", , , "Save As...", ".pcp")) Then
        mFilename = ret
    End If
    LoadSettingFromFile mFilename
End Sub

Private Sub mnuPreference_Click()
    frmOptions.Show 1, Me
End Sub

Private Sub mnuSave_Click()
    If mFilename = "" Then mnuSaveAs_Click: Exit Sub
    SaveSettingToFile mFilename
End Sub

Private Sub mnuSaveAs_Click()
    Dim ret As String
    Dim dlg As CCommonDialogLite
    Set dlg = New CCommonDialogLite
    ret = mFilename
    If (dlg.VBGetSaveFileName(ret, , , "PackCHM Project File(*.pcp)|*.pcp", , , "Save As...", ".pcp")) Then
        mFilename = ret
    End If
    SaveSettingToFile mFilename
End Sub


Private Sub txtBaseFolder_Change()
    Call BaseFolderSet
End Sub

Public Sub BaseFolderSet()
    Dim pathBase As String
    pathBase = txtBaseFolder.Text
    If FolderExists(pathBase) Then Call RefreshFileList(pathBase)
End Sub

Public Sub RefreshFileList(ByRef pathBase As String)
    Static pathLast As String
    If (pathBase = pathLast) Then Exit Sub
    
    Dim filenames() As String
    If (RGetFolderContent(filenames, pathBase, , vbDirectory, 4)) Then
        cboFileList.Clear
        Dim i As Long
        For i = LBound(filenames) To UBound(filenames)
            cboFileList.AddItem filenames(i)
        Next

    End If
End Sub

Private Sub txtContentFile_GotFocus()
    EditTextBox txtContentFile
End Sub

Private Sub txtDefaultFile_GotFocus()
    EditTextBox txtDefaultFile
End Sub

Private Sub EditTextBox(ByRef txtAny As TextBox, Optional ByVal fDone As Boolean = False)
    Static lastTextBox As TextBox
    If (fDone) Then
        lastTextBox.Text = cboFileList.Text
        cboFileList.Visible = False
        Set lastTextBox = Nothing
        Exit Sub
    End If
    Set lastTextBox = txtAny
    With txtAny
    cboFileList.Move .Left, .Top + .Height - cboFileList.Height, .Width
    cboFileList.Text = .Text
    End With
    cboFileList.Visible = True
    cboFileList.SetFocus
    
End Sub

Private Sub txtIndexFile_GotFocus()
    EditTextBox txtIndexFile
End Sub


Private Sub SaveSettingToFile(ByVal sFilename As String)
On Error GoTo error_SaveSetting
    Dim n As Integer
    n = FreeFile
    Open sFilename For Output As #n
    Print #n, txtProject.Text
    Print #n, txtBaseFolder.Text
    Print #n, txtContentFile.Text
    Print #n, txtIndexFile.Text
    Print #n, txtDefaultFile.Text
    Print #n, txtExt.Text
    Close #n
    Exit Sub
error_SaveSetting:
    MsgBox Err.Description, vbCritical
    On Error Resume Next
    Close #n

End Sub

Private Sub LoadSettingFromFile(ByVal sFilename As String)
On Error GoTo error_LoadSetting
    Dim n As Integer
    n = FreeFile
    Dim sBuffer As String
    Open sFilename For Input As #n
    Line Input #n, sBuffer: txtProject.Text = sBuffer
    Line Input #n, sBuffer: txtBaseFolder.Text = sBuffer
    Line Input #n, sBuffer: txtContentFile.Text = sBuffer
    Line Input #n, sBuffer: txtIndexFile.Text = sBuffer
    Line Input #n, sBuffer: txtDefaultFile.Text = sBuffer
    Line Input #n, sBuffer: txtExt.Text = sBuffer
    Close #n
    Exit Sub
error_LoadSetting:
    MsgBox Err.Description, vbCritical
    On Error Resume Next
    Close #n

End Sub

Private Sub LoadConsolePrograme(ByRef sCommand As String, Optional appStyle As VbAppWinStyle = vbNormalFocus)
    Shell sCommand, appStyle
End Sub

Private Sub LoadGuiPrograme(ByRef sCommand As String, Optional appStyle As VbAppWinStyle = vbNormalFocus)
    Shell sCommand, appStyle
End Sub
