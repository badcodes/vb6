VERSION 5.00
Begin VB.Form Mainfrm 
   AutoRedraw      =   -1  'True
   Caption         =   "xUNPack"
   ClientHeight    =   3252
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   6828
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3252
   ScaleWidth      =   6828
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frmDestination 
      Caption         =   "2、目标文件夹："
      Height          =   732
      Left            =   120
      TabIndex        =   7
      Top             =   1320
      Width           =   6540
      Begin VB.CommandButton cmdDstFolder 
         Caption         =   "选择"
         Height          =   300
         Left            =   5400
         TabIndex        =   9
         Top             =   240
         Width           =   972
      End
      Begin VB.TextBox txtDstPath 
         Appearance      =   0  'Flat
         Height          =   288
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   5124
      End
   End
   Begin VB.Frame frmSource 
      Caption         =   "1、源文件夹："
      Height          =   972
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   6540
      Begin VB.CheckBox chkDelete 
         Caption         =   "解压后删除"
         Height          =   300
         Left            =   480
         TabIndex        =   6
         Top             =   600
         Width           =   1812
      End
      Begin VB.TextBox txtSrcPath 
         Appearance      =   0  'Flat
         Height          =   288
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   5124
      End
      Begin VB.CheckBox chkSubdir 
         Caption         =   "包含子文件夹"
         Height          =   300
         Left            =   3000
         TabIndex        =   4
         Top             =   600
         Width           =   1692
      End
      Begin VB.CommandButton cmdSrcFolder 
         Caption         =   "选择"
         Height          =   300
         Left            =   5400
         TabIndex        =   3
         Top             =   240
         Width           =   972
      End
   End
   Begin VB.CommandButton cmdUnpack 
      Caption         =   "解压"
      Default         =   -1  'True
      Height          =   300
      Left            =   5544
      TabIndex        =   0
      Top             =   2244
      Width           =   972
   End
   Begin VB.Label lblStatus 
      Caption         =   "准备就绪。"
      Height          =   300
      Left            =   240
      TabIndex        =   10
      Top             =   2280
      Width           =   5148
   End
   Begin VB.Label txtLog 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   456
      Left            =   120
      TabIndex        =   1
      Top             =   2700
      Width           =   6492
   End
End
Attribute VB_Name = "Mainfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public WithEvents xUnzip As cUnzip
Attribute xUnzip.VB_VarHelpID = -1
Private inifile As String
Const MAX_STATUS_TEXT As Long = 30
Const MAX_LOG_TEXT As Long = 100
Private bStop As Boolean
Private Const c_tmpFolder As String = "$$1#3#4$8"
Private sDstfolder As String

Private Sub cmdDstFolder_Click()
Dim sPath As String
sPath = MDlgOpenDir.openDirDialog(Me.hWnd)
If sPath <> "" Then txtDstPath = sPath
End Sub

'Public bDelete As Boolean
'Public bSubdir As Boolean

Private Sub cmdSrcFolder_Click()
Dim sPath As String
sPath = MDlgOpenDir.openDirDialog(Me.hWnd)
If sPath <> "" Then txtSrcPath = sPath
End Sub

Private Sub cmdUnpack_Click()
    If cmdUnpack.Tag = "stop" Then
        cmdUnpack.Tag = "unpack'"
        cmdUnpack.Caption = "解压"
        bStop = True
    Else
        bStop = False
        cmdUnpack.Tag = "stop"
        cmdUnpack.Caption = "中止"
        unzipFolder txtSrcPath.Text, txtDstPath.Text, chkSubdir.Value, chkDelete.Value
        cmdUnpack.Tag = "unpack'"
        cmdUnpack.Caption = "解压"
        bStop = True
    End If
End Sub

'Private Sub cmdUnpack_Click()
'
'
'Dim fso As New FileSystemObject
'
'Dim thisfolder As Folder
'Dim fsofolders As Folders
'
'Dim workdir As String
'Dim delfile As String
'
'Dim ff As File
'Dim ft As TextStream
'Dim PackDir(512) As String
'Dim PNUMS As Integer
'Dim DoNUM As Integer
'Dim zfname(512) As String
'Dim parentdir As String
'
'
'If fso.FolderExists(txtSrcPath.Text) = False Then
'    AddLog txtSrcPath.Text & " not exists."
'    Exit Sub
'End If
'
'AddLog "Search in " & txtSrcPath.Text
'
'
'workdir = txtSrcPath.Text
'workdir = bddir(workdir)
'foundresult = Dir(workdir + "*.zip")
'Do Until foundresult = ""
'   PNUMS = PNUMS + 1
'   PackDir(PNUMS) = workdir
'   zfname(PNUMS) = workdir + foundresult
'   foundresult = Dir()
'Loop
'
'If chkSubdir.Value = 1 Then
'
'    Set fsofolders = fso.GetFolder(txtSrcPath.Text).SubFolders
'    parentdir = txtSrcPath.Text
'    parentdir = bddir(parentdir)
'    For Each thisfolder In fsofolders
'
'        workdir = thisfolder.Path
'        workdir = bddir(workdir)
'        foundresult = Dir(workdir + "*.zip")
'        Do Until foundresult = ""
'            PNUMS = PNUMS + 1
'            PackDir(PNUMS) = workdir
'            zfname(PNUMS) = workdir + foundresult
'            foundresult = Dir()
'        Loop
'    Next
'
'End If
'
'If PNUMS < 1 Then
'    AddLog "Error:No ZIP file found!"
'    Exit Sub
'End If
'
'
'
'For DoNUM = 1 To PNUMS
'
'    Set xUnzip = New cUnzip
'
'    With xUnzip
'    .FileToProcess = "*"
'    .zipfile = zfname(DoNUM)
'    .unzipFolder = PackDir(DoNUM)
'    End With
'
'    AddLog "unPacking " + fso.GetFileName(zfname(DoNUM))
'    Mainfrm.Caption = "PackPDG : Doing " + Str(DoNUM) + " of" + Str(PNUMS)
'
'    xUnzip.unzip
'
'    Set xUnzip = Nothing
'
'    AddLog "unPacked " + fso.GetFileName(zfname(DoNUM)) + " OK. "
'
'    If chkDelete.Value = 1 Then
'       fso.DeleteFile zfname(DoNUM)
'    AddLog fso.GetFileName(zfname(DoNUM)) + " Deleted."
'    End If
'
'    DoEvents
'
'Next
'AddLog "UnPacking Over!"
'
'
'End Sub

Private Sub Form_Load()
    Dim hSet As CSetting
    
    Set hSet = New CSetting
    
    inifile = bddir(App.Path) & App.EXEName & ".ini"
    With hSet
        .inifile = inifile
        .Load Me, SF_FORM
        .Load txtSrcPath, SF_TEXT
        .Load txtDstPath, SF_TEXT
        .Load chkDelete, SF_VALUE
        .Load Me.chkSubdir, SF_VALUE
    End With
    Set hSet = Nothing
    
    AddLog App.ProductName & " Start at " & Date$ & " " & Time$
    
    Set xUnzip = New cUnzip
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If cmdUnpack.Tag = "stop" Then
        bStop = True
        DoEvents
    End If
    Dim hSet As CSetting
    
    Set hSet = New CSetting
    With hSet
        .inifile = inifile
        .Save Me, SF_FORM
        .Save txtSrcPath, SF_TEXT
        .Save txtDstPath, SF_TEXT
        .Save chkDelete, SF_VALUE
        .Save Me.chkSubdir, SF_VALUE
    End With
    Set hSet = Nothing
    Set xUnzip = Nothing
End Sub



Private Sub xUnzip_Cancel(ByVal sMsg As String, bCancel As Boolean)
    AddLog sMsg
    If bStop Then bCancel = True
End Sub

Private Sub xUnzip_Progress(ByVal lCount As Long, ByVal sMsg As String)
    AddLog sMsg
End Sub

Public Sub AddLog(ByRef sMsg As String)
    txtLog.Caption = " " & Replace(strTail(sMsg, MAX_LOG_TEXT), c_tmpFolder, sDstfolder)
    DoEvents
End Sub

Public Function myUnzip(ByRef zipfile As String, ByVal dstFolder As String) As Boolean



    Dim fso As FileSystemObject
    Dim tmpFolder As String
    Dim sName As String
    Dim fd As Folder
    Dim fds As Folders
    Dim sFName As String
    Dim sSrc As String
    Dim sDST As String
    Dim mR As VbMsgBoxResult
    Set fso = New FileSystemObject
    
    sName = fso.GetBaseName(zipfile)
    sDstfolder = sName
    
    myUnzip = False
    
    If xUnzip Is Nothing Then Unload Me
    
    If fso.FileExists(zipfile) = False Then Exit Function
    xMkdir dstFolder
    If fso.FolderExists(dstFolder) = False Then Exit Function
    
    tmpFolder = fso.BuildPath(dstFolder, c_tmpFolder)
    
    If fso.FolderExists(tmpFolder) Then fso.DeleteFolder tmpFolder, True
    fso.CreateFolder tmpFolder
    
    With xUnzip
        .zipfile = zipfile
        .unzipFolder = tmpFolder
        .AddFileToPreocess "*"
        .OverwriteExisting = True
        .CaseSensitiveFileNames = False
        .UseFolderNames = True
    End With
    Dim xuzR As unzReturnCode
    xuzR = xUnzip.unzip
    
    If xuzR <> PK_OK And xuzR <> PK_COOL Then GoTo myUnzip_Force_stop
    If bStop Then GoTo myUnzip_Force_stop
    

    

    Set fd = fso.GetFolder(tmpFolder)
    
    If fd.SubFolders.Count + fd.Files.Count = 1 Then
        Set fds = fd.SubFolders
        For Each fd In fds
            sSrc = fd.Path
            Exit For
        Next
        Set fd = Nothing
        Set fds = Nothing
'        sFName = fso.GetBaseName(sSrc)
'        If sFName = sName Then
        sDST = fso.BuildPath(dstFolder, sName)
    Else
        sSrc = tmpFolder
        sDST = fso.BuildPath(dstFolder, sName)
    End If
    
    'Set fd = Nothing

    mR = vbYes
    If fso.FolderExists(sDST) Then
        'mR = MsgBox("文件夹" & sDST & "已经存在！" & vbCrLf & "覆盖？", vbYesNo)
        If mR = vbYes Then fso.DeleteFolder sDST
    End If
    
    If mR = vbYes Then
        fso.MoveFolder sSrc, sDST
    End If
    
    If fso.FolderExists(tmpFolder) Then fso.DeleteFolder tmpFolder, True
    Set fso = Nothing
    
    myUnzip = True
    
    Exit Function
    
myUnzip_Force_stop:
    Set fso = New FileSystemObject
    If fso.FolderExists(tmpFolder) Then fso.DeleteFolder tmpFolder, True
    Set fso = Nothing
End Function

Public Function unzipFolder(ByRef srcFolder As String, ByRef dstFolder As String, ByVal bSub As Boolean, ByVal bDel As Boolean) As Boolean
    unzipFolder = False
    
    Dim fso As FileSystemObject
    Dim fd As Folder
    Dim fds As Folders
    Dim fs As Files
    Dim f As File
    
    Set fso = New FileSystemObject
    
    On Error GoTo Error_unzipFolder
    
    xMkdir dstFolder
    If fso.FolderExists(srcFolder) = False Or fso.FolderExists(dstFolder) = False Then Exit Function
    
    Set fs = fso.GetFolder(srcFolder).Files

    
    For Each f In fs
        If bStop Then Exit Function
        If xUnzip Is Nothing Then Unload Me
        Debug.Print f.Path
        If xUnzip.ValidateZipFile(f.Path) Then
            lblStatus.Caption = "正在处理" & strTail(f.Path, MAX_STATUS_TEXT) ' & "..."
            myUnzip f.Path, dstFolder
            If bDel Then fso.DeleteFile f.Path, True
        End If
    Next

    If bSub Then
        Set fds = fso.GetFolder(srcFolder).SubFolders
        For Each fd In fds
            If bStop Then Exit Function
            If xUnzip Is Nothing Then Unload Me
            unzipFolder fd.Path, dstFolder, bSub, bDel
            If bDel Then fso.FolderExists fd.Path
        Next
    End If
    
    unzipFolder = True
    Exit Function
    
Error_unzipFolder:
    MsgBox "执行unzipfolder时发生以下错误:" & vbCrLf & Err.Description, vbCritical
    Resume Next
    
End Function

Public Function strTail(ByVal strSrc As String, ByVal tLength As Long) As String
    Dim iLen As Long
    iLen = Len(strSrc)
    If iLen < tLength Then strTail = strSrc: Exit Function
    If tLength <= 0 Then Exit Function
    
    strTail = Right$(strSrc, tLength)
    strTail = "..." & strTail
    
End Function
