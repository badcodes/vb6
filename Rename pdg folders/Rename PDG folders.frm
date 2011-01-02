VERSION 5.00
Begin VB.Form Mainfrm 
   ClientHeight    =   2745
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5730
   Icon            =   "Rename PDG folders.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   5730
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdProcess 
      Caption         =   "执行"
      Height          =   288
      Left            =   4356
      TabIndex        =   7
      Top             =   2172
      Width           =   1188
   End
   Begin VB.Frame frmOption 
      Caption         =   "2、设置命名规则："
      Height          =   960
      Left            =   144
      TabIndex        =   4
      Top             =   996
      Width           =   5424
      Begin VB.TextBox txtFormat 
         Height          =   288
         Left            =   168
         TabIndex        =   5
         Top             =   240
         Width           =   5136
      End
      Begin VB.Label lblFormat 
         Height          =   276
         Left            =   168
         TabIndex        =   6
         Top             =   636
         Width           =   5112
      End
   End
   Begin VB.Frame frmSrc 
      Caption         =   "1、选择pdg文件夹的上层目录："
      Height          =   684
      Left            =   132
      TabIndex        =   1
      Top             =   120
      Width           =   5424
      Begin VB.CommandButton cmdsrcPath 
         Caption         =   "选择"
         Height          =   288
         Left            =   4260
         TabIndex        =   3
         Top             =   240
         Width           =   996
      End
      Begin VB.TextBox txtFolder 
         Height          =   288
         Left            =   168
         TabIndex        =   2
         Top             =   240
         Width           =   3924
      End
   End
   Begin VB.Label lblLog 
      ForeColor       =   &H00800000&
      Height          =   432
      Left            =   204
      TabIndex        =   0
      Top             =   2124
      UseMnemonic     =   0   'False
      Width           =   3936
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "Mainfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdProcess_Click()

'Dim fso As FileSystemObject

'Dim thisfolder As Folder
'Dim fsofolders As Folders
Dim workdir As String
Dim parentdir As String
Dim srcdir As String
Dim dstdir As String
Dim srczip As String
Dim dstzip As String
Dim i As Integer
Dim ispdg As Boolean
Dim iszip As Boolean
Dim AllFolder() As String
Dim bookInfo As CBookInfo
Dim j As Long
Dim K As Long
Dim namef As String

On Error Resume Next


parentdir = txtFolder.Text
'Set fso = New FileSystemObject
If FolderExists(parentdir) = False Then
    logIt "错误路径 : " & parentdir
    Exit Sub
End If
Set bookInfo = New CBookInfo
bookInfo(SSF_AUTHOR) = "test"
bookInfo(SSF_Title) = "test"
bookInfo(SSF_SSID) = "test"

namef = pdgformat(bookInfo, txtFormat.Text)
If namef = "" Then
    logIt "错误的命名规则，必须至少含有%t,%a,%s,%p,%c,%d其中一项"
    Exit Sub
End If
Set bookInfo = Nothing


parentdir = bddir(parentdir)

Dim foldersCount As Long
foldersCount = subFolders(parentdir, AllFolder)


For j = 0 To foldersCount - 1
    
    workdir = AllFolder(j)
    logIt "正在处理：" & workdir & " ..."
    DoEvents
    workdir = bddir(workdir)
    If FileExists(workdir + "bookinfo.dat") Then
        
        Set bookInfo = New CBookInfo
    'ispdg = True
     'If fso.FileExists(workdir + "000001.pdg") Then ispdg = True Else ispdg = False
     'If Dir(workdir + "*.zip") = "" Then iszip = False Else iszip = True
    ' If ispdg Or iszip Then
        
        bookInfo.LoadFromFile workdir + "bookinfo.dat"


        namef = cleanFilename(pdgformat(bookInfo, txtFormat.Text))
        
        If namef <> "" Then
'            If iszip Then
'                srczip = Dir(workdir + "*.zip")
'                srczip = workdir + srczip
'                Dir (Environ("temp"))
'                dstzip = workdir + namef + ".zip"
'                If LCase(dstzip) <> LCase(srcdir) Then
'                    fso.MoveFile srczip, dstzip
'                End If
'            End If
            
                srcdir = workdir
                If Right(srcdir, 1) = "\" Then srcdir = Left(srcdir, Len(srcdir) - 1)
                dstdir = parentdir + namef
                
                srcdir = Replace$(srcdir, "/", "\")
                dstdir = Replace$(dstdir, "/", "\")
    
                If LCase$(dstdir) <> LCase$(srcdir) Then
                    If FolderExists(dstdir) Then
                        Dim mR As VbMsgBoxResult
                        mR = MsgBox("文件夹" & Chr$(34) & dstdir & Chr$(34) & "已经存在，覆盖？", vbYesNo)
                        If mR = vbYes Then
                            DeleteFolder dstdir
                        Else
                            GoTo For_Continue
                        End If
                    End If
                    Err.Clear
                    
                    If MoveFile(srcdir, dstdir) Then K = K + 1
                    
                    'fso.MoveFolder srcdir, dstdir
                End If
            End If
     End If
     
    

For_Continue:
Next

    logIt "完成。" & vbCrLf & "重新命名了" & Str(K) + "个文件夹。"
    
End Sub

Private Sub logIt(ByRef logText As String)
    lblLog.Caption = logText
End Sub

Private Sub cmdsrcPath_Click()
    Dim sPath As String
    Dim CDirSel As CFolderBrowser
    
    Set CDirSel = New CFolderBrowser
    CDirSel.InitDirectory = txtFolder.Text
    CDirSel.Owner = Me.hwnd
    CDirSel.Title = "将在选择的文件夹下对所有子文件夹进行批量命名."
    sPath = CDirSel.Browse()
    Set CDirSel = Nothing
    
    If sPath <> "" Then txtFolder.Text = sPath
End Sub

Private Sub Form_Load()
    Mainfrm.Caption = App.ProductName & " V" & App.Major & "." & App.Minor & App.Revision
    lblFormat.Caption = "%t = 书名, %a = 作者, %p = 页数, %d = 日期, %s = ss号, %c = 出版社"
    txtFolder.Text = GetSetting(App.EXEName, "ini", "folder")
    txtFormat.Text = GetSetting(App.EXEName, "ini", "format", "%a - %t")
End Sub



Private Sub Form_Unload(Cancel As Integer)
    SaveSetting App.EXEName, "ini", "folder", txtFolder.Text
    SaveSetting App.EXEName, "ini", "format", txtFormat.Text
End Sub


Private Sub lblFolder_Click()

End Sub


