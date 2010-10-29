VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Book info"
   ClientHeight    =   2580
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6300
   Icon            =   "pdgZF.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   6300
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5640
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command4 
      Caption         =   "更新INFO"
      Height          =   375
      Left            =   480
      TabIndex        =   9
      Top             =   1800
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Height          =   270
      Left            =   5280
      TabIndex        =   8
      Text            =   "1"
      Top             =   1920
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "打开正文"
      Height          =   375
      Left            =   4080
      TabIndex        =   7
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "打开目录"
      Height          =   375
      Left            =   2880
      TabIndex        =   6
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "打开封面"
      Height          =   375
      Left            =   1680
      TabIndex        =   5
      Top             =   1800
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1080
      TabIndex        =   3
      Top             =   1200
      Width           =   4455
   End
   Begin VB.Label Label1 
      Caption         =   "下载于："
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   4
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "页数："
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   5415
   End
   Begin VB.Label Label1 
      Caption         =   "作者："
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   5415
   End
   Begin VB.Label Label1 
      Caption         =   "书名："
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   5415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
  Dim thispdg As PDG
Dim hSet As New CAutoSetting

Private Sub Command1_Click()
checkpdg thispdg
openpdg "cov001.pdg"
End Sub

Private Sub Command2_Click()
checkpdg thispdg
openpdg "!00001.pdg"
End Sub

Private Sub Command3_Click()
checkpdg thispdg
openpdg String(6 - Len(Text2.Text), "0") + Text2.Text + ".pdg"
End Sub

Private Sub Command4_Click()

checkpdg thispdg
Dim fso As New FileSystemObject



With thispdg.info
If .vailable Then

    Dim ff As File
    Dim ft As TextStream


    Set ff = fso.GetFile(thispdg.infofile)
    Set ft = ff.OpenAsTextStream(ForWriting)
    .name = InputBox("书名：", "PDGinfo", .name)
    .author = InputBox("作者：", "PDGinfo", .author)
    .totalpage = InputBox("页数：", "PDGinfo", .totalpage)
    .download = InputBox("下载位置：", "PDGinfo", .download)
    ft.WriteLine "书名=" + .name
    ft.WriteLine "作者=" + .author
    ft.WriteLine "页数=" + .totalpage
    ft.WriteLine "下载位置=" + .download
    ft.Close
    
    updateZ
End If
End With


End Sub

Private Sub Form_Load()

    Dim tmpcmd As String
    tmpcmd = Command$
    
    If tmpcmd = "" Then
    
        CommonDialog1.ShowOpen
        
        tmpcmd = CommonDialog1.FileName
            
        If tmpcmd = "" Then End
    End If
    
    thispdg = getpdg(tmpcmd)
    If thispdg.iszip = True Then
    End If
    

    
    With thispdg.info
        Label1(1).Caption = Label1(1).Caption + "《" + .name + "》"
        Label1(2).Caption = Label1(2).Caption + " " + .author
        Label1(3).Caption = Label1(3).Caption + " " + .totalpage
        Text1.Text = .download
        Text2.Tag = .totalpage
    End With
    
    
    With hSet
    .fileNameSaveTo = bddir(App.Path) & "config.ini"
    .Add Form1, SF_Tag
    End With

End Sub


Private Sub Form_Unload(Cancel As Integer)
Set hSet = Nothing
End Sub

Private Sub Text2_Change()
Text2.Text = LTrim(Str(Val(Text2.Text)))
If Val(Text2.Text) < 1 Then Text2.Text = "1"
If Val(Text2.Text) > thispdg.info.totalpage Then Text2.Text = "1"
End Sub



Private Sub updateZ()

checkpdg thispdg
    Dim fso As New FileSystemObject
    Dim srcdir As String
    Dim dstdir As String
    Dim srczip As String
    Dim dstzip As String
If thispdg.iszip = False Then
    With thispdg
    srcdir = .infolder
    dstdir = bddir(fso.GetParentFolderName(.infolder))
    If LTrim(RTrim(.info.name)) <> "" Then dstdir = dstdir + "［" + .info.name + "］" + .info.author
    If LTrim(RTrim(.info.totalpage)) <> "" Then dstdir = dstdir + "(" + .info.totalpage + "页)"
    If LCase(dstdir) <> LCase(srcdir) Then
       fso.MoveFolder srcdir, dstdir
       .infolder = dstdir
       .infofile = bddir(.infolder) + "bookinfo.dat"
    End If
    End With

    With thispdg.info
        Label1(1).Caption = "书名：" + "《" + .name + "》"
        Label1(2).Caption = "作者 ：" + " " + .author
        Label1(3).Caption = "页数：" + " " + .totalpage
        Text1.Text = .download
    End With


Exit Sub
Else
    
    Dim xzip As New cZip
    With xzip
    .zipfile = thispdg.zipfile
    .FileToProcess = thispdg.infofile
    End With
    xzip.Zip


With thispdg
If .info.vailable Then

    srczip = .zipfile
    
    dstzip = bddir(.infolder)
    If LTrim(RTrim(.info.name)) <> "" Then dstzip = dstzip + "［" + .info.name + "］" + .info.author
    If LTrim(RTrim(.info.totalpage)) <> "" Then dstzip = dstzip + "(" + .info.totalpage + "页)"
    dstzip = dstzip + ".zip"
    
    srcdir = .infolder
    
    dstdir = bddir(fso.GetParentFolderName(.infolder))
    If LTrim(RTrim(.info.name)) <> "" Then dstdir = dstdir + "［" + .info.name + "］" + .info.author
    If LTrim(RTrim(.info.totalpage)) <> "" Then dstdir = dstdir + "(" + .info.totalpage + "页)"
            
            If LCase(dstzip) <> LCase(srczip) Then
                fso.MoveFile srczip, dstzip
                .zipfile = dstzip
            End If
            If LCase(dstdir) <> LCase(srcdir) Then
                fso.MoveFolder srcdir, dstdir

                .infolder = dstdir
                .zipfile = bddir(.infolder) + fso.GetFileName(.zipfile)
                .infofile = bddir(dstdir) + "bookinfo.dat"
            End If
    MsgBox "Updated!"
End If
End With

    With thispdg.info
        Label1(1).Caption = "书名：" + "《" + .name + "》"
        Label1(2).Caption = "作者 ：" + " " + .author
        Label1(3).Caption = "页数：" + " " + .totalpage
        Text1.Text = .download
    End With

End If
End Sub

Private Sub ZipComment(sComment As String)
Dim fso As New FileSystemObject
sComment = ""

If fso.FileExists(thispdg.infofile) Then

    Dim ffile As File
    Dim ftextstr As TextStream
    Set ffile = fso.GetFile(thispdg.infofile)
    Set ftextstr = ffile.OpenAsTextStream(ForReading)
    sComment = ftextstr.ReadAll
    ftextstr.Close
End If
sComment = sComment + Chr(13) + Chr(10)
sComment = sComment + "Packed Lin, MYPLACE Publishing Department."
sComment = sComment + Chr(13) + Chr(10)
sComment = sComment + Str(Date) + " " + Str(Time)

Dim xzip As New cZip
xzip.zipfile = thispdg.zipfile
xzip.ZipComment (sComment)
Set xzip = Nothing


End Sub


Sub openpdg(firstpage As String)

    Dim fso As New FileSystemObject
    Dim pdgpath As String
    If fso.FolderExists(thispdg.unzipfolder) Then
     fso.DeleteFolder thispdg.unzipfolder, True
    End If

pdgpath = thispdg.infolder

If thispdg.iszip = True Then

Dim xzresult As unzReturnCode

    Dim xzip1 As New cUnzip
    
    xzip1.zipfile = thispdg.zipfile
    xzip1.FileToProcess = "*"
    xzip1.unzipfolder = thispdg.unzipfolder
    xzresult = xzip1.unzip
    If xzresult <> PK_OK Then
        MsgBox "Error occured when unziping."
        Exit Sub
    End If
    pdgpath = thispdg.unzipfolder
    Set xzip1 = Nothing
    
End If

pdgpath = bddir(pdgpath)

If firstpage = "cov001.pdg" And Dir(pdgpath + "cov001.pdg") = "" Then firstpage = "bok001.pdg"
If firstpage = "bok001.pdg" And Dir(pdgpath + "bok001.pdg") = "" Then firstpage = "fow001.pdg"
If firstpage = "fow001.pdg" And Dir(pdgpath + "fow001.pdg") = "" Then firstpage = "!00001.pdg"
If firstpage = "!00001.pdg" And Dir(pdgpath + "!00001.pdg") = "" Then firstpage = "000001.pdg"

pdgpath = pdgpath + firstpage
Clipboard.Clear
Clipboard.SetText fso.GetFileName(thispdg.infolder) + ".pdf"
'Dim ssReader As String
'ssReader = Form1.Tag
'
'If ssReader = "" Or fso.FileExists(ssReader) = False Then ssReader = InputBox("输入SSREADER程序的路径")
'If ssReader = "" Then Exit Sub
'
'Form1.Tag = ssReader
'SaveSetting "PDGZF", "SSREADER", "PATH", ssReader
ShellExecute Me.hWnd, "open", pdgpath, "", "", SW_SHOWMAXIMIZED


Unload Me
    

End Sub

