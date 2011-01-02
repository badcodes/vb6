VERSION 5.00
Object = "{DB797681-40E0-11D2-9BD5-0060082AE372}#4.5#0"; "XceedZip.dll"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "PackPDG"
   ClientHeight    =   6255
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7410
   LinkTopic       =   "Form1"
   ScaleHeight     =   6255
   ScaleWidth      =   7410
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Text            =   "£Û%author£Ý%title(%pagesP)"
      Top             =   4920
      Width           =   6975
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "packpdg.frx":0000
      Left            =   4080
      List            =   "packpdg.frx":0010
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   960
      Width           =   2775
   End
   Begin VB.CheckBox Check1 
      Caption         =   "  Move the packed files ?"
      Height          =   495
      Left            =   4080
      TabIndex        =   5
      Top             =   1320
      Width           =   3015
   End
   Begin VB.FileListBox File1 
      Height          =   2610
      Left            =   4080
      TabIndex        =   3
      Top             =   1800
      Width           =   3015
   End
   Begin VB.DirListBox Dir1 
      Height          =   3660
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   3375
   End
   Begin VB.DriveListBox Drive1 
      Height          =   300
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   3375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Pack it"
      Height          =   375
      Left            =   4080
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label Label3 
      Caption         =   "Name Format:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Height          =   180
      Index           =   1
      Left            =   240
      TabIndex        =   8
      Top             =   5880
      Width           =   90
   End
   Begin XceedZipLibCtl.XceedZip xzip 
      Left            =   3600
      Top             =   2040
      BasePath        =   ""
      CompressionLevel=   6
      EncryptionPassword=   ""
      RequiredFileAttributes=   0
      ExcludedFileAttributes=   24
      FilesToProcess  =   ""
      FilesToExclude  =   ""
      MinDateToProcess=   2
      MaxDateToProcess=   2958465
      MinSizeToProcess=   0
      MaxSizeToProcess=   0
      SplitSize       =   0
      PreservePaths   =   0   'False
      ProcessSubfolders=   0   'False
      SkipIfExisting  =   0   'False
      SkipIfNotExisting=   0   'False
      SkipIfOlderDate =   0   'False
      SkipIfOlderVersion=   0   'False
      TempFolder      =   ""
      UseTempFile     =   -1  'True
      UnzipToFolder   =   ""
      ZipFilename     =   ""
      SpanMultipleDisks=   2
      ExtraHeaders    =   10
      ZipOpenedFiles  =   0   'False
      BackgroundProcessing=   0   'False
      SfxBinrayModule =   ""
      SfxDefaultPassword=   ""
      SfxDefaultUnzipToFolder=   ""
      SfxExistingFileBehavior=   0
      SfxReadmeFile   =   ""
      SfxExecuteAfter =   ""
      SfxInstallMode  =   0   'False
      SfxProgramGroup =   ""
      SfxProgramGroupItems=   ""
      SfxExtensionsToAssociate=   ""
      SfxIconFilename =   ""
   End
   Begin VB.Label Label2 
      Caption         =   "Select Compression Level:"
      Height          =   495
      Left            =   3960
      TabIndex        =   7
      Top             =   720
      Width           =   2535
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   7440
      Y1              =   5520
      Y2              =   5520
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Height          =   180
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Top             =   5640
      Width           =   90
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Command1_Click()


Dim fso As New FileSystemObject

Dim thisfolder As Folder
Dim fsofolders As Folders

Dim workdir As String
Dim delfile As String
Dim ResultCode As xcdError
Dim ff As File
Dim ft As TextStream
Dim PackDir(512) As String
Dim PNUMS As Integer
Dim DoNUM As Integer
Dim zfname(512) As String
Dim pdginfo As pdginfo
Dim parentdir As String



Set fsofolders = fso.GetFolder(Dir1.Path).SubFolders
parentdir = Dir1.Path
parentdir = bddir(parentdir)
For Each thisfolder In fsofolders
    workdir = thisfolder.Path
    workdir = bddir(workdir)
    If Dir(workdir + "*.pdg", vbNormal) <> "" Then
    
       PNUMS = PNUMS + 1
       PackDir(PNUMS) = workdir
       pdginfo = GETpdginfo(workdir + "bookinfo.dat")
       If pdginfo.vailable Then
           'zfname(PNUMS) = workdir + "£Û" + pdginfo.name + "£Ý" + pdginfo.author + "(" + pdginfo.totalpage + "Ò³)" + ".zip"
           zfname(PNUMS) = workdir + pdgformat(pdginfo, Text1.Text) + ".zip"
       Else
           zfname(PNUMS) = workdir + thisfolder.name + ".zip"
       End If
    End If
  
    
Next

If PNUMS < 1 Then
    MsgBox "Error:No pdg folder to pack!"
    Exit Sub
End If
Select Case Combo1.ListIndex
Case 0
    xzip.CompressionLevel = xclNone
Case 1
    xzip.CompressionLevel = xclLow
Case 2
    xzip.CompressionLevel = xclMedium
Case 3
    xzip.CompressionLevel = xclHigh
Case Else
    xzip.CompressionLevel = xclNone
End Select
    xzip.License "SFX45-GS44C-STWAN-YAAA"
For DoNUM = 1 To PNUMS

Command1.ToolTipText = PackDir(DoNUM) + "bookinfo.dat"
xzip.FilesToProcess = PackDir(DoNUM) + "*.*"
xzip.ZipFilename = zfname(DoNUM)
Dim showpath As String
showpath = fso.GetDriveName(PackDir(DoNUM)) + "\...\" + fso.GetBaseName(PackDir(DoNUM))
showpath = bddir(showpath)
Label1(0).Caption = "Packing " + showpath + fso.GetFileName(xzip.FilesToProcess)
Form1.Caption = "PackPDG : Doing " + Str(DoNUM) + " of" + Str(PNUMS)

xzip.Zip
Label1(0).Caption = "Save " + showpath + fso.GetFileName(xzip.ZipFilename) + " OK. "


delfile = Dir(PackDir(DoNUM) + "*.pdg")
Do While delfile <> ""
Label1(1).Caption = "Deleted " + showpath + delfile
fso.DeleteFile PackDir(DoNUM) + delfile, True
delfile = Dir()
Loop
Label1(1).Caption = showpath + "*.pdg Deleted"
srcdir = PackDir(DoNUM)
If Right(srcdir, 1) = "\" Then srcdir = Left(srcdir, Len(srcdir) - 1)
dstdir = parentdir + fso.GetBaseName(zfname(DoNUM))
If LCase(dstdir) <> LCase(srcdir) Then
    fso.MoveFolder srcdir, dstdir
    Label1(0).Caption = "Move " + fso.GetBaseName(srcdir) + " to " + fso.GetBaseName(dstdir)

End If
If Check1.Value = 1 Then
    If Not fso.FolderExists(parentdir + "£Û£Ûdone£Ý£Ý") Then fso.CreateFolder (parentdir + "£Û£Ûdone£Ý£Ý")
    fso.MoveFolder dstdir, parentdir + "£Û£Ûdone£Ý£Ý\" + fso.GetBaseName(dstdir)
End If
Next
MsgBox "Packing Over!"


End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path

End Sub

Private Sub Drive1_Change()

Dir1.Path = Drive1.Drive
End Sub

Private Sub xzip_ZipComment(sComment As String)
Dim fso As New FileSystemObject
sComment = ""

If fso.FileExists(Command1.ToolTipText) Then

    Dim ffile As File
    Dim ftextstr As TextStream
    Set ffile = fso.GetFile(Command1.ToolTipText)
    Set ftextstr = ffile.OpenAsTextStream(ForReading)
    sComment = ftextstr.ReadAll
    Command1.ToolTipText = ""
End If
sComment = sComment + Chr(13) + Chr(10)
sComment = sComment + "Packed Lin, MYPLACE Publishing Department."
sComment = sComment + Chr(13) + Chr(10)
sComment = sComment + Str(Date) + " " + Str(Time)
End Sub



