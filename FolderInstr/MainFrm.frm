VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form MainFrm 
   Caption         =   "Form1"
   ClientHeight    =   6624
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   9024
   LinkTopic       =   "Form1"
   ScaleHeight     =   6624
   ScaleWidth      =   9024
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDelete 
      Caption         =   "删除"
      Height          =   276
      Left            =   1944
      TabIndex        =   13
      Top             =   6120
      Width           =   1335
   End
   Begin VB.CommandButton cmdGetList 
      Caption         =   "Refresh"
      Height          =   288
      Left            =   7836
      TabIndex        =   12
      Top             =   6120
      Width           =   912
   End
   Begin VB.ComboBox txtDIRMove 
      Height          =   288
      Left            =   3576
      TabIndex        =   11
      Top             =   6120
      Width           =   4008
   End
   Begin VB.CommandButton cmdFNSearch 
      Caption         =   "Go"
      Height          =   288
      Left            =   6816
      TabIndex        =   10
      Top             =   1200
      Width           =   1935
   End
   Begin VB.TextBox txtFN 
      Height          =   300
      Left            =   1536
      TabIndex        =   8
      Top             =   1188
      Width           =   5055
   End
   Begin VB.CommandButton cmdMoveto 
      Caption         =   "移动"
      Height          =   276
      Left            =   348
      TabIndex        =   7
      Top             =   6132
      Width           =   1335
   End
   Begin MSComctlLib.ListView lvResult 
      Height          =   4212
      Left            =   348
      TabIndex        =   6
      Top             =   1680
      Width           =   8412
      _ExtentX        =   14838
      _ExtentY        =   7430
      SortKey         =   1
      View            =   3
      LabelEdit       =   1
      SortOrder       =   -1  'True
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      Checkboxes      =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "path"
         Text            =   "Path"
         Object.Width           =   11465
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "times"
         Text            =   "次数"
         Object.Width           =   1764
      EndProperty
   End
   Begin VB.CommandButton cmdCTSearch 
      Caption         =   "Go"
      Height          =   276
      Left            =   6816
      TabIndex        =   5
      Top             =   720
      Width           =   1935
   End
   Begin VB.TextBox txtCT 
      Height          =   288
      Left            =   1536
      TabIndex        =   2
      Top             =   708
      Width           =   5055
   End
   Begin VB.TextBox txtTimesMin 
      Height          =   288
      Left            =   7560
      TabIndex        =   1
      Text            =   "6"
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox txtDIRSearch 
      Height          =   288
      Left            =   360
      TabIndex        =   0
      Text            =   "F:\PERSONAL\LiN\Desktop\Download\"
      Top             =   240
      Width           =   6255
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "文件名包含："
      Height          =   192
      Left            =   336
      TabIndex        =   9
      Top             =   1236
      Width           =   1080
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "次数:"
      Height          =   180
      Left            =   6828
      TabIndex        =   4
      Top             =   288
      Width           =   456
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "内容包含："
      Height          =   192
      Left            =   336
      TabIndex        =   3
      Top             =   744
      Width           =   900
   End
End
Attribute VB_Name = "MainFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const inifile = "filestr.ini"
Dim hset As CAutoSetting

Private Sub cmdctsearch_Click()

lvResult.ListItems.Clear
folderSearch txtDIRSearch.Text, txtCT.Text, Val(txtTimesMin.Text)
End Sub

Public Function folderInstr(strFolder As String, strTxt As String, maxTime As Integer) As Integer
Dim fso As New FileSystemObject
Dim fds As Folders
Dim fd As Folder
Dim fs As Files
Dim f As File
Dim iCount As Long
If Not fso.FolderExists(strFolder) Then Exit Function

Set fds = fso.GetFolder(strFolder).SubFolders
For Each fd In fds
    iCount = iCount + folderInstr(fd.Path, strTxt, maxTime)
Next
Set fds = Nothing
Set fd = Nothing

Set fs = fso.GetFolder(strFolder).Files
For Each f In fs
    iCount = iCount + FileInStrTimes(f.Path, strTxt, maxTime, vbTextCompare)
Next
folderInstr = iCount
End Function



Public Sub folderSearch(thedir As String, thetext As String, MTimes As Integer)
Dim fso As New FileSystemObject
Dim fds As Folders
Dim fd As Folder
Dim a As Integer
Dim li As ListItem
Dim t As Integer
Dim m As Integer
Dim tt(3) As String

tt(2) = "-"
tt(1) = "\"
tt(3) = "/"
lvResult.Sorted = False
Set fds = fso.GetFolder(thedir).SubFolders
For Each fd In fds
m = m + 1
t = t + 1
If t > 3 Then t = 1
MainFrm.Caption = " " + tt(t) + FormatPercent(m / fds.Count) + " | " + fd.Path
a = folderInstr(fd.Path, thetext, MTimes)
If a > 0 Then
Set li = lvResult.ListItems.Add(, fd.Path, fd.Path)
li.ListSubItems.Add , "instr" + Str(a), StrNum(a, 4)
MainFrm.Refresh
End If
DoEvents
Next
lvResult.Sorted = True

End Sub

Public Sub FilenameSearch(thedir As String, theFN As String)
Dim fso As New FileSystemObject
Dim fsofolders As Folders
Dim fsofolder As Folder
Dim fsofiles As Files
Dim fsofile As File

Set fsofolder = fso.GetFolder(thedir)
Set fsofolders = fsofolder.SubFolders
Set fsofiles = fsofolder.Files
Dim a As Integer
Dim li As ListItem
Dim t As Integer
Dim m As Integer
Dim tt(3) As String
tt(2) = "-"
tt(1) = "\"
tt(3) = "/"
lvResult.Sorted = False
For Each fsofile In fsofiles
m = m + 1
t = t + 1
If t > 3 Then t = 1
MainFrm.Caption = " " + tt(t) + FormatPercent(m / fsofiles.Count) + " | " + fsofile.Path
a = MyInstr(fso.GetBaseName(fsofile.Path), theFN)
If a = True Then
a = 1
Set li = lvResult.ListItems.Add(, fsofile.Path, fsofile.Path)
li.ListSubItems.Add , "instr" + Str(a), LTrim(Str(a))
End If
MainFrm.Refresh
DoEvents
Next
lvResult.Sorted = True
'For Each fsofolder In fsofolders
'FilenameSearch fsofolder.Path, theFN
'Next

End Sub

Private Sub cmdDelete_Click()
Dim fso As New FileSystemObject
Dim uc As Integer
Dim i As Integer
uc = lvResult.ListItems.Count

For i = uc To 1 Step -1
If lvResult.ListItems(i).Selected = True Then
On Error Resume Next
fso.DeleteFolder lvResult.ListItems(i)
lvResult.ListItems.Remove i
End If
Next
End Sub

Private Sub cmdFNSearch_Click()
lvResult.ListItems.Clear
FilenameSearch txtDIRSearch.Text, txtFN.Text
End Sub

Private Sub cmdGetList_Click()
Dim sText As String
Dim fso As New FileSystemObject
Dim fd As Folder
Dim fds As Folders
sText = txtDIRMove.Text
If fso.FolderExists(sText) Then
    txtDIRMove.Clear
    Set fds = fso.GetFolder(sText).ParentFolder.SubFolders
    For Each fd In fds
    txtDIRMove.AddItem fd.Path
    Next
End If
txtDIRMove.Text = sText
End Sub

Private Sub cmdmoveTo_Click()
Dim fso As New FileSystemObject
Dim uc As Integer
Dim dstdir As String
Dim i As Long
uc = lvResult.ListItems.Count
dstdir = txtDIRMove
If fso.FolderExists(dstdir) = False Then fso.CreateFolder dstdir
For i = uc To 1 Step -1
If lvResult.ListItems(i).Selected = True Then
On Error Resume Next
fso.DeleteFolder fso.BuildPath(dstdir, fso.GetFileName(lvResult.ListItems(i)))
fso.MoveFolder lvResult.ListItems(i), fso.BuildPath(dstdir, fso.GetFileName(lvResult.ListItems(i)))
lvResult.ListItems.Remove i
End If
Next
End Sub

Private Sub Form_Load()
Set hset = New CAutoSetting
hset.fileNameSaveTo = Environ$("appdata") & "\" & inifile
With hset
.Add txtCT, SF_TEXT
.Add txtFN, SF_TEXT
.Add txtDIRMove, SF_TEXT
.Add txtDIRSearch, SF_TEXT
.Add txtTimesMin, SF_TEXT
End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set hset = Nothing
End Sub

Private Sub lvResult_ItemCheck(ByVal Item As MSComctlLib.ListItem)
ShellExecute Me.hWnd, "open", Item, "", "", 1
End Sub

