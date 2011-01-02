VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3330
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6645
   LinkTopic       =   "Form1"
   ScaleHeight     =   3330
   ScaleWidth      =   6645
   StartUpPosition =   3  '窗口缺省
   Begin VB.DriveListBox Drive1 
      Height          =   300
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   6255
   End
   Begin VB.TextBox Text1 
      Height          =   372
      Left            =   3360
      TabIndex        =   1
      Top             =   1320
      Width           =   3012
   End
   Begin VB.DirListBox Dir1 
      Height          =   1560
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   3012
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Delete"
      Height          =   372
      Left            =   3360
      TabIndex        =   0
      Top             =   1920
      Width           =   3012
   End
   Begin VB.Label Label1 
      Height          =   735
      Left            =   240
      TabIndex        =   4
      Top             =   2520
      Width           =   6255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim ff() As String '定义一个字符串数组用来保存找到的文件名称
Dim fn As Long '保存找到的文件数目
Dim i As Integer
fn = TreeSearch(Dir1.Path, Text1.Text, ff())
Rem If fn > 0 Then
'    Label1.Caption = "一共有" + Str(fn) + "个文件符合要求"
'    Text2.Text = ff(1)
'    If fn > 1 Then
'    For i = 2 To fn
'    Text2.Text = Text2.Text + Chr(13) + Chr(10) + ff(i)
'    Next
'    End If
'Else
'    Call MsgBox("            没有适合条件的结果             ", vbOKOnly, "提示信息")
'End If
End Sub

Private Function TreeSearch(ByVal sPath As String, ByVal sFileSpec As String, sFiles() As String) As Long
Static fstFiles As Long '文件数目
Dim sDir As String
Dim sSubDirs() As String '存放子目录名称
Dim fstIndex As Long
If Right(sPath, 1) <> "\" Then sPath = sPath + "\"
sDir = Dir(sPath + sFileSpec)
'获得当前目录下文件名和数目
Do While Len(sDir)
fstFiles = fstFiles + 1
ReDim Preserve sFiles(1 To fstFiles)
sFiles(fstFiles) = sPath + sDir

Label1.Caption = "Deleting " + sFiles(fstFiles)
Form1.Refresh
Kill sFiles(fstFiles)
sDir = Dir
Loop
'获得当前目录下的子目录名称
fstIndex = 0
sDir = Dir(sPath + "*.*", 16)
Do While Len(sDir)
If Left(sDir, 1) <> "." Then 'skip.and..
'找出子目录名
If GetAttr(sPath + sDir) = vbDirectory Then
fstIndex = fstIndex + 1
'保存子目录名
ReDim Preserve sSubDirs(1 To fstIndex)
sSubDirs(fstIndex) = sPath + sDir + "\"
End If
End If
sDir = Dir
Loop
For fstIndex = 1 To fstIndex '查找每一个子目录下文件，这里利用了递归
Call TreeSearch(sSubDirs(fstIndex), sFileSpec, sFiles())
Next fstIndex
TreeSearch = fstFiles
End Function
Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Call Command1_Click
End Sub
