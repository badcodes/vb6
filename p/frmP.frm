VERSION 5.00
Begin VB.Form frmP 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "P"
   ClientHeight    =   2790
   ClientLeft      =   30
   ClientTop       =   300
   ClientWidth     =   5835
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   5835
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.TextBox txtSearchExt 
      Height          =   456
      Left            =   156
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   1788
      Width           =   5508
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "SaveSetting"
      Height          =   312
      Left            =   4176
      TabIndex        =   1
      Top             =   2388
      Width           =   1476
   End
   Begin VB.TextBox txtSearchPath 
      Height          =   948
      Left            =   180
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   396
      Width           =   5508
   End
   Begin VB.Label Label2 
      Caption         =   "文件扩展名（以逗号隔开，如：exe,com,bat ）："
      Height          =   300
      Left            =   168
      TabIndex        =   4
      Top             =   1428
      Width           =   5520
   End
   Begin VB.Label Label1 
      Caption         =   "路径（以逗号隔开，如 d:\,C:\ 先后顺序，将影响速度。）："
      Height          =   300
      Left            =   216
      TabIndex        =   3
      Top             =   108
      Width           =   5436
   End
End
Attribute VB_Name = "frmP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Dim sAppIni As String
Dim sHistoryIni As String
Dim secName As String

Private extList() As String
Public scResultFound As New CStringVentor  ' CStringCollection
Const argSplitFlag = "|"


Private Sub cmdSave_Click()
iniSaveSetting sAppIni, "Search", "Path", txtSearchPath.Text
iniSaveSetting sAppIni, "Search", "Ext", txtSearchExt.Text
MsgBox "Setting Saved."
End Sub

Private Sub Form_Load()


Dim cmdLine As String
Dim args() As String
cmdLine = Command

cmdLine = linvblib.LeftRange(cmdLine, Chr$(34), Chr$(34), vbBinaryCompare, ReturnOriginalStr)
args = Split(cmdLine, argSplitFlag)


If UBound(args) > 0 Then
    cmdLine = args(0)
    sAppIni = args(1)
End If
If sAppIni = "" Then sAppIni = App.EXEName
secName = sAppIni

sAppIni = BuildPath(App.Path, sAppIni & ".ini")
sHistoryIni = BuildPath(App.Path, "p.History")
txtSearchPath.Text = iniGetSetting(sAppIni, "Search", "Path")
txtSearchExt.Text = iniGetSetting(sAppIni, "Search", "Ext")



If cmdLine = "" Or txtSearchPath.Text = "" Then
    Me.Visible = True
Else
    extList = Split(LCase(txtSearchExt.Text), ",")
    searchCommand cmdLine
    Unload Me
End If
End Sub

Private Sub searchCommand(ByVal cmdLine As String)

Dim sPreResult As String

cmdLine = LCase$(cmdLine)
sPreResult = iniGetSetting(sHistoryIni, secName, cmdLine)
If sPreResult <> "" And PathExists(sPreResult) Then
  ShellExecute Me.hWnd, "open", sPreResult, "", GetParentFolderName(sPreResult), 1
  Exit Sub
End If

Dim sSearchPath() As String
Dim sSearchPathCount As Long
Dim lBegin As Long
Const iDepth = 2

sSearchPath = Split(txtSearchPath, ",")
sSearchPathCount = UBound(sSearchPath)
If sSearchPathCount < 0 Then Exit Sub

If txtSearchExt.Text <> "" Then
    For lBegin = 0 To sSearchPathCount
        If searchFileToExec(sSearchPath(lBegin), cmdLine, iDepth) Then Exit Sub
    Next
End If

For lBegin = 0 To sSearchPathCount
    If searchPathToOpen(sSearchPath(lBegin), cmdLine, iDepth) Then Exit Sub
Next

If scResultFound.Count = 0 Then
MsgBox "搜索不到。可能是文件名的原因，请不要带扩展名。", vbOKOnly
Exit Sub
End If

Dim minLen As Long
Dim sLen As Long
Dim resultIndex As Long
minLen = 1024
For lBegin = 1 To scResultFound.Count
sLen = Len(GetFileName(scResultFound.Value(lBegin)))
If sLen < minLen Then minLen = sLen: resultIndex = lBegin
Next
sPreResult = scResultFound.Value(resultIndex)
If sPreResult = "" Then Exit Sub
If PathExists(sPreResult) = False Then Exit Sub
iniSaveSetting sHistoryIni, secName, cmdLine, sPreResult
ShellExecute Me.hWnd, "open", sPreResult, "", GetParentFolderName(sPreResult), 1
End Sub


Public Function searchFileToExec(ByRef sPathToSearch As String, ByRef cmdLine As String, ByVal iDepth As Integer) As Boolean

    Dim fso As New FileSystemObject
    Dim curFolder As Folder
    Dim fsoFolder As Folder
    Dim fsoFolders As Folders
    Dim fsoFiles As Files
    Dim eFile As File

    Dim sPath As String
    Dim sFilename As String
    Dim sExt As String
    
        

    On Error GoTo ExitCurSearch
    Set curFolder = fso.GetFolder(sPathToSearch)
    Set fsoFiles = curFolder.Files
    '搜索当前路径下的文件
    For Each eFile In fsoFiles
        sPath = eFile.Path
        sFilename = LCase$(fso.GetBaseName(sPath))
        sExt = LCase$(fso.GetExtensionName(sPath))
        If chkExt(sPath) Then
            If sFilename = cmdLine Then
                'If MsgBox("运行 " & sPath, vbOKCancel) = vbOK Then
                iniSaveSetting sHistoryIni, secName, cmdLine, eFile.Path
                ShellExecute Me.hWnd, "open", eFile.Path, "", eFile.ParentFolder.Path, 1
                'End If
                searchFileToExec = True
                Exit Function
            ElseIf InStr(sFilename, cmdLine) > 0 Then
                scResultFound.assign eFile.Path
            End If
        End If
    Next
    If iDepth <> 0 Then
        iDepth = iDepth - 1
        Set fsoFolders = curFolder.SubFolders
        For Each fsoFolder In fsoFolders
           If searchFileToExec(fsoFolder.Path, cmdLine, iDepth) Then
            searchFileToExec = True
            Exit Function
            End If
        Next
    End If
    
ExitCurSearch:
    
End Function

Public Function searchPathToOpen(ByRef sPathToSearch As String, ByRef cmdLine As String, ByVal iDepth As Integer) As Boolean

    Dim fso As New FileSystemObject
    Dim curFolder As Folder
    Dim fsoFolders As Folders

    
    Dim eFolder As Folder
    Dim sPath As String
    Dim sFilename As String

    
        

    On Error GoTo ExitCurSearch
    Set curFolder = fso.GetFolder(sPathToSearch)
    '搜索当前路径下的文件
    Set fsoFolders = curFolder.SubFolders
    
    For Each eFolder In fsoFolders
        sPath = eFolder.Path
        sFilename = LCase$(fso.GetBaseName(sPath))
        If sFilename = cmdLine Then
            iniSaveSetting sHistoryIni, secName, cmdLine, eFolder.Path
            ShellExecute Me.hWnd, "open", sPath, "", "", 1
            'End If
            searchPathToOpen = True
            Exit Function
        ElseIf InStr(sFilename, cmdLine) > 0 Then
            scResultFound.assign eFolder.Path
        End If
    Next
    If iDepth <> 0 Then
        iDepth = iDepth - 1
        Set fsoFolders = curFolder.SubFolders
        For Each eFolder In fsoFolders
           If searchPathToOpen(eFolder.Path, cmdLine, iDepth) Then
            searchPathToOpen = True
            Exit Function
            End If
        Next
    End If
    
ExitCurSearch:
    
End Function

Private Function chkExt(sFilename As String) As Boolean

Dim lEnd As Integer
Dim lBegin As Integer
Dim sExt As String
lEnd = UBound(extList)
If lEnd < 0 Then chkExt = True: Exit Function
sExt = LCase(GetExtensionName(sFilename))
For lBegin = 0 To lEnd
If sExt = extList(lBegin) Then chkExt = True: Exit Function
Next
End Function
