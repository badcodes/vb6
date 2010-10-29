VERSION 5.00
Begin VB.Form mainFrm 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2364
   ClientLeft      =   2328
   ClientTop       =   1836
   ClientWidth     =   5292
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2364
   ScaleWidth      =   5292
   Begin VB.Frame frmTextSign 
      Caption         =   "2、标志文本："
      Height          =   720
      Left            =   120
      TabIndex        =   4
      Top             =   972
      Width           =   5028
      Begin VB.TextBox txtTextSign 
         Height          =   288
         Left            =   132
         TabIndex        =   5
         Top             =   264
         Width           =   4716
      End
   End
   Begin VB.CommandButton cmdPartition 
      Caption         =   "分割"
      Height          =   288
      Left            =   4320
      TabIndex        =   3
      Top             =   1908
      Width           =   768
   End
   Begin VB.Frame frmSelectFile 
      Caption         =   "1、选择文件："
      Height          =   720
      Left            =   132
      TabIndex        =   0
      Top             =   144
      Width           =   5028
      Begin VB.CommandButton cmdSelect 
         Caption         =   "打开"
         Height          =   288
         Left            =   4116
         TabIndex        =   2
         Top             =   264
         Width           =   768
      End
      Begin VB.TextBox txtPath 
         Height          =   288
         Left            =   132
         TabIndex        =   1
         Top             =   264
         Width           =   3816
      End
   End
   Begin VB.Label lblStatus 
      Height          =   420
      Left            =   144
      TabIndex        =   6
      Top             =   1824
      Width           =   3972
   End
End
Attribute VB_Name = "mainFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim settingHND As CAutoSetting

Private Sub cmdPartition_Click()
If Not LiNVBLib.FileExists(txtPath.Text) Then
    MsgBox "文件不存在", vbInformation
    Exit Sub
End If
If txtTextSign.Text = "" Then
    MsgBox "标准文本为空", vbInformation
    Exit Sub
End If
splitFileByText txtPath.Text, txtTextSign.Text
MsgBox "分割完毕", vbInformation
End Sub

Private Sub cmdSelect_Click()
    Dim dlgOpen As CCommonDialogLite
    Dim sFilePath As String
    Set dlgOpen = New CCommonDialogLite
    sFilePath = txtPath.Text
    If (dlgOpen.VBGetOpenFileName(sFilePath)) Then
        txtPath.Text = sFilePath
    End If
    Set dlgOpen = Nothing
End Sub

Private Sub Form_Load()
    Set settingHND = New CAutoSetting
    With settingHND
    .Add Me.txtPath, SF_TEXT
    .Add Me.txtTextSign, SF_TEXT
    End With
End Sub
Public Sub splitFileByText(ByRef sFilename As String, ByRef sTextSign As String)
Dim fso As FileSystemObject
Dim srcTS As TextStream
Dim dstTS As TextStream
Dim iCount As Long
Dim sFolder As String
Dim sName As String
Dim sExt As String
Dim sLine As String
Dim sDstFile As String
Dim iLineCount As Long
Dim iFileCount As Integer

Set fso = New FileSystemObject
Debug.Print sFilename
If Not fso.FileExists(sFilename) Then Exit Sub
sName = fso.GetBaseName(sFilename)
sFolder = fso.GetParentFolderName(sFilename)
sFolder = fso.BuildPath(sFolder, sName)
sExt = fso.GetExtensionName(sFilename)
If sExt <> "" Then sExt = "." & sExt
If Not fso.FolderExists(sFolder) Then fso.CreateFolder sFolder


Set srcTS = fso.OpenTextFile(sFilename, ForReading, False)
sName = getLineNotEmpty(srcTS)
If sName = "" Then Exit Sub
iFileCount = 1
sDstFile = fso.BuildPath(sFolder, StrNum(iFileCount, 3) & "." & cleanFilename(sName) & sExt)
Debug.Print sDstFile

Set dstTS = fso.CreateTextFile(sDstFile, True)
dstTS.WriteLine sName

Do Until srcTS.AtEndOfStream
    sLine = srcTS.ReadLine
    If sLine Like sTextSign Then  'And iLineCount > 10 Then
        iFileCount = iFileCount + 1
        'iLineCount = 0
        dstTS.Close
        sName = ""
        If Not MyInstr(sTextSign, "*,?") Then sName = LTrim$(RTrim$(Replace$(sLine, sTextSign, "")))
        If sName = "" Then sName = getLineNotEmpty(srcTS)
        If sName = "" Then Exit Sub
        sDstFile = fso.BuildPath(sFolder, StrNum(iFileCount, 3) & "." & cleanFilename(sName) & sExt)
        Set dstTS = fso.CreateTextFile(sDstFile, True)
        dstTS.WriteLine sName
        Debug.Print sDstFile
    Else
'        If InStr(sLine, sTextSign) Then
'            iLineCount = 0
'        Else
'            iLineCount = iLineCount + 1
'        End If
        dstTS.WriteLine sLine
    End If
Loop

dstTS.Close
srcTS.Close
Set dstTS = Nothing
Set srcTS = Nothing
Set fso = Nothing

If sLine = "" Then sLine = LTrim(RTrim(srcTS.ReadLine))

End Sub
Public Function getLineNotEmpty(ByRef ts As TextStream) As String
    If ts Is Nothing Then Exit Function
    Do Until ts.AtEndOfStream
        getLineNotEmpty = LTrim$(RTrim$(ts.ReadLine))
        If getLineNotEmpty <> "" Then Exit Function
    Loop
End Function

Private Sub Form_Unload(Cancel As Integer)
    Set settingHND = Nothing
End Sub
