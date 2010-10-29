VERSION 5.00
Begin VB.Form mainfrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cleaner"
   ClientHeight    =   1488
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   6684
   Icon            =   "mainfrm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1488
   ScaleWidth      =   6684
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkDelete 
      Caption         =   "Recent"
      Height          =   375
      Index           =   3
      Left            =   4044
      TabIndex        =   6
      Tag             =   "Recent"
      Top             =   120
      Width           =   1095
   End
   Begin VB.CheckBox chkDelete 
      Caption         =   "History"
      Height          =   375
      Index           =   2
      Left            =   3096
      TabIndex        =   4
      Tag             =   "Local Settings\History\History.IE5"
      Top             =   120
      Width           =   1095
   End
   Begin VB.Frame frmInfo 
      Caption         =   "LOG"
      Height          =   852
      Left            =   96
      TabIndex        =   3
      Top             =   492
      Width           =   6456
      Begin VB.Label lblLog 
         Height          =   480
         Left            =   168
         TabIndex        =   5
         Top             =   228
         Width           =   6108
      End
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clean"
      Height          =   312
      Left            =   5340
      TabIndex        =   2
      Top             =   132
      Width           =   1215
   End
   Begin VB.CheckBox chkDelete 
      Caption         =   "Temporary Internet Files"
      Height          =   375
      Index           =   1
      Left            =   984
      TabIndex        =   1
      Tag             =   "Local Settings\Temporary Internet Files"
      Top             =   120
      Width           =   2220
   End
   Begin VB.CheckBox chkDelete 
      Caption         =   "Temp"
      Height          =   375
      Index           =   0
      Left            =   156
      TabIndex        =   0
      Tag             =   "Local Settings\Temp"
      Top             =   120
      Width           =   804
   End
End
Attribute VB_Name = "mainfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim f_s_c As Long
Dim f_f_c As Long
Dim fd_s_c As Long
Dim fd_f_c As Long


Private Sub cmdClear_Click()
Dim i As Long
Dim iStart As Long
Dim iEnd As Long
Dim strPath As String
iStart = chkDelete.LBound
iEnd = chkDelete.UBound
Dim fso As New FileSystemObject

f_s_c = 0
f_f_c = 0
fd_s_c = 0
fd_f_c = 0

For i = iStart To iEnd
    If chkDelete(i).Value = 1 Then
        strPath = chkDelete(i).Tag
        strPath = fso.BuildPath(Environ$("USERPROFILE"), strPath)
        DeleteFolder strPath
    End If
Next

AddLog f_s_c & " files deleted, while " & _
           f_f_c & " files remained." & vbCrLf & _
           fd_s_c & " folders deleted,while " & _
           fd_f_c & " folders remain." & vbCrLf
End Sub
Private Sub AddLog(ByRef msg As String)
    Randomize
    If Int(Rnd(1) * 10 + 1) = 2 Then DoEvents
    lblLog.Caption = msg
End Sub
Private Sub DeleteFolder(ByRef strPath As String)

    Dim fso As New FileSystemObject
    Dim fd As Folder
    Dim fds As Folders
    Dim fs As Files
    Dim f As File
    Dim s As String
    
    
    If fso.FolderExists(strPath) = False Then Exit Sub
    
    On Error Resume Next
    
    Set fs = fso.GetFolder(strPath).Files
    
    For Each f In fs
        s = f.Path
        fso.DeleteFile s, True
        If Err.Number = 0 Then
            AddLog "Deleting " & s & "... OK."
            f_s_c = f_s_c + 1
        Else
            AddLog "Deleting " & s & "... " & Err.Description
            f_f_c = f_f_c + 1
        End If
        Err.Clear
    Next
    
    Set fds = fso.GetFolder(strPath).SubFolders
    
    For Each fd In fds
        DeleteFolder fd.Path
        s = fd.Path
        fso.DeleteFolder s, True
        If Err.Number = 0 Then
            AddLog "Deleting " & s & "... OK."
            fd_s_c = fd_s_c + 1
        Else
            AddLog "Deleting " & s & "... " & Err.Description
            fd_f_c = fd_f_c + 1
        End If
        Err.Clear
    Next
    
    
End Sub

Private Sub Form_Load()
    Dim hSet As New CSetting
    Dim i As Long
    Dim iStart As Long
    Dim iEnd As Long
    iStart = chkDelete.LBound
    iEnd = chkDelete.UBound
    With hSet
        .iniFile = bddir(App.Path) & App.EXEName & ".ini"
        For i = iStart To iEnd
            .Load chkDelete(i), SF_VALUE, "chkDelete" & CStr(i)
        Next
    End With
    Set hSet = Nothing
End Sub


Private Sub Form_Unload(Cancel As Integer)
 Dim hSet As New CSetting
    Dim i As Long
    Dim iStart As Long
    Dim iEnd As Long
    iStart = chkDelete.LBound
    iEnd = chkDelete.UBound
    With hSet
        .iniFile = bddir(App.Path) & App.EXEName & ".ini"
        For i = iStart To iEnd
            .Save chkDelete(i), SF_VALUE, "chkDelete" & CStr(i)
        Next
    End With
    Set hSet = Nothing
End Sub

