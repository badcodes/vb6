VERSION 5.00
Begin VB.Form frmBatchTasksAdd 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Batch Tasks"
   ClientHeight    =   2220
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5175
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2220
   ScaleWidth      =   5175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkStartDownload 
      Caption         =   "立即开始下载"
      Height          =   345
      Left            =   180
      TabIndex        =   9
      Tag             =   "NoReseting"
      Top             =   1620
      Width           =   1740
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消"
      Height          =   375
      Left            =   3840
      TabIndex        =   8
      Top             =   1635
      Width           =   1110
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "确定"
      Default         =   -1  'True
      Height          =   375
      Left            =   2415
      TabIndex        =   7
      Top             =   1635
      Width           =   1110
   End
   Begin VB.ComboBox cboSavedIn 
      Height          =   315
      Left            =   135
      TabIndex        =   6
      Top             =   1035
      Width           =   3495
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "选择"
      Height          =   375
      Left            =   3825
      TabIndex        =   5
      Top             =   990
      Width           =   1110
   End
   Begin VB.TextBox txtSSIDLength 
      Height          =   300
      Left            =   3225
      TabIndex        =   2
      Text            =   "0"
      Top             =   165
      Width           =   1530
   End
   Begin VB.TextBox txtSSID 
      Height          =   300
      Left            =   960
      TabIndex        =   0
      Text            =   "0"
      Top             =   150
      Width           =   1530
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "保存目录:"
      Height          =   195
      Left            =   150
      TabIndex        =   4
      Top             =   705
      Width           =   765
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "数量:"
      Height          =   195
      Left            =   2805
      TabIndex        =   3
      Top             =   210
      Width           =   405
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "开始SSID:"
      Height          =   195
      Left            =   150
      TabIndex        =   1
      Top             =   210
      Width           =   780
   End
End
Attribute VB_Name = "frmBatchTasksAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const configFile As String = "taskdef.ini"

Private Sub cmdCancel_Click()
    Me.Hide
End Sub

Private Sub cmdOk_Click()


Dim iStart As Long
Dim iEnd As Long
Dim iStep As Integer
iStart = StringToLong(txtSSID.text)
iEnd = StringToLong(txtSSIDLength.text)
If iEnd = 0 Then Exit Sub

If iEnd < 1 Then
    iEnd = iStart + iEnd + 1
Else
    iEnd = iStart + iEnd - 1
End If

If iStart > iEnd Then iStep = -1 Else iStep = 1

Dim i As Long
Dim vArray() As String
vArray = SSLIB_CreateBookInfoArray()
vArray(SSLIBFields.SSF_ISJPGBOOK) = "1"
vArray(SSLIBFields.SSF_SAVEDIN) = cboSavedIn.text
For i = iStart To iEnd Step iStep
    vArray(SSLIBFields.SSF_SSID) = CStr(i)
    #If afTaskEbd = 1 Then
        frmMain.CallBack_AddTask "", vArray, chkStartDownload.Value
    #Else
        CallBack_AddTask "", vArray, chkStartDownload.Value
    #End If
Next




       ' Dim i As Long
       On Error Resume Next
        Dim c As Long
        Dim text As String
        
        
        text = cboSavedIn.text
    If text <> "" Then
        c = cboSavedIn.ListCount - 1
        For i = 0 To c
            If cboSavedIn.List(i) = text Or cboSavedIn.List(i) & "\" = text Then GoTo NoSaved
        Next
         cboSavedIn.AddItem text
NoSaved:
    End If
    Me.Hide
End Sub

Private Sub cmdSelect_Click()
    Dim vdirectory As String
    vdirectory = cboSavedIn.text
    Dim dlg As CFolderBrowser
    Set dlg = New CFolderBrowser
    If vdirectory <> "" Then dlg.InitDirectory = vdirectory
    dlg.Owner = Me.hWnd
    Dim r As String
    r = dlg.Browse
    If r <> "" Then
        cboSavedIn.text = r
    End If
End Sub

Private Sub Form_Load()
    Dim configHnd As CLiNInI
    Set configHnd = New CLiNInI
    With configHnd
        .Source = App.Path & "\" & configFile
        ComboxItemsFromString cboSavedIn, .GetSetting("SavedIn", "Path")
        'FormStateFromString Me, .GetSetting("Form", "State")
    End With
    Set configHnd = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Exit Sub
    Dim configHnd As CLiNInI
    Set configHnd = New CLiNInI
    With configHnd
        .Source = App.Path & "\" & configFile
        .SaveSetting "SavedIn", "Path", ComboxItemsToString(cboSavedIn)
        .Save
    End With
    Set configHnd = Nothing
End Sub
