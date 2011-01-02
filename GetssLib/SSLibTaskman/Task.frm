VERSION 5.00
Begin VB.Form frmTask 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GetSSLib"
   ClientHeight    =   6645
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9450
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6645
   ScaleWidth      =   9450
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer 
      Interval        =   500
      Left            =   4680
      Top             =   5580
   End
   Begin VB.CheckBox chkDetectClipboard 
      Caption         =   "自动检测剪贴板"
      Height          =   345
      Left            =   1695
      TabIndex        =   26
      Tag             =   "NoReseting"
      Top             =   5265
      Width           =   1785
   End
   Begin VB.ComboBox ComSavedIn 
      Height          =   315
      ItemData        =   "Task.frx":0000
      Left            =   105
      List            =   "Task.frx":0002
      TabIndex        =   25
      Top             =   4800
      Width           =   7890
   End
   Begin VB.CommandButton cmdResetAll 
      Caption         =   "重置"
      Height          =   360
      Left            =   5235
      TabIndex        =   24
      Top             =   6105
      Width           =   1125
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "检测剪贴板"
      Height          =   360
      Left            =   1215
      TabIndex        =   23
      Top             =   5820
      Width           =   2895
   End
   Begin VB.TextBox txtAddInfo 
      Height          =   1500
      Left            =   3600
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   7
      Top             =   2385
      Width           =   5685
   End
   Begin VB.TextBox txtPagesCount 
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   2880
      Width           =   3320
   End
   Begin VB.TextBox txtPublishedDate 
      Height          =   315
      Left            =   120
      TabIndex        =   5
      Top             =   3570
      Width           =   3320
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "退出"
      Height          =   360
      Left            =   8205
      TabIndex        =   11
      Top             =   6105
      Width           =   1125
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "添加任务"
      Height          =   360
      Left            =   6735
      TabIndex        =   12
      Top             =   6105
      Width           =   1125
   End
   Begin VB.CheckBox chkStartDownload 
      Caption         =   "立即开始下载"
      Height          =   345
      Left            =   75
      TabIndex        =   10
      Tag             =   "NoReseting"
      Top             =   5265
      Width           =   1530
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "选择..."
      Height          =   360
      Left            =   8145
      TabIndex        =   9
      Top             =   4770
      Width           =   1125
   End
   Begin VB.TextBox txtPublisher 
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   2205
      Width           =   3320
   End
   Begin VB.TextBox txtSsid 
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   3320
   End
   Begin VB.TextBox txtAuthor 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   3320
   End
   Begin VB.TextBox txtTitle 
      Height          =   315
      Left            =   105
      TabIndex        =   0
      Top             =   360
      Width           =   3320
   End
   Begin VB.TextBox txtHeader 
      Height          =   1530
      Left            =   3585
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   6
      Top             =   345
      Width           =   5685
   End
   Begin VB.TextBox txtUrl 
      Height          =   315
      Left            =   105
      TabIndex        =   8
      Top             =   4200
      Width           =   9180
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "备注:"
      Height          =   195
      Left            =   3615
      TabIndex        =   22
      Top             =   1995
      Width           =   420
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "页数:"
      Height          =   195
      Left            =   120
      TabIndex        =   21
      Top             =   2595
      Width           =   420
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "出版日期:"
      Height          =   195
      Left            =   120
      TabIndex        =   20
      Top             =   3315
      Width           =   780
   End
   Begin VB.Label lblSaveIN 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "保存位置:"
      Height          =   195
      Left            =   105
      TabIndex        =   19
      Top             =   4560
      Width           =   780
   End
   Begin VB.Label lblHeader 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "HTTP报头:"
      Height          =   195
      Left            =   3600
      TabIndex        =   18
      Top             =   105
      Width           =   795
   End
   Begin VB.Label lblPublisher 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "出版社:"
      Height          =   195
      Left            =   120
      TabIndex        =   17
      Top             =   1935
      Width           =   600
   End
   Begin VB.Label lblSsid 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SS号:"
      Height          =   195
      Left            =   120
      TabIndex        =   16
      Top             =   1320
      Width           =   420
   End
   Begin VB.Label lblAuthor 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "作者:"
      Height          =   195
      Left            =   120
      TabIndex        =   15
      Top             =   720
      Width           =   420
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "书名:"
      Height          =   195
      Left            =   120
      TabIndex        =   14
      Top             =   135
      Width           =   1425
   End
   Begin VB.Label lblURL 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "URL地址:"
      Height          =   195
      Left            =   105
      TabIndex        =   13
      Top             =   3960
      Width           =   705
   End
End
Attribute VB_Name = "frmTask"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mTask As CTask

Const configFile As String = "taskdef.ini"
Private m_bSolidMode As Boolean
Private m_bEditTaskMode As Boolean

Public Property Get EditTaskMode() As Boolean
    EditTaskMode = m_bEditTaskMode
End Property

Public Property Let EditTaskMode(ByVal bValue As Boolean)
    m_bEditTaskMode = bValue
End Property



Public Property Get SolidMode() As Boolean
    SolidMode = m_bSolidMode
End Property

Public Property Let SolidMode(ByVal bValue As Boolean)
    m_bSolidMode = bValue
End Property

Private Sub chkDetectClipboard_Click()
    Form_Activate
End Sub

Private Sub cmdCancel_Click()
    Timer.Enabled = False
    Me.Hide
End Sub

Private Sub cmdOK_Click()
Dim timerState As Boolean
timerState = Timer.Enabled
Timer.Enabled = False
If mTask Is Nothing Then
    frmMain.CallBack_AddTask txtTitle.text, txtAuthor.text, txtSsid.text, _
    txtPublisher.text, txtHeader.text, txtUrl.text, ComSavedIn.text, chkStartDownload.value, _
    txtPagesCount.text, txtPublishedDate.text, txtAddInfo.text
Else
    With mTask
    .Title = txtTitle.text
    .Author = txtAuthor.text
    .Publisher = txtPublisher.text
    .SSID = txtSsid.text
    .RootURL = txtUrl.text
    .HttpHeader = txtHeader.text
    .SavedIN = ComSavedIn.text
    .AdditionalText = txtAddInfo.text
    .PublishedDate = txtPublishedDate.text
    .PagesCount = txtPagesCount.text
    End With
    frmMain.CallBack_EditTask mTask
    Set mTask = Nothing
End If
Timer.Enabled = timerState
If m_bEditTaskMode Then
    m_bEditTaskMode = False
    Me.Hide
End If

'If Not m_bSolidMode Then Me.Hide
End Sub


Public Sub Init(ByRef vtask As CTask, Optional ByRef Disable As Boolean = False)
Me.Reset
cmdOK.Enabled = Not Disable
cmdResetAll.Enabled = Not Disable
Set mTask = Nothing

If vtask Is Nothing Then Exit Sub
With vtask
    If .Author <> "" Then txtAuthor.text = .Author
    If .Title <> "" Then txtTitle.text = .Title
    If .Publisher <> "" Then txtPublisher.text = .Publisher
    If .SSID <> "" Then txtSsid.text = .SSID
    If .HttpHeader <> "" Then txtHeader.text = .HttpHeader
    If .RootURL <> "" Then txtUrl.text = .RootURL
    If .SavedIN <> "" Then
        Dim i As Long
        Dim c As Long
        c = ComSavedIn.ListCount - 1
        For i = 0 To c
            If ComSavedIn.List(i) = .SavedIN Or ComSavedIn.List(i) = .SavedIN & "\" Then GoTo NoSaved
        Next

        ComSavedIn.AddItem .SavedIN
NoSaved:
        ComSavedIn.text = .SavedIN
    End If
    If .AdditionalText <> "" Then txtAddInfo.text = .AdditionalText
    If .PagesCount <> "" Then txtPagesCount.text = .PagesCount
    If .PublishedDate <> "" Then txtPublishedDate.text = .PublishedDate
End With
Set mTask = vtask



End Sub

Public Sub UpdateFromClipboard(Optional vText As String)
Dim vData As SSLIB_BOOKINFO
vData = SSLIB_ParseInfoText(vText)
With vData
    If .Author <> "" Then txtAuthor.text = .Author
    If .Title <> "" Then txtTitle.text = .Title
    If .Publisher <> "" Then txtPublisher.text = .Publisher
    If .SSID <> "" Then txtSsid.text = .SSID
    If .Header <> "" Then txtHeader.text = .Header
    If .URL <> "" Then txtUrl.text = .URL
    If .AddInfo <> "" Then txtAddInfo.text = .AddInfo
    If .PagesCount <> "" Then txtPagesCount.text = .PagesCount
    If .PublishedDate <> "" Then txtPublishedDate.text = .PublishedDate
End With
End Sub

Private Sub cmdResetAll_Click()
Me.Reset
End Sub

Private Sub cmdSelect_Click()
    Dim dlg As CFolderBrowser
    Set dlg = New CFolderBrowser
    If ComSavedIn.text <> "" Then dlg.InitDirectory = ComSavedIn.text
    dlg.Owner = Me.hwnd
    Dim r As String
    r = dlg.Browse
    If r <> "" Then

        Dim i As Long
        Dim c As Long
        c = ComSavedIn.ListCount - 1
        For i = 0 To c
            If ComSavedIn.List(i) = r Or ComSavedIn.List(i) = r & "\" Then Exit Sub
        Next
        ComSavedIn.AddItem r
        ComSavedIn.text = r
    End If
    
    
End Sub

Public Sub Reset()
    MFORM.ResetForm Me
End Sub

Private Sub cmdUpdate_Click()
    UpdateFromClipboard
End Sub



Private Sub Form_Activate()
    Timer.Enabled = False
    If chkDetectClipboard.value Then
        Timer.Enabled = True
    Else
        Timer.Enabled = False
    End If
    
    If m_bEditTaskMode Then
        cmdOK.Caption = "确认"
    Else
        cmdOK.Caption = "添加任务"
    End If
    
'    If m_bSolidMode Then
'        Me.BorderStyle = 0
'        cmdCancel.Enabled = False
'    Else
'        Me.BorderStyle = 1
'        cmdCancel.Enabled = True
'    End If
End Sub

Private Sub Form_Load()
    ComSavedIn.Tag = CST_FORM_FLAGS_NORESETING
    chkStartDownload.Tag = CST_FORM_FLAGS_NORESETING
    chkDetectClipboard.Tag = CST_FORM_FLAGS_NORESETING
    Timer.Enabled = False
    Dim configHnd As CLiNInI
    Set configHnd = New CLiNInI
    With configHnd
        .Source = App.Path & "\" & configFile
        ComboxItemsFromString ComSavedIn, .GetSetting("SavedIn", "Path")
    End With
    Set configHnd = Nothing
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim configHnd As CLiNInI
    Set configHnd = New CLiNInI
    With configHnd
        .Source = App.Path & "\" & configFile
        .SaveSetting "SavedIn", "Path", ComboxItemsToString(ComSavedIn)
        .Save
    End With
    Set configHnd = Nothing
End Sub

Private Sub Timer_Timer()
    DetectClipboard
End Sub

Private Sub DetectClipboard()
    Dim vText As String
    vText = Clipboard.GetText
    If vText <> "" Then
        'Clipboard.Clear
        UpdateFromClipboard vText
    End If
End Sub
