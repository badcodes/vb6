VERSION 5.00
Begin VB.Form frmTask 
   Caption         =   "SSLib Taskman"
   ClientHeight    =   5535
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   8310
   Icon            =   "Task.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5535
   ScaleWidth      =   8310
   Begin TaskMan.KeyValueEditor BookEditor 
      Height          =   4320
      Left            =   75
      TabIndex        =   7
      Top             =   135
      Width           =   8190
      _ExtentX        =   14446
      _ExtentY        =   7620
      Appearance      =   0
      BorderStyle     =   1
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TwoColumnMode   =   0   'False
   End
   Begin VB.Frame fraAddition 
      BorderStyle     =   0  'None
      Height          =   1140
      Left            =   345
      TabIndex        =   0
      Top             =   4260
      Width           =   7830
      Begin VB.CheckBox chkStartDownload 
         Caption         =   "立即开始下载"
         Height          =   345
         Left            =   15
         TabIndex        =   6
         Tag             =   "NoReseting"
         Top             =   105
         Width           =   1740
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "检测剪贴板"
         Height          =   360
         Left            =   375
         TabIndex        =   5
         Top             =   570
         Width           =   2895
      End
      Begin VB.CheckBox chkDetectClipboard 
         Caption         =   "自动检测剪贴板"
         Height          =   345
         Left            =   1980
         TabIndex        =   4
         Tag             =   "NoReseting"
         Top             =   105
         Width           =   1785
      End
      Begin VB.CommandButton cmdReset 
         Caption         =   "重置"
         Height          =   345
         Left            =   3975
         TabIndex        =   3
         Top             =   585
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "取消"
         Height          =   345
         Left            =   6735
         TabIndex        =   2
         Top             =   585
         Width           =   1095
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "确认"
         Height          =   345
         Left            =   5385
         TabIndex        =   1
         Top             =   585
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmTask"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private mAutoUnload As Boolean
Private mExitStatus As VbMsgBoxResult


Private mTask As CTask
Const configFile As String = "taskdef.ini"
Private m_bSolidMode As Boolean
Private m_bEditTaskMode As Boolean

Private WithEvents Timer As CTimer
Attribute Timer.VB_VarHelpID = -1
Private Const cst_Timer_Interval As Long = 400
Private mSavedInIndex As Long

Private Function N_(ByVal vField As SSLIBFields) As String
    N_ = SSLIB_ChnFieldName(vField)
End Function


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

Private Sub BookEditor_SelectDirectory(ByVal vKeyName As String, vDirectory As String)
    Dim dlg As CFolderBrowser
    Set dlg = New CFolderBrowser
    If vDirectory <> "" Then dlg.InitDirectory = vDirectory
    dlg.Owner = Me.hwnd
    Dim r As String
    r = dlg.Browse
    If r <> "" Then
        vDirectory = r
    End If
End Sub


Private Sub chkDetectClipboard_Click()
    If chkDetectClipboard.value Then
        Timer.Interval = cst_Timer_Interval
    Else
        Timer.Interval = 0
    End If
End Sub


Public Sub InitWithBookInfo(ByRef vBookInfo As CBookInfo)
    Me.Reset
    Set mTask = Nothing
    If vBookInfo Is Nothing Then Exit Sub
    Dim i As Long
    For i = CST_SSLIB_FIELDS_LBound To CST_SSLIB_FIELDS_UBound
    'Debug.Print N_(i) & " = " & bookInfo(i)
      If vBookInfo(i) <> "" Then BookEditor.SetField N_(i), vBookInfo(i)
    Next
    
End Sub


Public Sub Init(ByRef vtask As CTask, Optional ByRef Disable As Boolean = False)

Me.Reset

cmdOk.Enabled = Not Disable
cmdReset.Enabled = Not Disable
Set mTask = Nothing

If vtask Is Nothing Then Exit Sub
BookEditor.SetField "任务名", vtask.Name
Dim bookInfo As CBookInfo
Set bookInfo = vtask.bookInfo

Dim i As Long
For i = CST_SSLIB_FIELDS_LBound To CST_SSLIB_FIELDS_UBound
    'Debug.Print N_(i) & " = " & bookInfo(i)
      If bookInfo(i) <> "" Then BookEditor.SetField N_(i), bookInfo(i)
Next

'If Index = mSavedInIndex Then
        'Dim i As Long
'        Dim c As Long
'        Dim text As String
'        text = cboValue(mSavedInIndex).text
'        c = cboValue(mSavedInIndex).ListCount - 1
'        For i = 0 To c
'            If cboValue(mSavedInIndex).List(i) = text Or cboValue(mSavedInIndex).List(i) & "\" = text Then GoTo NoSaved
'        Next
'        cboValue(mSavedInIndex).AddItem text
'NoSaved:
    'End If




Set mTask = vtask



End Sub

Public Sub UpdateFromClipboard(Optional vText As String)
Dim vData() As String ' As SSLIB_BOOKINFO
vData() = SSLIB_ParseInfoText(vText)
If SafeUBound(vData()) < CST_SSLIB_FIELDS_LBound Then Exit Sub
On Error Resume Next
Dim i As Long
For i = CST_SSLIB_FIELDS_LBound To CST_SSLIB_FIELDS_UBound
    If vData(i) <> "" Then BookEditor.SetField N_(i), vData(i)
Next
'
'With vData
'    If .Author <> "" Then SetField N_(SSF_Author), .Author
'    If .title <> "" Then SetField N_(SSF_Title), .title
'    If .Publisher <> "" Then SetField N_(SSF_Publisher), .Publisher
'    If .SSID <> "" Then SetField N_(SSF_SSID), .SSID
'    If .Header <> "" Then SetField N_(SSF_HEADER), .Header
'    If .URL <> "" Then SetField N_(SSF_URL), .URL
'    If .About <> "" Then SetField N_(SSF_Comments), .About
'    If .Subject <> "" Then SetField N_(SSF_Subject), .Subject
'    If .PagesCount <> "" Then SetField N_(SSF_PagesCount), .PagesCount
'    If .PublishedDate <> "" Then SetField N_(SSF_PublishDate), .PublishedDate
'End With
End Sub




Public Sub Reset()
    MForms.ResetForm Me
    Dim i As Long
    For i = CST_SSLIB_FIELDS_LBound To CST_SSLIB_FIELDS_UBound
        If i <> SSLIBFields.SSF_SAVEDIN Then
            BookEditor.SetField N_(i), ""
        End If
    Next
End Sub

Private Sub cmdUpdate_Click()
    UpdateFromClipboard
End Sub



Private Sub Form_Activate()

'    SetKeyStyle SSLIB_ChnFieldName(SSF_SAVEDIN), VCT_Combox
'    SetKeyStyle SSLIB_ChnFieldName(SSF_SAVEDIN), VCT_DIR
'    SetKeyStyle SSLIB_ChnFieldName(SSF_HEADER), VCT_MultiLine
'

    'Timer.Interval = 0
    Set Timer = New CTimer
    If chkDetectClipboard.value Then
        Timer.Interval = cst_Timer_Interval ' True
    Else
        Timer.Interval = 0
    End If

    If m_bEditTaskMode Then
        cmdOk.Caption = "确认"
    Else
        cmdOk.Caption = "添加任务"
    End If

'    If m_bSolidMode Then
'        Me.BorderStyle = 0
'        cmdCancel.Enabled = False
'    Else
'        Me.BorderStyle = 1
'        cmdCancel.Enabled = True
'    End If

    'Form_Resize
End Sub

Private Sub Form_Initialize()


    'ReDim Map(0 To CST_SSLIB_FIELDS_TASKS_UBOUND - CST_SSLIB_FIELDS_LBound, 0 To 1) As String

   



End Sub

Private Sub Form_Load()
    'Me.Icon = frmMain.Icon
    'Set Timer = New CTimer
    'Timer.Interval = cst_Timer_Interval

    'If mKeyCount > 0 Then Me.ResetForm
     SSLIB_Init
    BookEditor.TwoColumnMode = True
    
    BookEditor.AddItem "任务名", VCT_NORMAL, "", False
    
    Dim i As Long
    For i = CST_SSLIB_FIELDS_LBound To CST_SSLIB_FIELDS_TASKS_UBOUND
        BookEditor.AddItem SSLIB_ChnFieldName(i), VCT_NORMAL, "", False
        'Map(i - CST_SSLIB_FIELDS_LBound, 0) = SSLIB_ChnFieldName(i)
    Next
    'Process Map()

    BookEditor.SetFieldStyle SSLIB_ChnFieldName(SSF_SAVEDIN), VCT_Combox + VCT_DIR, False
    BookEditor.SetFieldStyle SSLIB_ChnFieldName(SSF_HEADER), VCT_MultiLine, True
    
    
    mSavedInIndex = BookEditor.SearchIndex(N_(SSF_SAVEDIN))
    


    'cboValue(mSavedInIndex).Tag = CST_FORM_FLAGS_NORESETING
    chkStartDownload.Tag = CST_FORM_FLAGS_NORESETING
    chkDetectClipboard.Tag = CST_FORM_FLAGS_NORESETING
    'Timer.Interval = 0 ' False
    Dim configHnd As CLiNInI
    Set configHnd = New CLiNInI
    With configHnd
        .source = App.Path & "\" & configFile
        ComboxItemsFromString BookEditor.GetValueControl(mSavedInIndex), .GetSetting("SavedIn", "Path")
        FormStateFromString Me, .GetSetting("Form", "State")
    End With
    Set configHnd = Nothing


End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Exit Sub
    Dim configHnd As CLiNInI
    Set configHnd = New CLiNInI
    With configHnd
        .source = App.Path & "\" & configFile
        .SaveSetting "SavedIn", "Path", ComboxItemsToString(BookEditor.GetValueControl(mSavedInIndex))
        .SaveSetting "Form", "State", FormStateToString(Me)
        .Save
    End With
    Set configHnd = Nothing
    'Set Timer = Nothing
End Sub

'Private Sub Timer_ThatTimer()
'
'End Sub

Private Sub DetectClipboard()
    Dim vText As String
    vText = Clipboard.GetText
    If vText <> "" Then
        'Clipboard.Clear
        UpdateFromClipboard vText
    End If
End Sub




Public Property Get AutoUnload() As Boolean
    AutoUnload = mAutoUnload
End Property

Public Property Let AutoUnload(ByVal bValue As Boolean)
    mAutoUnload = bValue
End Property


Private Sub cmdCancel_Click()
    mExitStatus = vbCancel
    Me.Hide
    Set Timer = Nothing
    If mAutoUnload Then Unload Me 'Else Me.Hide
End Sub

Private Function CheckValue() As Boolean
    CheckValue = True
    If BookEditor.GetField(SSLIB_ChnFieldName(SSF_Title)) = "" And BookEditor.GetField(SSLIB_ChnFieldName(SSF_SSID)) = "" Then
        MsgBox QuoteString(SSLIB_ChnFieldName(SSF_Title)) & " " & QuoteString(SSLIB_ChnFieldName(SSF_SSID)) & " 不能都为空", vbCritical
        CheckValue = False
        'Exit Function
    End If
    
End Function

Private Sub cmdOK_Click()



If CheckValue() = False Then Exit Sub

Me.Enabled = False

Dim timerState As Boolean
timerState = (Timer.Interval = 0)
Timer.Interval = 0 ' False
Dim i As Long
Dim vArray() As String
    vArray = SSLIB_CreateBookInfoArray()
    For i = CST_SSLIB_FIELDS_LBound To CST_SSLIB_FIELDS_UBound
        vArray(i) = BookEditor.GetField(SSLIB_ChnFieldName(i))
    Next

If mTask Is Nothing Then
    #If afTaskEbd = 1 Then
        frmMain.CallBack_AddTask BookEditor.GetField("任务名"), vArray, chkStartDownload.value
    #Else
    CallBack_AddTask BookEditor.GetField("任务名"), vArray, chkStartDownload.value
    #End If
Else
    mTask.bookInfo.LoadFromArray vArray
    mTask.Name = BookEditor.GetField("任务名")
'    With mTask
'    .title = txtTitle.text
'    .Author = GetField(N_(SSF_Author))
'    .Publisher = GetField(N_(SSF_Publisher))
'    .SSID = GetField(N_(SSF_SSID))
'    .RootURL = txtUrl.text
'    .HttpHeader = txtHeader.text
'    .SavedIN = cboValue(mSavedInIndex).text
'    .AdditionalText = txtAddInfo.text
'    .PublishedDate = txtPublishedDate.text
'    .PagesCount = txtPagesCount.text
'    End With
    #If afTaskEbd = 1 Then
    frmMain.CallBack_EditTask mTask
    #Else
    CallBack_EditTask mTask
    #End If
    Set mTask = Nothing
End If


       ' Dim i As Long
       On Error Resume Next
        Dim c As Long
        Dim text As String
        Dim cboBox As Control
        Set cboBox = BookEditor.GetValueControl(mSavedInIndex)
        
        text = cboBox.text
    If text <> "" Then
        c = cboBox.ListCount - 1
        For i = 0 To c
            If cboBox.List(i) = text Or cboBox.List(i) & "\" = text Then GoTo NoSaved
        Next
         cboBox.AddItem text
NoSaved:
    End If


Me.Enabled = True
Reset

If timerState Then
    Timer.Interval = cst_Timer_Interval
Else
    Timer.Interval = 0
End If

If m_bEditTaskMode Then
    m_bEditTaskMode = False
    Set Timer = Nothing
    Me.Hide
End If

    If mAutoUnload Then
        Unload Me
    Else
        If chkDetectClipboard.value Then
            Clipboard.Clear
        End If
    End If

End Sub

Private Sub cmdReset_Click()
    Reset
End Sub

Private Sub Form_Resize()
    On Error Resume Next


    fraAddition.Top = Me.ScaleHeight - 120 - fraAddition.Height ' txtBox.Top + txtBox.Height + 120
    fraAddition.Left = Me.ScaleWidth - 120 - fraAddition.Width


    BookEditor.Move 120, 120, Me.ScaleWidth - 240, fraAddition.Top - 120



End Sub





Private Function SafeUBound(ByRef mArray() As String) As Long
    On Error GoTo ErrorSafeUbound
    SafeUBound = UBound(mArray())
    Exit Function

ErrorSafeUbound:
    SafeUBound = -1
End Function

Public Property Get ExitStatus() As VbMsgBoxResult
    ExitStatus = mExitStatus
End Property

Private Sub Form_Terminate()
    cmdCancel_Click
    'Me.Hide
    'Unload Me
End Sub




Private Sub Timer_ThatTime()
    If chkDetectClipboard.value Then DetectClipboard
End Sub


