VERSION 5.00
Begin VB.UserControl KeyValueEditor 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4320
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8190
   ScaleHeight     =   4320
   ScaleWidth      =   8190
   ToolboxBitmap   =   "KeyValueEditor.ctx":0000
   Begin VB.PictureBox fraContent 
      BorderStyle     =   0  'None
      Height          =   3705
      Left            =   315
      ScaleHeight     =   3705
      ScaleWidth      =   6930
      TabIndex        =   2
      Top             =   195
      Width           =   6930
      Begin VB.CheckBox chkValue 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "是的"
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   0
         Left            =   3090
         TabIndex        =   10
         Top             =   2610
         Visible         =   0   'False
         Width           =   1260
      End
      Begin VB.TextBox txtValueEx 
         Height          =   1425
         Index           =   0
         Left            =   420
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   9
         Top             =   615
         Visible         =   0   'False
         Width           =   5295
      End
      Begin VB.TextBox txtValue 
         Height          =   375
         Index           =   0
         Left            =   75
         TabIndex        =   8
         Top             =   555
         Visible         =   0   'False
         Width           =   5295
      End
      Begin VB.ComboBox cboValue 
         Height          =   315
         Index           =   0
         Left            =   0
         TabIndex        =   7
         Top             =   525
         Visible         =   0   'False
         Width           =   3120
      End
      Begin VB.CommandButton cmdDir 
         Caption         =   "目录..."
         Height          =   345
         Index           =   0
         Left            =   3390
         TabIndex        =   6
         Top             =   0
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdFile 
         Caption         =   "文件..."
         Height          =   345
         Index           =   0
         Left            =   4965
         TabIndex        =   5
         Top             =   15
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label lblKey 
         AutoSize        =   -1  'True
         Caption         =   "Key:"
         Height          =   195
         Index           =   0
         Left            =   255
         TabIndex        =   4
         Top             =   465
         Visible         =   0   'False
         Width           =   2115
      End
      Begin VB.Label lblValue 
         Caption         =   "Value"
         Height          =   255
         Index           =   0
         Left            =   1575
         TabIndex        =   3
         Top             =   375
         Visible         =   0   'False
         Width           =   1395
      End
   End
   Begin VB.Frame fraMain 
      BorderStyle     =   0  'None
      Height          =   4155
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7995
      Begin VB.VScrollBar VScroller 
         CausesValidation=   0   'False
         Height          =   4245
         LargeChange     =   10
         Left            =   7680
         Max             =   20
         TabIndex        =   1
         Top             =   120
         Width           =   270
      End
   End
End
Attribute VB_Name = "KeyValueEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'CSEH: ErrResumeNext
Option Explicit

'Option Explicit
'Private mMap As CStringMap
'Private mKeys() As String

Private mKeyValues() As String
Private mKeyType() As ValueObjectType
Private Const CST_ARRAY_TRUNK_NEW As Long = 16
Private mArrayUbound As Long
Private mKeyCount As Long


'Private mAutoUnload As Boolean
'Private mExitStatus As VbMsgBoxResult
Private mTwoColumnMode As Boolean

Public Enum ValueObjectType
    VCT_DEFAULT = 0
    VCT_NORMAL = 1
    VCT_MultiLine = 2
    VCT_FILE = 4
    VCT_DIR = 8
    VCT_Combox = 16
    VCT_HalfWidth = 32
    VCT_Label = 64
    VCT_Checked = 128
'    VCT_Disabled = 64
'    VCT_Enabled = 128
End Enum

Private Enum KeyObjectType
    KctLabel = 1
    kctTextBox = 2
    kctTextBoxEx = 4
    kctCombox = 8
    kctCmdFile = 16
    KctCmdDirectory = 32
    kctTextLabel = 64
    kctCheckBox = 128
End Enum

Public Enum ValueEditObjectType
    VectTextBox
    VectTextBoxMulti
    VectCombox
    VectLabel
    VectCheckBox
End Enum

Private Enum KeyObjectProperty
    KcpEnabled
    KcpAppearance
End Enum

Private Const CST_KeyObjectType_LBound As Long = 0
Private Const CST_KeyObjectType_UBound As Long = 6

'Event SelectDirectory(ByVal vKeyName As String, ByRef vDirectory As String)
'Event SelectFile(ByVal vKeyName As String, ByRef vFilename As String)
'Event KeyAdded(ByVal vKeyName As String)
'Event ValueChanged(ByVal vKeyName As String, ByVal vValue As String)
'Default Property Values:
'Const m_def_Result = 0
'Const m_def_Source = "0"
'Const m_def_TwoColumnMode = True
''Const m_def_Result = 0
''Const m_def_Source = 0
''Const m_def_TwoColumnMode = 0
''Property Variables:
'Dim m_Result As Variant
'Dim m_Source As String
'Dim m_TwoColumnMode As Boolean
'Dim m_Result As Variant
'Dim m_Source As Boolean
'Dim m_TwoColumnMode As Variant
'Event Declarations:
Event SelectDirectory(ByVal vKeyName As String, ByRef vdirectory As String)
Event SelectFile(ByVal vKeyName As String, ByRef vFilename As String)
Event KeyAdded(ByVal vKeyName As String)
Event ValueChanged(ByVal vKeyName As String, ByVal vValue As String)
'Event BeforeInitialized()


Private Sub SetValueObject(ByVal vIdx As Long, ByVal vType As ValueEditObjectType)
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>
    LoadObject vIdx, kctTextLabel, vType <> VectLabel
    LoadObject vIdx, kctTextBox, vType <> VectTextBox
    LoadObject vIdx, kctTextBoxEx, vType <> VectTextBoxMulti
    LoadObject vIdx, kctCombox, vType <> VectCombox
    LoadObject vIdx, kctCheckBox, vType <> VectCheckBox
End Sub

Public Function GetValueObject(ByVal vIndex As Long) As Object
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>
        
        Dim vcType As KeyObjectType
        
        If mKeyType(vIndex) And VCT_Combox Then
            vcType = kctCombox
        ElseIf mKeyType(vIndex) And VCT_Label Then
            vcType = kctTextLabel
        ElseIf mKeyType(vIndex) And VCT_MultiLine Then
            vcType = kctTextBoxEx
        ElseIf mKeyType(vIndex) And VCT_Checked Then
            vcType = kctCheckBox
        Else
            vcType = kctTextBox
        End If
    
        Set GetValueObject = GetKeyObject(vIndex, vcType)
        If GetValueObject Is Nothing Then Set GetValueObject = txtValue(vIndex)
   
    Exit Function

GetValueObject_Err:
    Err.Clear
    
    '</EhFooter>
End Function
'CSEH: ErrExit
Private Function GetKeyObject(ByVal vIdx As Long, ByVal vType As KeyObjectType) As Object
    '<EhHeader>
    On Error GoTo GetKeyObject_Err
    '</EhHeader>
    On Error Resume Next
    Dim ret As Object
    Select Case vType
    
        Case KeyObjectType.KctCmdDirectory
            Set ret = cmdDir(vIdx)
        Case KeyObjectType.kctCmdFile
            Set ret = cmdFile(vIdx)
        Case KeyObjectType.kctCombox
            Set ret = cboValue(vIdx)
        Case KeyObjectType.kctTextBox
            Set ret = txtValue(vIdx)
        Case KeyObjectType.kctTextBoxEx
            Set ret = txtValueEx(vIdx)
        Case KeyObjectType.kctTextLabel
            Set ret = lblValue(vIdx)
        Case KeyObjectType.kctCheckBox
            Set ret = chkValue(vIdx)
        Case Else
            Set ret = lblKey(vIdx)
    End Select
    Set GetKeyObject = ret
    '<EhFooter>
    Exit Function

GetKeyObject_Err:
    Err.Clear

    '</EhFooter>
End Function

Private Sub SetKeyObjectProperty(ByVal vIdx As Long, ByVal vProperty As KeyObjectProperty, ByVal vValue As Integer)
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>
    Dim I As Long
    On Error Resume Next
    Dim ctl As Object
    If vProperty = KcpAppearance Then
        For I = CST_KeyObjectType_LBound To CST_KeyObjectType_UBound
            Set ctl = GetKeyObject(vIdx, 2 ^ I)
            If Not ctl Is Nothing Then ctl.Appearance = vValue
        Next
    Else
        For I = CST_KeyObjectType_LBound To CST_KeyObjectType_UBound
            Set ctl = GetKeyObject(vIdx, 2 ^ I)
            If Not ctl Is Nothing Then ctl.Enabled = vValue
        Next
    End If
End Sub

Private Function LoadObject(ByVal vIdx As Long, ByVal vType As KeyObjectType, Optional vUnloadMode As Boolean = False) As Object
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>
    Dim ret As Object
    On Error Resume Next
    If vUnloadMode Then
        If vType And kctTextLabel Then
            Unload lblValue(vIdx)
        End If
        If vType And KctCmdDirectory Then
            Unload cmdDir(vIdx)
        End If
        If vType And kctCmdFile Then
            Unload cmdFile(vIdx)
        End If
        If vType And kctCombox Then
            Unload cboValue(vIdx)
        End If
        If vType And kctTextBox Then
            Unload txtValue(vIdx)
        End If
        If vType And kctTextBoxEx Then
            Unload txtValueEx(vIdx)
        End If
        If vType And KctLabel Then
            Unload lblKey(vIdx)
        End If
        If vType And kctCheckBox Then
            Unload chkValue(vIdx)
        End If
    Else
        If vType And kctTextLabel Then
            Load lblValue(vIdx)
            lblValue(vIdx).Visible = True
            lblValue(vIdx).Enabled = lblKey(vIdx).Enabled
            Set ret = lblValue(vIdx)
        End If
        If vType And KctCmdDirectory Then
            Load cmdDir(vIdx)
            cmdDir(vIdx).Visible = True
            cmdDir(vIdx).Enabled = lblKey(vIdx).Enabled
            Set ret = cmdDir(vIdx)
        End If
        If vType And kctCmdFile Then
            Load cmdFile(vIdx)
            cmdFile(vIdx).Visible = True
            cmdFile(vIdx).Enabled = lblKey(vIdx).Enabled
            Set ret = cmdFile(vIdx)
        End If
        If vType And kctCombox Then
            Load cboValue(vIdx)
            cboValue(vIdx).Visible = True
            cboValue(vIdx).Enabled = lblKey(vIdx).Enabled
            Set ret = cboValue(vIdx)
        End If
        If vType And kctTextBox Then
            Load txtValue(vIdx)
            txtValue(vIdx).Visible = True
            txtValue(vIdx).Enabled = lblKey(vIdx).Enabled
            Set ret = txtValue(vIdx)
        End If
        If vType And kctTextBoxEx Then
            Load txtValueEx(vIdx)
            txtValueEx(vIdx).Visible = True
            txtValueEx(vIdx).Enabled = lblKey(vIdx).Enabled
            Set ret = txtValueEx(vIdx)
        End If
        If vType And KctLabel Then
            Load lblKey(vIdx)
            lblKey(vIdx).Visible = True
            Set ret = lblKey(vIdx)
        End If
        If vType And kctCheckBox Then
            Load chkValue(vIdx)
            chkValue(vIdx).Visible = True
            chkValue(vIdx).Enabled = lblKey(vIdx).Enabled
            Set ret = chkValue(vIdx)
        End If
        Set LoadObject = ret
    End If
End Function
'Private mTask As CTask
'Const configFile As String = "taskdef.ini"
'Private m_bSolidMode As Boolean
'Private m_bEditTaskMode As Boolean

'Private WithEvents Timer As CTimer
'Private Const cst_Timer_Interval As Long = 400
'Private mSavedInIndex As Long


Public Property Get FieldEnabled(ByVal vKey As String) As Boolean
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>
    On Error Resume Next
    Dim I As Long
    I = SearchIndex(vKey)
    If I < 0 Then Exit Property
    FieldEnabled = lblKey(I).Enabled
End Property

Public Property Let FieldEnabled(ByVal vKey As String, ByVal vEnabled As Boolean)
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>
    On Error Resume Next
    Dim I As Long
    I = SearchIndex(vKey)
    If I < 0 Then Exit Property
    SetKeyObjectProperty I, KcpEnabled, vEnabled
    
End Property

Public Sub AddItem(ByVal vKey As String, _
    Optional vStyle As ValueObjectType = VCT_NORMAL, _
    Optional ByVal vValue As String = vbNullString, _
    Optional vRefresh As Boolean = True)
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>
    On Error Resume Next
    Dim I As Long
    I = SearchIndex(vKey)
    If I >= 0 Then
        SetItemStyle I, vStyle, vValue
    Else
        I = mKeyCount
        mKeyCount = mKeyCount + 1
        Do While mKeyCount > mArrayUbound
            ReDim Preserve mKeyValues(0 To 1, 0 To mArrayUbound + CST_ARRAY_TRUNK_NEW)
            ReDim Preserve mKeyType(0 To mArrayUbound + CST_ARRAY_TRUNK_NEW)
            mArrayUbound = mArrayUbound + CST_ARRAY_TRUNK_NEW
        Loop
        
        mKeyValues(0, mKeyCount - 1) = vKey
        
        Dim C As Object
        Set C = LoadObject(mKeyCount - 1, KctLabel)
        C.Caption = vKey & ":"
        SetItemStyle mKeyCount - 1, vStyle, vValue, vRefresh
    End If
End Sub


Public Property Get result() As String()
Attribute result.VB_MemberFlags = "400"
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>
    result = mKeyValues
End Property

Public Property Let Source(ByRef vSource() As String)
Attribute Source.VB_MemberFlags = "400"
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>
    Process vSource
End Property


Private Sub cmdDir_Click(Index As Integer)
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>
    On Error Resume Next
    Dim ret As String
    ret = GetText(Index)
    RaiseEvent SelectDirectory(mKeyValues(0, Index), ret)
    SetText Index, ret
End Sub

Private Sub cmdFile_Click(Index As Integer)
    'Static lastSelect As String
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>
    On Error Resume Next
    Dim ret As String
    ret = GetText(Index)
    RaiseEvent SelectFile(mKeyValues(0, Index), ret)
    SetText Index, ret
End Sub
''
'''Private Sub cmdOK_Click()
'''
'''    If ExitBeforeOK() Then Exit Sub
'''    Dim i As Long
'''    For i = 0 To mKeyCount - 1
'''        mKeyValues(i, 1) = GetText(i)
'''    Next
'''    mExitStatus = vbOK
'''
'''    If ExitAfterOK() Then Exit Sub
'''    If mAutoUnload Then Unload Me
'''
'''End Sub
''
'''Private Sub cmdReset_Click()
'''    On Error Resume Next
'''    If ExitBeforeReset Then Exit Sub
'''    Dim i As Long
'''    For i = 0 To mKeyCount - 1
'''        GetValueObject(i).text = vbNULLSTRING
'''    Next
'''    If ExitAfterReset Then Exit Sub
'''End Sub

Public Property Get TwoColumnMode() As Boolean
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>
    TwoColumnMode = mTwoColumnMode
End Property

Public Property Let TwoColumnMode(ByVal bValue As Boolean)
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>
    mTwoColumnMode = bValue
    PropertyChanged "TwoColumnMode"
End Property



Private Sub UserControl_Initialize()
'    Dim a(0 To 10, 0 To 1) As String
'    Dim i As Long
'    For i = 0 To 10
'        a(i, 0) = Chr(Asc("A") + i)
'        a(i, 1) = String$(10, a(i, 0))
'    Next
'    Me.TwoColumnMode = True
'    Process a()
'    UserControl_Resize
End Sub

Private Sub UserControl_Resize()
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>
    On Error Resume Next
    Dim I As Long
    Dim u As Long
    Static pResizing As Boolean
    
    If pResizing Then Exit Sub
    pResizing = True
    
    'me.
    
    
'    fraAddition.Top = UserControl.ScaleHeight - 120 - fraAddition.Height  ' txtBox.Top + txtBox.Height + 120
'    fraAddition.Left = Me.ScaleWidth - 120 - fraAddition.Width
    

    
    fraMain.Top = 0
    fraMain.Left = 0
    fraMain.Width = UserControl.ScaleWidth
    fraMain.Height = UserControl.ScaleHeight ' fraAddition.Top - 120
    
'    SplitLine.X1 = 0
'    SplitLine.X2 = Me.ScaleWidth
'    SplitLine.Y1 = fraAddition.Top - 60
'    SplitLine.Y2 = SplitLine.Y1
    
    VScroller.Left = fraMain.Width - VScroller.Width '- 20
    VScroller.Top = 0
    VScroller.Height = fraMain.Height ' - 140
    
    
    fraContent.Move 0, 0, VScroller.Left - 120
    
    
    Dim dist As Single
    Dim valueObject As Object
    
    Set valueObject = GetValueObject(0)
    
    Dim pW As Long
    Dim pH As Long
    Dim pL As Long
    Dim pT As Long
    
    
    Dim fullWidth As Long
    fullWidth = fraContent.Width
    Dim halfWidth As Long
    halfWidth = fullWidth / 2 '- 240
    Dim fullLeft As Long
    fullLeft = 120
    Dim halfLeft As Long
    halfLeft = halfWidth + 120
    Dim fullTop As Long
    fullTop = 120
    Dim halfTop As Long
    halfTop = 120
    
    If Not mTwoColumnMode Then
        pW = fullWidth
        pL = fullLeft
        pT = fullTop
    Else
        pW = halfWidth
        pL = halfLeft
        pT = halfTop
    End If
    

    lblKey(0).Move 120, 120
    valueObject.Move lblKey(0).Left, lblKey(0).Top + lblKey(0).Height + 120
    'lblKey(0).Move 60, 60
    cmdFile(0).Move pW - 120 - cmdFile(0).Width, valueObject.Top
    cmdDir(0).Move cmdFile(0).Left, cmdFile(0).Top
    'txtValue(0).Move lblKey(0).Left, cmdFile(0).Top + cmdFile(0).Height + 120
        If cmdFile(0).Visible Or cmdDir(0).Visible Then
            valueObject.Width = cmdFile(0).Left - 2 * valueObject.Left
        Else
            valueObject.Width = pW - 2 * valueObject.Left
        End If

    'txtValueEx(0).Move txtValue(0).Left, txtValue(0).Top, txtValue(0).Width
    
    Dim lastValueObject As Object
    Set lastValueObject = valueObject
    For I = 1 To mKeyCount - 1
            dist = lastValueObject.Top + lastValueObject.Height + 120 - lblKey(I - 1).Top
            Set valueObject = GetValueObject(I)
            lblKey(I).Move lblKey(I - 1).Left, lblKey(I - 1).Top + dist
            valueObject.Move lastValueObject.Left, lastValueObject.Top + dist
            If ObjectsVisible(cmdFile, I) Or ObjectsVisible(cmdDir, I) Then
                valueObject.Width = cmdFile(0).Left - 2 * valueObject.Left
            Else
                valueObject.Width = pW - 2 * valueObject.Left
            End If
            
            cmdFile(I).Move cmdFile(0).Left, valueObject.Top
            cmdDir(I).Move cmdDir(0).Left, valueObject.Top
                   
            Set lastValueObject = valueObject
    Next
'    End If
    
    I = mKeyCount - 1 'lblKey.UBound
'
    Dim txtBox As Object
    Set txtBox = GetValueObject(I)
    fraContent.Height = txtBox.Top + txtBox.Height ' + 240
    
    If mTwoColumnMode Then
        Dim splitTop As Long
        Dim splitIndex As Long
        splitIndex = -1
        splitTop = fraContent.Height / 2 - 60
        'Dim i As Long
        For I = 0 To mKeyCount - 1
            If lblKey(I).Top > splitTop Then
                splitIndex = I
                Exit For
            End If
        Next
        If splitIndex >= 0 Then
            splitTop = lblKey(splitIndex).Top - 120
            For I = splitIndex To mKeyCount - 1
                Set valueObject = GetValueObject(I)
                lblKey(I).Move lblKey(I).Left + halfWidth, lblKey(I).Top - splitTop
                valueObject.Move valueObject.Left + halfWidth, valueObject.Top - splitTop
                cmdFile(I).Move cmdFile(I).Left + halfWidth, cmdFile(I).Top - splitTop
                cmdDir(I).Move cmdDir(I).Left + halfWidth, cmdDir(I).Top - splitTop
            Next
            fraContent.Move fraContent.Left, fraContent.Top, fraContent.Width, splitTop + 360
        End If
    End If
    
    If fraMain.Height < fraContent.Height Then VScroller.Enabled = True Else VScroller.Enabled = False
    
'    If Me.ScaleHeight < fraAddition.Top + fraAddition.Height + 120 Then
'        Me.Height = fraAddition.Top + fraAddition.Height + 360
'    End If
'
'    If txtValue(i).Visible Then
'            cmdCancel.Move Me.ScaleWidth - cmdCancel.Width - 120, txtValue(i).Top + txtValue(i).Height + 120
'        Else
'           cmdCancel.Move Me.ScaleWidth - cmdCancel.Width - 120, txtValueEx(i).Top + txtValueEx(i).Height + 120
'    End If
'
'    cmdOK.Move cmdCancel.Left - cmdOK.Width - 240, cmdCancel.Top
    
    pResizing = False
    
    'Me.Enabled = True
End Sub

Private Sub SetItemStyle(ByVal vIndex As Long, Optional vStyle As ValueObjectType = VCT_DEFAULT, Optional vValue As String, Optional vRefresh As Boolean = True)
    '<EhHeader>
    'On Error Resume Next
    '</EhHeader>
    On Error Resume Next
    Dim C As Object
    If vIndex >= 0 And vIndex < mKeyCount Then
        If vValue = vbNullString Then vValue = GetText(vIndex) ' GetValueObject(vIndex).Text

    
        If Not vStyle = VCT_DEFAULT Then mKeyType(vIndex) = vStyle
        
        If vStyle = VCT_DEFAULT Then
            Dim txtObject As Object
            Set txtObject = GetValueObject(vIndex)
            If txtObject Is Nothing Then
                Set txtObject = LoadObject(vIndex, kctTextBox)
                'Load txtValue(vIndex)
                'Set txtObject = txtValue
            End If
            SetText vIndex, vValue
            'txtObject.Text = vValue
            txtObject.Visible = True
            txtObject.Enabled = lblKey(vIndex).Enabled
            Exit Sub
        End If
        
        
        If vStyle And ValueObjectType.VCT_NORMAL Then
            SetValueObject vIndex, VectTextBox
            SetText vIndex, vValue
            
            'Set c = LoadObject(vIndex, kctTextBox)
            'c.Text = vValue
            'LoadObject vIndex, kctTextBoxEx + kctCombox, True
'            'mKeyType(vIndex) = vStyle
'            Load txtValue(vIndex)
'            txtValueEx(vIndex).Visible = False
'            cboValue(vIndex).Visible = False
'            txtValue(vIndex).Visible = True
'            txtValue(vIndex).Text = vValue
            'Set c = Nothing
        End If
        
        If vStyle And ValueObjectType.VCT_Combox Then
            SetValueObject vIndex, VectCombox
            SetText vIndex, vValue
'
'            Set c = LoadObject(vIndex, kctCombox)
'            c.Text = vValue
'            LoadObject vIndex, kctTextBox + kctTextBoxEx, True
'            Set c = Nothing
'
'            txtValue(vIndex).Visible = False
'            txtValueEx(vIndex).Visible = False
'            Load cboValue(vIndex)
'            cboValue(vIndex).Visible = True
'            cboValue(vIndex).Text = vValue
            'mKeyType(vIndex) = vStyle
        End If
        
        If vStyle And ValueObjectType.VCT_MultiLine Then
            SetValueObject vIndex, VectTextBoxMulti
            SetText vIndex, vValue
'
'            txtValue(vIndex).Visible = False
'            cboValue(vIndex).Visible = False
'            Load txtValueEx(vIndex)
'            txtValueEx(vIndex).Visible = True
'            txtValueEx(vIndex).Text = vValue
'            'mKeyType(vIndex) = vStyle
        End If
        
        If vStyle And ValueObjectType.VCT_Label Then
            SetValueObject vIndex, VectLabel
            SetText vIndex, vValue
        End If
        If vStyle And ValueObjectType.VCT_DIR Then
            LoadObject vIndex, KctCmdDirectory
            
            'Load cmdDir(vIndex)
            cmdFile(vIndex).Visible = False
            'cmdDir(vIndex).Visible = True
            
        End If
        
        If vStyle And ValueObjectType.VCT_FILE Then
            LoadObject vIndex, kctCmdFile
            
            'Load cmdFile(vIndex)
            cmdDir(vIndex).Visible = False
            'cmdFile(vIndex).Visible = True
        End If
        
        If vStyle And VCT_Checked Then
            SetValueObject vIndex, VectCheckBox
            SetText vIndex, vValue
        End If
'        If vStyle And ValueObjectType.VCT_Enabled Then
'
'        End If
'
'        If vStyle And VCT_Disabled Then
'        End If
        
        If vRefresh Then UserControl_Resize
    End If
End Sub

Public Sub SetFieldStyle(ByVal vKey As String, _
                       Optional vStyle As ValueObjectType = VCT_DEFAULT, Optional vRefresh As Boolean = True)
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>

    SetItemStyle SearchIndex(vKey), vStyle, , vRefresh

End Sub
Public Function GetValueObjectByName(ByVal vName As String) As Object
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>
    Dim idx As Long
    idx = SearchIndex(vName)
    If idx >= 0 Then
        Set GetValueObjectByName = GetValueObject(idx)
    End If
End Function

'CSEH: ErrExit


Private Function ObjectsVisible(vObjectArray As Variant, ByVal vIndex As Long) As Boolean
    '<EhHeader>
    On Error GoTo ObjectsVisible_Err
    '</EhHeader>
    If ObjectsExists(vObjectArray, vIndex) = False Then Exit Function
    ObjectsVisible = vObjectArray(vIndex).Visible
    '<EhFooter>
    Exit Function

ObjectsVisible_Err:
    Err.Clear

    '</EhFooter>
End Function

'CSEH: ErrExit
Private Function ObjectsExists(vObjectArray As Variant, ByVal vIndex As Long) As Boolean
    '<EhHeader>
    On Error GoTo ObjectsExists_Err
    '</EhHeader>
        If vObjectArray(vIndex).Name <> vbNullString Then
            If Not Err Then ObjectsExists = True
        End If
    Err.Clear
    '<EhFooter>
    Exit Function

ObjectsExists_Err:
    Err.Clear

    '</EhFooter>
End Function

'Private Sub CopyPosition( _
'    ByRef vDest As Object, _
'    ByRef vSource As Object, _
'    Optional vLeft As Boolean = True, _
'    Optional vTop As Boolean = True, _
'    Optional vWidth As Boolean = True, _
'    Optional vHeight As Boolean = True)
'    If vDest Is Nothing Then Exit Sub
'    If vSource Is Nothing Then Exit Sub
'    On Error Resume Next
'    If vLeft Then vDest.Left = vSource.Left
'    If vTop Then vDest.Top = vSource.Top
'    If vHeight Then vDest.Height = vSource.Height
'    If vWidth Then vDest.Width = vSource.Width
'End Sub
'
'
Public Sub Reset()
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>
    On Error Resume Next

    cmdFile(0).Visible = False
    cmdDir(0).Visible = False
    cboValue(0).Visible = False
    txtValueEx(0).Visible = False
    txtValue(0).Visible = False
    Dim I As Long
    For I = 1 To mKeyCount - 1
        Unload cmdFile(I)
        Unload cmdDir(I)
        Unload cboValue(I)
        Unload txtValueEx(I)
        Unload lblValue(I)
    Next

    Erase mKeyValues
    Erase mKeyType

    mKeyCount = 0
    mArrayUbound = 0

End Sub

Public Sub Process(ByRef vKeyValues() As String)
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>

    Dim I As Long

    On Error Resume Next

    UserControl.Enabled = False

    Reset

    Dim iU As Long
    iU = SafeUBound(vKeyValues)
    For I = 0 To iU
        AddItem vKeyValues(I, 0), VCT_NORMAL, vKeyValues(I, 1), False
    Next

    UserControl.Enabled = True

    UserControl_Resize
End Sub

Public Property Get Field(ByVal vKey As String) As String
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>
    Dim I As Long
    I = SearchIndex(vKey)
    If I >= 0 Then Field = GetText(I)
End Property

Public Property Let Field(ByVal vKey As String, ByVal vValue As String)
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>
    Dim I As Long
    I = SearchIndex(vKey)
    If I >= 0 Then SetText I, vValue
End Property

Public Sub SetField(ByVal vKey As String, ByVal vValue As String)
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>
    Dim I As Long
    I = SearchIndex(vKey)
    If I >= 0 Then SetText I, vValue
End Sub

Public Function GetField(ByVal vKey As String) As String
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>
    Dim I As Long
    I = SearchIndex(vKey)
    GetField = GetText(I)
End Function

Public Function GetText(ByVal Index As Long) As String
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>
On Error Resume Next
    If Index > mKeyCount Then Exit Function
    If Index < 0 Then Exit Function
    Dim txtBox As Object
    Set txtBox = GetValueObject(Index)
    GetText = txtBox.Text
    If Err.Number <> 0 Then
        Err.Clear
        GetText = CStr(txtBox.Value)
    End If
    If Err.Number <> 0 Then
        Err.Clear
        GetText = txtBox.Caption
    End If

'    Exit Function
'ErrorGetText:
End Function

Public Sub SetText(ByVal Index As Long, ByRef vValue As String)
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>
On Error Resume Next
    If Index >= mKeyCount Then Exit Sub
    If Index < 0 Then Exit Sub
    Dim txtBox As Object
    Set txtBox = GetValueObject(Index)
    txtBox.Text = vValue
    If Err.Number <> 0 Then
        Err.Clear
        If vValue <> vbNullString Then
            txtBox.Value = 1
            'txtBox.Caption = "是"
        Else
            txtBox.Value = 0
            'txtBox.Caption = "否"
        End If
    End If
    If Err.Number <> 0 Then
        txtBox.Caption = vValue
        Err.Clear
    End If

End Sub
'CSEH: ErrExit
Public Function SearchIndex(ByVal vKey As String) As Long
    '<EhHeader>
    On Error GoTo SearchIndex_Err
    '</EhHeader>
    SearchIndex = -1
    Dim I As Long
    For I = 0 To mKeyCount - 1
        If mKeyValues(0, I) = vKey Then SearchIndex = I: Exit Function
    Next
    '<EhFooter>
    Exit Function

SearchIndex_Err:
    Err.Clear

    '</EhFooter>
End Function



'Public Sub SetMultiLine(ByVal vKey As String)
'    Dim i As Long
'    i = SearchIndex(vKey)
'    If i >= 0 Then
'        txtValueEx(i).Visible = True
'        txtValue(i).Visible = False
'        UserControl_Resize
'    End If
'End Sub
'
'Public Sub SetDirectory(ByVal vKey As String)
'    Dim i As Long
'    i = SearchIndex(vKey)
'    If i >= 0 Then
'        cmdFile(i).Visible = False
'        cmdDir(i).Visible = True
'        UserControl_Resize
'    End If
'End Sub
'
'Public Sub SetFile(ByVal vKey As String)
'    Dim i As Long
'    i = SearchIndex(vKey)
'    If i >= 0 Then
'        cmdFile(i).Visible = True
'        cmdDir(i).Visible = False
'        UserControl_Resize
'    End If
'End Sub
Private Function SafeUBound(ByRef mArray() As String) As Long
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>
    On Error GoTo ErrorSafeUbound
    SafeUBound = UBound(mArray())
    Exit Function
    
ErrorSafeUbound:
    SafeUBound = -1
End Function

Private Sub UserControl_Show()
    UserControl_Resize
End Sub

'Public Property Get ExitStatus() As VbMsgBoxResult
'    ExitStatus = mExitStatus
'End Property

'Private Sub Form_Terminate()
'    cmdCancel_Click
'    'Me.Hide
'    Unload Me
'End Sub

'Public Function SelectFile(ByVal vFilename As String) As String
''TODO
'End Function
'
'Public Function SelectDir(ByVal vDirectory As String) As String
'    Dim dlg As CFolderBrowser
'    Set dlg = New CFolderBrowser
'    If vDirectory <> vbNULLSTRING Then dlg.InitDirectory = vDirectory
'    dlg.Owner = Me.hWnd
'    Dim r As String
'    r = dlg.Browse
'    If r <> vbNULLSTRING Then
'        SelectDir = r
'    End If
''
''        Dim i As Long
''        Dim c As Long
''        c = cboValue(mSavedInIndex).ListCount - 1
''        For i = 0 To c
''            If cboValue(mSavedInIndex).List(i) = r Or cboValue(mSavedInIndex).List(i) = r & "\" Then Exit Sub
''        Next
''        cboValue(mSavedInIndex).AddItem r
''        cboValue(mSavedInIndex).text = r
''    End If 'TODO
'End Function

'Public Function ExitBeforeOK() As Boolean
''ToDo
'End Function
'
'Public Function ExitAfterOK() As Boolean
'Dim timerState As Boolean
'timerState = Timer.Interval = 0
'Timer.Interval = 0 ' False
'Dim i As Long
'Dim vArray() As String
'    vArray = SSLIB_CreateBookInfoArray()
'    For i = 0 To mKeyCount - 1
'        vArray(i + CST_SSLIB_FIELDS_LBound) = mKeyValues(i, 1)
'    Next
'
'If mTask Is Nothing Then
'
'    frmMain.CallBack_AddTask vArray, chkStartDownload.Value
'Else
'    mTask.bookInfo.LoadFromArray vArray
''    With mTask
''    .title = txtTitle.text
''    .Author = GetField(N_(SSF_Author))
''    .Publisher = GetField(N_(SSF_Publisher))
''    .SSID = GetField(N_(SSF_SSID))
''    .RootURL = txtUrl.text
''    .HttpHeader = txtHeader.text
''    .SavedIN = cboValue(mSavedInIndex).text
''    .AdditionalText = txtAddInfo.text
''    .PublishedDate = txtPublishedDate.text
''    .PagesCount = txtPagesCount.text
''    End With
'    frmMain.CallBack_EditTask mTask
'    Set mTask = Nothing
'End If
'If timerState Then
'    Timer.Interval = cst_Timer_Interval
'Else
'    Timer.Interval = 0
'End If
'
'       ' Dim i As Long
'        Dim c As Long
'        Dim text As String
'        text = cboValue(mSavedInIndex).text
'    If text <> vbNULLSTRING Then
'        c = cboValue(mSavedInIndex).ListCount - 1
'        For i = 0 To c
'            If cboValue(mSavedInIndex).List(i) = text Or cboValue(mSavedInIndex).List(i) & "\" = text Then GoTo NoSaved
'        Next
'        cboValue(mSavedInIndex).AddItem text
'NoSaved:
'    End If
'
'
'If m_bEditTaskMode Then
'    m_bEditTaskMode = False
'    Set Timer = Nothing
'    Me.Hide
'End If
'
''ToDo
'End Function

'Public Function ExitBeforeCancel() As Boolean
''ToDo
'End Function
'
'Public Function ExitAfterCancel() As Boolean
'    'Timer.Interval = 0 ' False
'    Me.Hide
'    Set Timer = Nothing
'End Function

'Public Function ExitBeforeReset() As Boolean
'    ExitBeforeReset = True
'    Me.Reset
'End Function
'
'Public Function ExitAfterReset() As Boolean
''ToDo
'End Function
'
'
'Private Sub Timer_ThatTime()
'    If Not chkDetectClipboard.Value Then Exit Sub
'    DetectClipboard
'End Sub

Private Sub VScroller_Change()
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>
    Dim p As Double
    Dim vAdd As Single
    Dim vTotal As Single
    vTotal = fraContent.Height - fraMain.Height - 240
    If vTotal < 0 Then Exit Sub
    vAdd = vTotal * (VScroller.Value / (VScroller.Max - VScroller.Min))
    fraContent.Top = -vAdd
    
End Sub


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Appearance
Public Property Get Appearance() As Integer
Attribute Appearance.VB_Description = "Returns/sets whether or not an object is painted at run time with 3-D effects."
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>
    Appearance = UserControl.Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As Integer)
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>
    On Error Resume Next
    UserControl.Appearance() = New_Appearance
    fraMain.Appearance = New_Appearance
    fraContent.Appearance = New_Appearance
    SetKeyObjectProperty 0, KcpAppearance, New_Appearance
'    Dim c As Object
'    Dim d As Object
'    Dim e As Object
'    For Each c In UserControl
'        Debug.Print c.Name
'        c.Appearance = New_Appearance
'        For Each d In c
'            d.Appearance = New_Appearance
'                For Each e In d
'                    e.Appearance = New_Appearance
'                Next
'        Next
'    Next
    Dim I As Long
    For I = 1 To mKeyCount - 1
        SetKeyObjectProperty I, KcpAppearance, New_Appearance
    Next
    PropertyChanged "Appearance"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BorderStyle
Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>
    UserControl.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackStyle
Public Property Get BackStyle() As Integer
Attribute BackStyle.VB_Description = "Indicates whether a Label or the background of a Shape is transparent or opaque."
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>
    BackStyle = UserControl.BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As Integer)
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>
    UserControl.BackStyle() = New_BackStyle
    PropertyChanged "BackStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=5
Public Sub Clear()
Attribute Clear.VB_Description = "Clear all values"
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>
     Dim I As Long
     For I = 0 To mKeyCount - 1
        SetText I, vbNullString
     Next
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,FillColor
Public Property Get FillColor() As OLE_COLOR
Attribute FillColor.VB_Description = "Returns/sets the color used to fill in shapes, circles, and boxes."
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>
    FillColor = UserControl.FillColor
End Property

Public Property Let FillColor(ByVal New_FillColor As OLE_COLOR)
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>
    On Error Resume Next
    UserControl.FillColor() = New_FillColor
'    Dim c As Object
'    For Each c In UserControl.Objects
'        c.FillColor = New_FillColor
'    Next
    PropertyChanged "FillColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,FillStyle
Public Property Get FillStyle() As Integer
Attribute FillStyle.VB_Description = "Returns/sets the fill style of a shape."
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>
    FillStyle = UserControl.FillStyle
End Property

Public Property Let FillStyle(ByVal New_FillStyle As Integer)
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>
    UserControl.FillStyle() = New_FillStyle
    PropertyChanged "FillStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>
    Set UserControl.Font = New_Font
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>
    ForeColor = UserControl.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>
    UserControl.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MemberInfo=14,0,0,0
'Public Property Get Result() As Variant
'    Result = m_Result
'End Property
'
'Public Property Let Result(ByVal New_Result As Variant)
'    m_Result = New_Result
'    PropertyChanged "Result"
'End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MemberInfo=0,0,0,0
'Public Property Get Source() As Boolean
'    Source = m_Source
'End Property
'
'Public Property Let Source(ByVal New_Source As Boolean)
'    m_Source = New_Source
'    PropertyChanged "Source"
'End Property



'Initialize Properties for User Object
Private Sub UserControl_InitProperties()
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>
    Set UserControl.Font = Ambient.Font
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>
    On Error Resume Next
    Appearance = PropBag.ReadProperty("Appearance", 1)
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    UserControl.BackStyle = PropBag.ReadProperty("BackStyle", 1)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    UserControl.FillColor = PropBag.ReadProperty("FillColor", &H0&)
    UserControl.FillStyle = PropBag.ReadProperty("FillStyle", 1)
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    UserControl.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    TwoColumnMode = PropBag.ReadProperty("TwoColumnMode", True)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>
    On Error Resume Next
    Call PropBag.WriteProperty("Appearance", UserControl.Appearance, 1)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 0)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("BackStyle", UserControl.BackStyle, 1)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("FillColor", UserControl.FillColor, &H0&)
    Call PropBag.WriteProperty("FillStyle", UserControl.FillStyle, 1)
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("ForeColor", UserControl.ForeColor, &H80000012)
    Call PropBag.WriteProperty("TwoColumnMode", mTwoColumnMode, True)
End Sub



