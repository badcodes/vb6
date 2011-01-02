VERSION 5.00
Begin VB.Form frmKeyValueEditor 
   Caption         =   "KeyValue Editor"
   ClientHeight    =   5535
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   8310
   LinkTopic       =   "Form1"
   ScaleHeight     =   5535
   ScaleWidth      =   8310
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cboValue 
      Height          =   315
      Index           =   0
      Left            =   150
      TabIndex        =   9
      Top             =   390
      Width           =   3120
   End
   Begin VB.Frame fraAddition 
      BorderStyle     =   0  'None
      Height          =   1140
      Left            =   345
      TabIndex        =   5
      Top             =   4260
      Width           =   7830
      Begin VB.CommandButton cmdReset 
         Caption         =   "重置"
         Height          =   345
         Left            =   3975
         TabIndex        =   8
         Top             =   585
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "取消"
         Height          =   345
         Left            =   6735
         TabIndex        =   7
         Top             =   585
         Width           =   1095
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "确认"
         Height          =   345
         Left            =   5385
         TabIndex        =   6
         Top             =   585
         Width           =   1095
      End
   End
   Begin VB.TextBox txtValueEx 
      Height          =   1170
      Index           =   0
      Left            =   150
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   390
      Width           =   5295
   End
   Begin VB.CommandButton cmdFile 
      Caption         =   "文件..."
      Height          =   345
      Index           =   0
      Left            =   6705
      TabIndex        =   3
      Top             =   390
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdDir 
      Caption         =   "目录..."
      Height          =   345
      Index           =   0
      Left            =   5580
      TabIndex        =   2
      Top             =   390
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtValue 
      Height          =   375
      Index           =   0
      Left            =   165
      TabIndex        =   1
      Top             =   375
      Width           =   5295
   End
   Begin VB.Label lblKey 
      AutoSize        =   -1  'True
      Caption         =   "Key:"
      Height          =   195
      Index           =   0
      Left            =   165
      TabIndex        =   0
      Top             =   165
      Width           =   2115
   End
End
Attribute VB_Name = "frmKeyValueEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Private mMap As CStringMap
'Private mKeys() As String
Private mKeyUBound As Long
Private mKeyValues() As String


Private mAutoUnload As Boolean
Private mExitStatus As VbMsgBoxResult

Public Enum ValueControlType
    VCT_Normal = 1
    VCT_MultiLine = 2
    VCT_FILE = 3
    VCT_DIR = 4
    VCT_Combox = 5
End Enum


Public Property Get AutoUnload() As Boolean
    AutoUnload = mAutoUnload
End Property

Public Property Let AutoUnload(ByVal bValue As Boolean)
    mAutoUnload = bValue
End Property


Private Sub cmdCancel_Click()
    If ExitBeforeCancel() Then Exit Sub
    mExitStatus = vbCancel
    If ExitAfterCancel() Then Exit Sub
    If mAutoUnload Then Unload Me 'Else Me.Hide
End Sub

Public Property Get Result() As String()
    Result = mKeyValues
End Property

Public Property Let Source(ByRef vSource() As String)
    Process vSource
End Property


Private Sub cmdDir_Click(Index As Integer)
    Static lastSelect As String
    Dim ret As String
    ret = SelectDir(lastSelect)
    If ret <> "" Then lastSelect = ret: SetText Index, ret
End Sub

Private Sub cmdFile_Click(Index As Integer)
    Static lastSelect As String
    Dim ret As String
    ret = SelectFile(lastSelect)
    If ret <> "" Then lastSelect = ret: SetText Index, ret
End Sub

Private Sub cmdOK_Click()
    
    If ExitBeforeOK() Then Exit Sub
    Dim i As Long
    For i = 0 To mKeyUBound
        mKeyValues(i, 1) = GetText(i)
    Next
    mExitStatus = vbOK

    If ExitAfterOK() Then Exit Sub
    If mAutoUnload Then Unload Me
    
End Sub

Private Sub cmdReset_Click()
    On Error Resume Next
    If ExitBeforeReset Then Exit Sub
    Dim i As Long
    For i = 0 To mKeyUBound
        GetValueControl(i).text = ""
    Next
    If ExitAfterReset Then Exit Sub
End Sub

'Private Sub Form_Activate()
'    mExitStatus = vbIgnore
'End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Dim i As Long
    Dim u As Long
    
    Dim dist As Single
    Dim valueControl As Control
    
    Set valueControl = GetValueControl(0)
    
    'lblKey(0).Move 120, 120
    
    'lblKey(0).Move 60, 60
    cmdFile(0).Move Me.ScaleWidth - 120 - cmdFile(0).Width, txtValue(0).Top
    cmdDir(0).Move cmdFile(0).Left, cmdFile(0).Top
    'txtValue(0).Move lblKey(0).Left, cmdFile(0).Top + cmdFile(0).Height + 120
        If cmdFile(0).Visible Or cmdDir(0).Visible Then
            valueControl.Width = cmdFile(0).Left - 2 * valueControl.Left
        Else
            valueControl.Width = Me.ScaleWidth - 2 * valueControl.Left
        End If

    'txtValueEx(0).Move txtValue(0).Left, txtValue(0).Top, txtValue(0).Width
    
    Dim lastValueControl As Control
    Set lastValueControl = valueControl
    For i = 1 To lblKey.UBound
        dist = lastValueControl.Top + lastValueControl.Height + 120 - lblKey(i - 1).Top
        Set valueControl = GetValueControl(i)
'        Debug.Print lastValueControl.Name
'        Debug.Print valueControl.Name
'
'        If txtValue(i - 1).Visible Then
'            dist = txtValue(i - 1).Top + txtValue(i - 1).Height + 120 - lblKey(i - 1).Top
'        Else
'            dist = txtValueEx(i - 1).Top + txtValueEx(i - 1).Height + 120 - lblKey(i - 1).Top
'        End If
        lblKey(i).Move lblKey(i - 1).Left, lblKey(i - 1).Top + dist
        valueControl.Move lastValueControl.Left, lastValueControl.Top + dist
        If ControlsVisible(cmdFile, i) Or ControlsVisible(cmdDir, i) Then
            valueControl.Width = cmdFile(0).Left - 2 * valueControl.Left
        Else
            valueControl.Width = Me.ScaleWidth - 2 * valueControl.Left
        End If
        
        cmdFile(i).Move cmdFile(0).Left, valueControl.Top
        cmdDir(i).Move cmdDir(0).Left, valueControl.Top
        
'        cmdFile(i).Move cmdFile(i - 1).Left, cmdFile(i - 1).Top + dist
'        cmdDir(i).Move cmdDir(i - 1).Left, cmdDir(i - 1).Top + dist
        
        Set lastValueControl = valueControl
    Next
    i = lblKey.UBound
    
    Dim txtBox As TextBox
    Set txtBox = GetTextValueControl(i)
    
    
    fraAddition.Top = txtBox.Top + txtBox.Height + 120
    fraAddition.Left = Me.ScaleWidth - 120 - fraAddition.Width
'
'    If txtValue(i).Visible Then
'            cmdCancel.Move Me.ScaleWidth - cmdCancel.Width - 120, txtValue(i).Top + txtValue(i).Height + 120
'        Else
'           cmdCancel.Move Me.ScaleWidth - cmdCancel.Width - 120, txtValueEx(i).Top + txtValueEx(i).Height + 120
'    End If
'
'    cmdOK.Move cmdCancel.Left - cmdOK.Width - 240, cmdCancel.Top
    
    
    
End Sub


Public Sub SetKeyStyle(ByRef vKey As String, Optional vStyle As ValueControlType = VCT_Normal)
    On Error Resume Next
    Dim i As Long
    i = SearchIndex(vKey)
    Dim value As String
    
    If i >= 0 Then
        value = GetValueControl(i).text
        Select Case vStyle
            Case ValueControlType.VCT_Combox
                txtValue(i).Visible = False
                txtValueEx(i).Visible = False
                Load cboValue(i)
                cboValue(i).Visible = True
                cboValue(i).text = value
            Case ValueControlType.VCT_MultiLine
                txtValue(i).Visible = False
                cboValue(i).Visible = False
                Load txtValueEx(i)
                txtValueEx(i).Visible = True
                txtValueEx(i).text = value
            Case ValueControlType.VCT_DIR
                Load cmdDir(i)
                cmdDir(i).Visible = True
                cmdFile(i).Visible = False
            Case ValueControlType.VCT_FILE
                Load cmdFile(i)
                cmdDir(i).Visible = False
                cmdFile(i).Visible = True
            Case ValueControlType.VCT_Normal
               ' Load txtValue(i)
                txtValueEx(i).Visible = False
                cboValue(i).Visible = False
                txtValue(i).Visible = True
                txtValue(i).text = value
        End Select
        
'        cmdFile(i).Visible = (vStyle = VCT_FILE)
'        cmdDir(i).Visible = (vStyle = VCT_DIR)
        
'        txtValueEx(i).Visible = (vStyle = VCT_MultiLine)
'        cboValue(i).Visible = (vStyle = VCT_Combox)
'        If vStyle = VCT_Normal Then txtValue(i).Visible = True
        Form_Resize
    End If
End Sub
Private Function GetValueControlByName(ByVal vName As String) As Control
    Dim idx As Long
    idx = SearchIndex("vName")
    If idx >= 0 Then
        Set GetValueControlByName = GetValueControl(idx)
    End If
End Function

Private Function GetValueControl(ByVal vIndex As Long) As Control
    'On Error Resume Next
    If ControlsVisible(cboValue, vIndex) Then
        Set GetValueControl = cboValue(vIndex)
    ElseIf ControlsVisible(txtValueEx, vIndex) Then
        Set GetValueControl = txtValueEx(vIndex)
    Else
        Set GetValueControl = txtValue(vIndex)
    End If
End Function

Private Function ControlsVisible(vControlArray As Variant, ByVal vIndex As Long) As Boolean
    If ControlsExists(vControlArray, vIndex) = False Then Exit Function
    ControlsVisible = vControlArray(vIndex).Visible
End Function

'CSEH: ErrExit
Private Function ControlsExists(vControlArray As Variant, ByVal vIndex As Long) As Boolean
    '<EhHeader>
    On Error GoTo ControlsExists_Err
    '</EhHeader>
        If vControlArray(vIndex).Name <> "" Then
            If Not Err Then ControlsExists = True
        End If
    Err.Clear
    '<EhFooter>
    Exit Function

ControlsExists_Err:
    ControlsExists = False
    Err.Clear

    '</EhFooter>
End Function

Private Sub CopyPosition( _
    ByRef vDest As Control, _
    ByRef vSource As Control, _
    Optional vLeft As Boolean = True, _
    Optional vTop As Boolean = True, _
    Optional vWidth As Boolean = True, _
    Optional vHeight As Boolean = True)
    If vDest Is Nothing Then Exit Sub
    If vSource Is Nothing Then Exit Sub
    On Error Resume Next
    If vLeft Then vDest.Left = vSource.Left
    If vTop Then vDest.Top = vSource.Top
    If vHeight Then vDest.Height = vSource.Height
    If vWidth Then vDest.Width = vSource.Width
End Sub



Public Sub Process(ByRef vKeyValues() As String)
    Dim i As Long
    
    On Error Resume Next
    
    cmdFile(0).Visible = False
    cmdDir(0).Visible = False
    cboValue(0).Visible = False
    txtValueEx(0).Visible = False
    txtValue(0).Visible = False
    
    For i = 1 To mKeyUBound
        Unload cmdFile(i)
        Unload cmdDir(i)
        Unload cboValue(i)
        Unload txtValueEx(i)
    Next
    
    mKeyValues = vKeyValues
    mKeyUBound = SafeUBound(mKeyValues())
    

    Dim i As Long
    For i = 0 To mKeyUBound
        Load lblKey(i)
        lblKey(i).Visible = True
        lblKey(i).Caption = mKeyValues(i, 0) & ":"
        Load txtValue(i)
        txtValue(i).Visible = True
        txtValue(i).text = mKeyValues(i, 1)
'        Load txtValueEx(i)
'        txtValueEx(i).Visible = False
'        txtValueEx(i).text = mKeyValues(i, 1)
'        Load cmdFile(i)
'        cmdFile(i).Visible = False
'        Load cmdDir(i)
'        cmdDir(i).Visible = False
    Next
    
    
    Form_Resize
End Sub

Public Sub SetField(ByRef vKey As String, ByRef vValue As String)
    Dim i As Long
    i = SearchIndex(vKey)
    If i >= 0 Then SetText i, vValue
End Sub

Public Function GetField(ByRef vKey As String) As String
    Dim i As Long
    i = SearchIndex(vKey)
    GetField = GetText(i)
End Function

Private Function GetText(ByVal Index As Long) As String
On Error GoTo ErrorGetText
    If Index > mKeyUBound Then Exit Function
    If Index < 0 Then Exit Function
    Dim txtBox As Control
    Set txtBox = GetValueControl(Index)
    GetText = txtBox.text

    Exit Function
ErrorGetText:
End Function

Private Sub SetText(ByVal Index As Long, ByRef vValue As String)
        '<EhHeader>
        On Error GoTo setText_Err
        '</EhHeader>
    If Index > mKeyUBound Then Exit Sub
    If Index < 0 Then Exit Sub
    Dim txtBox As Control
    Set txtBox = GetValueControl(Index)
    txtBox.text = vValue
        Exit Sub

setText_Err:

End Sub
Private Function SearchIndex(ByRef vKey As String) As Long
    Dim i As Long
    For i = 0 To mKeyUBound
        If mKeyValues(i, 0) = vKey Then SearchIndex = i: Exit Function
    Next
    SearchIndex = -1
End Function



'Public Sub SetMultiLine(ByRef vKey As String)
'    Dim i As Long
'    i = SearchIndex(vKey)
'    If i >= 0 Then
'        txtValueEx(i).Visible = True
'        txtValue(i).Visible = False
'        Form_Resize
'    End If
'End Sub
'
'Public Sub SetDirectory(ByRef vKey As String)
'    Dim i As Long
'    i = SearchIndex(vKey)
'    If i >= 0 Then
'        cmdFile(i).Visible = False
'        cmdDir(i).Visible = True
'        Form_Resize
'    End If
'End Sub
'
'Public Sub SetFile(ByRef vKey As String)
'    Dim i As Long
'    i = SearchIndex(vKey)
'    If i >= 0 Then
'        cmdFile(i).Visible = True
'        cmdDir(i).Visible = False
'        Form_Resize
'    End If
'End Sub
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
    Unload Me
End Sub

Public Function SelectFile(ByVal vFilename As String) As String
'TODO
End Function

Public Function SelectDir(ByVal vDirectory As String) As String
'TODO
End Function

Public Function ExitBeforeOK() As Boolean
'ToDo
End Function

Public Function ExitAfterOK() As Boolean
'ToDo
End Function

Public Function ExitBeforeCancel() As Boolean
'ToDo
End Function

Public Function ExitAfterCancel() As Boolean
'ToDo
End Function

Public Function ExitBeforeReset() As Boolean
'ToDo
End Function

Public Function ExitAfterReset() As Boolean
'ToDo
End Function
