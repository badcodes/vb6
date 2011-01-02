VERSION 5.00
Object = "*\ADropStack.vbp"
Begin VB.Form FSearch 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2640
   ClientLeft      =   2640
   ClientTop       =   2610
   ClientWidth     =   6585
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Search.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   6585
   Begin DropStack.XDropStack dropWith 
      Height          =   315
      Left            =   1680
      TabIndex        =   14
      Top             =   720
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin DropStack.XDropStack dropWhat 
      Height          =   315
      Left            =   1680
      TabIndex        =   13
      Top             =   240
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdReplace 
      Caption         =   "&Replace"
      Height          =   375
      Left            =   5295
      TabIndex        =   12
      Top             =   1215
      Width           =   1215
   End
   Begin VB.Frame frmMessage 
      Height          =   1200
      Left            =   204
      TabIndex        =   10
      Top             =   1116
      Width           =   1935
      Begin VB.Label lblMessage 
         Height          =   735
         Left            =   105
         TabIndex        =   11
         Top             =   315
         Width           =   1725
      End
   End
   Begin VB.CheckBox chkCase 
      Caption         =   "Match Ca&se"
      Height          =   300
      Left            =   2520
      TabIndex        =   5
      Top             =   1875
      Width           =   2400
   End
   Begin VB.CheckBox chkWord 
      Caption         =   "Find Whole Word &Only"
      Height          =   300
      Left            =   2508
      TabIndex        =   4
      Top             =   1545
      Width           =   2364
   End
   Begin VB.ComboBox cboDirection 
      Height          =   300
      ItemData        =   "Search.frx":0442
      Left            =   3645
      List            =   "Search.frx":044F
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1170
      Width           =   1215
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "&Help"
      Height          =   375
      Left            =   5280
      TabIndex        =   7
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton cmdAll 
      Caption         =   " Rep&lace All"
      Height          =   375
      Left            =   5280
      TabIndex        =   6
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5280
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton cmdFindNext 
      Caption         =   "Find &Next"
      Default         =   -1  'True
      Height          =   375
      Left            =   5280
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblDirection 
      Caption         =   "&Direction:"
      Height          =   300
      Left            =   2505
      TabIndex        =   9
      Top             =   1170
      Width           =   975
   End
   Begin VB.Label lblWith 
      Caption         =   "Replace &With: "
      Height          =   252
      Left            =   120
      TabIndex        =   0
      Top             =   708
      Width           =   1284
   End
   Begin VB.Label lblWhat 
      Caption         =   "&Find What:"
      Height          =   372
      Left            =   120
      TabIndex        =   8
      Top             =   240
      Width           =   1344
   End
End
Attribute VB_Name = "FSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RepId" ,"4FB51841-CEAF-11CF-A15E-00AA00A74D48-0073"


Option Explicit

Private fSetFocus As Integer
Private iFindStart As Integer
Private fBottomFile As Boolean, fNotFirst As Boolean
Private yControl2 As Double, yControl3 As Double
Private yControl4 As Double, yControl5 As Double, yButton5 As Double
Private fReplaceMode As Boolean
' Pointer for XEditor reference
Private pEditor As Long
Private fInCompleted As Boolean

Private ordReplaceStatus As Integer
Const ordFirst = 0
Const ordLast = 1
Const ordFail = 2

Private Sub cboDirection_Click()
    Editor.SearchOptionDirection = cboDirection.ListIndex
End Sub

Private Sub Form_Initialize()
    BugLocalMessage "FSearch Initialize"
End Sub

Private Sub Form_Load()
With Editor
    BugLocalMessage "FSearch Load"
    ' Initialize all control values from editor
    DrawForm
    .SearchActive = True
    chkWord.Value = -.SearchOptionWord
    chkCase.Value = -.SearchOptionCase
    cboDirection.ListIndex = .SearchOptionDirection
    fInCompleted = True
    Dim i As Long
    dropWhat.MaxCount = .FindWhatMax
    For i = .FindWhatCount To 1 Step -1
        dropWhat.Text = .FindWhat(i)
    Next
    dropWith.MaxCount = .ReplaceWithMax
    For i = .ReplaceWithCount To 1 Step -1
        dropWith.Text = .ReplaceWith(i)
    Next
    fInCompleted = False
    Show vbModeless
End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    BugLocalMessage "FSearch Unload"
    Editor.SearchActive = False
    Editor.SetFocus
End Sub

Private Sub Form_Terminate()
    BugLocalMessage "FSearch Terminate"
End Sub

Friend Property Get Editor() As XEditor
    Dim editorI As XEditor
    ' Turn editorI into an illegal, uncounted interface pointer
    CopyMemory editorI, pEditor, 4
    ' Do NOT hit the End button here! You will crash!
    ' Assign to legal reference (VB AddRefs it)
    Set Editor = editorI
    ' Still do NOT hit the End button! You will still crash!
    ' Destroy the illegal reference
    CopyMemory editorI, 0&, 4
    ' OK, hit the End button if you must
    ' Internal XEditor reference goes out of scope (VB Releases it)
End Property

Friend Property Set Editor(ByVal editorA As XEditor)
    ' Store an Editor pointer rather than an Editor object
    ' No AddRef for storing the pointer
    pEditor = ObjPtr(editorA)
End Property

Friend Property Get ReplaceMode() As Boolean
    ReplaceMode = fReplaceMode
End Property

Friend Property Let ReplaceMode(ByVal fReplaceModeA As Boolean)
    If fReplaceMode <> fReplaceModeA Then
        fReplaceMode = fReplaceModeA
    End If
End Property

Private Sub chkWord_Click()
    Static fInClick As Boolean
    If fInClick Then Exit Sub
    fInClick = True
    Editor.SearchOptionWord = Not Editor.SearchOptionWord
    fInClick = False
End Sub

Sub SearchChange(Kind As ESearchEvent)
With Editor
    Select Case Kind
    Case eseFindWhat
        If dropWhat.Text <> .FindWhat Then
            dropWhat.Text = .FindWhat
        End If
    Case eseReplaceWith
        If dropWith.Text <> .ReplaceWith Then
            dropWith.Text = .ReplaceWith
        End If
    Case eseCase
        chkCase.Value = -.SearchOptionCase
    Case eseWholeWord
        chkWord.Value = -.SearchOptionWord
    Case eseDirection
        cboDirection.ListIndex = .SearchOptionDirection
    End Select
End With
End Sub

Private Sub chkCase_Click()
    Static fInClick As Boolean
    If fInClick Then Exit Sub
    fInClick = True
    Editor.SearchOptionCase = Not Editor.SearchOptionCase
    fInClick = False
End Sub

Private Sub cmdAll_Click()
    With Editor

        ' Save position and start search from first
        Dim iPos As Integer, s As String

        ' Find first match, but don't highlight
        iPos = .FindNext(dropWhat)
        Do While iPos
            iPos = .ReplaceNext(dropWhat, dropWith)
        Loop

    End With
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdFindNext_Click()
With Editor
    Dim i As Integer
    ' Must be something to find
    If dropWhat.Text = sEmpty Then
        dropWhat.SetFocus
        Exit Sub
    End If
    ' When Replace user selects Next once, make it default
    cmdFindNext.Default = True
    ' Find next item
    i = .FindNext(dropWhat)
    ' Deal with failed search
    If i = 0 Then
        lblMessage.Caption = "Text not found"
        .SelLength = 0
    Else
        lblMessage.Caption = "Text found: " & .Line & "," & .Column
    End If
    dropWhat.SetFocus
End With
End Sub

Private Sub cmdReplace_Click()
    With Editor

        ' Turn Find dialog into Replace dialog
        If Not fReplaceMode Then
            fReplaceMode = True
            DrawForm
            If dropWhat.Text = sEmpty Then
                dropWhat.SetFocus
            Else
                dropWith.SetFocus
            End If
            Exit Sub
        End If

        ' Must be something to replace
        If dropWhat.Text = sEmpty Then
            dropWhat.SetFocus
            Exit Sub
        End If
        
        ' If you want to ask user to confirm empty replace string, do it here

        ' When user selects Replace once, make it default
        cmdReplace.Default = True

        If .ReplaceNext(dropWhat, dropWith) Then
            Call .FindNext(dropWhat)
        Else
            lblMessage.Caption = "Text not found"
            dropWhat.SetFocus
            Exit Sub
        End If

    End With
End Sub

Private Sub ReplaceText(sSrc As String, ByVal sReplace As String, iStart As Integer, iLen As Integer)
    sSrc = Left$(sSrc, iStart - 1) & sReplace & Mid$(sSrc, iStart + iLen)
End Sub

Private Sub dropWhat_Completed(Text As String)
    If fInCompleted Then Exit Sub
    fInCompleted = True
    lblMessage.Caption = sEmpty
    Editor.FindWhat = Text
    fInCompleted = False
End Sub

Private Sub dropWhat_Change()
    lblMessage.Caption = sEmpty
End Sub

Private Sub dropWith_Completed(Text As String)
    If fInCompleted Then Exit Sub
    fInCompleted = True
    lblMessage.Caption = sEmpty
    Editor.ReplaceWith = Text
    fInCompleted = False
End Sub

Private Sub dropWith_Change()
    lblMessage.Caption = Empty
End Sub

Private Sub DrawForm()

    ' Get initial button and control positions for later placement
    If fNotFirst = False Then
        fNotFirst = True
        yControl2 = dropWith.Top
        yControl3 = cboDirection.Top
        yControl4 = chkWord.Top
        yControl5 = chkCase.Top
        yButton5 = cmdHelp.Top
    End If

    ' Modify buttons and controls for current mode
    If fReplaceMode = False Then
        cmdFindNext.Caption = "&Next"
        cmdFindNext.Default = True
        cmdCancel.Caption = "&Close"
        cmdReplace.Caption = "&Replace..."
        Caption = "Find"
        cmdAll.Visible = False
        lblWith.Visible = False
        dropWith.Visible = False
        cmdHelp.Top = cmdAll.Top
        frmMessage.Top = yControl2
        cboDirection.Top = yControl2
        lblDirection.Top = yControl2
        chkWord.Top = yControl3
        chkCase.Top = yControl4
    Else ' Replace
        cmdCancel.Caption = "&Close"
        cmdReplace.Default = True
        cmdFindNext.Caption = "Find &Next"
        cmdReplace.Caption = "&Replace"
        Caption = "Replace"
        cmdAll.Visible = True
        lblWith.Visible = True
        dropWith.Visible = True
        cmdHelp.Top = yButton5
        frmMessage.Top = yControl3
        cboDirection.Top = yControl3
        lblDirection.Top = yControl3
        chkWord.Top = yControl4
        chkCase.Top = yControl5
    End If

    Dim dyTitleBar As Double, dyBtnLow As Double, dyBorder As Double
    ' Calculate height of title bar
    dyTitleBar = Height - ScaleHeight
    ' Add height of lowest element (help button moves up and down)
    dyBtnLow = cmdHelp.Top + cmdHelp.Height
    ' Add border around closest element (top button)
    dyBorder = cmdFindNext.Top
    ' Set height
    Height = dyTitleBar + dyBtnLow + dyBorder

End Sub

