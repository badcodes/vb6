VERSION 5.00
Begin VB.Form FCollectionWizard 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Collection Wizard"
   ClientHeight    =   6324
   ClientLeft      =   9060
   ClientTop       =   1548
   ClientWidth     =   6612
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   7.8
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "CollWiz.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6324
   ScaleWidth      =   6612
   Begin VB.ComboBox cboView 
      Height          =   315
      ItemData        =   "CollWiz.frx":0CFA
      Left            =   4200
      List            =   "CollWiz.frx":0D04
      Style           =   2  'Dropdown List
      TabIndex        =   24
      Top             =   1320
      Width           =   2295
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "View Class"
      Default         =   -1  'True
      Height          =   375
      Left            =   4200
      TabIndex        =   22
      Top             =   120
      Width           =   2292
   End
   Begin VB.CheckBox chkWalkPublic 
      Caption         =   "Make walker public"
      Height          =   252
      Left            =   4200
      TabIndex        =   20
      Top             =   1920
      Width           =   2172
   End
   Begin VB.OptionButton optContainer 
      Caption         =   "Object container"
      Height          =   252
      Index           =   2
      Left            =   4200
      TabIndex        =   19
      Top             =   2760
      Width           =   2172
   End
   Begin VB.OptionButton optContainer 
      Caption         =   "Variable container"
      Height          =   252
      Index           =   1
      Left            =   4200
      TabIndex        =   18
      Top             =   2520
      Width           =   2172
   End
   Begin VB.OptionButton optContainer 
      Caption         =   "Generic container"
      Height          =   252
      Index           =   0
      Left            =   4200
      TabIndex        =   17
      Top             =   2280
      Value           =   -1  'True
      Width           =   2172
   End
   Begin VB.CheckBox chkCollPublic 
      Caption         =   "Make collection public"
      Height          =   252
      Left            =   4200
      TabIndex        =   16
      Top             =   1680
      Value           =   1  'Checked
      Width           =   2295
   End
   Begin VB.TextBox txtBase 
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   480
      Width           =   1932
   End
   Begin VB.TextBox txtWalkFile 
      Height          =   375
      Left            =   2160
      TabIndex        =   10
      Top             =   2685
      Width           =   1932
   End
   Begin VB.TextBox txtWalkVar 
      Height          =   375
      Left            =   2160
      TabIndex        =   9
      Top             =   1965
      Width           =   1932
   End
   Begin VB.TextBox txtWalkClass 
      Height          =   375
      Left            =   2160
      TabIndex        =   8
      Top             =   1245
      Width           =   1932
   End
   Begin VB.TextBox txtCollClass 
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1245
      Width           =   1932
   End
   Begin VB.TextBox txtCollVar 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1965
      Width           =   1932
   End
   Begin VB.TextBox txtCollFile 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   2685
      Width           =   1932
   End
   Begin VB.CommandButton cmdFile 
      Caption         =   "&Save Class Files"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4200
      TabIndex        =   1
      Top             =   600
      Width           =   2292
   End
   Begin VB.TextBox txtCode 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2988
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   3120
      Width           =   6372
   End
   Begin VB.Label lblStatus 
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   6120
      Width           =   5415
   End
   Begin VB.Label lbl 
      Caption         =   "View:"
      Height          =   255
      Index           =   7
      Left            =   4200
      TabIndex        =   21
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label lbl 
      Caption         =   "Base name (blob for CBlobs collection)"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   15
      Top             =   240
      Width           =   3855
   End
   Begin VB.Label lbl 
      Caption         =   "Walker filename:"
      Height          =   255
      Index           =   5
      Left            =   2160
      TabIndex        =   13
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label lbl 
      Caption         =   "Walker variable:"
      Height          =   255
      Index           =   4
      Left            =   2160
      TabIndex        =   12
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label lbl 
      Caption         =   "Walker class name:"
      Height          =   255
      Index           =   0
      Left            =   2160
      TabIndex        =   11
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label lbl 
      Caption         =   "Collection class name:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label lbl 
      Caption         =   "Collection variable:"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   2295
   End
   Begin VB.Label lbl 
      Caption         =   "Collection filename:"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   5
      Top             =   2400
      Width           =   1815
   End
End
Attribute VB_Name = "FCollectionWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Enum EClassType
    ectCollection
    ectWalker
End Enum

Enum EContainer
    ecGeneric
    ecVariable
    ecObject
End Enum

Private Sub cboView_Click()
    CreateView
End Sub

Private Sub cmdView_Click()
    Dim s As String
    If cmdView.Caption = "View Class" Then
        If txtBase <> sEmpty Then
            cmdView.Caption = "Clear All"
            cmdFile.Enabled = True
            s = Left$(txtBase, 1)
            s = UCase$(s) & Mid$(txtBase, 2)
            If txtCollClass = sEmpty Then
                txtCollClass = "C" & s & "s"
            End If
            If txtCollVar = sEmpty Then
                txtCollVar = txtBase & "s"
            End If
            If txtCollFile = sEmpty Then
                txtCollFile = UCase$(s) & "S" & ".CLS"
            End If
            If txtWalkClass = sEmpty Then
                txtWalkClass = "C" & s & "Walker"
            End If
            If txtWalkVar = sEmpty Then
                txtWalkVar = txtBase
            End If
            If txtWalkFile = sEmpty Then
                txtWalkFile = UCase$(s) & "WLK" & ".CLS"
            End If
        End If
        CreateView
    Else
        cmdView.Caption = "View Class"
        cmdFile.Enabled = False
        txtBase = sEmpty
        txtCollClass = sEmpty
        txtCollVar = sEmpty
        txtCollFile = sEmpty
        txtWalkClass = sEmpty
        txtWalkVar = sEmpty
        txtWalkFile = sEmpty
        txtCode = sEmpty
        lblStatus = sEmpty
    End If
End Sub

Private Sub Form_Load()
    ChDir CurDir$
    Show
    cboView.ListIndex = ectCollection
    txtBase.SetFocus
End Sub

Private Sub chkCollPublic_Click()
    If chkCollPublic Then
        chkWalkPublic.Enabled = True
    Else
        chkWalkPublic = False
        chkWalkPublic.Enabled = False
    End If
    CreateView
End Sub

Private Sub chkWalkPublic_Click()
    CreateView
End Sub

Private Sub txtCollClass_LostFocus()
    CreateView
End Sub

Private Sub txtCollFile_LostFocus()
    CreateView
End Sub

Private Sub txtCollVar_LostFocus()
    CreateView
End Sub

Private Sub txtWalkClass_LostFocus()
    CreateView
End Sub

Private Sub txtWalkFile_LostFocus()
    CreateView
End Sub

Private Sub txtWalkVar_LostFocus()
    CreateView
End Sub

Private Sub optContainer_Click(Index As Integer)
    CreateView
End Sub

Sub CreateView()
    If txtBase = sEmpty Then
        lblStatus = "Must give base name"
        txtBase.SetFocus
        Exit Sub
    End If
    lblStatus = sEmpty
    If cboView.ListIndex = ectCollection Then
        txtCode.Text = MakeCollection(txtCollClass, txtWalkClass, _
                                      txtWalkVar)
    Else
        txtCode.Text = MakeWalker(txtWalkClass, txtCollClass, _
                                  txtCollVar)
    End If
End Sub

Private Sub cmdFile_Click()
    Dim result As VbMsgBoxResult, sCollBack As String, sWalkBack As String
    If txtCode = sEmpty Then CreateView
    If txtCollFile = sEmpty Then
        lblStatus.Caption = "Must have collection filename"
        Exit Sub
    ElseIf txtWalkFile = sEmpty Then
        lblStatus.Caption = "Must have walker filename"
        Exit Sub
    End If
    
    If ExistFile(txtCollFile) Then
        result = MsgBox("Files exists: " & txtCollFile & ". Make backup? ", vbYesNoCancel)
        Select Case result
        Case vbYes
            On Error Resume Next
            sCollBack = txtCollFile
            Mid$(sCollBack, Len(sCollBack)) = "K"
            If ExistFile(sCollBack) Then Kill sCollBack
            Name txtCollFile As sCollBack
            On Error GoTo 0
        Case vbNo
            ' Fall through
        Case vbCancel
            Exit Sub
        End Select
    End If
    SaveFileStr txtCollFile, MakeCollection(txtCollClass, txtWalkClass, txtWalkVar)
    
    If ExistFile(txtWalkFile) Then
        result = MsgBox("Files exists: " & txtWalkFile & ". Make backup? ", vbYesNoCancel)
        Select Case result
        Case vbYes
            On Error Resume Next
            sWalkBack = txtWalkFile
            Mid$(sWalkBack, Len(sWalkBack)) = "K"
            If ExistFile(sWalkBack) Then Kill sWalkBack
            Name txtWalkFile As sWalkBack
            On Error GoTo 0
        Case vbNo
            ' Fall through
        Case vbCancel
            Exit Sub
        End Select
    End If
    SaveFileStr txtWalkFile, MakeWalker(txtWalkClass, txtCollClass, txtCollVar)
End Sub

Private Function MakeCollection(sCollClass As String, sWalkClass As String, sWalkVar As String) As String

    Dim s As String, sCollName  As String
    sCollName = Mid$(sCollClass, 2)
    s = s & "VERSION 1.0 CLASS" & sCrLf & _
            "BEGIN" & sCrLf & _
            "  MultiUse = -1  'True" & sCrLf & _
            "END" & sCrLf & _
            "Attribute VB_Name = " & sQuote2 & sCollClass & sQuote2 & sCrLf & _
            "Attribute VB_GlobalNameSpace = False" & sCrLf & _
            "Attribute VB_Creatable = True" & sCrLf & _
            "Attribute VB_PredeclaredId = False" & sCrLf & _
            "Attribute VB_Exposed = " & IIf(chkCollPublic, "True", "False") & sCrLf & _
            "Option Explicit" & sCrLf & sCrLf

    s = s & "' Private data structure" & sCrLf & _
            "'!Private data() As DataType" & sCrLf & sCrLf

    s = s & "Private Sub Class_Initialize()" & sCrLf & _
            "    ' Initialize private data" & sCrLf & _
            "    '!data = initval" & sCrLf & _
            "End Sub" & sCrLf & sCrLf

    s = s & "' Friend properties to make data structure accessible to walker" & sCrLf & _
            "'!Friend Property Get " & sCollName & "(i As Long) '! As DataType" & sCrLf & _
            "'!    " & sCollName & " = data(i)" & sCrLf & _
            "'!End Property" & sCrLf & sCrLf
    
    s = s & "' NewEnum must have the procedure ID -4 in Procedure Attributes dialog" & sCrLf & _
            "' Create a new data walker object and connect to it" & sCrLf & _
            "Public Function NewEnum() As IEnumVARIANT" & sCrLf & _
            "Attribute NewEnum.VB_UserMemId = -4" & sCrLf & _
            "    ' Create a new iterator object" & sCrLf & _
            "    Dim " & sWalkVar & "walker As " & sWalkClass & sCrLf & _
            "    Set " & sWalkVar & "walker = New " & sWalkClass & sCrLf & _
            "    ' Connect it with collection data" & sCrLf & _
            "    " & sWalkVar & "walker.Attach Me" & sCrLf & _
            "    ' Return it" & sCrLf & _
            "    Set NewEnum = " & sWalkVar & "walker.NewEnum" & sCrLf & _
            "End Function" & sCrLf & sCrLf

    s = s & "Public Property Get Count() As Integer" & sCrLf & _
            "    '!Count = curcount" & sCrLf & _
            "End Property" & sCrLf & sCrLf
            
    s = s & "' Default property" & sCrLf & _
            "'!Public Property Get Item(vIndex As Variant) '! As DataType" & sCrLf & _
            "Attribute Item.VB_UserMemId = 0" & sCrLf
            
    Select Case GetOption(optContainer)
    Case ecVariable
    
        s = s & "    '!Item = data(vIndex)" & sCrLf & _
                "'!End Property" & sCrLf & sCrLf

    Case ecObject
    
        s = s & "    '!Set Item = data(vIndex)" & sCrLf & _
                "'!End Property" & sCrLf & sCrLf

    Case ecGeneric
    
        s = s & "    ' Generic containers must check and handle objects" & sCrLf & _
                "    '!If IsObject(data(vIndex)) Then" & sCrLf & _
                "    '!    Set Item = data(vIndex)" & sCrLf & _
                "    '!Else" & sCrLf & _
                "    '!    Item = data(vIndex)" & sCrLf & _
                "    '!End If" & sCrLf & _
                "'!End Property" & sCrLf & sCrLf

        s = s & "' Let and Set generally only required for generic containers" & sCrLf & _
                "Property Let Item(vIndex As Variant, curdataA) '! As DataType)" & sCrLf & _
                "    '!data(vIndex) = curdataA" & sCrLf & _
                "End Property" & sCrLf & sCrLf
    
        s = s & "Property Set Item(vIndex As Variant, curdataA) '! As DataType)" & sCrLf & _
                "    '!Set data(vIndex) = curdataA" & sCrLf & _
                "End Property" & sCrLf & sCrLf
            
    End Select

    s = s & "' Add other collection members such as Add and Remove" & sCrLf & sCrLf
    
    MakeCollection = s
    
End Function

Private Function MakeWalker(sWalkClass As String, sCollClass As String, _
                            sCollVar As String) As String

    Dim s As String, sCollName As String
    sCollName = Mid$(sCollClass, 2)
    ' Note that VB6 uses more properties here, but will add them automatically
    s = s & "VERSION 1.0 CLASS" & sCrLf & _
            "BEGIN" & sCrLf & _
            "  MultiUse = -1  'True" & sCrLf & _
            "END" & sCrLf & _
            "Attribute VB_Name = " & sQuote2 & sWalkClass & sQuote2 & sCrLf & _
            "Attribute VB_GlobalNameSpace = False" & sCrLf & _
            "Attribute VB_Creatable = True" & sCrLf & _
            "Attribute VB_PredeclaredId = False" & sCrLf & _
            "Attribute VB_Exposed = " & IIf(chkWalkPublic, "True", "False") & sCrLf & _
            "Option Explicit" & sCrLf & sCrLf
           
    s = s & "' Implement Basic-friendly version of IEnumVARIANT" & sCrLf & _
            "Implements IVariantWalker" & sCrLf & _
            "' Connect back to parent collection" & sCrLf & _
            "Private connect As " & sCollClass & sCrLf & sCrLf

    s = s & "' Private state data" & sCrLf & _
            "'!Private iCur As Long" & sCrLf & sCrLf

    s = s & "Private Sub Class_Initialize()" & sCrLf & _
            "    ' Initialize position in collection" & sCrLf & _
            "    '!iCur = 0" & sCrLf & _
            "End Sub" & sCrLf & sCrLf

    s = s & "' Receive connection from " & sCollClass & sCrLf & _
            "Sub Attach(connectA As " & sCollClass & ")" & sCrLf & _
            "    Set connect = connectA" & sCrLf & _
            "End Sub" & sCrLf & sCrLf

    ' Fixed based on reports from Kai Wagner and Elliot Witticar
    s = s & "' Return IEnumVARIANT (indirectly) to client collection" & sCrLf & _
            "Friend Function NewEnum() As stdole.IEnumVARIANT" & sCrLf & _
            "    ' Delegate to class that implements real IEnumVARIANT" & sCrLf & _
            "    Dim vars As CEnumVariant" & sCrLf & _
            "    ' Connect walker to CEnumVariant so it can call methods" & sCrLf & _
            "    Set vars = New CEnumVariant" & sCrLf & _
            "    vars.Attach Me" & sCrLf & _
            "    ' Return walker to collection data" & sCrLf & _
            "    Set NewEnum = vars" & sCrLf & _
            "End Function" & sCrLf & sCrLf

    s = s & "' Implement IVariantWalker methods" & sCrLf & _
            "Private Function IVariantWalker_More(v As Variant) As Boolean" & sCrLf & _
            "    ' Move to next element" & sCrLf & _
            "    '!iCur = iCur + 1" & sCrLf & _
            "    ' If more data, return True and update data" & sCrLf & _
            "    '!If iCur <= connect.Count Then" & sCrLf & _
            "        '!IVariantWalker_More = True" & sCrLf

    Select Case GetOption(optContainer)
    Case ecVariable
    
        s = s & "        '!v = connect." & sCollName & "(iCur)" & sCrLf

    Case ecObject
    
        s = s & "        '!Set v = connect." & sCollName & "(iCur)" & sCrLf

    Case ecGeneric
    
        s = s & "        '!If IsObject(connect." & sCollName & "(iCur)) Then" & sCrLf & _
                "        '!    Set v = connect." & sCollName & "(iCur)" & sCrLf & _
                "        '!Else" & sCrLf & _
                "        '!    v = connect." & sCollName & "(iCur)" & sCrLf & _
                "        '!End If" & sCrLf
    End Select
    
    s = s & "    '!End If" & sCrLf & _
            "End Function" & sCrLf & sCrLf

    s = s & "Private Sub IVariantWalker_Reset()" & sCrLf & _
            "    ' Move to first element" & sCrLf & _
            "    '!iCur = 0" & sCrLf & _
            "End Sub" & sCrLf & sCrLf

    s = s & "Private Sub IVariantWalker_Skip(c as Long)" & sCrLf & _
            "    ' Skip a given number of elements" & sCrLf & _
            "    '!iCur = iCur + c" & sCrLf & _
            "End Sub" & sCrLf & sCrLf
            
    MakeWalker = s
    
End Function

