VERSION 5.00
Begin VB.Form FTestEnum 
   Caption         =   "Test Enumerations"
   ClientHeight    =   3480
   ClientLeft      =   1992
   ClientTop       =   1416
   ClientWidth     =   5100
   Icon            =   "TestEnum.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3480
   ScaleWidth      =   5100
   Begin VB.CheckBox chkCase 
      Caption         =   "Case Sensitive"
      Height          =   312
      Left            =   144
      TabIndex        =   14
      Top             =   1176
      Width           =   1452
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "Find"
      Height          =   360
      Left            =   156
      TabIndex        =   12
      Top             =   2940
      Width           =   984
   End
   Begin VB.TextBox txtClass 
      Height          =   348
      Left            =   1200
      TabIndex        =   9
      Top             =   720
      Width           =   948
   End
   Begin VB.TextBox txtTitle 
      Height          =   348
      Left            =   144
      TabIndex        =   8
      Top             =   720
      Width           =   948
   End
   Begin VB.Frame fmEnum 
      Caption         =   "Enumerations"
      Height          =   1332
      Left            =   2388
      TabIndex        =   3
      Top             =   96
      Width           =   2388
      Begin VB.OptionButton opt 
         Caption         =   "Top Windows"
         Height          =   250
         Index           =   0
         Left            =   84
         TabIndex        =   7
         Top             =   252
         Width           =   1308
      End
      Begin VB.OptionButton opt 
         Caption         =   "Child Windows "
         Height          =   250
         Index           =   1
         Left            =   84
         TabIndex        =   6
         Top             =   492
         Width           =   1428
      End
      Begin VB.OptionButton opt 
         Caption         =   "Font Families"
         Height          =   250
         Index           =   2
         Left            =   84
         TabIndex        =   5
         Top             =   732
         Width           =   1356
      End
      Begin VB.OptionButton opt 
         Caption         =   "Resources"
         Height          =   250
         Index           =   3
         Left            =   84
         TabIndex        =   4
         Top             =   972
         Width           =   1428
      End
   End
   Begin VB.ListBox lstEnum 
      Height          =   1008
      Left            =   2388
      TabIndex        =   1
      Top             =   1584
      Width           =   2436
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   372
      Left            =   1260
      MaskColor       =   &H00000000&
      TabIndex        =   0
      Top             =   2940
      Width           =   972
   End
   Begin VB.Label lbl 
      Caption         =   "Find Any Window"
      Height          =   360
      Index           =   0
      Left            =   168
      TabIndex        =   15
      Top             =   120
      Width           =   2088
   End
   Begin VB.Label lblFound 
      Height          =   1176
      Left            =   144
      TabIndex        =   13
      Top             =   1632
      Width           =   2004
   End
   Begin VB.Label lbl 
      Caption         =   "Class:"
      Height          =   216
      Index           =   2
      Left            =   1188
      TabIndex        =   11
      Top             =   468
      Width           =   972
   End
   Begin VB.Label lbl 
      Caption         =   "Title:"
      Height          =   216
      Index           =   1
      Left            =   144
      TabIndex        =   10
      Top             =   480
      Width           =   972
   End
   Begin VB.Label lblResult 
      Height          =   216
      Left            =   2388
      TabIndex        =   2
      Top             =   3060
      Width           =   2424
   End
End
Attribute VB_Name = "FTestEnum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  
Enum EEnumerations
    eeTopWindows = 0
    eeChildWindows
    eeFontFamilies
    eeResources
End Enum

Private eeCur As EEnumerations

Private Sub Form_Load()
    Dim ab() As Byte, s As String
    s = "ABC"
    Set lstEnumRef = lstEnum
End Sub

Private Sub Form_Activate()
    opt(0) = True
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    Dim hWnd As Long, sClass As String, sTitle As String
    sTitle = txtTitle
    sClass = txtClass
    If sTitle = sEmpty And sClass = sEmpty Then Exit Sub
    ' Hide text so we won't find entry textbox
    txtTitle = sEmpty
    txtClass = sEmpty
    hWnd = FindAnyWindow(sTitle, sClass, chkCase)
    txtTitle = sTitle
    txtClass = sClass
    If hWnd Then
        lblFound = GetWndInfo(hWnd)
    Else
        lblFound = "No match"
    End If
End Sub

Private Sub lstEnum_Click()
    If eeCur = eeChildWindows Or eeCur = eeTopWindows Then
        lblFound = GetWndInfo(lstEnum.ItemData(lstEnum.ListIndex))
    End If
End Sub

Private Sub opt_Click(Index As Integer)
    Dim f As Long, c As Long
    eeCur = Index
    lstEnum.Clear
    Select Case eeCur
    Case eeTopWindows
        f = EnumWindows(AddressOf EnumWndProc, c)
        lblResult = "Window count: " & c
    Case eeChildWindows
        f = EnumChildWindows(GetDesktopWindow, AddressOf EnumWndProc, c)
        lblResult = "Window count: " & c
    Case eeFontFamilies
        f = EnumFontFamilies(hDC, sNullStr, AddressOf EnumFontFamProc, c)
        lblResult = "Font family count: " & c
    Case eeResources
        'Set nResType = New Collection
        f = EnumResourceTypes(App.hInstance, AddressOf EnumResTypeProc, c)
        lblResult = "Resource count: " & c
    End Select
End Sub

Private Sub txtTitle_GotFocus()
    cmdFind.Default = True
End Sub

Private Sub txtClass_GotFocus()
    cmdFind.Default = True
End Sub

