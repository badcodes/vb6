VERSION 5.00
Object = "{CFC13920-9EF4-11D0-B72F-0000C04D4C0A}#6.0#0"; "mswless.ocx"
Begin VB.Form frmAddIn 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "VBAlias"
   ClientHeight    =   7515
   ClientLeft      =   2175
   ClientTop       =   1935
   ClientWidth     =   9615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7515
   ScaleWidth      =   9615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VBAlias.MlcGrid MlcGrid1 
      Height          =   5535
      Left            =   240
      TabIndex        =   4
      Top             =   720
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   9763
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Cols            =   2
      Row             =   0
      Col             =   0
   End
   Begin VB.Label Label4 
      Caption         =   "Long String:"
      Height          =   255
      Left            =   4320
      TabIndex        =   21
      Top             =   1320
      Width           =   4095
   End
   Begin MSWLess.WLCheck Check2 
      Height          =   255
      Left            =   240
      TabIndex        =   20
      Top             =   7080
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   450
      _Version        =   393216
      Caption         =   "Space Enhancer"
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
   End
   Begin MSWLess.WLCheck Check1 
      Height          =   255
      Left            =   2160
      TabIndex        =   19
      Top             =   7080
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   450
      _Version        =   393216
      Caption         =   "Maximise on startup"
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
   End
   Begin MSWLess.WLFrame WLFrame2 
      Height          =   735
      Left            =   120
      TabIndex        =   18
      Top             =   6720
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   1296
      _Version        =   393216
      Caption         =   "Other Functions"
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      BorderStyle     =   1
      Appearance      =   1
   End
   Begin MSWLess.WLText WLText1 
      Height          =   2535
      Left            =   4320
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   1560
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   4471
      _Version        =   393216
      MultiLine       =   -1  'True
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
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
   Begin VB.Label Label3 
      Caption         =   "Caret Position:"
      Height          =   255
      Left            =   7680
      TabIndex        =   16
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   7680
      TabIndex        =   15
      Top             =   720
      Width           =   1455
   End
   Begin MSWLess.WLCheck WLCheck1 
      Height          =   315
      Left            =   4320
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   6000
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   556
      _Version        =   393216
      Caption         =   "The Macros are Enabled"
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
   End
   Begin MSWLess.WLCommand WLCommand6 
      Height          =   375
      Left            =   7800
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   5880
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Caption         =   "&Print List"
      ForeColor       =   -1
      BackColor       =   -2147483633
      Appearance      =   1
   End
   Begin MSWLess.WLCommand WLCommand5 
      Height          =   375
      Left            =   7800
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   5280
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Caption         =   "&Remove Entry"
      ForeColor       =   -1
      BackColor       =   -2147483633
      Appearance      =   1
   End
   Begin MSWLess.WLCommand WLCommand4 
      Height          =   375
      Left            =   7800
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   4680
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Caption         =   "Add &Entry"
      ForeColor       =   -1
      BackColor       =   -2147483633
      Appearance      =   1
   End
   Begin VB.Label Label6 
      Caption         =   "Programmer Name:"
      Height          =   255
      Left            =   4320
      TabIndex        =   10
      Top             =   5280
      Width           =   2655
   End
   Begin MSWLess.WLText WLText3 
      Height          =   315
      Left            =   4320
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   5520
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   556
      _Version        =   393216
      Text            =   "Programmer Name"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
   End
   Begin VB.Label Label5 
      Caption         =   "For embedding in the Long String:"
      Height          =   255
      Left            =   4320
      TabIndex        =   8
      Top             =   4440
      Width           =   2655
   End
   Begin MSWLess.WLCombo WLCombo1 
      Height          =   315
      Left            =   4320
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   4680
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   556
      _Version        =   393216
      ListCount       =   13844
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      List            =   "frmAddIn.frx":0000
   End
   Begin VB.Label Label2 
      Caption         =   "Short String:"
      Height          =   255
      Left            =   4320
      TabIndex        =   6
      Top             =   480
      Width           =   2655
   End
   Begin MSWLess.WLText WLText2 
      Height          =   315
      Left            =   4320
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   720
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   556
      _Version        =   393216
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
   End
   Begin MSWLess.WLFrame WLFrame1 
      Height          =   6495
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   11456
      _Version        =   393216
      Caption         =   "Macros"
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      BorderStyle     =   1
      Appearance      =   1
   End
   Begin MSWLess.WLCommand WLCommand1 
      Height          =   375
      Left            =   4440
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   6960
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Caption         =   "&Ok"
      ForeColor       =   -1
      BackColor       =   -2147483633
      Appearance      =   1
   End
   Begin MSWLess.WLCommand WLCommand2 
      Height          =   375
      Left            =   6240
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   6960
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Caption         =   "&Close"
      ForeColor       =   -1
      BackColor       =   -2147483633
      Appearance      =   1
   End
   Begin MSWLess.WLCommand WLCommand3 
      Height          =   375
      Left            =   8040
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   6960
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Caption         =   "&About"
      ForeColor       =   -1
      BackColor       =   -2147483633
      Appearance      =   1
   End
End
Attribute VB_Name = "frmAddIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public VBInstance As VBIDE.VBE
Public IDEExt As Connect

'macros
Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type


Private m_tP As POINTAPI
Private mWLText1SetStart As Long
Private mFastMacroTxt As String

'no close
Private Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Const MF_BYPOSITION = &H400&
Public ReadyToClose As Boolean

'on top
Private Declare Function SetWindowPos Lib "user32" _
        (ByVal hwnd As Long, _
        ByVal hWndInsertAfter As Long, _
        ByVal X As Long, ByVal Y As Long, _
        ByVal cx As Long, ByVal cy As Long, _
        ByVal wFlags As Long) As Long
Private Const SWP_NOMOVE = 2
Private Const SWP_NOSIZE = 1
Private Const SWP_WNDFLAGS = SWP_NOMOVE Or SWP_NOSIZE
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2

'get window text
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Function WindowTextFromWnd(ByVal hwnd As Long) As String
Dim c As Long, s As String
c = GetWindowTextLength(hwnd)
' Some windows return huge length--ignore 0 <= c < 4096
If c And &HFFFFF000 Then Exit Function
s = String$(c, 0)
c = GetWindowText(hwnd, s, c + 1)
WindowTextFromWnd = s
End Function

Public Sub MaximiseOnStartup()
'depending how are the the settings, maximise on startup
If Check1.Value = wlChecked Then
    Dim oCodePane As Window

    For Each oCodePane In VBInstance.ActiveVBProject.VBE.Windows
        If oCodePane.Type = vbext_wt_CodeWindow Or _
                oCodePane.Type = vbext_wt_Designer Then

            If Not oCodePane.WindowState = vbext_ws_Maximize Then
                oCodePane.WindowState = vbext_ws_Maximize
                Exit For
            Else
                Exit For
            End If
        End If
    Next oCodePane
End If
End Sub
Public Sub SetTopmost(frm As Form, bTopmost As Boolean)
Dim i As Long
If bTopmost = True Then
    i = SetWindowPos(frm.hwnd, HWND_TOPMOST, _
            0, 0, 0, 0, SWP_WNDFLAGS)
Else
    i = SetWindowPos(frm.hwnd, HWND_NOTOPMOST, _
            0, 0, 0, 0, SWP_WNDFLAGS)
End If
End Sub
Private Sub RemoveMenus(frm As Form, _
        remove_restore As Boolean, _
        remove_move As Boolean, _
        remove_size As Boolean, _
        remove_minimize As Boolean, _
        remove_maximize As Boolean, _
        remove_seperator As Boolean, _
        remove_close As Boolean)
Dim hMenu As Long

' Get the form's system menu handle.
hMenu = GetSystemMenu(frm.hwnd, False)

If remove_close Then DeleteMenu hMenu, 6, MF_BYPOSITION
If remove_seperator Then DeleteMenu hMenu, 5, MF_BYPOSITION
If remove_maximize Then DeleteMenu hMenu, 4, MF_BYPOSITION
If remove_minimize Then DeleteMenu hMenu, 3, MF_BYPOSITION
If remove_size Then DeleteMenu hMenu, 2, MF_BYPOSITION
If remove_move Then DeleteMenu hMenu, 1, MF_BYPOSITION
If remove_restore Then DeleteMenu hMenu, 0, MF_BYPOSITION
End Sub
Public Sub LoadListInArray()
On Error GoTo dlgerror
Dim sCurrent As String, sCurrent2 As String
Dim i As Integer, j As Integer, k As Integer, M As Integer
Dim sLocation As String


sLocation = App.Path & MACRO_FILE
Open sLocation For Input As #1
Line Input #1, sCurrent
M = Val(sCurrent)
ReDim mMacros(1 To 2, 0 To M)

Do Until EOF(1)
    Line Input #1, sCurrent

    If Left$(sCurrent, 2) = "**" Then
        k = 1
        i = i + 1
        
        If LenB(sCurrent2) Then
            mMacros(2, i - 2) = Left(sCurrent2, Len(sCurrent2) - 2)
            sCurrent2 = vbNullString
        End If
    Else
        If k = 1 Then
            mMacros(1, i - 1) = sCurrent
            k = k + 1
        ElseIf k > 1 Then
            sCurrent2 = sCurrent2 & sCurrent & vbCrLf
        End If
    End If

Loop

Close #1
'sCompactString mMacros()
'MsgBox mMacros(1, UBound(mMacros, 2) - 1)
Exit Sub


dlgerror:
MsgBox "An error has occured " & Err.Description
Exit Sub
End Sub
Public Sub PutSelectedText()
Dim startLine As Long, startCol As Long
Dim endLine As Long, endCol As Long
Dim codeText As String

Dim oCM As CodeModule
Dim oCP As CodePane

On Error Resume Next
Set oCM = VBInstance.ActiveCodePane.CodeModule
Set oCP = oCM.CodePane
' get a reference to the active code window and the underlying module
' exit if no one is available

'If Err Then Exit Sub

' get the current selection coordinates
oCP.GetSelection startLine, startCol, endLine, endCol
' exit if no text is highlighted
If startLine = endLine And startCol = endCol Then
    MsgBox "Please select some code and right click on the selection.", vbInformation, "Fast Macros"
    Exit Sub
End If
' get the code text
If startLine = endLine Then
    ' only one line is partially or fully highlighted
    codeText = Mid$(oCM.Lines(startLine, 1), startCol, endCol - startCol)
Else
    ' the selection spans multiple lines of code
    ' first, get the selection of the first line
    codeText = Mid$(oCM.Lines(startLine, 1), startCol) & vbCrLf
    ' then get the lines in the middle, that are fully highlighted
    If startLine + 1 < endLine Then
        codeText = codeText & oCM.Lines(startLine + 1, endLine - startLine - 1)
    End If
    ' finally, get the highlighted portion of the last line
    codeText = codeText & Left$(oCM.Lines(endLine, 1), endCol - 1)
End If

mFastMacroTxt = codeText
WLText1.Text = codeText
WLText2.Text = vbNullString
Me.Show
frmAddinDisplayed = True
WLText2.SetFocus
End Sub
Sub LoadSettings()
With Ini
    .Path = App.Path & "\Settings.ini"
    .Section = "Settings"

    .Key = "ProgrammerName"
    WLText3.Text = .Value

    .Key = "MacrosEnabled"
    If .Value = "True" Then
        WLCheck1.Value = wlChecked
        WLCheck1.Caption = "The Macros are Enabled"
        MacrosEnabled = True
        'HookMDIClient
    Else
        WLCheck1.Value = wlUnchecked
        WLCheck1.Caption = "The Macros are Disabled"
    End If

    .Key = "SpaceEnhancer"
    If .Value = "True" Then
        Check2.Value = wlChecked
        bSpaceEnhancer = True
        'HookMainWindow
    Else
        Check2.Value = wlUnchecked
    End If

    .Key = "Maximise"
    If .Value = "True" Then
        Check1.Value = wlChecked
    Else
        Check1.Value = wlUnchecked
    End If

End With

End Sub
Sub SaveSettings()
With Ini
    .Path = App.Path & "\Settings.ini"
    .Section = "Settings"

    .Key = "ProgrammerName"
    .Value = WLText3.Text

    .Key = "MacrosEnabled"
    If WLCheck1.Value = wlChecked Then
        .Value = "True"
    Else
        .Value = "False"
    End If

    .Key = "SpaceEnhancer"
    If Check2.Value = wlChecked Then
        .Value = "True"
    Else
        .Value = False
    End If

    .Key = "Maximise"
    If Check1.Value = wlChecked = True Then
        .Value = "True"
    Else
        .Value = "False"
    End If


End With

End Sub
Public Sub SaveListFromArray()
On Error GoTo dlgerror

Dim sCurrent As String
Dim i As Integer
Dim sLocation As String

sLocation = App.Path & MACRO_FILE
On Error Resume Next
Kill sLocation
On Error GoTo dlgerror

Open sLocation For Output As #1
Print #1, Trim$(Str$(UBound(mMacros, 2)))

For i = 0 To UBound(mMacros, 2) - 1
        Print #1, "**"
        Print #1, mMacros(1, i)
        Print #1, mMacros(2, i)
Next

Print #1, "**"
Close #1

Exit Sub

dlgerror:
MsgBox "An error has occured " & Err.Description
Exit Sub
End Sub
Public Sub SaveListFromGrid()
On Error GoTo dlgerror

Dim sCurrent As String
Dim i As Integer
Dim sLocation As String
Const cBakFile = "\DataBak.txt"

sLocation = App.Path & MACRO_FILE
On Error Resume Next
Kill sLocation
On Error GoTo dlgerror

ReDim mMacros(MlcGrid1.Rows)
'? mMacrosCount = MlcGrid1.Rows

With MlcGrid1
    .Sorted 1, .Rows, 1, True
    Open sLocation For Output As #1
    Print #1, Str$(.Rows)

    For i = 1 To .Rows
        .Row = i
        Print #1, "**"
        .Col = 1
        Print #1, .Text
        'mMacros(i - 1).sCode = .Text
        mMacros(1, i - 1) = .Text
        .Col = 2
        Print #1, .Text
        'mMacros(i - 1).sMacro = .Text
        mMacros(2, i - 1) = .Text
    Next
End With

Print #1, "**"
Close #1

Exit Sub

dlgerror:
MsgBox "An error has occured " & Err.Description
Exit Sub
End Sub

Public Sub LoadListInGrid()
On Error GoTo dlgerror
Dim sCurrent As String, sCurrent2 As String
Dim i As Integer, j As Integer, k As Integer ', m As Integer
Dim sLocation As String

sLocation = App.Path & MACRO_FILE
Open sLocation For Input As #1
Line Input #1, sCurrent

With MlcGrid1
    .Rows = Val(sCurrent)

    Do Until EOF(1)
        Line Input #1, sCurrent

        If Left$(sCurrent, 2) = "**" Then
            k = 1
            i = i + 1
            If LenB(sCurrent2) Then
                .TextMatrix(i - 1, 2) = Left(sCurrent2, Len(sCurrent2) - 2)
                sCurrent2 = vbNullString
            End If
        Else
            If k = 1 Then
                .TextMatrix(i, 1) = sCurrent
                k = k + 1
            ElseIf k > 1 Then
                sCurrent2 = sCurrent2 & sCurrent & vbCrLf
            End If
        End If
    Loop
End With

Close #1
Exit Sub

dlgerror:
MsgBox "An error has occured " & Err.Description
Exit Sub
End Sub
Public Sub LoadListInGridFromArr()

Dim i As Integer, tmp As Variant
On Error GoTo dlgerror

With MlcGrid1
    '.Rows = UBound(mMacros, 2)
    .Rows = 0
    
    For i = 1 To UBound(mMacros, 2)
        '.TextMatrix(i, 1) = mMacros(1, i - 1)
        '.TextMatrix(i, 2) = mMacros(2, i - 1)
        tmp = mMacros(1, i - 1) & vbTab & mMacros(2, i - 1)
        .AddItem tmp
    Next
End With

Exit Sub

dlgerror:
MsgBox "An error has occured " & Err.Description
End Sub
Public Function ReturnStringByChar(ByRef sText As String, ByRef sChar As String) As String
Dim i As Integer

If sText = vbNullString Then Exit Function

i = InStr(1, sText, sChar)
If i > 0 Then
    ReturnStringByChar = Left$(sText, i - 1)
    sText = Mid$(sText, i + Len(sChar))
Else
    ReturnStringByChar = sText
    sText = vbNullString
End If

End Function
Private Sub Check1_Click()

Select Case Check1.Value
Case 0
    With Ini
        .Path = App.Path & "\settings.ini"
        .Section = "Settings"
        .Key = "Maximise"
        .Value = "False"
    End With
Case 1
    With Ini
        .Path = App.Path & "\settings.ini"
        .Section = "Settings"
        .Key = "Maximise"
        .Value = "True"
    End With
End Select

End Sub
Private Sub Check2_Click()

Select Case Check2.Value
Case 0
    With Ini
        .Path = App.Path & "\settings.ini"
        .Section = "Settings"
        .Key = "SpaceEnhancer"
        .Value = "False"
    End With

    bSpaceEnhancer = False
    UnhookMainWindow
    
    'put the IDE in order
    If Not LinkedWindowsVisible Then ShowLinkedWindows
Case 1
    With Ini
        .Path = App.Path & "\settings.ini"
        .Section = "Settings"
        .Key = "SpaceEnhancer"
        .Value = "True"
    End With

    bSpaceEnhancer = True
    HookMainWindow
End Select

End Sub
Private Sub Form_Load()

RemoveMenus Me, True, False, True, True, True, True, True
frmAddinDisplayed = False
Hide
LoadListInArray

With MlcGrid1
    '.Left = 0
    .Top = 480
    .Height = 5910
    '.Width = 3375
    .Clear
    .ResetProperties
    .Cols = 2
    .Width = 3860
    .Scrolls = Vertical
    .Backcolor = .CellsBackColor
    .Titulo = "^Short|^Long String"
    .ColWidth(1) = 845
    .ColWidth(2) = 2700
    .AllowUserChangeColPos = True
    .AllowUserSortCol = True

    LoadListInGridFromArr

    .Sorted 1, .Rows, 1, True
    .Row = 1
    '    .Col = 1
End With

With WLCombo1
    .AddItem CMB_CURSOR
    .AddItem CMB_STARTSEL
    .AddItem CMB_ENDSEL
    .AddItem CMB_LASTWORD
    .AddItem CMB_DATE
    .AddItem CMB_TIME
    .AddItem CMB_PROCNAME
    .AddItem CMB_PROCKIND
    .AddItem CMB_PROCARG
    .AddItem CMB_PROCRETURNTYPE
    .AddItem CMB_PROCDESCRIPTION
    .AddItem CMB_MODULENAME
    .AddItem CMB_MODULEFILENAME
    .AddItem CMB_MODULEFILEPATH
    .AddItem CMB_MODULETYPE
    .AddItem CMB_PROJECTNAME
    .AddItem CMB_PROJECTFILENAME
    .AddItem CMB_PROJECTFILEPATH
    .AddItem CMB_PROJECTTYPE
    '    .AddItem CMB_INPUTBOX
    .AddItem CMB_PROGRAMMERNAME
End With

LinkedWindowsVisible = True
LoadSettings

End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim sLocation As String

Const cBakFile = "\DataBak.txt"
sLocation = App.Path & MACRO_FILE
Cancel = Not ReadyToClose

If FileExists(App.Path & cBakFile) = True Then Kill App.Path & cBakFile
FileCopy sLocation, App.Path & cBakFile
SaveSettings

End Sub

Private Sub MlcGrid1_Click()
Dim txtTemp As String, CaretPos As Long

IsFastMacro = False
mFastMacroTxt = vbNullString

WLText1 = MlcGrid1.TextMatrix(MlcGrid1.Row, 2)
WLText2 = MlcGrid1.TextMatrix(MlcGrid1.Row, 1)

txtTemp = WLText1
CaretPos = InStr(1, txtTemp, "#CURSOR#")

If CaretPos Then
    'WLText1.SelStart = CaretPos - 1
    'WLText1.SelLength = Len("#CURSOR#")
    Label1 = CaretPos - 1
Else
    Label1 = Len(WLText1)
End If

'WLText1.SetFocus
End Sub
Private Sub WLCheck1_Click()

If WLCheck1.Value = wlChecked Then
    WLCheck1.Caption = "The Macros are Enabled"
    MacrosEnabled = True
    HookMDIClient

Else
    WLCheck1.Caption = "The Macros are Disabled"
    MacrosEnabled = False
    UnhookMDIClient
End If

End Sub
Private Sub WLCombo1_Click()

With WLText1
    .SelStart = mWLText1SetStart
    .SelText = WLCombo1.Text
End With

End Sub
Private Sub WLCombo1_DropDown()
WLCombo1.Text = WLCombo1.List(0)
End Sub

Private Sub WLCommand2_Click()
SetTopmost Me, False
Hide
frmAddinDisplayed = False
IsFastMacro = False
mFastMacroTxt = vbNullString
MlcGrid1.Sorted 1, MlcGrid1.Rows, 1, True
WLText1.Text = vbNullString
WLText2.Text = vbNullString
Label1.Caption = vbNullString
MlcGrid1.Row = 1
End Sub

Private Sub WLCommand4_Click()
Dim i As Integer, j As Integer, NewUBound As Integer

If Len(WLText1) = 0 Then
    MsgBox "Missing Macro"
    WLText1.SetFocus
ElseIf Len(WLText2) = 0 Then
    MsgBox "Missing Code"
    WLText2.SetFocus
Else
    If Right(WLText1, 1) = Chr(32) Then
        WLText1 = WLText1 & CMB_CURSOR
    End If
    
    With MlcGrid1
        For i = 1 To .Rows
            If WLText2.Text = .TextMatrix(i, 1) Then
                .TextMatrix(i, 2) = WLText1.Text
                .Row = i
                
                For j = 0 To UBound(mMacros, 1) - 1
                    If mMacros(1, j) = WLText2.Text Then
                        mMacros(2, j) = WLText1.Text
                        Exit For
                    End If
                Next j
                
                SaveListFromArray
                Exit Sub
            End If
        Next i

        .AddItem WLText2.Text & vbTab & WLText1.Text, 1
        .Row = 1
    End With

    NewUBound = UBound(mMacros, 2) + 1
    ReDim Preserve mMacros(1 To 2, 0 To NewUBound)
    mMacros(1, NewUBound - 1) = WLText2.Text
    mMacros(2, NewUBound - 1) = WLText1.Text
    sTriQuickSortString mMacros()
    sCompactString mMacros()
    SaveListFromArray
End If
End Sub
Private Sub WLCommand5_Click()
Dim i As Integer

With MlcGrid1
    For i = 1 To .Rows
        If WLText2.Text = .TextMatrix(i, 1) Then
            .RemoveItem i
            Exit For
        End If
    Next
End With

For i = 0 To UBound(mMacros, 2) - 1
    If WLText2.Text = mMacros(1, i) Then
        mMacros(1, i) = vbNullString
        mMacros(2, i) = vbNullString
        WLText1.Text = vbNullString
        WLText2.Text = vbNullString
        Label1.Caption = vbNullString
        Exit For
    End If
Next



sCompactString mMacros()
SaveListFromArray
End Sub
Private Sub WLCommand6_Click()
frmPrint.Show

End Sub
Private Sub WLText1_KeyUp(KeyCode As Integer, Shift As Integer)
mWLText1SetStart = WLText1.SelStart
Label1 = mWLText1SetStart
End Sub
Private Sub WLText1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
mWLText1SetStart = WLText1.SelStart
Label1 = mWLText1SetStart
End Sub
Private Sub WLText2_Change()

Dim i As Integer

For i = 0 To UBound(mMacros, 2) - 1
    If WLText2.Text = mMacros(1, i) Then
        If IsFastMacro Then mFastMacroTxt = WLText1.Text
        WLText1.Text = mMacros(2, i)
        Label1 = vbNullString
        Exit For
    Else
        If Len(WLText1.Text) > 0 Then
            If IsFastMacro Then
                WLText1.Text = mFastMacroTxt
            Else
                WLText1.Text = vbNullString
            End If
        End If
    End If
Next

End Sub



