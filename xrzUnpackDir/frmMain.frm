VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "xrzUnpackDir"
   ClientHeight    =   7350
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   9195
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   9195
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtOptions 
      Height          =   330
      Left            =   120
      TabIndex        =   20
      Top             =   3360
      Width           =   8925
   End
   Begin VB.CheckBox chkSameDirectory 
      Caption         =   "Source Related"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   1920
      Value           =   1  'Checked
      Width           =   2055
   End
   Begin VB.Frame FrameOpt 
      BorderStyle     =   0  'None
      Height          =   615
      Index           =   1
      Left            =   120
      TabIndex        =   16
      Top             =   2400
      Width           =   9135
      Begin VB.TextBox txtExtension 
         Height          =   330
         Left            =   0
         TabIndex        =   11
         Top             =   240
         Width           =   8925
      End
      Begin VB.Label lblPath 
         Caption         =   "4. File Extensions to scan  (default to ""zip;7z;rar"")"
         Height          =   210
         Index           =   5
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   7290
      End
   End
   Begin VB.Frame FrameOpt 
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   15
      Top             =   2400
      Width           =   9135
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear Log"
      Height          =   375
      Left            =   7800
      TabIndex        =   12
      Top             =   6720
      Width           =   1215
   End
   Begin VB.CommandButton cmdProcess 
      Caption         =   "Start"
      Default         =   -1  'True
      Height          =   375
      Left            =   7800
      TabIndex        =   13
      Top             =   4680
      Width           =   1215
   End
   Begin VB.TextBox txtLog 
      Appearance      =   0  'Flat
      Height          =   1395
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   14
      Top             =   5160
      Width           =   8925
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "Select..."
      Height          =   375
      Index           =   2
      Left            =   7800
      TabIndex        =   9
      Top             =   1860
      Width           =   1215
   End
   Begin VB.TextBox txtPath 
      Height          =   330
      Index           =   2
      Left            =   2400
      TabIndex        =   8
      Top             =   1920
      Width           =   5205
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "Select..."
      Height          =   375
      Index           =   1
      Left            =   7800
      TabIndex        =   5
      Top             =   1140
      Width           =   1215
   End
   Begin VB.TextBox txtPath 
      Height          =   330
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   7485
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "Select..."
      Height          =   375
      Index           =   0
      Left            =   7800
      TabIndex        =   2
      Top             =   420
      Width           =   1215
   End
   Begin VB.TextBox txtPath 
      Height          =   330
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   7485
   End
   Begin VB.Label lblPath 
      Caption         =   "5. Additional Options:"
      Height          =   210
      Index           =   3
      Left            =   120
      TabIndex        =   21
      Top             =   3120
      Width           =   7290
   End
   Begin VB.Line Line 
      BorderColor     =   &H8000000D&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      Index           =   2
      X1              =   120
      X2              =   7440
      Y1              =   5055
      Y2              =   5055
   End
   Begin VB.Label LabelInfo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00400000&
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   2
      Left            =   120
      TabIndex        =   19
      Top             =   4800
      Width           =   7335
   End
   Begin VB.Line Line 
      BorderColor     =   &H8000000D&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      Index           =   1
      X1              =   120
      X2              =   7440
      Y1              =   4695
      Y2              =   4695
   End
   Begin VB.Label LabelInfo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00004000&
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   1
      Left            =   120
      TabIndex        =   18
      Top             =   4440
      Width           =   7335
   End
   Begin VB.Line Line 
      BorderColor     =   &H8000000D&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      Index           =   0
      X1              =   120
      X2              =   7440
      Y1              =   4340
      Y2              =   4340
   End
   Begin VB.Label LabelInfo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   0
      Left            =   120
      TabIndex        =   17
      Top             =   4080
      Width           =   7335
   End
   Begin VB.Label lblPath 
      Caption         =   "3.Target Directory:"
      Height          =   210
      Index           =   2
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   7290
   End
   Begin VB.Label lblPath 
      Caption         =   "2. Source Directory:"
      Height          =   210
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   6390
   End
   Begin VB.Label lblPath 
      Caption         =   "1. 7zip Directory:"
      Height          =   210
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   5610
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkSameDirectory_Click()
    Debug.Print "I'm click"
    If chkSameDirectory.Value = 1 Then
        lblPath(2).Enabled = False
        txtPath(2).Locked = True
        txtPath(2).Appearance = 0
        txtPath(2).Enabled = False
        cmdSelect(2).Enabled = False
    Else
        lblPath(2).Enabled = True
        txtPath(2).Enabled = True
        txtPath(2).Locked = False
        txtPath(2).Appearance = 1
        cmdSelect(2).Enabled = True
    End If
    txtPath_Change 1
End Sub

Private Sub cmdClear_Click()
txtLog.Text = ""
End Sub


Private Sub ScrollToEnd(ByRef vText As TextBox)
    MForms.TextBoxScroll vText, False
End Sub

Private Sub cmdProcess_Click()
Static sProcess As Boolean


If sProcess Then
    sProcess = False
    cmdProcess.Enabled = False
    cmdProcess.Caption = "Processing..."
Else
    Dim p7zip As String
    Dim pSrc As String
    Dim pDst As String
    p7zip = txtPath(0).Text
    Dim I As Long
    Dim Appname As String
    
    Appname = BuildPath(p7zip, "7zG.exe")
    If Not FileExists(Appname) Then Appname = BuildPath(p7zip, "7za.exe")
    
    If Not FileExists(Appname) Then
        MsgBox "7zip application not found!" & vbCrLf & "Maybe invalid path set!" & vbCrLf, vbCritical
        txtPath(0).SetFocus
        Exit Sub
    End If
    
    If Len(txtPath(1).Text) < 1 Then
        MsgBox "At least one option not specified " & vbCrLf & lblPath(1).Caption, vbCritical
        txtPath(1).SetFocus
        Exit Sub
    End If
    
    Dim appArg As String
    Dim fileExt As String
'
'    If Len(txtFormatOption.Text) > 0 Then
'        appArg = "-t" & txtFormatOption.Text
'        fileExt = txtFormatOption.Text
'    Else
'        For I = 0 To OptionFormat.ubound
'            If OptionFormat(I).Value Then
'                appArg = "-t" & OptionFormat(I).Caption
'                fileExt = OptionFormat(I).Caption
'                Exit For
'            End If
'        Next
'    End If
    
    If Len(txtOptions.Text) > 0 Then
        appArg = txtOptions.Text
    End If
    
    If Len(txtExtension.Text) > 0 Then
        fileExt = txtExtension.Text
    Else
        fileExt = "zip;7z;rar"
    End If
           
    
    'cmdProcess.Enabled = False
'    If OptPackAs(0).Value = True Then
'        p7zip = p7zip & " -afzip"
'    Else
'        p7zip = p7zip & " -afrar"
'    End If
    pSrc = BuildPath(txtPath(1).Text)
    pDst = BuildPath(txtPath(2).Text)
    
    Dim pDstFile As String
    
    sProcess = True
    cmdProcess.Caption = "Stop!"
    'cmdProcess.Enabled = True
    Dim pExts() As String
    Dim pArchives() As String
    Dim pCount As Long
    Dim pTemp As String
    pExts = Split(fileExt, ";")
    Dim iExt As Long
    Dim cExt As Long
    cExt = UBound(pExts)
    For iExt = 0 To cExt
        pTemp = Dir(pSrc & "*." & pExts(iExt))
        Do While (Len(pTemp) > 0)
            
            If Not (GetAttr(pSrc & pTemp) And vbDirectory) Then
            ReDim Preserve pArchives(0 To pCount) As String
            pArchives(pCount) = pTemp
            pCount = pCount + 1
            End If
            pTemp = Dir()
        Loop
    Next
    If pCount < 1 Then
        LabelInfo(0).Caption = "No archive found!"
        LabelInfo(1).Caption = ""
        LabelInfo(2).Caption = ""
    Else
        pCount = pCount - 1
        For I = 0 To pCount
            If sProcess = False Then
                LabelInfo(0).Caption = "Processing interrupted"
                MsgBox "Processing Interrupted", vbInformation
                GoTo Abort_Process
            End If
            pDstFile = pDst & GetBaseName(pArchives(I))
            LabelInfo(0).Caption = "Processing " & I + 1 & "/" & pCount + 1
            LabelInfo(1).Caption = pArchives(I)
            LabelInfo(2).Caption = pDstFile
            txtLog.Text = txtLog.Text & pArchives(I) & " ..."
            txtLog.Text = txtLog.Text & vbTab & M7zip.ShellExtract(Appname, appArg, pSrc & pArchives(I), pDstFile) & vbCrLf
            
        Next
        LabelInfo(0).Caption = pCount + 1 & " Archives Processed."
    End If

Abort_Process:
    sProcess = False
    cmdProcess.Caption = "Start"
    cmdProcess.Enabled = True
End If
End Sub


'CSEH: ErrMsgBox
Private Function RunOn(ByVal vWinRar As String, ByVal vSrc As String, ByVal vDst As String, Optional vZipMode As Boolean = False) As String
        '<EhHeader>
        On Error GoTo RunOn_Err
        '</EhHeader>
        Dim pCurDir As String
        Dim pCurDrive As String
100     pCurDir = CurDir$
102     pCurDrive = Left$(pCurDir, 1)
        Dim pDstName As String
104     pDstName = GetFileName(vSrc)
106     If vZipMode Then
108         vWinRar = vWinRar & " a -afzip"
110         pDstName = pDstName & ".zip"
        Else
112         vWinRar = vWinRar & " a -afrar"
114         pDstName = pDstName & ".rar"
        End If
116     ChDrive Left$(vSrc, 1)
118     ChDir vSrc
    
120     MShell32.ShellAndClose vWinRar & " -ibck -r " & QuoteString(vDst & pDstName) & " *.*", vbNormalNoFocus
122      ChDrive pCurDrive
124     ChDir pCurDir
126     RunOn = vDst & pDstName

        '<EhFooter>
        Exit Function

RunOn_Err:
        MsgBox Err.Description & vbCrLf & _
               "in FolderPacker.frmMain.RunOn " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        On Error Resume Next
    ChDrive pCurDrive
     ChDir pCurDir
     RunOn = vDst & pDstName
        '</EhFooter>
End Function

Private Sub cmdSelect_Click(Index As Integer)
    Dim ret As String
'    If Index = 0 Then
'        Dim dlgF As CCommonDialogLite
'        Set dlgF = New CCommonDialogLite
'        ret = txtPath(Index).Text
'        If dlgF.VBGetOpenFileName(ret, Filter:="EXE нд╪Ч(*.exe)|*.exe") Then
'            txtPath(Index).Text = ret
'        End If
'    Else
        Dim dlgD As CFolderBrowser
        Set dlgD = New CFolderBrowser
        ret = txtPath(Index).Text
        dlgD.InitDirectory = ret
        dlgD.Owner = Me.hWnd
        dlgD.Title = lblPath(Index).Caption
        ret = dlgD.Browse
        If ret <> "" Then txtPath(Index).Text = ret
'    End If
End Sub

Private Sub Form_Load()
    Dim cfgHnd As CVBSetting
    Set cfgHnd = New CVBSetting
    cfgHnd.Appname = App.EXEName
    cfgHnd.Section = "Config"
'    cfgHnd.ReadPropValue OptionMode(0), "OptMode0"
'    cfgHnd.ReadPropValue OptionMode(1), "OptMode1"
'    cfgHnd.ReadPropValue OptionFormat(0), "OptFormat0"
'    cfgHnd.ReadPropValue OptionFormat(1), "OptFormat1"
'    cfgHnd.ReadPropValue OptionFormat(2), "OptFormat2"
    cfgHnd.ReadPropTexts txtPath
'    cfgHnd.ReadPropText txtFormatOption, "ArchiveFormat", ""
    cfgHnd.ReadPropText txtExtension, "ArchiveExtension", ""
    cfgHnd.ReadPropText txtOptions, "ArchiveOptions", ""
'    cfgHnd.ReadPropText txtPath(0), "Path0"
'    cfgHnd.ReadPropText txtPath(1), "Path1"
'    cfgHnd.ReadPropText txtPath(2), "Path2"
    cfgHnd.ReadPropValue chkSameDirectory, "SameDirectory"
    
    chkSameDirectory_Click
    
    txtPath_Change 1
    If Command$ <> "" Then
        Dim pFile As String
        pFile = GetQToken(Command$, " " & Chr$(34))
        Do While pFile <> ""
            txtPath(1).Text = pFile
            txtPath(2).Text = pFile
'            OptionMode(0).Value = True
'            OptionFormat(0).Value = True
            chkSameDirectory.Value = 1
            cmdProcess_Click
            pFile = GetQToken("", " " & Chr$(34))
        Loop
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim cfgHnd As CVBSetting
    Set cfgHnd = New CVBSetting
    cfgHnd.Appname = App.EXEName
    cfgHnd.Section = "Config"
    
'    cfgHnd.WritePropValue OptionMode(0), "OptMode0"
'    cfgHnd.WritePropValue OptionMode(1), "OptMode1"
'    cfgHnd.WritePropValue OptionFormat(0), "OptFormat0"
'    cfgHnd.WritePropValue OptionFormat(1), "OptFormat1"
'    cfgHnd.WritePropValue OptionFormat(2), "OptFormat2"
    cfgHnd.WritePropTexts txtPath
    'cfgHnd.WritePropText txtFormatOption, "ArchiveFormat"
    cfgHnd.WritePropText txtExtension, "ArchiveExtension"
    cfgHnd.WritePropValue chkSameDirectory, "SameDirectory"
    cfgHnd.WritePropText txtOptions, "ArchiveOptions"
    
    
End Sub

Private Sub OptionMode_Click(Index As Integer)
    txtPath_Change 1
End Sub

Private Sub txtLog_Change()
        ScrollToEnd txtLog
End Sub

Private Sub txtPath_Change(Index As Integer)
    If Index = 1 And chkSameDirectory.Value = 1 Then
'        If OptionMode(0).Value = True Then
'            txtPath(2).Text = GetParentFolderName(txtPath(1).Text)
'        Else
            txtPath(2).Text = txtPath(1).Text
        'End If
    End If
End Sub
