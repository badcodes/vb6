VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Setting"
   ClientHeight    =   5715
   ClientLeft      =   465
   ClientTop       =   360
   ClientWidth     =   8880
   Icon            =   "frmOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   8880
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSaveCss 
      Caption         =   "Save As CSS"
      Height          =   300
      Left            =   7320
      TabIndex        =   22
      Top             =   1680
      Width           =   1188
   End
   Begin VB.CommandButton cmdClearAddressBar 
      Caption         =   "Clear AddressBar"
      Height          =   300
      Left            =   6120
      TabIndex        =   21
      Top             =   4320
      Width           =   2304
   End
   Begin VB.Frame frmEx 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   1188
      Left            =   288
      TabIndex        =   13
      Top             =   336
      Width           =   8250
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   192
         Index           =   3
         Left            =   3324
         TabIndex        =   17
         Top             =   312
         Width           =   492
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   192
         Index           =   2
         Left            =   2736
         TabIndex        =   16
         Top             =   624
         Width           =   492
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   192
         Index           =   1
         Left            =   2076
         TabIndex        =   15
         Top             =   360
         Width           =   492
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   192
         Index           =   0
         Left            =   1164
         TabIndex        =   14
         Top             =   420
         Width           =   492
      End
   End
   Begin VB.CommandButton cmdClearRecent 
      Caption         =   "Clear Recent Fille Menu"
      Height          =   300
      Left            =   6120
      TabIndex        =   12
      Top             =   3840
      Width           =   2304
   End
   Begin VB.Frame fmOption 
      Caption         =   "Option"
      Height          =   1530
      Left            =   120
      TabIndex        =   9
      Top             =   3324
      Width           =   8595
      Begin VB.TextBox txtHttpPort 
         Height          =   288
         Left            =   3720
         TabIndex        =   23
         Text            =   "50211"
         Top             =   1080
         Width           =   1410
      End
      Begin VB.TextBox txtInterval 
         Height          =   288
         Left            =   3720
         TabIndex        =   19
         Text            =   "1000"
         Top             =   684
         Width           =   1410
      End
      Begin VB.TextBox txtRecentMax 
         Height          =   288
         Left            =   3720
         TabIndex        =   10
         Text            =   "10"
         Top             =   336
         Width           =   1395
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Http Server Port:"
         Height          =   285
         Left            =   180
         TabIndex        =   24
         Top             =   1095
         Width           =   3480
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Interval of AutoRandom (ms) :"
         Height          =   285
         Left            =   180
         TabIndex        =   20
         Top             =   705
         Width           =   3480
      End
      Begin VB.Label lblRecentFile 
         Alignment       =   2  'Center
         Caption         =   "Numbers of Recent File List In Menu:"
         Height          =   285
         Left            =   180
         TabIndex        =   11
         Top             =   375
         Width           =   3480
      End
   End
   Begin VB.Frame txtFrame 
      Caption         =   "Display"
      Height          =   3144
      Left            =   60
      TabIndex        =   2
      Top             =   72
      Width           =   8640
      Begin VB.CheckBox chkUseTemplate 
         Caption         =   "Use Html Template Instead"
         Height          =   312
         Left            =   192
         TabIndex        =   18
         Top             =   2160
         Width           =   5916
      End
      Begin VB.TextBox txtTemplate 
         Height          =   320
         Left            =   192
         TabIndex        =   8
         Top             =   2532
         Width           =   6825
      End
      Begin VB.CommandButton cmdOpenPath 
         Caption         =   "File..."
         Height          =   300
         Left            =   7320
         TabIndex        =   7
         Top             =   2520
         Width           =   1116
      End
      Begin VB.ComboBox cboLineHeight 
         Height          =   300
         Left            =   3660
         TabIndex        =   6
         Text            =   "cboLineHeight"
         Top             =   1632
         Width           =   1104
      End
      Begin VB.CommandButton cmdBackColor 
         Caption         =   "BackColor"
         Height          =   300
         Left            =   2316
         TabIndex        =   5
         Top             =   1620
         Width           =   1155
      End
      Begin VB.CommandButton cmdForeColor 
         Caption         =   "ForeColor"
         Height          =   300
         Left            =   1188
         TabIndex        =   4
         Top             =   1620
         Width           =   960
      End
      Begin VB.CommandButton cmdFont 
         Caption         =   "Font"
         Height          =   300
         Left            =   204
         TabIndex        =   3
         Top             =   1620
         Width           =   756
      End
      Begin VB.Shape Shape1 
         Height          =   1224
         Left            =   216
         Top             =   240
         Width           =   8280
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   300
      Left            =   7545
      TabIndex        =   1
      Top             =   5040
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Save"
      Height          =   300
      Left            =   6210
      TabIndex        =   0
      Top             =   5040
      Width           =   1095
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBackColor_Click()

    Dim fResult As Boolean
    Dim iColor As Long
    iColor = frmEx.BackColor
    Dim cDLG As New CCommonDialogLite
    fResult = cDLG.VBChooseColor(iColor)
    If fResult Then frmEx.BackColor = iColor 'tempVS.BackColor
    Set cDLG = Nothing

End Sub

Private Sub cmdCancel_Click()

    Unload Me

End Sub

Private Sub cmdClearAddressBar_Click()
MainFrm.cmbAddress.Clear
End Sub

Private Sub cmdClearRecent_Click()

    Dim hMRU As New CMenuArrHandle
    hMRU.Menus = MainFrm.mnuFile_Recent
    hMRU.ReduceTo 1
    Set hMRU = Nothing

End Sub

Private Sub cmdFont_Click()

    Dim cDLG  As New CCommonDialogLite
    Dim fResult As Boolean
    Dim fontTemp As New StdFont
    'MStdFonts.CopytoIfont tempVS.Viewfont, fontTemp
    MStdFonts.FontEqual fontTemp, frmEx.Font
    fResult = cDLG.VBChooseFont(fontTemp)
    Set cDLG = Nothing

    If fResult Then
        MStdFonts.FontEqual frmEx.Font, fontTemp
        'tempVS.Viewfont = MStdFonts.toMYfont(fontTemp)
        Label1(0).Top = 0
        Dim i As Integer
        For i = 0 To 3
            MStdFonts.FontEqual Label1(i).Font, fontTemp
            Label1(i).Left = 0
        Next
        For i = 1 To 3
            Label1(i).Top = Label1(i - 1).Top + Label1(i - 1).Height * Val(cboLineHeight.text) / 100
        Next

    End If

End Sub

Private Sub cmdForeColor_Click()

    Dim cDLG As New CCommonDialogLite
    Dim fResult As Boolean
    Dim iColor As Long
    iColor = frmEx.ForeColor
    fResult = cDLG.VBChooseColor(iColor)
    Set cDLG = Nothing

    If fResult Then
        frmEx.ForeColor = iColor
        Dim i As Integer

        For i = 0 To 3
            Label1(i).ForeColor = frmEx.ForeColor
            Label1(i).Left = 0
        Next

    End If

End Sub
Public Sub MakeCss(ByRef fullpath As String)

    Dim fNum As Integer
    Dim sFile As String
    Dim sAppend As New CAppendString
    
    sFile = fullpath
    fNum = FreeFile
    Open sFile For Output As #fNum
    
    With sAppend
        .AppendLine "body {"
        .AppendLine "        background-color: #" + VBColorToRGB(frmEx.BackColor) + ";"
        .AppendLine "        margin-left:8;"
        .AppendLine "        margin-top:8;"
        .AppendLine "        margin-right:8;"
        .AppendLine "        margin-bottom:8;"
'        .AppendLine "        font-family:" + Chr$(34) + frmEx.Font.name + Chr$(34) + ";"
'        '.AppendLine "        font-size:" + Str$(frmEx.Font.Size) + "pt;"
'        .AppendLine "        color: #" + VBColorToRGB(frmEx.ForeColor) + ";"
        .AppendLine "        line-height:" + cboLineHeight.Tag + ";"
        .AppendLine "     }"
        .AppendLine "body,p,tr,td,.m_text {"
    End With

'FIXIT: Print method has no Visual Basic .NET equivalent and will not be upgraded.         FixIT90210ae-R7593-R67265
    Print #fNum, sAppend.Value;
    
    sAppend.Clear

    With sAppend
        .AppendLine "        font-family:" + Chr$(34) + frmEx.Font.name + Chr$(34) + ";"
        '.AppendLine "        font-size:" + Str$(frmEx.Font.Size) + "pt;"
        .AppendLine "        color: #" + VBColorToRGB(frmEx.ForeColor) + ";"
        '.AppendLine "        line-height:" + cboLineHeight.Tag + ";"
    End With

    With frmEx.Font
        If .Bold Then sAppend.AppendLine "        font-weight:Bold;"
        If .Italic Then sAppend.AppendLine "        font-style:Italic;"
        If .Underline Then sAppend.AppendLine "        text-decroation:underline;"
    End With

    With sAppend
        .AppendLine "}"
    End With

'FIXIT: Print method has no Visual Basic .NET equivalent and will not be upgraded.         FixIT90210ae-R7593-R67265
    Print #fNum, sAppend.Value;
'FIXIT: Print method has no Visual Basic .NET equivalent and will not be upgraded.         FixIT90210ae-R7593-R67265
    Print #fNum, "A {"
'FIXIT: Print method has no Visual Basic .NET equivalent and will not be upgraded.         FixIT90210ae-R7593-R67265
    Print #fNum, sAppend.Value;
    Close #fNum
    'Set fso = Nothing
    Set sAppend = Nothing
End Sub
Private Sub cmdOK_Click()

    'SaveViewerStyle , tempVS
'FIXIT: mnuFile_Recent(0).Tag property has no Visual Basic .NET equivalent and will not be upgraded.     FixIT90210ae-R7593-R67265
    MainFrm.mnuFile_Recent(0).Tag = txtRecentMax.text
    Dim mh As New CMenuArrHandle
    mh.Menus = MainFrm.mnuFile_Recent
'FIXIT: mnuFile_Recent(0).Tag property has no Visual Basic .NET equivalent and will not be upgraded.     FixIT90210ae-R7593-R67265
    mh.ReduceTo Val(MainFrm.mnuFile_Recent(0).Tag)
    Set mh = Nothing
    
    cboLineHeight.Tag = cboLineHeight.text
    MainFrm.IEView.Tag = Me.txtTemplate.text
    MainFrm.Tag = CInt(Me.chkUseTemplate.Value)
'FIXIT: Timer.Tag property has no Visual Basic .NET equivalent and will not be upgraded.     FixIT90210ae-R7593-R67265
    MainFrm.Timer.Tag = Me.txtInterval.text
    
    Dim hSetting As New CSetting
    hSetting.iniFile = zhtmIni
    With hSetting
    .Save frmEx, SF_COLOR
    .Save frmEx, SF_FONT
    .Save cboLineHeight, SF_Tag
    .Save MainFrm.IEView, SF_Tag
    .Save MainFrm.Timer, SF_Tag
    .Save MainFrm.mnuFile_Recent, SF_MENUARRAY
    End With
    Set hSetting = Nothing
    
'FIXIT: Timer.Tag property has no Visual Basic .NET equivalent and will not be upgraded.     FixIT90210ae-R7593-R67265
'FIXIT: Timer.Tag property has no Visual Basic .NET equivalent and will not be upgraded.     FixIT90210ae-R7593-R67265
    If Val(MainFrm.Timer.Tag) > 0 Then MainFrm.Timer.Interval = Val(MainFrm.Timer.Tag)
    
    MakeCss bddir(sConfigDir) + "style.css"
    MainFrm.GetView zhrStatus.sCur_zhSubFile
    
    

End Sub

Private Sub cmdOpenPath_Click()

    Dim fso As New GFileSystem
    Dim sInitDir As String
    Dim fResult As Boolean
    Dim sFilename As String
    Dim cDLG As New CCommonDialogLite
    If txtTemplate.text <> "" Then sInitDir = fso.GetParentFolderName(txtTemplate.text)
    Set fso = Nothing
    fResult = cDLG.VBGetOpenFileName( _
       FileName:=sFilename, _
       Filter:="Html File|*.htm|All File|*.*", _
       InitDir:=sInitDir, _
       DlgTitle:=cmdOpenPath.Caption, _
       Owner:=Me.hwnd)
    Set cDLG = Nothing

    If fResult Then
        If sFilename <> "" Then txtTemplate.text = sFilename
    End If

End Sub

Private Sub cboLineheight_Click()
    
    Dim iHeight As Long
    iHeight = Val(cboLineHeight.text)
    Dim i As Integer
    Label1(0).Top = 0

    For i = 1 To 3
        Label1(i).Top = Label1(i - 1).Top + Label1(i - 1).Height * iHeight / 100
    Next

End Sub

Private Sub cmdSaveCss_Click()
    Dim fso As New GFileSystem
    Dim sInitDir As String
    Dim fResult As Boolean
    Dim sFilename As String
    Dim cDLG As New CCommonDialogLite
    sInitDir = sConfigDir
    Set fso = Nothing
    fResult = cDLG.VBGetSaveFileName( _
       FileName:=sFilename, _
       Filter:="Style Sheet File|*.css|All File|*.*", _
       InitDir:=sInitDir, _
       DlgTitle:=cmdOpenPath.Caption, _
       Owner:=Me.hwnd, _
       DefaultExt:="css" _
       )
    Set cDLG = Nothing

    If fResult Then
        If sFilename <> "" Then MakeCss sFilename
    End If
    

End Sub

Private Sub Form_Load()

    Dim zhLocalize As New CLocalizer
    zhLocalize.Install Me, LanguageIni
    zhLocalize.loadFormStr
    Set zhLocalize = Nothing
    '置中窗体
    Me.Icon = MainFrm.Icon
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
    
    Dim hSetting As New CSetting
    hSetting.iniFile = zhtmIni
    With hSetting
    .Load frmEx, SF_COLOR
    .Load frmEx, SF_FONT
    .Load cboLineHeight, SF_Tag
    End With
    Set hSetting = Nothing
    
    Dim i As Integer
    
    Label1(0).Top = 0
    For i = 0 To 3
        MStdFonts.FontEqual Label1(i).Font, frmEx.Font
        'Label1(i).Font = frmEx.Font ' empVS.Viewfont
        Label1(i).ForeColor = frmEx.ForeColor ' tempVS.ForeColor
        Label1(i).Left = 0
    Next
    Label1(0).Caption = "天若有情天亦老凤凰台上凤凰游剑气萧瑟清澈山水秋风扫落叶天若有情天亦老凤凰台上凤凰游剑气萧瑟清澈山水秋风扫落叶天若有情天亦老凤凰台上凤凰游剑气萧瑟清澈山水秋风扫落叶"
    Label1(1).Caption = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"
    Label1(2).Caption = Label1(0).Caption
    Label1(3).Caption = Label1(1).Caption
    
    For i = 100 To 400 Step 10
        cboLineHeight.AddItem Str$(i) + "%"
    Next
    If cboLineHeight.Tag = "" Then cboLineHeight.Tag = " 100%"
    cboLineHeight.text = cboLineHeight.Tag
    
    If Val(MainFrm.Tag) <> 0 Then
        chkUseTemplate.Value = 1
    Else
        chkUseTemplate.Value = 0
    End If

    txtTemplate = MainFrm.IEView.Tag 'tempVS.TemplateFile
'FIXIT: mnuFile_Recent(0).Tag property has no Visual Basic .NET equivalent and will not be upgraded.     FixIT90210ae-R7593-R67265
    txtRecentMax.text = MainFrm.mnuFile_Recent(0).Tag  ' tempVS.RecentMax
'FIXIT: Timer.Tag property has no Visual Basic .NET equivalent and will not be upgraded.     FixIT90210ae-R7593-R67265
    txtInterval.text = MainFrm.Timer.Tag ' tempVS.AutoRandomInterval
    
    

End Sub



Private Sub Form_Unload(Cancel As Integer)
    MainFrm.Load_StyleSheetList
End Sub
