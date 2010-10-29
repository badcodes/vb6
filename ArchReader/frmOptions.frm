VERSION 5.00
Begin VB.Form frmOptions 
   Caption         =   "Setting"
   ClientHeight    =   5172
   ClientLeft      =   540
   ClientTop       =   432
   ClientWidth     =   6564
   Icon            =   "frmOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5172
   ScaleWidth      =   6564
   Begin VB.CommandButton cmdSaveCss 
      Caption         =   "Save As CSS"
      Height          =   300
      Left            =   4992
      TabIndex        =   22
      Top             =   1704
      Width           =   1188
   End
   Begin VB.CommandButton cmdClearAddressBar 
      Caption         =   "Clear AddressBar"
      Height          =   300
      Left            =   3864
      TabIndex        =   21
      Top             =   4020
      Width           =   2304
   End
   Begin VB.Frame frmEx 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   1188
      Left            =   288
      TabIndex        =   13
      Top             =   336
      Width           =   5868
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
      Left            =   3876
      TabIndex        =   12
      Top             =   3636
      Width           =   2304
   End
   Begin VB.Frame fmOption 
      Caption         =   "Option"
      Height          =   1176
      Left            =   120
      TabIndex        =   9
      Top             =   3324
      Width           =   6312
      Begin VB.TextBox txtInterval 
         Height          =   288
         Left            =   2772
         TabIndex        =   19
         Top             =   684
         Width           =   576
      End
      Begin VB.TextBox txtRecentMax 
         Height          =   288
         Left            =   2916
         TabIndex        =   10
         Top             =   336
         Width           =   432
      End
      Begin VB.Label Label2 
         Caption         =   "Interval of AutoRandom (ms) :"
         Height          =   288
         Left            =   192
         TabIndex        =   20
         Top             =   708
         Width           =   2976
      End
      Begin VB.Label lblRecentFile 
         Caption         =   "Numbers of Recent File List In Menu:"
         Height          =   288
         Left            =   180
         TabIndex        =   11
         Top             =   372
         Width           =   2976
      End
   End
   Begin VB.Frame txtFrame 
      Caption         =   "Display"
      Height          =   3144
      Left            =   60
      TabIndex        =   2
      Top             =   72
      Width           =   6360
      Begin VB.CheckBox chkUseTemplate 
         Caption         =   "Use Html Template Instead"
         Height          =   312
         Left            =   192
         TabIndex        =   18
         Top             =   2124
         Width           =   5916
      End
      Begin VB.TextBox txtTemplate 
         Height          =   320
         Left            =   192
         TabIndex        =   8
         Top             =   2532
         Width           =   4545
      End
      Begin VB.CommandButton cmdOpenPath 
         Caption         =   "File..."
         Height          =   300
         Left            =   5004
         TabIndex        =   7
         Top             =   2544
         Width           =   1116
      End
      Begin VB.ComboBox cboLineHeight 
         Height          =   288
         Left            =   3660
         Style           =   2  'Dropdown List
         TabIndex        =   6
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
         Width           =   5904
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   300
      Left            =   5148
      TabIndex        =   1
      Top             =   4680
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Save"
      Height          =   300
      Left            =   3804
      TabIndex        =   0
      Top             =   4680
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
    'MFont.CopytoIfont tempVS.Viewfont, fontTemp
    MFont.FontEqual fontTemp, frmEx.Font
    fResult = cDLG.VBChooseFont(fontTemp)
    Set cDLG = Nothing

    If fResult Then
        MFont.FontEqual frmEx.Font, fontTemp
        'tempVS.Viewfont = MFont.toMYfont(fontTemp)
        Label1(0).Top = 0
        Dim i As Integer
        For i = 0 To 3
            MFont.FontEqual Label1(i).Font, fontTemp
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
    Dim sFIle As String
    Dim sAppend As New CAppendString
    
    sFIle = fullpath
    fNum = FreeFile
    Open sFIle For Output As #fNum
    
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

    Print #fNum, sAppend.Value;
    Print #fNum, "A {"
    Print #fNum, sAppend.Value;
    Close #fNum
    'Set fso = Nothing
    Set sAppend = Nothing
End Sub
Private Sub cmdOK_Click()

    'SaveViewerStyle , tempVS
    MainFrm.mnuFile_Recent(0).Tag = txtRecentMax.text
    Dim mh As New CMenuArrHandle
    mh.Menus = MainFrm.mnuFile_Recent
    mh.ReduceTo Val(MainFrm.mnuFile_Recent(0).Tag)
    Set mh = Nothing
    
    cboLineHeight.Tag = cboLineHeight.text
    MainFrm.IEView.Tag = Me.txtTemplate.text
    MainFrm.Tag = CInt(Me.chkUseTemplate.Value)
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
    
    If Val(MainFrm.Timer.Tag) > 0 Then MainFrm.Timer.Interval = Val(MainFrm.Timer.Tag)
    
    MakeCss bddir(sConfigDir) + "style.css"
    MainFrm.GetView zhrStatus.sCur_zhSubFile
    
    

End Sub

Private Sub cmdOpenPath_Click()

    Dim fso As New gCFileSystem
    Dim sInitDir As String
    Dim fResult As Boolean
    Dim sFilename As String
    Dim cDLG As New CCommonDialogLite
    If txtTemplate.text <> "" Then sInitDir = fso.GetParentFolderName(txtTemplate.text)
    Set fso = Nothing
    fResult = cDLG.VBGetOpenFileName( _
       filename:=sFilename, _
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
    Dim fso As New gCFileSystem
    Dim sInitDir As String
    Dim fResult As Boolean
    Dim sFilename As String
    Dim cDLG As New CCommonDialogLite
    sInitDir = sConfigDir
    Set fso = Nothing
    fResult = cDLG.VBGetSaveFileName( _
       filename:=sFilename, _
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

    Dim zhLocalize As New CLocalize
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
        MFont.FontEqual Label1(i).Font, frmEx.Font
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
    txtRecentMax.text = MainFrm.mnuFile_Recent(0).Tag  ' tempVS.RecentMax
    txtInterval.text = MainFrm.Timer.Tag ' tempVS.AutoRandomInterval

End Sub



Private Sub Form_Unload(Cancel As Integer)
    MainFrm.Load_StyleSheetList
End Sub
