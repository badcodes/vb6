VERSION 5.00
Begin VB.Form frmOptions 
   Caption         =   "Setting"
   ClientHeight    =   2988
   ClientLeft      =   540
   ClientTop       =   432
   ClientWidth     =   6564
   Icon            =   "frmOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   2988
   ScaleWidth      =   6564
   Begin VB.CommandButton cmdClearRecent 
      Caption         =   "Clear Recent Fille Menu"
      Height          =   300
      Left            =   3816
      TabIndex        =   7
      Top             =   1776
      Width           =   2304
   End
   Begin VB.Frame fmOption 
      Caption         =   "Option"
      Height          =   1176
      Left            =   120
      TabIndex        =   4
      Top             =   1068
      Width           =   6312
      Begin VB.TextBox txtInterval 
         Height          =   288
         Left            =   2772
         TabIndex        =   8
         Top             =   684
         Width           =   576
      End
      Begin VB.TextBox txtRecentMax 
         Height          =   288
         Left            =   2916
         TabIndex        =   5
         Top             =   336
         Width           =   432
      End
      Begin VB.Label Label2 
         Caption         =   "Interval of AutoRandom (ms) :"
         Height          =   288
         Left            =   192
         TabIndex        =   9
         Top             =   708
         Width           =   2976
      End
      Begin VB.Label lblRecentFile 
         Caption         =   "Numbers of Recent File List In Menu:"
         Height          =   288
         Left            =   180
         TabIndex        =   6
         Top             =   384
         Width           =   2976
      End
   End
   Begin VB.Frame txtFrame 
      Caption         =   "Display"
      Height          =   864
      Left            =   60
      TabIndex        =   2
      Top             =   72
      Width           =   6360
      Begin VB.CommandButton cmdBackColor 
         Caption         =   "BackColor"
         Height          =   384
         Left            =   216
         TabIndex        =   3
         Top             =   288
         Width           =   1155
      End
      Begin VB.Label lblColor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   468
         Left            =   1572
         TabIndex        =   10
         Top             =   228
         Width           =   4608
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   300
      Left            =   5268
      TabIndex        =   1
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Save"
      Height          =   300
      Left            =   3804
      TabIndex        =   0
      Top             =   2400
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
    iColor = lblColor.BackColor
    Dim cDLG As New CCommonDialogLite
    fResult = cDLG.VBChooseColor(iColor)
    If fResult Then lblColor.BackColor = iColor  'tempVS.BackColor
    Set cDLG = Nothing

End Sub

Private Sub cmdCancel_Click()

    Unload Me

End Sub



Private Sub cmdClearRecent_Click()

    Dim hMRU As New CMenuArrHandle
    hMRU.Menus = MainFrm.mnuFile_Recent
    hMRU.ReduceTo 1
    Set hMRU = Nothing

End Sub




Private Sub cmdOK_Click()

    'SaveViewerStyle , tempVS
    MainFrm.mnuFile_Recent(0).Tag = txtRecentMax.text
    Dim mh As New CMenuArrHandle
    mh.Menus = MainFrm.mnuFile_Recent
    mh.ReduceTo Val(MainFrm.mnuFile_Recent(0).Tag)
    Set mh = Nothing
    
    MainFrm.Timer.Tag = Me.txtInterval.text
    MainFrm.theShow.BackColor = lblColor.BackColor
    
    Dim hSetting As New CSetting
    hSetting.iniFile = zhtmIni
    With hSetting
    .Save MainFrm.Timer, SF_Tag
    .Save MainFrm.mnuFile_Recent, SF_MENUARRAY
    .Save MainFrm.theShow, SF_COLOR
    End With
    Set hSetting = Nothing
    
    If Val(MainFrm.Timer.Tag) > 0 Then MainFrm.Timer.Interval = Val(MainFrm.Timer.Tag)
    
    MainFrm.GetView zhrStatus.sCur_zhSubFile
    
    

End Sub




Private Sub Form_Load()

    Dim zhLocalize As New CLocalize
    zhLocalize.Install Me, LanguageIni
    zhLocalize.loadFormStr
    Set zhLocalize = Nothing
    '÷√÷–¥∞ÃÂ
    Me.Icon = MainFrm.Icon
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
    
    lblColor.BackColor = MainFrm.theShow.BackColor
    Dim i As Integer
    
   
    
    txtRecentMax.text = MainFrm.mnuFile_Recent(0).Tag  ' tempVS.RecentMax
    txtInterval.text = MainFrm.Timer.Tag ' tempVS.AutoRandomInterval

End Sub


