VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Option"
   ClientHeight    =   3180
   ClientLeft      =   2565
   ClientTop       =   1500
   ClientWidth     =   6420
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   6420
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkTemplate 
      BackColor       =   &H00808080&
      Caption         =   "Html Template :"
      Height          =   495
      Left            =   285
      MaskColor       =   &H00808080&
      TabIndex        =   10
      Top             =   1050
      Width           =   2235
   End
   Begin VB.TextBox txtPort 
      Height          =   330
      Left            =   765
      TabIndex        =   8
      Top             =   2115
      Width           =   705
   End
   Begin VB.TextBox txtRootPath 
      Height          =   375
      Left            =   300
      TabIndex        =   5
      Top             =   630
      Width           =   4545
   End
   Begin VB.CommandButton cmdOpenDir 
      Caption         =   "Path..."
      Height          =   360
      Left            =   5040
      TabIndex        =   4
      Top             =   645
      Width           =   1050
   End
   Begin VB.CommandButton cmdOpenPath 
      Caption         =   "File..."
      Height          =   360
      Left            =   5055
      TabIndex        =   3
      Top             =   1575
      Width           =   1050
   End
   Begin VB.TextBox txtTemplate 
      Height          =   375
      Left            =   285
      TabIndex        =   2
      Top             =   1545
      Width           =   4545
   End
   Begin MSComDlg.CommonDialog Dlog1 
      Left            =   4080
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   360
      Left            =   5055
      TabIndex        =   1
      Top             =   2130
      Width           =   1050
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Save"
      Height          =   360
      Left            =   3795
      TabIndex        =   0
      Top             =   2130
      Width           =   1050
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00E0E0E0&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      X1              =   120
      X2              =   6285
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Label lblPort 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Port :"
      Height          =   195
      Left            =   330
      TabIndex        =   9
      Top             =   2190
      Width           =   375
   End
   Begin VB.Label lblAbout 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Any change will not be applied until the Server restarted."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00994411&
      Height          =   210
      Left            =   225
      TabIndex        =   7
      Top             =   2730
      Width           =   4725
   End
   Begin VB.Label lblRootPath 
      BackStyle       =   0  'Transparent
      Caption         =   "RootPath :"
      Height          =   270
      Left            =   330
      TabIndex        =   6
      Top             =   315
      Width           =   3540
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      BorderStyle     =   6  'Inside Solid
      Height          =   2955
      Left            =   90
      Top             =   120
      Width           =   6210
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim hssIni As HttpServerSet
    With hssIni
    .bUseTemplate = chkTemplate.Value
    .sTemplateFile = txtTemplate.Text
    .sRootPath = txtRootPath.Text
    If CLngStr(txtPort.Text) <> 0 Then .sPort = txtPort.Text
    End With
    modHttpServer.hs_saveServerSetting frmServer.iniHsszh, hssIni
    Unload Me
End Sub

Private Sub cmdOpenDir_Click()
    txtRootPath.Text = openDirDialog(Me.hwnd)
End Sub

Private Sub cmdOpenPath_Click()
    Dim fso As New FileSystemObject

    With Dlog1
        .Filter = "Html File|*.htm"

        If txtTemplate.Text <> "" Then .InitDir = fso.GetParentFolderName(txtTemplate.Text)
        .ShowOpen

        If .FileName <> "" Then txtTemplate.Text = .FileName
    End With

End Sub


Private Sub Form_Load()
    '÷√÷–¥∞ÃÂ
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
    
    Dim hsIni As HttpServerSet
    
    modHttpServer.hs_getServerSetting frmServer.iniHsszh, hsIni
    
    With hsIni
    txtTemplate.Text = .sTemplateFile
    If .bUseTemplate Then chkTemplate.Value = 1 Else chkTemplate.Value = 0
    txtRootPath.Text = .sRootPath
    txtPort.Text = .sPort
    End With
End Sub



Private Sub txtPort_Change()
    txtPort.Text = CLngStr(txtPort.Text)
End Sub


