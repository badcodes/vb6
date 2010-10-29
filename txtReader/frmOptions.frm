VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Setting"
   ClientHeight    =   3972
   ClientLeft      =   528
   ClientTop       =   420
   ClientWidth     =   6564
   Icon            =   "frmOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3972
   ScaleWidth      =   6564
   Begin VB.Frame frmEx 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   1188
      Left            =   288
      TabIndex        =   11
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
         TabIndex        =   15
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
         TabIndex        =   14
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
         TabIndex        =   13
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
         TabIndex        =   12
         Top             =   420
         Width           =   492
      End
   End
   Begin VB.CommandButton cmdClearRecent 
      Caption         =   "Clear Recent Fille Menu"
      Height          =   336
      Left            =   3840
      TabIndex        =   10
      Top             =   2700
      Width           =   2304
   End
   Begin VB.Frame fmOption 
      Caption         =   "Option"
      Height          =   888
      Left            =   84
      TabIndex        =   7
      Top             =   2388
      Width           =   6312
      Begin VB.TextBox txtRecentMax 
         Height          =   288
         Left            =   2916
         TabIndex        =   8
         Top             =   336
         Width           =   432
      End
      Begin VB.Label lblRecentFile 
         Caption         =   "Numbers of Recent File List In Menu:"
         Height          =   288
         Left            =   180
         TabIndex        =   9
         Top             =   372
         Width           =   2976
      End
   End
   Begin MSComDlg.CommonDialog dlgOpenFile 
      Left            =   3264
      Top             =   2796
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame txtFrame 
      Caption         =   "Display"
      Height          =   2184
      Left            =   60
      TabIndex        =   2
      Top             =   72
      Width           =   6360
      Begin VB.ComboBox cboLineHeight 
         Height          =   288
         Left            =   4752
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1656
         Width           =   1365
      End
      Begin VB.CommandButton cmdBackColor 
         Caption         =   "BackColor"
         Height          =   375
         Left            =   3300
         TabIndex        =   5
         Top             =   1632
         Width           =   1155
      End
      Begin VB.CommandButton cmdForeColor 
         Caption         =   "ForeColor"
         Height          =   375
         Left            =   1740
         TabIndex        =   4
         Top             =   1632
         Width           =   1275
      End
      Begin VB.CommandButton cmdFont 
         Caption         =   "Font"
         Height          =   375
         Left            =   204
         TabIndex        =   3
         Top             =   1620
         Width           =   1245
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
      Height          =   324
      Left            =   5304
      TabIndex        =   1
      Top             =   3444
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Save"
      Height          =   312
      Left            =   3972
      TabIndex        =   0
      Top             =   3468
      Width           =   1095
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim tempVS As ViewerStyle


Private Sub cmdBackColor_Click()

        With dlgOpenFile
            .Flags = cdlCCRGBInit
            .Color = tempVS.BackColor
            .ShowColor
            tempVS.BackColor = .Color
        End With

        frmEx.BackColor = tempVS.BackColor
End Sub

Private Sub cmdCancel_Click()

    Unload Me

End Sub


Private Sub cmdClearRecent_Click()
With MainFrm.zhRecentFile
.clear
.SaveToIni MainFrm.zhtmIni
.FillinMenu MainFrm.mnuFile_Recent
End With
End Sub

Private Sub cmdFont_Click()

        Setfont dlgOpenFile, tempVS.Viewfont

        With dlgOpenFile
            .Flags = cdlCFBoth
            .ShowFont
            tempVS.Viewfont.Bold = .FontBold
            tempVS.Viewfont.Italic = .FontItalic
            tempVS.Viewfont.name = .FontName
            tempVS.Viewfont.Size = .FontSize
            tempVS.Viewfont.Strikethrough = .FontStrikethru
            tempVS.Viewfont.Underline = .FontUnderline
        End With

        Label1(0).Top = 0
        Dim i As Integer
        For i = 0 To 3
            Setfont Label1(i), tempVS.Viewfont
            Label1(i).Left = 0
        Next

        For i = 1 To 3
            Label1(i).Top = Label1(i - 1).Top + Label1(i - 1).Height * tempVS.LineHeight / 100
        Next

End Sub

Private Sub cmdForeColor_Click()
      With dlgOpenFile
            .Flags = cdlCCRGBInit
            .Color = tempVS.ForeColor
            .ShowColor
            tempVS.ForeColor = .Color
        End With
        Dim i As Integer
        For i = 0 To 3
            Label1(i).ForeColor = tempVS.ForeColor
            Label1(i).Left = 0
        Next

End Sub

Private Sub cmdOK_Click()

    SaveViewerStyle MainFrm.zhtmIni, tempVS
    
    MainFrm.zhRecentFile.maxItem = tempVS.RecentMax
    MainFrm.zhRecentFile.clear
    MainFrm.zhRecentFile.LoadFromIni MainFrm.zhtmIni
    MainFrm.zhRecentFile.FillinMenu MainFrm.mnuFile_Recent
        
    With MainFrm.rtxtView
        .BackColor = tempVS.BackColor
        CopytoIfont tempVS.Viewfont, .Font
        .Tag = CStr(tempVS.ForeColor)
        '.SelColor = tempVS.ForeColor
    End With

    MainFrm.RtxtView_ChangeAppearance
    'GetView zhrStatus.sCur_zhSubFile, MainFrm.IEView(MainFrm.iCurIEView)

End Sub


Private Sub cboLineheight_Click()

    tempVS.LineHeight = Val(cboLineHeight.Text)
    Dim i As Integer
    Label1(0).Top = 0

    For i = 1 To 3
        Label1(i).Top = Label1(i - 1).Top + Label1(i - 1).Height * tempVS.LineHeight / 100
    Next

End Sub



Private Sub Form_Load()
    
    loadFormStr Me
    '置中窗体
    Me.Icon = MainFrm.Icon
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
    GetViewerStyle MainFrm.zhtmIni, tempVS
    Dim i As Integer

    For i = 100 To 400 Step 10
        cboLineHeight.AddItem Str$(i) + "%"
    Next

    cboLineHeight.Text = Str$(tempVS.LineHeight) + "%"
    Label1(0).Top = 0

    For i = 0 To 3
        Setfont Label1(i), tempVS.Viewfont
        Label1(i).ForeColor = tempVS.ForeColor
        Label1(i).Left = 0
    Next

    For i = 1 To 3
        Label1(i).Top = Label1(i - 1).Top + Label1(i - 1).Height * tempVS.LineHeight / 100
    Next
    
    Label1(0).Caption = "天若有情天亦老凤凰台上凤凰游剑气萧瑟清澈山水秋风扫落叶天若有情天亦老凤凰台上凤凰游剑气萧瑟清澈山水秋风扫落叶天若有情天亦老凤凰台上凤凰游剑气萧瑟清澈山水秋风扫落叶"
    Label1(1).Caption = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"
    Label1(2).Caption = Label1(0).Caption
    Label1(3).Caption = Label1(1).Caption
    frmEx.BackColor = tempVS.BackColor


    txtRecentMax.Text = tempVS.RecentMax

End Sub



Private Sub txtRecentMax_Change()
    tempVS.RecentMax = Val(txtRecentMax.Text)
End Sub

