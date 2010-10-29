VERSION 5.00
Begin VB.Form FTestEdges 
   Caption         =   "Test Edges"
   ClientHeight    =   3948
   ClientLeft      =   1260
   ClientTop       =   1608
   ClientWidth     =   5028
   Icon            =   "TEdge.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   329
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   419
   Begin VB.CheckBox chkSt 
      Caption         =   "Soft"
      Height          =   435
      Left            =   2265
      MaskColor       =   &H00000000&
      TabIndex        =   13
      Top             =   960
      Width           =   1230
   End
   Begin VB.CheckBox chkAj 
      Caption         =   "Adjust"
      Height          =   435
      Left            =   2265
      MaskColor       =   &H00000000&
      TabIndex        =   12
      Top             =   1380
      Width           =   1230
   End
   Begin VB.CheckBox chkFt 
      Caption         =   "Flat"
      Height          =   435
      Left            =   2265
      MaskColor       =   &H00000000&
      TabIndex        =   11
      Top             =   1800
      Width           =   1230
   End
   Begin VB.CheckBox chkMo 
      Caption         =   "Monochrome"
      Height          =   435
      Left            =   2265
      MaskColor       =   &H00000000&
      TabIndex        =   10
      Top             =   2220
      Width           =   1230
   End
   Begin VB.CheckBox chkTp 
      Caption         =   "Top"
      Height          =   435
      Left            =   3735
      MaskColor       =   &H00000000&
      TabIndex        =   9
      Top             =   960
      Value           =   1  'Checked
      Width           =   990
   End
   Begin VB.CheckBox chkRt 
      Caption         =   "Right"
      Height          =   435
      Left            =   3735
      MaskColor       =   &H00000000&
      TabIndex        =   8
      Top             =   1380
      Value           =   1  'Checked
      Width           =   990
   End
   Begin VB.CheckBox chkBt 
      Caption         =   "Bottom"
      Height          =   435
      Left            =   3735
      MaskColor       =   &H00000000&
      TabIndex        =   7
      Top             =   1800
      Value           =   1  'Checked
      Width           =   990
   End
   Begin VB.CheckBox chkMd 
      Caption         =   "Fill Middle"
      Height          =   435
      Left            =   2265
      MaskColor       =   &H00000000&
      TabIndex        =   6
      Top             =   540
      Width           =   1230
   End
   Begin VB.CheckBox chkDg 
      Caption         =   "Diagonal"
      Height          =   435
      Left            =   3735
      MaskColor       =   &H00000000&
      TabIndex        =   5
      Top             =   2220
      Width           =   990
   End
   Begin VB.CheckBox chkSI 
      Caption         =   "Sunken Inner"
      Height          =   435
      Left            =   180
      MaskColor       =   &H00000000&
      TabIndex        =   4
      Top             =   1800
      Width           =   1560
   End
   Begin VB.CheckBox chkRI 
      Caption         =   "Raised Inner"
      Height          =   435
      Left            =   180
      MaskColor       =   &H00000000&
      TabIndex        =   3
      Top             =   1380
      Value           =   1  'Checked
      Width           =   1515
   End
   Begin VB.CheckBox chkSO 
      Caption         =   "Sunken Outer"
      Height          =   435
      Left            =   180
      MaskColor       =   &H00000000&
      TabIndex        =   2
      Top             =   960
      Width           =   1530
   End
   Begin VB.CheckBox chkLf 
      Caption         =   "Left"
      Height          =   435
      Left            =   3735
      MaskColor       =   &H00000000&
      TabIndex        =   1
      Top             =   540
      Value           =   1  'Checked
      Width           =   990
   End
   Begin VB.CheckBox chkRO 
      Caption         =   "Raised Outer"
      Height          =   435
      Left            =   180
      MaskColor       =   &H00000000&
      TabIndex        =   0
      Top             =   540
      Value           =   1  'Checked
      Width           =   1545
   End
   Begin VB.Label lblButton 
      BackStyle       =   0  'Transparent
      Caption         =   "                                          &Click Me"
      Height          =   588
      Left            =   360
      TabIndex        =   16
      Top             =   2520
      Width           =   1260
   End
   Begin VB.Label lblBigOption 
      BackStyle       =   0  'Transparent
      Caption         =   "        Reset"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   312
      Index           =   1
      Left            =   3444
      TabIndex        =   19
      Top             =   3468
      Width           =   1740
   End
   Begin VB.Label lblBigOption 
      BackStyle       =   0  'Transparent
      Caption         =   "        Change"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   312
      Index           =   0
      Left            =   3444
      TabIndex        =   18
      Top             =   3000
      Width           =   1740
   End
   Begin VB.Label lblStyle 
      Caption         =   "Style flags:"
      Height          =   330
      Left            =   2265
      TabIndex        =   15
      Top             =   240
      Width           =   2475
   End
   Begin VB.Label lblBorder 
      Caption         =   "Border flags: "
      Height          =   285
      Left            =   180
      TabIndex        =   14
      Top             =   240
      Width           =   1530
   End
   Begin VB.Label lblBigCheck 
      BackStyle       =   0  'Transparent
      Caption         =   "        Cool"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1770
      TabIndex        =   17
      Top             =   3000
      Width           =   1740
   End
End
Attribute VB_Name = "FTestEdges"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private afOpFlags As Long
Private hLSysImages As Long, cLSysImages As Long
Private hSSysImages As Long, cSSysImages As Long
    
Private chk As RECT, afChk As Long
Private radChange As RECT, afRadChange As Long
Private radReset As RECT, afRadReset As Long
Private cmd As RECT
Private afBorder As Long, afBorderT As Long, afStyle As Long

Private Sub Form_Load()
    
    InitAll
    
End Sub

Private Sub InitAll()

    InitButton
    
    ' Fake check box
    With lblBigCheck
        chk.Left = .Left
        chk.Top = .Top
        chk.Right = .Left + .Height
        chk.bottom = .Top + .Height
    End With
    afChk = DFCS_BUTTONCHECK Or DFCS_CHECKED
    DrawFrameControl hDC, chk, DFC_BUTTON, afChk
    
    ' Fake option buttons
    With lblBigOption(0)
        radChange.Left = .Left
        radChange.Top = .Top
        radChange.Right = .Left + .Height
        radChange.bottom = .Top + .Height
    End With
    afRadChange = DFCS_BUTTONRADIO Or DFCS_CHECKED
    DrawFrameControl hDC, radChange, DFC_BUTTON, afRadChange
    With lblBigOption(1)
        radReset.Left = .Left
        radReset.Top = .Top
        radReset.Right = .Left + .Height
        radReset.bottom = .Top + .Height
    End With
    afRadReset = DFCS_BUTTONRADIO
    DrawFrameControl hDC, radReset, DFC_BUTTON, afRadReset
    
End Sub

Private Sub Form_Paint()
    DrawFrameControl hDC, chk, DFC_BUTTON, afChk
    DrawFrameControl hDC, radChange, DFC_BUTTON, afRadChange
    DrawFrameControl hDC, radReset, DFC_BUTTON, afRadReset
    DrawEdge hDC, cmd, afBorder, afStyle
End Sub

Private Sub chkRO_Click()
    UpdateButton
End Sub

Private Sub chkSO_Click()
    UpdateButton
End Sub

Private Sub chkRI_Click()
    UpdateButton
End Sub

Private Sub chkSI_Click()
    UpdateButton
End Sub

Private Sub chkLf_Click()
    UpdateButton
End Sub

Private Sub chkTp_Click()
    UpdateButton
End Sub

Private Sub chkRt_Click()
    UpdateButton
End Sub

Private Sub chkBt_Click()
    UpdateButton
End Sub

Private Sub chkDg_Click()
    UpdateButton
End Sub

Private Sub chkMd_Click()
    UpdateButton
End Sub

Private Sub chkSt_Click()
    UpdateButton
End Sub

Private Sub chkAj_Click()
    UpdateButton
End Sub

Private Sub chkFt_Click()
    UpdateButton
End Sub

Private Sub chkMo_Click()
    UpdateButton
End Sub

Private Sub lblBigCheck_Click()
    If afChk <> (DFCS_BUTTONCHECK Or DFCS_CHECKED) Then
        afChk = DFCS_BUTTONCHECK Or DFCS_CHECKED
    Else
        afChk = DFCS_BUTTONCHECK
    End If
    DrawFrameControl hDC, chk, DFC_BUTTON, afChk
End Sub

Private Sub lblBigOption_Click(Index As Integer)
    If Index = 0 Then
        If afRadChange <> (DFCS_BUTTONRADIO Or DFCS_CHECKED) Then
            afRadChange = DFCS_BUTTONRADIO Or DFCS_CHECKED
            afRadReset = DFCS_BUTTONRADIO
        Else
            afRadChange = DFCS_BUTTONRADIO
            afRadReset = DFCS_BUTTONRADIO Or DFCS_CHECKED
            InitButton
        End If
    Else
        If afRadReset <> (DFCS_BUTTONRADIO Or DFCS_CHECKED) Then
            afRadReset = DFCS_BUTTONRADIO Or DFCS_CHECKED
            afRadChange = DFCS_BUTTONRADIO
            InitButton
        Else
            afRadReset = DFCS_BUTTONRADIO
            afRadChange = DFCS_BUTTONRADIO Or DFCS_CHECKED
        End If
    End If
    DrawFrameControl hDC, radChange, DFC_BUTTON, afRadChange
    DrawFrameControl hDC, radReset, DFC_BUTTON, afRadReset
End Sub

Private Sub lblButton_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ToggleButton False
    lblBigCheck_Click
End Sub

Private Sub lblButton_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ToggleButton True
End Sub

Sub UpdateButton()
    If chkRO.Value = vbChecked Then
        afBorder = afBorder Or BDR_RAISEDOUTER
    Else
        afBorder = afBorder And Not BDR_RAISEDOUTER
    End If
    If chkSO.Value = vbChecked Then
        afBorder = afBorder Or BDR_SUNKENOUTER
    Else
        afBorder = afBorder And Not BDR_SUNKENOUTER
    End If
    If chkRI.Value = vbChecked Then
        afBorder = afBorder Or BDR_RAISEDINNER
    Else
        afBorder = afBorder And Not BDR_RAISEDINNER
    End If
    If chkSI.Value = vbChecked Then
        afBorder = afBorder Or BDR_SUNKENINNER
    Else
        afBorder = afBorder And Not BDR_SUNKENINNER
    End If
    If chkLf.Value = vbChecked Then
        afStyle = afStyle Or BF_LEFT
    Else
        afStyle = afStyle And Not BF_LEFT
    End If
    If chkTp.Value = vbChecked Then
        afStyle = afStyle Or BF_TOP
    Else
        afStyle = afStyle And Not BF_TOP
    End If
    If chkRt.Value = vbChecked Then
        afStyle = afStyle Or BF_RIGHT
    Else
        afStyle = afStyle And Not BF_RIGHT
    End If
    If chkBt.Value = vbChecked Then
        afStyle = afStyle Or BF_BOTTOM
    Else
        afStyle = afStyle And Not BF_BOTTOM
    End If
    If chkDg.Value = vbChecked Then
        afStyle = afStyle Or BF_DIAGONAL
    Else
        afStyle = afStyle And Not BF_DIAGONAL
    End If
    If chkMd.Value = vbChecked Then
        afStyle = afStyle Or BF_MIDDLE
    Else
        afStyle = afStyle And Not BF_MIDDLE
    End If
    If chkSt.Value = vbChecked Then
        afStyle = afStyle Or BF_SOFT
    Else
        afStyle = afStyle And Not BF_SOFT
    End If
    If chkAj.Value = vbChecked Then
        afStyle = afStyle Or BF_ADJUST
    Else
        afStyle = afStyle And Not BF_ADJUST
    End If
    If chkFt.Value = vbChecked Then
        afStyle = afStyle Or BF_FLAT
    Else
        afStyle = afStyle And Not BF_FLAT
    End If
    If chkMo.Value = vbChecked Then
        afStyle = afStyle Or BF_MONO
    Else
        afStyle = afStyle And Not BF_MONO
    End If
    lblBorder.Caption = "Border flags: &&H" & FmtHex(afBorder, 4)
    lblStyle.Caption = "Edge flags: &&H" & FmtHex(afStyle, 8)
    DrawEdge hDC, cmd, afBorder, afStyle
    Refresh
End Sub

Sub ToggleButton(ByVal fUp As Boolean)
    If fUp Then
        afBorder = afBorderT
    Else
        afBorderT = afBorder
        afBorder = (Not afBorder) And &HF
    End If
    lblBorder.Caption = "Border: &&H" & FmtHex(afBorder, 8)
    DrawEdge hDC, cmd, afBorder, afStyle
End Sub

Sub InitButton()
    ' Fake button
    With lblButton
        cmd.Left = .Left
        cmd.Top = .Top
        cmd.Right = .Left + .Width
        cmd.bottom = .Top + .Height
    End With
    chkRO.Value = vbChecked
    chkSO.Value = vbUnchecked
    chkRI.Value = vbChecked
    chkSI.Value = vbUnchecked
    chkLf.Value = vbChecked
    chkTp.Value = vbChecked
    chkRt.Value = vbChecked
    chkBt.Value = vbChecked
    chkDg.Value = vbUnchecked
    chkSt.Value = vbUnchecked
    chkAj.Value = vbUnchecked
    chkFt.Value = vbUnchecked
    chkMo.Value = vbUnchecked
    UpdateButton
End Sub


