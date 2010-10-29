VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FTestImageList 
   Caption         =   "Test ImageList"
   ClientHeight    =   5844
   ClientLeft      =   1416
   ClientTop       =   2148
   ClientWidth     =   7716
   Icon            =   "TImage.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5844
   ScaleWidth      =   7716
   Begin VB.CheckBox chkOverlay 
      Caption         =   "Fix Overlay"
      Height          =   330
      Left            =   2625
      TabIndex        =   37
      Top             =   4935
      Width           =   1275
   End
   Begin VB.CheckBox chkPicture 
      Caption         =   "Picture"
      Height          =   330
      Left            =   2625
      TabIndex        =   32
      Top             =   4620
      Value           =   1  'Checked
      Width           =   1275
   End
   Begin VB.Frame frame 
      Caption         =   "Flags"
      Height          =   1260
      Left            =   4200
      TabIndex        =   27
      Top             =   4440
      Width           =   3360
      Begin VB.CheckBox chk 
         Caption         =   "Mask"
         Height          =   435
         Index           =   0
         Left            =   255
         TabIndex        =   31
         Top             =   255
         Width           =   1275
      End
      Begin VB.CheckBox chk 
         Caption         =   "Focus"
         Height          =   435
         Index           =   3
         Left            =   1905
         TabIndex        =   30
         Top             =   660
         Width           =   1275
      End
      Begin VB.CheckBox chk 
         Caption         =   "Selected"
         Height          =   435
         Index           =   2
         Left            =   1905
         TabIndex        =   29
         Top             =   255
         Width           =   1275
      End
      Begin VB.CheckBox chk 
         Caption         =   "Transparent"
         Height          =   435
         Index           =   1
         Left            =   255
         TabIndex        =   28
         Top             =   660
         Width           =   1275
      End
   End
   Begin VB.PictureBox pb 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   4020
      Left            =   120
      ScaleHeight     =   3972
      ScaleWidth      =   3960
      TabIndex        =   0
      Top             =   450
      Width           =   4008
      Begin VB.Label lblBmpDraw 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2124
         TabIndex        =   26
         Top             =   2670
         Width           =   1215
      End
      Begin VB.Label lblIconDraw 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   252
         Left            =   828
         TabIndex        =   25
         Top             =   2670
         Width           =   1212
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Draw"
         ForeColor       =   &H00FFFFFF&
         Height          =   252
         Index           =   12
         Left            =   2160
         TabIndex        =   24
         Top             =   2376
         Width           =   1212
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Picture"
         ForeColor       =   &H00FFFFFF&
         Height          =   252
         Index           =   11
         Left            =   864
         TabIndex        =   23
         Top             =   852
         Width           =   1212
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Overlay"
         ForeColor       =   &H00FFFFFF&
         Height          =   252
         Index           =   10
         Left            =   816
         TabIndex        =   22
         Top             =   1608
         Width           =   1212
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Draw"
         ForeColor       =   &H00FFFFFF&
         Height          =   252
         Index           =   9
         Left            =   828
         TabIndex        =   21
         Top             =   2376
         Width           =   1212
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "ExtractIcon"
         ForeColor       =   &H00FFFFFF&
         Height          =   252
         Index           =   8
         Left            =   2136
         TabIndex        =   20
         Top             =   48
         Width           =   1212
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Picture"
         ForeColor       =   &H00FFFFFF&
         Height          =   252
         Index           =   7
         Left            =   2124
         TabIndex        =   19
         Top             =   840
         Width           =   1212
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Overlay"
         ForeColor       =   &H00FFFFFF&
         Height          =   252
         Index           =   6
         Left            =   2100
         TabIndex        =   18
         Top             =   1620
         Width           =   1212
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "ExtractIcon"
         ForeColor       =   &H00FFFFFF&
         Height          =   252
         Index           =   5
         Left            =   840
         TabIndex        =   17
         Top             =   48
         Width           =   1212
      End
      Begin VB.Image imgBmpOverlay 
         Height          =   492
         Left            =   2124
         Top             =   1836
         Width           =   516
      End
      Begin VB.Image imgIconOverlay 
         Height          =   492
         Left            =   828
         Top             =   1836
         Width           =   516
      End
      Begin VB.Image imgIconIcon 
         Height          =   528
         Left            =   828
         Top             =   300
         Width           =   600
      End
      Begin VB.Image imgIconPic 
         Height          =   492
         Left            =   828
         Top             =   1080
         Width           =   576
      End
      Begin VB.Image imgBmpIcon 
         Height          =   492
         Left            =   2136
         Top             =   300
         Width           =   576
      End
      Begin VB.Image imgBmpPic 
         Height          =   492
         Left            =   2124
         Top             =   1080
         Width           =   516
      End
   End
   Begin MSComctlLib.ImageList imlstBmps 
      Left            =   930
      Top             =   4620
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   12632256
      ImageWidth      =   16
      ImageHeight     =   16
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TImage.frx":0CFA
            Key             =   "Camera"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TImage.frx":0E0C
            Key             =   "Spelling"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TImage.frx":0F1E
            Key             =   "Network"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TImage.frx":1030
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TImage.frx":1142
            Key             =   "Help"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlstIcons 
      Left            =   300
      Top             =   4635
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483633
      ImageWidth      =   32
      ImageHeight     =   32
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TImage.frx":1254
            Key             =   "Music"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TImage.frx":156E
            Key             =   "Globe"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TImage.frx":1888
            Key             =   "Recycle"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TImage.frx":1BA2
            Key             =   "Network"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TImage.frx":1EBC
            Key             =   "Desktop"
         EndProperty
      EndProperty
   End
   Begin VB.Label lbl 
      Caption         =   "Back Color"
      Height          =   330
      Index           =   13
      Left            =   6510
      TabIndex        =   36
      Top             =   3675
      Width           =   855
   End
   Begin VB.Label lbl 
      Caption         =   "Back Color"
      Height          =   330
      Index           =   3
      Left            =   4515
      TabIndex        =   35
      Top             =   3675
      Width           =   855
   End
   Begin VB.Label lblBmpsBack 
      BorderStyle     =   1  'Fixed Single
      Height          =   225
      Left            =   6195
      TabIndex        =   34
      Top             =   3675
      Width           =   225
   End
   Begin VB.Label lblIconsBack 
      BorderStyle     =   1  'Fixed Single
      Height          =   225
      Left            =   4200
      TabIndex        =   33
      Top             =   3675
      Width           =   225
   End
   Begin VB.Image imgBall 
      Height          =   840
      Left            =   1455
      Picture         =   "TImage.frx":21D6
      Stretch         =   -1  'True
      Top             =   4620
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.Label lblIconsMask 
      BorderStyle     =   1  'Fixed Single
      Height          =   225
      Left            =   4200
      TabIndex        =   16
      Top             =   4095
      Width           =   225
   End
   Begin VB.Label lblBmpsMask 
      BorderStyle     =   1  'Fixed Single
      Height          =   225
      Left            =   6195
      TabIndex        =   15
      Top             =   4095
      Width           =   225
   End
   Begin VB.Label lbl 
      Caption         =   "Mask Color"
      Height          =   330
      Index           =   4
      Left            =   4515
      TabIndex        =   14
      Top             =   4095
      Width           =   855
   End
   Begin VB.Label lbl 
      Caption         =   "Mask Color"
      Height          =   330
      Index           =   2
      Left            =   6510
      TabIndex        =   13
      Top             =   4095
      Width           =   855
   End
   Begin VB.Label lbl 
      Caption         =   "Bitmaps"
      Height          =   315
      Index           =   1
      Left            =   6180
      TabIndex        =   12
      Top             =   30
      Width           =   855
   End
   Begin VB.Label lbl 
      Caption         =   "Icons"
      Height          =   315
      Index           =   0
      Left            =   4215
      TabIndex        =   11
      Top             =   30
      Width           =   855
   End
   Begin VB.Label lblBmps 
      Height          =   330
      Index           =   5
      Left            =   6615
      TabIndex        =   10
      Top             =   2940
      Width           =   855
   End
   Begin VB.Label lblBmps 
      Height          =   330
      Index           =   4
      Left            =   6615
      TabIndex        =   9
      Top             =   2310
      Width           =   855
   End
   Begin VB.Label lblBmps 
      Height          =   330
      Index           =   3
      Left            =   6615
      TabIndex        =   8
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label lblBmps 
      Height          =   330
      Index           =   2
      Left            =   6615
      TabIndex        =   7
      Top             =   1050
      Width           =   855
   End
   Begin VB.Label lblIcons 
      Height          =   330
      Index           =   5
      Left            =   4935
      TabIndex        =   6
      Top             =   2970
      Width           =   855
   End
   Begin VB.Label lblIcons 
      Height          =   330
      Index           =   4
      Left            =   4935
      TabIndex        =   5
      Top             =   2340
      Width           =   855
   End
   Begin VB.Label lblIcons 
      Height          =   330
      Index           =   3
      Left            =   4935
      TabIndex        =   4
      Top             =   1710
      Width           =   855
   End
   Begin VB.Label lblIcons 
      Height          =   330
      Index           =   2
      Left            =   4935
      TabIndex        =   3
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label lblBmps 
      Height          =   330
      Index           =   1
      Left            =   6615
      TabIndex        =   2
      Top             =   420
      Width           =   855
   End
   Begin VB.Label lblIcons 
      Height          =   330
      Index           =   1
      Left            =   4935
      TabIndex        =   1
      Top             =   450
      Width           =   855
   End
   Begin VB.Image imgBmps 
      Height          =   330
      Index           =   5
      Left            =   6195
      Top             =   2940
      Width           =   330
   End
   Begin VB.Image imgBmps 
      Height          =   330
      Index           =   4
      Left            =   6195
      Top             =   2310
      Width           =   330
   End
   Begin VB.Image imgBmps 
      Height          =   330
      Index           =   3
      Left            =   6195
      Top             =   1680
      Width           =   330
   End
   Begin VB.Image imgBmps 
      Height          =   330
      Index           =   2
      Left            =   6195
      Top             =   1050
      Width           =   330
   End
   Begin VB.Image imgBmps 
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Index           =   1
      Left            =   6195
      Top             =   420
      Width           =   330
   End
   Begin VB.Image imgIcons 
      Height          =   540
      Index           =   5
      Left            =   4200
      Top             =   2955
      Width           =   645
   End
   Begin VB.Image imgIcons 
      Height          =   540
      Index           =   4
      Left            =   4200
      Top             =   2340
      Width           =   645
   End
   Begin VB.Image imgIcons 
      Height          =   540
      Index           =   3
      Left            =   4200
      Top             =   1710
      Width           =   645
   End
   Begin VB.Image imgIcons 
      Height          =   540
      Index           =   2
      Left            =   4200
      Top             =   1080
      Width           =   645
   End
   Begin VB.Image imgIcons 
      BorderStyle     =   1  'Fixed Single
      Height          =   540
      Index           =   1
      Left            =   4200
      Top             =   450
      Width           =   645
   End
End
Attribute VB_Name = "FTestImageList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private iIcons As Integer
Private iBmps As Integer
Private iIconsLast As Integer
Private iBmpsLast As Integer
Private afDisplay As Long

Private Sub chkOverlay_Click()
    DrawIcons
End Sub

Private Sub Form_Load()
    Dim v As Variant
    For Each v In imlstIcons.ListImages
        imgIcons(v.Index).Picture = v.Picture
        lblIcons(v.Index).Caption = v.Key
    Next
    iIcons = 1
    iIconsLast = 5
    imlstIcons.BackColor = pb.BackColor
    imlstIcons.MaskColor = pb.BackColor
    lblIconsMask.BackColor = imlstIcons.MaskColor
    
    For Each v In imlstBmps.ListImages
        imgBmps(v.Index).Picture = v.Picture
        lblBmps(v.Index).Caption = v.Key
    Next
    iBmps = 1
    iBmpsLast = 5
    imlstBmps.BackColor = pb.BackColor
    imlstBmps.MaskColor = pb.BackColor
    lblBmpsMask.BackColor = imlstBmps.MaskColor
    
    Show
    chkPicture_Click
    DrawIcons
    DrawBmps
End Sub

Private Sub imgIcons_Click(Index As Integer)
    imgIcons(iIcons).BorderStyle = vbTransparent
    iIconsLast = iIcons
    iIcons = Index
    imgIcons(iIcons).BorderStyle = vbBSSolid
    DrawIcons
End Sub

Private Sub imgBmps_Click(Index As Integer)
    imgBmps(iBmps).BorderStyle = vbTransparent
    iBmpsLast = iBmps
    iBmps = Index
    imgBmps(iBmps).BorderStyle = vbBSSolid
    DrawBmps
End Sub

Private Sub lblIconsMask_MouseUp(Button As Integer, Shift As Integer, _
                                 X As Single, Y As Single)
    Dim getclr As New CColorPicker
    getclr.Color = lblIconsMask.BackColor
    getclr.Load Left + lblIconsMask.Left + X, Top + lblIconsMask.Top + Y
    imlstIcons.BackColor = getclr.Color
    lblIconsMask.BackColor = getclr.Color
    imgIcons_Click iIcons
End Sub

Private Sub lblBmpsMask_MouseUp(Button As Integer, Shift As Integer, _
                                X As Single, Y As Single)
    Dim getclr As New CColorPicker, clr As Long
    clr = imlstIcons.MaskColor
    getclr.Color = clr
    getclr.Load Left + lblBmpsMask.Left + X, Top + lblBmpsMask.Top + Y
    clr = getclr.Color
    lblBmpsMask.BackColor = clr
    imlstBmps.MaskColor = clr
    imgBmps_Click iBmps
End Sub

Private Sub lblIconsBack_MouseUp(Button As Integer, Shift As Integer, _
                                 X As Single, Y As Single)
    Dim getclr As New CColorPicker
    getclr.Color = lblIconsBack.BackColor
    getclr.Load Left + lblIconsBack.Left + X, Top + lblIconsBack.Top + Y
    imlstIcons.BackColor = getclr.Color
    lblIconsBack.BackColor = getclr.Color
    imgIcons_Click iIcons
End Sub

Private Sub lblBmpsBack_MouseUp(Button As Integer, Shift As Integer, _
                                X As Single, Y As Single)
    Dim getclr As New CColorPicker, clr As Long
    clr = imlstIcons.BackColor
    getclr.Color = clr
    getclr.Load Left + lblBmpsBack.Left + X, Top + lblBmpsBack.Top + Y
    clr = getclr.Color
    lblBmpsBack.BackColor = clr
    imlstBmps.BackColor = clr
    imgBmps_Click iBmps
End Sub

Private Sub Draw(afStyle As Long)
    
    Dim X As Long, Y As Long, dxy As Long
    pb.Cls
    X = lblIconDraw.Left
    Y = lblIconDraw.Top
    DrawImage imlstIcons, iIcons, pb.hDC, X, Y, afStyle
    dxy = imlstIcons.ImageHeight * Screen.TwipsPerPixelY
    X = lblBmpDraw.Left
    Y = lblBmpDraw.Top
    DrawImage imlstBmps, iBmps, pb.hDC, X, Y, afStyle

End Sub

Private Sub chkPicture_Click()
    If chkPicture.Value = vbChecked Then
        pb.Picture = imgBall.Picture
    Else
        pb.Picture = Nothing
    End If
End Sub

Private Sub chk_Click(Index As Integer)
    Select Case Index
    Case 0  ' Transparent
        If chk(Index).Value = vbChecked Then
            afDisplay = afDisplay Or ILD_MASK
        Else
            afDisplay = afDisplay And Not ILD_MASK
        End If
    Case 1 ' Mask
        If chk(Index).Value = vbChecked Then
            afDisplay = afDisplay Or ILD_TRANSPARENT
        Else
            afDisplay = afDisplay And Not ILD_TRANSPARENT
        End If
    Case 2 ' Selected
        If chk(Index).Value = vbChecked Then
            afDisplay = afDisplay Or ILD_SELECTED
        Else
            afDisplay = afDisplay And Not ILD_SELECTED
        End If
    Case 3 ' Focus
        If chk(Index).Value = vbChecked Then
            afDisplay = afDisplay Or ILD_FOCUS
        Else
            afDisplay = afDisplay And Not ILD_FOCUS
        End If
    End Select
    Draw afDisplay
End Sub

Sub DrawIcons()
    imgIconIcon.Picture = imlstIcons.ListImages(iIcons).ExtractIcon
    imgIconPic.Picture = imlstIcons.ListImages(iIcons).Picture
    With imlstIcons
        If chkOverlay.Value <> vbChecked Then
            ' Overlay without bug fix
            imgIconOverlay.Picture = .Overlay(iIconsLast, iIcons)
        Else
            ' Save old background and mask color
            Dim clrBack As Long, clrMask As Long
            clrBack = .BackColor: clrMask = .MaskColor
            ' Set color that does not occur in image
            .BackColor = vbMagenta: .MaskColor = vbMagenta
            ' Insert overlay, extract as icon, remove, and restore color
            .ListImages.Add 1, , .Overlay(iIconsLast, iIcons)
            imgIconOverlay.Picture = .ListImages(1).ExtractIcon
            .ListImages.Remove 1
            .BackColor = clrBack: .MaskColor = clrMask
        End If
    End With
    Draw afDisplay
End Sub

Sub DrawBmps()
    imgBmpIcon.Picture = imlstBmps.ListImages(iBmps).ExtractIcon
    imgBmpPic.Picture = imlstBmps.ListImages(iBmps).Picture
    imgBmpOverlay.Picture = imlstBmps.Overlay(iBmpsLast, iBmps)
    Draw afDisplay
End Sub

