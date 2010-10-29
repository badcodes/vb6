VERSION 5.00
Object = "{35DFF7C3-DEB3-11D0-8C50-00C04FC29CEC}#2.0#0"; "PictureGlass.ocx"
Begin VB.Form FFun 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fun 'n Games"
   ClientHeight    =   6660
   ClientLeft      =   1908
   ClientTop       =   2076
   ClientWidth     =   6708
   ClipControls    =   0   'False
   FillColor       =   &H00FF00FF&
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   7.8
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Fun.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   6708
   Begin PictureGlass.XPictureGlass pgControl 
      Height          =   1095
      Left            =   2400
      TabIndex        =   15
      Top             =   2040
      Width           =   1215
      _ExtentX        =   2138
      _ExtentY        =   1926
      BackColor       =   -2147483628
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ScaleWidth      =   1215
      ScaleMode       =   0
      ScaleHeight     =   1095
   End
   Begin VB.CommandButton cmdLargeBmp 
      Caption         =   "Large"
      Height          =   384
      Left            =   696
      TabIndex        =   14
      Top             =   5688
      Width           =   945
   End
   Begin VB.CommandButton cmdAniBmp 
      Caption         =   "Animate"
      Height          =   384
      Left            =   696
      TabIndex        =   13
      Top             =   6168
      Width           =   945
   End
   Begin VB.CommandButton cmdSmallBmp 
      Caption         =   "Small"
      Height          =   384
      Left            =   696
      TabIndex        =   12
      Top             =   5196
      Width           =   945
   End
   Begin VB.CheckBox chkAutoRedraw 
      Caption         =   "Auto Redraw"
      Height          =   360
      Left            =   4968
      TabIndex        =   11
      Top             =   6228
      Width           =   1575
   End
   Begin VB.CommandButton cmdAnimateCtl 
      Caption         =   "Animate &Control"
      Height          =   375
      Left            =   4968
      TabIndex        =   10
      Top             =   1533
      Width           =   1575
   End
   Begin VB.CommandButton cmdSpiralBmp 
      Caption         =   "Spiral Bitmap"
      Height          =   375
      Left            =   4968
      TabIndex        =   9
      Top             =   581
      Width           =   1575
   End
   Begin VB.CheckBox chkClip 
      Caption         =   "Clip Controls"
      Height          =   468
      Left            =   4968
      TabIndex        =   8
      Top             =   5892
      Value           =   1  'Checked
      Width           =   1575
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear Surface"
      Height          =   375
      Left            =   4968
      TabIndex        =   7
      Top             =   3913
      Width           =   1575
   End
   Begin VB.CommandButton cmdAnimate 
      Caption         =   "&Animate Picture"
      Height          =   375
      Left            =   4968
      TabIndex        =   6
      Top             =   1057
      Width           =   1575
   End
   Begin VB.CommandButton cmdStar 
      Caption         =   "Stars"
      Height          =   375
      Left            =   4968
      TabIndex        =   5
      Top             =   2009
      Width           =   1575
   End
   Begin VB.CommandButton cmdSort 
      Caption         =   "Sort Cards"
      Height          =   375
      Left            =   4968
      TabIndex        =   4
      Top             =   2961
      Width           =   1575
   End
   Begin VB.Timer tmrAnimate 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   240
      Top             =   6885
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Card Backs"
      Height          =   375
      Left            =   4968
      TabIndex        =   3
      Top             =   3437
      Width           =   1575
   End
   Begin VB.CommandButton cmdShuffle 
      Caption         =   "Shuffle Cards"
      Height          =   375
      Left            =   4968
      TabIndex        =   2
      Top             =   2485
      Width           =   1575
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   4968
      TabIndex        =   1
      Top             =   4380
      Width           =   1575
   End
   Begin VB.CommandButton cmdBmpSpiral 
      Caption         =   "Bitmap Spiral"
      Height          =   375
      Left            =   4968
      TabIndex        =   0
      Top             =   105
      Width           =   1575
   End
   Begin VB.Label lblMaskColorText 
      Caption         =   "Mask Color"
      Height          =   228
      Left            =   5280
      TabIndex        =   19
      Top             =   5664
      Width           =   1032
   End
   Begin VB.Label lblMaskColor 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   228
      Left            =   4956
      TabIndex        =   18
      Top             =   5664
      Width           =   228
   End
   Begin VB.Label lblFormColor 
      BorderStyle     =   1  'Fixed Single
      Height          =   228
      Left            =   4968
      TabIndex        =   17
      Top             =   5316
      Width           =   228
   End
   Begin VB.Label lblFormColorText 
      Caption         =   "Form Color"
      Height          =   252
      Left            =   5280
      TabIndex        =   16
      Top             =   5328
      Width           =   1116
   End
   Begin VB.Image imgAniBmp 
      BorderStyle     =   1  'Fixed Single
      Height          =   384
      Left            =   120
      Picture         =   "Fun.frx":0CFA
      Stretch         =   -1  'True
      Top             =   6168
      Width           =   468
   End
   Begin VB.Image imgLargeBmp 
      BorderStyle     =   1  'Fixed Single
      Height          =   384
      Left            =   120
      Picture         =   "Fun.frx":10C6
      Stretch         =   -1  'True
      Top             =   5688
      Width           =   456
   End
   Begin VB.Image imgSmallBmp 
      BorderStyle     =   1  'Fixed Single
      Height          =   384
      Left            =   120
      Picture         =   "Fun.frx":11508
      Stretch         =   -1  'True
      Top             =   5196
      Width           =   456
   End
   Begin VB.Image imgBig 
      Height          =   3072
      Left            =   9552
      Picture         =   "Fun.frx":1178A
      Top             =   2568
      Visible         =   0   'False
      Width           =   3072
   End
   Begin VB.Image imgLittle 
      Height          =   384
      Left            =   12288
      Picture         =   "Fun.frx":1980C
      Top             =   1068
      Visible         =   0   'False
      Width           =   384
   End
End
Attribute VB_Name = "FFun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

 Private dxCard As Long
Private dyCard As Long

Private xInc As Single, yInc As Single
Private xMove As Single, yMove As Single
Private aCards(0 To 51) As Variant
Private clrMask As Long

Private pgPicture As New CPictureGlass

Enum EAnimateType
    eatNone
    eatCardBacks
    eatPicture
    eatControl
End Enum

Private eatAnimate As EAnimateType

Private WithEvents timerAnimate As CTimer
Attribute timerAnimate.VB_VarHelpID = -1

Private Sub Form_Load()
    
    ChDir App.Path
           
    Set timerAnimate = New CTimer
    SetClipControls Me.hWnd, chkClip.Value = vbChecked
    clrMask = vbWhite
    ' Initialize CARDS library and array
    On Error GoTo CardsLoadFail
    If cdtInit(dxCard, dyCard) = 0 Then
        GoTo CardsLoadFail
        cmdShuffle.Enabled = False
        cmdBack.Enabled = False
    Else
        Dim i As Integer
        For i = 0 To 51
            aCards(i) = i
        Next
        ShuffleArray aCards()
    End If
    Exit Sub
    
CardsLoadFail:
    cmdSort.Enabled = False
    cmdShuffle.Enabled = False
    cmdBack.Enabled = False
    
End Sub

Private Sub chkClip_Click()
    SetClipControls Me.hWnd, chkClip.Value = vbChecked
End Sub

Private Sub chkAutoRedraw_Click()
    AutoRedraw = (chkAutoRedraw = vbChecked)
End Sub

Private Sub cmdAnimate_Click()
    If cmdAnimate.Caption = "&Animate Picture" Then
        With pgPicture
            ' Draw picture on center of form with white background
            .Create Me, imgAniBmp.Picture, clrMask, Width / 2, Height / 2
            ' Constant controls pace, sign controls direction
            xInc = .Width * 0.05
            yInc = -.Height * 0.05
        End With
        SetTimer eatPicture
        cmdAnimate.Caption = "Stop &Animate"
    Else
        SetTimer eatNone
        cmdAnimate.Caption = "&Animate Picture"
    End If
End Sub

Private Sub cmdAnimateCtl_Click()
    If cmdAnimateCtl.Caption = "Animate &Control" Then
        With pgControl
            Set .Picture = imgAniBmp.Picture
            .MaskColor = clrMask
            .Left = (ScaleWidth / 2) - (.Width / 2)
            .Top = (ScaleHeight / 2) - (.Height / 2)
            ' Constant controls pace, sign controls direction
            xInc = .Width * 0.05
            yInc = -.Height * 0.05
            .Visible = True
        End With
        SetTimer eatControl
        cmdAnimateCtl.Caption = "Stop &Animate"
    Else
        SetTimer eatNone
        cmdAnimateCtl.Caption = "Animate &Control"
    End If
End Sub

Private Sub cmdBack_Click()
    Dim ordScale As Integer
    ordScale = ScaleMode: ScaleMode = vbPixels
    SetTimer eatCardBacks
    Cls
    Dim X As Integer, Y As Integer, ecbBack As ECardBack
    ecbBack = ecbCrossHatch  ' First card back
    ' Draw cards in 4 by 4 grid
    For X = 0 To 3
        For Y = 0 To 3
            cdtDraw Me.hDC, (dxCard * 0.1) + (X * dxCard * 1.1), _
                    (dyCard * 0.1) + (Y * dyCard * 1.1), _
                    ecbBack, ectBacks, QBColor(Random(0, 15))
            ecbBack = ecbBack + 1
        Next
    Next
    ScaleMode = ordScale
End Sub

' Secret undocumented command for animating card backs a click at a time
Private Sub cmdBack_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Static f As Boolean
        If f = False Then
            cmdBack_Click
            f = True
            SetTimer eatNone
        End If
        AnimateBacks
    End If
    
End Sub

Private Sub cmdBmpSpiral_Click()
    SetTimer eatNone
    BmpSpiral Me, imgSmallBmp.Picture
    imgSmallBmp.Refresh
    imgLargeBmp.Refresh
    imgAniBmp.Refresh
End Sub

Private Sub cmdSpiralBmp_Click()
    SetTimer eatNone
    Dim X As Long, Y As Long
    X = ((Width - (cmdSpiralBmp.Width * 1.2)) / 2) - (ScaleX(imgLargeBmp.Picture.Width) / 2)
    Y = ((Height - (cmdSpiralBmp.Height * 1.2)) / 2) - (ScaleY(imgLargeBmp.Picture.Height) / 2)
    SpiralBmp Me, imgLargeBmp.Picture, X, Y
End Sub

Private Sub cmdClear_Click()
    SetTimer eatNone
    Cls
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSmallBmp_Click()
    Call FillPicture(imgSmallBmp)
    cmdClear_Click
End Sub

Private Sub cmdLargeBmp_Click()
    Call FillPicture(imgLargeBmp)
    cmdClear_Click
End Sub

Private Sub cmdAniBmp_Click()
    Call FillPicture(imgAniBmp)
    cmdClear_Click
End Sub

Private Sub cmdStar_Click()
    SetTimer eatNone
    Dim i As Integer, xMid As Long, yMid As Long, dxyRadius As Long
    For i = 1 To Random(5, 20)
        dxyRadius = Random(Height / 8, Height / 4)
        xMid = Random(1, Width): yMid = Random(1, Height)
        ' Black border and two random colors
        Star Me, xMid, yMid, dxyRadius, vbBlack, _
             QBColor(Random(1, 15)), QBColor(Random(1, 15))
        ' Black border and one random color
        'Star Me, xMid, yMid, dxyRadius, vbBlack, QBColor (Random(1, 15))
        ' One filled color
        'Star Me, xMid, yMid, dxyRadius, QBColor(Random(1, 15))
    Next
    imgSmallBmp.Refresh
    imgLargeBmp.Refresh
    imgAniBmp.Refresh
End Sub


Private Function FillPicture(pic As Control) As Boolean
    Dim opfile As New COpenPictureFile
    With opfile
        .InitDir = App.Path
        .FilterType = efpBitmap
        .Load Left + (Width * 0.3), Top + (Height * 0.4)
        If .FileName <> sEmpty And .picType = vbPicTypeBitmap Then
            pic.Picture = LoadPicture(.FileName)
        End If
    End With
End Function

Private Sub cmdShuffle_Click()
    SetTimer eatNone
    Dim ordScale As Integer
    ordScale = ScaleMode: ScaleMode = vbPixels
    ShuffleArray aCards()
    ShowCards aCards()
    ScaleMode = ordScale
End Sub

Private Sub cmdSort_Click()
    SetTimer eatNone
    Dim ordScale As Integer
    ordScale = ScaleMode: ScaleMode = vbPixels
    SortArray aCards(), 0, 51
    ShowCards aCards()
    ScaleMode = ordScale
End Sub

Private Sub Form_Paint()
    If eatAnimate = eatControl Then pgControl.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim f As Integer
    If cmdSort.Enabled Then cdtTerm
End Sub

Private Sub imgAniBmp_Click()
    cmdAniBmp_Click
End Sub

Private Sub imgLargeBmp_Click()
    cmdLargeBmp_Click
End Sub

Private Sub imgSmallBmp_Click()
    cmdSmallBmp_Click
End Sub

Sub AnimateBacks()
    Static X As Integer, Y As Integer
    Static ecbBack As ECardBack, iState As Integer
    
    ' Save scale mode and change to pixels
    Dim ordScale As Integer
    ordScale = ScaleMode: ScaleMode = vbPixels
    
    ' Adjust variables
    If ecbBack < ecbCrossHatch Or ecbBack > ecbO Then
        ecbBack = ecbCrossHatch
        X = 0: Y = 0
    End If
    If X = 4 Then X = 0
    If Y = 4 Then Y = 0: X = X + 1
    Select Case ecbBack
    Case ecbCrossHatch
        ' Change color of crosshatch
        cdtDraw Me.hDC, (dxCard * 0.1) + (X * dxCard * 1.1), _
                (dyCard * 0.1) + (Y * dyCard * 1.1), _
                ecbBack, ectBacks, QBColor(Random(0, 15))
    Case Else 'ecbRobot, ecbCastle, ecbBeach, ecbCardHand
        ' Step through animation states
        If cdtAnimate(Me.hDC, ecbBack, _
                      (dxCard * 0.1) + (X * dxCard * 1.1), _
                      (dyCard * 0.1) + (Y * dyCard * 1.1), iState) Then
            iState = iState + 1
            ScaleMode = ordScale
            Exit Sub    ' Don't move to next card until final state
        End If
        iState = 0
    ' Case Else
        ' Ignore other cards
    End Select
    ' Move to next card
    ecbBack = ecbBack + 1
    Y = Y + 1
    ' Restore
    ScaleMode = ordScale
End Sub

Private Sub AnimatePicture()
    With pgPicture
        If .Left + .Width > ScaleWidth Then xInc = -xInc
        If .Left <= Abs(xInc) Then xInc = -xInc
        If .Top + .Height > ScaleHeight Then yInc = -yInc
        If .Top <= Abs(yInc) Then yInc = -yInc
        .Move .Left + xInc, .Top + yInc
    End With
End Sub

Private Sub AnimateControl()
    With pgControl
        If .Left + .Width > ScaleWidth Then xInc = -xInc
        If .Left <= Abs(xInc) Then xInc = -xInc
        If .Top + .Height > ScaleHeight Then yInc = -yInc
        If .Top <= Abs(yInc) Then yInc = -yInc
        .Move .Left + xInc, .Top + yInc
    End With
End Sub

Private Sub SetTimer(eatAnimateA As Integer)
    eatAnimate = eatAnimateA
    Select Case eatAnimate
    Case eatNone
        timerAnimate.Interval = 0
        ' Hide XPictureGlass object
        pgControl.Visible = False
        cmdAnimate.Caption = "&Animate Picture"
        cmdAnimateCtl.Caption = "Animate &Control"
        ' Remove active CPictureGlass object from memory
        Set pgPicture = Nothing
    Case eatPicture, eatControl
        timerAnimate.Interval = 10
    Case eatCardBacks
        timerAnimate.Interval = 100
    End Select
End Sub

Private Sub ShowCards(aCards() As Variant)
    Dim iSuit As Integer, iCard As Integer
    
    For iCard = 0 To 12
        For iSuit = 0 To 3
            cdtDraw Me.hDC, (dxCard * 0.3) + (iSuit * dxCard), _
                    (dyCard * 0.1) + iCard * (dyCard / 3.7), _
                    aCards(iSuit + (iCard * 4)), 0, 0&
        Next
    Next
End Sub
    
Sub SetClipControls(hWnd As Long, f As Boolean)
    ' You want to do this:
    'Me.ClipControls = f
    ' But Visual Basic won't let you; do this instead:
    ChangeStyleBit hWnd, f, WS_CLIPCHILDREN
End Sub

Private Sub lblFormColor_MouseUp(Button As Integer, Shift As Integer, _
                                 X As Single, Y As Single)
    Dim getclr As New CColorPicker
    getclr.Color = lblFormColor.BackColor
    getclr.Load Left + lblFormColor.Left + X, Top + lblFormColor.Top + Y
    lblFormColor.BackColor = getclr.Color
    BackColor = getclr.Color
    chkClip.BackColor = getclr.Color
    chkAutoRedraw.BackColor = getclr.Color
    lblFormColorText.BackColor = getclr.Color
    lblMaskColorText.BackColor = getclr.Color
End Sub

Private Sub lblFormColorText_MouseUp(Button As Integer, Shift As Integer, _
                                     X As Single, Y As Single)
    lblFormColor_MouseUp Button, Shift, X, Y
End Sub

Private Sub lblMaskColor_MouseUp(Button As Integer, Shift As Integer, _
                                 X As Single, Y As Single)
    Dim getclr As New CColorPicker
    getclr.Color = lblMaskColor.BackColor
    getclr.Load Left + lblMaskColor.Left + X, Top + lblMaskColor.Top + Y
    lblMaskColor.BackColor = getclr.Color
    clrMask = getclr.Color
End Sub

Private Sub lblMaskColorText_MouseUp(Button As Integer, Shift As Integer, _
                                     X As Single, Y As Single)
    lblMaskColor_MouseUp Button, Shift, X, Y
End Sub

Private Sub timerAnimate_ThatTime()
    Select Case eatAnimate
    Case eatNone
        Exit Sub
    Case eatCardBacks
        AnimateBacks
    Case eatPicture
        AnimatePicture
    Case eatControl
        AnimateControl
    End Select
End Sub
'
