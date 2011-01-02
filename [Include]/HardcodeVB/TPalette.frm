VERSION 5.00
Begin VB.Form FTestPalette 
   AutoRedraw      =   -1  'True
   Caption         =   "Pal (and I don't mean Buddy)"
   ClientHeight    =   5484
   ClientLeft      =   1992
   ClientTop       =   1932
   ClientWidth     =   6984
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   7.8
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "TPalette.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   2  'Custom
   Picture         =   "TPalette.frx":0CFA
   ScaleHeight     =   5484
   ScaleWidth      =   6984
   Begin VB.HScrollBar hs 
      Height          =   204
      LargeChange     =   20
      Left            =   1992
      Max             =   20
      Min             =   500
      SmallChange     =   30
      TabIndex        =   16
      Top             =   984
      Value           =   500
      Width           =   852
   End
   Begin VB.OptionButton optPal 
      Caption         =   "Picture"
      Height          =   192
      Index           =   1
      Left            =   180
      TabIndex        =   15
      Top             =   720
      Width           =   972
   End
   Begin VB.OptionButton optPal 
      Caption         =   "Form"
      Height          =   192
      Index           =   0
      Left            =   180
      TabIndex        =   14
      Top             =   540
      Value           =   -1  'True
      Width           =   972
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "New Bitmap..."
      Height          =   396
      Left            =   1560
      TabIndex        =   13
      Top             =   504
      Width           =   1548
   End
   Begin VB.TextBox txtTotal 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2700
      MaxLength       =   3
      TabIndex        =   12
      Top             =   1716
      Width           =   396
   End
   Begin VB.PictureBox pbPal 
      Align           =   1  'Align Top
      AutoRedraw      =   -1  'True
      Height          =   360
      Left            =   0
      ScaleHeight     =   312
      ScaleWidth      =   6936
      TabIndex        =   9
      Top             =   0
      Width           =   6984
   End
   Begin VB.CommandButton cmdAcmeAnimate 
      BackColor       =   &H8000000C&
      Caption         =   "&< >"
      Height          =   375
      Index           =   3
      Left            =   1680
      TabIndex        =   8
      Top             =   1296
      Width           =   696
   End
   Begin VB.PictureBox pbBitmap 
      AutoSize        =   -1  'True
      Height          =   3168
      Left            =   192
      Picture         =   "TPalette.frx":BCD3C
      ScaleHeight     =   3120
      ScaleWidth      =   2880
      TabIndex        =   7
      Top             =   2076
      Visible         =   0   'False
      Width           =   2928
   End
   Begin VB.CommandButton cmdAcmeAnimate 
      BackColor       =   &H8000000C&
      Caption         =   "&Right"
      Height          =   375
      Index           =   1
      Left            =   2412
      TabIndex        =   6
      Top             =   1296
      Width           =   696
   End
   Begin VB.CommandButton cmdAcmeAnimate 
      BackColor       =   &H8000000C&
      Caption         =   "&> <"
      Height          =   375
      Index           =   2
      Left            =   936
      TabIndex        =   5
      Top             =   1296
      Width           =   696
   End
   Begin VB.TextBox txtTo 
      Height          =   285
      Left            =   1716
      MaxLength       =   3
      TabIndex        =   3
      Top             =   1716
      Width           =   396
   End
   Begin VB.TextBox txtFrom 
      Height          =   285
      Left            =   732
      MaxLength       =   3
      TabIndex        =   2
      Text            =   "0"
      Top             =   1704
      Width           =   396
   End
   Begin VB.Timer tmrAnimate 
      Left            =   10464
      Top             =   96
   End
   Begin VB.CheckBox chkContinuous 
      Caption         =   "Conti&nuous"
      Height          =   255
      Left            =   216
      TabIndex        =   1
      Top             =   960
      Width           =   1344
   End
   Begin VB.CommandButton cmdAcmeAnimate 
      Caption         =   "&Left"
      Height          =   375
      Index           =   0
      Left            =   204
      TabIndex        =   0
      Top             =   1296
      Width           =   696
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "S"
      Height          =   216
      Index           =   4
      Left            =   1824
      TabIndex        =   18
      Top             =   972
      Width           =   156
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "F"
      Height          =   216
      Index           =   3
      Left            =   2892
      TabIndex        =   17
      Top             =   972
      Width           =   156
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "&To:"
      ForeColor       =   &H00000000&
      Height          =   288
      Index           =   2
      Left            =   1212
      TabIndex        =   11
      Top             =   1704
      Width           =   504
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "Total:"
      ForeColor       =   &H00000000&
      Height          =   288
      Index           =   1
      Left            =   2196
      TabIndex        =   10
      Top             =   1716
      Width           =   504
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "&From:"
      ForeColor       =   &H00000000&
      Height          =   288
      Index           =   0
      Left            =   216
      TabIndex        =   4
      Top             =   1704
      Width           =   504
   End
End
Attribute VB_Name = "FTestPalette"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private fOnPicture As Boolean
Private pal As New CPalette
Private cPal As Long
Private aColors() As OLE_COLOR
Private ecd As ECycleDirection
Private iFormTo As Long, iFormFrom As Long
Private iPicTo As Long, iPicFrom As Long

Private Sub Form_Load()

    Show
    
    Dim xPixels As Long, yPixels As Long
    xPixels = Screen.Width / Screen.TwipsPerPixelX
    yPixels = Screen.Height / Screen.TwipsPerPixelY
    ' Use the largest size we can get away with
    If xPixels <= 640 Or yPixels <= 480 Then
        Width = 630 * Screen.TwipsPerPixelX
        Height = 470 * Screen.TwipsPerPixelY
    ElseIf xPixels < 800 Or yPixels < 600 Then
        Width = 790 * Screen.TwipsPerPixelX
        Height = 590 * Screen.TwipsPerPixelY
    ElseIf xPixels <= 1024 Or yPixels <= 768 Then
        Width = 1000 * Screen.TwipsPerPixelX
        Height = 750 * Screen.TwipsPerPixelY
    Else
        Width = 1032 * Screen.TwipsPerPixelX
        Height = 778 * Screen.TwipsPerPixelY
    End If
    
    ' Initialize exclusions
    iFormFrom = 0
    iFormTo = 233
    iPicFrom = 10
    iPicTo = 236
    
    ' Fake button clicks to initialize
    optPal_Click -fOnPicture
    chkContinuous_Click
        
End Sub

Private Sub Form_Unload(Cancel As Integer)
    pal.Destroy
End Sub

Private Sub hs_Change()
    tmrAnimate.Interval = hs.Value
End Sub

Private Sub optPal_Click(Index As Integer)
    On Error GoTo PalFail
    Select Case Index
    Case 0  ' Form
        Palette = Picture
        pbBitmap.Visible = False
        DrawPalette pbPal, Picture.hPal
        ' Create the palette and initialize the color array
        cPal = pal.Create(Picture.hPal, hWnd, aColors, iFormFrom, iFormTo)
        txtTotal = cPal
        txtTo = iFormTo
        txtFrom = iFormFrom
        fOnPicture = False
        
    Case 1  ' Picture
        pbBitmap.Visible = True
        Palette = pbBitmap.Picture
        DrawPalette pbPal, pbBitmap.Picture.hPal
        ' Create the palette and initialize the color array
        cPal = pal.Create(pbBitmap.Picture.hPal, pbBitmap.hWnd, _
                          aColors, iPicFrom, iPicTo)
        txtTotal = cPal
        If iPicTo = -1 Then iPicTo = cPal
        txtTo = iPicTo
        txtFrom = iPicFrom
        fOnPicture = True
        
    End Select
    Exit Sub
PalFail:
    MsgBox "Can't load palette: " & Err.Description
End Sub

Private Sub chkContinuous_Click()
    If chkContinuous.Value = vbChecked Then
        tmrAnimate.Interval = 154
        tmrAnimate.Enabled = True
    Else
        tmrAnimate.Enabled = False
    End If
End Sub

Private Sub cmdNew_Click()
    Dim opfile As New COpenPictureFile, fTimerOn As Boolean
    fTimerOn = tmrAnimate.Enabled
    tmrAnimate.Enabled = False
    With opfile
        .InitDir = WindowsDir
        .FilterType = efpBitmap
        .Load Left + (Width / 4), Top + (Height / 4)
        If .FileName <> sEmpty Then
            pbBitmap.Picture = LoadPicture(.FileName)
        End If
    End With
    iPicFrom = 0
    iPicTo = -1
    If pbBitmap.Picture.hPal <> hNull Then
        optPal_Click 1
        optPal(1).Value = True
    Else
        MsgBox "Bitmap does not have palette"
    End If
    tmrAnimate.Enabled = fTimerOn
End Sub

Private Sub txtFrom_LostFocus()
    If fOnPicture Then
        iPicFrom = CLng(txtFrom)
    Else
        iFormFrom = CLng(txtFrom)
    End If
    optPal_Click -fOnPicture
End Sub

Private Sub txtTo_LostFocus()
    If fOnPicture Then
        iPicTo = CLng(txtTo)
    Else
        iFormTo = CLng(txtTo)
    End If
    optPal_Click -fOnPicture
End Sub

Private Sub cmdAcmeAnimate_Click(Index As Integer)
    ' Index from ECycleDirection enum: left, right, inside out, outside in
    RotatePaletteArray aColors, Index
    pal.ModifyPalette aColors
    DrawPalette pbPal, pal.Handle
    ecd = Index
End Sub

Private Sub tmrAnimate_Timer()
    Call cmdAcmeAnimate_Click(CInt(ecd))
End Sub

Private Sub txtFrom_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 48 To 57, 8
        Case Else
            Beep
            KeyAscii = 0
    End Select
End Sub

Private Sub txtTo_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 48 To 57, 8
        Case Else
            Beep
            KeyAscii = 0
    End Select
End Sub


