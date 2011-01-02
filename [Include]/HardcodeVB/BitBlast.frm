VERSION 5.00
Begin VB.Form FBlit 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bit Blast"
   ClientHeight    =   5904
   ClientLeft      =   1356
   ClientTop       =   1836
   ClientWidth     =   8784
   FillStyle       =   7  'Diagonal Cross
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   7.8
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "BITBLAST.frx":0000
   LinkTopic       =   "frmBlt"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5904
   ScaleWidth      =   8784
   Begin VB.OptionButton optStretch 
      Caption         =   "Stretch Delete"
      Height          =   375
      Index           =   2
      Left            =   2040
      TabIndex        =   15
      Top             =   1200
      Value           =   -1  'True
      Width           =   1695
   End
   Begin VB.OptionButton optStretch 
      Caption         =   "Stretch Or"
      Height          =   375
      Index           =   1
      Left            =   2040
      TabIndex        =   14
      Top             =   840
      Width           =   1695
   End
   Begin VB.OptionButton optStretch 
      Caption         =   "Stretch And"
      Height          =   375
      Index           =   0
      Left            =   2040
      TabIndex        =   13
      Top             =   480
      Width           =   1695
   End
   Begin VB.CommandButton cmdMask 
      Caption         =   "Mask"
      Default         =   -1  'True
      Height          =   495
      Left            =   108
      TabIndex        =   10
      Top             =   720
      Width           =   1695
   End
   Begin VB.CheckBox chkBitBlt 
      Caption         =   "Use BitBlt"
      Height          =   375
      Left            =   2055
      TabIndex        =   9
      Top             =   120
      Value           =   1  'Checked
      Width           =   1695
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   2520
      Width           =   1695
   End
   Begin VB.CommandButton cmdStretch 
      Caption         =   "Stretch"
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   1305
      Width           =   1695
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear Destination"
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   1920
      Width           =   1695
   End
   Begin VB.CommandButton cmdBlit 
      Caption         =   "Blit"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   1695
   End
   Begin VB.ListBox lstROP 
      Height          =   1968
      Left            =   3840
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin VB.PictureBox pbTest 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   855
      Index           =   3
      Left            =   4425
      ScaleHeight     =   852
      ScaleWidth      =   852
      TabIndex        =   4
      Tag             =   "4"
      Top             =   3960
      Width           =   855
   End
   Begin VB.PictureBox pbTest 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      FillStyle       =   6  'Cross
      Height          =   855
      Index           =   2
      Left            =   3000
      ScaleHeight     =   852
      ScaleWidth      =   852
      TabIndex        =   3
      Tag             =   "3"
      Top             =   3960
      Width           =   855
   End
   Begin VB.PictureBox pbTest 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      FillColor       =   &H0000FF00&
      FillStyle       =   5  'Downward Diagonal
      Height          =   855
      Index           =   1
      Left            =   1560
      ScaleHeight     =   852
      ScaleWidth      =   852
      TabIndex        =   2
      Tag             =   "2"
      Top             =   3960
      Width           =   855
   End
   Begin VB.PictureBox pbTest 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      FillStyle       =   7  'Diagonal Cross
      Height          =   855
      Index           =   0
      Left            =   120
      ScaleHeight     =   57
      ScaleMode       =   0  'User
      ScaleWidth      =   57
      TabIndex        =   1
      Tag             =   "1"
      Top             =   3960
      Width           =   855
   End
   Begin VB.Image imgMarble 
      Height          =   2664
      Left            =   3660
      Picture         =   "BITBLAST.frx":0CFA
      Top             =   6396
      Width           =   2952
   End
   Begin VB.Image imgBlank 
      Height          =   975
      Left            =   7695
      Picture         =   "BITBLAST.frx":7904
      Stretch         =   -1  'True
      Top             =   6615
      Width           =   900
   End
   Begin VB.Image imgHead 
      Height          =   855
      Left            =   660
      Picture         =   "BITBLAST.frx":80C6
      Stretch         =   -1  'True
      Top             =   6465
      Width           =   885
   End
   Begin VB.Label lbl 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   5
      Left            =   1560
      TabIndex        =   19
      Top             =   3735
      Width           =   180
   End
   Begin VB.Label lbl 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   4
      Left            =   2985
      TabIndex        =   18
      Top             =   3735
      Width           =   180
   End
   Begin VB.Label lbl 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "4"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   3
      Left            =   4440
      TabIndex        =   17
      Top             =   3735
      Width           =   165
   End
   Begin VB.Label lbl 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      ForeColor       =   &H00000000&
      Height          =   180
      Index           =   2
      Left            =   150
      TabIndex        =   16
      Top             =   3765
      Width           =   180
   End
   Begin VB.Image imgSrcPoint 
      Height          =   384
      Left            =   336
      Picture         =   "BITBLAST.frx":8868
      Top             =   3492
      Width           =   384
   End
   Begin VB.Image imgDstPoint 
      Height          =   384
      Left            =   360
      Picture         =   "BITBLAST.frx":8B72
      Top             =   4812
      Width           =   384
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "Destination (right mouse button)"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   1
      Left            =   540
      TabIndex        =   12
      Top             =   5490
      Width           =   2775
   End
   Begin VB.Label lbl 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Source (left mouse button)"
      ForeColor       =   &H000000FF&
      Height          =   252
      Index           =   0
      Left            =   600
      TabIndex        =   11
      Top             =   3180
      Width           =   2772
   End
   Begin VB.Image imgBlob 
      Height          =   804
      Left            =   2292
      Picture         =   "BITBLAST.frx":8FB4
      Top             =   6480
      Width           =   744
   End
End
Attribute VB_Name = "FBlit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Blit Test Form"
Option Explicit

Private dxBlt As Long, dyBlt As Long
Private pbSrc As PictureBox, pbDst As PictureBox
    
Const ordSrc = 0
Const ordDst = 1

Private Sub Form_Load()

    Randomize
    ' Initializations
    pbTest(0).Picture = imgHead.Picture
    pbTest(1).Picture = imgBlob.Picture
    pbTest(2).Picture = imgMarble.Picture
    pbTest(3).Picture = imgBlank.Picture
    With pbTest(3)
        dxBlt = ScaleX(.Width, vbTwips, vbPixels)
        dyBlt = ScaleY(.Height, vbTwips, vbPixels)
        pbTest(3).Circle (.Width / 2, .Height / 2), .Width / 3
    End With
    InitRop chkBitBlt = 1
    Set pbSrc = pbTest(0)
    Set pbDst = pbTest(2)
    imgDstPoint.Left = pbDst.Left + (pbDst.Width / 2) - _
                       (imgDstPoint.Width / 2)
    imgSrcPoint.Left = pbSrc.Left + (pbSrc.Width / 2) - _
                       (imgSrcPoint.Width / 2)
    
#If 0 Then
    Show
    ' Default black to blue vertical fade on current form
    Fade Me
    ' Make it blue to black
    Fade Me, LightToDark:=False
    ' Red horizontal fade on FBlit
    Fade FBlit, Red:=True, Horizontal:=True
    ' Violet vertical fade on picture box
    Fade pbTest(0), Red:=True, Blue:=True
    ' Black to white diagonal fade on current form
    Fade Me, Horizontal:=True, Vertical:=True, _
         Red:=True, Green:=True, Blue:=True
#End If
  
End Sub

Private Sub chkBitBlt_Click()
    If chkBitBlt = vbChecked Then
        optStretch(0).Visible = True
        optStretch(1).Visible = True
        optStretch(2).Visible = True
    Else
        optStretch(0).Visible = False
        optStretch(1).Visible = False
        optStretch(2).Visible = False
    End If
End Sub

Private Sub cmdBlit_Click()
    Dim rop As Long
    rop = lstROP.ItemData(lstROP.ListIndex)
    If chkBitBlt.Value = vbChecked Then
        Call BitBlt(pbDst.hDC, 0, 0, dxBlt, dyBlt, _
                    pbSrc.hDC, 0, 0, rop)
        pbDst.Refresh
    Else
        pbSrc.Picture = pbSrc.Image
        pbDst.PaintPicture pbSrc.Picture, 0, 0, , , , , , , rop
    End If
End Sub

Private Sub cmdClear_Click()
    Select Case pbDst.Left
    Case pbTest(0).Left
        'pbDst.Picture = LoadResPicture(100, 0) ' vbResBitMap)
        pbDst.Picture = imgHead.Picture
    Case pbTest(1).Left
        'pbDst.Picture = LoadResPicture(101, vbResBitmap)
        pbDst.Picture = imgBlob.Picture
    Case pbTest(2).Left
        'pbDst.Picture = LoadResPicture(102, vbResBitmap)
        pbDst.Picture = imgMarble.Picture
    Case pbTest(3).Left
        'pbDst.Picture = LoadResPicture(103, vbResBitmap)
        pbDst.Picture = imgBlank.Picture
        pbTest(3).Circle (pbTest(3).Width / 2, _
                          pbTest(3).Height / 2), _
                          pbTest(3).Width / 3
    End Select
    pbDst.Picture = pbDst.Image
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

' Secret undocumented command!
Private Sub cmdClear_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 3 Then Form_Resize
End Sub

' Secret undocumented command to change fade pattern!
Private Sub cmdExit_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then Form_Resize
End Sub

Private Sub cmdMask_Click()
    Dim hdcMono As Long, hbmpMono As Long, hbmpOld As Long
    
    ' Create memory device context
    hdcMono = CreateCompatibleDC(0)
    ' Create monochrome bitmap and select it into DC
    hbmpMono = CreateCompatibleBitmap(hdcMono, dxBlt, dyBlt)
    hbmpOld = SelectObject(hdcMono, hbmpMono)
    ' Copy color bitmap to DC to create mono mask
    BitBlt hdcMono, 0, 0, dxBlt, dyBlt, pbSrc.hDC, 0, 0, SRCCOPY
    ' Copy mono memory mask to visible picture box
    BitBlt pbDst.hDC, 0, 0, dxBlt, dyBlt, hdcMono, 0, 0, SRCCOPY
    pbDst.Refresh
    ' Clean up
    Call SelectObject(hdcMono, hbmpOld)
    Call DeleteDC(hdcMono)
    Call DeleteObject(hbmpMono)
End Sub

Private Sub cmdStretch_Click()
    If chkBitBlt.Value = vbChecked Then
        ' Stretch inside out
        Call StretchBlt(hDC, ScaleX(Width, vbTwips, vbPixels) * 0.97, _
                             ScaleY(Height, vbTwips, vbPixels) * 0.9, _
                             -dxBlt * 3.5, -dyBlt * 6.5, _
                             pbSrc.hDC, 0, 0, dxBlt, dyBlt, vbSrcCopy)
        ' Compress backward
        Call StretchBlt(hDC, ScaleX(Width, vbTwips, vbPixels) * 0.75, _
                             ScaleY(Height, vbTwips, vbPixels) * 0.8, _
                             -dxBlt * 0.9, dyBlt * 0.5, _
                             pbSrc.hDC, 0, 0, dxBlt, dyBlt, vbSrcCopy)
        ' This line required in VB6 because of change in painting model
        Refresh
    Else
        With pbSrc
            .Picture = .Image
            ' Stretch inside out
            PaintPicture .Picture, Width * 0.97, Height * 0.9, _
                                   -.Width * 3.5, -.Height * 6.5
            ' Compress backward
            PaintPicture .Picture, Width * 0.75, Height * 0.8, _
                                   -.Width * 0.9, .Height * 0.5
        End With
    End If
End Sub

Sub InitRop(f As Boolean)
    With lstROP
        If f Then
            .AddItem "SrcCopy"
            .ItemData(.NewIndex) = SRCCOPY
            .AddItem "SrcPaint"
            .ItemData(.NewIndex) = SRCPAINT
            .AddItem "SrcAnd"
            .ItemData(.NewIndex) = SRCAND
            .AddItem "SrcInvert"
            .ItemData(.NewIndex) = SRCINVERT
            .AddItem "SrcErase"
            .ItemData(.NewIndex) = SRCERASE
            .AddItem "NotSrcCopy"
            .ItemData(.NewIndex) = NOTSRCCOPY
            .AddItem "NotSrcErase"
            .ItemData(.NewIndex) = NOTSRCERASE
            .AddItem "MergeCopy"
            .ItemData(.NewIndex) = MERGECOPY
            .AddItem "MergePaint"
            .ItemData(.NewIndex) = MERGEPAINT
            .AddItem "PatCopy"
            .ItemData(.NewIndex) = PATCOPY
            .AddItem "PatPaint"
            .ItemData(.NewIndex) = PATPAINT
            .AddItem "PatInvert"
            .ItemData(.NewIndex) = PATINVERT
            .AddItem "DstInvert"
            .ItemData(.NewIndex) = DSTINVERT
            .AddItem "Blackness"
            .ItemData(.NewIndex) = BLACKNESS
            .AddItem "Whiteness"
            .ItemData(.NewIndex) = WHITENESS
        Else
            .AddItem "SrcCopy"
            .ItemData(.NewIndex) = vbSrcCopy
            .AddItem "SrcPaint"
            .ItemData(.NewIndex) = vbSrcPaint
            .AddItem "SrcAnd"
            .ItemData(.NewIndex) = vbSrcAnd
            .AddItem "SrcInvert"
            .ItemData(.NewIndex) = vbSrcInvert
            .AddItem "SrcErase"
            .ItemData(.NewIndex) = vbSrcErase
            .AddItem "NotSrcCopy"
            .ItemData(.NewIndex) = vbNotSrcCopy
            .AddItem "NotSrcErase"
            .ItemData(.NewIndex) = vbNotSrcErase
            .AddItem "MergeCopy"
            .ItemData(.NewIndex) = vbMergeCopy
            .AddItem "MergePaint"
            .ItemData(.NewIndex) = vbMergePaint
            .AddItem "PatCopy"
            .ItemData(.NewIndex) = vbPatCopy
            .AddItem "PatPaint"
            .ItemData(.NewIndex) = vbPatPaint
            .AddItem "PatInvert"
            .ItemData(.NewIndex) = vbPatInvert
            .AddItem "DstInvert"
            .ItemData(.NewIndex) = vbDstInvert
            .AddItem "Blackness"
            .ItemData(.NewIndex) = vbBlackness
            .AddItem "Whiteness"
            .ItemData(.NewIndex) = vbWhiteness
        End If
        Dim i As Integer
        For i = 0 To .ListCount - 1
            If .List(i) = "SrcCopy" Then
                .ListIndex = i
                Exit For
            End If
        Next
    End With
        
End Sub


Private Sub Form_Resize()
    Dim fRed As Boolean, fGreen As Boolean, fBlue As Boolean
    ' If red, green, and blue are false, you get a black fade
    Do
        fRed = Random(0, 1)
        fGreen = Random(0, 1)
        fBlue = Random(0, 1)
    Loop Until (fRed = True) Or (fGreen = True) Or (fBlue = True)
    Fade Me, Red:=fRed, Green:=fGreen, Blue:=fBlue, _
             Horizontal:=Random(0, 1), Vertical:=Random(0, 1), _
             LightToDark:=Random(0, 1)
End Sub

Private Sub lstROP_DblClick()
    cmdBlit_Click
End Sub

Private Sub optStretch_Click(Index As Integer)
    Call SetStretchBltMode(Me.hDC, Index + 1)
End Sub

Private Sub pbTest_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim ord As Integer
    If Button And 2 Then
        ' Right mouse is destination
        Set pbDst = pbTest(Index)
        imgDstPoint.Left = pbDst.Left + (pbDst.Width / 2) - _
                           (imgDstPoint.Width / 2)
    Else
        ' Other mouse (probably left) is source
        Set pbSrc = pbTest(Index)
        imgSrcPoint.Left = pbSrc.Left + (pbSrc.Width / 2) - _
                           (imgSrcPoint.Width / 2)
    End If
End Sub
