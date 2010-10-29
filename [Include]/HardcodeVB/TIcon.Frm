VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FTestIcons 
   Caption         =   "Test Icons"
   ClientHeight    =   7140
   ClientLeft      =   5256
   ClientTop       =   3300
   ClientWidth     =   5256
   Icon            =   "TIcon.frx":0000
   LinkTopic       =   "Form1"
   MousePointer    =   99  'Custom
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7140
   ScaleWidth      =   5256
   Begin VB.TextBox txtY 
      Height          =   288
      Left            =   1200
      TabIndex        =   4
      Text            =   "32"
      Top             =   1320
      Width           =   372
   End
   Begin VB.ComboBox cboDepth 
      Height          =   288
      ItemData        =   "TIcon.frx":0442
      Left            =   240
      List            =   "TIcon.frx":0452
      TabIndex        =   2
      Top             =   948
      Width           =   1332
   End
   Begin VB.TextBox txtX 
      Height          =   288
      Left            =   480
      TabIndex        =   3
      Text            =   "32"
      Top             =   1320
      Width           =   372
   End
   Begin VB.ComboBox cboSize 
      Height          =   288
      ItemData        =   "TIcon.frx":0480
      Left            =   240
      List            =   "TIcon.frx":0493
      TabIndex        =   1
      Top             =   360
      Width           =   1332
   End
   Begin VB.CheckBox chkSmall 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Small"
      Height          =   330
      Left            =   930
      MaskColor       =   &H00FFFFFF&
      Picture         =   "TIcon.frx":04C7
      TabIndex        =   5
      Top             =   5496
      Width           =   780
   End
   Begin VB.PictureBox pb 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   540
      Index           =   0
      Left            =   3492
      ScaleHeight     =   540
      ScaleWidth      =   1272
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   516
      Width           =   1275
   End
   Begin VB.PictureBox pb 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   540
      Index           =   2
      Left            =   3492
      ScaleHeight     =   540
      ScaleWidth      =   1272
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   2676
      Width           =   1275
   End
   Begin VB.PictureBox pb 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   540
      Index           =   5
      Left            =   3492
      ScaleHeight     =   540
      ScaleWidth      =   1272
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   4656
      Width           =   1275
   End
   Begin VB.PictureBox pb 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   540
      Index           =   4
      Left            =   3504
      ScaleHeight     =   540
      ScaleWidth      =   1272
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   3996
      Width           =   1275
   End
   Begin VB.PictureBox pb 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   540
      Index           =   3
      Left            =   3492
      ScaleHeight     =   540
      ScaleWidth      =   1272
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   3336
      Width           =   1275
   End
   Begin VB.PictureBox pb 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   540
      Index           =   1
      Left            =   3492
      ScaleHeight     =   540
      ScaleWidth      =   1272
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   2016
      Width           =   1275
   End
   Begin MSComctlLib.ListView lstvwVB 
      Height          =   1050
      Left            =   150
      TabIndex        =   6
      Top             =   5895
      Width           =   1425
      _ExtentX        =   2498
      _ExtentY        =   1863
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin MSComctlLib.ListView lstvwFixed 
      Height          =   1050
      Left            =   1680
      TabIndex        =   7
      Top             =   5880
      Width           =   1350
      _ExtentX        =   2392
      _ExtentY        =   1863
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin MSComctlLib.ImageList imlstSmallFixed 
      Left            =   2040
      Top             =   7440
      _ExtentX        =   804
      _ExtentY        =   804
      BackColor       =   -2147483643
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imlstLargeFixed 
      Left            =   1440
      Top             =   7440
      _ExtentX        =   804
      _ExtentY        =   804
      BackColor       =   -2147483643
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imlstLargeVB 
      Left            =   120
      Top             =   7440
      _ExtentX        =   804
      _ExtentY        =   804
      BackColor       =   -2147483643
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imlstSmallVB 
      Left            =   840
      Top             =   7440
      _ExtentX        =   804
      _ExtentY        =   804
      BackColor       =   -2147483643
      _Version        =   393216
   End
   Begin VB.Label lblY 
      BackStyle       =   0  'Transparent
      Caption         =   "Y:"
      Enabled         =   0   'False
      Height          =   312
      Left            =   960
      TabIndex        =   26
      Top             =   1320
      Width           =   252
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "Color Depth:"
      Height          =   192
      Index           =   7
      Left            =   240
      TabIndex        =   25
      Top             =   720
      Width           =   972
   End
   Begin VB.Line Line1 
      X1              =   360
      X2              =   4800
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "Size:"
      Height          =   228
      Index           =   13
      Left            =   240
      TabIndex        =   24
      Top             =   120
      Width           =   900
   End
   Begin VB.Label lblX 
      BackStyle       =   0  'Transparent
      Caption         =   "X:"
      Enabled         =   0   'False
      Height          =   312
      Left            =   240
      TabIndex        =   23
      Top             =   1320
      Width           =   252
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "ListView"
      Height          =   252
      Index           =   5
      Left            =   132
      TabIndex        =   22
      Top             =   5532
      Width           =   816
   End
   Begin VB.Image img 
      Appearance      =   0  'Flat
      Height          =   540
      Index           =   0
      Left            =   1884
      Top             =   504
      Width           =   1272
   End
   Begin VB.Label lblData 
      BackStyle       =   0  'Transparent
      Height          =   855
      Left            =   3120
      TabIndex        =   20
      Top             =   5865
      Width           =   1875
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "Picture data:"
      Height          =   330
      Index           =   6
      Left            =   3120
      TabIndex        =   19
      Top             =   5535
      Width           =   1035
   End
   Begin VB.Image imgPattern 
      Height          =   1536
      Left            =   3240
      Picture         =   "TIcon.frx":2DA9
      Top             =   7440
      Width           =   1536
   End
   Begin VB.Image img 
      Appearance      =   0  'Flat
      Height          =   540
      Index           =   2
      Left            =   1872
      Top             =   2676
      Width           =   1272
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "Shell Size:"
      Height          =   432
      Index           =   4
      Left            =   300
      TabIndex        =   18
      Top             =   2676
      Width           =   852
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "Image Size:"
      Height          =   432
      Index           =   3
      Left            =   300
      TabIndex        =   16
      Top             =   3336
      Width           =   876
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "Small:"
      Height          =   432
      Index           =   2
      Left            =   300
      TabIndex        =   15
      Top             =   3996
      Width           =   852
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "Huge:"
      Height          =   432
      Index           =   1
      Left            =   300
      TabIndex        =   14
      Top             =   4656
      Width           =   756
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "Default:"
      Height          =   432
      Index           =   0
      Left            =   336
      TabIndex        =   13
      Top             =   2016
      Width           =   852
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "Picture Boxes"
      Height          =   336
      Index           =   9
      Left            =   3480
      TabIndex        =   12
      Top             =   84
      Width           =   1272
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "Images"
      Height          =   336
      Index           =   8
      Left            =   1860
      TabIndex        =   11
      Top             =   84
      Width           =   1272
   End
   Begin VB.Image img 
      Appearance      =   0  'Flat
      Height          =   540
      Index           =   5
      Left            =   1872
      Top             =   4656
      Width           =   1272
   End
   Begin VB.Image img 
      Appearance      =   0  'Flat
      Height          =   540
      Index           =   1
      Left            =   1872
      Top             =   2016
      Width           =   1272
   End
   Begin VB.Image img 
      Appearance      =   0  'Flat
      Height          =   540
      Index           =   3
      Left            =   1872
      Top             =   3336
      Width           =   1272
   End
   Begin VB.Image img 
      Appearance      =   0  'Flat
      Height          =   540
      Index           =   4
      Left            =   1872
      Top             =   3996
      Width           =   1272
   End
End
Attribute VB_Name = "FTestIcons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private ctlLast As Control
Private ctlCur As Control

Enum EIconPos
    eipBasic = 0
    eipDefault
    eipShell
    eipImage
    eipSmall
    eipHuge
End Enum

#If iVBVer > 5 Then
Private Sub cboDepth_Click()
    Static fNotFirst As Boolean
    If fNotFirst = False Then
        fNotFirst = True
        Exit Sub
    End If
    ' Load icon with Basic's LoadPicture and optional arguments
    Set img(eipBasic).Picture = BasicLoadPicture("i.ico")
    Set pb(eipBasic).Picture = BasicLoadPicture("i.ico")
End Sub

Private Sub cboSize_Click()
    Static fNotFirst As Boolean
    If fNotFirst = False Then
        fNotFirst = True
        Exit Sub
    End If
    ' Load icon with Basic's LoadPicture and optional arguments
    Set img(eipBasic).Picture = BasicLoadPicture("i.ico")
    Set pb(eipBasic).Picture = BasicLoadPicture("i.ico")
End Sub

Private Sub txtX_LostFocus()
    Set img(eipBasic).Picture = BasicLoadPicture("i.ico")
    Set pb(eipBasic).Picture = BasicLoadPicture("i.ico")
End Sub

Private Sub txtY_LostFocus()
    Set img(eipBasic).Picture = BasicLoadPicture("i.ico")
    Set pb(eipBasic).Picture = BasicLoadPicture("i.ico")
End Sub
#End If

Private Sub Form_Load()
    If Not HasShell Then
        MsgBox "This program requires the enhanced Windows user interface"
        End
    End If
    ChDir App.Path
    Set ctlLast = pb(eipBasic)
    Set ctlCur = img(eipBasic)
    Show
    
    cboSize.ListIndex = 0
    cboDepth.ListIndex = 0
            
#If iVBVer > 5 Then
    ' Load icon with Basic's LoadPicture and optional arguments
    Set img(eipBasic).Picture = BasicLoadPicture("i.ico")
    Set pb(eipBasic).Picture = BasicLoadPicture("i.ico")
#Else
    ' No such thing for VB5
    cboSize.Enabled = False
    cboDepth.Enabled = False
    txtX.Enabled = False
    txtY.Enabled = False
    Set img(eipBasic).Picture = LoadPicture("i.ico")
    Set pb(eipBasic).Picture = LoadPicture("i.ico")
#End If
    ' Load default icon (system metrics size)
    Set img(eipDefault).Picture = LoadAnyPicture("i.ico", eisDefault)
    Set pb(eipDefault).Picture = LoadAnyPicture("i.ico", eisDefault)
    ' Load icon stretched to system metrics size
    Set img(eipShell).Picture = LoadAnyPicture("i.ico", eisShell)
    Set pb(eipShell).Picture = LoadAnyPicture("i.ico", eisShell)
    ' Load first icon image in file with its real size
    Set img(eipImage).Picture = LoadAnyPicture("i.ico", eisImage)
    Set pb(eipImage).Picture = LoadAnyPicture("i.ico", eisImage)
    ' Load small (16 by 16) icon, squashing if necessary
    Set img(eipSmall).Picture = LoadAnyPicture("i.ico", eisSmall)
    Set pb(eipSmall).Picture = LoadAnyPicture("i.ico", eisSmall)
    ' Load huge (48 by 48) icon
    Set img(eipHuge).Picture = LoadAnyPicture("i.ico", eisHuge)
    Set pb(eipHuge).Picture = LoadAnyPicture("i.ico", eisHuge)
    PictureData
    
    ' Load small and large ImageLists and put in ListView
#If iVBVer > 5 Then
#If 1 Then
    ' This statement sometimes fails to load small icon, contrary to docs
    imlstSmallVB.ListImages.Add , , LoadPicture("i.ico", vbLPSmall)
#Else
    ' This statement does load small icon
    imlstSmallVB.ListImages.Add , , LoadPicture("i.ico", vbLPCustom, , 16, 16)
#End If
#Else
    imlstSmallVB.ListImages.Add , , LoadPicture("i.ico")
#End If
    imlstLargeVB.ListImages.Add , , LoadPicture("i.ico")
    Set lstvwVB.SmallIcons = imlstSmallVB
    Set lstvwVB.Icons = imlstLargeVB
    lstvwVB.ListItems.Add , , "Basic", 1, 1
    
    imlstSmallFixed.ListImages.Add , , LoadAnyPicture("i.ico", eisSmall)
    imlstLargeFixed.ListImages.Add , , LoadAnyPicture("i.ico", eisShell)
    Set lstvwFixed.SmallIcons = imlstSmallFixed
    Set lstvwFixed.Icons = imlstLargeFixed
    lstvwFixed.ListItems.Add , , "Fixed", 1, 1
    chkSmall_Click
    chkSmall_Click
End Sub

Private Sub Form_Paint()
    BmpTile Me, imgPattern.Picture
End Sub

Sub PictureData()
With ctlCur
    Dim s As String
    ctlLast.BorderStyle = 0
    ctlCur.BorderStyle = 1
    s = s & "Picture Width: " & Int(ScaleX(.Picture.Width, 8, vbPixels)) & sCrLf
    s = s & "Picture Height: " & Int(ScaleY(.Picture.Height, 8, vbPixels)) & sCrLf
    s = s & TypeName(ctlCur) & " Width: " & Int(ScaleX(.Width, vbTwips, vbPixels)) & sCrLf
    s = s & TypeName(ctlCur) & " Height: " & Int(ScaleY(.Height, vbTwips, vbPixels)) & sCrLf
    lblData.Caption = s
End With
End Sub

Private Sub chkSmall_Click()
    If chkSmall.Value = vbChecked Then
        lstvwVB.View = lvwSmallIcon
        lstvwFixed.View = lvwSmallIcon
    Else
        lstvwVB.View = lvwIcon
        lstvwFixed.View = lvwIcon
    End If
End Sub

Private Sub img_Click(Index As Integer)
    Set ctlLast = ctlCur
    Set ctlCur = img(Index)
    PictureData
End Sub

Private Sub pb_Click(Index As Integer)
    Set ctlLast = ctlCur
    Set ctlCur = pb(Index)
    PictureData
End Sub

#If iVBVer > 5 Then
Function BasicLoadPicture(sFile As String) As IPictureDisp
    Set BasicLoadPicture = LoadPicture(sFile, cboSize.ListIndex, _
                                              cboDepth.ListIndex, _
                                              txtX, txtY)
    If ctlCur = pb(eipBasic) Then PictureData
    If ctlCur = img(eipBasic) Then PictureData
End Function
#End If


