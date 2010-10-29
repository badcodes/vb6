VERSION 5.00
Begin VB.Form FTestResources 
   ClientHeight    =   4140
   ClientLeft      =   1092
   ClientTop       =   2076
   ClientWidth     =   5364
   Icon            =   "TRes.frx":0000
   LinkTopic       =   "frmRes"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4140
   ScaleWidth      =   5364
   Begin VB.CommandButton cmdExit 
      Height          =   495
      Left            =   3825
      TabIndex        =   3
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton cmdSound 
      Height          =   495
      Left            =   3840
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox txtDump 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2064
      Left            =   180
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   1944
      Width           =   4965
   End
   Begin VB.ListBox lstStrings 
      Height          =   816
      Left            =   1320
      TabIndex        =   0
      Top             =   435
      Width           =   2190
   End
   Begin VB.Label lblStrings 
      Height          =   240
      Left            =   1335
      TabIndex        =   4
      Top             =   45
      Width           =   1185
   End
   Begin VB.Image imgMascot 
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Left            =   300
      Top             =   270
      Width           =   885
   End
   Begin VB.Menu mnuFile 
      Caption         =   ""
      Begin VB.Menu mnuSound 
         Caption         =   ""
      End
      Begin VB.Menu mnuExit 
         Caption         =   ""
      End
   End
End
Attribute VB_Name = "FTestResources"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Enum ETestResource

    ordAppBmp = 101
    ordAppMeta = 201
    ordAppIcon = 301
    ordAppCursor = 401
    ordWavGrunt = 501
    ordTxtData = 601
    ordMetaTypewrite = 701
    ordAviDrill = 801

    ordFrmTitle = 1001
    ordMnuFile = 1101
    ordMnuGrunt = 1102
    ordMnuExit = 1103
    ordLstTitle = 1201
    ordLstWhat = 1301
    ordLstWhy = 1302
    ordLstWhere = 1303
    ordLstWho = 1304
    ordLstWhen  = 1305
End Enum

Private abWavGrunt() As Byte
Private abDump() As Byte
Private asText(1 To 5) As String

Private Sub Form_Load()

    App.Title = LoadResString(ordFrmTitle)
    Me.Caption = LoadResString(ordFrmTitle)
    Me.MousePointer = vbCustom
    Me.MouseIcon = LoadResPicture(ordAppCursor, vbResCursor)
    Me.Icon = LoadResPicture(ordAppIcon, vbResIcon)

    mnuFile.Caption = LoadResString(ordMnuFile)
    mnuSound.Caption = LoadResString(ordMnuGrunt)
    mnuExit.Caption = LoadResString(ordMnuExit)
    cmdSound.Caption = LoadResString(ordMnuGrunt)
    cmdExit.Caption = LoadResString(ordMnuExit)
    
    lblStrings.Caption = LoadResString(ordLstTitle)
    Dim i As Integer
    For i = 1 To 5
        lstStrings.AddItem LoadResString(ordLstWhat + i - 1)
    Next
    
    imgMascot.Picture = LoadResPicture(ordAppBmp, vbResBitmap)
    
    abDump = LoadResData(ordTxtData, "OURDATA")
    
    abWavGrunt = LoadResData(ordWavGrunt, "WAVE")
    txtDump.Text = HexDump(abDump, False)

End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Image1_Click()

End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub cmdSound_Click()
    PlayWave abWavGrunt
End Sub

Private Sub mnuSound_Click()
    PlayWave abWavGrunt
End Sub

Function PlayWave(ab() As Byte) As Boolean
    PlayWave = sndPlaySoundAsBytes(ab(0), SND_MEMORY Or SND_SYNC)
End Function
'
