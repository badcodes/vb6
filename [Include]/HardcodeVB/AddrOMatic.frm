VERSION 5.00
Begin VB.Form FAddressOMatic 
   Caption         =   "Address-o-matic"
   ClientHeight    =   3180
   ClientLeft      =   3225
   ClientTop       =   3450
   ClientWidth     =   3015
   Icon            =   "AddrOMatic.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3180
   ScaleWidth      =   3015
   Begin VB.TextBox txtBlocks 
      Height          =   324
      Left            =   96
      TabIndex        =   5
      Text            =   "1"
      Top             =   1104
      Width           =   2775
   End
   Begin VB.TextBox txtBaseAddr 
      Height          =   324
      Left            =   96
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   1884
      Width           =   2775
   End
   Begin VB.TextBox txtReserved 
      Height          =   324
      Left            =   96
      TabIndex        =   1
      Text            =   "65536"
      Top             =   360
      Width           =   2775
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "New Base Address"
      Default         =   -1  'True
      Height          =   495
      Left            =   732
      TabIndex        =   0
      Top             =   2472
      Width           =   1695
   End
   Begin VB.Label lbl 
      Caption         =   "64 KB blocks:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   204
      Index           =   2
      Left            =   96
      TabIndex        =   6
      Top             =   852
      Width           =   1452
   End
   Begin VB.Label lbl 
      Caption         =   "Base address:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   204
      Index           =   1
      Left            =   96
      TabIndex        =   3
      Top             =   1632
      Width           =   1812
   End
   Begin VB.Label lbl 
      Caption         =   "Bytes to reserve:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   204
      Index           =   0
      Left            =   96
      TabIndex        =   2
      Top             =   108
      Width           =   1452
   End
End
Attribute VB_Name = "FAddressOMatic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Function DllBaseAddress(Optional ByVal Size As Long = 65535) As String
    Dim iBase As Long, fDone As Boolean
    
    If Size < 65536 Then
        Size = 1
    Else
        ' Reduce Size by factor of 64K and round up
        Size = (Size \ 65536) - (Size Mod 65536 <> 0)
    End If
    
    Do
        ' Pick iBase from range available to component developers
        iBase = Random(256, 32768 - Size)
        ' Be sure iBase doesn't fall within unavailable ranges
        
        ' 0x00000000 - 0x0032FFFF  Crystal Reports
        If (iBase >= &H0) And (iBase + Size <= &H32) Then
            fDone = False
        ' 0x0F9A0000 - 0x0FFFFFFF  VBA components
        ElseIf (iBase >= &HF9A) And (iBase + Size <= &HFFF) Then
            fDone = False
        ' 0x0F000000 - 0x0F8BFFFF  Core VB components
        ElseIf (iBase >= &HF00) And (iBase + Size <= &HF8B) Then
            fDone = False
        ' 0x20000000 - 0x24FFFFFF  VB controls
        ElseIf (iBase >= &H2000) And (iBase + Size <= &H24FF) Then
            fDone = False
        ' 0x25000000 - 0x26FFFFFF  Crystal Reports
        ElseIf (iBase >= &H2500) And (iBase + Size <= &H26FF) Then
            fDone = False
        ' 0x2E8B0000 - 0x2E9AFFFF  Hardcore components
        ElseIf (iBase >= &H2E8B) And (iBase + Size <= &H2E9A) Then
            fDone = False
        ' 0x65000000 - 0x65FFFFFF  Office 97 components
        ElseIf (iBase >= &H6500) And (iBase + Size <= &H65FF) Then
            fDone = False
        ' Insert your range here
        'ElseIf (iBase >= &Hxxxx) And (iBase + Size <= &Hxxxx) Then
        '    fDone = False
        Else
            fDone = True
        End If
    Loop While Not fDone
    
    DllBaseAddress = "&H" & Right$(String$(4, "0") & Hex$(iBase), 4) & "0000"
End Function

Private Sub cmdNew_Click()
    txtBaseAddr.SetFocus
    txtBaseAddr = DllBaseAddress(txtReserved)
    Clipboard.SetText txtBaseAddr
End Sub

Private Sub txtBlocks_LostFocus()
    Dim cBlocks As Long, cBytes As Long
    cBlocks = Val(txtBlocks)
    cBytes = cBlocks * 65536
    txtReserved = CStr(cBytes)
End Sub

Private Sub txtReserved_LostFocus()
    Dim cBlocks As Long, cBytes As Long
    cBytes = txtReserved
    cBlocks = (cBytes \ 65536) - (cBytes Mod 65536 <> 0)
    txtBlocks = cBlocks
End Sub

Private Sub txtBlocks_GotFocus()
    txtBlocks.SelStart = 0
    txtBlocks.SelLength = 30
End Sub

Private Sub txtReserved_GotFocus()
    txtReserved.SelStart = 0
    txtReserved.SelLength = 30
End Sub

