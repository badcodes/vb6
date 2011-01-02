VERSION 5.00
Object = "{2DD06898-E157-11D0-8C51-00C04FC29CEC}#1.1#0"; "ListBoxPlus.ocx"
Begin VB.Form FCompletion 
   Caption         =   "Test Word Completion"
   ClientHeight    =   3192
   ClientLeft      =   912
   ClientTop       =   1476
   ClientWidth     =   4908
   Icon            =   "TCompletion.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3192
   ScaleWidth      =   4908
   Begin VB.TextBox txtLookup 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   1965
   End
   Begin ListBoxPlus.XListBoxPlus list 
      Height          =   2205
      Left            =   240
      TabIndex        =   0
      Top             =   705
      Width           =   1965
      _ExtentX        =   3471
      _ExtentY        =   3810
      BackColor       =   16777215
      ForeColor       =   -2147483640
      HiToLo          =   0   'False
      SortMode        =   2
      ListCount       =   13
      List0           =   "abacus"
      ItemData0       =   0
      List1           =   "absolutely"
      ItemData1       =   0
      List2           =   "abstract"
      ItemData2       =   0
      List3           =   "h"
      ItemData3       =   0
      List4           =   "he"
      ItemData4       =   0
      List5           =   "hell"
      ItemData5       =   0
      List6           =   "hello"
      ItemData6       =   0
      List7           =   "hot"
      ItemData7       =   0
      List8           =   "hotel"
      ItemData8       =   0
      List9           =   "one"
      ItemData9       =   0
      List10          =   "onerous"
      ItemData10      =   0
      List11          =   "only"
      ItemData11      =   0
      List12          =   "ontological"
      ItemData12      =   0
      ListIndex       =   -1
      Completion      =   0   'False
   End
   Begin VB.Label Label1 
      Caption         =   "Press Return in list box to complete the current word."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   2
      Left            =   2400
      TabIndex        =   4
      Top             =   1680
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Press Esc in text box to complete the current word."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   1
      Left            =   2400
      TabIndex        =   3
      Top             =   960
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Type in text box or list box. "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   2400
      TabIndex        =   2
      Top             =   240
      Width           =   2055
   End
End
Attribute VB_Name = "FCompletion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    list.Completion = True
    list.SortMode = esmlSortText
End Sub

Private Sub list_Click()
    'txtLookup = list.Text
End Sub

Private Sub list_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtLookup = list.Text
    Else
        txtLookup = list.PartialWord
        txtLookup.Refresh
    End If
End Sub

Private Sub txtLookup_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    ' Use Esc key to complete word
    Case 27
        KeyAscii = 0
        txtLookup = list.Text
        txtLookup.SelStart = 0
        txtLookup.SelLength = Len(txtLookup)
    ' Disable Enter key
    Case 13
        KeyAscii = 0
    End Select
End Sub

Private Sub txtLookup_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode <> 27 Then
        list.PartialWord = txtLookup
    End If
End Sub
