VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "FormatBlog"
   ClientHeight    =   8775
   ClientLeft      =   45
   ClientTop       =   450
   ClientWidth     =   11565
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8775
   ScaleWidth      =   11565
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog dlgFilename 
      Left            =   5535
      Top             =   4140
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtShow 
      Height          =   7335
      Index           =   0
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Text            =   "frmMain.frx":0000
      Top             =   1005
      Width           =   11280
   End
   Begin VB.TextBox txtShow 
      Height          =   7335
      Index           =   3
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Text            =   "frmMain.frx":0007
      Top             =   1005
      Width           =   11280
   End
   Begin VB.TextBox txtShow 
      Height          =   7335
      Index           =   2
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Text            =   "frmMain.frx":000D
      Top             =   1005
      Width           =   11280
   End
   Begin VB.TextBox txtShow 
      Height          =   7335
      Index           =   1
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Text            =   "frmMain.frx":0014
      Top             =   1005
      Width           =   11280
   End
   Begin ComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   8400
      Width           =   11565
      _ExtentX        =   20399
      _ExtentY        =   661
      Style           =   1
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   20346
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TabStrip TabStrip 
      Height          =   390
      Left            =   120
      TabIndex        =   3
      Top             =   615
      Width           =   11280
      _ExtentX        =   19897
      _ExtentY        =   688
      TabWidthStyle   =   2
      MultiRow        =   -1  'True
      ShowTips        =   0   'False
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Source"
            Key             =   "source"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "CNWeblog"
            Key             =   "cnweblog"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "TianYa"
            Key             =   "tianya"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "QZone"
            Key             =   "qzone"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdFilename 
      Appearance      =   0  'Flat
      Caption         =   "Select File..."
      Height          =   360
      Left            =   9465
      TabIndex        =   1
      Top             =   120
      Width           =   1950
   End
   Begin VB.TextBox txtFilename 
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   9135
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private mTextChanged() As Boolean
Private mSize As Long

Private Sub cmdFilename_Click()
    Dim filename As String
    Dim text As String
    dlgFilename.filename = txtFilename.text
    dlgFilename.ShowOpen
    If dlgFilename.filename <> "" Then
        filename = dlgFilename.filename
        txtFilename.text = filename
        Dim fNum As Integer
        fNum = FreeFile
        Open filename For Binary As #fNum
        text = String$(LOF(fNum), " ")
        Get #fNum, , text
        Close #fNum
        txtShow(0).text = text
        Call TabStrip_Click
    End If
End Sub

Private Sub Form_Load()

    ReDim mTextChanged(0 To txtShow.UBound)
    mTextChanged(0) = False
    
'    MLayout.Init Me
'    MLayout.PutControl txtFilename, "80%", 360, False
'    MLayout.PutControl cmdFilename, "0", 360, False
'    MLayout.PutControl TabStrip, "0", 480, True
'    MLayout.PutControl txtShow, "0", "0", True
End Sub

Private Sub TabStrip_Click()
'    Dim i As Integer
'    For i = 0 To txtShow.UBound
'        txtShow(i).ZOrder 1
'    Next

    If mTextChanged(TabStrip.SelectedItem.Index - 1) Then
        Call Text_Changed
    End If
    txtShow(TabStrip.SelectedItem.Index - 1).ZOrder 0
    mSize = Len(txtShow(TabStrip.SelectedItem.Index - 1).text)
    
    StatusBar.SimpleText = CStr(mSize) + " Charactors count"
End Sub

Private Sub Text_Changed()
    CallByName Me, TabStrip.SelectedItem.key, VbMethod, GetTabIndex(TabStrip.SelectedItem.key)
End Sub


Private Function GetTabIndex(ByVal key As String) As Integer
    GetTabIndex = TabStrip.Tabs(key).Index - 1
End Function

Private Sub txtShow_Change(Index As Integer)
    If Index = 0 Then
        Dim i As Integer
        For i = 1 To txtShow.UBound
            mTextChanged(i) = True
        Next
    End If
End Sub

Public Sub cnweblog(ByVal Index As Integer)
    Dim aSource() As String
    aSource = Split(txtShow(0).text, vbCrLf)
    Dim result As String
    Dim u As Integer
    Dim i As Integer
    u = UBound(aSource)
    For i = 0 To u
        aSource(i) = Trim$(aSource(i))
        If aSource(i) <> "" Then
            result = result + "<p>" + aSource(i) + "</p>" + vbCrLf + vbCrLf
        End If
    Next
    
    txtShow(Index).text = result
    
    mTextChanged(Index) = False
End Sub

Public Sub tianya(ByVal Index As Integer)
    Dim aSource() As String
    aSource = Split(txtShow(0).text, vbCrLf)
    Dim result As String
    Dim u As Integer
    Dim i As Integer
    u = UBound(aSource)
    For i = 0 To u
        aSource(i) = Trim$(aSource(i))
        If aSource(i) <> "" Then
            result = result + "¡¡¡¡" + aSource(i) + vbCrLf + vbCrLf
        End If
    Next
    
    txtShow(Index).text = result
    
    mTextChanged(Index) = False
End Sub

Public Sub qzone(ByVal Index As Integer)
    Call tianya(Index)
End Sub

