VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Wenxin"
   ClientHeight    =   2280
   ClientLeft      =   48
   ClientTop       =   348
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2280
   ScaleWidth      =   6000
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAddUrl 
      Caption         =   "AddUrl To OE"
      Default         =   -1  'True
      Height          =   288
      Left            =   4104
      TabIndex        =   9
      Top             =   1800
      Width           =   1776
   End
   Begin VB.TextBox txtPages 
      Height          =   288
      Left            =   852
      TabIndex        =   7
      Text            =   "60"
      Top             =   1776
      Width           =   1560
   End
   Begin VB.TextBox txtBaseUrl 
      Height          =   288
      Left            =   864
      TabIndex        =   5
      Text            =   "http://www.wenku.biz/bookroom.php?aclass=1&initial=*&page=$*$"
      Top             =   1368
      Width           =   5040
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "View"
      Height          =   288
      Left            =   4080
      TabIndex        =   1
      Top             =   564
      Width           =   1788
   End
   Begin VB.TextBox txtNUM 
      Height          =   288
      Left            =   852
      TabIndex        =   0
      Text            =   "25986"
      Top             =   564
      Width           =   1560
   End
   Begin VB.TextBox txtBase 
      Height          =   288
      Left            =   840
      TabIndex        =   2
      Text            =   "http://67.19.223.125/htmpage/"
      Top             =   108
      Width           =   5040
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "页数:"
      Height          =   240
      Left            =   108
      TabIndex        =   8
      Top             =   1812
      Width           =   624
   End
   Begin VB.Line Line1 
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      X1              =   108
      X2              =   5868
      Y1              =   1152
      Y2              =   1152
   End
   Begin VB.Label Label3 
      Caption         =   "基本地址:"
      Height          =   252
      Left            =   120
      TabIndex        =   6
      Top             =   1368
      Width           =   732
   End
   Begin VB.Label Label2 
      Caption         =   "文章序号:"
      Height          =   240
      Left            =   108
      TabIndex        =   4
      Top             =   600
      Width           =   624
   End
   Begin VB.Label Label1 
      Caption         =   "基本地址:"
      Height          =   252
      Left            =   120
      TabIndex        =   3
      Top             =   156
      Width           =   732
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const CONFIG = "config.ini"



Private Sub cmdAddUrl_Click()
Dim baseUrl As String
Dim iEnd As Integer
On Error Resume Next
baseUrl = txtBaseUrl.Text
iEnd = CInt(txtPages.Text)
makeWenxinPageforOE baseUrl, 1, iEnd
End Sub

Private Sub cmdView_Click()
Dim iNum As Long
Dim sAppend As String
If txtNUM.Text = "" Then Exit Sub
iNum = CLng(txtNUM.Text)
If iNum = 0 Then Exit Sub
txtNUM.Text = LTrim(CStr(iNum))
If txtNUM.Text = "" Then Exit Sub
If iNum < 1000 Then
    sAppend = Left$(txtNUM.Text, 1)
ElseIf iNum < 10000 Then
    sAppend = Left$(txtNUM.Text, 2)
ElseIf iNum < 100000 Then
    sAppend = Left$(txtNUM.Text, 3)
ElseIf iNum < 1000000 Then
    sAppend = Left$(txtNUM.Text, 4)
End If
sAppend = sAppend & "/" & txtNUM.Text & "/index.htm"
Shell "explorer.exe " & txtBase.Text & sAppend, vbNormalFocus
End Sub

Private Sub Form_Load()
Dim sIni As String
sIni = App.Path & "\" & CONFIG
txtBase.Text = iniGetSetting(sIni, "Last", "BaseText")
txtNUM.Text = iniGetSetting(sIni, "Last", "NUM")
MWindows.setPosition Me, HWND_TOPMOST
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim sIni As String
sIni = App.Path & "\" & CONFIG
iniSaveSetting sIni, "Last", "BaseText", txtBase.Text
iniSaveSetting sIni, "Last", "NUM", txtNUM.Text
End Sub

Public Function makeWenxinPageforOE(ByRef baseUrl As String, Optional iStart As Integer = 1, Optional iEnd As Integer = 8)
Dim i As Integer
Dim href As String
Dim hOE As New OELib.OfflineExplorerAddUrl
For i = iStart To iEnd
href = Replace(baseUrl, "$*$", LTrim(Str(i)))
hOE.AddUrl href, href, href
Next
Set hOE = Nothing
End Function
