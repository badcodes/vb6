VERSION 5.00
Begin VB.Form MainFrm 
   Caption         =   "SS Mirror Download Helper"
   ClientHeight    =   3516
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   7680
   Icon            =   "MainFrm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3516
   ScaleWidth      =   7680
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frmOption 
      Caption         =   "估计设置"
      Height          =   780
      Left            =   132
      TabIndex        =   9
      Top             =   2160
      Width           =   7368
      Begin VB.TextBox txtOption 
         Height          =   288
         Index           =   4
         Left            =   6672
         TabIndex        =   19
         Text            =   "20"
         Top             =   252
         Width           =   564
      End
      Begin VB.TextBox txtOption 
         Height          =   288
         Index           =   3
         Left            =   5196
         TabIndex        =   17
         Text            =   "15"
         Top             =   264
         Width           =   564
      End
      Begin VB.TextBox txtOption 
         Height          =   288
         Index           =   2
         Left            =   3744
         TabIndex        =   15
         Text            =   "1"
         Top             =   264
         Width           =   564
      End
      Begin VB.TextBox txtOption 
         Height          =   288
         Index           =   1
         Left            =   2280
         TabIndex        =   13
         Text            =   "3"
         Top             =   276
         Width           =   564
      End
      Begin VB.TextBox txtOption 
         Height          =   288
         Index           =   0
         Left            =   804
         TabIndex        =   11
         Text            =   "3"
         Top             =   276
         Width           =   564
      End
      Begin VB.Label lblOption 
         Alignment       =   1  'Right Justify
         Caption         =   "目录数："
         Height          =   264
         Index           =   4
         Left            =   6036
         TabIndex        =   18
         Top             =   264
         Width           =   600
      End
      Begin VB.Label lblOption 
         Alignment       =   1  'Right Justify
         Caption         =   "前言数："
         Height          =   264
         Index           =   3
         Left            =   4584
         TabIndex        =   16
         Top             =   288
         Width           =   600
      End
      Begin VB.Label lblOption 
         Alignment       =   1  'Right Justify
         Caption         =   "版权页："
         Height          =   264
         Index           =   2
         Left            =   3108
         TabIndex        =   14
         Top             =   276
         Width           =   600
      End
      Begin VB.Label lblOption 
         Alignment       =   1  'Right Justify
         Caption         =   "书名页："
         Height          =   264
         Index           =   1
         Left            =   1644
         TabIndex        =   12
         Top             =   288
         Width           =   600
      End
      Begin VB.Label lblOption 
         Alignment       =   1  'Right Justify
         Caption         =   "封面页："
         Height          =   264
         Index           =   0
         Left            =   156
         TabIndex        =   10
         Top             =   288
         Width           =   600
      End
   End
   Begin VB.TextBox txtPages 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1260
      TabIndex        =   7
      Top             =   1680
      Width           =   6240
   End
   Begin VB.TextBox txtLocation 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1260
      TabIndex        =   5
      Top             =   1212
      Width           =   6240
   End
   Begin VB.TextBox txtTitle 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1260
      TabIndex        =   3
      Top             =   756
      Width           =   6240
   End
   Begin VB.Timer Timer 
      Interval        =   1000
      Left            =   132
      Top             =   1476
   End
   Begin VB.CommandButton cmdDownload 
      Caption         =   "下载 - Flashget"
      Height          =   324
      Left            =   5856
      TabIndex        =   1
      Top             =   3072
      Width           =   1632
   End
   Begin VB.TextBox txtUrl 
      Appearance      =   0  'Flat
      Height          =   468
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   144
      Width           =   7368
   End
   Begin VB.Label lblCopyright 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Left            =   0
      TabIndex        =   8
      Top             =   3060
      Width           =   5568
   End
   Begin VB.Label lblpages 
      Alignment       =   1  'Right Justify
      Caption         =   "页数："
      Height          =   252
      Left            =   168
      TabIndex        =   6
      Top             =   1716
      Width           =   924
   End
   Begin VB.Label lblLocation 
      Alignment       =   1  'Right Justify
      Caption         =   "地址："
      Height          =   252
      Left            =   168
      TabIndex        =   4
      Top             =   1260
      Width           =   924
   End
   Begin VB.Label lblTitle 
      Alignment       =   1  'Right Justify
      Caption         =   "书名："
      Height          =   252
      Left            =   168
      TabIndex        =   2
      Top             =   792
      Width           =   924
   End
End
Attribute VB_Name = "MainFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
'    Component  : MainFrm
'    Project    : SMDH
'
'    Description: ssLib Mirror Download Helper
'
'    Author   : xrLin
'
'    Date     : 2006-09
'--------------------------------------------------------------------------------
Option Explicit
Dim hSet As New CAutoSetting

Private Sub cmdDownload_Click()

Dim cText As String
Dim sHref As String
Dim sTitle As String
Dim iPages As Integer

Dim covs As Integer
Dim boks As Integer
Dim legs As Integer
Dim fows As Integer
Dim cats As Integer

Dim i As Integer

Dim pList()
Dim pCount As Long

'On Error Resume Next

covs = CInt(txtOption(0).Text)
boks = CInt(txtOption(1).Text)
legs = CInt(txtOption(2).Text)
fows = CInt(txtOption(3).Text)
cats = CInt(txtOption(4).Text)

'Clipboard.Clear
sTitle = txtTitle.Text
sHref = txtLocation.Text
iPages = CInt(txtPages.Text)

If LCase$(Left$(sHref, 5)) <> "http:" Or iPages = 0 Then
    MsgBox "无效链接或页数为0", vbCritical
    Exit Sub
End If

    Clipboard.Clear
    Clipboard.SetText sTitle

ReDim pList(0)
pList(0) = sHref

For i = 1 To covs
    cText = sHref & "/" & "cov" & StrNum(i, 3) & ".pdg"
    pCount = pCount + 1
    ReDim Preserve pList((pCount) * 2)
    pList((pCount - 1) * 2 + 1) = cText
    pList((pCount - 1) * 2 + 2) = ""
Next

For i = 1 To boks
    cText = sHref & "/" & "bok" & StrNum(i, 3) & ".pdg"
    pCount = pCount + 1
    ReDim Preserve pList((pCount) * 2)
    pList((pCount - 1) * 2 + 1) = cText
    pList((pCount - 1) * 2 + 2) = ""
Next

For i = 1 To legs
    cText = sHref & "/" & "leg" & StrNum(i, 3) & ".pdg"
    pCount = pCount + 1
    ReDim Preserve pList((pCount) * 2)
    pList((pCount - 1) * 2 + 1) = cText
    pList((pCount - 1) * 2 + 2) = ""
Next

For i = 1 To fows
    cText = sHref & "/" & "fow" & StrNum(i, 3) & ".pdg"
    pCount = pCount + 1
    ReDim Preserve pList((pCount) * 2)
    pList((pCount - 1) * 2 + 1) = cText
    pList((pCount - 1) * 2 + 2) = ""
Next

For i = 1 To cats
    cText = sHref & "/" & "!" & StrNum(i, 5) & ".pdg"
    pCount = pCount + 1
    ReDim Preserve pList((pCount) * 2)
    pList((pCount - 1) * 2 + 1) = cText
    pList((pCount - 1) * 2 + 2) = ""
Next

For i = 1 To iPages
    cText = sHref & "/" & StrNum(i, 6) & ".pdg"
    pCount = pCount + 1
    ReDim Preserve pList((pCount) * 2)
    pList((pCount - 1) * 2 + 1) = cText
    pList((pCount - 1) * 2 + 2) = ""
Next



Dim jet As New JetCarNetscape
jet.AddUrlList pList()

Set jet = Nothing


End Sub

Private Sub Form_Load()
With App
lblCopyright.Caption = .ProductName & " v" & .Major & "." & .Minor
End With


Dim i As Integer

With hSet
    .fileNameSaveTo = bddir(App.Path) & App.EXEName & ".ini"
    For i = 0 To txtOption.Count
        .Add txtOption(i), SF_TEXT
    Next
    .Add txtTitle, SF_TEXT
    .Add txtLocation, SF_TEXT
    .Add txtPages, SF_TEXT
End With

End Sub

Private Sub Form_Unload(Cancel As Integer)

Set hSet = Nothing

End Sub

Private Sub Timer_Timer()
Dim cText As String
Dim sUrl As String
On Error Resume Next
sUrl = txtUrl.Text
cText = Clipboard.GetText

'book://ssreader/e0?url=http://dl3.lib.tongji.edu.cn/cx210k/07/diskjsj/js106/01/!00001.pdg&&&&&pages=493&bookname=计算实用教程Visual
If LCase$(Left$(cText, 5)) = "book:" And sUrl <> cText Then
    txtUrl.Text = cText
    MainFrm.Show
End If

If txtLocation.Text <> "" And txtPages <> "" Then cmdDownload.Enabled = True Else cmdDownload.Enabled = False
If txtLocation.Text = "" Then
    lblLocation.ForeColor = &HFF
    lblLocation.Caption = "请指定位置："
Else
    lblLocation.ForeColor = &H0
    lblLocation.Caption = "位置："
End If

If txtPages.Text = "" Then
    lblpages.ForeColor = &HFF
    lblpages.Caption = "请指定页数："
Else
    lblpages.ForeColor = &H0
    lblpages.Caption = "页数："
End If

End Sub

Private Sub txtUrl_Change()
Dim sPage As String
Dim sHref As String
Dim sTitle As String

sHref = LeftRange(txtUrl.Text, "e0?url=", "/!", vbTextCompare, ReturnEmptyStr)
sTitle = LeftRange(txtUrl.Text, "bookname=", "&", vbTextCompare, ReturnEmptyStr)
If sTitle = "" Then sTitle = LeftRight(txtUrl.Text, "bookname=", vbTextCompare, ReturnEmptyStr)

sPage = LeftRange(txtUrl.Text, "pages=", "&", vbTextCompare, ReturnEmptyStr)
If sPage = "" Then sPage = LeftRight(txtUrl.Text, "pages", vbTextCompare, ReturnEmptyStr)


txtTitle.Text = sTitle
txtLocation.Text = sHref
txtPages.Text = sPage

End Sub
