VERSION 5.00
Begin VB.Form cmdFrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BookUrl"
   ClientHeight    =   2040
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   5664
   Icon            =   "cmdFrm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   5664
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmUrl 
      Caption         =   "原始URL∶"
      Height          =   864
      Left            =   120
      TabIndex        =   6
      Top             =   132
      Width           =   5412
      Begin VB.TextBox txtUrl 
         Appearance      =   0  'Flat
         Height          =   516
         Left            =   108
         MultiLine       =   -1  'True
         TabIndex        =   7
         Top             =   204
         Width           =   5160
      End
   End
   Begin VB.CommandButton cmdReadText 
      Appearance      =   0  'Flat
      Caption         =   "->正文"
      Height          =   320
      Left            =   2364
      TabIndex        =   5
      Top             =   1536
      Width           =   828
   End
   Begin VB.CommandButton cmdReadCata 
      Appearance      =   0  'Flat
      Caption         =   "->目录"
      Height          =   320
      Left            =   1284
      TabIndex        =   4
      Top             =   1536
      Width           =   828
   End
   Begin VB.CommandButton cmdRandom 
      Appearance      =   0  'Flat
      Caption         =   "->随机页"
      Height          =   320
      Left            =   3444
      TabIndex        =   3
      Top             =   1536
      Width           =   828
   End
   Begin VB.TextBox txtBookName 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   852
      TabIndex        =   0
      Top             =   1080
      Width           =   4680
   End
   Begin VB.CommandButton cmdReadCov 
      Appearance      =   0  'Flat
      Caption         =   "->封面"
      Height          =   320
      Left            =   204
      TabIndex        =   2
      Top             =   1536
      Width           =   828
   End
   Begin VB.CommandButton cmdDownload 
      Appearance      =   0  'Flat
      Caption         =   "下载"
      Default         =   -1  'True
      Height          =   320
      Left            =   4524
      TabIndex        =   1
      Top             =   1536
      Width           =   828
   End
   Begin VB.Label Label1 
      Caption         =   "书名更正∶"
      Height          =   264
      Left            =   132
      TabIndex        =   8
      Top             =   1116
      Width           =   780
   End
End
Attribute VB_Name = "cmdFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Dim urlHander As CbookUrl
Const c_mdbFile As String = "download.mdb"

Private Sub cmdDownload_Click()
     
    urlHander.setHeader toDownload
    'MsgBox urlHander.toString
    
    ShellExecute Me.hwnd, "open", urlHander.toString, "", "", 1
    Call downloadLog(urlHander)
    Unload Me
End Sub

Private Sub cmdRandom_Click()

    Dim L As Integer
    On Error Resume Next
    Randomize
    L = CInt(urlHander.Pages)
    L = Int(Rnd() * L + 1)
    
    urlHander.Page = urlHander.page_Text(L)
    urlHander.setHeader toRead
    
    ShellExecute Me.hwnd, "open", urlHander.toString, "", "", 1
    Unload Me
    
    
End Sub

Private Sub cmdReadCov_Click()
    urlHander.Page = urlHander.page_Cover
    urlHander.setHeader toRead
    ShellExecute Me.hwnd, "open", urlHander.toString, "", "", 1
    Unload Me
End Sub

Private Sub cmdReadCata_Click()
    urlHander.Page = urlHander.page_Catalog
    urlHander.setHeader toRead
    ShellExecute Me.hwnd, "open", urlHander.toString, "", "", 1
    Unload Me
End Sub

Private Sub cmdReadText_Click()
    urlHander.Page = urlHander.page_Text(1)
    urlHander.setHeader toRead
    ShellExecute Me.hwnd, "open", urlHander.toString, "", "", 1
    Unload Me
End Sub

Private Sub Form_Load()
    Dim sUrl As String
    Dim sName As String
    Dim sText As String
    
    Dim sUrlPart() As String
    Dim iU As Long
    Dim iL As Long
    
    On Error Resume Next
    
    Set urlHander = New CbookUrl
    sUrl = Command$
    If Left$(sUrl, 1) = Chr(34) And Right$(sUrl, 1) = Chr(34) Then
        sUrl = Mid$(sUrl, 2, Len(sUrl) - 2)
    End If
    
    sUrlPart = Split(sUrl, "||")
    iU = UBound(sUrlPart)
    iL = LBound(sUrlPart)
    
    If iL > -1 And iU > iL Then
        sUrl = sUrlPart(iL)
        sName = sUrlPart(iL + 1)
        If iU > iL + 1 Then sText = sUrlPart(iL + 1 + 1)
    End If
        
    txtUrl.Text = sUrl
    urlHander.Initialize sUrl
    urlHander.Candownload = "1"
    
    If sText <> "" Then parseText sText, urlHander

    Randomize
    If urlHander.SS = "" Then urlHander.SS = CStr(str2lng(urlHander.Location) * 10 + CLng(urlHander.Pages))
    If urlHander.Author = "" Then urlHander.Author = Environ$("USERNAME")
    If sName = "" Then
        txtBookName.Text = urlHander.Bookname
    Else
        txtBookName.Text = sName
    End If
    
    
    
    Dim sClipboard As String
    sClipboard = Clipboard.GetText()
    If InStr(sClipboard, txtBookName.Text) > 0 Then
        txtBookName.Text = sClipboard
        Clipboard.Clear
    End If
    
    MWindows.setPosition Me, HWND_TOPMOST
End Sub

Private Sub Form_Terminate()
    Set urlHander = Nothing
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set urlHander = Nothing
End Sub

Public Sub Connect(ByVal sUrl As String)
    urlHander.Initialize sUrl
    urlHander.Candownload = "1"
    txtBookName.Text = urlHander.Bookname
End Sub

Private Sub txtBookName_Change()
    urlHander.Bookname = txtBookName.Text
End Sub

Private Function str2lng(ByRef str As String) As Long
    Dim i As Long
    Dim iEnd As Long
    Dim c As String
    Dim r As Long
    
    r = 0
    iEnd = Len(str)
    For i = 1 To iEnd
        c = Mid$(str, i, 1)
        r = r + AscW(c)
    Next
    If r < 0 Then r = 0 - r
    str2lng = r
    
End Function

Public Sub downloadLog(ByRef bUrl As CbookUrl)
    Dim sMDB As String
    Dim dbase As Database
    Dim tdef As TableDef
    Dim rc As Recordset
    
    sMDB = App.Path & "\" & c_mdbFile
    
    Set dbase = openDatabase(sMDB)
     
    If dbase Is Nothing Then
        Set dbase = newDatabase(sMDB)
        If dbase Is Nothing Then Exit Sub
        Set tdef = newTable(dbase, "Log")
        addField tdef, "ssid"
        addField tdef, "bookname"
        addField tdef, "author"
        addField tdef, "pages"
        addField tdef, "date"
        addField tdef, "url"
        addField tdef, "download"
        addTable dbase, tdef
        Set tdef = Nothing
    End If
    
    If dbase Is Nothing Then Exit Sub
    
    Set tdef = getTable(dbase, "Log")
    Set rc = getRecord(tdef)
    
    rc.AddNew
    With rc
        .Fields("download").Value = Date$ & " " & Time$
        .Fields("ssid").Value = bUrl.realSS
        .Fields("bookname").Value = bUrl.Bookname
        .Fields("author").Value = bUrl.Author
        .Fields("pages").Value = bUrl.Pages
        .Fields("date").Value = bUrl.getParam("date")
        .Fields("url").Value = bUrl.Url
    End With
    
    rc.Update
        
    Set rc = Nothing
    Set tdef = Nothing
    Set dbase = Nothing
    
End Sub


Private Sub parseText(ByVal sText As String, ByRef bUrl As CbookUrl)
' "※ 五代史通俗演义 发表评论  添加个人书签  点击次数：0  请双击书名浏览
'作者：蔡东藩　  出版社： 出版日期：1981年2月第1版  页数：584"
Dim sAuthor As String
Dim sDate As String
Dim sPage As String

sAuthor = RTrim$(LTrim$(LeftRange(sText, "作者：", "出版社", ReturnEmptyStr)))
sDate = RTrim$(LTrim$(LeftRange(sText, "出版日期：", "页数：", ReturnEmptyStr)))
sPage = RTrim$(LTrim$(LeftRight(sText, "页数：", ReturnEmptyStr)))

If sAuthor <> "" Then bUrl.Author = sAuthor
If sDate <> "" Then bUrl.setParam "date", sDate
If sPage <> "" Then bUrl.Pages = sPage

'MsgBox sAuthor & vbCrLf & sDate & vbCrLf & sPage

End Sub
